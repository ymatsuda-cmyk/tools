import { fetchBin, updateBin } from './jsonbin.js'
import { loadSettings, saveSettings } from './llm-settings.js'
import { exportPromptOverrides, importPromptOverrides } from './prompts.js'
import { encryptConfig, decryptConfig, isEnvelope } from './jsonbin-crypto.js'

const KEY = 'videos:config'
const PASS_KEY = 'videos:jsonbin-pass' // Binの中身を暗号化するパスフレーズ。configには含めない(この端末にだけ置く)

const DEFAULTS = {
  gasUrl: '',
  accessToken: '', // GASの ACCESS_TOKEN と一致させる共有トークン(Notionのシークレットではない)
  code: '', // 権限コード。GAS側スクリプトプロパティ "code" と照合する
  role: '', // 検証済みの権限。'xYz' は管理者、'err' は権限なし
  dataUrl: '', // 一覧JSONの置き場所。空なら同じリポジトリの data/clipstock/ を見る
  useJsonbin: false, // trueなら、この設定オブジェクト自体をJSONBinとも同期する(他端末と共有するため)
  jsonbinBinId: '', // jsonbin.io のBin ID
  jsonbinApiKey: '', // jsonbin.io のAPIキー(X-Master-Key)
}

export function loadConfig() {
  try {
    const raw = localStorage.getItem(KEY)
    return raw ? { ...DEFAULTS, ...JSON.parse(raw) } : { ...DEFAULTS }
  } catch {
    return { ...DEFAULTS }
  }
}

export function saveConfig(c) {
  localStorage.setItem(KEY, JSON.stringify(c))
}

/** JSONBinへ同期しに行けるだけの情報が揃っているか */
export function jsonbinReady(c = loadConfig()) {
  return Boolean(c.useJsonbin && c.jsonbinBinId && c.jsonbinApiKey)
}

/** Binを暗号化するパスフレーズ。configには含めず、この端末のブラウザにだけ保存する */
export function loadJsonbinPassphrase() {
  try {
    return localStorage.getItem(PASS_KEY) || ''
  } catch {
    return ''
  }
}

export function saveJsonbinPassphrase(p) {
  try {
    if (p) localStorage.setItem(PASS_KEY, p)
    else localStorage.removeItem(PASS_KEY)
  } catch {
    // 保存できなくてもこのセッションの同期自体は続ける
  }
}

/**
 * アプリ起動時に呼ぶ。JSONBinを使う設定になっていれば、クラウド側の内容で
 * localStorageの設定を上書きしてから返す。使っていない/取得に失敗したときは
 * ローカルの設定をそのまま返す(オフラインでも今まで通り動かすため)。
 *
 * useJsonbin・jsonbinBinId・jsonbinApiKey はクラウドへ聞きに行くための鍵なので、
 * クラウド側の値では上書きしない(常にこの端末で入力したものを使う)。
 *
 * Binには config だけでなく、AI接続(llm-settings.js)とAIへの指示(prompts.jsの上書き)も
 * 並べて入れてあるので、そちらも合わせて取り込む。
 * 中身が暗号文(封筒)なら、この端末のパスフレーズで復号してから取り込む。
 */
export async function syncConfigFromJsonbin() {
  const local = loadConfig()
  if (!jsonbinReady(local)) return local
  try {
    return await pullFromJsonbin(local, loadJsonbinPassphrase())
  } catch (err) {
    console.warn('JSONBinから設定を取得できなかったため、この端末の設定を使います:', err)
    return local
  }
}

/**
 * Binを読み、(暗号文なら復号して)この端末のlocalStorageへ取り込む。失敗したら例外を投げる。
 * 起動時の同期(失敗してもローカルで続行)と、初回設定画面(失敗を画面に出す)の両方から使う。
 */
async function pullFromJsonbin(local, passphrase) {
  let remote = await fetchBin(local.jsonbinBinId, local.jsonbinApiKey)
  if (isEnvelope(remote)) {
    if (!passphrase) throw new Error('JSONBinの中身は暗号化されています。パスフレーズを入力してください')
    remote = await decryptConfig(remote, passphrase)
  }
  const { llmSettings, promptOverrides, ...remoteConfig } = remote || {}
  const merged = {
    ...local,
    ...remoteConfig,
    useJsonbin: local.useJsonbin,
    jsonbinBinId: local.jsonbinBinId,
    jsonbinApiKey: local.jsonbinApiKey,
  }
  saveConfig(merged)
  if (llmSettings) saveSettings(llmSettings)
  if (promptOverrides) importPromptOverrides(promptOverrides)
  return merged
}

/**
 * 初回設定画面から使う。Bin ID・APIキー・パスフレーズでBinを読み、
 * 保存してある設定をこの端末に取り込んで、以降の変更もJSONBinへ書く状態(useJsonbin)にする。
 * 読めなかった(キー違い・パスフレーズ違いなど)ときは、この端末の設定を何も変えずに例外を投げる。
 */
export async function connectJsonbin({ binId, apiKey, passphrase }) {
  const probe = { ...loadConfig(), useJsonbin: true, jsonbinBinId: binId, jsonbinApiKey: apiKey }
  const merged = await pullFromJsonbin(probe, passphrase) // 復号できて初めて何か書く
  saveJsonbinPassphrase(passphrase)
  return merged
}

/**
 * 設定を保存したときに呼ぶ。JSONBinを使う設定なら、いま保存した内容をクラウド側にも書き込む。
 * 使わない設定なら何もしない。失敗したときは例外を投げるので、呼び出し側で状況を伝えること
 * (この端末への保存はすでに済んでいるので、ここで失敗してもデータは失われない)。
 *
 * config に加えて llmSettings(AI接続)と promptOverrides(AIへの指示)もまとめて書き込む。
 * パスフレーズが設定されていれば暗号化して登録する。
 */
export async function pushConfigToJsonbin(c) {
  if (!jsonbinReady(c)) return
  const payload = {
    ...c,
    llmSettings: loadSettings(),
    promptOverrides: exportPromptOverrides(),
  }
  const passphrase = loadJsonbinPassphrase()
  const body = passphrase ? await encryptConfig(payload, passphrase) : payload
  await updateBin(c.jsonbinBinId, c.jsonbinApiKey, body)
}

export function isConfigured(c) {
  return Boolean(c.gasUrl && c.accessToken)
}

export const ADMIN_ROLE = 'xYz'

/** 全機能(生成・編集・状態変更)を使える管理者か */
export function isAdmin(c) {
  return c.role === ADMIN_ROLE
}

/** 権限が無い(未入力または未登録コード)か */
export function isDenied(c) {
  return !c.role || c.role === 'err'
}

/**
 * 生成・編集ができるか。
 * コードを使わない運用(個人で使う場合)でも触れるように、
 * 「明示的に弾かれたときだけ読み取り専用」にしている。
 */
export function canEdit(c) {
  return c.role !== 'err'
}
