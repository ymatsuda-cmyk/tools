import { fetchBin, updateBin } from './jsonbin.js'

const KEY = 'videos:config'

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

/**
 * アプリ起動時に呼ぶ。JSONBinを使う設定になっていれば、クラウド側の内容で
 * localStorageの設定を上書きしてから返す。使っていない/取得に失敗したときは
 * ローカルの設定をそのまま返す(オフラインでも今まで通り動かすため)。
 *
 * useJsonbin・jsonbinBinId・jsonbinApiKey はクラウドへ聞きに行くための鍵なので、
 * クラウド側の値では上書きしない(常にこの端末で入力したものを使う)。
 */
export async function syncConfigFromJsonbin() {
  const local = loadConfig()
  if (!jsonbinReady(local)) return local
  try {
    const remote = await fetchBin(local.jsonbinBinId, local.jsonbinApiKey)
    const merged = {
      ...local,
      ...remote,
      useJsonbin: local.useJsonbin,
      jsonbinBinId: local.jsonbinBinId,
      jsonbinApiKey: local.jsonbinApiKey,
    }
    saveConfig(merged)
    return merged
  } catch (err) {
    console.warn('JSONBinから設定を取得できなかったため、この端末の設定を使います:', err)
    return local
  }
}

/**
 * 設定を保存したときに呼ぶ。JSONBinを使う設定なら、いま保存した内容をクラウド側にも書き込む。
 * 使わない設定なら何もしない。失敗したときは例外を投げるので、呼び出し側で状況を伝えること
 * (この端末への保存はすでに済んでいるので、ここで失敗してもデータは失われない)。
 */
export async function pushConfigToJsonbin(c) {
  if (!jsonbinReady(c)) return
  await updateBin(c.jsonbinBinId, c.jsonbinApiKey, c)
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
