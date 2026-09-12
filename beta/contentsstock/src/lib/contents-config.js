const KEY = 'contents:config'

const DEFAULTS = {
  gasUrl: '',
  accessToken: '', // GASの ACCESS_TOKEN と一致させる共有トークン(Notionのシークレットではない)
  code: '', // 権限コード。GAS側スクリプトプロパティ "code" と照合する
  role: '', // 検証済みの権限。'xYz' は管理者、'err' は権限なし
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

export function isConfigured(c) {
  return Boolean(c.gasUrl && c.accessToken)
}

export const ADMIN_ROLE = 'xYz'

export function isAdmin(c) {
  return c.role === ADMIN_ROLE
}

/**
 * 生成・編集ができるか。
 * コードを使わない運用(個人で使う場合)でも触れるように、
 * 「明示的に弾かれたときだけ読み取り専用」にしている。
 */
export function canEdit(c) {
  return c.role !== 'err'
}
