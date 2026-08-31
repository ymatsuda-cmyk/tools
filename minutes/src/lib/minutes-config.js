const KEY = 'minutes:config'

const DEFAULTS = {
  gasUrl: '',
  notionToken: '', // GASの ACCESS_TOKEN と一致させる共有トークン(Notionのシークレットではない)
  code: '', // 権限コード。GAS側スクリプトプロパティ "code" と照合する
  role: '', // 検証済みの権限。'xYz' は管理者、'err' は権限なし
  lastSyncAt: null,
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
  return Boolean(c.gasUrl && c.notionToken)
}

export const ADMIN_ROLE = 'xYz'

/** 全機能を使える管理者か */
export function isAdmin(c) {
  return c.role === ADMIN_ROLE
}

/** 権限が無い(未入力または未登録コード)か */
export function isDenied(c) {
  return !c.role || c.role === 'err'
}
