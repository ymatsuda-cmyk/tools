const KEY = 'minutes:config'

const DEFAULTS = {
  gasUrl: '',
  notionToken: '', // GASの ACCESS_TOKEN と一致させる共有トークン(Notionのシークレットではない)
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
