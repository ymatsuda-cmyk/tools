const KEY = 'gemma-chat.settings'

export const DEFAULT_SETTINGS = {
  baseUrl: 'https://pregnant-vindicate-deacon.ngrok-free.dev/v1',
  apiKey: '',
  model: 'gemma4:12b',
  numCtx: 32768,
  temperature: 0.7,
  systemPrompt: '必ず日本語で回答してください。',
}

export function loadSettings() {
  try {
    const raw = localStorage.getItem(KEY)
    return raw ? { ...DEFAULT_SETTINGS, ...JSON.parse(raw) } : { ...DEFAULT_SETTINGS }
  } catch {
    return { ...DEFAULT_SETTINGS }
  }
}

export function saveSettings(s) {
  localStorage.setItem(KEY, JSON.stringify(s))
}

/**
 * APIキーの検証。
 * コピペで末尾に改行や全角文字が混入すると、ブラウザからは
 * ネットワークエラーとしか見えず原因究明に時間がかかる。入口で弾く。
 */
export function validateApiKey(raw) {
  const value = String(raw).trim()
  if (!value) return { ok: false, value, error: 'APIキーを入力してください' }
  if (!/^[\x21-\x7e]+$/.test(value)) {
    return { ok: false, value, error: '使用できない文字が含まれています（改行・空白・全角など）' }
  }
  return { ok: true, value }
}

export function normalizeBaseUrl(raw) {
  return String(raw).trim().replace(/\/+$/, '')
}
