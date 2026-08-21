const KEY = 'gemma-chat.settings'

function makeId() {
  return Math.random().toString(36).slice(2, 8)
}

export function newProfile(over = {}) {
  return {
    id: makeId(),          // GAS の ENDPOINTS の id と一致させる
    label: '新しい接続',
    baseUrl: '',
    apiKey: '',
    model: '',
    numCtx: 32768,
    ...over,
  }
}

export const DEFAULT_SETTINGS = {
  profiles: [],
  activeId: null,
  gasUrl: '',
  controlToken: '',
  temperature: 0.7,
  systemPrompt: '必ず日本語で回答してください。',
}

/** 単一接続だった旧形式を profiles 配列へ移行する */
function migrate(s) {
  if (Array.isArray(s.profiles) && s.profiles.length) return s
  if (!s.baseUrl && !s.apiKey) return { ...s, profiles: [], activeId: null }
  const p = newProfile({
    id: 'a',
    label: s.model || 'メイン',
    baseUrl: s.baseUrl || '',
    apiKey: s.apiKey || '',
    model: s.model || '',
    numCtx: s.numCtx || 32768,
  })
  const out = { ...s, profiles: [p], activeId: p.id }
  delete out.baseUrl
  delete out.apiKey
  delete out.model
  delete out.numCtx
  return out
}

export function loadSettings() {
  try {
    const raw = localStorage.getItem(KEY)
    const parsed = raw ? { ...DEFAULT_SETTINGS, ...JSON.parse(raw) } : { ...DEFAULT_SETTINGS }
    const s = migrate(parsed)
    if (!s.activeId && s.profiles.length) s.activeId = s.profiles[0].id
    return s
  } catch {
    return { ...DEFAULT_SETTINGS }
  }
}

export function saveSettings(s) {
  localStorage.setItem(KEY, JSON.stringify(s))
}

/** 選択中のプロファイル。無ければ null */
export function activeProfile(s) {
  return s.profiles.find((p) => p.id === s.activeId) ?? s.profiles[0] ?? null
}

/** client.js に渡す形（プロファイル＋共通設定） */
export function connectionOf(s) {
  const p = activeProfile(s)
  if (!p) return null
  return {
    baseUrl: p.baseUrl,
    apiKey: p.apiKey,
    model: p.model,
    numCtx: p.numCtx,
    temperature: s.temperature,
  }
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
