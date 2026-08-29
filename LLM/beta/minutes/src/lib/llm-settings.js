// gemma-chat アプリと同じ localStorage キーを共有する。
// プロファイル(接続先LLMの一覧)を一箇所で管理し、どちらのアプリからも
// 追加・切り替えができるようにするための意図的な共有。
// 中身は gemma-chat 側の src/lib/settings.js と同一に保つこと。

const KEY = 'gemma-chat.settings'

function makeId() {
  return Math.random().toString(36).slice(2, 8)
}

export function newProfile(over = {}) {
  return {
    id: makeId(),
    label: '新しい接続',
    baseUrl: '',
    apiKey: '',
    model: '',
    numCtx: 32768,
    gasUrl: '',
    controlToken: '',
    think: null,
    ...over,
  }
}

export const DEFAULT_SETTINGS = {
  profiles: [],
  activeId: null,
  temperature: 0.7,
  systemPrompt: '必ず日本語で回答してください。',
}

function migrate(s) {
  if (Array.isArray(s.profiles) && s.profiles.length) {
    if (s.gasUrl || s.controlToken) {
      for (const p of s.profiles) {
        if (!p.gasUrl) p.gasUrl = s.gasUrl || ''
        if (!p.controlToken) p.controlToken = s.controlToken || ''
      }
    }
    const out = { ...s }
    delete out.gasUrl
    delete out.controlToken
    return out
  }
  if (!s.baseUrl && !s.apiKey) return { ...s, profiles: [], activeId: null }
  const p = newProfile({
    id: 'a',
    label: s.model || 'メイン',
    baseUrl: s.baseUrl || '',
    apiKey: s.apiKey || '',
    model: s.model || '',
    numCtx: s.numCtx || 32768,
    gasUrl: s.gasUrl || '',
    controlToken: s.controlToken || '',
  })
  const out = { ...s, profiles: [p], activeId: p.id }
  delete out.baseUrl
  delete out.apiKey
  delete out.model
  delete out.numCtx
  delete out.gasUrl
  delete out.controlToken
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

export function activeProfile(s) {
  return s.profiles.find((p) => p.id === s.activeId) ?? s.profiles[0] ?? null
}

export function connectionOf(s) {
  const p = activeProfile(s)
  if (!p) return null
  return {
    baseUrl: p.baseUrl,
    apiKey: p.apiKey,
    model: p.model,
    numCtx: p.numCtx,
    temperature: s.temperature,
    think: p.think ?? null,
  }
}
