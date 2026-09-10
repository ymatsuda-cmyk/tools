/**
 * AI生成のプロンプト。
 *
 * 既定値はアプリと同じ場所に置いた setting.json から読む。コードに埋め込まず
 * ファイルに出しているのは、文言を直すたびにJSを触らなくて済むようにするため。
 * 画面から編集した内容は localStorage に上書きとして持ち、setting.json より優先する
 * (既定に戻せば上書きは消える)。
 */

const KEY = 'videos:prompts'
const SETTING_URL = new URL('../../setting.json', import.meta.url)

export const PROMPT_IDS = ['summary', 'mindmap', 'fields', 'apply', 'ideas']

let defaults = {}
let overrides = loadOverrides()

function loadOverrides() {
  try {
    return JSON.parse(localStorage.getItem(KEY) || '{}')
  } catch {
    return {}
  }
}

function persist() {
  try {
    localStorage.setItem(KEY, JSON.stringify(overrides))
  } catch {
    // 容量超過。保存できなくてもこのセッションの編集内容は生きている
  }
}

/** setting.json を読む。失敗しても画面は動かす(生成しようとしたときにだけ落ちる) */
export async function initPrompts() {
  try {
    const res = await fetch(SETTING_URL, { cache: 'no-cache' })
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    const json = await res.json()
    defaults = json?.prompts || {}
  } catch (err) {
    console.error('setting.json を読み込めませんでした:', err)
  }
}

export function promptLabel(id) {
  return defaults[id]?.label || id
}

export function defaultPromptOf(id) {
  return defaults[id]?.system ?? ''
}

export function promptOf(id) {
  return typeof overrides[id] === 'string' ? overrides[id] : defaultPromptOf(id)
}

export function isCustomPrompt(id) {
  return typeof overrides[id] === 'string' && overrides[id] !== defaultPromptOf(id)
}

/** 既定と同じ内容なら上書きを持たない。setting.json を直したときに追従させるため */
export function savePrompt(id, text) {
  const value = String(text ?? '')
  if (!value.trim() || value === defaultPromptOf(id)) delete overrides[id]
  else overrides[id] = value
  persist()
}

export function resetPrompt(id) {
  delete overrides[id]
  persist()
}
