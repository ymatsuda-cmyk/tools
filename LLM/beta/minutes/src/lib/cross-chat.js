const DATA_KEY = 'minutes:crossChatData'
const SPACES_KEY = 'minutes:crossChatSpaces'

// Gemma(コンテキスト長 32K トークン前提)を想定した警告しきい値。
// 日本語は概ね1文字≈1.3〜1.5トークンなので、プロンプトの余白を見て
// 本文20,000字程度を上限の目安とする。あくまで警告であり生成は止めない。
export const GEMMA_WARN_CHARS = 20000

function makeId() {
  return Math.random().toString(36).slice(2, 8)
}

/** @returns {{createdAt: string, count: number, chars: number, items: object[]} | null} */
export function loadCrossChatData() {
  try {
    const raw = localStorage.getItem(DATA_KEY)
    return raw ? JSON.parse(raw) : null
  } catch {
    return null
  }
}

export function saveCrossChatData(data) {
  localStorage.setItem(DATA_KEY, JSON.stringify(data))
}

export function clearCrossChatData() {
  localStorage.removeItem(DATA_KEY)
}

/** 1件分のchat用データの文字数を見積る(agenda/decisions/todosのみ) */
export function estimateItemChars(entry) {
  return JSON.stringify({
    title: entry.title,
    agenda: entry.agenda,
    decisions: entry.decisions,
    todos: entry.todos,
  }).length
}

// --- チャットスペース ---

export function loadSpaces() {
  try {
    const raw = localStorage.getItem(SPACES_KEY)
    return raw ? JSON.parse(raw) : []
  } catch {
    return []
  }
}

export function saveSpaces(spaces) {
  localStorage.setItem(SPACES_KEY, JSON.stringify(spaces))
}

export function newSpace(name) {
  return { id: makeId(), name: name || '新しいスペース', messages: [] }
}
