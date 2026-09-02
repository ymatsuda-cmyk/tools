const DATA_KEY = 'minutes:crossChatData' // 旧形式(全スペース共通)。移行のためだけに残す
const SPACES_KEY = 'minutes:crossChatSpaces'

// Gemma(コンテキスト長 32K トークン前提)を想定した警告しきい値。
// 日本語は概ね1文字≈1.3〜1.5トークンなので、プロンプトの余白を見て
// 本文20,000字程度を上限の目安とする。あくまで警告であり生成は止めない。
export const GEMMA_WARN_CHARS = 20000

// 対象データの選択上限。件数を増やすとコンテキストが膨らみ、
// Gemma前提の想定を超えやすくなるため一律この件数に固定する。
export const MAX_CROSS_CHAT_ITEMS = 20

function makeId() {
  return Math.random().toString(36).slice(2, 8)
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
// 対象データ(選んだ議事録一覧)はスペースごとに個別に持つ。
// 新規スペースは空のデータから始まり、作成時に対象を選ぶ。

export function loadSpaces() {
  try {
    const raw = localStorage.getItem(SPACES_KEY)
    const spaces = raw ? JSON.parse(raw) : []
    return migrateSpaces(spaces)
  } catch {
    return []
  }
}

export function saveSpaces(spaces) {
  localStorage.setItem(SPACES_KEY, JSON.stringify(spaces))
}

export function newSpace(name) {
  return { id: makeId(), name: name || '新しいスペース', messages: [], data: null }
}

/**
 * 旧形式(全スペース共通の1つの対象データ)から、
 * スペース単位に対象データを持たせる形式への移行。
 * 既存スペースには、旧共通データをそのままコピーする(選び直し不要)。
 * 移行後は旧データを削除し、二重管理を防ぐ。
 */
function migrateSpaces(spaces) {
  let changed = false
  const legacy = (() => {
    try {
      const raw = localStorage.getItem(DATA_KEY)
      return raw ? JSON.parse(raw) : null
    } catch {
      return null
    }
  })()

  const migrated = spaces.map((s) => {
    if (s.data !== undefined) return s // 既に新形式
    changed = true
    return { ...s, data: legacy }
  })

  if (changed) {
    saveSpaces(migrated)
    if (legacy) localStorage.removeItem(DATA_KEY)
  }
  return migrated
}
