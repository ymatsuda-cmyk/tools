const DETAIL_PREFIX = 'minutes:detail:'

/**
 * @returns {{cardSummary, detail, model, generatedAt, cachedAt}|null}
 */
export function getDetailCache(pageId) {
  try {
    const raw = localStorage.getItem(DETAIL_PREFIX + pageId)
    if (!raw) return null
    const parsed = JSON.parse(raw)
    // 旧形式(ToDoが文字列配列)のキャッシュを新形式に揃える
    const todos = parsed.detail?.todos
    if (Array.isArray(todos) && typeof todos[0] === 'string') {
      parsed.detail.todos = todos.map((text) => ({ text, done: false }))
    }
    return parsed
  } catch {
    return null
  }
}

export function setDetailCache(pageId, data) {
  const record = { ...data, cachedAt: new Date().toISOString() }
  localStorage.setItem(DETAIL_PREFIX + pageId, JSON.stringify(record))
  return record
}

/**
 * キャッシュが Notion 側の更新より新しければ再利用可能。
 * updatedAt (Notionページの last_edited_time) を基準に比較する。
 * generatedAt(要約生成日時)は「要約を再生成」した時にしか変わらないため、
 * 決定事項/ToDo/論点/議事/サマリを個別編集した場合の変更を検知できない。
 * updatedAt はどのプロパティを変更しても更新されるため、こちらを基準にする。
 */
export function isCacheFresh(cache, remoteUpdatedAt) {
  if (!cache || !remoteUpdatedAt) return false
  if (!cache.updatedAt) return false // 旧キャッシュにupdatedAtが無ければ安全側に倒して再取得する
  return new Date(cache.updatedAt).getTime() >= new Date(remoteUpdatedAt).getTime()
}
