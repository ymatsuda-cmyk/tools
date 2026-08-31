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
 * generatedAt (要約が生成された日時) を基準に比較する。
 */
export function isCacheFresh(cache, remoteGeneratedAt) {
  if (!cache || !remoteGeneratedAt) return false
  return new Date(cache.generatedAt).getTime() >= new Date(remoteGeneratedAt).getTime()
}
