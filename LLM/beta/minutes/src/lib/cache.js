const DETAIL_PREFIX = 'minutes:detail:'

/**
 * @returns {{cardSummary, detail, model, generatedAt, cachedAt}|null}
 */
export function getDetailCache(pageId) {
  try {
    const raw = localStorage.getItem(DETAIL_PREFIX + pageId)
    return raw ? JSON.parse(raw) : null
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
