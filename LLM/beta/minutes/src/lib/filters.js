import { getDetailCache } from './cache.js'

/** index.json 全体から既知のタグを集める(タグ選択モーダルの候補に使う) */
export function allKnownTags(items) {
  const set = new Set()
  items.forEach((i) => (i.tags || []).forEach((t) => set.add(t)))
  return [...set].sort()
}

/** 選択した状態のいずれかに一致するものだけ(未選択なら全件通す) */
export function filterByStatus(items, selectedStatuses) {
  if (!selectedStatuses.size) return items
  return items.filter((i) => selectedStatuses.has(i.status))
}

/** 状態が「削除」のものを除外する */
export function excludeDeleted(items) {
  return items.filter((i) => i.status !== '削除')
}

/** items のうち year-month(0始まりではなく "YYYY-MM") が一致するものだけ */
export function filterByMonth(items, monthKey) {
  return items.filter((i) => i.date.slice(0, 7) === monthKey)
}

/** タイトルとキャッシュ済み要約(cardSummary)だけを対象にした検索。本文は含まない */
export function filterBySearch(items, query) {
  const q = query.trim().toLowerCase()
  if (!q) return items
  return items.filter((i) => {
    if (i.title.toLowerCase().includes(q)) return true
    const cache = getDetailCache(i.key)
    return Boolean(cache?.cardSummary?.toLowerCase().includes(q))
  })
}

/** 選択タグをすべて含む(AND)アイテムのみ */
export function filterByTags(items, selectedTags) {
  if (!selectedTags.size) return items
  return items.filter((i) => {
    const tags = i.tags || []
    return [...selectedTags].every((t) => tags.includes(t))
  })
}

/**
 * ツールバー用のタグ候補を、選択可否付きで組み立てる。
 * baseItems(タグ絞り込み前、月・検索は適用済み)から
 * 「そのタグを追加選択したら該当件数」を数え、0件なら選択不可にする。
 */
export function buildTagOptions(baseItems, selectedTags) {
  const allTags = new Set()
  baseItems.forEach((i) => (i.tags || []).forEach((t) => allTags.add(t)))

  return [...allTags].sort().map((tag) => {
    const selected = selectedTags.has(tag)
    const candidate = new Set(selectedTags)
    candidate.add(tag)
    const count = filterByTags(baseItems, candidate).length
    return { tag, selected, disabled: !selected && count === 0 }
  })
}
