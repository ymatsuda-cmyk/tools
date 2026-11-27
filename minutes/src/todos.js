// src/todos.js — ToDo の識別子・完了・経過、再要約したときの突き合わせ
//
// ToDo は s.json の todos[] に入っていて、完了と経過は n.json に ToDo の識別子（key）ごとに入っている：
//   n.done = { <key>: true }       n.logs = { <key>: [{ d: 時刻(ms), x: 本文 }] }
// key は todo.id。無いもの（今の minutes から取り込んだものなど）は、並び順から 'i0', 'i1' … とみなす。

export const todoKey = (t, i) => (t && t.id) || 'i' + i
export const newTodoKey = () => 'n' + Math.random().toString(36).slice(2, 8)

export function doneOf(n, key) {
  const d = n && n.done
  if (!d) return false
  if (key in d) return !!d[key]
  const m = /^i(\d+)$/.exec(key)
  return m ? !!d[m[1]] : false   // 古い形式（並び順の番号）の完了も読む
}

export const logsOf = (n, key) => (n && n.logs && n.logs[key]) || []

/** 比較用に、空白・記号・定型の言い回しを取り除く */
function normalize(text) {
  return String(text || '')
    .replace(/[\s　、。,.・()（）「」【】\-—/:：]/g, '')
    .replace(/(を確認する|を確認|を実施する|を実施|を検討する|を依頼する|を行う|について|に関して|に対して|します|する|した|して)/g, '')
    .replace(/[のをがにはでとへもや]/g, '')
    .toLowerCase()
}

/** 文字の2-gramによる Dice 係数。日本語でも分かち書き無しで近さを測れる */
export function similarity(a, b) {
  const x = normalize(a), y = normalize(b)
  if (!x || !y) return 0
  if (x === y) return 1
  const grams = (s) => { const m = new Map(); for (let i = 0; i < s.length - 1; i++) { const g = s.slice(i, i + 2); m.set(g, (m.get(g) || 0) + 1) } return m }
  const gx = grams(x), gy = grams(y)
  let overlap = 0
  gx.forEach((n, g) => { overlap += Math.min(n, gy.get(g) || 0) })
  const nx = Math.max(x.length - 1, 0), ny = Math.max(y.length - 1, 0)
  if (!nx || !ny) return 0
  const dice = (2 * overlap) / (nx + ny)
  const ratio = overlap / Math.min(nx, ny)
  return dice >= 0.25 ? Math.max(dice, ratio * 0.9) : dice
}

const MATCH = 0.5

/**
 * 再要約で作り直した ToDo を、前の ToDo と文言の近さで突き合わせて、識別子（＝完了・経過）を引き継ぐ。
 * どれにも対応しなかった前の ToDo のうち、完了か経過を持つものは、捨てずに末尾へ残す（orphan）。
 */
export function mergeTodos(prevTodos, nextTodos, n) {
  const prev = (prevTodos || []).map((t, i) => ({ ...t, id: todoKey(t, i) }))
  const next = nextTodos || []
  const pairs = []
  next.forEach((t, ni) => prev.forEach((p, pi) => { const sc = similarity(t.text, p.text); if (sc >= MATCH) pairs.push({ ni, pi, sc }) }))
  pairs.sort((a, b) => b.sc - a.sc)
  const toPrev = new Map(), used = new Set()
  pairs.forEach(({ ni, pi }) => { if (!toPrev.has(ni) && !used.has(pi)) { toPrev.set(ni, pi); used.add(pi) } })
  const out = next.map((t, ni) => ({ ...t, id: toPrev.has(ni) ? prev[toPrev.get(ni)].id : (t.id || newTodoKey()) }))
  prev.forEach((p, pi) => {
    if (used.has(pi)) return
    if (doneOf(n, p.id) || logsOf(n, p.id).length) out.push({ ...p, orphan: true })
  })
  return out
}
