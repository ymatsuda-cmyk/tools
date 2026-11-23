// src/marks.js — マーカー（書籍アプリと同じ：選んだ文字に色を付けて、ノートに集める）
//
// マーカーは、文章そのものには手を入れず、n.json の marks[] に位置で持つ：
//   { id, loc, s, e, c, q, tt?, m?, o?, at }
//   loc … 文章の場所。'summary' / 'decision:2' / 'agenda:1:topic' / 'agenda:1:point:3' / 'agenda:1:outcome'
//         / 'topic:0' / 'todo:<ToDo の key>' / 'tr:12'（文字起こしの12行目）
//   s,e … その文章の中の [s, e) 番目の文字    c … 色 1〜3    q … 選んだ文字（位置がずれたときの探し直し・ノート表示用）
//   tt  … 文字起こしの時刻（秒）  m … ノートに書いたメモ  o … 元の文章が変わって、位置が分からなくなった
import { todoKey } from './todos.js'

export const kindOf = (loc) => String(loc).split(':')[0]
const KIND_ORDER = { summary: 0, decision: 1, agenda: 2, topic: 3, todo: 4, tr: 5 }
export const KIND_LABEL = { summary: 'サマリ', decision: '決定事項', agenda: '論点別', topic: '未解決の論点', todo: 'ToDo', tr: '文字起こし' }
const SUB = { topic: 0, point: 1, outcome: 2 }
const SUB_LABEL = { topic: '議題', outcome: '結論' }

export const SUMMARY_KINDS = ['summary', 'decision', 'agenda', 'topic', 'todo']

/** 見る場所（詳細のタブ）：文字起こし→tr、ToDo→todo、ほかは要約 */
export const locTab = (loc) => ({ tr: 'tr', todo: 'todo' }[kindOf(loc)] || 'sum')

export function locLabel(loc) {
  const p = String(loc).split(':')
  const base = KIND_LABEL[p[0]] || p[0]
  if (p[0] === 'agenda' && p[2] && SUB_LABEL[p[2]]) return `${base}（${SUB_LABEL[p[2]]}）`
  return base
}

export function locCompare(a, b) {
  const pa = String(a).split(':'), pb = String(b).split(':')
  const ka = KIND_ORDER[pa[0]] ?? 9, kb = KIND_ORDER[pb[0]] ?? 9
  if (ka !== kb) return ka - kb
  for (let i = 1; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] ?? '', y = pb[i] ?? ''
    if (x === y) continue
    if (x in SUB && y in SUB) return SUB[x] - SUB[y]
    const nx = Number(x), ny = Number(y)
    if (x !== '' && y !== '' && !isNaN(nx) && !isNaN(ny)) return nx - ny
    return x < y ? -1 : 1
  }
  return 0
}

export const newMarkId = () => 'k' + Math.random().toString(36).slice(2, 10)
export const marksOf = (n) => (n && Array.isArray(n.marks) ? n.marks : [])

export function groupMarks(marks) {
  const g = new Map()
  for (const m of marks || []) { if (!g.has(m.loc)) g.set(m.loc, []); g.get(m.loc).push(m) }
  return g
}

/** 文章と合っているマーカーだけ（位置順・重なりなし）。合わないものは表示しない */
export function validFor(text, list) {
  const out = []
  let end = 0
  for (const m of [...(list || [])].sort((a, b) => a.s - b.s)) {
    if (m.o || m.s < end || m.e > text.length || text.slice(m.s, m.e) !== m.q) continue
    out.push(m); end = m.e
  }
  return out
}

const hl = (s, q, esc) => (q ? s.split(q).map(esc).join(`<mark class="hit">${esc(q)}</mark>`) : esc(s))

/** 文章を、マーカーの色付きの HTML にする。q があれば、検索語も強調する */
export function renderMarked(text, list, esc, q = '') {
  let html = '', pos = 0
  for (const m of list || []) {
    if (m.s > pos) html += hl(text.slice(pos, m.s), q, esc)
    html += `<span class="marker c${m.c}" data-mid="${esc(m.id)}">${hl(text.slice(m.s, m.e), q, esc)}</span>`
    pos = m.e
  }
  return html + hl(text.slice(pos), q, esc)
}

/**
 * [s, e) に色 c を付ける。重なる既存のマーカーは切り詰める（c が null なら、その範囲のマーカーを消す）。
 * 選んだ範囲の前後の空白は外す。
 */
export function applyRange(marks, loc, s, e, c, text, at = Date.now()) {
  while (s < e && /\s/.test(text[s])) s++
  while (e > s && /\s/.test(text[e - 1])) e--
  if (e <= s) return marks
  const out = []
  for (const m of marks || []) {
    if (m.loc !== loc || m.o || m.e <= s || m.s >= e) { out.push(m); continue }
    if (m.s < s) out.push({ ...m, e: s, q: text.slice(m.s, s) })
    if (m.e > e) out.push({ ...m, id: m.s < s ? newMarkId() : m.id, s: e, q: text.slice(e, m.e), m: m.s < s ? undefined : m.m })
  }
  if (c) out.push({ id: newMarkId(), loc, s, e, c, q: text.slice(s, e), at })
  return out
}

/** ある場所の、1つのマーカーの色を変える */
export const recolor = (marks, id, c) => (marks || []).map((m) => (m.id === id ? { ...m, c } : m))

function nearest(text, q, from) {
  let best = -1, bd = Infinity, i = text.indexOf(q)
  while (i >= 0) { const d = Math.abs(i - from); if (d < bd) { bd = d; best = i } i = text.indexOf(q, i + 1) }
  return best
}

/**
 * 本文を直したあと、マーカーの位置を探し直す。
 * 同じ場所に同じ文字があればそこ（複数なら元の位置に近いもの）、無ければ同じ種類の別の場所、
 * それでも無ければ「元の文章が変わった」印（o）を付けて、ノートには残す。
 * kinds に無い種類のマーカーには手を付けない。
 */
export function reconcile(marks, texts, kinds) {
  const moved = (marks || []).map((m) => {
    const k = kindOf(m.loc)
    if (kinds && !kinds.includes(k)) return m
    const t = texts[m.loc]
    if (t != null && !m.o && t.slice(m.s, m.e) === m.q) return m
    if (t != null) { const i = nearest(t, m.q, m.s); if (i >= 0) { const { o, ...r } = m; return { ...r, s: i, e: i + m.q.length } } }
    for (const [loc, tx] of Object.entries(texts)) {
      if (kindOf(loc) !== k || loc === m.loc) continue
      const i = tx.indexOf(m.q)
      if (i >= 0) { const { o, ...r } = m; return { ...r, loc, s: i, e: i + m.q.length } }
    }
    return { ...m, o: 1 }
  })
  // 探し直した結果、同じ場所で重なったものは、あとのものを「位置不明」にする
  const seen = new Map()
  return moved.map((m) => {
    if (m.o) return m
    const list = seen.get(m.loc) || []
    if (list.some((x) => m.s < x.e && m.e > x.s)) return { ...m, o: 1 }
    list.push(m); seen.set(m.loc, list)
    return m
  })
}

/** 要約の各文章を、場所（loc）ごとに取り出す */
export function summaryTexts(s) {
  const o = { summary: s.summary || s.cardSummary || '' }
  ;(s.decisions || []).forEach((d, i) => { o['decision:' + i] = d.text })
  ;(s.agenda || []).forEach((a, i) => {
    o[`agenda:${i}:topic`] = a.topic
    if (a.outcome) o[`agenda:${i}:outcome`] = a.outcome
    ;(a.points || []).forEach((p, j) => { o[`agenda:${i}:point:${j}`] = p.text })
  })
  ;(s.topics || []).forEach((t, i) => { o['topic:' + i] = t })
  ;(s.todos || []).forEach((t, i) => { o['todo:' + todoKey(t, i)] = t.text })
  return o
}

export const trTexts = (segments) => Object.fromEntries((segments || []).map((x, i) => ['tr:' + i, x.s]))

/** 文字起こしのマーカーに、その行の時刻を入れ直す */
export function withTimes(marks, segments) {
  return (marks || []).map((m) => {
    if (kindOf(m.loc) !== 'tr' || m.o) return m
    const seg = segments[Number(m.loc.split(':')[1])]
    return seg && seg.t != null ? { ...m, tt: seg.t } : m
  })
}
