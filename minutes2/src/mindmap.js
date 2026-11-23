// src/mindmap.js — マインドマップ（markmap 用の Markdown を、枝を開け閉めできる木にして見せる）
//
// 今の minutes と同じ形式：  # 中心のテーマ  /  ## 大項目  /  - 子（2字下げで、さらに下の階層）
// 保存済みのもの（s.mindmap）があればそれを、無ければ要約の構成からそのまま組み立てる。
import { todoKey } from './todos.js'

const TAGS = /<\/?m[123]>/g
const clean = (t) => String(t ?? '').replace(TAGS, '').replace(/\s+/g, ' ').trim()

/** Markdown → 木 { text, children }。見出し（#）と箇条書き（- * + 1.）の入れ子を読む */
export function parseOutline(md) {
  const root = { text: '', children: [], level: 0 }
  const stack = [root]
  let base = 0   // 直前の見出しの深さ（箇条書きは、その下に入る）
  for (const raw of String(md || '').replace(/\r\n?/g, '\n').split('\n')) {
    if (!raw.trim()) continue
    let level, text
    let m = raw.match(/^(#{1,6})\s+(.*)$/)
    if (m) { level = m[1].length; base = level; text = m[2] }
    else if ((m = raw.match(/^(\s*)(?:[-*+]|\d+\.)\s+(.*)$/))) { level = base + 1 + Math.floor(m[1].replace(/\t/g, '  ').length / 2); text = m[2] }
    else continue
    text = clean(text)
    if (!text) continue
    while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop()
    const node = { text, children: [], level }
    stack[stack.length - 1].children.push(node)
    stack.push(node)
  }
  // 先頭の1件が中心テーマ（# が1つ）なら、それを根にする
  if (root.children.length === 1) return root.children[0]
  return { text: '', children: root.children, level: 0 }
}

/** 要約の構成をそのままマインドマップにする（AI は使わない） */
export function outlineFromSummary(title, s) {
  const node = (text, children = []) => ({ text: clean(text), children: children.filter((c) => c.text) })
  const branches = []
  const sentences = (clean(s.summary || s.cardSummary).match(/[^。]+。?/g) || []).map((x) => x.trim()).filter(Boolean)   // 後読みは古い Safari で使えないので使わない
  if (sentences.length) branches.push(node('サマリ', sentences.map((x) => node(x))))
  ;(s.agenda || []).forEach((a) => {
    const kids = (a.points || []).map((p) => node(p.text))
    if (a.outcome) kids.push(node('結論：' + a.outcome))
    branches.push(node(a.topic || '議題', kids))
  })
  if ((s.decisions || []).length) branches.push(node('決定事項', s.decisions.map((d) => node(d.text))))
  if ((s.todos || []).length) branches.push(node('ToDo', s.todos.map((t) => node(t.text))))
  if ((s.topics || []).length) branches.push(node('未解決の論点', s.topics.map((t) => node(t))))
  return { text: clean(title) || '議事録', children: branches.filter((b) => b.text), level: 1 }
}

export function mindmapTree(title, s) {
  if (s && s.mindmap) {
    const t = parseOutline(s.mindmap)
    if (t.children.length) return { ...t, text: t.text || clean(title) || '議事録' }
  }
  return outlineFromSummary(title, s || {})
}

const COLORS = 5
/** 木 → HTML。2階層目までは開き、それより下は閉じておく（押すと開く）。枝ごとに色分け */
export function renderTree(tree, esc, openLevel = 2) {
  const branch = (n, depth, color) => {
    const c = `c${color}`
    if (!n.children.length) return `<div class="mm-b leaf ${c}"><div class="mm-n">${esc(n.text)}</div></div>`
    return `<details class="mm-b ${c}" ${depth < openLevel ? 'open' : ''}><summary class="mm-n">${esc(n.text)}<span class="mm-c">${n.children.length}</span></summary>
      <div class="mm-k">${n.children.map((k) => branch(k, depth + 1, color)).join('')}</div></details>`
  }
  const kids = tree.children.map((k, i) => branch(k, 1, i % COLORS)).join('')
  return `<div class="mm"><div class="mm-root">${esc(tree.text)}</div><div class="mm-k root">${kids}</div></div>`
}
