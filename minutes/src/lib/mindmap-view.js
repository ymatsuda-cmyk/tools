// 要約(サマリ・議事・決定事項・ToDo・論点)をマインドマップとして描画する。
// 描画本体は api/mindmap/mindmap.api.js (window.MindMap) を利用する。
import { plainTextOf } from './markers.js'

const MAX_LEN = 10 // MindMap側のノード文字数上限。超過分は自前で「…」に丸める
const MAX_CHILDREN = 8

const treeByKey = {} // item.key -> 編集後のツリー(タブ切替やノード編集の結果を保つ)
let host = null // mount()は1度だけ。この要素をタブのDOMへ移動して使い回す
let currentKey = null

function short(text) {
  const s = plainTextOf(String(text ?? '')).replace(/\s+/g, ' ').trim().replace(/[。．.]+$/, '')
  if (s.length <= MAX_LEN) return s
  return s.slice(0, MAX_LEN - 1) + '…'
}

function node(text, color, children) {
  const n = { text: short(text) || '(空)' }
  if (color) n.color = color
  if (children && children.length) n.children = children.slice(0, MAX_CHILDREN)
  return n
}

export function buildTreeFromSummary(item, summary) {
  const d = summary?.detail || {}
  const children = []

  const sentences = plainTextOf(summary?.cardSummary || '')
    .split(/[。\n]/)
    .map((x) => x.trim())
    .filter(Boolean)
  if (sentences.length) {
    children.push(node('サマリ', 'purple', sentences.map((x) => node(x))))
  }

  ;(d.agenda || []).slice(0, MAX_CHILDREN).forEach((a) => {
    const points = (a.points || []).map((p) => node(p))
    if (a.outcome) points.push(node(a.outcome, 'green'))
    children.push(node(a.topic || '議題', null, points))
  })

  if (d.decisions?.length) children.push(node('決定事項', 'orange', d.decisions.map((x) => node(x))))
  if (d.todos?.length) children.push(node('ToDo', 'yellow', d.todos.map((t) => node(t?.text ?? t))))
  if (d.topics?.length) children.push(node('論点', 'pink', d.topics.map((x) => node(x))))

  return { text: short(item.title) || '議事録', children }
}

/** マインドマップタブのDOMに描画領域を差し込み、ツールバーを配線する */
export function setupMindmapTab(target, item, state) {
  const slot = target.querySelector('#mindmap-slot')
  if (!slot) return

  const MM = window.MindMap
  if (!MM) {
    slot.innerHTML = '<p class="empty-section">マインドマップを読み込めませんでした</p>'
    return
  }

  if (!host) {
    host = document.createElement('div')
    host.className = 'mm-viewport'
    MM.mount(host)
    MM.on('change', (tree) => {
      if (currentKey && tree) treeByKey[currentKey] = tree
    })
    MM.on('zoom', (z) => {
      document.querySelectorAll('.mindmap-zoom-level').forEach((el) => {
        el.textContent = `${Math.round(z * 100)}%`
      })
    })
  }
  slot.appendChild(host)
  currentKey = item.key

  const draw = (tree) => {
    MM.render(tree)
    requestAnimationFrame(() => MM.fit())
  }
  draw(treeByKey[item.key] || buildTreeFromSummary(item, state.summary))

  target.querySelector('.btn-mm-rebuild')?.addEventListener('click', () => {
    delete treeByKey[item.key]
    draw(buildTreeFromSummary(item, state.summary))
  })
  target.querySelectorAll('[data-mm]').forEach((btn) => {
    btn.addEventListener('click', () => {
      if (btn.dataset.mm === 'in') MM.zoomIn()
      else if (btn.dataset.mm === 'out') MM.zoomOut()
      else MM.fit()
    })
  })
}
