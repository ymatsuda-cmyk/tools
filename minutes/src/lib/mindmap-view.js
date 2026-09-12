// 要約(サマリ・議事・決定事項・ToDo・論点)からマインドマップのツリーを作り、描画する。
// 描画本体は api/mindmap/mindmap.api.js (window.MindMap)。
// ツリーのJSONはNotionの「マインドマップ」カラムに保存する(保存処理は main.js 側)。
import { plainTextOf } from './markers.js'
import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'

const MAX_LEN = 10 // ノードの文字数上限。超える分は「…」で丸める
const MAX_CHILDREN = 8
const MAX_DEPTH = 2 // 中心テーマを0とした階層の深さ(合計3階層)
const COLORS = ['purple', 'green', 'orange', 'yellow', 'pink']

let host = null // mount()は1度だけ。この要素をタブのDOMへ移動して使い回す
let changeHandler = null
let suppressChange = false // 初期描画のemitを「編集」と誤認しないための抑止フラグ

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

const INSTRUCTION = `次の原文を日本語のマインドマップに構造化してください。
出力はJSONのみ。前置き・コードフェンス・説明は一切書かないこと。
形式: {"text":"中心テーマ","children":[{"text":"見出し","children":[{"text":"要点"}]}]}
制約:
- すべてのtextは日本語10文字以内。超えそうなら削って体言止めにする
- 第1階層は3〜7個、第2階層は各0〜4個、3階層まで
- 原文にない情報を足さない`

/** 要約をAIへ渡すプレーンテキストにする */
function summaryText(item, summary) {
  const d = summary?.detail || {}
  const lines = [`会議名: ${item.title}`]
  if (summary?.cardSummary) lines.push('サマリ: ' + plainTextOf(summary.cardSummary))
  ;(d.agenda || []).forEach((a, i) => {
    lines.push(`議題${i + 1}: ${plainTextOf(a.topic || '')}`)
    ;(a.points || []).forEach((p) => lines.push('- ' + plainTextOf(p)))
    if (a.outcome) lines.push('結論: ' + plainTextOf(a.outcome))
  })
  if (d.decisions?.length) lines.push('決定事項: ' + d.decisions.map((x) => plainTextOf(x)).join(' / '))
  if (d.todos?.length) lines.push('ToDo: ' + d.todos.map((t) => plainTextOf(t?.text ?? t)).join(' / '))
  if (d.topics?.length) lines.push('論点: ' + d.topics.map((x) => plainTextOf(x)).join(' / '))
  return lines.join('\n')
}

function extractJson(text) {
  const trimmed = text.trim()
  const start = trimmed.indexOf('{')
  const end = trimmed.lastIndexOf('}')
  if (start === -1 || end === -1) throw new Error('LLM応答からJSONを抽出できませんでした')
  return JSON.parse(trimmed.slice(start, end + 1))
}

/** AIの出力を文字数・階層・色の制約に収める */
function sanitizeNode(raw, depth) {
  const children = depth < MAX_DEPTH && Array.isArray(raw?.children)
    ? raw.children.map((c) => sanitizeNode(c, depth + 1))
    : []
  return node(raw?.text, COLORS.includes(raw?.color) ? raw.color : null, children)
}

/** 要約をもとにAIでマインドマップのツリーを生成する */
export async function generateTreeWithAI(item, summary) {
  const connection = connectionOf(loadSettings())
  if (!connection) throw new Error('LLM接続が未設定です。設定から接続先とモデルを追加してください。')

  const messages = [
    { role: 'user', content: INSTRUCTION + '\n\n原文:\n' + summaryText(item, summary).slice(0, 30000) },
  ]

  let full = ''
  for await (const chunk of streamChat(connection, messages)) {
    if (chunk.delta) full += chunk.delta
  }

  const tree = sanitizeNode(extractJson(full), 0)
  if (!tree.children?.length) throw new Error('ノードが生成されませんでした')
  return tree
}

/** 「マインドマップ」カラムの文字列をツリーに戻す。空・壊れていればnull */
export function parseTree(json) {
  if (!json) return null
  try {
    const tree = JSON.parse(json)
    return tree && typeof tree.text === 'string' ? tree : null
  } catch {
    return null
  }
}

/**
 * マインドマップタブのDOMに描画領域を差し込み、ツールバーを配線する。
 * @param {(tree: object) => void} onChange ノード編集のたびに最新ツリーを受け取る
 */
export function setupMindmapTab(target, tree, onChange) {
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
    MM.on('change', (edited) => {
      if (!suppressChange && edited) changeHandler?.(edited)
    })
    MM.on('zoom', (z) => {
      document.querySelectorAll('.mindmap-zoom-level').forEach((el) => {
        el.textContent = `${Math.round(z * 100)}%`
      })
    })
  }
  changeHandler = onChange
  slot.appendChild(host)

  suppressChange = true
  MM.render(tree)
  suppressChange = false
  requestAnimationFrame(() => MM.fit())

  target.querySelectorAll('[data-mm]').forEach((btn) => {
    btn.addEventListener('click', () => {
      if (btn.dataset.mm === 'in') MM.zoomIn()
      else if (btn.dataset.mm === 'out') MM.zoomOut()
      else MM.fit()
    })
  })
}
