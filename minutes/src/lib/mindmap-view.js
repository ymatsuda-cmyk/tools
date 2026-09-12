/**
 * マインドマップ(動画ナレッジ clipstock と同じ仕組み)。
 *
 * Notionの「マインドマップ」カラムには markmap 用のMarkdownを保存する。
 * HTMLやツリーJSONではなくMarkdownにしているのは、Notion上でもそのまま読めて、
 * 描画ライブラリを差し替えても中身が生き残るため。
 */
import { plainTextOf } from './markers.js'
import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'

// 中心テーマ(深さ0)と大項目(##、深さ1)まで開き、第3階層以降は畳んだ状態で描く
const EXPAND_LEVEL = 2

/**
 * ファイル名を直接指定しないこと。
 * ブラウザ向けの実体は markmap-view が dist/browser/index.js なのに対し、
 * markmap-lib は dist/browser/index.iife.js と名前が違う。
 * パッケージ名だけを指定すれば、CDNが package.json の jsdelivr フィールドを見て
 * 正しいファイルを返す。読み込み順も lib(Transformer) -> view(Markmap) から変えない。
 */
const CDN = [
  { url: 'https://cdn.jsdelivr.net/npm/d3@7', check: () => Boolean(window.d3), name: 'd3' },
  {
    url: 'https://cdn.jsdelivr.net/npm/markmap-lib@0.18',
    check: () => Boolean(window.markmap?.Transformer),
    name: 'markmap-lib',
  },
  {
    url: 'https://cdn.jsdelivr.net/npm/markmap-view@0.18',
    check: () => Boolean(window.markmap?.Markmap),
    name: 'markmap-view',
  },
]

let loading = null

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve()
    const el = document.createElement('script')
    el.src = src
    el.onload = () => resolve()
    el.onerror = () => reject(new Error(`読み込みに失敗しました: ${src}`))
    document.head.appendChild(el)
  })
}

/** markmap一式を読み込む。どれが欠けたか分かるよう1つずつ検証する */
function loadMarkmap() {
  if (!loading) {
    loading = (async () => {
      for (const dep of CDN) {
        await loadScript(dep.url)
        if (!dep.check()) throw new Error(`${dep.name} を読み込めませんでした (${dep.url})`)
      }
      return window.markmap
    })().catch((err) => {
      loading = null // 次回リトライできるようにする
      throw err
    })
  }
  return loading
}

/**
 * container の中にマインドマップを描画する。
 * @param {{onChange?: (markdown: string) => void, autoFocus?: boolean}} options
 */
export async function renderMindmap(container, markdown, options = {}) {
  const raw = String(markdown ?? '').trim()
  container.innerHTML = ''
  if (!raw) return

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
  svg.classList.add('mindmap-svg')
  container.appendChild(svg)

  try {
    const { Markmap, Transformer } = await loadMarkmap()
    const transformer = new Transformer()
    const toRoot = (md) => transformer.transform(md).root
    const state = { markdown: raw, root: toRoot(raw) }
    const mm = Markmap.create(svg, { duration: 200, spacingVertical: 6, paddingX: 12, initialExpandLevel: EXPAND_LEVEL }, state.root)
    bindCursor(container, svg, mm, state, options, toRoot)
  } catch (err) {
    // 描画できなくても内容は読めるようにしておく
    container.innerHTML = `
      <p class="error-text">${err.message || err}</p>
      <pre class="mindmap-fallback">${raw.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</pre>
    `
  }
}

// ---- カーソル操作とその場編集 ----
//
// 位置は「非空行の何番目か」で持つ。markmapの木を前順でたどった順番と一致するので、
// 描き直しても同じ枝に戻れる。
// (動画ナレッジ clipstock の src/lib/mindmap.js と同じ作り。片方を直したら両方に反映すること)

const NEW_LABEL = '新しい項目'

/** ノードになる行だけを並び順で返す */
function nodeLineIndexes(markdown) {
  const lines = String(markdown ?? '').split('\n')
  const indexes = []
  lines.forEach((line, i) => {
    if (line.trim()) indexes.push(i)
  })
  return { lines, indexes }
}

/** 行頭の "## " や "  - " と、そのあとのラベルを分ける */
function splitPrefix(line) {
  const m = String(line).match(/^(\s*(?:#{1,6}\s+|[-*+]\s+|\d+\.\s+)?)([\s\S]*)$/)
  return { prefix: m[1], label: m[2] }
}

/** その行がマップの何段目になるか。見出しの数とリストのインデントから決まる */
function lineDepth(line) {
  const heading = line.match(/^(#{1,6})\s/)
  if (heading) return heading[1].length - 1
  const item = line.match(/^(\s*)[-*+]\s/)
  if (item) return 2 + Math.floor(item[1].length / 2)
  return 99
}

/** その枝の1段下に足すときの行頭 */
function childPrefixOf(line) {
  const heading = line.match(/^(#{1,6})\s/)
  if (heading) return heading[1].length < 2 ? '#'.repeat(heading[1].length + 1) + ' ' : '- '
  const item = line.match(/^(\s*)[-*+]\s/)
  return item ? ' '.repeat(item[1].length + 2) + '- ' : '- '
}

/** pos番目のノードの、ぶら下がりを含めた最後の位置 */
function subtreeEnd(lines, indexes, pos) {
  const depth = lineDepth(lines[indexes[pos]])
  let last = pos
  for (let i = pos + 1; i < indexes.length; i++) {
    if (lineDepth(lines[indexes[i]]) <= depth) break
    last = i
  }
  return last
}

/** 内容のあるノードを前順で。Markdownの非空行と同じ並びになる */
function contentNodes(root) {
  const out = []
  ;(function walk(node) {
    if (String(node?.content ?? '').trim()) out.push(node)
    ;(node?.children || []).forEach(walk)
  })(root)
  return out
}

function bindCursor(container, svg, mm, state, options, toRoot) {
  const editable = typeof options.onChange === 'function'
  let all = contentNodes(state.root)
  if (!all.length) return

  let { lines, indexes } = nodeLineIndexes(state.markdown)
  let current = all[0]
  let editing = false

  container.tabIndex = 0
  const gOf = (node) =>
    [...svg.querySelectorAll('g.markmap-node')].find((g) => window.d3?.select(g).datum() === node)
  const posOf = (node) => all.indexOf(node)

  function paint() {
    svg.querySelectorAll('g.mm-current').forEach((g) => g.classList.remove('mm-current'))
    gOf(current)?.classList.add('mm-current')
  }

  /**
   * Markdownを差し替える。描き直しではなく markmap にデータだけ渡すので、
   * 表示位置と拡大率はそのままで、増えた枝だけが現れる。
   * setData は initialExpandLevel を当て直してしまうため、開閉は自分で持ち回して
   * 新しい木へ写し、以後は -1(データの指定に従う)に切り替える。
   */
  function apply(nextMarkdown, cursorPos, opts = {}) {
    const folds = all.map((n) => (n.payload?.fold ? 1 : 0))
    const nextRoot = toRoot(nextMarkdown)
    contentNodes(nextRoot).forEach((node, i) => {
      const from = opts.insertedAt == null || i < opts.insertedAt ? i : i === opts.insertedAt ? -1 : i - 1
      node.payload = { ...(node.payload || {}), fold: from >= 0 ? folds[from] || 0 : 0 }
    })

    state.markdown = nextMarkdown
    state.root = nextRoot
    mm.setData(nextRoot, { initialExpandLevel: -1 })

    // setData はノードを複製するので、DOMに結び付いた実体を取り直す
    all = contentNodes(state.root)
    ;({ lines, indexes } = nodeLineIndexes(state.markdown))
    current = all[Math.min(Math.max(cursorPos, 0), all.length - 1)] || all[0]
    paint()
    if (opts.changed) options.onChange?.(nextMarkdown)
    if (opts.edit) requestAnimationFrame(() => startEdit())
  }

  function visible() {
    const out = []
    ;(function walk(node) {
      if (String(node?.content ?? '').trim()) out.push(node)
      if (node?.payload?.fold) return
      ;(node?.children || []).forEach(walk)
    })(state.root)
    return out
  }

  function move(delta) {
    const list = visible()
    const at = list.indexOf(current)
    current = list[Math.min(list.length - 1, Math.max(0, at + delta))] || current
    paint()
  }

  async function toggle() {
    if (!current.children?.length) return
    await mm.toggleNode(current)
    paint()
  }

  function startEdit() {
    const g = gOf(current)
    const div = g?.querySelector('.markmap-foreign') || g?.querySelector('foreignObject div')
    if (!editable || editing || !div) return
    const pos = posOf(current)
    const at = indexes[pos]
    const original = splitPrefix(lines[at]).label.trim()

    // SVGのforeignObject内は文字入力を受け付けないブラウザがあるため、
    // 枝と同じ位置にHTMLの入力欄を重ねて編集する
    const rect = div.getBoundingClientRect()
    const base = container.getBoundingClientRect()
    const scale = div.offsetWidth ? rect.width / div.offsetWidth : 1
    const input = document.createElement('input')
    input.className = 'mm-editor'
    input.value = original
    input.style.left = `${rect.left - base.left}px`
    input.style.top = `${rect.top - base.top}px`
    input.style.minWidth = `${Math.max(rect.width + 16, 80)}px`
    input.style.height = `${Math.max(rect.height, 20)}px`
    input.style.fontSize = `${parseFloat(getComputedStyle(div).fontSize || '14') * scale}px`
    container.appendChild(input)

    editing = true
    input.focus()
    input.select()

    const finish = (commit) => {
      if (!editing) return
      editing = false
      const text = input.value.replace(/\s+/g, ' ').trim()
      input.remove()
      if (commit && text && text !== original) {
        const next = [...lines]
        next[at] = splitPrefix(next[at]).prefix + text
        apply(next.join('\n'), pos, { changed: true })
      }
      container.focus({ preventScroll: true })
    }

    input.addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.isComposing) return // 変換中のEnterは入力の確定に使わせる
      if (e.key === 'Enter') { e.preventDefault(); finish(true) }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false) }
    })
    input.addEventListener('blur', () => finish(true))
  }

  function addNode(kind) {
    if (!editable || editing) return
    const pos = posOf(current)
    const line = lines[indexes[pos]]
    // 中心テーマに兄弟を足すとh1が2つになり、根が分かれてしまうので子として足す
    const asChild = kind === 'child' || pos === 0
    const prefix = asChild ? childPrefixOf(line) : splitPrefix(line).prefix
    const endPos = subtreeEnd(lines, indexes, pos)
    const next = [...lines]
    next.splice(indexes[endPos] + 1, 0, prefix + NEW_LABEL)
    if (asChild && current.payload?.fold) current.payload = { ...current.payload, fold: 0 }
    apply(next.join('\n'), endPos + 1, { changed: true, edit: true, insertedAt: endPos + 1 })
  }

  svg.addEventListener('click', (e) => {
    const g = e.target.closest('g.markmap-node')
    const node = g && window.d3?.select(g).datum()
    if (node && posOf(node) !== -1) { current = node; paint() }
    if (!editing) container.focus({ preventScroll: true })
  })

  container.addEventListener('keydown', (e) => {
    if (editing) return
    // マップにカーソルがある間は、画面側のキー操作へ流さない
    e.stopPropagation()
    if (e.key === 'ArrowDown') { e.preventDefault(); move(1) }
    else if (e.key === 'ArrowUp') { e.preventDefault(); move(-1) }
    else if (e.key === 'ArrowRight') {
      e.preventDefault()
      if (current.payload?.fold) toggle()
      else if (current.children?.length) { current = current.children[0]; paint() }
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault()
      if (!current.payload?.fold && current.children?.length) toggle()
      else {
        const parent = all.find((n) => (n.children || []).includes(current))
        if (parent) { current = parent; paint() }
      }
    } else if (e.key === ' ') { e.preventDefault(); startEdit() }
    else if (e.key === 'Tab') { e.preventDefault(); addNode('child') }
    // 編集中のEnterは入力欄側で「決定」に使う
    else if (e.key === 'Enter') { e.preventDefault(); addNode('sibling') }
  })

  paint()
  if (options.autoFocus !== false) container.focus({ preventScroll: true })
}

/** マインドマップタブの描画先に、保存済みのMarkdownを描く */
export function renderMindmapTab(target, markdown, onChange) {
  const host = target.querySelector('#mindmap-host')
  if (host) renderMindmap(host, markdown, { onChange })
}

function label(text) {
  return plainTextOf(String(text ?? '')).replace(/\s+/g, ' ').trim()
}

/**
 * 枝をmarkmap用のMarkdownに組み立てる。
 * 見出し記号やインデントを機械的に出すので、モデルが形を崩す余地がない。
 */
function branchesToMarkdown(title, branches) {
  const lines = [`# ${label(title) || '議事録'}`]

  function walk(nodes, depth) {
    for (const raw of Array.isArray(nodes) ? nodes : []) {
      const text = label(raw?.label ?? raw?.text)
      if (!text) continue
      lines.push(depth === 0 ? `## ${text}` : `${'  '.repeat(depth - 1)}- ${text}`)
      walk(raw?.children, depth + 1)
    }
  }

  walk(branches, 0)
  return lines.join('\n')
}

/** AIを使わず、要約の構成をそのままマインドマップにする */
export function buildMarkdownFromSummary(item, summary) {
  const d = summary?.detail || {}
  const branches = []

  const sentences = plainTextOf(summary?.cardSummary || '')
    .split(/[。\n]/)
    .map((x) => x.trim())
    .filter(Boolean)
  if (sentences.length) branches.push({ label: 'サマリ', children: sentences.map((x) => ({ label: x })) })

  ;(d.agenda || []).forEach((a) => {
    const children = (a.points || []).map((p) => ({ label: p }))
    if (a.outcome) children.push({ label: a.outcome })
    branches.push({ label: a.topic || '議題', children })
  })

  if (d.decisions?.length) branches.push({ label: '決定事項', children: d.decisions.map((x) => ({ label: x })) })
  if (d.todos?.length) branches.push({ label: 'ToDo', children: d.todos.map((t) => ({ label: t?.text ?? t })) })
  if (d.topics?.length) branches.push({ label: '論点', children: d.topics.map((x) => ({ label: x })) })

  return branchesToMarkdown(item.title, branches)
}

const INSTRUCTION = `次の原文を日本語のマインドマップに構造化してください。
出力はJSONのみ。前置き・コードフェンス・説明は一切書かないこと。
形式: {"title":"中心テーマ","branches":[{"label":"見出し","children":[{"label":"要点"}]}]}
原文にない情報を足さないこと。`

function extractJson(text) {
  const trimmed = text.trim()
  const start = trimmed.indexOf('{')
  const end = trimmed.lastIndexOf('}')
  if (start === -1 || end === -1) throw new Error('LLM応答からJSONを抽出できませんでした')
  return JSON.parse(trimmed.slice(start, end + 1))
}

/** 文字起こし全文をもとにAIでマインドマップのMarkdownを生成する */
export async function generateMarkdownWithAI(item, transcript) {
  const connection = connectionOf(loadSettings())
  if (!connection) throw new Error('LLM接続が未設定です。設定から接続先とモデルを追加してください。')

  const source = `会議名: ${item.title}\n\n${String(transcript ?? '')}`.slice(0, 30000)
  const messages = [{ role: 'user', content: INSTRUCTION + '\n\n原文:\n' + source }]

  let full = ''
  for await (const chunk of streamChat(connection, messages)) {
    if (chunk.delta) full += chunk.delta
  }

  const parsed = extractJson(full)
  const branches = Array.isArray(parsed?.branches) ? parsed.branches : []
  if (!branches.length) throw new Error('ノードが生成されませんでした')
  return branchesToMarkdown(parsed?.title || item.title, branches)
}
