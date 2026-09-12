/**
 * マインドマップの描画。
 *
 * Notionには「markmap用のMarkdown」を保存する方針にしている。
 * HTMLを丸ごと保存する方式だと、
 *  - CDNのscriptタグ分だけで文字数を食う
 *  - Notion上で開いても中身が読めない
 *  - 描画ライブラリを差し替えられない
 * ため。ただし旧スキルがHTMLを書き込んだページも残っているので、
 * '<' で始まる値は旧形式としてiframeにそのまま流し込む。
 */

import { parseTimecode, youtubeUrlAt, splitLabel, withTimecode } from './timecode.js'
import { MARKER_COLORS } from './markers.js'

/**
 * ファイル名を直接指定しないこと。
 * ブラウザ向けの実体は markmap-view が dist/browser/index.js なのに対し、
 * markmap-lib は dist/browser/index.iife.js と名前が違う。
 * パッケージ名だけを指定すれば、CDNが package.json の jsdelivr フィールドを見て
 * 正しいファイルを返すので、この取り違えが起きない。
 *
 * 読み込み順も変えないこと。両方が window.markmap に生えるが、
 * lib(Transformer) -> view(Markmap) の順が公式ドキュメントの前提。
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

/**
 * markmap一式を読み込む。2回目以降は同じPromiseを返す。
 * どれが欠けたのか分かるよう、1つずつ読み込んで直後に検証する。
 * まとめて読んでから確認すると「初期化に失敗」としか言えず、原因を追えない。
 */
function loadMarkmap() {
  if (!loading) {
    loading = (async () => {
      for (const dep of CDN) {
        await loadScript(dep.url)
        if (!dep.check()) {
          throw new Error(`${dep.name} を読み込めませんでした (${dep.url})`)
        }
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
 * マインドマップ内の "[12:34]" をMarkdownリンクに変える。
 * markmapはMarkdownのリンクをそのままクリック可能にするので、
 * 描画側に手を入れずに枝から動画へ飛べる。
 */
function linkTimecodes(markdown, videoUrl) {
  if (!videoUrl) return markdown
  return String(markdown).replace(/\[((?:\d{1,2}:)?\d{1,3}:\d{2})\]/g, (m, label) => {
    const at = parseTimecode(m)
    const href = at === null ? null : youtubeUrlAt(videoUrl, at)
    return href ? `[${label}](${href})` : m
  })
}

export function isLegacyHtml(value) {
  return String(value ?? '').trim().startsWith('<')
}

// ---- ノードのマーカー ----
//
// 他のタブと同じ <m1>…</m1> を行のラベルに直接埋める。座標を別に持たないので
// あとでMarkdownを直しても壊れず、既存の plainTextOf でそのまま外せる。

const MARKER_TAG = /<m([123])>([\s\S]*?)<\/m\1>/g

/** ノードになる行だけを並び順で返す。markmapの木を前順でたどった順番と一致する */
function nodeLineIndexes(markdown) {
  const lines = String(markdown ?? '').split('\n')
  const out = []
  lines.forEach((line, i) => {
    if (line.trim()) out.push(i)
  })
  return { lines, indexes: out }
}

/** 行頭の "## " や "  - " と、そのあとのラベルを分ける */
function splitPrefix(line) {
  const m = String(line).match(/^(\s*(?:#{1,6}\s+|[-*+]\s+|\d+\.\s+)?)([\s\S]*)$/)
  return { prefix: m[1], label: m[2] }
}

/** nodeIndex 番目のノードに今引かれている色。無ければ null */
export function nodeMarkerOf(markdown, nodeIndex) {
  const { lines, indexes } = nodeLineIndexes(markdown)
  const at = indexes[nodeIndex]
  if (at === undefined) return null
  const m = splitPrefix(lines[at]).label.match(/<m([123])>/)
  return m ? Number(m[1]) : null
}

/**
 * nodeIndex 番目のノードにマーカーを引く。colorIndex が null なら消す。
 * 末尾のタイムコードは外してから包む(包むとリンク化が壊れるため)。
 */
export function markNodeLine(markdown, nodeIndex, colorIndex) {
  const { lines, indexes } = nodeLineIndexes(markdown)
  const at = indexes[nodeIndex]
  if (at === undefined) return String(markdown ?? '')

  const { prefix, label } = splitPrefix(lines[at])
  const { text, at: seconds } = splitLabel(label)
  const bare = text.replace(MARKER_TAG, '$2')
  const marked = colorIndex ? `<m${colorIndex}>${bare}</m${colorIndex}>` : bare
  lines[at] = prefix + withTimecode(marked, seconds)
  return lines.join('\n')
}

/** markmapに渡す前に、マーカーのタグを色付きのspanにする */
function markerToHtml(markdown) {
  return String(markdown).replace(
    MARKER_TAG,
    (_, color, text) => `<span class="mm-mark" style="background:${MARKER_COLORS[color]}">${text}</span>`
  )
}

/**
 * container の中にマインドマップを描画する。
 * @param {HTMLElement} container
 * @param {string} value Notionの「マインドマップ」プロパティの生値
 * @param {string} videoUrl
 * @param {{onNodeClick?: (nodeIndex: number) => void, onChange?: (markdown: string) => void, autoFocus?: boolean}} options
 */
export async function renderMindmap(container, value, videoUrl = '', options = {}) {
  const raw = String(value ?? '').trim()
  container.innerHTML = ''

  if (!raw) {
    container.innerHTML = '<p class="empty-section">マインドマップはまだありません</p>'
    return
  }

  if (isLegacyHtml(raw)) {
    const frame = document.createElement('iframe')
    frame.className = 'mindmap-frame'
    frame.setAttribute('sandbox', 'allow-scripts')
    frame.srcdoc = raw
    container.appendChild(frame)
    return
  }

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
  svg.classList.add('mindmap-svg')
  container.appendChild(svg)

  try {
    const { Markmap, Transformer } = await loadMarkmap()
    const transformer = new Transformer()
    const toRoot = (md) => transformer.transform(markerToHtml(linkTimecodes(md, videoUrl))).root
    const state = { markdown: raw, root: toRoot(raw) }
    const mm = Markmap.create(svg, { duration: 200, spacingVertical: 6, paddingX: 12 }, state.root)
    if (options.onNodeClick) bindNodeClick(svg, state, options.onNodeClick)
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
// 描き直しても同じ枝に戻れる。行の中身(マーカー・タイムコード)はMarkdown側で扱う。

const NEW_LABEL = '新しい項目'

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

/** 編集欄に出す、マーカーもタイムコードも外した文言 */
function bareLabel(line) {
  return splitLabel(splitPrefix(line).label).text.replace(MARKER_TAG, '$2').trim()
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
  const lineOf = (node) => indexes[posOf(node)]

  function paint() {
    svg.querySelectorAll('g.mm-current').forEach((g) => g.classList.remove('mm-current'))
    gOf(current)?.classList.add('mm-current')
  }

  // markmapは枝の大きさを測ってから描き、描き直しのたびにclassを付け直す。
  // 一度付けただけだとカーソルが消えるので、描画が落ち着くまで付け直す
  function paintSoon() {
    paint()
    requestAnimationFrame(paint)
    setTimeout(paint, 300)
  }

  /**
   * カーソルの枝が見えていなければ、見える位置まで盤面を動かす。
   * markmapのzoomをd3のtransition経由で動かすので、滅入するように寄る。
   * 枝の遷移中は位置が定まらないため、呼び出し側が待ち時間を指定する。
   */
  let revealTimer = null
  function reveal(delay = 0) {
    clearTimeout(revealTimer)
    revealTimer = setTimeout(() => {
      const g = gOf(current)
      const d3 = window.d3
      if (!g || !d3 || !mm.svg || !mm.zoom) return
      const box = g.getBoundingClientRect()
      const view = container.getBoundingClientRect()
      const margin = 48
      let dx = 0
      let dy = 0
      if (box.left < view.left + margin) dx = view.left + margin - box.left
      else if (box.right > view.right - margin) dx = view.right - margin - box.right
      if (box.top < view.top + margin) dy = view.top + margin - box.top
      else if (box.bottom > view.bottom - margin) dy = view.bottom - margin - box.bottom
      if (!dx && !dy) return
      const t = d3.zoomTransform(mm.svg.node())
      mm.svg.transition().duration(320).call(mm.zoom.transform, t.translate(dx / t.k, dy / t.k))
    }, delay)
  }

  /**
   * Markdownを差し替える。描き直しではなく markmap にデータだけ渡すので、
   * 表示位置と拡大率はそのままで、増えた枝だけが現れる。
   * setData は initialExpandLevel を当て直してしまうため、開閉は自分で持ち回して
   * 新しい木へ写し、以後は -1(データの指定に従う)に切り替える。
   */
  async function apply(nextMarkdown, cursorPos, opts = {}) {
    const folds = all.map((n) => (n.payload?.fold ? 1 : 0))
    const nextRoot = toRoot(nextMarkdown)
    const nextAll = contentNodes(nextRoot)
    nextAll.forEach((node, i) => {
      const from = opts.insertedAt == null || i < opts.insertedAt ? i : i === opts.insertedAt ? -1 : i - 1
      node.payload = { ...(node.payload || {}), fold: from >= 0 ? folds[from] || 0 : 0 }
    })

    state.markdown = nextMarkdown
    state.root = nextRoot
    if (opts.changed) options.onChange?.(nextMarkdown)

    // setData は大きさを測ってから描くため、終わるのを待たないと枝がまだDOMに無い
    await mm.setData(nextRoot, { initialExpandLevel: -1 })

    // setData はノードを複製するので、DOMに結び付いた実体を取り直す
    all = contentNodes(state.root)
    ;({ lines, indexes } = nodeLineIndexes(state.markdown))
    current = all[Math.min(Math.max(cursorPos, 0), all.length - 1)] || all[0]
    paintSoon()
    // 入力欄は枝の位置に重ねるので、枝の移動と盤面の寄せが終わってから出す
    if (opts.edit) {
      reveal(160)
      setTimeout(startEdit, 500)
    } else {
      reveal(260)
    }
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
    reveal()
  }

  async function toggle() {
    if (!current.children?.length) return
    await mm.toggleNode(current)
    paintSoon()
    reveal(260)
  }

  function startEdit() {
    const g = gOf(current)
    const div = g?.querySelector('.markmap-foreign') || g?.querySelector('foreignObject div')
    if (!editable || editing || !div) return
    const at = lineOf(current)
    const pos = posOf(current)
    const original = bareLabel(lines[at])

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
        const { prefix, label } = splitPrefix(next[at])
        next[at] = prefix + withTimecode(text, splitLabel(label).at)
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
    // 詳細画面は左右キーでタブを切り替えるので、マップにカーソルがある間は上へ流さない
    e.stopPropagation()
    if (e.key === 'ArrowDown') { e.preventDefault(); move(1) }
    else if (e.key === 'ArrowUp') { e.preventDefault(); move(-1) }
    else if (e.key === 'ArrowRight') {
      e.preventDefault()
      if (current.payload?.fold) toggle()
      else if (current.children?.length) { current = current.children[0]; paint(); reveal() }
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault()
      if (!current.payload?.fold && current.children?.length) toggle()
      else {
        const parent = all.find((n) => (n.children || []).includes(current))
        if (parent) { current = parent; paint(); reveal() }
      }
    } else if (e.key === ' ') { e.preventDefault(); startEdit() }
    else if (e.key === 'Tab') { e.preventDefault(); addNode('child') }
    // 編集中のEnterは入力欄側で「決定」に使う
    else if (e.key === 'Enter') { e.preventDefault(); addNode('sibling') }
  })

  paint()
  if (options.autoFocus !== false) container.focus({ preventScroll: true })
}

/**
 * ノードのクリックを、そのノードが何行目から作られたかに変換して渡す。
 * 木を前順でたどった順番と、Markdownの空行を除いた行の順番は一致する。
 * 折りたたみの丸とリンクは markmap 側の操作なので拾わない。
 */
function bindNodeClick(svg, state, onNodeClick) {
  svg.addEventListener('click', (e) => {
    if (e.target.closest('a') || e.target.closest('circle')) return
    const g = e.target.closest('g.markmap-node')
    if (!g) return
    // 枝を足したあとは木が入れ替わるため、対応表は作り置きせずその都度たどる
    const index = contentNodes(state.root).indexOf(window.d3?.select(g).datum())
    if (index < 0) return
    const label = g.querySelector('.markmap-foreign') || g.querySelector('foreignObject div')
    if (label) onNodeClick(index, label)
  })
}
