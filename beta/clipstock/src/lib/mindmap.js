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
 * @param {{onNodeClick?: (nodeIndex: number) => void}} options
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
    const { root } = new Transformer().transform(markerToHtml(linkTimecodes(raw, videoUrl)))
    Markmap.create(svg, { duration: 200, spacingVertical: 6, paddingX: 12 }, root)
    if (options.onNodeClick) bindNodeClick(svg, root, options.onNodeClick)
  } catch (err) {
    // 描画できなくても内容は読めるようにしておく
    container.innerHTML = `
      <p class="error-text">${err.message || err}</p>
      <pre class="mindmap-fallback">${raw.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</pre>
    `
  }
}

/**
 * ノードのクリックを、そのノードが何行目から作られたかに変換して渡す。
 * 木を前順でたどった順番と、Markdownの空行を除いた行の順番は一致する。
 * 折りたたみの丸とリンクは markmap 側の操作なので拾わない。
 */
function bindNodeClick(svg, root, onNodeClick) {
  const indexOf = new Map()
  // 先頭が見出しでないMarkdownでは、中身の無い根が1つ足される。数えると1つずれる
  ;(function walk(node) {
    if (!node) return
    if (String(node.content ?? '').trim()) indexOf.set(node, indexOf.size)
    ;(node.children || []).forEach(walk)
  })(root)

  svg.addEventListener('click', (e) => {
    if (e.target.closest('a') || e.target.closest('circle')) return
    const g = e.target.closest('g.markmap-node')
    if (!g) return
    const index = indexOf.get(window.d3?.select(g).datum())
    if (index === undefined) return
    const label = g.querySelector('.markmap-foreign') || g.querySelector('foreignObject div')
    if (label) onNodeClick(index, label)
  })
}
