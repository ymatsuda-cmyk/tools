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

/** container の中にマインドマップを描画する */
export async function renderMindmap(container, markdown) {
  const raw = String(markdown ?? '').trim()
  container.innerHTML = ''
  if (!raw) return

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
  svg.classList.add('mindmap-svg')
  container.appendChild(svg)

  try {
    const { Markmap, Transformer } = await loadMarkmap()
    const { root } = new Transformer().transform(raw)
    Markmap.create(svg, { duration: 200, spacingVertical: 6, paddingX: 12 }, root)
  } catch (err) {
    // 描画できなくても内容は読めるようにしておく
    container.innerHTML = `
      <p class="error-text">${err.message || err}</p>
      <pre class="mindmap-fallback">${raw.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</pre>
    `
  }
}

/** マインドマップタブの描画先に、保存済みのMarkdownを描く */
export function renderMindmapTab(target, markdown) {
  const host = target.querySelector('#mindmap-host')
  if (host) renderMindmap(host, markdown)
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
