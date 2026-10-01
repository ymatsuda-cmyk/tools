/**
 * マインドマップの生成と表示。
 * 描画と編集は共通API(api/mindmap2)に任せ、ここは要約からの組み立てだけを持つ。
 *
 * Notionの「マインドマップ」カラムには markmap 用のMarkdownを保存する。
 * HTMLやツリーJSONではなくMarkdownにしているのは、Notion上でもそのまま読めて、
 * 描画ライブラリを差し替えても中身が生き残るため。
 */
import { renderMindmap } from '../../../api/mindmap2/mindmap2.js'
import { plainTextOf } from './markers.js'
import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'

// 中心テーマ(深さ0)と大項目(##、深さ1)まで開き、第3階層以降は畳んだ状態で描く
const EXPAND_LEVEL = 2

/** マインドマップタブの描画先に、保存済みのMarkdownを描く */
export function renderMindmapTab(target, markdown, onChange) {
  const host = target.querySelector('#mindmap-host')
  if (host) renderMindmap(host, markdown, { onChange, initialExpandLevel: EXPAND_LEVEL })
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
