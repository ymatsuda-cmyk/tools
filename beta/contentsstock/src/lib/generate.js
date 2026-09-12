/**
 * 文字起こし全文からAI生成物を作る。
 *
 * 段ごとに分けてあるのは「この項目だけ作り直す」を1タブ単位でできるようにするため。
 * どの段もJSONで受け取り、Notionに保存する形(セクション形式 / markmap用Markdown)へは
 * こちら側で組み立てる。モデルに整形させると形が崩れて読めなくなるため。
 */
import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'
import { serializeSections } from './sections.js'

export const STAGES = [
  { id: 'summary', label: 'サマリ' },
  { id: 'mindmap', label: 'マインドマップ' },
  { id: 'fields', label: '分野別' },
  { id: 'apply', label: '応用' },
  { id: 'ideas', label: '活用' },
]

const TRANSCRIPT_LIMIT = 40000
const NO_FENCE = '前後に説明文やコードフェンス(```)を付けず、JSONだけを出力してください。日本語で書いてください。'
const SHARED_SYSTEM = '日本語で回答してください。'

function requireConnection() {
  const connection = connectionOf(loadSettings())
  if (!connection) throw new Error('LLM接続が未設定です。設定から接続先とモデルを追加してください。')
  return connection
}

async function ask(connection, system, user, onProgress) {
  const messages = [
    { role: 'system', content: system },
    { role: 'user', content: user },
  ]
  let full = ''
  for await (const chunk of streamChat(connection, messages)) {
    if (chunk.delta) {
      full += chunk.delta
      onProgress?.(full)
    }
  }
  return full
}

function jsonOf(text) {
  const trimmed = String(text ?? '').trim()
  const start = trimmed.indexOf('{')
  const end = trimmed.lastIndexOf('}')
  if (start === -1 || end === -1) throw new Error('AIの応答からJSONを取り出せませんでした')
  return JSON.parse(trimmed.slice(start, end + 1))
}

/**
 * タグは既存の語彙に寄せさせる。生成は1本ずつ独立に走るため、何も縛らないと
 * 「AI」「生成AI」「LLM」のように語が際限なく増え、溜まるほど絞り込みが効かなくなる。
 */
function tagRule(knownTags) {
  if (!knownTags.length) return '- tags: 内容を表す短い語を3〜6件。長い説明ではなく分類語にすること。'
  return `- tags: 3〜6件。まず下の「既存のタグ」から当てはまるものを選ぶこと。
  綴りは1文字も変えずにそのまま使う。既存のどれにも当てはまらない観点が中心的な主題である場合にかぎり、
  新しいタグを1件だけ作ってよい。迷ったら既存のタグを選ぶ。

既存のタグ:
${knownTags.map((t) => `- ${t}`).join('\n')}`
}

const PROMPTS = {
  summary: (knownTags) => `あなたは動画の文字起こしを整理するアシスタントです。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "summary": "この動画が何を扱い、何を主張しているかを300〜500字で。「この動画では」といった枕詞は書かず内容から始める",
  "tags": ["内容を表す短い語"]
}

tags の決め方:
${tagRule(knownTags)}`,

  mindmap: `あなたは動画の内容をマインドマップに整理するアシスタントです。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "title": "動画の主題",
  "branches": [
    { "label": "見出しの語句(40字以内)", "children": [ { "label": "同じ形。無ければ省略可" } ] }
  ]
}

制約:
- branches(大項目)は3〜6個。各branchのchildrenは中項目、そのchildrenは詳細(最大3段)。
- label は文ではなく要点の語句。40字以内。
- 原文にない情報を足さない。`,

  fields: `あなたは動画の内容を分野ごとに切り分けて整理するアシスタントです。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "fields": [
    {
      "name": "分野名(例: 技術/経営/マーケティング/組織/法務/学習 など、内容に合うもの)",
      "summary": "その分野の観点から見たこの動画の要点を80〜150字で",
      "points": ["押さえるべき具体的な事実・数値・手順を1行で"]
    }
  ]
}

制約:
- 分野は内容から自然に立つものだけを2〜5件。無理に埋めない。
- points は分野ごとに2〜4件。
- 実際に語られていないことは書かない。推測は入れない。
- 全体で1800字以内に収める。`,

  apply: `あなたは、動画から得た知識を実務に落とし込む企画者です。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "apply": [
    {
      "title": "ビジネス展開の案。名詞句で短く",
      "summary": "誰のどんな課題をどう解くのかを80〜150字で",
      "steps": ["明日から着手できる具体的な一歩を1件1行で。2〜4件"]
    }
  ]
}

制約:
- apply は2〜4件。一般論ではなく、この動画の内容が効いている案にする。
- 全体で1800字以内に収める。`,

  ideas: `あなたは、動画から得た知識を遊びや暮らしに落とし込む発想者です。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "ideas": [
    { "title": "面白い活用方法。実用性より発想の飛び方を優先した案", "summary": "何をするとどう面白いのかを60〜120字で" }
  ]
}

制約:
- ideas は3〜5件。個人の趣味・家庭・遊び・学習など、仕事以外の文脈も歓迎する。
- 全体で1200字以内に収める。`,
}

/**
 * 段ごとの指示は system ではなく本文の末尾に置く。
 * ローカルLLMは「直前のプロンプトと先頭から一致している範囲」の読み込みを省けるので、
 * 原文を先・指示を後ろにすれば、2段目以降は原文ぶんの読み込みを飛ばせる。
 */
function withInstructions(context, instructions) {
  return `${context}\n\n----------------\n\n${instructions}`
}

/** mindmapのJSONをmarkmap用のMarkdownに組み立てる */
function mindmapToText(parsed) {
  const lines = [`# ${String(parsed?.title ?? '').trim() || '動画の主題'}`]

  function walk(nodes, depth) {
    for (const raw of Array.isArray(nodes) ? nodes : []) {
      const label = String(raw?.label ?? '').replace(/\s+/g, ' ').trim()
      if (!label) continue
      lines.push(depth === 0 ? `## ${label}` : `${'  '.repeat(depth - 1)}- ${label}`)
      if (Array.isArray(raw?.children) && raw.children.length) walk(raw.children, depth + 1)
    }
  }

  walk(parsed?.branches, 0)
  return lines.join('\n')
}

function fieldsToText(list) {
  return serializeSections(
    (Array.isArray(list) ? list : []).map((f) => ({
      heading: String(f?.name ?? '').trim(),
      body: String(f?.summary ?? '').trim(),
      points: (Array.isArray(f?.points) ? f.points : []).map((p) =>
        String((typeof p === 'string' ? p : p?.text) ?? '').trim()
      ),
    }))
  )
}

// 星は生成時には付けない(全件星0)。読んだあとに自分で付けるため
function applyToText(list) {
  return serializeSections(
    (Array.isArray(list) ? list : []).map((a) => ({
      heading: String(a?.title ?? '').trim(),
      body: String(a?.summary ?? '').trim(),
      points: Array.isArray(a?.steps) ? a.steps : [],
    }))
  )
}

function ideasToText(list) {
  return serializeSections(
    (Array.isArray(list) ? list : []).map((a) => ({
      heading: String(a?.title ?? '').trim(),
      body: String(a?.summary ?? '').trim(),
      points: [],
    }))
  )
}

/** 第3・4段のコンテキスト。原文全文ではなくサマリ+分野別で足りるので軽い */
function applyContext(title, summaryText, fieldsText) {
  return [`動画タイトル: ${title}`, '', '# サマリ', summaryText, '', '# 分野別要約', fieldsText]
    .join('\n')
    .slice(0, 8000)
}

/**
 * 1段だけ生成する。
 * @param {'summary'|'mindmap'|'fields'|'apply'|'ideas'} stageId
 * @param {{title: string, transcript: string, summary?: string, fields?: string, knownTags?: string[]}} ctx
 * @returns {Promise<{detail: object, model: string}>} detail は saveGenerated にそのまま渡せる形
 */
export async function generateStage(stageId, ctx, onProgress) {
  const connection = requireConnection()
  const transcript = String(ctx.transcript ?? '').slice(0, TRANSCRIPT_LIMIT)
  const transcriptInput = `動画タイトル: ${ctx.title}\n\n${transcript}`

  if (stageId === 'summary') {
    const knownTags = Array.isArray(ctx.knownTags) ? ctx.knownTags : []
    const parsed = jsonOf(await ask(connection, SHARED_SYSTEM, withInstructions(transcriptInput, PROMPTS.summary(knownTags)), onProgress))
    return {
      model: connection.model,
      detail: {
        summary: String(parsed.summary || '').trim(),
        tags: (Array.isArray(parsed.tags) ? parsed.tags : []).map((t) => String(t).trim()).filter(Boolean),
      },
    }
  }

  if (stageId === 'mindmap') {
    const parsed = jsonOf(await ask(connection, SHARED_SYSTEM, withInstructions(transcriptInput, PROMPTS.mindmap), onProgress))
    return { model: connection.model, detail: { mindmap: mindmapToText(parsed) } }
  }

  if (stageId === 'fields') {
    const parsed = jsonOf(await ask(connection, SHARED_SYSTEM, withInstructions(transcriptInput, PROMPTS.fields), onProgress))
    return { model: connection.model, detail: { fields: fieldsToText(parsed.fields) } }
  }

  if (stageId === 'apply' || stageId === 'ideas') {
    const context = applyContext(ctx.title, ctx.summary || '', ctx.fields || '')
    const parsed = jsonOf(await ask(connection, SHARED_SYSTEM, withInstructions(context, PROMPTS[stageId]), onProgress))
    return {
      model: connection.model,
      detail: stageId === 'apply' ? { apply: applyToText(parsed.apply) } : { ideas: ideasToText(parsed.ideas) },
    }
  }

  throw new Error('unknown stage: ' + stageId)
}

/**
 * 全段をまとめて生成する。段が終わるたびに onStage で通知するので、
 * 呼び出し側はその都度保存できる(途中で失敗しても手前は残る)。
 * 同じ原文を見る段を続けて直列に走らせるのは意図的で、ローカルLLMでは
 * 同時に投げるより順番に投げたほうが速い。
 */
export async function generateAll(ctx, onStage, onProgress) {
  const acc = { ...ctx }
  for (const stage of STAGES) {
    const { detail, model } = await generateStage(stage.id, acc, (text) => onProgress?.(stage.id, text))
    Object.assign(acc, detail)
    await onStage?.(stage.id, detail, model)
  }
  return acc
}
