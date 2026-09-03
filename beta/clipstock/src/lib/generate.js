import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'
import { serializeSections } from './sections.js'
import { reconcileTags } from './tags.js'
import { splitTranscript, resolveQuote, withTimecode, hasTimecodes } from './timecode.js'

/**
 * 生成は3段に分けている。1回のJSONに全部詰めると、
 *  - 小さいモデルでは後半が切れて丸ごと欠落する
 *  - 途中で失敗したとき全部やり直しになる
 *  - 「分野別だけ作り直す」ができない
 * ため。段ごとに保存できるので、途中で止まっても手前の結果は残る。
 */
export const STAGES = [
  { id: 'core', label: 'サマリ・マインドマップ・タグ' },
  { id: 'fields', label: '分野別要約' },
  { id: 'apply', label: '応用と活用アイデア' },
]

// 文字起こしをそのまま渡す上限。Gemma(32K)想定で余白を見た値
const TRANSCRIPT_LIMIT = 30000

function jsonOf(text) {
  const trimmed = String(text ?? '').trim()
  const start = trimmed.indexOf('{')
  const end = trimmed.lastIndexOf('}')
  if (start === -1 || end === -1) throw new Error('LLM応答からJSONを抽出できませんでした')
  return JSON.parse(trimmed.slice(start, end + 1))
}

function requireConnection() {
  const connection = connectionOf(loadSettings())
  if (!connection) {
    throw new Error('AI接続が未設定です。設定から接続先とモデルを追加してください。')
  }
  return connection
}

/** HTTPステータスをメッセージ文字列から拾う。streamChatは "HTTP 429 ..." の形で投げてくる */
function httpStatusOf(err) {
  const m = String(err?.message ?? '').match(/HTTP (\d{3})/)
  return m ? Number(m[1]) : null
}

/**
 * 429には2種類ある。
 *  - 分あたりの上限(RPM): 少し待てば直る一時的なもの
 *  - 日あたりの上限(RPD): 待っても直らない。深夜(太平洋時間)まで復旧しない
 * メッセージ本文で見分けている。判別できない場合は一時的な方として扱い、
 * 待っても直らなければ最終的にリトライ上限で諦める。
 */
function isDailyQuota(err) {
  return /quota|exceeded your current quota/i.test(String(err?.message ?? '')) && !/per minute|rpm/i.test(String(err?.message ?? ''))
}

const RATE_LIMIT_RETRIES = 4
const RATE_LIMIT_BASE_MS = 15000

async function ask(connection, system, user, onProgress) {
  const messages = [
    { role: 'system', content: system },
    { role: 'user', content: user },
  ]

  for (let attempt = 0; ; attempt++) {
    try {
      let full = ''
      for await (const chunk of streamChat(connection, messages)) {
        if (chunk.delta) {
          full += chunk.delta
          onProgress?.(full)
        }
      }
      if (!full.trim()) throw new Error('AIから応答がありませんでした')
      return full
    } catch (err) {
      if (httpStatusOf(err) !== 429 || attempt >= RATE_LIMIT_RETRIES) throw err
      if (isDailyQuota(err)) {
        // 待っても直らない種類なので、リトライで時間を浪費せずすぐ諦める
        throw new Error('1日の利用上限に達しています。日付が変わるまで待つか、プランを見直してください。')
      }
      // 分あたりの上限。指数的に間を空けて再試行する
      const wait = RATE_LIMIT_BASE_MS * 2 ** attempt
      onProgress?.(`(利用制限のため ${Math.round(wait / 1000)} 秒待って再試行します: ${attempt + 1}/${RATE_LIMIT_RETRIES})`)
      await new Promise((r) => setTimeout(r, wait))
    }
  }
}

const NO_FENCE = '前後に説明文やコードフェンス(```)を付けず、JSONだけを出力してください。日本語で書いてください。'

// ---- 第1段: サマリ / マインドマップ / タグ ----

/**
 * タグは既存の語彙に寄せさせる。生成は動画1本ずつ独立に走るため、
 * 何も縛らないと「AI」「生成AI」「LLM」「人工知能」のように語が際限なく増え、
 * 溜まるほど絞り込みが効かなくなる。
 * 語彙が1件も無い初回だけは、自由に付けさせて土台を作る。
 */
function tagRule(knownTags) {
  if (!knownTags.length) {
    return `- tags: 内容を表す短い語を3〜6件。長い説明ではなく分類語にすること。`
  }
  return `- tags: 3〜6件。まず下の「既存のタグ」から当てはまるものを選ぶこと。
  綴りは1文字も変えずにそのまま使う(「AI」を「Ai」や「生成AI」に書き換えない)。
  既存のどれにも当てはまらない観点が中心的な主題である場合にかぎり、新しいタグを1件だけ作ってよい。
  「近い意味だから」で新語を作らないこと。迷ったら既存のタグを選ぶ。

既存のタグ:
${knownTags.map((t) => `- ${t}`).join('\n')}`
}

/**
 * マインドマップは自由記述のMarkdown文字列ではなく構造化JSONで受け取る。
 * 理由は2つ: (1) 分野別と同じ「引用→時刻」を枝ごとに割り当てるには、
 * どの行がどの枝かをこちらで把握できる形でないと後処理できない。
 * (2) 自由記述だとインデントや見出し記号をモデルが崩し、markmapの
 * パースに失敗する事故が起きやすい。Markdown文字列はこちらで組み立てる。
 */
function mindmapNodeShape(withQuotes) {
  const quote = withQuotes
    ? `, "quote": "この項目の根拠になった原文の一文。原文から一字一句そのまま写す。無ければ空文字"`
    : ''
  return `{ "label": "見出しの語句(40字以内)"${quote}, "children": [ /* 同じ形。無ければ省略可 */ ] }`
}

function coreSystem(knownTags, withQuotes) {
  return `あなたは動画の文字起こしを整理するアシスタントです。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "summary": "この動画が何を扱い、何を主張しているかを300〜500字で。前置きや「この動画では」といった枕詞は書かず、内容そのものから始める",
  "tags": ["内容を表す短い語"],
  "mindmap": {
    "title": "動画の主題",
    "branches": [${mindmapNodeShape(withQuotes)}]
  }
}

tags の決め方:
${tagRule(knownTags)}

mindmap の決め方:
- branches(大項目)は3〜6個。各branchのchildrenは中項目、そのchildrenは詳細(最大3段)。
- label は文ではなく要点の語句。40字以内。
- 全体で40〜80項目を超えないこと(枝を無理に増やさない)。${
    withQuotes
      ? `
- quote は原文に存在する文字列でなければならない。20〜60字程度で写す。自分で言い換えた文を書いてはいけない。
  該当が無い、または要約や見出しとして作った項目(原文の特定の一文に対応しない)には quote を空文字にする。
- 時刻や秒数は書かないこと。こちらで原文から割り出す。`
      : ''
  }`
}

// ---- 第2段: 分野別要約 ----

/**
 * 時刻は聞かない。要点の根拠になった原文の引用だけを出させて、
 * こちら側で文字列一致から時刻を割り当てる(resolveQuote)。
 * 「何分何秒か」を直接聞くとモデルは平然と作るので、その道は塞ぐ。
 */
function fieldsSystem(withQuotes) {
  const pointShape = withQuotes
    ? `{ "text": "押さえるべき具体的な事実・数値・手順を1行で", "quote": "その根拠になった原文の一文。原文から一字一句そのまま写す(要約・言い換え・省略をしない)" }`
    : `"押さえるべき具体的な事実・数値・手順を1行で"`

  return `あなたは動画の内容を分野ごとに切り分けて整理するアシスタントです。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "fields": [
    {
      "name": "分野名(例: 技術/経営/マーケティング/組織/法務/学習 など、内容に合うもの)",
      "summary": "その分野の観点から見たこの動画の要点を80〜150字で",
      "points": [${pointShape}]
    }
  ]
}

制約:
- 分野は内容から自然に立つものだけを2〜5件挙げる。無理に埋めない。
- points は分野ごとに2〜4件。
- 動画で実際に語られていないことは書かない。推測は入れない。${
    withQuotes
      ? `
- quote は原文に存在する文字列でなければならない。20〜60字程度で写す。
  自分で言い換えた文を quote に書いてはいけない。該当が無ければ quote は空文字にする。
- 時刻や秒数は書かないこと。こちらで原文から割り出す。`
      : ''
  }
- 全体で1800字以内に収める。`
}

// ---- 第3段: 応用(ビジネス展開) / 活用アイデア ----

const APPLY_SYSTEM = `あなたは、動画から得た知識を実務と遊びの両面に落とし込む企画者です。
次のJSON形式のみで回答してください。${NO_FENCE}

{
  "apply": [
    {
      "title": "ビジネス展開の案。名詞句で短く",
      "summary": "誰のどんな課題をどう解くのかを80〜150字で",
      "steps": ["明日から着手できる具体的な一歩を1件1行で。2〜4件"]
    }
  ],
  "ideas": [
    {
      "title": "面白い活用方法。実用性より発想の飛び方を優先した案",
      "summary": "何をするとどう面白いのかを60〜120字で"
    }
  ]
}

制約:
- apply は2〜4件。「AIを活用する」のような一般論ではなく、この動画の内容が効いている案にする。
- ideas は2〜4件。個人の趣味・家庭・遊び・学習など、仕事以外の文脈も歓迎する。
- 全体で1800字以内に収める。`

// ---- マインドマップの組み立て ----

/**
 * mindmapのJSONをmarkmap用のMarkdownに組み立てる。
 * quote が原文で見つかった枝にだけ末尾へ [mm:ss] を付ける
 * (renderMindmap側がそれをリンクに変換する)。
 * 見出し記号やインデントをこちらで機械的に出すので、モデルが崩す余地がない。
 */
function mindmapToText(parsed, segments) {
  const title = String(parsed?.title ?? '').trim() || '動画の主題'
  const lines = [`# ${title}`]

  function walk(nodes, depth) {
    for (const raw of Array.isArray(nodes) ? nodes : []) {
      const label = String(raw?.label ?? '').trim()
      if (!label) continue
      const quote = String(raw?.quote ?? '')
      const at = segments.length ? resolveQuote(quote, segments) : null
      const text = withTimecode(label, at)

      if (depth === 0) {
        lines.push(`## ${text}`)
      } else {
        lines.push(`${'  '.repeat(depth - 1)}- ${text}`)
      }
      if (Array.isArray(raw?.children) && raw.children.length) walk(raw.children, depth + 1)
    }
  }

  walk(parsed?.branches, 0)
  return lines.join('\n')
}

// ---- セクション形式への変換 ----

/**
 * 分野別をセクション形式に変換する。
 * points は文字列でも {text, quote} でも受ける(モデルが形を崩しても落ちないように)。
 * quote が原文に見つかった項目にだけ末尾へ [mm:ss] を足す。
 */
function fieldsToText(list, segments) {
  return serializeSections(
    (Array.isArray(list) ? list : []).map((f) => {
      const points = []
      let earliest = null

      ;(Array.isArray(f?.points) ? f.points : []).forEach((p) => {
        const text = String((typeof p === 'string' ? p : p?.text) ?? '').trim()
        if (!text) return
        const quote = typeof p === 'string' ? '' : String(p?.quote ?? '')
        const at = segments.length ? resolveQuote(quote, segments) : null
        if (at !== null && (earliest === null || at < earliest)) earliest = at
        points.push(withTimecode(text, at))
      })

      return {
        heading: withTimecode(String(f?.name ?? '').trim(), earliest),
        body: String(f?.summary ?? '').trim(),
        points,
      }
    })
  )
}

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

/** 第3段のコンテキスト。原文全文ではなくサマリ+分野別で足りるので軽い */
function applyContext(title, summaryText, fieldsText) {
  return [`動画タイトル: ${title}`, '', '# サマリ', summaryText, '', '# 分野別要約', fieldsText]
    .join('\n')
    .slice(0, 8000)
}

/**
 * 1段だけ生成する。
 * @param {string} stageId 'core' | 'fields' | 'apply'
 * @param {{title: string, transcript: string, summary?: string, fields?: string}} ctx
 * @param {(text: string) => void} [onProgress] ストリーミング中の生テキスト
 * @returns {Promise<{detail: object, model: string}>} detail は saveGenerated にそのまま渡せる形
 */
export async function generateStage(stageId, ctx, onProgress) {
  const connection = requireConnection()
  const transcript = String(ctx.transcript ?? '').slice(0, TRANSCRIPT_LIMIT)

  if (stageId === 'core') {
    const knownTags = Array.isArray(ctx.knownTags) ? ctx.knownTags : []
    // 文字起こしにタイムスタンプが無い(旧データ)なら引用も求めない。無駄に出力が伸びるだけ
    const timed = hasTimecodes(transcript)
    const segments = timed ? splitTranscript(transcript) : []
    const raw = await ask(connection, coreSystem(knownTags, timed), `動画タイトル: ${ctx.title}\n\n${transcript}`, onProgress)
    const parsed = jsonOf(raw)
    // プロンプトで縛っても表記ゆれは残るので、既存の綴りへ機械的に寄せ直す
    const reconciled = reconcileTags(parsed.tags, knownTags)
    return {
      model: connection.model,
      tagReport: reconciled,
      detail: {
        summary: String(parsed.summary || '').trim(),
        mindmap: mindmapToText(parsed.mindmap, segments),
        tags: reconciled.tags,
      },
    }
  }

  if (stageId === 'fields') {
    // 文字起こしにタイムスタンプが無い(旧データ)なら引用も求めない。無駄に出力が伸びるだけ
    const timed = hasTimecodes(transcript)
    const segments = timed ? splitTranscript(transcript) : []
    const raw = await ask(connection, fieldsSystem(timed), `動画タイトル: ${ctx.title}\n\n${transcript}`, onProgress)
    const parsed = jsonOf(raw)
    return { model: connection.model, detail: { fields: fieldsToText(parsed.fields, segments) } }
  }

  if (stageId === 'apply') {
    const context = applyContext(ctx.title, ctx.summary || '', ctx.fields || '')
    const raw = await ask(connection, APPLY_SYSTEM, context, onProgress)
    const parsed = jsonOf(raw)
    return {
      model: connection.model,
      detail: { apply: applyToText(parsed.apply), ideas: ideasToText(parsed.ideas) },
    }
  }

  throw new Error('unknown stage: ' + stageId)
}

/**
 * 全段をまとめて生成する。段が終わるたびに onStage で通知するので、
 * 呼び出し側はその都度Notionへ保存できる(途中で失敗しても手前は残る)。
 * @param {(stageId: string, detail: object, model: string) => Promise<void>|void} onStage
 */
export async function generateAll(ctx, { onStage, onStageStart, onProgress } = {}) {
  const acc = { summary: ctx.summary || '', fields: ctx.fields || '' }
  let model = null
  let tagReport = null

  for (const stage of STAGES) {
    onStageStart?.(stage)
    const res = await generateStage(
      stage.id,
      { ...ctx, summary: acc.summary, fields: acc.fields },
      (text) => onProgress?.(stage, text)
    )
    const detail = res.detail
    model = res.model
    if (res.tagReport) tagReport = res.tagReport
    if (typeof detail.summary === 'string' && detail.summary) acc.summary = detail.summary
    if (typeof detail.fields === 'string' && detail.fields) acc.fields = detail.fields
    await onStage?.(stage.id, detail, model)
  }

  return { model, tagReport }
}
