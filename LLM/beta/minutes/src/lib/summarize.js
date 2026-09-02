import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf } from './llm-settings.js'

const SYSTEM_PROMPT = `あなたは会議の文字起こしを要約するアシスタントです。
必ず次のJSON形式のみで回答してください。前後に説明文やコードフェンスを付けないこと。

{
  "cardSummary": "一覧カードに表示する100字程度の要約",
  "agenda": [
    {
      "topic": "議題名",
      "points": ["その議題で話された経緯や論拠を、時系列に沿って1件1文字列で"],
      "outcome": "その議題の結論。結論が出ていなければ「継続検討」等と書く"
    }
  ],
  "decisions": ["決定事項を1件1文字列で"],
  "todos": ["ToDoを1件1文字列で。担当者や期限が分かれば含める"],
  "topics": ["未解決の論点や気になる点を1件1文字列で"]
}

agenda は必須項目です。省略や空配列は認められません。会議で話された議題を
必ず1つ以上、議題ごとに分けて出力してください。points には「なぜそうなったか」が
後から追えるよう、背景・提起された問題・検討された選択肢を順に並べてください。
decisions / todos / topics は該当が無ければ空配列で構いません。
5つのキー(cardSummary, agenda, decisions, todos, topics)すべてを含めてください。
日本語で出力してください。`

function extractJson(text) {
  const trimmed = text.trim()
  const start = trimmed.indexOf('{')
  const end = trimmed.lastIndexOf('}')
  if (start === -1 || end === -1) throw new Error('LLM応答からJSONを抽出できませんでした')
  return JSON.parse(trimmed.slice(start, end + 1))
}

/**
 * 文字起こし全文からカード要約と詳細を生成する。
 * @param {string} transcriptText
 * @param {(partial: string) => void} [onProgress] ストリーミング中のテキストを都度受け取る
 * @returns {Promise<{cardSummary: string, decisions: string[], todos: string[], topics: string[], model: string}>}
 */
export async function generateSummary(transcriptText, onProgress) {
  const settings = loadSettings()
  const connection = connectionOf(settings)
  if (!connection) {
    throw new Error('LLM接続が未設定です。設定から接続先とモデルを追加してください。')
  }

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'user', content: transcriptText.slice(0, 30000) }, // モデルの実質上限に合わせて切り詰め
  ]

  let full = ''
  for await (const chunk of streamChat(connection, messages)) {
    if (chunk.delta) {
      full += chunk.delta
      onProgress?.(full)
    }
  }

  const parsed = extractJson(full)
  const agenda = Array.isArray(parsed.agenda) ? parsed.agenda : []
  return {
    cardSummary: String(parsed.cardSummary || ''),
    agenda,
    decisions: Array.isArray(parsed.decisions) ? parsed.decisions : [],
    todos: (Array.isArray(parsed.todos) ? parsed.todos : []).map((t) => ({ text: String(t), done: false })),
    topics: Array.isArray(parsed.topics) ? parsed.topics : [],
    model: connection.model,
    // 議事は入れ子構造のため小さいモデルでは省略されることがある。
    // 黙って空のまま保存すると「議事が未登録」に見えてしまうため、呼び出し側へ通知する。
    agendaMissing: agenda.length === 0,
  }
}
