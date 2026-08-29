import { streamChat } from './llm-client.js'
import { loadSettings, connectionOf, activeProfile } from './llm-settings.js'

const SYSTEM_PROMPT = `あなたは会議の文字起こしを要約するアシスタントです。
必ず次のJSON形式のみで回答してください。前後に説明文やコードフェンスを付けないこと。

{
  "cardSummary": "一覧カードに表示する100字程度の要約",
  "decisions": ["決定事項を1件1文字列で"],
  "todos": ["ToDoを1件1文字列で。担当者や期限が分かれば含める"],
  "topics": ["未解決の論点や気になる点を1件1文字列で"]
}

該当する項目が無い場合は空配列にしてください。日本語で出力してください。`

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
  const profile = activeProfile(settings)
  if (!connection) {
    throw new Error('LLM接続プロファイルが未設定です。設定から接続先を追加してください。')
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
  return {
    cardSummary: String(parsed.cardSummary || ''),
    decisions: Array.isArray(parsed.decisions) ? parsed.decisions : [],
    todos: Array.isArray(parsed.todos) ? parsed.todos : [],
    topics: Array.isArray(parsed.topics) ? parsed.topics : [],
    model: profile?.model || connection.model || 'unknown',
  }
}
