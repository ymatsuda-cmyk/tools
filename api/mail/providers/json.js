/**
 * JSONで決め打ちしたメールを出す取得元。
 *
 * Outlookと違ってサインインが要らない。account.mails に配列を直接書くか、
 * account.url を書けば毎回そこから取りに行く（同一オリジンかCORS許可が必要）。
 * どちらも書いてあれば url を優先する。
 *
 * account.maxAgeDays でこのアカウントだけ有効期間（何日前まで）を変えられる。
 * 省略するとカード側（mail.api.js の recent() に渡す既定日数）に従う。
 */

export const id = 'json'
export const label = 'JSON（固定のメール・共有ファイル）'

export const FIELDS = [
  {
    key: 'url',
    label: 'JSONのURL（任意）',
    placeholder: 'https://example.com/mails.json',
    hint: '空欄なら account.mails をそのまま使います。書くと毎回そこから取得します（同一オリジンかCORSが必要）'
  },
  {
    key: 'maxAgeDays',
    label: '何日前まで有効か（任意）',
    placeholder: '7',
    hint: '省略するとカード側の既定日数を使います'
  }
]

export function redirectUri() {
  return ''
}

/* サインインの概念が無いので、常に使える状態として扱う */
export async function status() {
  return { signedIn: true }
}

export async function signIn() {
  return { signedIn: true }
}

export async function signOut() {
  return { signedIn: true }
}

export async function messages(account, { since }) {
  const raw = account.url ? await fetchList(account.url) : Array.isArray(account.mails) ? account.mails : []
  return raw
    .map(toMail)
    .filter((e) => e.receivedAt)
    .filter((e) => e.receivedAt >= since)
}

async function fetchList(url) {
  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`メールのJSONを取得できません（HTTP ${res.status}）`)
  return parseFeed(await res.text())
}

/**
 * 3通りの書き方を受け付ける。
 *   1. 配列                 [ {...}, {...} ]
 *   2. 包んだオブジェクト    { "mails": [ {...} ] }（mails/events/items/data/value/values でも可）
 *   3. 1件そのもの           {...}（subject か receivedAt を持つオブジェクト）
 *   4. 1行に1件のNDJSON      {...}\n{...}\n…（GASなどが1件ずつ追記していく運用向け）
 */
function parseFeed(text) {
  const trimmed = text.trim()
  if (!trimmed) return []

  try {
    return unwrap(JSON.parse(trimmed))
  } catch (err) {
    if (err instanceof FeedShapeError) throw err
    // ひとつのJSONとしては読めない → 1行1件のNDJSONとして読み直す
  }

  return trimmed
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, i) => {
      try {
        return JSON.parse(line)
      } catch {
        throw new Error(`メールのJSON ${i + 1}行目を読み取れません: ${line.slice(0, 60)}`)
      }
    })
}

class FeedShapeError extends Error {}

/** 配列・包まれたオブジェクト・単体のメールオブジェクトのいずれかを配列に揃える */
const WRAP_KEYS = ['mails', 'events', 'items', 'data', 'value', 'values']

function unwrap(json) {
  if (Array.isArray(json)) return json
  if (json && typeof json === 'object') {
    const key = WRAP_KEYS.find((k) => Array.isArray(json[k]))
    if (key) return json[key]
    if ('subject' in json || 'title' in json || 'receivedAt' in json) return [json]
    const keys = Object.keys(json).join(', ') || '(空のオブジェクト)'
    throw new FeedShapeError(
      `メールのJSONを読み取れません（${WRAP_KEYS.map((k) => `"${k}"`).join('/')} に配列を入れる形も可）。実際のキー: ${keys}`
    )
  }
  throw new FeedShapeError('メールのJSONを読み取れません')
}

/** タイムゾーンの表記が無ければUTCとして読む（Outlookの生のGraph応答と揃える） */
function parseDate(value) {
  if (!value) return null
  const raw = String(value)
  const hasZone = /(Z|[+-]\d{2}:\d{2})$/.test(raw)
  const d = new Date(hasZone ? raw : raw + 'Z')
  return isNaN(d.getTime()) ? null : d
}

function toMail(e, i) {
  return {
    id: e.id || `json-${i}`,
    subject: (e.subject || e.title || '').trim() || '(件名なし)',
    from: e.from || e.sender || '',
    receivedAt: parseDate(e.receivedAt || e.receivedDateTime || e.date),
    isRead: e.isRead !== false,
    preview: e.preview || e.bodyPreview || '',
    url: e.url || e.webLink || ''
  }
}
