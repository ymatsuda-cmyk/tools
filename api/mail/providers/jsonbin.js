/**
 * JSONBinから読み込む取得元（読み込み専用）。
 *
 * メール本文の書き込みはここでは行わない。GAS・共有メールボックスの転送設定など、
 * ダッシュボードの外側から https://api.jsonbin.io/v3/b/{binId} に PUT しておく運用を
 * 想定している（書き込み方の例は api/mail/README.md を参照）。
 *
 * account に binId / apiKey（X-Master-Key、またはjsonbin.ioのAccess Key）を
 * それぞれ持たせる。Bin1つにつきメールボックス1つ、という対応になる
 * （アカウントごとに別々のBin/キーを設定できる）。
 */

export const id = 'jsonbin'
export const label = 'JSONBin（外部から書き込む共有メール）'

export const FIELDS = [
  {
    key: 'binId',
    label: 'JSONBin Bin ID',
    placeholder: '例: 6512abcd1f2e3a4b5c6d7e8f',
    hint: 'jsonbin.io で作成したBinのID'
  },
  {
    key: 'apiKey',
    label: 'JSONBin X-Master-Key（またはAccess Key）',
    placeholder: '$2a$10$...',
    hint: '読み込み専用に使う。書き込みは外部（GAS・手動など）で行う'
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

/* サインインの概念は無い。Bin ID/キーが揃っているかだけ見る */
export async function status(account) {
  if (!account.binId || !account.apiKey) {
    return { signedIn: false, error: 'Bin IDとキーが未設定です' }
  }
  return { signedIn: true }
}

export async function signIn(account) {
  return status(account)
}

export async function signOut() {
  return { signedIn: true }
}

export async function messages(account, { since }) {
  if (!account.binId || !account.apiKey) {
    throw new Error('Bin IDとキーが未設定です')
  }

  const url = `https://api.jsonbin.io/v3/b/${encodeURIComponent(account.binId)}/latest`
  const res = await fetch(url, {
    headers: { 'X-Master-Key': account.apiKey, 'X-Bin-Meta': 'false' },
    cache: 'no-store'
  })
  if (res.status === 404) return [] // まだ何も登録されていない
  if (!res.ok) throw new Error(`JSONBinを取得できません（HTTP ${res.status}）`)

  const raw = unwrap(await res.json())
  return raw
    .map(toMail)
    .filter((e) => e.receivedAt)
    .filter((e) => e.receivedAt >= since)
}

/**
 * 2通りの書き方を受け付ける（NDJSONはJSONBinの性質上想定しない。常に1つの正しいJSON値）。
 *   1. 配列                 [ {...}, {...} ]
 *   2. 包んだオブジェクト    { "mails": [ {...} ] }（mails/events/items/data/value/values でも可）
 *   3. 1件そのもの           {...}（subject か receivedAt を持つオブジェクト）
 */
const WRAP_KEYS = ['mails', 'events', 'items', 'data', 'value', 'values']

function unwrap(json) {
  if (Array.isArray(json)) return json
  if (json && typeof json === 'object') {
    const key = WRAP_KEYS.find((k) => Array.isArray(json[k]))
    if (key) return json[key]
    if ('subject' in json || 'title' in json || 'receivedAt' in json) return [json]
    const keys = Object.keys(json).join(', ') || '(空のオブジェクト)'
    throw new Error(
      `JSONBinの中身を読み取れません（${WRAP_KEYS.map((k) => `"${k}"`).join('/')} に配列を入れる形も可）。実際のキー: ${keys}`
    )
  }
  throw new Error('JSONBinの中身を読み取れません')
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
    id: e.id || `jsonbin-${i}`,
    subject: (e.subject || e.title || '').trim() || '(件名なし)',
    from: e.from || e.sender || '',
    receivedAt: parseDate(e.receivedAt || e.receivedDateTime || e.date),
    isRead: e.isRead !== false,
    preview: e.preview || e.bodyPreview || '',
    url: e.url || e.webLink || ''
  }
}
