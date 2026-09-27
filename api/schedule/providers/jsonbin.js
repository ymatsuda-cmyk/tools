/**
 * JSONBinから読み込む取得元（読み込み専用）。
 *
 * 予定の書き込みはここでは行わない。GAS・手動など、ダッシュボードの外側から
 * https://api.jsonbin.io/v3/b/{binId} に PUT しておく運用を想定している
 * （書き込み方の例は api/schedule/README.md を参照）。
 *
 * account に binId / apiKey（X-Master-Key、またはjsonbin.ioのAccess Key）を
 * それぞれ持たせる。Bin1つにつき予定表1つ、という対応になる
 * （アカウントごとに別々のBin/キーを設定できる）。
 */

export const id = 'jsonbin'
export const label = 'JSONBin（外部から書き込む共有予定）'

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

export async function events(account, { from, to }) {
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
    .map(toEvent)
    .filter((e) => e.start)
    .filter((e) => {
      // allDayは終了時刻を持たないことがあるので、翌日0時を仮の終わりとして扱う
      const effEnd = e.end || (e.allDay ? new Date(e.start.getTime() + 24 * 3600 * 1000) : e.start)
      return e.start < to && effEnd > from
    })
}

/**
 * 2通りの書き方を受け付ける（NDJSONはJSONBinの性質上想定しない。常に1つの正しいJSON値）。
 *   1. 配列                 [ {...}, {...} ]
 *   2. 包んだオブジェクト    { "events": [ {...} ] }（events/items/data/value/values でも可）
 *   3. 1件そのもの           {...}（title か start を持つオブジェクト）
 */
const WRAP_KEYS = ['events', 'items', 'data', 'value', 'values']

function unwrap(json) {
  if (Array.isArray(json)) return json
  if (json && typeof json === 'object') {
    const key = WRAP_KEYS.find((k) => Array.isArray(json[k]))
    if (key) return json[key]
    if ('title' in json || 'subject' in json || 'start' in json) return [json]
    const keys = Object.keys(json).join(', ') || '(空のオブジェクト)'
    throw new Error(
      `JSONBinの中身を読み取れません（${WRAP_KEYS.map((k) => `"${k}"`).join('/')} に配列を入れる形も可）。実際のキー: ${keys}`
    )
  }
  throw new Error('JSONBinの中身を読み取れません')
}

/** allDayは日付だけでもよい（"2026-09-20"）。時刻付きは、タイムゾーンが無ければUTCとして読む
 * （Outlookの生のGraph応答と同じで、"2026-09-16T04:30:00.0000000" にはタイムゾーンが付いていない） */
function parseDate(value, allDay) {
  if (!value) return null
  const raw = String(value)
  if (allDay && !raw.includes('T')) return new Date(raw + 'T00:00:00')
  const hasZone = /(Z|[+-]\d{2}:\d{2})$/.test(raw)
  const d = new Date(hasZone ? raw : raw + 'Z')
  return isNaN(d.getTime()) ? null : d
}

function toEvent(e, i) {
  const allDay = !!(e.allDay || e.isAllDay)
  const start = parseDate(e.start && e.start.dateTime ? e.start.dateTime : e.start || e.date, allDay)
  const end = parseDate(e.end && e.end.dateTime ? e.end.dateTime : e.end, allDay) || (allDay ? null : start)
  return {
    id: e.id || `jsonbin-${i}`,
    title: (e.title || e.subject || '').trim() || '(件名なし)',
    start,
    end,
    allDay,
    cancelled: !!e.cancelled,
    free: !!e.free,
    declined: false,
    location: (e.location && e.location.displayName) || e.location || '',
    organizer: e.organizer || '',
    onlineUrl: e.onlineUrl || '',
    url: e.url || ''
  }
}
