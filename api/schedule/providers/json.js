/**
 * JSONで決め打ちした予定を出す取得元。
 *
 * Outlookと違ってサインインが要らない。account.events に配列を直接書くか、
 * account.url を書けば毎回そこから取りに行く（同一オリジンかCORS許可が必要）。
 * どちらも書いてあれば url を優先する。
 */

export const id = 'json'
export const label = 'JSON（固定の予定・共有ファイル）'

export const FIELDS = [
  {
    key: 'url',
    label: 'JSONのURL（任意）',
    placeholder: 'https://example.com/events.json',
    hint: '空欄なら account.events をそのまま使います。書くと毎回そこから取得します（同一オリジンかCORSが必要）'
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

export async function events(account, { from, to }) {
  const raw = account.url ? await fetchList(account.url) : Array.isArray(account.events) ? account.events : []
  return raw
    .map(toEvent)
    .filter((e) => e.start)
    .filter((e) => {
      // allDayは終了時刻を持たないことがあるので、翌日0時を仮の終わりとして扱う
      const effEnd = e.end || (e.allDay ? new Date(e.start.getTime() + 24 * 3600 * 1000) : e.start)
      return e.start < to && effEnd > from
    })
}

async function fetchList(url) {
  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`予定のJSONを取得できません（HTTP ${res.status}）`)
  return parseFeed(await res.text())
}

/**
 * 3通りの書き方を受け付ける。
 *   1. 配列                 [ {...}, {...} ]
 *   2. 包んだオブジェクト    { "events": [ {...} ] }（events/items/data/value/values）
 *   3. 1件そのもの           {...}（title か start を持つオブジェクト）
 *   4. 1行に1件のNDJSON      {...}\n{...}\n…（GASなどが追記していく運用を想定）
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
        throw new Error(`予定のJSON ${i + 1}行目を読み取れません: ${line.slice(0, 60)}`)
      }
    })
}

class FeedShapeError extends Error {}

/** 配列・包まれたオブジェクト・単体の予定オブジェクトのいずれかを配列に揃える */
const WRAP_KEYS = ['events', 'items', 'data', 'value', 'values']

function unwrap(json) {
  if (Array.isArray(json)) return json
  if (json && typeof json === 'object') {
    const key = WRAP_KEYS.find((k) => Array.isArray(json[k]))
    if (key) return json[key]
    if ('title' in json || 'subject' in json || 'start' in json) return [json]
    const keys = Object.keys(json).join(', ') || '(空のオブジェクト)'
    throw new FeedShapeError(
      `予定のJSONを読み取れません（${WRAP_KEYS.map((k) => `"${k}"`).join('/')} に配列を入れる形も可）。実際のキー: ${keys}`
    )
  }
  throw new FeedShapeError('予定のJSONを読み取れません')
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
    id: e.id || `json-${i}`,
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
