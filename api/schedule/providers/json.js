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
  const json = await res.json()
  if (!Array.isArray(json)) throw new Error('予定のJSONは配列で書いてください')
  return json
}

/** allDayは日付だけでもよい（"2026-09-20"）。時刻付きはそのままDateに渡す */
function parseDate(value, allDay) {
  if (!value) return null
  const d = new Date(allDay && !String(value).includes('T') ? value + 'T00:00:00' : value)
  return isNaN(d.getTime()) ? null : d
}

function toEvent(e, i) {
  const allDay = !!e.allDay
  const start = parseDate(e.start || e.date, allDay)
  const end = parseDate(e.end, allDay) || (allDay ? null : start)
  return {
    id: e.id || `json-${i}`,
    title: (e.title || '').trim() || '(件名なし)',
    start,
    end,
    allDay,
    cancelled: !!e.cancelled,
    free: !!e.free,
    declined: false,
    location: e.location || '',
    organizer: e.organizer || '',
    onlineUrl: e.onlineUrl || '',
    url: e.url || ''
  }
}
