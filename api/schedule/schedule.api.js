/**
 * 今日の予定をまとめて取るAPI。
 *
 * 取得元(プロバイダ)は差し替えられる。いまは Outlook だけだが、providers/ に
 * 同じ形のモジュールを足して LOADERS に並べれば増やせる。アカウントは何個でもよく、
 * 同じ取得元を別アカウントで並べても混ざらない。
 *
 * どの画面からでも同じ形で使えるように、設定（アカウント一覧）は呼び出し側が持つ。
 * ここが覚えるのはサインインの控えだけ。
 *
 * 使い方:
 *   const schedule = await import('/api/schedule/schedule.api.js')
 *   const accounts = schedule.normalizeAccounts([
 *     { id:'work', label:'仕事', provider:'outlook', clientId:'...' },
 *     { id:'home', label:'個人', provider:'outlook', clientId:'...', tenant:'consumers' }
 *   ])
 *   await schedule.signIn(accounts[0])          // ポップアップでサインイン
 *   const { events, accounts: state } = await schedule.today(accounts)
 */

const LOADERS = {
  outlook: () => import('./providers/outlook.js')
}

/** アカウントに色を指定しなかったときの並び順 */
const PALETTE = ['#5FA8A0', '#8F7BD6', '#E0B252', '#D9694F', '#6F9FD8', '#B58AC4']

const loaded = new Map()

function loadProvider(providerId) {
  const key = providerId || 'outlook'
  const load = LOADERS[key]
  if (!load) throw new Error(`知らない取得元です: ${key}`)
  if (!loaded.has(key)) loaded.set(key, load())
  return loaded.get(key)
}

/** 選べる取得元の一覧。設定画面を作るときに使う */
export async function providers() {
  const list = await Promise.all(
    Object.keys(LOADERS).map(async (key) => {
      const p = await loadProvider(key)
      return { id: p.id, label: p.label, fields: p.FIELDS || [] }
    })
  )
  return list
}

/** Azureなどに登録するリダイレクトURI。画面に出して控えてもらう用 */
export async function redirectUri(providerId = 'outlook') {
  const p = await loadProvider(providerId)
  return p.redirectUri ? p.redirectUri() : ''
}

/** 設定に書き漏らした項目を埋める。以降の処理はこの形だけを見る */
export function normalizeAccounts(list) {
  return (Array.isArray(list) ? list : []).map((raw, i) => {
    const a = raw && typeof raw === 'object' ? raw : {}
    const id = String(a.id || `account-${i + 1}`)
    return {
      ...a,
      id,
      provider: a.provider || 'outlook',
      label: a.label || id,
      color: a.color || PALETTE[i % PALETTE.length]
    }
  })
}

export async function signIn(account) {
  const [one] = normalizeAccounts([account])
  const p = await loadProvider(one.provider)
  return p.signIn(one)
}

export async function signOut(account) {
  const [one] = normalizeAccounts([account])
  const p = await loadProvider(one.provider)
  return p.signOut(one)
}

/** アカウントごとのサインイン状態。予定は取りに行かない */
export async function status(accounts) {
  const list = normalizeAccounts(accounts)
  return Promise.all(
    list.map(async (a) => {
      try {
        const p = await loadProvider(a.provider)
        const st = await p.status(a)
        return { id: a.id, label: a.label, color: a.color, provider: a.provider, ...st }
      } catch (err) {
        return {
          id: a.id,
          label: a.label,
          color: a.color,
          provider: a.provider,
          signedIn: false,
          error: err.message || String(err)
        }
      }
    })
  )
}

/** その日の 0:00 から翌 0:00 まで（端末のタイムゾーン） */
export async function today(accounts, options = {}) {
  const base = options.date ? new Date(options.date) : new Date()
  const from = new Date(base.getFullYear(), base.getMonth(), base.getDate())
  const to = new Date(from.getFullYear(), from.getMonth(), from.getDate() + 1)
  return range(accounts, { from, to })
}

/**
 * 期間内の予定を、全アカウントぶんまとめて取る。
 * 1件が落ちても他は返す。落ちた理由は accounts[].error に入る。
 */
export async function range(accounts, { from, to }) {
  const list = normalizeAccounts(accounts)
  const picked = await Promise.all(list.map((a) => collect(a, from, to)))

  const events = picked.flatMap((r) => r.events).sort(byStart)
  return {
    from,
    to,
    dayKey: dayKey(from),
    events,
    accounts: picked.map((r) => r.info),
    signedIn: picked.some((r) => r.info.signedIn),
    errors: picked.filter((r) => r.info.error).map((r) => ({ id: r.info.id, message: r.info.error }))
  }
}

async function collect(account, from, to) {
  const info = {
    id: account.id,
    label: account.label,
    color: account.color,
    provider: account.provider,
    signedIn: false,
    username: '',
    count: 0,
    error: '',
    needsSignIn: false
  }
  try {
    const p = await loadProvider(account.provider)
    const st = await p.status(account)
    info.signedIn = !!st.signedIn
    info.username = st.username || ''
    if (!info.signedIn) {
      info.needsSignIn = true
      if (st.error) info.error = st.error
      return { info, events: [] }
    }
    const raw = await p.events(account, { from, to })
    const events = raw
      .filter((e) => e && e.start)
      .map((e) => ({
        ...e,
        key: `${account.id}:${e.id}`,
        accountId: account.id,
        accountLabel: account.label,
        color: account.color,
        provider: account.provider
      }))
    info.count = events.length
    return { info, events }
  } catch (err) {
    info.error = err.message || String(err)
    info.needsSignIn = !!err.needsSignIn
    return { info, events: [] }
  }
}

/* 終日は先頭に。あとは開始が早い順 */
function byStart(a, b) {
  if (a.allDay !== b.allDay) return a.allDay ? -1 : 1
  return a.start - b.start
}

export function dayKey(date) {
  return new Date(date).toLocaleDateString('sv-SE') // YYYY-MM-DD
}

export function formatTime(date) {
  if (!date) return ''
  return new Date(date).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', hour12: false })
}

/** いま進行中か */
export function ongoing(event, now = new Date()) {
  if (!event || event.allDay || !event.start) return false
  const end = event.end || new Date(event.start.getTime() + 30 * 60 * 1000)
  return event.start <= now && now < end
}

/** これから始まる中で一番近いもの */
export function nextEvent(events, now = new Date()) {
  return (events || []).find((e) => !e.allDay && e.start && e.start > now) || null
}
