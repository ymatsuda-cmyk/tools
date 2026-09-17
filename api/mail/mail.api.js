/**
 * 受信トレイ（トップフォルダ）のメールをまとめて取るAPI。api/schedule/schedule.api.js と同じ作り。
 *
 * 取得元(プロバイダ)は差し替えられる。いまは Outlook と、サインイン不要のJSONが入っている。
 * アカウントは何個でもよく、同じ取得元を別アカウントで並べても混ざらない。
 *
 * 使い方:
 *   const mail = await import('/api/mail/mail.api.js')
 *   const accounts = mail.normalizeAccounts([
 *     { id:'work', label:'仕事', provider:'outlook', clientId:'...' },
 *     { id:'home', label:'個人', provider:'outlook', clientId:'...', tenant:'consumers' },
 *     { id:'plan', label:'共有箱', provider:'json', maxAgeDays:3,
 *       mails:[{ subject:'お知らせ', from:'info@example.com', receivedAt:'2026-09-16T10:00:00+09:00' }] }
 *   ])
 *   await mail.signIn(accounts[0])                          // ポップアップでサインイン
 *   const { mails, accounts: state } = await mail.recent(accounts, { maxAgeDays: 7 })
 */

const LOADERS = {
  outlook: () => import('./providers/outlook.js'),
  json: () => import('./providers/json.js')
}

/** アカウントに色を指定しなかったときの並び順 */
const PALETTE = ['#5FA8A0', '#8F7BD6', '#E0B252', '#D9694F', '#6F9FD8', '#B58AC4']

/** カードやアカウントに maxAgeDays が無いときの既定値 */
const DEFAULT_MAX_AGE_DAYS = 7

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

/** アカウントごとのサインイン状態。メールは取りに行かない */
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

/**
 * 期間内（既定 options.maxAgeDays 日、アカウント側の maxAgeDays があればそちらを優先）の
 * 受信トレイのメールを、全アカウントぶんまとめて取る。1件が落ちても他は返す。
 */
export async function recent(accounts, options = {}) {
  const list = normalizeAccounts(accounts)
  const defaultMaxAge = Math.max(1, Number(options.maxAgeDays) || DEFAULT_MAX_AGE_DAYS)
  const picked = await Promise.all(list.map((a) => collect(a, defaultMaxAge)))

  const mails = picked.flatMap((r) => r.mails).sort(byReceivedDesc)
  return {
    mails,
    accounts: picked.map((r) => r.info),
    signedIn: picked.some((r) => r.info.signedIn),
    errors: picked.filter((r) => r.info.error).map((r) => ({ id: r.info.id, message: r.info.error }))
  }
}

async function collect(account, defaultMaxAge) {
  const maxAgeDays = Math.max(1, Number(account.maxAgeDays) || defaultMaxAge)
  const since = new Date(Date.now() - maxAgeDays * 24 * 3600 * 1000)
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
      return { info, mails: [] }
    }
    const raw = await p.messages(account, { since })
    const mails = raw
      .filter((e) => e && e.receivedAt && e.receivedAt >= since)
      .map((e) => ({
        ...e,
        key: `${account.id}:${e.id}`,
        accountId: account.id,
        accountLabel: account.label,
        color: account.color,
        provider: account.provider
      }))
    info.count = mails.length
    return { info, mails }
  } catch (err) {
    info.error = err.message || String(err)
    info.needsSignIn = !!err.needsSignIn
    return { info, mails: [] }
  }
}

function byReceivedDesc(a, b) {
  return b.receivedAt - a.receivedAt
}

export function formatTime(date) {
  if (!date) return ''
  return new Date(date).toLocaleString('ja-JP', {
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false
  })
}
