/**
 * Outlook（Microsoft Graph）から予定を取る取得元。
 *
 * schedule.api.js から呼ばれる。ここが持つのはMicrosoft固有の事情だけで、
 * 予定の形をそろえるのは schedule.api.js の仕事。
 *
 * アカウントは何個でもよい。アプリ登録(clientId)が同じでも、サインインした
 * Microsoftアカウントごとに別物として扱う。どのカードのアカウント設定が
 * どのMicrosoftアカウントに結びついているかだけ、この場で覚えておく。
 */

/* MSALはESMで配っていないので、必要になった時点で読み込んでグローバルから使う */
const MSAL_SRC = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/lib/msal-browser.min.js'
const GRAPH = 'https://graph.microsoft.com/v1.0'
const SCOPES = ['User.Read', 'Calendars.Read']

/* サインインが終わったポップアップが戻ってくる先。Azureにこれを登録する */
const REDIRECT_URI = new URL('../auth-redirect.html', import.meta.url).href

/** どのMicrosoftアカウントに繋いだかの控え。トークンはMSALが別に持つ */
const LINK_KEY = 'schedule.outlook.links'

export const id = 'outlook'
export const label = 'Outlook（Microsoft 365 / Outlook.com）'

/** 設定画面を作るときに使う。アカウント1件に何を書けばいいか */
export const FIELDS = [
  {
    key: 'clientId',
    label: 'アプリケーション (クライアント) ID',
    required: true,
    placeholder: '00000000-0000-0000-0000-000000000000',
    hint: 'Azure のアプリ登録で発行されるID。2つのアカウントで同じIDを使い回せます'
  },
  {
    key: 'tenant',
    label: 'テナント',
    placeholder: 'common',
    hint: '職場/学校と個人の両方なら common、職場だけなら organizations、個人だけなら consumers'
  }
]

export function redirectUri() {
  return REDIRECT_URI
}

/** サインインし直しが要るエラー。呼び出し側はこの印を見てボタンを出す */
function needSignIn(message) {
  const err = new Error(message)
  err.needsSignIn = true
  return err
}

/* ---------- MSAL ---------- */

let msalLoading = null

function loadMsal() {
  if (msalLoading) return msalLoading
  msalLoading = new Promise((resolve, reject) => {
    if (window.msal) return resolve(window.msal)
    const el = document.createElement('script')
    el.src = MSAL_SRC
    el.async = true
    el.crossOrigin = 'anonymous'
    el.onload = () =>
      window.msal ? resolve(window.msal) : reject(new Error('サインインの部品を読み込めませんでした'))
    el.onerror = () =>
      reject(new Error('サインインの部品を読み込めませんでした（ネットワークを確認してください）'))
    document.head.appendChild(el)
  })
  return msalLoading
}

function authorityOf(account) {
  return 'https://login.microsoftonline.com/' + (account.tenant || 'common')
}

/* アプリ登録ごとに1つ。同じ登録を使うアカウントはこの中に並ぶ */
const instances = new Map()

function instanceFor(account) {
  if (!account.clientId) throw new Error('clientId が設定されていません')
  // 見本のIDのままだとMicrosoft側で AADSTS700038 になり、原因が分かりにくい
  if (/^[0-]+$/.test(String(account.clientId).trim())) {
    throw new Error('clientId が見本のままです。Azure のアプリ登録で発行したIDに置き換えてください')
  }
  const key = account.clientId + '|' + authorityOf(account)
  if (!instances.has(key)) instances.set(key, createInstance(account))
  return instances.get(key)
}

async function createInstance(account) {
  const msal = await loadMsal()
  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: account.clientId,
      authority: authorityOf(account),
      redirectUri: REDIRECT_URI,
      navigateToLoginRequestUrl: false
    },
    // タブを開き直してもサインインしたままにする
    cache: { cacheLocation: 'localStorage', temporaryCacheLocation: 'sessionStorage' }
  })
  await pca.initialize()
  return pca
}

/* ---------- アカウントの結びつけ ---------- */

function links() {
  try {
    return JSON.parse(localStorage.getItem(LINK_KEY)) || {}
  } catch {
    return {}
  }
}

function saveLinks(all) {
  localStorage.setItem(LINK_KEY, JSON.stringify(all))
}

function linkTo(accountId, msalAccount) {
  const all = links()
  all[accountId] = { homeAccountId: msalAccount.homeAccountId, username: msalAccount.username }
  saveLinks(all)
}

function unlink(accountId) {
  const all = links()
  delete all[accountId]
  saveLinks(all)
}

function sameName(a, b) {
  return String(a || '').toLowerCase() === String(b || '').toLowerCase()
}

/** 設定のアカウント1件に対応するMicrosoftアカウントを、MSALの控えから探す */
function msalAccountOf(pca, account) {
  const all = pca.getAllAccounts()
  if (!all.length) return null

  const saved = links()[account.id]
  if (saved) {
    const hit = all.find((a) => a.homeAccountId === saved.homeAccountId)
    if (hit) return hit
    if (saved.username) {
      const byName = all.find((a) => sameName(a.username, saved.username))
      if (byName) return byName
    }
    return null
  }

  if (account.username) return all.find((a) => sameName(a.username, account.username)) || null

  // まだ結びつけていない設定に、まだ誰にも使われていないサインインが1つだけ残っていればそれを使う
  const used = Object.values(links()).map((x) => x.homeAccountId)
  const free = all.filter((a) => !used.includes(a.homeAccountId))
  return free.length === 1 ? free[0] : null
}

/* ---------- 公開する操作 ---------- */

export async function status(account) {
  try {
    const pca = await instanceFor(account)
    const acc = msalAccountOf(pca, account)
    if (!acc) return { signedIn: false }
    return { signedIn: true, username: acc.username, name: acc.name || '' }
  } catch (err) {
    return { signedIn: false, error: err.message || String(err) }
  }
}

export async function signIn(account) {
  const pca = await instanceFor(account)
  // アカウントが複数あるので、毎回どれでサインインするかを選ばせる
  const res = await pca.loginPopup({ scopes: SCOPES, prompt: 'select_account' })
  if (!res || !res.account) throw new Error('サインインできませんでした')

  const taken = Object.entries(links()).find(
    ([key, v]) => key !== account.id && v.homeAccountId === res.account.homeAccountId
  )
  if (taken) throw new Error(`${res.account.username} は「${taken[0]}」で使っています。別のアカウントを選んでください`)

  linkTo(account.id, res.account)
  return { signedIn: true, username: res.account.username, name: res.account.name || '' }
}

export async function signOut(account) {
  const pca = await instanceFor(account)
  const acc = msalAccountOf(pca, account)
  unlink(account.id)
  if (!acc) return { signedIn: false }
  await pca.logoutPopup({ account: acc, mainWindowRedirectUri: REDIRECT_URI })
  return { signedIn: false }
}

async function tokenFor(account) {
  const msal = await loadMsal()
  const pca = await instanceFor(account)
  const acc = msalAccountOf(pca, account)
  if (!acc) throw needSignIn('サインインしてください')
  try {
    const res = await pca.acquireTokenSilent({ scopes: SCOPES, account: acc })
    return res.accessToken
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      throw needSignIn('サインインし直してください')
    }
    throw err
  }
}

/**
 * 期間内の予定。時刻はUTCで受け取って、表示側のタイムゾーンに任せる
 * （Graphのタイムゾーン名の違いを持ち込まないため）
 */
export async function events(account, { from, to }) {
  const token = await tokenFor(account)
  const query = [
    'startDateTime=' + encodeURIComponent(from.toISOString()),
    'endDateTime=' + encodeURIComponent(to.toISOString()),
    '$select=' +
      encodeURIComponent(
        'id,subject,start,end,isAllDay,isCancelled,showAs,location,organizer,onlineMeeting,onlineMeetingUrl,webLink,responseStatus'
      ),
    '$orderby=' + encodeURIComponent('start/dateTime'),
    '$top=50'
  ].join('&')

  let url = `${GRAPH}/me/calendarView?${query}`
  const out = []

  // 1日ぶんなのでまず足りるが、繰り返しの多い日に備えて数ページだけ追う
  for (let page = 0; url && page < 5; page++) {
    const res = await fetch(url, {
      headers: { Authorization: 'Bearer ' + token, Prefer: 'outlook.timezone="UTC"' }
    })
    if (res.status === 401) throw needSignIn('権限が切れました。サインインし直してください')
    if (!res.ok) throw new Error(await errorText(res))
    const json = await res.json()
    ;(json.value || []).forEach((e) => out.push(toEvent(e)))
    url = json['@odata.nextLink'] || ''
  }
  return out
}

async function errorText(res) {
  let detail = ''
  try {
    const json = await res.json()
    detail = (json.error && json.error.message) || ''
  } catch {
    /* 本文がJSONでないことがある */
  }
  if (res.status === 403) return detail || 'カレンダーを読む権限がありません（Calendars.Read）'
  return detail ? `${detail}（HTTP ${res.status}）` : `予定を取得できません（HTTP ${res.status}）`
}

/** Graphは "2026-09-16T01:30:00.0000000" のように時差なしで返すので、UTCとして読む */
function parseTime(t) {
  if (!t || !t.dateTime) return null
  const raw = t.dateTime
  const hasZone = /(Z|[+-]\d{2}:\d{2})$/.test(raw)
  const d = new Date(hasZone ? raw : raw + 'Z')
  return isNaN(d.getTime()) ? null : d
}

function toEvent(e) {
  return {
    id: e.id,
    title: (e.subject || '').trim() || '(件名なし)',
    start: parseTime(e.start),
    end: parseTime(e.end),
    allDay: !!e.isAllDay,
    cancelled: !!e.isCancelled,
    free: e.showAs === 'free',
    declined: e.responseStatus && e.responseStatus.response === 'declined',
    location: (e.location && e.location.displayName) || '',
    organizer: (e.organizer && e.organizer.emailAddress && e.organizer.emailAddress.name) || '',
    onlineUrl: (e.onlineMeeting && e.onlineMeeting.joinUrl) || e.onlineMeetingUrl || '',
    url: e.webLink || ''
  }
}
