// src/app.js — MinuteStock（議事録 v2。ClipStock と同じ作り：暗号化して GitHub private に保管、GAS 経由で書く）
import { GAS_URL, POLL_MS, CLAUDE_DEFAULT_URL, WORKER_EVERY_SEC, MARK_TAGS } from './config.js'
import { openBundle, importPriv, newDek, sealJson, openJson, wrapDek, unwrapDek, rnd } from './crypto.js'
import { gas, repo } from './store.js'
import { bioOk, bioRegistered, bioRegister, bioUnlock, bioForget } from './bio.js'
import { autoEnabled, setAutoEnabled, autoSaved, autoSave, autoLoad, autoForget } from './autoopen.js'
import { fmtT, parseT, normalizeSummary, segmentsFromText, summarizeWithAI, transcriptText } from './summary.js'
import { parseSummaryText, summaryTypeLabel } from './summary-text.js'
import { loadSettings, saveSettings as saveSettingsLocal, newConnection, activeConnection, activeModelName, allModels, connectionOf } from './llm-settings.js'
import { streamChat } from './llm-client.js'
import { groupMarks, marksOf, validFor, renderMarked, applyRange, recolor, reconcile, summaryTexts, trTexts, withTimes, locCompare, locLabel, locTab, kindOf, SUMMARY_KINDS } from './marks.js'
import { todoKey, doneOf, logsOf, mergeTodos } from './todos.js'
import { draftFromSummary, editorHTML, readDraft, summaryFromDraft, blank, trToText, textToSegments } from './edit.js'
import { initScrollHide } from './scrollhide.js'
import { mindmapMarkdown } from './mindmap.js'
import { renderMindmap, isLegacyHtml } from '../../api/mindmap2/mindmap2.js'

/* ---------- 小物 ---------- */
const $ = (s, el = document) => el.querySelector(s)
const $$ = (s, el = document) => [...el.querySelectorAll(s)]
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v) } catch { return d } },
  set(k, v) { localStorage.setItem(k, JSON.stringify(v)); if (SYNC_KEYS.has(k)) schedulePush() },
  del(k) { localStorage.removeItem(k); if (SYNC_KEYS.has(k)) schedulePush() },
}
const pad = (n) => String(n).padStart(2, '0')
const WD = ['日', '月', '火', '水', '木', '金', '土']
const monthKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`
const todayKey = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` }
const toDate = (s) => new Date(String(s || '').length >= 16 ? s : (s || '1970-01-01') + 'T00:00')
const md = (s) => { const d = toDate(s); return `${d.getMonth() + 1}/${d.getDate()}` }
const hm = (s) => String(s || '').slice(11, 16)
const fmtDur = (sec) => (sec ? fmtT(sec) : '—')
const nowLocal = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}` }
const newId = () => 'm_' + [...rnd(12)].map((b) => 'abcdefghijklmnopqrstuvwxyz0123456789'[b % 36]).join('')
const ago = (ms) => { const m = Math.round((Date.now() - ms) / 60000); return m < 1 ? 'たった今' : m < 60 ? `${m}分前` : m < 1440 ? `${Math.round(m / 60)}時間前` : `${Math.round(m / 1440)}日前` }

const ADMIN_ROLE = 'xYz'
const SRC = { audio: '音声', text: 'テキスト', plaud: 'PLAUD' }
const ST = {
  tw: { l: '文字起こし待ち', c: 'st-tw' }, tr: { l: '文字起こし中', c: 'st-tr' },
  sw: { l: '要約待ち', c: 'st-sw' }, ok: { l: '要約済み', c: 'st-ok' }, err: { l: '失敗', c: 'st-err' },
}
const USERS_SVG = '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3.2"/><path d="M3.5 19a5.5 5.5 0 0 1 11 0"/><circle cx="17" cy="9" r="2.6"/><path d="M15.5 14.2A4.6 4.6 0 0 1 21 19"/></svg>'
const LOCK_SVG = '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>'

/* ---------- 状態 ---------- */
const S = {
  code: null, role: null, admin: false, gid: null, adminGid: null, roles: {}, privs: {}, gh: null, R: null,
  session: null, exp: 0, pub: {}, raw: [], items: [], byId: new Map(), worker: null, sig: '', syncedAt: 0, syncErr: '',
  tab: 'list', view: ls.get('m2.view', 'timeline'), month: monthKey(new Date()), filter: 'all', favOnly: false, scope: 'all', tagOpen: null, q: '', sel: null, dtab: 'sum',
  cfilter: 'all', checked: new Set(), busy: {}, chat: {}, trq: '', seek: null, cnote: '',
  range: ls.get('m2.range', '3m'), tfilter: 'open', towner: '', nfilter: 0, nq: '', logOpen: null, mkMemo: null,
  edit: null, jump: null, wq: {}, sec: 'summary', secJump: null, spyLock: 0,
  files: {}, fav: new Set(ls.get('m2.fav', [])), seen: ls.get('m2.seen', {}), timer: 0, prefsReady: false,
}
const DEK = new Map()
const META = new Map()

/* ---------- 文字の大きさ（この端末に保存。次回も同じ大きさで開く） ---------- */
const FS_LEVELS = [{ n: '小', v: 0.88 }, { n: '標準', v: 1 }, { n: '大', v: 1.15 }, { n: '特大', v: 1.3 }, { n: '最大', v: 1.5 }]
{
  const saved = Number(ls.get('m2.fs', 1))
  S.fs = Number.isInteger(saved) && saved >= 0 && saved < FS_LEVELS.length ? saved : 1
}
const applyFs = () => document.documentElement.style.setProperty('--rs', String(FS_LEVELS[S.fs].v))

function fsCtl() {
  return `<span class="fsctl" role="group" aria-label="文字の大きさ"><button type="button" class="btn sm" data-act="fs" data-d="-1" aria-label="文字を小さく" ${S.fs <= 0 ? 'disabled' : ''}>A−</button><span class="fsl">${FS_LEVELS[S.fs].n}</span><button type="button" class="btn sm" data-act="fs" data-d="1" aria-label="文字を大きく" ${S.fs >= FS_LEVELS.length - 1 ? 'disabled' : ''}>A＋</button></span>`
}

function setFs(i) {
  S.fs = Math.max(0, Math.min(FS_LEVELS.length - 1, i))
  ls.set('m2.fs', S.fs)
  applyFs()
  captureEdit()   // 編集中の入力を失わない
  if (S.tab === 'list') paintDetail()
  else if (S.tab === 'todo' || S.tab === 'notes') renderMain()
  $$('#fsset [data-v]').forEach((b) => { b.classList.toggle('main', Number(b.dataset.v) === S.fs); b.setAttribute('aria-pressed', Number(b.dataset.v) === S.fs) })
}

/* ---------- ログイン ---------- */
// 開き方：auto（この端末の控え）／typed（入力）／bio（Face ID）／recv（他の端末から受け取り）／change（設定で変更）
const DL_APP = 'minutestock'
const DL = () => window.DeviceLink
// index.html の <script> が古いキャッシュで読まれていないときに備えて、ここからも読み込む
let dlLoading = null
function ensureDL() {
  if (window.DeviceLink) return Promise.resolve(window.DeviceLink)
  if (!dlLoading) {
    dlLoading = new Promise((resolve) => {
      const el = document.createElement('script')
      el.src = new URL('../common/devicelink.js', location.href).href
      el.onload = () => resolve(window.DeviceLink || null)
      el.onerror = () => { dlLoading = null; resolve(null) }
      document.head.append(el)
    })
  }
  return dlLoading
}
const DL_MISSING = '受け渡しの部品（common/devicelink.js）を読み込めませんでした。オンラインで開き直してください。'
const LOCKED = 'm2.locked'   // 鍵をかけたあとは、このタブを閉じるまで自動で開かない

function gateShell(inner) {
  clearInterval(S.timer)
  $('#app').hidden = true
  $('#gate').hidden = false
  $('#gate').innerHTML = `<div class="gate-card"><h1 class="gate-title">MinuteStock</h1>${inner}</div>`
}

function showGate(msg = '') {
  const bio = bioOk() && bioRegistered()
  const saved = autoEnabled() && autoSaved()
  gateShell(`
      <p class="gate-sub">暗号化して保管しています。開くには権限コードが必要です。</p>
      ${saved ? '<button class="btn main wide" data-act="auto-open">開く</button><div class="gate-or">または</div>' : ''}
      ${bio ? `<button class="btn ${saved ? '' : 'main'} wide gate-alt" data-act="bio">Face ID・Touch ID で開く</button><div class="gate-or">または</div>` : ''}
      <form id="loginForm" class="gate-form">
        <label for="code">権限コード</label>
        <input id="code" type="password" autocomplete="current-password" required>
        <button class="btn ${bio || saved ? '' : 'main'} wide" type="submit">${bio || saved ? '権限コードで開く' : '開く'}</button>
      </form>
      <button class="link gate-link" data-act="recv">コードで受け取る（別の端末から）</button>
      <p class="gate-msg" id="gateMsg" role="alert">${esc(msg)}</p>
      ${bio || saved || ls.get('m2.seen', null) ? '<button class="link danger" data-act="reset">この端末の設定を消して、最初から設定する</button>' : ''}`)
  $('#loginForm').addEventListener('submit', async (e) => {
    e.preventDefault()
    const code = $('#code').value.trim()
    if (code) await doLogin(code, 'typed')
  })
  if (!bio && !saved) $('#code')?.focus()
}

/** 他の端末から受け取る。pkg があれば（QR・リンク、またはコードで受け取り済み）パスフレーズだけ聞く */
function showReceive(pkg = null, msg = '') {
  gateShell(`
      <p class="gate-sub">${pkg ? '別の端末から設定を受け取りました。渡した端末に表示されたパスフレーズを入れてください。'
        : '渡す端末の 設定 →「他端末へ設定共有」→「コードを発行」で表示された6桁のコードと、パスフレーズを入れてください。'}</p>
      <form id="recvForm" class="gate-form">
        ${pkg ? '' : '<label for="dlCode">コード（6桁）</label><input id="dlCode" inputmode="numeric" autocomplete="one-time-code" maxlength="7" required>'}
        <label for="dlPass">パスフレーズ</label>
        <input id="dlPass" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="例：ABCDE-FGHJK" required>
        <button class="btn main wide" type="submit">受け取って開く</button>
      </form>
      <p class="gate-msg" id="gateMsg" role="alert">${esc(msg)}</p>
      <button class="link gate-link" data-act="gate">権限コードで開く</button>`)
  $(pkg ? '#dlPass' : '#dlCode')?.focus()
  $('#recvForm').addEventListener('submit', async (e) => {
    e.preventDefault()
    const m = $('#gateMsg')
    const dl = await ensureDL()
    if (!dl) { m.textContent = DL_MISSING; return }
    $$('#gate button').forEach((b) => { b.disabled = true })
    let got = pkg
    try {
      m.textContent = '受け取っています…'
      if (!got) got = await dl.receiveCode(DL_APP, $('#dlCode').value)
      let data
      try { data = await dl.open(got, normPass($('#dlPass').value)) } catch (err) {
        // コードは1回きりなので、受け取った包みを持ったまま、パスフレーズだけ聞き直す
        if (err.message === 'BADPASS') { showReceive(got, 'パスフレーズが違います。渡した端末に表示されているものを入れてください。'); return }
        throw err
      }
      if (!data || typeof data.code !== 'string' || !data.code) throw new Error('受け取った内容に権限コードがありません。')
      const aiMsg = takeAi(data.ai)
      if (await doLogin(data.code, 'recv') && aiMsg) setTimeout(() => toast(aiMsg, false, 6000), 4200)
    } catch (err) {
      showReceive(got, err.message || '受け取れませんでした。')
    }
  })
}

async function doLogin(code, how = 'typed') {
  const m = $('#gateMsg')
  if (m) m.textContent = '確認しています…'
  $$('#gate button').forEach((b) => { b.disabled = true })
  try {
    if (GAS_URL.includes('ここに')) throw new Error('src/config.js の GAS_URL が未設定です。')
    await login(code)
    sessionStorage.removeItem(LOCKED)
    if (how !== 'auto' && how !== 'bio' && autoEnabled()) await autoSave(code).catch(() => {})
    await start()
    if (how === 'typed' && !autoEnabled() && bioOk() && !bioRegistered() && !ls.get('m2.bioAsked', false)) askBio()
    if (how === 'recv') toast(`この端末で使えるようにしました（グループ：${S.admin ? '管理者' : S.role}）。`)
    return true
  } catch (e) {
    S.code = null
    let msg = e.message || '開けませんでした。'
    if (how === 'auto' && /権限コードが違います/.test(msg)) {
      await autoForget()
      msg = 'この端末に覚えていた権限コードが使えなくなりました（変更された可能性があります）。新しい権限コードで開くか、別の端末から受け取ってください。'
    }
    showGate(msg)
    return false
  }
}

async function login(code) {
  const d = await gas('login', { code })
  let b
  try { b = await openBundle(code, d.bundle) } catch { throw new Error('鍵束を開けませんでした（管理者に連絡してください）') }
  S.code = code
  S.role = d.role; S.admin = !!d.admin; S.gid = d.gid; S.adminGid = d.adminGid
  S.session = d.session; S.exp = d.exp
  S.gh = b.gh; S.R = repo(b.gh)
  S.roles = b.roles || { [d.role]: d.gid }
  S.privs = {}
  for (const [g, jwk] of Object.entries(b.keys || {})) S.privs[g] = await importPriv(jwk)
}

async function ensureSession() {
  if (Date.now() < S.exp - 120000) return
  const d = await gas('login', { code: S.code })
  S.session = d.session; S.exp = d.exp
}

async function gasS(action, body) {
  await ensureSession()
  return gas(action, { session: S.session, ...body })
}

/** 読み込んだ議事録・鍵を手放す（画面はそのまま） */
function clearData() {
  clearInterval(S.timer)
  clearTimeout(prefsTimer)
  hideBar()
  Object.assign(S, { code: null, role: null, admin: false, gid: null, adminGid: null, roles: {}, privs: {}, gh: null, R: null, session: null, exp: 0, pub: {}, raw: [], items: [], files: {}, chat: {}, sel: null, sig: '', edit: null, prefsReady: false })
  S.byId = new Map(); S.checked = new Set(); DEK.clear(); META.clear()
  closeModal()
}

function lock() {
  sessionStorage.setItem(LOCKED, '1')
  clearData()
  showGate()
}

function askBio() {
  ls.set('m2.bioAsked', true)
  modal('Face ID・Touch ID', `
    <p>次から、権限コードを入れずに Face ID・Touch ID で開けるようにしますか？ 権限コードはこの端末の中だけに、生体認証で包んで置きます。</p>
    <div class="acts"><button class="btn main" data-act="bioReg">登録する</button><button class="btn" data-act="close">今はしない</button></div>
    <p class="msg" id="bioMsg"></p>`)
}

/** 起動：リンク・QR で開かれたら受け取り、控えがあれば自動で開く */
async function boot() {
  let pkg = null
  if (/^#dl=/.test(location.hash)) await ensureDL()
  try { pkg = DL()?.fromURL(DL_APP) || null } catch (e) {
    showGate(e.message === 'DeviceLink の形式ではありません' ? 'リンクを読み取れませんでした。渡す端末でもう一度コピーし直してください。' : e.message)
    return
  }
  if (pkg) { showReceive(pkg); return }
  if (autoEnabled() && autoSaved() && !sessionStorage.getItem(LOCKED)) {
    gateShell('<p class="gate-sub">開いています…</p><p class="gate-msg" id="gateMsg" role="status"></p>')
    const code = await autoLoad()
    if (code) { await doLogin(code, 'auto'); return }
  }
  showGate()
}

/* ---------- 設定の保管（GitHub private の settings/<グループ>.json。暗号化して GAS 経由で書く） ----------
   中身：Claude.ai のページ・文字の大きさ・一覧の表示・ToDo の期間・お気に入り
   AI の接続先（URL・API キー）は個人の設定なので入れない。この端末だけに置き、他の端末へは「他端末へ設定共有」で一緒に渡す
   会議と同じく「使い捨ての鍵（DEK）で暗号化 → グループの公開鍵で包む」。開けるのは同じグループの端末だけ */
const SYNC_KEYS = new Set(['m2.fs', 'm2.view', 'm2.range', 'm2.fav', 'm2.claudeUrl'])
let prefsTimer = 0
const prefsPath = () => `settings/${S.gid}.json`
const prefsUpdOf = () => (ls.get('m2.prefsUpd', {}) || {})[S.gid] || 0
function setPrefsUpd(v) { const m = ls.get('m2.prefsUpd', {}) || {}; m[S.gid] = v; localStorage.setItem('m2.prefsUpd', JSON.stringify(m)) }

function localPrefs() {
  return { claudeUrl: ls.get('m2.claudeUrl', ''), fs: S.fs, view: S.view, range: S.range, fav: [...S.fav] }
}

function applyPrefs(p) {
  const raw = (k, v) => (v == null || v === '' ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)))   // 送り返さないよう ls.set は使わない
  raw('m2.claudeUrl', p.claudeUrl || '')
  if (Number.isInteger(p.fs) && p.fs >= 0 && p.fs < FS_LEVELS.length) { S.fs = p.fs; raw('m2.fs', p.fs); applyFs() }
  if (p.view === 'timeline' || p.view === 'cards') { S.view = p.view; raw('m2.view', p.view) }
  if (typeof p.range === 'string' && p.range) { S.range = p.range; raw('m2.range', p.range) }
  if (Array.isArray(p.fav)) { S.fav = new Set(p.fav); raw('m2.fav', p.fav) }
}

/** ログイン直後：GitHub の設定が新しければ取り込み、手元が新しければ送る */
async function pullPrefs() {
  try {
    const f = await S.R.json(prefsPath())
    const local = prefsUpdOf()
    if (f && f.upd > local && f.k && f.d && S.privs[S.gid]) {
      const p = await openJson(await unwrapDek(f.k, S.privs[S.gid]), f.d)
      applyPrefs(p)
      setPrefsUpd(f.upd)
      render()
      if (p.ai) { S.prefsReady = true; await pushPrefs() }   // 以前の形式で入っていた AI の接続先を、GitHub から消す
    } else if (!f || local > f.upd) {
      S.prefsReady = true
      await pushPrefs()
    }
  } catch (e) {
    console.warn('設定の読み込みに失敗:', e)
  } finally {
    S.prefsReady = true
  }
}

function schedulePush() {
  if (!S.R || !S.prefsReady) return
  clearTimeout(prefsTimer)
  prefsTimer = setTimeout(() => { pushPrefs().catch((e) => toast('設定を GitHub に保存できませんでした：' + e.message, true)) }, 1500)
}

async function pushPrefs() {
  if (!S.R || !S.prefsReady) return
  const pub = S.pub[S.gid]
  if (!pub) throw new Error('グループの公開鍵が見つかりません')
  const upd = Date.now()
  const dek = newDek()
  const file = { v: 1, upd, k: await wrapDek(dek, pub), d: await sealJson(dek, localPrefs()) }
  await gasS('putSettings', { file })
  setPrefsUpd(upd)
}

// AI の接続先は個人の設定：この端末にだけ保存し、GitHub（グループ共有）には送らない
const saveSettings = (s) => saveSettingsLocal(s)

/** 他端末へ渡す AI の接続先（空の接続は省く。QR に収まるよう小さくする） */
function aiForShare() {
  const s = loadSettings()
  const connections = s.connections.filter((c) => c.baseUrl || c.apiKey || (c.models || []).length)
    .map(({ id, label, baseUrl, apiKey, numCtx, models, think }) => ({ id, label, baseUrl, apiKey, numCtx, models, think }))
  if (!connections.length) return null
  return { connections, activeConnectionId: s.activeConnectionId, activeModel: s.activeModel, temperature: s.temperature, systemPrompt: s.systemPrompt }
}

/** 受け取った AI の接続先を入れる。この端末にすでに接続先があれば上書きしない */
function takeAi(ai) {
  if (!ai || !Array.isArray(ai.connections) || !ai.connections.length) return ''
  const mine = loadSettings()
  if (mine.connections.some((c) => c.baseUrl)) return 'この端末の AI の接続先はそのままにしました。'
  saveSettingsLocal({ ...mine, ...ai, connections: ai.connections.map((c) => newConnection(c)) })
  return 'AI の接続先も受け取りました。'
}

/* ---------- 他端末へ設定共有（common/devicelink.js） ----------
   渡すのは権限コードだけ。その場で作るパスフレーズ（10文字）で暗号化する。設定は GitHub から自動で入る */
const PASS_ABC = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
const newPass = () => { const s = [...rnd(10)].map((x) => PASS_ABC[x % 32]).join(''); return `${s.slice(0, 5)}-${s.slice(5)}` }
const normPass = (v) => { const s = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); return s.length === 10 ? `${s.slice(0, 5)}-${s.slice(5)}` : String(v || '').trim() }
let dlPanel = null
let dlShare = null   // { pass, pkg } 設定画面を開いている間だけ同じパスフレーズを使う

async function mountShare() {
  dlPanel?.destroy(); dlPanel = null
  let el = $('#devLink')
  if (!el) return
  const dl = await ensureDL()
  el = $('#devLink')   // 読み込みを待つ間に設定を閉じたり開き直したりしたとき
  if (!el || el.dataset.mounted) return
  if (!dl) { el.innerHTML = `<p class="msg">${esc(DL_MISSING)}</p>`; return }
  el.dataset.mounted = '1'
  dlShare = { pass: newPass(), pkg: null }
  dlPanel = dl.panel(el, {
    app: DL_APP,
    getPackage: async () => {
      if (!S.code) throw new Error('ログインし直してください。')
      if (!dlShare.pkg) {
        const ai = aiForShare()
        dlShare.pkg = await dl.seal({ v: 1, code: S.code, ai }, dlShare.pass, { app: DL_APP, label: 'MinuteStock' })
        dlShare.noAi = false
        // QR・中継に収まらないほど大きいときは、権限コードだけ渡す（AI の接続先は受け取った端末で設定）
        if (ai && dlShare.pkg.length > 1800) {
          dlShare.pkg = await dl.seal({ v: 1, code: S.code }, dlShare.pass, { app: DL_APP, label: 'MinuteStock' })
          dlShare.noAi = true
        }
      }
      const p = $('#dlPassBox')
      if (p) {
        p.hidden = false; $('#dlPassVal').textContent = dlShare.pass
        $('#dlAiNote').textContent = dlShare.noAi ? 'AI の接続先は大きすぎて入れられませんでした。受け取った端末の設定で入れてください。' : ''
      }
      return dlShare.pkg
    },
  })
}

/* ---------- 読み込み ---------- */
async function start() {
  $('#gate').hidden = true
  $('#app').hidden = false
  await refresh(true)
  await pullPrefs()
  clearInterval(S.timer)
  S.timer = setInterval(() => { if (!document.hidden) refresh() }, POLL_MS)
}

async function refresh(first = false) {
  if (!S.R) return
  try {
    const [idx, pub, wk] = await Promise.all([
      S.R.json('index.json'),
      first || !Object.keys(S.pub).length ? S.R.json('pub.json') : Promise.resolve(null),
      S.R.json('worker.json').catch(() => null),
    ])
    if (pub) S.pub = pub
    S.worker = wk
    const raw = Array.isArray(idx) ? idx : []
    const sig = JSON.stringify(raw.map((e) => [e.id, e.upd, e.st]))
    S.syncErr = ''
    S.syncedAt = Date.now()
    if (first || sig !== S.sig) {
      S.sig = sig
      S.raw = raw
      await decryptAll()
      render()
    } else {
      paintSync()
      if (S.tab === 'convert') renderMain()
    }
  } catch (e) {
    S.syncErr = e.message
    if (first) render()
    paintSync()
  }
}

async function decryptAll() {
  const out = []
  for (const e of S.raw) {
    const gid = Object.keys(e.k || {}).find((g) => S.privs[g])
    const base = { ...e, day: String(e.date || '').slice(0, 10) }
    if (!gid) { out.push({ ...base, locked: true, title: '非公開の会議', tags: [] }); continue }
    try {
      let dek = DEK.get(e.id)
      if (!dek) { dek = await unwrapDek(e.k[gid], S.privs[gid]); DEK.set(e.id, dek) }
      let meta = META.get(e.id)
      if (!meta || meta.iv !== e.m?.iv) { meta = { iv: e.m.iv, v: await openJson(dek, e.m) }; META.set(e.id, meta) }
      out.push({ ...base, locked: false, title: meta.v.title || '（無題）', tags: meta.v.tags || [], file: meta.v.file || null })
    } catch {
      out.push({ ...base, locked: true, title: '開けない議事録', tags: [] })
    }
  }
  out.sort((a, b) => (b.date || '').localeCompare(a.date || ''))
  S.items = out
  S.byId = new Map(out.map((x) => [x.id, x]))
  if (S.sel && !S.byId.has(S.sel)) S.sel = null
}

/** 会議のファイル（t: 文字起こし, s: 要約, n: メモ）。復号して、更新されるまでは使い回す */
function cached(it, name) {
  if (name === 'n' && !it.nt) return null   // メモ・マーカー・経過・完了のある会議だけ n.json がある
  const c = S.files[it.id]
  return c && c.upd === it.upd && name in c ? c[name] : undefined
}

function file(it, name) {
  if (it.locked) return Promise.resolve(null)
  if (name === 'n' && !it.nt) return Promise.resolve(null)
  let c = S.files[it.id]
  if (!c || c.upd !== it.upd) c = S.files[it.id] = { upd: it.upd, p: {} }
  if (name in c) return Promise.resolve(c[name])
  if (c.p[name]) return c.p[name]
  c.p[name] = (async () => {
    try {
      const v = await S.R.json(`m/${it.id}/${name}.json`)
      const o = v ? await openJson(DEK.get(it.id), v) : null
      if (S.files[it.id] === c) c[name] = o
      return o
    } finally { delete c.p[name] }
  })()
  return c.p[name]
}

/** 書き込みの結果を一覧に反映する（GitHub を読み直さずにすぐ表示を変える） */
async function applyEntry(entry, plain = {}, deleted = false, silent = false) {
  const i = S.raw.findIndex((x) => x.id === (entry?.id || plain.id))
  if (deleted) { if (i >= 0) S.raw.splice(i, 1) } else if (i >= 0) S.raw[i] = entry; else S.raw.push(entry)
  if (!deleted) {
    const prev = S.files[entry.id] || {}
    const keep = {}
    for (const k of ['t', 's', 'n']) if (k in prev && !(k in plain)) keep[k] = prev[k]
    S.files[entry.id] = { upd: entry.upd, p: {}, ...keep, ...plain }
  }
  if (!deleted && entry) { S.seen[entry.id] = entry.upd; ls.set('m2.seen', S.seen) }   // 自分で書いたものは「新着」にしない
  S.sig = JSON.stringify(S.raw.map((e) => [e.id, e.upd, e.st]))
  await decryptAll()
  if (!silent) render()
}

/* ---------- 表示：共通 ---------- */
function render(reset = false, noSheet = false) {
  paintHeader()
  renderMain(reset, noSheet)
  paintSync()
}

function paintSync() {
  const el = $('#sync')
  if (!el) return
  el.textContent = S.syncErr ? `同期できません：${S.syncErr}` : S.syncedAt ? `同期済み・${ago(S.syncedAt)}` : ''
  el.classList.toggle('bad', !!S.syncErr)
}

function procCount() {
  return S.items.filter((x) => !x.locked && ['tw', 'tr', 'sw', 'err'].includes(x.st)).length
}

function paintHeader() {
  const n = procCount()
  $$('[data-tab]').forEach((b) => {
    b.classList.toggle('on', b.dataset.tab === S.tab)
    b.setAttribute('aria-current', b.dataset.tab === S.tab ? 'page' : 'false')
    const badge = b.querySelector('.badge')
    if (badge) { badge.textContent = n; badge.hidden = !n }
  })
  $('#searchBox').hidden = S.tab !== 'list'
  $('#role').textContent = S.admin ? '管理者' : (S.role || '')
}

/** 通信の安全設定（CSP）で style 属性は使えないので、棒の長さ・高さは data-w / data-h に持たせて、描いたあとに当てる */
function applyDyn(root = document) {
  root.querySelectorAll('[data-w]').forEach((el) => { el.style.width = el.dataset.w + 'px' })
  root.querySelectorAll('[data-h]').forEach((el) => { el.style.height = el.dataset.h + 'px' })
}

function renderMain(reset = false, noSheet = false) {
  const m = $('#main')
  const keep = reset ? null : { main: m.scrollTop, list: $('.listpane')?.scrollTop, sheet: $('.sheet-scroll')?.scrollTop }
  if (S.tab === 'list') m.innerHTML = listHTML()
  else if (S.tab === 'todo') m.innerHTML = todoPageHTML()
  else if (S.tab === 'notes') m.innerHTML = notesPageHTML()
  else m.innerHTML = convertHTML()
  if (keep) {
    m.scrollTop = keep.main || 0
    const lp = $('.listpane'); if (lp && keep.list) lp.scrollTop = keep.list
    const sh = $('.sheet-scroll'); if (sh && keep.sheet && !noSheet) sh.scrollTop = keep.sheet
  }
  applyDyn(m)
  if (S.tab === 'list') afterList()
  if (S.tab === 'list' && S.sel && S.dtab === 'map') mountMap(S.byId.get(S.sel))
  if (S.tab === 'todo') kickTodo()
  if (S.tab === 'notes') kickNotes()
}

const stChip = (it) => {
  if (S.busy[it.id]) return `<span class="chip st-busy" data-busy="${it.id}">${esc(S.busy[it.id])}</span>`
  const s = ST[it.st] || ST.tw
  return `<span class="chip ${s.c}">${s.l}</span>`
}
const isNew = (it) => !it.locked && Date.now() - (it.upd || 0) < 86400000 && S.seen[it.id] !== it.upd

/* ---------- 一覧 ---------- */
// 閲覧範囲（管理者だけに見せる）：管理者のみ／ほかのグループと共有
const roleOfGid = () => Object.fromEntries(Object.entries(S.roles || {}).map(([r, g]) => [g, r]))
const shareGids = (it) => Object.keys(it.k || {}).filter((g) => g !== S.adminGid)
const scopeOf = (it) => (shareGids(it).length ? 'share' : 'admin')
const scopeNames = (it) => { const m = roleOfGid(); return shareGids(it).map((g) => m[g] || '不明なグループ') }
/** 題名の横に置く閲覧権限のアイコン（管理者だけ）。鍵＝管理者のみ、人＝共有（グループ名つき） */
function scopeIcon(it) {
  if (!S.admin || it.locked) return ''
  if (scopeOf(it) === 'admin') return `<span class="scope" title="閲覧：管理者のみ" aria-label="閲覧：管理者のみ">${LOCK_SVG}</span>`
  const names = scopeNames(it)
  return `<span class="scope sh" title="閲覧：${esc(names.join('・'))}" aria-label="閲覧：${esc(names.join('・'))}">${USERS_SVG}<span class="scope-n">${esc(names.join('・'))}</span></span>`
}

/** 権限の絞り込み（管理者だけ）：すべて／管理者のみ／グループごと */
function scopeMatch(x) {
  if (S.scope === 'all') return true
  if (x.locked) return false
  if (S.scope === 'admin') return scopeOf(x) === 'admin'
  return shareGids(x).includes(S.scope)
}

function scopeSelectHTML(mi) {
  const open = mi.filter((x) => !x.locked)
  const groups = Object.entries(S.roles || {}).filter(([, g]) => g !== S.adminGid)
  const opt = (v, l, n) => `<option value="${esc(v)}" ${S.scope === v ? 'selected' : ''}>${esc(l)}（${n}）</option>`
  return `<label class="scopesel ${S.scope !== 'all' ? 'on' : ''}">${S.scope === 'admin' ? LOCK_SVG : USERS_SVG}<span class="sr">閲覧権限で絞り込み</span>
    <select id="scopeSel">${opt('all', '権限：すべて', open.length)}${opt('admin', '管理者のみ', open.filter((x) => scopeOf(x) === 'admin').length)}${groups.map(([r, g]) => opt(g, r, open.filter((x) => shareGids(x).includes(g)).length)).join('')}</select></label>`
}
function monthItems() {
  return S.items.filter((x) => (x.date || '').startsWith(S.month))
}

function filtered() {
  const q = S.q.trim().toLowerCase()
  return monthItems().filter((x) => {
    if (S.filter === 'ok' && x.st !== 'ok') return false
    if (S.filter === 'proc' && !['tw', 'tr', 'sw', 'err'].includes(x.st)) return false
    if (S.favOnly && !S.fav.has(x.id)) return false
    if (S.admin && !scopeMatch(x)) return false
    if (q) {
      if (x.locked) return false
      const s = cached(x, 's')
      const hay = [x.title, ...(x.tags || []), s?.cardSummary || '', s?.summary || ''].join(' ').toLowerCase()
      if (!hay.includes(q)) return false
    }
    return true
  })
}

function byDay(items) {
  const g = new Map()
  for (const it of items) { const k = it.day || '日付なし'; if (!g.has(k)) g.set(k, []); g.get(k).push(it) }
  for (const v of g.values()) v.sort((a, b) => (a.date || '').localeCompare(b.date || ''))
  return [...g.entries()]
}

function dayLabel(k) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(k)) return { d: k, w: '' }
  const d = toDate(k)
  return { d: `${d.getMonth() + 1}月${d.getDate()}日`, w: WD[d.getDay()] + (k === todayKey() ? '・今日' : '') }
}

function listHTML() {
  const mi = monthItems()
  const cnt = { all: mi.length, ok: mi.filter((x) => x.st === 'ok').length, proc: mi.filter((x) => ['tw', 'tr', 'sw', 'err'].includes(x.st)).length, fav: mi.filter((x) => S.fav.has(x.id)).length }
  const [y, mo] = S.month.split('-')
  const cards = S.view === 'cards'
  const items = filtered()
  const sel = S.sel ? S.byId.get(S.sel) : null
  // 状態はコンボボックス、お気に入りは押して切り替え（状態と組み合わせて絞り込める）
  const chips = `<label class="scopesel ${S.filter !== 'all' ? 'on' : ''}"><span class="sr">状態で絞り込み</span><select id="stSel">${[['all', 'すべて'], ['ok', '要約済み'], ['proc', '処理中']]
    .map(([k, l]) => `<option value="${k}" ${S.filter === k ? 'selected' : ''}>${l}（${cnt[k]}）</option>`).join('')}</select></label>`
    + `<button class="fchip ${S.favOnly ? 'on' : ''}" data-act="favonly" aria-pressed="${S.favOnly}">★ お気に入り ${cnt.fav}</button>`
  return `
    <div class="toolbar ${sel ? 'hide-m' : ''}">
      <div class="monthnav">
        <button class="ib" data-act="month" data-d="-1" aria-label="前の月">‹</button>
        <span class="month">${Number(y)}年${Number(mo)}月</span>
        <button class="ib" data-act="month" data-d="1" aria-label="次の月">›</button>
      </div>
      <div class="fchips">${chips}${S.admin ? scopeSelectHTML(mi) : ''}</div>
      <span class="grow"></span>
      <div class="seg" role="group" aria-label="表示の切り替え">
        <button data-act="view" data-v="timeline" aria-pressed="${!cards}" class="${!cards ? 'on' : ''}">タイムライン</button>
        <button data-act="view" data-v="cards" aria-pressed="${cards}" class="${cards ? 'on' : ''}">カード</button>
      </div>
    </div>
    ${cards ? stripHTML(mi).replace('class="strip"', `class="strip ${sel ? 'hide-m' : ''}"`) : ''}
    <div class="split ${sel ? 'has-sel' : ''} ${cards ? 'cards' : 'tl'}">
      <section class="listpane" aria-label="議事録の一覧">${items.length ? (cards ? cardsHTML(items) : timelineHTML(items)) : emptyListHTML(mi.length)}</section>
      <section class="detail" id="detail" aria-label="議事録の詳細">${detailHTML(sel)}</section>
    </div>`
}

function emptyListHTML(any) {
  if (any) return '<p class="empty">条件に合う議事録がありません。絞り込み（状態・お気に入り・権限）を戻してください。</p>'
  return '<div class="empty"><p>この月の議事録はまだありません。</p><button class="btn main" data-act="upload">議事録を登録</button></div>'
}

function timelineHTML(items) {
  return byDay(items).map(([k, list]) => {
    const l = dayLabel(k)
    return `<div class="day" id="d-${esc(k)}">
      <div class="dayhead"><span class="dd">${l.d}</span><span class="dw">${l.w}</span><span class="rule"></span><span class="dc">${list.length}件</span></div>
      ${list.map((it) => `
        <div class="tl-row">
          <div class="tl-time">${hm(it.date) || '—'}</div>
          <div class="tl-rail"><span class="tl-line"></span><span class="tl-dot ${(ST[it.st] || ST.tw).c}"></span><span class="tl-line grow"></span></div>
          <button class="tl-card ${S.sel === it.id ? 'sel' : ''}" data-act="open" data-id="${it.id}">
            <span class="tl-top"><span class="tl-title">${esc(it.title)}</span>${scopeIcon(it)}${it.locked ? LOCK_SVG : ''}${isNew(it) ? '<span class="new">新着</span>' : ''}${stChip(it)}</span>
            <span class="tl-meta"><span class="bar ${(ST[it.st] || ST.tw).c}" data-w="${Math.round(Math.min(it.dur || 0, 7200) / 7200 * 110) + 6}"></span>${fmtDur(it.dur)}${it.tags?.length ? ' ・ ' + esc(it.tags.join('・')) : ''}</span>
          </button>
        </div>`).join('')}
    </div>`
  }).join('')
}

function cardLead(it) {
  if (it.locked) return '閲覧できるグループに入ると表示されます。'
  const s = cached(it, 's')
  if (s?.cardSummary || s?.summary) return s.cardSummary || s.summary
  return { tw: 'Mac で文字起こしの順番待ちです。', tr: 'Mac で文字起こししています。', sw: '文字起こし済み。要約待ちです。', err: '処理に失敗しました。変換管理から再実行できます。', ok: '要約を読み込んでいます…' }[it.st] || ''
}

function cardsHTML(items) {
  return byDay(items).map(([k, list]) => {
    const l = dayLabel(k)
    return `<div class="day" id="d-${esc(k)}">
      <div class="dayhead"><span class="dd">${l.d}</span><span class="dw">${l.w}</span><span class="rule"></span></div>
      <div class="cardgrid">${list.map((it) => `
        <button class="mcard ${S.sel === it.id ? 'sel' : ''}" data-act="open" data-id="${it.id}">
          <span class="mc-top"><span class="mc-time">${hm(it.date) || '—'}</span><span class="mc-dur">${fmtDur(it.dur)}</span><span class="grow"></span>${stChip(it)}</span>
          <span class="mc-title">${esc(it.title)}${scopeIcon(it)}${it.locked ? LOCK_SVG : ''}${isNew(it) ? '<span class="new">新着</span>' : ''}</span>
          <span class="mc-lead">${esc(cardLead(it))}</span>
        </button>`).join('')}</div>
    </div>`
  }).join('')
}

function stripHTML(mi) {
  const [y, m] = S.month.split('-').map(Number)
  const days = new Date(y, m, 0).getDate()
  const by = new Map()
  for (const it of mi) { if (!by.has(it.day)) by.set(it.day, []); by.get(it.day).push(it) }
  const durOf = (it) => it.dur || 1800   // 長さが分からない会議は、30分として描く
  // 棒の高さは、その月でいちばん会議時間の長い日を最大（40px）にそろえる。会議ごとに積み上げ、色は状態
  const maxTotal = Math.max(1, ...[...by.values()].map((l) => l.reduce((a, it) => a + durOf(it), 0)))
  const today = todayKey()
  let h = ''
  for (let d = 1; d <= days; d++) {
    const k = `${S.month}-${pad(d)}`
    const list = by.get(k) || []
    const total = list.reduce((a, it) => a + durOf(it), 0)
    const bars = list.map((it) => `<span class="sb ${(ST[it.st] || ST.tw).c}" data-h="${Math.max(5, Math.round(durOf(it) / maxTotal * 40))}"></span>`).join('')
    const label = `${m}月${d}日 ${list.length}件${list.length ? '・計' + fmtT(total) : ''}`
    h += `<button class="sday ${k === today ? 'today' : ''}" ${list.length ? `data-act="jump" data-day="${k}"` : 'disabled'} aria-label="${label}" title="${label}"><span class="sbars">${bars}</span><span class="sn">${d}</span></button>`
  }
  return `<div class="strip"><div class="strip-in">${h}</div></div>`
}

/** 一覧を描いたあと：カードの要約・検索用に要約を先読みし、詳細のファイルを読む */
function afterList() {
  const sel = S.sel ? S.byId.get(S.sel) : null
  applySecJump()
  centerActiveTab()
  if (sel) loadDetail(sel)
  if (S.view === 'cards' || S.q) prefetch(monthItems().filter((x) => !x.locked && x.st === 'ok' && cached(x, 's') === undefined), 's')
}

let prefetchRun = 0
async function prefetch(list, name) {
  if (!list.length) return
  const run = ++prefetchRun
  for (let i = 0; i < list.length; i += 4) {
    await Promise.all(list.slice(i, i + 4).map((it) => file(it, name).catch(() => null)))
    if (run !== prefetchRun) return
  }
  if (S.tab === 'list') { const keep = $('.listpane')?.scrollTop; renderMain(); const lp = $('.listpane'); if (lp && keep) lp.scrollTop = keep }
}

/* ---------- 詳細 ---------- */
const NEED = { sum: ['s', 'n'], tr: ['t', 'n'], chat: ['s', 't'], map: ['s'] }

function loadDetail(it) {
  if (it.locked) return
  const need = NEED[S.dtab].filter((n) => cached(it, n) === undefined)
  if (!need.length) return
  Promise.all(need.map((n) => file(it, n))).then(() => {
    if (S.sel === it.id && S.tab === 'list') paintDetail()
  }).catch((e) => toast(e.message, true))
}

function paintDetail(reset = false) {
  const el = $('#detail')
  if (!el) return
  const it = S.sel ? S.byId.get(S.sel) : null
  const sc = reset ? 0 : el.querySelector('.sheet-scroll')?.scrollTop
  el.innerHTML = detailHTML(it)
  if (reset) scrollHide.reset()
  const sh = el.querySelector('.sheet-scroll')
  if (sc && sh) sh.scrollTop = sc
  const f = S.seek && !S.seek.scrolled && el.querySelector('.seg-row.flash')
  if (f) { S.seek.scrolled = true; f.scrollIntoView({ behavior: 'smooth', block: 'center' }) }
  if (S.jump) {
    const target = el.querySelector(`[data-mid="${S.jump.id}"]`)
    if (target) { target.classList.add('flash'); target.scrollIntoView({ behavior: 'smooth', block: 'center' }); S.jump = null }
    else if (it && NEED[S.dtab].every((n) => cached(it, n) !== undefined)) S.jump = null   // 見つからない（文章が変わった）
  }
  if (S.edit?.focus) { el.querySelector(S.edit.focus)?.focus(); S.edit.focus = null }
  applySecJump()
  centerActiveTab()
  if (it && S.dtab === 'map') mountMap(it)
  if (it) loadDetail(it)
}

/* ---------- タグ（よく使うタグを押して付け外し） ---------- */
function dtagsHTML(it, canEdit) {
  const tags = it.tags || []
  if (!tags.length && !canEdit) return ''
  return `<div class="dtags">${tags.map((t) => `<span class="ttag">${esc(t)}</span>`).join('')}${canEdit ? `<button class="ttag add" data-act="tags" data-id="${it.id}" aria-expanded="${S.tagOpen === it.id}">${S.tagOpen === it.id ? '閉じる' : tags.length ? '＋ タグ' : '＋ タグを付ける'}</button>` : ''}</div>`
}

function tagStats() {
  const m = new Map()
  for (const x of S.items) if (!x.locked) for (const t of x.tags || []) m.set(t, (m.get(t) || 0) + 1)
  return [...m].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'ja'))
}

function tagPanelHTML(it) {
  const cur = new Set(it.tags || [])
  const stats = tagStats().filter(([t]) => !cur.has(t))
  const s = cached(it, 's')
  const hay = [it.title, s?.cardSummary || '', s?.summary || '', ...((s?.agenda || []).map((a) => a.topic || ''))].join(' ')
  const rec = stats.filter(([t]) => t.length >= 2 && hay.includes(t)).slice(0, 6).map(([t]) => t)
  const freq = stats.filter(([t]) => !rec.includes(t)).slice(0, 18)
  const on = (t) => `<button class="ttag on" data-act="tag-off" data-id="${it.id}" data-t="${esc(t)}" aria-label="${esc(t)} を外す">${esc(t)}<span aria-hidden="true">×</span></button>`
  const off = (t, n) => `<button class="ttag off" data-act="tag-on" data-id="${it.id}" data-t="${esc(t)}" aria-label="${esc(t)} を付ける">${esc(t)}${n ? `<span class="tn">${n}</span>` : ''}</button>`
  return `<div class="tagpanel" role="region" aria-label="タグの付け外し">
    <div class="tp-row"><span class="tp-l">付いている</span><div class="tp-c">${cur.size ? [...cur].map(on).join('') : '<span class="muted sm">まだありません</span>'}</div></div>
    ${rec.length ? `<div class="tp-row"><span class="tp-l">この会議に合いそう</span><div class="tp-c">${rec.map((t) => off(t)).join('')}</div></div>` : ''}
    <div class="tp-row"><span class="tp-l">よく使う</span><div class="tp-c">${freq.length ? freq.map(([t, n]) => off(t, n)).join('') : '<span class="muted sm">ほかの会議のタグがここに並びます</span>'}</div></div>
    <form class="tp-new" data-tagform="${it.id}"><label for="tagNew" class="sr">新しいタグ</label><input id="tagNew" maxlength="20" placeholder="新しいタグ" autocomplete="off"><button class="btn sm" type="submit" data-act="tag-add" data-id="${it.id}">追加</button></form>
    <p class="muted sm">押すとすぐ保存します。「この会議に合いそう」は、ほかの会議で使ったタグのうち、題名や要約に出てくる言葉です。</p>
  </div>`
}

let tagQueue = Promise.resolve()
const normTag = (t) => String(t || '').replace(/[、,，]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 20)

/** タグを変える。先に画面へ反映し（META を差し替える）、保存は順番に送る */
function setTags(id, next) {
  const m = META.get(id)
  if (!m) return
  const meta = { ...m.v, tags: next }
  META.set(id, { iv: m.iv, v: meta })   // 保存が返るまで、decryptAll がこの内容を使う
  const it = S.byId.get(id); if (it) it.tags = next
  paintDetail()
  tagQueue = tagQueue.then(async () => {
    const latest = META.get(id)?.v || meta
    const r = await gasS('put', { id, patch: { m: await sealJson(DEK.get(id), latest) } })
    await applyEntry(r.entry, {}, false, true)
    renderMain()
  }).catch((e) => { toast('タグを保存できませんでした：' + e.message, true); refresh(true) })
}

function toggleTag(id, t, add) {
  const it = S.byId.get(id); if (!it) return
  const cur = it.tags || []
  if (add && !cur.includes(t)) setTags(id, [...cur, t])
  if (!add) setTags(id, cur.filter((x) => x !== t))
}

function addTagFromInput(id) {
  const t = normTag($('#tagNew')?.value)
  if (!t) { $('#tagNew')?.focus(); return }
  toggleTag(id, t, true)
  setTimeout(() => $('#tagNew')?.focus(), 0)
}

function metaLine(it) {
  const s = cached(it, 's')
  const tail = it.st === 'ok' ? (s?.model || '要約済み') : (ST[it.st] || ST.tw).l
  return `${it.src === 'text' ? 'テキスト' : it.src === 'plaud' ? 'PLAUD' : '会議'} ・ ${md(it.date)} ${hm(it.date)} ・ ${fmtDur(it.dur)} ・ ${esc(tail)}`
}

function detailHTML(it) {
  const cards = S.view === 'cards'
  if (!it) return cards ? '' : '<div class="empty big"><p>一覧から議事録を選んでください。</p></div>'
  const back = `<button class="btn back" data-act="close-detail">${cards ? '✕ 閉じる' : '‹ 一覧へ'}</button>`
  if (it.locked) {
    return `<div class="dhead">${back}<div class="dtitle"><h1>${esc(it.title)}</h1><div class="dmeta">${md(it.date)} ${hm(it.date)} ・ ${fmtDur(it.dur)}</div></div></div>
      <div class="lockpanel">${LOCK_SVG}<p class="strong">この議事録を見る権限がありません</p><p>日時と長さだけ表示しています。タイトルと内容は、閲覧できるグループの鍵がある端末でだけ開けます。</p></div>`
  }
  const canEdit = S.admin || it.by === S.gid
  const canDel = S.admin || (it.by === S.gid && ['tw', 'err'].includes(it.st))
  const hasT = ['sw', 'ok'].includes(it.st)
  const acts = [
    `<button class="btn sm" data-act="fav" data-id="${it.id}" aria-pressed="${S.fav.has(it.id)}">${S.fav.has(it.id) ? '★ お気に入り' : '☆ お気に入り'}</button>`,
    hasT ? `<button class="btn sm" data-act="summarize" data-id="${it.id}" ${S.busy[it.id] ? 'disabled' : ''}>${it.st === 'ok' ? '再要約' : '要約する'}</button>` : '',
    it.st === 'sw' ? `<button class="btn sm" data-act="claude" data-id="${it.id}">Claude.ai で要約</button>` : '',
    canEdit ? `<button class="btn sm" data-act="edit" data-id="${it.id}">タイトル・日時</button>` : '',
    S.admin ? `<button class="btn sm" data-act="groups" data-id="${it.id}">閲覧グループ</button>` : '',
    canDel ? `<button class="btn sm danger" data-act="delete" data-id="${it.id}">削除</button>` : '',
  ].join('')
  const nT = cached(it, 's')?.todos?.length
  const onSum = S.dtab === 'sum'
  const secs = [['summary', 'サマリ'], ['agenda', '論点別'], ['decisions', '決定事項'], ['todo', 'ToDo'], ['memo', 'メモ']]
  const views = [['tr', '文字起こし'], ['chat', 'チャット'], ['map', 'マインドマップ']]
  const tabs = secs.map(([k, l]) => `<button class="dtab ${onSum && S.sec === k ? 'on' : ''}" data-act="sec" data-v="${k}" aria-pressed="${onSum && S.sec === k}">${l}${k === 'todo' && nT ? ' ' + nT : ''}</button>`).join('')
    + '<span class="dtab-sep" aria-hidden="true"></span>'
    + views.map(([k, l]) => `<button class="dtab ${S.dtab === k ? 'on' : ''}" data-act="dtab" data-v="${k}" aria-pressed="${S.dtab === k}">${l}</button>`).join('')
  const backTab = '<button class="dtab-back" data-act="close-detail" aria-label="一覧へ">‹</button>'
  return `
    <div class="dhead">${back}
      <div class="dtitle"><h1>${esc(it.title)}${scopeIcon(it)}${isNew(it) ? '<span class="new">新着</span>' : ''}</h1><div class="dmeta">${metaLine(it)}</div>
        ${dtagsHTML(it, canEdit)}</div>
      <div class="dacts">${acts}</div>
    </div>
    ${canEdit && S.tagOpen === it.id ? tagPanelHTML(it) : ''}
    <nav class="dtabs" aria-label="詳細の切り替え">${backTab}${tabs}</nav>
    <div class="sheet-scroll"><div class="sheet">${tabBody(it)}</div></div>`
}

const tchip = (t) => (t && parseT(t) != null ? `<button class="tchip" data-act="seek" data-t="${esc(t)}" aria-label="${esc(t)} の発言へ">${esc(t)}</button>` : '')
const loadingHTML = '<p class="muted">読み込んでいます…</p>'

function tabBody(it) {
  switch (S.dtab) {
    case 'tr': return trBody(it)
    case 'chat': return chatBody(it)
    case 'map': return mapBody(it)
    default: return bodyAll(it)
  }
}

/* ---------- 本文：マーカー・ToDo の部品 ---------- */
function markerFn(it) {
  const g = groupMarks(marksOf(cached(it, 'n')))
  return (loc, text, q = '') => `<span class="mk" data-m="${it.id}" data-loc="${esc(loc)}">${renderMarked(text, validFor(text, g.get(loc)), esc, q)}</span>`
}
const tagName = (c) => (MARK_TAGS.find((t) => t.id === c) || {}).name || `色${c}`
const fmtLog = (ms) => { const d = new Date(ms); return `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}` }
const canEditBody = (it) => !it.locked

/** 要約の本文：サマリ・論点別・決定事項・ToDo・メモを、1枚のスクロールに並べる（タブは、その見出しへ飛ぶ） */
function bodyAll(it) {
  const s = cached(it, 's'), n = cached(it, 'n')
  if (s === undefined || n === undefined) return loadingHTML
  const memo = memoSection(it, n)
  if (!s) {
    let wait
    if (['tw', 'tr'].includes(it.st)) wait = '<p class="muted">Mac の文字起こしが終わると、要約できるようになります。</p>'
    else if (it.st === 'err') wait = `<p class="muted">処理に失敗しました${it.err ? `（${esc(it.err)}）` : ''}。変換管理から再実行できます。</p>`
    else wait = `<div class="sumwait"><p>まだ要約されていません。</p><div class="acts"><button class="btn main" data-act="summarize" data-id="${it.id}" ${S.busy[it.id] ? 'disabled' : ''}>${S.busy[it.id] ? esc(S.busy[it.id]) : '要約する（' + esc(activeModelName(loadSettings()) || 'AI未設定') + '）'}</button><button class="btn" data-act="claude" data-id="${it.id}">Claude.ai で要約</button></div></div>`
    return `<div class="read body">${wait}${memo}</div>`
  }
  if (S.edit && S.edit.id === it.id && S.edit.kind === 'sum') return `<div class="read body">${editorHTML(S.edit.draft)}${memo}</div>`
  const m = markerFn(it)
  const sec = (key, title, inner) => `<section class="ssec" data-sec="${key}"><h2>${title}</h2>${inner}</section>`
  const none = (t) => `<p class="muted">${t}</p>`
  const todos = s.todos || []
  const open = todos.filter((t, i) => !doneOf(n, todoKey(t, i))).length
  return `<div class="read body">
    <div class="sumbar"><span class="hint">文字を選ぶとマーカーを付けられます。付けた箇所は「ノート」に集まります。</span>${fsCtl()}${canEditBody(it) ? `<button class="btn sm" data-act="edit-sum" data-id="${it.id}">本文を編集</button>` : ''}</div>
    ${sec('summary', 'サマリ', `<p>${m('summary', s.summary || s.cardSummary || '')}</p>`)}
    ${sec('agenda', '論点別', `${s.agenda.length ? s.agenda.map((a, ai) => `
      <div class="topic">
        <h3>${m(`agenda:${ai}:topic`, a.topic)} ${tchip(a.t)}</h3>
        ${a.points.length ? `<ul class="dots">${a.points.map((p, j) => `<li><span>${m(`agenda:${ai}:point:${j}`, p.text)} ${tchip(p.t)}</span></li>`).join('')}</ul>` : ''}
        ${a.outcome ? `<p class="outcome"><span class="ol">結論</span>${m(`agenda:${ai}:outcome`, a.outcome)}</p>` : ''}
      </div>`).join('') : none('論点はありません。')}
      ${s.topics.length ? `<h3 class="sub">未解決の論点</h3><ul class="dots">${s.topics.map((t, i) => `<li><span>${m('topic:' + i, t)}</span></li>`).join('')}</ul>` : ''}`)}
    ${sec('decisions', '決定事項', s.decisions.length ? `<ul class="dots">${s.decisions.map((p, i) => `<li><span>${m('decision:' + i, p.text)} ${tchip(p.t)}</span></li>`).join('')}</ul>` : none('決定事項はありません。'))}
    ${sec('todo', 'ToDo', todos.length ? `<p class="muted sm">残り ${open} / ${todos.length} 件 ・ 経過は追記だけできます（消せません）</p><ul class="todos">${todos.map((t, i) => todoItem(it, t, i, n, m)).join('')}</ul>` : none('この会議の ToDo はありません。「本文を編集」で追加することもできます。'))}
    ${memo}
    <p class="foot">要約：${esc(s.model || '—')}${s.edited ? ' ・ 手直し済み' : ''} ・ 時刻を押すと、その発言を文字起こしで表示します</p>
  </div>`
}

function memoSection(it, n) {
  return `<section class="ssec" data-sec="memo"><h2>メモ</h2>
    <label for="memoIn" class="sr">メモ</label>
    <textarea id="memoIn" class="memoin" rows="6" placeholder="この会議のメモ（同じグループの人にも見えます）">${esc(n?.memo || '')}</textarea>
    <div class="acts"><button class="btn main" data-act="memo-save" data-id="${it.id}">メモを保存</button><span class="muted" id="memoMsg"></span></div>
  </section>`
}

/* ---------- マインドマップ（共通部品 api/mindmap2。markmap はリポジトリの vendor/ から読む） ----------
   表示：保存済みのマインドマップ（s.mindmap）があればそれ、無ければ要約の構成から作る
   編集：管理者と登録した人だけ。項目を押して選び、ボタンかキー（Space・Tab・Enter・Delete）で直す。直すと少し待って自動で保存 */
const MM = { id: null, src: '', md: '', el: null, mm: null, dirty: false, saving: false, timer: 0, status: '' }
const canEditIt = (it) => S.admin || it.by === S.gid

function mapSource(it, s) {
  return s.mindmap && !isLegacyHtml(s.mindmap) ? String(s.mindmap) : mindmapMarkdown(it.title, s)
}

function mapBody(it) {
  const s = cached(it, 's')
  if (s === undefined) return loadingHTML
  if (!s) return '<p class="muted">要約すると、マインドマップが作られます。</p>'
  const ed = canEditIt(it)
  const saved = s.mindmap && !isLegacyHtml(s.mindmap)
  const keys = [['Space', '名前を変える', ' '], ['Tab', '子を追加', 'Tab'], ['Enter', '下に追加', 'Enter'], ['Delete', '削除', 'Delete']]
  return `<div class="mmwrap">
    <div class="acts mmbar">
      <button class="btn sm" data-act="mm-fit">全体を表示</button>
      ${ed ? keys.map(([k, l, key]) => `<button class="btn sm" data-act="mm-key" data-k="${key}" title="${k}">${l}</button>`).join('') : ''}
      ${ed && saved ? `<button class="btn sm danger" data-act="mm-reset" data-id="${it.id}">要約から作り直す</button>` : ''}
      <span class="muted sm" id="mmStatus" role="status">${esc(MM.id === it.id && MM.status ? MM.status : (saved ? '保存済みのマインドマップ' : '要約の構成から作っています'))}</span>
    </div>
    <div id="mmSlot" class="mm2" data-id="${it.id}"></div>
    <p class="muted sm">${ed ? '項目を押して選び、上のボタンかキー（Space・Tab・Enter・Delete、矢印で移動）で直します。直すと自動で保存します。' : '丸を押すと開け閉めできます。ドラッグで移動、ピンチやホイールで拡大縮小。'}</p>
  </div>`
}

function ensureMapCss() {
  if (document.querySelector('link[data-mm2]')) return
  const l = document.createElement('link')
  l.rel = 'stylesheet'; l.href = new URL('../api/mindmap2/mindmap2.css', location.href).href; l.dataset.mm2 = '1'
  document.head.append(l)
}

/** paintDetail のあとに呼ぶ。同じ会議・同じ中身なら、描いた地図をそのまま差し戻す（位置や開閉を保つ） */
async function mountMap(it) {
  const slot = $('#mmSlot')
  if (!slot || !it) return
  const s = cached(it, 's')
  if (!s) return
  const src = mapSource(it, s)
  if (MM.el && MM.id === it.id && (MM.dirty || MM.saving || MM.md === src)) { slot.replaceWith(MM.el); return }
  clearTimeout(MM.timer)
  const el = document.createElement('div')
  el.className = 'mm2'; el.id = 'mmHost'; el.setAttribute('aria-label', 'マインドマップ')
  slot.replaceWith(el)
  Object.assign(MM, { id: it.id, src, md: src, el, mm: null, dirty: false, saving: false, status: '' })
  ensureMapCss()
  const ed = canEditIt(it)
  MM.mm = await renderMindmap(el, src, {
    autoFocus: false,
    initialExpandLevel: 3,
    emptyText: 'マインドマップにする内容がありません。',
    onChange: ed ? (md) => {
      if (MM.el !== el) return
      MM.md = md; MM.dirty = true
      mmStatus('保存待ち…')
      clearTimeout(MM.timer)
      MM.timer = setTimeout(() => saveMap(it.id), 1500)
    } : undefined,
  })
}

function mmStatus(t) { MM.status = t; const e = $('#mmStatus'); if (e) e.textContent = t }

/** 選んでいる項目に、キー操作と同じことをする（スマホなどキーボードが無いとき用） */
function mmKey(key) {
  if (!MM.el) return
  MM.el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
}

async function saveMap(id) {
  const it = S.byId.get(id)
  const s = it && cached(it, 's')
  if (!s || MM.id !== id || MM.saving) return
  MM.saving = true; MM.dirty = false
  const md = MM.md
  mmStatus('保存しています…')
  try {
    await saveFile(it, 's', { ...s, mindmap: md })
    mmStatus('保存しました')
  } catch (e) {
    MM.dirty = true
    mmStatus('保存できませんでした：' + e.message)
  } finally {
    MM.saving = false
    if (MM.dirty && MM.md !== md) { clearTimeout(MM.timer); MM.timer = setTimeout(() => saveMap(id), 1500) }
  }
}

async function resetMap(id) {
  const it = S.byId.get(id)
  const s = it && cached(it, 's')
  if (!s || !confirm('手で直したマインドマップを消して、要約の構成から作り直します。よろしいですか？')) return
  clearTimeout(MM.timer)
  MM.el = null; MM.dirty = false
  try { const next = { ...s }; delete next.mindmap; await saveFile(it, 's', next); toast('要約から作り直しました。') } catch (e) { toast(e.message, true) }
}

/** ToDo 1件：完了・本文（マーカーつき）・担当と期限・経過・経過の追記 */
function todoItem(it, t, i, n, m) {
  const key = todoKey(t, i), done = doneOf(n, key), logs = logsOf(n, key)
  const open = S.logOpen === `${it.id}|${key}`
  const meta = `${t.owner ? '担当：' + esc(t.owner) : ''}${t.owner && t.due ? ' ・ ' : ''}${t.due ? '期限：' + esc(t.due) : ''}`
  return `<li class="todo ${done ? 'done' : ''}">
    <div class="trow"><input type="checkbox" data-todo="${it.id}" data-k="${esc(key)}" aria-label="完了にする" ${done ? 'checked' : ''}><span class="tt">${m('todo:' + key, t.text)}</span>${meta ? `<span class="tm">${meta}</span>` : ''}</div>
    ${t.orphan ? '<div class="orph">再要約で対応する ToDo が見つからなかったため、経過ごと残しています</div>' : ''}
    ${logs.length ? `<div class="tlogs">${logs.map((l) => `<div class="tlog"><span class="tld">${esc(fmtLog(l.d))}</span>${esc(l.x)}</div>`).join('')}</div>` : ''}
    ${open
      ? `<form class="logform" data-log="${it.id}" data-k="${esc(key)}"><label class="sr" for="logIn">経過を追記</label><input id="logIn" placeholder="経過を追記（Enter で追記）" autocomplete="off"><button class="btn sm main" type="submit">追記</button><button class="btn sm" type="button" data-act="log-close">閉じる</button></form>`
      : `<button class="btn sm tlbtn" data-act="log-open" data-id="${it.id}" data-k="${esc(key)}">経過を追記${logs.length ? `（${logs.length}件）` : ''}</button>`}
  </li>`
}

function trBody(it) {
  const t = cached(it, 't'), n = cached(it, 'n')
  if (t === undefined || n === undefined) return loadingHTML
  if (!t?.segments?.length) return `<p class="muted">${['tw', 'tr'].includes(it.st) ? 'まだ文字起こしされていません。Mac の次の確認で始まります。' : '文字起こしがありません。'}</p>`
  if (S.edit && S.edit.id === it.id && S.edit.kind === 'tr') {
    return `<div class="read tr"><p class="muted sm">1行が1つの発言です。行頭の [1:23] は発言の時刻です（直さなければそのまま残ります）。行を足す・消すこともできます。</p>
      <label class="sr" for="trEdit">文字起こしの本文</label><textarea id="trEdit" class="treditor" rows="22">${esc(S.edit.text)}</textarea>
      <div class="acts"><button class="btn main" data-act="e-save">保存</button><button class="btn" data-act="e-cancel">キャンセル</button><span class="msg" id="edMsg" role="alert"></span></div></div>`
  }
  const m = markerFn(it)
  const q = S.trq.trim()
  let target = -1
  if (S.seek) t.segments.forEach((x, i) => { if (x.t != null && x.t <= S.seek.t) target = i })
  if (S.seek && target < 0) target = 0
  return `<div class="read tr">
    <div class="trbar"><label for="trq" class="sr">文字起こし内を検索</label><input id="trq" type="search" placeholder="文字起こし内を検索" value="${esc(q)}">${fsCtl()}${canEditBody(it) ? `<button class="btn sm" data-act="edit-tr" data-id="${it.id}">本文を編集</button>` : ''}<button class="btn sm" data-act="copy-tr" data-id="${it.id}">コピー</button></div>
    <p class="hint">文字を選ぶとマーカーを付けられます。</p>
    <div class="segs">${t.segments.map((x, i) => {
      const hit = q && x.s.includes(q)
      return `<div class="seg-row ${hit ? 'hit' : ''} ${i === target ? 'flash' : ''}" data-t="${x.t ?? ''}" id="seg-${i}"><span class="seg-t">${x.t != null ? fmtT(x.t) : ''}</span><span class="seg-s">${m('tr:' + i, x.s, q)}</span></div>`
    }).join('')}</div>
  </div>`
}

function chatBody(it) {
  const msgs = S.chat[it.id] || []
  return `<div class="read chat">
    <div class="msgs" id="msgs">${msgs.length ? msgs.map((m) => `<div class="msg-b ${m.role}">${esc(m.content) || '…'}</div>`).join('') : '<p class="muted">この会議について質問できます（' + esc(activeModelName(loadSettings()) || 'AI未設定') + '）。答えには根拠の時刻を添えます。</p>'}</div>
    <form class="ask" data-chat="${it.id}"><label for="askIn" class="sr">質問</label><textarea id="askIn" rows="2" placeholder="この会議について質問する"></textarea><button class="btn main" type="submit">送信</button></form>
  </div>`
}

/* ---------- ToDo（会議をまたいで）と ノート（マーカーの集まり） ---------- */
const RANGES = [['1m', '今月'], ['3m', '直近3か月'], ['all', 'すべて']]
function inRange(it) {
  if (S.range === 'all') return true
  const now = new Date()
  const from = new Date(now.getFullYear(), now.getMonth() - (S.range === '1m' ? 0 : 2), 1)
  return toDate(it.date) >= from
}
const rangeSelect = () => `<label class="sr" for="rangeSel">期間</label><select id="rangeSel">${RANGES.map(([k, l]) => `<option value="${k}" ${S.range === k ? 'selected' : ''}>${l}</option>`).join('')}</select>`

const FAILED = new Set()
let kickRun = 0
/** 画面に出す会議のファイルを、少しずつ読み込む（読めた分から画面に反映する） */
function kickLoad(list, names, tab) {
  const key = (it, nm) => `${it.id}:${nm}:${it.upd}`
  const need = list.filter((it) => names.some((nm) => cached(it, nm) === undefined && !FAILED.has(key(it, nm))))
  if (!need.length) return
  const run = ++kickRun
  ;(async () => {
    for (let i = 0; i < need.length; i += 6) {
      await Promise.all(need.slice(i, i + 6).flatMap((it) => names.map((nm) => file(it, nm).catch(() => { FAILED.add(key(it, nm)); return null }))))
      if (run !== kickRun) return
      if (S.tab === tab && !$('#main input:focus, #main textarea:focus')) renderMain()
    }
  })()
}

const todoEntries = () => S.items.filter((x) => !x.locked && x.st === 'ok' && x.td !== 0 && inRange(x))
function kickTodo() { kickLoad(todoEntries(), ['s', 'n'], 'todo') }

function todoPageHTML() {
  const list = todoEntries()
  const ready = list.filter((it) => cached(it, 's') !== undefined && cached(it, 'n') !== undefined)
  const data = ready.map((it) => ({ it, s: cached(it, 's'), n: cached(it, 'n') })).filter((x) => x.s && x.s.todos && x.s.todos.length)
  let open = 0, done = 0
  const owners = new Set()
  data.forEach(({ s, n }) => s.todos.forEach((t, i) => { if (doneOf(n, todoKey(t, i))) done++; else open++; if (t.owner) owners.add(t.owner) }))
  const chips = [['open', '未完了', open], ['done', '完了', done], ['all', 'すべて', open + done]]
    .map(([k, l, c]) => `<button class="fchip ${S.tfilter === k ? 'on' : ''}" data-act="tfilter" data-v="${k}" aria-pressed="${S.tfilter === k}">${l} ${c}</button>`).join('')
  const blocks = data.map(({ it, s, n }) => {
    const m = markerFn(it)
    const rows = s.todos.map((t, i) => ({ t, i })).filter(({ t, i }) => {
      const d = doneOf(n, todoKey(t, i))
      if (S.tfilter === 'open' && d) return false
      if (S.tfilter === 'done' && !d) return false
      return !S.towner || t.owner === S.towner
    })
    if (!rows.length) return ''
    const left = s.todos.filter((t, i) => !doneOf(n, todoKey(t, i))).length
    return `<section class="nblock">
      <div class="nhead"><button class="link" data-act="open" data-id="${it.id}" data-go="todo">${esc(it.title)}</button><span class="muted">${md(it.date)} ・ 残り ${left} / ${s.todos.length}</span></div>
      <ul class="todos">${rows.map(({ t, i }) => todoItem(it, t, i, n, m)).join('')}</ul></section>`
  }).join('')
  const loading = ready.length < list.length ? `<p class="muted sm">読み込んでいます… ${ready.length}/${list.length}</p>` : ''
  const emptyMsg = S.tfilter === 'open' ? '未完了の ToDo はありません。' : 'この条件に合う ToDo はありません。'
  return `<div class="page">
    <div class="toolbar"><h1 class="ptitle">ToDo</h1>${chips}
      ${owners.size ? `<label class="sr" for="towner">担当</label><select id="towner"><option value="">担当：全員</option>${[...owners].sort().map((o) => `<option ${S.towner === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>` : ''}
      <span class="grow"></span>${fsCtl()}${rangeSelect()}</div>
    <div class="notes">${blocks || (loading ? '' : `<p class="empty">${emptyMsg}</p>`)}${loading}</div>
  </div>`
}

const noteEntries = () => S.items.filter((x) => !x.locked && x.nt && inRange(x))
function kickNotes() { kickLoad(noteEntries(), ['n'], 'notes') }

function notesPageHTML() {
  const list = noteEntries()
  const all = []
  let pending = 0
  list.forEach((it) => {
    const n = cached(it, 'n')
    if (n === undefined) { pending++; return }
    marksOf(n).forEach((mk) => all.push({ it, mk }))
  })
  const q = S.nq.trim().toLowerCase()
  const counts = { 0: all.length }
  MARK_TAGS.forEach((t) => { counts[t.id] = all.filter((x) => x.mk.c === t.id).length })
  const shown = all.filter(({ mk }) => (!S.nfilter || mk.c === S.nfilter) && (!q || (mk.q + ' ' + (mk.m || '')).toLowerCase().includes(q)))
  const chips = [{ id: 0, name: 'すべて' }, ...MARK_TAGS]
    .map((t) => `<button class="fchip ${S.nfilter === t.id ? 'on' : ''}" data-act="nfilter" data-v="${t.id}" aria-pressed="${S.nfilter === t.id}">${t.id ? `<span class="mdot c${t.id}"></span>` : ''}${esc(t.name)} ${counts[t.id]}</button>`).join('')
  const byMeeting = new Map()
  shown.forEach((x) => { if (!byMeeting.has(x.it.id)) byMeeting.set(x.it.id, { it: x.it, items: [] }); byMeeting.get(x.it.id).items.push(x.mk) })
  const blocks = [...byMeeting.values()].map(({ it, items }) => `<section class="nblock">
      <div class="nhead"><button class="link" data-act="open" data-id="${it.id}">${esc(it.title)}</button><span class="muted">${md(it.date)}</span></div>
      ${items.sort((a, b) => locCompare(a.loc, b.loc) || a.s - b.s).map((mk) => `
        <div class="mkcard c${mk.c}" data-mid="${esc(mk.id)}">
          <blockquote class="mq">${esc(mk.q)}</blockquote>
          <div class="mmeta"><span class="mtag c${mk.c}">${esc(tagName(mk.c))}</span><span>${esc(locLabel(mk.loc))}</span>${mk.tt != null ? `<span class="tld">${fmtT(mk.tt)}</span>` : ''}${mk.o ? '<span class="orph">元の文章が変わりました</span>' : ''}</div>
          ${mk.m ? `<p class="mmemo">${esc(mk.m)}</p>` : ''}
          ${S.mkMemo === mk.id ? `<form class="logform" data-mkmemo="${it.id}" data-mid="${esc(mk.id)}"><label class="sr" for="mkMemoIn">ノート</label><input id="mkMemoIn" value="${esc(mk.m || '')}" placeholder="メモを書く（Enter で保存）" autocomplete="off"><button class="btn sm main" type="submit">保存</button><button class="btn sm" type="button" data-act="mk-memo-close">閉じる</button></form>` : ''}
          <div class="mact">
            ${mk.o ? '' : `<button class="btn sm" data-act="note-open" data-id="${it.id}" data-mid="${esc(mk.id)}">元の位置を開く</button>`}
            <button class="btn sm" data-act="mk-memo" data-mid="${esc(mk.id)}">メモ</button>
            <span class="mcolors" role="group" aria-label="色を変える">${MARK_TAGS.map((t) => `<button class="mdot c${t.id} ${t.id === mk.c ? 'on' : ''}" data-act="mk-color" data-id="${it.id}" data-mid="${esc(mk.id)}" data-c="${t.id}" aria-label="${esc(t.name)}にする" aria-pressed="${t.id === mk.c}"></button>`).join('')}</span>
            <button class="btn sm danger" data-act="mk-del" data-id="${it.id}" data-mid="${esc(mk.id)}">消す</button>
          </div>
        </div>`).join('')}
    </section>`).join('')
  const loading = pending ? `<p class="muted sm">読み込んでいます… ${list.length - pending}/${list.length}</p>` : ''
  return `<div class="page">
    <div class="toolbar"><h1 class="ptitle">ノート</h1>${chips}<span class="grow"></span>
      <label class="sr" for="nq">ノート内を検索</label><input id="nq" class="nsearch" type="search" placeholder="ノート内を検索" value="${esc(S.nq)}">${fsCtl()}${rangeSelect()}</div>
    <div class="notes">${blocks || (pending ? '' : '<p class="empty">マーカーはまだありません。議事録の文字を選ぶと、マーカーを付けられます。付けた箇所は、ここに集まります。</p>')}${loading}</div>
  </div>`
}

/* ---------- 変換管理 ---------- */
function workerLine() {
  const w = S.worker
  if (!w?.at) return { ok: false, t: 'Mac から、まだ確認の記録が届いていません。', s: 'Mac で minutes_worker を設定してください。' }
  const every = (w.every || WORKER_EVERY_SEC) * 1000
  const late = Date.now() - w.at > every * 4
  const busy = w.busy ? S.byId.get(w.busy) : null
  return {
    ok: !late,
    t: late ? 'Mac から応答がありません' : (busy ? `Mac が文字起こし中：${busy.title}` : 'Mac mini 監視中'),
    s: `最終確認 ${ago(w.at)}（${Math.round(every / 60000)}分おき）${w.asr ? ' ・ ' + esc(w.asr) : ''}`,
  }
}

function notionLine() {
  const ns = notionState()
  if (!ns) return ''
  if (!ns.ok) return `<div class="sm bad-t">Notion の作業DB：書き出せていません（${esc(ns.err)}）</div>`
  return `<div class="muted sm">Notion の作業DB：要約待ち ${(ns.ids || []).length} 件を書き出し済み</div>`
}

function convertHTML() {
  const vis = S.items
  const c = (k) => vis.filter((x) => x.st === k).length
  const steps = [['tw', '文字起こし待ち', 'Mac の次の確認で始まります'], ['tr', '文字起こし中', ''], ['sw', '要約待ち', 'AI を選んで要約'], ['ok', '要約済み', ''], ['err', '失敗', '再実行できます']]
  const w = workerLine()
  const settings = loadSettings()
  const models = allModels(settings)
  let rows = vis.filter((x) => S.cfilter === 'all' ? x.st !== 'ok' || Date.now() - (x.upd || 0) < 7 * 86400000 : x.st === S.cfilter)
  rows = rows.slice(0, 200)
  const swSel = [...S.checked].filter((id) => S.byId.get(id)?.st === 'sw')
  const swAll = vis.filter((x) => x.st === 'sw' && !x.locked)
  return `<div class="page conv">
    <div class="cv-top">
      <div class="panel wk"><span class="dot ${w.ok ? 'ok' : 'bad'}"></span><div class="grow"><div class="strong">${esc(w.t)}</div><div class="muted sm">${w.s}</div>${notionLine()}</div></div>
      <div class="panel"><div class="grow"><div class="strong">要約AI</div><div class="muted sm">URL で指定した AI（この端末の設定）</div></div>
        <label for="aiSel" class="sr">要約AI</label>
        <select id="aiSel">${models.length ? models.map((m) => `<option value="${esc(m.connectionId)}|${esc(m.model)}" ${m.active ? 'selected' : ''}>${esc(m.connectionLabel)}：${esc(m.model)}</option>`).join('') : '<option>未設定</option>'}</select>
        <button class="btn sm" data-act="settings">設定</button></div>
    </div>
    <div class="steps">${steps.map(([k, l, sub]) => `<button class="step ${S.cfilter === k ? 'on' : ''} ${k}" data-act="cfilter" data-v="${S.cfilter === k ? 'all' : k}" aria-pressed="${S.cfilter === k}"><span class="sl">${l}</span><span class="sn">${c(k)}</span><span class="ss">${sub}</span></button>`).join('')}</div>
    <div class="cv-acts">
      <button class="btn" data-act="check-sw">要約待ちを全選択</button>
      <button class="btn main" data-act="sum-selected" ${swSel.length ? '' : 'disabled'}>選択を要約する（${swSel.length}件）</button>
      <span class="grow"></span>
      <button class="btn gold" data-act="claude">Claude.ai で要約待ちを変換（${swAll.length}件）</button>
    </div>
    <p class="cnote" id="cnote" role="status" ${S.cnote ? '' : 'hidden'}>${esc(S.cnote)}</p>
    <div class="tablebox"><table class="ctable">
      <thead><tr><th><span class="sr">選択</span></th><th>日時</th><th>タイトル</th><th>長さ</th><th>登録</th><th>状態</th><th></th></tr></thead>
      <tbody>${rows.length ? rows.map((it) => {
        const own = it.by === S.gid
        let act = ''
        if (it.locked) act = ''
        else if (it.st === 'sw') act = `<button class="btn sm" data-act="summarize" data-id="${it.id}" ${S.busy[it.id] ? 'disabled' : ''}>要約する</button>`
        else if (it.st === 'err') act = `<button class="btn sm" data-act="retry" data-id="${it.id}">再実行</button>${S.admin || own ? ` <button class="btn sm" data-act="delete" data-id="${it.id}">取り消す</button>` : ''}`
        else if (it.st === 'tw' && (S.admin || own)) act = `<button class="btn sm" data-act="delete" data-id="${it.id}">取り消す</button>`
        else if (it.st === 'ok') act = `<button class="btn sm" data-act="open" data-id="${it.id}">開く</button>`
        return `<tr>
          <td>${it.st === 'sw' && !it.locked ? `<input type="checkbox" data-chk="${it.id}" ${S.checked.has(it.id) ? 'checked' : ''} aria-label="${esc(it.title)} を選ぶ">` : ''}</td>
          <td class="nw">${md(it.date)} ${hm(it.date)}</td>
          <td class="strong">${esc(it.title)}${it.locked ? LOCK_SVG : ''}</td>
          <td>${fmtDur(it.dur)}</td>
          <td>${SRC[it.src] || '音声'}</td>
          <td>${stChip(it)}${it.st === 'err' && it.err && !it.locked ? `<div class="muted sm">${esc(it.err)}</div>` : ''}</td>
          <td class="nw r">${act}</td></tr>`
      }).join('') : '<tr><td colspan="7" class="muted c">該当する議事録はありません。</td></tr>'}</tbody>
    </table></div>
  </div>`
}

/* ---------- 要約 ---------- */
async function saveFile(it, name, obj, patch) {
  const p = { ...(patch || {}) }
  if (name === 's') p.td = (obj.todos || []).length   // 一覧の側に ToDo の件数を持たせる（ToDo 一覧で、読まなくてよい会議を飛ばすため）
  const sealed = await sealJson(DEK.get(it.id), obj)
  const r = await gasS('put', { id: it.id, patch: p, files: { [name]: sealed } })
  await applyEntry(r.entry, { [name]: obj })
}

async function runSummary(id) {
  const it = S.byId.get(id)
  if (!it || it.locked || S.busy[id]) return
  S.busy[id] = '要約中…'
  render()
  try {
    const t = await file(it, 't')
    if (!t?.segments?.length) throw new Error('文字起こしがありません。')
    const [prevS, prevN] = await Promise.all([file(it, 's'), file(it, 'n')])
    let s = await summarizeWithAI(t.segments, (n) => {
      S.busy[id] = `要約中…（${n}字）`
      $$(`[data-busy="${id}"]`).forEach((el) => { el.textContent = S.busy[id] })
    })
    // 再要約：前の ToDo と文言の近さで突き合わせ、完了と経過を引き継ぐ
    if (prevS?.todos?.length) s = { ...s, todos: mergeTodos(prevS.todos, s.todos, prevN) }
    delete S.busy[id]
    await saveFile(S.byId.get(id), 's', s, { st: 'ok' })
    if (marksOf(prevN).length) await mutateNotes(S.byId.get(id), (n) => { n.marks = reconcile(n.marks, summaryTexts(s), SUMMARY_KINDS) })
    toast(`「${it.title}」を要約しました。`)
  } catch (e) {
    delete S.busy[id]
    render()
    toast(`要約できませんでした：${e.message}`, true)
  }
}

async function sumSelected() {
  const ids = [...S.checked].filter((id) => S.byId.get(id)?.st === 'sw')
  S.checked.clear()
  for (const id of ids) await runSummary(id)
}


/** Claude.ai の要約ページの URL。会議から開いたときは、その会議の ID を #m-xxxx で渡す（そのページで、その会議だけを選ぶ） */
function claudeUrl(id = '') {
  const saved = ls.get('m2.claudeUrl', '')
  if (saved) return saved
  return CLAUDE_DEFAULT_URL + (id ? '#' + id.replace(/_/g, '-') : '')
}

/** Mac が Notion の作業DBに書き出した状態（worker.json）。古い Mac のときは null */
const notionState = () => (S.worker && S.worker.notion) || null

function openClaude(a) {
  const id = a?.dataset?.id || ''
  const ns = notionState()
  if (id) {
    // 会議から開いた：その会議が作業DBに無ければ、開いても別の会議しか出ないので先に伝える
    if (ns && !ns.ok) { toast('Mac が Notion の作業DBに書き出せていません：' + ns.err, true, 10000); return }
    if (ns && !(ns.ids || []).includes(id)) {
      toast('この会議は、まだ Notion の作業DBに書き出されていません。Mac の次の確認（5分以内）のあとに押してください。', true, 8000)
      return
    }
    window.open(claudeUrl(id), '_blank', 'noopener,noreferrer')
    toast('Claude.ai を開きました。要約が済むと、Mac の次の確認でここに反映されます。', false, 6000)
    return
  }
  const n = S.items.filter((x) => x.st === 'sw' && !x.locked).length
  if (!n) { setNote('要約待ちの議事録はありません。'); return }
  if (ns && !ns.ok) { setNote('Mac が Notion の作業DBに書き出せていません：' + ns.err); return }
  window.open(claudeUrl(), '_blank', 'noopener,noreferrer')
  const inDb = ns ? (ns.ids || []).length : null
  setNote(`Claude.ai を開きました（要約待ち ${n} 件${inDb != null && inDb !== n ? `、うち Notion の作業DBに書き出し済み ${inDb} 件` : ''}）。開かないときは、ブラウザのポップアップの許可を確かめてください。`
    + ' Claude.ai が要約を Notion の作業DBに書き込むと、Mac が次の確認（5分以内）で回収し、ここが「要約済み」になります。')
}

function setNote(text) {
  S.cnote = text
  const el = $('#cnote')
  if (el) { el.textContent = text; el.hidden = !text } else toast(text, false, 8000)
}

/* ---------- チャット ---------- */
async function sendChat(id, text) {
  const it = S.byId.get(id)
  const conn = connectionOf(loadSettings())
  const msgs = S.chat[id] ||= []
  msgs.push({ role: 'user', content: text })
  if (!conn) { msgs.push({ role: 'assistant', content: 'AI の接続先が未設定です。設定で URL とモデルを登録してください。' }); paintDetail(); return }
  const [s, t] = await Promise.all([file(it, 's'), file(it, 't')])
  const sys = `あなたは会議の議事録について答えるアシスタントです。次の要約と文字起こしだけを根拠に、日本語で簡潔に答えてください。根拠になった発言の時刻を [m:ss] で添えてください。分からないことは分からないと答えてください。\n\n# 会議\n${it.title}（${it.date}）\n\n# 要約\n${s ? JSON.stringify({ summary: s.summary, decisions: s.decisions, agenda: s.agenda, todos: s.todos }) : 'なし'}\n\n# 文字起こし\n${transcriptText(t?.segments).slice(0, 24000)}`
  const out = { role: 'assistant', content: '' }
  msgs.push(out)
  paintDetail()
  try {
    for await (const ch of streamChat(conn, [{ role: 'system', content: sys }, ...msgs.slice(0, -1)])) {
      if (ch.delta) {
        out.content += ch.delta
        const last = $('#msgs .msg-b.assistant:last-child')
        if (last && S.sel === id) last.textContent = out.content
      }
    }
  } catch (e) {
    out.content = out.content || `答えられませんでした：${e.message}`
  }
  if (S.sel === id) paintDetail()
}

/* ---------- ノート（メモ・マーカー・ToDo の完了と経過）の書き込み ---------- */
// n.json = { memo, done: {key: bool}, logs: {key: [{d, x}]}, marks: [...] }
const normN = (n) => {
  const o = n || {}
  o.marks = Array.isArray(o.marks) ? o.marks : []
  o.logs = o.logs && typeof o.logs === 'object' ? o.logs : {}
  o.done = o.done && typeof o.done === 'object' ? o.done : {}
  return o
}
const hasN = (n) => !!(n.memo || n.marks.length || Object.keys(n.logs).length || Object.values(n.done).some(Boolean) || n.todoLogs)

/** 今の画面を描き直す（スクロール位置は保つ） */
function repaintBodies() {
  if (S.tab === 'list') paintDetail()
  else if (S.tab === 'todo' || S.tab === 'notes') renderMain()
}

/**
 * ノートを書き換える。まず手元の表示にすぐ反映し、そのあと GitHub の最新を読み直して同じ変更を当てて保存する
 * （別の人が同時に書いていても、その人の変更を消さないため）。同じ会議の書き込みは順番に行う。
 */
async function mutateNotes(it, fn, { quiet = false } = {}) {
  const id = it.id
  let painted = false
  if (!it.locked) {
    let c = S.files[id]
    if (!c || c.upd !== it.upd) c = S.files[id] = { upd: it.upd, p: {} }
    if (!it.nt || c.n !== undefined) {   // 今の中身が分かっているときだけ、先に画面へ反映する
      const copy = normN(JSON.parse(JSON.stringify((it.nt ? c.n : null) || {})))
      fn(copy)
      c.n = copy
      it.nt = 1
      const r = S.raw.find((x) => x.id === id); if (r) r.nt = 1
      if (!quiet) { repaintBodies(); painted = true }
    }
  }
  const run = async () => {
    const v = await S.R.json(`m/${id}/n.json`)
    const n = normN(v ? await openJson(DEK.get(id), v) : {})
    fn(n)
    const r = await gasS('put', { id, patch: { nt: hasN(n) ? 1 : 0 }, files: { n: await sealJson(DEK.get(id), n) } })
    // 先に画面へ反映してあるときは、保存の完了では描き直さない（続けて文字を選んでいるときに、選択が消えるため）
    await applyEntry(r.entry, { n }, false, quiet || painted)
    return n
  }
  const p = (S.wq[id] || Promise.resolve()).catch(() => {}).then(run)
  S.wq[id] = p.catch(() => {})
  try { return await p } catch (e) {
    if (S.files[id]) delete S.files[id].n   // 先走りの表示を捨てて、読み直す
    repaintBodies()
    toast(`保存できませんでした：${e.message}`, true)
    throw e
  }
}

async function toggleTodo(id, key, on) {
  try { await mutateNotes(S.byId.get(id), (n) => { n.done[key] = !!on }) } catch { /* mutateNotes が知らせて、表示を戻す */ }
}

async function appendLog(id, key, text) {
  const it = S.byId.get(id)
  const entry = { d: Date.now(), x: text }
  S.logOpen = null
  await mutateNotes(it, (n) => {
    const cur = n.logs[key] || []
    const last = cur[cur.length - 1]
    if (!(last && last.d === entry.d && last.x === entry.x)) n.logs[key] = [...cur, entry]
  }).catch(() => {})
}

/* ---------- マーカー（選んだ文字に色を付けて、ノートに集める） ---------- */
let PEND = null   // いま選んでいる範囲 { m: 会議, loc, s, e, text, tt?, id?（既存のマーカー）, fromMarker? }
const barEl = () => $('#mkbar')

function initBar() {
  const b = barEl()
  if (!b || b.dataset.ready) return
  b.dataset.ready = '1'
  b.innerHTML = `<span id="mkText" class="mk-text"></span>${MARK_TAGS.map((t) => `<button type="button" class="mkb" data-act="mk" data-c="${t.id}"><span class="mdot c${t.id}"></span>${esc(t.name)}</button>`).join('')}<button type="button" class="mkb erase" data-act="mk" data-c="0">消す</button><button type="button" class="ib" data-act="mk-close" aria-label="閉じる">×</button>`
  b.addEventListener('mousedown', (e) => e.preventDefault())   // ボタンを押しても、選択を外さない
}

function hideBar() { PEND = null; const b = barEl(); if (b) b.hidden = true }

function showBar(p) {
  initBar()
  PEND = p
  const b = barEl()
  b.hidden = false
  b.classList.toggle('existing', !!p.id)
  const t = p.text.slice(p.s, p.e).replace(/\s+/g, ' ')
  $('#mkText').textContent = p.id ? 'マーカー' : `「${t.length > 14 ? t.slice(0, 14) + '…' : t}」`
}

/** 画面で選んでいる文字 → { 会議, 場所, 開始, 終了 }。マーカーを付けられる文章の中だけ */
function readSelection() {
  const sel = window.getSelection && window.getSelection()
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null
  const r = sel.getRangeAt(0)
  const node = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement
  const root = node && node.closest && node.closest('.mk')
  if (!root) return null
  const upTo = (c, o) => { const x = document.createRange(); x.selectNodeContents(root); x.setEnd(c, o); return x.toString().length }
  const s = upTo(r.startContainer, r.startOffset)
  const e = root.contains(r.endContainer) ? upTo(r.endContainer, r.endOffset) : root.textContent.length   // 複数の文にまたがるときは、最初の文の終わりまで
  if (e <= s) return null
  const row = root.closest('.seg-row')
  const tt = row && row.dataset.t !== '' && row.dataset.t != null ? Number(row.dataset.t) : null
  return { m: root.dataset.m, loc: root.dataset.loc, s, e, text: root.textContent, tt }
}

async function applyMarker(c) {
  const p = PEND
  if (!p) return
  const it = S.byId.get(p.m)
  if (!it) return
  hideBar()
  if (window.getSelection) window.getSelection().removeAllRanges()
  try {
    await mutateNotes(it, (n) => {
      if (p.id) { n.marks = c ? recolor(n.marks, p.id, c) : n.marks.filter((x) => x.id !== p.id); return }
      const before = new Set(n.marks.map((x) => x.id))
      let marks = applyRange(n.marks, p.loc, p.s, p.e, c || null, p.text)
      if (c && p.tt != null) marks = marks.map((x) => (before.has(x.id) ? x : { ...x, tt: p.tt }))
      n.marks = marks
    })
  } catch { /* mutateNotes が知らせる */ }
}

/* ---------- 本文の見出しへ飛ぶ・読んでいる場所をタブに映す ---------- */
/** 選んでいるタブが、横にはみ出して見えないとき（スマホ）は、見える位置までタブの列を動かす */
function centerActiveTab() {
  const nav = $('.dtabs')
  const on = nav && nav.querySelector('.dtab.on')
  if (!on || nav.scrollWidth <= nav.clientWidth) return
  nav.scrollLeft = Math.max(0, on.offsetLeft - (nav.clientWidth - on.offsetWidth) / 2)
}

function setSec(key) {
  S.sec = key
  $$('.dtab[data-act="sec"]').forEach((b) => { const on = S.dtab === 'sum' && b.dataset.v === key; b.classList.toggle('on', on); b.setAttribute('aria-pressed', on) })
  centerActiveTab()
}

function scrollToSection(key, smooth = true) {
  const sc = $('.sheet-scroll')
  const el = sc && sc.querySelector(`[data-sec="${key}"]`)
  if (!el) return
  setSec(key)
  S.spyLock = performance.now() + (smooth ? 800 : 200)   // 動いている間は、通り過ぎる見出しにタブを追従させない
  const top = el.getBoundingClientRect().top - sc.getBoundingClientRect().top + sc.scrollTop - 6
  if (smooth && sc.scrollTo) sc.scrollTo({ top, behavior: 'smooth' }); else sc.scrollTop = top
}

function applySecJump() {
  if (!S.secJump) return
  const sc = $('.sheet-scroll')
  if (!sc || !sc.querySelector(`[data-sec="${S.secJump}"]`)) return
  const key = S.secJump
  S.secJump = null
  scrollToSection(key, false)
}

function goSection(key) {
  hideBar()
  if (S.dtab !== 'sum') {   // 文字起こし・チャット・マインドマップから戻るときは、本文を出してから飛ぶ
    captureEdit()
    S.dtab = 'sum'; S.sec = key; S.secJump = key; S.seek = null
    paintDetail(true)
    return
  }
  scrollToSection(key, true)
}

let spyRaf = 0
document.addEventListener('scroll', (e) => {
  const sc = e.target
  if (!(sc instanceof Element) || !sc.classList.contains('sheet-scroll') || S.dtab !== 'sum') return
  cancelAnimationFrame(spyRaf)
  spyRaf = requestAnimationFrame(() => {
    if (performance.now() < S.spyLock) return
    const secs = [...sc.querySelectorAll('[data-sec]')]
    if (!secs.length) return
    const line = sc.getBoundingClientRect().top + 28
    let cur = secs[0].dataset.sec
    for (const el of secs) if (el.getBoundingClientRect().top <= line) cur = el.dataset.sec
    if (sc.scrollTop > 0 && sc.scrollTop + sc.clientHeight >= sc.scrollHeight - 4) cur = secs[secs.length - 1].dataset.sec
    if (cur !== S.sec) setSec(cur)
  })
}, true)

/* ---------- 本文の手直し ---------- */
function captureEdit() {
  if (!S.edit) return
  if (S.edit.kind === 'sum' && $('#editor')) S.edit.draft = readDraft($('#editor'))
  if (S.edit.kind === 'tr' && $('#trEdit')) S.edit.text = $('#trEdit').value
}
function leaveEdit() {
  if (!S.edit) return true
  captureEdit()
  if (!confirm('編集中の内容はまだ保存されていません。破棄しますか？')) return false
  S.edit = null
  return true
}

function eAdd(a) {
  const root = $('#editor'); if (!root || !S.edit) return
  const d = readDraft(root)
  const list = a.dataset.list
  let focus = ''
  if (list === 'points') { const ai = Number(a.dataset.ai); d.agenda[ai].points.push(blank.points()); focus = `.eagenda[data-ai="${ai}"] .erow[data-list="points"][data-i="${d.agenda[ai].points.length - 1}"] [data-f="text"]` }
  else if (list === 'agenda') { d.agenda.push(blank.agenda()); focus = `.eagenda[data-ai="${d.agenda.length - 1}"] [data-f="topic"]` }
  else if (list === 'topics') { d.topics.push(blank.topics()); focus = `.erow[data-list="topics"][data-i="${d.topics.length - 1}"] [data-f="text"]` }
  else { d[list].push(blank[list]()); focus = `.erow[data-list="${list}"][data-i="${d[list].length - 1}"] [data-f="text"]` }
  S.edit.draft = d; S.edit.focus = focus
  paintDetail()
}
function eDel(a) {
  const root = $('#editor'); if (!root || !S.edit) return
  const d = readDraft(root)
  const list = a.dataset.list, i = Number(a.dataset.i)
  if (list === 'points') d.agenda[Number(a.dataset.ai)].points.splice(i, 1)
  else d[list].splice(i, 1)
  S.edit.draft = d
  paintDetail()
}

async function saveEditSum() {
  const ed = S.edit
  const it = ed && S.byId.get(ed.id)
  const root = $('#editor')
  if (!it || !root) return
  const draft = readDraft(root)
  ed.draft = draft
  const prev = cached(it, 's'), n = cached(it, 'n')
  const r = summaryFromDraft(draft, prev)
  const msg = $('#edMsg')
  if (r.err) { msg.textContent = r.err; return }
  const keep = new Set(r.s.todos.map((t) => t.id))
  const lost = (prev.todos || []).filter((t, i) => !keep.has(todoKey(t, i)) && logsOf(n, todoKey(t, i)).length)
  if (lost.length && !confirm(`次の ToDo には経過が記録されています。消すと、経過も見えなくなります。\n\n${lost.map((t) => '・' + t.text).join('\n')}\n\nこのまま保存しますか？`)) return
  msg.textContent = '保存しています…'
  $$('#editor button').forEach((b) => { b.disabled = true })
  try {
    await saveFile(it, 's', r.s)
    if (marksOf(n).length) await mutateNotes(S.byId.get(ed.id), (nn) => { nn.marks = reconcile(nn.marks, summaryTexts(r.s), SUMMARY_KINDS) })
    S.edit = null
    paintDetail()
    toast('本文を保存しました。')
  } catch (e) {
    const m2 = $('#edMsg'); if (m2) m2.textContent = e.message
    $$('#editor button').forEach((b) => { b.disabled = false })
  }
}

async function saveEditTr() {
  const ed = S.edit
  const it = ed && S.byId.get(ed.id)
  const ta = $('#trEdit')
  if (!it || !ta) return
  ed.text = ta.value
  const t = cached(it, 't'), n = cached(it, 'n')
  const segs = textToSegments(ta.value, t.segments)
  const msg = $('#edMsg')
  if (!segs.length) { msg.textContent = '文字起こしを空にはできません。'; return }
  msg.textContent = '保存しています…'
  $$('.treditor, #main .acts button').forEach((b) => { b.disabled = true })
  try {
    await saveFile(it, 't', { ...t, segments: segs, edited: Date.now() })
    if (marksOf(n).some((x) => kindOf(x.loc) === 'tr')) await mutateNotes(S.byId.get(ed.id), (nn) => { nn.marks = withTimes(reconcile(nn.marks, trTexts(segs), ['tr']), segs) })
    S.edit = null
    paintDetail()
    toast('文字起こしを保存しました。')
  } catch (e) {
    const m2 = $('#edMsg'); if (m2) m2.textContent = e.message
    $$('.treditor, #main .acts button').forEach((b) => { b.disabled = false })
  }
}

/* ---------- 登録（音声・文字起こし） ---------- */
const CHUNK = 3 * 1024 * 1024
const blobB64 = (blob) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = () => rej(new Error('ファイルを読めませんでした')); r.readAsDataURL(blob) })
const audioDuration = (f) => new Promise((res) => {
  const a = new Audio(); const u = URL.createObjectURL(f)
  const done = (v) => { URL.revokeObjectURL(u); res(v) }
  a.preload = 'metadata'; a.onloadedmetadata = () => done(isFinite(a.duration) ? Math.round(a.duration) : 0); a.onerror = () => done(0)
  setTimeout(() => done(0), 8000); a.src = u
})

function groupsPicker(selected) {
  if (!S.admin) {
    return `<div class="chips"><span class="gchip mine">${esc(S.role)}</span><span class="gchip">管理者</span></div>
      <p class="muted sm">ログイン中のグループで登録されます（変更できません）。ほかのグループへの公開は管理者があとで設定します。</p>`
  }
  const roles = Object.keys(S.roles).filter((r) => r !== ADMIN_ROLE)
  return `<div class="chips">${roles.map((r) => `<label class="gpick"><input type="checkbox" name="grp" value="${esc(r)}" ${selected.includes(r) ? 'checked' : ''}>${esc(r)}</label>`).join('')}<span class="gchip">管理者（常に閲覧可）</span></div>
    <p class="muted sm">選んだグループと管理者が閲覧できます。選ばなければ管理者だけが閲覧できます。</p>`
}

function openUpload() {
  modal('議事録を登録', `
    <div class="kinds" role="group" aria-label="登録するもの">
      <button class="kind on" data-act="kind" data-v="audio" aria-pressed="true"><span class="strong">① 音声ファイル</span><span class="muted sm">Mac が文字起こしします</span></button>
      <button class="kind" data-act="kind" data-v="text" aria-pressed="false"><span class="strong">② 文字起こしデータ</span><span class="muted sm">要約テキストも一緒に登録できます（任意）</span></button>
    </div>
    <div class="frow">
      <div class="fcol grow2"><label for="upTitle">タイトル</label><input id="upTitle" placeholder="空のときはファイル名を使います"></div>
      <div class="fcol"><label for="upDate">開始日時</label><input id="upDate" type="datetime-local" value="${nowLocal()}"></div>
    </div>
    <div class="fcol"><label for="upTags">タグ（読点・カンマ区切り）</label><input id="upTags" placeholder="定例、営業"></div>
    <div class="k-audio">
      <span class="lbl">音声ファイル</span>
      <label class="drop" id="upDrop"><input id="upFile" type="file" accept="audio/*,.m4a,.mp3,.wav,.aac,.ogg,.webm" class="sr">
        <span>ここにドロップ、または <span class="linkish">ファイルを選ぶ</span></span><span class="muted sm">m4a・mp3・wav など</span></label>
      <div id="upFileRow" class="filerow" hidden></div>
    </div>
    <div class="k-text" hidden>
      <span class="lbl">会議時間</span>
      <div class="hms"><label class="sr" for="upH">時間</label><input id="upH" inputmode="numeric" placeholder="1">時間<label class="sr" for="upM">分</label><input id="upM" inputmode="numeric" placeholder="0">分<label class="sr" for="upS">秒</label><input id="upS" inputmode="numeric" placeholder="0">秒</div>
      <div class="lblrow"><label for="upRaw">文字起こし（原文）</label><span class="req">必須</span><span class="grow"></span><span class="muted sm" id="upRawN">0 文字</span></div>
      <textarea id="upRaw" rows="6" placeholder="発言をそのまま貼り付け。行頭の [00:01:12] は時刻として読みます。テキストファイルをドロップしても読み込めます"></textarea>
      <div class="lblrow"><label for="upSum">要約テキスト</label><span class="opt">任意</span></div>
      <textarea id="upSum" rows="5" placeholder="議事メモ（タイプ1）・会議のメモ:（タイプ2）・会議メモ＋議題の見出し（タイプ3）の要約を貼り付け"></textarea>
      <p class="sm" id="upSumMsg"></p>
    </div>
    <div class="panel grp"><span class="strong">${LOCK_SVG} 閲覧できる人</span>${groupsPicker([])}</div>
    <div class="dlg-foot"><span class="grow sm" id="upState">登録後の状態：<b>文字起こし待ち</b> ・ Mac が次の確認で文字起こしします</span>
      <button class="btn" data-act="close">キャンセル</button><button class="btn main" data-act="up-go" id="upGo">アップロード</button></div>
    <div class="progress" id="upProg" hidden><span></span></div>
    <p class="msg" id="upMsg" role="alert"></p>`, { wide: true, sticky: true })
  U.kind = 'audio'; U.file = null; U.dur = 0
  const f = $('#upFile')
  f.addEventListener('change', () => setUpFile(f.files[0]))
  const drop = $('#upDrop')
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over') })
  drop.addEventListener('dragleave', () => drop.classList.remove('over'))
  drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); setUpFile(e.dataTransfer.files[0]) })
  const raw = $('#upRaw')
  raw.addEventListener('input', () => { $('#upRawN').textContent = `${raw.value.length} 文字` })
  raw.addEventListener('dragover', (e) => e.preventDefault())
  raw.addEventListener('drop', async (e) => { e.preventDefault(); const fl = e.dataTransfer.files[0]; if (fl) { raw.value = await fl.text(); raw.dispatchEvent(new Event('input')) } })
  $('#upSum').addEventListener('input', paintUpState)
}
const U = { kind: 'audio', file: null, dur: 0 }

async function setUpFile(f) {
  if (!f) return
  U.file = f
  U.dur = await audioDuration(f)
  const row = $('#upFileRow')
  row.hidden = false
  row.innerHTML = `<span class="strong grow">${esc(f.name)}</span><span class="muted sm">${U.dur ? fmtT(U.dur) + ' ・ ' : ''}${(f.size / 1048576).toFixed(1)}MB</span><button class="ib" data-act="up-clear" aria-label="ファイルを外す">×</button>`
}

function paintUpState() {
  const el = $('#upState'); if (!el) return
  if (U.kind === 'audio') { el.innerHTML = '登録後の状態：<b>文字起こし待ち</b> ・ Mac が次の確認で文字起こしします'; return }
  const txt = $('#upSum').value.trim()
  const msg = $('#upSumMsg')
  if (!txt) { msg.textContent = ''; el.innerHTML = '登録後の状態：<b>要約待ち</b> ・ 要約テキストを入れると要約済みで登録します'; return }
  try {
    const r = parseSummaryText(txt)
    msg.className = 'sm okmsg'
    msg.textContent = `${summaryTypeLabel(r.type)}として読み取りました：議題 ${r.agenda.length} 件・決定事項 ${r.decisions.length} 件・ToDo ${r.todos.length} 件・論点 ${r.topics.length} 件`
    el.innerHTML = '登録後の状態：<b class="okc">要約済み</b>'
  } catch (e) {
    msg.className = 'sm errmsg'; msg.textContent = e.message
    el.innerHTML = '登録後の状態：<b>要約テキストを読み取れません</b>'
  }
}

async function submitUpload() {
  const msg = $('#upMsg'); msg.textContent = ''
  const go = $('#upGo')
  const title0 = $('#upTitle').value.trim()
  const date = $('#upDate').value
  const tags = $('#upTags').value.split(/[、,，]/).map((s) => s.trim()).filter(Boolean)
  if (!date) { msg.textContent = '開始日時を入れてください。'; return }
  const roles = S.admin ? $$('input[name="grp"]:checked').map((x) => x.value) : []
  const gids = [...new Set([S.adminGid, ...(S.admin ? roles.map((r) => S.roles[r]) : [S.gid])].filter(Boolean))]
  for (const g of gids) if (!S.pub[g]) { msg.textContent = 'グループの公開鍵が見つかりません（管理者に連絡してください）。'; return }
  const id = newId()
  const dek = newDek()
  const k = {}
  go.disabled = true
  try {
    for (const g of gids) k[g] = await wrapDek(dek, S.pub[g])
    if (U.kind === 'audio') {
      const f = U.file
      if (!f) throw new Error('音声ファイルを選んでください。')
      if (!f.size) throw new Error('空のファイルは登録できません。')
      const ext = ((f.name.match(/\.([a-z0-9]{1,5})$/i) || [])[1] || 'm4a').toLowerCase()
      const fname = `${id}.${ext}`
      const prog = $('#upProg'); prog.hidden = false
      msg.textContent = 'アップロードしています…'
      const { sessionUrl } = await gasS('initUpload', { filename: fname, mimeType: f.type || 'application/octet-stream', size: f.size })
      let off = 0, res = null
      while (off < f.size) {
        const blob = f.slice(off, off + CHUNK)
        res = await gasS('putChunk', { sessionUrl, chunk: await blobB64(blob), offset: off, total: f.size })
        off += blob.size
        prog.firstElementChild.style.width = `${Math.round(off / f.size * 100)}%`
      }
      if (!res?.done) throw new Error('アップロードが最後まで終わりませんでした。もう一度試してください。')
      const meta = { title: title0 || f.name.replace(/\.[^.]+$/, ''), tags, file: { id: res.file.id, name: fname, orig: f.name } }
      msg.textContent = '登録しています…'
      const r = await gasS('put', { id, entry: { date, dur: U.dur, st: 'tw', src: 'audio', k, m: await sealJson(dek, meta) } })
      DEK.set(id, dek)
      closeModal()
      await applyEntry(r.entry)
      toast('アップロードしました。Mac が次の確認で文字起こしします。')
    } else {
      if (!title0) throw new Error('タイトルを入れてください。')
      const raw = $('#upRaw').value
      if (!raw.trim()) throw new Error('文字起こし（原文）を貼り付けてください。')
      const segments = segmentsFromText(raw)
      const hms = Number($('#upH').value || 0) * 3600 + Number($('#upM').value || 0) * 60 + Number($('#upS').value || 0)
      const lastT = [...segments].reverse().find((x) => x.t != null)?.t || 0
      const sumText = $('#upSum').value.trim()
      let s = null
      if (sumText) s = normalizeSummary(parseSummaryText(sumText), 'テキスト取り込み')
      const files = { t: await sealJson(dek, { segments }) }
      if (s) files.s = await sealJson(dek, s)
      msg.textContent = '登録しています…'
      const r = await gasS('put', { id, entry: { date, dur: hms || lastT, st: s ? 'ok' : 'sw', src: 'text', k, m: await sealJson(dek, { title: title0, tags }), td: s ? s.todos.length : undefined }, files })
      DEK.set(id, dek)
      closeModal()
      await applyEntry(r.entry, { t: { segments }, s, n: null })
      toast(s ? '要約済みで登録しました。' : '要約待ちで登録しました。変換管理から要約できます。')
    }
  } catch (e) {
    msg.textContent = e.message
    go.disabled = false
  }
}

/* ---------- 編集・閲覧グループ・削除（管理者／登録した人） ---------- */
function openEdit(id) {
  const it = S.byId.get(id)
  modal('タイトル・日時', `
    <div class="fcol"><label for="edTitle">タイトル</label><input id="edTitle" value="${esc(it.title)}"></div>
    <p class="muted sm">タグは、見出しの「＋ タグ」から付け外しします。</p>
    <div class="fcol"><label for="edDate">開始日時</label><input id="edDate" type="datetime-local" value="${esc(it.date)}"></div>
    <div class="dlg-foot"><span class="grow"></span><button class="btn" data-act="close">キャンセル</button><button class="btn main" data-act="edit-save" data-id="${id}">保存</button></div>
    <p class="msg" id="edMsg" role="alert"></p>`)
}

async function saveEdit(id) {
  const it = S.byId.get(id)
  const title = $('#edTitle').value.trim()
  if (!title) { $('#edMsg').textContent = 'タイトルを入れてください。'; return }
  const meta = { ...(META.get(id)?.v || {}), title }
  try {
    const r = await gasS('put', { id, patch: { m: await sealJson(DEK.get(id), meta), date: $('#edDate').value || it.date } })
    closeModal()
    await applyEntry(r.entry)
    toast('保存しました。')
  } catch (e) { $('#edMsg').textContent = e.message }
}

function openGroups(id) {
  const it = S.byId.get(id)
  const byGid = Object.fromEntries(Object.entries(S.roles).map(([r, g]) => [g, r]))
  const cur = Object.keys(it.k || {}).map((g) => byGid[g]).filter(Boolean)
  modal('閲覧グループ', `<p class="muted sm">「${esc(it.title)}」を見られるグループを選びます。管理者は常に閲覧できます。</p>
    ${groupsPicker(cur)}
    <div class="dlg-foot"><span class="grow"></span><button class="btn" data-act="close">キャンセル</button><button class="btn main" data-act="groups-save" data-id="${id}">保存</button></div>
    <p class="msg" id="grMsg" role="alert"></p>`)
}

async function saveGroups(id) {
  const it = S.byId.get(id)
  const roles = $$('input[name="grp"]:checked').map((x) => x.value)
  const want = [...new Set([S.adminGid, ...roles.map((r) => S.roles[r])].filter(Boolean))]
  const k = {}
  try {
    for (const g of want) {
      if (it.k[g]) k[g] = it.k[g]
      else { if (!S.pub[g]) throw new Error('公開鍵が見つからないグループがあります。'); k[g] = await wrapDek(DEK.get(id), S.pub[g]) }
    }
    const r = await gasS('put', { id, patch: { k } })
    closeModal()
    await applyEntry(r.entry)
    toast('閲覧グループを変えました。外したグループの端末に、変更前に読み込んだ内容が残っている場合があります。', false, 7000)
  } catch (e) { $('#grMsg').textContent = e.message }
}

async function del(id) {
  const it = S.byId.get(id)
  if (!confirm(`「${it.title}」を削除します。元に戻せません。よろしいですか？`)) return
  try {
    await gasS('put', { id, del: true })
    if (S.sel === id) S.sel = null
    await applyEntry(null, { id }, true)
    toast('削除しました。')
  } catch (e) { toast(e.message, true) }
}

async function retry(id) {
  try {
    const r = await gasS('put', { id, patch: { st: 'tw' } })
    await applyEntry(r.entry)
    toast('文字起こし待ちに戻しました。')
  } catch (e) { toast(e.message, true) }
}

/* ---------- 設定 ---------- */
// タブで分ける：表示／AI／ログイン／他端末へ設定共有。最後に開いたタブを覚えておく
const SET_TABS = [['view', '表示'], ['ai', 'AI'], ['login', 'ログイン'], ['share', '他端末へ設定共有']]
let setTab = 'view'

function openSettings(tab) {
  if (tab && SET_TABS.some(([k]) => k === tab)) setTab = tab
  const s = loadSettings()
  const c = activeConnection(s) || newConnection()
  const bio = bioOk()
  const auto = autoEnabled()
  const where = S.gh ? `${S.gh.owner}/${S.gh.repo}` : 'データ用リポジトリ'
  const pane = (k, html) => `<div class="spane" data-pane="${k}" role="tabpanel" aria-labelledby="stab-${k}" ${setTab === k ? '' : 'hidden'}>${html}</div>`
  modal('設定', `
    <div class="stabs" role="tablist" aria-label="設定の種類">${SET_TABS.map(([k, l]) => `<button type="button" role="tab" id="stab-${k}" class="stab ${setTab === k ? 'on' : ''}" data-act="stab" data-v="${k}" aria-selected="${setTab === k}">${l}</button>`).join('')}</div>

    ${pane('view', `
      <section class="sset"><h3>文字の大きさ</h3>
        <div class="fsset" id="fsset" role="group" aria-label="文字の大きさ">${FS_LEVELS.map((l, i) => `<button type="button" class="btn ${i === S.fs ? 'main' : ''}" data-act="fs-set" data-v="${i}" aria-pressed="${i === S.fs}">${l.n}</button>`).join('')}</div>
        <p class="muted sm">要約・文字起こし・ToDo・ノート・マインドマップの文字の大きさです。</p>
      </section>`)}

    ${pane('ai', `
      <section class="sset"><h3>要約・チャットに使う AI <span class="tag">個人の設定</span></h3>
        <div class="srow"><label for="stConn" class="sr">接続</label><select id="stConn">${s.connections.map((x) => `<option value="${esc(x.id)}" ${x.id === c.id ? 'selected' : ''}>${esc(x.label || '名前なし')}</option>`).join('')}</select><button type="button" class="btn" data-act="ai-add">＋ 接続を追加</button></div>
        <div class="sgrid">
          <div class="fcol"><label for="stLabel">名前</label><input id="stLabel" value="${esc(c.label)}"></div>
          <div class="fcol narrow"><label for="stCtx">num_ctx</label><input id="stCtx" inputmode="numeric" value="${esc(c.numCtx)}"></div>
          <div class="fcol full"><label for="stUrl">URL <span class="muted">（OpenAI 互換。例：https://xxxx.ngrok-free.app/v1）</span></label><input id="stUrl" value="${esc(c.baseUrl)}" autocapitalize="off" spellcheck="false"></div>
          <div class="fcol full"><label for="stKey">API キー</label><div class="srow"><input id="stKey" type="password" value="${esc(c.apiKey)}" autocomplete="off"><button type="button" class="btn" data-act="key-eye" aria-pressed="false">表示</button></div></div>
          <div class="fcol full"><label for="stModels">モデル <span class="muted">（カンマ区切り。先頭を使います）</span></label><input id="stModels" value="${esc((c.models || []).join(', '))}" autocapitalize="off" spellcheck="false"></div>
        </div>
        <div class="sfoot"><button class="btn main" data-act="ai-save">保存</button>${s.connections.length > 1 ? '<span class="grow"></span><button class="btn danger" data-act="ai-del">この接続を消す</button>' : ''}</div>
        <p class="muted sm">この端末にだけ保存します（グループで共有する設定には入れません）。自分の別の端末へは「他端末へ設定共有」で渡せます。gemma-chat・今の議事録と同じ設定です。</p>
      </section>
      <section class="sset"><h3>Claude.ai で要約するページ</h3>
        <div class="srow"><label for="stClaude" class="sr">リンク</label><input id="stClaude" value="${esc(ls.get('m2.claudeUrl', ''))}" placeholder="空欄なら標準のページ" autocapitalize="off" spellcheck="false"><button class="btn" data-act="claude-save">保存</button></div>
        <p class="muted sm">ふだんは空欄のままで構いません。別のページで要約したいときだけ、そのリンクを入れます。</p>
      </section>
      <p class="msg" id="stMsg" role="status"></p>`)}

    ${pane('login', `
      <section class="sset"><h3>権限コード</h3>
        <p class="sm">いまのグループ <span class="tag strong">${esc(S.admin ? '管理者' : S.role)}</span></p>
        <form id="codeForm" class="srow" autocomplete="off">
          <label for="stCode" class="sr">新しい権限コード</label>
          <input id="stCode" type="password" autocomplete="new-password" placeholder="新しい権限コード" required>
          <button class="btn" type="submit">変更する</button>
        </form>
        <p class="msg" id="codeMsg" role="status"></p>
      </section>
      <section class="sset"><h3>開き方</h3>
        <label class="chk"><input type="checkbox" id="stAuto" ${auto ? 'checked' : ''}> この端末では自動で開く</label>
        <p class="muted sm">${auto ? '権限コードを、この端末の中だけで取り出せる鍵で暗号化して覚えておきます。端末を人に使われる心配があるときはオフにしてください。'
          : '開くたびに、権限コード（または Face ID・Touch ID）を使います。'}</p>
        <div class="srow between"><span class="sm">Face ID・Touch ID：${!bio ? '<span class="muted">このブラウザでは使えません</span>' : bioRegistered() ? '登録済み' : '未登録'}</span>
          ${!bio ? '' : bioRegistered() ? '<button class="btn" data-act="bio-off">解除</button>' : '<button class="btn" data-act="bioReg">登録する</button>'}</div>
        ${bio && auto ? '<p class="muted sm">自動で開く間は使いません。鍵をかけたあとに開くときに使えます。</p>' : ''}
        <p class="msg" id="bioMsg"></p>
      </section>
      <div class="sfoot"><button class="btn" data-act="lock">鍵をかける</button><span class="grow"></span><button class="btn danger" data-act="reset">この端末の設定を消す</button></div>`)}

    ${pane('share', `
      <section class="sset">
        <p class="muted sm">別の端末を、権限コードを入れずに使えるようにします。渡すのは権限コードと、この端末の AI の接続先です。表示の設定（${esc(where)} の settings/${esc(S.gid || '')}.json に暗号化して保存）は、同じ権限コードで開いた端末に自動で反映されます。</p>
        <div id="devLink" class="dl" aria-live="polite"></div>
        <div id="dlPassBox" class="dlpass" hidden><span class="sm">パスフレーズ</span><span id="dlPassVal" class="dlpass-v"></span><span class="muted sm">受け取る端末で入力します。この画面を閉じると変わります。</span><span id="dlAiNote" class="sm bad-t"></span></div>
        <p class="hint">中身はパスフレーズで暗号化してあります。ホーム画面に追加した MinuteStock で受け取るときは、QR やリンクではなくコードを使ってください（Safari とホーム画面のアプリは保存場所が別のため）。</p>
      </section>
      <div class="sfoot"><button type="button" class="btn danger" data-act="forget-conn">この端末から接続情報を消去してロック</button></div>`)}
    `, { wide: true })
  $('#stConn').addEventListener('change', (e) => {
    const s2 = loadSettings(); s2.activeConnectionId = e.target.value
    const c2 = activeConnection(s2); s2.activeModel = c2?.models?.[0] ?? null
    saveSettings(s2); openSettings()
  })
  $('#codeForm').addEventListener('submit', (e) => { e.preventDefault(); changeCode() })
  mountShare()
}

/** 設定から権限コードを変える（この端末で使うコードを入れ替える。グループが変われば見られる会議も変わる） */
async function changeCode() {
  const code = $('#stCode')?.value.trim()
  const m = $('#codeMsg')
  if (!code || !m) return
  if (code === S.code) { m.textContent = 'いまと同じ権限コードです。'; return }
  if (S.edit) { m.textContent = '編集中の内容を保存するか取り消してから変えてください。'; return }
  m.textContent = '確認しています…'
  try { await gas('login', { code }) } catch (e) { m.textContent = e.message; return }   // 先に確かめて、違っていたら今のまま
  if (bioRegistered()) bioForget()   // Face ID が包んでいるのは古いコードなので外す
  clearData()
  gateShell('<p class="gate-sub">権限コードを切り替えています…</p><p class="gate-msg" id="gateMsg" role="status"></p>')
  if (await doLogin(code, 'change')) toast(`権限コードを変更しました（グループ：${S.admin ? '管理者' : S.role}）。`)
}

async function toggleAuto(on) {
  setAutoEnabled(on)
  try {
    if (on && S.code) await autoSave(S.code)
    if (!on) await autoForget()
  } catch (e) { toast('この端末に覚えられませんでした：' + e.message, true) }
  openSettings()
}

function saveAI() {
  const s = loadSettings()
  const c = activeConnection(s)
  const url = $('#stUrl').value.trim().replace(/\/+$/, '')
  if (url && !/^https:\/\//.test(url) && !/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//.test(url + '/')) { $('#stMsg').textContent = 'URL は https で始まるものにしてください（localhost は http も可）。'; return }
  Object.assign(c, {
    label: $('#stLabel').value.trim() || '接続', baseUrl: url, apiKey: $('#stKey').value.trim(),
    numCtx: Number($('#stCtx').value) || 32768, models: $('#stModels').value.split(',').map((x) => x.trim()).filter(Boolean),
  })
  if (!c.models.includes(s.activeModel)) s.activeModel = c.models[0] ?? null
  saveSettings(s)
  $('#stMsg').textContent = ''; toast('保存しました。')
  if (S.tab === 'convert') renderMain()
}

/* ---------- モーダル・トースト ---------- */
function modal(title, body, opt = {}) {
  $('#modal').innerHTML = `<div class="ov"><div class="dlg ${opt.wide ? 'wide' : ''}" role="dialog" aria-modal="true" aria-labelledby="dlgT">
    <div class="dlg-h"><h2 id="dlgT">${esc(title)}</h2><button class="ib x" data-act="close" aria-label="閉じる">×</button></div>
    <div class="dlg-b">${body}</div></div></div>`
  const first = $('#modal input, #modal textarea, #modal select, #modal button:not(.x)')
  first?.focus()
}
function closeModal() { dlPanel?.destroy(); dlPanel = null; dlShare = null; $('#modal').innerHTML = '' }

let toastTimer = 0
function toast(text, bad = false, ms = 4000) {
  const t = $('#toast')
  t.textContent = text
  t.className = bad ? 'show bad' : 'show'
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => { t.className = '' }, ms)
}

/* ---------- 操作 ---------- */
const ACT = {
  bio: async () => {
    try { const code = await bioUnlock(); await doLogin(code, 'bio') } catch (e) { showGate(e?.name === 'NotAllowedError' ? '取り消されました。権限コードでも開けます。' : e.message) }
  },
  bioReg: async () => {
    const m = $('#bioMsg')
    try { await bioRegister(S.code); if (m) m.textContent = ''; closeModal(); toast('登録しました。次から Face ID・Touch ID で開けます。') } catch (e) { if (m) m.textContent = e?.name === 'NotAllowedError' ? '取り消されました。' : e.message }
  },
  'bio-off': () => { bioForget(); openSettings() },
  stab: (a) => {
    setTab = a.dataset.v
    $$('.stab').forEach((b) => { const on = b.dataset.v === setTab; b.classList.toggle('on', on); b.setAttribute('aria-selected', on) })
    $$('.spane').forEach((p) => { p.hidden = p.dataset.pane !== setTab })
  },
  'key-eye': (a) => {
    const k = $('#stKey'); if (!k) return
    const show = k.type === 'password'
    k.type = show ? 'text' : 'password'; a.textContent = show ? '隠す' : '表示'; a.setAttribute('aria-pressed', show)
  },
  reset: async () => {
    if (!confirm('この端末の設定（覚えている権限コード・Face ID の登録・お気に入り・表示の設定）を消します。AI の接続先と、GitHub に保存した設定は残ります。よろしいですか？')) return
    clearData()   // 先に接続を切って、消した設定が GitHub に送られないようにする
    for (const k of ['m2.bio', 'm2.fav', 'm2.seen', 'm2.view', 'm2.bioAsked', 'm2.claudeUrl', 'm2.fs', 'm2.prefsUpd', 'm2.autoOff']) localStorage.removeItem(k)
    await autoForget()
    S.fav = new Set(); S.seen = {}; S.fs = 1; applyFs()
    lock()
  },
  'auto-open': async () => {
    const code = await autoLoad()
    if (code) await doLogin(code, 'auto'); else showGate('この端末に覚えていた権限コードを読めませんでした。権限コードで開いてください。')
  },
  recv: () => showReceive(),
  gate: () => showGate(),
  'code-change': changeCode,
  'forget-conn': async () => {
    if (!confirm('この端末から接続情報（覚えている権限コード・Face ID の登録）を消去してロックします。設定そのものは GitHub に残ります。')) return
    await autoForget()
    bioForget()
    lock()
  },
  lock,
  close: closeModal,
  tab: (a) => { if (!leaveEdit()) return; S.tab = a.dataset.tab; hideBar(); window.scrollTo(0, 0); scrollHide.reset(); render(true) },
  upload: openUpload,
  settings: openSettings,
  month: (a) => {
    const [y, m] = S.month.split('-').map(Number)
    S.month = monthKey(new Date(y, m - 1 + Number(a.dataset.d), 1))
    renderMain()
  },
  favonly: () => { S.favOnly = !S.favOnly; renderMain() },
  tags: (a) => { S.tagOpen = S.tagOpen === a.dataset.id ? null : a.dataset.id; paintDetail(); if (S.tagOpen) $('#tagNew')?.focus() },
  'tag-on': (a) => toggleTag(a.dataset.id, a.dataset.t, true),
  'tag-off': (a) => toggleTag(a.dataset.id, a.dataset.t, false),
  'tag-add': (a) => addTagFromInput(a.dataset.id),
  'mm-key': (a) => mmKey(a.dataset.k),
  'mm-fit': () => MM.mm?.fit?.(),
  'mm-reset': (a) => resetMap(a.dataset.id),
  view: (a) => { S.view = a.dataset.v; ls.set('m2.view', S.view); if (S.view === 'cards') S.sel = null; renderMain() },
  jump: (a) => { $(`#d-${a.dataset.day}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' }) },
  open: (a) => {
    const it = S.byId.get(a.dataset.id)
    if (!it) return
    if (S.edit && S.edit.id !== it.id && !leaveEdit()) return
    const tabChanged = S.tab !== 'list'
    S.sel = it.id
    S.seek = null
    S.jump = null
    S.tab = 'list'
    S.sec = 'summary'
    if (a.dataset.go === 'todo') { S.dtab = 'sum'; S.sec = 'todo'; S.secJump = 'todo' }
    else if (a.dataset.go) S.dtab = a.dataset.go
    if (!it.locked && S.seen[it.id] !== it.upd) { S.seen[it.id] = it.upd; ls.set('m2.seen', S.seen) }
    if (it.date && !it.date.startsWith(S.month)) S.month = it.date.slice(0, 7)
    hideBar()
    scrollHide.reset()
    render(tabChanged, true)
    if (matchMedia('(max-width: 820px)').matches) window.scrollTo(0, 0)
  },
  'close-detail': () => { if (!leaveEdit()) return; S.sel = null; hideBar(); scrollHide.reset(); renderMain() },
  dtab: (a) => { captureEdit(); S.dtab = a.dataset.v; S.seek = null; hideBar(); paintDetail(true) },
  sec: (a) => goSection(a.dataset.v),
  fs: (a) => setFs(S.fs + Number(a.dataset.d)),
  'fs-set': (a) => setFs(Number(a.dataset.v)),
  fav: (a) => {
    const id = a.dataset.id
    if (S.fav.has(id)) S.fav.delete(id); else S.fav.add(id)
    ls.set('m2.fav', [...S.fav]); renderMain()
  },
  summarize: (a) => runSummary(a.dataset.id),
  claude: openClaude,
  seek: (a) => {
    S.seek = { t: parseT(a.dataset.t), scrolled: false }
    S.dtab = 'tr'
    paintDetail()
  },
  'copy-tr': async (a) => {
    const t = cached(S.byId.get(a.dataset.id), 't')
    try { await navigator.clipboard.writeText(transcriptText(t?.segments)); toast('文字起こしをコピーしました。') } catch { toast('コピーできませんでした。', true) }
  },
  'memo-save': async (a) => {
    const it = S.byId.get(a.dataset.id)
    const memo = $('#memoIn').value
    $('#memoMsg').textContent = '保存しています…'
    try { await mutateNotes(it, (n) => { n.memo = memo }, { quiet: true }); toast('メモを保存しました。'); const m = $('#memoMsg'); if (m) m.textContent = '' } catch (e) { const m = $('#memoMsg'); if (m) m.textContent = e.message }
  },
  edit: (a) => openEdit(a.dataset.id),
  'edit-save': (a) => saveEdit(a.dataset.id),
  groups: (a) => openGroups(a.dataset.id),
  'groups-save': (a) => saveGroups(a.dataset.id),
  delete: (a) => del(a.dataset.id),
  retry: (a) => retry(a.dataset.id),
  cfilter: (a) => { S.cfilter = a.dataset.v; renderMain() },
  'check-sw': () => { S.items.filter((x) => x.st === 'sw' && !x.locked).forEach((x) => S.checked.add(x.id)); renderMain() },
  'sum-selected': sumSelected,
  kind: (a) => {
    U.kind = a.dataset.v
    $$('.kind').forEach((b) => { const on = b.dataset.v === U.kind; b.classList.toggle('on', on); b.setAttribute('aria-pressed', on) })
    $('.k-audio').hidden = U.kind !== 'audio'
    $('.k-text').hidden = U.kind !== 'text'
    $('#upTitle').placeholder = U.kind === 'audio' ? '空のときはファイル名を使います' : '会議名'
    $('#upGo').textContent = U.kind === 'audio' ? 'アップロード' : '登録'
    paintUpState()
  },
  'up-clear': () => { U.file = null; $('#upFile').value = ''; $('#upFileRow').hidden = true },
  'up-go': submitUpload,
  'ai-save': saveAI,
  'ai-add': () => {
    const s = loadSettings(); const c = newConnection({ label: '新しい接続' })
    s.connections.push(c); s.activeConnectionId = c.id; s.activeModel = null; saveSettings(s); openSettings()
  },
  'ai-del': () => {
    const s = loadSettings()
    s.connections = s.connections.filter((x) => x.id !== s.activeConnectionId)
    s.activeConnectionId = s.connections[0]?.id ?? null; s.activeModel = s.connections[0]?.models?.[0] ?? null
    saveSettings(s); openSettings()
  },
  'claude-save': () => {
    const u = $('#stClaude').value.trim()
    if (u) { try { const o = new URL(u); if (o.protocol !== 'https:' || !(o.hostname === 'claude.ai' || o.hostname.endsWith('.claude.ai'))) throw 0 } catch { $('#stMsg').textContent = 'claude.ai の https のリンクだけ使えます。'; return } }
    if (u) ls.set('m2.claudeUrl', u); else ls.del('m2.claudeUrl')
    $('#stMsg').textContent = ''; toast('保存しました。')
  },
}

Object.assign(ACT, {
  'edit-sum': (a) => {
    const it = S.byId.get(a.dataset.id), sm = it && cached(it, 's')
    if (!sm) return
    S.edit = { id: it.id, kind: 'sum', draft: draftFromSummary(sm), focus: '#edSum' }
    hideBar(); paintDetail()
  },
  'edit-tr': (a) => {
    const it = S.byId.get(a.dataset.id), t = it && cached(it, 't')
    if (!t) return
    S.edit = { id: it.id, kind: 'tr', text: trToText(t.segments), focus: '#trEdit' }
    hideBar(); paintDetail()
  },
  'e-cancel': () => { S.edit = null; paintDetail() },
  'e-add': eAdd,
  'e-del': eDel,
  'e-save': () => (S.edit && S.edit.kind === 'tr' ? saveEditTr() : saveEditSum()),
  'log-open': (a) => { S.logOpen = `${a.dataset.id}|${a.dataset.k}`; repaintBodies(); const i = $('#logIn'); if (i) i.focus() },
  'log-close': () => { S.logOpen = null; repaintBodies() },
  mk: (a) => applyMarker(Number(a.dataset.c)),
  'mk-close': () => { hideBar(); if (window.getSelection) window.getSelection().removeAllRanges() },
  tfilter: (a) => { S.tfilter = a.dataset.v; renderMain() },
  nfilter: (a) => { S.nfilter = Number(a.dataset.v); renderMain() },
  'note-open': (a) => {
    const it = S.byId.get(a.dataset.id)
    const mk = it && marksOf(cached(it, 'n')).find((x) => x.id === a.dataset.mid)
    if (!mk) return
    ACT.open({ dataset: { id: it.id, go: locTab(mk.loc) } })
    S.jump = { id: mk.id }
    paintDetail()
  },
  'mk-memo': (a) => { S.mkMemo = a.dataset.mid; renderMain(); const i = $('#mkMemoIn'); if (i) i.focus() },
  'mk-memo-close': () => { S.mkMemo = null; renderMain() },
  'mk-color': (a) => { mutateNotes(S.byId.get(a.dataset.id), (n) => { n.marks = recolor(n.marks, a.dataset.mid, Number(a.dataset.c)) }).catch(() => {}) },
  'mk-del': (a) => { mutateNotes(S.byId.get(a.dataset.id), (n) => { n.marks = n.marks.filter((x) => x.id !== a.dataset.mid) }).catch(() => {}) },
})

function saveMkMemo(id, mid, text) {
  S.mkMemo = null
  mutateNotes(S.byId.get(id), (n) => { n.marks = n.marks.map((x) => (x.id === mid ? { ...x, m: text || undefined } : x)) }).catch(() => {})
}

// マーカー：文字を選ぶと、下にバーが出る。付けたマーカーを押すと、色の変更・消去ができる
let selTimer = 0
document.addEventListener('selectionchange', () => {
  clearTimeout(selTimer)
  selTimer = setTimeout(() => {
    const p = readSelection()
    if (p) { showBar(p); return }
    if (PEND && !PEND.fromMarker) hideBar()
  }, 180)
})

document.addEventListener('click', (e) => {
  const mkEl = e.target.closest && e.target.closest('.marker')
  const sel = window.getSelection && window.getSelection()
  if (mkEl && (!sel || sel.isCollapsed)) {
    const root = mkEl.closest('.mk')
    const it = root && S.byId.get(root.dataset.m)
    const mk = it && marksOf(cached(it, 'n')).find((x) => x.id === mkEl.dataset.mid)
    if (mk) { showBar({ m: it.id, loc: mk.loc, s: mk.s, e: mk.e, text: root.textContent, id: mk.id, fromMarker: true }); return }
  }
  if (PEND && PEND.fromMarker && !(e.target.closest && e.target.closest('#mkbar'))) hideBar()
})

document.addEventListener('click', (e) => {
  const a = e.target.closest('[data-act]')
  if (!a || a.disabled) return
  const f = ACT[a.dataset.act]
  if (f) { e.preventDefault(); f(a, e) }
})

document.addEventListener('input', (e) => {
  if (e.target.id === 'q') {
    S.q = e.target.value
    clearTimeout(ACT._qt)
    ACT._qt = setTimeout(() => { if (S.tab === 'list') { renderMain(); const q = $('#q'); q?.focus() } }, 250)
  }
  if (e.target.id === 'nq') {
    S.nq = e.target.value
    clearTimeout(ACT._nt)
    ACT._nt = setTimeout(() => { if (S.tab === 'notes') { renderMain(); const q = $('#nq'); if (q) { q.focus(); q.setSelectionRange(q.value.length, q.value.length) } } }, 250)
  }
  if (e.target.id === 'trq') {
    S.trq = e.target.value
    clearTimeout(ACT._tt)
    ACT._tt = setTimeout(() => {
      paintDetail()
      const el = $('#trq'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length) }
      $('.seg-row.hit')?.scrollIntoView({ block: 'center' })
    }, 300)
  }
})

document.addEventListener('change', (e) => {
  const t = e.target
  if (t.dataset.chk) { if (t.checked) S.checked.add(t.dataset.chk); else S.checked.delete(t.dataset.chk); renderMain() }
  if (t.dataset.todo) { toggleTodo(t.dataset.todo, t.dataset.k, t.checked); t.closest('li')?.classList.toggle('done', t.checked) }
  if (t.id === 'stAuto') toggleAuto(t.checked)
  if (t.id === 'scopeSel') { S.scope = t.value; renderMain() }
  if (t.id === 'stSel') { S.filter = t.value; renderMain() }
  if (t.id === 'rangeSel') { S.range = t.value; ls.set('m2.range', S.range); renderMain() }
  if (t.id === 'towner') { S.towner = t.value; renderMain() }
  if (t.id === 'aiSel' && t.value.includes('|')) {
    const [cid, model] = t.value.split('|')
    const s = loadSettings(); s.activeConnectionId = cid; s.activeModel = model; saveSettings(s)
  }
})

document.addEventListener('submit', (e) => {
  const f = e.target
  if (f.matches && f.matches('form.editor')) { e.preventDefault(); return }
  if (f.matches && f.matches('form[data-tagform]')) { e.preventDefault(); return }   // 追加は送信ボタンの data-act で行う
  const lg = f.closest && f.closest('form[data-log]')
  if (lg) { e.preventDefault(); const v = $('#logIn').value.trim(); if (v) appendLog(lg.dataset.log, lg.dataset.k, v); return }
  const mm = f.closest && f.closest('form[data-mkmemo]')
  if (mm) { e.preventDefault(); saveMkMemo(mm.dataset.mkmemo, mm.dataset.mid, $('#mkMemoIn').value.trim()); return }
  const ch = f.closest && f.closest('form[data-chat]')
  if (!ch) return
  e.preventDefault()
  const text = $('#askIn').value.trim()
  if (text) sendChat(ch.dataset.chat, text)
})

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && PEND) { hideBar(); if (window.getSelection) window.getSelection().removeAllRanges() }
  if (e.key === 'Escape' && $('#modal').innerHTML && !$('#upGo')) closeModal()
  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && e.target.id === 'askIn') e.target.closest('form')?.requestSubmit()
})

document.addEventListener('visibilitychange', () => { if (!document.hidden && S.R) refresh() })

applyFs()
const scrollHide = initScrollHide($('#app'), matchMedia('(max-width: 1024px), (pointer: coarse)'))
boot()
