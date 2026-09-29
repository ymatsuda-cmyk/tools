/**
 * WBSカンバン タスク状況API
 *
 * WBSカンバンアドイン（addin/kanban）が使っている Excel の wbs シートを、
 * Power Automate + Office Script（office-script/export-wbs.ts）が JSON にし、
 * Mac（mac/scripts/jsonbin/jsonbin_sync.py）が暗号化して JSONBin に置いたものを読む。
 * ダッシュボードなど他の画面から「遅延・本日〆・今週・来週」の件数と中身を出すための入口。
 *
 * JSONBin の中身は暗号文だけ（形式は README.md の「暗号化の形式」）。
 * 復号は Web Crypto（AES-256-GCM / PBKDF2-SHA256）でブラウザ内だけで行い、
 * パスフレーズはこのモジュールでは保存しない（保存するかどうかは呼び出し側が決める）。
 *
 * 使い方:
 *   const kanban = await import('/api/kanban/kanban.api.js')
 *   const data = await kanban.load({ binId, apiKey, keyType: 'access', passphrase })
 *   const rows = kanban.classify(data.tasks)          // 描画のたびに呼ぶ（日付で変わる）
 *   rows.late.rest / rows.today.rest / rows.today.done / rows.week.total ...
 *
 * 判定はカンバン（addin/kanban/kanban.js）と同じ列・同じ記号を使う。
 *   状態: 実績完了(S)あり→完了 / 備考(O)に▲→保留 / 実績開始(R)あり→対応中 / それ以外→未着手
 *   ★  : 備考の先頭が★
 */

export const SCHEMA = 'wbs-tasks/v1'
export const ENC = 'kanban-aesgcm/v1'

/* 行の並びと見出し。group が同じ行どうしでバーの長さの基準を揃える */
export const ROWS = [
  { key: 'late',  group: 1, label: '遅延',   total: false },
  { key: 'today', group: 1, label: '本日〆', total: true },
  { key: 'week',  group: 2, label: '今週',   total: true },
  { key: 'next',  group: 2, label: '来週',   total: true }
]

export const STATUS_LABELS = { todo: '未着手', doing: '対応中', held: '保留', done: '完了' }

/* ============================================================
   取得
   ============================================================ */

/**
 * JSONBin から読み、暗号文なら復号して { schema, updatedAt, tasks } を返す。
 *   config.binId      必須
 *   config.apiKey     任意。公開Binなら不要（中身は暗号文なので公開でも読まれて困らない）
 *   config.keyType    'access'（既定。X-Access-Key）/ 'master'（X-Master-Key）
 *   config.passphrase 暗号文のときに必須
 *   config.url        binId の代わりに直接URLを指定（テスト・ミラー用）
 * 失敗の理由は err.code で分かる: 'config' / 'fetch' / 'empty' / 'locked' / 'badpass' / 'format'
 */
export async function load(config = {}) {
  const url = config.url ||
    (config.binId ? `https://api.jsonbin.io/v3/b/${encodeURIComponent(config.binId)}/latest` : '')
  if (!url) throw fail('config', 'Bin ID が未設定です')

  const headers = { 'X-Bin-Meta': 'false' }
  if (config.apiKey) headers[config.keyType === 'master' ? 'X-Master-Key' : 'X-Access-Key'] = config.apiKey

  let res
  try {
    res = await fetch(url, { headers, cache: 'no-store' })
  } catch (err) {
    throw fail('fetch', `JSONBinに接続できません（${err.message || err}）`)
  }
  if (res.status === 404) throw fail('empty', 'まだ何も登録されていません')
  if (res.status === 401 || res.status === 403) throw fail('fetch', `JSONBinのキーが違うか、読み取り権限がありません（HTTP ${res.status}）`)
  if (!res.ok) throw fail('fetch', `JSONBinを取得できません（HTTP ${res.status}）`)

  let body = await res.json()
  if (body && body.record && !body.enc && !body.tasks) body = body.record  // X-Bin-Meta が効かなかったとき

  if (isEnvelope(body)) {
    if (!config.passphrase) throw fail('locked', 'パスフレーズを入力してください')
    body = await decrypt(body, config.passphrase)
  }
  return normalizeData(body)
}

function normalizeData(json) {
  if (typeof json === 'string') {
    try { json = JSON.parse(json) } catch { throw fail('format', 'JSONを読み取れません') }
  }
  if (Array.isArray(json)) json = { tasks: json }
  if (!json || !Array.isArray(json.tasks)) {
    throw fail('format', `タスクの一覧が見つかりません（schema ${SCHEMA} の tasks 配列が要ります）`)
  }
  return {
    schema: json.schema || '',
    updatedAt: json.updatedAt || null,
    tasks: json.tasks.map(normalizeTask).filter(t => t.title)
  }
}

function normalizeTask(t, i) {
  const note = String(t.note == null ? '' : t.note)
  const task = {
    id: t.id == null || t.id === '' ? `row-${t.row || i}` : String(t.id),
    title: String(t.title == null ? '' : t.title).trim(),
    category: String(t.category || ''),
    classification: String(t.classification || ''),
    user: t.user === '#' ? '' : String(t.user || ''),
    row: t.row || null,
    start: toDate(t.start),
    end: toDate(t.end),
    actualStart: toDate(t.actualStart),
    actualEnd: toDate(t.actualEnd),
    star: note.startsWith('★'),
    held: note.includes('▲')
  }
  task.status = task.actualEnd ? 'done' : task.held ? 'held' : task.actualStart ? 'doing' : 'todo'
  return task
}

/**
 * Excelのシリアル値・"2026-09-29"・ISO文字列をその日の 0:00（ブラウザの現地時刻）にする。
 * シリアル値と日付だけの文字列はUTCで作ってから年月日を取り出すので、時差で日付がずれない。
 */
export function toDate(v) {
  if (v === null || v === undefined || v === '') return null
  let d
  if (typeof v === 'number' || /^\d+(\.\d+)?$/.test(String(v))) {
    const u = new Date(Math.round((Number(v) - 25569) * 86400000))
    d = new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate())
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(String(v))) {
    const [y, m, day] = String(v).split('-').map(Number)
    d = new Date(y, m - 1, day)
  } else {
    d = new Date(v)
    if (!isNaN(d.getTime())) d.setHours(0, 0, 0, 0)
  }
  return d && !isNaN(d.getTime()) ? d : null
}

/* ============================================================
   分類
   ============================================================ */

/**
 * 4行ぶんに振り分ける。日付で結果が変わるので、描画のたびに呼ぶこと。
 *   late : 未完了で予定完了日が今日より前（件数のみ。総数は持たない）
 *   today: 予定完了日が今日（完了済みも総数に含む）
 *   week : 予定期間が今週（月〜日）に重なる（完了済みも総数に含む）
 *   next : 予定期間が来週に重なる（同上）
 * options.user を渡すとその担当者だけにする。
 * 戻り値: { late:{rest,done,total,from,to}, today:{...}, week:{...}, next:{...} }
 */
export function classify(tasks, options = {}) {
  const today = startOfDay(options.now ? new Date(options.now) : new Date())
  const mon = mondayOf(today)
  const sun = addDays(mon, 6)
  const nmon = addDays(mon, 7)
  const nsun = addDays(mon, 13)
  const list = options.user ? tasks.filter(t => t.user === options.user) : tasks

  const overlap = (t, a, b) => t.start && t.end && t.start <= b && t.end >= a
  const rules = {
    late:  { from: null, to: null, f: t => t.status !== 'done' && t.end && t.end < today },
    today: { from: today, to: today, f: t => sameDay(t.end, today) },
    week:  { from: mon, to: sun, f: t => overlap(t, mon, sun) },
    next:  { from: nmon, to: nsun, f: t => overlap(t, nmon, nsun) }
  }

  const out = {}
  ROWS.forEach(row => {
    const r = rules[row.key]
    const hit = list.filter(r.f)
    const rest = hit.filter(t => t.status !== 'done').sort(byEnd)
    const done = row.total ? hit.filter(t => t.status === 'done').sort(byActualEndDesc) : []
    out[row.key] = { rest, done, total: rest.length + done.length, from: r.from, to: r.to }
  })
  out.today.date = today
  return out
}

/** 期限切れの日数（未完了のみ）。期限切れでなければ 0 */
export function overdueDays(task, now = new Date()) {
  const today = startOfDay(new Date(now))
  if (task.status === 'done' || !task.end || task.end >= today) return 0
  return Math.round((today - task.end) / 86400000)
}

/** 表示中の担当者一覧（'#' と空は除く） */
export function users(tasks) {
  return [...new Set(tasks.map(t => t.user).filter(Boolean))]
}

export function formatMd(d) {
  return d ? `${d.getMonth() + 1}/${d.getDate()}` : ''
}

function startOfDay(d) { const t = new Date(d); t.setHours(0, 0, 0, 0); return t }
function addDays(d, n) { const t = new Date(d); t.setDate(t.getDate() + n); return t }
function mondayOf(d) { const t = startOfDay(d); const w = t.getDay(); t.setDate(t.getDate() - w + (w === 0 ? -6 : 1)); return t }
function sameDay(a, b) { return !!a && !!b && a.getTime() === b.getTime() }
function byEnd(a, b) { return (a.end ? a.end.getTime() : Infinity) - (b.end ? b.end.getTime() : Infinity) }
function byActualEndDesc(a, b) { return (b.actualEnd ? b.actualEnd.getTime() : 0) - (a.actualEnd ? a.actualEnd.getTime() : 0) }

/* ============================================================
   暗号化（AES-256-GCM / PBKDF2-SHA256）
   Mac側の jsonbin_sync.py と同じ形式。片方だけ変えると復号できなくなる
   ============================================================ */

const PBKDF2_ITER = 250000

export function isEnvelope(x) {
  return !!x && typeof x === 'object' && x.enc === ENC && typeof x.ct === 'string'
}

/** 暗号文 → 元のJSON（オブジェクト）。パスフレーズ違い・改ざんは err.code === 'badpass' */
export async function decrypt(envelope, passphrase) {
  if (!isEnvelope(envelope)) throw fail('format', '暗号化の形式が違います')
  if (envelope.kdf !== 'PBKDF2-SHA256') throw fail('format', `未対応の鍵導出です: ${envelope.kdf}`)
  const key = await deriveKey(passphrase, b64d(envelope.salt), Number(envelope.iter) || PBKDF2_ITER, ['decrypt'])
  let plain
  try {
    plain = await crypto.subtle.decrypt(
      { name: 'AES-GCM', iv: b64d(envelope.iv), additionalData: utf8(ENC), tagLength: 128 },
      key, b64d(envelope.ct))
  } catch {
    throw fail('badpass', 'パスフレーズが違うか、データが壊れています')
  }
  try {
    return JSON.parse(new TextDecoder().decode(plain))
  } catch {
    throw fail('format', '復号した中身がJSONではありません')
  }
}

/** 元のJSON → 暗号文。ブラウザ側から登録・テストするとき用（通常はMac側が暗号化する） */
export async function encrypt(json, passphrase, iter = PBKDF2_ITER) {
  const salt = crypto.getRandomValues(new Uint8Array(16))
  const iv = crypto.getRandomValues(new Uint8Array(12))
  const key = await deriveKey(passphrase, salt, iter, ['encrypt'])
  const ct = await crypto.subtle.encrypt(
    { name: 'AES-GCM', iv, additionalData: utf8(ENC), tagLength: 128 },
    key, utf8(JSON.stringify(json)))
  return { enc: ENC, kdf: 'PBKDF2-SHA256', iter, salt: b64e(salt), iv: b64e(iv), ct: b64e(new Uint8Array(ct)) }
}

async function deriveKey(passphrase, salt, iter, usages) {
  if (!crypto || !crypto.subtle) throw fail('format', 'このブラウザでは復号できません（https で開いてください）')
  const base = await crypto.subtle.importKey('raw', utf8(String(passphrase)), 'PBKDF2', false, ['deriveKey'])
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', hash: 'SHA-256', salt, iterations: iter },
    base, { name: 'AES-GCM', length: 256 }, false, usages)
}

function utf8(s) { return new TextEncoder().encode(s) }
function b64e(bytes) { let s = ''; bytes.forEach(b => { s += String.fromCharCode(b) }); return btoa(s) }
function b64d(str) { const s = atob(String(str)); const out = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i); return out }

function fail(code, message) {
  const err = new Error(message)
  err.code = code
  return err
}
