/**
 * 議事録（Notion）一覧・権限取得API
 *
 * 議事録アプリ(minutes/)が Notion に貯めている一覧を、ダッシュボードなど他の画面からも
 * 使えるようにする薄い入口。一覧は GitHub Pages に公開された index.json を読むだけで
 * 認証は要らない。権限（Notionの「権限」multi_select）の更新だけ、議事録用GAS
 * (minutes/gas/Code.gs) の savePermissions アクションを呼ぶ。
 *
 * 使い方:
 *   const minutes = await import('/api/minutes/minutes.api.js')
 *   const { items } = await minutes.recent({ limit: 10 })
 *   await minutes.savePermissions(
 *     { gasUrl: 'https://script.google.com/macros/s/.../exec', token: '...' },
 *     [items[0].notionPageId], ['jba'], 'add'
 *   )
 */

const INDEX_URL = 'https://ymatsuda-cmyk.github.io/tools/data/minutes/index.json'
const DEFAULT_LIMIT = 10

/**
 * 議事録DBの一覧を新着順（date降順）で取る。GASは呼ばない、認証不要の静的JSON。
 * config.indexUrl を指定すると、既定のURLの代わりにそこから読む（社内ミラーなど向け）。
 */
export async function recent(config = {}) {
  const url = (config && config.indexUrl) || INDEX_URL
  const limit = Math.max(1, Number(config && config.limit) || DEFAULT_LIMIT)

  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`議事録一覧を取得できません（HTTP ${res.status}）`)

  const raw = await res.json()
  if (!Array.isArray(raw)) throw new Error('議事録一覧の形式が正しくありません')

  const items = raw
    .slice()
    .sort(byDateDesc)
    .slice(0, limit)
  return { items }
}

function byDateDesc(a, b) {
  return new Date(b.date) - new Date(a.date)
}

/**
 * GAS(minutes/gas/Code.gs)の doPost を呼ぶ。
 * Content-Type は必ず text/plain にすること — application/json にすると
 * ブラウザが CORS preflight (OPTIONS) を送るが、GAS は OPTIONS に応答できず必ず失敗する。
 */
async function callGas(config, action, params = {}) {
  if (!config || !config.gasUrl || !config.token) {
    throw new Error('GASのURLと共有トークンが未設定です')
  }
  const res = await fetch(config.gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token: config.token, ...params })
  })
  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)

  const json = await res.json()
  if (!json.ok) throw new Error(json.error || 'GASがエラーを返しました')
  return json.data
}

/**
 * 複数ページの権限(Notionの「権限」multi_select)をまとめて更新する。
 * mode: 'add' 既存に追加 / 'remove' 指定分を除去 / 'replace' 置き換え（既定 'add'）
 */
export function savePermissions(config, pageIds, permissions, mode = 'add') {
  return callGas(config, 'savePermissions', { pageIds, permissions, mode })
}

/** ミーティング名(タイトル)を更新する */
export function saveTitle(config, pageId, title) {
  return callGas(config, 'saveTitle', { pageId, title })
}
