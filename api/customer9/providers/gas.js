/**
 * GAS（Google Apps Script）Web App から顧客一覧を取得する取得元。
 *
 * 事前にスプレッドシートを読む doGet を持つ GAS プロジェクトをデプロイし、
 * その `.../exec` の URL を config.url に設定して使う（README.md 参照）。
 * doGet は次の形の JSON を返す想定:
 *   { "customers": [ { "code":"C001", "name":"…", "email":"…", "phone":"…", "status":"…" } ] }
 */

export const id = 'gas'
export const label = 'GAS Web App（スプレッドシート連携）'

export const FIELDS = [
  {
    key: 'url',
    label: 'GAS Web App の URL',
    placeholder: 'https://script.google.com/macros/s/xxx/exec',
    hint: 'Apps Script を「ウェブアプリ」としてデプロイしたときに発行される exec のURL'
  }
]

export async function list(config) {
  const url = String(config && config.url || '').trim()
  if (!url) throw new Error('GAS Web App の URL が設定されていません')

  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`顧客一覧を取得できません（HTTP ${res.status}）`)

  const body = await res.json()
  return normalize(body)
}

/** レスポンスの包み方が多少違っても拾えるようにする */
function normalize(body) {
  const raw = Array.isArray(body) ? body
    : Array.isArray(body && body.customers) ? body.customers
    : Array.isArray(body && body.items) ? body.items
    : []

  return raw.map((c, i) => ({
    code: String(c.code || c.id || `C${String(i + 1).padStart(3, '0')}`),
    name: String(c.name || c.customerName || '（名称未設定）'),
    email: String(c.email || ''),
    phone: String(c.phone || c.tel || ''),
    status: String(c.status || '')
  }))
}
