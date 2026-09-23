/**
 * JSONで決め打ちした顧客一覧を出す取得元。
 *
 * GAS未設定のときの確認用。config.url を書けば毎回そこから取得し（同一オリジンかCORS許可が必要）、
 * 空欄なら config.customers（無ければ組み込みのサンプル）をそのまま使う。
 */

export const id = 'json'
export const label = 'JSON（固定の顧客データ・確認用）'

export const FIELDS = [
  {
    key: 'url',
    label: 'JSONのURL（任意）',
    placeholder: 'https://example.com/customers.json',
    hint: '空欄なら組み込みのサンプルデータを使います'
  }
]

/** GAS未接続でも画面を確認できるようにするための組み込みサンプル */
const SAMPLE_CUSTOMERS = [
  { code: 'C001', name: '株式会社サンプル商事', email: 'contact@sample-shoji.example', phone: '03-0000-0001', status: '取引中' },
  { code: 'C002', name: '有限会社テスト工業', email: 'info@test-kogyo.example', phone: '03-0000-0002', status: '取引中' },
  { code: 'C003', name: '合同会社デモシステムズ', email: 'sales@demo-systems.example', phone: '03-0000-0003', status: '休止中' }
]

export async function list(config) {
  const url = String(config && config.url || '').trim()
  const raw = url ? await fetchList(url) : Array.isArray(config && config.customers) ? config.customers : SAMPLE_CUSTOMERS
  return raw.map((c, i) => ({
    code: String(c.code || c.id || `C${String(i + 1).padStart(3, '0')}`),
    name: String(c.name || c.customerName || '（名称未設定）'),
    email: String(c.email || ''),
    phone: String(c.phone || c.tel || ''),
    status: String(c.status || '')
  }))
}

async function fetchList(url) {
  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`顧客一覧のJSONを取得できません（HTTP ${res.status}）`)
  const body = await res.json()
  return Array.isArray(body) ? body : Array.isArray(body.customers) ? body.customers : Array.isArray(body.items) ? body.items : []
}
