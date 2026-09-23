/**
 * 顧客一覧をまとめて取るAPI（customer9ページ用）。
 *
 * 取得元(プロバイダ)は差し替えられる。既定は GAS（Google Apps Script Web App）で、
 * 未設定でも確認できるように JSON（組み込みサンプル）も用意している。
 * providers/ に同じ形のモジュールを足して LOADERS に並べれば取得元を増やせる。
 *
 * 使い方:
 *   const customer9 = await import('/api/customer9/customer9.api.js')
 *   const config = customer9.normalizeConfig({ provider: 'gas', url: 'https://script.google.com/macros/s/xxx/exec' })
 *   const { customers } = await customer9.list(config)
 */

const LOADERS = {
  gas: () => import('./providers/gas.js'),
  json: () => import('./providers/json.js')
}

const loaded = new Map()

function loadProvider(providerId) {
  const key = providerId || 'json'
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

/** 設定に書き漏らした項目を埋める。以降の処理はこの形だけを見る */
export function normalizeConfig(raw) {
  const c = raw && typeof raw === 'object' ? raw : {}
  return {
    ...c,
    provider: c.provider || 'json',
    url: c.url || ''
  }
}

/** 顧客一覧を取得する */
export async function list(config) {
  const conf = normalizeConfig(config)
  const p = await loadProvider(conf.provider)
  const customers = await p.list(conf)
  return { provider: conf.provider, customers }
}
