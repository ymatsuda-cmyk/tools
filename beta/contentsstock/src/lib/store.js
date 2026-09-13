import { loadConfig } from './contents-config.js'
import { listContents as gasListContents, listIdeas as gasListIdeas } from './gas.js'

/**
 * 一覧・アイデア一覧の読み込み口。
 *
 * これまでは画面を開くたびに GAS 経由で Notion を全件クエリしていて、件数が増えるほど
 * 最初の描画までが遅くなっていた。Mac 側の cron(mac/scripts/contents/build_contentsstock_json.py)が
 * 書き出した静的JSONを先に読み、取れなかったときだけ従来どおり Notion に問い合わせる。
 *
 * 詳細画面(タブごとの本文)は鮮度が要るので、これまでどおり Notion から取る。
 * JSONは cron の周期ぶん古い。編集直後の見え方は画面側で反映しているので、
 * 「古いJSON + その場の反映」で足りる。
 */

// 既定は同じリポジトリの data/contentsstock/。設定で別の場所を指せる
const DEFAULT_BASE = new URL('../../../../data/contentsstock/', import.meta.url).href

const LIST_FILE = 'index-doc.json'
const IDEA_FILE = 'idea-doc.json'

function baseUrl() {
  const raw = String(loadConfig().dataUrl || '').trim()
  if (!raw) return DEFAULT_BASE
  return new URL(raw.endsWith('/') ? raw : raw + '/', location.href).href
}

async function fetchJson(name) {
  const url = new URL(name, baseUrl())
  // GitHub Pages / CDN のキャッシュに引っかかると更新が反映されないため毎回変える
  url.searchParams.set('t', Date.now())
  const res = await fetch(url, { cache: 'no-store' })
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  const json = await res.json()
  if (!Array.isArray(json.items)) throw new Error(`${name} の形式が不正です`)
  return json
}

export async function listContents() {
  try {
    const json = await fetchJson(LIST_FILE)
    const items = [...json.items].sort((a, b) =>
      String(b.createdAt || '').localeCompare(String(a.createdAt || ''))
    )
    return { items, fetchedAt: json.generatedAt || null, source: 'json' }
  } catch (err) {
    console.warn('一覧JSONを使えないため Notion から直接読み込みます:', err)
    const data = await gasListContents()
    return { ...data, source: 'notion' }
  }
}

export async function listIdeas() {
  try {
    const json = await fetchJson(IDEA_FILE)
    return { items: json.items, fetchedAt: json.generatedAt || null, source: 'json' }
  } catch (err) {
    console.warn('アイデアJSONを使えないため Notion から直接読み込みます:', err)
    const data = await gasListIdeas()
    return { ...data, source: 'notion' }
  }
}
