import { loadConfig } from './videos-config.js'
import { listVideos as gasListVideos, listIdeas as gasListIdeas } from './gas.js'

/**
 * 一覧・アイデア一覧の読み込み口。
 *
 * これまでは開くたびに GAS 経由で Notion を全件クエリしていて、件数が増えるほど
 * 最初の描画までが遅くなっていた。Mac 側の cron が書き出した静的JSONを先に読み、
 * 取れなかったときだけ従来どおり Notion に問い合わせる。
 * 詳細画面(タブごとの本文)は鮮度が要るので、これまでどおり Notion から取る。
 */

// 既定は同じリポジトリの data/clipstock/。設定で別の場所を指せる
const DEFAULT_BASE = new URL('../../../../data/clipstock/', import.meta.url).href

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
  return res.json()
}

export async function listVideos() {
  try {
    const json = await fetchJson('index.json')
    if (!Array.isArray(json.items)) throw new Error('index.json の形式が不正です')
    return { items: json.items, fetchedAt: json.generatedAt || null, source: 'json' }
  } catch (err) {
    console.warn('index.json を使えないため Notion から直接読み込みます:', err)
    const data = await gasListVideos()
    return { ...data, source: 'notion' }
  }
}

export async function listIdeas() {
  try {
    const json = await fetchJson('ideas.json')
    if (!Array.isArray(json.items)) throw new Error('ideas.json の形式が不正です')
    return { items: json.items, source: 'json' }
  } catch (err) {
    console.warn('ideas.json を使えないため Notion から直接読み込みます:', err)
    const data = await gasListIdeas()
    return { ...data, source: 'notion' }
  }
}
