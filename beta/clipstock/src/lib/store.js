import { loadConfig } from './videos-config.js'
import { listVideos as gasListVideos, listIdeas as gasListIdeas, registerPageSources } from './gas.js'

/**
 * 一覧・アイデア一覧の読み込み口。
 *
 * これまでは開くたびに GAS 経由で Notion を全件クエリしていて、件数が増えるほど
 * 最初の描画までが遅くなっていた。Mac 側の cron が書き出した静的JSONを先に読み、
 * 取れなかったときだけ従来どおり Notion に問い合わせる。
 * 詳細画面(タブごとの本文)は鮮度が要るので、これまでどおり Notion から取る。
 *
 * 一覧は取り込み元ごとにファイルが分かれている(movie.json / web.json)。
 * 元のNotionが別で、更新も別々に走るため、片方が古くても・落ちても
 * もう片方はそのまま出せるようにしている。
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

/** 取れなければ null。web.json のように「無くても一覧は出す」ファイル用 */
async function fetchOptionalJson(name) {
  try {
    const json = await fetchJson(name)
    return Array.isArray(json.items) ? json : null
  } catch {
    return null
  }
}

/** 動画の一覧。movie.json を置くまでの間は旧 index.json を読む */
async function fetchMovieJson() {
  try {
    return await fetchJson('movie.json')
  } catch {
    return await fetchJson('index.json')
  }
}

/**
 * 取り込み元ごとのJSONを1つの一覧にする。
 * 旧 index.json には web の分も入っているので、キーが重複したら先勝ちで落とす。
 */
function mergeLists(sources) {
  const byKey = new Map()
  sources.forEach(({ json, source }) => {
    if (!json) return
    json.items.forEach((item) => {
      if (!item || !item.key || byKey.has(item.key)) return
      byKey.set(item.key, { ...item, source: item.source || source })
    })
  })
  return [...byKey.values()].sort((a, b) =>
    String(b.createdAt || '').localeCompare(String(a.createdAt || ''))
  )
}

export async function listVideos() {
  try {
    const [movie, web] = await Promise.all([fetchMovieJson(), fetchOptionalJson('web.json')])
    if (!Array.isArray(movie.items)) throw new Error('movie.json の形式が不正です')
    const items = mergeLists([
      { json: movie, source: 'video' },
      { json: web, source: 'web' },
    ])
    registerPageSources(items)
    return { items, fetchedAt: movie.generatedAt || null, source: 'json' }
  } catch (err) {
    console.warn('一覧JSONを使えないため Notion から直接読み込みます:', err)
    const data = await gasListVideos()
    registerPageSources(data.items)
    return { ...data, source: 'notion' }
  }
}

export async function listIdeas() {
  try {
    const json = await fetchJson('ideas.json')
    if (!Array.isArray(json.items)) throw new Error('ideas.json の形式が不正です')
    registerPageSources(json.items)
    return { items: json.items, source: 'json' }
  } catch (err) {
    console.warn('ideas.json を使えないため Notion から直接読み込みます:', err)
    const data = await gasListIdeas()
    registerPageSources(data.items)
    return { ...data, source: 'notion' }
  }
}
