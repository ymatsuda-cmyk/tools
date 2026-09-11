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
 * JSONは取り込み元ごとに分かれている(index-video / index-web、idea-video / idea-web)。
 * 元のNotionが別で、更新も別々に走るため、片方が古くても・落ちても
 * もう片方はそのまま出せるようにしている。
 */

// 既定は同じリポジトリの data/clipstock/。設定で別の場所を指せる
const DEFAULT_BASE = new URL('../../../../data/clipstock/', import.meta.url).href

// 取り込み元ごとのファイル。legacy は分割前の1本だけだった頃のファイル名
const LIST_FILES = [
  { name: 'index-video.json', source: 'video' },
  { name: 'index-web.json', source: 'web' },
]
const LIST_LEGACY = 'index.json'
const IDEA_FILES = [
  { name: 'idea-video.json', source: 'video' },
  { name: 'idea-web.json', source: 'web' },
]
const IDEA_LEGACY = 'ideas.json'

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

/** 取れなければ null。片方のファイルが無くても、もう片方だけで一覧を出すため */
async function fetchOptionalJson(name) {
  try {
    const json = await fetchJson(name)
    return Array.isArray(json.items) ? json : null
  } catch {
    return null
  }
}

/**
 * 取り込み元ごとのJSONを1つにまとめる。
 * 分割前のファイルには web の分も入っているので、キーが重複したら先勝ちで落とす。
 */
function mergeParts(parts) {
  const byKey = new Map()
  parts.forEach(({ json, source }) => {
    if (!json) return
    json.items.forEach((item) => {
      if (!item || !item.key || byKey.has(item.key)) return
      byKey.set(item.key, { ...item, source: item.source || source })
    })
  })
  return [...byKey.values()]
}

/** 分割後のファイルを読む。どちらも無ければ分割前の1本を読む(無ければ例外) */
async function loadParts(files, legacyName) {
  const parts = await Promise.all(
    files.map(async (f) => ({ ...f, json: await fetchOptionalJson(f.name) }))
  )
  if (parts.some((p) => p.json)) {
    return { items: mergeParts(parts), generatedAt: (parts.find((p) => p.json).json || {}).generatedAt || null }
  }
  const legacy = await fetchJson(legacyName)
  if (!Array.isArray(legacy.items)) throw new Error(`${legacyName} の形式が不正です`)
  return { items: mergeParts([{ json: legacy, source: 'video' }]), generatedAt: legacy.generatedAt || null }
}

export async function listVideos() {
  try {
    const { items, generatedAt } = await loadParts(LIST_FILES, LIST_LEGACY)
    items.sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')))
    registerPageSources(items)
    return { items, fetchedAt: generatedAt, source: 'json' }
  } catch (err) {
    console.warn('一覧JSONを使えないため Notion から直接読み込みます:', err)
    const data = await gasListVideos()
    registerPageSources(data.items)
    return { ...data, source: 'notion' }
  }
}

export async function listIdeas() {
  try {
    const { items } = await loadParts(IDEA_FILES, IDEA_LEGACY)
    registerPageSources(items)
    return { items, source: 'json' }
  } catch (err) {
    console.warn('アイデアJSONを使えないため Notion から直接読み込みます:', err)
    const data = await gasListIdeas()
    registerPageSources(data.items)
    return { ...data, source: 'notion' }
  }
}
