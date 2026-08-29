import { renderList, renderDetailHtml, renderToolbar, escapeHtml } from './ui/render.js'
import { fetchSummary, fetchTranscript, saveSummary, saveTags } from './lib/gas.js'
import { getDetailCache, setDetailCache, isCacheFresh } from './lib/cache.js'
import { generateSummary } from './lib/summarize.js'
import { loadConfig, saveConfig, isConfigured } from './lib/minutes-config.js'
import { loadSettings, saveSettings, newProfile } from './lib/llm-settings.js'
import { filterByMonth, filterBySearch, filterByTags, buildTagOptions } from './lib/filters.js'

const listEl = document.getElementById('list')
const listItemsEl = document.getElementById('list-items')
const toolbarEl = document.getElementById('toolbar')
const detailEl = document.getElementById('detail')
const syncStatusEl = document.getElementById('sync-status')

let items = []
let selectedKey = null
const tagsByKey = {} // pageId(notionPageId) -> string[]、タグ編集の楽観更新用

// --- 一覧の絞り込み状態 ---
let currentMonthKey = monthKeyOf(new Date()) // "YYYY-MM"
let searchQuery = ''
const selectedTags = new Set()
let showTags = false

function monthKeyOf(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`
}
function monthLabelOf(monthKey) {
  const [y, m] = monthKey.split('-')
  return `${y}年${Number(m)}月`
}
function shiftMonth(monthKey, delta) {
  const [y, m] = monthKey.split('-').map(Number)
  const d = new Date(y, m - 1 + delta, 1)
  return monthKeyOf(d)
}

const MOBILE_BREAKPOINT = 720

function isMobile() {
  return window.innerWidth <= MOBILE_BREAKPOINT
}

const INDEX_URL = 'https://ymatsuda-cmyk.github.io/tools/data/minutes/index.json'

async function loadIndex() {
  syncStatusEl.textContent = '読み込み中...'
  try {
    const res = await fetch(INDEX_URL, { cache: 'no-store' })
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    items = await res.json()
  } catch (err) {
    syncStatusEl.textContent = 'index.json の読み込みに失敗しました'
    items = []
  }
  refresh()
}

/**
 * フィルタ状態(月・検索・タグ)に基づいて一覧とツールバーを再描画する。
 * タグの選択可否は「月・検索を適用した後、AND追加しても0件にならないか」で判定する。
 */
function currentFilteredItems() {
  const byMonth = filterByMonth(items, currentMonthKey)
  const byMonthAndSearch = filterBySearch(byMonth, searchQuery)
  return filterByTags(byMonthAndSearch, selectedTags)
}

function refresh() {
  const byMonth = filterByMonth(items, currentMonthKey)
  const baseItems = filterBySearch(byMonth, searchQuery) // タグ絞り込み前(タグ候補の母集団)
  const filteredItems = filterByTags(baseItems, selectedTags)
  const tagOptions = buildTagOptions(baseItems, selectedTags)

  renderToolbar(toolbarEl, {
    monthLabel: monthLabelOf(currentMonthKey),
    query: searchQuery,
    showTags,
    tagOptions,
  }, {
    onPrevMonth: () => { currentMonthKey = shiftMonth(currentMonthKey, -1); refresh() },
    onNextMonth: () => { currentMonthKey = shiftMonth(currentMonthKey, 1); refresh() },
    onSearch: (q) => { searchQuery = q; refresh() },
    onToggleTag: (tag) => {
      selectedTags.has(tag) ? selectedTags.delete(tag) : selectedTags.add(tag)
      refresh()
    },
    onToggleShowTags: (v) => { showTags = v; refresh() },
  })

  renderList(listItemsEl, filteredItems, selectedKey, onSelect, showTags)
  syncStatusEl.textContent = `${filteredItems.length}件`
}

/**
 * 詳細を表示するターゲット要素を決める。
 * モバイルは選択した行の直後にインライン挿入、PCは右ペインに固定表示。
 */
function detailTarget(rowEl) {
  if (!isMobile()) {
    detailEl.classList.add('side-panel')
    return detailEl
  }
  // 既存のインライン展開を除去してから作り直す
  document.querySelectorAll('.inline-detail').forEach((el) => el.remove())
  const wrap = document.createElement('div')
  wrap.className = 'inline-detail'
  wrap.innerHTML = '<div class="detail-pane-inner"></div>'
  rowEl.after(wrap)
  return wrap.querySelector('.detail-pane-inner')
}

function paintDetail(target, item, state) {
  target.innerHTML = renderDetailHtml(item, { ...state, tags: tagsByKey[item.notionPageId] })
  const generateBtn = target.querySelector('.btn-generate, .btn-regenerate')
  generateBtn?.addEventListener('click', () => runGenerate(target, item))
  target.querySelector('.btn-retry')?.addEventListener('click', () => onSelect(item, findRow(item.key)))
  target.querySelector('.btn-raw')?.addEventListener('click', () => showRawTranscript(item))

  target.querySelectorAll('.tag-remove').forEach((el) => {
    el.addEventListener('click', (e) => {
      const tag = e.target.closest('.tag-chip').dataset.tag
      const next = (tagsByKey[item.notionPageId] || []).filter((t) => t !== tag)
      commitTags(target, item, state, next)
    })
  })
  target.querySelector('.tag-add-btn')?.addEventListener('click', () => {
    const input = prompt('追加するタグを入力してください')
    const tag = input?.trim()
    if (!tag) return
    const current = tagsByKey[item.notionPageId] || []
    if (current.includes(tag)) return
    commitTags(target, item, state, [...current, tag])
  })
}

async function commitTags(target, item, state, nextTags) {
  const prev = tagsByKey[item.notionPageId]
  tagsByKey[item.notionPageId] = nextTags // 楽観的に即反映
  paintDetail(target, item, state)
  try {
    await saveTags(item.notionPageId, nextTags)
  } catch (err) {
    tagsByKey[item.notionPageId] = prev // 失敗したら元に戻す
    paintDetail(target, item, state)
    alert('タグの保存に失敗しました: ' + (err.message || err))
  }
}

function findRow(key) {
  return listEl.querySelector(`.list-item[data-key="${key}"]`)
}

/** 一覧全体を作り直さず、指定アイテムの行のバッジ表示だけを更新する */
function updateRowBadge(item) {
  const row = findRow(item.key)
  const badge = row?.querySelector('.badge')
  if (!badge) return
  badge.textContent = item.status
  badge.className = `badge status-${item.status}`
}

async function onSelect(item, rowEl) {
  selectedKey = item.key
  renderList(listItemsEl, currentFilteredItems(), selectedKey, onSelect, showTags)
  const target = detailTarget(isMobile() ? findRow(item.key) : rowEl)

  paintDetail(target, item, { phase: 'loading' })

  const cache = getDetailCache(item.key)
  try {
    const remote = await fetchSummary(item.notionPageId)
    tagsByKey[item.notionPageId] = remote.tags || []

    if (!remote.generatedAt) {
      paintDetail(target, item, { phase: 'no-summary' })
      return
    }

    const fresh = isCacheFresh(cache, remote.generatedAt) ? cache : setDetailCache(item.key, remote)
    paintDetail(target, item, { phase: 'ready', summary: fresh })
  } catch (err) {
    // 通信に失敗してもキャッシュがあればそれを出す(タグは未確定のため編集UIは出さない)
    if (cache) {
      paintDetail(target, item, { phase: 'ready', summary: cache })
    } else {
      paintDetail(target, item, { phase: 'error', message: String(err.message || err) })
    }
  }
}

async function runGenerate(target, item) {
  paintDetail(target, item, { phase: 'generating' })
  try {
    const { text } = await fetchTranscript(item.notionPageId)
    const result = await generateSummary(text, (partial) => {
      paintDetail(target, item, { phase: 'generating', progress: partial.slice(0, 200) })
    })
    await saveSummary(item.notionPageId, result.cardSummary, {
      decisions: result.decisions,
      todos: result.todos,
      topics: result.topics,
    }, result.model)

    // Notion側の状態も"要約"に変わっているはずなので、画面側も合わせる
    item.status = '要約'
    updateRowBadge(item)

    const saved = setDetailCache(item.key, {
      cardSummary: result.cardSummary,
      detail: { decisions: result.decisions, todos: result.todos, topics: result.topics },
      model: result.model,
      generatedAt: new Date().toISOString(),
    })
    paintDetail(target, item, { phase: 'ready', summary: saved })
  } catch (err) {
    paintDetail(target, item, { phase: 'error', message: String(err.message || err) })
  }
}

// --- 原文表示モーダル ---
async function showRawTranscript(item) {
  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="raw-modal-overlay">
      <div class="raw-modal">
        <div class="raw-modal-header">
          <span>${escapeHtml(item.title)} — 原文</span>
          <button id="raw-close" class="btn-ghost" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
        </div>
        <div id="raw-body" class="raw-modal-body">
          <p style="color:var(--text-muted);font-size:13px">読み込み中...</p>
        </div>
      </div>
    </div>
  `
  document.getElementById('raw-close').addEventListener('click', () => (root.innerHTML = ''))

  try {
    const { text } = await fetchTranscript(item.notionPageId)
    document.getElementById('raw-body').innerHTML =
      `<pre class="raw-text">${escapeHtml(text || '(本文が空です)')}</pre>`
  } catch (err) {
    document.getElementById('raw-body').innerHTML =
      `<p class="error-text">${escapeHtml(String(err.message || err))}</p>`
  }
}

// --- 設定モーダル(簡易版) ---
document.getElementById('open-settings').addEventListener('click', openSettings)

function openSettings() {
  const config = loadConfig()
  const settings = loadSettings()
  const profile = settings.profiles[0] || newProfile()

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div style="position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center;z-index:10">
      <div style="background:var(--surface-2);border-radius:12px;padding:20px;width:320px;max-width:90vw">
        <h2 style="font-size:15px;margin:0 0 12px">設定</h2>
        <label style="font-size:12px;color:var(--text-secondary)">GAS URL</label>
        <input id="cfg-gas" value="${config.gasUrl}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">共有トークン</label>
        <input id="cfg-token" value="${config.notionToken}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">LLM baseUrl</label>
        <input id="cfg-llm-url" value="${profile.baseUrl}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">LLM APIキー</label>
        <input id="cfg-llm-key" value="${profile.apiKey}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">モデル名</label>
        <input id="cfg-llm-model" value="${profile.model}" style="width:100%;margin-bottom:16px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="cfg-cancel" class="btn">キャンセル</button>
          <button id="cfg-save" class="btn">保存</button>
        </div>
      </div>
    </div>
  `
  document.getElementById('cfg-cancel').addEventListener('click', () => (root.innerHTML = ''))
  document.getElementById('cfg-save').addEventListener('click', () => {
    saveConfig({
      ...config,
      gasUrl: document.getElementById('cfg-gas').value.trim(),
      notionToken: document.getElementById('cfg-token').value.trim(),
    })

    const p = { ...profile,
      baseUrl: document.getElementById('cfg-llm-url').value.trim(),
      apiKey: document.getElementById('cfg-llm-key').value.trim(),
      model: document.getElementById('cfg-llm-model').value.trim(),
    }
    const nextProfiles = settings.profiles.length ? [p, ...settings.profiles.slice(1)] : [p]
    saveSettings({ ...settings, profiles: nextProfiles, activeId: p.id })

    root.innerHTML = ''
  })
}

loadIndex()
