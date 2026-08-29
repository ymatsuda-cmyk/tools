import { renderList, renderDetailHtml, renderToolbar, escapeHtml } from './ui/render.js'
import { fetchSummary, fetchTranscript, saveSummary, saveTags, saveTitle, saveDetail } from './lib/gas.js'
import { getDetailCache, setDetailCache, isCacheFresh } from './lib/cache.js'
import { generateSummary } from './lib/summarize.js'
import { loadConfig, saveConfig, isConfigured } from './lib/minutes-config.js'
import { loadSettings, saveSettings, newProfile } from './lib/llm-settings.js'
import { filterByMonth, filterBySearch, filterByTags, buildTagOptions, allKnownTags, excludeDeleted } from './lib/filters.js'

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
  const byMonth = filterByMonth(excludeDeleted(items), currentMonthKey)
  const byMonthAndSearch = filterBySearch(byMonth, searchQuery)
  return filterByTags(byMonthAndSearch, selectedTags)
}

function refresh() {
  const byMonth = filterByMonth(excludeDeleted(items), currentMonthKey)
  const baseItems = filterBySearch(byMonth, searchQuery) // タグ絞り込み前(タグ候補の母集団)
  const filteredItems = filterByTags(baseItems, selectedTags)
  const tagOptions = buildTagOptions(baseItems, selectedTags)

  renderToolbar(toolbarEl, {
    monthLabel: monthLabelOf(currentMonthKey),
    tagOptions,
  }, {
    onPrevMonth: () => { currentMonthKey = shiftMonth(currentMonthKey, -1); refresh() },
    onNextMonth: () => { currentMonthKey = shiftMonth(currentMonthKey, 1); refresh() },
    onToggleTag: (tag) => {
      selectedTags.has(tag) ? selectedTags.delete(tag) : selectedTags.add(tag)
      refresh()
    },
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
  target.querySelector('.btn-edit-title')?.addEventListener('click', () => editTitle(target, item, state))

  target.querySelectorAll('.btn-edit').forEach((el) => {
    el.addEventListener('click', () => openFieldEditor(target, item, state, el.dataset.field))
  })
  target.querySelectorAll('.todo-check').forEach((el) => {
    el.addEventListener('change', () => toggleTodo(target, item, state, Number(el.dataset.index), el.checked))
  })

  target.querySelectorAll('.tag-remove').forEach((el) => {
    el.addEventListener('click', (e) => {
      const tag = e.target.closest('.tag-chip').dataset.tag
      const next = (tagsByKey[item.notionPageId] || []).filter((t) => t !== tag)
      commitTags(target, item, state, next)
    })
  })
  target.querySelector('.tag-add-btn')?.addEventListener('click', () => {
    openTagPicker(target, item, state)
  })
}

/** ToDoのチェック状態を変更してNotionに保存する */
async function toggleTodo(target, item, state, index, done) {
  const summary = state.summary
  const prev = summary.detail.todos
  const next = prev.map((t, i) => (i === index ? { ...t, done } : t))

  const apply = (todos) => {
    const updated = { ...summary, detail: { ...summary.detail, todos } }
    setDetailCache(item.key, updated)
    paintDetail(target, item, { ...state, summary: updated })
    return updated
  }

  const updated = apply(next)
  try {
    await saveDetail(item.notionPageId, updated.cardSummary, updated.detail)
  } catch (err) {
    apply(prev)
    alert('ToDoの更新に失敗しました: ' + (err.message || err))
  }
}

/**
 * 項目ごとの編集モーダル。
 * 配列項目は1行1件のテキストとして編集させ、保存時に配列へ戻す。
 * 議事だけは入れ子構造のためJSONを直接編集する。
 */
function openFieldEditor(target, item, state, field) {
  const summary = state.summary
  const labels = {
    cardSummary: 'サマリ',
    agenda: '議事',
    decisions: '決定事項',
    todos: 'ToDo',
    topics: '論点',
  }

  let initial
  let hint
  if (field === 'cardSummary') {
    initial = summary.cardSummary || ''
    hint = '一覧カードにも表示されます'
  } else if (field === 'agenda') {
    initial = JSON.stringify(summary.detail.agenda || [], null, 2)
    hint = '議題ごとの入れ子構造のため、JSON形式で編集します'
  } else if (field === 'todos') {
    initial = (summary.detail.todos || []).map((t) => t.text).join('\n')
    hint = '1行に1件。チェック状態は保持されます'
  } else {
    initial = (summary.detail[field] || []).join('\n')
    hint = '1行に1件'
  }

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="raw-modal-overlay">
      <div class="raw-modal" style="width:min(560px,100%)">
        <div class="raw-modal-header">
          <span>${labels[field]}を編集</span>
          <button id="edit-close" class="btn-ghost" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
        </div>
        <div class="raw-modal-body">
          <p class="edit-hint">${hint}</p>
          <textarea id="edit-textarea" class="edit-textarea" rows="14">${escapeHtml(initial)}</textarea>
          <p id="edit-error" class="error-text" style="display:none"></p>
        </div>
        <div class="raw-modal-footer">
          <button id="edit-cancel" class="btn">キャンセル</button>
          <button id="edit-save" class="btn">保存</button>
        </div>
      </div>
    </div>
  `

  const close = () => (root.innerHTML = '')
  root.querySelector('#edit-close').addEventListener('click', close)
  root.querySelector('#edit-cancel').addEventListener('click', close)

  root.querySelector('#edit-save').addEventListener('click', async () => {
    const raw = root.querySelector('#edit-textarea').value
    const errorEl = root.querySelector('#edit-error')

    let updated
    if (field === 'cardSummary') {
      updated = { ...summary, cardSummary: raw.trim() }
    } else if (field === 'agenda') {
      try {
        const parsed = JSON.parse(raw)
        if (!Array.isArray(parsed)) throw new Error('配列である必要があります')
        updated = { ...summary, detail: { ...summary.detail, agenda: parsed } }
      } catch (err) {
        errorEl.textContent = 'JSONの形式が不正です: ' + (err.message || err)
        errorEl.style.display = 'block'
        return
      }
    } else if (field === 'todos') {
      const lines = raw.split('\n').map((l) => l.trim()).filter(Boolean)
      const prevTodos = summary.detail.todos || []
      // 同じ本文のToDoが残っていればチェック状態を引き継ぐ
      const next = lines.map((text) => {
        const found = prevTodos.find((t) => t.text === text)
        return { text, done: found ? found.done : false }
      })
      updated = { ...summary, detail: { ...summary.detail, todos: next } }
    } else {
      const lines = raw.split('\n').map((l) => l.trim()).filter(Boolean)
      updated = { ...summary, detail: { ...summary.detail, [field]: lines } }
    }

    close()
    setDetailCache(item.key, updated)
    paintDetail(target, item, { ...state, summary: updated })

    try {
      await saveDetail(item.notionPageId, updated.cardSummary, updated.detail)
    } catch (err) {
      setDetailCache(item.key, summary)
      paintDetail(target, item, { ...state, summary })
      alert('保存に失敗しました: ' + (err.message || err))
    }
  })
}

/**
 * 既存タグから複数選択できるモーダル。新規タグの追加もここで行う。
 * 候補は index.json 全体から集めるため、Notionへの追加問い合わせは不要。
 */
function openTagPicker(target, item, state) {
  const root = document.getElementById('modal-root')
  const current = new Set(tagsByKey[item.notionPageId] || [])
  const candidates = allKnownTags(items)
  // 他レコードに存在しない独自タグも候補に含める
  current.forEach((t) => { if (!candidates.includes(t)) candidates.push(t) })

  const draft = new Set(current)

  function paint() {
    root.innerHTML = `
      <div class="raw-modal-overlay">
        <div class="raw-modal" style="width:min(420px,100%)">
          <div class="raw-modal-header">
            <span>タグを選択</span>
            <button id="tag-close" class="btn-ghost" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
          </div>
          <div class="raw-modal-body">
            <div class="tag-picker-list">
              ${candidates.map((t) => `
                <span class="tag-chip picker-chip ${draft.has(t) ? 'selected' : ''}" data-tag="${escapeHtml(t)}">
                  ${draft.has(t) ? '<i class="ti ti-check" aria-hidden="true"></i>' : ''}${escapeHtml(t)}
                </span>
              `).join('')}
            </div>
            <div class="tag-new-row">
              <input type="text" id="tag-new-input" placeholder="新しいタグを追加" />
              <button id="tag-new-add" class="btn">追加</button>
            </div>
          </div>
          <div class="raw-modal-footer">
            <button id="tag-cancel" class="btn">キャンセル</button>
            <button id="tag-save" class="btn">保存</button>
          </div>
        </div>
      </div>
    `

    root.querySelectorAll('.picker-chip').forEach((el) => {
      el.addEventListener('click', () => {
        const t = el.dataset.tag
        draft.has(t) ? draft.delete(t) : draft.add(t)
        paint()
      })
    })
    root.querySelector('#tag-close').addEventListener('click', () => (root.innerHTML = ''))
    root.querySelector('#tag-cancel').addEventListener('click', () => (root.innerHTML = ''))
    root.querySelector('#tag-new-add').addEventListener('click', () => {
      const input = root.querySelector('#tag-new-input')
      const t = input.value.trim()
      if (!t) return
      if (!candidates.includes(t)) candidates.push(t)
      draft.add(t)
      paint()
    })
    root.querySelector('#tag-save').addEventListener('click', () => {
      root.innerHTML = ''
      commitTags(target, item, state, [...draft])
    })
  }

  paint()
}

async function editTitle(target, item, state) {
  const next = prompt('ミーティング名を入力してください', item.title)?.trim()
  if (!next || next === item.title) return

  const prev = item.title
  item.title = next // 楽観的に即反映
  paintDetail(target, item, state)
  updateRowTitle(item)

  try {
    await saveTitle(item.notionPageId, next)
  } catch (err) {
    item.title = prev
    paintDetail(target, item, state)
    updateRowTitle(item)
    alert('タイトルの保存に失敗しました: ' + (err.message || err))
  }
}

/** 一覧の該当行のタイトル表示だけを更新する */
function updateRowTitle(item) {
  const titleEl = findRow(item.key)?.querySelector('.list-item-title')
  if (titleEl) titleEl.textContent = item.title
}

async function commitTags(target, item, state, nextTags) {
  const prev = tagsByKey[item.notionPageId]
  const prevItemTags = item.tags

  const apply = (tags, itemTags) => {
    tagsByKey[item.notionPageId] = tags
    item.tags = itemTags
    refresh() // 一覧のタグ表示・フィルタを更新(モバイルではインライン詳細が消える)
    const t = detailTarget(findRow(item.key))
    paintDetail(t, item, state)
    return t
  }

  apply(nextTags, nextTags) // 楽観的に即反映
  try {
    await saveTags(item.notionPageId, nextTags)
  } catch (err) {
    apply(prev, prevItemTags) // 失敗したら元に戻す
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

/**
 * 1件分の要約を生成してNotionに保存し、キャッシュも更新する。
 * 単体実行(runGenerate)と一括実行(runBulkSummarize)の共通処理。
 * @param {(partial: string) => void} [onProgress]
 */
async function generateAndSave(item, onProgress) {
  const { text } = await fetchTranscript(item.notionPageId)
  const result = await generateSummary(text, onProgress)

  await saveSummary(item.notionPageId, result.cardSummary, {
    agenda: result.agenda,
    decisions: result.decisions,
    todos: result.todos,
    topics: result.topics,
  }, result.model)

  item.status = '要約'

  return setDetailCache(item.key, {
    cardSummary: result.cardSummary,
    detail: {
      agenda: result.agenda,
      decisions: result.decisions,
      todos: result.todos,
      topics: result.topics,
    },
    model: result.model,
    generatedAt: new Date().toISOString(),
  })
}

async function runGenerate(target, item) {
  paintDetail(target, item, { phase: 'generating' })
  try {
    const saved = await generateAndSave(item, (partial) => {
      paintDetail(target, item, { phase: 'generating', progress: partial.slice(0, 200) })
    })
    updateRowBadge(item)
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

// --- 一括要約 ---
const bulkProgressEl = document.getElementById('bulk-progress')
let bulkCancelled = false
let bulkRunning = false

/**
 * 現在の絞り込み結果のうち、議事(agenda)が空のものをまとめて要約する。
 * 「議事が空」かどうかはNotion側の実データを見ないと分からないため、
 * 生成の前に対象確認フェーズ(fetchSummaryを1件ずつ呼ぶだけの軽い問い合わせ)を挟む。
 * LLMエンドポイントの同時実行を避けるため生成は直列に処理し、1件失敗しても続行する。
 */
async function runBulkSummarize() {
  if (bulkRunning) { bulkCancelled = true; return }

  const candidates = currentFilteredItems()
  if (!candidates.length) {
    alert('現在の絞り込みに議事録がありません')
    return
  }

  bulkRunning = true
  bulkCancelled = false

  // --- フェーズ1: 議事が空のものだけを対象に絞る ---
  const targets = []
  for (let i = 0; i < candidates.length; i++) {
    if (bulkCancelled) break
    paintBulkChecking(i + 1, candidates.length)
    try {
      const remote = await fetchSummary(candidates[i].notionPageId)
      if (!remote.generatedAt || !(remote.detail?.agenda?.length)) {
        targets.push(candidates[i])
      }
    } catch {
      // 確認に失敗したものは対象外にする(生成フェーズで無駄打ちしない)
    }
  }

  if (bulkCancelled) {
    bulkRunning = false
    bulkProgressEl.innerHTML = ''
    return
  }
  if (!targets.length) {
    bulkRunning = false
    bulkProgressEl.innerHTML = ''
    alert('対象がありません(議事が空の議事録が現在の絞り込みに含まれていません)')
    return
  }
  if (!confirm(`${targets.length}件を要約します。よろしいですか?`)) {
    bulkRunning = false
    bulkProgressEl.innerHTML = ''
    return
  }

  // --- フェーズ2: 生成 ---
  const results = { done: 0, failed: 0, errors: [] }
  for (let i = 0; i < targets.length; i++) {
    if (bulkCancelled) break
    const item = targets[i]
    paintBulkProgress({ current: i + 1, total: targets.length, title: item.title, ...results })

    try {
      await generateAndSave(item)
      results.done++
      updateRowBadge(item)
    } catch (err) {
      results.failed++
      results.errors.push(`${item.title}: ${err.message || err}`)
    }
  }

  bulkRunning = false
  paintBulkProgress({ finished: true, cancelled: bulkCancelled, total: targets.length, ...results })
  refresh()
}

function paintBulkChecking(current, total) {
  const pct = Math.round((current / total) * 100)
  bulkProgressEl.innerHTML = `
    <div class="bulk-bar">
      <div class="bulk-track"><div class="bulk-fill" style="width:${pct}%"></div></div>
      <span class="bulk-status">対象を確認中 ${current} / ${total}</span>
      <button class="btn bulk-cancel">中断</button>
    </div>
  `
  bulkProgressEl.querySelector('.bulk-cancel').addEventListener('click', () => {
    bulkCancelled = true
  })
}

function paintBulkProgress(state) {
  if (state.finished) {
    const label = state.cancelled ? '中断しました' : '完了しました'
    bulkProgressEl.innerHTML = `
      <div class="bulk-bar">
        <span class="bulk-status">${label} — 成功 ${state.done}件 / 失敗 ${state.failed}件</span>
        ${state.errors.length ? `<button class="btn bulk-errors">失敗を表示</button>` : ''}
        <button class="btn bulk-close">閉じる</button>
      </div>
    `
    bulkProgressEl.querySelector('.bulk-close').addEventListener('click', () => (bulkProgressEl.innerHTML = ''))
    bulkProgressEl.querySelector('.bulk-errors')?.addEventListener('click', () => {
      alert(state.errors.join('\n'))
    })
    return
  }

  const pct = Math.round((state.current / state.total) * 100)
  bulkProgressEl.innerHTML = `
    <div class="bulk-bar">
      <div class="bulk-track"><div class="bulk-fill" style="width:${pct}%"></div></div>
      <span class="bulk-status">${state.current} / ${state.total} — ${escapeHtml(state.title)}</span>
      <button class="btn bulk-cancel">中断</button>
    </div>
  `
  bulkProgressEl.querySelector('.bulk-cancel').addEventListener('click', () => {
    bulkCancelled = true
  })
}

document.getElementById('bulk-summarize').addEventListener('click', runBulkSummarize)

// 検索欄とタグ表示トグルは再生成しない永続DOMなので、初回に一度だけ結線する。
// (毎回 innerHTML で作り直すと入力のたびにフォーカスが外れ、1文字しか打てなくなる)
document.getElementById('search-input').addEventListener('input', (e) => {
  searchQuery = e.target.value
  refresh()
})
document.getElementById('show-tags-checkbox').addEventListener('change', (e) => {
  showTags = e.target.checked
  refresh()
})

loadIndex()
