import { renderList, renderDetailHtml, renderToolbar, escapeHtml } from './ui/render.js'
import { fetchSummary, fetchTranscript, saveSummary, saveTags, saveTitle, saveDetail, requestRetranscribe, verifyCode, savePermissions, saveMemo } from './lib/gas.js'
import { getDetailCache, setDetailCache, isCacheFresh } from './lib/cache.js'
import { generateSummary } from './lib/summarize.js'
import { loadConfig, saveConfig, isConfigured, isAdmin, isDenied } from './lib/minutes-config.js'
import { loadSettings, saveSettings, newProfile, activeProfile } from './lib/llm-settings.js'
import { filterByMonth, filterBySearch, filterByTags, filterByStatus, filterByPermission, filterByPermissionTags, buildTagOptions, buildPermissionOptions, allKnownTags, excludeDeleted } from './lib/filters.js'
import { loadCrossChatData, saveCrossChatData, clearCrossChatData, estimateItemChars, GEMMA_WARN_CHARS, loadSpaces, saveSpaces, newSpace } from './lib/cross-chat.js'
import { streamChat } from './lib/llm-client.js'
import { connectionOf } from './lib/llm-settings.js'

const listEl = document.getElementById('list')
const listItemsEl = document.getElementById('list-items')
const toolbarEl = document.getElementById('toolbar')
const detailEl = document.getElementById('detail')
const syncStatusEl = document.getElementById('sync-status')

let items = []
let selectedKey = null
const tagsByKey = {} // pageId(notionPageId) -> string[]、タグ編集の楽観更新用
const memoByKey = {} // pageId(notionPageId) -> string、メモの楽観更新用
const activeTabByKey = {} // item.key -> 'summary'|'decisions'|'todos'|'memo'、選択中タブの記憶

// --- 一覧の絞り込み状態 ---
let currentMonthKey = monthKeyOf(new Date()) // "YYYY-MM"
let searchQuery = ''
const selectedTags = new Set()
const selectedPermissionFilters = new Set() // 管理者専用の権限フィルタ(タグの前段)
const selectedStatuses = new Set() // 空 = 全ステータス表示
let assignMode = false // 管理者の権限一括割り当てモード
const selectedIds = new Set() // 一括割り当ての選択中キー(全期間をまたいで保持する)
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
/** 権限フィルタと削除除外を適用した、この利用者が見てよい全期間の議事録 */
function visibleItems() {
  return filterByPermission(excludeDeleted(items), loadConfig().role)
}

function currentFilteredItems() {
  const byMonth = filterByMonth(visibleItems(), currentMonthKey)
  const byMonthAndSearch = filterBySearch(byMonth, searchQuery)
  const byStatus = filterByStatus(byMonthAndSearch, selectedStatuses)
  const byPerm = filterByPermissionTags(byStatus, selectedPermissionFilters)
  return filterByTags(byPerm, selectedTags)
}

/** タイトルバー右端に、現在アクティブなAI接続プロファイルのモデル名を表示する */
function paintActiveModel() {
  const settings = loadSettings()
  const p = activeProfile(settings)
  const el = document.getElementById('active-model')
  el.textContent = p?.model ? `AI: ${p.model}` : ''
}

function refresh() {
  paintActiveModel()
  const config = loadConfig()
  const admin = isAdmin(config)

  if (isDenied(config)) {
    renderDenied()
    return
  }

  // 割り当てモードでは全期間から選べるようにするため月フィルタを外す
  const scoped = assignMode ? visibleItems() : filterByMonth(visibleItems(), currentMonthKey)
  const byMonthAndSearch = filterBySearch(scoped, searchQuery)
  const baseItems = filterByStatus(byMonthAndSearch, selectedStatuses) // タグ絞り込み前(タグ候補の母集団)
  const byPerm = admin ? filterByPermissionTags(baseItems, selectedPermissionFilters) : baseItems
  const filteredItems = filterByTags(byPerm, selectedTags)
  const tagOptions = buildTagOptions(byPerm, selectedTags)
  const permissionOptions = admin ? buildPermissionOptions(baseItems, selectedPermissionFilters) : []

  renderToolbar(toolbarEl, {
    monthLabel: assignMode ? '全期間' : monthLabelOf(currentMonthKey),
    tagOptions,
    permissionOptions,
  }, {
    onPrevMonth: () => { if (assignMode) return; currentMonthKey = shiftMonth(currentMonthKey, -1); refresh() },
    onNextMonth: () => { if (assignMode) return; currentMonthKey = shiftMonth(currentMonthKey, 1); refresh() },
    onToggleTag: (tag) => {
      selectedTags.has(tag) ? selectedTags.delete(tag) : selectedTags.add(tag)
      refresh()
    },
    onTogglePermission: (perm) => {
      selectedPermissionFilters.has(perm) ? selectedPermissionFilters.delete(perm) : selectedPermissionFilters.add(perm)
      refresh()
    },
  })

  renderList(listItemsEl, filteredItems, selectedKey, onSelect, showTags, {
    showPermissions: admin,
    selectable: assignMode,
    selectedIds,
  })

  if (assignMode) {
    listItemsEl.querySelectorAll('.row-select').forEach((el) => {
      el.addEventListener('change', () => {
        el.checked ? selectedIds.add(el.dataset.key) : selectedIds.delete(el.dataset.key)
        paintAssignBar()
      })
    })
  }

  document.getElementById('bulk-summarize').style.display = admin ? '' : 'none'
  document.getElementById('assign-permission').style.display = admin ? '' : 'none'

  syncStatusEl.textContent = `${filteredItems.length}件`
}

/** 権限が無いときは一覧・詳細を出さずメッセージのみ表示する */
function renderDenied() {
  toolbarEl.innerHTML = ''
  listItemsEl.innerHTML = ''
  detailEl.innerHTML = `
    <div class="empty-state">
      <i class="ti ti-lock" aria-hidden="true"></i>
      <p>権限がないため表示できません</p>
      <p style="font-size:12px">設定画面でコードを入力してください</p>
    </div>
  `
  syncStatusEl.textContent = ''
  document.getElementById('bulk-summarize').style.display = 'none'
  document.getElementById('assign-permission').style.display = 'none'
}

/**
 * 詳細を表示するターゲット要素を決める。
 * モバイルは選択した行の直後にインライン挿入、PCは右ペインに固定表示。
 */
function detailTarget(rowEl) {
  if (!isMobile()) {
    detailEl.classList.add('side-panel')
    return document.getElementById('detail-content')
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
  target.innerHTML = renderDetailHtml(item, {
    ...state,
    tags: tagsByKey[item.notionPageId],
    memo: memoByKey[item.notionPageId],
    activeTab: activeTabByKey[item.key],
    canEdit: isAdmin(loadConfig()), // タグ・タイトル・文字起こし・要約生成は管理者のみ
    canEditContent: true, // サマリ/議事/決定事項/ToDo/論点の編集は誰でも可能
  })
  const generateBtn = target.querySelector('.btn-generate, .btn-regenerate')
  generateBtn?.addEventListener('click', () => runGenerate(target, item))
  target.querySelector('.btn-retry')?.addEventListener('click', () => onSelect(item, findRow(item.key)))
  target.querySelector('.btn-raw')?.addEventListener('click', () => showRawTranscript(item))
  target.querySelector('.btn-edit-title')?.addEventListener('click', () => editTitle(target, item, state))
  target.querySelector('.btn-retranscribe')?.addEventListener('click', () => retranscribeItem(target, item, state))

  target.querySelectorAll('.detail-tab').forEach((el) => {
    el.addEventListener('click', () => {
      activeTabByKey[item.key] = el.dataset.tab
      paintDetail(target, item, state)
    })
  })
  target.querySelector('.btn-memo-save')?.addEventListener('click', () => saveMemoField(target, item, state))

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

/** メモを保存する。要約とは独立した項目なので単独でNotionへ反映する */
async function saveMemoField(target, item, state) {
  const textarea = target.querySelector('.memo-textarea')
  const statusEl = target.querySelector('#memo-save-status')
  const value = textarea.value
  const prev = memoByKey[item.notionPageId]

  memoByKey[item.notionPageId] = value // 楽観的に即反映
  if (statusEl) statusEl.textContent = '保存中...'

  try {
    await saveMemo(item.notionPageId, value)
    if (statusEl) statusEl.textContent = '保存しました'
    setTimeout(() => { if (statusEl) statusEl.textContent = '' }, 2000)
  } catch (err) {
    memoByKey[item.notionPageId] = prev
    if (statusEl) statusEl.textContent = ''
    alert('メモの保存に失敗しました: ' + (err.message || err))
  }
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

/**
 * 状態を「再取得」にし、Mac mini側のバッチ処理(retranscribe.py)による
 * 文字起こしのやり直しをリクエストする。楽観的に即バッジを更新し、
 * 失敗時は元の状態に戻す。
 */
async function retranscribeItem(target, item, state) {
  if (!confirm(`「${item.title}」を再文字起こし対象にしますか?\n状態が「再取得」に変わり、次回のバッチ処理で音声から文字起こしをやり直します。`)) return

  const prev = item.status
  item.status = '再取得' // 楽観的に即反映
  updateRowBadge(item)
  paintDetail(target, item, state)

  try {
    await requestRetranscribe(item.notionPageId)
  } catch (err) {
    item.status = prev
    updateRowBadge(item)
    paintDetail(target, item, state)
    alert('状態の更新に失敗しました: ' + (err.message || err))
  }
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
    memoByKey[item.notionPageId] = remote.memo || ''

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
  // 下書き。ここで編集し、保存時にまとめて反映する(キャンセル時は破棄)
  const draftProfiles = settings.profiles.length ? settings.profiles.map((p) => ({ ...p })) : [newProfile()]
  let draftActiveId = settings.activeId || draftProfiles[0].id

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div style="position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center;z-index:10">
      <div style="background:var(--surface-2);border-radius:12px;padding:20px;width:380px;max-width:90vw;max-height:85vh;overflow-y:auto">
        <h2 style="font-size:15px;margin:0 0 12px">設定</h2>
        <label style="font-size:12px;color:var(--text-secondary)">GAS URL</label>
        <input id="cfg-gas" value="${config.gasUrl}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">共有トークン</label>
        <input id="cfg-token" value="${config.notionToken}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">コード</label>
        <div style="display:flex;gap:6px;margin-bottom:4px">
          <input id="cfg-code" value="${config.code}" style="flex:1;min-width:0;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
          <button id="cfg-verify" class="btn">確認</button>
        </div>
        <div id="cfg-role" style="font-size:11px;margin-bottom:14px">${roleLabel(config.role)}</div>

        <div style="display:flex;align-items:center;margin-bottom:8px">
          <label style="font-size:12px;color:var(--text-secondary);flex:1">AI接続プロファイル</label>
          <button id="cfg-llm-add" class="btn" style="font-size:11px;padding:3px 9px">+ 追加</button>
        </div>
        <div id="cfg-llm-list"></div>

        <details style="margin:12px 0">
          <summary style="font-size:12px;color:var(--text-secondary);cursor:pointer">JSON文字列で一括設定</summary>
          <textarea id="cfg-json" rows="8" style="width:100%;margin-top:6px;font-family:var(--font-mono,monospace);font-size:11px;padding:6px;border:0.5px solid var(--border);border-radius:6px"></textarea>
          <div style="display:flex;gap:8px;margin-top:6px">
            <button id="cfg-json-export" class="btn" style="font-size:12px">現在の設定を書き出す</button>
            <button id="cfg-json-import" class="btn" style="font-size:12px">この内容を反映</button>
          </div>
        </details>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="cfg-cancel" class="btn">キャンセル</button>
          <button id="cfg-save" class="btn">保存</button>
        </div>
      </div>
    </div>
  `

  /** プロファイル一覧を描画。各カードは開閉式で、使用中はラジオで選ぶ */
  function paintProfiles() {
    const listEl = document.getElementById('cfg-llm-list')
    listEl.innerHTML = draftProfiles.map((p, i) => `
      <div class="profile-card ${p.id === draftActiveId ? 'active' : ''}" style="border:0.5px solid var(--border);border-radius:8px;padding:8px 10px;margin-bottom:8px">
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
          <input type="radio" name="cfg-active" class="profile-active" data-id="${p.id}" ${p.id === draftActiveId ? 'checked' : ''} />
          <input type="text" class="profile-label" data-id="${p.id}" value="${escapeHtml(p.label)}" placeholder="表示名(例: Gemini)" style="flex:1;min-width:0;font-size:12px;padding:4px 6px;border:0.5px solid var(--border);border-radius:6px" />
          ${draftProfiles.length > 1 ? `<button class="btn profile-delete" data-id="${p.id}" style="font-size:11px;padding:3px 7px">削除</button>` : ''}
        </div>
        <input type="text" class="profile-baseurl" data-id="${p.id}" value="${escapeHtml(p.baseUrl)}" placeholder="baseUrl (例: https://generativelanguage.googleapis.com/v1beta/openai)" style="width:100%;margin-bottom:5px;font-size:11px;padding:5px 6px;border:0.5px solid var(--border);border-radius:6px" />
        <input type="text" class="profile-apikey" data-id="${p.id}" value="${escapeHtml(p.apiKey)}" placeholder="APIキー" style="width:100%;margin-bottom:5px;font-size:11px;padding:5px 6px;border:0.5px solid var(--border);border-radius:6px" />
        <input type="text" class="profile-model" data-id="${p.id}" value="${escapeHtml(p.model)}" placeholder="モデル名 (例: gemini-2.5-flash)" style="width:100%;font-size:11px;padding:5px 6px;border:0.5px solid var(--border);border-radius:6px" />
      </div>
    `).join('')

    listEl.querySelectorAll('.profile-active').forEach((el) => {
      el.addEventListener('change', () => {
        draftActiveId = el.dataset.id
        listEl.querySelectorAll('.profile-card').forEach((c) => c.classList.remove('active'))
        el.closest('.profile-card').classList.add('active')
      })
    })
    listEl.querySelectorAll('.profile-label, .profile-baseurl, .profile-apikey, .profile-model').forEach((el) => {
      el.addEventListener('input', () => {
        const p = draftProfiles.find((p) => p.id === el.dataset.id)
        if (!p) return
        if (el.classList.contains('profile-label')) p.label = el.value
        if (el.classList.contains('profile-baseurl')) p.baseUrl = el.value
        if (el.classList.contains('profile-apikey')) p.apiKey = el.value
        if (el.classList.contains('profile-model')) p.model = el.value
      })
    })
    listEl.querySelectorAll('.profile-delete').forEach((el) => {
      el.addEventListener('click', () => {
        const idx = draftProfiles.findIndex((p) => p.id === el.dataset.id)
        if (idx === -1) return
        draftProfiles.splice(idx, 1)
        if (draftActiveId === el.dataset.id) draftActiveId = draftProfiles[0].id
        paintProfiles()
      })
    })
  }
  paintProfiles()

  document.getElementById('cfg-llm-add').addEventListener('click', () => {
    const p = newProfile({ label: `プロファイル${draftProfiles.length + 1}` })
    draftProfiles.push(p)
    paintProfiles()
  })

  document.getElementById('cfg-cancel').addEventListener('click', () => (root.innerHTML = ''))

  document.getElementById('cfg-json-export').addEventListener('click', () => {
    const json = {
      gasUrl: document.getElementById('cfg-gas').value.trim(),
      notionToken: document.getElementById('cfg-token').value.trim(),
      code: document.getElementById('cfg-code').value.trim(),
      llmProfiles: draftProfiles.map(({ label, baseUrl, apiKey, model }) => ({ label, baseUrl, apiKey, model })),
      activeLlmLabel: draftProfiles.find((p) => p.id === draftActiveId)?.label,
    }
    document.getElementById('cfg-json').value = JSON.stringify(json, null, 2)
  })

  document.getElementById('cfg-json-import').addEventListener('click', () => {
    let parsed
    try {
      parsed = JSON.parse(document.getElementById('cfg-json').value)
    } catch (err) {
      alert('JSONの形式が不正です: ' + (err.message || err))
      return
    }
    if (parsed.gasUrl !== undefined) document.getElementById('cfg-gas').value = parsed.gasUrl
    if (parsed.notionToken !== undefined) document.getElementById('cfg-token').value = parsed.notionToken
    if (parsed.code !== undefined) document.getElementById('cfg-code').value = parsed.code

    if (Array.isArray(parsed.llmProfiles) && parsed.llmProfiles.length) {
      draftProfiles.length = 0
      parsed.llmProfiles.forEach((p) => draftProfiles.push(newProfile(p)))
      const match = draftProfiles.find((p) => p.label === parsed.activeLlmLabel)
      draftActiveId = match ? match.id : draftProfiles[0].id
      paintProfiles()
    }

    if (parsed.code) document.getElementById('cfg-verify').click()
    alert('反映しました。内容を確認して「保存」を押してください。')
  })

  let verifiedRole = config.role
  document.getElementById('cfg-verify').addEventListener('click', async () => {
    const gasUrl = document.getElementById('cfg-gas').value.trim()
    const code = document.getElementById('cfg-code').value.trim()
    const roleEl = document.getElementById('cfg-role')
    if (!gasUrl) { roleEl.innerHTML = '<span style="color:var(--text-danger)">GAS URLを先に入力してください</span>'; return }
    roleEl.textContent = '確認中...'
    try {
      const res = await verifyCode(gasUrl, code)
      verifiedRole = res.role
      roleEl.innerHTML = roleLabel(res.role)
    } catch (err) {
      roleEl.innerHTML = `<span style="color:var(--text-danger)">${escapeHtml(String(err.message || err))}</span>`
    }
  })
  document.getElementById('cfg-save').addEventListener('click', () => {
    saveConfig({
      ...config,
      gasUrl: document.getElementById('cfg-gas').value.trim(),
      notionToken: document.getElementById('cfg-token').value.trim(),
      code: document.getElementById('cfg-code').value.trim(),
      role: verifiedRole,
    })

    saveSettings({ ...settings, profiles: draftProfiles, activeId: draftActiveId })

    root.innerHTML = ''
    refresh()
  })
}

/** 認証結果の表示ラベル */
function roleLabel(role) {
  if (!role) return '<span style="color:var(--text-muted)">未確認</span>'
  if (role === 'err') return '<span style="color:var(--text-danger)">権限がありません</span>'
  if (role === 'xYz') return '<span style="color:var(--text-accent)">管理者 — 全機能</span>'
  return `<span style="color:var(--text-success)">権限: ${escapeHtml(role)}</span>`
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

// --- 権限の一括割り当て(管理者のみ) ---
const assignBarEl = document.getElementById('assign-bar')

function toggleAssignMode() {
  assignMode = !assignMode
  selectedIds.clear()
  refresh()
  paintAssignBar()
}

function paintAssignBar() {
  if (!assignMode) {
    assignBarEl.innerHTML = ''
    return
  }
  // 権限の選択肢は既存データから集め、新規入力もできるようにする
  const known = new Set(['xYz'])
  items.forEach((i) => (i.permissions || []).forEach((p) => known.add(p)))

  assignBarEl.innerHTML = `
    <div class="bulk-bar">
      <span class="bulk-status">${selectedIds.size}件を選択中 — 全期間から選べます</span>
      <div style="flex:1"></div>
      <input id="assign-value" list="assign-options" placeholder="権限" style="width:110px;font-size:12px;padding:5px 8px" />
      <datalist id="assign-options">${[...known].sort().map((p) => `<option value="${escapeHtml(p)}"></option>`).join('')}</datalist>
      <button class="btn" id="assign-add">割り当て</button>
      <button class="btn" id="assign-remove">解除</button>
      <button class="btn" id="assign-exit">終了</button>
    </div>
  `
  document.getElementById('assign-add').addEventListener('click', () => applyPermissions('add'))
  document.getElementById('assign-remove').addEventListener('click', () => applyPermissions('remove'))
  document.getElementById('assign-exit').addEventListener('click', toggleAssignMode)
}

async function applyPermissions(mode) {
  const value = document.getElementById('assign-value').value.trim()
  if (!value) { alert('権限を入力してください'); return }
  if (!selectedIds.size) { alert('対象を選択してください'); return }

  const targets = items.filter((i) => selectedIds.has(i.key))
  const label = mode === 'add' ? '割り当て' : '解除'
  if (!confirm(`${targets.length}件に「${value}」を${label}します。よろしいですか?`)) return

  // GASの6分上限に収まるよう小さめに分割して送る
  const CHUNK = 20
  const chunks = []
  for (let i = 0; i < targets.length; i += CHUNK) chunks.push(targets.slice(i, i + CHUNK))

  let updated = 0
  const errors = []
  for (let i = 0; i < chunks.length; i++) {
    paintBulkProgress({ current: i + 1, total: chunks.length, title: `${label}中(${chunks[i].length}件ずつ処理)`, done: updated, failed: errors.length, errors: [] })
    try {
      const res = await savePermissions(chunks[i].map((t) => t.notionPageId), [value], mode)
      updated += res.updated
      if (res.errors?.length) errors.push(...res.errors)
    } catch (err) {
      errors.push(String(err.message || err))
    }
  }
  paintBulkProgress({ finished: true, total: targets.length, done: updated, failed: errors.length, errors })

  // 手元のデータにも反映(index.jsonの再取得を待たずに一覧へ出すため)
  targets.forEach((t) => {
    const current = t.permissions || []
    t.permissions = mode === 'add'
      ? [...new Set([...current, value])]
      : current.filter((p) => p !== value)
  })

  selectedIds.clear()
  refresh()
  paintAssignBar()
}

document.getElementById('assign-permission').addEventListener('click', toggleAssignMode)

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
document.querySelectorAll('.status-filter-chip').forEach((el) => {
  el.addEventListener('click', () => {
    const status = el.dataset.status
    selectedStatuses.has(status) ? selectedStatuses.delete(status) : selectedStatuses.add(status)
    el.classList.toggle('selected')
    refresh()
  })
})

// --- 一覧の幅リサイズ ---
const LIST_WIDTH_KEY = 'minutes:listWidth'
const LIST_WIDTH_MIN = 200
const LIST_WIDTH_MAX = 600

function restoreListWidth() {
  const saved = Number(localStorage.getItem(LIST_WIDTH_KEY))
  if (saved) listEl.style.width = `${saved}px`
}

function setupResizeHandle() {
  const handle = document.getElementById('resize-handle')
  let startX = 0
  let startWidth = 0

  const onMove = (e) => {
    const clientX = e.touches ? e.touches[0].clientX : e.clientX
    const next = Math.min(LIST_WIDTH_MAX, Math.max(LIST_WIDTH_MIN, startWidth + (clientX - startX)))
    listEl.style.width = `${next}px`
  }
  const onUp = () => {
    handle.classList.remove('dragging')
    document.removeEventListener('mousemove', onMove)
    document.removeEventListener('mouseup', onUp)
    document.removeEventListener('touchmove', onMove)
    document.removeEventListener('touchend', onUp)
    localStorage.setItem(LIST_WIDTH_KEY, String(listEl.getBoundingClientRect().width))
  }
  const onDown = (e) => {
    startX = e.touches ? e.touches[0].clientX : e.clientX
    startWidth = listEl.getBoundingClientRect().width
    handle.classList.add('dragging')
    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp)
    document.addEventListener('touchmove', onMove, { passive: true })
    document.addEventListener('touchend', onUp)
  }

  handle.addEventListener('mousedown', onDown)
  handle.addEventListener('touchstart', onDown, { passive: true })
}

// --- 詳細エリアのズーム(他のエリアには影響させない) ---
const DETAIL_ZOOM_KEY = 'minutes:detailZoom'
const ZOOM_MIN = 0.8
const ZOOM_MAX = 2.0
const ZOOM_STEP = 0.1

function setupDetailZoom() {
  const content = document.getElementById('detail-content')
  const levelEl = document.getElementById('detail-zoom-level')
  let zoom = Number(localStorage.getItem(DETAIL_ZOOM_KEY)) || 1

  const apply = () => {
    content.style.zoom = zoom
    levelEl.textContent = `${Math.round(zoom * 100)}%`
    localStorage.setItem(DETAIL_ZOOM_KEY, String(zoom))
  }
  apply()

  document.getElementById('detail-zoom-in').addEventListener('click', () => {
    zoom = Math.min(ZOOM_MAX, Math.round((zoom + ZOOM_STEP) * 10) / 10)
    apply()
  })
  document.getElementById('detail-zoom-out').addEventListener('click', () => {
    zoom = Math.max(ZOOM_MIN, Math.round((zoom - ZOOM_STEP) * 10) / 10)
    apply()
  })
}

restoreListWidth()
setupResizeHandle()
setupDetailZoom()

loadIndex()

// ============ 横断チャット ============

let crossChatView = 'select' // 'select' | 'chat'
let crossChatSelection = { fromMonth: '', toMonth: '', tags: new Set(), excluded: new Set() }
let activeSpaceId = null
let crossChatBusy = false

document.getElementById('cross-chat-btn').addEventListener('click', openCrossChat)

function openCrossChat() {
  const data = loadCrossChatData()
  crossChatView = data ? 'chat' : 'select'
  if (data && !activeSpaceId) {
    const spaces = loadSpaces()
    activeSpaceId = spaces[0]?.id || null
  }
  paintCrossChat()
}

function closeCrossChat() {
  document.getElementById('modal-root').innerHTML = ''
}

function paintCrossChat() {
  const root = document.getElementById('modal-root')
  const data = loadCrossChatData()

  root.innerHTML = `
    <div class="raw-modal-overlay">
      <div class="raw-modal cross-chat-modal">
        <div class="raw-modal-header">
          <span>横断チャット${data ? ` — ${data.count}件を読み込み済み` : ''}</span>
          <div style="display:flex;gap:6px;align-items:center">
            ${data ? '<button id="cc-manage" class="btn" style="font-size:11px;padding:3px 9px">データ管理</button>' : ''}
            <button id="cc-close" class="btn-ghost" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
          </div>
        </div>
        <div id="cc-body" class="raw-modal-body" style="padding:0"></div>
      </div>
    </div>
  `
  document.getElementById('cc-close').addEventListener('click', closeCrossChat)
  document.getElementById('cc-manage')?.addEventListener('click', () => { crossChatView = 'select'; paintCrossChat() })

  if (crossChatView === 'select') {
    paintCrossChatSelect()
  } else {
    paintCrossChatSpaces()
  }
}

// --- 対象選択・データ作成 ---

function paintCrossChatSelect() {
  const body = document.getElementById('cc-body')
  const candidates = visibleItems() // 権限フィルタ済み、削除除外済み
  const tags = allKnownTags(candidates)

  const inRange = (item) => {
    const m = item.date.slice(0, 7)
    if (crossChatSelection.fromMonth && m < crossChatSelection.fromMonth) return false
    if (crossChatSelection.toMonth && m > crossChatSelection.toMonth) return false
    return true
  }
  const matchesTags = (item) => {
    if (!crossChatSelection.tags.size) return true
    return (item.tags || []).some((t) => crossChatSelection.tags.has(t))
  }

  const filtered = candidates.filter((i) => inRange(i) && matchesTags(i))
  const selectable = filtered.filter((i) => i.status === '要約')
  const selectedCount = selectable.filter((i) => !crossChatSelection.excluded.has(i.key)).length

  body.innerHTML = `
    <div style="padding:12px 16px;border-bottom:0.5px solid var(--border)">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap">
        <span style="font-size:11px;color:var(--text-secondary)">期間</span>
        <input type="month" id="cc-from" value="${crossChatSelection.fromMonth}" style="font-size:12px;padding:5px 7px" />
        <span style="font-size:11px;color:var(--text-muted)">〜</span>
        <input type="month" id="cc-to" value="${crossChatSelection.toMonth}" style="font-size:12px;padding:5px 7px" />
      </div>
      <div style="display:flex;align-items:baseline;gap:8px;flex-wrap:wrap">
        <span style="font-size:11px;color:var(--text-secondary)">タグ</span>
        <div style="display:flex;gap:5px;flex-wrap:wrap;flex:1">
          ${tags.map((t) => `<span class="tag-chip filter-chip cc-tag-chip ${crossChatSelection.tags.has(t) ? 'selected' : ''}" data-tag="${escapeHtml(t)}">${escapeHtml(t)}</span>`).join('') || '<span style="font-size:11px;color:var(--text-muted)">タグがありません</span>'}
        </div>
      </div>
    </div>
    <div style="padding:6px 16px;border-bottom:0.5px solid var(--border);display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="cc-select-all" ${selectedCount === selectable.length && selectable.length ? 'checked' : ''} />
      <span style="font-size:11px;color:var(--text-secondary)">すべて選択</span>
      <div style="flex:1"></div>
      <span style="font-size:11px;color:var(--text-muted)">${filtered.length}件中 ${selectedCount}件を選択</span>
    </div>
    <div style="max-height:280px;overflow-y:auto;padding:4px 16px">
      ${filtered.map((i) => {
        const ok = i.status === '要約'
        const checked = ok && !crossChatSelection.excluded.has(i.key)
        return `
          <label style="display:flex;align-items:flex-start;gap:8px;padding:7px 0;border-bottom:0.5px solid var(--border);cursor:${ok ? 'pointer' : 'default'}">
            <input type="checkbox" class="cc-item-check" data-key="${escapeHtml(i.key)}" ${checked ? 'checked' : ''} ${ok ? '' : 'disabled'} style="margin-top:3px" />
            <div style="flex:1;min-width:0">
              <div style="font-size:12px;font-weight:500;${ok ? '' : 'color:var(--text-muted)'}">${escapeHtml(i.title)}</div>
              <div style="font-size:10px;color:var(--text-muted);margin-top:2px">${i.date.slice(0, 10)} · ${(i.tags || []).join(', ') || 'タグなし'}${ok ? '' : ' · 要約が未生成のため対象外'}</div>
            </div>
          </label>
        `
      }).join('') || '<p style="font-size:12px;color:var(--text-muted);padding:12px 0">該当する議事録がありません</p>'}
    </div>
    <div style="display:flex;align-items:center;gap:10px;padding:10px 16px;border-top:0.5px solid var(--border);background:var(--surface-1);flex-wrap:wrap">
      <span style="font-size:11px;color:var(--text-secondary);flex:1">${selectedCount}件を選択中(データ作成後に正確な文字数を表示します)</span>
      <button id="cc-create" class="btn" ${selectedCount ? '' : 'disabled'}><i class="ti ti-download" style="font-size:13px;vertical-align:-2px;margin-right:4px" aria-hidden="true"></i>データを作成</button>
    </div>
    <div id="cc-create-progress"></div>
  `

  document.getElementById('cc-from').addEventListener('change', (e) => { crossChatSelection.fromMonth = e.target.value; paintCrossChatSelect() })
  document.getElementById('cc-to').addEventListener('change', (e) => { crossChatSelection.toMonth = e.target.value; paintCrossChatSelect() })
  body.querySelectorAll('.cc-tag-chip').forEach((el) => {
    el.addEventListener('click', () => {
      const t = el.dataset.tag
      crossChatSelection.tags.has(t) ? crossChatSelection.tags.delete(t) : crossChatSelection.tags.add(t)
      paintCrossChatSelect()
    })
  })
  document.getElementById('cc-select-all').addEventListener('change', (e) => {
    if (e.target.checked) selectable.forEach((i) => crossChatSelection.excluded.delete(i.key))
    else selectable.forEach((i) => crossChatSelection.excluded.add(i.key))
    paintCrossChatSelect()
  })
  body.querySelectorAll('.cc-item-check').forEach((el) => {
    el.addEventListener('change', () => {
      el.checked ? crossChatSelection.excluded.delete(el.dataset.key) : crossChatSelection.excluded.add(el.dataset.key)
      paintCrossChatSelect()
    })
  })
  document.getElementById('cc-create').addEventListener('click', () => {
    const targets = selectable.filter((i) => !crossChatSelection.excluded.has(i.key))
    runCrossChatDataCreation(targets)
  })
}

async function runCrossChatDataCreation(targets) {
  const progressEl = document.getElementById('cc-create-progress')
  const createBtn = document.getElementById('cc-create')
  createBtn.disabled = true

  const entries = []
  let totalChars = 0

  for (let i = 0; i < targets.length; i++) {
    const item = targets[i]
    progressEl.innerHTML = `
      <div class="bulk-bar">
        <div class="bulk-track"><div class="bulk-fill" style="width:${Math.round(((i + 1) / targets.length) * 100)}%"></div></div>
        <span class="bulk-status">${i + 1} / ${targets.length} — ${escapeHtml(item.title)}</span>
      </div>
    `
    try {
      const remote = await fetchSummary(item.notionPageId)
      const entry = {
        key: item.key,
        title: item.title,
        date: item.date,
        tags: item.tags || [],
        agenda: remote.detail?.agenda || [],
        decisions: remote.detail?.decisions || [],
        todos: (remote.detail?.todos || []).map((t) => t.text),
      }
      entries.push(entry)
      totalChars += estimateItemChars(entry)
    } catch {
      // 取得失敗はスキップ(件数が減るだけで処理は継続)
    }
  }

  const data = { createdAt: new Date().toISOString(), count: entries.length, chars: totalChars, items: entries }
  saveCrossChatData(data)

  const warn = totalChars > GEMMA_WARN_CHARS
  progressEl.innerHTML = `
    <div class="bulk-bar" style="flex-direction:column;align-items:stretch;gap:6px">
      <span class="bulk-status">${entries.length}件を読み込みました(約${totalChars.toLocaleString()}字)</span>
      ${warn ? `<span style="font-size:11px;color:var(--text-danger)">Gemmaのコンテキスト上限の目安(約${GEMMA_WARN_CHARS.toLocaleString()}字)を超えています。件数を減らすことをおすすめします。</span>` : ''}
    </div>
  `
  setTimeout(() => {
    crossChatView = 'chat'
    if (!activeSpaceId) {
      const spaces = loadSpaces()
      activeSpaceId = spaces[0]?.id || null
    }
    paintCrossChat()
  }, warn ? 2500 : 800)
}

// --- チャットスペース ---

function paintCrossChatSpaces() {
  const body = document.getElementById('cc-body')
  const data = loadCrossChatData()
  const spaces = loadSpaces()
  if (!activeSpaceId && spaces.length) activeSpaceId = spaces[0].id
  const active = spaces.find((s) => s.id === activeSpaceId)

  body.innerHTML = `
    <div style="display:grid;grid-template-columns:150px minmax(0,1fr);height:420px">
      <div style="border-right:0.5px solid var(--border);background:var(--surface-1);overflow-y:auto">
        <div style="padding:8px 10px;border-bottom:0.5px solid var(--border)">
          <button id="cc-new-space" style="width:100%;font-size:11px;padding:5px 0"><i class="ti ti-plus" style="font-size:13px;vertical-align:-2px;margin-right:3px" aria-hidden="true"></i>新しいスペース</button>
        </div>
        ${spaces.map((s) => `
          <div class="cc-space-item ${s.id === activeSpaceId ? 'active' : ''}" data-id="${s.id}" style="padding:8px 10px;cursor:pointer;${s.id === activeSpaceId ? 'border-left:2px solid var(--border-accent);background:var(--bg-accent)' : 'border-left:2px solid transparent'}">
            <div style="font-size:12px;font-weight:500;${s.id === activeSpaceId ? 'color:var(--text-accent)' : ''}">${escapeHtml(s.name)}</div>
            <div style="font-size:10px;color:${s.id === activeSpaceId ? 'var(--text-accent)' : 'var(--text-muted)'};margin-top:2px">${s.messages.length}件のやり取り</div>
          </div>
        `).join('') || '<p style="font-size:11px;color:var(--text-muted);padding:8px 10px">スペースがありません</p>'}
      </div>
      <div style="display:flex;flex-direction:column;min-height:0">
        <div id="cc-messages" style="flex:1;overflow-y:auto;padding:14px 13px"></div>
        <div style="display:flex;gap:7px;padding:10px 13px;border-top:0.5px solid var(--border)">
          <input type="text" id="cc-input" placeholder="${data ? `${data.count}件の議事録に質問する` : 'まずデータを作成してください'}" style="flex:1;min-width:0;font-size:12px;padding:7px 9px" ${data && active ? '' : 'disabled'} />
          <button id="cc-send" class="btn" ${data && active ? '' : 'disabled'} aria-label="送信"><i class="ti ti-send" style="font-size:14px" aria-hidden="true"></i></button>
        </div>
      </div>
    </div>
  `

  paintCrossChatMessages(active)

  document.getElementById('cc-new-space').addEventListener('click', () => {
    const name = prompt('スペース名を入力してください', `スペース${spaces.length + 1}`)
    if (!name) return
    const s = newSpace(name)
    const next = [...spaces, s]
    saveSpaces(next)
    activeSpaceId = s.id
    paintCrossChatSpaces()
  })
  body.querySelectorAll('.cc-space-item').forEach((el) => {
    el.addEventListener('click', () => { activeSpaceId = el.dataset.id; paintCrossChatSpaces() })
  })

  const send = () => sendCrossChatMessage()
  document.getElementById('cc-send')?.addEventListener('click', send)
  document.getElementById('cc-input')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.isComposing) send()
  })
}

function paintCrossChatMessages(space) {
  const el = document.getElementById('cc-messages')
  if (!el) return
  if (!space) {
    el.innerHTML = '<p style="font-size:12px;color:var(--text-muted)">左のスペース一覧から選ぶか、新しいスペースを作成してください</p>'
    return
  }
  el.innerHTML = space.messages.map((m) => m.role === 'user'
    ? `<div style="display:flex;justify-content:flex-end;margin-bottom:10px"><div style="background:var(--surface-1);border-radius:12px;padding:8px 12px;max-width:78%;font-size:13px;line-height:1.7">${escapeHtml(m.content)}</div></div>`
    : `<div style="font-size:13px;line-height:1.8;margin-bottom:14px;white-space:pre-wrap">${escapeHtml(m.content)}</div>`
  ).join('')
  el.scrollTop = el.scrollHeight
}

async function sendCrossChatMessage() {
  if (crossChatBusy) return
  const input = document.getElementById('cc-input')
  const text = input.value.trim()
  if (!text) return

  const data = loadCrossChatData()
  const spaces = loadSpaces()
  const space = spaces.find((s) => s.id === activeSpaceId)
  if (!data || !space) return

  input.value = ''
  space.messages.push({ role: 'user', content: text })
  saveSpaces(spaces)
  paintCrossChatMessages(space)

  const connection = connectionOf(loadSettings())
  if (!connection) {
    space.messages.push({ role: 'assistant', content: 'LLM接続プロファイルが未設定です。設定から接続先を追加してください。' })
    saveSpaces(spaces)
    paintCrossChatMessages(space)
    return
  }

  const systemPrompt = `あなたは複数の会議議事録を横断して質問に答えるアシスタントです。
以下のJSONが対象データです。各要素は1件の会議を表し、agenda(議題ごとの経緯)・decisions(決定事項)・todos(ToDo)を持ちます。
このデータの範囲内で答え、無い情報は「分かりません」と答えてください。日本語で回答してください。

${JSON.stringify(data.items)}`

  const messages = [
    { role: 'system', content: systemPrompt },
    ...space.messages.map((m) => ({ role: m.role, content: m.content })),
  ]

  crossChatBusy = true
  space.messages.push({ role: 'assistant', content: '' })
  try {
    let full = ''
    for await (const chunk of streamChat(connection, messages)) {
      if (chunk.delta) {
        full += chunk.delta
        space.messages[space.messages.length - 1].content = full
        paintCrossChatMessages(space)
      }
    }
    if (!full) space.messages[space.messages.length - 1].content = '(応答がありませんでした)'
  } catch (err) {
    space.messages[space.messages.length - 1].content = 'エラーが発生しました: ' + (err.message || err)
  }
  saveSpaces(spaces)
  paintCrossChatMessages(space)
  crossChatBusy = false
}
