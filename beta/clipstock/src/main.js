import {
  escapeHtml,
  renderTagRail,
  renderLibrary,
  detailHtml,
  renderIdeas,
  flattenIdeas,
  TABS,
} from './ui/render.js'
import { openSettings, openEditor } from './ui/settings.js'
import { openVocabPanel } from './ui/vocab.js'
import {
  listVideos,
  listIdeas,
  fetchDetail,
  fetchTranscript,
  saveGenerated,
  saveField,
  saveMemo,
  saveTags,
  saveTitle,
  setStatus,
  updateRawCount,
  mergeTag,
} from './lib/gas.js'
import { loadConfig, isConfigured, canEdit } from './lib/videos-config.js'
import { loadSettings, saveSettings, activeModelName, allModels, connectionOf } from './lib/llm-settings.js'
import { generateAll, generateStage, STAGES } from './lib/generate.js'
import { renderMindmap } from './lib/mindmap.js'
import { hasTimecodes } from './lib/timecode.js'
import { getDetailCache, setDetailCache, isCacheFresh, markSeen, isSeen } from './lib/cache.js'
import {
  excludeExcluded,
  filterByStatus,
  filterByTags,
  filterBySearch,
  buildTagOptions,
  statusCounts,
  STATUS_ORDER,
  STATUS_DONE,
  STATUS_NEW,
  STATUS_SUMMARIZED,
  STATUS_EXCLUDED,
} from './lib/filters.js'
import { knownTagsOf } from './lib/tags.js'
import { loadDismissed, dismissPair, clearDismissed } from './lib/vocab.js'
import { streamChat } from './lib/llm-client.js'
import {
  loadChat,
  saveChat,
  renderQA,
  wireComposer,
  loadSpaces,
  saveSpaces,
  newSpace,
  MAX_CROSS_ITEMS,
  WARN_CHARS,
} from './lib/chat.js'

const $ = (id) => document.getElementById(id)
const stageEl = $('stage')
const railEl = $('tag-rail')
const syncEl = $('sync-status')

// ---- アプリの状態 ----
let items = []
let view = 'library' // 'library' | 'detail' | 'ideas' | 'crosschat'
let selectedKey = null
let searchQuery = ''
let showTags = true
const selectedStatuses = new Set()
const selectedTags = new Set()

// 詳細画面の状態。動画を切り替えるたび作り直す
let detail = null

// アイデア一覧の状態
const ideasState = { phase: 'idle', items: [], kind: 'all', shuffleSeed: 0, message: '' }

// ============ 一覧の取得 ============

async function loadList() {
  const config = loadConfig()
  if (!isConfigured(config)) {
    stageEl.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-plug-connected-x" aria-hidden="true"></i>
        <p>まず接続の設定が必要です</p>
        <p class="empty-hint">右上の歯車から GAS URL と共有トークンを入れてください</p>
      </div>`
    return
  }
  syncEl.textContent = '読み込み中...'
  try {
    const data = await listVideos()
    items = data.items
    syncEl.textContent = `${items.length}件`
  } catch (err) {
    items = []
    syncEl.textContent = ''
    stageEl.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-alert-triangle" aria-hidden="true"></i>
        <p>一覧を取得できませんでした</p>
        <p class="empty-hint">${escapeHtml(String(err.message || err))}</p>
      </div>`
    return
  }
  refresh()
}

/** 「除外」を外した、この端末で見る対象の全件 */
function visibleItems() {
  return excludeExcluded(items)
}

function currentItems() {
  const byStatus = filterByStatus(visibleItems(), selectedStatuses)
  const bySearch = filterBySearch(byStatus, searchQuery)
  return filterByTags(bySearch, selectedTags)
}

function itemOf(key) {
  return items.find((i) => i.key === key)
}

// ============ 画面の描画 ============

function refresh() {
  paintStatusChips()
  paintActiveModel()

  // 詳細・横断チャットは内側でスクロールさせるので、外側のスクロールは切る
  stageEl.classList.toggle('no-scroll', view === 'detail' || view === 'crosschat')

  if (view === 'detail') {
    railEl.innerHTML = ''
    paintDetail()
    return
  }
  if (view === 'crosschat') {
    railEl.innerHTML = ''
    paintCrossChat()
    return
  }

  const base = filterBySearch(filterByStatus(visibleItems(), selectedStatuses), searchQuery)
  renderTagRail(railEl, { tagOptions: buildTagOptions(base, selectedTags), selectedTags }, {
    onToggleTag: (tag) => {
      selectedTags.has(tag) ? selectedTags.delete(tag) : selectedTags.add(tag)
      refresh()
    },
    onClearTags: () => {
      selectedTags.clear()
      refresh()
    },
  })

  if (view === 'ideas') {
    paintIdeas()
    return
  }

  renderLibrary(
    stageEl,
    currentItems(),
    { searchQuery, showTags, seen: isSeen },
    { onOpen: openDetail }
  )
}

function paintStatusChips() {
  const counts = statusCounts(visibleItems())
  const el = $('status-filter')
  el.innerHTML = STATUS_ORDER.map(
    (s) => `<button class="chip ${selectedStatuses.has(s) ? 'on' : ''}" data-status="${escapeHtml(s)}">
      ${escapeHtml(s)}<span class="chip-count">${counts.get(s) || 0}</span>
    </button>`
  ).join('')
  el.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const s = chip.dataset.status
      selectedStatuses.has(s) ? selectedStatuses.delete(s) : selectedStatuses.add(s)
      refresh()
    })
  })
}

function paintActiveModel() {
  const model = activeModelName(loadSettings())
  $('active-model').textContent = model ? `AI: ${model}` : 'AI未設定'
}

function setView(next) {
  view = next
  selectedKey = view === 'detail' ? selectedKey : null
  document.querySelectorAll('.viewtab').forEach((t) => t.classList.toggle('on', t.dataset.view === next))
  refresh()
}

// ============ 詳細 ============

async function openDetail(key) {
  const item = itemOf(key)
  if (!item) return
  selectedKey = key
  view = 'detail'
  transcriptInFlight = null
  markSeen(key)

  detail = {
    phase: 'loading',
    detail: null,
    tags: item.tags || [],
    activeTab: 'summary',
    transcript: undefined, // undefined=未取得 / null=取得中 / string=取得済み
    memoDraft: null,
    memoDirty: false,
    busyStage: null,
    busyLabel: '',
    busyText: '',
    message: '',
    canEdit: canEdit(loadConfig()),
  }
  refresh()

  // キャッシュが Notion の更新より新しければそれを使う
  const cached = getDetailCache(key)
  if (isCacheFresh(cached, item.editedAt)) {
    detail.detail = cached
    detail.tags = cached.tags || item.tags || []
    detail.phase = 'ready'
    paintDetail()
    return
  }

  try {
    const d = await fetchDetail(key)
    if (selectedKey !== key) return
    setDetailCache(key, d)
    detail.detail = d
    detail.tags = d.tags || []
    detail.phase = 'ready'
  } catch (err) {
    detail.phase = 'error'
    detail.message = String(err.message || err)
  }
  paintDetail()
}

function tabHasContent(id) {
  const d = detail?.detail || {}
  if (id === 'chat') return loadChat(selectedKey).length > 0
  if (id === 'raw') return Boolean(itemOf(selectedKey)?.rawCount)
  if (id === 'memo') return Boolean(detail?.memoDraft || d.memo)
  return Boolean(d[id])
}

function paintDetail() {
  const item = itemOf(selectedKey)
  if (!item) {
    setView('library')
    return
  }
  stageEl.innerHTML = detailHtml(item, { ...detail, tabHasContent })
  wireDetail(item)
}

function switchTab(id) {
  detail.activeTab = id
  paintDetail()
}

function wireDetail(item) {
  stageEl.querySelector('.btn-back')?.addEventListener('click', () => setView('library'))
  stageEl.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => switchTab(t.dataset.tab)))
  stageEl.querySelector('.btn-retry')?.addEventListener('click', () => openDetail(item.key))
  stageEl.querySelector('.btn-generate-all')?.addEventListener('click', () => runGenerateAll(item))
  stageEl.querySelector('.btn-regen')?.addEventListener('click', (e) => runStage(item, e.currentTarget.dataset.stage))
  stageEl.querySelector('.btn-edit-field')?.addEventListener('click', () => editCurrentField(item))
  stageEl.querySelector('.btn-more')?.addEventListener('click', (e) => openMoreMenu(e.currentTarget, item))
  stageEl.querySelector('.btn-tag-add')?.addEventListener('click', () => addTag(item))
  stageEl.querySelector('.btn-copy')?.addEventListener('click', () => copyCurrentTab())
  stageEl.querySelector('.btn-mm-full')?.addEventListener('click', () => openMindmapFull())
  stageEl.querySelectorAll('.tag-edit i').forEach((x) =>
    x.addEventListener('click', (e) => removeTag(item, e.target.closest('.tag').dataset.tag))
  )

  if (detail.phase !== 'ready' || detail.busyStage) return

  if (detail.activeTab === 'mindmap') {
    const host = stageEl.querySelector('#mindmap-host')
    if (host) renderMindmap(host, detail.detail?.mindmap, item.url)
  }
  if (detail.activeTab === 'memo') setupMemo(item)
  if (detail.activeTab === 'chat') setupChat(item)
  if (detail.activeTab === 'raw') ensureTranscript(item)
}

// ---- タブごとの仕込み ----

function setupMemo(item) {
  const input = stageEl.querySelector('#memo-input')
  const saveBtn = stageEl.querySelector('.btn-memo-save')
  if (!input || !saveBtn) return
  input.addEventListener('input', () => {
    detail.memoDraft = input.value
    const dirty = input.value !== (detail.detail?.memo ?? '')
    if (dirty !== detail.memoDirty) {
      detail.memoDirty = dirty
      $('memo-status').textContent = dirty ? '未保存の変更があります' : ''
    }
  })
  saveBtn.addEventListener('click', async () => {
    const value = input.value
    saveBtn.disabled = true
    saveBtn.textContent = '保存中...'
    try {
      await saveMemo(item.key, value)
      detail.detail = { ...detail.detail, memo: value }
      setDetailCache(item.key, { ...detail.detail, updatedAt: new Date().toISOString() })
      detail.memoDraft = value
      detail.memoDirty = false
      item.has = { ...(item.has || {}), memo: Boolean(value) }
      saveBtn.textContent = '保存しました'
      $('memo-status').textContent = ''
      setTimeout(() => {
        saveBtn.textContent = 'メモを保存'
        saveBtn.disabled = false
      }, 1200)
    } catch (err) {
      alert('保存できませんでした: ' + (err.message || err))
      saveBtn.textContent = 'メモを保存'
      saveBtn.disabled = false
    }
  })
}

let transcriptInFlight = null

async function ensureTranscript(item) {
  if (typeof detail.transcript === 'string') return detail.transcript
  if (transcriptInFlight) return transcriptInFlight
  detail.transcript = null
  transcriptInFlight = (async () => {
    try {
      const { text } = await fetchTranscript(item.key)
      detail.transcript = text
      if (!item.rawCount && text.length) {
        item.rawCount = text.length
        updateRawCount(item.key, text.length).catch(() => {})
      }
    } catch (err) {
      detail.transcript = ''
      console.error('原文の取得に失敗しました:', err)
    }
    transcriptInFlight = null
    if (detail.activeTab === 'raw' && !detail.busyStage) paintDetail()
    return detail.transcript
  })()
  return transcriptInFlight
}

/** 要約タブ相当のテキスト。チャットの軽いコンテキスト用 */
function summaryContext(d) {
  return ['# サマリ', d.summary || '', '', '# 分野別要約', d.fields || '', '', '# 応用', d.apply || '']
    .join('\n')
    .trim()
}

function setupChat(item) {
  const logEl = stageEl.querySelector('#chat-log')
  const inputEl = stageEl.querySelector('#chat-input')
  const sendBtn = stageEl.querySelector('#chat-send')
  const countEl = stageEl.querySelector('#ctx-count')
  if (!logEl || !inputEl || !sendBtn) return

  const messages = loadChat(item.key)
  renderQA(logEl, messages, item.url)

  let ctxMode = 'raw'
  let busy = false

  function paintCount() {
    const sum = summaryContext(detail.detail || {}).length
    const raw = typeof detail.transcript === 'string' ? detail.transcript.length : item.rawCount || 0
    countEl.textContent = `要約 約${sum.toLocaleString()}字 ／ 原文 ${raw ? `約${raw.toLocaleString()}字` : '未取得'}`
  }
  paintCount()

  stageEl.querySelectorAll('.ctx-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      ctxMode = btn.dataset.ctx
      stageEl.querySelectorAll('.ctx-btn').forEach((b) => b.classList.toggle('on', b === btn))
    })
  })

  wireComposer(inputEl, sendBtn, async () => {
    if (busy) return
    const text = inputEl.value.trim()
    if (!text) return
    inputEl.value = ''
    inputEl.style.height = 'auto'

    messages.push({ role: 'user', content: text })
    saveChat(item.key, messages)
    renderQA(logEl, messages, item.url)

    const connection = connectionOf(loadSettings())
    if (!connection) {
      messages.push({ role: 'assistant', content: 'AI接続が未設定です。設定から接続先とモデルを追加してください。' })
      saveChat(item.key, messages)
      renderQA(logEl, messages, item.url)
      return
    }

    busy = true
    // 本文取得より先にプレースホルダーを積んで「考え中」を即座に見せる
    messages.push({ role: 'assistant', content: '' })
    renderQA(logEl, messages, item.url)

    try {
      let context
      if (ctxMode === 'summary') {
        context = summaryContext(detail.detail || {})
      } else {
        context = (await ensureTranscript(item)) || summaryContext(detail.detail || {})
        paintCount()
      }

      // 原文にタイムスタンプがあるときだけ引用元の時刻を添えさせる。
      // 時刻の無い原文で頼むと、それらしい数字を作られるだけになる
      const citeRule = hasTimecodes(context)
        ? `
根拠になった箇所には、原文にあるタイムスタンプを [12:34] の形でそのまま添えてください。
原文に無い時刻を書いてはいけません。該当が分からなければ時刻は書かないでください。`
        : ''

      const system = `あなたは動画の内容について質問に答えるアシスタントです。
以下は「${item.title}」の${ctxMode === 'summary' ? '要約' : '文字起こし全文'}です。この内容の範囲で答え、書かれていないことは「分かりません」と答えてください。
Markdown(見出し・箇条書き・強調)を使って読みやすく整理して構いません。日本語で回答してください。${citeRule}

${context.slice(0, 30000)}`

      let full = ''
      for await (const chunk of streamChat(connection, [
        { role: 'system', content: system },
        ...messages.slice(0, -1).map((m) => ({ role: m.role, content: m.content })),
      ])) {
        if (chunk.delta) {
          full += chunk.delta
          messages[messages.length - 1].content = full
          renderQA(logEl, messages, item.url)
        }
      }
      if (!full) messages[messages.length - 1].content = '(応答がありませんでした)'
    } catch (err) {
      messages[messages.length - 1].content = 'エラーが発生しました: ' + (err.message || err)
    }
    saveChat(item.key, messages)
    renderQA(logEl, messages, item.url)
    busy = false
  })
}

// ---- 生成 ----

/**
 * 一時的なお知らせ。一括生成のバーと同じ場所を使い回す。
 * 新語が増えたことに気づけないと語彙が静かに膨らんでいくため、その通知に使う。
 */
let noticeTimer = null
function notice(text) {
  const bar = $('bulk')
  bar.innerHTML = `<div class="bulk"><span>${escapeHtml(text)}</span><span class="grow"></span><button class="btn btn-sm" id="notice-close">閉じる</button></div>`
  $('notice-close').addEventListener('click', () => {
    bar.innerHTML = ''
    clearTimeout(noticeTimer)
  })
  clearTimeout(noticeTimer)
  noticeTimer = setTimeout(() => {
    if (bar.querySelector('#notice-close')) bar.innerHTML = ''
  }, 8000)
}

/** 新しいタグが作られたときだけ知らせる。既存タグへ寄せられた分は黙って通す */
function reportTags(tagReport) {
  if (tagReport?.created?.length) {
    notice(`新しいタグを追加しました: ${tagReport.created.join('、')}`)
  }
}

function paintBusy(label, text) {
  detail.busyLabel = label
  detail.busyText = text
  const pre = stageEl.querySelector('.gen-stream')
  const p = stageEl.querySelector('.gen-progress .muted')
  if (pre && p) {
    p.textContent = label
    pre.textContent = (text || '').slice(-1200)
    pre.scrollTop = pre.scrollHeight
  } else {
    paintDetail()
  }
}

/** 段が終わるたびにNotionへ保存する。途中で失敗しても手前の段は残る */
async function persistStage(item, stageDetail, model) {
  const rawCount = typeof detail.transcript === 'string' ? detail.transcript.length : item.rawCount || 0
  await saveGenerated(item.key, stageDetail, model, rawCount)

  const next = { ...(detail.detail || {}), ...stageDetail, model, generatedAt: new Date().toISOString() }
  if (Array.isArray(stageDetail.tags)) {
    next.tags = stageDetail.tags
    detail.tags = stageDetail.tags
    item.tags = stageDetail.tags
  }
  detail.detail = next
  setDetailCache(item.key, { ...next, updatedAt: new Date().toISOString() })

  // 一覧カードの表示も追従させる
  item.status = STATUS_SUMMARIZED
  item.model = model
  if (typeof stageDetail.summary === 'string') item.summary = stageDetail.summary
  item.has = {
    ...(item.has || {}),
    mindmap: Boolean(next.mindmap),
    fields: Boolean(next.fields),
    apply: Boolean(next.apply),
    ideas: Boolean(next.ideas),
  }
}

async function generateContext(item) {
  const transcript = (await ensureTranscript(item)) || ''
  return {
    title: item.title,
    transcript,
    summary: detail.detail?.summary || '',
    fields: detail.detail?.fields || '',
    // 既存のタグを渡して語彙を縛る。渡さないと動画ごとに表記が増えていく
    knownTags: knownTagsOf(items),
  }
}

async function runGenerateAll(item) {
  if (detail.busyStage) return
  detail.busyStage = 'core'
  paintBusy('原文を読み込んでいます', '')
  try {
    const ctx = await generateContext(item)
    if (!ctx.transcript) throw new Error('原文がありません。状態が「完了」になるまで待ってください')

    const { tagReport } = await generateAll(ctx, {
      onStageStart: (stage) => paintBusy(`${stage.label} を生成中`, ''),
      onProgress: (stage, text) => paintBusy(`${stage.label} を生成中`, text),
      onStage: async (stageId, stageDetail, model) => {
        paintBusy('Notionに保存中', '')
        await persistStage(item, stageDetail, model)
      },
    })
    detail.busyStage = null
    paintDetail()
    reportTags(tagReport)
  } catch (err) {
    detail.busyStage = null
    paintDetail()
    alert('生成できませんでした: ' + (err.message || err))
  }
}

async function runStage(item, stageId) {
  if (detail.busyStage) return
  const stage = STAGES.find((s) => s.id === stageId)
  detail.busyStage = stageId
  paintBusy(`${stage?.label ?? stageId} を生成中`, '')
  try {
    const ctx = await generateContext(item)
    if (stageId !== 'apply' && !ctx.transcript) {
      throw new Error('原文がありません。状態が「完了」になるまで待ってください')
    }
    const { detail: stageDetail, model, tagReport } = await generateStage(stageId, ctx, (text) =>
      paintBusy(`${stage?.label ?? stageId} を生成中`, text)
    )
    paintBusy('Notionに保存中', '')
    await persistStage(item, stageDetail, model)
    detail.busyStage = null
    paintDetail()
    reportTags(tagReport)
  } catch (err) {
    detail.busyStage = null
    paintDetail()
    alert('生成できませんでした: ' + (err.message || err))
  }
}

// ---- 手編集・タグ・状態 ----

const FIELD_LABEL = {
  summary: 'サマリ',
  mindmap: 'マインドマップ',
  fields: '分野別要約',
  apply: '応用',
  ideas: '活用アイデア',
}

function editCurrentField(item) {
  const field = detail.activeTab
  if (!FIELD_LABEL[field]) return
  const hint =
    field === 'mindmap'
      ? 'markmap用のMarkdownです。# が中心、## が大項目、- が枝になります'
      : field === 'summary'
        ? ''
        : '## で項目名、次の行に説明、- で箇条書きです'

  openEditor({
    title: `${FIELD_LABEL[field]}を直す`,
    value: detail.detail?.[field] ?? '',
    hint,
    onSave: async (value) => {
      try {
        await saveField(item.key, field, value)
        const next = { ...(detail.detail || {}), [field]: value }
        detail.detail = next
        setDetailCache(item.key, { ...next, updatedAt: new Date().toISOString() })
        if (field === 'summary') item.summary = value
        item.has = { ...(item.has || {}), [field]: Boolean(value) }
        paintDetail()
      } catch (err) {
        alert('保存できませんでした: ' + (err.message || err))
      }
    },
  })
}

async function commitTags(item, next) {
  const before = detail.tags
  detail.tags = next
  item.tags = next
  paintDetail()
  try {
    await saveTags(item.key, next)
    setDetailCache(item.key, { ...(detail.detail || {}), tags: next, updatedAt: new Date().toISOString() })
  } catch (err) {
    detail.tags = before
    item.tags = before
    paintDetail()
    alert('タグを保存できませんでした: ' + (err.message || err))
  }
}

function addTag(item) {
  const known = [...new Set(items.flatMap((i) => i.tags || []))].sort((a, b) => a.localeCompare(b, 'ja'))
  const root = $('modal-root')
  root.innerHTML = `
    <div class="overlay">
      <div class="modal">
        <h2 class="modal-title">タグを追加</h2>
        <div class="row">
          <input id="tag-new" class="input grow" placeholder="新しいタグ" />
          <button id="tag-add" class="btn btn-primary">追加</button>
        </div>
        ${known.length ? `<div class="tag-row picker">${known.map((t) => `<button class="chip ${(detail.tags || []).includes(t) ? 'on' : ''}" data-tag="${escapeHtml(t)}">${escapeHtml(t)}</button>`).join('')}</div>` : ''}
        <div class="modal-foot"><button id="tag-close" class="btn">閉じる</button></div>
      </div>
    </div>
  `
  const close = () => (root.innerHTML = '')
  $('tag-close').addEventListener('click', close)
  $('tag-add').addEventListener('click', () => {
    const name = $('tag-new').value.trim()
    if (!name || (detail.tags || []).includes(name)) return close()
    commitTags(item, [...(detail.tags || []), name])
    close()
  })
  root.querySelectorAll('.chip[data-tag]').forEach((chip) =>
    chip.addEventListener('click', () => {
      const t = chip.dataset.tag
      const cur = detail.tags || []
      commitTags(item, cur.includes(t) ? cur.filter((x) => x !== t) : [...cur, t])
      close()
    })
  )
}

function removeTag(item, tag) {
  commitTags(item, (detail.tags || []).filter((t) => t !== tag))
}

function openMoreMenu(anchor, item) {
  document.querySelector('.popmenu')?.remove()
  const menu = document.createElement('div')
  menu.className = 'popmenu'
  menu.innerHTML = `
    <button data-act="title"><i class="ti ti-edit" aria-hidden="true"></i>タイトルを直す</button>
    <button data-act="retry"><i class="ti ti-microphone" aria-hidden="true"></i>文字起こしをやり直す</button>
    <button data-act="exclude" class="danger"><i class="ti ti-archive" aria-hidden="true"></i>一覧から除外する</button>
  `
  document.body.appendChild(menu)
  const rect = anchor.getBoundingClientRect()
  menu.style.top = `${rect.bottom + 4}px`
  menu.style.right = `${window.innerWidth - rect.right}px`

  const close = () => menu.remove()
  menu.querySelectorAll('button').forEach((btn) =>
    btn.addEventListener('click', async () => {
      close()
      const act = btn.dataset.act
      if (act === 'title') {
        openEditor({
          title: 'タイトルを直す',
          value: item.title,
          multiline: false,
          onSave: async (value) => {
            const next = value.trim()
            if (!next) return
            try {
              await saveTitle(item.key, next)
              item.title = next
              paintDetail()
            } catch (err) {
              alert('保存できませんでした: ' + (err.message || err))
            }
          },
        })
      }
      if (act === 'retry') {
        if (!confirm('状態を「新規」に戻します。次回のバッチで文字起こしをやり直します。')) return
        try {
          await setStatus(item.key, STATUS_NEW)
          item.status = STATUS_NEW
          paintDetail()
        } catch (err) {
          alert('変更できませんでした: ' + (err.message || err))
        }
      }
      if (act === 'exclude') {
        if (!confirm('一覧から除外します。Notionのページは残ります。')) return
        try {
          await setStatus(item.key, STATUS_EXCLUDED)
          item.status = STATUS_EXCLUDED
          setView('library')
        } catch (err) {
          alert('変更できませんでした: ' + (err.message || err))
        }
      }
    })
  )
  setTimeout(() => {
    document.addEventListener('click', function once(e) {
      if (!menu.contains(e.target)) {
        close()
        document.removeEventListener('click', once)
      }
    })
  }, 0)
}

function copyCurrentTab() {
  const d = detail.detail || {}
  const text = detail.activeTab === 'raw' ? detail.transcript || '' : d[detail.activeTab] || ''
  if (!text) return
  navigator.clipboard.writeText(text)
  const btn = stageEl.querySelector('.btn-copy')
  if (btn) {
    const before = btn.innerHTML
    btn.textContent = 'コピーしました'
    setTimeout(() => (btn.innerHTML = before), 1500)
  }
}

function openMindmapFull() {
  const root = $('modal-root')
  root.innerHTML = `
    <div class="overlay">
      <div class="modal modal-wide">
        <div class="modal-head">
          <h2 class="modal-title">${escapeHtml(itemOf(selectedKey)?.title || '')}</h2>
          <button id="mm-close" class="btn-ghost" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
        </div>
        <div id="mm-full" class="mindmap-host mindmap-full"></div>
      </div>
    </div>
  `
  $('mm-close').addEventListener('click', () => (root.innerHTML = ''))
  renderMindmap($('mm-full'), detail.detail?.mindmap, itemOf(selectedKey)?.url)
}

// ============ アイデア一覧 ============

async function paintIdeas() {
  if (ideasState.phase === 'idle') {
    ideasState.phase = 'loading'
    renderIdeas(stageEl, [], ideasState, {})
    try {
      const data = await listIdeas()
      ideasState.items = flattenIdeas(data.items)
      ideasState.phase = 'ready'
    } catch (err) {
      ideasState.phase = 'error'
      ideasState.message = String(err.message || err)
    }
    if (view !== 'ideas') return
  }

  let entries = ideasState.items
  if (ideasState.kind !== 'all') entries = entries.filter((e) => e.kind === ideasState.kind)
  if (selectedTags.size) entries = entries.filter((e) => [...selectedTags].every((t) => e.tags.includes(t)))
  if (searchQuery.trim()) {
    const q = searchQuery.trim().toLowerCase()
    entries = entries.filter(
      (e) =>
        e.heading.toLowerCase().includes(q) ||
        e.body.toLowerCase().includes(q) ||
        e.videoTitle.toLowerCase().includes(q)
    )
  }
  if (ideasState.shuffleSeed) entries = pickRandom(entries, 6, ideasState.shuffleSeed)

  renderIdeas(stageEl, entries, ideasState, {
    onKind: (kind) => {
      ideasState.kind = kind
      ideasState.shuffleSeed = 0
      paintIdeas()
    },
    onShuffle: () => {
      ideasState.shuffleSeed = ideasState.shuffleSeed ? 0 : Date.now()
      paintIdeas()
    },
    onOpen: (key) => openDetail(key),
  })
  stageEl.querySelector('.btn-retry')?.addEventListener('click', () => {
    ideasState.phase = 'idle'
    paintIdeas()
  })
}

/** 決め打ちのシードで並べ替えて先頭n件。再描画しても同じ並びになる */
function pickRandom(list, n, seed) {
  const scored = list.map((e, i) => {
    let h = seed + i * 2654435761
    h = (h ^ (h >>> 15)) * 2246822507
    return { e, r: (h ^ (h >>> 13)) >>> 0 }
  })
  return scored.sort((a, b) => a.r - b.r).slice(0, n).map((s) => s.e)
}

// ============ 横断チャット ============

let crossSpaceId = null

function paintCrossChat() {
  const spaces = loadSpaces()
  const space = spaces.find((s) => s.id === crossSpaceId) || spaces[0] || null

  if (!space) {
    stageEl.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-messages" aria-hidden="true"></i>
        <p>スペースがありません</p>
        <p class="empty-hint">複数の動画をまとめて対象にして質問できます</p>
        <button class="btn btn-primary btn-space-new">スペースを作る</button>
      </div>`
    stageEl.querySelector('.btn-space-new').addEventListener('click', createSpace)
    return
  }
  crossSpaceId = space.id

  stageEl.innerHTML = `
    <div class="detail">
      <div class="detail-fixed">
        <div class="detail-head">
          <button class="btn-ghost btn-back" aria-label="一覧に戻る"><i class="ti ti-arrow-left" aria-hidden="true"></i></button>
          <div class="detail-headline">
            <h2 class="detail-title">${escapeHtml(space.name)}</h2>
            <div class="detail-meta"><span>対象 ${space.targets.length}件</span><span>約${space.targets.reduce((a, t) => a + (t.chars || 0), 0).toLocaleString()}字</span></div>
          </div>
          <select id="space-select" class="input sm">
            ${spaces.map((s) => `<option value="${s.id}" ${s.id === space.id ? 'selected' : ''}>${escapeHtml(s.name)}</option>`).join('')}
          </select>
          <button class="btn-ghost btn-space-new" aria-label="スペースを作る"><i class="ti ti-plus" aria-hidden="true"></i></button>
        </div>
        <div class="cross-targets">
          ${space.targets.map((t) => `<span class="tag">${escapeHtml(t.title)}</span>`).join('') || '<span class="muted">対象未選択</span>'}
          <button class="btn btn-sm btn-pick">対象を選ぶ</button>
        </div>
      </div>
      <div class="detail-scroll"><div id="cross-log" class="chat-log"></div></div>
      <div class="detail-foot">
        <div class="composer">
          <div class="composer-row">
            <textarea id="cross-input" class="chat-input" rows="1" placeholder="選んだ動画をまとめて質問する(Shift+Enterで改行)"></textarea>
            <button id="cross-send" class="btn btn-primary" aria-label="送信"><i class="ti ti-send" aria-hidden="true"></i></button>
          </div>
        </div>
      </div>
    </div>
  `

  stageEl.querySelector('.btn-back').addEventListener('click', () => setView('library'))
  stageEl.querySelector('.btn-space-new').addEventListener('click', createSpace)
  stageEl.querySelector('.btn-pick').addEventListener('click', () => pickTargets(space.id))
  $('space-select').addEventListener('change', (e) => {
    crossSpaceId = e.target.value
    paintCrossChat()
  })

  const logEl = $('cross-log')
  renderQA(logEl, space.messages)

  let busy = false
  wireComposer($('cross-input'), $('cross-send'), async () => {
    if (busy) return
    const input = $('cross-input')
    const text = input.value.trim()
    if (!text) return
    if (!space.targets.length) {
      alert('先に対象の動画を選んでください')
      return
    }
    input.value = ''
    input.style.height = 'auto'

    space.messages.push({ role: 'user', content: text })
    persistSpace(space)
    renderQA(logEl, space.messages)

    const connection = connectionOf(loadSettings())
    if (!connection) {
      space.messages.push({ role: 'assistant', content: 'AI接続が未設定です。設定から接続先とモデルを追加してください。' })
      persistSpace(space)
      renderQA(logEl, space.messages)
      return
    }

    busy = true
    space.messages.push({ role: 'assistant', content: '' })
    renderQA(logEl, space.messages)

    try {
      const context = space.targets
        .map((t) => `## ${t.title}\n${t.text}`)
        .join('\n\n')
        .slice(0, 60000)
      const system = `あなたは複数の動画の要約を横断して質問に答えるアシスタントです。
以下は対象の動画それぞれの要約です。どの動画の話かが分かるよう、答えの中で動画タイトルに触れてください。
書かれていないことは「分かりません」と答えてください。Markdownで整理し、日本語で回答してください。

${context}`

      let full = ''
      for await (const chunk of streamChat(connection, [
        { role: 'system', content: system },
        ...space.messages.slice(0, -1).map((m) => ({ role: m.role, content: m.content })),
      ])) {
        if (chunk.delta) {
          full += chunk.delta
          space.messages[space.messages.length - 1].content = full
          renderQA(logEl, space.messages)
        }
      }
      if (!full) space.messages[space.messages.length - 1].content = '(応答がありませんでした)'
    } catch (err) {
      space.messages[space.messages.length - 1].content = 'エラーが発生しました: ' + (err.message || err)
    }
    persistSpace(space)
    renderQA(logEl, space.messages)
    busy = false
  })
}

function persistSpace(space) {
  const spaces = loadSpaces()
  const i = spaces.findIndex((s) => s.id === space.id)
  if (i === -1) spaces.push(space)
  else spaces[i] = space
  saveSpaces(spaces)
}

function createSpace() {
  const name = prompt('スペースの名前')
  if (name === null) return
  const space = newSpace(name.trim())
  persistSpace(space)
  crossSpaceId = space.id
  view = 'crosschat'
  paintCrossChat()
}

/**
 * 横断チャットの対象を選ぶ。原文ではなく要約(サマリ+分野別+応用)を積むので、
 * 20件でもコンテキストに収まる。
 */
function pickTargets(spaceId) {
  const candidates = visibleItems().filter((i) => i.summary)
  const chosen = new Set((loadSpaces().find((s) => s.id === spaceId)?.targets || []).map((t) => t.key))
  const root = $('modal-root')

  const paint = () => {
    root.innerHTML = `
      <div class="overlay">
        <div class="modal modal-wide">
          <div class="modal-head">
            <h2 class="modal-title">対象の動画を選ぶ</h2>
            <span class="foot-note" id="pick-count"></span>
          </div>
          <div class="pick-list">
            ${candidates
              .map(
                (i) => `<label class="pick-row">
                  <input type="checkbox" data-key="${escapeHtml(i.key)}" ${chosen.has(i.key) ? 'checked' : ''} />
                  <span class="pick-title">${escapeHtml(i.title)}</span>
                  <span class="foot-note">${escapeHtml(i.status)}</span>
                </label>`
              )
              .join('') || '<p class="empty-section">サマリのある動画がまだありません</p>'}
          </div>
          <div class="modal-foot">
            <button id="pick-cancel" class="btn">キャンセル</button>
            <button id="pick-ok" class="btn btn-primary">この対象で作る</button>
          </div>
        </div>
      </div>
    `
    const countEl = $('pick-count')
    const paintCount = () => {
      countEl.textContent = `${chosen.size} / ${MAX_CROSS_ITEMS}件`
      countEl.classList.toggle('warn', chosen.size > MAX_CROSS_ITEMS)
    }
    paintCount()

    root.querySelectorAll('.pick-row input').forEach((cb) =>
      cb.addEventListener('change', () => {
        const key = cb.dataset.key
        if (cb.checked) {
          if (chosen.size >= MAX_CROSS_ITEMS) {
            cb.checked = false
            alert(`対象は${MAX_CROSS_ITEMS}件までです`)
            return
          }
          chosen.add(key)
        } else {
          chosen.delete(key)
        }
        paintCount()
      })
    )
    $('pick-cancel').addEventListener('click', () => (root.innerHTML = ''))
    $('pick-ok').addEventListener('click', async () => {
      const ok = $('pick-ok')
      ok.disabled = true
      ok.textContent = '要約を集めています...'
      const targets = []
      let total = 0
      for (const key of chosen) {
        const item = itemOf(key)
        if (!item) continue
        let d = getDetailCache(key)
        if (!isCacheFresh(d, item.editedAt)) {
          try {
            d = setDetailCache(key, await fetchDetail(key))
          } catch {
            continue
          }
        }
        const text = summaryContext(d)
        total += text.length
        targets.push({ key, title: item.title, text, chars: text.length })
      }
      const spaces = loadSpaces()
      const space = spaces.find((s) => s.id === spaceId)
      if (space) {
        space.targets = targets
        saveSpaces(spaces)
      }
      root.innerHTML = ''
      if (total > WARN_CHARS) {
        alert(`対象の合計が約${total.toLocaleString()}字です。モデルのコンテキスト長を超えると答えが途切れることがあります。`)
      }
      paintCrossChat()
    })
  }
  paint()
}

// ============ 一括生成 ============

async function runBulkGenerate() {
  const targets = visibleItems().filter((i) => i.status === STATUS_DONE)
  if (!targets.length) {
    alert('状態が「完了」の動画がありません(文字起こし済みで要約待ちのものが対象です)')
    return
  }
  if (!confirm(`${targets.length}件を順番に生成します。時間がかかります。続けますか?`)) return

  const bar = $('bulk')
  let cancelled = false
  const paint = (i, label) => {
    bar.innerHTML = `
      <div class="bulk">
        <span>${i}/${targets.length} ${escapeHtml(label)}</span>
        <div class="bulk-track"><div class="bulk-fill" style="width:${(i / targets.length) * 100}%"></div></div>
        <button class="btn btn-sm" id="bulk-cancel">中止</button>
      </div>`
    $('bulk-cancel').addEventListener('click', () => {
      cancelled = true
      bar.innerHTML = '<div class="bulk"><span>中止しています...</span></div>'
    })
  }

  const failures = []
  // 一括実行の途中で新しく生まれたタグも語彙に加える。こうしないと、
  // 同じ回の中で似た動画がそれぞれ別の新語を作ってしまう
  let vocabulary = knownTagsOf(items)
  for (let i = 0; i < targets.length; i++) {
    if (cancelled) break
    const item = targets[i]
    paint(i, item.title)
    try {
      const { text: transcript } = await fetchTranscript(item.key)
      if (!transcript) throw new Error('原文が空です')
      await generateAll(
        { title: item.title, transcript, summary: '', fields: '', knownTags: vocabulary },
        {
          onStage: async (stageId, stageDetail, model) => {
            await saveGenerated(item.key, stageDetail, model, transcript.length)
            if (typeof stageDetail.summary === 'string') item.summary = stageDetail.summary
            if (Array.isArray(stageDetail.tags)) {
              item.tags = stageDetail.tags
              stageDetail.tags.forEach((t) => {
                if (!vocabulary.includes(t)) vocabulary = [...vocabulary, t]
              })
            }
          },
        }
      )
      item.status = STATUS_SUMMARIZED
      item.rawCount = transcript.length
    } catch (err) {
      failures.push(`${item.title}: ${err.message || err}`)
    }
  }

  bar.innerHTML = ''
  ideasState.phase = 'idle' // アイデア一覧を作り直させる
  refresh()
  if (failures.length) alert(`${failures.length}件でエラーが出ました:\n\n` + failures.slice(0, 8).join('\n'))
}

// ============ トップバーの配線 ============

$('search-input').addEventListener('input', (e) => {
  searchQuery = e.target.value
  refresh()
})
$('open-settings').addEventListener('click', () => openSettings(() => loadList()))
$('open-vocab').addEventListener('click', () => {
  openVocabPanel(visibleItems(), {
    dismissed: () => loadDismissed(),
    onKeep: (key) => dismissPair(key),
    onResetDismissed: () => clearDismissed(),
    onMerge: async (from, to) => {
      const res = await mergeTag(from, to)
      // 一覧の手元の状態も揃える。再取得を待たずに候補が消えるようにする
      items.forEach((i) => {
        if (!(i.tags || []).includes(from)) return
        i.tags = [...new Set(i.tags.map((t) => (t === from ? to : t)))]
      })
      ideasState.phase = 'idle'
      refresh()
      return res
    },
  })
})
$('bulk-generate').addEventListener('click', runBulkGenerate)
$('cross-chat').addEventListener('click', () => setView('crosschat'))
$('reload').addEventListener('click', () => {
  ideasState.phase = 'idle'
  loadList()
})
$('toggle-tags').addEventListener('click', () => {
  showTags = !showTags
  $('toggle-tags').classList.toggle('on', showTags)
  refresh()
})
document.querySelectorAll('.viewtab').forEach((t) => t.addEventListener('click', () => setView(t.dataset.view)))

$('active-model').addEventListener('click', () => {
  const settings = loadSettings()
  const models = allModels(settings)
  if (!models.length) {
    openSettings(() => loadList())
    return
  }
  document.querySelector('.popmenu')?.remove()
  const menu = document.createElement('div')
  menu.className = 'popmenu'
  let lastConn = null
  menu.innerHTML = models
    .map((m) => {
      const header = m.connectionId !== lastConn ? `<div class="popmenu-group">${escapeHtml(m.connectionLabel)}</div>` : ''
      lastConn = m.connectionId
      return `${header}<button data-conn="${m.connectionId}" data-model="${escapeHtml(m.model)}" class="${m.active ? 'on' : ''}">${escapeHtml(m.model)}</button>`
    })
    .join('')
  document.body.appendChild(menu)
  const rect = $('active-model').getBoundingClientRect()
  menu.style.top = `${rect.bottom + 4}px`
  menu.style.right = `${window.innerWidth - rect.right}px`
  menu.querySelectorAll('button[data-model]').forEach((btn) =>
    btn.addEventListener('click', () => {
      saveSettings({ ...settings, activeConnectionId: btn.dataset.conn, activeModel: btn.dataset.model })
      menu.remove()
      paintActiveModel()
    })
  )
  setTimeout(() => {
    document.addEventListener('click', function once(e) {
      if (!menu.contains(e.target) && e.target !== $('active-model')) {
        menu.remove()
        document.removeEventListener('click', once)
      }
    })
  }, 0)
})

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if ($('modal-root').innerHTML) {
      $('modal-root').innerHTML = ''
      return
    }
    if (view === 'detail' || view === 'crosschat') setView('library')
  }
  // 詳細を開いているとき、左右キーでタブを移動する
  if (view === 'detail' && detail?.phase === 'ready' && !e.metaKey && !e.ctrlKey) {
    const target = e.target
    if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT') return
    const i = TABS.findIndex((t) => t.id === detail.activeTab)
    if (e.key === 'ArrowRight' && i < TABS.length - 1) switchTab(TABS[i + 1].id)
    if (e.key === 'ArrowLeft' && i > 0) switchTab(TABS[i - 1].id)
  }
})

$('toggle-tags').classList.toggle('on', showTags)
loadList()
