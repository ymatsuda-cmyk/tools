import {
  renderLibrary,
  renderDetail,
  renderIdeaGallery,
  renderMindmapGallery,
  flattenIdeas,
  escapeHtml,
} from './ui/render.js'
import { openSettings } from './ui/settings.js'
import { openVocabPanel } from './ui/vocab.js'
import {
  fetchDetail,
  fetchTranscript,
  saveGenerated,
  saveField,
  saveMemo,
  saveTags,
  saveTitle,
  setStatus,
  setPublic,
  mergeTag,
  deleteContent,
} from './lib/gas.js'
import { listContents, listIdeas } from './lib/store.js'
import { loadConfig, isConfigured, canEdit } from './lib/contents-config.js'
import { loadSettings, connectionOf, activeModelName } from './lib/llm-settings.js'
import { streamChat } from './lib/llm-client.js'
import { generateStage, generateAll, STAGES } from './lib/generate.js'
import { setSectionRank, setSectionHidden } from './lib/sections.js'
import { loadDismissed, dismissPair, clearDismissed } from './lib/vocab.js'
import { renderMindmap } from '../../../api/mindmap2/mindmap2.js'
import { uploadVideo } from './upload.js'

const stageEl = document.getElementById('stage')
const statusEl = document.getElementById('sync-status')
const $ = (id) => document.getElementById(id)

let view = 'library' // 'library' | 'ideas' | 'mindmaps' | 'detail'
let lastListView = 'library' // 詳細から戻る先
let items = []
let library = { phase: 'idle', message: '' }
let searchQuery = ''
const selectedTags = new Set()

// アイデア一覧は開いたときに1度だけ読む(一覧JSONとは別ファイル)
let ideas = { phase: 'idle', source: [], entries: [], kind: 'all', message: '' }

let detail = null // { key, item, phase, detail, activeTab, transcript, memoDraft, busyStage, busyText }

// 詳細は開くたびに取り直さず、最後に読んだものを覚えておく
const detailCache = new Map()

// ============ 一覧 ============

async function loadList() {
  library = { phase: 'loading' }
  ideas = { ...ideas, phase: 'idle' }
  paint()
  try {
    const data = await listContents()
    items = (data.items || []).filter((i) => i.status !== '除外')
    library = { phase: 'ready', source: data.source }
  } catch (err) {
    library = { phase: 'error', message: String(err.message || err) }
  }
  paint()
}

function filtered() {
  let list = items
  if (selectedTags.size) list = list.filter((i) => [...selectedTags].every((t) => (i.tags || []).includes(t)))
  const q = searchQuery.trim().toLowerCase()
  if (q) {
    list = list.filter((i) =>
      [i.title, i.summary, i.file, (i.tags || []).join(' ')].join(' ').toLowerCase().includes(q)
    )
  }
  return list
}

function paintTags() {
  const all = new Map()
  items.forEach((i) => (i.tags || []).forEach((t) => all.set(t, (all.get(t) || 0) + 1)))
  const box = $('tag-bar')
  box.innerHTML = [...all.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([tag, n]) => `<button class="tag-chip ${selectedTags.has(tag) ? 'on' : ''}" data-tag="${escapeHtml(tag)}">${escapeHtml(tag)}<span class="tag-n">${n}</span></button>`)
    .join('')
  box.querySelectorAll('.tag-chip').forEach((el) => {
    el.addEventListener('click', () => {
      const tag = el.dataset.tag
      selectedTags.has(tag) ? selectedTags.delete(tag) : selectedTags.add(tag)
      paint()
    })
  })
}

function paint() {
  $('active-model').textContent = activeModelName(loadSettings()) ? `AI: ${activeModelName(loadSettings())}` : '(AI未設定)'
  document.querySelectorAll('.viewtab').forEach((t) => t.classList.toggle('on', t.dataset.view === lastListView))

  if (view === 'detail' && detail) {
    paintDetail()
    return
  }
  paintTags()

  if (view === 'ideas') {
    paintIdeas()
    return
  }
  if (view === 'mindmaps') {
    paintMindmaps()
    return
  }

  const list = filtered()
  statusEl.textContent = library.phase === 'ready'
    ? `${list.length}件${library.source === 'notion' ? ' ・Notion直読み' : ''}`
    : ''
  renderLibrary(stageEl, list, library, { onOpen: openDetail, onRetry: loadList })
}

function setView(next) {
  view = next
  lastListView = next
  detail = null
  paint()
}

// ============ マインドマップ一覧 ============

/** 「公開」をONにしたものだけを並べる。マップ自体はカードを開いたときに描く */
function paintMindmaps() {
  const list = filtered().filter((i) => i.isPublic && i.has?.mindmap)
  statusEl.textContent = `${list.length}件`
  renderMindmapGallery(stageEl, list, { phase: library.phase, canEdit: canEdit(loadConfig()) }, {
    onOpen: openMindmapViewer,
    onHide: hideMindmap,
  })
}

/** 一覧からその場で公開をやめる。戻すときは詳細のマインドマップタブから */
async function hideMindmap(key) {
  const item = items.find((i) => i.key === key)
  if (!item) return
  item.isPublic = false
  paint()
  try {
    await setPublic(key, false)
  } catch (err) {
    item.isPublic = true
    paint()
    alert('公開の切り替えができませんでした: ' + (err.message || err))
  }
}

async function openMindmapViewer(key) {
  const item = items.find((i) => i.key === key)
  if (!item) return
  const root = $('mm-root')
  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal modal-full">
        <div class="modal-head"><span>${escapeHtml(item.title)}</span><button class="btn-ghost btn-close" aria-label="閉じる"><i class="ti ti-x"></i></button></div>
        <div id="mm-full" class="mindmap-host mindmap-full"><p class="muted">読み込んでいます...</p></div>
      </div>
    </div>
  `
  root.querySelector('.btn-close').addEventListener('click', () => (root.innerHTML = ''))

  try {
    let data = detailCache.get(key)
    if (!data) {
      data = await fetchDetail(key)
      detailCache.set(key, data)
    }
    // 読んでいる間に閉じられていることがある
    if ($('mm-full')) renderMindmap($('mm-full'), data.mindmap, { emptyText: 'マインドマップはまだありません', initialExpandLevel: 2 })
  } catch (err) {
    if ($('mm-full')) $('mm-full').innerHTML = `<p class="error-text">${escapeHtml(String(err.message || err))}</p>`
  }
}

// ============ アイデア一覧 ============

async function paintIdeas() {
  if (ideas.phase === 'idle') {
    ideas = { ...ideas, phase: 'loading' }
    renderIdeaGallery(stageEl, [], ideas, {})
    try {
      const data = await listIdeas()
      ideas = { ...ideas, phase: 'ready', source: data.items || [], entries: flattenIdeas(data.items || []) }
    } catch (err) {
      ideas = { ...ideas, phase: 'error', message: String(err.message || err) }
    }
    if (view !== 'ideas') return
  }

  let entries = ideas.entries.filter((e) => e.isPublic)
  if (ideas.kind !== 'all') entries = entries.filter((e) => e.kind === ideas.kind)
  if (selectedTags.size) entries = entries.filter((e) => [...selectedTags].every((t) => e.tags.includes(t)))
  const q = searchQuery.trim().toLowerCase()
  if (q) {
    entries = entries.filter((e) =>
      [e.heading, e.body, e.points.join(' '), e.contentTitle].join(' ').toLowerCase().includes(q)
    )
  }
  entries = [...entries].sort((a, b) => b.rank - a.rank)

  statusEl.textContent = ideas.phase === 'ready' ? `${entries.length}件` : ''
  renderIdeaGallery(stageEl, entries, { ...ideas, canEdit: canEdit(loadConfig()) }, {
    onKind: (kind) => {
      ideas.kind = kind
      paintIdeas()
    },
    onOpen: (id) => openDetail(ideas.entries.find((e) => e.id === id)?.key),
    onHide: hideIdea,
  })
  stageEl.querySelector('.btn-retry')?.addEventListener('click', () => {
    ideas.phase = 'idle'
    paintIdeas()
  })
}

/** アイデア1件を一覧から外す。戻すときは詳細の応用 / 活用タブから */
async function hideIdea(id) {
  const entry = ideas.entries.find((e) => e.id === id)
  const src = ideas.source.find((v) => v.key === entry?.key)
  if (!entry || !src) return
  const before = src[entry.kind] || ''
  const next = setSectionHidden(before, entry.sec, true)
  entry.isPublic = false
  src[entry.kind] = next
  paint()
  try {
    await saveField(entry.key, entry.kind, next)
    // 詳細を開き直したときに古い本文が出ないようにする
    const cached = detailCache.get(entry.key)
    if (cached) detailCache.set(entry.key, { ...cached, [entry.kind]: next })
  } catch (err) {
    entry.isPublic = true
    src[entry.kind] = before
    paint()
    alert('一覧から外せませんでした: ' + (err.message || err))
  }
}

// ============ タグの整理 ============

function openVocab() {
  let dismissed = loadDismissed()
  openVocabPanel(items, {
    dismissed: () => dismissed,
    onKeep: (key) => {
      dismissed = dismissPair(key)
    },
    onResetDismissed: () => {
      clearDismissed()
      dismissed = loadDismissed()
    },
    onMerge: async (from, to) => {
      const res = await mergeTag(from, to)
      // Notion 側は書き換わっている。JSON の再生成を待たずに手元も揃えておく
      items.forEach((i) => {
        if (!(i.tags || []).includes(from)) return
        i.tags = [...new Set(i.tags.map((t) => (t === from ? to : t)))]
      })
      ideas.entries.forEach((e) => {
        if (!e.tags.includes(from)) return
        e.tags = [...new Set(e.tags.map((t) => (t === from ? to : t)))]
      })
      selectedTags.delete(from)
      detailCache.clear()
      paint()
      return res
    },
  })
}

// ============ 詳細 ============

async function openDetail(key) {
  const item = items.find((i) => i.key === key)
  if (!item) return
  view = 'detail'
  detail = {
    key,
    item,
    phase: 'loading',
    detail: detailCache.get(key) || null,
    activeTab: 'summary',
    transcript: undefined,
    memoDraft: undefined,
    memoDirty: false,
    canEdit: canEdit(loadConfig()),
  }
  paintDetail()

  try {
    const data = await fetchDetail(key)
    detailCache.set(key, data)
    detail.detail = data
    detail.phase = 'ready'
  } catch (err) {
    detail.phase = 'ready'
    detail.detail = detail.detail || {}
    alert('詳細を読み込めませんでした: ' + (err.message || err))
  }
  paintDetail()
}

function paintDetail() {
  renderDetail(stageEl, detail.item, detail)
  wireDetail()
}

function switchTab(tab) {
  detail.activeTab = tab
  paintDetail()
  if (tab === 'raw' && detail.transcript === undefined) ensureTranscript()
}

function wireDetail() {
  stageEl.querySelector('.btn-back')?.addEventListener('click', () => setView(lastListView))
  stageEl.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => switchTab(t.dataset.tab)))
  stageEl.querySelector('.btn-generate-all')?.addEventListener('click', runGenerateAll)
  stageEl.querySelector('.btn-regen')?.addEventListener('click', (e) => runStage(e.currentTarget.dataset.stage))
  stageEl.querySelector('.btn-publish')?.addEventListener('click', (e) => togglePublic(e.currentTarget.dataset.on !== '1'))
  stageEl.querySelector('.btn-copy')?.addEventListener('click', copyCurrentTab)
  stageEl.querySelector('.btn-edit-title')?.addEventListener('click', editTitle)
  stageEl.querySelector('.btn-more')?.addEventListener('click', openMoreMenu)
  stageEl.querySelector('.tag-add')?.addEventListener('click', addTag)
  stageEl.querySelectorAll('.tag-remove').forEach((el) =>
    el.addEventListener('click', () => removeTag(el.dataset.tag))
  )
  stageEl.querySelectorAll('.rank .star[data-sec]').forEach((btn) =>
    btn.addEventListener('click', () => changeRank(detail.activeTab, Number(btn.dataset.sec), Number(btn.dataset.rank)))
  )
  stageEl.querySelectorAll('.sec-pub').forEach((btn) =>
    btn.addEventListener('click', () => changeSectionHidden(detail.activeTab, Number(btn.dataset.sec), btn.dataset.on === '1'))
  )

  if (detail.phase !== 'ready' || detail.busyStage) return

  if (detail.activeTab === 'mindmap') {
    const host = stageEl.querySelector('#mindmap-host')
    if (host) {
      renderMindmap(host, detail.detail?.mindmap, {
        emptyText: 'マインドマップはまだありません',
        initialExpandLevel: 2,
        onChange: detail.canEdit ? (markdown) => saveMindmapEdit(markdown) : null,
      })
    }
  }
  if (detail.activeTab === 'memo') setupMemo()
  if (detail.activeTab === 'chat') setupChat()
}

async function ensureTranscript() {
  detail.transcript = null
  paintDetail()
  try {
    const { text } = await fetchTranscript(detail.key)
    detail.transcript = text
  } catch (err) {
    detail.transcript = '(読み込めませんでした: ' + (err.message || err) + ')'
  }
  paintDetail()
}

// ---- AI生成 ----

function generateContext() {
  return {
    title: detail.detail?.title || detail.item.title,
    transcript: detail.transcript || '',
    summary: detail.detail?.summary || '',
    fields: detail.detail?.fields || '',
    knownTags: [...new Set(items.flatMap((i) => i.tags || []))],
  }
}

/** 原文が手元に無ければ取ってくる。生成はどの段も原文が起点になる */
async function needTranscript() {
  if (!detail.transcript) {
    const { text } = await fetchTranscript(detail.key)
    detail.transcript = text
  }
  if (!detail.transcript.trim()) throw new Error('原文(文字起こし)がまだありません')
  return detail.transcript
}

async function runStage(stageId) {
  if (!connectionOf(loadSettings())) {
    alert('AI接続が未設定です。設定から接続先とモデルを追加してください。')
    return
  }
  const label = STAGES.find((s) => s.id === stageId)?.label || stageId
  detail.busyStage = stageId
  detail.busyLabel = `${label}を生成しています...`
  detail.busyText = ''
  paintDetail()

  try {
    await needTranscript()
    const { detail: generated, model } = await generateStage(stageId, generateContext(), (text) => {
      detail.busyText = text.slice(-1500)
      const pre = stageEl.querySelector('.stream')
      if (pre) pre.textContent = detail.busyText
    })
    await saveGenerated(detail.key, generated, model, detail.transcript.length)
    applyGenerated(generated, model)
  } catch (err) {
    alert(`${label}の生成に失敗しました: ` + (err.message || err))
  }
  detail.busyStage = null
  paintDetail()
}

async function runGenerateAll() {
  if (!connectionOf(loadSettings())) {
    alert('AI接続が未設定です。設定から接続先とモデルを追加してください。')
    return
  }
  detail.busyStage = 'all'
  detail.busyText = ''
  paintDetail()

  try {
    await needTranscript()
    await generateAll(
      generateContext(),
      // 段ごとに保存する。途中で失敗しても手前の段は残る
      async (stageId, generated, model) => {
        await saveGenerated(detail.key, generated, model, detail.transcript.length)
        applyGenerated(generated, model)
      },
      (stageId, text) => {
        detail.busyLabel = `${STAGES.find((s) => s.id === stageId)?.label}を生成しています...`
        detail.busyText = text.slice(-1500)
        const pre = stageEl.querySelector('.stream')
        if (pre) pre.textContent = detail.busyText
      }
    )
  } catch (err) {
    alert('生成に失敗しました: ' + (err.message || err))
  }
  detail.busyStage = null
  paintDetail()
}

/** 生成結果を画面と一覧カードへ反映する */
function applyGenerated(generated, model) {
  detail.detail = { ...detail.detail, ...generated, model, generatedAt: new Date().toISOString() }
  detailCache.set(detail.key, detail.detail)
  const item = detail.item
  if (typeof generated.summary === 'string') item.summary = generated.summary
  if (Array.isArray(generated.tags)) item.tags = generated.tags
  item.has = {
    ...(item.has || {}),
    mindmap: Boolean(detail.detail.mindmap),
    fields: Boolean(detail.detail.fields),
    apply: Boolean(detail.detail.apply),
    ideas: Boolean(detail.detail.ideas),
  }
  item.status = '要約済み'
  item.model = model
}

// ---- 人手の編集 ----

async function saveMindmapEdit(markdown) {
  const before = detail.detail?.mindmap ?? ''
  if (markdown === before) return
  detail.detail = { ...detail.detail, mindmap: markdown }
  detailCache.set(detail.key, detail.detail)
  try {
    await saveField(detail.key, 'mindmap', markdown)
  } catch (err) {
    detail.detail = { ...detail.detail, mindmap: before }
    paintDetail()
    alert('マインドマップを保存できませんでした: ' + (err.message || err))
  }
}

async function changeRank(field, index, rank) {
  const before = detail.detail?.[field] ?? ''
  const next = setSectionRank(before, index, rank)
  if (next === before) return
  detail.detail = { ...detail.detail, [field]: next }
  detailCache.set(detail.key, detail.detail)
  paintDetail()
  try {
    await saveField(detail.key, field, next)
  } catch (err) {
    detail.detail = { ...detail.detail, [field]: before }
    paintDetail()
    alert('星を変えられませんでした: ' + (err.message || err))
  }
}

/** アイデア1件をアイデア一覧に出すかどうか。見出しの印で持つ */
async function changeSectionHidden(field, index, hidden) {
  const before = detail.detail?.[field] ?? ''
  const next = setSectionHidden(before, index, hidden)
  if (next === before) return
  detail.detail = { ...detail.detail, [field]: next }
  detailCache.set(detail.key, detail.detail)
  // アイデア一覧は別のJSONから読んでいるので、次に開いたときに取り直す
  ideas.phase = 'idle'
  paintDetail()
  try {
    await saveField(detail.key, field, next)
  } catch (err) {
    detail.detail = { ...detail.detail, [field]: before }
    paintDetail()
    alert('切り替えられませんでした: ' + (err.message || err))
  }
}

/** マインドマップ一覧に並べるかどうか */
async function togglePublic(next) {
  detail.detail = { ...detail.detail, isPublic: next }
  detail.item.isPublic = next
  detailCache.set(detail.key, detail.detail)
  paintDetail()
  try {
    await setPublic(detail.key, next)
  } catch (err) {
    detail.detail = { ...detail.detail, isPublic: !next }
    detail.item.isPublic = !next
    paintDetail()
    alert('公開の切り替えができませんでした: ' + (err.message || err))
  }
}

function setupMemo() {
  const input = stageEl.querySelector('#memo-input')
  const saveBtn = stageEl.querySelector('.btn-memo-save')
  if (!input || !saveBtn) return
  input.addEventListener('input', () => {
    detail.memoDraft = input.value
    detail.memoDirty = input.value !== (detail.detail?.memo ?? '')
    const el = stageEl.querySelector('#memo-status')
    if (el) el.textContent = detail.memoDirty ? '未保存の変更があります' : ''
  })
  saveBtn.addEventListener('click', async () => {
    const value = input.value
    const el = stageEl.querySelector('#memo-status')
    if (el) el.textContent = '保存中...'
    try {
      await saveMemo(detail.key, value)
      detail.detail = { ...detail.detail, memo: value }
      detailCache.set(detail.key, detail.detail)
      detail.memoDraft = undefined
      detail.memoDirty = false
      if (el) el.textContent = '保存しました'
    } catch (err) {
      if (el) el.textContent = '未保存の変更があります'
      alert('メモを保存できませんでした: ' + (err.message || err))
    }
  })
}

async function editTitle() {
  const next = prompt('タイトル', detail.detail?.title || detail.item.title)
  if (next == null) return
  const title = next.trim()
  if (!title) return
  try {
    await saveTitle(detail.key, title)
    detail.detail = { ...detail.detail, title }
    detail.item.title = title
    paintDetail()
  } catch (err) {
    alert('タイトルを保存できませんでした: ' + (err.message || err))
  }
}

async function addTag() {
  const name = prompt('追加するタグ')?.trim()
  if (!name) return
  const next = [...new Set([...(detail.detail?.tags || detail.item.tags || []), name])]
  await commitTags(next)
}

async function removeTag(tag) {
  const next = (detail.detail?.tags || detail.item.tags || []).filter((t) => t !== tag)
  await commitTags(next)
}

async function commitTags(next) {
  const before = detail.detail?.tags || detail.item.tags || []
  detail.detail = { ...detail.detail, tags: next }
  detail.item.tags = next
  paintDetail()
  try {
    await saveTags(detail.key, next)
    detailCache.set(detail.key, detail.detail)
  } catch (err) {
    detail.detail = { ...detail.detail, tags: before }
    detail.item.tags = before
    paintDetail()
    alert('タグを保存できませんでした: ' + (err.message || err))
  }
}

function copyCurrentTab() {
  const d = detail.detail || {}
  const text = detail.activeTab === 'raw' ? detail.transcript || '' : d[detail.activeTab] || ''
  navigator.clipboard.writeText(text)
  statusEl.textContent = 'コピーしました'
  setTimeout(() => paint(), 1500)
}

function openMoreMenu(e) {
  const menu = document.createElement('div')
  menu.className = 'menu'
  menu.innerHTML = `
    <button data-act="exclude">一覧から除外する</button>
    <button data-act="delete" class="danger">Notionから削除する</button>
  `
  document.body.appendChild(menu)
  const rect = e.currentTarget.getBoundingClientRect()
  menu.style.top = `${rect.bottom + 4}px`
  menu.style.right = `${window.innerWidth - rect.right}px`

  menu.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', async () => {
    menu.remove()
    if (btn.dataset.act === 'exclude') {
      if (!confirm('この動画を一覧から除外します。よろしいですか?')) return
      await setStatus(detail.key, '除外')
    } else {
      if (!confirm('Notionのページをゴミ箱へ移します。よろしいですか?')) return
      await deleteContent(detail.key)
    }
    items = items.filter((i) => i.key !== detail.key)
    setView(lastListView)
  }))

  setTimeout(() => {
    const once = (ev) => {
      if (!menu.contains(ev.target)) {
        menu.remove()
        document.removeEventListener('click', once)
      }
    }
    document.addEventListener('click', once)
  }, 0)
}

// ---- チャット ----

const chatByKey = new Map()

function setupChat() {
  const box = stageEl.querySelector('#chat-messages')
  const input = stageEl.querySelector('#chat-input')
  const send = stageEl.querySelector('#chat-send')
  if (!box || !input || !send) return

  const messages = chatByKey.get(detail.key) || []
  chatByKey.set(detail.key, messages)
  paintChat(box, messages)

  const post = async () => {
    const text = input.value.trim()
    if (!text) return
    const connection = connectionOf(loadSettings())
    if (!connection) {
      alert('AI接続が未設定です。')
      return
    }
    input.value = ''
    messages.push({ role: 'user', content: text })
    messages.push({ role: 'assistant', content: '' })
    paintChat(box, messages)

    try {
      await needTranscript()
      const system = `あなたは動画の内容について質問に答えるアシスタントです。
以下は「${detail.detail?.title || detail.item.title}」の文字起こしです。この範囲で答え、無い情報は「分かりません」と答えてください。

${detail.transcript.slice(0, 30000)}`
      let full = ''
      for await (const chunk of streamChat(connection, [
        { role: 'system', content: system },
        ...messages.slice(0, -1).map((m) => ({ role: m.role, content: m.content })),
      ])) {
        if (chunk.delta) {
          full += chunk.delta
          messages[messages.length - 1].content = full
          paintChat(box, messages)
        }
      }
      if (!full) messages[messages.length - 1].content = '(応答がありませんでした)'
    } catch (err) {
      messages[messages.length - 1].content = 'エラーが発生しました: ' + (err.message || err)
    }
    paintChat(box, messages)
  }

  send.addEventListener('click', post)
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
      e.preventDefault()
      post()
    }
  })
}

function paintChat(box, messages) {
  box.innerHTML = messages.map((m) => `
    <div class="chat-msg ${m.role}">
      <div class="bubble">${escapeHtml(m.content) || '<span class="muted">考えています...</span>'}</div>
    </div>
  `).join('')
  box.scrollTop = box.scrollHeight
}

// ============ アップロード ============

function openUpload() {
  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal">
        <div class="modal-head"><span>動画をアップロード</span><button class="btn-ghost btn-close" aria-label="閉じる"><i class="ti ti-x"></i></button></div>
        <div class="modal-body">
          <label>タイトル(省略時はファイル名)</label>
          <input id="up-title" class="input" />
          <label>動画ファイル</label>
          <input id="up-file" type="file" accept="video/*,audio/*" />
          <p class="foot-note">Google Drive の inbox へ送ります。文字起こしはMac側で行われ、終わると一覧に並びます。</p>
          <div class="progress"><div id="up-bar" class="progress-bar"></div></div>
          <p id="up-status" class="foot-note"></p>
        </div>
        <div class="modal-foot">
          <button class="btn btn-cancel">閉じる</button>
          <button class="btn btn-primary btn-send">アップロード</button>
        </div>
      </div>
    </div>
  `
  const close = () => (root.innerHTML = '')
  root.querySelector('.btn-close').addEventListener('click', close)
  root.querySelector('.btn-cancel').addEventListener('click', close)
  root.querySelector('.btn-send').addEventListener('click', async () => {
    const file = root.querySelector('#up-file').files?.[0]
    const status = root.querySelector('#up-status')
    const bar = root.querySelector('#up-bar')
    if (!file) {
      status.textContent = 'ファイルを選んでください'
      return
    }
    root.querySelector('.btn-send').disabled = true
    status.textContent = 'アップロードしています...'
    try {
      await uploadVideo(file, {
        title: root.querySelector('#up-title').value.trim(),
        onProgress: (ratio) => {
          bar.style.width = `${Math.round(ratio * 100)}%`
          status.textContent = `アップロード中 ${Math.round(ratio * 100)}%`
        },
      })
      status.textContent = '送信しました。文字起こしが終わると一覧に並びます。'
    } catch (err) {
      status.textContent = '失敗しました: ' + (err.message || err)
    }
    root.querySelector('.btn-send').disabled = false
  })
}

// ============ 起動 ============

$('search').addEventListener('input', (e) => {
  searchQuery = e.target.value
  if (view !== 'detail') paint()
})
document.querySelectorAll('.viewtab').forEach((t) => t.addEventListener('click', () => setView(t.dataset.view)))
$('open-vocab').addEventListener('click', openVocab)
$('open-upload').addEventListener('click', openUpload)
$('open-settings').addEventListener('click', () => openSettings(() => loadList()))
$('reload').addEventListener('click', loadList)

if (!isConfigured(loadConfig())) {
  library = { phase: 'error', message: '設定からGASのURLと共有トークンを入力してください' }
  paint()
  openSettings(() => loadList())
} else {
  loadList()
}
