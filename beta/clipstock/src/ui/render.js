import { parseSections } from '../lib/sections.js'
import { splitLabel, splitTranscript, formatTimecode, youtubeUrlAt } from '../lib/timecode.js'
import { STATUS_SUMMARIZED, STATUS_DONE, STATUS_NEW } from '../lib/filters.js'

export function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]))
}

/** 検索語を大文字小文字区別なくハイライトしたhtmlを返す */
export function highlightText(text, query) {
  const safe = escapeHtml(text ?? '')
  const q = (query ?? '').trim()
  if (!q) return safe
  const escaped = escapeHtml(q).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return safe.replace(new RegExp(escaped, 'gi'), (m) => `<mark>${m}</mark>`)
}

function fmtDate(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`
}

/** YouTubeのURLから動画IDを取り、サムネイルURLを組み立てる(サムネイル未設定の保険) */
function fallbackThumb(url) {
  const m = String(url ?? '').match(/(?:v=|youtu\.be\/|shorts\/|embed\/)([\w-]{11})/)
  return m ? `https://i.ytimg.com/vi/${m[1]}/mqdefault.jpg` : ''
}

function thumbHtml(item, extraClass = '') {
  const src = item.thumb || fallbackThumb(item.url)
  if (!src) {
    return `<div class="thumb thumb-blank ${extraClass}"><i class="ti ti-video-off" aria-hidden="true"></i></div>`
  }
  return `<div class="thumb ${extraClass}"><img src="${escapeHtml(src)}" alt="" loading="lazy" /></div>`
}

// ============ タグレール ============

export function renderTagRail(container, state, handlers) {
  const { tagOptions, selectedTags } = state
  if (!tagOptions.length) {
    container.innerHTML = ''
    return
  }
  container.innerHTML = `
    <div class="tag-rail">
      ${selectedTags.size ? '<button class="chip chip-clear">絞り込みを外す</button>' : ''}
      ${tagOptions
        .map(
          (o) => `<button class="chip ${o.selected ? 'on' : ''} ${o.disabled ? 'off' : ''}" data-tag="${escapeHtml(o.tag)}" ${o.disabled ? 'disabled' : ''}>
            ${escapeHtml(o.tag)}<span class="chip-count">${o.count}</span>
          </button>`
        )
        .join('')}
    </div>
  `
  container.querySelector('.chip-clear')?.addEventListener('click', handlers.onClearTags)
  container.querySelectorAll('.chip[data-tag]').forEach((el) => {
    el.addEventListener('click', () => handlers.onToggleTag(el.dataset.tag))
  })
}

// ============ ライブラリ(サムネイルのグリッド) ============

/** 進捗の見え方。カードの帯とドットで「どこまで出来ているか」を出す */
function progressHtml(item) {
  const steps = [
    { on: item.status === STATUS_SUMMARIZED || Boolean(item.summary), label: 'サマリ' },
    { on: item.has?.mindmap, label: 'マインドマップ' },
    { on: item.has?.fields, label: '分野別' },
    { on: item.has?.apply, label: '応用' },
    { on: item.has?.ideas, label: '活用' },
  ]
  return `<span class="dots" title="${steps.map((s) => `${s.label}:${s.on ? '有' : '無'}`).join(' / ')}">${steps
    .map((s) => `<i class="dot ${s.on ? 'on' : ''}"></i>`)
    .join('')}</span>`
}

export function renderLibrary(container, items, state, handlers) {
  if (!items.length) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-movie" aria-hidden="true"></i>
        <p>該当する動画がありません</p>
        <p class="empty-hint">Notionの動画DBにURLを追加すると、次のバッチで文字起こしまで進みます</p>
      </div>`
    return
  }

  container.innerHTML = `
    <div class="grid">
      ${items
        .map(
          (item) => `
        <article class="card ${state.seen(item.key) ? '' : 'unseen'}" data-key="${escapeHtml(item.key)}" tabindex="0">
          ${thumbHtml(item)}
          <div class="card-body">
            <h3 class="card-title">${highlightText(item.title, state.searchQuery)}</h3>
            <p class="card-summary">${item.summary ? highlightText(item.summary.slice(0, 110), state.searchQuery) : '<span class="muted">要約はまだありません</span>'}</p>
            <div class="card-foot">
              <span class="badge s-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
              ${progressHtml(item)}
              <span class="card-date">${fmtDate(item.createdAt)}</span>
            </div>
            ${state.showTags && item.tags?.length ? `<div class="card-tags">${item.tags.map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>` : ''}
          </div>
        </article>`
        )
        .join('')}
    </div>
  `

  container.querySelectorAll('.card').forEach((el) => {
    const open = () => handlers.onOpen(el.dataset.key)
    el.addEventListener('click', open)
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault()
        open()
      }
    })
  })
}

// ============ セクション表示(分野別 / 応用 / 活用アイデア) ============

/**
 * "[12:34]" を動画のその時刻へ飛ぶリンクにする。
 * 動画URLが無い、または時刻が無い項目には何も出さない
 * (リンクの有無自体が「根拠が原文で確認できたか」の印になっている)。
 */
function timeChip(at, videoUrl) {
  if (at === null || at === undefined) return ''
  const href = youtubeUrlAt(videoUrl, at)
  const label = formatTimecode(at)
  if (!href) return `<span class="tc">${label}</span>`
  return `<a class="tc tc-link" href="${escapeHtml(href)}" target="_blank" rel="noopener" title="${label} から再生"><i class="ti ti-player-play" aria-hidden="true"></i>${label}</a>`
}

function sectionsHtml(text, { numbered = false, videoUrl = '' } = {}) {
  const sections = parseSections(text)
  if (!sections.length) return '<p class="empty-section">まだありません</p>'
  return `<div class="sections">${sections
    .map((s, i) => {
      const head = splitLabel(s.heading || '(無題)')
      return `
    <section class="sec">
      <h4 class="sec-head">${numbered ? `<span class="sec-num">${i + 1}</span>` : ''}${escapeHtml(head.text)}${timeChip(head.at, videoUrl)}</h4>
      ${s.body ? `<p class="sec-body">${escapeHtml(s.body).replace(/\n/g, '<br />')}</p>` : ''}
      ${s.points.length
        ? `<ul class="sec-points">${s.points
            .map((p) => {
              const point = splitLabel(p)
              return `<li>${escapeHtml(point.text)}${timeChip(point.at, videoUrl)}</li>`
            })
            .join('')}</ul>`
        : ''}
    </section>`
    })
    .join('')}</div>`
}

/** 原文タブ。タイムスタンプがあれば各かたまりの頭を再生リンクにする */
function transcriptHtml(text, videoUrl) {
  const segments = splitTranscript(text)
  if (!segments.length) return '<p class="empty-section">原文がまだありません</p>'
  if (segments.every((seg) => seg.at === null)) {
    return `<pre class="raw-text">${escapeHtml(text)}</pre>`
  }
  return `<div class="tr-list">${segments
    .map(
      (seg) => `<div class="tr-row">
        ${timeChip(seg.at, videoUrl)}
        <p class="tr-text">${escapeHtml(seg.text)}</p>
      </div>`
    )
    .join('')}</div>`
}

// ============ 詳細 ============

export const TABS = [
  { id: 'summary', label: 'サマリ' },
  { id: 'mindmap', label: 'マインドマップ' },
  { id: 'fields', label: '分野別' },
  { id: 'apply', label: '応用' },
  { id: 'ideas', label: '活用' },
  { id: 'memo', label: 'メモ' },
  { id: 'chat', label: 'チャット' },
  { id: 'raw', label: '原文' },
]

/** どのタブがAI生成物か。ここに載っているタブには「作り直す」ボタンを出す */
const STAGE_OF_TAB = { summary: 'core', mindmap: 'core', fields: 'fields', apply: 'apply', ideas: 'apply' }

export function detailHtml(item, state) {
  const d = state.detail || {}
  const tab = state.activeTab
  const canEdit = state.canEdit

  const editable = ['summary', 'mindmap', 'fields', 'apply', 'ideas'].includes(tab)
  const stage = STAGE_OF_TAB[tab]

  const head = `
    <div class="detail-head">
      <button class="btn-ghost btn-back" aria-label="一覧に戻る"><i class="ti ti-arrow-left" aria-hidden="true"></i></button>
      <div class="detail-headline">
        <h2 class="detail-title">${escapeHtml(item.title)}</h2>
        <div class="detail-meta">
          <span class="badge s-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
          <span>${fmtDate(item.createdAt)}</span>
          ${d.rawCount ? `<span>原文 ${d.rawCount.toLocaleString()}字</span>` : ''}
          ${d.model ? `<span class="model-badge">${escapeHtml(d.model)}</span>` : ''}
        </div>
      </div>
      ${item.url ? `<a class="btn-ghost" href="${escapeHtml(item.url)}" target="_blank" rel="noopener" aria-label="YouTubeで開く"><i class="ti ti-external-link" aria-hidden="true"></i></a>` : ''}
      ${canEdit ? '<button class="btn-ghost btn-more" aria-label="その他の操作"><i class="ti ti-dots" aria-hidden="true"></i></button>' : ''}
    </div>
    <div class="detail-hero">
      ${item.url ? `<a href="${escapeHtml(item.url)}" target="_blank" rel="noopener">${thumbHtml(item, 'thumb-hero')}</a>` : thumbHtml(item, 'thumb-hero')}
      <div class="hero-side">
        <div class="tag-row">
          ${(state.tags || []).map((t) => `<span class="tag${canEdit ? ' tag-edit' : ''}" data-tag="${escapeHtml(t)}">${escapeHtml(t)}${canEdit ? '<i class="ti ti-x" aria-hidden="true"></i>' : ''}</span>`).join('')}
          ${canEdit ? '<button class="btn-ghost btn-tag-add" aria-label="タグを追加"><i class="ti ti-plus" aria-hidden="true"></i></button>' : ''}
        </div>
        ${canEdit ? '<button class="btn btn-primary btn-generate-all"><i class="ti ti-sparkles" aria-hidden="true"></i>すべて生成</button>' : ''}
      </div>
    </div>
    <nav class="tabs">
      ${TABS.map((t) => `<button class="tab ${t.id === tab ? 'on' : ''}" data-tab="${t.id}">${t.label}${state.tabHasContent(t.id) ? '' : '<i class="tab-empty" aria-hidden="true"></i>'}</button>`).join('')}
    </nav>
  `

  let panel
  if (state.phase === 'loading') {
    panel = '<p class="muted">読み込み中...</p>'
  } else if (state.phase === 'error') {
    panel = `<p class="error-text">${escapeHtml(state.message)}</p><button class="btn btn-retry">もう一度読み込む</button>`
  } else if (state.busyStage) {
    panel = `
      <div class="gen-progress">
        <p class="muted">${escapeHtml(state.busyLabel || '生成中')}</p>
        <pre class="gen-stream">${escapeHtml((state.busyText || '').slice(-1200))}</pre>
      </div>`
  } else {
    panel = renderPanel(item, state, tab, d)
  }

  const foot =
    tab === 'memo'
      ? `<span class="foot-note" id="memo-status">${state.memoDirty ? '未保存の変更があります' : ''}</span>
         <span class="grow"></span>
         <button class="btn btn-memo-save">メモを保存</button>`
      : tab === 'chat'
        ? `<div class="composer">
             <div class="composer-row">
               <div class="ctx-switch">
                 <button class="ctx-btn" data-ctx="summary">要約</button>
                 <button class="ctx-btn on" data-ctx="raw">原文</button>
               </div>
               <span class="foot-note" id="ctx-count"></span>
             </div>
             <div class="composer-row">
               <textarea id="chat-input" class="chat-input" rows="1" placeholder="この動画について質問する(Shift+Enterで改行)"></textarea>
               <button id="chat-send" class="btn btn-primary" aria-label="送信"><i class="ti ti-send" aria-hidden="true"></i></button>
             </div>
           </div>`
        : `${canEdit && stage ? `<button class="btn btn-regen" data-stage="${stage}"><i class="ti ti-refresh" aria-hidden="true"></i>この項目を作り直す</button>` : ''}
           ${canEdit && editable ? '<button class="btn btn-edit-field"><i class="ti ti-edit" aria-hidden="true"></i>手で直す</button>' : ''}
           <span class="grow"></span>
           ${tab === 'mindmap' && state.tabHasContent('mindmap') ? '<button class="btn btn-mm-full"><i class="ti ti-arrows-maximize" aria-hidden="true"></i>大きく見る</button>' : ''}
           ${['summary', 'mindmap', 'fields', 'apply', 'ideas', 'raw'].includes(tab) ? '<button class="btn btn-copy"><i class="ti ti-copy" aria-hidden="true"></i>コピー</button>' : ''}`

  return `
    <div class="detail">
      <div class="detail-fixed">${head}</div>
      <div class="detail-scroll" id="detail-scroll">${panel}</div>
      <div class="detail-foot">${foot}</div>
    </div>
  `
}

function renderPanel(item, state, tab, d) {
  switch (tab) {
    case 'summary':
      return d.summary
        ? `<p class="prose">${escapeHtml(d.summary).replace(/\n/g, '<br />')}</p>`
        : emptyPanel(item, 'サマリ', 'core')
    case 'mindmap':
      return '<div id="mindmap-host" class="mindmap-host"></div>'
    case 'fields':
      return d.fields ? sectionsHtml(d.fields, { videoUrl: item.url }) : emptyPanel(item, '分野別要約', 'fields')
    case 'apply':
      return d.apply
        ? sectionsHtml(d.apply, { numbered: true, videoUrl: item.url })
        : emptyPanel(item, '応用', 'apply')
    case 'ideas':
      return d.ideas ? sectionsHtml(d.ideas, { videoUrl: item.url }) : emptyPanel(item, '活用アイデア', 'apply')
    case 'memo':
      return `<textarea id="memo-input" class="memo-input" placeholder="気づいたこと、あとで試すこと、関連する話などを自由に">${escapeHtml(state.memoDraft ?? d.memo ?? '')}</textarea>`
    case 'chat':
      return '<div id="chat-log" class="chat-log"></div>'
    case 'raw':
      return state.transcript === null || state.transcript === undefined
        ? '<p class="muted">原文を読み込んでいます...</p>'
        : state.transcript
          ? transcriptHtml(state.transcript, item.url)
          : '<p class="empty-section">原文がまだありません。状態が「完了」になるとここに入ります</p>'
    default:
      return ''
  }
}

function emptyPanel(item, label, stage) {
  if (item.status === STATUS_NEW) {
    return `<p class="empty-section">${escapeHtml(label)}はまだありません。まず文字起こしが必要です(状態が「完了」になるまで待つ)</p>`
  }
  if (item.status === STATUS_DONE) {
    return `<p class="empty-section">${escapeHtml(label)}はまだありません。下の「この項目を作り直す」で生成できます</p>`
  }
  return `<p class="empty-section">${escapeHtml(label)}はまだありません</p>`
}

// ============ アイデア一覧(横断して掘り起こす画面) ============

/**
 * 応用と活用アイデアを動画をまたいで1本のフィードにする。
 * 動画単位で見ていると「あの話、何かに使えそう」で終わってしまうため、
 * アイデア側から入って動画に戻れる導線を作るのがこの画面の役目。
 */
export function renderIdeas(container, entries, state, handlers) {
  if (state.phase === 'loading') {
    container.innerHTML = '<p class="muted">アイデアを集めています...</p>'
    return
  }
  if (state.phase === 'error') {
    container.innerHTML = `<p class="error-text">${escapeHtml(state.message)}</p><button class="btn btn-retry">もう一度読み込む</button>`
    return
  }
  if (!entries.length) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-bulb" aria-hidden="true"></i>
        <p>まだアイデアがありません</p>
        <p class="empty-hint">ライブラリで動画を開いて「すべて生成」すると、ここに応用と活用アイデアが集まります</p>
      </div>`
    return
  }

  container.innerHTML = `
    <div class="idea-toolbar">
      <div class="seg">
        <button class="seg-btn ${state.kind === 'all' ? 'on' : ''}" data-kind="all">すべて</button>
        <button class="seg-btn ${state.kind === 'apply' ? 'on' : ''}" data-kind="apply">ビジネス</button>
        <button class="seg-btn ${state.kind === 'ideas' ? 'on' : ''}" data-kind="ideas">面白い活用</button>
      </div>
      <button class="btn btn-shuffle"><i class="ti ti-dice" aria-hidden="true"></i>掘り起こす</button>
    </div>
    <div class="idea-feed">
      ${entries
        .map(
          (e) => `
        <article class="idea" data-key="${escapeHtml(e.key)}">
          <div class="idea-kind ${e.kind}">${e.kind === 'apply' ? 'ビジネス' : '活用'}</div>
          <h4 class="idea-title">${escapeHtml(e.heading)}</h4>
          ${e.body ? `<p class="idea-body">${escapeHtml(e.body)}</p>` : ''}
          ${e.points.length ? `<ul class="sec-points">${e.points.map((p) => `<li>${escapeHtml(splitLabel(p).text)}</li>`).join('')}</ul>` : ''}
          <button class="idea-source">
            <i class="ti ti-movie" aria-hidden="true"></i>${escapeHtml(e.videoTitle)}
          </button>
        </article>`
        )
        .join('')}
    </div>
  `

  container.querySelectorAll('.seg-btn').forEach((el) => {
    el.addEventListener('click', () => handlers.onKind(el.dataset.kind))
  })
  container.querySelector('.btn-shuffle')?.addEventListener('click', handlers.onShuffle)
  container.querySelectorAll('.idea-source').forEach((el) => {
    el.addEventListener('click', () => handlers.onOpen(el.closest('.idea').dataset.key))
  })
}

/** listIdeas の結果を1件1アイデアのフィード用配列に展開する */
export function flattenIdeas(items) {
  const out = []
  items.forEach((v) => {
    ;['apply', 'ideas'].forEach((kind) => {
      parseSections(v[kind]).forEach((s, i) => {
        out.push({
          key: v.key,
          videoTitle: v.title,
          tags: v.tags || [],
          kind,
          heading: s.heading || '(無題)',
          body: s.body,
          points: s.points,
          id: `${v.key}:${kind}:${i}`,
        })
      })
    })
  })
  return out
}
