import { parseSections, visibleHeading, sectionRank, isSectionHidden } from '../lib/sections.js'

export function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]))
}

function fmtDate(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`
}

/** 生成の進み具合。5つの点が サマリ/マインドマップ/分野別/応用/活用 に対応する */
function dotsHtml(item) {
  const flags = [
    { on: Boolean(item.summary), label: 'サマリ' },
    { on: item.has?.mindmap, label: 'マインドマップ' },
    { on: item.has?.fields, label: '分野別' },
    { on: item.has?.apply, label: '応用' },
    { on: item.has?.ideas, label: '活用' },
  ]
  const title = flags.map((f) => `${f.label}:${f.on ? '有' : '無'}`).join(' / ')
  return `<span class="dots" title="${title}">${flags.map((f) => `<i class="dot ${f.on ? 'on' : ''}"></i>`).join('')}</span>`
}

// ============ ライブラリ(一覧) ============

export function renderLibrary(container, items, state, handlers) {
  if (state.phase === 'loading') {
    container.innerHTML = '<p class="muted">読み込んでいます...</p>'
    return
  }
  if (state.phase === 'error') {
    container.innerHTML = `<p class="error-text">${escapeHtml(state.message)}</p><button class="btn btn-retry">もう一度読み込む</button>`
    container.querySelector('.btn-retry')?.addEventListener('click', handlers.onRetry)
    return
  }
  if (!items.length) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-movie" aria-hidden="true"></i>
        <p>動画がまだありません</p>
        <p class="empty-hint">右上の「アップロード」から動画を置くと、Mac側で文字起こしされてここに並びます</p>
      </div>`
    return
  }

  container.innerHTML = `
    <div class="cards">
      ${items.map((item) => `
        <article class="card" data-key="${escapeHtml(item.key)}">
          <div class="card-kind"><i class="ti ti-movie" aria-hidden="true"></i>${escapeHtml(item.kind || '動画')}</div>
          <h3 class="card-title">${escapeHtml(item.title)}</h3>
          ${item.summary ? `<p class="card-summary">${escapeHtml(item.summary)}</p>` : '<p class="card-summary muted">要約はまだありません</p>'}
          <div class="card-tags">${(item.tags || []).map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>
          <div class="card-foot">
            <span class="badge status-${escapeHtml(item.status || '')}">${escapeHtml(item.status || '')}</span>
            ${dotsHtml(item)}
            <span class="grow"></span>
            <span class="card-date">${fmtDate(item.createdAt)}</span>
          </div>
        </article>
      `).join('')}
    </div>
  `

  container.querySelectorAll('.card').forEach((el) => {
    el.addEventListener('click', () => handlers.onOpen(el.dataset.key))
  })
}

// ============ マインドマップ一覧(公開ONのものだけ) ============

/** マップ自体は重いのでカードには描かない。開いたときに1枚だけ描く */
export function renderMindmapGallery(container, items, state, handlers) {
  if (state.phase === 'loading') {
    container.innerHTML = '<p class="muted">読み込んでいます...</p>'
    return
  }
  if (!items.length) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="ti ti-sitemap" aria-hidden="true"></i>
        <p>公開中のマインドマップがありません</p>
        <p class="empty-hint">詳細のマインドマップタブで「公開する」を押すと、ここに並びます</p>
      </div>`
    return
  }

  container.innerHTML = `
    <div class="cards">
      ${items.map((item) => `
        <article class="card mm-card" data-key="${escapeHtml(item.key)}">
          <div class="card-kind"><i class="ti ti-sitemap" aria-hidden="true"></i>マインドマップ</div>
          <h3 class="card-title">${escapeHtml(item.title)}</h3>
          <div class="card-tags">${(item.tags || []).map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>
          <div class="card-foot">
            <span class="grow"></span>
            <span class="card-date">${fmtDate(item.createdAt)}</span>
            ${state.canEdit ? '<button class="btn-ghost btn-hide" aria-label="一覧から外す"><i class="ti ti-eye-off"></i></button>' : ''}
          </div>
        </article>
      `).join('')}
    </div>
  `

  container.querySelectorAll('.card').forEach((el) => {
    el.addEventListener('click', () => handlers.onOpen(el.dataset.key))
  })
  container.querySelectorAll('.btn-hide').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation() // カードのクリック(大きく見る)と競合させない
      handlers.onHide(btn.closest('.card').dataset.key)
    })
  })
}

// ============ アイデア一覧 ============

/** listIdeas の結果を1件1アイデアの配列に展開する */
export function flattenIdeas(items) {
  const out = []
  items.forEach((v) => {
    ;['apply', 'ideas'].forEach((kind) => {
      parseSections(v[kind]).forEach((s, i) => {
        out.push({
          id: `${v.key}:${kind}:${i}`,
          key: v.key,
          contentTitle: v.title,
          tags: v.tags || [],
          kind,
          sec: i,
          isPublic: !isSectionHidden(s.heading),
          rank: sectionRank(s.heading),
          heading: visibleHeading(s.heading) || '(無題)',
          body: s.body,
          points: s.points,
        })
      })
    })
  })
  return out
}

const IDEA_KINDS = [
  { id: 'all', label: 'すべて' },
  { id: 'apply', label: '応用' },
  { id: 'ideas', label: '活用' },
]

export function renderIdeaGallery(container, entries, state, handlers) {
  if (state.phase === 'loading') {
    container.innerHTML = '<p class="muted">読み込んでいます...</p>'
    return
  }
  if (state.phase === 'error') {
    container.innerHTML = `<p class="error-text">${escapeHtml(state.message)}</p><button class="btn btn-retry">もう一度読み込む</button>`
    return
  }

  const filters = `
    <div class="idea-filters">
      ${IDEA_KINDS.map((k) => `<button class="chip idea-kind ${k.id === state.kind ? 'on' : ''}" data-kind="${k.id}">${k.label}</button>`).join('')}
      <span class="grow"></span>
      <span class="foot-note">${entries.length}件</span>
    </div>`

  const body = entries.length
    ? `<div class="cards">
        ${entries.map((e) => `
          <article class="card idea-card" data-id="${escapeHtml(e.id)}">
            <div class="card-kind">
              <i class="ti ti-bulb" aria-hidden="true"></i>${e.kind === 'apply' ? '応用' : '活用'}
              <span class="grow"></span>
              ${e.rank ? `<span class="rank">${[1, 2, 3].map((n) => `<span class="star ${n <= e.rank ? 'on' : ''}">★</span>`).join('')}</span>` : ''}
            </div>
            <h3 class="card-title">${escapeHtml(e.heading)}</h3>
            ${e.body ? `<p class="card-summary">${escapeHtml(e.body)}</p>` : ''}
            ${e.points.length ? `<ul class="sec-points">${e.points.map((p) => `<li>${escapeHtml(p)}</li>`).join('')}</ul>` : ''}
            <div class="card-tags">${e.tags.map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>
            <div class="card-foot">
              <span class="idea-from">${escapeHtml(e.contentTitle || '')}</span>
              <span class="grow"></span>
              ${state.canEdit ? '<button class="btn-ghost btn-hide" aria-label="一覧から外す"><i class="ti ti-eye-off"></i></button>' : ''}
            </div>
          </article>
        `).join('')}
      </div>`
    : `<div class="empty-state">
         <i class="ti ti-bulb" aria-hidden="true"></i>
         <p>公開中のアイデアがありません</p>
         <p class="empty-hint">詳細の「応用」「活用」タブで生成すると、ここに並びます</p>
       </div>`

  container.innerHTML = filters + body

  container.querySelectorAll('.idea-kind').forEach((btn) =>
    btn.addEventListener('click', () => handlers.onKind(btn.dataset.kind))
  )
  container.querySelectorAll('.idea-card').forEach((el) => {
    el.addEventListener('click', () => handlers.onOpen(el.dataset.id))
  })
  container.querySelectorAll('.btn-hide').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation() // カードのクリック(詳細を開く)と競合させない
      handlers.onHide(btn.closest('.idea-card').dataset.id)
    })
  })
}

// ============ 詳細 ============

export const TABS = [
  { id: 'summary', label: 'サマリ' },
  { id: 'mindmap', label: 'マインドマップ' },
  { id: 'fields', label: '分野別' },
  { id: 'apply', label: '応用' },
  { id: 'ideas', label: '活用' },
  { id: 'raw', label: '原文' },
  { id: 'memo', label: 'メモ' },
  { id: 'chat', label: 'チャット' },
]

/** 分野別・応用・活用の本文。見出し+本文+箇条書きで出す */
function sectionsHtml(text, { ranked = false, editable = false } = {}) {
  const sections = parseSections(text)
  if (!sections.length) return '<p class="muted">まだありません</p>'
  return sections.map((s, i) => {
    const shown = !isSectionHidden(s.heading)
    return `
    <section class="sec ${shown ? '' : 'sec-hidden'}" data-sec="${i}">
      <div class="sec-head">
        <h4>${escapeHtml(visibleHeading(s.heading) || '(無題)')}</h4>
        ${ranked ? rankHtml(i, sectionRank(s.heading), editable) : ''}
        ${ranked && editable ? `<button class="btn-ghost sec-pub" data-sec="${i}" data-on="${shown ? '1' : '0'}" aria-label="${shown ? 'アイデア一覧から外す' : 'アイデア一覧に戻す'}"><i class="ti ti-eye${shown ? '' : '-off'}"></i></button>` : ''}
      </div>
      ${s.body ? `<p class="sec-body">${escapeHtml(s.body).replace(/\n/g, '<br>')}</p>` : ''}
      ${s.points.length ? `<ul class="sec-points">${s.points.map((p) => `<li>${escapeHtml(p)}</li>`).join('')}</ul>` : ''}
    </section>
  `
  }).join('')
}

/** マインドマップ一覧に並べるかの切り替え */
function publishHtml(on) {
  return `<button class="btn btn-publish" data-on="${on ? '1' : '0'}"><i class="ti ti-${on ? 'eye-off' : 'eye'}" aria-hidden="true"></i>${on ? '一覧から外す' : 'マインドマップ一覧に公開'}</button>`
}

/** 価値の目安の3つ星。編集できるときは星を押して変えられる */
function rankHtml(index, rank, editable) {
  const stars = [1, 2, 3].map((n) => (editable
    ? `<button class="star ${n <= rank ? 'on' : ''}" data-sec="${index}" data-rank="${n === rank ? 0 : n}" aria-label="星${n}">★</button>`
    : `<span class="star ${n <= rank ? 'on' : ''}">★</span>`)).join('')
  return `<span class="rank" title="${rank ? `星${rank}` : '未設定'}">${stars}</span>`
}

export function renderDetail(container, item, state) {
  const d = state.detail || {}
  const tab = state.activeTab || 'summary'
  const canEdit = state.canEdit

  const head = `
    <div class="detail-head">
      <button class="btn btn-back"><i class="ti ti-arrow-left" aria-hidden="true"></i>一覧</button>
      <h2 class="detail-title">${escapeHtml(d.title || item.title)}</h2>
      ${canEdit ? '<button class="btn-ghost btn-edit-title" aria-label="タイトルを編集"><i class="ti ti-edit" aria-hidden="true"></i></button>' : ''}
      <span class="grow"></span>
      ${item.driveUrl ? `<a class="btn" href="${escapeHtml(item.driveUrl)}" target="_blank" rel="noopener"><i class="ti ti-player-play" aria-hidden="true"></i>動画を開く</a>` : ''}
      ${canEdit ? '<button class="btn btn-generate-all"><i class="ti ti-sparkles" aria-hidden="true"></i>すべて生成</button>' : ''}
      <button class="btn-ghost btn-more" aria-label="その他"><i class="ti ti-dots" aria-hidden="true"></i></button>
    </div>
    <div class="detail-meta">
      ${escapeHtml(d.file || item.file || '')}
      ${d.model ? ` · ${escapeHtml(d.model)}` : ''}
      ${d.rawCount ? ` · 原文${Number(d.rawCount).toLocaleString()}字` : ''}
    </div>
    <div class="tags-row">
      ${(d.tags || item.tags || []).map((t) => `<span class="tag">${escapeHtml(t)}${canEdit ? `<i class="ti ti-x tag-remove" data-tag="${escapeHtml(t)}"></i>` : ''}</span>`).join('')}
      ${canEdit ? '<button class="tag-add"><i class="ti ti-plus" aria-hidden="true"></i></button>' : ''}
    </div>
    <div class="tabs">
      ${TABS.map((t) => `<button class="tab ${t.id === tab ? 'on' : ''}" data-tab="${t.id}">${t.label}</button>`).join('')}
    </div>
  `

  let panel
  if (state.busyStage) {
    panel = `<p class="muted">${escapeHtml(state.busyLabel || '生成しています...')}</p><pre class="stream">${escapeHtml(state.busyText || '')}</pre>`
  } else if (tab === 'summary') {
    panel = d.summary ? `<p class="sec-body">${escapeHtml(d.summary).replace(/\n/g, '<br>')}</p>` : '<p class="muted">まだありません</p>'
  } else if (tab === 'mindmap') {
    panel = '<div id="mindmap-host" class="mindmap-host"></div><p class="hint">↑↓で移動 / ←→で開閉 / スペースで編集 / Tabで子を追加 / Enterで同じ階層に追加 / Deleteで削除</p>'
  } else if (tab === 'fields') {
    panel = sectionsHtml(d.fields)
  } else if (tab === 'apply' || tab === 'ideas') {
    panel = sectionsHtml(d[tab], { ranked: true, editable: canEdit })
  } else if (tab === 'raw') {
    panel = state.transcript === null
      ? '<p class="muted">読み込んでいます...</p>'
      : `<pre class="raw">${escapeHtml(state.transcript || '(原文がありません)')}</pre>`
  } else if (tab === 'memo') {
    panel = `<textarea id="memo-input" class="memo" placeholder="自由に記入できます">${escapeHtml(state.memoDraft ?? d.memo ?? '')}</textarea>`
  } else {
    panel = '<div id="chat-messages" class="chat-messages"></div>'
  }

  const stage = TABS.find((t) => t.id === tab && ['summary', 'mindmap', 'fields', 'apply', 'ideas'].includes(t.id))
  const foot = tab === 'memo'
    ? `<span class="foot-note" id="memo-status">${state.memoDirty ? '未保存の変更があります' : ''}</span>
       <span class="grow"></span>
       <button class="btn btn-memo-save">メモを保存</button>`
    : tab === 'chat'
      ? `<div class="composer">
           <textarea id="chat-input" class="chat-input" rows="1" placeholder="この動画について質問する(Shift+Enterで改行)"></textarea>
           <button id="chat-send" class="btn" aria-label="送信"><i class="ti ti-send" aria-hidden="true"></i></button>
         </div>`
      : `${canEdit && stage ? `<button class="btn btn-regen" data-stage="${stage.id}"><i class="ti ti-refresh" aria-hidden="true"></i>この項目を作り直す</button>` : ''}
         ${canEdit && tab === 'mindmap' ? publishHtml(d.isPublic ?? item.isPublic) : ''}
         <span class="grow"></span>
         <button class="btn btn-copy"><i class="ti ti-copy" aria-hidden="true"></i>コピー</button>`

  container.innerHTML = `
    <div class="detail">
      <div class="detail-fixed">${head}</div>
      <div class="detail-scroll" id="detail-scroll">
        ${state.phase === 'loading' ? '<p class="muted">読み込んでいます...</p>' : `<div class="panel">${panel}</div>`}
      </div>
      <div class="detail-foot">${foot}</div>
    </div>
  `
}
