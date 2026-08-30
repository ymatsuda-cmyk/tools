/**
 * ツールバー(月ナビ / タグフィルタ)を描画する。
 * 検索欄とタグ表示トグルはトップバーの永続DOMに移したためここには含まない
 * (毎回作り直すと入力中にフォーカスが外れるため)。
 * handlers: { onPrevMonth, onNextMonth, onToggleTag(tag) }
 */
export function renderToolbar(container, state, handlers) {
  container.innerHTML = `
    <div class="toolbar">
      <div class="toolbar-row">
        <div class="month-nav">
          <button class="month-prev btn-ghost" aria-label="前月"><i class="ti ti-chevron-left" aria-hidden="true"></i></button>
          <span class="month-label">${escapeHtml(state.monthLabel)}</span>
          <button class="month-next btn-ghost" aria-label="翌月"><i class="ti ti-chevron-right" aria-hidden="true"></i></button>
        </div>
      </div>
      ${state.tagOptions.length ? `<div class="tag-filter-row">${state.tagOptions.map((o) => `
        <span class="tag-chip filter-chip ${o.selected ? 'selected' : ''} ${o.disabled ? 'disabled' : ''}" data-tag="${escapeHtml(o.tag)}">
          ${escapeHtml(o.tag)}${o.selected ? '<i class="ti ti-x" aria-hidden="true"></i>' : ''}
        </span>
      `).join('')}</div>` : ''}
    </div>
  `

  container.querySelector('.month-prev').addEventListener('click', handlers.onPrevMonth)
  container.querySelector('.month-next').addEventListener('click', handlers.onNextMonth)
  container.querySelectorAll('.filter-chip:not(.disabled)').forEach((el) => {
    el.addEventListener('click', () => handlers.onToggleTag(el.dataset.tag))
  })
}

function fmtDate(iso) {
  const d = new Date(iso)
  return d.toLocaleDateString('ja-JP', { month: 'long', day: 'numeric', weekday: 'short' })
}
function fmtTime(iso) {
  const d = new Date(iso)
  return d.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
}
function groupKey(iso) {
  return new Date(iso).toDateString()
}

/**
 * 一覧を描画する。onSelect(item, itemEl) がクリック時に呼ばれる。
 */
export function renderList(container, items, selectedKey, onSelect, showTags = false, opts = {}) {
  const { showPermissions = false, selectable = false, selectedIds = new Set() } = opts
  container.innerHTML = ''
  const sorted = [...items].sort((a, b) => new Date(b.date) - new Date(a.date))

  let lastGroup = null
  for (const item of sorted) {
    const g = groupKey(item.date)
    if (g !== lastGroup) {
      const label = document.createElement('div')
      label.className = 'date-group-label'
      label.textContent = fmtDate(item.date)
      container.appendChild(label)
      lastGroup = g
    }

    const tagsLine = showTags && item.tags?.length
      ? `<div class="list-item-tags">${item.tags.map((t) => `<span class="tag-chip small">${escapeHtml(t)}</span>`).join('')}</div>`
      : ''

    const permBadge = showPermissions
      ? ((item.permissions || []).length
          ? (item.permissions || []).map((p) => `<span class="perm-chip">${escapeHtml(p)}</span>`).join('')
          : '<span class="perm-chip none">未設定</span>')
      : ''

    const checkbox = selectable
      ? `<input type="checkbox" class="row-select" data-key="${escapeHtml(item.key)}" ${selectedIds.has(item.key) ? 'checked' : ''} />`
      : ''

    const row = document.createElement('div')
    row.className = 'list-item' + (item.key === selectedKey ? ' selected' : '')
    row.dataset.key = item.key
    if (selectable) row.classList.add('selectable')
    row.innerHTML = `
      ${checkbox}
      <div class="list-item-time">${fmtTime(item.date)}</div>
      <div class="list-item-body">
        <div class="list-item-title">${escapeHtml(item.title)}${permBadge}</div>
        <div class="list-item-meta">
          <span class="badge status-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
          ${escapeHtml(item.duration || '')}
        </div>
        ${tagsLine}
      </div>
    `
    row.addEventListener('click', (e) => {
      if (e.target.classList.contains('row-select')) return
      onSelect(item, row)
    })
    container.appendChild(row)
  }

  if (!sorted.length) {
    container.innerHTML = '<div class="empty-state"><i class="ti ti-inbox" aria-hidden="true"></i><p>該当する議事録がありません</p></div>'
  }
}

/**
 * 詳細ペインの中身を組み立てて返す(HTML文字列)。
 * state.phase: 'loading' | 'no-summary' | 'ready' | 'generating' | 'error'
 */
function tagsHtml(tags, canEdit) {
  if (tags === undefined) return ''
  const chips = tags.map((t) =>
    canEdit
      ? `<span class="tag-chip" data-tag="${escapeHtml(t)}">${escapeHtml(t)}<i class="ti ti-x tag-remove" aria-hidden="true"></i></span>`
      : `<span class="tag-chip">${escapeHtml(t)}</span>`
  ).join('')
  const addBtn = canEdit
    ? '<button class="tag-add-btn" aria-label="タグを追加"><i class="ti ti-plus" aria-hidden="true"></i></button>'
    : ''
  if (!chips && !addBtn) return ''
  return `<div class="tag-row">${chips}${addBtn}</div>`
}

export function renderDetailHtml(item, state) {
  const editBtn = (field, label) =>
    state.canEdit
      ? `<button class="btn-ghost btn-edit" data-field="${field}" aria-label="${label}を編集"><i class="ti ti-edit" aria-hidden="true"></i></button>`
      : ''

  const header = `
    <div class="detail-header">
      <span class="detail-title">${escapeHtml(item.title)}</span>
      ${state.canEdit ? '<button class="btn-ghost btn-edit-title" aria-label="タイトルを編集"><i class="ti ti-edit" aria-hidden="true"></i></button>' : ''}
      <span class="badge status-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
    </div>
    <div class="detail-meta">${fmtDate(item.date)} ${fmtTime(item.date)} · ${escapeHtml(item.duration || '')}</div>
    ${tagsHtml(state.tags, state.canEdit)}
  `

  if (state.phase === 'loading') {
    return header + `<p class="summary-text">読み込み中...</p>`
  }

  if (state.phase === 'error') {
    return header + `<p class="error-text">${escapeHtml(state.message)}</p>
      <div class="actions-row">
        <button class="btn btn-retry"><i class="ti ti-refresh" aria-hidden="true"></i>再試行</button>
        ${state.canEdit ? '<button class="btn btn-retranscribe"><i class="ti ti-microphone" aria-hidden="true"></i>文字起こし</button>' : ''}
        <button class="btn btn-raw"><i class="ti ti-file-text" aria-hidden="true"></i>原文表示</button>
      </div>`
  }

  if (state.phase === 'generating') {
    return header + `<p class="summary-text">${escapeHtml(state.progress || '要約を生成しています...')}</p>`
  }

  if (state.phase === 'no-summary') {
    return header + `
      <p class="summary-text" style="color:var(--text-muted)">この議事録の要約はまだありません。</p>
      <div class="actions-row">
        <button class="btn btn-generate"><i class="ti ti-sparkles" aria-hidden="true"></i>要約を生成</button>
        ${state.canEdit ? '<button class="btn btn-retranscribe"><i class="ti ti-microphone" aria-hidden="true"></i>文字起こし</button>' : ''}
        <button class="btn btn-raw"><i class="ti ti-file-text" aria-hidden="true"></i>原文表示</button>
      </div>
    `
  }

  // ready
  const s = state.summary
  const agenda = s.detail?.agenda || []
  const decisions = s.detail?.decisions || []
  const todos = (s.detail?.todos || []).map((t) => (typeof t === 'string' ? { text: t, done: false } : t))
  const topics = s.detail?.topics || []

  const agendaHtml = agenda.length ? `
    <div class="section-label">議事${editBtn('agenda', '議事')}</div>
    <div class="agenda-list">
      ${agenda.map((a, i) => `
        <div class="agenda-item">
          <div class="agenda-topic"><span class="agenda-num">${i + 1}</span>${escapeHtml(a.topic || '')}</div>
          ${(a.points || []).length ? `<ul class="agenda-points">${a.points.map((p) => `<li>${escapeHtml(p)}</li>`).join('')}</ul>` : ''}
          ${a.outcome ? `<div class="agenda-outcome"><i class="ti ti-arrow-narrow-right" aria-hidden="true"></i>${escapeHtml(a.outcome)}</div>` : ''}
        </div>
      `).join('')}
    </div>
  ` : `<div class="section-label">議事${editBtn('agenda', '議事')}</div>
       <p class="empty-section">未登録</p>`

  const doneCount = todos.filter((t) => t.done).length

  return header + `
    <div class="stat-grid">
      <div class="stat-card"><div class="stat-label">決定事項</div><div class="stat-value">${decisions.length}</div></div>
      <div class="stat-card"><div class="stat-label">ToDo</div><div class="stat-value">${doneCount}/${todos.length}</div></div>
      <div class="stat-card"><div class="stat-label">論点</div><div class="stat-value">${topics.length}</div></div>
    </div>
    <div class="section-label">サマリ${editBtn('cardSummary', 'サマリ')}</div>
    <p class="summary-text">${escapeHtml(s.cardSummary || '')}</p>
    ${agendaHtml}
    <div class="section-label">決定事項${editBtn('decisions', '決定事項')}</div>
    ${decisions.length ? `<ul class="plain-list">${decisions.map((d) => `<li>${escapeHtml(d)}</li>`).join('')}</ul>` : '<p class="empty-section">未登録</p>'}
    <div class="section-label">ToDo${editBtn('todos', 'ToDo')}</div>
    ${todos.length ? `<div class="todo-box">${todos.map((t, i) => `
      <label class="todo-row ${t.done ? 'done' : ''}">
        <input type="checkbox" class="todo-check" data-index="${i}" ${t.done ? 'checked' : ''} ${state.canEdit ? '' : 'disabled'} />
        <span>${escapeHtml(t.text)}</span>
      </label>
    `).join('')}</div>` : '<p class="empty-section">未登録</p>'}
    <div class="section-label">論点${editBtn('topics', '論点')}</div>
    ${topics.length ? `<ul class="plain-list">${topics.map((t) => `<li>${escapeHtml(t)}</li>`).join('')}</ul>` : '<p class="empty-section">未登録</p>'}
    <div class="actions-row">
      <button class="btn btn-regenerate"><i class="ti ti-refresh" aria-hidden="true"></i>要約を再生成</button>
      ${state.canEdit ? '<button class="btn btn-retranscribe"><i class="ti ti-microphone" aria-hidden="true"></i>文字起こし</button>' : ''}
      <button class="btn btn-raw"><i class="ti ti-file-text" aria-hidden="true"></i>原文表示</button>
      <span style="flex:1"></span>
      <span style="font-size:11px;color:var(--text-muted);align-self:center">${escapeHtml(s.model || '')}</span>
    </div>
  `
}

export function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]))
}
