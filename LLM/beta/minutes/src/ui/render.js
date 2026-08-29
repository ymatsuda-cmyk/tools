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
export function renderList(container, items, selectedKey, onSelect) {
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

    const row = document.createElement('div')
    row.className = 'list-item' + (item.key === selectedKey ? ' selected' : '')
    row.dataset.key = item.key
    row.innerHTML = `
      <div class="list-item-time">${fmtTime(item.date)}</div>
      <div class="list-item-body">
        <div class="list-item-title">${escapeHtml(item.title)}</div>
        <div class="list-item-meta">
          <span class="badge status-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
          ${escapeHtml(item.duration || '')}
        </div>
      </div>
    `
    row.addEventListener('click', () => onSelect(item, row))
    container.appendChild(row)
  }

  if (!sorted.length) {
    container.innerHTML = '<div class="empty-state"><i class="ti ti-inbox" aria-hidden="true"></i><p>議事録がありません</p></div>'
  }
}

/**
 * 詳細ペインの中身を組み立てて返す(HTML文字列)。
 * state.phase: 'loading' | 'no-summary' | 'ready' | 'generating' | 'error'
 */
export function renderDetailHtml(item, state) {
  const header = `
    <div class="detail-header">
      <span class="detail-title">${escapeHtml(item.title)}</span>
      <span class="badge status-${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>
    </div>
    <div class="detail-meta">${fmtDate(item.date)} ${fmtTime(item.date)} · ${escapeHtml(item.duration || '')}</div>
  `

  if (state.phase === 'loading') {
    return header + `<p class="summary-text">読み込み中...</p>`
  }

  if (state.phase === 'error') {
    return header + `<p class="error-text">${escapeHtml(state.message)}</p>
      <div class="actions-row"><button class="btn btn-retry"><i class="ti ti-refresh" aria-hidden="true"></i>再試行</button></div>`
  }

  if (state.phase === 'generating') {
    return header + `<p class="summary-text">${escapeHtml(state.progress || '要約を生成しています...')}</p>`
  }

  if (state.phase === 'no-summary') {
    return header + `
      <p class="summary-text" style="color:var(--text-muted)">この議事録の要約はまだありません。</p>
      <div class="actions-row"><button class="btn btn-generate"><i class="ti ti-sparkles" aria-hidden="true"></i>要約を生成</button></div>
    `
  }

  // ready
  const s = state.summary
  const decisions = s.detail?.decisions || []
  const todos = s.detail?.todos || []
  const topics = s.detail?.topics || []

  return header + `
    <div class="stat-grid">
      <div class="stat-card"><div class="stat-label">決定事項</div><div class="stat-value">${decisions.length}</div></div>
      <div class="stat-card"><div class="stat-label">ToDo</div><div class="stat-value">${todos.length}</div></div>
      <div class="stat-card"><div class="stat-label">論点</div><div class="stat-value">${topics.length}</div></div>
    </div>
    <div class="section-label">サマリ</div>
    <p class="summary-text">${escapeHtml(s.cardSummary || '')}</p>
    ${decisions.length ? `<div class="section-label">決定事項</div><ul class="plain-list">${decisions.map((d) => `<li>${escapeHtml(d)}</li>`).join('')}</ul>` : ''}
    ${todos.length ? `<div class="section-label">ToDo</div><div class="todo-box">${todos.map((t) => `<div class="todo-row"><i class="ti ti-square" aria-hidden="true"></i><span>${escapeHtml(t)}</span></div>`).join('')}</div>` : ''}
    ${topics.length ? `<div class="section-label">論点</div><ul class="plain-list">${topics.map((t) => `<li>${escapeHtml(t)}</li>`).join('')}</ul>` : ''}
    <div class="actions-row">
      <button class="btn btn-regenerate"><i class="ti ti-refresh" aria-hidden="true"></i>要約を再生成</button>
      <span style="flex:1"></span>
      <span style="font-size:11px;color:var(--text-muted);align-self:center">${escapeHtml(s.model || '')}</span>
    </div>
  `
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]))
}
