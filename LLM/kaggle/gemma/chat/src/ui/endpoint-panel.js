import { h, clear } from '../lib/dom.js'
import { deriveState, formatHours, getStatusAll, startKernel, stopKernel } from '../lib/gas.js'

const POLL_MS = 30000
const OPEN_KEY = 'gemma-chat.endpointPanelOpen'

const DOT_COLOR = {
  ready: 'var(--fill-success)',
  booting: 'var(--text-accent)',
  stopping: 'var(--text-accent)',
  stopped: 'var(--text-muted)',
  unknown: 'var(--text-muted)',
}

/**
 * サイドバー上部のエンドポイント切替パネル。
 * 折りたたみ時は選択中の1件を1行で、展開時は全件を一覧で出す。
 * 30秒ごとに全件の状態をまとめて取得する（GAS 呼び出しは1回で済む）。
 *
 * @param {HTMLElement} root
 * @param {{ getSettings: () => object, onSelect: (id) => void }} ctx
 */
export function createEndpointPanel(root, ctx) {
  let open = localStorage.getItem(OPEN_KEY) === '1'
  let list = []
  let busy = null
  let timer = null

  function activeOf(items) {
    const s = ctx.getSettings()
    return items.find((e) => e.id === s.activeId) ?? items[0] ?? null
  }

  function row(e, { compact = false } = {}) {
    const s = ctx.getSettings()
    const state = deriveState(e)
    const isActive = e.id === s.activeId
    const working = busy === e.id

    const dot = h('span', {
      class: 'ep-dot',
      style: { background: DOT_COLOR[state.key] },
    })

    const remainClass = e.success && e.weeklyRemainMin < 120 ? 'ep-remain low' : 'ep-remain'
    const remain = h('span', {
      class: remainClass,
      text: e.success ? formatHours(e.weeklyRemainMin) : '—',
    })

    if (compact) {
      return h(
        'div',
        { class: 'ep-row ep-compact', onClick: () => togglePanel(true) },
        dot,
        h('span', { class: 'ep-name', text: e.label || e.id }),
        remain,
        h('i', { class: 'ti ti-chevron-down', 'aria-hidden': 'true' }),
      )
    }

    const actions = []
    if (working) {
      actions.push(h('span', { class: 'ep-hint', text: '送信中…' }))
    } else if (state.canStart) {
      actions.push(
        h('button', {
          class: 'ep-btn',
          text: '起動',
          onClick: (ev) => {
            ev.stopPropagation()
            act(e.id, startKernel)
          },
        }),
      )
    } else if (state.canStop) {
      actions.push(
        h('button', {
          class: 'ep-btn',
          text: '停止',
          onClick: (ev) => {
            ev.stopPropagation()
            act(e.id, stopKernel)
          },
        }),
      )
    }

    return h(
      'div',
      {
        class: isActive ? 'ep-row on' : 'ep-row',
        onClick: () => {
          if (!isActive) ctx.onSelect(e.id)
        },
      },
      dot,
      h(
        'div',
        { class: 'ep-mid' },
        h('div', { class: 'ep-name', text: e.label || e.id }),
        h('div', { class: 'ep-state', text: state.label }),
      ),
      remain,
      h('div', { class: 'ep-actions' }, ...actions),
    )
  }

  async function act(id, fn) {
    busy = id
    render()
    try {
      await fn(ctx.getSettings(), id)
    } catch (e) {
      alert(`操作に失敗しました: ${e.message}`)
    } finally {
      busy = null
      load()
    }
  }

  function togglePanel(force) {
    open = force ?? !open
    localStorage.setItem(OPEN_KEY, open ? '1' : '0')
    render()
  }

  function render() {
    clear(root)

    if (!ctx.getSettings().gasUrl) {
      root.append(h('div', { class: 'ep-panel', text: '起動制御は未設定' }))
      return
    }

    if (!list.length) {
      root.append(h('div', { class: 'ep-panel ep-loading', text: '読み込み中…' }))
      return
    }

    if (!open) {
      const a = activeOf(list)
      root.append(h('div', { class: 'ep-panel' }, a ? row(a, { compact: true }) : null))
      return
    }

    root.append(
      h(
        'div',
        { class: 'ep-panel ep-open' },
        h(
          'div',
          { class: 'ep-header', onClick: () => togglePanel(false) },
          h('span', { text: 'エンドポイント' }),
          h('i', { class: 'ti ti-chevron-up', 'aria-hidden': 'true' }),
        ),
        ...list.map((e) => row(e)),
      ),
    )
  }

  async function load() {
    clearTimeout(timer)
    try {
      const res = await getStatusAll(ctx.getSettings())
      list = res.endpoints ?? []
    } catch (e) {
      list = []
      root.replaceChildren(h('div', { class: 'ep-panel warn', text: e.message }))
    }
    render()
    timer = setTimeout(load, POLL_MS)
  }

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) clearTimeout(timer)
    else load()
  })

  render()
  load()

  return {
    refresh: load,
    rerender: render,
    /** 選択中エンドポイントが起動可能な状態（停止中）かどうか */
    activeIsStopped() {
      const a = activeOf(list)
      return a ? deriveState(a).key === 'stopped' : false
    },
  }
}
