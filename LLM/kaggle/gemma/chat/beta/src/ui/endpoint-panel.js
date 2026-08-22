import { h, clear } from '../lib/dom.js'
import { activeProfile } from '../lib/settings.js'
import { deriveState, formatElapsed, formatHours, getStatus, startKernel, stopKernel } from '../lib/gas.js'

const POLL_MS = 30000
const OPEN_KEY = 'gemma-chat.endpointPanelOpen'

const DOT_COLOR = {
  ready: 'var(--fill-success, #2e9e5b)',
  booting: 'var(--accent, #185fa5)',
  stopping: 'var(--accent, #185fa5)',
  stopped: 'var(--text-muted)',
  unknown: 'var(--text-muted)',
}

/**
 * サイドバー上部のエンドポイント切替パネル。
 * 各プロファイルは個別の GAS プロジェクトを持つため、
 * 状態取得は profile ごとに並列でリクエストする（GAS を1回にまとめられない）。
 *
 * @param {HTMLElement} root
 * @param {{ getSettings: () => object, onSelect: (id) => void }} ctx
 */
export function createEndpointPanel(root, ctx) {
  let open = localStorage.getItem(OPEN_KEY) === '1'
  // profileId -> 直近のステータス（{ profile, status, error }）
  let byId = new Map()
  let loaded = false
  let busy = null
  let timer = null

  function activeEntry() {
    const p = activeProfile(ctx.getSettings())
    return p ? byId.get(p.id) : null
  }

  function row(entry, { compact = false } = {}) {
    const s = ctx.getSettings()
    const { profile, status } = entry
    const state = deriveState(status)
    const isActive = profile.id === s.activeId
    const working = busy === profile.id

    const dot = h('span', { class: 'ep-dot', style: { background: DOT_COLOR[state.key] } })
    const remainOk = status?.success && typeof status.weeklyRemainMin === 'number'
    const remainClass = remainOk && status.weeklyRemainMin < 120 ? 'ep-remain low' : 'ep-remain'
    const remain = h('span', { class: remainClass, text: remainOk ? formatHours(status.weeklyRemainMin) : '—' })

    if (compact) {
      return h(
        'div',
        { class: 'ep-row ep-compact', onClick: () => togglePanel(true) },
        dot,
        h('span', { class: 'ep-name', text: profile.label || profile.id }),
        remain,
        h('i', { class: 'ti ti-chevron-down', 'aria-hidden': 'true' }),
      )
    }

    const actions = []
    if (working) {
      actions.push(h('span', { class: 'ep-hint', text: '送信中…' }))
    } else if (!profile.gasUrl) {
      actions.push(h('span', { class: 'ep-hint', text: '未設定' }))
    } else if (state.canStart) {
      actions.push(
        h('button', {
          class: 'ep-btn',
          text: '起動',
          onClick: (ev) => {
            ev.stopPropagation()
            act(profile, startKernel)
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
            act(profile, stopKernel)
          },
        }),
      )
    }

    return h(
      'div',
      {
        class: isActive ? 'ep-row on' : 'ep-row',
        onClick: () => {
          if (!isActive) ctx.onSelect(profile.id)
        },
      },
      dot,
      h(
        'div',
        { class: 'ep-mid' },
        h('div', { class: 'ep-name', text: profile.label || profile.id }),
        h('div', { class: 'ep-state', text: stateLine(status, state) }),
      ),
      remain,
      h('div', { class: 'ep-actions' }, ...actions),
    )
  }

  /** 状態ラベル。稼働中は今回の経過時間も併記する */
  function stateLine(status, state) {
    if (status?.success === false) return status.error
    const elapsed = formatElapsed(status?.currentSessionMin)
    return elapsed ? `${state.label} · ${elapsed}` : state.label
  }

  async function act(profile, fn) {
    busy = profile.id
    render()
    try {
      const res = await fn(profile)
      // 強制停止に切り替わった場合は黙って済ませず知らせる
      if (res && res.mode === 'forced') alert(res.message || '強制停止しました')
      else if (res && res.success === false) alert(res.message || '操作できませんでした')
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
    const profiles = ctx.getSettings().profiles

    if (!profiles.length) {
      root.append(h('div', { class: 'ep-panel', text: '接続先が未設定' }))
      return
    }

    if (!loaded) {
      root.append(h('div', { class: 'ep-panel ep-loading', text: '読み込み中…' }))
      return
    }

    if (!open) {
      const a = activeEntry()
      root.append(h('div', { class: 'ep-panel' }, a ? row(a, { compact: true }) : null))
      return
    }

    const totals = profiles.reduce((sum, p) => {
      const t = byId.get(p.id)?.status?.totalUsedMin
      return sum + (typeof t === 'number' ? t : 0)
    }, 0)

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
        ...profiles.map((p) => row(byId.get(p.id) ?? { profile: p, status: null })),
        totals > 0 &&
          h('div', { class: 'ep-total', text: `累計 ${formatHours(totals)}（GAS 経由の起動分）` }),
      ),
    )
  }

  async function load() {
    clearTimeout(timer)
    const profiles = ctx.getSettings().profiles

    const results = await Promise.allSettled(
      profiles.map((p) => (p.gasUrl ? getStatus(p) : Promise.reject(new Error('未設定')))),
    )

    const next = new Map()
    profiles.forEach((p, i) => {
      const r = results[i]
      if (r.status === 'fulfilled') {
        next.set(p.id, { profile: p, status: r.value })
      } else {
        next.set(p.id, { profile: p, status: { success: false, error: r.reason.message } })
      }
    })
    byId = next
    loaded = true
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
      const e = activeEntry()
      return e ? deriveState(e.status).key === 'stopped' : false
    },
  }
}
