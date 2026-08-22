import { h, clear } from '../lib/dom.js'
import { deriveState, formatQuota, getStatus, startKernel, stopKernel } from '../lib/gas.js'

const POLL_IDLE_MS = 20000
const POLL_TRANSITION_MS = 6000

/**
 * @param {HTMLElement} root
 * @param {{ getSettings: () => object, onStateChange?: (state) => void }} ctx
 */
export function createStatusBar(root, ctx) {
  let status = null
  let state = deriveState(null)
  let message = null
  let busy = false
  let timer = null

  function schedule() {
    clearTimeout(timer)
    const s = ctx.getSettings()
    if (!s.gasUrl || document.hidden) return
    const wait = state.key === 'booting' || state.key === 'stopping'
      ? POLL_TRANSITION_MS
      : POLL_IDLE_MS
    timer = setTimeout(poll, wait)
  }

  async function poll() {
    const s = ctx.getSettings()
    if (!s.gasUrl) {
      render()
      return
    }
    try {
      status = await getStatus(s)
      message = null
    } catch (e) {
      status = null
      message = e.message
    }
    state = deriveState(status)
    ctx.onStateChange?.(state)
    render()
    schedule()
  }

  async function act(fn, label) {
    busy = true
    message = null
    render()
    try {
      const res = await fn(ctx.getSettings())
      message = res?.message ?? null
      // 直後は状態が変わる途中なので短い間隔で追う
      state = { ...state, key: 'booting' }
    } catch (e) {
      message = `${label}に失敗: ${e.message}`
    } finally {
      busy = false
      render()
      clearTimeout(timer)
      timer = setTimeout(poll, 2000)
    }
  }

  function render() {
    const s = ctx.getSettings()
    clear(root)

    if (!s.gasUrl) {
      root.append(h('span', { class: 'hint', text: '起動制御は未設定' }))
      return
    }

    root.append(
      h('span', { class: `chip ${state.tone}` }, h('span', { class: 'dot' }), state.label),
    )

    const quota = status?.gpuQuota

    if (quota) {

      const blocks =
        Math.round(quota.remainPct / 10)

      const bar =
        '■'.repeat(blocks) +
        '□'.repeat(10 - blocks)

      root.append(
        h(
          'span',
          {
            class: 'hint quota-display'
          },
          `GPU ${bar} ${quota.remainPct}% `,
          `${quota.remainHour.toFixed(1)}h/${quota.totalHour.toFixed(0)}h`
        )
      )
    }

    if (busy) {
      root.append(h('span', { class: 'hint', text: '送信中…' }))
    } else if (state.canStart) {
      root.append(h('button', { text: '起動', onClick: () => act(startKernel, '起動') }))
    } else if (state.canStop) {
      root.append(h('button', { text: '停止', onClick: () => act(stopKernel, '停止') }))
    }

    if (message) root.append(h('span', { class: 'hint msg', text: message }))
  }

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) clearTimeout(timer)
    else poll()
  })

  render()

  return {
    refresh: poll,
    isReady: () => state.key === 'ready',
    getState: () => state,
  }
}
