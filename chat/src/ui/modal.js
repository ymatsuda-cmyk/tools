import { h, clear } from '../lib/dom.js'

const root = () => document.getElementById('modal-root')

/**
 * build({ close }) が返した要素をオーバーレイに載せる。
 * 背景クリックと Esc で閉じる。
 */
export function openModal(build) {
  const host = root()
  const close = () => {
    clear(host)
    document.removeEventListener('keydown', onKey)
  }
  const onKey = (e) => {
    if (e.key === 'Escape') close()
  }
  document.addEventListener('keydown', onKey)

  const inner = build({ close })
  inner.addEventListener('click', (e) => e.stopPropagation())

  clear(host)
  host.append(h('div', { class: 'modal-bg', onClick: close }, inner))
  return { close }
}
