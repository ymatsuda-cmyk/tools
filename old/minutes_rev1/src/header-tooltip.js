// src/header-tooltip.js — ヘッダー(.topbar)のアイコンボタンにマウスを乗せたとき、
// aria-label の内容をツールチップで表示する。
//
// .topbar / .topbar-scroll は overflow:hidden/auto なので、ボタンの子要素として
// 出すと切れてしまう。body 直下に position:fixed の要素を1つだけ置いて使い回す。
// タッチ操作では出さない(タップ後に残り続けるのを防ぐ)。キーボードのフォーカスでは出す。

const SHOW_DELAY_MS = 250
const MARGIN = 6

let tip = null
let timer = null
let current = null

function ensureTip() {
  if (tip) return tip
  tip = document.createElement('div')
  tip.className = 'hdr-tip'
  tip.setAttribute('role', 'tooltip')
  tip.id = 'hdr-tip'
  document.body.appendChild(tip)
  return tip
}

function place(btn) {
  const r = btn.getBoundingClientRect()
  const t = ensureTip()
  const w = t.offsetWidth
  const center = r.left + r.width / 2
  // 画面端ではみ出さないよう左右を寄せ、矢印だけはボタン中央を指す
  const left = Math.min(Math.max(MARGIN, center - w / 2), window.innerWidth - w - MARGIN)
  t.style.left = `${left}px`
  t.style.top = `${r.bottom + MARGIN + 2}px`
  t.style.setProperty('--arrow-x', `${center - left}px`)
}

function show(btn) {
  const label = btn.getAttribute('aria-label')
  if (!label || btn.offsetParent === null) return // 非表示ボタンには出さない
  const t = ensureTip()
  t.textContent = label
  current = btn
  btn.setAttribute('aria-describedby', 'hdr-tip')
  place(btn)
  t.classList.add('show')
}

function hide() {
  clearTimeout(timer)
  if (current) current.removeAttribute('aria-describedby')
  current = null
  if (tip) tip.classList.remove('show')
}

function buttonFrom(e) {
  const btn = e.target.closest?.('.topbar button[aria-label]')
  return btn && !btn.disabled ? btn : null
}

function init() {
  const bar = document.querySelector('.topbar')
  if (!bar) return

  bar.addEventListener('pointerover', (e) => {
    if (e.pointerType === 'touch') return
    const btn = buttonFrom(e)
    if (!btn || btn === current) return
    hide()
    timer = setTimeout(() => show(btn), SHOW_DELAY_MS)
  })
  bar.addEventListener('pointerout', (e) => {
    const btn = buttonFrom(e)
    if (btn && !btn.contains(e.relatedTarget)) hide()
  })
  bar.addEventListener('focusin', (e) => {
    const btn = buttonFrom(e)
    if (btn && btn.matches(':focus-visible')) show(btn)
  })
  bar.addEventListener('focusout', hide)
  bar.addEventListener('pointerdown', hide) // クリックしたら消す(モーダルが開くため)
  bar.querySelector('.topbar-scroll')?.addEventListener('scroll', hide, { passive: true })
  window.addEventListener('resize', hide)
  window.addEventListener('scroll', hide, { passive: true, capture: true })
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init)
} else {
  init()
}
