// src/scrollhide.js — スマホ・タブレットで本文を読むとき、ヘッダーを自動でしまう
// 下へスクロールすると、アプリのヘッダーと会議のタイトル部分が上へ消え、少し上へ戻すと出てくる。
export function initScrollHide(app, mq, now = () => performance.now()) {
  let last = 0, el = null, lockUntil = 0
  const set = (on) => {
    if (app.classList.contains('collapsed') === on) return
    if (on) {
      const top = app.querySelector('.top')
      if (top) app.style.setProperty('--top-h', top.offsetHeight + 'px')
    }
    app.classList.toggle('collapsed', on)
    lockUntil = now() + 380  // 表示が詰まる・広がる間の見かけのスクロールは、判定に使わない
  }
  document.addEventListener('scroll', (e) => {
    const t = e.target
    if (!(t instanceof Element) || !t.classList.contains('sheet-scroll')) return
    if (!mq.matches) { set(false); return }
    const y = t.scrollTop
    if (t !== el) { el = t; last = y }
    if (now() < lockUntil) { last = y; return }
    if (y < 24) { set(false); last = y; return }
    const dy = y - last
    if (dy > 14) { set(true); last = y } else if (dy < -14) { set(false); last = y }
  }, true)
  return { reset() { set(false); el = null; last = 0 } }
}
