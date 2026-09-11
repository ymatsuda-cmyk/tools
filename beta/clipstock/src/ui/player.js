/**
 * 途中再生の小窓。
 *
 * 時刻付きのリンク([12:34])を踏むたびに別タブへ飛ぶと、読んでいた場所に戻るのが
 * 手間で「根拠を確かめる」という用途に合わない。画面の隅に小さく出して、
 * 本文を見ながら該当箇所だけ確認できるようにする。
 */
import { formatTimecode } from '../lib/timecode.js'

const HOST_ID = 'mini-player'

/** YouTubeの視聴URLのうち、開始時刻が付いているものだけを小窓の対象とみなす */
export function seekTargetOf(href) {
  const raw = String(href ?? '')
  const id = raw.match(/(?:v=|youtu\.be\/|shorts\/|embed\/)([\w-]{11})/)?.[1]
  if (!id) return null
  const t = raw.match(/[?&]t=(\d+)s?/)?.[1]
  if (t === undefined) return null
  return { id, at: Number(t) }
}

export function closeMiniPlayer() {
  document.getElementById(HOST_ID)?.remove()
}

// 閉じて開き直しても同じ場所に出す。null のうちは CSS の既定位置(右下)のまま
let placed = null

function clampAndPlace(host, left, top) {
  const w = host.offsetWidth
  const h = host.offsetHeight
  placed = {
    left: Math.min(Math.max(left, 0), Math.max(0, window.innerWidth - w)),
    top: Math.min(Math.max(top, 0), Math.max(0, window.innerHeight - h)),
  }
  host.style.left = `${placed.left}px`
  host.style.top = `${placed.top}px`
  host.style.right = 'auto'
  host.style.bottom = 'auto'
}

/** バーを掴んで動かす。再生中のiframeにポインタを取られないようキャプチャする */
function enableDrag(host) {
  const bar = host.querySelector('.mp-bar')
  let from = null

  bar.addEventListener('pointerdown', (e) => {
    if (e.target.closest('.mp-btn')) return
    const box = host.getBoundingClientRect()
    from = { x: e.clientX - box.left, y: e.clientY - box.top }
    bar.setPointerCapture(e.pointerId)
    host.classList.add('dragging')
    e.preventDefault()
  })

  bar.addEventListener('pointermove', (e) => {
    if (!from) return
    clampAndPlace(host, e.clientX - from.x, e.clientY - from.y)
  })

  const end = () => {
    from = null
    host.classList.remove('dragging')
  }
  bar.addEventListener('pointerup', end)
  bar.addEventListener('pointercancel', end)
}

export function openMiniPlayer(id, at) {
  let host = document.getElementById(HOST_ID)
  if (!host) {
    host = document.createElement('div')
    host.id = HOST_ID
    host.className = 'mini-player'
    host.innerHTML = `
      <div class="mp-bar">
        <i class="ti ti-grip-vertical mp-grip" aria-hidden="true"></i>
        <span class="mp-time"></span>
        <span class="grow"></span>
        <a class="mp-btn mp-open" target="_blank" rel="noopener" aria-label="YouTubeで開く"><i class="ti ti-external-link" aria-hidden="true"></i></a>
        <button class="mp-btn mp-close" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
      </div>
      <div class="mp-frame"></div>`
    document.body.appendChild(host)
    host.querySelector('.mp-close').addEventListener('click', closeMiniPlayer)
    enableDrag(host)
    if (placed) clampAndPlace(host, placed.left, placed.top)
  }

  host.querySelector('.mp-time').textContent = `${formatTimecode(at)} から再生`
  host.querySelector('.mp-open').href = `https://www.youtube.com/watch?v=${id}&t=${at}s`
  // src を差し替えるだけだと同じ動画のときに巻き直らないので、iframe ごと作り直す
  host.querySelector('.mp-frame').innerHTML = `
    <iframe
      src="https://www.youtube-nocookie.com/embed/${id}?start=${at}&autoplay=1&rel=0"
      title="動画の再生"
      allow="autoplay; encrypted-media; picture-in-picture"
      allowfullscreen
      referrerpolicy="strict-origin-when-cross-origin"></iframe>`
}
