/**
 * 途中再生の小窓。
 *
 * 時刻付きのリンク([12:34])を踏むたびに別タブへ飛ぶと、読んでいた場所に戻るのが
 * 手間で「根拠を確かめる」という用途に合わない。画面の隅に小さく出して、
 * 本文を見ながら該当箇所だけ確認できるようにする。
 *
 * Drive のプレビューは外から seek できないため、時刻が変わるたびに iframe ごと
 * 作り直す。YouTube も同じ扱いにして、再生元による違いを呼び出し側に見せない。
 */
import { formatTimecode, embedUrlAt, openUrlAt } from '../lib/timecode.js'

const HOST_ID = 'mini-player'

export function closeMiniPlayer() {
  document.getElementById(HOST_ID)?.remove()
}

// 閉じて開き直しても同じ場所・同じ大きさで出す。null のうちは CSS の既定のまま
let placed = null
let sized = null

const MIN_WIDTH = 220
const MAX_WIDTH = 900

function clampAndPlace(host, left, top) {
  placed = {
    left: Math.min(Math.max(left, 0), Math.max(0, window.innerWidth - host.offsetWidth)),
    top: Math.min(Math.max(top, 0), Math.max(0, window.innerHeight - host.offsetHeight)),
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

/** 右下の角で横幅を変える。高さは 16:9 のまま追従する */
function enableResize(host) {
  const grip = host.querySelector('.mp-resize')
  let from = null

  grip.addEventListener('pointerdown', (e) => {
    const box = host.getBoundingClientRect()
    // 右下基準のままだと掴んだ角と反対側が動いてしまうので、先に左上基準に直す
    if (!placed) clampAndPlace(host, box.left, box.top)
    from = { x: e.clientX, width: box.width }
    grip.setPointerCapture(e.pointerId)
    host.classList.add('dragging')
    e.preventDefault()
  })

  grip.addEventListener('pointermove', (e) => {
    if (!from) return
    const max = Math.min(MAX_WIDTH, window.innerWidth - (placed?.left ?? 0) - 8)
    sized = Math.min(Math.max(from.width + (e.clientX - from.x), MIN_WIDTH), Math.max(MIN_WIDTH, max))
    host.style.width = `${sized}px`
  })

  const end = () => {
    from = null
    host.classList.remove('dragging')
  }
  grip.addEventListener('pointerup', end)
  grip.addEventListener('pointercancel', end)
}

/**
 * @param {string} url Driveリンク or YouTubeのURL
 * @param {number} at 開始秒
 * @param {string} title 小窓のバーに出す名前
 */
export function openMiniPlayer(url, at = 0, title = '') {
  const src = embedUrlAt(url, at)
  if (!src) return

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
        <a class="mp-btn mp-open" target="_blank" rel="noopener" aria-label="別タブで開く"><i class="ti ti-external-link" aria-hidden="true"></i></a>
        <button class="mp-btn mp-close" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
      </div>
      <div class="mp-frame"></div>
      <div class="mp-resize" aria-hidden="true"></div>`
    document.body.appendChild(host)
    host.querySelector('.mp-close').addEventListener('click', closeMiniPlayer)
    enableDrag(host)
    enableResize(host)
    if (sized) host.style.width = `${sized}px`
    if (placed) clampAndPlace(host, placed.left, placed.top)
  }

  host.querySelector('.mp-time').textContent = at ? `${formatTimecode(at)} から再生` : title || '再生'
  host.querySelector('.mp-time').title = title || ''
  host.querySelector('.mp-open').href = openUrlAt(url, at)
  // src を差し替えるだけだと同じ動画のときに巻き直らないので、iframe ごと作り直す
  host.querySelector('.mp-frame').innerHTML = `
    <iframe
      src="${src}"
      title="動画の再生"
      allow="autoplay; encrypted-media; picture-in-picture"
      allowfullscreen
      referrerpolicy="strict-origin-when-cross-origin"></iframe>`
}
