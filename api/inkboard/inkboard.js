/*!
 * inkboard.js — 手書きボード（Apple Pencil／指／マウス）の共通API
 *
 *   import { createBoard, renderBoard, openBoard } from '../api/inkboard/inkboard.js'
 *
 *   const value = createBoard({ paper: 'A4', orientation: 'portrait' })   // 用紙固定
 *   const value = createBoard({ paper: 'infinite' })                      // 無限キャンバス
 *
 *   // 参照表示。クリック（タップ）で全画面の編集画面が開く
 *   renderBoard(el, value, { onChange: (v) => save(v) })
 *
 *   // 全画面の編集画面だけを開く。閉じたときの値で解決する
 *   const next = await openBoard(value, { onChange: (v) => save(v) })
 *
 * 扱う値は改行を含まない1行のJSON文字列。線の座標は差分を短い文字列に詰めてあるので、
 * Markdownのコードブロックや1つのセルにそのまま置ける。
 * 参照表示で見せる範囲（view）も値の中に持つ。編集画面の「範囲」ツールで指定する。
 *
 * 表示だけなら options は不要。onChange を渡したときだけ描ける（渡さなければ閲覧のみ）。
 */

// ---------------------------------------------------------------- 定数

/** 用紙の大きさ（96dpi の CSS ピクセル。縦向きの値） */
export const PAPERS = {
  A3: { label: 'A3', w: 1123, h: 1587 },
  A4: { label: 'A4', w: 794, h: 1123 },
  A5: { label: 'A5', w: 559, h: 794 },
  B5: { label: 'B5', w: 688, h: 971 },
  W169: { label: '16:9', w: 720, h: 1280 },
  W43: { label: '4:3', w: 768, h: 1024 },
}

export const BACKGROUNDS = { plain: '無地', grid: '方眼', lines: '罫線', dots: 'ドット' }

const COLORS = ['#1f1f1f', '#2f5fd0', '#d33a2c', '#2e8b57', '#f2b100', '#8a4fd6']
const WIDTHS = { pen: [1.5, 3, 6], marker: [12, 20, 32], eraser: [6, 14, 30] }
const TOOL_LABEL = { pen: 'ペン', marker: 'マーカー', eraser: '消しゴム', area: '参照表示の範囲', hand: '移動' }
const DEFAULT_VIEW = { x: 0, y: 0, w: 800, h: 450 }
const Q = 4 // 座標は 1/4 ピクセル単位で保存
const MIN_SCALE = 0.1
const MAX_SCALE = 8

const clamp = (v, a, b) => Math.min(b, Math.max(a, v))

// ---------------------------------------------------------------- 値 <-> データ

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_'
const B64I = Object.fromEntries([...B64].map((c, i) => [c, i]))

/** 整数列を「ジグザグ符号化 + 1文字5ビットの可変長」で文字列にする */
function encInts(list) {
  let s = ''
  for (const v of list) {
    let n = v < 0 ? -v * 2 - 1 : v * 2
    do {
      let c = n % 32
      n = Math.floor(n / 32)
      if (n) c += 32
      s += B64[c]
    } while (n)
  }
  return s
}
function decInts(s) {
  const out = []
  let n = 0
  let mul = 1
  for (const ch of s) {
    const c = B64I[ch]
    if (c === undefined) throw new Error('線のデータが壊れています')
    n += (c & 31) * mul
    if (c & 32) mul *= 32
    else {
      out.push(n % 2 ? -(n + 1) / 2 : n / 2)
      n = 0
      mul = 1
    }
  }
  return out
}

/** 点列 [x, y, 筆圧, x, y, 筆圧, ...] を文字列にする（各成分は前の点との差分） */
function encPoints(pts) {
  const ints = []
  let px = 0, py = 0, pp = 0
  for (let i = 0; i < pts.length; i += 3) {
    const x = Math.round(pts[i] * Q), y = Math.round(pts[i + 1] * Q), p = Math.round(pts[i + 2] * 63)
    ints.push(x - px, y - py, p - pp)
    px = x; py = y; pp = p
  }
  return encInts(ints)
}
function decPoints(s) {
  const ints = decInts(s)
  const pts = new Array(ints.length - (ints.length % 3))
  let x = 0, y = 0, p = 0
  for (let i = 0; i + 2 < ints.length; i += 3) {
    x += ints[i]; y += ints[i + 1]; p += ints[i + 2]
    pts[i] = x / Q; pts[i + 1] = y / Q; pts[i + 2] = p / 63
  }
  return pts
}

function strokeBox(s) {
  const p = s.pts
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (let i = 0; i < p.length; i += 3) {
    if (p[i] < x0) x0 = p[i]
    if (p[i] > x1) x1 = p[i]
    if (p[i + 1] < y0) y0 = p[i + 1]
    if (p[i + 1] > y1) y1 = p[i + 1]
  }
  const pad = s.width
  s.box = { x0: x0 - pad, y0: y0 - pad, x1: x1 + pad, y1: y1 + pad }
  return s
}

function paperSize(key, orient) {
  const p = PAPERS[key] || PAPERS.A4
  return orient === 'landscape' ? { w: p.h, h: p.w } : { w: p.w, h: p.h }
}

function normalizeView(v) {
  if (!Array.isArray(v) || v.length !== 4 || !(v[2] > 0) || !(v[3] > 0)) return null
  return { x: +v[0], y: +v[1], w: +v[2], h: +v[3] }
}

/**
 * 値を作る。
 * @param {{ paper?: 'infinite'|'A3'|'A4'|'A5'|'B5'|'W169'|'W43', orientation?: 'portrait'|'landscape', background?: 'plain'|'grid'|'lines'|'dots' }} [opts]
 */
export function createBoard(opts = {}) {
  const infinite = opts.paper === 'infinite'
  const paper = infinite ? 'A4' : (PAPERS[opts.paper] ? opts.paper : 'A4')
  const orient = opts.orientation === 'landscape' || (!opts.orientation && /^W/.test(paper)) ? 'landscape' : 'portrait'
  return serializeBoard({
    mode: infinite ? 'infinite' : 'paper',
    paper,
    orient,
    bg: BACKGROUNDS[opts.background] ? opts.background : (infinite ? 'grid' : 'plain'),
    view: null,
    strokes: [],
  })
}

/** 値をデータに戻す。空の値は A4 縦の白紙として扱う */
export function parseBoard(value) {
  const raw = String(value ?? '').trim()
  const j = JSON.parse(raw || createBoard())
  if (!j || j.v !== 1) throw new Error('手書きボードの形式ではありません')
  const mode = j.mode === 'infinite' ? 'infinite' : 'paper'
  const paper = PAPERS[j.paper] ? j.paper : 'A4'
  const orient = j.orient === 'landscape' ? 'landscape' : 'portrait'
  const strokes = (Array.isArray(j.s) ? j.s : []).map(([tool, color, width, enc]) =>
    strokeBox({ tool: tool === 'marker' ? 'marker' : 'pen', color: String(color), width: +width || 2, pts: decPoints(String(enc)), enc: String(enc) })
  ).filter((s) => s.pts.length >= 3)
  return { mode, paper, orient, bg: BACKGROUNDS[j.bg] ? j.bg : 'plain', view: normalizeView(j.view), strokes }
}

/** データを値にする */
export function serializeBoard(d) {
  const r = (n) => Math.round(n * 10) / 10
  return JSON.stringify({
    v: 1,
    mode: d.mode,
    paper: d.paper,
    orient: d.orient,
    bg: d.bg,
    view: d.view ? [r(d.view.x), r(d.view.y), r(d.view.w), r(d.view.h)] : null,
    s: d.strokes.map((s) => [s.tool, s.color, s.width, s.enc || (s.enc = encPoints(s.pts))]),
  })
}

/** 描いた線全体の範囲（線が無ければ null） */
export function contentBounds(d) {
  if (!d.strokes.length) return null
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (const s of d.strokes) {
    x0 = Math.min(x0, s.box.x0); y0 = Math.min(y0, s.box.y0)
    x1 = Math.max(x1, s.box.x1); y1 = Math.max(y1, s.box.y1)
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

/**
 * 参照表示で見せる範囲。指定があればそれ、無ければ用紙全体（無限キャンバスは描いた線の全体）。
 * 受け取るのは値でもデータでもよい。呼び出し側で枠の縦横比を合わせたいときに使う。
 */
export function previewArea(valueOrData) {
  const d = typeof valueOrData === 'string' || valueOrData == null ? parseBoard(valueOrData) : valueOrData
  if (d.view) return { ...d.view }
  if (d.mode === 'paper') return { x: 0, y: 0, ...paperSize(d.paper, d.orient) }
  const b = contentBounds(d)
  if (!b) return { ...DEFAULT_VIEW }
  const pad = 24
  return { x: b.x - pad, y: b.y - pad, w: Math.max(b.w + pad * 2, 160), h: Math.max(b.h + pad * 2, 90) }
}

// ---------------------------------------------------------------- 描画

function pressureWidth(s, p) {
  return s.width * (0.3 + 0.9 * p)
}

/** 1本の線を描く。ペンは筆圧で太さを変え、マーカーは一定の太さで下の線を透かす */
function drawStroke(ctx, s) {
  const p = s.pts
  const n = p.length / 3
  ctx.strokeStyle = s.color
  ctx.fillStyle = s.color
  ctx.lineCap = 'round'
  ctx.lineJoin = 'round'
  if (s.tool === 'marker') {
    ctx.save()
    ctx.globalAlpha = 0.45
    ctx.globalCompositeOperation = 'multiply'
    ctx.lineWidth = s.width
    ctx.beginPath()
    ctx.moveTo(p[0], p[1])
    if (n === 1) ctx.lineTo(p[0] + 0.01, p[1])
    for (let i = 1; i < n - 1; i++) {
      ctx.quadraticCurveTo(p[i * 3], p[i * 3 + 1], (p[i * 3] + p[i * 3 + 3]) / 2, (p[i * 3 + 1] + p[i * 3 + 4]) / 2)
    }
    if (n > 1) ctx.lineTo(p[n * 3 - 3], p[n * 3 - 2])
    ctx.stroke()
    ctx.restore()
    return
  }
  if (n === 1) {
    ctx.beginPath()
    ctx.arc(p[0], p[1], pressureWidth(s, p[2]) / 2, 0, Math.PI * 2)
    ctx.fill()
    return
  }
  let sx = p[0], sy = p[1]
  for (let i = 1; i < n; i++) {
    const last = i === n - 1
    const ex = last ? p[i * 3] : (p[i * 3] + p[i * 3 + 3]) / 2
    const ey = last ? p[i * 3 + 1] : (p[i * 3 + 1] + p[i * 3 + 4]) / 2
    ctx.lineWidth = pressureWidth(s, (p[i * 3 - 1] + p[i * 3 + 2]) / 2)
    ctx.beginPath()
    ctx.moveTo(sx, sy)
    if (last) ctx.lineTo(ex, ey)
    else ctx.quadraticCurveTo(p[i * 3], p[i * 3 + 1], ex, ey)
    ctx.stroke()
    sx = ex; sy = ey
  }
}

function cssVar(el, name, fallback) {
  const v = getComputedStyle(el).getPropertyValue(name).trim()
  return v || fallback
}

/**
 * 世界座標の rect（見えている範囲）に背景と線を描く。
 * ctx は既に「世界座標 -> 画面」の変換が掛かっている前提。
 */
function drawScene(ctx, d, rect, scale, colors, opts = {}) {
  const paper = d.mode === 'paper' ? { x: 0, y: 0, ...paperSize(d.paper, d.orient) } : null
  if (paper && !opts.noDesk) {
    ctx.fillStyle = colors.desk
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
    ctx.save()
    ctx.shadowColor = 'rgba(0,0,0,.18)'
    ctx.shadowBlur = 12 * scale
    ctx.fillStyle = colors.paper
    ctx.fillRect(paper.x, paper.y, paper.w, paper.h)
    ctx.restore()
  } else {
    ctx.fillStyle = colors.paper
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h)
  }
  ctx.save()
  const area = paper ? intersect(rect, paper) : rect
  if (paper) {
    ctx.beginPath()
    ctx.rect(paper.x, paper.y, paper.w, paper.h)
    ctx.clip()
  }
  if (area) drawPattern(ctx, d.bg, area, scale, colors.rule)
  for (const s of d.strokes) {
    const b = s.box
    if (b.x1 < rect.x || b.y1 < rect.y || b.x0 > rect.x + rect.w || b.y0 > rect.y + rect.h) continue
    drawStroke(ctx, s)
  }
  ctx.restore()
}

function intersect(a, b) {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y)
  const w = Math.min(a.x + a.w, b.x + b.w) - x, h = Math.min(a.y + a.h, b.y + b.h) - y
  return w > 0 && h > 0 ? { x, y, w, h } : null
}

function drawPattern(ctx, bg, r, scale, color) {
  if (bg === 'plain') return
  const step = bg === 'lines' ? 32 : 24
  if (step * scale < 5) return
  const x0 = Math.floor(r.x / step) * step, y0 = Math.floor(r.y / step) * step
  const x1 = r.x + r.w, y1 = r.y + r.h
  ctx.save()
  ctx.strokeStyle = color
  ctx.fillStyle = color
  ctx.lineWidth = 1 / scale
  if (bg === 'dots') {
    const rad = Math.max(1.1 / scale, 0.8)
    ctx.beginPath()
    for (let x = x0; x <= x1; x += step) for (let y = y0; y <= y1; y += step) {
      ctx.moveTo(x + rad, y)
      ctx.arc(x, y, rad, 0, Math.PI * 2)
    }
    ctx.fill()
  } else {
    ctx.beginPath()
    for (let y = y0; y <= y1; y += step) { ctx.moveTo(r.x, y); ctx.lineTo(x1, y) }
    if (bg === 'grid') for (let x = x0; x <= x1; x += step) { ctx.moveTo(x, r.y); ctx.lineTo(x, y1) }
    ctx.stroke()
  }
  ctx.restore()
}

function sceneColors(el) {
  return {
    paper: cssVar(el, '--ib-paper', '#ffffff'),
    desk: cssVar(el, '--ib-desk', '#e9e7e1'),
    rule: cssVar(el, '--ib-rule', 'rgba(60,90,160,.18)'),
  }
}

/**
 * 値の一部を画像にする。
 * @param {string} value
 * @param {{ area?: {x:number,y:number,w:number,h:number}, scale?: number, type?: string }} [opts]
 * @returns {Promise<Blob>}
 */
export function exportImage(value, opts = {}) {
  const d = parseBoard(value)
  const area = opts.area || previewArea(d)
  const scale = opts.scale || 2
  const c = document.createElement('canvas')
  c.width = Math.max(1, Math.round(area.w * scale))
  c.height = Math.max(1, Math.round(area.h * scale))
  const ctx = c.getContext('2d')
  ctx.setTransform(scale, 0, 0, scale, -area.x * scale, -area.y * scale)
  drawScene(ctx, d, area, scale, sceneColors(document.documentElement), { noDesk: true })
  return new Promise((res, rej) => c.toBlob((b) => (b ? res(b) : rej(new Error('画像を作れませんでした'))), opts.type || 'image/png'))
}

// ---------------------------------------------------------------- 参照表示

/**
 * container の中に、参照表示の範囲を縦横比を保って描く。
 * クリック（タップ）で全画面の編集画面を開き、変更は onChange と表示の両方に反映する。
 *
 * @param {HTMLElement} container
 * @param {string} value
 * @param {{
 *   onChange?: (value: string) => void,   // 渡したときだけ編集できる
 *   onOpen?: () => void,
 *   onClose?: (value: string) => void,
 *   clickToOpen?: boolean,                // 既定 true。false にすると controller.open() でだけ開く
 *   emptyText?: string,
 *   title?: string,                       // 編集画面の見出し
 *   changeDelay?: number,                 // 書いてから onChange を呼ぶまでの待ち（ms、既定 800）
 * }} [options]
 * @returns {{ value: string, update(value: string): void, open(): Promise<string>, destroy(): void }}
 */
export function renderBoard(container, value, options = {}) {
  container.innerHTML = ''
  container.classList.add('ib-host')
  let current = String(value ?? '')
  let data
  const canvas = document.createElement('canvas')
  canvas.className = 'ib-preview'
  const empty = document.createElement('div')
  empty.className = 'ib-empty'
  container.append(canvas, empty)

  function parse() {
    try {
      data = parseBoard(current)
      container.classList.remove('ib-broken')
    } catch (err) {
      data = null
      container.classList.add('ib-broken')
      empty.textContent = `手書きボードを読み込めません（${err.message || err}）`
    }
  }

  function draw() {
    const w = container.clientWidth, h = container.clientHeight
    if (!w || !h) return
    const dpr = window.devicePixelRatio || 1
    canvas.width = Math.round(w * dpr)
    canvas.height = Math.round(h * dpr)
    const ctx = canvas.getContext('2d')
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, canvas.width, canvas.height)
    if (!data) return
    empty.textContent = data.strokes.length ? '' : (options.emptyText ?? (options.onChange ? 'タップして手書き' : ''))
    const a = previewArea(data)
    const scale = Math.min(w / a.w, h / a.h)
    const ox = (w - a.w * scale) / 2, oy = (h - a.h * scale) / 2
    ctx.setTransform(dpr * scale, 0, 0, dpr * scale, dpr * (ox - a.x * scale), dpr * (oy - a.y * scale))
    ctx.beginPath()
    ctx.rect(a.x, a.y, a.w, a.h)
    ctx.clip()
    drawScene(ctx, data, a, scale, sceneColors(container), { noDesk: true })
  }

  parse()
  let raf = 0
  const ro = new ResizeObserver(() => { cancelAnimationFrame(raf); raf = requestAnimationFrame(draw) })
  ro.observe(container)
  draw()

  let opened = null
  const ctl = {
    get value() { return current },
    update(v) {
      if (String(v ?? '') === current) return
      current = String(v ?? '')
      parse()
      draw()
    },
    open() {
      if (opened) return opened
      options.onOpen?.()
      opened = openBoard(current, {
        title: options.title,
        changeDelay: options.changeDelay,
        readOnly: typeof options.onChange !== 'function',
        onChange: (v) => {
          current = v
          parse()
          draw()
          options.onChange?.(v)
        },
      }).then((v) => {
        opened = null
        options.onClose?.(v)
        return v
      })
      return opened
    },
    destroy() {
      ro.disconnect()
      container.innerHTML = ''
      container.classList.remove('ib-host', 'ib-clickable', 'ib-broken')
    },
  }
  if (options.clickToOpen !== false) {
    container.classList.add('ib-clickable')
    container.title = options.onChange ? 'タップで全画面表示して手書き' : 'タップで全画面表示'
    container.addEventListener('click', () => { if (data) ctl.open() })
  }
  return ctl
}

// ---------------------------------------------------------------- 全画面の編集画面

const ICON = {
  pen: '<path d="M4 20l4-1 11-11-3-3L5 16z"/><path d="M14 6l3 3"/>',
  marker: '<path d="M9 15l-3 3v2h5l2-2"/><path d="M9 15l8-8 3 3-8 8z"/><path d="M4 21h16" opacity=".5"/>',
  eraser: '<path d="M8 20h12"/><path d="M4.5 15.5l9-9a2 2 0 0 1 2.8 0l3.2 3.2a2 2 0 0 1 0 2.8L13 19H8z"/><path d="M9 11l5 5"/>',
  area: '<path d="M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4"/><rect x="8" y="8" width="8" height="8" rx="1" opacity=".5"/>',
  hand: '<path d="M8 13V5.5a1.5 1.5 0 0 1 3 0V12M11 11.5V4a1.5 1.5 0 0 1 3 0v8M14 11.5V5.5a1.5 1.5 0 0 1 3 0V14c0 4-2.5 7-6 7-2.5 0-4-1.2-5.5-3.5L4 15a1.5 1.5 0 0 1 2.5-1.6L8 15"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
  minus: '<path d="M5 12h14"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
}
const svg = (k) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`

const PREF_KEY = 'inkboard.prefs'
function loadPrefs() {
  try { return JSON.parse(localStorage.getItem(PREF_KEY)) || {} } catch { return {} }
}
function savePrefs(p) {
  try { localStorage.setItem(PREF_KEY, JSON.stringify(p)) } catch { /* 保存できなくても動かす */ }
}

let openCount = 0

/**
 * 全画面の編集画面を開く。閉じると、その時点の値で解決する。
 * @param {string} value
 * @param {{ onChange?: (value: string) => void, readOnly?: boolean, title?: string, changeDelay?: number }} [options]
 * @returns {Promise<string>}
 */
export function openBoard(value, options = {}) {
  return new Promise((resolve) => {
    const ed = new Editor(value, options, resolve)
    ed.mount()
  })
}

class Editor {
  constructor(value, options, resolve) {
    this.options = options
    this.resolve = resolve
    this.readOnly = !!options.readOnly || typeof options.onChange !== 'function'
    this.data = parseBoard(value)
    this.lastValue = serializeBoard(this.data)
    this.prefs = loadPrefs()
    this.tool = 'pen'
    this.colors = { pen: this.prefs.penColor || COLORS[0], marker: this.prefs.markerColor || COLORS[4] }
    this.widthIdx = { pen: this.prefs.penWidth ?? 1, marker: this.prefs.markerWidth ?? 1, eraser: this.prefs.eraserWidth ?? 1 }
    // 指で描くか。Apple Pencil を一度でも使ったら、指は移動と拡大縮小に回す
    this.fingerDraw = this.prefs.fingerDraw ?? !this.prefs.penSeen
    this.undoStack = []
    this.redoStack = []
    this.cam = { x: 0, y: 0, scale: 1 }
    this.pointers = new Map()
    this.action = null
    this.space = false
    this.dirty = false
  }

  // ---- 画面を組み立てる
  mount() {
    const root = (this.root = document.createElement('div'))
    root.className = 'ib-overlay'
    root.setAttribute('role', 'dialog')
    root.setAttribute('aria-label', this.options.title || '手書きボード')
    const ro = this.readOnly
    root.innerHTML = `
      <div class="ib-bar">
        <button type="button" class="ib-done" data-op="close">${ro ? '閉じる' : '完了'}</button>
        <span class="ib-title"></span>
        ${ro ? '' : `
        <div class="ib-group ib-tools">
          ${['pen', 'marker', 'eraser', 'area', 'hand'].map((t) => `<button type="button" class="ib-ic" data-tool="${t}" title="${TOOL_LABEL[t]}" aria-label="${TOOL_LABEL[t]}">${svg(t)}</button>`).join('')}
        </div>
        <div class="ib-group ib-colors">
          ${COLORS.map((c) => `<button type="button" class="ib-sw" data-color="${c}" style="--c:${c}" aria-label="色 ${c}"></button>`).join('')}
        </div>
        <div class="ib-group ib-widths">
          ${[0, 1, 2].map((i) => `<button type="button" class="ib-ic ib-w" data-width="${i}" aria-label="太さ ${i + 1}"><i style="--d:${[4, 7, 11][i]}px"></i></button>`).join('')}
        </div>
        <div class="ib-group">
          <button type="button" class="ib-ic" data-op="undo" title="元に戻す（⌘Z）" aria-label="元に戻す">${svg('undo')}</button>
          <button type="button" class="ib-ic" data-op="redo" title="やり直す（⇧⌘Z）" aria-label="やり直す">${svg('redo')}</button>
        </div>`}
        <div class="ib-group">
          <button type="button" class="ib-ic" data-op="zoomout" title="縮小" aria-label="縮小">${svg('minus')}</button>
          <button type="button" class="ib-zoom" data-op="fit" title="全体を表示（0）">100%</button>
          <button type="button" class="ib-ic" data-op="zoomin" title="拡大" aria-label="拡大">${svg('plus')}</button>
        </div>
        ${ro ? '' : `
        <div class="ib-group">
          <select data-op="paper" aria-label="用紙">
            <option value="infinite">無限キャンバス</option>
            ${Object.entries(PAPERS).map(([k, p]) => `<option value="${k}">${p.label}</option>`).join('')}
          </select>
          <select data-op="orient" aria-label="用紙の向き"><option value="portrait">縦</option><option value="landscape">横</option></select>
          <select data-op="bg" aria-label="背景">${Object.entries(BACKGROUNDS).map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select>
        </div>
        <div class="ib-group">
          <button type="button" class="ib-text" data-op="finger" aria-pressed="false">指で描く</button>
          <button type="button" class="ib-ic ib-danger" data-op="clear" title="すべて消す" aria-label="すべて消す">${svg('trash')}</button>
        </div>`}
      </div>
      <div class="ib-areabar" hidden>
        <span>ドラッグして、参照表示で見せる範囲を囲んでください</span>
        <button type="button" data-op="areaScreen">今の画面を範囲にする</button>
        <button type="button" data-op="areaReset">指定を解除</button>
      </div>
      <div class="ib-stage">
        <canvas class="ib-main"></canvas>
        <canvas class="ib-live"></canvas>
        <div class="ib-viewrect" hidden><span>参照表示の範囲</span></div>
        <div class="ib-cursor" hidden></div>
        <div class="ib-toast" hidden></div>
      </div>`
    root.querySelector('.ib-title').textContent = this.options.title || ''
    this.stage = root.querySelector('.ib-stage')
    this.main = root.querySelector('.ib-main')
    this.live = root.querySelector('.ib-live')
    this.viewRectEl = root.querySelector('.ib-viewrect')
    this.cursorEl = root.querySelector('.ib-cursor')
    this.toastEl = root.querySelector('.ib-toast')

    this.prevOverflow = document.documentElement.style.overflow
    document.documentElement.style.overflow = 'hidden'
    this.prevFocus = document.activeElement
    document.body.appendChild(root)
    openCount++
    this.bind()
    this.syncUi()
    this.resize()
    this.fit()
    root.querySelector('[data-op="close"]').focus({ preventScroll: true })
  }

  bind() {
    const root = this.root, stage = this.stage
    root.addEventListener('click', (e) => {
      const b = e.target.closest('button')
      if (!b) return
      if (b.dataset.tool) this.setTool(b.dataset.tool)
      else if (b.dataset.color) this.setColor(b.dataset.color)
      else if (b.dataset.width) this.setWidth(+b.dataset.width)
      else this.command(b.dataset.op)
    })
    root.addEventListener('change', (e) => {
      const op = e.target.dataset.op
      if (op === 'paper' || op === 'orient' || op === 'bg') this.setMeta(op, e.target.value)
    })
    this.onKey = (e) => this.key(e)
    this.onKeyUp = (e) => { if (e.code === 'Space') { this.space = false; this.updateCursor() } }
    window.addEventListener('keydown', this.onKey, true)
    window.addEventListener('keyup', this.onKeyUp, true)
    this.onVis = () => { if (document.visibilityState === 'hidden') this.flush() }
    document.addEventListener('visibilitychange', this.onVis)
    this.ro = new ResizeObserver(() => this.resize())
    this.ro.observe(stage)

    stage.addEventListener('pointerdown', (e) => this.down(e))
    stage.addEventListener('pointermove', (e) => this.move(e))
    stage.addEventListener('pointerup', (e) => this.up(e))
    stage.addEventListener('pointercancel', (e) => this.up(e, true))
    stage.addEventListener('pointerleave', () => { this.cursorEl.hidden = true })
    stage.addEventListener('wheel', (e) => this.wheel(e), { passive: false })
    // iOS の拡大鏡・長押しメニュー・Safari のピンチによるページ拡大を止める
    for (const t of ['gesturestart', 'gesturechange', 'gestureend']) stage.addEventListener(t, (e) => this.gesture(e))
    stage.addEventListener('contextmenu', (e) => e.preventDefault())
    stage.addEventListener('touchstart', (e) => e.preventDefault(), { passive: false })
  }

  syncUi() {
    const q = (s) => this.root.querySelectorAll(s)
    q('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === this.tool))
    const colorTool = this.tool === 'marker' ? 'marker' : 'pen'
    q('[data-color]').forEach((b) => b.classList.toggle('on', b.dataset.color === this.colors[colorTool]))
    const wt = WIDTHS[this.tool] ? this.tool : 'pen'
    q('[data-width]').forEach((b) => b.classList.toggle('on', +b.dataset.width === this.widthIdx[wt]))
    const colorsEl = this.root.querySelector('.ib-colors')
    if (colorsEl) colorsEl.classList.toggle('ib-dim', this.tool === 'eraser' || this.tool === 'area' || this.tool === 'hand')
    const widthsEl = this.root.querySelector('.ib-widths')
    if (widthsEl) widthsEl.classList.toggle('ib-dim', this.tool === 'area' || this.tool === 'hand')
    const fb = this.root.querySelector('[data-op="finger"]')
    if (fb) { fb.classList.toggle('on', this.fingerDraw); fb.setAttribute('aria-pressed', String(this.fingerDraw)); fb.title = this.fingerDraw ? '指で描く：オン（指1本で描く）' : '指で描く：オフ（指は移動・拡大縮小）' }
    const u = this.root.querySelector('[data-op="undo"]'), r = this.root.querySelector('[data-op="redo"]')
    if (u) u.disabled = !this.undoStack.length
    if (r) r.disabled = !this.redoStack.length
    const d = this.data
    const ps = this.root.querySelector('[data-op="paper"]')
    if (ps) {
      ps.value = d.mode === 'infinite' ? 'infinite' : d.paper
      const os = this.root.querySelector('[data-op="orient"]')
      os.value = d.orient
      os.hidden = d.mode === 'infinite'
      this.root.querySelector('[data-op="bg"]').value = d.bg
    }
    this.root.querySelector('.ib-areabar').hidden = this.tool !== 'area'
    this.stage.dataset.tool = this.tool
    this.updateCursor()
  }

  updateCursor() {
    this.stage.classList.toggle('ib-panning', this.space || this.tool === 'hand' || this.readOnly)
  }

  toast(msg) {
    const t = this.toastEl
    t.textContent = msg
    t.hidden = false
    clearTimeout(this.toastTimer)
    this.toastTimer = setTimeout(() => { t.hidden = true }, 1800)
  }

  // ---- ツールと設定
  setTool(t) {
    this.tool = t
    this.syncUi()
    this.drawOverlay()
  }
  setColor(c) {
    const t = this.tool === 'marker' ? 'marker' : 'pen'
    if (this.tool !== 'pen' && this.tool !== 'marker') this.tool = 'pen'
    this.colors[t] = c
    this.prefs[t + 'Color'] = c
    savePrefs(this.prefs)
    this.syncUi()
  }
  setWidth(i) {
    if (!WIDTHS[this.tool]) this.tool = 'pen'
    this.widthIdx[this.tool] = i
    this.prefs[this.tool + 'Width'] = i
    savePrefs(this.prefs)
    this.syncUi()
  }
  setMeta(op, v) {
    const d = this.data
    this.pushUndo()
    if (op === 'paper') {
      if (v === 'infinite') d.mode = 'infinite'
      else {
        const prev = d.mode === 'paper' ? d.paper : ''
        d.mode = 'paper'
        d.paper = v
        // 16:9 と 4:3 は横向きで使うことが多いので、選び直したときは横にする
        if (/^W/.test(v) && !/^W/.test(prev)) d.orient = 'landscape'
      }
    } else if (op === 'orient') d.orient = v
    else d.bg = v
    this.changed()
    this.syncUi()
    if (op !== 'bg') this.fit()
    else this.redraw()
  }

  command(op) {
    switch (op) {
      case 'close': return this.close()
      case 'undo': return this.undo()
      case 'redo': return this.redo()
      case 'zoomin': return this.zoomAt(this.w / 2, this.h / 2, 1.25)
      case 'zoomout': return this.zoomAt(this.w / 2, this.h / 2, 0.8)
      case 'fit': return this.fit()
      case 'finger':
        this.fingerDraw = !this.fingerDraw
        this.prefs.fingerDraw = this.fingerDraw
        savePrefs(this.prefs)
        this.toast(this.fingerDraw ? '指1本で描けます（2本指で移動・拡大縮小）' : '指は移動・拡大縮小に使います')
        return this.syncUi()
      case 'clear':
        if (!this.data.strokes.length || !confirm('このボードの線をすべて消しますか？（元に戻すで戻せます）')) return
        this.pushUndo()
        this.data.strokes = []
        this.changed()
        return this.redraw()
      case 'areaScreen': {
        this.pushUndo()
        this.data.view = this.screenRect()
        this.changed()
        this.toast('今の画面を参照表示の範囲にしました')
        return this.drawOverlay()
      }
      case 'areaReset':
        if (!this.data.view) return this.toast('範囲は指定されていません（全体を表示します）')
        this.pushUndo()
        this.data.view = null
        this.changed()
        this.toast('範囲の指定を解除しました（全体を表示します）')
        return this.drawOverlay()
    }
  }

  // ---- 取り消し（線の配列は参照ごと保存する。線そのものは書き換えない）
  snapshot() {
    const d = this.data
    return { strokes: d.strokes.slice(), view: d.view, bg: d.bg, mode: d.mode, paper: d.paper, orient: d.orient }
  }
  restore(s) {
    const meta = s.mode !== this.data.mode || s.paper !== this.data.paper || s.orient !== this.data.orient
    Object.assign(this.data, s, { strokes: s.strokes.slice() })
    this.changed()
    this.syncUi()
    if (meta) this.fit()
    else this.redraw()
  }
  pushUndo() {
    this.undoStack.push(this.snapshot())
    if (this.undoStack.length > 200) this.undoStack.shift()
    this.redoStack = []
  }
  undo() {
    if (!this.undoStack.length) return
    this.redoStack.push(this.snapshot())
    this.restore(this.undoStack.pop())
  }
  redo() {
    if (!this.redoStack.length) return
    this.undoStack.push(this.snapshot())
    this.restore(this.redoStack.pop())
  }

  // ---- 変更の通知（線を書くたびではなく、少し待ってからまとめて）
  changed() {
    this.dirty = true
    this.syncUiButtons()
    if (this.readOnly) return
    clearTimeout(this.timer)
    this.timer = setTimeout(() => this.flush(), this.options.changeDelay ?? 800)
  }
  syncUiButtons() {
    const u = this.root.querySelector('[data-op="undo"]'), r = this.root.querySelector('[data-op="redo"]')
    if (u) u.disabled = !this.undoStack.length
    if (r) r.disabled = !this.redoStack.length
  }
  flush() {
    clearTimeout(this.timer)
    if (!this.dirty || this.readOnly) return
    this.dirty = false
    const v = serializeBoard(this.data)
    if (v === this.lastValue) return
    this.lastValue = v
    try { this.options.onChange?.(v) } catch (err) { console.error(err) }
  }
  close() {
    if (this.action?.type === 'draw') this.finishStroke()
    this.flush()
    window.removeEventListener('keydown', this.onKey, true)
    window.removeEventListener('keyup', this.onKeyUp, true)
    document.removeEventListener('visibilitychange', this.onVis)
    this.ro.disconnect()
    this.root.remove()
    openCount--
    if (!openCount) document.documentElement.style.overflow = this.prevOverflow
    try { this.prevFocus?.focus?.({ preventScroll: true }) } catch { /* 無視 */ }
    this.resolve(this.lastValue)
  }

  // ---- 座標
  resize() {
    const r = this.stage.getBoundingClientRect()
    this.w = r.width
    this.h = r.height
    this.left = r.left
    this.top = r.top
    this.dpr = window.devicePixelRatio || 1
    for (const c of [this.main, this.live]) {
      c.width = Math.max(1, Math.round(this.w * this.dpr))
      c.height = Math.max(1, Math.round(this.h * this.dpr))
    }
    this.colorsCss = sceneColors(this.root)
    this.redraw()
  }
  toWorld(cx, cy) {
    const { x, y, scale } = this.cam
    return { x: x + (cx - this.left) / scale, y: y + (cy - this.top) / scale }
  }
  screenRect() {
    const { x, y, scale } = this.cam
    return { x, y, w: this.w / scale, h: this.h / scale }
  }
  setCam(x, y, scale) {
    this.cam = { x, y, scale: clamp(scale, MIN_SCALE, MAX_SCALE) }
    const z = this.root.querySelector('.ib-zoom')
    if (z) z.textContent = Math.round(this.cam.scale * 100) + '%'
    this.requestRedraw()
  }
  zoomAt(sx, sy, factor) {
    const { x, y, scale } = this.cam
    const ns = clamp(scale * factor, MIN_SCALE, MAX_SCALE)
    const wx = x + sx / scale, wy = y + sy / scale
    this.setCam(wx - sx / ns, wy - sy / ns, ns)
  }
  /** 用紙は用紙全体、無限キャンバスは範囲の指定か線全体が収まるように */
  fit() {
    const d = this.data
    const a = d.mode === 'paper' ? { x: 0, y: 0, ...paperSize(d.paper, d.orient) } : previewArea(d)
    const m = 24
    const s = clamp(Math.min((this.w - m * 2) / a.w, (this.h - m * 2) / a.h), MIN_SCALE, MAX_SCALE)
    this.setCam(a.x - (this.w / s - a.w) / 2, a.y - (this.h / s - a.h) / 2, s)
  }

  // ---- 描画
  requestRedraw() {
    if (this.raf) return
    this.raf = requestAnimationFrame(() => { this.raf = 0; this.redraw() })
  }
  applyCam(ctx) {
    const { x, y, scale } = this.cam, k = this.dpr * scale
    ctx.setTransform(k, 0, 0, k, -x * k, -y * k)
  }
  redraw() {
    if (!this.w) return
    const ctx = this.main.getContext('2d')
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, this.main.width, this.main.height)
    this.applyCam(ctx)
    drawScene(ctx, this.data, this.screenRect(), this.cam.scale, this.colorsCss)
    this.drawLive()
    this.drawOverlay()
  }
  /** 書き足した1本だけを本体の絵に重ねる（全体は描き直さない） */
  drawOne(s) {
    const ctx = this.main.getContext('2d')
    this.applyCam(ctx)
    ctx.save()
    if (this.data.mode === 'paper') {
      const p = paperSize(this.data.paper, this.data.orient)
      ctx.beginPath()
      ctx.rect(0, 0, p.w, p.h)
      ctx.clip()
    }
    drawStroke(ctx, s)
    ctx.restore()
  }
  drawLive() {
    const ctx = this.live.getContext('2d')
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, this.live.width, this.live.height)
    const a = this.action
    if (!a || a.type !== 'draw' || !a.stroke.pts.length) return
    this.applyCam(ctx)
    drawStroke(ctx, a.stroke)
  }
  drawOverlay() {
    const el = this.viewRectEl
    const a = this.action
    const r = a?.type === 'area' ? a.rect : this.data.view
    if (!r || (this.readOnly && !this.data.view)) { el.hidden = true; return }
    const { x, y, scale } = this.cam
    el.hidden = false
    el.classList.toggle('ib-active', this.tool === 'area' || a?.type === 'area')
    Object.assign(el.style, {
      left: (r.x - x) * scale + 'px',
      top: (r.y - y) * scale + 'px',
      width: r.w * scale + 'px',
      height: r.h * scale + 'px',
    })
  }

  // ---- 入力
  isDrawInput(e) {
    if (this.readOnly) return false
    if (e.pointerType === 'pen') return true
    if (e.pointerType === 'mouse') return e.button === 0 && !this.space && this.tool !== 'hand'
    return this.fingerDraw && this.tool !== 'hand'
  }

  down(e) {
    if (e.pointerType === 'pen') {
      this.lastPen = performance.now()
      // 手のひらが先に触れて移動や指描きが始まっていたら、Pencil を優先する
      if (this.action && this.action.pointerType === 'touch') {
        if (this.action.type === 'draw') this.cancelStroke()
        this.action = null
        this.stage.classList.remove('ib-grabbing')
      }
      for (const [id, p] of this.pointers) if (p.type === 'touch') this.pointers.delete(id)
    } else if (e.pointerType === 'touch' && (this.action?.pointerType === 'pen' || performance.now() - (this.lastPen || 0) < 500)) {
      return // Pencil を使っている最中・直後の指（手のひら）は無視する
    }
    if (e.pointerType === 'pen' && !this.prefs.penSeen) {
      this.prefs.penSeen = true
      if (this.prefs.fingerDraw === undefined && this.fingerDraw) {
        this.fingerDraw = false
        this.toast('Apple Pencil で描き、指で移動・拡大縮小します')
        this.syncUi()
      }
      savePrefs(this.prefs)
    }
    try { this.stage.setPointerCapture(e.pointerId) } catch { /* 合成イベントなど */ }
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY, type: e.pointerType })
    const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')

    if (e.pointerType === 'touch' && touches.length >= 2) {
      // 2本目の指が来たら、指で描きかけた線は捨てて拡大縮小に切り替える
      if (this.action?.type === 'draw' && this.action.pointerType === 'touch') this.cancelStroke()
      if (this.action && this.action.pointerType === 'touch' && this.action.type !== 'pinch') this.action = null
      if (!this.action) this.startPinch()
      return
    }
    if (this.action) return // Pencil で描いている最中の指などは無視する

    if (this.isDrawInput(e)) {
      if (this.tool === 'pen' || this.tool === 'marker') this.startStroke(e)
      else if (this.tool === 'eraser') this.startErase(e)
      else if (this.tool === 'area') this.startArea(e)
    } else this.startPan(e)
  }

  move(e) {
    if (e.pointerType === 'pen') this.lastPen = performance.now()
    const p = this.pointers.get(e.pointerId)
    if (p) { p.x = e.clientX; p.y = e.clientY }
    this.hover(e)
    const a = this.action
    if (!a) return
    if (a.type === 'pinch') return this.movePinch()
    if (a.pointerId !== e.pointerId) return
    if (a.type === 'draw') this.moveStroke(e)
    else if (a.type === 'erase') this.moveErase(e)
    else if (a.type === 'area') this.moveArea(e)
    else if (a.type === 'pan') this.setCam(a.cx - (e.clientX - a.sx) / this.cam.scale, a.cy - (e.clientY - a.sy) / this.cam.scale, this.cam.scale)
  }

  up(e, cancelled) {
    if (e.pointerType === 'pen') this.lastPen = performance.now()
    this.pointers.delete(e.pointerId)
    this.stage.classList.remove('ib-grabbing')
    const a = this.action
    if (!a) return
    if (a.type === 'pinch') {
      const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')
      if (touches.length < 2) this.action = null
      return
    }
    if (a.pointerId !== e.pointerId) return
    if (a.type === 'draw') cancelled ? this.cancelStroke() : this.finishStroke()
    else if (a.type === 'erase') this.finishErase()
    else if (a.type === 'area') this.finishArea()
    this.action = null
  }

  hover(e) {
    const el = this.cursorEl
    if (this.readOnly || this.tool !== 'eraser' || e.pointerType === 'touch') { el.hidden = true; return }
    const r = WIDTHS.eraser[this.widthIdx.eraser] * this.cam.scale
    el.hidden = false
    Object.assign(el.style, { left: e.clientX - this.left + 'px', top: e.clientY - this.top + 'px', width: r * 2 + 'px', height: r * 2 + 'px' })
  }

  // 線
  startStroke(e) {
    const tool = this.tool
    const s = { tool, color: this.colors[tool], width: WIDTHS[tool][this.widthIdx[tool]], pts: [] }
    this.action = { type: 'draw', pointerId: e.pointerId, pointerType: e.pointerType, stroke: s }
    this.addPoint(e)
    this.drawLive()
  }
  addPoint(e) {
    const s = this.action.stroke
    const { x, y } = this.toWorld(e.clientX, e.clientY)
    const pr = e.pointerType === 'pen' ? clamp(e.pressure || 0.5, 0.05, 1) : 0.5
    const n = s.pts.length
    if (n) {
      const dx = x - s.pts[n - 3], dy = y - s.pts[n - 2]
      if (dx * dx + dy * dy < (0.6 / this.cam.scale) ** 2) { s.pts[n - 1] = Math.max(s.pts[n - 1], pr); return }
    }
    s.pts.push(x, y, pr)
  }
  moveStroke(e) {
    const list = e.getCoalescedEvents?.() || []
    for (const ce of list.length ? list : [e]) this.addPoint(ce)
    if (!this.liveRaf) this.liveRaf = requestAnimationFrame(() => { this.liveRaf = 0; this.drawLive() })
  }
  finishStroke() {
    const s = this.action.stroke
    this.action = null
    cancelAnimationFrame(this.liveRaf)
    this.liveRaf = 0
    this.drawLive()
    if (!s.pts.length) return
    strokeBox(s)
    this.pushUndo()
    this.data.strokes.push(s)
    this.drawOne(s)
    this.changed()
  }
  cancelStroke() {
    this.action = null
    this.drawLive()
  }

  // 消しゴム（線ごと消す）
  startErase(e) {
    this.action = { type: 'erase', pointerId: e.pointerId, pointerType: e.pointerType, before: this.snapshot(), hit: false, last: null }
    this.moveErase(e)
  }
  moveErase(e) {
    const a = this.action
    const list = e.getCoalescedEvents?.() || []
    const r = WIDTHS.eraser[this.widthIdx.eraser]
    let removed = false
    for (const ce of list.length ? list : [e]) {
      const pt = this.toWorld(ce.clientX, ce.clientY)
      const from = a.last || pt
      a.last = pt
      const keep = []
      for (const s of this.data.strokes) {
        if (hitStroke(s, from, pt, r)) removed = true
        else keep.push(s)
      }
      if (removed) this.data.strokes = keep
    }
    if (removed) {
      a.hit = true
      this.requestRedraw()
    }
  }
  finishErase() {
    const a = this.action
    if (!a.hit) return
    this.undoStack.push(a.before)
    this.redoStack = []
    this.changed()
  }

  // 参照表示の範囲
  startArea(e) {
    const p = this.toWorld(e.clientX, e.clientY)
    this.action = { type: 'area', pointerId: e.pointerId, pointerType: e.pointerType, p0: p, rect: { x: p.x, y: p.y, w: 0, h: 0 } }
    this.drawOverlay()
  }
  moveArea(e) {
    const a = this.action
    const p = this.toWorld(e.clientX, e.clientY)
    a.rect = { x: Math.min(a.p0.x, p.x), y: Math.min(a.p0.y, p.y), w: Math.abs(p.x - a.p0.x), h: Math.abs(p.y - a.p0.y) }
    this.drawOverlay()
  }
  finishArea() {
    const r = this.action.rect
    this.action = null
    if (r.w * this.cam.scale < 16 || r.h * this.cam.scale < 16) { this.drawOverlay(); return }
    this.pushUndo()
    this.data.view = r
    this.changed()
    this.drawOverlay()
    this.toast('参照表示の範囲を設定しました')
  }

  // 移動・拡大縮小
  startPan(e) {
    this.action = { type: 'pan', pointerId: e.pointerId, pointerType: e.pointerType, sx: e.clientX, sy: e.clientY, cx: this.cam.x, cy: this.cam.y }
    this.stage.classList.add('ib-grabbing')
  }
  startPinch() {
    const [a, b] = [...this.pointers.values()].filter((p) => p.type === 'touch')
    const mid = { x: (a.x + b.x) / 2 - this.left, y: (a.y + b.y) / 2 - this.top }
    this.action = {
      type: 'pinch',
      pointerType: 'touch',
      d0: Math.hypot(a.x - b.x, a.y - b.y) || 1,
      mid0: mid,
      cam0: { ...this.cam },
    }
  }
  movePinch() {
    const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')
    if (touches.length < 2) return
    const [a, b] = touches
    const p = this.action
    const mid = { x: (a.x + b.x) / 2 - this.left, y: (a.y + b.y) / 2 - this.top }
    const s = clamp(p.cam0.scale * (Math.hypot(a.x - b.x, a.y - b.y) / p.d0), MIN_SCALE, MAX_SCALE)
    const wx = p.cam0.x + p.mid0.x / p.cam0.scale, wy = p.cam0.y + p.mid0.y / p.cam0.scale
    this.setCam(wx - mid.x / s, wy - mid.y / s, s)
  }
  wheel(e) {
    e.preventDefault()
    if (e.ctrlKey || e.metaKey) this.zoomAt(e.clientX - this.left, e.clientY - this.top, Math.exp(-e.deltaY * 0.01))
    else this.setCam(this.cam.x + e.deltaX / this.cam.scale, this.cam.y + e.deltaY / this.cam.scale, this.cam.scale)
  }
  gesture(e) {
    // Mac の Safari のトラックパッドのピンチ
    e.preventDefault()
    if (e.type === 'gesturestart') this.gs = this.cam.scale
    else if (e.type === 'gesturechange' && this.gs) {
      const f = (this.gs * e.scale) / this.cam.scale
      this.zoomAt(e.clientX - this.left, e.clientY - this.top, f)
    }
  }

  key(e) {
    if (e.target.closest?.('select')) {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.close() }
      return
    }
    e.stopPropagation()
    const mod = e.metaKey || e.ctrlKey, k = e.key.toLowerCase()
    if (e.key === 'Escape') { e.preventDefault(); return this.close() }
    if (mod && k === 'z') { e.preventDefault(); return e.shiftKey ? this.redo() : this.undo() }
    if (mod && k === 'y') { e.preventDefault(); return this.redo() }
    if (mod && k === 's') { e.preventDefault(); return this.flush() }
    if (mod) return
    if (e.code === 'Space') { e.preventDefault(); this.space = true; return this.updateCursor() }
    if (k === '0') return this.fit()
    if (k === '+' || k === '=' || k === ';') return this.zoomAt(this.w / 2, this.h / 2, 1.25)
    if (k === '-') return this.zoomAt(this.w / 2, this.h / 2, 0.8)
    if (this.readOnly) return
    const tools = { p: 'pen', m: 'marker', e: 'eraser', r: 'area', h: 'hand' }
    if (tools[k]) this.setTool(tools[k])
  }
}

/** 線分 a-b の太さ r の帯に、線 s が触れているか */
function hitStroke(s, a, b, r) {
  const pad = r + s.width / 2
  const b0 = s.box
  if (Math.max(a.x, b.x) < b0.x0 - pad || Math.min(a.x, b.x) > b0.x1 + pad || Math.max(a.y, b.y) < b0.y0 - pad || Math.min(a.y, b.y) > b0.y1 + pad) return false
  const p = s.pts
  const lim = (r + s.width / 2) ** 2
  if (p.length === 3) return segDist2(p[0], p[1], a, b) <= lim
  for (let i = 0; i + 3 < p.length; i += 3) {
    if (segSegDist2(p[i], p[i + 1], p[i + 3], p[i + 4], a, b) <= lim) return true
  }
  return false
}
function segDist2(px, py, a, b) {
  const dx = b.x - a.x, dy = b.y - a.y
  const l = dx * dx + dy * dy
  const t = l ? clamp(((px - a.x) * dx + (py - a.y) * dy) / l, 0, 1) : 0
  const x = a.x + t * dx - px, y = a.y + t * dy - py
  return x * x + y * y
}
function segSegDist2(x1, y1, x2, y2, a, b) {
  return Math.min(
    segDist2(x1, y1, a, b),
    segDist2(x2, y2, a, b),
    segDist2(a.x, a.y, { x: x1, y: y1 }, { x: x2, y: y2 }),
    segDist2(b.x, b.y, { x: x1, y: y1 }, { x: x2, y: y2 })
  )
}
