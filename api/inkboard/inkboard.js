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
 * 手書きの線（s）に加えて、図形・矢印・文字（o）とグループ（各要素の g）を持てる。
 * 参照表示で見せる範囲（view）も値の中に持つ。編集画面の「範囲」で指定する。
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

/** 図形の種類と、タップで置いたときの大きさ */
export const SHAPES = {
  rect: { label: '四角（処理）', w: 120, h: 64 },
  round: { label: '角丸（開始・終了）', w: 120, h: 48 },
  diamond: { label: 'ひし形（分岐）', w: 128, h: 80 },
  ellipse: { label: '楕円', w: 112, h: 64 },
  para: { label: '平行四辺形（入力・出力）', w: 136, h: 64 },
  doc: { label: '書類', w: 120, h: 72 },
  cyl: { label: '円柱（データ）', w: 104, h: 84 },
  note: { label: '付箋', w: 128, h: 104 },
  frame: { label: '枠（囲み）', w: 288, h: 192 },
}

const DEFAULT_COLORS = ['#1f1f1f', '#d33a2c', '#2f5fd0', '#2e8b57', '#f2b100']
const PRESET_COLORS = ['#1f1f1f', '#5f5e5a', '#a6a49c', '#d33a2c', '#e8833a', '#f2b100', '#2e8b57', '#1d9e75', '#2f5fd0', '#5aa7e8', '#8a4fd6', '#d4537e']
const WIDTHS = { pen: [1.5, 3, 6], marker: [12, 20, 32], eraser: [6, 14, 30] }
const LINE_WIDTHS = [1.2, 2, 3.5]
const TOOL_LABEL = { pen: 'ペン', marker: 'マーカー', eraser: '消しゴム', select: '選択（囲んで選ぶ）', shape: '図形と線', text: '文字', area: '参照表示の範囲', hand: '移動' }
const DEFAULT_VIEW = { x: 0, y: 0, w: 800, h: 450 }
const Q = 4 // 座標は 1/4 ピクセル単位で保存
const MIN_SCALE = 0.1
const MAX_SCALE = 8
const GRID = 24 // グリッド吸着の間隔（背景の方眼・ドットと同じ）
const FONT = '-apple-system, BlinkMacSystemFont, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif'
const NOTE_FILL = '#fff4c2'
const SIDES = ['t', 'r', 'b', 'l']
const SIDE_DIR = { t: [0, -1], r: [1, 0], b: [0, 1], l: [-1, 0] }

const clamp = (v, a, b) => Math.min(b, Math.max(a, v))
const r1 = (n) => Math.round(n * 10) / 10
let seq = 0
const uid = () => (seq++).toString(36) + Math.random().toString(36).slice(2, 7)

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

const num = (v, d = 0) => (Number.isFinite(+v) ? +v : d)
const str = (v, d = '') => (typeof v === 'string' ? v : d)
const groups = (g) => (Array.isArray(g) && g.length ? g.map(String) : undefined)

function normEnd(e) {
  if (!e || typeof e !== 'object') return { x: 0, y: 0 }
  const out = { x: num(e.x), y: num(e.y) }
  if (e.s && SIDE_DIR[e.sd]) { out.s = String(e.s); out.sd = e.sd }
  return out
}

/** 保存されている図形・線・文字を、扱える形にそろえる（知らない種類は捨てる） */
function normObj(o) {
  if (!o || typeof o !== 'object' || !o.id) return null
  const base = { id: String(o.id) }
  const g = groups(o.g)
  if (g) base.g = g
  if (o.lk) base.lk = 1
  if (o.k === 'r') {
    if (!SHAPES[o.t]) return null
    return {
      ...base, k: 'r', t: o.t,
      x: num(o.x), y: num(o.y), w: Math.max(4, num(o.w, 120)), h: Math.max(4, num(o.h, 64)),
      c: str(o.c, '#1f1f1f'), f: str(o.f, o.t === 'frame' ? '' : 'w'), lw: num(o.lw, 2), d: o.d ? 1 : 0,
      tx: str(o.tx), tc: str(o.tc, '#1f1f1f'), fs: num(o.fs, o.t === 'frame' ? 14 : 15),
    }
  }
  if (o.k === 't') {
    return { ...base, k: 't', x: num(o.x), y: num(o.y), w: Math.max(16, num(o.w, 200)), h: Math.max(8, num(o.h, 24)), tx: str(o.tx), c: str(o.c, '#1f1f1f'), fs: num(o.fs, 16) }
  }
  if (o.k === 'l') {
    return {
      ...base, k: 'l', a: normEnd(o.a), b: normEnd(o.b),
      sh: ['s', 'e', 'c'].includes(o.sh) ? o.sh : 's', hd: ['e', 'b', 'n'].includes(o.hd) ? o.hd : 'e',
      d: o.d ? 1 : 0, c: str(o.c, '#1f1f1f'), lw: num(o.lw, 2),
    }
  }
  return null
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
    objs: [],
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
  const strokes = (Array.isArray(j.s) ? j.s : []).map(([tool, color, width, enc, meta]) => {
    const s = strokeBox({ tool: tool === 'marker' ? 'marker' : 'pen', color: String(color), width: +width || 2, pts: decPoints(String(enc)), enc: String(enc), id: uid() })
    if (meta && typeof meta === 'object') {
      if (meta.p) s.p = String(meta.p)
      const g = groups(meta.g)
      if (g) s.g = g
      if (meta.lk) s.lk = 1
    }
    return s
  }).filter((s) => s.pts.length >= 3)
  const objs = (Array.isArray(j.o) ? j.o : []).map(normObj).filter(Boolean)
  return { mode, paper, orient, bg: BACKGROUNDS[j.bg] ? j.bg : 'plain', view: normalizeView(j.view), strokes, objs }
}

/** データを値にする */
export function serializeBoard(d) {
  const objs = d.objs || []
  const idx = new Map(objs.map((o) => [o.id, o]))
  const out = {
    v: 1,
    mode: d.mode,
    paper: d.paper,
    orient: d.orient,
    bg: d.bg,
    view: d.view ? [r1(d.view.x), r1(d.view.y), r1(d.view.w), r1(d.view.h)] : null,
    s: d.strokes.map((s) => {
      const row = [s.tool, s.color, s.width, s.enc || (s.enc = encPoints(s.pts))]
      const meta = {}
      if (s.p) meta.p = s.p
      if (s.g?.length) meta.g = s.g
      if (s.lk) meta.lk = 1
      if (Object.keys(meta).length) row.push(meta)
      return row
    }),
  }
  if (objs.length) out.o = objs.map((o) => packObj(o, idx))
  return JSON.stringify(out)
}

function packEnd(e, idx) {
  const at = e.s && idx.get(e.s)
  const p = at ? sidePoint(at, e.sd) : e
  const out = { x: r1(p.x), y: r1(p.y) }
  if (at) { out.s = e.s; out.sd = e.sd }
  return out
}
function packObj(o, idx) {
  const out = { k: o.k, id: o.id }
  if (o.k === 'r') Object.assign(out, { t: o.t, x: r1(o.x), y: r1(o.y), w: r1(o.w), h: r1(o.h), c: o.c, f: o.f, lw: o.lw })
  else if (o.k === 't') Object.assign(out, { x: r1(o.x), y: r1(o.y), w: r1(o.w), h: r1(o.h), c: o.c, fs: o.fs })
  else Object.assign(out, { a: packEnd(o.a, idx), b: packEnd(o.b, idx), sh: o.sh, hd: o.hd, c: o.c, lw: o.lw })
  if (o.d) out.d = 1
  if (o.k !== 'l' && o.tx) out.tx = o.tx
  if (o.k === 'r' && (o.tx || o.tc !== '#1f1f1f')) { out.tc = o.tc; out.fs = o.fs }
  if (o.g?.length) out.g = o.g
  if (o.lk) out.lk = 1
  return out
}

// ---------------------------------------------------------------- 図形と線の形

function sidePoint(o, sd) {
  switch (sd) {
    case 't': return { x: o.x + o.w / 2, y: o.y }
    case 'r': return { x: o.x + o.w, y: o.y + o.h / 2 }
    case 'b': return { x: o.x + o.w / 2, y: o.y + o.h }
    default: return { x: o.x, y: o.y + o.h / 2 }
  }
}
function nearestSide(o, p) {
  let best = 't', bd = Infinity
  for (const sd of SIDES) {
    const q = sidePoint(o, sd)
    const dd = (q.x - p.x) ** 2 + (q.y - p.y) ** 2
    if (dd < bd) { bd = dd; best = sd }
  }
  return best
}
const objMap = (d) => new Map((d.objs || []).map((o) => [o.id, o]))

function endPos(e, idx) {
  const at = e.s && idx.get(e.s)
  if (at && at.k === 'r') { const p = sidePoint(at, e.sd); return { x: p.x, y: p.y, dir: SIDE_DIR[e.sd], att: true } }
  return { x: e.x, y: e.y, dir: null, att: false }
}
function axisDir(dx, dy) {
  return Math.abs(dx) >= Math.abs(dy) ? [Math.sign(dx) || 1, 0] : [0, Math.sign(dy) || 1]
}

/** 線の形。pts は描く点列（曲線は細かく区切った近似）、cubic は曲線の制御点 */
function lineGeom(o, idx) {
  const A = endPos(o.a, idx), B = endPos(o.b, idx)
  const da = A.dir || axisDir(B.x - A.x, B.y - A.y)
  const db = B.dir || axisDir(A.x - B.x, A.y - B.y)
  if (o.sh === 'c') {
    const k = Math.max(30, Math.hypot(B.x - A.x, B.y - A.y) * 0.4)
    const c1 = { x: A.x + da[0] * k, y: A.y + da[1] * k }, c2 = { x: B.x + db[0] * k, y: B.y + db[1] * k }
    const pts = []
    for (let i = 0; i <= 24; i++) {
      const t = i / 24, u = 1 - t
      pts.push({ x: u * u * u * A.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t * t * t * B.x, y: u * u * u * A.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t * t * t * B.y })
    }
    return { A, B, pts, cubic: [c1, c2], endA: c1, endB: c2 }
  }
  if (o.sh === 'e') {
    const g = 18
    const A1 = A.att ? { x: A.x + da[0] * g, y: A.y + da[1] * g } : { x: A.x, y: A.y }
    const B1 = B.att ? { x: B.x + db[0] * g, y: B.y + db[1] * g } : { x: B.x, y: B.y }
    const ha = da[0] !== 0, hb = db[0] !== 0
    let mid
    // 両端が同じ向きに出るときは、図形を横切らないように外側を回る
    if (ha && hb) {
      const mx = da[0] === db[0] ? (da[0] > 0 ? Math.max(A1.x, B1.x) : Math.min(A1.x, B1.x)) : (A1.x + B1.x) / 2
      mid = [{ x: mx, y: A1.y }, { x: mx, y: B1.y }]
    } else if (!ha && !hb) {
      const my = da[1] === db[1] ? (da[1] > 0 ? Math.max(A1.y, B1.y) : Math.min(A1.y, B1.y)) : (A1.y + B1.y) / 2
      mid = [{ x: A1.x, y: my }, { x: B1.x, y: my }]
    }
    else if (ha) mid = [{ x: B1.x, y: A1.y }]
    else mid = [{ x: A1.x, y: B1.y }]
    const raw = [A, A1, ...mid, B1, B]
    const pts = []
    for (const p of raw) {
      const q = pts[pts.length - 1]
      if (!q || Math.abs(q.x - p.x) > 0.01 || Math.abs(q.y - p.y) > 0.01) pts.push({ x: p.x, y: p.y })
    }
    return { A, B, pts, endA: pts[1] || B, endB: pts[pts.length - 2] || A }
  }
  return { A, B, pts: [{ x: A.x, y: A.y }, { x: B.x, y: B.y }], endA: B, endB: A }
}

function itemBox(it, idx) {
  if (!it.k) return it.box
  if (it.k === 'l') {
    const { pts } = lineGeom(it, idx)
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
    for (const p of pts) { x0 = Math.min(x0, p.x); y0 = Math.min(y0, p.y); x1 = Math.max(x1, p.x); y1 = Math.max(y1, p.y) }
    const pad = 6 + it.lw * 2
    return { x0: x0 - pad, y0: y0 - pad, x1: x1 + pad, y1: y1 + pad }
  }
  return { x0: it.x, y0: it.y, x1: it.x + it.w, y1: it.y + it.h }
}
function unionBox(boxes) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (const b of boxes) {
    if (!b) continue
    x0 = Math.min(x0, b.x0); y0 = Math.min(y0, b.y0); x1 = Math.max(x1, b.x1); y1 = Math.max(y1, b.y1)
  }
  return x0 === Infinity ? null : { x0, y0, x1, y1 }
}

/** 描いた線と図形の全体の範囲（何も無ければ null） */
export function contentBounds(d) {
  const idx = objMap(d)
  const b = unionBox([...d.strokes.map((s) => s.box), ...(d.objs || []).map((o) => itemBox(o, idx))])
  return b ? { x: b.x0, y: b.y0, w: b.x1 - b.x0, h: b.y1 - b.y0 } : null
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
  ctx.setLineDash([])
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

function roundRectPath(ctx, x, y, w, h, r) {
  r = Math.max(0, Math.min(r, w / 2, h / 2))
  ctx.moveTo(x + r, y)
  ctx.arcTo(x + w, y, x + w, y + h, r)
  ctx.arcTo(x + w, y + h, x, y + h, r)
  ctx.arcTo(x, y + h, x, y, r)
  ctx.arcTo(x, y, x + w, y, r)
  ctx.closePath()
}

/** 図形の輪郭のパスを作る（塗りと線に使う） */
function shapePath(ctx, o) {
  const { x, y, w, h } = o
  ctx.beginPath()
  switch (o.t) {
    case 'rect': roundRectPath(ctx, x, y, w, h, Math.min(8, w / 4, h / 4)); break
    case 'round': roundRectPath(ctx, x, y, w, h, h / 2); break
    case 'frame': roundRectPath(ctx, x, y, w, h, 8); break
    case 'diamond':
      ctx.moveTo(x + w / 2, y); ctx.lineTo(x + w, y + h / 2); ctx.lineTo(x + w / 2, y + h); ctx.lineTo(x, y + h / 2); ctx.closePath(); break
    case 'ellipse': ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2); break
    case 'para': {
      const k = Math.min(w * 0.18, h * 0.5)
      ctx.moveTo(x + k, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w - k, y + h); ctx.lineTo(x, y + h); ctx.closePath(); break
    }
    case 'doc': {
      const wv = Math.min(h * 0.14, 12)
      ctx.moveTo(x, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w, y + h - wv)
      ctx.bezierCurveTo(x + w * 0.7, y + h - wv * 3, x + w * 0.3, y + h + wv, x, y + h - wv)
      ctx.closePath(); break
    }
    case 'cyl': {
      const ry = Math.min(h * 0.15, 14), cx = x + w / 2
      ctx.moveTo(x, y + ry); ctx.lineTo(x, y + h - ry)
      ctx.ellipse(cx, y + h - ry, w / 2, ry, 0, Math.PI, 0, true)
      ctx.lineTo(x + w, y + ry)
      ctx.ellipse(cx, y + ry, w / 2, ry, 0, 0, Math.PI, true)
      ctx.closePath(); break
    }
    case 'note': {
      const f = Math.min(16, w / 4, h / 4)
      ctx.moveTo(x, y); ctx.lineTo(x + w - f, y); ctx.lineTo(x + w, y + f); ctx.lineTo(x + w, y + h); ctx.lineTo(x, y + h); ctx.closePath(); break
    }
  }
}

function fillOf(o, colors) {
  if (o.f === 'w') return o.t === 'note' ? NOTE_FILL : colors.paper
  return o.f || ''
}

const measureCtx = (() => {
  let c = null
  return () => (c ||= document.createElement('canvas').getContext('2d'))
})()

/** 文字を幅に合わせて折り返す（日本語は文字単位、英単語はなるべく切らない） */
function wrapText(ctx, text, maxW) {
  const out = []
  for (const para of String(text).split('\n')) {
    if (!para) { out.push(''); continue }
    const tokens = para.match(/[A-Za-z0-9_.,'’\-]+|\s+|./gu) || []
    let line = ''
    for (const tk of tokens) {
      const t = line + tk
      if (line && ctx.measureText(t).width > maxW) {
        out.push(line.trimEnd())
        if (ctx.measureText(tk).width > maxW) {
          line = ''
          for (const ch of tk) {
            if (line && ctx.measureText(line + ch).width > maxW) { out.push(line); line = ch } else line += ch
          }
        } else line = tk.trimStart()
      } else line = t
    }
    out.push(line)
  }
  return out
}

/** 図形の中で文字を置ける範囲 */
function textArea(o) {
  const { x, y, w, h } = o
  switch (o.t) {
    case 'diamond': return { x: x + w * 0.2, y: y + h * 0.2, w: w * 0.6, h: h * 0.6 }
    case 'ellipse': return { x: x + w * 0.14, y: y + h * 0.12, w: w * 0.72, h: h * 0.76 }
    case 'para': { const k = Math.min(w * 0.18, h * 0.5); return { x: x + k, y: y + 4, w: w - k * 2, h: h - 8 } }
    case 'cyl': { const ry = Math.min(h * 0.15, 14); return { x: x + 8, y: y + ry * 2, w: w - 16, h: h - ry * 3 } }
    case 'doc': { const wv = Math.min(h * 0.14, 12); return { x: x + 8, y: y + 4, w: w - 16, h: h - wv - 8 } }
    case 'frame': return { x: x + 10, y: y + 6, w: w - 20, h: 22 }
    default: return { x: x + 8, y: y + 6, w: w - 16, h: h - 12 }
  }
}

function drawTextBlock(ctx, text, area, fs, color, align, valign) {
  if (!text) return
  ctx.save()
  ctx.font = `${fs}px ${FONT}`
  ctx.fillStyle = color
  ctx.textBaseline = 'middle'
  ctx.textAlign = align
  const lh = fs * 1.35
  const lines = wrapText(ctx, text, Math.max(8, area.w))
  const total = lines.length * lh
  let y = valign === 'top' ? area.y + lh / 2 : area.y + (area.h - total) / 2 + lh / 2
  const x = align === 'center' ? area.x + area.w / 2 : area.x
  for (const l of lines) { ctx.fillText(l, x, y); y += lh }
  ctx.restore()
}

function drawShape(ctx, o, colors, hideText) {
  ctx.save()
  ctx.lineJoin = 'round'
  shapePath(ctx, o)
  const fill = fillOf(o, colors)
  if (fill) { ctx.fillStyle = fill; ctx.fill() }
  ctx.strokeStyle = o.c
  ctx.lineWidth = o.lw
  ctx.setLineDash(o.t === 'frame' ? [8, 5] : o.d ? [6, 4] : [])
  ctx.stroke()
  ctx.setLineDash(o.d ? [6, 4] : [])
  if (o.t === 'cyl') {
    const ry = Math.min(o.h * 0.15, 14)
    ctx.beginPath()
    ctx.ellipse(o.x + o.w / 2, o.y + ry, o.w / 2, ry, 0, Math.PI, 0, true)
    ctx.stroke()
  } else if (o.t === 'note') {
    const f = Math.min(16, o.w / 4, o.h / 4)
    ctx.beginPath()
    ctx.moveTo(o.x + o.w - f, o.y); ctx.lineTo(o.x + o.w - f, o.y + f); ctx.lineTo(o.x + o.w, o.y + f)
    ctx.stroke()
  }
  ctx.restore()
  if (!hideText && o.tx) {
    const frame = o.t === 'frame'
    drawTextBlock(ctx, o.tx, textArea(o), o.fs, o.tc, frame ? 'left' : 'center', frame ? 'top' : 'middle')
  }
}

function drawArrowHead(ctx, tip, from, size) {
  const a = Math.atan2(tip.y - from.y, tip.x - from.x)
  ctx.beginPath()
  ctx.moveTo(tip.x, tip.y)
  ctx.lineTo(tip.x - size * Math.cos(a - 0.42), tip.y - size * Math.sin(a - 0.42))
  ctx.lineTo(tip.x - size * Math.cos(a + 0.42), tip.y - size * Math.sin(a + 0.42))
  ctx.closePath()
  ctx.fill()
}

function drawLine(ctx, o, idx) {
  const g = lineGeom(o, idx)
  ctx.save()
  ctx.strokeStyle = o.c
  ctx.fillStyle = o.c
  ctx.lineWidth = o.lw
  ctx.lineCap = 'round'
  ctx.lineJoin = 'round'
  ctx.setLineDash(o.d ? [7, 5] : [])
  ctx.beginPath()
  ctx.moveTo(g.A.x, g.A.y)
  if (g.cubic) ctx.bezierCurveTo(g.cubic[0].x, g.cubic[0].y, g.cubic[1].x, g.cubic[1].y, g.B.x, g.B.y)
  else for (const p of g.pts.slice(1)) ctx.lineTo(p.x, p.y)
  ctx.stroke()
  ctx.setLineDash([])
  const size = 8 + o.lw * 2
  if (o.hd === 'e' || o.hd === 'b') drawArrowHead(ctx, g.B, g.endB, size)
  if (o.hd === 'b') drawArrowHead(ctx, g.A, g.endA, size)
  ctx.restore()
}

function drawTextObj(ctx, o) {
  drawTextBlock(ctx, o.tx, { x: o.x + 4, y: o.y + 2, w: o.w - 8, h: o.h - 4 }, o.fs, o.c, 'left', 'top')
}

/** 図形・線・文字を描く順番：枠 → 図形 → 手書き → 線 → 文字 */
function drawObjects(ctx, d, rect, colors, phase, hide) {
  const objs = d.objs || []
  if (!objs.length) return
  const idx = objMap(d)
  const vis = (b) => !(b.x1 < rect.x || b.y1 < rect.y || b.x0 > rect.x + rect.w || b.y0 > rect.y + rect.h)
  for (const o of objs) {
    if (phase === 'under') {
      if (o.k === 'r' && o.t === 'frame' && vis(itemBox(o, idx))) drawShape(ctx, o, colors, hide === o.id)
    } else if (phase === 'shapes') {
      if (o.k === 'r' && o.t !== 'frame' && vis(itemBox(o, idx))) drawShape(ctx, o, colors, hide === o.id)
    } else if (phase === 'over') {
      if (o.k === 'l' && vis(itemBox(o, idx))) drawLine(ctx, o, idx)
    } else if (phase === 'text') {
      if (o.k === 't' && hide !== o.id && vis(itemBox(o, idx))) drawTextObj(ctx, o)
    }
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
  if (area) {
    drawPattern(ctx, d.bg, area, scale, colors.rule)
    if (opts.snapDots && (d.bg === 'plain' || d.bg === 'lines')) drawPattern(ctx, 'dots', area, scale, colors.rule)
  }
  drawObjects(ctx, d, rect, colors, 'under', opts.hideText)
  drawObjects(ctx, d, rect, colors, 'shapes', opts.hideText)
  for (const s of d.strokes) {
    const b = s.box
    if (b.x1 < rect.x || b.y1 < rect.y || b.x0 > rect.x + rect.w || b.y0 > rect.y + rect.h) continue
    drawStroke(ctx, s)
  }
  drawObjects(ctx, d, rect, colors, 'over', opts.hideText)
  drawObjects(ctx, d, rect, colors, 'text', opts.hideText)
  ctx.restore()
}

function intersect(a, b) {
  const x = Math.max(a.x, b.x), y = Math.max(a.y, b.y)
  const w = Math.min(a.x + a.w, b.x + b.w) - x, h = Math.min(a.y + a.h, b.y + b.h) - y
  return w > 0 && h > 0 ? { x, y, w, h } : null
}

function drawPattern(ctx, bg, r, scale, color) {
  if (bg === 'plain') return
  const step = bg === 'lines' ? 32 : GRID
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
 *   prefs?: object | (() => object),      // 端末をまたいで共有したい設定（色・グリッド吸着など）
 *   onPrefsChange?: (prefs: object) => void,
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
    empty.textContent = data.strokes.length || data.objs.length ? '' : (options.emptyText ?? (options.onChange ? 'タップして手書き' : ''))
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
        prefs: options.prefs,
        onPrefsChange: options.onPrefsChange,
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
  select: '<path d="M12 4c4.6 0 8 2.2 8 5s-3.4 5-8 5c-1 0-2-.1-2.9-.3"/><path d="M6.2 12.4C4.8 11.5 4 10.3 4 9c0-1.5 1-2.8 2.6-3.7"/><path d="M8.5 13.5l-1 3.5 3.5-1"/><path d="M9.5 4.4l1.2-.3" stroke-dasharray="2 2"/>',
  shape: '<rect x="3.5" y="4" width="9" height="7" rx="1.6"/><circle cx="16.5" cy="16.5" r="4"/><path d="M8 11v3.5a2 2 0 0 0 2 2h2.5"/>',
  text: '<path d="M5 6V4.5h14V6"/><path d="M12 4.5v15"/><path d="M9 19.5h6"/>',
  area: '<path d="M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4"/><rect x="8" y="8" width="8" height="8" rx="1" opacity=".5"/>',
  hand: '<path d="M8 13V5.5a1.5 1.5 0 0 1 3 0V12M11 11.5V4a1.5 1.5 0 0 1 3 0v8M14 11.5V5.5a1.5 1.5 0 0 1 3 0V14c0 4-2.5 7-6 7-2.5 0-4-1.2-5.5-3.5L4 15a1.5 1.5 0 0 1 2.5-1.6L8 15"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
  minus: '<path d="M5 12h14"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
  grid: '<circle cx="6" cy="6" r=".9"/><circle cx="12" cy="6" r=".9"/><circle cx="18" cy="6" r=".9"/><circle cx="6" cy="12" r=".9"/><circle cx="12" cy="12" r=".9"/><circle cx="18" cy="12" r=".9"/><circle cx="6" cy="18" r=".9"/><circle cx="12" cy="18" r=".9"/><circle cx="18" cy="18" r=".9"/><rect x="9" y="9" width="9" height="9" rx="1.5" opacity=".6"/>',
  dots: '<circle cx="5.5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="18.5" cy="12" r="1"/>',
  grip: '<circle cx="9" cy="6" r=".9"/><circle cx="15" cy="6" r=".9"/><circle cx="9" cy="12" r=".9"/><circle cx="15" cy="12" r=".9"/><circle cx="9" cy="18" r=".9"/><circle cx="15" cy="18" r=".9"/>',
  exit: '<path d="M4 14h6v6M20 10h-6V4M14 10l6-6M4 20l6-6"/>',
  dash: '<path d="M3 12h3M10.5 12h3M18 12h3"/>',
}
const svg = (k) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`

/** パレットに並べる図形の小さな絵 */
function shapeIcon(t) {
  const s = 'fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"'
  const body = {
    rect: `<rect x="5" y="7" width="26" height="16" rx="3.5" ${s}/>`,
    round: `<rect x="4" y="8" width="28" height="14" rx="7" ${s}/>`,
    diamond: `<path d="M18 4L32 15L18 26L4 15z" ${s}/>`,
    ellipse: `<ellipse cx="18" cy="15" rx="14" ry="9" ${s}/>`,
    para: `<path d="M10 7H32L26 23H4z" ${s}/>`,
    doc: `<path d="M5 5H31V21Q24.5 16 18 21T5 21z" ${s}/>`,
    cyl: `<path d="M7 8V22A11 3.5 0 0 0 29 22V8" ${s}/><ellipse cx="18" cy="8" rx="11" ry="3.5" ${s}/>`,
    note: `<path d="M6 4H26L30 8V26H6z" ${s}/><path d="M26 4V8H30" ${s}/>`,
    frame: `<rect x="4" y="5" width="28" height="20" rx="3" ${s} stroke-dasharray="3 2.5"/>`,
  }[t]
  return `<svg viewBox="0 0 36 30" aria-hidden="true">${body}</svg>`
}
const LINE_ICON = {
  s: '<path d="M4 18L20 6"/>',
  e: '<path d="M4 18h8V6h8"/>',
  c: '<path d="M4 18C4 8 20 16 20 6"/>',
}
const HEAD_ICON = {
  e: '<path d="M4 12h15"/><path d="M15 8l4 4-4 4"/>',
  b: '<path d="M5 12h14"/><path d="M15 8l4 4-4 4M9 8l-4 4 4 4"/>',
  n: '<path d="M4 12h16"/>',
}
const lineSvg = (p) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`

const PREF_KEY = 'inkboard.prefs'
const LOCAL_ONLY = ['penSeen', 'fingerDraw', 'palPos']
function loadPrefs() {
  try { return JSON.parse(localStorage.getItem(PREF_KEY)) || {} } catch { return {} }
}
function savePrefs(p) {
  try { localStorage.setItem(PREF_KEY, JSON.stringify(p)) } catch { /* 保存できなくても動かす */ }
}
const sharedPrefs = (p) => Object.fromEntries(Object.entries(p).filter(([k]) => !LOCAL_ONLY.includes(k)))

let openCount = 0
/** コピーした図形や線（同じページの中なら、別のボードにも貼り付けられる） */
let clip = null

/**
 * 全画面の編集画面を開く。閉じると、その時点の値で解決する。
 * @param {string} value
 * @param {{ onChange?: (value: string) => void, readOnly?: boolean, title?: string, changeDelay?: number, prefs?: object | (() => object), onPrefsChange?: (prefs: object) => void }} [options]
 * @returns {Promise<string>}
 */
export function openBoard(value, options = {}) {
  return new Promise((resolve) => {
    const ed = new Editor(value, options, resolve)
    ed.mount()
  })
}

const isStroke = (it) => !it.k
function startsWith(g, e) {
  g = g || []
  if (g.length < e.length) return false
  for (let i = 0; i < e.length; i++) if (g[i] !== e[i]) return false
  return true
}
function inPoly(x, y, poly) {
  let inside = false
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const a = poly[i], b = poly[j]
    if ((a.y > y) !== (b.y > y) && x < ((b.x - a.x) * (y - a.y)) / (b.y - a.y) + a.x) inside = !inside
  }
  return inside
}
function distToPolyline(p, pts) {
  let best = Infinity
  for (let i = 0; i + 1 < pts.length; i++) best = Math.min(best, segDist2(p.x, p.y, pts[i], pts[i + 1]))
  return Math.sqrt(best)
}

/** 線や図形を、座標の写し方 f で写した新しいものを作る（元は書き換えない） */
function mapItem(it, f, k, idx) {
  if (!it.k) {
    const pts = it.pts.slice()
    for (let i = 0; i < pts.length; i += 3) { const q = f(pts[i], pts[i + 1]); pts[i] = q.x; pts[i + 1] = q.y }
    return strokeBox({ ...it, pts, enc: undefined })
  }
  if (it.k === 'l') {
    const m = (e) => (e.s && idx.get(e.s)?.k === 'r' ? e : { ...e, ...f(e.x, e.y) })
    return { ...it, a: m(it.a), b: m(it.b) }
  }
  const p0 = f(it.x, it.y), p1 = f(it.x + it.w, it.y + it.h)
  const out = { ...it, x: Math.min(p0.x, p1.x), y: Math.min(p0.y, p1.y), w: Math.max(8, Math.abs(p1.x - p0.x)), h: Math.max(8, Math.abs(p1.y - p0.y)) }
  if (k !== 1) out.fs = clamp(Math.round(it.fs * k * 2) / 2, 6, 200)
  return out
}
const translate = (dx, dy) => (x, y) => ({ x: x + dx, y: y + dy })

/** 文字の量に合わせて、文字ボックスの高さを決める */
function fitTextHeight(o) {
  const ctx = measureCtx()
  ctx.font = `${o.fs}px ${FONT}`
  const lines = wrapText(ctx, o.tx || ' ', Math.max(8, o.w - 8))
  return { ...o, h: Math.max(o.fs * 1.35 + 4, lines.length * o.fs * 1.35 + 4) }
}

class Editor {
  constructor(value, options, resolve) {
    this.options = options
    this.resolve = resolve
    this.readOnly = !!options.readOnly || typeof options.onChange !== 'function'
    this.data = parseBoard(value)
    this.lastValue = serializeBoard(this.data)
    const ext = typeof options.prefs === 'function' ? options.prefs() : options.prefs
    this.prefs = { ...loadPrefs(), ...(ext && typeof ext === 'object' ? sharedPrefs(ext) : {}) }
    const p = this.prefs
    this.tool = 'pen'
    this.colors = { pen: p.penColor || DEFAULT_COLORS[0], marker: p.markerColor || '#f2b100' }
    this.palette = Array.isArray(p.colors) && p.colors.length ? p.colors.filter((c) => /^#[0-9a-f]{6}$/i.test(c)).slice(0, 12) : DEFAULT_COLORS.slice()
    if (!this.palette.length) this.palette = DEFAULT_COLORS.slice()
    this.widthIdx = { pen: p.penWidth ?? 1, marker: p.markerWidth ?? 1, eraser: p.eraserWidth ?? 1 }
    this.lwIdx = clamp(p.lineWidth ?? 1, 0, 2)
    this.shapeKind = SHAPES[p.shapeKind] || p.shapeKind === 'line' ? p.shapeKind : 'rect'
    this.lineStyle = { sh: 's', hd: 'e', d: 0, ...(p.lineStyle || {}) }
    this.snap = !!p.snap
    this.palMode = p.palMode === 'auto' ? 'auto' : 'always'
    this.exitSide = p.exitSide === 'right' ? 'right' : 'left'
    // 指で描くか。Apple Pencil を一度でも使ったら、指は移動と拡大縮小に回す
    this.fingerDraw = p.fingerDraw ?? !p.penSeen
    this.sel = new Set()
    this.gpath = [] // 中に入っているグループ（外側から順に）
    this.undoStack = []
    this.redoStack = []
    this.cam = { x: 0, y: 0, scale: 1 }
    this.pointers = new Map()
    this.action = null
    this.space = false
    this.dirty = false
    this.palOpen = this.palMode === 'always'
    this.reindex()
  }

  // ---- 画面を組み立てる
  mount() {
    const root = (this.root = document.createElement('div'))
    root.className = 'ib-overlay'
    root.setAttribute('role', 'dialog')
    root.setAttribute('aria-label', this.options.title || '手書きボード')
    const ro = this.readOnly
    const toolBtn = (t) => `<button type="button" class="ib-ic" data-tool="${t}" title="${TOOL_LABEL[t]}" aria-label="${TOOL_LABEL[t]}">${svg(t)}</button>`
    root.innerHTML = `
      <div class="ib-stage">
        <canvas class="ib-main"></canvas>
        <canvas class="ib-live"></canvas>
        <canvas class="ib-ui"></canvas>
        <div class="ib-viewrect" hidden><span>参照表示の範囲</span></div>
        <div class="ib-cursor" hidden></div>
        <textarea class="ib-edit" hidden spellcheck="false" aria-label="文字"></textarea>
      </div>
      <div class="ib-menu" role="toolbar" aria-label="選んだものの操作" hidden></div>
      ${ro ? '' : `
      <div class="ib-pal" role="toolbar" aria-label="ペンと図形">
        <button type="button" class="ib-grip" title="ドラッグでパレットを動かす" aria-label="パレットを動かす">${svg('grip')}</button>
        <div class="ib-palrow">
          ${['pen', 'marker', 'eraser', 'select'].map(toolBtn).join('')}
          <button type="button" class="ib-ic" data-tool="shape" title="図形と線" aria-label="図形と線" aria-haspopup="true">${svg('shape')}</button>
          ${toolBtn('text')}
          <span class="ib-sep"></span>
          <div class="ib-colors"></div>
          <button type="button" class="ib-ic ib-add" data-op="addColor" title="色を追加" aria-label="色を追加">${svg('plus')}</button>
          <span class="ib-sep"></span>
          <button type="button" class="ib-ic ib-wbtn" data-op="width" title="太さ" aria-label="太さ"><i></i></button>
          <button type="button" class="ib-ic" data-op="undo" title="元に戻す（⌘Z・2本指でタップ）" aria-label="元に戻す">${svg('undo')}</button>
          <button type="button" class="ib-ic" data-op="redo" title="やり直す（⇧⌘Z・3本指でタップ）" aria-label="やり直す">${svg('redo')}</button>
          <button type="button" class="ib-ic" data-op="snap" title="グリッドに吸着" aria-label="グリッドに吸着" aria-pressed="false">${svg('grid')}</button>
          <button type="button" class="ib-ic" data-op="more" title="その他" aria-label="その他" aria-haspopup="true">${svg('dots')}</button>
        </div>
      </div>
      <button type="button" class="ib-fab" data-op="openPal" title="パレットを開く" aria-label="パレットを開く" hidden>${svg('pen')}</button>
      <div class="ib-pop ib-shapes" hidden>
        <div class="ib-shapegrid">${Object.entries(SHAPES).map(([k, s]) => `<button type="button" data-shape="${k}" title="${s.label}" aria-label="${s.label}">${shapeIcon(k)}</button>`).join('')}</div>
        <div class="ib-poprow">
          ${Object.entries({ s: '直線', e: 'カギ線', c: '曲線' }).map(([k, l]) => `<button type="button" class="ib-ic" data-line="${k}" title="${l}で線を引く" aria-label="${l}">${lineSvg(LINE_ICON[k])}</button>`).join('')}
          <span class="ib-sep"></span>
          ${Object.entries({ e: '片矢印', b: '両矢印', n: '矢印なし' }).map(([k, l]) => `<button type="button" class="ib-ic" data-head="${k}" title="${l}" aria-label="${l}">${lineSvg(HEAD_ICON[k])}</button>`).join('')}
          <button type="button" class="ib-ic" data-op="dash" title="点線" aria-label="点線" aria-pressed="false">${svg('dash')}</button>
        </div>
      </div>
      <div class="ib-pop ib-colorpop" hidden></div>
      <div class="ib-pop ib-fillpop" hidden>
        <button type="button" data-fill="">塗りなし</button>
        <button type="button" data-fill="w">白</button>
        <button type="button" data-fill="tint">薄い色</button>
      </div>`}
      <div class="ib-pop ib-more" hidden>
        ${ro ? '' : `
        <div class="ib-sec"><span>パレット</span><div class="ib-seg"><button type="button" data-pal="always">常に表示</button><button type="button" data-pal="auto">ペンで表示</button></div></div>
        <div class="ib-sec"><span>閉じるボタン</span><div class="ib-seg"><button type="button" data-exit="left">左下</button><button type="button" data-exit="right">右下</button></div></div>
        <div class="ib-sec"><span>指で描く</span><button type="button" class="ib-tog" data-op="finger" aria-pressed="false"></button></div>
        <div class="ib-sec"><span>用紙</span><div class="ib-selects">
          <select data-op="paper" aria-label="用紙"><option value="infinite">無限キャンバス</option>${Object.entries(PAPERS).map(([k, p]) => `<option value="${k}">${p.label}</option>`).join('')}</select>
          <select data-op="orient" aria-label="用紙の向き"><option value="portrait">縦</option><option value="landscape">横</option></select>
          <select data-op="bg" aria-label="背景">${Object.entries(BACKGROUNDS).map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select>
        </div></div>`}
        <div class="ib-sec"><span>表示</span><div class="ib-zoomrow">
          <button type="button" class="ib-ic" data-op="zoomout" title="縮小" aria-label="縮小">${svg('minus')}</button>
          <button type="button" class="ib-zoom" data-op="fit" title="全体を表示（0）">100%</button>
          <button type="button" class="ib-ic" data-op="zoomin" title="拡大" aria-label="拡大">${svg('plus')}</button>
        </div></div>
        ${ro ? '' : `
        <div class="ib-sec ib-wide">
          <button type="button" class="ib-line" data-tool="area">${svg('area')}参照表示の範囲を指定</button>
          <button type="button" class="ib-line" data-tool="hand">${svg('hand')}移動ツール</button>
          <button type="button" class="ib-line ib-danger" data-op="clear">${svg('trash')}すべて消す</button>
        </div>`}
      </div>
      <div class="ib-areabar" hidden>
        <span>ドラッグして、参照表示で見せる範囲を囲んでください</span>
        <button type="button" data-op="areaScreen">今の画面を範囲にする</button>
        <button type="button" data-op="areaReset">指定を解除</button>
        <button type="button" data-op="areaDone">終わる</button>
      </div>
      <button type="button" class="ib-exit" data-op="close" title="閉じる（Esc）" aria-label="閉じる">${svg('exit')}</button>
      ${ro ? `<button type="button" class="ib-romore" data-op="more" title="表示" aria-label="表示">${svg('dots')}</button>` : ''}
      <div class="ib-toast" hidden></div>`
    const $ = (s) => root.querySelector(s)
    this.stage = $('.ib-stage')
    this.main = $('.ib-main')
    this.live = $('.ib-live')
    this.ui = $('.ib-ui')
    this.viewRectEl = $('.ib-viewrect')
    this.cursorEl = $('.ib-cursor')
    this.toastEl = $('.ib-toast')
    this.editEl = $('.ib-edit')
    this.menuEl = $('.ib-menu')
    this.palEl = $('.ib-pal')
    this.fabEl = $('.ib-fab')
    this.exitEl = $('.ib-exit')

    this.prevOverflow = document.documentElement.style.overflow
    document.documentElement.style.overflow = 'hidden'
    this.prevFocus = document.activeElement
    document.body.appendChild(root)
    openCount++
    this.renderColors()
    this.bind()
    this.syncUi()
    this.resize()
    this.fit()
    this.placePalette()
    this.exitEl.focus({ preventScroll: true })
  }

  bind() {
    const root = this.root, stage = this.stage
    root.addEventListener('click', (e) => {
      const b = e.target.closest('button')
      if (!b || !root.contains(b)) return
      if (b === this.exitEl) {
        // ペン先が触れて閉じてしまわないように、閉じるのは指・マウス・キーボードだけ
        if (this.exitPen) { this.exitPen = false; return }
        return this.close()
      }
      if (b.dataset.tool) return this.toolClick(b.dataset.tool, b)
      if (b.dataset.color) return this.colorClick(b.dataset.color, +b.dataset.slot)
      if (b.dataset.shape) return this.pickShape(b.dataset.shape)
      if (b.dataset.line) return this.pickLine(b.dataset.line)
      if (b.dataset.head) return this.setLineStyle({ hd: b.dataset.head })
      if (b.dataset.pal) return this.setPalMode(b.dataset.pal)
      if (b.dataset.exit) return this.setExitSide(b.dataset.exit)
      if (b.dataset.fill !== undefined) return this.applyFill(b.dataset.fill)
      if (b.dataset.menu) return this.menuCommand(b.dataset.menu, b)
      if (b.dataset.preset) return this.colorPick(b.dataset.preset)
      this.command(b.dataset.op, b)
    })
    root.addEventListener('change', (e) => {
      const op = e.target.dataset.op
      if (op === 'paper' || op === 'orient' || op === 'bg') this.setMeta(op, e.target.value)
      else if (op === 'colorInput') this.colorPick(e.target.value)
    })
    root.addEventListener('pointerdown', (e) => {
      if (e.target.closest('.ib-pal, .ib-pop, .ib-menu')) this.palTouched()
      if (!e.target.closest('.ib-pop') && !e.target.closest('[aria-haspopup], [data-op="addColor"], [data-color], [data-menu="fill"]')) this.closePops()
    }, true)
    this.exitEl.addEventListener('pointerdown', (e) => { this.exitPen = e.pointerType === 'pen' })
    this.bindPalette()
    this.editEl.addEventListener('keydown', (e) => {
      e.stopPropagation()
      if (e.key === 'Escape' || (e.key === 'Enter' && (e.metaKey || e.ctrlKey))) { e.preventDefault(); this.commitText() }
    })
    this.editEl.addEventListener('input', () => this.layoutEdit())
    this.editEl.addEventListener('blur', () => { if (this.editing) this.commitText() })
    this.onKey = (e) => this.key(e)
    this.onKeyUp = (e) => { if (e.code === 'Space') { this.space = false; this.updateCursor() } }
    window.addEventListener('keydown', this.onKey, true)
    window.addEventListener('keyup', this.onKeyUp, true)
    this.onVis = () => { if (document.visibilityState === 'hidden') this.flush() }
    document.addEventListener('visibilitychange', this.onVis)
    this.ro = new ResizeObserver(() => { this.resize(); this.placePalette() })
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
    stage.addEventListener('touchstart', (e) => { if (e.target !== this.editEl) e.preventDefault() }, { passive: false })
  }

  // ---- パレット
  bindPalette() {
    const pal = this.palEl
    if (!pal) return
    const grip = pal.querySelector('.ib-grip')
    grip.addEventListener('pointerdown', (e) => {
      e.preventDefault()
      grip.setPointerCapture(e.pointerId)
      const r = pal.getBoundingClientRect(), rr = this.root.getBoundingClientRect()
      const ox = e.clientX - r.left, oy = e.clientY - r.top
      const mv = (ev) => {
        this.palPos = { x: ev.clientX - ox - rr.left, y: ev.clientY - oy - rr.top }
        this.placePalette()
      }
      const upFn = () => {
        grip.removeEventListener('pointermove', mv)
        grip.removeEventListener('pointerup', upFn)
        grip.removeEventListener('pointercancel', upFn)
        if (this.palPos) { this.prefs.palPos = this.palPos; this.savePrefs() }
      }
      grip.addEventListener('pointermove', mv)
      grip.addEventListener('pointerup', upFn)
      grip.addEventListener('pointercancel', upFn)
    })
    grip.addEventListener('dblclick', () => { this.palPos = null; delete this.prefs.palPos; this.savePrefs(); this.placePalette() })
    // 色は長押し（右クリック）で変更
    const colors = pal.querySelector('.ib-colors')
    colors.addEventListener('pointerdown', (e) => {
      const b = e.target.closest('[data-color]')
      if (!b) return
      clearTimeout(this.lpTimer)
      this.lpFired = false
      this.lpTimer = setTimeout(() => { this.lpFired = true; this.openColorPop(+b.dataset.slot, b) }, 500)
    })
    for (const t of ['pointerup', 'pointercancel', 'pointerleave']) colors.addEventListener(t, () => clearTimeout(this.lpTimer))
    colors.addEventListener('contextmenu', (e) => {
      const b = e.target.closest('[data-color]')
      if (!b) return
      e.preventDefault()
      this.lpFired = true
      this.openColorPop(+b.dataset.slot, b)
    })
  }
  placePalette() {
    const pal = this.palEl
    if (!pal) return
    this.palPos ||= this.prefs.palPos || null
    const rw = this.root.clientWidth, rh = this.root.clientHeight
    const pw = pal.offsetWidth, ph = pal.offsetHeight
    let x, y
    if (this.palPos) { x = this.palPos.x; y = this.palPos.y } else { x = (rw - pw) / 2; y = 10 }
    x = clamp(x, 6, Math.max(6, rw - pw - 6))
    y = clamp(y, 6, Math.max(6, rh - ph - 6))
    pal.style.left = x + 'px'
    pal.style.top = y + 'px'
  }
  renderColors() {
    const box = this.root.querySelector('.ib-colors')
    if (!box) return
    box.innerHTML = this.palette.map((c, i) => `<button type="button" class="ib-sw" data-color="${c}" data-slot="${i}" style="--c:${c}" title="${c}（長押しで変更）" aria-label="色 ${c}"></button>`).join('')
    this.syncUi()
  }
  setPalMode(m) {
    this.palMode = m
    this.prefs.palMode = m
    this.savePrefs()
    this.palOpen = m === 'always'
    this.toast(m === 'auto' ? 'ペンで書き始めるとパレットが出て、しばらくすると丸く縮みます' : 'パレットを常に表示します')
    this.syncUi()
  }
  setExitSide(s) {
    this.exitSide = s
    this.prefs.exitSide = s
    this.savePrefs()
    this.syncUi()
  }
  /** 自動表示のとき：パレットを開いて、しばらく触らなければ縮める */
  openPal() {
    if (this.palMode !== 'auto' || this.readOnly) return
    if (!this.palOpen) { this.palOpen = true; this.syncPal() }
    this.schedulePalClose()
  }
  palTouched() { if (this.palMode === 'auto' && this.palOpen) this.schedulePalClose() }
  schedulePalClose(ms = 3500) {
    clearTimeout(this.palTimer)
    if (this.palMode !== 'auto') return
    this.palTimer = setTimeout(() => {
      if (this.root.querySelector('.ib-pop:not([hidden])') || this.action?.type === 'draw') return this.schedulePalClose(1500)
      this.palOpen = false
      this.syncPal()
    }, ms)
  }
  syncPal() {
    if (!this.palEl) return
    const show = this.palMode === 'always' || this.palOpen
    this.palEl.hidden = !show
    this.fabEl.hidden = show
    if (show) this.placePalette()
  }

  // ---- ポップアップ
  showPop(sel, anchor) {
    const pop = this.root.querySelector(sel)
    if (!pop) return
    const was = !pop.hidden
    this.closePops()
    if (was) return
    pop.hidden = false
    const rr = this.root.getBoundingClientRect(), ar = anchor.getBoundingClientRect()
    const pw = pop.offsetWidth, ph = pop.offsetHeight
    let x = ar.left + ar.width / 2 - pw / 2 - rr.left
    let y = ar.bottom + 8 - rr.top
    if (y + ph > rr.height - 8) y = ar.top - ph - 8 - rr.top
    pop.style.left = clamp(x, 8, Math.max(8, rr.width - pw - 8)) + 'px'
    pop.style.top = clamp(y, 8, Math.max(8, rr.height - ph - 8)) + 'px'
    anchor.setAttribute('aria-expanded', 'true')
  }
  closePops() {
    this.root.querySelectorAll('.ib-pop').forEach((p) => { p.hidden = true })
    this.root.querySelectorAll('[aria-expanded="true"]').forEach((b) => b.setAttribute('aria-expanded', 'false'))
  }
  openColorPop(slot, anchor) {
    const pop = this.root.querySelector('.ib-colorpop')
    this.colorSlot = slot // -1 は追加
    const cur = slot >= 0 ? this.palette[slot] : this.colors.pen
    pop.innerHTML = `
      <div class="ib-presets">${PRESET_COLORS.map((c) => `<button type="button" class="ib-sw" data-preset="${c}" style="--c:${c}" aria-label="${c}"></button>`).join('')}</div>
      <div class="ib-poprow">
        <label class="ib-custom" title="好きな色を選ぶ"><input type="color" data-op="colorInput" value="${cur}"><span>ほかの色</span></label>
        ${slot >= 0 && this.palette.length > 1 ? `<button type="button" class="ib-line ib-danger" data-op="colorDel">${svg('trash')}この色を外す</button>` : ''}
      </div>
      <p class="ib-hint">${slot >= 0 ? 'パレットのこの色を置き換えます' : 'パレットに色を追加します（最大12色）'}</p>`
    pop.hidden = true
    this.showPop('.ib-colorpop', anchor)
  }
  colorPick(c) {
    if (!/^#[0-9a-f]{6}$/i.test(c)) return
    const slot = this.colorSlot
    if (slot >= 0) this.palette[slot] = c
    else if (this.palette.length < 12) this.palette.push(c)
    else this.toast('パレットは12色までです')
    this.prefs.colors = this.palette.slice()
    this.closePops()
    this.renderColors()
    this.colorClick(c, -1)
  }

  // ---- 状態を画面に反映
  syncUi() {
    if (!this.root) return
    const q = (s) => this.root.querySelectorAll(s)
    q('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === this.tool))
    const colorTool = this.tool === 'marker' ? 'marker' : 'pen'
    q('[data-color]').forEach((b) => b.classList.toggle('on', b.dataset.color === this.colors[colorTool]))
    const wb = this.root.querySelector('.ib-wbtn i')
    if (wb) {
      const wt = WIDTHS[this.tool] ? this.tool : null
      const i = wt ? this.widthIdx[wt] : this.lwIdx
      wb.style.setProperty('--d', [4, 7, 11][i] + 'px')
      wb.parentElement.title = wt ? `太さ（${['細', '中', '太'][i]}）` : `線の太さ（${['細', '中', '太'][i]}）`
    }
    q('[data-shape]').forEach((b) => b.classList.toggle('on', this.shapeKind === b.dataset.shape))
    q('[data-line]').forEach((b) => b.classList.toggle('on', this.shapeKind === 'line' && this.lineStyle.sh === b.dataset.line))
    q('[data-head]').forEach((b) => b.classList.toggle('on', this.lineStyle.hd === b.dataset.head))
    const dash = this.root.querySelector('[data-op="dash"]')
    if (dash) { dash.classList.toggle('on', !!this.lineStyle.d); dash.setAttribute('aria-pressed', String(!!this.lineStyle.d)) }
    const snap = this.root.querySelector('[data-op="snap"]')
    if (snap) { snap.classList.toggle('on', this.snap); snap.setAttribute('aria-pressed', String(this.snap)) }
    q('[data-pal]').forEach((b) => b.classList.toggle('on', b.dataset.pal === this.palMode))
    q('[data-exit]').forEach((b) => b.classList.toggle('on', b.dataset.exit === this.exitSide))
    this.exitEl.classList.toggle('ib-right', this.exitSide === 'right')
    const fb = this.root.querySelector('[data-op="finger"]')
    if (fb) { fb.classList.toggle('on', this.fingerDraw); fb.setAttribute('aria-pressed', String(this.fingerDraw)); fb.textContent = this.fingerDraw ? 'オン' : 'オフ' }
    this.syncUiButtons()
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
    this.syncPal()
    this.updateCursor()
  }
  syncUiButtons() {
    const u = this.root.querySelector('[data-op="undo"]'), r = this.root.querySelector('[data-op="redo"]')
    if (u) u.disabled = !this.undoStack.length
    if (r) r.disabled = !this.redoStack.length
  }
  updateCursor() {
    this.stage.classList.toggle('ib-panning', this.space || this.tool === 'hand' || this.readOnly)
  }
  toast(msg) {
    const t = this.toastEl
    t.textContent = msg
    t.hidden = false
    clearTimeout(this.toastTimer)
    this.toastTimer = setTimeout(() => { t.hidden = true }, 2200)
  }
  savePrefs() {
    savePrefs(this.prefs)
    try { this.options.onPrefsChange?.(sharedPrefs(this.prefs)) } catch (err) { console.error(err) }
  }

  // ---- ツールと設定
  setTool(t) {
    if (this.editing) this.commitText()
    this.tool = t
    if (!['select', 'shape', 'text'].includes(t)) { this.sel.clear(); this.gpath = [] }
    this.syncUi()
    this.drawUi()
    this.drawOverlay()
  }
  toolClick(t, btn) {
    if (t === 'shape') {
      if (this.tool !== 'shape') this.setTool('shape')
      return this.showPop('.ib-shapes', btn)
    }
    this.closePops()
    this.setTool(t)
  }
  pickShape(k) {
    this.shapeKind = k
    this.prefs.shapeKind = k
    this.savePrefs()
    if (this.tool !== 'shape') this.setTool('shape')
    this.closePops()
    this.syncUi()
    this.toast(`${SHAPES[k].label}：タップで置く・ドラッグで大きさを決める`)
  }
  pickLine(sh) {
    this.shapeKind = 'line'
    this.prefs.shapeKind = 'line'
    this.setLineStyle({ sh })
    if (this.tool !== 'shape') this.setTool('shape')
    this.closePops()
    this.toast('ドラッグで線を引きます（図形の上で離すとつながります）')
  }
  setLineStyle(patch) {
    Object.assign(this.lineStyle, patch)
    this.prefs.lineStyle = { ...this.lineStyle }
    this.savePrefs()
    // 線を選んでいれば、その線も変える
    const lines = this.selItems().filter((it) => it.k === 'l' && !it.lk)
    if (lines.length) this.updateItems(lines, (o) => ({ ...o, ...patch }))
    this.syncUi()
  }
  colorClick(c, slot) {
    if (this.lpFired) { this.lpFired = false; return }
    // 図形などを選んでいるときは、選んだものの色を変える
    const items = this.selItems().filter((it) => !it.lk)
    if (items.length && ['select', 'shape', 'text'].includes(this.tool)) {
      this.updateItems(items, (it) => (it.k ? { ...it, c } : { ...it, color: c }))
    }
    const t = this.tool === 'marker' ? 'marker' : 'pen'
    if (this.tool === 'eraser' || this.tool === 'area' || this.tool === 'hand') this.tool = 'pen'
    this.colors[t] = c
    this.prefs[t + 'Color'] = c
    this.savePrefs()
    this.syncUi()
  }
  cycleWidth() {
    const wt = WIDTHS[this.tool] ? this.tool : null
    if (wt) {
      this.widthIdx[wt] = (this.widthIdx[wt] + 1) % 3
      this.prefs[wt + 'Width'] = this.widthIdx[wt]
    } else {
      this.lwIdx = (this.lwIdx + 1) % 3
      this.prefs.lineWidth = this.lwIdx
      const items = this.selItems().filter((it) => it.k && it.k !== 't' && !it.lk)
      if (items.length) this.updateItems(items, (o) => ({ ...o, lw: LINE_WIDTHS[this.lwIdx] }))
    }
    this.savePrefs()
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

  command(op, btn) {
    switch (op) {
      case 'close': return this.close()
      case 'undo': return this.undo()
      case 'redo': return this.redo()
      case 'zoomin': return this.zoomAt(this.w / 2, this.h / 2, 1.25)
      case 'zoomout': return this.zoomAt(this.w / 2, this.h / 2, 0.8)
      case 'fit': return this.fit()
      case 'width': return this.cycleWidth()
      case 'more': return this.showPop('.ib-more', btn)
      case 'addColor': return this.openColorPop(-1, btn)
      case 'openPal': this.palOpen = true; this.syncPal(); return this.schedulePalClose(5000)
      case 'dash': return this.setLineStyle({ d: this.lineStyle.d ? 0 : 1 })
      case 'colorDel':
        if (this.colorSlot >= 0 && this.palette.length > 1) {
          this.palette.splice(this.colorSlot, 1)
          this.prefs.colors = this.palette.slice()
          this.savePrefs()
          this.closePops()
          this.renderColors()
        }
        return
      case 'snap':
        this.snap = !this.snap
        this.prefs.snap = this.snap
        this.savePrefs()
        this.toast(this.snap ? 'グリッドに吸着します' : 'グリッドへの吸着をやめました')
        this.syncUi()
        return this.redraw()
      case 'finger':
        this.fingerDraw = !this.fingerDraw
        this.prefs.fingerDraw = this.fingerDraw
        this.savePrefs()
        this.toast(this.fingerDraw ? '指1本で描けます（2本指で移動・拡大縮小）' : '指は移動・拡大縮小に使います')
        return this.syncUi()
      case 'clear':
        this.closePops()
        if ((!this.data.strokes.length && !this.data.objs.length) || !confirm('このボードの線と図形をすべて消しますか？（元に戻すで戻せます）')) return
        this.pushUndo()
        this.data.strokes = []
        this.data.objs = []
        this.sel.clear()
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
      case 'areaDone': return this.setTool('pen')
    }
  }

  // ---- 取り消し（配列は参照ごと保存する。線・図形そのものは書き換えずに差し替える）
  snapshot() {
    const d = this.data
    return { strokes: d.strokes.slice(), objs: d.objs.slice(), view: d.view, bg: d.bg, mode: d.mode, paper: d.paper, orient: d.orient }
  }
  restore(s) {
    const meta = s.mode !== this.data.mode || s.paper !== this.data.paper || s.orient !== this.data.orient
    Object.assign(this.data, s, { strokes: s.strokes.slice(), objs: s.objs.slice() })
    this.reindex()
    for (const id of [...this.sel]) if (!this.idx.has(id)) this.sel.delete(id)
    this.changed()
    this.syncUi()
    if (meta) this.fit()
    else this.redraw()
  }
  pushUndo(snap) {
    this.undoStack.push(snap || this.snapshot())
    if (this.undoStack.length > 200) this.undoStack.shift()
    this.redoStack = []
  }
  undo() {
    if (this.editing) this.commitText()
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
    this.reindex()
    this.dirty = true
    this.syncUiButtons()
    if (this.readOnly) return
    clearTimeout(this.timer)
    this.timer = setTimeout(() => this.flush(), this.options.changeDelay ?? 800)
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
    if (this.editing) this.commitText()
    if (this.action?.type === 'draw') this.finishStroke()
    this.flush()
    clearTimeout(this.palTimer)
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

  // ---- データの出し入れ
  reindex() {
    const m = new Map()
    for (const s of this.data.strokes) m.set(s.id, s)
    for (const o of this.data.objs) m.set(o.id, o)
    this.idx = m
  }
  selItems() { return [...this.sel].map((id) => this.idx.get(id)).filter(Boolean) }
  /** 線・図形を差し替える（元の並び順のまま） */
  replaceItems(map) {
    const d = this.data
    d.strokes = d.strokes.map((s) => map.get(s.id) || s)
    d.objs = d.objs.map((o) => map.get(o.id) || o)
    for (const [id, it] of map) this.idx.set(id, it)
  }
  updateItems(items, fn) {
    this.pushUndo()
    const map = new Map()
    for (const it of items) map.set(it.id, fn(it))
    this.replaceItems(map)
    this.changed()
    this.redraw()
  }
  parentOf(it) {
    if (it && isStroke(it) && it.p) { const p = this.idx.get(it.p); if (p) return p }
    return it
  }
  /** タップしたものから、選ぶ単位（今いる階層のグループ全体、または1つ）を決める */
  resolveIds(it) {
    const base = this.parentOf(it)
    const g = base.g || []
    if (!startsWith(g, this.gpath)) this.gpath = []
    const L = this.gpath.length
    if (g.length > L) {
      const gid = g[L]
      return [...this.idx.values()].filter((x) => (x.g || [])[L] === gid && startsWith(x.g, this.gpath)).map((x) => x.id)
    }
    return [base.id]
  }
  /** 選んだものに、図形の中の手書き（と、動かすときは枠の中身）を足す */
  expandIds(ids, forMove) {
    const out = new Set(ids)
    if (forMove) {
      for (const id of ids) {
        const f = this.idx.get(id)
        if (f?.k !== 'r' || f.t !== 'frame') continue
        const fb = itemBox(f, this.idx)
        for (const it of this.idx.values()) {
          if (it === f || it.lk) continue
          const b = itemBox(it, this.idx)
          if (b.x0 >= fb.x0 - 1 && b.y0 >= fb.y0 - 1 && b.x1 <= fb.x1 + 1 && b.y1 <= fb.y1 + 1) out.add(it.id)
        }
      }
    }
    for (const s of this.data.strokes) if (s.p && out.has(s.p)) out.add(s.id)
    return out
  }
  selBox(ids = this.sel) {
    return unionBox([...ids].map((id) => this.idx.get(id)).filter(Boolean).map((it) => itemBox(it, this.idx)))
  }
  isGroupSel() {
    const items = this.selItems()
    if (items.length < 2) return false
    const L = this.gpath.length
    const gid = (items[0].g || [])[L]
    return !!gid && items.every((it) => (it.g || [])[L] === gid)
  }
  /** 選んだものの「単位」の数（グループは1つと数える） */
  selUnits() {
    const L = this.gpath.length
    const units = new Set()
    for (const it of this.selItems()) {
      if (isStroke(it) && it.p && this.sel.has(it.p)) continue
      units.add((it.g || [])[L] || it.id)
    }
    return units.size
  }
  /** 手書きの線が図形の中に収まっていれば、その図形の中身にする */
  reparent(ids) {
    const shapes = this.data.objs.filter((o) => o.k === 'r' && o.t !== 'frame')
    const map = new Map()
    for (const id of ids) {
      const s = this.idx.get(id)
      if (!s || !isStroke(s)) continue
      const b = s.box, w = s.width
      let parent = null
      for (let i = shapes.length - 1; i >= 0; i--) {
        const o = shapes[i]
        if (b.x0 + w >= o.x - 4 && b.y0 + w >= o.y - 4 && b.x1 - w <= o.x + o.w + 4 && b.y1 - w <= o.y + o.h + 4) { parent = o; break }
      }
      const p = parent?.id
      if ((s.p || undefined) !== p) {
        const n = { ...s }
        if (p) n.p = p
        else delete n.p
        map.set(id, n)
      }
    }
    if (map.size) this.replaceItems(map)
  }

  // ---- 選んだものの操作
  setSel(ids) {
    this.sel = new Set(ids)
    this.drawUi()
  }
  clearSel() {
    if (!this.sel.size && !this.gpath.length) return
    this.sel.clear()
    this.gpath = []
    this.redraw()
  }
  deleteSel() {
    const items = this.selItems()
    if (!items.length) return
    if (items.some((it) => it.lk)) return this.toast('ロック中のものは消せません')
    this.pushUndo()
    this.removeIds(this.expandIds(this.sel))
    this.sel.clear()
    this.changed()
    this.redraw()
  }
  removeIds(ids) {
    const d = this.data
    // 消す図形につながっていた線は、その位置で切り離して残す
    const free = (e) => {
      if (!e.s || !ids.has(e.s)) return e
      const at = this.idx.get(e.s)
      const p = at ? sidePoint(at, e.sd) : e
      return { x: p.x, y: p.y }
    }
    d.strokes = d.strokes.filter((s) => !ids.has(s.id))
    d.objs = d.objs.filter((o) => !ids.has(o.id)).map((o) => (o.k === 'l' && ((o.a.s && ids.has(o.a.s)) || (o.b.s && ids.has(o.b.s))) ? { ...o, a: free(o.a), b: free(o.b) } : o))
    this.reindex()
  }
  copySel(cut) {
    const ids = this.expandIds(this.sel)
    if (!ids.size) return
    const items = [...ids].map((id) => this.idx.get(id)).filter(Boolean)
    const fixEnd = (e) => {
      if (e.s && ids.has(e.s)) return { ...e }
      const p = endPos(e, this.idx)
      return { x: p.x, y: p.y }
    }
    clip = {
      items: items.map((it) => (isStroke(it) ? { ...it, pts: it.pts.slice() } : it.k === 'l' ? { ...it, a: fixEnd(it.a), b: fixEnd(it.b) } : { ...it })),
      box: this.selBox(ids),
      n: 0,
    }
    if (cut) {
      if (items.some((it) => it.lk)) return this.toast('ロック中のものは切り取れません')
      this.pushUndo()
      this.removeIds(ids)
      this.sel.clear()
      this.changed()
      this.redraw()
    } else this.toast('コピーしました')
  }
  /** コピーしたものを置く。at を渡すとその位置に、無ければ少しずらして */
  pasteClip(at, src = clip) {
    if (!src) return
    const b = src.box
    let dx, dy
    if (at) { dx = at.x - (b.x0 + b.x1) / 2; dy = at.y - (b.y0 + b.y1) / 2 } else { src.n++; dx = dy = GRID * src.n }
    if (this.snap) { dx = Math.round(dx / GRID) * GRID; dy = Math.round(dy / GRID) * GRID }
    const ids = new Map(src.items.map((it) => [it.id, uid()]))
    const gids = new Map()
    const regroup = (g) => g?.map((x) => gids.get(x) || (gids.set(x, uid()), gids.get(x)))
    const f = translate(dx, dy)
    const fresh = src.items.map((it) => {
      let n = { ...it, id: ids.get(it.id) }
      if (it.g) n.g = regroup(it.g)
      if (isStroke(it)) {
        if (it.p) { if (ids.has(it.p)) n.p = ids.get(it.p); else delete n.p }
        return mapItem(n, f, 1, new Map())
      }
      if (it.k === 'l') {
        const end = (e) => (e.s && ids.has(e.s) ? { ...e, s: ids.get(e.s), ...f(e.x, e.y) } : { x: e.x + dx, y: e.y + dy })
        return { ...n, a: end(it.a), b: end(it.b) }
      }
      return { ...n, x: it.x + dx, y: it.y + dy }
    })
    this.pushUndo()
    for (const it of fresh) {
      delete it.lk
      if (isStroke(it)) this.data.strokes.push(it)
      else this.data.objs.push(it)
    }
    this.changed()
    this.gpath = []
    const top = fresh.filter((it) => !(isStroke(it) && it.p && ids.size && [...ids.values()].includes(it.p)))
    this.setSel(top.map((it) => it.id))
    this.redraw()
  }
  duplicateSel() {
    const keep = clip
    this.copySel(false)
    const dup = clip
    clip = keep
    if (dup) this.pasteClip(null, dup)
  }
  groupSel() {
    if (this.selUnits() < 2) return
    const gid = uid(), L = this.gpath.length
    const items = this.selItems().filter((it) => !(isStroke(it) && it.p && this.sel.has(it.p)))
    this.updateItems(items, (it) => {
      const g = it.g || []
      return { ...it, g: [...g.slice(0, L), gid, ...g.slice(L)] }
    })
    this.toast('グループにしました（ダブルタップで中を個別に選べます）')
  }
  ungroupSel() {
    if (!this.isGroupSel()) return
    const L = this.gpath.length
    const items = this.selItems()
    this.updateItems(items, (it) => {
      const g = (it.g || []).filter((_, i) => i !== L)
      const n = { ...it }
      if (g.length) n.g = g
      else delete n.g
      return n
    })
    this.toast('グループを解除しました')
  }
  lockSel(on) {
    const ids = this.expandIds(this.sel)
    const items = [...ids].map((id) => this.idx.get(id)).filter(Boolean)
    this.updateItems(items, (it) => {
      const n = { ...it }
      if (on) n.lk = 1
      else delete n.lk
      return n
    })
    this.toast(on ? 'ロックしました（動かしたり消したりできません）' : 'ロックを解除しました')
  }
  orderSel(front) {
    const ids = this.expandIds(this.sel)
    this.pushUndo()
    const d = this.data
    const pick = (arr) => [arr.filter((x) => !ids.has(x.id)), arr.filter((x) => ids.has(x.id))]
    const [so, si] = pick(d.strokes), [oo, oi] = pick(d.objs)
    d.strokes = front ? [...so, ...si] : [...si, ...so]
    d.objs = front ? [...oo, ...oi] : [...oi, ...oo]
    this.changed()
    this.redraw()
  }
  applyFill(f) {
    const shapes = this.selItems().filter((it) => it.k === 'r' && !it.lk)
    const fill = f === 'tint' ? this.colors.pen + '2e' : f
    if (shapes.length) this.updateItems(shapes, (o) => ({ ...o, f: fill }))
    this.closePops()
  }

  menuCommand(op, btn) {
    switch (op) {
      case 'copy': return this.copySel(false)
      case 'cut': return this.copySel(true)
      case 'dup': return this.duplicateSel()
      case 'del': return this.deleteSel()
      case 'group': return this.groupSel()
      case 'ungroup': return this.ungroupSel()
      case 'front': return this.orderSel(true)
      case 'back': return this.orderSel(false)
      case 'lock': return this.lockSel(true)
      case 'unlock': return this.lockSel(false)
      case 'fill': return this.showPop('.ib-fillpop', btn)
      case 'text': { const it = this.selItems()[0]; if (it) this.editText(it); return }
      case 'paste': this.menuEl.hidden = true; return this.pasteClip(this.pasteAt)
    }
  }
  showMenu() {
    const m = this.menuEl
    if (this.readOnly || this.action || this.editing || !this.sel.size || !['select', 'shape', 'text'].includes(this.tool)) { if (!this.pasteAt) m.hidden = true; return }
    this.pasteAt = null
    const items = this.selItems()
    const locked = items.some((it) => it.lk)
    const b = (op, label, cls = '') => `<button type="button" data-menu="${op}"${cls ? ` class="${cls}"` : ''}>${label}</button>`
    const out = []
    if (locked) out.push(b('copy', 'コピー'), b('unlock', 'ロック解除'))
    else {
      out.push(b('copy', 'コピー'), b('cut', '切り取り'), b('dup', '複製'))
      if (this.selUnits() >= 2) out.push(b('group', 'グループ化'))
      if (this.isGroupSel()) out.push(b('ungroup', 'グループ解除'))
      const single = items.length === 1 ? items[0] : null
      if (single && (single.k === 'r' || single.k === 't')) out.push(b('text', '文字'))
      if (items.some((it) => it.k === 'r')) out.push(b('fill', '塗り'))
      out.push(b('front', '前面へ'), b('back', '背面へ'), b('lock', 'ロック'), b('del', '削除', 'ib-danger'))
    }
    m.innerHTML = out.join('')
    m.hidden = false
    this.placeMenu(this.selBox())
  }
  placeMenu(box, pt) {
    const m = this.menuEl
    const mw = m.offsetWidth, mh = m.offsetHeight
    let x, y
    if (box) {
      const s0 = this.toScreen(box.x0, box.y0), s1 = this.toScreen(box.x1, box.y1)
      x = (s0.x + s1.x) / 2 - mw / 2
      y = s0.y - mh - 22
      if (y < 60) y = s1.y + 22
      if (y + mh > this.h - 8) y = Math.max(8, s0.y + 8)
    } else { x = pt.x - mw / 2; y = pt.y - mh - 14 }
    m.style.left = clamp(x + this.stage.offsetLeft, 8, Math.max(8, this.w - mw - 8)) + 'px'
    m.style.top = clamp(y + this.stage.offsetTop, 8, Math.max(8, this.h - mh - 8)) + 'px'
  }
  showPasteMenu(sp, wp) {
    if (!clip || this.readOnly) return
    this.pasteAt = wp
    const m = this.menuEl
    m.innerHTML = '<button type="button" data-menu="paste">ここに貼り付け</button>'
    m.hidden = false
    this.placeMenu(null, sp)
  }

  // ---- 文字
  editText(o, before) {
    if (this.editing) this.commitText()
    if (o.lk) return this.toast('ロック中は編集できません')
    this.editing = { id: o.id, before: before || this.snapshot(), isNew: !!before }
    this.setSel([o.id])
    const ta = this.editEl
    ta.value = o.tx || ''
    ta.hidden = false
    this.menuEl.hidden = true
    this.layoutEdit()
    this.redraw()
    ta.focus({ preventScroll: true })
    ta.setSelectionRange(ta.value.length, ta.value.length)
  }
  layoutEdit() {
    const ed = this.editing
    if (!ed) return
    const o = this.idx.get(ed.id)
    if (!o) return
    const ta = this.editEl, s = this.cam.scale
    const isShape = o.k === 'r'
    const a = isShape ? textArea(o) : { x: o.x + 4, y: o.y + 2, w: o.w - 8, h: o.h - 4 }
    const p = this.toScreen(a.x, a.y)
    const fs = o.fs * s
    Object.assign(ta.style, {
      left: p.x + 'px', top: p.y + 'px', width: Math.max(40, a.w * s) + 'px',
      fontSize: fs + 'px', lineHeight: '1.35', color: isShape ? o.tc : o.c,
      textAlign: isShape && o.t !== 'frame' ? 'center' : 'left',
    })
    ta.style.height = 'auto'
    const need = Math.max(fs * 1.35 + 4, ta.scrollHeight)
    const boxH = isShape && o.t !== 'frame' ? Math.max(a.h * s, need) : need
    ta.style.height = boxH + 'px'
    // 図形の中は上下の真ん中に寄せる
    if (isShape && o.t !== 'frame') {
      const lines = Math.max(1, Math.round((ta.scrollHeight - 2) / (fs * 1.35)))
      ta.style.paddingTop = Math.max(0, (a.h * s - lines * fs * 1.35) / 2) + 'px'
    } else ta.style.paddingTop = '0px'
  }
  commitText() {
    const ed = this.editing
    if (!ed) return
    this.editing = null
    const ta = this.editEl
    const text = ta.value.replace(/\s+$/, '')
    ta.hidden = true
    const o = this.idx.get(ed.id)
    if (o) {
      if (o.k === 't' && !text) {
        // 空の文字ボックスは残さない
        if (!ed.isNew) this.pushUndo(ed.before)
        this.data.objs = this.data.objs.filter((x) => x.id !== o.id)
        this.sel.delete(o.id)
        this.changed()
      } else if (text !== (o.tx || '') || ed.isNew) {
        this.pushUndo(ed.before)
        let n = { ...o, tx: text }
        if (n.k === 't') n = fitTextHeight(n)
        this.replaceItems(new Map([[n.id, n]]))
        this.changed()
      }
    }
    this.redraw()
    this.stage.focus?.()
  }
  newText(wp) {
    const before = this.snapshot()
    const fs = 18
    let x = wp.x - 4, y = wp.y - fs * 0.7
    if (this.snap) { x = Math.round(x / GRID) * GRID; y = Math.round(y / GRID) * GRID }
    const o = { k: 't', id: uid(), x, y, w: 240, h: fs * 1.35 + 4, tx: '', c: this.colors.pen, fs }
    this.data.objs.push(o)
    this.reindex()
    this.editText(o, before)
  }

  // ---- 座標
  resize() {
    const r = this.stage.getBoundingClientRect()
    this.w = r.width
    this.h = r.height
    this.left = r.left
    this.top = r.top
    this.dpr = window.devicePixelRatio || 1
    for (const c of [this.main, this.live, this.ui]) {
      c.width = Math.max(1, Math.round(this.w * this.dpr))
      c.height = Math.max(1, Math.round(this.h * this.dpr))
    }
    this.colorsCss = sceneColors(this.root)
    this.selColor = cssVar(this.root, '--ib-sel', '#2f7bd0')
    this.redraw()
  }
  toWorld(cx, cy) {
    const { x, y, scale } = this.cam
    return { x: x + (cx - this.left) / scale, y: y + (cy - this.top) / scale }
  }
  toScreen(wx, wy) {
    const { x, y, scale } = this.cam
    return { x: (wx - x) * scale, y: (wy - y) * scale }
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
    const top = this.palEl && !this.palEl.hidden ? 56 : 0
    let s = clamp(Math.min((this.w - m * 2) / a.w, (this.h - m * 2 - top) / a.h), MIN_SCALE, MAX_SCALE)
    // 何も描いていない無限キャンバスは、拡大せず等倍で始める
    if (d.mode === 'infinite' && !d.view && !d.strokes.length && !d.objs.length) s = Math.min(s, 1)
    this.setCam(a.x - (this.w / s - a.w) / 2, a.y - ((this.h - top) / s - a.h) / 2 - top / s, s)
  }
  snapV(v) { return this.snap ? Math.round(v / GRID) * GRID : v }

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
    drawScene(ctx, this.data, this.screenRect(), this.cam.scale, this.colorsCss, { snapDots: this.snap && !this.readOnly, hideText: this.editing?.id })
    this.drawLive()
    this.drawOverlay()
    this.drawUi()
    if (this.editing) this.layoutEdit()
  }
  /** 書き足した1本だけを本体の絵に重ねる（図形や線があるときは、重なり順を守るため全体を描き直す） */
  drawOne(s) {
    if (this.data.objs.length) return this.redraw()
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
  /** 選んだものの枠・つまみ・つなぐ点・投げなわなど（画面座標で描く） */
  drawUi() {
    const ctx = this.ui.getContext('2d')
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, this.ui.width, this.ui.height)
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0)
    const accent = this.selColor
    const a = this.action
    ctx.lineWidth = 1.2
    // 中に入っているグループの範囲
    if (this.gpath.length) {
      const L = this.gpath.length
      const ids = [...this.idx.values()].filter((x) => (x.g || [])[L - 1] === this.gpath[L - 1]).map((x) => x.id)
      const b = this.selBox(ids)
      if (b) {
        const p0 = this.toScreen(b.x0, b.y0), p1 = this.toScreen(b.x1, b.y1)
        ctx.strokeStyle = 'rgba(120,120,120,.7)'
        ctx.setLineDash([3, 4])
        ctx.strokeRect(p0.x - 8, p0.y - 8, p1.x - p0.x + 16, p1.y - p0.y + 16)
      }
    }
    this.handles = []
    if (this.sel.size && !(a && ['move', 'resize'].includes(a.type) && a.started)) {
      const items = this.selItems()
      const b = this.selBox()
      if (b) {
        const p0 = this.toScreen(b.x0, b.y0), p1 = this.toScreen(b.x1, b.y1)
        const locked = items.some((it) => it.lk)
        const single = items.length === 1 ? items[0] : null
        ctx.strokeStyle = accent
        ctx.setLineDash(single?.k === 'l' ? [] : [5, 4])
        if (single?.k !== 'l') ctx.strokeRect(p0.x - 4, p0.y - 4, p1.x - p0.x + 8, p1.y - p0.y + 8)
        ctx.setLineDash([])
        if (this.isGroupSel() || locked) {
          const label = (this.isGroupSel() ? 'グループ' : '') + (locked ? (this.isGroupSel() ? '・' : '') + 'ロック中' : '')
          ctx.font = `11px ${FONT}`
          const tw = ctx.measureText(label).width + 12
          ctx.fillStyle = accent
          ctx.fillRect(p0.x - 4, p0.y - 22, tw, 17)
          ctx.fillStyle = '#fff'
          ctx.textBaseline = 'middle'
          ctx.fillText(label, p0.x + 2, p0.y - 13)
        }
        if (!locked && !this.readOnly) {
          if (single?.k === 'l') {
            const g = lineGeom(single, this.idx)
            for (const [which, pt] of [['a', g.A], ['b', g.B]]) {
              const sp = this.toScreen(pt.x, pt.y)
              this.handles.push({ type: 'end', which, x: sp.x, y: sp.y })
              ctx.beginPath(); ctx.arc(sp.x, sp.y, 6, 0, Math.PI * 2)
              ctx.fillStyle = '#fff'; ctx.fill(); ctx.strokeStyle = accent; ctx.stroke()
            }
          } else {
            const corners = [[p0.x - 4, p0.y - 4], [p1.x + 4, p0.y - 4], [p1.x + 4, p1.y + 4], [p0.x - 4, p1.y + 4]]
            corners.forEach(([x, y], i) => {
              this.handles.push({ type: 'resize', corner: i, x, y })
              ctx.fillStyle = '#fff'; ctx.fillRect(x - 5, y - 5, 10, 10)
              ctx.strokeStyle = accent; ctx.strokeRect(x - 5, y - 5, 10, 10)
            })
            if (single?.k === 'r') {
              for (const sd of SIDES) {
                const q = sidePoint(single, sd), sp = this.toScreen(q.x, q.y)
                const off = 14
                const x = sp.x + SIDE_DIR[sd][0] * off, y = sp.y + SIDE_DIR[sd][1] * off
                this.handles.push({ type: 'conn', sd, x, y })
                ctx.beginPath(); ctx.arc(x, y, 6, 0, Math.PI * 2)
                ctx.fillStyle = accent; ctx.globalAlpha = 0.9; ctx.fill(); ctx.globalAlpha = 1
                ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5
                ctx.beginPath(); ctx.moveTo(x - 3, y); ctx.lineTo(x + 3, y); ctx.moveTo(x, y - 3); ctx.lineTo(x, y + 3); ctx.stroke()
                ctx.lineWidth = 1.2
              }
            }
          }
        }
      }
    }
    if (!a) return
    if (a.type === 'lasso' && a.pts.length > 1) {
      ctx.beginPath()
      a.pts.forEach((p, i) => { const s = this.toScreen(p.x, p.y); i ? ctx.lineTo(s.x, s.y) : ctx.moveTo(s.x, s.y) })
      ctx.closePath()
      ctx.fillStyle = 'rgba(47,123,208,.08)'; ctx.fill()
      ctx.strokeStyle = accent; ctx.setLineDash([5, 4]); ctx.stroke(); ctx.setLineDash([])
    }
    // 図形・線の下書きは世界座標で描く
    if (a.type === 'create' || a.type === 'connect' || a.type === 'endpoint') {
      ctx.save()
      const k = this.dpr * this.cam.scale
      ctx.setTransform(k, 0, 0, k, -this.cam.x * k, -this.cam.y * k)
      if (a.type === 'create' && a.rect) {
        ctx.globalAlpha = 0.6
        drawShape(ctx, { k: 'r', t: this.shapeKind, ...a.rect, c: accent, f: '', lw: 1.5 / this.cam.scale, d: 1, tx: '' }, this.colorsCss)
      }
      if (a.type === 'connect' || a.type === 'endpoint') {
        if (a.target) {
          ctx.strokeStyle = accent
          ctx.lineWidth = 2 / this.cam.scale
          ctx.setLineDash([])
          shapePath(ctx, a.target)
          ctx.stroke()
        }
        if (a.preview) { ctx.globalAlpha = 0.85; drawLine(ctx, a.preview, this.idx) }
      }
      ctx.restore()
    }
  }

  // ---- 入力
  /** 'draw'：道具として使う　'touch'：指（物の上なら選ぶ、何もない所なら移動）　'pan'：移動 */
  inputKind(e) {
    if (this.readOnly) return 'pan'
    const objTool = ['select', 'shape', 'text'].includes(this.tool)
    if (this.tool === 'hand') return 'pan'
    if (e.pointerType === 'mouse') return e.button === 0 && !this.space ? 'draw' : 'pan'
    if (e.pointerType === 'pen') return 'draw'
    if (objTool) return this.fingerDraw ? 'draw' : 'touch'
    return this.fingerDraw ? 'draw' : 'pan'
  }

  down(e) {
    if (e.target === this.editEl) return
    if (e.pointerType === 'pen') {
      this.lastPen = performance.now()
      // 手のひらが先に触れて移動や指描きが始まっていたら、Pencil を優先する
      if (this.action && this.action.pointerType === 'touch') {
        if (this.action.type === 'draw') this.cancelStroke()
        else if (this.action.before && this.action.started) this.restore(this.undoStack.pop())
        this.action = null
        this.stage.classList.remove('ib-grabbing')
      }
      for (const [id, p] of this.pointers) if (p.type === 'touch') this.pointers.delete(id)
      this.openPal()
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
    if (this.editing) { this.commitText(); return }
    if (!this.menuEl.hidden && this.pasteAt) { this.pasteAt = null; this.menuEl.hidden = true }
    try { this.stage.setPointerCapture(e.pointerId) } catch { /* 合成イベントなど */ }
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY, type: e.pointerType })
    const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')

    if (e.pointerType === 'touch' && touches.length >= 2) {
      // 2本目の指が来たら、指でしていた操作をやめて拡大縮小に切り替える
      const a = this.action
      if (a && a.pointerType === 'touch' && a.type !== 'pinch') {
        if (a.type === 'draw') this.cancelStroke()
        else if (a.before && a.started) this.restore(this.undoStack.pop())
        this.action = null
      }
      if (!this.action) this.startPinch()
      else if (this.action.type === 'pinch') this.action.max = Math.max(this.action.max, touches.length)
      return
    }
    if (this.action) return // Pencil で描いている最中の指などは無視する
    this.menuEl.hidden = true

    const kind = this.inputKind(e)
    if (kind === 'pan') return this.startPan(e)
    const t = this.tool
    if (t === 'pen' || t === 'marker') this.startStroke(e)
    else if (t === 'eraser') this.startErase(e)
    else if (t === 'area') this.startArea(e)
    else this.startObj(e, kind === 'touch')
  }

  move(e) {
    if (e.target === this.editEl && !this.action) return
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
    else this.moveObj(e)
  }

  up(e, cancelled) {
    if (e.pointerType === 'pen') { this.lastPen = performance.now(); this.schedulePalClose() }
    this.pointers.delete(e.pointerId)
    this.stage.classList.remove('ib-grabbing')
    const a = this.action
    if (!a) return
    if (a.type === 'pinch') {
      const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')
      if (touches.length < 2) {
        this.action = null
        // 2本指で軽くタップ：元に戻す　3本指：やり直す
        if (!a.moved && performance.now() - a.t0 < 320 && !this.readOnly) {
          if (a.max >= 3) { this.redo(); this.toast('やり直しました') } else { this.undo(); this.toast('元に戻しました') }
        }
      }
      return
    }
    if (a.pointerId !== e.pointerId) return
    if (a.type === 'draw') cancelled ? this.cancelStroke() : this.finishStroke()
    else if (a.type === 'erase') this.finishErase()
    else if (a.type === 'area') this.finishArea()
    else if (a.type !== 'pan') this.upObj(e, cancelled)
    this.action = null
    this.drawUi()
    this.showMenu()
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
    const s = { tool, color: this.colors[tool], width: WIDTHS[tool][this.widthIdx[tool]], pts: [], id: uid() }
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
    this.reindex()
    this.reparent([s.id])
    this.drawOne(this.idx.get(s.id))
    this.changed()
  }
  cancelStroke() {
    this.action = null
    this.drawLive()
  }

  // 消しゴム（手書きの線を1本ごと消す。図形は選んで消す）
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
        if (!s.lk && hitStroke(s, from, pt, r)) removed = true
        else keep.push(s)
      }
      if (removed) this.data.strokes = keep
    }
    if (removed) {
      a.hit = true
      this.reindex()
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

  // ---- 図形・線・文字・選択
  hitHandle(sp) {
    let best = null, bd = 18 * 18
    for (const h of this.handles || []) {
      const d = (h.x - sp.x) ** 2 + (h.y - sp.y) ** 2
      if (d <= bd) { bd = d; best = h }
    }
    return best
  }
  hitItem(wp) {
    const tol = 9 / this.cam.scale
    const objs = this.data.objs, idx = this.idx
    for (let i = objs.length - 1; i >= 0; i--) {
      const o = objs[i]
      if (o.k === 't' && wp.x >= o.x - tol && wp.x <= o.x + o.w + tol && wp.y >= o.y - tol && wp.y <= o.y + o.h + tol) return o
    }
    for (let i = objs.length - 1; i >= 0; i--) {
      const o = objs[i]
      if (o.k === 'l' && distToPolyline(wp, lineGeom(o, idx).pts) <= tol + o.lw) return o
    }
    for (let i = this.data.strokes.length - 1; i >= 0; i--) {
      const s = this.data.strokes[i]
      if (hitStroke(s, wp, wp, tol * 0.8)) return s
    }
    for (let i = objs.length - 1; i >= 0; i--) {
      const o = objs[i]
      if (o.k === 'r' && o.t !== 'frame' && insideShape(o, wp, tol)) return o
    }
    for (let i = objs.length - 1; i >= 0; i--) {
      const o = objs[i]
      if (o.k !== 'r' || o.t !== 'frame') continue
      const inOuter = wp.x >= o.x - tol && wp.x <= o.x + o.w + tol && wp.y >= o.y - tol && wp.y <= o.y + o.h + tol
      const inInner = wp.x > o.x + tol && wp.x < o.x + o.w - tol && wp.y > o.y + 28 && wp.y < o.y + o.h - tol
      if (inOuter && !inInner) return o
    }
    return null
  }
  /** 線をつなぐ先の図形 */
  attachTarget(wp, exclude) {
    const pad = 12 / this.cam.scale
    const objs = this.data.objs
    for (const pass of [false, true]) {
      for (let i = objs.length - 1; i >= 0; i--) {
        const o = objs[i]
        if (o.k !== 'r' || o.id === exclude || (o.t === 'frame') !== pass) continue
        if (wp.x >= o.x - pad && wp.x <= o.x + o.w + pad && wp.y >= o.y - pad && wp.y <= o.y + o.h + pad) {
          if (o.t === 'frame' && wp.x > o.x + pad && wp.x < o.x + o.w - pad && wp.y > o.y + pad && wp.y < o.y + o.h - pad) continue
          return o
        }
      }
    }
    return null
  }
  newLine(a, b) {
    return { k: 'l', id: uid(), a, b, sh: this.lineStyle.sh, hd: this.lineStyle.hd, d: this.lineStyle.d ? 1 : 0, c: this.colors.pen, lw: LINE_WIDTHS[this.lwIdx] }
  }

  startObj(e, touch) {
    const sp = { x: e.clientX - this.left, y: e.clientY - this.top }
    const wp = this.toWorld(e.clientX, e.clientY)
    const base = { pointerId: e.pointerId, pointerType: e.pointerType, sp, wp, t0: performance.now() }
    const h = this.tool !== 'text' && this.hitHandle(sp)
    if (h) {
      const items = this.selItems()
      if (h.type === 'conn') {
        const src = items[0]
        this.action = { ...base, type: 'connect', from: { s: src.id, sd: h.sd, ...sidePoint(src, h.sd) }, srcId: src.id }
        return
      }
      if (h.type === 'end') {
        const line = items[0]
        this.action = { ...base, type: 'endpoint', id: line.id, which: h.which, orig: line, before: this.snapshot() }
        return
      }
      const ids = this.expandIds(this.sel)
      const box = this.selBox()
      const single = items.length === 1 && items[0].k ? items[0] : null
      this.action = { ...base, type: 'resize', corner: h.corner, ids, box, single, uniform: [...ids].some((id) => isStroke(this.idx.get(id))) || (items.length > 1 && items.some((it) => it.k === 't')), orig: new Map([...ids].map((id) => [id, this.idx.get(id)])), before: this.snapshot() }
      return
    }
    const hit = this.hitItem(wp)
    // 線を引く設定のときは、図形の上から始めても線を引く（図形は動かさない）
    const from = this.tool === 'shape' && this.shapeKind === 'line' && this.parentOf(hit)
    if (from && from.k === 'r') {
      const sd = nearestSide(from, wp)
      this.action = { ...base, type: 'connect', from: { s: from.id, sd, ...sidePoint(from, sd) }, srcId: from.id }
      return
    }
    if (hit) {
      if (this.tool === 'text' && (hit.k === 'r' || hit.k === 't')) { this.action = { ...base, type: 'press', hit, ids: [hit.id], textTap: true }; return }
      const baseIt = this.parentOf(hit)
      const already = this.sel.has(baseIt.id)
      const ids = already ? [...this.sel] : this.resolveIds(hit)
      if (!already) { this.sel = new Set(ids); this.drawUi() }
      this.action = { ...base, type: 'press', hit, ids, already }
      return
    }
    // 何もない所
    if (touch) { this.action = { ...base, type: 'pending', cx: this.cam.x, cy: this.cam.y, sx: e.clientX, sy: e.clientY }; return }
    if (this.tool === 'select') { this.action = { ...base, type: 'lasso', pts: [wp] }; return }
    if (this.tool === 'shape') {
      if (this.shapeKind === 'line') { this.action = { ...base, type: 'connect', from: { x: this.snapV(wp.x), y: this.snapV(wp.y) } }; return }
      this.action = { ...base, type: 'create', p0: wp, rect: null }
      return
    }
    this.action = { ...base, type: 'press', hit: null, ids: [] }
  }

  moveObj(e) {
    const a = this.action
    const sp = { x: e.clientX - this.left, y: e.clientY - this.top }
    const wp = this.toWorld(e.clientX, e.clientY)
    const far = Math.hypot(sp.x - a.sp.x, sp.y - a.sp.y) > (a.pointerType === 'touch' ? 8 : 4)
    if (a.type === 'pending') {
      if (far) { this.action = { type: 'pan', pointerId: a.pointerId, pointerType: a.pointerType, sx: a.sx, sy: a.sy, cx: a.cx, cy: a.cy }; this.clearSel(); this.stage.classList.add('ib-grabbing') }
      return
    }
    if (a.type === 'press') {
      if (!far || !a.hit || a.textTap) return
      const items = a.ids.map((id) => this.idx.get(id)).filter(Boolean)
      if (items.some((it) => it.lk)) { if (!a.warned) { a.warned = true; this.toast('ロック中は動かせません') } return }
      const ids = this.expandIds(new Set(a.ids), true)
      const shapes = [...ids].map((id) => this.idx.get(id)).filter((it) => it && it.k && it.k !== 'l')
      Object.assign(a, { type: 'move', moveIds: ids, orig: new Map([...ids].map((id) => [id, this.idx.get(id)])), box: this.selBox(shapes.length ? shapes.map((s) => s.id) : ids), before: this.snapshot() })
    }
    if (a.type === 'move') {
      let dx = wp.x - a.wp.x, dy = wp.y - a.wp.y
      if (this.snap && a.box) { dx = this.snapV(a.box.x0 + dx) - a.box.x0; dy = this.snapV(a.box.y0 + dy) - a.box.y0 }
      if (!a.started) { if (!dx && !dy) return; a.started = true; this.pushUndo(a.before) }
      const f = translate(dx, dy)
      const map = new Map()
      for (const [id, it] of a.orig) map.set(id, mapItem(it, f, 1, this.idx))
      this.replaceItems(map)
      this.requestRedraw()
      return
    }
    if (a.type === 'resize') {
      const b = a.box
      const corners = [[b.x0, b.y0], [b.x1, b.y0], [b.x1, b.y1], [b.x0, b.y1]]
      const [ax, ay] = corners[(a.corner + 2) % 4], [cx, cy] = corners[a.corner]
      let nx = this.snapV(wp.x), ny = this.snapV(wp.y)
      const min = 16
      if (Math.abs(nx - ax) < min) nx = ax + Math.sign(cx - ax) * min
      if (Math.abs(ny - ay) < min) ny = ay + Math.sign(cy - ay) * min
      let sx = (nx - ax) / (cx - ax || 1), sy = (ny - ay) / (cy - ay || 1)
      if (sx <= 0) sx = min / Math.abs(cx - ax || 1)
      if (sy <= 0) sy = min / Math.abs(cy - ay || 1)
      if (a.uniform || e.shiftKey) sx = sy = Math.max(sx, sy)
      if (!a.started) { a.started = true; this.pushUndo(a.before) }
      const f = (x, y) => ({ x: ax + (x - ax) * sx, y: ay + (y - ay) * sy })
      const textOnly = a.single?.k === 't'
      const k = a.single ? 1 : Math.sqrt(sx * sy)
      const map = new Map()
      for (const [id, it] of a.orig) {
        let n = mapItem(it, f, textOnly ? 1 : k, this.idx)
        if (n.k === 't') n = fitTextHeight(n)
        map.set(id, n)
      }
      this.replaceItems(map)
      this.requestRedraw()
      return
    }
    if (a.type === 'lasso') {
      const last = a.pts[a.pts.length - 1]
      if (Math.hypot(wp.x - last.x, wp.y - last.y) * this.cam.scale > 3) a.pts.push(wp)
      this.drawUi()
      return
    }
    if (a.type === 'create') {
      const x0 = this.snapV(Math.min(a.p0.x, wp.x)), y0 = this.snapV(Math.min(a.p0.y, wp.y))
      const x1 = this.snapV(Math.max(a.p0.x, wp.x)), y1 = this.snapV(Math.max(a.p0.y, wp.y))
      a.rect = far ? { x: x0, y: y0, w: Math.max(8, x1 - x0), h: Math.max(8, y1 - y0) } : null
      this.drawUi()
      return
    }
    if (a.type === 'connect' || a.type === 'endpoint') {
      if (!far && !a.started) return
      a.started = true
      const exclude = a.type === 'connect' ? a.srcId : null
      a.target = this.attachTarget(wp, exclude)
      const end = a.target ? { s: a.target.id, sd: nearestSide(a.target, wp), ...sidePoint(a.target, nearestSide(a.target, wp)) } : { x: this.snapV(wp.x), y: this.snapV(wp.y) }
      if (a.type === 'connect') a.preview = { ...this.newLine(a.from, end), id: '_preview' }
      else {
        a.preview = { ...a.orig, [a.which]: end }
        a.end = end
      }
      a.endPt = end
      this.drawUi()
    }
  }

  upObj(e, cancelled) {
    const a = this.action
    const now = performance.now()
    if (cancelled) {
      if (a.started && a.before) this.restore(this.undoStack.pop())
      return
    }
    if (a.type === 'move' || a.type === 'resize') {
      if (a.started) {
        this.reparent([...a.orig.keys()])
        this.changed()
        this.redraw()
      }
      return
    }
    if (a.type === 'pending' || (a.type === 'press' && !a.hit)) return this.tapEmpty(a)
    if (a.type === 'press') return this.tapItem(a, now)
    if (a.type === 'lasso') {
      const bb = unionBox(a.pts.map((p) => ({ x0: p.x, y0: p.y, x1: p.x, y1: p.y })))
      if (!bb || (bb.x1 - bb.x0) * this.cam.scale < 8 && (bb.y1 - bb.y0) * this.cam.scale < 8) return this.tapEmpty(a)
      return this.lassoSelect(a.pts)
    }
    if (a.type === 'create') return this.createShape(a)
    if (a.type === 'connect') {
      if (!a.started || !a.endPt) return
      const end = a.endPt
      const len = Math.hypot((end.x ?? 0) - a.from.x, (end.y ?? 0) - a.from.y)
      if (!end.s && len * this.cam.scale < 12) return
      const line = this.newLine(a.from, end)
      this.pushUndo()
      this.data.objs.push(line)
      this.changed()
      this.setSel([line.id])
      this.redraw()
      return
    }
    if (a.type === 'endpoint') {
      if (!a.started || !a.end) return
      this.pushUndo(a.before)
      this.replaceItems(new Map([[a.id, { ...a.orig, [a.which]: a.end }]]))
      this.changed()
      this.redraw()
    }
  }
  tapEmpty(a) {
    const wp = a.wp
    if (this.tool === 'text') return this.newText(wp)
    if (this.tool === 'shape' && this.shapeKind !== 'line') {
      if (this.sel.size) return this.clearSel()
      return this.createShape({ p0: wp, rect: null })
    }
    if (this.tool === 'shape' && this.shapeKind === 'line') {
      if (this.sel.size) return this.clearSel()
      return this.toast('ドラッグで線を引きます（図形の上で離すとつながります）')
    }
    if (this.sel.size || this.gpath.length) return this.clearSel()
    if (clip) this.showPasteMenu(a.sp, wp)
  }
  tapItem(a, now) {
    const hit = a.hit
    const base = this.parentOf(hit)
    if (a.textTap) return this.editText(base)
    const dbl = this.lastTap && this.lastTap.id === base.id && now - this.lastTap.t < 420
    this.lastTap = { id: base.id, t: now }
    if (dbl) {
      this.lastTap = null
      const L = this.gpath.length
      if ((base.g || []).length > L && startsWith(base.g, this.gpath)) {
        // グループの中へ1段入る
        this.gpath = base.g.slice(0, L + 1)
        this.setSel(this.resolveIds(base))
        this.toast('グループの中を選んでいます（何もない所をタップで抜けます）')
        return this.redraw()
      }
      if (base.k === 'r' || base.k === 't') return this.editText(base)
      return
    }
    if (a.already && !this.isGroupSel() && this.sel.size > 1) this.setSel(this.resolveIds(hit))
  }
  lassoSelect(poly) {
    const found = new Set()
    const L = this.gpath.length
    for (const it of this.idx.values()) {
      if (it.lk) continue
      if (!startsWith(it.g, this.gpath)) continue
      if (isStroke(it)) {
        const p = it.pts
        let n = 0, inside = 0
        const step = Math.max(3, Math.floor(p.length / 3 / 40) * 3)
        for (let i = 0; i < p.length; i += step) { n++; if (inPoly(p[i], p[i + 1], poly)) inside++ }
        if (n && inside / n >= 0.5) found.add(it.id)
      } else {
        const b = itemBox(it, this.idx)
        if (inPoly((b.x0 + b.x1) / 2, (b.y0 + b.y1) / 2, poly)) found.add(it.id)
      }
    }
    const ids = new Set()
    for (const id of found) {
      const it = this.idx.get(id)
      if (isStroke(it) && it.p && found.has(it.p)) continue
      const g = it.g || []
      if (g.length > L) for (const x of this.idx.values()) { if ((x.g || [])[L] === g[L]) ids.add(x.id) }
      else ids.add(id)
    }
    this.setSel([...ids])
    this.redraw()
  }
  createShape(a) {
    const kind = this.shapeKind
    const def = SHAPES[kind]
    let r = a.rect
    if (!r) {
      const x = this.snapV(a.p0.x - def.w / 2), y = this.snapV(a.p0.y - def.h / 2)
      r = { x, y, w: def.w, h: def.h }
    }
    const o = {
      k: 'r', id: uid(), t: kind, ...r,
      c: this.colors.pen, f: kind === 'frame' ? '' : 'w', lw: LINE_WIDTHS[this.lwIdx], d: 0,
      tx: '', tc: '#1f1f1f', fs: kind === 'frame' ? 14 : 15,
    }
    this.pushUndo()
    if (kind === 'frame') this.data.objs.unshift(o)
    else this.data.objs.push(o)
    this.changed()
    this.gpath = []
    this.setSel([o.id])
    this.redraw()
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
      t0: performance.now(),
      max: 2,
      moved: false,
    }
  }
  movePinch() {
    const touches = [...this.pointers.values()].filter((p) => p.type === 'touch')
    if (touches.length < 2) return
    const [a, b] = touches
    const p = this.action
    const mid = { x: (a.x + b.x) / 2 - this.left, y: (a.y + b.y) / 2 - this.top }
    const d = Math.hypot(a.x - b.x, a.y - b.y)
    if (!p.moved && (Math.abs(d - p.d0) > 10 || Math.hypot(mid.x - p.mid0.x, mid.y - p.mid0.y) > 10)) p.moved = true
    if (!p.moved) return
    const s = clamp(p.cam0.scale * (d / p.d0), MIN_SCALE, MAX_SCALE)
    const wx = p.cam0.x + p.mid0.x / p.cam0.scale, wy = p.cam0.y + p.mid0.y / p.cam0.scale
    this.setCam(wx - mid.x / s, wy - mid.y / s, s)
  }
  wheel(e) {
    e.preventDefault()
    if (e.ctrlKey || e.metaKey) this.zoomAt(e.clientX - this.left, e.clientY - this.top, Math.exp(-e.deltaY * 0.01))
    else this.setCam(this.cam.x + e.deltaX / this.cam.scale, this.cam.y + e.deltaY / this.cam.scale, this.cam.scale)
    if (this.sel.size) this.showMenu()
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
    if (e.target === this.editEl) return
    if (e.target.closest?.('select, input')) {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.closePops() }
      return
    }
    e.stopPropagation()
    const mod = e.metaKey || e.ctrlKey, k = e.key.toLowerCase()
    if (e.key === 'Escape') {
      e.preventDefault()
      if (this.root.querySelector('.ib-pop:not([hidden])')) return this.closePops()
      if (this.sel.size || this.gpath.length) { this.clearSel(); return this.showMenu() }
      return this.close()
    }
    if (mod && k === 'z') { e.preventDefault(); return e.shiftKey ? this.redo() : this.undo() }
    if (mod && k === 'y') { e.preventDefault(); return this.redo() }
    if (mod && k === 's') { e.preventDefault(); return this.flush() }
    if (this.readOnly) {
      if (k === '0') return this.fit()
      return
    }
    const objTool = ['select', 'shape', 'text'].includes(this.tool)
    if (mod && k === 'a') { e.preventDefault(); if (!objTool) this.setTool('select'); this.gpath = []; this.setSel(this.data.objs.map((o) => o.id).concat(this.data.strokes.filter((s) => !s.p).map((s) => s.id)).filter((id) => !this.idx.get(id).lk)); return this.showMenu() }
    if (mod && k === 'c') { e.preventDefault(); return this.copySel(false) }
    if (mod && k === 'x') { e.preventDefault(); return this.copySel(true) }
    if (mod && k === 'v') { e.preventDefault(); if (!objTool) this.setTool('select'); this.pasteClip(null); return this.showMenu() }
    if (mod && k === 'd') { e.preventDefault(); this.duplicateSel(); return this.showMenu() }
    if (mod && k === 'g') { e.preventDefault(); e.shiftKey ? this.ungroupSel() : this.groupSel(); return this.showMenu() }
    if (mod) return
    if ((e.key === 'Delete' || e.key === 'Backspace') && this.sel.size) { e.preventDefault(); return this.deleteSel() }
    if (e.key === 'Enter' && this.sel.size === 1) { const it = this.selItems()[0]; if (it.k === 'r' || it.k === 't') { e.preventDefault(); return this.editText(it) } }
    if (e.key.startsWith('Arrow') && this.sel.size) {
      e.preventDefault()
      const items = [...this.expandIds(this.sel, true)].map((id) => this.idx.get(id))
      if (items.some((it) => it.lk)) return
      const st = this.snap ? GRID : e.shiftKey ? 10 : 1
      const dx = e.key === 'ArrowLeft' ? -st : e.key === 'ArrowRight' ? st : 0, dy = e.key === 'ArrowUp' ? -st : e.key === 'ArrowDown' ? st : 0
      this.updateItems(items, (it) => mapItem(it, translate(dx, dy), 1, this.idx))
      return this.showMenu()
    }
    if (e.code === 'Space') { e.preventDefault(); this.space = true; return this.updateCursor() }
    if (k === '0') return this.fit()
    if (k === '+' || k === '=' || k === ';') return this.zoomAt(this.w / 2, this.h / 2, 1.25)
    if (k === '-') return this.zoomAt(this.w / 2, this.h / 2, 0.8)
    const tools = { p: 'pen', m: 'marker', e: 'eraser', v: 'select', s: 'shape', t: 'text', r: 'area', h: 'hand' }
    if (tools[k]) this.setTool(tools[k])
    if (k === 'g') this.command('snap')
  }
}

function insideShape(o, p, tol) {
  const hw = o.w / 2 + tol, hh = o.h / 2 + tol
  const dx = Math.abs(p.x - (o.x + o.w / 2)), dy = Math.abs(p.y - (o.y + o.h / 2))
  if (dx > hw || dy > hh) return false
  if (o.t === 'diamond') return dx / hw + dy / hh <= 1
  if (o.t === 'ellipse') return (dx / hw) ** 2 + (dy / hh) ** 2 <= 1
  return true
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
