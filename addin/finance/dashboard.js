/* ============================================================
 * 財務ダッシュボード dashboard.js
 * ------------------------------------------------------------
 * 計算とレンダリングのみを担当する。Excelの読み書きは app.js。
 * DATA は app.js が setData() で差し込む（Excel実データ／デモ共通）。
 *
 * DATA の形:
 *   { months:["2025/10",…12件],
 *     pl:[{name, annualPlan, v:[12]}], bs:[{name, v:[12]}],
 *     fy36:[{name, amount}], assum:{salesPlan,sgaPlan,cogsRate},
 *     bonusPeople:[{name,amt}], mnashi:0.5, termLabel:"第37期" }
 * ============================================================ */

/* ---------- 状態 ---------- */
let DATA = null;
let M = [], N = 12, plMap = {}, SGA = [], B = {}, bsIdx = [];
let PLAN_SALES = 0, PLAN_OP = 0, PLAN_SGA = 0;
let UNMAPPED = [];

const INK = '#1b1b1b', SHU = '#b03a24', SHU2 = '#d98b76', GREY = '#b8b3aa',
      WASH = '#e5e1d9', MID = '#7c7770', LIGHT = '#cfc9be';

const CAT_DEF = [
  ['人件費', ['役員報酬', '給与手当', '給料手当', '賞与', '法定福利費', '福利厚生費']],
  ['外注費', ['業務委託費', '外注費']],
  ['地代家賃', ['地代家賃']],
  ['販促費', ['広告宣伝費', '接待交際費']],
  ['一般管理費', ['荷造発送費', '荷造運賃', '会議費', '旅費交通費', '通信費', '消耗品費',
    '水道光熱費', '新聞図書費', '諸会費', '保険料', '支払保険料', '事務用品費', '雑費', '車両費', '修繕費']],
  ['税金・専門家費用', ['支払手数料', '租税公課', '支払報酬', '支払報酬料', '貸倒引当金繰入額']],
];
const CATS = CAT_DEF.map(x => x[0]);
const CATCOL = { '人件費': INK, '外注費': '#6b6660', '地代家賃': '#9a948b', '販促費': SHU, '一般管理費': LIGHT, '税金・専門家費用': SHU2 };
const CAT_LOOKUP = (() => { const m = {}; CAT_DEF.forEach(([c, ns]) => ns.forEach(n => m[n] = c)); return m; })();

/* ---------- 表示ヘルパー ---------- */
const yen = n => (n === null || n === undefined || isNaN(n)) ? '—' : Math.round(n).toLocaleString('ja-JP');
const syen = n => (n >= 0 ? '' : '△') + yen(Math.abs(n));
const pct = (n, d = 1) => (n === null || !isFinite(n)) ? '—' : (n * 100).toFixed(d) + '%';
const mio = n => { const v = n / 1000000; return (Math.abs(v) >= 10 ? v.toFixed(0) : v.toFixed(1)) + '百万'; };
const sum = a => a.reduce((s, x) => s + (x || 0), 0);
const monthLabel = m => (m.split('/')[1] || m) + '月';
const normName = s => String(s || '').replace(/[（(]旧[:：][^）)]*[）)]/g, '').replace(/\s/g, '').trim();
function catOf(name) {
  const n = normName(name);
  if (CAT_LOOKUP[n]) return CAT_LOOKUP[n];
  if (UNMAPPED.indexOf(name) < 0) UNMAPPED.push(name);
  return '一般管理費';
}

/* ============================================================
   データ準備
   ============================================================ */
const ZERO = name => ({ name: name, annualPlan: 0, v: new Array(12).fill(0) });
let K = {};

function setData(d) {
  DATA = d;
  UNMAPPED = [];
  M = DATA.months.slice();
  N = M.length;
  plMap = {}; DATA.pl.forEach(p => { plMap[p.name] = p; });

  const findPL = pred => DATA.pl.find(pred) || null;
  K.sales = findPL(p => normName(p.name) === '売上高') || ZERO('売上高');
  K.cogs = findPL(p => normName(p.name) === '売上原価') || ZERO('売上原価');
  K.gp = findPL(p => normName(p.name) === '売上総利益') || ZERO('売上総利益');
  K.sgaTot = findPL(p => /販管費/.test(p.name) && /合計/.test(p.name)) || ZERO('販管費 合計');
  K.op = findPL(p => normName(p.name) === '営業利益') || ZERO('営業利益');
  K.ord = findPL(p => normName(p.name) === '経常利益') || ZERO('経常利益');
  K.noIn = findPL(p => /^営業外収益/.test(p.name)) || ZERO('営業外収益');
  K.noEx = findPL(p => /^営業外費用/.test(p.name)) || ZERO('営業外費用');

  SGA = DATA.pl.filter(p => p.isSga);
  if (!SGA.length) {
    // isSga が無い場合は既知の費目名で拾う
    SGA = DATA.pl.filter(p => CAT_LOOKUP[normName(p.name)]);
  }

  PLAN_SALES = K.sales.annualPlan || (DATA.assum && DATA.assum.salesPlan) || 0;
  PLAN_SGA = K.sgaTot.annualPlan || 0;
  PLAN_OP = K.op.annualPlan || (PLAN_SALES - (K.cogs.annualPlan || 0) - PLAN_SGA);

  B = {}; DATA.bs.forEach(b => { B[b.name] = b.v; });
  const assets = B['資産 合計'] || B['資産合計'] || new Array(N).fill(0);
  bsIdx = []; for (let i = 0; i < N; i++) if (assets[i] && Math.abs(assets[i]) > 0) bsIdx.push(i);
  if (!bsIdx.length) bsIdx = [0];

  S.last = clampIdx(S.last === null ? defaultLastActual() : S.last);
}
const bsKey = (...names) => { for (const n of names) if (B[n]) return B[n]; return new Array(N).fill(0); };
function clampIdx(i) { return Math.max(0, Math.min(N - 1, i)); }
function defaultLastActual() { return bsIdx[bsIdx.length - 1]; }
function actualCogsRate(last) {
  let ss = 0, cc = 0;
  for (let i = 0; i <= last; i++) { ss += K.sales.v[i] || 0; cc += K.cogs.v[i] || 0; }
  return ss ? cc / ss : 0.10;
}

/* ============================================================
   チャート
   ============================================================ */
function niceTicks(min, max, count) {
  if (min === max) max = min + 1;
  const span = max - min, raw = span / count, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag; let step;
  if (norm <= 1) step = 1; else if (norm <= 2) step = 2; else if (norm <= 2.5) step = 2.5;
  else if (norm <= 5) step = 5; else step = 10;
  step *= mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const t = []; for (let v = lo; v <= hi + step * 1e-9; v += step) t.push(Math.round(v * 1e6) / 1e6);
  return t;
}
function drawChart(host, spec) {
  if (!host) return;
  host._spec = spec;
  const W = Math.max(260, Math.round(host.clientWidth || 900));
  const narrow = W < 560;
  const H = Math.round((spec.height || 320) * (narrow ? 0.9 : 1));
  const ml = spec.ml != null ? (narrow ? Math.min(spec.ml, 46) : spec.ml) : (narrow ? 40 : 86);
  const mr = narrow ? 8 : 18, mt = 14, mb = narrow ? 30 : 38;
  const fs = narrow ? 9.5 : 11.5;
  const pw = W - ml - mr, ph = H - mt - mb;
  const labels = narrow && spec.compactLabels !== false
    ? spec.labels.map(x => String(x).replace('月', '')) : spec.labels;
  const n = labels.length;
  const bars = spec.series.filter(s => s.type === 'bar'), lines = spec.series.filter(s => s.type === 'line');
  let lo = 0, hi = 0;
  if (spec.stacked) {
    for (let i = 0; i < n; i++) { let p = 0, m = 0; bars.forEach(s => { const v = s.values[i] || 0; v >= 0 ? p += v : m += v; }); hi = Math.max(hi, p); lo = Math.min(lo, m); }
  } else {
    bars.forEach(s => s.values.forEach(v => { if (v == null) return; hi = Math.max(hi, v); lo = Math.min(lo, v); }));
  }
  lines.forEach(s => s.values.forEach(v => { if (v == null) return; hi = Math.max(hi, v); lo = Math.min(lo, v); }));
  if (spec.zeroBase !== false) { hi = Math.max(hi, 0); lo = Math.min(lo, 0); }
  const ticks = niceTicks(lo, hi, spec.ticks || (narrow ? 4 : 5));
  const y0 = ticks[0], y1 = ticks[ticks.length - 1];
  const Y = v => mt + ph - ((v - y0) / (y1 - y0)) * ph;
  const band = pw / n, X = i => ml + band * i + band / 2;
  const fmtY = spec.yfmt === 'pct' ? (v => (v * 100).toFixed(0) + '%')
    : spec.yfmt === 'raw' ? (v => v.toFixed(1))
      : (v => v === 0 ? '0' : (narrow ? (v / 1000000).toFixed(1) : mio(v)));
  const fmtV = spec.vfmt === 'pct' ? (v => pct(v, 1))
    : spec.vfmt === 'raw' ? (v => v.toFixed(2) + 'ヶ月') : (v => syen(v));
  let s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="' + W + '" height="' + H + '" role="img">';
  ticks.forEach(t => {
    const y = Y(t), z = Math.abs(t) < 1e-9;
    s += '<line x1="' + ml + '" x2="' + (W - mr) + '" y1="' + y + '" y2="' + y + '" stroke="' + (z ? INK : '#eeebe5') + '"/>';
    s += '<text x="' + (ml - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="' + fs + '" fill="#918d86">' + fmtY(t) + '</text>';
  });
  const nb = bars.length, gw = band * 0.64, bw = spec.stacked ? gw : gw / Math.max(nb, 1);
  if (spec.stacked) {
    for (let i = 0; i < n; i++) {
      let accP = 0, accM = 0;
      bars.forEach(sr => {
        const v = sr.values[i] || 0; if (!v) return;
        const from = v >= 0 ? accP : accM, to = from + v;
        const yA = Y(Math.max(from, to)), yB = Y(Math.min(from, to));
        s += '<rect x="' + (X(i) - gw / 2) + '" y="' + yA + '" width="' + gw + '" height="' + Math.max(0, yB - yA) + '" fill="' + sr.color + '"/>';
        if (v >= 0) accP = to; else accM = to;
      });
    }
  } else {
    bars.forEach((sr, bi) => {
      for (let i = 0; i < n; i++) {
        const v = sr.values[i]; if (v == null) continue;
        const x = X(i) - gw / 2 + bi * bw, yA = Y(Math.max(v, 0)), yB = Y(Math.min(v, 0));
        const col = (sr.colors && sr.colors[i]) || sr.color;
        s += '<rect x="' + x + '" y="' + yA + '" width="' + Math.max(1, bw - 1.5) + '" height="' + Math.max(1, yB - yA) + '" fill="' + col + '"/>';
      }
    });
  }
  lines.forEach(sr => {
    let d = '', started = false;
    for (let i = 0; i < n; i++) { const v = sr.values[i]; if (v == null) { started = false; continue; } d += (started ? 'L' : 'M') + X(i) + ' ' + Y(v) + ' '; started = true; }
    s += '<path d="' + d + '" fill="none" stroke="' + sr.color + '" stroke-width="' + (sr.w || 2) + '" stroke-linejoin="round" stroke-linecap="round"' + (sr.dash ? ' stroke-dasharray="6 4"' : '') + '/>';
    if (sr.dots !== false && !narrow) for (let i = 0; i < n; i++) { const v = sr.values[i]; if (v == null) continue; s += '<circle cx="' + X(i) + '" cy="' + Y(v) + '" r="3" fill="#fff" stroke="' + sr.color + '" stroke-width="1.6"/>'; }
  });
  const skip = narrow && n > 8 ? 1 : 0;
  labels.forEach((L, i) => {
    if (skip && band < 22 && i % 2 === 1) return;
    s += '<text x="' + X(i) + '" y="' + (H - mb + 20) + '" text-anchor="middle" font-size="' + fs + '" fill="#5c5a56">' + L + '</text>';
  });
  for (let i = 0; i < n; i++) s += '<rect class="hb" data-i="' + i + '" x="' + (ml + band * i) + '" y="' + mt + '" width="' + band + '" height="' + ph + '" fill="transparent"/>';
  s += '</svg>';

  const legend = '<div class="legend">' + spec.series.filter(x => x.name).map(x =>
    '<span><i class="' + (x.type === 'line' ? 'ln' : '') + '" style="' + (x.type === 'line' ? 'border-color:' + x.color : 'background:' + x.color) + '"></i>' + x.name + '</span>').join('') + '</div>';
  const unit = spec.yfmt || 'yen';
  const unitNote = unit === 'yen' ? '<p class="note">縦軸の単位：百万円</p>' : '';
  host.innerHTML = (spec.legend === false ? '' : legend) + '<div class="chart">' + s + '<div class="tip"></div></div>'
    + (spec.note ? '<p class="note">' + spec.note + '</p>' : '') + (narrow ? unitNote : '');

  const box = host.querySelector('.chart'), tip = host.querySelector('.tip');
  box.addEventListener('mousemove', e => {
    const t = e.target;
    if (!t.classList || !t.classList.contains('hb')) { tip.style.opacity = 0; return; }
    const i = +t.dataset.i;
    let h = '<b>' + spec.labels[i] + '</b>';
    spec.series.forEach(sr => {
      if (!sr.name) return; const v = sr.values[i]; if (v == null) return;
      h += '<br><span class="sw" style="background:' + ((sr.colors && sr.colors[i]) || sr.color) + '"></span>' + sr.name + '　' + fmtV(v);
    });
    tip.innerHTML = h; tip.style.opacity = 1;
    const r = box.getBoundingClientRect(), x = e.clientX - r.left, y = e.clientY - r.top;
    tip.style.left = Math.max(0, Math.min(x + 12, box.clientWidth - tip.offsetWidth - 4)) + 'px';
    tip.style.top = Math.max(0, y - tip.offsetHeight - 10) + 'px';
  });
  box.addEventListener('mouseleave', () => tip.style.opacity = 0);
}

/* ============================================================
   設定（チューナー）
   ============================================================ */
const S = {
  last: null, method: 'sheet', cogsMode: 'method', salesAdj: 0, cogsRate: 10,
  sgaAdj: {}, base: 'fc37', growth: 0, season: 100, nyCogs: 10, nySgaAdj: {}
};
CATS.forEach(c => { S.sgaAdj[c] = 0; S.nySgaAdj[c] = 0; });

/* ============================================================
   計算
   ============================================================ */
function estimate(item, i) {
  if (S.method === 'plan') return (item.annualPlan || 0) / 12;
  if (S.method === 'ma3') {
    let a = 0, k = 0;
    for (let j = Math.max(0, S.last - 2); j <= S.last; j++) { a += item.v[j] || 0; k++; }
    return k ? a / k : 0;
  }
  return item.v[i] || 0;
}
function compute() {
  const isAct = i => i <= S.last;
  const sales = [], cogs = [], sgaItems = {}, sgaCat = {}, sgaTot = [], op = [], gp = [];
  CATS.forEach(c => sgaCat[c] = new Array(N).fill(0));
  for (let i = 0; i < N; i++) {
    const sv = isAct(i) ? (K.sales.v[i] || 0) : estimate(K.sales, i) * (1 + S.salesAdj / 100);
    sales.push(Math.round(sv));
    cogs.push(isAct(i) ? (K.cogs.v[i] || 0)
      : (S.cogsMode === 'rate' ? Math.round(sv * S.cogsRate / 100) : Math.round(estimate(K.cogs, i))));
  }
  SGA.forEach(it => {
    const cat = catOf(it.name), arr = [];
    for (let i = 0; i < N; i++) {
      const v = isAct(i) ? (it.v[i] || 0) : Math.round(estimate(it, i) * (1 + S.sgaAdj[cat] / 100));
      arr.push(v); sgaCat[cat][i] += v;
    }
    sgaItems[it.name] = arr;
  });
  for (let i = 0; i < N; i++) {
    let t = 0; CATS.forEach(c => t += sgaCat[c][i]);
    sgaTot.push(t); gp.push(sales[i] - cogs[i]); op.push(sales[i] - cogs[i] - t);
  }
  const nIn = K.noIn.v.map(v => v || 0), nEx = K.noEx.v.map(v => v || 0);
  const ord = op.map((v, i) => v + nIn[i] - nEx[i]);
  return {
    sales, cogs, gp, sgaItems, sgaCat, sgaTot, op, ord, isAct,
    tot: { sales: sum(sales), cogs: sum(cogs), gp: sum(gp), sga: sum(sgaTot), op: sum(op), ord: sum(ord) }
  };
}
function fy36Of(name) { const x = (DATA.fy36 || []).find(y => normName(y.name) === normName(name)); return x ? x.amount : 0; }
function computeNext(cur) {
  const b36 = fy36Of('売上高');
  const base = S.base === 'fy36' ? b36 : S.base === 'avg' ? (b36 + cur.tot.sales) / 2 : cur.tot.sales;
  const annual = base * (1 + S.growth / 100);
  const w = cur.tot.sales ? cur.sales.map(v => v / cur.tot.sales) : new Array(N).fill(1 / N);
  const a = S.season / 100, wt = w.map(x => a * x + (1 - a) / N);
  const sales = wt.map(x => Math.round(annual * x));
  const cogs = sales.map(v => Math.round(v * S.nyCogs / 100));
  const sgaCat = {}, sgaTot = new Array(N).fill(0);
  CATS.forEach(c => {
    const cy = cur.sgaCat[c], t = sum(cy), nt = t * (1 + S.nySgaAdj[c] / 100);
    const ww = t ? cy.map(v => v / t) : new Array(N).fill(1 / N);
    sgaCat[c] = ww.map(x => Math.round(nt * x));
    sgaCat[c].forEach((v, i) => sgaTot[i] += v);
  });
  const op = sales.map((v, i) => v - cogs[i] - sgaTot[i]);
  const fixed = sum(sgaTot), vr = S.nyCogs / 100;
  return {
    sales, cogs, sgaCat, sgaTot, op, base, annual,
    tot: { sales: sum(sales), cogs: sum(cogs), sga: fixed, op: sum(op) },
    bep: vr < 1 ? fixed / (1 - vr) : null
  };
}

/* ============================================================
   決算賞与
   ============================================================ */
const BN = {
  people: [], peopleDirty: false,
  socRate: 15, socWhen: 'this', mnashi: 50, kintou: 180000,
  empSoc: 15, empTax: 10.21, tgtRate: 3, resMonths: 3
};
const btTotal = () => BN.people.reduce((a, p) => a + (p.amt || 0), 0);

function corpTax(income, kintou) {
  if (!(income > 0)) return { houjin: 0, chihou: 0, jumin: 0, jigyo: 0, toku: 0, kintou: kintou, total: kintou };
  const h = Math.min(income, 8000000) * 0.15 + Math.max(0, income - 8000000) * 0.232;
  const c = h * 0.103, j = h * 0.07;
  const g = Math.min(income, 4000000) * 0.035
    + Math.min(Math.max(income - 4000000, 0), 4000000) * 0.053
    + Math.max(income - 8000000, 0) * 0.07;
  const t = g * 0.37;
  return { houjin: h, chihou: c, jumin: j, jigyo: g, toku: t, kintou: kintou, total: h + c + j + g + t + kintou };
}
function maxBonus(f, hi) {
  if (f(0) < 0) return 0;
  if (f(hi) >= 0) return hi;
  let lo = 0, h = hi;
  for (let k = 0; k < 50; k++) { const m = (lo + h) / 2; if (f(m) >= 0) lo = m; else h = m; }
  return lo;
}
function bonusModel(cur) {
  const sInc = cur.tot.sales, sEx = sInc / 1.1;
  const ct = sInc * 10 / 110 * (1 - BN.mnashi / 100);
  const A0 = cur.tot.op - ct, NO = cur.tot.ord - cur.tot.op, r = BN.socRate / 100;
  const lastB = bsIdx[bsIdx.length - 1];
  const cash = bsKey('現金預金 計', '現金預金計')[lastB] || 0;
  const kari = bsKey('仮払消費税')[lastB] || 0;
  const ctRemain = Math.max(0, ct - kari);
  const monthlyFixed = cur.tot.sga / 12, reserve = monthlyFixed * BN.resMonths;
  const book = bt => BN.socWhen === 'this' ? bt * (1 + r) : bt;
  const at = bt => {
    const A1 = A0 - book(bt), PT = A1 + NO, T = corpTax(PT, BN.kintou), NI = PT - T.total;
    return { bt, A1, PT, T, NI, rate: sEx ? A1 / sEx : 0, cashout: bt * (1 + r), endCash: cash - bt * (1 + r) - T.total - ctRemain };
  };
  const base = at(0), now = at(btTotal());
  const limZero = maxBonus(b => at(b).NI, 20000000);
  const limTgt = Math.max(0, (A0 - BN.tgtRate / 100 * sEx) / (1 + r));
  const limCash = maxBonus(b => at(b).endCash - reserve, 20000000);
  return {
    sInc, sEx, ct, A0, NO, base, now, at, limZero, limTgt, limCash,
    rec: Math.min(limZero, limTgt, limCash), cash, kari, ctRemain, reserve, monthlyFixed, r
  };
}

/* ============================================================
   レンダリング
   ============================================================ */
function kpiCard(label, val, sub, cls) {
  return '<div class="kpi"><div class="k-label">' + label + '</div><div class="k-val' + (cls ? ' ' + cls : '') + '">' + val + '</div><div class="k-sub">' + (sub || '') + '</div></div>';
}
const $ = id => document.getElementById(id);

function render() {
  if (!DATA) return;
  const cur = compute(), nx = computeNext(cur);
  const lastB = bsIdx[bsIdx.length - 1];
  const assets = bsKey('資産 合計', '資産合計'), equity = bsKey('純資産 合計', '純資産合計');
  const cash = bsKey('現金預金 計', '現金預金計')[lastB] || 0;
  const eq = assets[lastB] ? equity[lastB] / assets[lastB] : null;
  const sga36 = fy36Of('販管費 合計') || PLAN_SGA;
  const op36 = fy36Of('営業利益'), sales36 = fy36Of('売上高');

  /* ---- 現状 ---- */
  $('kpi').innerHTML =
    kpiCard('通期売上（着地見込み）', yen(cur.tot.sales), '年間計画 ' + yen(PLAN_SALES) + '　差 ' + syen(cur.tot.sales - PLAN_SALES), cur.tot.sales < PLAN_SALES ? 'down' : '') +
    kpiCard('通期営業利益（着地見込み）', syen(cur.tot.op), '計画 ' + yen(PLAN_OP) + '　差 ' + syen(cur.tot.op - PLAN_OP), cur.tot.op < 0 ? 'down' : '') +
    kpiCard('営業利益率', pct(cur.tot.sales ? cur.tot.op / cur.tot.sales : 0), '前期実績 ' + pct(sales36 ? op36 / sales36 : null)) +
    kpiCard('現金預金（' + M[lastB] + '末）', yen(cash), '自己資本比率 ' + pct(eq));
  $('kpiNote').textContent = '実績：' + M[0] + '〜' + M[S.last] + '（' + (S.last + 1) + 'ヶ月）／ 見込み：'
    + (S.last < N - 1 ? M[S.last + 1] + '〜' + M[N - 1] : 'なし') + '　※税込・試算表ベース';

  const cumA = [], cumP = []; let a = 0;
  for (let i = 0; i < N; i++) { a += cur.sales[i]; cumA.push(a); cumP.push(PLAN_SALES / N * (i + 1)); }
  $('cumLede').textContent = '年間計画 ' + yen(PLAN_SALES) + ' 円を12等分した累計ラインと、実績＋見込みの累計を重ねています。';
  drawChart($('c_cum'), {
    labels: M.map(monthLabel), height: 300, series: [
      { type: 'line', name: '累計 計画', values: cumP, color: GREY, dash: true, dots: false },
      { type: 'line', name: '累計 実績＋見込み', values: cumA, color: INK, w: 2.4 }]
  });
  drawChart($('c_op'), {
    labels: M.map(monthLabel), height: 280,
    series: [{ type: 'bar', name: '営業利益', values: cur.op, color: INK, colors: cur.op.map((v, i) => cur.isAct(i) ? INK : SHU) }],
    note: '黒＝実績　朱＝見込み'
  });
  $('findings').innerHTML = '<div class="flag">売上は年間計画に対して ' + pct(PLAN_SALES ? cur.tot.sales / PLAN_SALES : null)
    + '（' + syen(cur.tot.sales - PLAN_SALES) + '）。販管費は ' + yen(cur.tot.sga) + '（前期実績 ' + yen(sga36) + ' に対し '
    + syen(cur.tot.sga - sga36) + '）。営業利益の着地見込みは ' + syen(cur.tot.op) + ' です。</div>';

  /* ---- 損益 ---- */
  drawChart($('c_sales'), {
    labels: M.map(monthLabel), height: 300, series: [
      { type: 'bar', name: '売上高', values: cur.sales, color: INK, colors: cur.sales.map((v, i) => cur.isAct(i) ? INK : SHU2) },
      { type: 'line', name: '月次計画', values: new Array(N).fill(PLAN_SALES / N), color: GREY, dash: true, dots: false }],
    note: '濃色＝実績　淡朱＝見込み'
  });
  drawChart($('c_sga'), {
    labels: M.map(monthLabel), height: 310, stacked: true,
    series: CATS.map(c => ({ type: 'bar', name: c, values: cur.sgaCat[c], color: CATCOL[c] }))
  });
  drawChart($('c_margin'), {
    labels: M.map(monthLabel), height: 260, yfmt: 'pct', vfmt: 'pct', zeroBase: false, series: [
      { type: 'line', name: '売上総利益率', values: cur.gp.map((v, i) => cur.sales[i] ? v / cur.sales[i] : null), color: INK },
      { type: 'line', name: '営業利益率', values: cur.op.map((v, i) => cur.sales[i] ? v / cur.sales[i] : null), color: SHU }]
  });

  const sel = $('acctSel');
  const wantNames = SGA.map(it => it.name).join('|');
  if (sel.dataset.sig !== wantNames) {
    sel.innerHTML = SGA.map(it => '<option>' + it.name + '</option>').join('');
    sel.dataset.sig = wantNames;
    if (SGA.some(it => it.name === '役員報酬')) sel.value = '役員報酬';
  }
  const nm = sel.value || (SGA[0] && SGA[0].name);
  if (nm) {
    const it = plMap[nm], vals = cur.sgaItems[nm] || new Array(N).fill(0);
    drawChart($('c_acct'), {
      labels: M.map(monthLabel), height: 260, series: [
        { type: 'bar', name: nm, values: vals, color: INK, colors: vals.map((v, i) => cur.isAct(i) ? INK : SHU2) },
        { type: 'line', name: '月次計画', values: new Array(N).fill((it.annualPlan || 0) / 12), color: GREY, dash: true, dots: false }]
    });
    $('acctNote').textContent = '年間計画 ' + yen(it.annualPlan || 0) + '　着地見込み ' + yen(sum(vals))
      + '　差 ' + syen(sum(vals) - (it.annualPlan || 0));
  }

  /* ---- 健全性 ---- */
  const L = bsIdx.map(i => monthLabel(M[i])), pick = k => bsIdx.map(i => k[i] || 0);
  const cashArr = bsKey('現金預金 計', '現金預金計'), ar = bsKey('売掛金');
  const curAsset = bsKey('流動資産 計', '流動資産計'), fixAsset = bsKey('固定資産 計', '固定資産計');
  const liab = bsKey('負債 合計', '負債合計');
  const other = bsIdx.map(i => (curAsset[i] || 0) - (cashArr[i] || 0) - (ar[i] || 0));
  drawChart($('c_bs'), {
    labels: L, height: 310, stacked: true, series: [
      { type: 'bar', name: '現金預金', values: pick(cashArr), color: INK },
      { type: 'bar', name: '売掛金', values: pick(ar), color: MID },
      { type: 'bar', name: 'その他流動資産', values: other, color: LIGHT },
      { type: 'bar', name: '固定資産', values: pick(fixAsset), color: WASH },
      { type: 'line', name: '負債合計', values: pick(liab), color: SHU }]
  });
  drawChart($('c_cash'), {
    labels: L, height: 280, series: [
      { type: 'line', name: '現金預金', values: pick(cashArr), color: INK, w: 2.4 },
      { type: 'line', name: '売掛金', values: pick(ar), color: SHU }]
  });
  drawChart($('c_ratio'), {
    labels: L, height: 260, yfmt: 'pct', vfmt: 'pct', zeroBase: false, series: [
      { type: 'line', name: '自己資本比率', values: bsIdx.map(i => assets[i] ? equity[i] / assets[i] : null), color: INK },
      { type: 'line', name: '流動比率（負債合計ベース）', values: bsIdx.map(i => liab[i] ? curAsset[i] / liab[i] : null), color: SHU }],
    note: '流動比率は負債合計を分母にした簡便計算です。'
  });
  drawChart($('c_turn'), {
    labels: L, height: 240, yfmt: 'raw', vfmt: 'raw',
    series: [{ type: 'bar', name: '売掛金回転月数', values: bsIdx.map(i => cur.sales[i] ? (ar[i] || 0) / cur.sales[i] : null), color: MID }]
  });

  /* ---- 予測 ---- */
  $('v_salesAdj').textContent = (S.salesAdj > 0 ? '+' : '') + S.salesAdj + '%';
  $('v_cogs').textContent = S.cogsMode === 'rate' ? S.cogsRate.toFixed(1) + '%' : '実績累計 ' + pct(actualCogsRate(S.last));
  $('cogsRate').disabled = S.cogsMode !== 'rate';
  $('cogsRate').style.opacity = S.cogsMode === 'rate' ? 1 : .35;
  $('v_growth').textContent = (S.growth > 0 ? '+' : '') + S.growth + '%';
  $('v_season').textContent = S.season + '%';
  $('v_nyCogs').textContent = S.nyCogs.toFixed(1) + '%';
  $('baseHint').textContent = '前期実績 ' + yen(fy36Of('売上高')) + '／今期着地見込み ' + yen(cur.tot.sales);

  const rem = N - 1 - S.last;
  $('kpi4').innerHTML =
    kpiCard('売上（残' + rem + 'ヶ月の見込み）', yen(sum(cur.sales.slice(S.last + 1))), '実績累計 ' + yen(sum(cur.sales.slice(0, S.last + 1)))) +
    kpiCard('通期 売上', yen(cur.tot.sales), '計画比 ' + pct(PLAN_SALES ? cur.tot.sales / PLAN_SALES : null)) +
    kpiCard('通期 販管費', yen(cur.tot.sga), '前期比 ' + syen(cur.tot.sga - sga36)) +
    kpiCard('通期 営業利益', syen(cur.tot.op), '営業利益率 ' + pct(cur.tot.sales ? cur.tot.op / cur.tot.sales : 0), cur.tot.op < 0 ? 'down' : '');
  drawChart($('c_fc'), {
    labels: M.map(monthLabel), height: 310, series: [
      { type: 'bar', name: '売上高', values: cur.sales, color: INK, colors: cur.sales.map((v, i) => cur.isAct(i) ? INK : SHU2) },
      { type: 'bar', name: '原価＋販管費', values: cur.sgaTot.map((v, i) => v + cur.cogs[i]), color: LIGHT },
      { type: 'line', name: '営業利益', values: cur.op, color: SHU }]
  });
  $('fcTable').innerHTML = tbl(['科目'].concat(M.map(monthLabel), ['通期', '年間計画', '差額']), [
    ['売上高', cur.sales, PLAN_SALES],
    ['売上原価', cur.cogs, K.cogs.annualPlan],
    ['売上総利益', cur.gp, K.gp.annualPlan]
  ].concat(CATS.map(c => [c, cur.sgaCat[c], null]), [
    ['販管費 合計', cur.sgaTot, PLAN_SGA, 'sum'],
    ['営業利益', cur.op, PLAN_OP, 'sum']]), cur.isAct);

  $('kpi38').innerHTML =
    kpiCard('次期 売上計画', yen(nx.tot.sales), '基準 ' + yen(nx.base) + ' × ' + (S.growth > 0 ? '+' : '') + S.growth + '%') +
    kpiCard('次期 販管費', yen(nx.tot.sga), '今期着地比 ' + syen(nx.tot.sga - cur.tot.sga)) +
    kpiCard('次期 営業利益', syen(nx.tot.op), '営業利益率 ' + pct(nx.tot.sales ? nx.tot.op / nx.tot.sales : 0), nx.tot.op < 0 ? 'down' : '') +
    kpiCard('損益分岐点売上高', yen(nx.bep), nx.tot.sales >= nx.bep ? '計画売上が ' + yen(nx.tot.sales - nx.bep) + ' 上回る' : '計画売上が ' + yen(nx.bep - nx.tot.sales) + ' 不足', nx.tot.sales < nx.bep ? 'down' : '');
  drawChart($('c_ny'), {
    labels: M.map(monthLabel), height: 310, series: [
      { type: 'bar', name: '次期 売上計画', values: nx.sales, color: INK },
      { type: 'bar', name: '原価＋販管費', values: nx.sgaTot.map((v, i) => v + nx.cogs[i]), color: LIGHT },
      { type: 'line', name: '営業利益', values: nx.op, color: SHU }]
  });
  $('nyTable').innerHTML = tbl(['科目'].concat(M.map(monthLabel), ['通期', '今期着地', '差額']), [
    ['売上高', nx.sales, cur.tot.sales],
    ['売上原価', nx.cogs, cur.tot.cogs],
    ['売上総利益', nx.sales.map((v, i) => v - nx.cogs[i]), cur.tot.gp]
  ].concat(CATS.map(c => [c, nx.sgaCat[c], sum(cur.sgaCat[c])]), [
    ['販管費 合計', nx.sgaTot, cur.tot.sga, 'sum'],
    ['営業利益', nx.op, cur.tot.op, 'sum']]), () => true);

  drawChart($('c_3y'), {
    labels: ['売上高', '売上総利益', '販管費', '営業利益'], height: 300, ml: 96, compactLabels: false, series: [
      { type: 'bar', name: '前期 実績', values: [sales36, fy36Of('売上総利益'), sga36, op36], color: LIGHT },
      { type: 'bar', name: '今期 着地見込み', values: [cur.tot.sales, cur.tot.gp, cur.tot.sga, cur.tot.op], color: INK },
      { type: 'bar', name: '次期 計画', values: [nx.tot.sales, nx.tot.sales - nx.tot.cogs, nx.tot.sga, nx.tot.op], color: SHU }]
  });

  /* ---- 明細 ---- */
  $('plTable').innerHTML = tbl(['科目'].concat(M.map(monthLabel), ['通期', '年間計画', '差額']), [
    ['売上高', cur.sales, PLAN_SALES],
    ['売上原価', cur.cogs, K.cogs.annualPlan],
    ['売上総利益', cur.gp, K.gp.annualPlan, 'sum'],
    ['【販売費及び一般管理費】']
  ].concat(SGA.map(it => [it.name, cur.sgaItems[it.name], it.annualPlan]), [
    ['販管費 合計', cur.sgaTot, PLAN_SGA, 'sum'],
    ['営業利益', cur.op, PLAN_OP, 'sum'],
    ['経常利益', cur.ord, K.ord.annualPlan, 'sum']]), cur.isAct);
  $('bsTable').innerHTML = tblBS(DATA.bs.filter(b => !/^※/.test(b.name) && !/再掲/.test(b.name) && !/回転月数/.test(b.name) && b.name !== '【参考指標】'));
  renderCaveats(cur);

  /* ---- 決算賞与 ---- */
  renderBonus(cur);
}

function renderCaveats(cur) {
  const ct = cur.tot.sales * 10 / 110 * (1 - BN.mnashi / 100);
  const yakuin = plMap['役員報酬'];
  let h = '';
  if (yakuin) {
    const zero = [];
    for (let i = 0; i <= S.last; i++) if (!yakuin.v[i]) zero.push(M[i]);
    if (zero.length) h += '<div class="flag">役員報酬が ' + zero.join('・') + ' でゼロになっています。組み替えなのか未計上なのか、確認が必要です。</div>';
  }
  if (UNMAPPED.length) h += '<div class="warn">カテゴリー未定義の費目を「一般管理費」に含めています：' + UNMAPPED.join('、') + '</div>';
  h += '<ul class="lede"><li>金額はすべて税込・試算表ベース。消費税の発生見込み（通期 約 ' + yen(ct) + ' 円）は営業利益に含めていません。</li>'
    + '<li>BSは試算表を入力済みの ' + M[bsIdx[0]] + '〜' + M[bsIdx[bsIdx.length - 1]] + ' のみ表示しています。</li>'
    + '<li>次期の月別配分は今期の季節性を使った推計です。過去複数年の月次データを取り込めば精度が上がります。</li></ul>';
  $('caveats').innerHTML = h;
}

function tbl(head, rows, isAct) {
  let h = '<div class="tblwrap"><table><thead><tr>' + head.map((x, i) => '<th' + (i === 0 ? ' class="nm"' : '') + '>' + x + '</th>').join('') + '</tr></thead><tbody>';
  rows.forEach(r => {
    if (r.length === 1) { h += '<tr class="grp"><td class="nm">' + r[0] + '</td><td colspan="' + (head.length - 1) + '"></td></tr>'; return; }
    const nm = r[0], vals = r[1] || [], plan = r[2], kind = r[3], tot = sum(vals);
    h += '<tr' + (kind === 'sum' ? ' class="sum"' : '') + '><td class="nm">' + nm + '</td>';
    vals.forEach((v, i) => { h += '<td class="' + (isAct(i) ? '' : 'fc') + (v < 0 ? ' neg' : '') + '">' + (v ? syen(v) : '—') + '</td>'; });
    h += '<td><b>' + syen(tot) + '</b></td><td>' + (plan == null ? '—' : yen(plan)) + '</td>';
    h += '<td class="' + (plan != null && tot - plan < 0 ? 'neg' : '') + '">' + (plan == null ? '—' : syen(tot - plan)) + '</td></tr>';
  });
  return h + '</tbody></table></div>';
}
function tblBS(rows) {
  let h = '<div class="tblwrap"><table><thead><tr><th class="nm">科目</th>'
    + bsIdx.map(i => '<th>' + monthLabel(M[i]) + '</th>').join('') + '<th>期首比</th></tr></thead><tbody>';
  rows.forEach(b => {
    if (/^【/.test(b.name)) { h += '<tr class="grp"><td class="nm">' + b.name + '</td><td colspan="' + (bsIdx.length + 1) + '"></td></tr>'; return; }
    const sm = b.name.indexOf('合計') >= 0 || b.name.indexOf(' 計') >= 0;
    h += '<tr' + (sm ? ' class="sum"' : '') + '><td class="nm">' + b.name + '</td>';
    bsIdx.forEach(i => { const v = b.v[i]; h += '<td class="' + (v < 0 ? 'neg' : '') + '">' + (v ? syen(v) : '—') + '</td>'; });
    const d = (b.v[bsIdx[bsIdx.length - 1]] || 0) - (b.v[bsIdx[0]] || 0);
    h += '<td class="' + (d < 0 ? 'neg' : '') + '">' + (d ? syen(d) : '—') + '</td></tr>';
  });
  return h + '</tbody></table></div>';
}

/* ---------- 賞与タブ ---------- */
function buildPeople() {
  const tb = $('peopleBody');
  tb.innerHTML = BN.people.map((p, i) =>
    '<tr><td class="nm"><input type="text" data-i="' + i + '" class="pn" value="' + String(p.name).replace(/"/g, '&quot;') + '"></td>'
    + '<td><input type="number" data-i="' + i + '" class="pa" step="10000" min="0" value="' + p.amt + '"></td>'
    + '<td class="ps"></td><td class="pt"></td><td class="ph"><b></b></td>'
    + '<td>' + (BN.people.length > 1 ? '<button class="btn" data-del="' + i + '" style="padding:2px 8px;font-size:11px">削除</button>' : '') + '</td></tr>').join('');
  tb.querySelectorAll('.pa').forEach(el => el.addEventListener('input', e => {
    BN.people[+e.target.dataset.i].amt = Math.max(0, +e.target.value || 0);
    BN.peopleDirty = true;
    $('btRange').value = Math.min(4000000, btTotal()); render();
  }));
  tb.querySelectorAll('.pn').forEach(el => el.addEventListener('input', e => {
    BN.people[+e.target.dataset.i].name = e.target.value; BN.peopleDirty = true; render();
  }));
  tb.querySelectorAll('[data-del]').forEach(el => el.addEventListener('click', e => {
    BN.people.splice(+e.currentTarget.dataset.del, 1); BN.peopleDirty = true; buildPeople(); render();
  }));
}
function renderBonus(cur) {
  const m = bonusModel(cur), bt = btTotal();
  $('v_bt').textContent = yen(bt);
  $('v_soc').textContent = BN.socRate.toFixed(1) + '%';
  $('v_mnashi').textContent = BN.mnashi + '%';
  $('v_kintou').textContent = yen(BN.kintou);
  $('v_esoc').textContent = BN.empSoc.toFixed(1) + '%';
  $('v_etax').textContent = BN.empTax.toFixed(2) + '%';
  $('v_tgt').textContent = BN.tgtRate.toFixed(1) + '%';
  $('v_res').textContent = BN.resMonths + 'ヶ月';
  $('preRateEcho').textContent = pct(m.base.rate);

  $('kpiBonus').innerHTML =
    kpiCard('賞与前 営業利益率', pct(m.base.rate), '調整後営業利益 ' + syen(m.A0)) +
    kpiCard('臨時賞与 総額', yen(bt), '会社の総コスト ' + yen(m.now.cashout)) +
    kpiCard('賞与後 営業利益率', pct(m.now.rate), '調整後営業利益 ' + syen(m.now.A1), m.now.rate < 0 ? 'down' : '') +
    kpiCard('正味の現金流出', yen(m.now.cashout - (m.base.T.total - m.now.T.total)),
      '節税 ' + yen(m.base.T.total - m.now.T.total) + ' を差引後');

  BN.people.forEach((p, i) => {
    const soc = p.amt * BN.empSoc / 100, tax = (p.amt - soc) * BN.empTax / 100;
    const tr = $('peopleBody').children[i]; if (!tr) return;
    tr.querySelector('.ps').textContent = yen(soc);
    tr.querySelector('.pt').textContent = yen(tax);
    tr.querySelector('.ph b').textContent = yen(p.amt - soc - tax);
  });

  const mn = m.rec;
  const lc = (lab, val, sub, pick) => kpiCard(lab, yen(val), sub, pick ? 'down' : '');
  $('limits').innerHTML =
    lc('税引後利益がゼロになる額', m.limZero, 'これを超えると当期は赤字', Math.abs(m.limZero - mn) < 1) +
    lc('営業利益率 ' + BN.tgtRate.toFixed(1) + '% を守る額', m.limTgt, '消費税調整後・対税抜売上', Math.abs(m.limTgt - mn) < 1) +
    lc('運転資金 ' + BN.resMonths + 'ヶ月を残す額', m.limCash, '現預金 ' + yen(m.cash) + ' が起点', Math.abs(m.limCash - mn) < 1) +
    lc('この3つで見た上限', mn, '現在の設定は ' + yen(bt) + '。' + (bt <= mn ? '余裕 ' + yen(mn - bt) : '超過 ' + yen(bt - mn)), false);

  const soc = (m.now.cashout - bt) * (BN.socWhen === 'this' ? 1 : 0);
  const rows = [
    ['税込売上（通期見込み）', m.sInc, null, m.sInc],
    ['税抜売上（÷1.1）', m.sEx, null, m.sEx],
    ['営業利益（税込PL）', cur.tot.op, null, cur.tot.op],
    ['消費税 発生見込み（△）', -m.ct, null, -m.ct],
    ['臨時賞与（△）', 0, -bt, -bt],
    ['賞与の会社負担社会保険料（△）', 0, -soc, -soc],
    ['調整後 営業利益', m.A0, m.now.A1 - m.A0, m.now.A1, 'sum'],
    ['調整後 営業利益率（対税抜売上）', m.base.rate, null, m.now.rate, 'pct'],
    ['営業外損益', m.NO, null, m.NO],
    ['税引前当期純利益', m.base.PT, m.now.PT - m.base.PT, m.now.PT, 'sum'],
    ['　法人税・地方法人税', -(m.base.T.houjin + m.base.T.chihou), null, -(m.now.T.houjin + m.now.T.chihou)],
    ['　法人住民税（法人税割＋均等割）', -(m.base.T.jumin + m.base.T.kintou), null, -(m.now.T.jumin + m.now.T.kintou)],
    ['　事業税・特別法人事業税', -(m.base.T.jigyo + m.base.T.toku), null, -(m.now.T.jigyo + m.now.T.toku)],
    ['法人税等 合計（△）', -m.base.T.total, m.base.T.total - m.now.T.total, -m.now.T.total],
    ['税引後 当期純利益', m.base.NI, m.now.NI - m.base.NI, m.now.NI, 'sum'],
    ['期末後に残る現預金', m.base.endCash, m.now.endCash - m.base.endCash, m.now.endCash, 'sum']
  ];
  let h = '<div class="tblwrap"><table><thead><tr><th class="nm">項目</th><th>賞与なし</th><th>増減</th><th>賞与あり</th></tr></thead><tbody>';
  rows.forEach(r => {
    const nm = r[0], a = r[1], d = r[2], b = r[3], k = r[4];
    const f = v => v === null ? '—' : (k === 'pct' ? pct(v) : syen(Math.round(v)));
    h += '<tr' + (k === 'sum' ? ' class="sum"' : '') + '><td class="nm">' + nm + '</td>'
      + '<td class="' + (a < 0 ? 'neg' : '') + '">' + f(a) + '</td>'
      + '<td class="' + (d < 0 ? 'neg' : '') + '">' + (d === null || k === 'pct' ? '—' : syen(Math.round(d))) + '</td>'
      + '<td class="' + (b < 0 ? 'neg' : '') + '">' + f(b) + '</td></tr>';
  });
  $('bonusTable').innerHTML = h + '</tbody></table></div>';

  const step = 250000, xs = []; for (let i = 0; i < 13; i++) xs.push(i * step);
  const pts = xs.map(x => m.at(x)), labs = xs.map(x => (x / 10000) + '万');
  drawChart($('c_bcurve'), {
    labels: labs, height: 280, yfmt: 'pct', vfmt: 'pct', zeroBase: false, compactLabels: false, series: [
      { type: 'line', name: '賞与後 営業利益率', values: pts.map(p => p.rate), color: INK },
      { type: 'line', name: '確保したい利益率', values: xs.map(() => BN.tgtRate / 100), color: SHU, dash: true, dots: false }],
    note: '横軸は賞与総額。'
  });
  drawChart($('c_bcash'), {
    labels: labs, height: 280, compactLabels: false, series: [
      { type: 'bar', name: '期末後に残る現預金', values: pts.map(p => p.endCash), color: LIGHT, colors: pts.map(p => p.endCash >= m.reserve ? LIGHT : SHU2) },
      { type: 'line', name: '残したい運転資金', values: xs.map(() => m.reserve), color: SHU, dash: true, dots: false },
      { type: 'line', name: '税引後 当期純利益', values: pts.map(p => p.NI), color: INK }],
    note: '朱色の棒は運転資金の下限を割り込む水準です。'
  });

  $('bonusNotes').innerHTML = '<h2>試算の前提</h2><ul class="lede">'
    + '<li>法人税等は中小法人の標準税率で概算（所得800万円以下15%・超過分23.2%、地方法人税10.3%、住民税法人税割7.0%、事業税3.5/5.3/7.0%、特別法人事業税37%、均等割 ' + yen(BN.kintou) + '）。事業税の翌期損金算入と繰越欠損金は考慮していません。実額は顧問税理士にご確認ください。</li>'
    + '<li>消費税は簡易課税・みなし仕入率' + BN.mnashi + '%を仮定した発生見込み ' + yen(m.ct) + ' を控除。中間納付済み ' + yen(m.kari) + ' を除いた残納付は ' + yen(m.ctRemain) + ' です。</li>'
    + '<li>決算賞与を損金にするには、期末までに各人へ支給額を書面通知し、翌日から1ヶ月以内に通知どおり支給し、当期に未払計上する必要があります。</li>'
    + '<li>会社負担の社会保険料は、期末後の支給なら税務上は翌期の費用です。' + (BN.socWhen === 'this' ? '現在は当期に含める設定（保守的）' : '現在は翌期に計上する設定') + 'です。</li>'
    + '<li>厚生年金保険料の賞与上限は1回150万円、健康保険は年間573万円です。</li>'
    + '<li>現預金は ' + M[bsIdx[bsIdx.length - 1]] + '末の ' + yen(m.cash) + ' が起点で、それ以降の入出金は反映していません。</li></ul>';
}
