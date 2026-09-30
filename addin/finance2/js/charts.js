/* ============================================================
 * charts.js — 依存なしのSVGチャート（タスクペイン幅に追従）
 * ============================================================ */
(function (global) {
  "use strict";
  const C = { navy: "#1F3F6E", navyS: "#DCE4F0", gray: "#8A939C", light: "#B9C0C7", grid: "#E3E6E4", muted: "#56606B", amber: "#9A560A" };
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function niceScale(min, max, ticks = 4) {
    if (min === max) { max = min + 1; }
    const span = max - min, raw = span / ticks;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map(k => k * mag).find(s => s >= raw);
    return { lo: Math.floor(min / step) * step, hi: Math.ceil(max / step) * step, step };
  }
  const fmtTick = (v, unit) => {
    const d = unit === "百万" ? v / 1e6 : unit === "千" ? v / 1e3 : v;
    return Math.abs(d) >= 100 || Number.isInteger(d) ? Math.round(d).toLocaleString() : d.toFixed(1);
  };

  /**
   * labels: x軸ラベル, bars: [{values, color}], lines: [{values, color, width, dash, dots}],
   * band: {lo:[], hi:[]}, marker: index（実績/予測の境界）
   */
  function chart(o) {
    const W = 360, H = o.height || 180, L = 62, R = 6, T = 10, B = 22;
    const all = [];
    (o.bars || []).forEach(b => b.values.forEach(v => v != null && all.push(v)));
    (o.lines || []).forEach(l => l.values.forEach(v => v != null && all.push(v)));
    if (o.band) { o.band.lo.forEach(v => v != null && all.push(v)); o.band.hi.forEach(v => v != null && all.push(v)); }
    if (!all.length) return `<div class="chart-empty">表示できるデータがありません</div>`;
    let mn = Math.min(...all), mx = Math.max(...all);
    if (o.zero !== false) { mn = Math.min(0, mn); mx = Math.max(0, mx); }
    const sc = niceScale(mn, mx);
    const n = o.labels.length, step = (W - L - R) / n;
    const y = (v) => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    const cx = (i) => L + step * i + step / 2;
    const s = [`<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="${esc(o.label || "グラフ")}" class="chart">`];
    for (let v = sc.lo; v <= sc.hi + 1e-9; v += sc.step) {
      s.push(`<line x1="${L}" x2="${W - R}" y1="${y(v).toFixed(1)}" y2="${y(v).toFixed(1)}" stroke="${v === 0 ? C.light : C.grid}"/>`);
      s.push(`<text x="${L - 5}" y="${(y(v) + 3.5).toFixed(1)}" font-size="9" fill="${C.muted}" text-anchor="end">${fmtTick(v, o.unit)}</text>`);
    }
    if (o.band) {
      const pts = [];
      o.band.hi.forEach((v, i) => v != null && pts.push([cx(i), y(v)]));
      for (let i = o.band.lo.length - 1; i >= 0; i--) if (o.band.lo[i] != null) pts.push([cx(i), y(o.band.lo[i])]);
      if (pts.length > 2) s.push(`<polygon points="${pts.map(p => p[0].toFixed(1) + "," + p[1].toFixed(1)).join(" ")}" fill="${C.navyS}"/>`);
    }
    const nb = (o.bars || []).length;
    (o.bars || []).forEach((b, bi) => {
      const bw = step * 0.62 / nb;
      b.values.forEach((v, i) => {
        if (v == null) return;
        const x = L + step * i + step * 0.19 + bw * bi, y0 = y(0), y1 = y(v);
        s.push(`<rect x="${x.toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(1, Math.abs(y1 - y0)).toFixed(1)}" rx="1.5" fill="${b.color || C.navy}"><title>${esc(o.labels[i])}：${Math.round(v).toLocaleString()}</title></rect>`);
      });
    });
    (o.lines || []).forEach(l => {
      let seg = [];
      const flush = () => { if (seg.length > 1) s.push(`<polyline points="${seg.join(" ")}" fill="none" stroke="${l.color}" stroke-width="${l.width || 2}" ${l.dash ? `stroke-dasharray="${l.dash}"` : ""} stroke-linejoin="round"/>`); seg = []; };
      l.values.forEach((v, i) => { if (v == null) flush(); else seg.push(cx(i).toFixed(1) + "," + y(v).toFixed(1)); });
      flush();
      if (l.dots) l.values.forEach((v, i) => { if (v != null) s.push(`<circle cx="${cx(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="3" fill="${l.color}"><title>${esc(o.labels[i])}：${Math.round(v).toLocaleString()}</title></circle>`); });
    });
    if (o.marker != null) {
      const x = L + step * (o.marker + 1);
      s.push(`<line x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${T}" y2="${H - B}" stroke="${C.muted}" stroke-dasharray="2 2"/>`);
      s.push(`<text x="${(x + 3).toFixed(1)}" y="${T + 9}" font-size="9.5" fill="${C.muted}">予測</text>`);
    }
    o.labels.forEach((t, i) => s.push(`<text x="${cx(i).toFixed(1)}" y="${H - 7}" font-size="9.5" fill="${C.muted}" text-anchor="middle">${esc(t)}</text>`));
    s.push("</svg>");
    return s.join("");
  }

  global.FinCharts = { chart, C };
})(window);
