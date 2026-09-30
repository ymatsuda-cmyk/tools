/* LedgerPilot SVG チャート（依存なし） */
(function (root) {
  "use strict";
  var NS = "http://www.w3.org/2000/svg";

  function man(n) { // 軸ラベル用の万円表記
    var a = Math.abs(n);
    if (a >= 1e8) return (n / 1e8).toFixed(1).replace(/\.0$/, "") + "億";
    if (a >= 1e4) return Math.round(n / 1e4).toLocaleString("ja-JP") + "万";
    return String(Math.round(n));
  }
  function niceMax(v) {
    if (v <= 0) return 1;
    var p = Math.pow(10, Math.floor(Math.log10(v))), f = v / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  }
  // 0を必ず目盛に含む、きりの良い軸
  function axis(lo, hi) {
    var step = niceMax(Math.max(hi - lo, 1) / 4);
    var min = Math.floor(lo / step) * step, max = Math.ceil(hi / step) * step;
    if (max === min) max = min + step;
    return { min: min, max: max, steps: Math.round((max - min) / step) };
  }
  function el(tag, attrs, parent) {
    var e = document.createElementNS(NS, tag);
    for (var k in attrs) e.setAttribute(k, attrs[k]);
    if (parent) parent.appendChild(e);
    return e;
  }
  function text(parent, x, y, s, attrs) {
    var t = el("text", Object.assign({ x: x, y: y }, attrs || {}), parent);
    t.textContent = s; return t;
  }
  function tip(parent, s) { var t = el("title", {}, parent); t.textContent = s; }
  function yen(n) { return Math.round(n).toLocaleString("ja-JP") + "円"; }

  /** 棒（複数系列）＋折れ線 */
  function barLine(host, o) {
    host.innerHTML = "";
    var W = Math.max(host.clientWidth || 600, 280), H = o.height || 220;
    var padL = 44, padR = 8, padT = 10, padB = 26;
    var n = o.labels.length; if (!n) return;
    var vals = [];
    o.bars.forEach(function (b) { vals = vals.concat(b.values); });
    if (o.line) vals = vals.concat(o.line.values);
    var ax = axis(Math.min.apply(null, vals.concat([0])), Math.max.apply(null, vals.concat([0])));
    var max = ax.max, min = ax.min, steps = ax.steps;
    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, role: "img", "aria-label": o.title || "グラフ" }, host);
    var iw = W - padL - padR, ih = H - padT - padB;
    var y = function (v) { return padT + ih * (max - v) / (max - min); };
    // 目盛
    for (var i = 0; i <= steps; i++) {
      var v = min + (max - min) * i / steps;
      el("line", { x1: padL, x2: W - padR, y1: y(v), y2: y(v), class: v === 0 ? "axis" : "grid" }, svg);
      text(svg, padL - 6, y(v) + 4, man(v), { "text-anchor": "end", "font-size": 10 });
    }
    if (min < 0) el("line", { x1: padL, x2: W - padR, y1: y(0), y2: y(0), class: "axis" }, svg);
    var slot = iw / n, nb = o.bars.length;
    var bw = Math.max(3, Math.min(26, slot * 0.7 / nb));
    var labelEvery = Math.ceil(n / Math.max(1, Math.floor(iw / 44)));
    o.labels.forEach(function (lab, i) {
      var g = el("g", { class: o.onClick ? "hit" : "" }, svg);
      var x0 = padL + slot * i;
      el("rect", { x: x0, y: padT, width: slot, height: ih, fill: i === o.active ? "#eceefb" : "transparent", class: "bar-bg" }, g);
      var tips = [lab];
      o.bars.forEach(function (b, j) {
        var v = b.values[i] || 0;
        var bx = x0 + (slot - bw * nb) / 2 + bw * j;
        el("rect", { x: bx, y: Math.min(y(v), y(0)), width: bw, height: Math.max(1, Math.abs(y(v) - y(0))),
          fill: b.color, rx: 1.5 }, g);
        tips.push(b.name + " " + yen(v));
      });
      if (o.line) tips.push(o.line.name + " " + yen(o.line.values[i] || 0));
      tip(g, tips.join("\n"));
      if (i % labelEvery === 0) text(svg, x0 + slot / 2, H - 8, lab, { "text-anchor": "middle", "font-size": 10 });
      if (o.onClick) g.addEventListener("click", function () { o.onClick(i); });
    });
    if (o.line) {
      var pts = o.line.values.map(function (v, i) { return (padL + slot * i + slot / 2) + "," + y(v || 0); });
      el("polyline", { points: pts.join(" "), fill: "none", stroke: o.line.color, "stroke-width": 2, "pointer-events": "none" }, svg);
      o.line.values.forEach(function (v, i) {
        el("circle", { cx: padL + slot * i + slot / 2, cy: y(v || 0), r: 3, fill: "#fff", stroke: o.line.color, "stroke-width": 2, "pointer-events": "none" }, svg);
      });
    }
  }

  /** 損益の階段（ウォーターフォール） steps: [{label, value, kind: 'total'|'down'|'up'}] */
  function waterfall(host, steps, onClick) {
    host.innerHTML = "";
    var W = Math.max(host.clientWidth || 600, 280), H = 200;
    var padL = 44, padR = 8, padT = 18, padB = 34;
    var run = 0, bars = [];
    steps.forEach(function (s) {
      if (s.kind === "total") { bars.push({ s: s, a: 0, b: s.value }); run = s.value; }
      else if (s.kind === "down") { bars.push({ s: s, a: run - s.value, b: run }); run -= s.value; }
      else { bars.push({ s: s, a: run, b: run + s.value }); run += s.value; }
    });
    var hi = Math.max.apply(null, bars.map(function (b) { return Math.max(b.a, b.b); }).concat([0]));
    var lo = Math.min.apply(null, bars.map(function (b) { return Math.min(b.a, b.b); }).concat([0]));
    var ax = axis(lo, hi), max = ax.max, min = ax.min;
    var svg = el("svg", { viewBox: "0 0 " + W + " " + H, role: "img", "aria-label": "損益の内訳" }, host);
    var iw = W - padL - padR, ih = H - padT - padB;
    var y = function (v) { return padT + ih * (max - v) / (max - min); };
    for (var i = 0; i <= ax.steps; i++) {
      var v = min + (max - min) * i / ax.steps;
      el("line", { x1: padL, x2: W - padR, y1: y(v), y2: y(v), class: v === 0 ? "axis" : "grid" }, svg);
      text(svg, padL - 6, y(v) + 4, man(v), { "text-anchor": "end", "font-size": 10 });
    }
    var slot = iw / bars.length, bw = Math.min(54, slot * 0.62);
    bars.forEach(function (b, i) {
      var x = padL + slot * i + (slot - bw) / 2;
      var color = b.s.kind === "total" ? (b.b < 0 ? "#c23a22" : "#34398f") : b.s.kind === "down" ? "#b9bdd6" : "#7fb89b";
      var g = el("g", { class: onClick && b.s.key ? "hit" : "" }, svg);
      el("rect", { x: x, y: y(Math.max(b.a, b.b)), width: bw, height: Math.max(1, Math.abs(y(b.a) - y(b.b))), fill: color, rx: 2 }, g);
      tip(g, b.s.label + " " + yen(b.s.value));
      if (i < bars.length - 1) {
        var end = b.s.kind === "down" ? b.a : b.b;
        el("line", { x1: x + bw, x2: x + slot, y1: y(end), y2: y(end), stroke: "#7b8497", "stroke-dasharray": "2 2" }, svg);
      }
      text(svg, x + bw / 2, y(Math.max(b.a, b.b)) - 4, (b.s.kind === "down" ? "−" : b.s.kind === "up" ? "+" : "") + man(b.s.value),
        { "text-anchor": "middle", "font-size": 10, "font-weight": b.s.kind === "total" ? 700 : 400 });
      var lab = b.s.label;
      text(svg, x + bw / 2, H - 18, lab, { "text-anchor": "middle", "font-size": slot < 60 ? 9 : 10 });
      if (onClick && b.s.key) g.addEventListener("click", function () { onClick(b.s.key); });
    });
  }

  root.LPCharts = { barLine: barLine, waterfall: waterfall, man: man };
})(window);
