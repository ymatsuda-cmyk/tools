/* ============================================================
 * yojitsu.js — 予実タブ（①売上 ②コスト ③利益率）
 * ------------------------------------------------------------
 * ①売上   ：累計の積み上げ（実績＝面、見込み＝破線、計画＝点線）。
 *            保守／開発（既存）／開発（新規）はボタンでONにした区分だけ
 *            下から 保守→既存→新規 の順に色分けし、OFFの区分は「その他」として上に積む。
 * ②コスト ：原価・販管費の月別棒グラフ。計画を横線で重ね、突発（計画比○%超かつ○円超）を赤で示す。
 * ③利益率 ：年度末の着地。受注確定ベース（実績＋保守＋受注残）と着地見込み（＋営業見込み）を
 *            計画と並べ、調整後営業利益率（税抜）を目標と比べる。見込み部分は斜線。
 *
 * 区分の内訳は「売上区分」シートから読む（保守・開発（新規）の実績と計画。既存＝売上高−保守−新規）。
 * 見込みの前提（受注残・追加見込み・原価率など）はブックの設定に保存する。
 * ============================================================ */
(function (global) {
  "use strict";
  const SEG_SHEET = "売上区分";
  const SKEY = "finance2.yojitsu";
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const yen = (n) => (n == null || !isFinite(n) ? "—" : (n < 0 ? "−" : "") + Math.abs(Math.round(n)).toLocaleString());
  const sgn = (n) => (n == null || !isFinite(n) ? "—" : (n > 0 ? "+" : n < 0 ? "−" : "±") + Math.abs(Math.round(n)).toLocaleString());
  const sum = (a) => a.reduce((t, v) => t + (v || 0), 0);
  const cum = (a) => { let s = 0; return a.map(v => (s += v || 0)); };

  const CAT = {
    m: { name: "保守", col: "#2E8B6E", fill: "rgba(46,139,110,0.75)" },
    e: { name: "開発（既存）", col: "#1F3F6E", fill: "rgba(31,63,110,0.75)" },
    n: { name: "開発（新規）", col: "#6A5BB8", fill: "rgba(106,91,184,0.75)" }
  };
  const KEYS = ["m", "e", "n"];
  const COL = { plan: "#A9B3BF", fc: "#C9781A", cogs: "#7B848D", sga: "#D98A6A", spike: "#C0392B", firm: "#8FA9CC", act: "#1F3F6E", planBar: "#C9D5E6", ok: "#2E7D6B", bad: "#8E2A2A", muted: "#56606B", grid: "#E3E6E4" };

  /* ---------- 設定（ブックに保存） ---------- */
  const DEF = { bl: 0, ex: null, nw: null, cr: null, sr: 100, tgt: 15, mr: 50, sp: { cp: 120, ca: 200000, sp: 120, sa: 200000 }, sel: { m: false, e: false, n: false }, open: {} };
  function loadParams() {
    let p = null;
    try { const s = Office.context.document.settings.get(SKEY); if (s) p = JSON.parse(s); } catch (e) {}
    if (!p) { try { p = JSON.parse(localStorage.getItem(SKEY) || "null"); } catch (e) {} }
    const o = JSON.parse(JSON.stringify(DEF));
    if (p) { Object.assign(o, p); o.sp = Object.assign({}, DEF.sp, p.sp || {}); o.sel = Object.assign({}, DEF.sel, p.sel || {}); o.open = Object.assign({}, p.open || {}); }
    return o;
  }
  let saveTimer = null;
  function saveParams(p) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const s = JSON.stringify(p);
      try { localStorage.setItem(SKEY, s); } catch (e) {}
      try { const st = Office.context.document.settings; st.set(SKEY, s); st.saveAsync(() => {}); } catch (e) {}
    }, 500);
  }

  /* ---------- 売上区分シート ---------- */
  async function readSegments() {
    return Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets; wss.load("items/name"); await ctx.sync();
      const name = wss.items.map(w => w.name).find(n => n.indexOf(SEG_SHEET) === 0);
      if (!name) return { exists: false };
      const ws = wss.getItem(name);
      const used = ws.getUsedRange(true).getBoundingRect(ws.getRange("A1"));
      used.load("values"); await ctx.sync();
      const vals = used.values;
      const hdr = MonthlySheets.findHeader(vals);
      if (!hdr) return { exists: true, sheet: name, error: "売上区分シートの見出し行（2026/10 形式の月）が見つかりません" };
      const periods = hdr.labels.map(l => { const [y, m] = l.split("/").map(Number); return y + "-" + String(m).padStart(2, "0"); });
      const out = { act: { m: {}, n: {} }, plan: { m: {}, n: {} } };
      let mode = null;
      for (let r = hdr.row + 1; r < vals.length; r++) {
        const a = String(vals[r][0] == null ? "" : vals[r][0]).replace(/[\s\u3000]/g, "");
        if (!a) continue;
        if (/実績/.test(a) && /^【/.test(a)) { mode = "act"; continue; }
        if (/計画/.test(a) && /^【/.test(a)) { mode = "plan"; continue; }
        if (!mode) continue;
        const k = /保守/.test(a) ? "m" : /新規/.test(a) ? "n" : null;
        if (!k) continue;
        hdr.cols.forEach((c, i) => { const v = vals[r][c]; if (typeof v === "number") out[mode][k][periods[i]] = v; });
      }
      return { exists: true, sheet: name, periods, ...out };
    });
  }
  async function createSegments(M, fy) {
    const months = M.fyMonths(fy);
    const lab = (p) => p.slice(0, 4) + "/" + Number(p.slice(5));
    const W = 13;
    const row = (a) => { const r = Array(W).fill(""); a.forEach((v, i) => (r[i] = v)); return r; };
    const g = [
      row(["売上区分（保守・開発の内訳）"]),
      row(["対象期", M.fyLabel(fy)]),
      row(["保守と開発（新規）の金額を入れてください。開発（既存）は「売上高 − 保守 − 開発（新規）」でfinance2が計算します。"]),
      row(["区分"].concat(months.map(lab))),
      row(["【実績】"]),
      row(["保守"].concat(months.map(() => 0))),
      row(["開発（新規）"].concat(months.map(() => 0))),
      row([""]),
      row(["【計画】"]),
      row(["保守"].concat(months.map(() => 0))),
      row(["開発（新規）"].concat(months.map(() => 0)))
    ];
    await Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets;
      const old = wss.getItemOrNullObject(SEG_SHEET); await ctx.sync();
      if (!old.isNullObject) throw new Error("売上区分シートはすでにあります");
      const ws = wss.add(SEG_SHEET);
      const rg = ws.getRangeByIndexes(0, 0, g.length, W);
      rg.formulas = g;
      ws.getRangeByIndexes(5, 1, 2, 12).numberFormat = Array(2).fill(Array(12).fill("#,##0"));
      ws.getRangeByIndexes(9, 1, 2, 12).numberFormat = Array(2).fill(Array(12).fill("#,##0"));
      ws.getRange("A1").format.font.bold = true;
      ws.getRangeByIndexes(3, 0, 1, W).format.fill.color = "#E4EAF3";
      ws.getRangeByIndexes(3, 0, 1, W).format.font.bold = true;
      [4, 8].forEach(r => { ws.getRangeByIndexes(r, 0, 1, 1).format.font.bold = true; });
      [5, 6, 9, 10].forEach(r => { const x = ws.getRangeByIndexes(r, 1, 1, 12); x.format.fill.color = "#FFF4D6"; x.format.font.color = "#1F3F6E"; });
      try { ws.getRange("A:A").format.columnWidth = 120; ws.getRange("B:M").format.columnWidth = 78; } catch (e) {}
      ws.activate();
      await ctx.sync();
    });
  }

  /* ---------- 計算 ---------- */
  function compute(c) {
    const { M, fy, plan, seg, p } = c;
    const K = M.K, months = M.fyMonths(fy), last = M.latestIn(fy);
    const cur = last ? months.indexOf(last) + 1 : 0, r = 12 - cur;
    const pv = (key) => (plan && key && plan[key] ? plan[key].values.map(v => v || 0) : Array(12).fill(0));
    const pS = pv(K.sales), pC = pv(K.cogs), pG = pv(K.sga);
    const hasSeg = !!(seg && seg.exists && !seg.error);
    const sv = (src, k) => months.map(q => (hasSeg && seg[src][k][q]) || 0);
    const planM = sv("plan", "m"), planN = sv("plan", "n"), planE = pS.map((v, i) => v - planM[i] - planN[i]);
    const actM = sv("act", "m"), actN = sv("act", "n");
    const mv = (key, i) => (key && i < cur ? (M.month(key, months[i]) || 0) : 0);
    // 既定値：計画の残り月平均・計画原価率
    const avgRem = (arr) => (r > 0 ? sum(arr.slice(cur)) / r : 0);
    const ex = p.ex == null ? Math.round(avgRem(planE)) : p.ex;
    const nw = p.nw == null ? Math.round(avgRem(planN)) : p.nw;
    const crP = p.cr == null ? (sum(pS) ? sum(pC) / sum(pS) * 100 : 20) : p.cr;
    const cr = crP / 100, sr = p.sr / 100, bl = r > 0 ? p.bl : 0;
    const mon = { m: [], e: [], n: [] }, mc = [], mg = [], ms = [];
    for (let i = 0; i < 12; i++) {
      if (i < cur) {
        const s = mv(K.sales, i);
        mon.m.push(actM[i]); mon.n.push(actN[i]); mon.e.push(s - actM[i] - actN[i]);
        mc.push(mv(K.cogs, i)); mg.push(mv(K.sga, i)); ms.push(s);
      } else {
        const m = planM[i], e = bl / r + ex, n = nw;
        mon.m.push(m); mon.e.push(e); mon.n.push(n);
        mc.push((m + e + n) * cr); mg.push(pG[i] * sr); ms.push(m + e + n);
      }
    }
    const taxK = (10 / 110) * (1 - p.mr / 100), tg = p.tgt / 100;
    const adj = (R, C) => (R > 0 ? (R - C - R * taxK) / (R / 1.1) * 100 : 0);
    const Ract = sum(ms.slice(0, cur)), cAct = sum(mc.slice(0, cur)), gAct = sum(mg.slice(0, cur)), gRem = sum(mg.slice(cur));
    const maintRem = sum(planM.slice(cur)), firm = maintRem + bl, pipe = (ex + nw) * r, S = gAct + gRem;
    const R1 = Ract + firm, C1 = cAct + cr * firm, R2 = R1 + pipe, C2 = C1 + cr * pipe;
    const PR = sum(pS), PC = sum(pC), PG = sum(pG);
    return {
      months, cur, r, last, hasSeg, pS, pC, pG, planM, planN, planE, mon, mc, mg, ms,
      ex, nw, crP, cr, bl, taxK, tg, adj, Ract, cAct, gAct, gRem, S, maintRem, firm, pipe, R1, C1, R2, C2, PR, PC, PG,
      m1: adj(R1, C1 + S), m2: adj(R2, C2 + S), mP: adj(PR, PC + PG)
    };
  }

  /* ---------- SVG ---------- */
  function niceScale(min, max, ticks = 4) {
    if (min === max) max = min + 1;
    const raw = (max - min) / ticks, mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map(k => k * mag).find(s => s >= raw);
    return { lo: Math.floor(min / step) * step, hi: Math.ceil(max / step) * step, step };
  }
  const tick = (v) => (Math.abs(v) >= 1e4 ? (v / 1e4).toLocaleString() + "万" : Math.round(v).toLocaleString());
  function frame(H, mx, mn) {
    const W = 360, L = 46, R = 6, T = 12, B = 22;
    const sc = niceScale(Math.min(0, mn || 0), mx);
    const y = (v) => T + (H - T - B) * (1 - (v - sc.lo) / (sc.hi - sc.lo));
    const s = [];
    for (let v = sc.lo; v <= sc.hi + 1e-9; v += sc.step) {
      s.push(`<line x1="${L}" x2="${W - R}" y1="${y(v).toFixed(1)}" y2="${y(v).toFixed(1)}" stroke="${v === 0 ? "#B9C0C7" : COL.grid}"/>`);
      s.push(`<text x="${L - 4}" y="${(y(v) + 3.5).toFixed(1)}" font-size="9" fill="${COL.muted}" text-anchor="end">${tick(v)}</text>`);
    }
    return { W, L, R, T, B, H, y, s };
  }
  const marker = (f, x) => `<line x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${f.T}" y2="${f.H - f.B}" stroke="${COL.muted}" stroke-dasharray="2 2"/><text x="${(x + 3).toFixed(1)}" y="${f.T + 8}" font-size="9" fill="${COL.muted}">見込み →</text>`;

  // 累計の積み上げ面＋線
  function areaChart(o) {
    const all = [];
    o.layers.forEach(l => l.top.forEach(v => v != null && all.push(v)));
    o.lines.forEach(l => l.values.forEach(v => v != null && all.push(v)));
    const f = frame(o.height || 200, Math.max(1, ...all));
    const n = o.labels.length, step = (f.W - f.L - f.R) / n, cx = (i) => f.L + step * i + step / 2;
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="${esc(o.label)}">`, ...f.s];
    let prev = Array(n).fill(0);
    o.layers.forEach(l => {
      const idx = l.top.map((v, i) => (v != null ? i : -1)).filter(i => i >= 0);
      if (idx.length >= 2) {
        const pts = idx.map(i => cx(i).toFixed(1) + "," + f.y(l.top[i]).toFixed(1)).concat(idx.slice().reverse().map(i => cx(i).toFixed(1) + "," + f.y(prev[i]).toFixed(1)));
        s.push(`<polygon points="${pts.join(" ")}" fill="${l.fill}" stroke="${l.col}" stroke-width="0.8"/>`);
      } else if (idx.length === 1) {
        const i = idx[0]; s.push(`<rect x="${(cx(i) - 3).toFixed(1)}" y="${f.y(l.top[i]).toFixed(1)}" width="6" height="${(f.y(prev[i]) - f.y(l.top[i])).toFixed(1)}" fill="${l.fill}"/>`);
      }
      prev = prev.map((v, i) => (l.top[i] != null ? l.top[i] : v));
    });
    o.lines.forEach(l => {
      let seg = [];
      const flush = () => { if (seg.length > 1) s.push(`<polyline points="${seg.join(" ")}" fill="none" stroke="${l.color}" stroke-width="${l.width || 1.5}" ${l.dash ? `stroke-dasharray="${l.dash}"` : ""} stroke-linejoin="round"/>`); seg = []; };
      l.values.forEach((v, i) => { if (v == null) flush(); else seg.push(cx(i).toFixed(1) + "," + f.y(v).toFixed(1)); });
      flush();
    });
    if (o.marker != null && o.marker >= 0 && o.marker < n - 1) s.push(marker(f, f.L + step * (o.marker + 1)));
    o.labels.forEach((t, i) => s.push(`<text x="${cx(i).toFixed(1)}" y="${f.H - 7}" font-size="9.5" fill="${COL.muted}" text-anchor="middle">${esc(t)}</text>`));
    s.push("</svg>");
    return s.join("");
  }

  // 月別の棒（原価・販管費）＋計画の横線
  function costChart(o) {
    const all = [];
    o.bars.forEach(b => { b.values.forEach(v => all.push(v || 0)); b.plan.forEach(v => all.push(v || 0)); });
    const f = frame(o.height || 190, Math.max(1, ...all));
    const n = o.labels.length, step = (f.W - f.L - f.R) / n, bw = step * 0.36;
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="${esc(o.label)}">`, ...f.s];
    o.bars.forEach((b, bi) => {
      b.values.forEach((v, i) => {
        const x = f.L + step * i + step * 0.14 + bw * bi, y0 = f.y(0), y1 = f.y(v || 0), h = Math.max(1, y0 - y1);
        const fc = i >= o.cur;
        const fill = fc ? "#FFFFFF" : (b.spike[i] ? COL.spike : b.color);
        s.push(`<rect x="${x.toFixed(1)}" y="${y1.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" rx="1.5" fill="${fill}" ${fc ? `stroke="${b.color}" stroke-width="1" stroke-dasharray="2 1.5"` : ""}><title>${esc(o.labels[i])}月 ${b.name}：${yen(v)}円（計画 ${yen(b.plan[i])}円）</title></rect>`);
        const py = f.y(b.plan[i] || 0);
        s.push(`<line x1="${(x - 1).toFixed(1)}" x2="${(x + bw + 1).toFixed(1)}" y1="${py.toFixed(1)}" y2="${py.toFixed(1)}" stroke="#1B2430" stroke-width="1.6"/>`);
      });
    });
    if (o.cur > 0 && o.cur < n) s.push(marker(f, f.L + step * o.cur));
    o.labels.forEach((t, i) => s.push(`<text x="${(f.L + step * i + step / 2).toFixed(1)}" y="${f.H - 7}" font-size="9.5" fill="${COL.muted}" text-anchor="middle">${esc(t)}</text>`));
    s.push("</svg>");
    return s.join("");
  }

  // 年度末の着地（計画／受注確定ベース／着地見込み × 売上・コスト）
  function landingChart(o) {
    const groups = o.groups, all = [];
    groups.forEach(g => { all.push(sum(g.rev.map(x => x.v)), sum(g.cost.map(x => x.v))); });
    const f = frame(o.height || 220, Math.max(1, ...all) * 1.12);
    const n = groups.length, step = (f.W - f.L - f.R) / n, bw = step * 0.3;
    const pat = (id, c) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="6" height="6" fill="#FFFFFF"/><line x1="0" y1="0" x2="0" y2="6" stroke="${c}" stroke-width="2.2"/></pattern>`;
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="${esc(o.label)}">`,
      `<defs>${pat("yjHf", COL.firm)}${pat("yjHc", COL.cogs)}${pat("yjHg", COL.sga)}</defs>`, ...f.s];
    groups.forEach((g, gi) => {
      [["rev", 0.12], ["cost", 0.12 + 0.3 + 0.04]].forEach(([k, off]) => {
        let base = 0;
        const x = f.L + step * gi + step * off;
        g[k].forEach(seg => {
          if (!seg.v) return;
          const y1 = f.y(base + seg.v), y0 = f.y(base);
          s.push(`<rect x="${x.toFixed(1)}" y="${y1.toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(0.5, y0 - y1).toFixed(1)}" fill="${seg.hatch ? `url(#${seg.hatch})` : seg.color}" ${seg.hatch ? `stroke="${seg.color}" stroke-width="0.8"` : ""}><title>${esc(seg.name)}：${yen(seg.v)}円</title></rect>`);
          base += seg.v;
        });
      });
      const top = f.y(Math.max(sum(g.rev.map(x => x.v)), sum(g.cost.map(x => x.v))));
      const cx = f.L + step * gi + step / 2;
      s.push(`<text x="${cx.toFixed(1)}" y="${(top - 14).toFixed(1)}" font-size="10.5" font-weight="700" text-anchor="middle" fill="${g.ok ? COL.ok : COL.bad}">${g.m.toFixed(1)}%</text>`);
      s.push(`<text x="${cx.toFixed(1)}" y="${(top - 3).toFixed(1)}" font-size="8" text-anchor="middle" fill="${COL.muted}">営業利益 ${tick(g.op)}</text>`);
      s.push(`<text x="${cx.toFixed(1)}" y="${f.H - 7}" font-size="10" text-anchor="middle" fill="${COL.muted}">${esc(g.label)}</text>`);
    });
    s.push("</svg>");
    return s.join("");
  }

  /* ---------- 画面 ---------- */
  let last = null; // 最後に描いた文脈（スライダー操作時の再描画用）
  const lgBox = (c) => `<i class="lg" style="background:${c}"></i>`;
  const lgLine = (c, dash) => `<i class="lg lg-line" style="border-top:2px ${dash ? "dashed" : "solid"} ${c}"></i>`;
  const lgHatch = (c) => `<i class="lg" style="border:1px solid ${c};background:repeating-linear-gradient(135deg,${c} 0 1.5px,#fff 1.5px 4px)"></i>`;
  const badge = (kind, t) => `<span class="yj-st yj-${kind}">${esc(t)}</span>`;

  function sliders(c, d) {
    const p = c.p, mx = Math.max(1, ...d.pS) * 2, step = 50000;
    const rng = (id, label, val, min, max, st, fmt) => `<div class="yj-rng"><label for="yj-${id}">${label}<b id="yjv-${id}">${fmt(val)}</b></label><input type="range" id="yj-${id}" data-yj="${id}" min="${min}" max="${max}" step="${st}" value="${val}"></div>`;
    const y = (v) => yen(v) + "円";
    return `<section class="card yj-ctl">
  <div class="card-head"><h2>見込みの前提</h2><button type="button" class="link" data-yact="reset">既定に戻す</button></div>
  <p class="muted small">${d.cur ? `${Number(d.last.slice(5))}月まで実績（${d.cur}/12か月）。残り${d.r}か月を見込みで計算します。` : "実績はまだありません。12か月すべてを見込みで計算します。"}</p>
  ${d.r > 0 ? rng("bl", "受注残（受注済・未計上）", p.bl, 0, Math.max(d.PR, p.bl), step, y) : ""}
  ${d.r > 0 ? rng("ex", "開発（既存）追加見込み／月", d.ex, 0, Math.max(mx, d.ex), step, y) : ""}
  ${d.r > 0 ? rng("nw", "開発（新規）見込み／月", d.nw, 0, Math.max(mx, d.nw), step, y) : ""}
  ${d.r > 0 ? rng("cr", "今後の原価率", Math.round(d.crP * 2) / 2, 0, 60, 0.5, v => Number(v).toFixed(1) + "%") : ""}
  ${d.r > 0 ? rng("sr", "販管費（計画比）", p.sr, 80, 120, 1, v => v + "%") : ""}
</section>`;
  }

  function body(c, d) {
    const p = c.p, M = c.M, labels = d.months.map(q => String(Number(q.slice(5)))), j = d.cur - 1;
    const open = (k) => (p.open[k] ? "open" : "");
    const A = (arr, act) => arr.map((v, i) => (act ? (i < d.cur ? v : null) : (i >= Math.max(0, d.cur - 1) ? v : null)));
    // ① 売上
    const on = KEYS.filter(k => p.sel[k] && d.hasSeg), off = KEYS.filter(k => !on.includes(k));
    const planOf = { m: d.planM, e: d.planE, n: d.planN };
    const layersDef = on.map(k => ({ name: CAT[k].name, keys: [k], col: CAT[k].col, fill: CAT[k].fill }));
    if (off.length) layersDef.push({ name: on.length ? "その他（" + off.map(k => CAT[k].name).join("・") + "）" : "売上高", keys: off, col: on.length ? "#8A939C" : CAT.e.col, fill: on.length ? "rgba(138,147,156,0.5)" : "rgba(31,63,110,0.55)" });
    let run = Array(12).fill(0), runP = Array(12).fill(0);
    const layers = [], lines = [];
    layersDef.forEach((ly, li) => {
      const c1 = cum(d.mon.m.map((_, i) => ly.keys.reduce((t, k) => t + d.mon[k][i], 0)));
      const cp = cum(d.mon.m.map((_, i) => ly.keys.reduce((t, k) => t + planOf[k][i], 0)));
      run = run.map((v, i) => v + c1[i]); runP = runP.map((v, i) => v + cp[i]);
      const top = run.slice(), topP = runP.slice(), lastL = li === layersDef.length - 1;
      layers.push({ top: A(top, true), col: ly.col, fill: ly.fill });
      if (d.r > 0) lines.push({ values: A(top, false), color: lastL ? COL.fc : ly.col, width: lastL ? 2.2 : 1.3, dash: "5 3" });
      lines.push({ values: topP, color: COL.plan, width: lastL ? 1.5 : 1, dash: "2 2" });
    });
    const end = run[11], goal = d.PR, gap = goal - end, pNow = d.cur ? runP[j] : 0;
    const st1 = gap <= 0 ? badge("ok", "達成見込み") : badge(gap < goal * 0.06 ? "warn" : "bad", "不足 " + yen(gap) + "円");
    const segBtns = d.hasSeg
      ? `<div class="toggles" role="group" aria-label="色分け">${KEYS.map(k => `<button type="button" class="tgl ${p.sel[k] ? "on" : ""}" data-yact="sel" data-k="${k}" aria-pressed="${!!p.sel[k]}">${CAT[k].name}</button>`).join("")}</div>`
      : `<p class="muted small">保守・開発（既存）・開発（新規）で色分けするには「売上区分」シートが必要です。<button type="button" class="link" data-yact="seg-create">売上区分シートを作成</button></p>`;
    const catCards = d.hasSeg ? KEYS.map(k => {
      const cm = cum(d.mon[k]), a = d.cur ? cm[j] : 0, pn = d.cur ? sum(planOf[k].slice(0, d.cur)) : 0, an = sum(planOf[k]), ef = cm[11];
      const r = pn ? a / pn * 100 : null, need = d.r > 0 ? Math.max(0, (an - a) / d.r) : 0;
      return `<div class="yj-cat"><div class="strong">${lgBox(CAT[k].col)} ${CAT[k].name}</div>
        <div class="yj-kv"><span>実績累計</span><b>${yen(a)}円</b></div>
        <div class="yj-kv"><span>同月計画</span><b>${yen(pn)}円</b></div>
        <div class="yj-kv"><span>計画比</span><b class="${r == null ? "" : r >= 97 ? "up-good" : "up-bad"}">${r == null ? "—" : r.toFixed(1) + "%"}</b></div>
        <div class="yj-kv"><span>年度末見込み</span><b>${yen(ef)}円</b></div>
        <div class="yj-kv"><span>年間計画</span><b>${yen(an)}円</b></div>
        <div class="yj-kv"><span>達成に必要／月</span><b>${yen(need)}円</b></div></div>`;
    }).join("") : "";
    const sec1 = `<section class="card">
  <div class="card-head"><h2>① 売上</h2>${st1}</div>
  <p class="muted small">目的：売上目標（${yen(goal)}円）を達成する。計画どおり積み上がっているか。</p>
  ${segBtns}
  ${areaChart({ label: "売上の累計", labels, layers, lines, marker: d.r > 0 ? j : null })}
  <div class="legend">${layersDef.map(l => `<span>${lgBox(l.col)}${esc(l.name)}</span>`).join("")}${d.r > 0 ? `<span>${lgLine(COL.fc, true)}見込み</span>` : ""}<span>${lgLine(COL.plan, true)}計画</span></div>
  <div class="yj-box yj-${gap > 0 ? "warn" : "ok"}">現時点 <b>${yen(d.cur ? run[j] : 0)}円</b>（同月計画 ${yen(pNow)}円の <b>${pNow ? Math.round((d.cur ? run[j] : 0) / pNow * 100) : "—"}%</b>）<br>年度末見込み <b>${yen(end)}円</b>${gap > 0 ? (d.r > 0 ? `　→ 残り${d.r}か月で <b>あと${yen(gap)}円</b>（月 +${yen(gap / d.r)}円）の追加受注が必要` : `　→ 目標に ${yen(gap)}円 届きませんでした`) : "　→ 目標達成の見込み"}</div>
  ${d.hasSeg ? `<details data-yjd="cat" ${open("cat")}><summary>区分別の詳細（保守・開発（既存）・開発（新規））</summary><div class="yj-cats">${catCards}</div><p class="muted small">開発（既存）＝売上高−保守−開発（新規）。<button type="button" class="link" data-yact="seg-open">売上区分シートを開く</button></p></details>` : ""}
</section>`;

    // ② コスト
    const sp = p.sp;
    const isSp = (v, pl, pc, pa) => v > pl * pc / 100 && v - pl > pa;
    const spC = d.mc.map((v, i) => i < d.cur && isSp(v, d.pC[i], sp.cp, sp.ca));
    const spG = d.mg.map((v, i) => i < d.cur && isSp(v, d.pG[i], sp.sp, sp.sa));
    const list = [];
    for (let i = 0; i < d.cur; i++) {
      if (spC[i]) list.push(`${labels[i]}月 原価 ${yen(d.mc[i])}円（計画比 ${sgn(d.mc[i] - d.pC[i])}円）`);
      if (spG[i]) list.push(`${labels[i]}月 販管費 ${yen(d.mg[i])}円（計画比 ${sgn(d.mg[i] - d.pG[i])}円）`);
    }
    const sC = d.cAct, sCP = sum(d.pC.slice(0, d.cur)), sG = d.gAct, sGP = sum(d.pG.slice(0, d.cur));
    const over = (sCP && sC > sCP * 1.05) || (sGP && sG > sGP * 1.05);
    const k2 = list.length ? "bad" : over ? "warn" : "ok";
    const rt = (a, b) => (b ? Math.round(a / b * 100) + "%" : "—");
    const sec2 = `<section class="card">
  <div class="card-head"><h2>② コスト（原価・販管費）</h2>${badge(k2, list.length ? "突発 " + list.length + "件" : over ? "計画超過" : "計画内")}</div>
  <p class="muted small">目的：費用を計画の範囲内に収める。計画どおりか、突発が発生していないか。</p>
  ${costChart({ label: "原価・販管費の月別", labels, cur: d.cur, bars: [{ name: "原価", color: COL.cogs, values: d.mc, plan: d.pC, spike: spC }, { name: "販管費", color: COL.sga, values: d.mg, plan: d.pG, spike: spG }] })}
  <div class="legend"><span>${lgBox(COL.cogs)}原価</span><span>${lgBox(COL.sga)}販管費</span><span>${lgBox(COL.spike)}突発</span><span><i class="lg" style="border:1px dashed ${COL.cogs};background:#fff"></i>見込み</span><span>${lgLine("#1B2430")}計画</span></div>
  <div class="yj-box yj-${k2}">累計　原価 <b>${yen(sC)}円</b>／計画 ${yen(sCP)}円（${rt(sC, sCP)}）<br>販管費 <b>${yen(sG)}円</b>／計画 ${yen(sGP)}円（${rt(sG, sGP)}）<br>${list.length ? "<b>突発</b>：" + list.map(esc).join("、") : "突発なし"}</div>
  <details data-yjd="spk" ${open("spk")}><summary>突発の条件</summary>
    <div class="yj-spk"><span></span><span>計画比（%超）</span><span>超過額（円超）</span>
    <span>原価</span><input type="number" class="inp" data-yjc="cp" value="${sp.cp}" min="100" step="5"><input type="number" class="inp" data-yjc="ca" value="${sp.ca}" min="0" step="10000">
    <span>販管費</span><input type="number" class="inp" data-yjc="sp" value="${sp.sp}" min="100" step="5"><input type="number" class="inp" data-yjc="sa" value="${sp.sa}" min="0" step="10000"></div>
    <p class="muted small">両方の条件を満たした実績月を突発として赤で示します。</p></details>
</section>`;

    // ③ 利益率（年度末の着地）
    const tgtP = p.tgt, okOf = (m) => m >= tgtP;
    const groups = [
      { label: "計画", m: d.mP, ok: okOf(d.mP), op: d.PR - d.PC - d.PG, rev: [{ name: "計画売上", v: d.PR, color: COL.planBar }], cost: [{ name: "原価（計画）", v: d.PC, color: COL.cogs }, { name: "販管費（計画）", v: d.PG, color: COL.sga }] },
      { label: "受注確定ベース", m: d.m1, ok: okOf(d.m1), op: d.R1 - d.C1 - d.S, rev: [{ name: "売上 実績", v: d.Ract, color: COL.act }, { name: "保守＋受注残", v: d.firm, color: COL.firm }], cost: [{ name: "原価 実績", v: d.cAct, color: COL.cogs }, { name: "原価見込み", v: d.C1 - d.cAct, color: COL.cogs, hatch: "yjHc" }, { name: "販管費 実績", v: d.gAct, color: COL.sga }, { name: "販管費見込み", v: d.gRem, color: COL.sga, hatch: "yjHg" }] },
      { label: "着地見込み", m: d.m2, ok: okOf(d.m2), op: d.R2 - d.C2 - d.S, rev: [{ name: "売上 実績", v: d.Ract, color: COL.act }, { name: "保守＋受注残", v: d.firm, color: COL.firm }, { name: "営業見込み", v: d.pipe, color: COL.firm, hatch: "yjHf" }], cost: [{ name: "原価 実績", v: d.cAct, color: COL.cogs }, { name: "原価見込み", v: d.C2 - d.cAct, color: COL.cogs, hatch: "yjHc" }, { name: "販管費 実績", v: d.gAct, color: COL.sga }, { name: "販管費見込み", v: d.gRem, color: COL.sga, hatch: "yjHg" }] }
    ];
    const ok = okOf(d.m2);
    const k3 = ok ? (okOf(d.m1) ? "ok" : "warn") : "bad";
    const kk = 1 - d.cr - d.taxK - d.tg / 1.1, Rneed = kk > 0 ? (d.cAct - d.cr * d.Ract + d.S) / kk : Infinity, add = Rneed - d.R2;
    // 今後の案件で許容できる原価率の上限
    const futR = d.firm + d.pipe, allow = (d.R2 - d.S - d.R2 * d.taxK - d.tg * d.R2 / 1.1) - d.cAct;
    let capHtml = "", capK = "ok";
    if (d.r > 0 && futR > 0) {
      if (allow <= 0) { capK = "bad"; capHtml = `<b class="yj-big">確保不可</b><span class="small">売上見込みと販管費だけで${tgtP}%を下回ります。売上の上積みか販管費の見直しが必要です。</span>`; }
      else {
        const cap = allow / futR * 100, room = allow - futR * d.cr;
        capK = d.crP <= cap ? (cap - d.crP < 3 ? "warn" : "ok") : "bad";
        capHtml = `<b class="yj-big">${cap.toFixed(1)}%</b><div class="yj-meter"><div style="width:${Math.max(0, Math.min(100, cap / 60 * 100)).toFixed(1)}%"></div><i style="left:${Math.min(100, d.crP / 60 * 100).toFixed(1)}%"></i></div><span class="small">今後の売上 ${yen(futR)}円（受注残＋見込み）に対し、原価は <b>${yen(allow)}円以内</b>。今の想定 ${d.crP.toFixed(1)}% では${room >= 0 ? ` あと <b>${yen(room)}円</b> の余裕` : ` <b>${yen(-room)}円</b> 超過`}。</span>`;
      }
    }
    const sec3 = `<section class="card">
  <div class="card-head"><h2>③ 利益率（年度末の着地）</h2>${badge(k3, ok ? (okOf(d.m1) ? "受注確定で" + tgtP + "%確保" : "見込み込みで" + tgtP + "%") : "着地 " + d.m2.toFixed(1) + "%")}</div>
  <p class="muted small">目的：調整後営業利益率（税抜）${tgtP}%以上を確保する。受注残・営業見込みを含めて達成できるか、計画との乖離。</p>
  <div class="yj-kpis">
    <div class="kpi"><span class="kpi-l">受注確定ベース</span><span class="kpi-v ${okOf(d.m1) ? "up-good" : "up-bad"}">${d.m1.toFixed(1)}<small>%</small></span><span class="kpi-c">実績＋保守＋受注残</span></div>
    <div class="kpi"><span class="kpi-l">着地見込み</span><span class="kpi-v ${ok ? "up-good" : "up-bad"}">${d.m2.toFixed(1)}<small>%</small></span><span class="kpi-c">＋営業見込み</span></div>
    <div class="kpi"><span class="kpi-l">計画</span><span class="kpi-v">${d.mP.toFixed(1)}<small>%</small></span><span class="kpi-c">計画シート</span></div>
  </div>
  ${landingChart({ label: "年度末の着地", groups })}
  <div class="legend"><span>${lgBox(COL.act)}売上 実績</span><span>${lgBox(COL.firm)}保守＋受注残</span><span>${lgHatch(COL.firm)}営業見込み</span><span>${lgBox(COL.planBar)}計画売上</span><span>${lgBox(COL.cogs)}原価</span><span>${lgHatch(COL.cogs)}原価見込み</span><span>${lgBox(COL.sga)}販管費</span><span>${lgHatch(COL.sga)}販管費見込み</span></div>
  <div class="yj-box yj-${ok ? "ok" : "bad"}">計画との乖離（年度末）：売上 <b>${sgn(d.R2 - d.PR)}円</b>、コスト <b>${sgn(d.C2 + d.S - d.PC - d.PG)}円</b>、利益率 <b>${(d.m2 - d.mP >= 0 ? "+" : "−") + Math.abs(d.m2 - d.mP).toFixed(1)}pt</b><br>${ok ? `着地見込みで${tgtP}%を確保（受注確定ベースは ${d.m1.toFixed(1)}%、営業見込み ${yen(d.pipe)}円の受注が前提）` : isFinite(add) && d.r > 0 ? `${tgtP}%確保には、営業見込みに加えて <b>あと${yen(add)}円</b> の受注が必要（原価率 ${d.crP.toFixed(1)}%前提）` : `${tgtP}%に届きません`}</div>
  ${capHtml ? `<div class="yj-box yj-${capK}"><span class="small">今後の案件で許容できる原価率の上限（${tgtP}%確保ライン）</span>${capHtml}</div>` : ""}
  <details data-yjd="tax" ${open("tax")}><summary>利益率の計算条件</summary>
    <div class="yj-spk"><span>目標（%）</span><input type="number" class="inp" data-yjc="tgt" value="${p.tgt}" min="0" step="0.5"><span></span>
    <span>みなし仕入率（%）</span><input type="number" class="inp" data-yjc="mr" value="${p.mr}" min="0" max="90" step="10"><span></span></div>
    <p class="muted small">調整後営業利益＝営業利益−消費税の納付見込み（売上×10/110×（1−みなし仕入率））。利益率は税抜売上（売上÷1.1）に対する割合です。簡易課税を前提にしています。</p></details>
</section>`;
    return sec1 + sec2 + sec3;
  }

  function render(c) {
    const M = c.M;
    if (!M.latest) return c.empty;
    if (!c.plan || !c.plan[M.K.sales]) {
      return `<div class="toolbar">${c.fySelect}</div><div class="note">${esc(M.fyLabel(c.fy))}の計画シートがありません。予実は計画シートの売上高・売上原価・販管費と比べます。計画タブで計画シートを作成してください。<br><button type="button" class="btn primary" data-act="tab" data-tab="plan" style="margin-top:8px">計画タブを開く</button></div>`;
    }
    if (c.seg && c.seg.error) c.segError = c.seg.error;
    const d = compute(c);
    last = { c, d };
    return `<div class="toolbar">${c.fySelect}</div>
${c.segError ? `<div class="banner err">${esc(c.segError)}</div>` : ""}
${sliders(c, d)}
<div id="yj-out" class="yj-out">${body(c, d)}</div>`;
  }
  function refresh() {
    if (!last) return;
    const d = compute(last.c);
    last.d = d;
    const o = document.getElementById("yj-out");
    if (o) o.innerHTML = body(last.c, d);
  }

  /* ---------- 操作 ---------- */
  document.addEventListener("input", (e) => {
    const t = e.target, id = t.dataset && t.dataset.yj;
    if (!id || !last) return;
    const p = last.c.p, v = Number(t.value);
    p[id] = v;
    const lab = document.getElementById("yjv-" + id);
    if (lab) lab.textContent = id === "cr" ? v.toFixed(1) + "%" : id === "sr" ? v + "%" : yen(v) + "円";
    saveParams(p); refresh();
  });
  document.addEventListener("change", (e) => {
    const t = e.target, id = t.dataset && t.dataset.yjc;
    if (!id || !last) return;
    const p = last.c.p, v = Number(t.value);
    if (!isFinite(v)) return;
    if (["cp", "ca", "sp", "sa"].includes(id)) p.sp[id] = v; else p[id] = v;
    saveParams(p); refresh();
  });
  document.addEventListener("toggle", (e) => {
    const t = e.target;
    if (!t || !t.dataset || !t.dataset.yjd || !last) return;
    last.c.p.open[t.dataset.yjd] = t.open; saveParams(last.c.p);
  }, true);
  document.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-yact]");
    if (!b || !last) return;
    const c = last.c, p = c.p;
    switch (b.dataset.yact) {
      case "sel": p.sel[b.dataset.k] = !p.sel[b.dataset.k]; saveParams(p); b.classList.toggle("on", p.sel[b.dataset.k]); b.setAttribute("aria-pressed", p.sel[b.dataset.k]); refresh(); break;
      case "reset": Object.assign(p, { bl: 0, ex: null, nw: null, cr: null, sr: 100 }); saveParams(p); c.rerender(); break;
      case "seg-open": try { await MonthlySheets.activateByPrefix(SEG_SHEET); } catch (err) { c.toast("シートを開けませんでした", true); } break;
      case "seg-create":
        try { await createSegments(c.M, c.fy); c.toast("売上区分シートを作成しました。保守と開発（新規）の金額を入れて再読み込みしてください"); await c.reload(); }
        catch (err) { console.error(err); c.toast("売上区分シートを作れませんでした：" + (err.message || err), true); }
        break;
    }
  });

  global.Yojitsu = { render, refresh, readSegments, createSegments, loadParams, saveParams, compute, SEG_SHEET };
})(window);
