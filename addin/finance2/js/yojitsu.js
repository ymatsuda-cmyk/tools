/* ============================================================
 * yojitsu.js — 予実タブ（①売上 ②コスト ③利益率）
 * ------------------------------------------------------------
 * ①売上   ：累計の積み上げ。全体・保守・開発（既存）・開発（新規）のタイルで区分を切り替え、
 *            選んだ区分の実績（濃い面）・受注確定（薄い面）・パイプライン込み（点線）・計画（破線）を比べる。
 * ②コスト ：原価＋販管費の月別積み上げ棒。計画を横線で重ね、突発（その月までの累計が計画累計の○%超）を赤で示す。
 * どのグラフもカーソルを当てる（タッチはタップ）と、その月・その棒の内訳（実績／計画／差）を表示する。
 * ③利益率 ：年度末の着地。受注確定ベース（実績＋月次PLの見込み月の売上）と着地見込み（＋パイプライン）を
 *            計画と並べ、調整後営業利益率（税抜）を目標と比べる。見込み部分は斜線。
 *
 * 区分の内訳は 計画シート・月次PL の売上高の下の「保守」「開発(既存)」「開発(新規)」の行から読む
 * （行がない月は売上高をすべて開発（既存）として扱う。内訳の合計と売上高の差は開発（既存）で調整して表示）。
 * 見込み月：売上は月次PLの売上高（受注確定）とパイプラインの行。原価・販管費は月次PLに値があればその値、
 *   なければ計画シートの値。パイプラインの原価は売上原価の下の「パイプライン」の行。表示の設定はブックに保存する。
 * 実績の確定月：月次PL／月次BSの4行目に印（空でない値）がある月まで実績、それ以降は見込み。
 *   印がなければ、データのある最新月までを実績とする。
 * ============================================================ */
(function (global) {
  "use strict";
  const SKEY = "finance2.yojitsu";
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const yen = (n) => (n == null || !isFinite(n) ? "—" : (n < 0 ? "−" : "") + Math.abs(Math.round(n)).toLocaleString());
  const sgn = (n) => (n == null || !isFinite(n) ? "—" : (n > 0 ? "+" : n < 0 ? "−" : "±") + Math.abs(Math.round(n)).toLocaleString());
  const sum = (a) => a.reduce((t, v) => t + (v || 0), 0);
  const cum = (a) => { let s = 0; return a.map(v => (s += v || 0)); };

  const CAT = {
    m: { name: "保守", short: "保守", col: "#2E8B6E", fill: "rgba(46,139,110,0.75)" },
    e: { name: "開発（既存）", short: "開発既存", col: "#1F3F6E", fill: "rgba(31,63,110,0.75)" },
    n: { name: "開発（新規）", short: "開発新規", col: "#6A5BB8", fill: "rgba(106,91,184,0.75)" }
  };
  const KEYS = ["m", "e", "n"];
  const COL = { ink: "#1B2430", plan: "#A9B3BF", fc: "#C9781A", pipe: "#8E44AD", cogs: "#7B848D", sga: "#D98A6A", spike: "#C0392B", firm: "#8FA9CC", act: "#1F3F6E", planBar: "#C9D5E6", ok: "#2E7D6B", bad: "#8E2A2A", muted: "#56606B", grid: "#E3E6E4" };

  /* ---------- 設定（ブックに保存） ---------- */
  const DEF = { tgt: 15, mr: 50, sp: { cp: 120, sp: 120 }, seg: "all", open: {} };
  function loadParams() {
    let p = null;
    try { const s = Office.context.document.settings.get(SKEY); if (s) p = JSON.parse(s); } catch (e) {}
    if (!p) { try { p = JSON.parse(localStorage.getItem(SKEY) || "null"); } catch (e) {} }
    const o = JSON.parse(JSON.stringify(DEF));
    if (p) { Object.assign(o, p); o.sp = { cp: (p.sp && p.sp.cp) || DEF.sp.cp, sp: (p.sp && p.sp.sp) || DEF.sp.sp }; o.seg = ["all", "m", "e", "n"].includes(p.seg) ? p.seg : "all"; delete o.sel; o.open = Object.assign({}, p.open || {}); }
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

  /* ---------- 計算 ---------- */
  function compute(c) {
    const { M, fy, plan, segPlan, segAct, p, confirmed } = c;
    const K = M.K, months = M.fyMonths(fy);
    let cur, byMark = false;
    if (confirmed) {
      byMark = true;
      cur = confirmed < months[0] ? 0 : confirmed > months[11] ? 12 : months.indexOf(confirmed) + 1;
    } else { const lt = M.latestIn(fy); cur = lt ? months.indexOf(lt) + 1 : 0; }
    const last = cur ? months[cur - 1] : null, r = 12 - cur;
    const pv = (key) => (plan && key && plan[key] ? plan[key].values.map(v => v || 0) : Array(12).fill(0));
    const pS = pv(K.sales), pC = pv(K.cogs), pG = pv(K.sga);
    const hasSeg = !!(segPlan && (segPlan.m || segPlan.e || segPlan.n));
    const planM = months.map((_, i) => (hasSeg && segPlan.m ? segPlan.m[i] || 0 : 0));
    const planN = months.map((_, i) => (hasSeg && segPlan.n ? segPlan.n[i] || 0 : 0));
    const planE = months.map((_, i) => (hasSeg && segPlan.e && segPlan.e[i] != null ? segPlan.e[i] : pS[i] - planM[i] - planN[i]));
    // 計画の内訳の合計と売上高の差
    const planDiff = months.map((_, i) => pS[i] - planM[i] - planE[i] - planN[i]).map((v, i) => ({ i, v })).filter(x => Math.abs(x.v) >= 1);
    const av = (k, q) => (segAct && segAct[k] && typeof segAct[k][q] === "number" ? segAct[k][q] : null);
    const pa = (k, q) => (c.pipeAct && c.pipeAct[k] && typeof c.pipeAct[k][q] === "number" ? c.pipeAct[k][q] : 0);
    const pca = (k, q) => (c.pipeCost && c.pipeCost[k] && typeof c.pipeCost[k][q] === "number" ? c.pipeCost[k][q] : 0);
    const mv = (key, i) => (key && i < cur ? (M.month(key, months[i]) || 0) : 0);
    const sheetVal = (key, i) => { if (!key || !M.has(months[i])) return null; const v = M.month(key, months[i]); return v ? v : null; };
    const PR = sum(pS), PC = sum(pC), PG = sum(pG);
    const crP = PR ? PC / PR * 100 : 0, cr = crP / 100;   // 計画の原価率（目標に届かないときの上積み額の計算に使う）
    const pipeNoCost = [];
    const mon = { m: [], e: [], n: [] }, pmon = { m: [], e: [], n: [] }, mc = [], mg = [], ms = [], pc = [], actDiff = [], actMissing = [];
    const split = (s, q, isAct, i) => {
      const am = av("m", q), ae = av("e", q), an = av("n", q);
      let m = 0, e = s, n = 0;
      if (am != null || ae != null || an != null) {
        m = am || 0; n = an || 0; e = ae == null ? s - m - n : ae;
        const df = s - m - e - n;
        if (Math.abs(df) >= 1) { if (isAct || s) actDiff.push({ i, v: df }); e += df; }
      } else if (isAct && hasSeg && s) actMissing.push(i);
      return { m, e, n };
    };
    for (let i = 0; i < 12; i++) {
      const q = months[i];
      if (i < cur) {
        const s = mv(K.sales, i), x = split(s, q, true, i);
        mon.m.push(x.m); mon.e.push(x.e); mon.n.push(x.n);
        pmon.m.push(0); pmon.e.push(0); pmon.n.push(0);
        mc.push(mv(K.cogs, i)); mg.push(mv(K.sga, i)); ms.push(s); pc.push(0);
      } else {
        // 見込み月：受注確定＝月次PLの売上高、パイプライン＝パイプラインの行
        const s = sheetVal(K.sales, i) || 0, x = split(s, q, false, i);
        mon.m.push(x.m); mon.e.push(x.e); mon.n.push(x.n);
        const pm = pa("m", q), pe = pa("e", q), pn = pa("n", q);
        pmon.m.push(pm); pmon.e.push(pe); pmon.n.push(pn);
        const sc = sheetVal(K.cogs, i), sg = sheetVal(K.sga, i);
        mc.push(sc != null ? sc : pC[i]); mg.push(sg != null ? sg : pG[i]); ms.push(s);
        const pcost = pca("m", q) + pca("e", q) + pca("n", q);
        pc.push(pcost);
        if (pm + pe + pn && !pcost) pipeNoCost.push(i);
      }
    }
    const taxK = (10 / 110) * (1 - p.mr / 100), tg = p.tgt / 100;
    const adj = (R, C) => (R > 0 ? (R - C - R * taxK) / (R / 1.1) * 100 : 0);
    const Ract = sum(ms.slice(0, cur)), cAct = sum(mc.slice(0, cur)), gAct = sum(mg.slice(0, cur)), gRem = sum(mg.slice(cur));
    const firm = sum(ms.slice(cur)), cFirm = sum(mc.slice(cur));
    const pipe = sum(pmon.m) + sum(pmon.e) + sum(pmon.n), cPipe = sum(pc), S = gAct + gRem;
    const R1 = Ract + firm, C1 = cAct + cFirm, R2 = R1 + pipe, C2 = C1 + cPipe;
    return {
      months, cur, r, last, pipeNoCost, byMark, hasSeg, planDiff, actDiff, actMissing, pS, pC, pG, planM, planN, planE, mon, pmon, mc, mg, ms, pc,
      crP, cr, taxK, tg, adj, Ract, cAct, gAct, gRem, S, firm, cFirm, pipe, cPipe, R1, C1, R2, C2, PR, PC, PG,
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
  const man = (n) => (n == null || !isFinite(n) ? "—" : (n < 0 ? "−" : "") + Math.round(Math.abs(n) / 1e4).toLocaleString() + "万");
  const sman = (n) => (n == null || !isFinite(n) ? "—" : (Math.round(n / 1e4) > 0 ? "+" : Math.round(n / 1e4) < 0 ? "−" : "±") + Math.abs(Math.round(n / 1e4)).toLocaleString() + "万");
  function frame(H, mx, mn) {
    const W = 360, L = 40, R = 6, T = 12, B = 22;
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

  /* ---------- ホバーの内訳（ツールチップ） ---------- */
  let TIPS = [], gid = 0;
  const tipId = (html) => (TIPS.push(html), TIPS.length - 1);
  const sw = (c, hatch) => `<i class="yj-sw" style="${hatch ? `border:1px solid ${c};background:repeating-linear-gradient(135deg,${c} 0 1.5px,#fff 1.5px 4px)` : `background:${c}`}"></i>`;
  const trow = (lab, a, b, d) => `<div class="r"><span>${lab}</span><span>${a}</span><span>${b == null ? "" : b}</span>${d === undefined ? "<span></span>" : `<span class="${d < 0 ? "t-err" : "t-ok"}">${sman(d)}</span>`}</div>`;
  const tsep = `<hr>`;
  // セルのコメント（月次PL）をツールチップ用に並べる。新しい月から、多いときは件数で省略
  const KIND_TAG = { firm: "", pipe: "P", pcost: "P原価", sales: "", cogs: "", sga: "" };
  function noteHtml(list, max) {
    if (!list.length) return "";
    max = max || 5;
    const cut = (t) => (t.length > 60 ? t.slice(0, 59) + "…" : t);
    const rows = list.slice(0, max).map(x => `<div class="yj-nt"><span>${esc(x.when)} ${esc(x.who)}</span>${esc(cut(x.text))}</div>`).join("");
    return `<div class="yj-nts">${rows}${list.length > max ? `<div class="yj-tn2">ほか${list.length - max}件</div>` : ""}</div>`;
  }
  // 透明な当たり判定（data-yjh：内訳の番号、data-g / data-gx：縦ガイド）
  const hit = (x, y, w, h, id, g, gx) => `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(1, w).toFixed(1)}" height="${Math.max(1, h).toFixed(1)}" fill="transparent" data-yjh="${id}"${g ? ` data-g="${g}" data-gx="${gx.toFixed(1)}"` : ""}/>`;

  // コメントのある月：月のラベルの右上に小さな印
  function noteDots(s, f, noted, cx) {
    if (!noted) return;
    noted.forEach((v, i) => { if (v) s.push(`<circle cx="${(cx(i) + 7).toFixed(1)}" cy="${f.H - 15}" r="2.2" fill="${COL.fc}"/>`); });
  }

  // ① 売上の累計（区分ごとの積み上げ面＋計画線）
  function salesChart(o) {
    const { d, keys, plan, pipe, labels } = o;
    const n = 12, cur = d.cur;
    const tops = [];
    let base = Array(n).fill(0);
    keys.forEach(k => { const t = base.map((v, i) => v + o.act[k][i]); tops.push({ k, bot: base, top: t }); base = t; });
    const withPipe = base.map((v, i) => v + pipe[i]);
    const f = frame(o.height || 200, Math.max(1, ...withPipe, ...plan) * 1.05);
    const step = (f.W - f.L - f.R) / n, cx = (i) => f.L + step * i + step / 2;
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="売上の累計">`, ...f.s];
    const poly = (top, bot, col, op, a, b) => {
      if (b - a < 1) return;
      const idx = []; for (let i = a; i < b; i++) idx.push(i);
      if (idx.length === 1) { const i = idx[0]; s.push(`<rect x="${(cx(i) - 3).toFixed(1)}" y="${f.y(top[i]).toFixed(1)}" width="6" height="${Math.max(0, f.y(bot[i]) - f.y(top[i])).toFixed(1)}" fill="${col}" fill-opacity="${op}"/>`); return; }
      const pts = idx.map(i => cx(i).toFixed(1) + "," + f.y(top[i]).toFixed(1)).concat(idx.slice().reverse().map(i => cx(i).toFixed(1) + "," + f.y(bot[i]).toFixed(1)));
      s.push(`<polygon points="${pts.join(" ")}" fill="${col}" fill-opacity="${op}"/>`);
    };
    tops.forEach(t => { poly(t.top, t.bot, CAT[t.k].col, 0.85, 0, cur); poly(t.top, t.bot, CAT[t.k].col, 0.25, Math.max(0, cur - 1), n); });
    const pc = o.lineColor;
    s.push(`<polyline points="${plan.map((v, i) => cx(i).toFixed(1) + "," + f.y(v).toFixed(1)).join(" ")}" fill="none" stroke="${pc}" stroke-width="1.6" stroke-dasharray="4 3"/>`);
    if (cur < n && pipe[n - 1]) s.push(`<polyline points="${withPipe.map((v, i) => cx(i).toFixed(1) + "," + f.y(v).toFixed(1)).slice(Math.max(0, cur - 1)).join(" ")}" fill="none" stroke="${pc}" stroke-width="1.6" stroke-dasharray="1.5 2.5"/>`);
    if (cur > 0 && cur < n) s.push(marker(f, f.L + step * cur));
    labels.forEach((t, i) => s.push(`<text x="${cx(i).toFixed(1)}" y="${f.H - 7}" font-size="9.5" fill="${COL.muted}" text-anchor="middle">${esc(t)}</text>`));
    noteDots(s, f, o.noted, cx);
    const g = "yjg" + (++gid);
    s.push(`<line id="${g}" x1="0" x2="0" y1="${f.T}" y2="${f.H - f.B}" stroke="${COL.ink}" stroke-opacity="0.45" visibility="hidden" pointer-events="none"/>`);
    for (let i = 0; i < n; i++) s.push(hit(f.L + step * i, 0, step, f.H - f.B, o.tip(i), g, cx(i)));
    s.push("</svg>");
    return s.join("");
  }

  // ② 月別の原価＋販管費（積み上げ）と計画の横線
  function costChart(o) {
    const { d, labels } = o, n = 12;
    const tot = d.mc.map((v, i) => (v || 0) + (d.mg[i] || 0)), pt = d.pC.map((v, i) => (v || 0) + (d.pG[i] || 0));
    const f = frame(o.height || 180, Math.max(1, ...tot, ...pt) * 1.05);
    const step = (f.W - f.L - f.R) / n, bw = Math.min(16, step * 0.55);
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="原価・販管費の月別">`, ...f.s];
    for (let i = 0; i < n; i++) {
      const x = f.L + step * i + (step - bw) / 2, fc = i >= d.cur, op = fc ? 0.3 : 1;
      // 下に販管費、上に原価を積む
      const c = d.mc[i] || 0, g = d.mg[i] || 0, y0 = f.y(0), y1 = f.y(g), y2 = f.y(g + c);
      if (g) s.push(`<rect x="${x.toFixed(1)}" y="${(y1 + 0.5).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(0.5, y0 - y1 - 0.5).toFixed(1)}" rx="1.5" fill="${o.spG[i] ? COL.spike : COL.sga}" fill-opacity="${op}"/>`);
      if (c) s.push(`<rect x="${x.toFixed(1)}" y="${y2.toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(0.5, y1 - y2 - 1).toFixed(1)}" rx="1.5" fill="${o.spC[i] ? COL.spike : COL.cogs}" fill-opacity="${op}"/>`);
      const py = f.y(pt[i]);
      s.push(`<line x1="${(x - 2).toFixed(1)}" x2="${(x + bw + 2).toFixed(1)}" y1="${py.toFixed(1)}" y2="${py.toFixed(1)}" stroke="${COL.ink}" stroke-width="1.6"/>`);
      s.push(`<text x="${(x + bw / 2).toFixed(1)}" y="${f.H - 7}" font-size="9.5" fill="${COL.muted}" text-anchor="middle">${esc(labels[i])}</text>`);
      if (o.noted && o.noted[i]) noteDots(s, f, o.noted.map((v, k) => k === i && v), () => x + bw / 2);
      s.push(hit(f.L + step * i, 0, step, f.H - f.B, o.tip(i)));
    }
    if (d.cur > 0 && d.cur < n) s.push(marker(f, f.L + step * d.cur));
    s.push("</svg>");
    return s.join("");
  }

  // ③ 年度末の着地（計画／受注確定ベース／着地見込み × 売上・コスト）
  function landingChart(o) {
    const groups = o.groups, all = [];
    groups.forEach(g => { all.push(sum(g.rev.map(x => x.v)), sum(g.cost.map(x => x.v))); });
    const f = frame(o.height || 200, Math.max(1, ...all) * 1.15);
    const n = groups.length, step = (f.W - f.L - f.R) / n, bw = step * 0.3;
    const pat = (id, c) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="6" height="6" fill="#FFFFFF"/><line x1="0" y1="0" x2="0" y2="6" stroke="${c}" stroke-width="2.2"/></pattern>`;
    const s = [`<svg viewBox="0 0 ${f.W} ${f.H}" width="100%" class="chart" role="img" aria-label="年度末の着地">`,
      `<defs>${pat("yjHf", COL.firm)}${pat("yjHc", COL.cogs)}${pat("yjHg", COL.sga)}</defs>`, ...f.s];
    groups.forEach((g, gi) => {
      [["rev", 0.12, "売上"], ["cost", 0.12 + 0.3 + 0.04, "コスト"]].forEach(([k, off, nm]) => {
        let base = 0;
        const x = f.L + step * gi + step * off;
        g[k].forEach(seg => {
          if (!seg.v) return;
          const y1 = f.y(base + seg.v), y0 = f.y(base);
          s.push(`<rect x="${x.toFixed(1)}" y="${y1.toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(0.5, y0 - y1 - 0.5).toFixed(1)}" fill="${seg.hatch ? `url(#${seg.hatch})` : seg.color}" ${seg.hatch ? `stroke="${seg.color}" stroke-width="0.8"` : ""}/>`);
          base += seg.v;
        });
        const rows = g[k].filter(x => x.v).slice().reverse().map(x => trow(sw(x.color, !!x.hatch) + esc(x.name), man(x.v))).join("");
        const id = tipId(`<div class="t">${esc(g.label)}・${nm}</div>${rows}${tsep}${trow("<b>合計</b>", "<b>" + man(base) + "</b>")}${g.notes && g.notes[k] ? g.notes[k] : ""}`);
        s.push(hit(x, f.y(base), bw, f.y(0) - f.y(base), id));
      });
      const top = f.y(Math.max(sum(g.rev.map(x => x.v)), sum(g.cost.map(x => x.v))));
      const cx = f.L + step * gi + step / 2;
      s.push(`<text x="${cx.toFixed(1)}" y="${(top - 5).toFixed(1)}" font-size="11" font-weight="700" text-anchor="middle" fill="${g.ok ? COL.ok : COL.bad}">${g.m.toFixed(1)}%</text>`);
      s.push(`<text x="${cx.toFixed(1)}" y="${f.H - 7}" font-size="10" text-anchor="middle" fill="${COL.muted}">${esc(g.label)}</text>`);
    });
    s.push("</svg>");
    return s.join("");
  }

  /* ---------- 画面 ---------- */
  let last = null; // 最後に描いた文脈（条件変更時の再描画用）
  const lgBox = (c) => `<i class="lg" style="background:${c}"></i>`;
  const lgLine = (c, style) => `<i class="lg lg-line" style="border-top:2px ${style || "solid"} ${c}"></i>`;
  const lgHatch = (c) => `<i class="lg" style="border:1px solid ${c};background:repeating-linear-gradient(135deg,${c} 0 1.5px,#fff 1.5px 4px)"></i>`;
  const badge = (kind, t) => `<span class="yj-st yj-${kind}">${esc(t)}</span>`;
  const rateCls = (r) => (r == null ? "" : r >= 97 ? "yj-ok" : r >= 90 ? "yj-warn" : "yj-bad");

  // 実績・見込みの範囲（1行）と、データの注意だけを出す
  function info(c, d) {
    const notes = [];
    if (!d.byMark) notes.push("月次PL／月次BSの4行目に印がないため、データのある最新月までを実績にしています。");
    if (d.pipeNoCost.length) notes.push(`パイプラインの原価が未入力の月：${d.pipeNoCost.map(i => Number(d.months[i].slice(5)) + "月").join("・")}`);
    return `<p class="yj-line">${d.cur ? `${Number(d.last.slice(5))}月まで実績（${d.cur}/12）・残り${d.r}か月は見込み` : "実績はまだありません（12か月すべて見込み）"}</p>
${notes.length ? `<p class="small t-amber yj-line">${notes.map(esc).join("<br>")}</p>` : ""}`;
  }

  function body(c, d) {
    TIPS = [];
    const p = c.p, labels = d.months.map(q => String(Number(q.slice(5)))), j = d.cur - 1;
    const mlab = (i) => labels[i] + "月";
    // 月次PLのコメント：月の位置（0〜11）ごと
    const NT = d.months.map(q => ((c.notes || {})[q] || []).map(x => Object.assign({ when: labels[d.months.indexOf(q)] + "月", who: (x.k ? CAT[x.k].short : x.label) + (KIND_TAG[x.kind] ? "（" + KIND_TAG[x.kind] + "）" : "") }, x)));
    // from〜to の月のコメントのうち条件に合うもの（新しい月から）
    const notesIn = (from, to, ok) => { const r = []; for (let i = to; i >= from; i--) NT[i].forEach(x => { if (ok(x)) r.push(x); }); return r; };
    const when = (i) => (i < d.cur ? "実績" : "見込み");
    const open = (k) => (p.open[k] ? "open" : "");

    // ① 売上（全体・保守・開発（既存）・開発（新規）それぞれの予実）
    const planOf = { m: d.planM, e: d.planE, n: d.planN };
    const act = {}, pip = {}, pln = {};
    KEYS.forEach(k => { act[k] = cum(d.mon[k]); pip[k] = cum(d.pmon[k]); pln[k] = cum(planOf[k]); });
    const sumK = (o, i) => KEYS.reduce((t, k) => t + o[k][i], 0);
    act.all = act.m.map((_, i) => sumK(act, i)); pip.all = pip.m.map((_, i) => sumK(pip, i)); pln.all = cum(d.pS);
    const seg = d.hasSeg && ["all", "m", "e", "n"].includes(p.seg) ? p.seg : "all";
    const nameOf = (k) => (k === "all" ? (d.hasSeg ? "全体" : "売上高") : CAT[k].name);
    const stat = (k) => {
      const a = d.cur ? act[k][j] : 0, pn = d.cur ? pln[k][j] : 0, endF = act[k][11], end = endF + pip[k][11], plan = pln[k][11];
      return { a, pn, r: pn ? a / pn * 100 : null, endF, end, plan, gap: plan - end, gapF: plan - endF };
    };
    const st = stat(seg);
    const st1 = st.gap <= 0 ? badge(st.gapF <= 0 ? "ok" : "warn", st.gapF <= 0 ? "受注確定で達成" : "パイプライン込みで達成") : badge(st.gap < st.plan * 0.06 ? "warn" : "bad", "不足 " + man(st.gap));
    const tiles = d.hasSeg ? `<div class="yj-tiles" role="group" aria-label="区分">${["all", ...KEYS].map(k => {
      const x = stat(k), rc = rateCls(x.r);
      return `<button type="button" class="yj-tile ${seg === k ? "on" : ""}" data-yact="seg" data-k="${k}" aria-pressed="${seg === k}">
        <span class="yj-tn">${k === "all" ? "" : lgBox(CAT[k].col)}${esc(k === "all" ? "全体" : CAT[k].short)}</span>
        <span class="yj-tv">${d.cur ? man(x.a) : man(x.end)}</span>
        <span class="yj-tr ${rc}">${x.r == null ? (d.cur ? "—" : "計画比 " + (x.plan ? Math.round(x.end / x.plan * 100) + "%" : "—")) : Math.round(x.r) + "%"}</span>
        <span class="yj-bar"><b class="${rc}" style="width:${x.r == null ? 0 : Math.min(100, x.r).toFixed(0)}%"></b></span></button>`;
    }).join("")}</div>` : "";
    const keys = seg === "all" ? KEYS : [seg];
    const salesNote = (x) => (x.kind === "firm" || x.kind === "pipe") ? keys.includes(x.k) : (x.kind === "sales" && seg === "all");
    const noted1 = NT.map(l => l.some(salesNote));
    const lineColor = seg === "all" ? COL.muted : CAT[seg].col;
    const rowsOf = seg === "all" && d.hasSeg ? ["n", "e", "m"] : [];
    const tip1 = (i) => {
      const rw = (k) => { const a = act[k][i] + pip[k][i], pl = pln[k][i]; return trow(k === "all" ? "<b>" + esc(nameOf(k)) + "</b>" : sw(CAT[k].col) + esc(CAT[k].name), man(a), man(pl), a - pl); };
      const pp = pip[seg][i];
      const nts = notesIn(0, i, salesNote);
      return tipId(`<div class="t">${mlab(i)} 累計（${when(i)}）<span class="yj-th">実績 / 計画　差</span></div>${rowsOf.map(rw).join("")}${rowsOf.length ? tsep : ""}${rw(seg)}${pp ? `<div class="yj-tn2">うちパイプライン ${man(pp)}</div>` : ""}${noteHtml(nts)}`);
    };
    const segNotes = [];
    if (d.hasSeg && d.actMissing.length) segNotes.push(`内訳がない月（${d.actMissing.map(mlab).join("・")}）は開発（既存）に含めています。`);
    if (d.actDiff.length) segNotes.push(`月次PLの内訳と売上高の差（${d.actDiff.map(x => mlab(x.i) + " " + sman(x.v)).join("、")}）は開発（既存）に含めています。`);
    if (d.planDiff.length) segNotes.push(`計画シートの内訳と売上高に差があります（${d.planDiff.map(x => mlab(x.i) + " " + sman(x.v)).join("、")}）。`);
    const lg1 = (seg === "all" ? (d.hasSeg ? KEYS.map(k => `<span>${lgBox(CAT[k].col)}${CAT[k].name}</span>`).join("") : `<span>${lgBox(CAT.e.col)}実績</span>`) : `<span>${lgBox(CAT[seg].col)}実績</span>`)
      + (d.r > 0 ? `<span><i class="lg" style="background:${seg === "all" ? CAT.e.col : CAT[seg].col};opacity:.3"></i>受注確定</span>` : "")
      + (d.r > 0 && pip[seg][11] ? `<span>${lgLine(lineColor, "dotted")}パイプライン込み</span>` : "")
      + `<span>${lgLine(lineColor, "dashed")}計画</span>`;
    const sec1 = `<section class="card yj-sec">
  <div class="card-head"><h2>① 売上</h2>${st1}</div>
  ${tiles}
  <div class="yj-sum"><span>${d.cur ? `${mlab(j)}まで ${man(st.a)} / 計画 ${man(st.pn)}` : "実績なし"}</span><span>年度末 ${man(st.end)} / ${man(st.plan)}</span></div>
  ${salesChart({ d, keys, act, plan: pln[seg], pipe: pip[seg], labels, lineColor, tip: tip1, noted: noted1 })}
  <div class="legend">${lg1}</div>
  ${!d.hasSeg ? `<p class="muted small">区分（保守・開発(既存)・開発(新規)）で見るには、計画シートと月次PLの売上高の下にその行を追加してください。</p>` : ""}
  ${segNotes.length ? `<p class="small t-amber">${segNotes.map(esc).join("<br>")}</p>` : ""}
  <div class="yj-tip" hidden></div>
</section>`;

    // ② コスト（その月までの累計実績が計画の累計の○%を超えた実績月を突発とする）
    const sp = p.sp;
    const cA = { c: cum(d.mc), g: cum(d.mg) }, cP = { c: cum(d.pC), g: cum(d.pG) };
    const isSp = (a, pl, pc) => a > 0 && a > pl * pc / 100;
    const spC = d.mc.map((_, i) => i < d.cur && isSp(cA.c[i], cP.c[i], sp.cp));
    const spG = d.mg.map((_, i) => i < d.cur && isSp(cA.g[i], cP.g[i], sp.sp));
    const nSp = spC.filter(Boolean).length + spG.filter(Boolean).length;
    const sC = d.cAct, sCP = sum(d.pC.slice(0, d.cur)), sG = d.gAct, sGP = sum(d.pG.slice(0, d.cur));
    const over = (sCP && sC > sCP * 1.05) || (sGP && sG > sGP * 1.05);
    const k2 = nSp ? "bad" : over ? "warn" : "ok";
    const pct = (a, b) => (b ? Math.round(a / b * 100) + "%" : "—");
    const tip2 = (i) => {
      const t = (d.mc[i] || 0) + (d.mg[i] || 0), pt = (d.pC[i] || 0) + (d.pG[i] || 0);
      const spk = [];
      if (spC[i]) spk.push(`原価の累計が計画累計の${pct(cA.c[i], cP.c[i])}（基準${sp.cp}%）`);
      if (spG[i]) spk.push(`販管費の累計が計画累計の${pct(cA.g[i], cP.g[i])}（基準${sp.sp}%）`);
      return tipId(`<div class="t">${mlab(i)}（${when(i)}）${spk.length ? ' <span class="t-err">突発</span>' : ""}<span class="yj-th">${when(i)} / 計画　差</span></div>`
        + trow(sw(spC[i] ? COL.spike : COL.cogs) + "原価", man(d.mc[i]), man(d.pC[i]), (d.mc[i] || 0) - (d.pC[i] || 0))
        + trow(sw(spG[i] ? COL.spike : COL.sga) + "販管費", man(d.mg[i]), man(d.pG[i]), (d.mg[i] || 0) - (d.pG[i] || 0))
        + tsep + trow("<b>合計</b>", "<b>" + man(t) + "</b>", man(pt), t - pt)
        + (i < d.cur ? `<div class="yj-tn2">累計の計画比　販管費 ${pct(cA.g[i], cP.g[i])}・原価 ${pct(cA.c[i], cP.c[i])}</div>` : "")
        + (spk.length ? `<div class="t-err yj-tn2">${spk.join("<br>")}</div>` : "")
        + noteHtml(notesIn(i, i, costNote)));
    };
    const costNote = (x) => ["cogs", "sga", "pcost"].includes(x.kind);
    const sec2 = `<section class="card yj-sec">
  <div class="card-head"><h2>② コスト</h2>${badge(k2, nSp ? "突発 " + nSp + "件" : over ? "計画超過" : "計画内")}</div>
  <div class="yj-sum"><span>${d.cur ? `${mlab(j)}まで ${man(sC + sG)} / 計画 ${man(sCP + sGP)}（${pct(sC + sG, sCP + sGP)}）` : "実績なし"}</span></div>
  ${costChart({ d, labels, spC, spG, tip: tip2, noted: NT.map(l => l.some(costNote)) })}
  <div class="legend"><span>${lgBox(COL.sga)}販管費</span><span>${lgBox(COL.cogs)}原価</span><span>${lgBox(COL.spike)}突発</span><span>${lgLine(COL.ink)}計画</span></div>
  <details data-yjd="spk" ${open("spk")}><summary>突発の条件</summary>
    <div class="yj-spk yj-spk2"><span>原価（計画累計の%超）</span><input type="number" class="inp" data-yjc="cp" value="${sp.cp}" min="100" step="5">
    <span>販管費（計画累計の%超）</span><input type="number" class="inp" data-yjc="sp" value="${sp.sp}" min="100" step="5"></div></details>
  <div class="yj-tip" hidden></div>
</section>`;

    // ③ 利益率（年度末の着地）
    const tgtP = p.tgt, okOf = (m) => m >= tgtP;
    const groups = [
      { label: "計画", m: d.mP, ok: okOf(d.mP), rev: [{ name: "計画売上", v: d.PR, color: COL.planBar }], cost: [{ name: "原価", v: d.PC, color: COL.cogs }, { name: "販管費", v: d.PG, color: COL.sga }] },
      { label: "受注確定", m: d.m1, ok: okOf(d.m1), rev: [{ name: "売上 実績", v: d.Ract, color: COL.act }, { name: "受注確定", v: d.firm, color: COL.firm }], cost: [{ name: "原価 実績", v: d.cAct, color: COL.cogs }, { name: "原価 見込み", v: d.C1 - d.cAct, color: COL.cogs, hatch: "yjHc" }, { name: "販管費 実績", v: d.gAct, color: COL.sga }, { name: "販管費 見込み", v: d.gRem, color: COL.sga, hatch: "yjHg" }] },
      { label: "着地見込み", m: d.m2, ok: okOf(d.m2), rev: [{ name: "売上 実績", v: d.Ract, color: COL.act }, { name: "受注確定", v: d.firm, color: COL.firm }, { name: "パイプライン", v: d.pipe, color: COL.firm, hatch: "yjHf" }], cost: [{ name: "原価 実績", v: d.cAct, color: COL.cogs }, { name: "原価 見込み", v: d.C2 - d.cAct, color: COL.cogs, hatch: "yjHc" }, { name: "販管費 実績", v: d.gAct, color: COL.sga }, { name: "販管費 見込み", v: d.gRem, color: COL.sga, hatch: "yjHg" }] }
    ];
    const fut = (ok) => (d.cur < 12 ? notesIn(d.cur, 11, ok) : []);
    groups[1].notes = { rev: noteHtml(fut(x => x.kind === "firm" || x.kind === "sales")), cost: noteHtml(fut(costNote)) };
    groups[2].notes = { rev: noteHtml(fut(x => x.kind === "firm" || x.kind === "pipe" || x.kind === "sales"), 6), cost: noteHtml(fut(costNote)) };
    const ok = okOf(d.m2), k3 = ok ? (okOf(d.m1) ? "ok" : "warn") : "bad";
    // 目標に届かないときの上積み額（計画の原価率で受注した場合）
    const kk = 1 - d.cr - d.taxK - d.tg / 1.1, add = kk > 0 ? (d.C2 + d.S - d.R2 * (1 - d.taxK - d.tg / 1.1)) / kk : Infinity;
    // パイプライン案件で許容できる原価率の上限
    const allow = (d.R2 - d.S - d.R2 * d.taxK - d.tg * d.R2 / 1.1) - d.C1;
    const lines3 = [];
    if (!ok) lines3.push(isFinite(add) && d.r > 0 ? `${tgtP}%確保には あと <b>${man(add)}円</b> の受注が必要` : `${tgtP}%に届きません`);
    if (d.r > 0 && d.pipe > 0) lines3.push(allow <= 0 ? `パイプラインの原価をゼロにしても${tgtP}%に届きません` : `パイプラインの原価率の上限 <b>${(allow / d.pipe * 100).toFixed(1)}%</b>（入力済み ${(d.cPipe / d.pipe * 100).toFixed(1)}%）`);
    const sec3 = `<section class="card yj-sec">
  <div class="card-head"><h2>③ 利益率（着地）</h2>${badge(k3, "着地 " + d.m2.toFixed(1) + "%")}</div>
  <div class="yj-sum"><span>目標 ${tgtP}%・受注確定 ${d.m1.toFixed(1)}%・計画 ${d.mP.toFixed(1)}%</span></div>
  ${landingChart({ groups })}
  <div class="legend"><span>${lgHatch(COL.muted)}見込み</span></div>
  ${lines3.length ? `<p class="small yj-line">${lines3.join("<br>")}</p>` : ""}
  <details data-yjd="tax" ${open("tax")}><summary>利益率の条件</summary>
    <div class="yj-spk yj-spk2"><span>目標（%）</span><input type="number" class="inp" data-yjc="tgt" value="${p.tgt}" min="0" step="0.5">
    <span>みなし仕入率（%）</span><input type="number" class="inp" data-yjc="mr" value="${p.mr}" min="0" max="90" step="10"></div>
    <p class="muted small">調整後営業利益率（税抜）＝（営業利益−消費税の納付見込み）÷ 税抜売上。簡易課税が前提です。</p></details>
  <div class="yj-tip" hidden></div>
</section>`;
    return sec1 + sec2 + sec3;
  }

  function render(c) {
    const M = c.M;
    if (!M.latest) return c.empty;
    if (!c.plan || !c.plan[M.K.sales]) {
      return `<div class="toolbar">${c.fySelect}</div><div class="note">${esc(M.fyLabel(c.fy))}の計画シートがありません。予実は計画シートの売上高・売上原価・販管費と比べます。計画タブで計画シートを作成してください。<br><button type="button" class="btn primary" data-act="tab" data-tab="plan" style="margin-top:8px">計画タブを開く</button></div>`;
    }
    const d = compute(c);
    last = { c, d };
    return `<div class="toolbar">${c.fySelect}</div>
${info(c, d)}
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
  // 数値の条件は入力中も反映する（再描画後も入力欄のフォーカスとカーソル位置を保つ）
  const applyNum = (t) => {
    const id = t.dataset && t.dataset.yjc;
    if (!id || !last) return;
    if (t.value === "") return;
    const p = last.c.p, v = Number(t.value);
    if (!isFinite(v)) return;
    if (["cp", "sp"].includes(id)) { if (p.sp[id] === v) return; p.sp[id] = v; } else { if (p[id] === v) return; p[id] = v; }
    saveParams(p);
    const pos = t.selectionStart;
    refresh();
    const n = document.querySelector(`[data-yjc="${id}"]`);
    if (n) { n.focus(); try { if (pos != null) n.setSelectionRange(pos, pos); } catch (err) {} }
  };
  document.addEventListener("input", (e) => applyNum(e.target));
  document.addEventListener("change", (e) => applyNum(e.target));
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
      case "seg": p.seg = b.dataset.k; saveParams(p); refresh(); break;
    }
  });

  // ホバー（タッチではタップ）で、その月・その棒に積み上がっている内訳を出す
  let tipOn = null;
  function hideTip() {
    if (!tipOn) return;
    tipOn.tip.hidden = true;
    if (tipOn.guide) tipOn.guide.setAttribute("visibility", "hidden");
    tipOn = null;
  }
  function showTip(e) {
    const h = e.target && e.target.closest ? e.target.closest("[data-yjh]") : null;
    if (!h) { hideTip(); return; }
    const sec = h.closest(".yj-sec"), tip = sec && sec.querySelector(".yj-tip"), html = TIPS[Number(h.dataset.yjh)];
    if (!tip || html == null) { hideTip(); return; }
    if (!tipOn || tipOn.h !== h) {
      hideTip();
      tip.innerHTML = html; tip.hidden = false;
      const guide = h.dataset.g ? document.getElementById(h.dataset.g) : null;
      if (guide) { guide.setAttribute("x1", h.dataset.gx); guide.setAttribute("x2", h.dataset.gx); guide.setAttribute("visibility", "visible"); }
      tipOn = { h, tip, guide };
    }
    const r = sec.getBoundingClientRect(), tw = tip.offsetWidth, th = tip.offsetHeight;
    let x = e.clientX - r.left + 12, y = e.clientY - r.top + 12;
    if (x + tw > r.width - 4) x = Math.max(4, e.clientX - r.left - tw - 12);
    if (y + th > r.height - 4) y = Math.max(4, e.clientY - r.top - th - 12);
    tip.style.left = x + "px"; tip.style.top = y + "px";
  }
  document.addEventListener("pointermove", showTip);
  document.addEventListener("pointerdown", showTip);
  document.addEventListener("scroll", hideTip, true);

  global.Yojitsu = { render, refresh, loadParams, saveParams, compute };
})(window);
