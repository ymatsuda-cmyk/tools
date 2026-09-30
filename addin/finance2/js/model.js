/* ============================================================
 * model.js — TB_明細 から会計年度・月次系列・予測を組み立てる
 * ============================================================ */
(function (global) {
  "use strict";
  const norm = (s) => String(s || "").replace(/[\s\u3000]/g, "");
  const ym = (p) => { const [y, m] = p.split("-").map(Number); return { y, m }; };
  const pstr = (y, m) => y + "-" + String(m).padStart(2, "0");
  const addM = (p, n) => { const { y, m } = ym(p); const t = y * 12 + (m - 1) + n; return pstr(Math.floor(t / 12), (t % 12) + 1); };
  const mIndex = (p) => { const { y, m } = ym(p); return y * 12 + m - 1; };

  const FIND = {
    sales: ["純売上高", "売上高合計", "売上高計", "売上高"],
    cogs: ["売上原価"],
    gross: ["売上総利益", "売上総損失"],
    sga: ["販売費及び一般管理費", "販売費及び一般管理費計"],
    op: ["営業利益", "営業損失"],
    ord: ["経常利益", "経常損失"],
    net: ["当期純利益", "当期純損失"],
    cash: ["現金預金計", "現金及び預金", "現金・預金計", "現金及び預金計"],
    assets: ["資産合計", "資産の部合計"],
    curA: ["流動資産計", "流動資産合計"],
    fixA: ["固定資産計", "固定資産合計"],
    defA: ["繰延資産計", "繰延資産合計"],
    curL: ["流動負債計", "流動負債合計"],
    fixL: ["固定負債計", "固定負債合計"],
    liab: ["負債合計", "負債の部合計"],
    equity: ["純資産合計", "純資産の部合計"]
  };

  function build(tb, settings) {
    const start = settings.期首月 || 10;
    const byP = {};             // period -> key -> row
    const meta = {};            // key -> {sheet,name,kind,code,order} （最新期間の値）
    const periods = [...new Set(tb.map(r => r.period))].filter(p => /^\d{4}-\d{2}$/.test(p)).sort();
    tb.forEach(r => { (byP[r.period] = byP[r.period] || {})[r.key] = r; });
    periods.forEach(p => Object.values(byP[p]).forEach(r => { meta[r.key] = { key: r.key, sheet: r.sheet, name: r.name, kind: r.kind, code: r.code, order: r.order }; }));

    const fyOf = (p) => { const { y, m } = ym(p); return start === 1 ? y : (m >= start ? y + 1 : y); };
    const endMonth = start === 1 ? 12 : start - 1;
    const fyLabel = (fy) => `${fy}年${endMonth}月期`;
    const fyMonths = (fy) => { const first = start === 1 ? pstr(fy, 1) : pstr(fy - 1, start); return Array.from({ length: 12 }, (_, i) => addM(first, i)); };
    const fys = [...new Set(periods.map(fyOf))].sort();
    const latest = periods[periods.length - 1] || null;

    const keyByNames = (sheet, names) => {
      for (const n of names) {
        const hit = Object.values(meta).filter(m => m.sheet === sheet && norm(m.name) === n);
        if (hit.length) { const agg = hit.find(h => h.kind === "集計"); return (agg || hit[0]).key; }
      }
      return null;
    };
    const K = {};
    Object.entries(FIND).forEach(([k, names]) => {
      const sheet = ["sales", "cogs", "gross", "sga", "op", "ord", "net"].includes(k) ? "PL" : "BS";
      K[k] = keyByNames(sheet, names);
    });

    const row = (key, p) => (byP[p] && byP[p][key]) || null;
    const bal = (key, p) => { const r = row(key, p); return r ? r.bal : null; };
    const month = (key, p) => { const r = row(key, p); return r ? r.bal - r.open : null; };
    const open = (key, p) => { const r = row(key, p); return r ? r.open : null; };
    // PLは残高＝期首からの累計、BSは残高＝月末残高。月次値：PLは当月発生、BSは月末残高
    const series = (key, p) => (meta[key] && meta[key].sheet === "BS" ? bal(key, p) : month(key, p));
    const has = (p) => !!byP[p];

    // 「増えると良い」科目か（色分け用）
    const goodUp = (key) => {
      const m = meta[key]; if (!m) return true;
      if (m.sheet === "BS") return null;
      const n = norm(m.name);
      return /(売上|収益|利益|受取|雑収入|益$)/.test(n) && !/損失|原価|費/.test(n);
    };

    return {
      start, periods, fys, latest, meta, K, byP,
      fyOf, fyLabel, fyMonths, addM, has, row, bal, month, open, series, goodUp, mIndex,
      keysOf: (sheet) => Object.values(meta).filter(m => m.sheet === sheet).sort((a, b) => a.order - b.order),
      latestIn: (fy) => { const ps = periods.filter(p => fyOf(p) === fy); return ps[ps.length - 1] || null; }
    };
  }

  /* ---------- 予測 ----------
   * 対象：月次値（PL＝当月発生）。直近のFYの残り月（残りがなければ翌期12か月）
   * yoy : 直近6か月の「今年／前年同月」合計比 × 前年同月
   * ma  : 直近3か月の平均
   * reg : 直近12か月の最小二乗直線
   * 幅   : 当てはめ残差の標準偏差 × 1.28（約80%）
   */
  function forecast(M, key, method) {
    const res = { ok: false, reason: "", points: [], horizon: [], method };
    if (!M.latest) { res.reason = "データがありません。"; return res; }
    const fy = M.fyOf(M.latest);
    const months = M.fyMonths(fy);
    let horizon = months.filter(p => p > M.latest);
    let targetFy = fy;
    if (!horizon.length) { targetFy = fy + 1; horizon = M.fyMonths(fy + 1); }
    res.targetFy = targetFy; res.horizon = horizon;
    const v = (p) => M.month(key, p);
    const actual = M.periods.filter(p => p <= M.latest && v(p) !== null);
    const last = (n) => actual.slice(-n);
    const std = (arr) => { if (arr.length < 2) return null; const mu = arr.reduce((a, b) => a + b, 0) / arr.length; return Math.sqrt(arr.reduce((a, b) => a + (b - mu) ** 2, 0) / (arr.length - 1)); };
    let f = null, sigma = null;

    if (method === "yoy") {
      const pairs = last(6).filter(p => v(M.addM(p, -12)) !== null);
      if (!pairs.length) { res.reason = "前年同月のデータがないため、前年同月比では予測できません。前期のPDFを取り込むと使えます。"; return res; }
      const a = pairs.reduce((s, p) => s + v(p), 0), b = pairs.reduce((s, p) => s + v(M.addM(p, -12)), 0);
      if (Math.abs(b) < 1) { res.reason = "前年同月の値が0のため比率を計算できません。"; return res; }
      const r = a / b;
      res.ratio = r; res.basis = pairs;
      const base = (p) => { const q = M.addM(p, -12); const bv = v(q); return bv !== null ? bv : null; };
      if (horizon.some(p => base(p) === null)) { res.reason = "予測する月の前年同月データが揃っていません。"; return res; }
      f = (p) => base(p) * r;
      sigma = std(pairs.map(p => v(p) - base(p) * r));
    } else if (method === "ma") {
      const l = last(3);
      if (!l.length) { res.reason = "月次の実績がありません。"; return res; }
      const mu = l.reduce((s, p) => s + v(p), 0) / l.length;
      f = () => mu;
      sigma = std(last(6).map(p => v(p)));
      res.basis = l;
    } else {
      const l = last(12);
      if (l.length < 3) { res.reason = "傾向線には3か月以上の実績が必要です。"; return res; }
      const xs = l.map(p => M.mIndex(p)), ys = l.map(v);
      const mx = xs.reduce((a, b) => a + b) / xs.length, my = ys.reduce((a, b) => a + b) / ys.length;
      const sxx = xs.reduce((a, x) => a + (x - mx) ** 2, 0), sxy = xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0);
      const slope = sxx ? sxy / sxx : 0, icpt = my - slope * mx;
      f = (p) => icpt + slope * M.mIndex(p);
      sigma = std(l.map((p, i) => ys[i] - (icpt + slope * xs[i])));
      res.slope = slope; res.basis = l;
    }
    if (sigma === null) sigma = Math.abs(f(horizon[0])) * 0.1;
    const z = 1.28;
    res.points = horizon.map(p => ({ p, v: f(p), lo: f(p) - z * sigma, hi: f(p) + z * sigma }));
    const sum = res.points.reduce((s, x) => s + x.v, 0);
    const spread = z * sigma * Math.sqrt(horizon.length);
    const actualInTarget = M.fyMonths(targetFy).filter(p => p <= M.latest);
    // 当期の実績累計（最新月の残高＝期首からの累計）
    const cum = targetFy === fy ? (M.bal(key, M.latest) || 0) : 0;
    res.actualCum = cum; res.actualMonths = actualInTarget;
    res.total = cum + sum; res.lo = res.total - spread; res.hi = res.total + spread;
    res.ok = true;
    return res;
  }

  global.FinModel = { build, forecast, norm, addM };
})(window);
