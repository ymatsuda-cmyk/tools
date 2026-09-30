/* LedgerPilot アドイン本体 */
(function () {
  "use strict";
  var APP_VERSION = "rev_20260930_lp001";
  window.APP_VERSION = APP_VERSION;
  var L = window.LedgerCore, DV = window.LedgerDiffView, CH = window.LPCharts;
  var esc = DV.esc;
  var $ = function (id) { return document.getElementById(id); };

  var PL_CATS = ["売上高", "売上原価", "販管費", "営業外収益", "営業外費用", "特別利益", "特別損失", "法人税等"];
  var COLORS = { sales: "#34398f", cost: "#b9bdd6", profit: "#1b7a4e", dr: "#5157c4", cr: "#c9a36b" };

  var S = {
    entries: [], postings: [], master: {}, months: [], history: [], byYm: {}, plYm: {},
    tab: "overview", range: { from: null, to: null }, month: null,
    acc: { code: null, ym: null, q: "", limit: 200, breakdown: "partner" },
    imp: { parsed: null, diff: null, view: null, name: "" }
  };

  // ===== 表示ユーティリティ =====
  function yen(n) {
    n = Math.round(n || 0);
    if (n < 0) return '<span class="neg">△' + (-n).toLocaleString("ja-JP") + "</span>";
    return n.toLocaleString("ja-JP");
  }
  function pct(a, b) { return b ? (a / b * 100).toFixed(1) + "%" : "—"; }
  // invert=true: 費用科目（増加が悪化）。符号は実際の増減のまま、色だけ反転
  function diffYen(n, invert) {
    n = Math.round(n || 0);
    if (!n) return '<span class="muted">±0</span>';
    return '<span class="' + ((n > 0) !== !!invert ? "pos" : "neg") + '">' + (n > 0 ? "+" : "△") + Math.abs(n).toLocaleString("ja-JP") + "</span>";
  }
  function ymLabel(ym) { return ym ? ym.slice(0, 4) + "年" + Number(ym.slice(5)) + "月" : ""; }
  function ymShort(ym) { return ym.slice(2, 4) + "/" + ym.slice(5); }
  function status(cls, text) {
    var el = $("status");
    if (!text) { el.className = "lp-note hidden"; return; }
    el.className = "lp-note " + cls; el.textContent = text;
  }
  function isPL(cat) { return PL_CATS.indexOf(cat) >= 0; }
  function catOf(code) { return (S.master[code] && S.master[code].cat) || L.defaultCategory(code); }
  function nameOf(code) {
    if (S.master[code] && S.master[code].name) return S.master[code].name;
    var p = S.postings.find(function (x) { return x.code === code; });
    return p ? p.name : code;
  }
  function saveUi() {
    try { localStorage.setItem("lp-ui", JSON.stringify({ tab: S.tab, range: S.range, month: S.month, acc: S.acc.code })); } catch (e) { /* noop */ }
  }
  function restoreUi() {
    try {
      var u = JSON.parse(localStorage.getItem("lp-ui") || "{}");
      if (u.tab) S.tab = u.tab;
      if (u.range) S.range = u.range;
      if (u.month) S.month = u.month;
      if (u.acc) S.acc.code = u.acc;
    } catch (e) { /* noop */ }
  }

  // ===== Excel 読み込み =====
  async function loadBook() {
    return Excel.run(async function (ctx) {
      var wss = ctx.workbook.worksheets;
      wss.load("items/name");
      await ctx.sync();
      var targets = [], masterR = null, histR = null;
      wss.items.forEach(function (ws) {
        if (L.ymFromSheet(ws.name)) {
          var r = ws.getUsedRangeOrNullObject(true); r.load("values,rowIndex");
          targets.push({ name: ws.name, r: r });
        } else if (ws.name === L.MASTER_SHEET) {
          masterR = ws.getUsedRangeOrNullObject(true); masterR.load("values");
        } else if (ws.name === L.HISTORY_SHEET) {
          histR = ws.getUsedRangeOrNullObject(true); histR.load("values");
        }
      });
      await ctx.sync();
      var entries = [], header = null;
      targets.forEach(function (t) {
        if (t.r.isNullObject) return;
        var v = t.r.values, width = v[0].length - L.META_N;
        if (!header) header = v[0].slice(L.META_N);
        v.slice(1).forEach(function (row, i) {
          if (!row[0]) return;
          entries.push(L.entryFromSheetRow(row, t.r.rowIndex + i + 2, t.name, width));
        });
      });
      var master = {};
      if (masterR && !masterR.isNullObject) {
        masterR.values.slice(1).forEach(function (r) {
          if (r[0] !== "") master[String(r[0])] = { name: r[1], cat: r[2] || L.defaultCategory(r[0]) };
        });
      }
      var history = histR && !histR.isNullObject ? histR.values.slice(1).filter(function (r) { return r[0] !== ""; }) : [];
      return { entries: entries, master: master, history: history, header: header };
    });
  }

  function rebuild(data) {
    S.entries = data.entries; S.master = data.master; S.history = data.history;
    S.postings = L.toPostings(S.entries, S.master);
    S.byYm = {};
    S.postings.forEach(function (p) { (S.byYm[p.ym] = S.byYm[p.ym] || []).push(p); });
    S.months = Object.keys(S.byYm).sort();
    S.plYm = {};
    S.months.forEach(function (ym) { S.plYm[ym] = L.plOf(S.byYm[ym]); });
    if (!S.range.from || S.months.indexOf(S.range.from) < 0) S.range.from = S.months[0] || null;
    if (!S.range.to || S.months.indexOf(S.range.to) < 0) S.range.to = S.months[S.months.length - 1] || null;
    if (!S.month || S.months.indexOf(S.month) < 0) S.month = S.months[S.months.length - 1] || null;
  }

  async function reload() {
    status("", "台帳を読み込んでいます…");
    try {
      rebuild(await loadBook());
      S.loaded = true;
      status("", "");
      render();
    } catch (e) {
      console.error(e);
      status("err", "台帳を読み込めませんでした: " + (e.message || e));
    }
  }

  // ===== Excel 行ジャンプ（シート切替と選択は別sync） =====
  async function jumpTo(sheet, row) {
    if (!sheet || !row) return;
    try {
      await Excel.run(async function (ctx) {
        var ws = ctx.workbook.worksheets.getItem(sheet);
        ws.activate();
        await ctx.sync();
        ws.getRange(row + ":" + row).select();
        await ctx.sync();
      });
    } catch (e) { status("err", "シートへ移動できませんでした: " + (e.message || e)); }
  }

  // ===== タブ =====
  function setTab(tab) {
    S.tab = tab;
    document.querySelectorAll(".tabs button").forEach(function (b) { b.classList.toggle("active", b.dataset.tab === tab); });
    ["overview", "month", "account", "import"].forEach(function (t) { $("pane-" + t).classList.toggle("hidden", t !== tab); });
    saveUi();
    render();
  }
  function render() {
    if (!S.loaded && S.tab !== "import") return;
    if (S.tab === "overview") renderOverview();
    else if (S.tab === "month") renderMonth();
    else if (S.tab === "account") renderAccount();
    else renderImport();
  }
  function emptyState(pane) {
    $(pane).innerHTML = '<div class="lp-panel empty"><b>まだ仕訳がありません</b>取込タブから弥生会計の仕訳日記帳CSVを取り込むと、ここに損益が表示されます。<div style="margin-top:10px"><button class="lp-btn primary" data-go="import">CSVを取り込む</button></div></div>';
    $(pane).querySelector("[data-go]").onclick = function () { setTab("import"); };
  }

  // ===== 全体 =====
  function rangeMonths() {
    return S.months.filter(function (m) { return m >= S.range.from && m <= S.range.to; });
  }
  function monthOptions(sel) {
    return S.months.map(function (m) { return '<option value="' + m + '"' + (m === sel ? " selected" : "") + ">" + ymLabel(m) + "</option>"; }).join("");
  }
  function plSheet(pl, prev) {
    var rows = [
      ["売上高", pl["売上高"], "", "key", "売上高"],
      ["売上原価", pl["売上原価"], "", "", "売上原価"],
      ["売上総利益", pl["売上総利益"], "粗利率 " + pct(pl["売上総利益"], pl["売上高"]), "profit"],
      ["販売費及び一般管理費", pl["販管費"], "", "", "販管費"],
      ["営業利益", pl["営業利益"], "利益率 " + pct(pl["営業利益"], pl["売上高"]), "profit key"],
      ["営業外収益", pl["営業外収益"], "", "", "営業外収益"],
      ["営業外費用", pl["営業外費用"], "", "", "営業外費用"],
      ["経常利益", pl["経常利益"], "利益率 " + pct(pl["経常利益"], pl["売上高"]), "profit"]
    ];
    if (pl["特別利益"] || pl["特別損失"] || pl["法人税等"]) {
      rows.push(["特別損益", pl["特別利益"] - pl["特別損失"], "", ""]);
      rows.push(["法人税等", pl["法人税等"], "", ""]);
      rows.push(["当期純利益", pl["当期純利益"], "", "profit"]);
    }
    var keyMap = { "売上高": "売上高", "売上原価": "売上原価", "売上総利益": "売上総利益", "販売費及び一般管理費": "販管費",
      "営業利益": "営業利益", "営業外収益": "営業外収益", "営業外費用": "営業外費用", "経常利益": "経常利益",
      "特別損益": null, "法人税等": "法人税等", "当期純利益": "当期純利益" };
    return '<div class="pl-sheet">' + rows.map(function (r) {
      var sub = r[2];
      if (prev) {
        var k = keyMap[r[0]];
        var pv = k ? prev[k] : prev["特別利益"] - prev["特別損失"];
        sub = "前月比 " + diffYen(r[1] - pv, /原価|管理費|費用|法人税/.test(r[0]));
      }
      return '<div class="pl-line ' + r[3] + '"><span class="lbl">' + r[0] + '</span><span class="amt">' + yen(r[1]) +
        '</span><span class="sub">' + sub + "</span></div>";
    }).join("") + "</div>";
  }
  function waterfallSteps(pl) {
    var s = [
      { label: "売上高", value: pl["売上高"], kind: "total" },
      { label: "売上原価", value: pl["売上原価"], kind: "down" },
      { label: "売上総利益", value: pl["売上総利益"], kind: "total" },
      { label: "販管費", value: pl["販管費"], kind: "down" },
      { label: "営業利益", value: pl["営業利益"], kind: "total" }
    ];
    var nonop = pl["営業外収益"] - pl["営業外費用"];
    if (nonop) s.push({ label: "営業外", value: Math.abs(nonop), kind: nonop > 0 ? "up" : "down" });
    s.push({ label: "経常利益", value: pl["経常利益"], kind: "total" });
    return s;
  }
  function accountTotals(postings, cats) {
    var m = {};
    postings.forEach(function (p) {
      if (cats.indexOf(p.cat) < 0) return;
      m[p.code] = (m[p.code] || 0) + L.signed(p);
    });
    return Object.keys(m).map(function (c) { return { code: c, name: nameOf(c), cat: catOf(c), v: m[c] }; })
      .sort(function (a, b) { return b.v - a.v; });
  }
  function rankHtml(list, cls, limit) {
    var top = list.slice(0, limit || 12);
    var max = Math.max.apply(null, top.map(function (x) { return Math.abs(x.v); }).concat([1]));
    return '<div class="rank">' + top.map(function (x, i) {
      return '<div class="rank-row" data-i="' + i + '" title="' + esc(x.name) + '"><span class="nm">' + esc(x.name) +
        '</span><span class="track"><span class="fill ' + (cls || "") + '" style="width:' + (Math.abs(x.v) / max * 100).toFixed(1) + '%"></span></span><span class="v">' + yen(x.v) + "</span></div>";
    }).join("") + "</div>";
  }
  function bindRank(host, list, fn) {
    host.querySelectorAll(".rank-row").forEach(function (r) {
      r.addEventListener("click", function () { fn(list[Number(r.dataset.i)]); });
    });
  }
  function partnerTotals(postings, cat) {
    var m = {};
    postings.forEach(function (p) {
      if (p.cat !== cat) return;
      var k = L.memoKey(p.memo);
      m[k] = (m[k] || 0) + L.signed(p);
    });
    return Object.keys(m).map(function (k) { return { name: k, v: m[k] }; }).sort(function (a, b) { return b.v - a.v; });
  }

  function renderOverview() {
    var pane = $("pane-overview");
    if (!S.months.length) return emptyState("pane-overview");
    var ms = rangeMonths();
    var ps = [];
    ms.forEach(function (m) { ps = ps.concat(S.byYm[m]); });
    var pl = L.plOf(ps);
    var vcount = {};
    ps.forEach(function (p) { vcount[p.vkey] = 1; });

    pane.innerHTML =
      '<div class="toolbar">' +
        '<label class="lp-field">開始月<select id="ov-from">' + monthOptions(S.range.from) + "</select></label>" +
        '<label class="lp-field">終了月<select id="ov-to">' + monthOptions(S.range.to) + "</select></label>" +
        '<button class="lp-btn small" id="ov-all">全期間</button>' +
        '<span class="spacer"></span><span class="muted" style="font-size:12px">' + ms.length + "ヶ月・" + Object.keys(vcount).length + "伝票</span>" +
      "</div>" +
      '<div class="grid-2">' +
        '<div class="lp-panel"><h2>損益 <small>' + ymLabel(ms[0]) + "〜" + ymLabel(ms[ms.length - 1]) + "</small></h2>" + plSheet(pl) + "</div>" +
        '<div class="lp-panel"><h2>売上から経常利益まで</h2><div class="chart" id="ov-wf"></div></div>' +
      "</div>" +
      '<div class="lp-panel"><h2>月次推移 <small>棒をクリックで月別へ</small></h2>' +
        '<div class="legend-row"><span><i style="background:' + COLORS.sales + '"></i>売上高</span><span><i style="background:' + COLORS.cost + '"></i>費用（原価＋販管費）</span><span><i class="line" style="background:' + COLORS.profit + '"></i>営業利益</span></div>' +
        '<div class="chart" id="ov-trend"></div></div>' +
      '<div class="lp-panel"><h2>月次損益表 <small>月をクリックで月別へ</small></h2><div class="lp-scroll" id="ov-table"></div></div>' +
      '<div class="grid-2">' +
        '<div class="lp-panel"><h2>費用の大きい科目 <small>クリックで科目へ</small></h2><div id="ov-cost"></div></div>' +
        '<div class="lp-panel"><h2>売上の相手先 <small>摘要から推定</small></h2><div id="ov-sales"></div></div>' +
      "</div>";

    $("ov-from").onchange = function () { S.range.from = this.value; if (S.range.from > S.range.to) S.range.to = S.range.from; saveUi(); renderOverview(); };
    $("ov-to").onchange = function () { S.range.to = this.value; if (S.range.to < S.range.from) S.range.from = S.range.to; saveUi(); renderOverview(); };
    $("ov-all").onclick = function () { S.range.from = S.months[0]; S.range.to = S.months[S.months.length - 1]; saveUi(); renderOverview(); };

    CH.waterfall($("ov-wf"), waterfallSteps(pl));
    var goMonth = function (i) { S.month = ms[i]; setTab("month"); };
    CH.barLine($("ov-trend"), {
      labels: ms.map(ymShort), active: -1, onClick: goMonth,
      bars: [
        { name: "売上高", color: COLORS.sales, values: ms.map(function (m) { return S.plYm[m]["売上高"]; }) },
        { name: "費用", color: COLORS.cost, values: ms.map(function (m) { return S.plYm[m]["売上原価"] + S.plYm[m]["販管費"]; }) }
      ],
      line: { name: "営業利益", color: COLORS.profit, values: ms.map(function (m) { return S.plYm[m]["営業利益"]; }) }
    });

    // 月次損益表
    var lines = [["売上高", ""], ["売上原価", ""], ["売上総利益", "subtotal"], ["販管費", ""], ["営業利益", "subtotal"],
      ["営業外収益", ""], ["営業外費用", ""], ["経常利益", "total"]];
    var t = '<table class="lp-table pl-table"><thead><tr><th></th>' +
      ms.map(function (m) { return '<th class="num m" data-ym="' + m + '">' + ymShort(m) + "</th>"; }).join("") +
      '<th class="num">合計</th></tr></thead><tbody>';
    lines.forEach(function (ln) {
      t += '<tr class="' + ln[1] + '"><td>' + ln[0] + "</td>" + ms.map(function (m) { return '<td class="num">' + yen(S.plYm[m][ln[0]]) + "</td>"; }).join("") +
        '<td class="num">' + yen(pl[ln[0]]) + "</td></tr>";
    });
    $("ov-table").innerHTML = t + "</tbody></table>";
    $("ov-table").querySelectorAll("th[data-ym]").forEach(function (th) {
      th.onclick = function () { S.month = th.dataset.ym; setTab("month"); };
    });

    var costs = accountTotals(ps, ["売上原価", "販管費"]);
    $("ov-cost").innerHTML = rankHtml(costs, "cost", 12);
    bindRank($("ov-cost"), costs, function (x) { openAccount(x.code, null); });
    var sales = partnerTotals(ps, "売上高");
    $("ov-sales").innerHTML = sales.length ? rankHtml(sales, "", 12) : '<div class="muted">売上の仕訳がありません。</div>';
    bindRank($("ov-sales"), sales, function (x) {
      var sc = accountTotals(ps, ["売上高"])[0];
      if (sc) openAccount(sc.code, null, x.name);
    });
  }

  // ===== 月別 =====
  function renderMonth() {
    var pane = $("pane-month");
    if (!S.months.length) return emptyState("pane-month");
    var ym = S.month, idx = S.months.indexOf(ym);
    var prevYm = S.months[idx - 1] || null;
    var ps = S.byYm[ym] || [], pps = prevYm ? S.byYm[prevYm] : [];
    var pl = S.plYm[ym], ppl = prevYm ? S.plYm[prevYm] : null;
    var vset = {};
    ps.forEach(function (p) { vset[p.vkey] = 1; });

    pane.innerHTML =
      '<div class="toolbar"><span class="month-nav">' +
        '<button class="lp-btn small" id="m-prev" ' + (idx <= 0 ? "disabled" : "") + ' aria-label="前月">‹</button>' +
        '<span class="cur">' + ymLabel(ym) + "</span>" +
        '<button class="lp-btn small" id="m-next" ' + (idx >= S.months.length - 1 ? "disabled" : "") + ' aria-label="翌月">›</button>' +
      '</span><select id="m-sel" class="lp-field" style="padding:4px">' + monthOptions(ym) + "</select>" +
      '<span class="spacer"></span><span class="muted" style="font-size:12px">' + Object.keys(vset).length + "伝票" + (prevYm ? "・比較: " + ymLabel(prevYm) : "") + "</span></div>" +
      '<div class="grid-2">' +
        '<div class="lp-panel"><h2>損益</h2>' + plSheet(pl, ppl) + "</div>" +
        '<div class="lp-panel"><h2>売上から経常利益まで</h2><div class="chart" id="m-wf"></div></div>' +
      "</div>" +
      '<div class="lp-panel"><h2>科目別 <small>クリックで科目へ</small></h2><div class="lp-scroll" id="m-acc"></div></div>' +
      '<div class="lp-panel"><h2>金額の大きい伝票 <small>クリックでExcelの行へ</small></h2><div class="lp-scroll" id="m-big"></div></div>';

    $("m-prev").onclick = function () { S.month = S.months[idx - 1]; saveUi(); renderMonth(); };
    $("m-next").onclick = function () { S.month = S.months[idx + 1]; saveUi(); renderMonth(); };
    $("m-sel").onchange = function () { S.month = this.value; saveUi(); renderMonth(); };
    CH.waterfall($("m-wf"), waterfallSteps(pl));

    // 科目別：当月・前月・増減
    var cur = {}, prv = {};
    ps.forEach(function (p) { if (isPL(p.cat)) cur[p.code] = (cur[p.code] || 0) + L.signed(p); });
    pps.forEach(function (p) { if (isPL(p.cat)) prv[p.code] = (prv[p.code] || 0) + L.signed(p); });
    var codes = Object.keys(Object.assign({}, cur, prv));
    var html = '<table class="lp-table"><thead><tr><th>科目</th><th class="num">当月</th><th class="num">前月</th><th class="num">増減</th></tr></thead><tbody>';
    PL_CATS.forEach(function (cat) {
      var cs = codes.filter(function (c) { return catOf(c) === cat; })
        .sort(function (a, b) { return Math.abs(cur[b] || 0) - Math.abs(cur[a] || 0); });
      if (!cs.length) return;
      var sc = 0, sp = 0;
      cs.forEach(function (c) { sc += cur[c] || 0; sp += prv[c] || 0; });
      html += '<tr class="subtotal"><td>' + cat + '</td><td class="num">' + yen(sc) + '</td><td class="num">' + (prevYm ? yen(sp) : "—") + '</td><td class="num">' + (prevYm ? diffYen(sc - sp, !L.CREDIT_NATURE[cat]) : "") + "</td></tr>";
      cs.forEach(function (c) {
        html += '<tr class="clickable" data-code="' + esc(c) + '"><td>' + esc(nameOf(c)) + '</td><td class="num">' + yen(cur[c]) +
          '</td><td class="num">' + (prevYm ? yen(prv[c]) : "—") + '</td><td class="num">' + (prevYm ? diffYen((cur[c] || 0) - (prv[c] || 0), !L.CREDIT_NATURE[cat]) : "") + "</td></tr>";
      });
    });
    $("m-acc").innerHTML = html + "</tbody></table>";
    $("m-acc").querySelectorAll("tr[data-code]").forEach(function (tr) {
      tr.onclick = function () { openAccount(tr.dataset.code, ym); };
    });

    // 金額の大きい伝票
    var byV = L.groupByVoucher(S.entries.filter(function (e) { return L.ymOf(e.date) === ym; }));
    var vs = Object.keys(byV).map(function (k) {
      var rows = byV[k], amt = 0, memos = [], accs = {};
      rows.forEach(function (r) {
        amt += Number(r.fields[L.C.D_AMT] || 0);
        if (r.fields[L.C.MEMO] && memos.indexOf(r.fields[L.C.MEMO]) < 0) memos.push(r.fields[L.C.MEMO]);
        if (r.fields[L.C.D_NAME]) accs[r.fields[L.C.D_NAME]] = 1;
      });
      var memo = memos.slice(0, 2).join(" / ") + (memos.length > 2 ? " ほか" + (memos.length - 2) + "件" : "");
      return { date: rows[0].date, amt: amt, memo: memo, acc: Object.keys(accs).join("・"), sheet: rows[0].sheet, row: rows[0].rowIndex };
    }).sort(function (a, b) { return b.amt - a.amt; }).slice(0, 10);
    $("m-big").innerHTML = '<table class="lp-table"><thead><tr><th>日付</th><th>借方科目</th><th>摘要</th><th class="num">金額</th></tr></thead><tbody>' +
      vs.map(function (v, i) {
        return '<tr class="clickable" data-i="' + i + '"><td class="num">' + v.date.slice(5) + "</td><td>" + esc(v.acc) + "</td><td>" + esc(v.memo) + '</td><td class="num">' + yen(v.amt) + "</td></tr>";
      }).join("") + "</tbody></table>";
    $("m-big").querySelectorAll("tr[data-i]").forEach(function (tr) {
      tr.onclick = function () { var v = vs[Number(tr.dataset.i)]; jumpTo(v.sheet, v.row); };
    });
  }

  // ===== 科目 =====
  function openAccount(code, ym, q) {
    S.acc.code = code; S.acc.ym = ym || null; S.acc.q = q || ""; S.acc.limit = 200;
    setTab("account");
  }
  function renderAccount() {
    var pane = $("pane-account");
    if (!S.months.length) return emptyState("pane-account");
    var codes = {};
    S.postings.forEach(function (p) { codes[p.code] = 1; });
    var all = Object.keys(codes).sort();
    if (!S.acc.code || !codes[S.acc.code]) {
      var top = accountTotals(S.postings, ["販管費"])[0];
      S.acc.code = top ? top.code : all[0];
    }
    var code = S.acc.code, cat = catOf(code), pl = isPL(cat);
    var mine = S.postings.filter(function (p) { return p.code === code; });

    var opts = PL_CATS.concat(["資産", "負債", "純資産"]).map(function (c) {
      var cs = all.filter(function (x) { return catOf(x) === c; });
      if (!cs.length) return "";
      return '<optgroup label="' + c + '">' + cs.map(function (x) {
        return '<option value="' + esc(x) + '"' + (x === code ? " selected" : "") + ">" + esc(x + " " + nameOf(x)) + "</option>";
      }).join("") + "</optgroup>";
    }).join("");

    pane.innerHTML =
      '<div class="acc-head"><label class="lp-field">勘定科目<select id="a-sel">' + opts + "</select></label>" +
        (S.acc.ym ? '<span class="chip">' + ymLabel(S.acc.ym) + ' <button id="a-clear-ym" aria-label="月の絞り込みを解除">×</button></span>' : "") +
        '<span class="muted" style="font-size:12px">区分: ' + cat + "（科目マスタで変更できます）</span></div>" +
      '<div class="lp-panel"><h2>月次推移 <small>棒をクリックで月を絞り込み</small></h2>' +
        (pl ? "" : '<div class="legend-row"><span><i style="background:' + COLORS.dr + '"></i>借方</span><span><i style="background:' + COLORS.cr + '"></i>貸方</span><span><i class="line" style="background:' + COLORS.profit + '"></i>純増減</span></div>') +
        '<div class="chart" id="a-chart"></div></div>' +
      '<div class="lp-panel"><h2>内訳 <span class="tabs-mini" id="a-bd">' +
        '<button data-bd="partner">相手先（摘要）</button><button data-bd="counter">相手科目</button><button data-bd="sub">補助科目</button></span></h2><div id="a-breakdown"></div></div>' +
      '<div class="lp-panel"><h2>明細 <small>クリックでExcelの行へ</small></h2>' +
        '<input class="search" id="a-q" placeholder="摘要・相手科目で絞り込み" value="' + esc(S.acc.q) + '">' +
        '<div class="lp-scroll" id="a-list" style="margin-top:8px"></div></div>';

    $("a-sel").onchange = function () { S.acc.code = this.value; S.acc.q = ""; S.acc.limit = 200; saveUi(); renderAccount(); };
    if ($("a-clear-ym")) $("a-clear-ym").onclick = function () { S.acc.ym = null; renderAccount(); };

    // 月次グラフ
    var ms = S.months;
    var dr = ms.map(function () { return 0; }), cr = dr.slice(), net = dr.slice();
    mine.forEach(function (p) {
      var i = ms.indexOf(p.ym);
      if (p.side === "D") dr[i] += p.amt; else cr[i] += p.amt;
      net[i] += L.signed(p);
    });
    var onBar = function (i) { S.acc.ym = S.acc.ym === ms[i] ? null : ms[i]; S.acc.limit = 200; renderAccount(); };
    CH.barLine($("a-chart"), pl ? {
      labels: ms.map(ymShort), active: ms.indexOf(S.acc.ym), onClick: onBar,
      bars: [{ name: nameOf(code), color: COLORS.sales, values: net }]
    } : {
      labels: ms.map(ymShort), active: ms.indexOf(S.acc.ym), onClick: onBar,
      bars: [{ name: "借方", color: COLORS.dr, values: dr }, { name: "貸方", color: COLORS.cr, values: cr }],
      line: { name: "純増減", color: COLORS.profit, values: net }
    });

    var scoped = mine.filter(function (p) { return !S.acc.ym || p.ym === S.acc.ym; });

    // 内訳
    function drawBreakdown() {
      document.querySelectorAll("#a-bd button").forEach(function (b) { b.classList.toggle("active", b.dataset.bd === S.acc.breakdown); });
      var m = {};
      scoped.forEach(function (p) {
        var k = S.acc.breakdown === "partner" ? L.memoKey(p.memo) : S.acc.breakdown === "counter" ? (p.counter || "（なし）") : (p.sub || "（補助なし）");
        m[k] = (m[k] || 0) + L.signed(p);
      });
      var list = Object.keys(m).map(function (k) { return { name: k, v: m[k] }; }).sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); });
      $("a-breakdown").innerHTML = list.length ? rankHtml(list, pl && cat !== "売上高" ? "cost" : "", 15) : '<div class="muted">明細がありません。</div>';
      bindRank($("a-breakdown"), list, function (x) {
        S.acc.q = x.name.replace(/^（.*）$/, ""); S.acc.limit = 200;
        $("a-q").value = S.acc.q; drawList();
      });
    }
    document.querySelectorAll("#a-bd button").forEach(function (b) {
      b.onclick = function () { S.acc.breakdown = b.dataset.bd; drawBreakdown(); };
    });
    drawBreakdown();

    // 明細
    function drawList() {
      var q = S.acc.q.trim().normalize("NFKC");
      var rows = scoped.filter(function (p) {
        if (!q) return true;
        return (String(p.memo || "") + " " + p.counter + " " + (p.sub || "")).normalize("NFKC").indexOf(q) >= 0;
      }).sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
      var sd = 0, sc = 0;
      rows.forEach(function (p) { if (p.side === "D") sd += p.amt; else sc += p.amt; });
      var shown = rows.slice(0, S.acc.limit);
      $("a-list").innerHTML = '<table class="lp-table"><thead><tr><th>日付</th><th>相手科目</th><th>摘要</th><th class="num">借方</th><th class="num">貸方</th></tr></thead><tbody>' +
        shown.map(function (p, i) {
          return '<tr class="clickable" data-i="' + i + '"><td class="num">' + p.date.slice(2) + "</td><td>" + esc(p.counter) + (p.sub ? '<div class="muted" style="font-size:11px">' + esc(p.sub) + "</div>" : "") +
            "</td><td>" + esc(p.memo) + '</td><td class="num">' + (p.side === "D" ? yen(p.amt) : "") + '</td><td class="num">' + (p.side === "C" ? yen(p.amt) : "") + "</td></tr>";
        }).join("") +
        '<tr class="total"><td colspan="3">' + rows.length + "件の合計</td><td class=\"num\">" + yen(sd) + '</td><td class="num">' + yen(sc) + "</td></tr></tbody></table>" +
        (rows.length > shown.length ? '<div class="more"><button class="lp-btn small" id="a-more">さらに ' + Math.min(200, rows.length - shown.length) + " 件表示</button></div>" : "");
      $("a-list").querySelectorAll("tr[data-i]").forEach(function (tr) {
        tr.onclick = function () { var p = shown[Number(tr.dataset.i)]; jumpTo(p.sheet, p.rowIndex); };
      });
      if ($("a-more")) $("a-more").onclick = function () { S.acc.limit += 200; drawList(); };
    }
    var qTimer;
    $("a-q").oninput = function () {
      var v = this.value; clearTimeout(qTimer);
      qTimer = setTimeout(function () { S.acc.q = v; S.acc.limit = 200; drawList(); }, 200);
    };
    drawList();
    saveUi();
  }

  // ===== 取込 =====
  function renderImport() {
    var h = S.history.slice().reverse().slice(0, 50);
    $("imp-history").innerHTML = h.length ?
      '<table class="lp-table"><thead><tr><th>取込日時</th><th>年月</th><th class="num">追加</th><th class="num">変更</th><th class="num">削除</th><th class="num">行数</th><th>CSV</th></tr></thead><tbody>' +
      h.map(function (r) {
        return "<tr><td style=\"white-space:nowrap\">" + esc(r[0]) + "</td><td>" + esc(r[3]) + '</td><td class="num">' + esc(r[4]) + '</td><td class="num">' + esc(r[5]) +
          '</td><td class="num">' + esc(r[6]) + '</td><td class="num">' + esc(r[8]) + "</td><td>" + esc(r[1]) + "</td></tr>";
      }).join("") + "</tbody></table>" : '<div class="muted">まだ取り込んでいません。</div>';
  }
  function impMsg(cls, text) { $("imp-msg").innerHTML = text ? '<div class="lp-note ' + cls + '">' + esc(text) + "</div>" : ""; }
  function readCsv(file) {
    file.arrayBuffer().then(function (buf) {
      var parsed = L.buildEntries(L.decodeCsvBuffer(buf));
      S.imp.parsed = parsed; S.imp.name = file.name;
      $("csv-name").textContent = file.name + "（" + parsed.voucherCount + "伝票・" + parsed.entries.length + "行）";
      $("p-start").value = L.monthStart(parsed.minDate).replace(/\//g, "-");
      $("p-end").value = L.monthEnd(parsed.maxDate).replace(/\//g, "-");
      $("csv-info").textContent = "CSVの日付範囲: " + parsed.minDate + "〜" + parsed.maxDate + "。月の途中までのCSVなら終了日を合わせてください。";
      $("imp-period").classList.remove("hidden");
      impMsg("", "");
      runDiff();
    }).catch(function (e) { impMsg("err", e.message || String(e)); });
  }
  function runDiff() {
    if (!S.imp.parsed) return;
    var p = { start: $("p-start").value.replace(/-/g, "/"), end: $("p-end").value.replace(/-/g, "/") };
    if (!p.start || !p.end || p.start > p.end) { impMsg("err", "期間の開始日と終了日を確認してください。"); return; }
    S.imp.diff = L.computeDiff(S.entries, S.imp.parsed, p);
    $("imp-diff-panel").classList.remove("hidden");
    S.imp.view = DV.render($("imp-diff"), S.imp.diff, {
      onChange: function (sel) {
        $("btn-apply").disabled = !sel.length;
        $("apply-msg").textContent = sel.length ? sel.join("、") + " を差し替えます" : "";
      }
    });
  }
  function colName(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function nowStr() {
    var d = new Date(), p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "/" + p(d.getMonth() + 1) + "/" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
  }

  // 1シートずつ Excel.run を分け、400行ごとに sync（大きな一括書き込みで画面が白くなるのを避ける）
  async function writeMonthSheet(ym, rows, header, width, at) {
    var name = L.sheetNameOf(ym);
    await Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
      ws.load("isNullObject");
      await ctx.sync();
      if (!rows.length) {
        if (!ws.isNullObject) { ws.delete(); await ctx.sync(); }
        return;
      }
      if (ws.isNullObject) {
        ws = ctx.workbook.worksheets.add(name);
      } else {
        var used = ws.getUsedRangeOrNullObject();
        used.load("isNullObject");
        await ctx.sync();
        if (!used.isNullObject) used.clear();
      }
      var ncol = L.META_N + width, last = colName(ncol);
      var hdr = ws.getRange("A1:" + last + "1");
      hdr.values = [L.META_COLS.concat(header)];
      hdr.format.font.bold = true;
      hdr.format.fill.color = "#ECEEFB";
      await ctx.sync();
      var fm = L.columnFormats(width);
      for (var i = 0; i < rows.length; i += 400) {
        var chunk = rows.slice(i, i + 400);
        var rg = ws.getRange("A" + (i + 2) + ":" + last + (i + 1 + chunk.length));
        rg.numberFormat = chunk.map(function () { return fm; });
        rg.values = chunk.map(function (e) { return L.entryToSheetRow(e, at); });
        await ctx.sync();
      }
      ws.freezePanes.freezeRows(1);
      ws.getRange("A:D").columnHidden = true; // 伝票キー等の管理列
      ws.getRange(colName(L.META_N + L.C.MEMO + 1) + ":" + colName(L.META_N + L.C.MEMO + 1)).format.columnWidth = 220;
      await ctx.sync();
    });
  }
  async function writeMasterAndHistory(parsed, histRows) {
    var acc = L.collectAccounts(parsed.entries);
    Object.keys(acc).forEach(function (code) {
      if (!S.master[code]) S.master[code] = { name: acc[code], cat: L.defaultCategory(code) };
    });
    var mRows = [["科目コード", "勘定科目名", "区分"]].concat(Object.keys(S.master).sort().map(function (c) {
      return [String(c), S.master[c].name, S.master[c].cat];
    }));
    await Excel.run(async function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(L.MASTER_SHEET);
      ws.load("isNullObject"); await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add(L.MASTER_SHEET);
      var used = ws.getUsedRangeOrNullObject(); used.load("isNullObject"); await ctx.sync();
      if (!used.isNullObject) used.clear();
      var rg = ws.getRange("A1:C" + mRows.length);
      rg.numberFormat = mRows.map(function () { return ["@", "@", "@"]; });
      rg.values = mRows;
      ws.getRange("A1:C1").format.font.bold = true;
      ws.getRange("A1:C1").format.fill.color = "#ECEEFB";
      await ctx.sync();

      var hs = ctx.workbook.worksheets.getItemOrNullObject(L.HISTORY_SHEET);
      hs.load("isNullObject"); await ctx.sync();
      var start = 2;
      if (hs.isNullObject) {
        hs = ctx.workbook.worksheets.add(L.HISTORY_SHEET);
        var h = hs.getRange("A1:I1");
        h.values = [["取込日時", "CSVファイル", "指定期間", "年月", "追加", "変更", "削除", "同一", "反映後の行数"]];
        h.format.font.bold = true; h.format.fill.color = "#ECEEFB";
      } else {
        var u = hs.getUsedRangeOrNullObject(true); u.load("isNullObject,rowCount,rowIndex"); await ctx.sync();
        start = u.isNullObject ? 2 : u.rowIndex + u.rowCount + 1;
      }
      if (histRows.length) {
        var hr = hs.getRange("A" + start + ":I" + (start + histRows.length - 1));
        hr.numberFormat = histRows.map(function () { return ["@", "@", "@", "@", "0", "0", "0", "0", "0"]; });
        hr.values = histRows;
      }
      await ctx.sync();
    });
  }
  async function orderSheets() {
    await Excel.run(async function (ctx) {
      var wss = ctx.workbook.worksheets; wss.load("items/name"); await ctx.sync();
      var names = wss.items.map(function (w) { return w.name; });
      var months = names.filter(function (n) { return L.ymFromSheet(n); }).sort();
      // 年月順に末尾へ送る（取込履歴・科目マスタなど他のシートは前側に残る）
      months.forEach(function (n) { wss.getItem(n).position = names.length - 1; });
      await ctx.sync();
    });
  }
  async function applyImport() {
    var sel = S.imp.view.getSelected();
    if (!sel.length) return;
    var diff = S.imp.diff, parsed = S.imp.parsed;
    var plan = L.planApply(S.entries, diff, sel);
    if (plan.duplicates.length) {
      impMsg("err", "同じ仕訳が複数の月に入るため中止しました（" + plan.duplicates.length + "件）。連動している月をすべて選んでください。");
      return;
    }
    $("btn-apply").disabled = true;
    var at = nowStr(), yms = Object.keys(plan.sheets).sort();
    try {
      for (var i = 0; i < yms.length; i++) {
        $("apply-msg").textContent = yms[i] + " を書き込み中…（" + (i + 1) + "/" + yms.length + "）";
        await writeMonthSheet(yms[i], plan.sheets[yms[i]], diff.header, diff.width, at);
      }
      var hist = sel.map(function (ym) {
        var m = diff.months[ym] || { add: 0, chg: 0, del: 0, same: 0 };
        return [at, S.imp.name, diff.period.start + "〜" + diff.period.end, ym, m.add, m.chg, m.del, m.same, (plan.sheets[ym] || []).length];
      });
      await writeMasterAndHistory(parsed, hist);
      await orderSheets();
      rebuild(await loadBook());
      impMsg("ok", sel.join("、") + " を差し替えました。");
      runDiff(); // 差し替え後は差分0になることを確認表示
      renderImport();
    } catch (e) {
      console.error(e);
      impMsg("err", "書き込み中に止まりました: " + (e.message || e) + "。↻で台帳を読み直し、差分を確認してからもう一度実行してください。");
      $("btn-apply").disabled = false;
    }
    $("apply-msg").textContent = "";
  }

  // ===== 起動 =====
  function bind() {
    document.querySelectorAll(".tabs button").forEach(function (b) { b.onclick = function () { setTab(b.dataset.tab); }; });
    $("btn-reload").onclick = reload;
    var d = $("drop-csv"), inp = $("file-csv");
    inp.onchange = function () { if (inp.files[0]) readCsv(inp.files[0]); inp.value = ""; };
    d.addEventListener("dragover", function (e) { e.preventDefault(); d.classList.add("over"); });
    d.addEventListener("dragleave", function () { d.classList.remove("over"); });
    d.addEventListener("drop", function (e) { e.preventDefault(); d.classList.remove("over"); if (e.dataTransfer.files[0]) readCsv(e.dataTransfer.files[0]); });
    $("btn-diff").onclick = runDiff;
    $("btn-apply").onclick = applyImport;
    var rt;
    window.addEventListener("resize", function () {
      clearTimeout(rt);
      rt = setTimeout(function () { if (S.tab !== "import") render(); }, 250);
    });
    $("version-label").textContent = APP_VERSION + " / core " + L.CORE_VERSION;
  }

  Office.onReady(function () {
    restoreUi();
    bind();
    setTab(S.tab);
    reload();
  });
  window.__lp = S;
})();
