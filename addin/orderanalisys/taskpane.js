/* 受注件数分析アドイン：作業ウィンドウ */
/* global Office, Excel, Core */
(function () {
  "use strict";

  const C = Core;
  const COLOR = {
    autoFill: "#FFF4D6", autoFont: "#6B4A00", headFill: "#F1F7F3",
    man: "#2F55C8", edi: "#17917E", sub: "#9AA3AF"
  };

  const st = {
    sources: [],          // [{name, records}]
    noteMap: new Map(),
    groupRows: [],        // parseGroupSheet の結果
    groupMap: new Map(),  // cd → グループ名
    hasGroupSheet: false,
    suggestion: null,     // グループシートがないときの候補
    period: null,
    view: "group",
    pending: 0,
    handler: null,
    writing: false,
    quietUntil: 0
  };

  const $ = (id) => document.getElementById(id);
  const fmt = (n) => Math.round(n).toLocaleString("ja-JP");
  const pct = (v) => v.toFixed(1) + "%";
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function say(text, isErr) {
    const m = $("msg");
    m.textContent = text || "";
    m.className = "msg" + (isErr ? " err" : "");
  }

  async function run(label, fn) {
    say("");
    try { await fn(); }
    catch (e) {
      console.error(e);
      say(label + "できませんでした：" + (e && e.message ? e.message : e), true);
    }
  }

  /** アドイン自身の書き込みで変更イベントが起きないようにする。
      enableEvents は ExcelApi 1.8 以上のため、1.7 では一定時間イベントを無視して代わりにする */
  async function setEvents(ctx, on) {
    if (!on) st.writing = true;
    else { st.writing = false; st.quietUntil = Date.now() + 1500; }
    try { ctx.runtime.enableEvents = on; await ctx.sync(); } catch (e) { /* ExcelApi 1.8 未満 */ }
  }

  /* ---------- 読み込み ---------- */

  async function readWorkbook() {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const reads = [];
      let groupRange = null;
      sheets.items.forEach((ws) => {
        if (ws.name === C.GROUP_SHEET) {
          groupRange = ws.getUsedRangeOrNullObject(true);
          groupRange.load("values");
        } else if (!C.isSystemSheet(ws.name)) {
          const r = ws.getUsedRangeOrNullObject(true);
          r.load("values");
          reads.push({ name: ws.name, range: r });
        }
      });
      await ctx.sync();

      st.sources = [];
      reads.forEach((x) => {
        if (x.range.isNullObject) return;
        const h = C.findHeader(x.range.values);
        if (!h || h.isDaily) return;
        const records = C.parseRecords(x.range.values, h);
        if (records.length) st.sources.push({ name: x.name, records: records });
      });
      st.noteMap = C.buildNoteMap(st.sources);

      st.hasGroupSheet = !!groupRange && !groupRange.isNullObject;
      st.groupRows = st.hasGroupSheet ? C.parseGroupSheet(groupRange.values) : [];
      st.groupMap = new Map(st.groupRows.filter((r) => r.group).map((r) => [r.cd, r.group]));
      st.suggestion = st.hasGroupSheet ? null
        : C.suggestGroups(C.collectCustomers(st.sources, C.primarySource(st.sources)));
    });
  }

  /* ---------- グループシート ---------- */

  async function createGroupSheet() {
    await Excel.run(async (ctx) => {
      const exists = ctx.workbook.worksheets.getItemOrNullObject(C.GROUP_SHEET);
      await ctx.sync();
      if (!exists.isNullObject) throw new Error("「" + C.GROUP_SHEET + "」シートはすでにあります");

      const primary = C.primarySource(st.sources);
      const customers = C.collectCustomers(st.sources, primary);
      const sug = C.suggestGroups(customers);
      const rows = C.buildInitialGroupRows(customers, sug);

      await setEvents(ctx, false);
      const ws = ctx.workbook.worksheets.add(C.GROUP_SHEET);
      const head = [["請求先CD", "請求先", "グループ名", "受注件数（" + primary.name + "）", "自動入力値"]];
      const body = rows.map((r) => [r.cd, r.name, r.group, r.n, r.group]);
      const n = body.length;
      ws.getRange("A1:E1").values = head;
      ws.getRange("A2:E" + (n + 1)).values = body;

      const hr = ws.getRange("A1:E1");
      hr.format.font.bold = true;
      hr.format.fill.color = COLOR.headFill;
      ws.getRange("D2:D" + (n + 1)).numberFormat = Array(n).fill(["#,##0"]);
      ws.getRange("A:A").format.columnWidth = 70;
      ws.getRange("B:B").format.columnWidth = 300;
      ws.getRange("C:C").format.columnWidth = 170;
      ws.getRange("D:D").format.columnWidth = 150;
      ws.getRange("E:E").columnHidden = true;
      ws.freezePanes.freezeRows(1);
      ws.autoFilter.apply(ws.getRange("A1:D" + (n + 1)));

      const autoCount = rows.filter((r) => r.group).length;
      if (autoCount) {
        const ar = ws.getRange("C2:C" + (autoCount + 1));
        ar.format.fill.color = COLOR.autoFill;
        ar.format.font.color = COLOR.autoFont;
      }
      await refreshList(ctx, ws, n, Array.from(new Set(rows.map((r) => r.group).filter(Boolean))));
      ws.activate();
      await ctx.sync();
      await setEvents(ctx, true);
      say("グループシートを作成しました。黄色のセルが自動で入力したグループ名です。");
    });
    await afterGroupChange();
  }

  /** プルダウン用のグループ名一覧と、C列の入力規則を更新する */
  async function refreshList(ctx, ws, nRows, names) {
    let ls = ctx.workbook.worksheets.getItemOrNullObject(C.LIST_SHEET);
    await ctx.sync();
    if (ls.isNullObject) ls = ctx.workbook.worksheets.add(C.LIST_SHEET);
    ls.visibility = Excel.SheetVisibility.hidden;
    ls.getRange("A:A").clear(Excel.ClearApplyTo.contents);
    if (!names.length) return;
    names.sort((a, b) => a.localeCompare(b, "ja"));
    ls.getRange("A1:A" + names.length).values = names.map((x) => [x]);
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.8")) return; // 入力規則（プルダウン）は 1.8 以上
    const target = ws.getRange("C2:C" + Math.max(nRows + 1, 2));
    target.dataValidation.clear();
    target.dataValidation.rule = {
      list: { inCellDropDown: true, source: "='" + C.LIST_SHEET + "'!$A$1:$A$" + names.length }
    };
    target.dataValidation.errorAlert = { showAlert: false };
  }

  /** C列の色を状態に合わせて塗り直す（同じ状態が続く範囲をまとめて処理） */
  function paintRuns(ws, rows) {
    let i = 0;
    while (i < rows.length) {
      const isAuto = !!rows[i].group && rows[i].auto === rows[i].group;
      let j = i;
      while (j + 1 < rows.length && rows[j + 1].row === rows[j].row + 1 &&
        (!!rows[j + 1].group && rows[j + 1].auto === rows[j + 1].group) === isAuto) j++;
      const r = ws.getRange("C" + (rows[i].row + 1) + ":C" + (rows[j].row + 1));
      if (isAuto) { r.format.fill.color = COLOR.autoFill; r.format.font.color = COLOR.autoFont; }
      else { r.format.fill.clear(); r.format.font.color = "#000000"; }
      i = j + 1;
    }
  }

  /** 再集計の前に、グループシートを整える（新しい請求先の追加・色・プルダウン） */
  async function syncGroupSheet() {
    if (!st.hasGroupSheet) return 0;
    let added = 0;
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(C.GROUP_SHEET);
      await setEvents(ctx, false);

      const known = new Set(st.groupRows.map((r) => r.cd));
      const customers = C.collectCustomers(st.sources, C.primarySource(st.sources));
      const fresh = customers.filter((c) => !known.has(c.cd)).sort((a, b) => b.n - a.n);
      if (fresh.length) {
        const sug = C.suggestForNew(fresh, st.groupRows);
        const used = ws.getUsedRange(true);
        used.load("rowCount");
        await ctx.sync();
        const start = used.rowCount + 1;
        ws.getRange("A" + start + ":E" + (start + fresh.length - 1)).values =
          fresh.map((c) => [c.cd, c.name, sug.get(c.cd) || "", c.n, sug.get(c.cd) || ""]);
        fresh.forEach((c, k) => st.groupRows.push({ row: start - 1 + k, cd: c.cd, name: c.name, group: sug.get(c.cd) || "", auto: sug.get(c.cd) || "" }));
        added = fresh.length;
      }
      paintRuns(ws, st.groupRows);
      const lastRow = st.groupRows.length ? st.groupRows[st.groupRows.length - 1].row + 1 : 1;
      await refreshList(ctx, ws, lastRow - 1, C.groupStatus(st.groupRows).names);
      await ctx.sync();
      await setEvents(ctx, true);
    });
    st.groupMap = new Map(st.groupRows.filter((r) => r.group).map((r) => [r.cd, r.group]));
    return added;
  }

  /** グループシートの編集を検知する */
  async function watchGroupSheet() {
    if (!st.hasGroupSheet || st.handler) return;
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(C.GROUP_SHEET);
      st.handler = ws.onChanged.add(onGroupChanged);
      await ctx.sync();
    });
  }

  function cellsInColumnC(address) {
    const a = address.split("!").pop().replace(/\$/g, "");
    const m = a.match(/^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/);
    if (!m) return 1;
    const col = (s) => s ? s.split("").reduce((t, ch) => t * 26 + ch.charCodeAt(0) - 64, 0) : null;
    const c1 = col(m[1]) || 1, c2 = col(m[3]) || col(m[1]) || 16384;
    if (c1 > 3 || c2 < 3) return 0;
    const r1 = Math.max(Number(m[2] || 1), 2), r2 = Number(m[4] || m[2] || 1048576);
    return Math.max(0, Math.min(r2, 1048576) - r1 + 1);
  }

  async function onGroupChanged(ev) {
    if (st.writing || Date.now() < (st.quietUntil || 0)) return;
    const n = cellsInColumnC(ev.address || "");
    if (!n) return;
    st.pending += Math.min(n, 9999);
    $("bnChgText").textContent = "グループシートが変更されました（" + fmt(st.pending) + "か所）";
    $("bnChanged").hidden = false;
    // 編集したセルの色だけすぐ直す
    try {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(C.GROUP_SHEET);
        const r = ws.getRange(ev.address.split("!").pop()).getIntersectionOrNullObject(ws.getRange("C2:C1048576"));
        r.load("rowIndex,rowCount");
        await ctx.sync();
        if (r.isNullObject || r.rowCount > 2000) return;
        const blk = ws.getRangeByIndexes(r.rowIndex, 0, r.rowCount, 5);
        blk.load("values");
        await ctx.sync();
        await setEvents(ctx, false);
        paintRuns(ws, C.parseGroupSheet([[]].concat(blk.values)).map((x) => ({ ...x, row: r.rowIndex + x.row - 1 })));
        await ctx.sync();
        await setEvents(ctx, true);
      });
    } catch (e) { console.warn(e); }
  }

  async function confirmAll(selectedOnly) {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(C.GROUP_SHEET);
      let rowIndex = 1, rowCount = Math.max(st.groupRows.length, 1);
      if (selectedOnly) {
        const sel = ctx.workbook.getSelectedRange();
        sel.load("rowIndex,rowCount,worksheet/name");
        await ctx.sync();
        if (sel.worksheet.name !== C.GROUP_SHEET) throw new Error("グループシートで行を選んでください");
        rowIndex = Math.max(sel.rowIndex, 1);
        rowCount = sel.rowCount - (rowIndex - sel.rowIndex);
        if (rowCount <= 0) return;
      }
      await setEvents(ctx, false);
      const r = ws.getRangeByIndexes(rowIndex, 4, rowCount, 1);
      r.clear(Excel.ClearApplyTo.contents);
      const c = ws.getRangeByIndexes(rowIndex, 2, rowCount, 1);
      c.format.fill.clear();
      c.format.font.color = "#000000";
      await ctx.sync();
      await setEvents(ctx, true);
    });
    await recalc();
    say(selectedOnly ? "選んだ行を確認済みにしました。" : "黄色のセルをすべて確認済みにしました。");
  }

  async function openGroupSheet() {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(C.GROUP_SHEET).activate();
      await ctx.sync();
    });
  }

  /* ---------- 集計・表示 ---------- */

  function currentRecords() {
    const s = st.sources.find((x) => x.name === st.period) || st.sources[0];
    return s ? s.records : [];
  }

  function render() {
    const recs = currentRecords();
    const view = st.hasGroupSheet ? st.view : "customer";
    const agg = C.aggregate(recs, st.groupMap, st.noteMap, view);
    const base = C.aggregate(recs, st.groupMap, st.noteMap, "customer");

    // バナー
    const sugCount = st.suggestion ? new Set(st.suggestion.values()).size : 0;
    $("bnSuggest").hidden = st.hasGroupSheet || !sugCount;
    if (!st.hasGroupSheet && sugCount) {
      const preview = C.aggregate(recs, st.suggestion, st.noteMap, "group");
      const big = preview.units.filter((u) => u.isGroup).slice(0, 2).map((u) => u.name + "（" + u.members.length + "拠点）").join("、");
      $("bnSugTitle").textContent = "同じ会社の請求先が" + sugCount + "組あります";
      $("bnSugDesc").textContent = big + "など。まとめると上位10のシェアは " + pct(C.topShare(base, 10)) + " → " + pct(C.topShare(preview, 10)) + " になります。";
    }
    $("bnChanged").hidden = st.pending === 0;

    // グループシートの状態
    $("gstat").hidden = !st.hasGroupSheet;
    if (st.hasGroupSheet) {
      const s = C.groupStatus(st.groupRows);
      $("stAuto").textContent = fmt(s.auto);
      $("stOk").textContent = fmt(s.confirmed);
      $("stBlank").textContent = fmt(s.blank);
      $("btnConfirmAll").hidden = s.auto === 0;
      $("btnConfirmSel").hidden = s.auto === 0;
    }

    // 切り替え
    $("vGroup").disabled = !st.hasGroupSheet;
    $("vGroup").title = st.hasGroupSheet ? "" : "グループシートを作ると使えます";
    $("vGroup").setAttribute("aria-pressed", view === "group");
    $("vCust").setAttribute("aria-pressed", view === "customer");

    // KPI
    const t10 = C.topShare(agg, 10), t20 = C.topShare(agg, 20);
    $("k10").textContent = pct(t10);
    $("k20").textContent = pct(t20);
    const diff = (a, b) => view === "group" && Math.abs(a - b) >= 0.05 ? "請求先別より " + (a > b ? "+" : "") + (a - b).toFixed(1) + "pt" : "";
    $("k10d").textContent = diff(t10, C.topShare(base, 10));
    $("k20d").textContent = diff(t20, C.topShare(base, 20));
    $("kTot").textContent = fmt(agg.total) + "件";
    $("kUnits").textContent = view === "group" ? fmt(agg.units.length) + "単位（請求先 " + fmt(base.units.length) + "）" : fmt(agg.units.length) + "請求先";
    $("kMan").textContent = agg.total ? pct((agg.total - agg.linked) / agg.total * 100) : "—";
    $("kManN").textContent = fmt(agg.total - agg.linked) + "件";

    renderPareto(agg, view === "group" ? base : null);
    renderRanking(agg);
    $("footer").hidden = !recs.length;
  }

  function curve(agg, N, W, H) {
    const pts = ["M0 " + H];
    for (let i = 0; i < N; i++) {
      const u = agg.units[i];
      const c = u ? u.cum : 100;
      pts.push("L" + ((i + 1) / N * W).toFixed(1) + " " + (H - c / 100 * H).toFixed(1));
    }
    return pts.join(" ");
  }

  function renderPareto(agg, base) {
    const W = 316, H = 150, N = 100;
    const k = Number($("rk").value);
    const u = agg.units[Math.min(k, agg.units.length) - 1];
    const v = u ? u.cum : 0;
    const x = k / N * W, y = H - v / 100 * H;
    const lx = x > W - 110 ? x - 8 : x + 8, anchor = x > W - 110 ? "end" : "start";
    $("pareto").innerHTML =
      '<svg viewBox="-10 -8 340 178" role="img" aria-label="上位100までの累計シェア">' +
      '<line x1="0" y1="' + H + '" x2="' + W + '" y2="' + H + '" stroke="#C9CFD6"/>' +
      [25, 50, 75, 100].map((p) => '<line x1="0" y1="' + (H - p / 100 * H) + '" x2="' + W + '" y2="' + (H - p / 100 * H) + '" stroke="#EEF1F4"/>' +
        '<text x="-4" y="' + (H - p / 100 * H + 3) + '" font-size="9" fill="#7A8492" text-anchor="end">' + p + '</text>').join("") +
      (base ? '<path d="' + curve(base, N, W, H) + '" fill="none" stroke="' + COLOR.sub + '" stroke-width="2" stroke-dasharray="4 3"/>' : "") +
      '<path d="' + curve(agg, N, W, H) + '" fill="none" stroke="' + COLOR.man + '" stroke-width="2.5" stroke-linejoin="round"/>' +
      '<line x1="' + x + '" y1="0" x2="' + x + '" y2="' + H + '" stroke="#A65F0A" stroke-dasharray="2 3"/>' +
      '<circle cx="' + x + '" cy="' + y + '" r="5" fill="' + COLOR.man + '" stroke="#fff" stroke-width="2"/>' +
      '<text x="' + lx + '" y="' + Math.max(y - 6, 10) + '" font-size="11" fill="#1B2430" text-anchor="' + anchor + '">上位' + k + ' → ' + pct(v) + '</text>' +
      '<text x="0" y="166" font-size="9" fill="#7A8492">1</text><text x="' + W + '" y="166" font-size="9" fill="#7A8492" text-anchor="end">100</text>' +
      "</svg>";
    $("paretoLg").innerHTML = base ? '<span style="color:' + COLOR.man + '">━ グループ別</span><span style="color:#7A8492">┅ 請求先別</span>' : "";
    $("rkOut").textContent = "上位" + k + " → " + pct(v);
  }

  function renderRanking(agg) {
    const N = Number($("topn").value);
    const top = agg.units.slice(0, N);
    const max = top.length ? top[0].n : 1;
    $("rank").innerHTML = top.map((u, i) => {
      const m = u.n - u.linked;
      const tip = u.name + "：" + fmt(u.n) + "件（連携なし " + fmt(m) + "／連携 " + fmt(u.linked) + "）" +
        (u.isGroup ? "\n" + u.members.slice(0, 15).map((x) => "・" + x.name + " " + fmt(x.n)).join("\n") + (u.members.length > 15 ? "\nほか" + (u.members.length - 15) + "件" : "") : "");
      return '<div class="rk" title="' + esc(tip) + '"><span class="no">' + (i + 1) + '</span>' +
        '<span class="nm">' + esc(u.name) + (u.isGroup ? "<small>" + u.members.length + "</small>" : "") + "</span>" +
        '<span class="bar" style="width:' + (u.n / max * 100).toFixed(1) + '%">' +
        (m ? '<span style="flex:' + m + ';background:' + COLOR.man + '"></span>' : "") +
        (u.linked ? '<span style="flex:' + u.linked + ';background:' + COLOR.edi + '"></span>' : "") +
        '</span><span class="n">' + fmt(u.n) + "</span></div>";
    }).join("") || '<div class="msg">この期間のデータがありません。</div>';
  }

  /* ---------- 集計シートへの出力 ---------- */

  async function output() {
    const view = st.hasGroupSheet ? st.view : "customer";
    const agg = C.aggregate(currentRecords(), st.groupMap, st.noteMap, view);
    const label = view === "group" ? "グループ別" : "請求先別";
    const name = (C.OUTPUT_PREFIX + label + "_" + st.period).slice(0, 31).replace(/[\\/?*[\]:]/g, "");

    await Excel.run(async (ctx) => {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name);
      await ctx.sync();
      if (!old.isNullObject) old.delete();
      const ws = ctx.workbook.worksheets.add(name);

      ws.getRange("A1").values = [["受注件数の集計（" + label + "）　期間：" + st.period]];
      ws.getRange("A1").format.font.bold = true;
      ws.getRange("A1").format.font.size = 14;
      ws.getRange("A2:H2").values = [["受注件数", agg.total, "連携なし", agg.total - agg.linked, "上位10のシェア", C.topShare(agg, 10) / 100, "上位20のシェア", C.topShare(agg, 20) / 100]];
      ws.getRange("B2").numberFormat = [["#,##0"]];
      ws.getRange("D2").numberFormat = [["#,##0"]];
      ws.getRange("F2").numberFormat = [["0.0%"]];
      ws.getRange("H2").numberFormat = [["0.0%"]];

      const head = [["順位", label === "グループ別" ? "グループ／請求先" : "請求先", "請求先数", "受注件数", "連携なし", "連携", "割合", "累計割合"]];
      const rows = agg.units.map((u, i) => [i + 1, u.name, u.members.length, u.n, u.n - u.linked, u.linked, u.share / 100, u.cum / 100]);
      const n = rows.length;
      ws.getRange("A4:H4").values = head;
      ws.getRange("A5:H" + (n + 4)).values = rows;
      const hr = ws.getRange("A4:H4");
      hr.format.font.bold = true;
      hr.format.fill.color = COLOR.headFill;
      ws.getRange("D5:F" + (n + 4)).numberFormat = Array(n).fill(["#,##0", "#,##0", "#,##0"]);
      ws.getRange("G5:H" + (n + 4)).numberFormat = Array(n).fill(["0.0%", "0.0%"]);
      ws.getRange("A:A").format.columnWidth = 40;
      ws.getRange("B:B").format.columnWidth = 260;
      ws.getRange("C:H").format.columnWidth = 70;
      ws.freezePanes.freezeRows(4);

      // グラフ用のデータ（右側の列。グラフの元データなので非表示にしない）
      const tn = Math.min(20, n), pn = Math.min(100, n);
      ws.getRange("J4:L4").values = [["上位20", "連携なし", "連携"]];
      ws.getRange("J5:L" + (tn + 4)).values = agg.units.slice(0, tn).map((u) => [u.name, u.n - u.linked, u.linked]);
      ws.getRange("N4:O4").values = [["順位", "累計割合"]];
      ws.getRange("N5:O" + (pn + 4)).values = agg.units.slice(0, pn).map((u, i) => [i + 1, Math.round(u.cum * 10) / 10]);
      ws.getRange("J4:O4").format.font.color = "#7A8492";

      const bar = ws.charts.add(Excel.ChartType.barStacked, ws.getRange("J4:L" + (tn + 4)), Excel.ChartSeriesBy.columns);
      bar.title.text = "受注件数ランキング（上位" + tn + "）";
      bar.setPosition("Q4", "Z30");
      bar.axes.categoryAxis.reversePlotOrder = true;
      bar.legend.position = Excel.ChartLegendPosition.top;
      bar.series.getItemAt(0).format.fill.setSolidColor(COLOR.man);
      bar.series.getItemAt(1).format.fill.setSolidColor(COLOR.edi);

      const line = ws.charts.add(Excel.ChartType.line, ws.getRange("O4:O" + (pn + 4)), Excel.ChartSeriesBy.columns);
      line.title.text = "累計シェア（上位" + pn + "まで・%）";
      line.setPosition("Q32", "Z52");
      line.legend.visible = false;
      line.series.getItemAt(0).setXAxisValues(ws.getRange("N5:N" + (pn + 4)));
      line.series.getItemAt(0).format.line.color = COLOR.man;
      line.axes.valueAxis.maximum = 100;
      line.axes.valueAxis.minimum = 0;

      ws.activate();
      await ctx.sync();
    });
    say("「" + name + "」シートに出力しました。");
  }

  /* ---------- 操作 ---------- */

  async function recalc() {
    await readWorkbook();
    const added = await syncGroupSheet();
    st.pending = 0;
    fillPeriods();
    render();
    if (added) say("新しい請求先" + added + "件をグループシートの末尾に追加しました。");
  }

  async function afterGroupChange() {
    await readWorkbook();
    await syncGroupSheet();
    await watchGroupSheet();
    st.view = "group";
    st.pending = 0;
    render();
  }

  function fillPeriods() {
    const sel = $("period");
    const names = st.sources.map((s) => s.name);
    if (!names.includes(st.period)) st.period = (C.primarySource(st.sources) || {}).name || null;
    sel.innerHTML = names.map((n) => '<option value="' + esc(n) + '"' + (n === st.period ? " selected" : "") + ">" + esc(n) + "</option>").join("");
  }

  function bind() {
    $("period").onchange = (e) => { st.period = e.target.value; render(); };
    $("vGroup").onclick = () => { st.view = "group"; render(); };
    $("vCust").onclick = () => { st.view = "customer"; render(); };
    $("rk").oninput = () => render();
    $("topn").onchange = () => render();
    $("btnCreate").onclick = () => run("グループシートを作成", createGroupSheet);
    $("btnRecalc").onclick = () => run("再集計", async () => { await recalc(); if (!$("msg").textContent) say("再集計しました。"); });
    $("btnRecalc2").onclick = $("btnRecalc").onclick;
    $("btnOpenSheet").onclick = () => run("シートを開く", openGroupSheet);
    $("btnConfirmAll").onclick = () => run("確認済みに", () => confirmAll(false));
    $("btnConfirmSel").onclick = () => run("確認済みに", () => confirmAll(true));
    $("btnOutput").onclick = () => run("集計シートに出力", output);
  }

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) {
      $("loading").textContent = "このアドインはExcelで使います。";
      return;
    }
    // リボンの「グループシート」ボタン（共有ランタイムなので作業ウィンドウと同じページで受ける）
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("openGroupSheet", (event) => {
        openGroupSheet().catch((e) => console.error(e)).then(() => { if (event && event.completed) event.completed(); });
      });
    }
    bind();
    await run("シートの読み込み", async () => {
      await readWorkbook();
      if (!st.sources.length) {
        $("loading").textContent = "「請求先CD」「請求先」「受注件数」の見出しがあるシートが見つかりませんでした。";
        return;
      }
      st.view = st.hasGroupSheet ? "group" : "customer";
      if (st.hasGroupSheet) { await syncGroupSheet(); await watchGroupSheet(); }
      fillPeriods();
      $("loading").hidden = true;
      $("main").hidden = false;
      render();
    });
  });

  // テスト用
  if (typeof window !== "undefined") window.__taskpane = { cellsInColumnC: cellsInColumnC };
})();
