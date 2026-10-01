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
    origin: "手作業",
    openKey: null,
    showAll: false,
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

  /* 入力起因の色：手作業は青、外部連携は緑系（明るさを変えて区別） */
  const EXT_COLORS = ["#0F6B5C", "#17917E", "#4DB6A4", "#86CFC2", "#B4E2D9"];
  function originColors(summary) {
    const m = new Map();
    let k = 0;
    summary.list.forEach((o) => m.set(o.origin, o.origin === C.MANUAL ? COLOR.man : EXT_COLORS[Math.min(k++, EXT_COLORS.length - 1)]));
    return m;
  }

  function currentView() {
    const recsAll = currentRecords();
    const summary = C.originSummary(recsAll, st.noteMap);
    if (!summary.list.some((o) => o.origin === st.origin)) {
      st.origin = summary.list.length ? summary.list[0].origin : C.MANUAL;
    }
    const recs = C.filterByOrigin(recsAll, st.noteMap, st.origin);
    const view = st.hasGroupSheet ? st.view : "customer";
    return {
      summary: summary, colors: originColors(summary), recs: recs, view: view,
      origin: summary.list.find((o) => o.origin === st.origin) || { origin: st.origin, label: C.originLabel(st.origin), short: C.originShort(st.origin), n: 0, customers: 0, share: 0 },
      agg: C.aggregate(recs, st.groupMap, st.noteMap, view),
      base: C.aggregate(recs, st.groupMap, st.noteMap, "customer")
    };
  }

  function render() {
    const v = currentView();
    const { agg, base, view } = v;
    const color = v.colors.get(st.origin) || COLOR.man;

    renderOrigins(v);

    // バナー
    const sugCount = st.suggestion ? new Set(st.suggestion.values()).size : 0;
    $("bnSuggest").hidden = st.hasGroupSheet || !sugCount;
    if (!st.hasGroupSheet && sugCount) {
      const preview = C.aggregate(v.recs, st.suggestion, st.noteMap, "group");
      const big = preview.units.filter((u) => u.isGroup).slice(0, 2).map((u) => u.name + "（" + u.members.length + "拠点）").join("、");
      $("bnSugTitle").textContent = "同じ会社の請求先が" + sugCount + "組あります";
      $("bnSugDesc").textContent = (big ? big + "など。" : "") + "まとめると上位10のシェアは " + pct(C.topShare(base, 10)) + " → " + pct(C.topShare(preview, 10)) + " になります（" + v.origin.short + "）。";
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
    $("k10").textContent = agg.total ? pct(t10) : "—";
    $("k20").textContent = agg.total ? pct(t20) : "—";
    const diff = (a, b) => view === "group" && Math.abs(a - b) >= 0.05 ? "請求先別より " + (a > b ? "+" : "") + (a - b).toFixed(1) + "pt" : "";
    $("k10d").textContent = diff(t10, C.topShare(base, 10));
    $("k20d").textContent = diff(t20, C.topShare(base, 20));

    $("paretoTitle").textContent = "累計シェア：" + v.origin.label + "（上位100まで）";
    renderPareto(agg, view === "group" ? base : null, color);
    renderRanking(agg, color);
    $("footer").hidden = !currentRecords().length;
  }

  /** ドーナツグラフと切り替えボタン */
  function renderOrigins(v) {
    const list = v.summary.list, T = v.summary.total;
    const cx = 80, cy = 80, R = 70, r = 44;
    const pt = (rad, ang) => (cx + rad * Math.cos(ang * Math.PI / 180)).toFixed(1) + " " + (cy + rad * Math.sin(ang * Math.PI / 180)).toFixed(1);
    let a = -90;
    const arcs = list.map((o) => {
      const s = T ? o.n / T * 360 : 0;
      const b = a + Math.min(s, 359.99);
      const la = s > 180 ? 1 : 0;
      const d = "M" + pt(R, a) + " A" + R + " " + R + " 0 " + la + " 1 " + pt(R, b) + " L" + pt(r, b) + " A" + r + " " + r + " 0 " + la + " 0 " + pt(r, a) + " Z";
      a += s;
      const on = o.origin === st.origin;
      return '<path data-o="' + esc(o.origin) + '" d="' + d + '" fill="' + (on ? v.colors.get(o.origin) : "#D5DAE0") + '" stroke="#fff" stroke-width="2"><title>' + esc(o.label + "：" + fmt(o.n) + "件（" + pct(o.share) + "）") + "</title></path>";
    }).join("");
    $("donut").innerHTML = '<svg viewBox="0 0 160 160" role="img" aria-label="入力起因の構成比">' + arcs +
      '<text x="80" y="74" font-size="12" fill="#5B6573" text-anchor="middle">' + esc(v.origin.short) + "</text>" +
      '<text x="80" y="96" font-size="20" font-weight="700" fill="#1B2430" text-anchor="middle">' + pct(v.origin.share) + "</text></svg>";
    $("orgList").innerHTML = list.map((o) =>
      '<button data-o="' + esc(o.origin) + '" aria-pressed="' + (o.origin === st.origin) + '"><i class="sw" style="width:9px;height:9px;background:' + v.colors.get(o.origin) + '"></i>' +
      '<span class="t">' + esc(o.label) + '</span><span class="c">' + fmt(o.n) + "</span></button>").join("");

    const ext = list.filter((o) => o.origin !== C.MANUAL).map((o) => o.short);
    let note;
    if (st.origin === C.MANUAL) {
      note = ext.length ? "外部連携（" + ext.join("・") + "）は累計シェアの対象外です。入力起因を押すと、その起因だけで集計します。"
        : "この期間に外部連携の受注はありません。";
    } else {
      const g = C.aggregate(v.recs, st.groupMap, st.noteMap, "group");
      const groups = g.units.filter((u) => u.isGroup);
      note = v.origin.short + "の" + fmt(v.origin.customers) + "請求先を集計しています。" +
        (st.hasGroupSheet && groups.length ? "グループ別では" + fmt(g.units.length) + "単位（" + groups.slice(0, 2).map((u) => u.name + " " + fmt(u.n) + "件").join("、") + (groups.length > 2 ? "など" : "") + "）です。" : "") +
        "手作業に戻すには「手作業（備考なし）」を押します。";
    }
    $("orgNote").textContent = note;
  }

  function curve(agg, N, W, H) {
    const pts = ["M0 " + H];
    for (let i = 0; i < N; i++) {
      const u = agg.units[i];
      const c = u ? u.cum : (agg.units.length ? 100 : 0);
      pts.push("L" + ((i + 1) / N * W).toFixed(1) + " " + (H - c / 100 * H).toFixed(1));
    }
    return pts.join(" ");
  }

  function renderPareto(agg, base, color) {
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
      '<path d="' + curve(agg, N, W, H) + '" fill="none" stroke="' + color + '" stroke-width="2.5" stroke-linejoin="round"/>' +
      '<line x1="' + x + '" y1="0" x2="' + x + '" y2="' + H + '" stroke="#A65F0A" stroke-dasharray="2 3"/>' +
      '<circle cx="' + x + '" cy="' + y + '" r="5" fill="' + color + '" stroke="#fff" stroke-width="2"/>' +
      '<text x="' + lx + '" y="' + Math.max(y - 6, 10) + '" font-size="11" fill="#1B2430" text-anchor="' + anchor + '">上位' + k + ' → ' + pct(v) + '</text>' +
      '<text x="0" y="166" font-size="9" fill="#7A8492">1</text><text x="' + W + '" y="166" font-size="9" fill="#7A8492" text-anchor="end">100</text>' +
      "</svg>";
    $("paretoLg").innerHTML = base ? '<span style="color:' + color + '">━ グループ別</span><span style="color:#7A8492">┅ 請求先別</span>' : "";
    $("rkOut").textContent = "上位" + k + " → " + pct(v);
  }

  const MEMBER_LIMIT = 8;

  function renderRanking(agg, color) {
    const N = Number($("topn").value);
    const top = agg.units.slice(0, N);
    const max = top.length ? top[0].n : 1;
    if (st.openKey && !top.some((u) => u.key === st.openKey && u.isGroup)) { st.openKey = null; st.showAll = false; }
    $("rank").innerHTML = top.map((u, i) => {
      const bar = '<span class="bar" style="width:' + (u.n / max * 100).toFixed(1) + '%"><span style="flex:1;background:' + color + '"></span></span>';
      const no = '<span class="no">' + (i + 1) + "</span>";
      const num = '<span class="n">' + fmt(u.n) + "</span>";
      if (!u.isGroup) {
        return '<div class="rk-i"><div class="rk" title="' + esc(u.name + "：" + fmt(u.n) + "件（" + pct(u.share) + "）") + '">' + no +
          '<span class="cv"></span><span class="nm">' + esc(u.name) + "</span>" + bar + num + "</div></div>";
      }
      const open = st.openKey === u.key;
      let body = "";
      if (open) {
        const top1 = u.members[0] ? u.members[0].n : 1;
        const list = st.showAll ? u.members : u.members.slice(0, MEMBER_LIMIT);
        body = '<div class="acc" id="acc-' + i + '">' +
          '<div class="acc-h"><span>' + u.members.length + "件の請求先</span><span>グループ内の割合</span></div>" +
          list.map((m) => '<div class="acc-r" title="' + esc(m.cd + "　" + m.name) + '">' +
            '<span class="acc-n">' + esc(C.shortMemberName(m.name, u.name)) + "</span>" +
            '<span class="acc-b"><span style="width:' + (m.n / top1 * 100).toFixed(1) + "%;background:" + color + '"></span></span>' +
            '<span class="acc-c">' + fmt(m.n) + '</span><span class="acc-p">' + pct(u.n ? m.n / u.n * 100 : 0) + "</span></div>").join("") +
          '<div class="acc-f">' +
          (u.members.length > MEMBER_LIMIT ? '<button class="lnk sm" data-more="1">' + (st.showAll ? "上位" + MEMBER_LIMIT + "件だけ表示" : "残り" + (u.members.length - MEMBER_LIMIT) + "件を表示") + "</button>" : "<span></span>") +
          '<button class="lnk sm" data-sheet="' + esc(u.name) + '">グループシートで表示</button></div></div>';
      }
      return '<div class="rk-i' + (open ? " open" : "") + '"><button class="rk rk-g" data-key="' + esc(u.key) + '" aria-expanded="' + open + '"' + (open ? ' aria-controls="acc-' + i + '"' : "") + ">" + no +
        '<span class="cv" aria-hidden="true">' + (open ? "▼" : "▶") + "</span>" +
        '<span class="nm">' + esc(u.name) + "<small>" + u.members.length + "</small></span>" + bar + num + "</button>" + body + "</div>";
    }).join("") || '<div class="msg">この入力起因の受注はありません。</div>';
  }

  /** グループシートで、そのグループの行だけを表示する */
  async function showGroupInSheet(groupName) {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(C.GROUP_SHEET);
      ws.activate();
      const rows = st.groupRows.filter((r) => r.group === groupName);
      const last = st.groupRows.length ? st.groupRows[st.groupRows.length - 1].row + 1 : 1;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
        ws.autoFilter.apply(ws.getRange("A1:D" + last), 2, { filterOn: Excel.FilterOn.values, values: [groupName] });
        say("グループシートを「" + groupName + "」で絞り込みました。解除はC列のフィルターから行います。");
      } else if (rows.length) {
        ws.getRange("A" + (rows[0].row + 1) + ":D" + (rows[0].row + 1)).select();
        say("グループシートの「" + groupName + "」の先頭行を選びました。");
      }
      await ctx.sync();
    });
  }

  /* ---------- 集計シートへの出力 ---------- */

  async function output() {
    const v = currentView();
    const { agg, view } = v;
    const label = view === "group" ? "グループ別" : "請求先別";
    const name = (C.OUTPUT_PREFIX + label + "_" + v.origin.short + "_" + st.period).replace(/[\\/?*[\]:]/g, "").slice(0, 31);

    await Excel.run(async (ctx) => {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name);
      await ctx.sync();
      if (!old.isNullObject) old.delete();
      const ws = ctx.workbook.worksheets.add(name);

      ws.getRange("A1").values = [["受注件数の集計（" + label + "）　入力起因：" + v.origin.label + "　期間：" + st.period]];
      ws.getRange("A1").format.font.bold = true;
      ws.getRange("A1").format.font.size = 14;
      ws.getRange("A2:F2").values = [["受注件数", agg.total, "上位10のシェア", C.topShare(agg, 10) / 100, "上位20のシェア", C.topShare(agg, 20) / 100]];
      ws.getRange("B2").numberFormat = [["#,##0"]];
      ws.getRange("D2").numberFormat = [["0.0%"]];
      ws.getRange("F2").numberFormat = [["0.0%"]];

      const head = [["順位", view === "group" ? "グループ／請求先" : "請求先", "請求先数", "受注件数", "割合", "累計割合"]];
      const rows = agg.units.map((u, i) => [i + 1, u.name, u.members.length, u.n, u.share / 100, u.cum / 100]);
      const n = rows.length;
      ws.getRange("A4:F4").values = head;
      const hr = ws.getRange("A4:F4");
      hr.format.font.bold = true;
      hr.format.fill.color = COLOR.headFill;
      if (n) {
        ws.getRange("A5:F" + (n + 4)).values = rows;
        ws.getRange("D5:D" + (n + 4)).numberFormat = Array(n).fill(["#,##0"]);
        ws.getRange("E5:F" + (n + 4)).numberFormat = Array(n).fill(["0.0%", "0.0%"]);
      }
      ws.getRange("A:A").format.columnWidth = 40;
      ws.getRange("B:B").format.columnWidth = 260;
      ws.getRange("C:F").format.columnWidth = 70;
      ws.freezePanes.freezeRows(4);

      // グラフ用のデータ（グラフの元データなので非表示にしない）
      const os = v.summary.list, on = os.length;
      ws.getRange("H4:I4").values = [["入力起因", "受注件数"]];
      ws.getRange("H5:I" + (on + 4)).values = os.map((o) => [o.label, o.n]);
      const tn = Math.min(20, n), pn = Math.min(100, n);
      ws.getRange("K4:L4").values = [["上位" + tn, "受注件数"]];
      ws.getRange("N4:O4").values = [["順位", "累計割合"]];
      ws.getRange("H4:O4").format.font.color = "#7A8492";

      const pie = ws.charts.add(Excel.ChartType.doughnut, ws.getRange("H4:I" + (on + 4)), Excel.ChartSeriesBy.columns);
      pie.title.text = "入力起因の構成比";
      pie.setPosition("Q4", "Z22");
      pie.legend.position = Excel.ChartLegendPosition.right;
      pie.series.getItemAt(0).hasDataLabels = true;

      if (n) {
        ws.getRange("K5:L" + (tn + 4)).values = agg.units.slice(0, tn).map((u) => [u.name, u.n]);
        ws.getRange("N5:O" + (pn + 4)).values = agg.units.slice(0, pn).map((u, i) => [i + 1, Math.round(u.cum * 10) / 10]);

        const color = v.colors.get(st.origin) || COLOR.man;
        const bar = ws.charts.add(Excel.ChartType.barClustered, ws.getRange("K4:L" + (tn + 4)), Excel.ChartSeriesBy.columns);
        bar.title.text = "受注件数ランキング（" + v.origin.short + "・上位" + tn + "）";
        bar.setPosition("Q24", "Z48");
        bar.axes.categoryAxis.reversePlotOrder = true;
        bar.legend.visible = false;
        bar.series.getItemAt(0).format.fill.setSolidColor(color);

        const line = ws.charts.add(Excel.ChartType.line, ws.getRange("O4:O" + (pn + 4)), Excel.ChartSeriesBy.columns);
        line.title.text = "累計シェア（" + v.origin.short + "・上位" + pn + "まで・%）";
        line.setPosition("Q50", "Z70");
        line.legend.visible = false;
        line.series.getItemAt(0).setXAxisValues(ws.getRange("N5:N" + (pn + 4)));
        line.series.getItemAt(0).format.line.color = color;
        line.axes.valueAxis.maximum = 100;
        line.axes.valueAxis.minimum = 0;
      }

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
    const pickOrigin = (e) => { const el = e.target.closest("[data-o]"); if (el) { st.origin = el.getAttribute("data-o"); render(); } };
    $("orgList").onclick = pickOrigin;
    $("donut").onclick = pickOrigin;
    $("rank").onclick = (e) => {
      const more = e.target.closest("[data-more]");
      if (more) { st.showAll = !st.showAll; render(); return; }
      const sh = e.target.closest("[data-sheet]");
      if (sh) { run("グループシートで表示", () => showGroupInSheet(sh.getAttribute("data-sheet"))); return; }
      const row = e.target.closest("[data-key]");
      if (row) {
        const k = row.getAttribute("data-key");
        st.openKey = st.openKey === k ? null : k;
        st.showAll = false;
        render();
      }
    };
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
