/* LedgerPilot 仕訳取込（単体ツール / SheetJS）
 * 台帳xlsxを読み込み → CSVとの差分確認 → 選択月を差し替えた台帳xlsxをダウンロード
 */
(function () {
  "use strict";
  var APP_VERSION = "rev_20260930_lp001";
  var L = window.LedgerCore, DV = window.LedgerDiffView;
  var $ = function (id) { return document.getElementById(id); };

  var state = { wb: null, bookName: "仕訳台帳.xlsx", existing: [], master: {}, history: [],
    parsed: null, csvName: "", diff: null, view: null };

  $("version").textContent = APP_VERSION + " / core " + L.CORE_VERSION;

  function msg(el, cls, text) { $(el).innerHTML = text ? '<div class="lp-note ' + cls + '">' + DV.esc(text) + "</div>" : ""; }
  function setupDrop(dropId, inputId, handler) {
    var d = $(dropId), inp = $(inputId);
    inp.addEventListener("change", function () { if (inp.files[0]) handler(inp.files[0]); });
    d.addEventListener("dragover", function (e) { e.preventDefault(); d.classList.add("over"); });
    d.addEventListener("dragleave", function () { d.classList.remove("over"); });
    d.addEventListener("drop", function (e) {
      e.preventDefault(); d.classList.remove("over");
      if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]);
    });
  }

  // ===== 台帳の読み込み =====
  function readBook(file) {
    file.arrayBuffer().then(function (buf) {
      var wb = XLSX.read(buf, { type: "array" });
      var existing = [], master = {}, history = [];
      wb.SheetNames.forEach(function (name) {
        var ws = wb.Sheets[name];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
        if (L.ymFromSheet(name)) {
          var width = Math.max(0, (aoa[0] || []).length - L.META_N);
          aoa.slice(1).forEach(function (r, i) {
            if (!r[0]) return;
            existing.push(L.entryFromSheetRow(r, i + 2, name, width));
          });
        } else if (name === L.MASTER_SHEET) {
          aoa.slice(1).forEach(function (r) { if (r[0] !== "") master[String(r[0])] = { name: r[1], cat: r[2] || L.defaultCategory(r[0]) }; });
        } else if (name === L.HISTORY_SHEET) {
          history = aoa.slice(1).filter(function (r) { return r[0] !== ""; });
        }
      });
      state.wb = wb; state.bookName = file.name; state.existing = existing; state.master = master; state.history = history;
      var months = {};
      existing.forEach(function (e) { months[L.ymOf(e.date)] = 1; });
      $("book-name").textContent = file.name + "（" + Object.keys(months).length + "ヶ月・" + existing.length + "行）";
      msg("load-msg", "", "");
      state.diff = null; $("diff").innerHTML = ""; $("btn-apply").disabled = true;
    }).catch(function (e) { msg("load-msg", "err", "台帳を読み込めませんでした: " + e.message); });
  }

  // ===== CSVの読み込み =====
  function readCsv(file) {
    file.arrayBuffer().then(function (buf) {
      var parsed = L.buildEntries(L.decodeCsvBuffer(buf));
      state.parsed = parsed; state.csvName = file.name;
      $("csv-name").textContent = file.name + "（" + parsed.voucherCount + "伝票・" + parsed.entries.length + "行）";
      $("p-start").value = L.monthStart(parsed.minDate).replace(/\//g, "-");
      $("p-end").value = L.monthEnd(parsed.maxDate).replace(/\//g, "-");
      $("csv-info").textContent = "CSVの日付範囲: " + parsed.minDate + " 〜 " + parsed.maxDate +
        "（月初〜月末に広げています。月の途中までのCSVなら終了日を合わせてください）";
      $("step2").classList.remove("off");
      msg("load-msg", "", "");
      runDiff();
    }).catch(function (e) { msg("load-msg", "err", e.message); });
  }

  function period() {
    return { start: $("p-start").value.replace(/-/g, "/"), end: $("p-end").value.replace(/-/g, "/") };
  }

  function runDiff() {
    if (!state.parsed) return;
    var p = period();
    if (!p.start || !p.end || p.start > p.end) { msg("load-msg", "err", "期間の開始日と終了日を確認してください。"); return; }
    state.diff = L.computeDiff(state.existing, state.parsed, p);
    $("step3").classList.remove("off");
    state.view = DV.render($("diff"), state.diff, {
      onChange: function (sel) {
        $("btn-apply").disabled = !sel.length;
        $("apply-msg").textContent = sel.length ? sel.join("、") + " を差し替えます" : "";
      }
    });
  }

  // ===== 書き出し =====
  function monthSheet(rows, header, width) {
    var importedAt = nowStr();
    var aoa = [L.META_COLS.concat(header)];
    rows.forEach(function (e) { aoa.push(L.entryToSheetRow(e, importedAt)); });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var fmts = L.columnFormats(width);
    for (var r = 1; r < aoa.length; r++) {
      for (var c = 0; c < aoa[r].length; c++) {
        var ref = XLSX.utils.encode_cell({ r: r, c: c });
        var cell = ws[ref]; if (!cell) continue;
        if (fmts[c] === "@" ) { cell.t = "s"; cell.v = String(cell.v); }
        else if (cell.t === "n") cell.z = fmts[c];
      }
    }
    ws["!cols"] = aoa[0].map(function (h, i) {
      var w = i < L.META_N ? 10 : 10;
      if (h === "摘要") w = 36; else if (/勘定科目名|補助科目名/.test(h)) w = 16; else if (/金額/.test(h)) w = 12;
      else if (/登録日付/.test(h)) w = 18; else if (h === "日付") w = 11;
      return { wch: w };
    });
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(aoa.length - 1, 1), c: aoa[0].length - 1 } }) };
    return ws;
  }
  function nowStr() {
    var d = new Date(), p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "/" + p(d.getMonth() + 1) + "/" + p(d.getDate()) + " " + p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
  }

  function apply() {
    var sel = state.view.getSelected();
    if (!sel.length) return;
    var plan = L.planApply(state.existing, state.diff, sel);
    if (plan.duplicates.length) {
      msg("load-msg", "err", "同じ仕訳が複数の月に入るため中止しました（" + plan.duplicates.length + "件）。連動している月をすべて選んでください。");
      return;
    }
    var wb = state.wb || XLSX.utils.book_new();
    var header = state.diff.header, width = state.diff.width;
    var at = nowStr();

    Object.keys(plan.sheets).sort().forEach(function (ym) {
      var name = L.sheetNameOf(ym), rows = plan.sheets[ym];
      var idx = wb.SheetNames.indexOf(name);
      if (!rows.length) {
        if (idx >= 0) { wb.SheetNames.splice(idx, 1); delete wb.Sheets[name]; }
        return;
      }
      var ws = monthSheet(rows, header, width);
      if (idx >= 0) wb.Sheets[name] = ws;
      else XLSX.utils.book_append_sheet(wb, ws, name);
    });
    // 仕訳シートを年月順に並べ替え（その他のシートは先頭側に維持）
    var others = wb.SheetNames.filter(function (n) { return !L.ymFromSheet(n) && n !== L.MASTER_SHEET && n !== L.HISTORY_SHEET; });
    var months = wb.SheetNames.filter(function (n) { return L.ymFromSheet(n); }).sort();

    // 科目マスタ（既存の区分は保持、新しい科目だけ追加）
    var acc = L.collectAccounts(state.parsed.entries);
    Object.keys(acc).forEach(function (code) {
      if (!state.master[code]) state.master[code] = { name: acc[code], cat: L.defaultCategory(code) };
    });
    var mAoa = [["科目コード", "勘定科目名", "区分"]];
    Object.keys(state.master).sort().forEach(function (code) {
      mAoa.push([String(code), state.master[code].name, state.master[code].cat]);
    });
    var mws = XLSX.utils.aoa_to_sheet(mAoa);
    mws["!cols"] = [{ wch: 10 }, { wch: 20 }, { wch: 12 }];
    wb.Sheets[L.MASTER_SHEET] = mws;

    // 取込履歴
    var p = state.diff.period;
    sel.forEach(function (ym) {
      var m = state.diff.months[ym] || { add: 0, chg: 0, del: 0, same: 0, incomingRows: 0 };
      state.history.push([at, state.csvName, p.start + "〜" + p.end, ym, m.add, m.chg, m.del, m.same, (plan.sheets[ym] || []).length]);
    });
    var hAoa = [["取込日時", "CSVファイル", "指定期間", "年月", "追加", "変更", "削除", "同一", "反映後の行数"]].concat(state.history);
    var hws = XLSX.utils.aoa_to_sheet(hAoa);
    hws["!cols"] = [{ wch: 19 }, { wch: 28 }, { wch: 24 }, { wch: 9 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 12 }];
    wb.Sheets[L.HISTORY_SHEET] = hws;

    wb.SheetNames = others.concat([L.HISTORY_SHEET, L.MASTER_SHEET], months);
    XLSX.writeFile(wb, state.bookName);

    // 保存後の状態で差分を取り直す（同じCSVなら差分0になる）
    var existing = [];
    months.forEach(function (name) {
      XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" }).slice(1).forEach(function (r, i) {
        if (r[0]) existing.push(L.entryFromSheetRow(r, i + 2, name, width));
      });
    });
    state.wb = wb; state.existing = existing;
    msg("load-msg", "ok", sel.join("、") + " を差し替えて「" + state.bookName + "」を保存しました。");
    runDiff();
  }

  setupDrop("drop-book", "file-book", readBook);
  setupDrop("drop-csv", "file-csv", readCsv);
  $("btn-diff").addEventListener("click", runDiff);
  $("btn-apply").addEventListener("click", apply);
  window.__lp = state; // デバッグ用
})();
