/* LedgerPilot 共通コア（取込ツール／アドイン共用）
 * - 弥生会計 仕訳日記帳CSV（cp932, 57列）の解析
 * - 伝票キー・行キー・内容ハッシュ
 * - 期間指定の差分（追加/変更/削除/同一）と月別マージ
 * - 科目区分と損益集計
 */
(function (root) {
  "use strict";

  var CORE_VERSION = "rev_20260930_lp003";

  // ===== シート構成 =====
  var SHEET_PREFIX = "仕訳_";          // 仕訳_2025-10
  var HISTORY_SHEET = "取込履歴";
  var MASTER_SHEET = "科目マスタ";
  var META_COLS = ["伝票キー", "行キー", "内容ハッシュ", "取込日時"]; // A〜D
  var META_N = META_COLS.length;

  // CSV列インデックス（元CSV基準）
  var C = {
    KUGIRI: 0, GYO: 1, DATE: 2,
    D_CODE: 8, D_NAME: 9, D_SUBC: 10, D_SUB: 11, D_TAXKBN: 13, D_RATE: 16, D_PARTNER: 23, D_AMT: 24, D_TAX: 25,
    C_CODE: 26, C_NAME: 27, C_SUBC: 28, C_SUB: 29, C_TAXKBN: 31, C_RATE: 34, C_PARTNER: 41, C_AMT: 42, C_TAX: 43,
    MEMO: 44, NEW_AT: 48, NEW_BY: 49, UPD_AT: 51, LAST_AT: 54
  };
  var NUMERIC_COLS = [16, 24, 25, 34, 42, 43]; // 税率・金額・税額

  var CATEGORIES = ["売上高", "売上原価", "販管費", "営業外収益", "営業外費用",
    "特別利益", "特別損失", "法人税等", "資産", "負債", "純資産"];
  var CREDIT_NATURE = { "売上高": 1, "営業外収益": 1, "特別利益": 1, "負債": 1, "純資産": 1 };

  // ===== 文字コード =====
  function decodeCsvBuffer(buf) {
    var bytes = new Uint8Array(buf);
    // UTF-8 BOM
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      return new TextDecoder("utf-8").decode(bytes.subarray(3));
    }
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder("shift_jis").decode(bytes);
    }
  }

  // ===== CSVパーサ（RFC4180, 改行入り引用符対応） =====
  function parseCsv(text) {
    var rows = [], row = [], f = "", q = false, i = 0, n = text.length, ch;
    while (i < n) {
      ch = text[i];
      if (q) {
        if (ch === '"') {
          if (text[i + 1] === '"') { f += '"'; i += 2; continue; }
          q = false; i++; continue;
        }
        f += ch; i++; continue;
      }
      if (ch === '"') { q = true; i++; continue; }
      if (ch === ",") { row.push(f); f = ""; i++; continue; }
      if (ch === "\r") { i++; continue; }
      if (ch === "\n") { row.push(f); rows.push(row); row = []; f = ""; i++; continue; }
      f += ch; i++;
    }
    if (f !== "" || row.length) { row.push(f); rows.push(row); }
    return rows.filter(function (r) { return r.length > 1 || (r[0] || "") !== ""; });
  }

  // ===== 正規化 =====
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  function serialToDateStr(v) {
    var d = new Date(Math.round((v - 25569) * 86400000));
    return d.getUTCFullYear() + "/" + pad2(d.getUTCMonth() + 1) + "/" + pad2(d.getUTCDate());
  }
  function serialToDateTimeStr(v) {
    var d = new Date(Math.round((v - 25569) * 86400000));
    return serialToDateStr(v) + " " + pad2(d.getUTCHours()) + ":" + pad2(d.getUTCMinutes()) + ":" + pad2(d.getUTCSeconds());
  }
  function normDate(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") return serialToDateStr(v);
    var m = String(v).trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    return m ? m[1] + "/" + pad2(+m[2]) + "/" + pad2(+m[3]) : String(v).trim();
  }
  function dateStrToSerial(s) {
    var m = String(s).match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (!m) return s;
    return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
  }
  function normField(i, v) {
    if (v === null || v === undefined) return "";
    if (i === C.DATE) return normDate(v);
    if ((i === C.NEW_AT || i === C.UPD_AT || i === C.LAST_AT) && typeof v === "number") return serialToDateTimeStr(v);
    if (NUMERIC_COLS.indexOf(i) >= 0) {
      if (v === "") return "";
      var n = Number(String(v).replace(/,/g, ""));
      return isNaN(n) ? String(v).trim() : String(n);
    }
    return String(v).trim();
  }
  function normRow(raw, width) {
    var out = [];
    for (var i = 0; i < width; i++) out.push(normField(i, raw[i]));
    return out;
  }

  // ===== ハッシュ（FNV-1a 32bit） =====
  function fnv1a(str) {
    var h = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("0000000" + h.toString(16)).slice(-8);
  }
  function rowHash(fields) { return fnv1a(fields.join("\u0001")); }

  // ===== CSV → 仕訳行 =====
  // 伝票キー = 新規登録日付|新規登録者（伝票No.が空のため）。同一秒の衝突時のみ |日付|#n を付与。
  function buildEntries(csvText) {
    var rows = parseCsv(csvText);
    if (!rows.length) throw new Error("CSVが空です。");
    var header = rows[0].map(function (s) { return s.trim(); });
    if (header[C.DATE] !== "日付" || header[C.D_NAME] !== "借方勘定科目名" || header[C.NEW_AT] !== "新規登録日付") {
      throw new Error("弥生会計の仕訳日記帳CSV（57列）ではありません。1行目の見出しを確認してください。");
    }
    var width = header.length;
    var vouchers = [], cur = null;
    rows.slice(1).forEach(function (raw, idx) {
      var f = normRow(raw, width);
      if (f[C.KUGIRI] === "*" || !cur) {
        cur = { base: f[C.NEW_AT] + "|" + f[C.NEW_BY], date: f[C.DATE], rows: [] };
        vouchers.push(cur);
      }
      cur.rows.push({ fields: f, csvLine: idx + 2 });
    });
    var baseCount = {};
    vouchers.forEach(function (v) { baseCount[v.base] = (baseCount[v.base] || 0) + 1; });
    var seq = {};
    var entries = [];
    vouchers.forEach(function (v) {
      var key = v.base;
      if (baseCount[v.base] > 1) {
        var s = v.base + "|" + v.date;
        seq[s] = (seq[s] || 0) + 1;
        key = s + "|#" + seq[s];
      }
      v.rows.forEach(function (r) {
        entries.push({
          vkey: key,
          rkey: key + "|" + r.fields[C.GYO],
          hash: rowHash(r.fields),
          fields: r.fields,
          date: r.fields[C.DATE],
          csvLine: r.csvLine
        });
      });
    });
    var dates = entries.map(function (e) { return e.date; }).filter(Boolean).sort();
    return { header: header, width: width, entries: entries, voucherCount: vouchers.length,
      minDate: dates[0] || "", maxDate: dates[dates.length - 1] || "" };
  }

  // ===== 期間 =====
  function ymOf(dateStr) { return dateStr.slice(0, 4) + "-" + dateStr.slice(5, 7); }
  function sheetNameOf(ym) { return SHEET_PREFIX + ym; }
  function ymFromSheet(name) { var m = name.match(/^仕訳_(\d{4}-\d{2})$/); return m ? m[1] : null; }
  function monthStart(d) { return d.slice(0, 8) + "01"; }
  function monthEnd(d) {
    var y = +d.slice(0, 4), m = +d.slice(5, 7);
    var last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return d.slice(0, 8) + pad2(last);
  }
  function monthsBetween(start, end) {
    var out = [], y = +start.slice(0, 4), m = +start.slice(5, 7);
    var ey = +end.slice(0, 4), em = +end.slice(5, 7);
    while (y < ey || (y === ey && m <= em)) {
      out.push(y + "-" + pad2(m));
      m++; if (m > 12) { m = 1; y++; }
    }
    return out;
  }
  function inPeriod(d, p) { return d >= p.start && d <= p.end; }

  // ===== シート行 ⇔ 仕訳行 =====
  function entryFromSheetRow(values, rowIndex, sheetName, width) {
    var f = normRow(values.slice(META_N), width);
    return { vkey: String(values[0]), rkey: String(values[1]), hash: String(values[2]),
      importedAt: values[3], fields: f, date: f[C.DATE], sheet: sheetName, rowIndex: rowIndex };
  }
  // 書き込み用: 日付はシリアル値、金額は数値、それ以外は文字列
  function entryToSheetRow(e, importedAt) {
    var row = [e.vkey, e.rkey, e.hash, e.importedAt || importedAt];
    e.fields.forEach(function (v, i) {
      if (i === C.DATE && v) row.push(dateStrToSerial(v));
      else if (NUMERIC_COLS.indexOf(i) >= 0 && v !== "" && !isNaN(Number(v))) row.push(Number(v));
      else row.push(v);
    });
    return row;
  }
  function columnFormats(width) {
    var f = META_COLS.map(function () { return "@"; });
    for (var i = 0; i < width; i++) {
      if (i === C.DATE) f.push("yyyy/mm/dd");
      else if (i === C.D_AMT || i === C.D_TAX || i === C.C_AMT || i === C.C_TAX) f.push("#,##0");
      else if (NUMERIC_COLS.indexOf(i) >= 0) f.push("General");
      else f.push("@");
    }
    return f;
  }

  // ===== 差分 =====
  function groupByVoucher(entries) {
    var map = {};
    entries.forEach(function (e) { (map[e.vkey] = map[e.vkey] || []).push(e); });
    return map;
  }
  function voucherSummary(rows) {
    var f0 = rows[0].fields, dr = 0, cr = 0, memos = [];
    rows.forEach(function (r) {
      dr += Number(r.fields[C.D_AMT] || 0);
      cr += Number(r.fields[C.C_AMT] || 0);
      if (r.fields[C.MEMO] && memos.indexOf(r.fields[C.MEMO]) < 0) memos.push(r.fields[C.MEMO]);
    });
    return { date: f0[C.DATE], debit: dr, credit: cr, memo: memos.join(" / "),
      accounts: rows.map(function (r) { return (r.fields[C.D_NAME] || "") + "／" + (r.fields[C.C_NAME] || ""); }) };
  }
  function fieldDiffs(header, oldRows, newRows) {
    var out = [], om = {}, nm = {};
    oldRows.forEach(function (r) { om[r.rkey] = r; });
    newRows.forEach(function (r) { nm[r.rkey] = r; });
    newRows.forEach(function (r) {
      var o = om[r.rkey];
      if (!o) { out.push({ gyo: r.fields[C.GYO], kind: "行追加" }); return; }
      if (o.hash === r.hash) return;
      r.fields.forEach(function (v, i) {
        if ((o.fields[i] || "") !== v) out.push({ gyo: r.fields[C.GYO], kind: "変更", col: header[i], before: o.fields[i], after: v });
      });
    });
    oldRows.forEach(function (r) { if (!nm[r.rkey]) out.push({ gyo: r.fields[C.GYO], kind: "行削除" }); });
    return out;
  }

  /**
   * existing: シートから読んだ全仕訳行（全月）
   * parsed:   buildEntries() の結果
   * period:   {start:"YYYY/MM/DD", end:"YYYY/MM/DD"} — この期間内はCSVで置き換える
   */
  function computeDiff(existing, parsed, period) {
    var incoming = parsed.entries.filter(function (e) { return inPeriod(e.date, period); });
    var outOfPeriodCsv = parsed.entries.length - incoming.length;
    var oldIn = existing.filter(function (e) { return inPeriod(e.date, period); });
    var oldG = groupByVoucher(oldIn), newG = groupByVoucher(incoming);
    var allOldG = groupByVoucher(existing);
    var vouchers = [];
    Object.keys(newG).forEach(function (k) {
      var n = newG[k], o = oldG[k] || allOldG[k];
      var s = voucherSummary(n);
      if (!o) { vouchers.push({ vkey: k, status: "追加", ym: ymOf(s.date), summary: s, rows: n }); return; }
      var same = o.length === n.length && n.every(function (r) {
        return o.some(function (x) { return x.rkey === r.rkey && x.hash === r.hash; });
      });
      var oldYm = ymOf(o[0].date);
      vouchers.push({ vkey: k, status: same ? "同一" : "変更", ym: ymOf(s.date), oldYm: oldYm,
        summary: s, oldSummary: voucherSummary(o), rows: n,
        diffs: same ? [] : fieldDiffs(parsed.header, o, n) });
    });
    Object.keys(oldG).forEach(function (k) {
      if (newG[k]) return;
      var o = oldG[k];
      vouchers.push({ vkey: k, status: "削除", ym: ymOf(o[0].date), summary: voucherSummary(o), rows: o });
    });
    vouchers.sort(function (a, b) { return a.summary.date < b.summary.date ? -1 : a.summary.date > b.summary.date ? 1 : 0; });

    var months = {};
    monthsBetween(period.start, period.end).forEach(function (ym) {
      months[ym] = { ym: ym, existingRows: 0, incomingRows: 0, add: 0, chg: 0, del: 0, same: 0,
        oldDebit: 0, newDebit: 0, moveOut: 0, linked: [] };
    });
    oldIn.forEach(function (e) { var m = months[ymOf(e.date)]; if (m) { m.existingRows++; m.oldDebit += Number(e.fields[C.D_AMT] || 0); } });
    incoming.forEach(function (e) { var m = months[ymOf(e.date)]; if (m) { m.incomingRows++; m.newDebit += Number(e.fields[C.D_AMT] || 0); } });
    vouchers.forEach(function (v) {
      var m = months[v.ym]; if (!m) return;
      if (v.status === "追加") m.add++; else if (v.status === "変更") m.chg++;
      else if (v.status === "削除") m.del++; else m.same++;
      if (v.oldYm && v.oldYm !== v.ym && months[v.oldYm]) months[v.oldYm].moveOut++;
      if (v.oldYm && v.oldYm !== v.ym && m.linked.indexOf(v.oldYm) < 0) {
        m.linked.push(v.oldYm);
        if (months[v.oldYm] && months[v.oldYm].linked.indexOf(v.ym) < 0) months[v.oldYm].linked.push(v.ym);
      }
    });
    Object.keys(months).forEach(function (k) {
      var m = months[k]; m.changed = m.add + m.chg + m.del + m.moveOut > 0;
    });
    return { period: period, months: months, vouchers: vouchers, incoming: incoming,
      outOfPeriodCsv: outOfPeriodCsv, header: parsed.header, width: parsed.width };
  }

  /**
   * 選択月の最終状態を計算する。
   * 選択月 M: (既存M のうち期間外の行) + (CSVのうち M かつ期間内の行)
   * 期間内でも未選択の月は既存のまま。全シート横断で伝票キー重複があれば中止（2重計上防止）。
   */
  function planApply(existing, diff, selectedYms) {
    var sel = {};
    selectedYms.forEach(function (ym) { sel[ym] = 1; });
    var p = diff.period;
    // 選択月に入るCSV伝票のキー（他の月・期間外にある旧版は取り除く＝月移動対応）
    var takeKeys = {};
    diff.incoming.forEach(function (e) { if (sel[ymOf(e.date)]) takeKeys[e.vkey] = 1; });
    var byYm = {}, touched = {};
    function put(e) { var ym = ymOf(e.date); (byYm[ym] = byYm[ym] || []).push(e); }
    existing.forEach(function (e) {
      var ym = ymOf(e.date);
      if (sel[ym] && inPeriod(e.date, p)) { touched[ym] = 1; return; } // 置き換え対象
      if (takeKeys[e.vkey]) { touched[ym] = 1; return; }                // 月移動した旧版
      put(e);
    });
    diff.incoming.forEach(function (e) { if (sel[ymOf(e.date)]) put(e); });

    // 2重計上チェック（全月横断で行キー重複を検出）
    var seen = {}, dups = [];
    Object.keys(byYm).forEach(function (ym) {
      byYm[ym].forEach(function (e) {
        if (seen[e.rkey]) dups.push({ rkey: e.rkey, a: seen[e.rkey], b: ym });
        seen[e.rkey] = ym;
      });
    });
    var sheetsToWrite = {};
    Object.keys(sel).concat(Object.keys(touched)).forEach(function (ym) {
      sheetsToWrite[ym] = (byYm[ym] || []).slice().sort(function (a, b) {
        if (a.date !== b.date) return a.date < b.date ? -1 : 1;
        if (a.vkey !== b.vkey) return a.vkey < b.vkey ? -1 : 1;
        return Number(a.fields[C.GYO]) - Number(b.fields[C.GYO]);
      });
    });
    return { sheets: sheetsToWrite, duplicates: dups };
  }

  // ===== 科目区分 =====
  function defaultCategory(code) {
    var n = parseInt(code, 10);
    if (isNaN(n)) return "資産";
    if (n < 300) return "資産";
    if (n < 400) return "負債";
    if (n < 500) return "純資産";
    if (n < 600) return "売上高";
    if (n < 700) return "売上原価";
    if (n < 800) return "販管費";
    if (n < 850) return "営業外収益";
    if (n < 900) return "営業外費用";
    if (n < 950) return "特別利益";
    if (n < 990) return "特別損失";
    return "法人税等";
  }
  function collectAccounts(entries) {
    var map = {};
    entries.forEach(function (e) {
      var f = e.fields;
      if (f[C.D_CODE]) map[f[C.D_CODE]] = f[C.D_NAME];
      if (f[C.C_CODE]) map[f[C.C_CODE]] = f[C.C_NAME];
    });
    return map;
  }

  // ===== 集計 =====
  // 仕訳行 → 借方・貸方の明細（posting）へ展開
  function toPostings(entries, master, mode) {
    var out = [];
    var byV = groupByVoucher(entries);
    entries.forEach(function (e) {
      var f = e.fields, ym = ymOf(e.date);
      var vrows = byV[e.vkey];
      function counter(side) {
        // 同じ行の反対側。無ければ伝票内の反対側科目（1種類なら科目名、複数なら諸口）
        var code = side === "D" ? f[C.C_CODE] : f[C.D_CODE];
        var name = side === "D" ? f[C.C_NAME] : f[C.D_NAME];
        if (code) return name;
        var names = {};
        vrows.forEach(function (r) {
          var nm = side === "D" ? r.fields[C.C_NAME] : r.fields[C.D_NAME];
          if (nm) names[nm] = 1;
        });
        var ks = Object.keys(names);
        return ks.length === 1 ? ks[0] : "諸口";
      }
      if (f[C.D_CODE]) out.push(withTax({ side: "D", code: f[C.D_CODE], name: f[C.D_NAME], sub: f[C.D_SUB],
        counter: counter("D"), memo: f[C.MEMO], date: e.date, ym: ym,
        cat: (master[f[C.D_CODE]] || {}).cat || defaultCategory(f[C.D_CODE]), vkey: e.vkey,
        sheet: e.sheet, rowIndex: e.rowIndex }, f[C.D_AMT], f[C.D_TAX], f[C.D_RATE], f[C.D_TAXKBN]));
      if (f[C.C_CODE]) out.push(withTax({ side: "C", code: f[C.C_CODE], name: f[C.C_NAME], sub: f[C.C_SUB],
        counter: counter("C"), memo: f[C.MEMO], date: e.date, ym: ym,
        cat: (master[f[C.C_CODE]] || {}).cat || defaultCategory(f[C.C_CODE]), vkey: e.vkey,
        sheet: e.sheet, rowIndex: e.rowIndex }, f[C.C_AMT], f[C.C_TAX], f[C.C_RATE], f[C.C_TAXKBN]));
    });
    setTaxMode(out, mode || "in");
    return out;
  }

  /* 税込／税抜の金額を持たせる（その明細自身の税区分・税率で判定）
   * - 税率なし（対象外・非課税）: 税込＝税抜＝本体金額
   * - 消費税の科目（仮払消費税・未払消費税など）: 換算しない
   * - 消費税額あり（税抜経理・別記）: 税抜＝本体、税込＝本体＋税額
   * - 消費税額0（税込経理・内税）: 税込＝本体、税額＝本体×税率/(100+税率) 切り捨て、税抜＝本体−税額 */
  function withTax(p, amt, taxAmt, rate, kbn) {
    var a = Number(amt || 0), t = Number(taxAmt || 0), r = Number(rate || 0);
    p.rate = r; p.taxKbn = kbn || "";
    if (!r || /消費税/.test(p.name)) { p.amtIn = a; p.amtEx = a; p.tax = 0; }
    else if (t) { p.amtEx = a; p.amtIn = a + t; p.tax = t; }
    else {
      var tax = (a < 0 ? -1 : 1) * Math.floor(Math.abs(a) * r / (100 + r));
      p.amtIn = a; p.amtEx = a - tax; p.tax = tax;
    }
    return p;
  }
  // mode: "in"=税込 / "ex"=税抜。p.amt を切り替える（集計はすべて p.amt を使う）
  function setTaxMode(postings, mode) {
    postings.forEach(function (p) { p.amt = mode === "ex" ? p.amtEx : p.amtIn; });
    return postings;
  }
  // 区分の性質に合わせた符号（収益・負債は貸方プラス）
  function signed(p) {
    var credit = CREDIT_NATURE[p.cat];
    return (p.side === "D" ? 1 : -1) * (credit ? -1 : 1) * p.amt;
  }
  function plOf(postings) {
    var t = { "売上高": 0, "売上原価": 0, "販管費": 0, "営業外収益": 0, "営業外費用": 0,
      "特別利益": 0, "特別損失": 0, "法人税等": 0 };
    postings.forEach(function (p) { if (p.cat in t) t[p.cat] += signed(p); });
    t["売上総利益"] = t["売上高"] - t["売上原価"];
    t["営業利益"] = t["売上総利益"] - t["販管費"];
    t["経常利益"] = t["営業利益"] + t["営業外収益"] - t["営業外費用"];
    t["税引前利益"] = t["経常利益"] + t["特別利益"] - t["特別損失"];
    t["当期純利益"] = t["税引前利益"] - t["法人税等"];
    return t;
  }

  // 摘要から相手先を推定（例: "108027　キャスト(10月売上高)" → "キャスト"）
  function memoKey(memo) {
    var s = String(memo || "").normalize("NFKC");
    s = s.replace(/[\(（][^\)）]*[\)）]?/g, " ")
      .replace(/^\s*\d{3,}\s*/, "")
      .replace(/\d{1,2}月分?/g, " ")
      .replace(/(売上計上|売上高|支払|振込|入金)/g, " ")
      .trim();
    var t = s.split(/[\s\u3000\/／]+/).filter(Boolean)[0];
    return t || "（摘要なし）";
  }

  var api = {
    CORE_VERSION: CORE_VERSION, C: C, SHEET_PREFIX: SHEET_PREFIX, HISTORY_SHEET: HISTORY_SHEET,
    MASTER_SHEET: MASTER_SHEET, META_COLS: META_COLS, META_N: META_N, CATEGORIES: CATEGORIES,
    CREDIT_NATURE: CREDIT_NATURE,
    decodeCsvBuffer: decodeCsvBuffer, parseCsv: parseCsv, buildEntries: buildEntries,
    normDate: normDate, dateStrToSerial: dateStrToSerial, serialToDateStr: serialToDateStr,
    ymOf: ymOf, sheetNameOf: sheetNameOf, ymFromSheet: ymFromSheet,
    monthStart: monthStart, monthEnd: monthEnd, monthsBetween: monthsBetween, inPeriod: inPeriod,
    entryFromSheetRow: entryFromSheetRow, entryToSheetRow: entryToSheetRow, columnFormats: columnFormats,
    computeDiff: computeDiff, planApply: planApply,
    defaultCategory: defaultCategory, collectAccounts: collectAccounts,
    toPostings: toPostings, setTaxMode: setTaxMode, memoKey: memoKey, signed: signed, plOf: plOf, groupByVoucher: groupByVoucher
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.LedgerCore = api;
})(typeof window !== "undefined" ? window : this);
