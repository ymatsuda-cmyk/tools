/* ============================================================
 * monthly-sheets.js — 月次PL／月次BSシートへの反映と検算
 * ------------------------------------------------------------
 * ・シートは名前の前方一致（「月次PL」「月次BS」）。なければ新規作成。
 * ・見出し行（2025/10 形式のセルが6つ以上ある行）から該当月の列を探す。
 * ・A列の科目名とPDFの科目を対応づけて、その月の列に書き込む。
 *     PL＝当月発生額（残高−繰越残高）、BS＝月末残高
 *   数式の入ったセル（合計・利益の行）には書き込まない。
 * ・対応は「科目対応」シートに保存。手で直すと次回から優先される。
 * ・どの行にも対応しないPDFの科目（金額あり）は行を追加する。
 * ・書き込み後、指定の合計行をPDFの集計値と照合する。
 * ============================================================ */
(function (global) {
  "use strict";
  const MAP_SHEET = "科目対応";
  const MAP_HEAD = ["シート", "行の科目名", "取込元（PDFの科目名。複数は＋でつなぐ）", "方法"];

  // 照合キー：（旧:〜）注記と空白を除き、残りの括弧書きも除く
  const base = (s) => String(s == null ? "" : s).replace(/[（(]旧[:：][^）)]*[）)]/g, "").replace(/[\s\u3000]/g, "");
  const mkey = (s) => base(s).replace(/[（(][^）)]*[）)]/g, "");
  // 売上高の内訳行（手入力。PDFからは書き込まず、科目としても扱わない）
  const SEG_NAMES = { m: ["保守"], e: ["開発(既存)", "既存開発"], n: ["開発(新規)", "新規開発"] };
  const segOf = (s) => { const t = base(s).replace(/（/g, "(").replace(/）/g, ")"); return Object.keys(SEG_NAMES).find(k => SEG_NAMES[k].includes(t)) || null; };

  const CHECKS = {
    PL: [
      { label: "売上総利益", row: ["売上総利益"], pdf: ["売上総利益", "売上総損失"] },
      { label: "販管費合計", row: ["販管費合計", "販売費及び一般管理費", "販売費及び一般管理費計"], pdf: ["販売費及び一般管理費", "販売費及び一般管理費計"] },
      { label: "営業利益", row: ["営業利益", "営業損失"], pdf: ["営業利益", "営業損失"] },
      { label: "経常利益", row: ["経常利益", "経常損失"], pdf: ["経常利益", "経常損失"] }
    ],
    BS: [
      { label: "現金預金 計", row: ["現金預金計"], pdf: ["現金預金計", "現金及び預金", "現金及び預金計"] },
      { label: "流動資産 計", row: ["流動資産計"], pdf: ["流動資産計", "流動資産合計"] },
      { label: "固定資産 計", row: ["固定資産計"], pdf: ["固定資産計", "固定資産合計"] },
      { label: "資産 合計", row: ["資産合計"], pdf: ["資産合計", "資産の部合計"] },
      { label: "負債 合計", row: ["負債合計"], pdf: ["負債合計", "負債の部合計"] },
      { label: "純資産 合計", row: ["純資産合計"], pdf: ["純資産合計", "純資産の部合計"] },
      { label: "負債・純資産 合計", row: ["負債・純資産合計", "負債純資産合計"], pdf: ["負債・純資産合計", "負債純資産合計", "負債及び純資産合計"] }
    ]
  };
  // 名前だけでは決まらない行の既定の取込元（照合キー → PDF科目名）
  const DEFAULTS = {
    "有形・無形固定資産": ["有形固定資産計", "無形固定資産計"],
    "保証金（投資その他）": ["投資その他の資産計"],
    "その他・固定負債": ["買掛金", "未払金", "固定負債計"],
    "販管費合計": ["販売費及び一般管理費"],
    "営業外収益（受取利息等）": ["営業外収益"]
  };

  const isMonthCell = (v) => /^\d{4}\/\d{1,2}$/.test(String(v == null ? "" : v).trim());
  function monthText(v) {
    if (typeof v === "number" && v > 20000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + "/" + (d.getUTCMonth() + 1);
    }
    return String(v == null ? "" : v).trim();
  }
  // 相対参照の行番号をずらす（$付きの行は固定、他シート参照は対象外）
  function shiftRows(f, d) {
    return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, ca, col, ra, row) => ca + col + ra + (ra ? row : String(Number(row) + d)));
  }
  // 数式の月列と行を別の位置へ写す（計画シート用）
  function moveFormula(f, dRow, fromCol, toCol) {
    return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, ca, col, ra, row) => {
      const c = !ca && col === fromCol ? toCol : col;
      return ca + c + ra + (ra ? row : String(Number(row) + dRow));
    });
  }
  const toMonthLabel = (period) => period.slice(0, 4) + "/" + Number(period.slice(5));
  const colLetter = (n) => { let s = ""; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  function findHeader(vals) {
    for (let r = 0; r < Math.min(vals.length, 30); r++) {
      const cols = [];
      (vals[r] || []).forEach((c, i) => { if (isMonthCell(monthText(c))) cols.push(i); });
      if (cols.length >= 6) return { row: r, cols, labels: cols.map(c => monthText(vals[r][c])) };
    }
    return null;
  }

  // PDF（TB_明細）の行から、名前→金額・名前→行 の表を作る
  function pdfIndex(rows, sheet) {
    const list = rows.filter(r => r.sheet === sheet).sort((a, b) => a.order - b.order);
    const val = (r) => (sheet === "PL" ? r.bal - r.open : r.bal);
    const byKey = {};
    list.forEach(r => { const k = mkey(r.name); if (!(k in byKey)) byKey[k] = r; });
    // 明細行の直近の親（次に現れる集計行）
    list.forEach((r, i) => { r._parent = null; if (r.kind !== "集計") { const p = list.slice(i + 1).find(x => x.kind === "集計"); r._parent = p || null; } });
    return { list, byKey, val };
  }

  function createLayout(sheet, rows, fyMonths) {
    const list = rows.filter(r => r.sheet === sheet).sort((a, b) => a.order - b.order);
    const head = ["科目"].concat(fyMonths.map(toMonthLabel));
    const body = list.map(r => [r.kind === "集計" ? r.name : r.name].concat(fyMonths.map(() => "")));
    return [head].concat(body);
  }

  /**
   * period の行（TB_明細の形）を月次PL・月次BSへ反映する
   * opt.fyMonths: その会計年度の12か月（新規作成時の列見出し）
   */
  async function reflect(period, rows, opt) {
    const result = { period, sheets: {}, errors: [], inserted: [], mapAdded: 0 };
    await Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets;
      wss.load("items/name");
      await ctx.sync();
      const names = wss.items.map(w => w.name);
      const pick = (frag) => names.find(n => n.indexOf(frag) === 0) || names.find(n => n.indexOf(frag) >= 0);
      const target = { PL: pick("月次PL"), BS: pick("月次BS") };

      // なければ作る
      for (const sh of ["PL", "BS"]) {
        if (target[sh]) continue;
        const nm = "月次" + sh;
        const ws = wss.add(nm);
        const lay = createLayout(sh, rows, opt.fyMonths);
        const rg = ws.getRangeByIndexes(0, 0, lay.length, lay[0].length);
        rg.numberFormat = lay.map((r, i) => r.map((_, c) => (i === 0 || c === 0 ? "@" : "#,##0;-#,##0")));
        rg.values = lay;
        ws.getRangeByIndexes(0, 0, 1, lay[0].length).format.font.bold = true;
        target[sh] = nm;
        result.sheets[sh] = { created: true };
      }
      // 科目対応シート
      let mapWs = wss.getItemOrNullObject(MAP_SHEET);
      await ctx.sync();
      if (mapWs.isNullObject) {
        mapWs = wss.add(MAP_SHEET);
        const h = mapWs.getRangeByIndexes(0, 0, 1, MAP_HEAD.length);
        h.values = [MAP_HEAD];
        h.format.font.bold = true;
        h.format.fill.color = "#E4EAF3";
      }
      const mapUsed = mapWs.getUsedRangeOrNullObject(true);
      mapUsed.load("values");
      const grab = {};
      for (const sh of ["PL", "BS"]) {
        const ws = wss.getItem(target[sh]);
        const used = ws.getUsedRange(true).getBoundingRect(ws.getRange("A1"));
        used.load("values,formulas,rowCount,columnCount");
        grab[sh] = { ws, used };
      }
      await ctx.sync();

      const mapVals = mapUsed.isNullObject ? [MAP_HEAD] : mapUsed.values;
      const mapping = {};
      mapVals.slice(1).forEach(r => { if (r[0] && r[1]) mapping[r[0] + "|" + mkey(r[1])] = { sources: String(r[2] || "").split(/[＋+]/).map(s => s.trim()).filter(Boolean), method: String(r[3] || "") }; });
      const newMap = [];

      for (const sh of ["PL", "BS"]) {
        const { ws, used } = grab[sh];
        const sheetName = target[sh];
        const vals = used.values, fmls = used.formulas;
        const info = result.sheets[sh] = Object.assign(result.sheets[sh] || {}, { name: sheetName, written: 0, checks: [], noSource: [] });
        const hdr = findHeader(vals);
        if (!hdr) { result.errors.push(`${sheetName}：見出し行（2025/10 形式の月）が見つかりません`); info.error = true; continue; }
        const mLabel = toMonthLabel(period);
        const ci = hdr.cols.find((c, i) => hdr.labels[i] === mLabel);
        if (ci == null) { result.errors.push(`${sheetName}：${mLabel} の列がありません。列を追加してから「月次シートへ反映し直す」を実行してください`); info.error = true; continue; }
        info.col = colLetter(ci);

        // 対象行（PLは経常利益まで、BSは■や【参考】の手前まで）
        const srows = [];
        for (let r = hdr.row + 1; r < vals.length; r++) {
          const raw = String(vals[r][0] == null ? "" : vals[r][0]).trim();
          if (sh === "BS" && (/^■/.test(raw) || /^【参考/.test(raw))) break;
          if (!raw || /^[【※]/.test(raw)) { srows.push({ r, section: true, raw }); continue; }
          if (segOf(raw)) continue;   // 売上高の内訳（保守・開発）は手入力のまま
          const f = String(fmls[r][ci] == null ? "" : fmls[r][ci]);
          srows.push({ r, raw, key: mkey(raw), bkey: base(raw), formula: f.startsWith("="), section: false });
          if (sh === "PL" && ["経常利益", "経常損失"].includes(mkey(raw))) break;
        }
        const P = pdfIndex(rows, sh);
        const findPdf = (nm) => P.byKey[mkey(nm)] || null;
        const checkKeys = new Set(CHECKS[sh].flatMap(c => c.row));

        // 各行の取込元を決める（自分の行がある科目は既定の合算から外す＝二重計上防止）
        const ownRow = new Set(srows.filter(s => !s.section).map(s => s.key));
        const used2 = new Set();     // 参照されたPDF科目（照合キー）
        srows.forEach(s => {
          if (s.section) return;
          const m = mapping[sheetName + "|" + s.key] || mapping["月次" + sh + "|" + s.key];
          if (m) { s.sources = m.sources; s.method = m.method || "手入力"; }
          else if (DEFAULTS[s.bkey]) { s.sources = DEFAULTS[s.bkey].filter(n => mkey(n) === s.key || !ownRow.has(mkey(n))); s.method = "既定"; }
          else if (findPdf(s.raw)) { s.sources = [findPdf(s.raw).name]; s.method = "名前一致"; }
          else { s.sources = []; s.method = "対応なし"; }
          if (!m) newMap.push([sheetName, s.raw, s.sources.join("＋"), s.method]);
          s.sources.forEach(n => used2.add(mkey(n)));
        });
        const writableKeys = new Set(srows.filter(s => !s.section && !s.formula && s.sources.length).flatMap(s => s.sources.map(mkey)));

        // 追加が必要な科目：明細・金額あり・どの行からも参照されず、直近の親も書込行で扱われていない
        const nonzero = (r) => r.bal !== 0 || r.dr !== 0 || r.cr !== 0 || r.open !== 0;
        const adds = [];
        P.list.forEach((r, idx) => {
          if (r.kind === "集計" || !nonzero(r)) return;
          const k = mkey(r.name);
          if (used2.has(k)) return;
          if (r._parent && writableKeys.has(mkey(r._parent.name))) return;
          // 同名で取込元が空の行があれば、そこを使う
          const same = srows.find(s => !s.section && s.key === k);
          if (same) { same.sources = [r.name]; same.method = "名前一致"; used2.add(k); const nm = newMap.find(x => x[0] === sheetName && mkey(x[1]) === k); if (nm) { nm[2] = r.name; nm[3] = "名前一致"; } else newMap.push([sheetName, same.raw, r.name, "名前一致"]); return; }
          // 挿入位置：後ろに続く科目のうち、最初に書込行がある行の直前。なければ前の書込行の直前（合計の範囲内に入れるため）
          const rowOfPdf = (x) => srows.find(s => !s.section && !s.formula && s.sources.map(mkey).includes(mkey(x.name)));
          let at = null;
          for (let j = idx + 1; j < P.list.length; j++) {
            const x = P.list[j];
            if (x.kind === "集計" && x === r._parent) { /* 親の集計行を越えたら探索終了 */ break; }
            const s = rowOfPdf(x); if (s) { at = s.r; break; }
          }
          if (at == null) for (let j = idx - 1; j >= 0; j--) { const s = rowOfPdf(P.list[j]); if (s) { at = s.r; break; } }
          if (at == null) { const sec = srows.find(s => !s.section && checkKeys.has(s.key)); at = sec ? sec.r : vals.length; }
          // ブロックの先頭行の直前に入れると SUM(先頭:末尾) の範囲外になるため、2行目の直前にずらす
          const detailAt = (row) => srows.find(s => s.r === row && !s.section && !s.formula);
          const prevRow = srows.find(s => s.r === at - 1);
          if (detailAt(at) && (!prevRow || prevRow.section || prevRow.formula)) {
            const second = srows.find(s => s.r === at + 1 && !s.section && !s.formula);
            if (second) at = second.r;
          }
          adds.push({ at, pdf: r, order: r.order });
        });

        // 下から挿入（同じ位置はPDFの順が保たれるよう後ろのものから）
        adds.sort((a, b) => b.at - a.at || b.order - a.order);
        adds.forEach(a => { ws.getRange(`${a.at + 1}:${a.at + 1}`).insert("Down"); });
        const shift = (r) => r + adds.filter(a => a.at <= r).length;
        // 挿入後の行番号：同じ位置に入った行は PDF順で並ぶ
        const byAt = {};
        adds.slice().sort((a, b) => a.at - b.at || a.order - b.order).forEach(a => { (byAt[a.at] = byAt[a.at] || []).push(a); });
        Object.entries(byAt).forEach(([at, list]) => {
          const startRow = Number(at) + adds.filter(a => a.at < Number(at)).length;
          list.forEach((a, i) => { a.newR = startRow + i; });
        });

        // 値の書き込み
        const pdfVal = (nm) => { const x = findPdf(nm); return x ? P.val(x) : null; };
        srows.forEach(s => {
          if (s.section || s.formula || !s.sources.length) { if (!s.section && !s.formula && !s.sources.length) info.noSource.push(s.raw); return; }
          const parts = s.sources.map(pdfVal);
          const v = parts.reduce((t, x) => t + (x || 0), 0);
          const cell = ws.getRange(info.col + (shift(s.r) + 1));
          cell.values = [[v]];
          cell.numberFormat = [["#,##0;-#,##0"]];
          info.written++;
        });
        const monthCols = new Set(hdr.cols);
        adds.forEach(a => {
          const rr = a.newR + 1;
          // 月以外の列（累計・計画比など）の数式を上の行からコピー。他シート参照を含む数式はコピーしない
          const srcR = a.at - 1;
          if (srcR > hdr.row) {
            const fr = fmls[srcR] || [];
            fr.forEach((f, c) => {
              if (c === 0 || monthCols.has(c)) return;
              const t = String(f == null ? "" : f);
              if (!t.startsWith("=") || t.includes("!")) return;
              ws.getRange(colLetter(c) + rr).formulas = [[shiftRows(t, rr - (srcR + 1))]];
            });
          }
          const nm = ws.getRange("A" + rr);
          nm.values = [[a.pdf.name]];
          const cell = ws.getRange(info.col + rr);
          cell.values = [[P.val(a.pdf)]];
          cell.numberFormat = [["#,##0;-#,##0"]];
          ws.getRange(`A${rr}:${info.col}${rr}`).format.fill.color = "#FBEEDC";
          result.inserted.push({ sheet: sheetName, name: a.pdf.name, row: rr });
          newMap.push([sheetName, a.pdf.name, a.pdf.name, "追加"]);
          info.written++;
        });
        info.checkRows = CHECKS[sh].map(c => {
          const s = srows.find(x => !x.section && c.row.includes(x.key));
          const pdf = c.pdf.map(pdfVal).find(v => v !== null);
          return { label: c.label, r: s ? shift(s.r) + 1 : null, formula: s ? s.formula : false, pdf: pdf == null ? null : pdf };
        });
      }

      // 科目対応シートへ追記
      if (newMap.length) {
        const r0 = mapVals.length;
        const rg = mapWs.getRangeByIndexes(r0, 0, newMap.length, MAP_HEAD.length);
        rg.numberFormat = newMap.map(() => MAP_HEAD.map(() => "@"));
        rg.values = newMap;
        result.mapAdded = newMap.length;
      }
      await ctx.sync();

      // 検算：数式の再計算後の値を読む
      const reads = [];
      for (const sh of ["PL", "BS"]) {
        const info = result.sheets[sh];
        if (!info || info.error || !info.checkRows) continue;
        const ws = wss.getItem(target[sh]);
        info.checkRows.forEach(c => { if (c.r) { const rg = ws.getRange(info.col + c.r); rg.load("values"); reads.push({ c, rg }); } });
      }
      await ctx.sync();
      reads.forEach(({ c, rg }) => { const v = rg.values[0][0]; c.sheet = typeof v === "number" ? v : (v === "" || v == null ? 0 : Number(String(v).replace(/,/g, ""))); });
      for (const sh of ["PL", "BS"]) {
        const info = result.sheets[sh];
        if (!info || !info.checkRows) continue;
        info.checks = info.checkRows.map(c => {
          if (!c.r) return { label: c.label, status: "none", note: "行が見つかりません" };
          if (c.pdf == null) return { label: c.label, status: "none", sheet: c.sheet, note: "PDFに該当行がありません" };
          const d = Math.round(c.sheet - c.pdf);
          return { label: c.label, status: d === 0 ? "ok" : "ng", sheet: c.sheet, pdf: c.pdf, diff: d, note: c.formula ? "" : "数式ではない行（値を直接比較）" };
        });
      }
    });
    result.ngCount = Object.values(result.sheets).reduce((t, s) => t + (s.checks || []).filter(c => c.status === "ng").length, 0);
    result.ok = !result.errors.length && result.ngCount === 0;
    return result;
  }

  /**
   * 月次PL／月次BSを読み、ダッシュボード用の行（TB_明細と同じ形）を作る。
   * シートに値がない月（前期以前など）は TB_明細 を科目対応で集計して補う。
   * 戻り値：{ found, rows, layout }
   */
  async function readModel(tb, startMonth) {
    return Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets;
      wss.load("items/name");
      await ctx.sync();
      const names = wss.items.map(w => w.name);
      const pick = (frag) => names.find(n => n.indexOf(frag) === 0) || names.find(n => n.indexOf(frag) >= 0);
      const target = { PL: pick("月次PL"), BS: pick("月次BS") };
      if (!target.PL && !target.BS) return { found: false, rows: [], layout: {} };
      const grab = {};
      for (const sh of ["PL", "BS"]) {
        if (!target[sh]) continue;
        const ws = wss.getItem(target[sh]);
        const used = ws.getUsedRange(true).getBoundingRect(ws.getRange("A1"));
        used.load("values,formulas");
        grab[sh] = used;
      }
      const mapWs = wss.getItemOrNullObject(MAP_SHEET);
      await ctx.sync();
      let mapVals = [];
      if (!mapWs.isNullObject) { const u = mapWs.getUsedRangeOrNullObject(true); u.load("values"); await ctx.sync(); if (!u.isNullObject) mapVals = u.values; }
      const mapping = {};
      mapVals.slice(1).forEach(r => { if (r[0] && r[1]) mapping[r[0] + "|" + mkey(r[1])] = String(r[2] || "").split(/[＋+]/).map(x => x.trim()).filter(Boolean); });

      const out = [], layout = {}, seg = {};
      const tbBy = {};
      tb.forEach(r => { ((tbBy[r.period] = tbBy[r.period] || {})[r.sheet] = tbBy[r.period][r.sheet] || []).push(r); });
      for (const sh of ["PL", "BS"]) {
        const used = grab[sh]; if (!used) continue;
        const vals = used.values, fmls = used.formulas;
        const hdr = findHeader(vals); if (!hdr) continue;
        const periods = hdr.labels.map(l => { const [y, m] = l.split("/").map(Number); return y + "-" + String(m).padStart(2, "0"); });
        const srows = [], segRows = [];
        for (let r = hdr.row + 1; r < vals.length; r++) {
          const raw = String(vals[r][0] == null ? "" : vals[r][0]).trim();
          if (sh === "BS" && (/^■/.test(raw) || /^【参考/.test(raw))) break;
          if (!raw || /^[【※]/.test(raw)) { srows.push({ r, section: true, raw }); continue; }
          const sk = sh === "PL" ? segOf(raw) : null;
          if (sk) {
            const v = {};
            hdr.cols.forEach((c, i) => { const x = vals[r][c]; if (typeof x === "number") v[periods[i]] = x; });
            segRows.push({ r, raw, k: sk, vals: v });
            if (!seg[sk]) seg[sk] = v;
            continue;
          }
          const f0 = String(fmls[r][hdr.cols[0]] == null ? "" : fmls[r][hdr.cols[0]]);
          srows.push({ r, raw, key: mkey(raw), bkey: base(raw), formula: f0.startsWith("="), f0 });
          if (sh === "PL" && ["経常利益", "経常損失"].includes(mkey(raw))) break;
        }
        const ownRow = new Set(srows.filter(s => !s.section).map(s => s.key));
        const checkAlias = {};
        CHECKS[sh].forEach(c => c.row.forEach(k => { checkAlias[k] = c.pdf; }));
        srows.forEach(s => {
          if (s.section) return;
          const m = mapping[target[sh] + "|" + s.key];
          if (m) s.sources = m;
          else if (DEFAULTS[s.bkey]) s.sources = DEFAULTS[s.bkey].filter(n => mkey(n) === s.key || !ownRow.has(mkey(n)));
          else if (checkAlias[s.key]) s.sources = checkAlias[s.key].slice(0, 1);
          else s.sources = [s.raw];
        });
        layout[sh] = { sheetName: target[sh], hdrRow: hdr.row, cols: hdr.cols, periods, rows: srows, segRows };
        const items = srows.filter(s => !s.section);
        const key = (s) => sh + ":" + s.key;
        // シートの値
        const have = {};  // period -> key -> value
        periods.forEach((p, i) => {
          const c = hdr.cols[i];
          items.forEach(s => {
            const v = vals[s.r][c];
            if (v === "" || v == null || typeof v !== "number") return;
            (have[p] = have[p] || {})[key(s)] = v;
          });
        });
        // 列が空の月（全科目が空）は、TBから補う
        const tbPeriods = Object.keys(tbBy).filter(p => tbBy[p][sh]);
        const allPeriods = [...new Set([...periods.filter(p => have[p] && Object.values(have[p]).some(v => v !== 0)), ...tbPeriods])].sort();
        const fromTb = (p, s, field) => {
          const list = (tbBy[p] || {})[sh] || [];
          let t = 0, hit = false;
          s.sources.forEach(n => { const x = list.find(r => mkey(r.name) === mkey(n)); if (x) { hit = true; t += field(x); } });
          return hit ? t : null;
        };
        const fyOf = (p) => { const y = Number(p.slice(0, 4)), m = Number(p.slice(5)); return startMonth === 1 ? y : (m >= startMonth ? y + 1 : y); };
        const cum = {};   // PLの期首からの累計
        allPeriods.forEach(p => {
          const onSheet = have[p] && Object.values(have[p]).some(v => v !== 0);
          items.forEach(s => {
            const k = key(s);
            let month = null, bal = null, open = null;
            if (onSheet) {
              const v = have[p][k];
              if (v == null && !(s.formula)) { month = 0; } else month = v == null ? null : v;
              if (month == null) return;
              if (sh === "PL") {
                const prevP = allPeriods[allPeriods.indexOf(p) - 1];
                const base = prevP && fyOf(prevP) === fyOf(p) && cum[prevP] && cum[prevP][k] != null ? cum[prevP][k] : 0;
                open = base; bal = base + month;
              } else {
                bal = month;
                const prevP = allPeriods[allPeriods.indexOf(p) - 1];
                open = prevP && cum[prevP] && cum[prevP][k] != null ? cum[prevP][k] : null;
              }
            } else {
              if (sh === "PL") { bal = fromTb(p, s, x => x.bal); open = fromTb(p, s, x => x.open); }
              else { bal = fromTb(p, s, x => x.bal); open = fromTb(p, s, x => x.open); }
              if (bal == null) return;
              month = bal - (open || 0);
            }
            (cum[p] = cum[p] || {})[k] = bal;
            out.push({ period: p, sheet: sh, order: s.r, kind: s.formula ? "集計" : "明細", key: k, code: "", name: s.raw.replace(/[（(]旧[:：][^）)]*[）)]/g, ""), open: open == null ? null : open, bal, month, src: onSheet ? "sheet" : "tb" });
          });
        });
      }
      return { found: true, rows: out, layout, seg: Object.keys(seg).length ? seg : null };
    });
  }

  async function activateByPrefix(prefix) {
    return Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets; wss.load("items/name"); await ctx.sync();
      const n = wss.items.map(w => w.name).find(x => x.indexOf(prefix) === 0);
      if (!n) return false;
      wss.getItem(n).activate(); await ctx.sync(); return true;
    });
  }

  global.MonthlySheets = { reflect, readModel, activateByPrefix, mkey, base, segOf, moveFormula, findHeader, CHECKS, DEFAULTS };
})(window);
