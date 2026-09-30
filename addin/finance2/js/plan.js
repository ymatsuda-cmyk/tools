/* ============================================================
 * plan.js — 計画シート（前期実績ベース）の作成と読み込み
 * ------------------------------------------------------------
 * 月次PLと同じ並びで行を作る。上段＝計画、下段＝前期実績（参照用）。
 *   各月の計画 ＝ ROUND(前期同月 ×(1+調整率) + 調整額/12, 0)
 *   合計・利益の行は月次PLの数式を同じ位置関係で写す
 * 月のセルを直接上書きしてもよい（読み込みは値で行う）。
 * ============================================================ */
(function (global) {
  "use strict";
  const SHEET = "計画";
  const colLetter = (n) => { let s = ""; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const toLabel = (p) => p.slice(0, 4) + "/" + Number(p.slice(5));
  const MC0 = 6; // G列から12か月
  const NCOL = 19;

  /**
   * layout: MonthlySheets.readModel の layout.PL
   * M: 表示用モデル, targetFy: 計画する期, fyLabel: 表示名
   */
  async function create(layout, M, targetFy) {
    const baseFy = targetFy - 1;
    const tMonths = M.fyMonths(targetFy), bMonths = M.fyMonths(baseFy);
    const rows = layout.rows;
    const firstR = rows[0].r, span = rows[rows.length - 1].r - firstR + 1;
    const fromCol = colLetter(layout.cols[0]);
    const MAIN = 5;                         // 0始まりの行番号（Excel 6行目）
    const BASE_TITLE = MAIN + span + 2, BASE_HDR = BASE_TITLE + 1, BASE = BASE_HDR + 1;
    const total = BASE + span;
    const grid = Array.from({ length: total }, () => Array(NCOL).fill(""));
    const fmt = Array.from({ length: total }, () => Array(NCOL).fill("General"));
    const bold = [], inputs = [], sections = [];
    grid[0][0] = "計画（前期実績ベース）";
    grid[1] = ["対象期", M.fyLabel(targetFy), "前期", M.fyLabel(baseFy), "作成", new Date().toLocaleDateString("ja-JP")].concat(Array(NCOL - 6).fill(""));
    grid[2][0] = "C列（調整率）・D列（年間の調整額）を入れると各月の計画が変わります。月のセルを直接上書きしても構いません（finance2は値を読みます）。";
    grid[4] = ["科目", "前期実績（年間）", "調整率", "調整額（年間）", "年間計画", "前期比"].concat(tMonths.map(toLabel)).concat(["備考"]);
    grid[BASE_TITLE][0] = `▼ 前期実績（${M.fyLabel(baseFy)}・月次）― 計画の基準。ここは編集しない`;
    grid[BASE_HDR] = ["科目", "前期実績（年間）", "", "", "", ""].concat(bMonths.map(toLabel)).concat([""]);
    bold.push(4, BASE_HDR);
    const warn = [];
    rows.forEach(s => {
      const pr = MAIN + (s.r - firstR), br = BASE + (s.r - firstR);
      const PR = pr + 1, BR = br + 1, sheetRow = s.r + 1;
      grid[pr][0] = s.raw; grid[br][0] = s.raw;
      if (s.section) { if (s.raw) { sections.push(pr, br); } return; }
      const key = "PL:" + s.key;
      const sumB = `=SUM(G${BR}:R${BR})`;
      grid[pr][1] = sumB; grid[br][1] = sumB;
      grid[pr][4] = `=SUM(G${PR}:R${PR})`;
      grid[pr][5] = `=IF(B${PR}=0,"",E${PR}/B${PR}-1)`;
      fmt[pr][5] = "0.0%";
      [1, 3, 4].forEach(c => { fmt[pr][c] = "#,##0;-#,##0"; });
      fmt[br][1] = "#,##0;-#,##0";
      for (let i = 0; i < 12; i++) {
        const col = colLetter(MC0 + i);
        fmt[pr][MC0 + i] = fmt[br][MC0 + i] = "#,##0;-#,##0";
        if (s.formula) {
          if (String(s.f0).includes("!")) { warn.push(s.raw); continue; }
          grid[pr][MC0 + i] = MonthlySheets.moveFormula(s.f0, PR - sheetRow, fromCol, col);
          grid[br][MC0 + i] = MonthlySheets.moveFormula(s.f0, BR - sheetRow, fromCol, col);
        } else {
          const v = M.series(key, bMonths[i]);
          grid[br][MC0 + i] = v == null ? 0 : Math.round(v);
          grid[pr][MC0 + i] = `=ROUND(${col}${BR}*(1+$C${PR})+$D${PR}/12,0)`;
        }
      }
      if (s.formula) bold.push(pr, br);
      else { grid[pr][2] = 0; grid[pr][3] = 0; fmt[pr][2] = "0.0%"; fmt[pr][3] = "#,##0;-#,##0"; inputs.push(pr); }
    });

    await Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets;
      const old = wss.getItemOrNullObject(SHEET);
      await ctx.sync();
      if (!old.isNullObject) { old.delete(); await ctx.sync(); }
      const ws = wss.add(SHEET);
      const rg = ws.getRangeByIndexes(0, 0, total, NCOL);
      rg.numberFormat = fmt;
      rg.formulas = grid;
      ws.getRange("A1").format.font.bold = true;
      bold.forEach(r => { ws.getRangeByIndexes(r, 0, 1, NCOL).format.font.bold = true; });
      sections.forEach(r => { ws.getRangeByIndexes(r, 0, 1, 1).format.font.bold = true; });
      [4, BASE_HDR].forEach(r => { ws.getRangeByIndexes(r, 0, 1, NCOL).format.fill.color = "#E4EAF3"; });
      inputs.forEach(r => { const x = ws.getRangeByIndexes(r, 2, 1, 2); x.format.fill.color = "#FFF4D6"; x.format.font.color = "#1F3F6E"; });
      ws.getRangeByIndexes(BASE_TITLE, 0, span + 2, NCOL).format.font.color = "#56606B";
      try { ws.freezePanes.freezeAt(ws.getRange("A5")); } catch (e) {}
      try { ws.getRange("A:A").format.columnWidth = 170; ws.getRange("B:R").format.columnWidth = 80; } catch (e) {}
      ws.activate();
      await ctx.sync();
    });
    return { rows: span, warn };
  }

  /** 計画シートを読む。{ exists, fyLabel, periods, rows:[{key,name,kind,values}] } */
  async function read() {
    return Excel.run(async (ctx) => {
      const wss = ctx.workbook.worksheets;
      wss.load("items/name");
      await ctx.sync();
      const name = wss.items.map(w => w.name).find(n => n === SHEET) || wss.items.map(w => w.name).find(n => n.indexOf(SHEET) === 0);
      if (!name) return { exists: false };
      const ws = wss.getItem(name);
      const used = ws.getUsedRange(true).getBoundingRect(ws.getRange("A1"));
      used.load("values,formulas");
      await ctx.sync();
      const vals = used.values, fmls = used.formulas;
      const hdr = MonthlySheets.findHeader(vals);
      if (!hdr) return { exists: true, error: "計画シートの見出し行（2026/10 形式の月）が見つかりません" };
      const periods = hdr.labels.map(l => { const [y, m] = l.split("/").map(Number); return y + "-" + String(m).padStart(2, "0"); });
      const rows = [];
      for (let r = hdr.row + 1; r < vals.length; r++) {
        const raw = String(vals[r][0] == null ? "" : vals[r][0]).trim();
        if (/^▼/.test(raw)) break;
        if (!raw || /^[【※]/.test(raw)) continue;
        const values = hdr.cols.map(c => (typeof vals[r][c] === "number" ? vals[r][c] : null));
        const isF = String(fmls[r][hdr.cols[0]] || "").startsWith("=") && !/ROUND\(/.test(String(fmls[r][hdr.cols[0]] || ""));
        rows.push({ key: "PL:" + MonthlySheets.mkey(raw), name: raw.replace(/[（(]旧[:：][^）)]*[）)]/g, ""), kind: isF ? "集計" : "明細", values });
      }
      return { exists: true, sheet: name, fyLabel: String(vals[1] && vals[1][1] || ""), periods, rows };
    });
  }

  global.Plan = { create, read, SHEET };
})(window);
