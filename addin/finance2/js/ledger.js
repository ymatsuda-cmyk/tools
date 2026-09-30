/* ============================================================
 * ledger.js — 財務台帳（Excelブック）の読み書き
 * ------------------------------------------------------------
 * TB_明細 : 1行＝1か月×1科目。「期間＋キー」は常に1行だけ。
 *           差し替えは月単位（旧行を TB_退避 に移してから削除→追記）
 * 取込履歴 : 取込ID・期間・ファイル名・版・状態（有効／置換済）
 * 科目マスタ: キーごとの区分・行種別（明細／集計）。手で直せば次回から優先
 * 設定     : 期首月・会社名・金額区分
 * 月次PL／月次BSシートへの反映は monthly-sheets.js
 * ============================================================ */
(function (global) {
  "use strict";
  const S = { TB: "TB_明細", HIST: "取込履歴", ARCH: "TB_退避", MASTER: "科目マスタ", SET: "設定" };
  const TB_HEAD = ["期間", "区分", "表示順", "行種別", "キー", "科目コード", "科目名", "繰越残高", "借方", "貸方", "残高", "当月", "取込ID"];
  const HIST_HEAD = ["取込ID", "期間", "ファイル名", "会社名", "行数", "取込日時", "版", "状態", "理由", "月次シート検算"];
  const MASTER_HEAD = ["キー", "区分", "科目コード", "科目名", "行種別", "初回期間"];
  const SET_HEAD = ["項目", "値", "説明"];
  const DEFAULT_SETTINGS = [
    ["期首月", 10, "会計年度の開始月（1〜12）"],
    ["会社名", "", "最初に取り込んだPDFの会社名"],
    ["金額区分", "", "PDFの【税込】／【税抜】"],
    ["異常値の下限額", 50000, "この金額より小さい動きは異常値にしない（円）"]
  ];

  const colLetter = (n) => { let s = ""; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const nowStr = () => { const d = new Date(), p = (v) => String(v).padStart(2, "0"); return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };

  // Excelが日付に変換してしまった期間セルも "YYYY-MM" に戻す
  function periodText(v) {
    if (typeof v === "number" && v > 20000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0");
    }
    const s = String(v || "").trim();
    const m = s.match(/^(\d{4})[-/年](\d{1,2})/);
    return m ? m[1] + "-" + m[2].padStart(2, "0") : s;
  }
  const n0 = (v) => (typeof v === "number" ? v : Number(String(v || "0").replace(/,/g, "")) || 0);

  function readSheet(ctx, name) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    return { ws, name };
  }

  /** 台帳全体を読む。台帳シートがなければ exists=false */
  async function readAll() {
    return Excel.run(async (ctx) => {
      const names = [S.TB, S.HIST, S.MASTER, S.SET];
      const refs = names.map(n => readSheet(ctx, n));
      await ctx.sync();
      const used = refs.map(r => {
        if (r.ws.isNullObject) return null;
        const u = r.ws.getUsedRangeOrNullObject(true);
        u.load("values");
        return u;
      });
      await ctx.sync();
      const vals = used.map(u => (u && !u.isNullObject ? u.values : null));
      const [tbV, hV, mV, sV] = vals;
      const tb = (tbV || []).slice(1).filter(r => r[0] !== "" && r[4] !== "").map(r => ({
        period: periodText(r[0]), sheet: String(r[1]), order: n0(r[2]), kind: String(r[3] || "明細"),
        key: String(r[4]), code: String(r[5] || ""), name: String(r[6] || ""),
        open: n0(r[7]), dr: n0(r[8]), cr: n0(r[9]), bal: n0(r[10]), month: n0(r[11]), impId: String(r[12] || "")
      }));
      const history = (hV || []).slice(1).filter(r => r[0] !== "").map(r => ({
        id: String(r[0]), period: periodText(r[1]), file: String(r[2]), company: String(r[3]), rows: n0(r[4]),
        at: String(r[5]), ver: n0(r[6]), status: String(r[7]), reason: String(r[8] || ""), check: String(r[9] || "")
      }));
      const master = {};
      (mV || []).slice(1).forEach(r => { if (r[0] !== "") master[String(r[0])] = { key: String(r[0]), sheet: String(r[1]), code: String(r[2] || ""), name: String(r[3]), kind: String(r[4] || "明細") }; });
      const settings = { 期首月: 10, 会社名: "", 金額区分: "" };
      (sV || []).slice(1).forEach(r => { if (r[0] !== "") settings[String(r[0])] = r[1]; });
      settings.期首月 = Math.min(12, Math.max(1, n0(settings.期首月) || 10));
      return { exists: !!tbV, tb, history, master, settings };
    });
  }

  async function ensureSheets(ctx) {
    const defs = [[S.TB, TB_HEAD], [S.HIST, HIST_HEAD], [S.ARCH, TB_HEAD.concat(["退避日時"])], [S.MASTER, MASTER_HEAD], [S.SET, SET_HEAD]];
    const refs = defs.map(([n]) => ctx.workbook.worksheets.getItemOrNullObject(n));
    await ctx.sync();
    defs.forEach(([n, head], i) => {
      if (!refs[i].isNullObject) return;
      const ws = ctx.workbook.worksheets.add(n);
      const h = ws.getRangeByIndexes(0, 0, 1, head.length);
      h.values = [head];
      h.format.font.bold = true;
      h.format.fill.color = "#E4EAF3";
      try { ws.freezePanes.freezeRows(1); } catch (e) {}
      if (n === S.SET) {
        const r = ws.getRangeByIndexes(1, 0, DEFAULT_SETTINGS.length, 3);
        r.values = DEFAULT_SETTINGS;
      }
    });
    await ctx.sync();
  }

  function usedCount(ctx, name) {
    const ws = ctx.workbook.worksheets.getItem(name);
    const u = ws.getUsedRangeOrNullObject(true);
    u.load("rowCount,values");
    return { ws, u };
  }

  /**
   * 1か月分を書き込む。mode: "new" | "replace"
   * parsed: TBParser の結果, master: 既存マスタ（行種別の上書き用）
   */
  async function writePeriod(parsed, opt) {
    const { mode, reason = "", master = {} } = opt;
    const period = parsed.period;
    return Excel.run(async (ctx) => {
      await ensureSheets(ctx);
      const tbR = usedCount(ctx, S.TB), hR = usedCount(ctx, S.HIST), mR = usedCount(ctx, S.MASTER), arR = usedCount(ctx, S.ARCH), sR = usedCount(ctx, S.SET);
      await ctx.sync();
      const tbVals = tbR.u.isNullObject ? [TB_HEAD] : tbR.u.values;
      const hVals = hR.u.isNullObject ? [HIST_HEAD] : hR.u.values;

      // 取込ID
      let maxId = 0;
      hVals.slice(1).forEach(r => { const m = String(r[0]).match(/(\d+)$/); if (m) maxId = Math.max(maxId, Number(m[1])); });
      const impId = "IMP-" + String(maxId + 1).padStart(4, "0");
      const ver = hVals.slice(1).filter(r => periodText(r[1]) === period).length + 1;

      // 既存行（同じ期間）
      const idx = [];
      tbVals.forEach((r, i) => { if (i > 0 && periodText(r[0]) === period) idx.push(i); });
      if (mode === "new" && idx.length) throw new Error(period + " は登録済みです。差分確認から差し替えてください。");

      let tbCount = tbVals.length;
      if (idx.length) {
        // 退避シートへコピー
        const arCount = arR.u.isNullObject ? 1 : arR.u.values.length;
        const stamp = nowStr();
        const arch = idx.map(i => tbVals[i].slice(0, TB_HEAD.length).concat([stamp]));
        const ar = arR.ws.getRangeByIndexes(arCount, 0, arch.length, TB_HEAD.length + 1);
        ar.numberFormat = arch.map(r => r.map((_, c) => (c === 0 || c === 5 ? "@" : "General")));
        ar.values = arch;
        // 連続区間ごとに下から削除
        const runs = [];
        idx.forEach(i => { const last = runs[runs.length - 1]; if (last && last[1] === i - 1) last[1] = i; else runs.push([i, i]); });
        runs.reverse().forEach(([a, b]) => { tbR.ws.getRangeByIndexes(a, 0, b - a + 1, TB_HEAD.length).delete("Up"); });
        tbCount -= idx.length;
        // 履歴の状態を置換済に
        hVals.forEach((r, i) => {
          if (i > 0 && periodText(r[1]) === period && String(r[7]) === "有効") hR.ws.getRange("H" + (i + 1)).values = [["置換済"]];
        });
      }

      // 追記
      const rows = parsed.rows.map(r => {
        const kind = (master[r.key] && master[r.key].kind) || r.kind;
        return [period, r.sheet, r.order, kind, r.key, r.code, r.name, r.open, r.dr, r.cr, r.bal, r.bal - r.open, impId];
      });
      const tgt = tbR.ws.getRangeByIndexes(tbCount, 0, rows.length, TB_HEAD.length);
      tgt.numberFormat = rows.map(() => TB_HEAD.map((_, c) => (c === 0 || c === 5 ? "@" : (c >= 7 && c <= 11 ? "#,##0" : "General"))));
      tgt.values = rows;

      // 履歴
      const hRow = [[impId, period, parsed.fileName, parsed.company, rows.length, nowStr(), ver, "有効", reason, ""]];
      const hr = hR.ws.getRangeByIndexes(hVals.length, 0, 1, HIST_HEAD.length);
      hr.numberFormat = [HIST_HEAD.map((_, c) => (c === 1 || c === 5 ? "@" : "General"))];
      hr.values = hRow;

      // マスタ（未登録キーだけ追加）
      const mVals = mR.u.isNullObject ? [MASTER_HEAD] : mR.u.values;
      const have = new Set(mVals.slice(1).map(r => String(r[0])));
      const add = parsed.rows.filter(r => !have.has(r.key)).map(r => [r.key, r.sheet, r.code, r.name, r.kind, period]);
      if (add.length) {
        const mr = mR.ws.getRangeByIndexes(mVals.length, 0, add.length, MASTER_HEAD.length);
        mr.numberFormat = add.map(() => MASTER_HEAD.map((_, c) => (c === 2 || c === 5 ? "@" : "General")));
        mr.values = add;
      }

      // 設定（会社名・金額区分が空なら埋める）
      if (!sR.u.isNullObject) {
        sR.u.values.forEach((r, i) => {
          if (i === 0) return;
          if (r[0] === "会社名" && !r[1] && parsed.company) sR.ws.getRange("B" + (i + 1)).values = [[parsed.company]];
          if (r[0] === "金額区分" && !r[1] && parsed.taxMode) sR.ws.getRange("B" + (i + 1)).values = [[parsed.taxMode]];
        });
      }
      await ctx.sync();
      return { impId, ver, rows: rows.length };
    });
  }

  /** 取込履歴の「月次シート検算」欄を更新する */
  async function setHistCheck(impId, text) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(S.HIST);
      await ctx.sync();
      if (ws.isNullObject) return;
      const u = ws.getUsedRangeOrNullObject(true); u.load("values"); await ctx.sync();
      if (u.isNullObject) return;
      const i = u.values.findIndex(r => String(r[0]) === impId);
      if (i > 0) { ws.getRange("J" + (i + 1)).values = [[text]]; await ctx.sync(); }
    });
  }

  async function activate(name, cell) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
      await ctx.sync();
      if (ws.isNullObject) return;
      ws.activate();
      await ctx.sync();
      if (cell) { ws.getRange(cell).select(); await ctx.sync(); }
    });
  }

  global.Ledger = { S, readAll, writePeriod, setHistCheck, activate, periodText };
})(window);
