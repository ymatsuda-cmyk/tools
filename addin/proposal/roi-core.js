/* ============================================================
 * roi-core.js — 議事録・ROI試算の共有ライブラリ
 * ------------------------------------------------------------
 * 「営業報告」アドインと「提案ナレッジ」アドインの両方から読み込む。
 * 同じGitHub Pagesドメイン配下に置くこと（両アドインとも
 * ymatsuda-cmyk.github.io なので相互に script タグで読み込める）。
 *
 * このファイルは DOM を一切触らない「データ層」。
 * ・営業報告アドイン　　→ 簡単な操作のみ（紐づけ・参照・提案作成・提案参照）
 * ・提案ナレッジアドイン → ROIマスタ編集、抽出結果の手動レビュー、
 *                        プロンプト組み立てなど詳細機能
 * という役割分担で、両方がこのファイルの関数を呼ぶ。
 *
 * 読み込み例:
 *   <script src="https://ymatsuda-cmyk.github.io/tools/addin/roi-knowledge/roi-core.js"></script>
 *   window.RoiCore.xxx(...) で呼び出す。
 * ============================================================ */

(function (global) {

  const EIGYO_SHEET = "営業報告";
  const CUST_SHEET = "顧客マスタ";

  const MASTER_SHEET = "ROIマスタ";
  const MASTER_COLUMNS = ["課題カテゴリ", "項目ID", "項目名", "区分", "単位", "デフォルト値", "数式", "信頼度初期値"];

  const HEARING_SHEET = "議事録";
  // 参照URL: PLAUD/Notion等、既存の議事録ビューアに保管されているテキストへの
  // リンクだけを持たせるケースを想定した列。発言テキストは空でもよい。
  // 議事録ID: ROI試算側から「どの議事録から抽出したか」を辿るための一意キー。
  const HEARING_COLUMNS = ["案件ID", "課題カテゴリ", "発言者", "参照URL", "発言テキスト", "登録日時", "議事録ID"];

  const CALC_SHEET = "ROI試算";
  // 根拠議事録ID: そのカテゴリを作成するときに参照した議事録IDのカンマ区切り。
  // 「提案書データの元データリンク一覧」表示に使う。
  const CALC_COLUMNS = ["案件ID", "課題カテゴリ", "項目ID", "項目名", "区分", "値", "単位", "信頼度", "選択", "更新日時", "根拠議事録ID"];

  const SOLUTION_SHEET = "ソリューションDB";
  const SOLUTION_COLUMNS = ["課題カテゴリ", "解決策名", "概要"];

  const CONF_LEVELS = ["確定", "推定", "未確認"];

  const MASTER_SEED = [
    ["在庫管理", "stk_people", "棚卸人数", "入力", "人", 3, "", "未確認"],
    ["在庫管理", "stk_hours", "棚卸時間", "入力", "時間", 4, "", "未確認"],
    ["在庫管理", "stk_freq", "棚卸回数/年", "入力", "回", 12, "", "未確認"],
    ["在庫管理", "stk_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["在庫管理", "stk_improve", "改善率", "入力", "%", 60, "", "推定"],
    ["在庫管理", "stk_hours_yr", "年間棚卸工数", "出力", "時間", "", "stk_people*stk_hours*stk_freq", ""],
    ["在庫管理", "stk_cost_yr", "年間棚卸コスト", "出力", "円", "", "stk_hours_yr*stk_wage", ""],
    ["在庫管理", "stk_saving", "削減額", "出力", "円", "", "stk_cost_yr*stk_improve/100", ""],
    ["ロット管理", "lot_hours", "追跡時間", "入力", "時間", 2, "", "未確認"],
    ["ロット管理", "lot_freq", "追跡回数/年", "入力", "回", 100, "", "未確認"],
    ["ロット管理", "lot_people", "担当人数", "入力", "人", 2, "", "未確認"],
    ["ロット管理", "lot_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["ロット管理", "lot_improve", "改善率", "入力", "%", 60, "", "推定"],
    ["ロット管理", "lot_hours_yr", "年間追跡工数", "出力", "時間", "", "lot_hours*lot_freq*lot_people", ""],
    ["ロット管理", "lot_saving", "削減額", "出力", "円", "", "lot_hours_yr*lot_wage*lot_improve/100", ""],
    ["AI議事録", "min_meetings", "会議回数/月", "入力", "回", 8, "", "未確認"],
    ["AI議事録", "min_people", "参加人数", "入力", "人", 3, "", "未確認"],
    ["AI議事録", "min_hours", "議事録作成時間", "入力", "時間", 1, "", "未確認"],
    ["AI議事録", "min_wage", "時給", "入力", "円", 3000, "", "推定"],
    ["AI議事録", "min_hours_yr", "年間工数", "出力", "時間", "", "min_meetings*12*min_people*min_hours", ""],
    ["AI議事録", "min_saving", "削減額", "出力", "円", "", "min_hours_yr*min_wage", ""],
  ];

  const SOLUTION_SEED = [
    ["在庫管理", "在庫管理システム", "棚卸・在庫差異の自動集計"],
    ["ロット管理", "ロット管理システム化", "追跡工数の削減とトレーサビリティ確保"],
    ["AI議事録", "AI議事録", "文字起こし・要約・タスク抽出の自動化"],
  ];

  let masterItems = null; // 初回 ensureAllSheets() 後にキャッシュ

  /* ---------- 設定（AIエンドポイント） ----------
   * localStorage にはエンドポインURLと簡易トークンのみ保存する。
   * APIキー本体はサーバ（GAS等）側にのみ保持し、ここには保存しない。
   * 営業報告・提案ナレッジの両アドインは同一オリジンなので localStorage を共有する。 */
  function getConfig() {
    try { return JSON.parse(localStorage.getItem("roiAddinConfig") || "{}"); }
    catch (e) { return {}; }
  }
  function setConfig(cfg) { localStorage.setItem("roiAddinConfig", JSON.stringify(cfg)); }

  /* ---------- シートの自動作成 ---------- */
  function colLetterOf(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  async function getOrCreateSheet(ctx, name, columns, seedRows) {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let ws = sheets.items.find(s => s.name === name);
    if (!ws) {
      ws = sheets.add(name);
      const lastCol = colLetterOf(columns.length);
      const hdr = ws.getRange(`A1:${lastCol}1`);
      hdr.values = [columns];
      hdr.format.fill.color = "#44546A";
      hdr.format.font.color = "#FFFFFF";
      hdr.format.font.bold = true;
      if (seedRows && seedRows.length) ws.getRange(`A2:${lastCol}${seedRows.length + 1}`).values = seedRows;
      await ctx.sync();
    }
    return ws;
  }

  function rowToMasterItem(r) {
    return { category: r[0], itemId: r[1], name: r[2], kind: r[3], unit: r[4], defaultVal: r[5], formula: r[6], confDefault: r[7] };
  }

  /* すべての必要シートを用意し、ROIマスタをキャッシュして返す。
   * 営業報告・提案ナレッジどちらの起動時にも最初に呼ぶ。 */
  async function ensureAllSheets() {
    if (!global.Office || !global.Excel) {
      masterItems = MASTER_SEED.map(rowToMasterItem);
      return { demo: true, masterItems };
    }
    await Excel.run(async ctx => {
      await getOrCreateSheet(ctx, MASTER_SHEET, MASTER_COLUMNS, MASTER_SEED);
      await getOrCreateSheet(ctx, HEARING_SHEET, HEARING_COLUMNS, null);
      await getOrCreateSheet(ctx, CALC_SHEET, CALC_COLUMNS, null);
      await getOrCreateSheet(ctx, SOLUTION_SHEET, SOLUTION_COLUMNS, SOLUTION_SEED);
      const rng = ctx.workbook.worksheets.getItem(MASTER_SHEET).getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      masterItems = rng.values.slice(1).filter(r => r[1]).map(rowToMasterItem);
    });
    return { demo: false, masterItems };
  }

  function getMasterItems() { return masterItems || []; }
  function getCategories() { return Array.from(new Set(getMasterItems().map(m => m.category))); }

  /* ---------- 案件IDの候補（営業報告シートのID列） ---------- */
  async function listCaseIds() {
    if (!global.Office || !global.Excel) return [];
    const ids = new Set();
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      if (!sheets.items.find(s => s.name === EIGYO_SHEET)) return;
      const sheet = ctx.workbook.worksheets.getItem(EIGYO_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const lastRow = Math.min(Math.max(used.rowCount, 1), 1000);
      if (lastRow < 2) return;
      const rng = sheet.getRange(`A2:A${lastRow}`);
      rng.load("values");
      await ctx.sync();
      rng.values.forEach(r => { if (r[0]) ids.add(String(r[0])); });
    });
    return Array.from(ids);
  }

  /* ---------- 議事録：紐づけ・参照 ---------- */
  async function appendHearingLog(caseId, category, speaker, { text = "", url = "" } = {}) {
    const hearingId = genId("H");
    const row = [caseId, category, speaker, url, text, nowStr(), hearingId];
    if (!global.Office || !global.Excel) return hearingId; // デモモードは書き込みなし
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const nextRow = Math.max(used.rowCount, 1) + 1;
      sheet.getRange(`A${nextRow}:G${nextRow}`).values = [row];
      await ctx.sync();
    });
    return hearingId;
  }

  async function listHearingLogs(caseId) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
    return rows.map(r => ({ category: r[1], speaker: r[2], url: r[3], text: r[4], registeredAt: r[5], hearingId: r[6] }));
  }

  function genId(prefix) {
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  }

  /* ---------- AI抽出（GAS等のWebhook） ----------
   * url が指定されている場合はサーバ側（GAS）でページ内容を取得させる。
   * レスポンス: { items: [{ itemId, value, confidence }] } */
  async function callExtractionWebhook(caseId, category, { text = "", url = "" } = {}) {
    const cfg = getConfig();
    if (!cfg.webhookUrl) throw new Error("AI連携エンドポイントが未設定です");
    const items = getMasterItems().filter(m => m.category === category && m.kind === "入力");
    const res = await fetch(cfg.webhookUrl, {
      method: "POST",
      body: JSON.stringify({
        token: cfg.token || "",
        caseId, category, text, url,
        items: items.map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
      }),
    });
    const raw = await res.text();
    const data = JSON.parse(raw);
    return data.items || [];
  }

  /* ---------- ROIマスタの数式(項目IDのトークン式)をExcel数式に変換 ---------- */
  function translateFormula(masterFormula, rowNum) {
    if (!masterFormula) return "";
    const translated = masterFormula.replace(/[a-zA-Z_][a-zA-Z0-9_]*/g, tok =>
      `INDEX($F$2:$F$9999,MATCH(1,($A$2:$A$9999=$A${rowNum})*($C$2:$C$9999="${tok}"),0))`);
    return "=" + translated;
  }

  /* ---------- ROI試算シートへの反映（カテゴリ丸ごと） ----------
   * 1. 対象カテゴリの全項目がこの案件の行としてまだ無ければROIマスタからコピー
   *   （出力行は数式化）。
   * 2. 渡された抽出値・信頼度を、該当する入力行の値・信頼度に上書きする。 */
  async function applyCategoryToCalcSheet(caseId, category, extractedItems, hearingIds = []) {
    if (!global.Office || !global.Excel) return;
    const provenance = hearingIds.filter(Boolean).join(",");
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("values, rowCount");
      await ctx.sync();

      const existingRows = used.values.slice(1);
      const keyOf = r => `${r[0]}__${r[2]}`;
      const existingIndex = {};
      existingRows.forEach((r, i) => { existingIndex[keyOf(r)] = i + 2; });

      const catItems = getMasterItems().filter(m => m.category === category);
      const toAppend = [];
      const nowTs = nowStr();

      catItems.forEach(m => {
        const key = `${caseId}__${m.itemId}`;
        if (existingIndex[key]) return;
        const rowNum = used.rowCount + 1 + toAppend.length;
        let valueCell;
        if (m.kind === "出力") {
          valueCell = translateFormula(m.formula, rowNum);
        } else {
          const ex = extractedItems.find(e => e.itemId === m.itemId);
          valueCell = ex && ex.value != null ? ex.value : m.defaultVal;
        }
        const confCell = m.kind === "出力" ? "" : ((extractedItems.find(e => e.itemId === m.itemId) || {}).confidence || m.confDefault);
        toAppend.push([caseId, m.category, m.itemId, m.name, m.kind, valueCell, m.unit, confCell, "FALSE", nowTs, provenance]);
      });

      if (toAppend.length) {
        const startRow = used.rowCount + 1;
        const endRow = startRow + toAppend.length - 1;
        const rangeAll = sheet.getRange(`A${startRow}:K${endRow}`);
        rangeAll.values = toAppend.map(r => r.map((v, i) => (i === 5 && typeof v === "string" && v.startsWith("=")) ? "" : v));
        const formulaCol = sheet.getRange(`F${startRow}:F${endRow}`);
        formulaCol.formulas = toAppend.map(r => [typeof r[5] === "string" && r[5].startsWith("=") ? r[5] : (r[5] === "" ? "" : r[5])]);
        await ctx.sync();
      }

      const updates = [];
      extractedItems.forEach(e => {
        const rowNum = existingIndex[`${caseId}__${e.itemId}`];
        if (rowNum) updates.push({ rowNum, value: e.value, confidence: e.confidence });
      });
      updates.forEach(u => {
        sheet.getRange(`F${u.rowNum}`).values = [[u.value]];
        sheet.getRange(`H${u.rowNum}`).values = [[u.confidence]];
        sheet.getRange(`J${u.rowNum}`).values = [[nowTs]];
      });
      if (updates.length) await ctx.sync();
    });
  }

  /* ---------- 「簡単な操作」向け：抽出→即反映を1呼び出しでまとめる ----------
   * 営業報告アドインの「ROI提案」ボタンはこれだけを呼ぶ。
   * レビュー画面は挟まず、AIの抽出結果をそのまま保存する（詳細な確認・修正が
   * 必要な場合は提案ナレッジアドイン側の runExtractionForReview を使う）。 */
  async function quickCreateProposal(caseId, category, { text = "", url = "" } = {}) {
    let hearingId = null;
    if (text || url) hearingId = await appendHearingLog(caseId, category, "顧客", { text, url });
    const items = await callExtractionWebhook(caseId, category, { text, url });
    await applyCategoryToCalcSheet(caseId, category, items, hearingId ? [hearingId] : []);
    return getProposalSummaryForCase(caseId, category);
  }

  /* ---------- 複数カテゴリの一括抽出（営業報告の「作成」アイコン用） ----------
   * この案件に紐づく議事録すべて＋メモ（営業報告シートの備考等）を渡し、
   * AIに「どの課題カテゴリが当てはまるか」ごと判定させ、該当カテゴリを
   * まとめてROI試算シートに反映する。
   * 精度優先の簡易実装のため、根拠議事録IDはこの案件の全議事録IDを
   * まとめて記録する（カテゴリ単位でどの発言が根拠かまでは追跡しない）。 */
  async function autoExtractProposals(caseId, memoText = "") {
    const cfg = getConfig();
    if (!cfg.webhookUrl) throw new Error("AI連携エンドポイントが未設定です");
    const logs = await listHearingLogs(caseId);
    const combinedText = [memoText, ...logs.map(l => l.text || l.url || "")].filter(Boolean).join("\n\n");
    if (!combinedText) return [];
    const hearingIds = logs.map(l => l.hearingId).filter(Boolean);

    const categoryDefs = getCategories().map(cat => ({
      category: cat,
      items: getMasterItems().filter(m => m.category === cat && m.kind === "入力").map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
    }));

    const res = await fetch(cfg.webhookUrl, {
      method: "POST",
      body: JSON.stringify({ mode: "auto", token: cfg.token || "", caseId, text: combinedText, categories: categoryDefs }),
    });
    const data = JSON.parse(await res.text());
    const results = data.results || [];
    for (const r of results) {
      if (r.items && r.items.length) await applyCategoryToCalcSheet(caseId, r.category, r.items, hearingIds);
    }
    return results.map(r => r.category);
  }

  /* 選択済みカテゴリの元データ（議事録）リンク一覧を返す。
   * ROI試算の「根拠議事録ID」列に記録されたIDから議事録を引く。 */
  async function getSourceEntriesForCategory(caseId, category) {
    if (!global.Office || !global.Excel) return [];
    let provenance = "";
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const row = rng.values.slice(1).find(r => String(r[0]) === caseId && r[1] === category && r[10]);
      provenance = row ? row[10] : "";
    });
    const ids = new Set(provenance.split(",").filter(Boolean));
    if (!ids.size) return [];
    const logs = await listHearingLogs(caseId);
    return logs.filter(l => ids.has(l.hearingId));
  }

  /* 詳細レビュー用：抽出だけ行い、保存はしない（提案ナレッジのレビューUIが使う） */
  async function runExtractionForReview(caseId, category, { text = "", url = "" } = {}) {
    return callExtractionWebhook(caseId, category, { text, url });
  }
  /* レビュー後、編集済みの値で保存する（提案ナレッジのレビューUIが使う） */
  async function applyReviewedItems(caseId, category, items) {
    return applyCategoryToCalcSheet(caseId, category, items);
  }

  /* ---------- 提案サマリの参照（両アドイン共通） ----------
   * category を指定すればそのカテゴリだけ、省略すれば案件の全カテゴリを返す。 */
  async function getProposalSummaryForCase(caseId, onlyCategory) {
    if (!global.Office || !global.Excel) return [];
    let calcRows = [], solutions = [];
    await Excel.run(async ctx => {
      const calcSheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const r1 = calcSheet.getUsedRange(true);
      r1.load("values");
      const solSheet = ctx.workbook.worksheets.getItem(SOLUTION_SHEET);
      const r2 = solSheet.getUsedRange(true);
      r2.load("values");
      await ctx.sync();
      calcRows = r1.values.slice(1).filter(r => String(r[0]) === caseId);
      solutions = r2.values.slice(1);
    });

    const byCategory = {};
    calcRows.forEach(r => {
      const cat = r[1];
      byCategory[cat] = byCategory[cat] || { inputs: [], outputs: [] };
      (r[4] === "出力" ? byCategory[cat].outputs : byCategory[cat].inputs).push(r);
    });

    const cats = onlyCategory ? [onlyCategory] : Object.keys(byCategory);
    return cats.filter(c => byCategory[c]).map(cat => {
      const g = byCategory[cat];
      const sol = solutions.find(s => s[0] === cat);
      const savingRow = g.outputs.find(r => String(r[2]).endsWith("_saving")) || g.outputs[0];
      const conf = g.inputs.some(r => r[7] === "未確認") ? "未確認"
        : g.inputs.some(r => r[7] === "推定") ? "推定" : "確定";
      const selected = g.outputs.concat(g.inputs).some(r => String(r[8]).toUpperCase() === "TRUE");
      return {
        category: cat,
        solutionName: sol ? sol[1] : null,
        solutionDesc: sol ? sol[2] : null,
        saving: savingRow ? savingRow[5] : null,
        unit: savingRow ? savingRow[6] : "",
        confidence: conf,
        selected,
      };
    });
  }

  /* プロンプト出力など、生の明細行が必要な場面向け（提案ナレッジ側で使用）。
   * onlySelected=true なら「選択」列がTRUEのカテゴリの行のみを返す。 */
  async function getCalcRowsForCase(caseId, { onlySelected = false } = {}) {
    if (!global.Office || !global.Excel) return [];
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
    if (onlySelected) rows = rows.filter(r => String(r[8]).toUpperCase() === "TRUE");
    return rows.map(r => ({
      caseId: r[0], category: r[1], itemId: r[2], name: r[3], kind: r[4],
      value: r[5], unit: r[6], confidence: r[7], selected: String(r[8]).toUpperCase() === "TRUE",
    }));
  }

  async function toggleSelection(caseId, category, checked) {
    if (!global.Office || !global.Excel) return;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rng.values.forEach((r, i) => {
        if (i === 0) return;
        if (String(r[0]) === caseId && r[1] === category) sheet.getRange(`I${i + 1}`).values = [[checked ? "TRUE" : "FALSE"]];
      });
      await ctx.sync();
    });
  }

  async function getSolutions() {
    if (!global.Office || !global.Excel) return SOLUTION_SEED.map(r => ({ category: r[0], name: r[1], desc: r[2] }));
    let rows = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SOLUTION_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1);
    });
    return rows.map(r => ({ category: r[0], name: r[1], desc: r[2] }));
  }

  async function getCustomerInfo(caseId) {
    if (!global.Office || !global.Excel) return null;
    let row = null;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(CUST_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      const code = caseId.split("-")[0];
      row = rng.values.slice(1).find(r => r[0] === code) || null;
    });
    if (!row) return null;
    return { code: row[0], name: row[1], contact: row[2] };
  }

  function nowStr() {
    const d = new Date();
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  global.RoiCore = {
    CONF_LEVELS,
    getConfig, setConfig,
    ensureAllSheets, getMasterItems, getCategories,
    listCaseIds,
    appendHearingLog, listHearingLogs,
    callExtractionWebhook, runExtractionForReview, applyReviewedItems,
    quickCreateProposal, autoExtractProposals, getSourceEntriesForCategory,
    getProposalSummaryForCase, getCalcRowsForCase, toggleSelection,
    getCustomerInfo, getSolutions,
  };

})(window);
