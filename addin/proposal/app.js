/* ============================================================
 * 提案ナレッジ アドイン app.js
 * ------------------------------------------------------------
 * 既存の「営業報告」アドインとは別アドイン（別タスクペイン）。
 * 同じワークブックに、以下のシートを「無ければ自動作成」する。
 * 既存シート（営業報告・wbs・課題・顧客マスタ・体制・確度設定 等）は
 * 一切変更しない。
 *
 * 【ROIマスタ】  課題カテゴリ・項目ID・項目名・区分(入力/出力)・単位・
 *                デフォルト値・数式（項目ID同士のトークン式）・信頼度初期値
 *                → 運用側がここを直接編集して計算ロジックを調整する。
 * 【議事録】      案件ID・課題カテゴリ・発言者・発言テキスト・登録日時
 * 【ROI試算】     案件ID・課題カテゴリ・項目ID・項目名・区分・値・単位・
 *                信頼度・選択・更新日時
 *                → 出力行の「値」は Excel の数式（=INDEX/MATCH）で
 *                  同じ案件内の入力行を参照して自動計算される。
 *                  計算ロジックはこのアドインのJSではなく、あくまで
 *                  ROIマスタの数式列とExcelの数式評価に持たせる。
 * 【ソリューションDB】 課題カテゴリ・解決策名・概要
 *
 * 案件ID は既存の「営業報告」シートのID列（顧客コード-通番）と
 * 揃える前提。案件ID入力欄の候補（datalist）は営業報告シートの
 * ID列と顧客マスタから作る。
 * ============================================================ */

const APP_VERSION = "roi_knowledge_rev1";

const EIGYO_SHEET = "営業報告";     // 既存シート（読み取りのみ）
const CUST_SHEET = "顧客マスタ";    // 既存シート（読み取りのみ）

const MASTER_SHEET = "ROIマスタ";
const MASTER_COLUMNS = ["課題カテゴリ", "項目ID", "項目名", "区分", "単位", "デフォルト値", "数式", "信頼度初期値"];

const HEARING_SHEET = "議事録";
const HEARING_COLUMNS = ["案件ID", "課題カテゴリ", "発言者", "発言テキスト", "登録日時"];

const CALC_SHEET = "ROI試算";
const CALC_COLUMNS = ["案件ID", "課題カテゴリ", "項目ID", "項目名", "区分", "値", "単位", "信頼度", "選択", "更新日時"];

const SOLUTION_SHEET = "ソリューションDB";
const SOLUTION_COLUMNS = ["課題カテゴリ", "解決策名", "概要"];

const CONF_LEVELS = ["確定", "推定", "未確認"];

/* ---------- ROIマスタの初期シード（運用側は保存後にシート上で自由に追加・編集する） ---------- */
const MASTER_SEED = [
  // 在庫管理
  ["在庫管理", "stk_people", "棚卸人数", "入力", "人", 3, "", "未確認"],
  ["在庫管理", "stk_hours", "棚卸時間", "入力", "時間", 4, "", "未確認"],
  ["在庫管理", "stk_freq", "棚卸回数/年", "入力", "回", 12, "", "未確認"],
  ["在庫管理", "stk_wage", "時給", "入力", "円", 3000, "", "推定"],
  ["在庫管理", "stk_improve", "改善率", "入力", "%", 60, "", "推定"],
  ["在庫管理", "stk_hours_yr", "年間棚卸工数", "出力", "時間", "", "stk_people*stk_hours*stk_freq", ""],
  ["在庫管理", "stk_cost_yr", "年間棚卸コスト", "出力", "円", "", "stk_hours_yr*stk_wage", ""],
  ["在庫管理", "stk_saving", "削減額", "出力", "円", "", "stk_cost_yr*stk_improve/100", ""],
  // ロット管理
  ["ロット管理", "lot_hours", "追跡時間", "入力", "時間", 2, "", "未確認"],
  ["ロット管理", "lot_freq", "追跡回数/年", "入力", "回", 100, "", "未確認"],
  ["ロット管理", "lot_people", "担当人数", "入力", "人", 2, "", "未確認"],
  ["ロット管理", "lot_wage", "時給", "入力", "円", 3000, "", "推定"],
  ["ロット管理", "lot_improve", "改善率", "入力", "%", 60, "", "推定"],
  ["ロット管理", "lot_hours_yr", "年間追跡工数", "出力", "時間", "", "lot_hours*lot_freq*lot_people", ""],
  ["ロット管理", "lot_saving", "削減額", "出力", "円", "", "lot_hours_yr*lot_wage*lot_improve/100", ""],
  // AI議事録
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

/* ---------- 状態 ---------- */
let demoMode = false;
let masterItems = [];   // ROIマスタ全行
let caseIdOptions = [];
let currentReview = null; // AI抽出のレビュー中データ

/* ============================================================
   起動
   ============================================================ */
if (window.Office) {
  Office.onReady(() => whenDomReady(init));
} else {
  window.addEventListener("DOMContentLoaded", () => init());
}

function whenDomReady(fn) {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", fn, { once: true });
  } else {
    fn();
  }
}

async function init() {
  bindStaticUI();
  await loadAll();
  renderCategorySelect();
  renderCaseIdList();
}

function bindStaticUI() {
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });
  document.getElementById("reload-btn").addEventListener("click", () => location.reload());
  document.getElementById("settings-btn").addEventListener("click", openSettings);
  document.getElementById("cfg-close-btn").addEventListener("click", closeSettings);
  document.getElementById("cfg-save-btn").addEventListener("click", saveSettings);

  document.getElementById("save-log-btn").addEventListener("click", saveHearingLog);
  document.getElementById("extract-btn").addEventListener("click", runExtraction);
  document.getElementById("apply-extract-btn").addEventListener("click", applyExtraction);
  document.getElementById("cancel-extract-btn").addEventListener("click", cancelExtraction);
  document.getElementById("case-id").addEventListener("change", loadHearingLogForCase);

  document.getElementById("load-proposal-btn").addEventListener("click", loadProposal);
  document.getElementById("build-prompt-btn").addEventListener("click", buildPrompt);
}

function switchTab(tab) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".pane").forEach(p => p.classList.remove("active"));
  document.getElementById("pane-" + tab).classList.add("active");
  // 3タブ間で案件IDを引き継ぐ
  const cur = document.getElementById("case-id").value;
  ["case-id-2", "case-id-3"].forEach(id => { document.getElementById(id).value = cur; });
}

/* ============================================================
   設定（AIエンドポイント）
   localStorage にはエンドポインURLと簡易トークンのみ保存する。
   APIキー本体は保存しない（サーバ側にのみ保持する）。
   ============================================================ */
function openSettings() {
  const cfg = getConfig();
  document.getElementById("cfg-webhook").value = cfg.webhookUrl || "";
  document.getElementById("cfg-token").value = cfg.token || "";
  document.getElementById("settings-modal").style.display = "flex";
}
function closeSettings() { document.getElementById("settings-modal").style.display = "none"; }
function saveSettings() {
  const cfg = {
    webhookUrl: document.getElementById("cfg-webhook").value.trim(),
    token: document.getElementById("cfg-token").value.trim(),
  };
  localStorage.setItem("roiAddinConfig", JSON.stringify(cfg));
  closeSettings();
}
function getConfig() {
  try { return JSON.parse(localStorage.getItem("roiAddinConfig") || "{}"); }
  catch (e) { return {}; }
}

/* ============================================================
   Excel 読み書き：シートの自動作成
   ============================================================ */
async function loadAll() {
  if (!window.Office || !window.Excel) {
    demoMode = true;
    document.getElementById("demo-badge").style.display = "";
    seedDemoMaster();
    return;
  }
  try {
    await ensureMasterSheet();
    await ensureHearingSheet();
    await ensureCalcSheet();
    await ensureSolutionSheet();
    await loadCaseIdCandidates();
    demoMode = false;
  } catch (e) {
    console.warn("Excel読込に失敗。デモモードで起動します。", e);
    demoMode = true;
    seedDemoMaster();
  }
  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
}

function seedDemoMaster() {
  masterItems = MASTER_SEED.map(r => rowToMasterItem(r));
  caseIdOptions = ["KM-01", "OF-02"];
}

function rowToMasterItem(r) {
  return { category: r[0], itemId: r[1], name: r[2], kind: r[3], unit: r[4], defaultVal: r[5], formula: r[6], confDefault: r[7] };
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
    if (seedRows && seedRows.length) {
      ws.getRange(`A2:${lastCol}${seedRows.length + 1}`).values = seedRows;
    }
    await ctx.sync();
  }
  return ws;
}

function colLetterOf(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function ensureMasterSheet() {
  await Excel.run(async ctx => {
    await getOrCreateSheet(ctx, MASTER_SHEET, MASTER_COLUMNS, MASTER_SEED);
    const rng = ctx.workbook.worksheets.getItem(MASTER_SHEET).getUsedRange(true);
    rng.load("values");
    await ctx.sync();
    const rows = rng.values.slice(1); // ヘッダー除く
    masterItems = rows.filter(r => r[1]).map(rowToMasterItem);
  });
}

async function ensureHearingSheet() {
  await Excel.run(async ctx => { await getOrCreateSheet(ctx, HEARING_SHEET, HEARING_COLUMNS, null); });
}

async function ensureCalcSheet() {
  await Excel.run(async ctx => { await getOrCreateSheet(ctx, CALC_SHEET, CALC_COLUMNS, null); });
}

async function ensureSolutionSheet() {
  await Excel.run(async ctx => { await getOrCreateSheet(ctx, SOLUTION_SHEET, SOLUTION_COLUMNS, SOLUTION_SEED); });
}

/* ---------- 案件IDの候補（営業報告シートのID列＋顧客マスタの顧客コード） ---------- */
async function loadCaseIdCandidates() {
  const ids = new Set();
  await Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    if (sheets.items.find(s => s.name === EIGYO_SHEET)) {
      const sheet = ctx.workbook.worksheets.getItem(EIGYO_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const lastRow = Math.min(Math.max(used.rowCount, 1), 1000);
      if (lastRow >= 2) {
        const rng = sheet.getRange(`A2:A${lastRow}`);
        rng.load("values");
        await ctx.sync();
        rng.values.forEach(r => { if (r[0]) ids.add(String(r[0])); });
      }
    }
  });
  caseIdOptions = Array.from(ids);
}

function renderCaseIdList() {
  const dl = document.getElementById("case-id-list");
  dl.innerHTML = caseIdOptions.map(id => `<option value="${escAttr(id)}">`).join("");
}

function renderCategorySelect() {
  const cats = Array.from(new Set(masterItems.map(m => m.category)));
  const sel = document.getElementById("category-select");
  sel.innerHTML = cats.map(c => `<option value="${escAttr(c)}">${escHtml(c)}</option>`).join("");
}

/* ============================================================
   ① 議事録入力
   ============================================================ */
async function saveHearingLog() {
  const caseId = document.getElementById("case-id").value.trim();
  const category = document.getElementById("category-select").value;
  const speaker = document.getElementById("speaker-select").value;
  const text = document.getElementById("hearing-text").value.trim();
  if (!caseId) { setStatus("案件IDを入力してください"); return; }
  if (!text) { setStatus("発言・メモを入力してください"); return; }

  const row = [caseId, category, speaker, text, nowStr()];
  if (demoMode) {
    setStatus("デモモードのため保存はシミュレーションのみです");
  } else {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const nextRow = Math.max(used.rowCount, 1) + 1;
      sheet.getRange(`A${nextRow}:E${nextRow}`).values = [row];
      await ctx.sync();
    });
  }
  document.getElementById("hearing-text").value = "";
  setStatus("議事録に記録しました");
  await loadHearingLogForCase();
}

async function loadHearingLogForCase() {
  const caseId = document.getElementById("case-id").value.trim();
  const list = document.getElementById("hearing-log-list");
  if (!caseId) { list.innerHTML = ""; return; }
  let rows = [];
  if (demoMode) {
    rows = [];
  } else {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(HEARING_SHEET);
      const rng = sheet.getUsedRange(true);
      rng.load("values");
      await ctx.sync();
      rows = rng.values.slice(1).filter(r => String(r[0]) === caseId);
    });
  }
  list.innerHTML = rows.map(r => `
    <div class="log-item">
      <div class="meta">${escHtml(r[1] || "")} ／ ${escHtml(r[2] || "")} ／ ${escHtml(r[4] || "")}</div>
      <div class="text">${escHtml(r[3] || "")}</div>
    </div>`).join("") || `<div class="meta">まだ記録がありません</div>`;
}

/* ---------- AI抽出（GAS等のWebhookを叩き、項目ID・値・信頼度のJSON配列を受け取る） ---------- */
async function runExtraction() {
  const caseId = document.getElementById("case-id").value.trim();
  const category = document.getElementById("category-select").value;
  const text = document.getElementById("hearing-text").value.trim();
  if (!caseId) { setStatus("案件IDを入力してください"); return; }
  if (!text) { setStatus("発言・メモを入力してください（記録前でも抽出だけ試せます）"); return; }

  const cfg = getConfig();
  if (!cfg.webhookUrl) { setStatus("設定（⚙）でAI連携エンドポイントを登録してください"); return; }

  const items = masterItems.filter(m => m.category === category && m.kind === "入力");
  setStatus("AIに抽出を依頼中…");
  try {
    const res = await fetch(cfg.webhookUrl, {
      method: "POST",
      body: JSON.stringify({
        token: cfg.token || "",
        caseId, category,
        text,
        items: items.map(i => ({ itemId: i.itemId, name: i.name, unit: i.unit })),
      }),
    });
    const raw = await res.text();
    const data = JSON.parse(raw);
    // 期待するレスポンス: { items: [{ itemId, value, confidence }] }
    currentReview = { caseId, category, items: data.items || [] };
    renderExtractReview();
    setStatus("");
  } catch (e) {
    console.warn(e);
    setStatus("抽出に失敗しました。エンドポイントの設定を確認してください");
  }
}

function renderExtractReview() {
  const box = document.getElementById("extract-review");
  const wrap = document.getElementById("extract-items");
  if (!currentReview || !currentReview.items.length) {
    box.style.display = "none";
    return;
  }
  const master = masterItems.filter(m => m.category === currentReview.category && m.kind === "入力");
  wrap.innerHTML = currentReview.items.map((it, idx) => {
    const def = master.find(m => m.itemId === it.itemId);
    return `
    <div class="review-item" data-idx="${idx}">
      <div class="item-head"><span>${escHtml(def ? def.name : it.itemId)}</span><span>${escHtml(def ? def.unit : "")}</span></div>
      <div class="item-row">
        <input type="text" class="rv-value" value="${escAttr(it.value ?? "")}">
        <select class="rv-conf">
          ${CONF_LEVELS.map(c => `<option value="${c}" ${c === it.confidence ? "selected" : ""}>${c}</option>`).join("")}
        </select>
      </div>
    </div>`;
  }).join("");
  box.style.display = "";
}

function cancelExtraction() {
  currentReview = null;
  document.getElementById("extract-review").style.display = "none";
}

async function applyExtraction() {
  if (!currentReview) return;
  // レビューUIの編集値を反映
  document.querySelectorAll("#extract-items .review-item").forEach(el => {
    const idx = Number(el.dataset.idx);
    currentReview.items[idx].value = el.querySelector(".rv-value").value;
    currentReview.items[idx].confidence = el.querySelector(".rv-conf").value;
  });
  await applyCategoryToCalcSheet(currentReview.caseId, currentReview.category, currentReview.items);
  setStatus("ROI試算シートに反映しました");
  cancelExtraction();
}

/* ============================================================
   ROI試算シートへの反映
   1. 対象カテゴリの全項目（入力・出力）が、この案件の行として
      まだ無ければ ROIマスタから行をコピーする（出力行は数式化）。
   2. 渡された入力値・信頼度を、該当行の「値」「信頼度」に上書きする。
   ============================================================ */
async function applyCategoryToCalcSheet(caseId, category, extractedItems) {
  if (demoMode) return; // デモモードでは書き込みをスキップ

  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
    const used = sheet.getUsedRange(true);
    used.load("values, rowCount");
    await ctx.sync();

    const existingRows = used.values.slice(1); // ヘッダー除く
    const existingKey = (r) => `${r[0]}__${r[2]}`; // 案件ID__項目ID
    const existingIndex = {};
    existingRows.forEach((r, i) => { existingIndex[existingKey(r)] = i + 2; }); // シート上の行番号（ヘッダー=1行目）

    const catItems = masterItems.filter(m => m.category === category);
    const toAppend = [];
    const nowTs = nowStr();

    catItems.forEach(m => {
      const key = `${caseId}__${m.itemId}`;
      if (existingIndex[key]) return; // 既にこの案件の行がある → 新規追加はしない
      const rowNum = used.rowCount + 1 + toAppend.length; // 追加予定行の見込み
      let valueCell;
      if (m.kind === "出力") {
        valueCell = translateFormula(m.formula, rowNum);
      } else {
        const ex = extractedItems.find(e => e.itemId === m.itemId);
        valueCell = ex ? ex.value : m.defaultVal;
      }
      const confCell = m.kind === "出力" ? "" : (extractedItems.find(e => e.itemId === m.itemId)?.confidence || m.confDefault);
      toAppend.push([caseId, m.category, m.itemId, m.name, m.kind, valueCell, m.unit, confCell, "FALSE", nowTs]);
    });

    if (toAppend.length) {
      const startRow = used.rowCount + 1;
      const endRow = startRow + toAppend.length - 1;
      // F列（値）は数式混在のため formulas で、その他は values で書く
      const rangeAll = sheet.getRange(`A${startRow}:J${endRow}`);
      rangeAll.values = toAppend.map(r => r.map((v, i) => (i === 5 && typeof v === "string" && v.startsWith("=")) ? "" : v));
      const formulaCol = sheet.getRange(`F${startRow}:F${endRow}`);
      formulaCol.formulas = toAppend.map(r => [typeof r[5] === "string" && r[5].startsWith("=") ? r[5] : (r[5] === "" ? "" : r[5])]);
      await ctx.sync();
    }

    // 既存行への値・信頼度の上書き（入力項目のみ）
    const updates = [];
    extractedItems.forEach(e => {
      const key = `${caseId}__${e.itemId}`;
      const rowNum = existingIndex[key];
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

/* ---------- ROIマスタの数式（項目IDのトークン式）をExcel数式へ変換 ----------
 * 例: "stk_people*stk_hours*stk_freq" →
 *   =INDEX($F$2:$F$9999,MATCH(1,($A$2:$A$9999=$A{row})*($C$2:$C$9999="stk_people"),0))
 *    * 同様に stk_hours, stk_freq も INDEX/MATCH に置換
 * 「同じ案件ID・その項目IDの行」をROI試算シート内から検索して値を引く。
 * 計算そのものはExcelの数式評価に委ねる（JS側では計算しない）。 */
function translateFormula(masterFormula, rowNum) {
  if (!masterFormula) return "";
  const translated = masterFormula.replace(/[a-zA-Z_][a-zA-Z0-9_]*/g, (tok) => {
    return `INDEX($F$2:$F$9999,MATCH(1,($A$2:$A$9999=$A${rowNum})*($C$2:$C$9999="${tok}"),0))`;
  });
  return "=" + translated;
}

/* ============================================================
   ② 提案（課題×解決案）
   ============================================================ */
async function loadProposal() {
  const caseId = document.getElementById("case-id-2").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  let calcRows = [];
  let solutions = [];
  if (!demoMode) {
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
  } else {
    solutions = SOLUTION_SEED;
  }

  const byCategory = {};
  calcRows.forEach(r => {
    const cat = r[1];
    byCategory[cat] = byCategory[cat] || { inputs: [], outputs: [] };
    (r[4] === "出力" ? byCategory[cat].outputs : byCategory[cat].inputs).push(r);
  });

  const list = document.getElementById("proposal-list");
  const cats = Object.keys(byCategory);
  if (!cats.length) {
    list.innerHTML = `<div class="meta">この案件のROI試算データがまだありません。①でAI抽出を行ってください。</div>`;
    document.getElementById("proposal-summary").style.display = "none";
    return;
  }

  list.innerHTML = cats.map(cat => {
    const sol = solutions.find(s => s[0] === cat);
    const savingRow = byCategory[cat].outputs.find(r => String(r[2]).endsWith("_saving")) || byCategory[cat].outputs[0];
    const saving = savingRow ? savingRow[5] : "";
    const conf = byCategory[cat].inputs.some(r => r[7] === "未確認") ? "未確認"
      : byCategory[cat].inputs.some(r => r[7] === "推定") ? "推定" : "確定";
    const badgeClass = conf === "確定" ? "badge-confirmed" : conf === "推定" ? "badge-estimated" : "badge-unknown";
    const checked = byCategory[cat].outputs.concat(byCategory[cat].inputs).some(r => String(r[8]).toUpperCase() === "TRUE");
    return `
    <div class="proposal-card" data-cat="${escAttr(cat)}">
      <div class="p-head">
        <input type="checkbox" class="p-check" ${checked ? "checked" : ""}>
        <div>
          <div class="p-title">${escHtml(cat)}</div>
          <div class="p-sub">${escHtml((byCategory[cat].inputs[0] || [])[3] || "")}</div>
        </div>
      </div>
      ${sol ? `<div class="p-solution">→ 解決策：${escHtml(sol[1])}</div>` : `<div class="p-solution">解決策が未登録です（ソリューションDBに追加してください）</div>`}
      <div class="p-metrics">
        <span>削減額 <b>${saving !== "" ? fmtNum(saving) + "円/年" : "—"}</b></span>
        <span class="badge ${badgeClass}">${conf}</span>
      </div>
    </div>`;
  }).join("");

  list.querySelectorAll(".p-check").forEach(cb => {
    cb.addEventListener("change", async (e) => {
      const cat = e.target.closest(".proposal-card").dataset.cat;
      await toggleSelection(caseId, cat, e.target.checked);
    });
  });

  const summary = document.getElementById("proposal-summary");
  summary.style.display = "";
  const selectedCount = list.querySelectorAll(".p-check:checked").length;
  document.getElementById("proposal-summary-text").textContent = `選択中 ${selectedCount}件`;
}

async function toggleSelection(caseId, category, checked) {
  if (demoMode) return;
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
    const rng = sheet.getUsedRange(true);
    rng.load("values");
    await ctx.sync();
    const rows = rng.values;
    rows.forEach((r, i) => {
      if (i === 0) return;
      if (String(r[0]) === caseId && r[1] === category) {
        sheet.getRange(`I${i + 1}`).values = [[checked ? "TRUE" : "FALSE"]];
      }
    });
    await ctx.sync();
  });
}

/* ============================================================
   ③ プロンプト出力
   ============================================================ */
async function buildPrompt() {
  const caseId = document.getElementById("case-id-3").value.trim();
  document.getElementById("case-id").value = caseId;
  if (!caseId) return;

  let calcRows = [], solutions = [], custRow = null;
  if (!demoMode) {
    await Excel.run(async ctx => {
      const calcSheet = ctx.workbook.worksheets.getItem(CALC_SHEET);
      const r1 = calcSheet.getUsedRange(true);
      r1.load("values");
      const solSheet = ctx.workbook.worksheets.getItem(SOLUTION_SHEET);
      const r2 = solSheet.getUsedRange(true);
      r2.load("values");
      const custSheet = ctx.workbook.worksheets.getItem(CUST_SHEET);
      const r3 = custSheet.getUsedRange(true);
      r3.load("values");
      await ctx.sync();
      calcRows = r1.values.slice(1).filter(r => String(r[0]) === caseId && String(r[8]).toUpperCase() === "TRUE");
      solutions = r2.values.slice(1);
      const code = caseId.split("-")[0];
      custRow = r3.values.slice(1).find(r => r[0] === code) || null;
    });
  }

  if (!calcRows.length) {
    document.getElementById("prompt-output").innerHTML = `<div class="meta">②で提案書に含める課題を選択（チェック）してください。</div>`;
    return;
  }

  const byCategory = {};
  calcRows.forEach(r => {
    byCategory[r[1]] = byCategory[r[1]] || { inputs: [], outputs: [] };
    (r[4] === "出力" ? byCategory[r[1]].outputs : byCategory[r[1]].inputs).push(r);
  });
  const cats = Object.keys(byCategory);

  const custInputText = custRow
    ? `取引先：${custRow[1] || ""}\n窓口：${custRow[2] || ""}`
    : `案件ID：${caseId}`;

  const roiText = cats.map(cat => {
    const lines = byCategory[cat].outputs.map(r => `${r[3]} ${fmtNum(r[5])}${r[6] || ""}`).join("\n");
    return `[${cat}]\n${lines}`;
  }).join("\n\n");

  const outlineText = cats.map((cat, i) => {
    const sol = solutions.find(s => s[0] === cat);
    return `${i + 1}. ${cat}\n解決の型：${sol ? sol[1] : "（未設定）"}\n${sol ? sol[2] : ""}`;
  }).join("\n\n");

  const promptText =
`中小企業向けの提案書を、現状課題→解決の型→ROI→費用・体制の4章構成で作成してください。
トーンは平易で、数値の根拠（信頼度）を明示してください。`;

  const blocks = [
    { label: "生成プロンプト", body: promptText },
    { label: "顧客インプット", body: custInputText },
    { label: "ROI数値", body: roiText },
    { label: "章立て骨子", body: outlineText },
  ];

  const out = document.getElementById("prompt-output");
  out.innerHTML = blocks.map((b, i) => `
    <div class="prompt-block" data-idx="${i}">
      <div class="pb-head">
        <span class="pb-label">${escHtml(b.label)}</span>
        <button class="btn pb-copy">コピー</button>
      </div>
      <div class="pb-body">${escHtml(b.body)}</div>
    </div>`).join("");

  out.querySelectorAll(".pb-copy").forEach((btn, i) => {
    btn.addEventListener("click", () => copyToClipboard(blocks[i].body));
  });
}

function copyToClipboard(text) {
  if (navigator.clipboard) navigator.clipboard.writeText(text);
}

/* ============================================================
   ユーティリティ
   ============================================================ */
function setStatus(msg) { document.getElementById("extract-status").textContent = msg; }
function nowStr() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function fmtNum(v) {
  const n = Number(v);
  return isNaN(n) ? String(v) : n.toLocaleString("ja-JP");
}
function escHtml(s) {
  return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function escAttr(s) { return escHtml(s); }
