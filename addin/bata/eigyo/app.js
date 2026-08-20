/* ============================================================
 * 営業報告アドイン app.js（ステージタブ版 rev_e）
 * ------------------------------------------------------------
 * 対象シート: 「営業報告」（1案件1行、ヘッダー行=1行目）
 * カラム定義は SHEET_COLUMNS（列レター・見出し・用途）に一元化。
 * 起動時にヘッダー行（A1:AI1）を照合し、見出しが無い／異なる列が
 * あれば自動で正規の見出しに書き直す（列の並び順・位置は不変）。
 *
 * 【基本情報】       A:ID  B:取引先  C:No(未使用)  D:種別  E:状態
 *                    F:発生日  G:完了日  H:担当者  I:窓口  J:優先度
 * 【見積・受注】     K:見積工数  L:見積金額  M:受注区分  N:納品日
 * 【内容・メモ】     O:問合せ・提案内容  P:進捗状況  Q:備考  R:(未使用)
 * 【ステージ詳細】   S:区分  T:着手日  U:見積根拠  V:商談状況
 *                    W:確認状況  X:計上日  Y:最終工数  Z:最終価格  AA:受注条件
 * 【管理・完了日】   AB:起票者  AC:見積完了日  AD:検討完了日
 *                    AE:商談完了日  AF:確認完了日
 * 【更新管理】       AJ:最終更新日（保存の都度、自動打刻）
 * 【追加】           AK:見積有効期限（見積完了時は必須）
 *                    AL:保留（TRUE/空。状態(E)は本来の進捗を保持したまま保留を表す）
 * 【工数（改修）】   AM:対応工数（人日）　AN:受託工数（人日）
 *   ・保守対応／瑕疵対応／調整の「対応工数」は、旧K列（見積工数）の流用をやめ
 *     AM列を専用に使用する（K列は見積り／プリセールスの見積工数のまま）。
 *   ・受託工数（AN）は見積り／プリセールスの「受託中」ステージでのみ使用する。
 * 顧客マスタ: 「顧客マスタ」シート（無ければ自動作成）
 * ============================================================ */

const APP_VERSION = "rev_20260819_b3f7a01";
const SHEET_NAME = "営業報告";
const CUST_SHEET = "顧客マスタ";
const MAX_ROWS = 500;
const TAX_RATE = 0.10;

/* カンバンのドラッグ＆ドロップ制御（true にすると有効化） */
const ENABLE_KANBAN_DND = false;

/* ---------- ワークフロー定義 ----------
 * 見積り／プリセールスは 受注 の後に「受託中」（実際の作業実施期間）を経て「完了」となる。
 * 失注は確認中タブの選択で即座に確定するため、チェーン上の位置に関わらず特別扱い（saveEditRecord参照）。 */
const WORKFLOWS = {
  "保守対応":     { steps: ["新規", "対応中"],                              terminals: ["完了"] },
  "瑕疵対応":     { steps: ["新規", "対応中"],                              terminals: ["完了"] },
  "見積り":       { steps: ["新規", "見積中", "確認中", "受注", "受託中"],   terminals: ["完了", "失注"] },
  "プリセールス": { steps: ["新規", "検討中", "商談中", "確認中", "受注", "受託中"], terminals: ["完了", "失注"] },
  "調整":         { steps: ["新規", "対応中"],                              terminals: ["完了"] },
};
const TYPES = Object.keys(WORKFLOWS);
const HOLD = "保留";
const QUOTE_TYPES = ["見積り", "プリセールス"];
/* 受注確定済み（以降のステージも含む）とみなす状態 */
const ORDER_CONFIRMED_STATUSES = ["受注", "受託中", "完了"];

/* ---------- 保留の扱い ----------
 * 保留は「状態」ではなく独立したフラグ（AL列）で持つ。
 * 状態(E列)には常に本来の進捗（見積中・商談中など）を残すため、
 * 保留を解除すればそのステージからそのまま再開できる。
 * 表示・集計・カンバンのレーン分けは effectiveStatus() に一元化する。 */
function isHold(rec) { return !!(rec && rec.hold); }
function effectiveStatus(rec) { return isHold(rec) ? HOLD : rec.status; }

function stageTabsOf(type) {
  if (type === "見積り") return ["起票", "見積中", "確認中", "受注", "受託中"];
  if (type === "プリセールス") return ["起票", "検討中", "商談中", "確認中", "受注", "受託中"];
  return ["起票", "対応中"];
}
function firstStageOf(type) {
  if (type === "見積り") return "見積中";
  if (type === "プリセールス") return "検討中";
  return "対応中";
}

const LEGACY_STATUS = {
  "未着手": "新規",
  "作成中": "検討中",
  "見積作成中": "見積中", "見積提出済み": "確認中",
  "調整中": "対応中",
  "完了(受注)": "受注", "完了(失注)": "失注",
};

/* 正規の列見出し（A〜AL、38列）。既存列の並び順・位置は不変（末尾に2列追加）。 */
const SHEET_COLUMNS = [
  "ID", "取引先", "No（未使用）", "種別", "状態",
  "発生日", "完了日", "担当者", "窓口", "優先度",
  "見積工数（人日）", "見積金額（税抜）", "受注区分", "納品日",
  "問合せ・提案内容", "進捗状況", "備考", "（未使用）",
  "区分（問合せ／改修）", "着手日", "見積根拠", "商談状況", "確認状況",
  "計上日", "最終工数（人日）", "最終価格（税抜）", "受注条件",
  "起票者", "見積完了日", "検討完了日", "商談完了日", "確認完了日",
  "受注確定日", "受託開始日", "完了予定日", "最終更新日",
  "見積有効期限", "保留",
  "対応工数（人日）", "受託工数（人日）",
];
/* シート範囲の最終列（列を増やすときはここだけ変える） */
const LAST_COL = "AN";
/* 旧バージョンで使っていた見出し文言（読み込み時の判定に使用、書込みはしない） */
const EXT_HEADERS = SHEET_COLUMNS.slice(18); // S列以降（互換維持用）


/* ---------- 期（会計年度）：10月〜翌9月、第37期=2025/10〜2026/09 ---------- */
function termOfDate(d) { return d.getMonth() + 1 >= 10 ? d.getFullYear() - 1988 : d.getFullYear() - 1989; }
function fiscalMonths(term) {
  const sy = term + 1988;
  const out = [];
  for (let i = 0; i < 12; i++) {
    const m = 10 + i;
    out.push(monthKeyYM(m > 12 ? sy + 1 : sy, m > 12 ? m - 12 : m));
  }
  return out;
}
function termLabel(term) {
  const sy = term + 1988;
  return `第${term}期（${sy}/10〜${sy + 1}/09）`;
}
let currentTerm = termOfDate(new Date());

/* ---------- 状態 ---------- */
let records = [];
let customers = [];
let demoMode = false;
let editingRec = null;
let currentStageTab = null;
let inputType = "保守対応";
let currentKanbanType = "保守対応";
let dragId = null;
let filters = { q: "", status: [], client: [], owner: "", lastWeekOnly: false, thisWeekOnly: false };
let editDirty = false;      // 詳細画面で変更があったか
let selectedId = null;      // 一覧で選択中の案件ID（ハイライト用）
/* 一覧: 種別グループの開閉状態（種別名の集合、閉じているものだけ保持） */
let collapsedTypes = new Set();
try { collapsedTypes = new Set(JSON.parse(localStorage.getItem("eigyo_collapsed_types") || "[]")); } catch (e) {}

/* ---------- Cookie 保存/復元 ---------- */
const COOKIE_DAYS = 365;
function setCookie(name, value) {
  const exp = new Date(Date.now() + COOKIE_DAYS * 86400000).toUTCString();
  document.cookie = `${name}=${encodeURIComponent(value)};expires=${exp};path=/;SameSite=Lax`;
}
function getCookie(name) {
  const m = document.cookie.match(new RegExp("(?:^|; )" + name + "=([^;]*)"));
  return m ? decodeURIComponent(m[1]) : null;
}
function saveFiltersCookie() {
  try { localStorage.setItem("eigyo_filters", JSON.stringify(filters)); } catch (e) {}
}
function restoreFiltersCookie() {
  try {
    const raw = localStorage.getItem("eigyo_filters");
    if (!raw) return;
    const f = JSON.parse(raw);
    filters = {
      q: f.q || "",
      status: Array.isArray(f.status) ? f.status : (f.status ? [f.status] : []),
      client: Array.isArray(f.client) ? f.client : (f.client ? [f.client] : []),
      owner: f.owner || "",
      lastWeekOnly: !!f.lastWeekOnly,
      thisWeekOnly: !!f.thisWeekOnly,
    };
  } catch (e) {}
}
function saveCollapsedTypes() {
  try { localStorage.setItem("eigyo_collapsed_types", JSON.stringify([...collapsedTypes])); } catch (e) {}
}
function toggleTypeGroup(type) {
  if (collapsedTypes.has(type)) collapsedTypes.delete(type); else collapsedTypes.add(type);
  saveCollapsedTypes();
  renderList();
}

/* ---------- 前週実績（先週実績）判定 ---------- */
function lastWeekRange() {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const day = now.getDay(); // 0=日
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const thisMonday = new Date(now); thisMonday.setDate(now.getDate() + diffToMonday);
  const lastMonday = new Date(thisMonday); lastMonday.setDate(thisMonday.getDate() - 7);
  const lastSunday = new Date(thisMonday); lastSunday.setDate(thisMonday.getDate() - 1);
  lastSunday.setHours(23, 59, 59, 999);
  return { start: lastMonday, end: lastSunday };
}
function isLastWeekUpdate(rec) {
  if (!rec.lastUpdate) return false;
  const { start, end } = lastWeekRange();
  return rec.lastUpdate >= start && rec.lastUpdate <= end;
}
/* ---------- 今週実績判定 ---------- */
function thisWeekRange() {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const day = now.getDay(); // 0=日
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const monday = new Date(now); monday.setDate(now.getDate() + diffToMonday);
  const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);
  return { start: monday, end: sunday };
}
function isThisWeekUpdate(rec) {
  if (!rec.lastUpdate) return false;
  const { start, end } = thisWeekRange();
  return rec.lastUpdate >= start && rec.lastUpdate <= end;
}

/* 先週実績・今週実績は排他（両方ONにしても該当0件になるだけなので、
   一方をONにしたら他方を自動でOFFにする） */
function toggleLastWeek() {
  filters.lastWeekOnly = !filters.lastWeekOnly;
  if (filters.lastWeekOnly) filters.thisWeekOnly = false;
  saveFiltersCookie();
  updateWeekBtns();
  renderCurrentPane();
}
function toggleThisWeek() {
  filters.thisWeekOnly = !filters.thisWeekOnly;
  if (filters.thisWeekOnly) filters.lastWeekOnly = false;
  saveFiltersCookie();
  updateWeekBtns();
  renderCurrentPane();
}
function updateWeekBtns() {
  const lw = document.getElementById("btn-lastweek");
  if (lw) lw.classList.toggle("active", !!filters.lastWeekOnly);
  const tw = document.getElementById("btn-thisweek");
  if (tw) tw.classList.toggle("active", !!filters.thisWeekOnly);
}
/* 旧名（他から呼ばれていた場合の互換） */
function updateLastWeekBtn() { updateWeekBtns(); }

/* 実績フィルタで「対象外」として薄く表示するか */
function isDimmedByWeek(rec) {
  if (filters.lastWeekOnly) return !isLastWeekUpdate(rec);
  if (filters.thisWeekOnly) return !isThisWeekUpdate(rec);
  return false;
}
function saveGanttCookie() {
  try {
    localStorage.setItem("eigyo_gantt",
      JSON.stringify({ zoom: ganttZoom, term: ganttTerm, hideDone: ganttHideDone }));
  } catch (e) {}
}
function restoreGanttCookie() {
  try {
    const raw = localStorage.getItem("eigyo_gantt");
    if (!raw) return;
    const g = JSON.parse(raw);
    if ([12, 6, 3, 1].includes(g.zoom)) ganttZoom = g.zoom;
    if (Number.isInteger(g.term)) ganttTerm = g.term;
    if (typeof g.hideDone === "boolean") ganttHideDone = g.hideDone;
  } catch (e) {}
}

/* ---------- 共通スライドメニュー ---------- */
const COMMON_BASE = "https://ymatsuda-cmyk.github.io/tools/common";
let menuReady = null;
function openMenu() {
  if (!menuReady) {
    menuReady = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = COMMON_BASE + "/slide-menu.js";
      s.onload = () => {
        SlideMenu.init({
          appName: "営業報告",
          version: APP_VERSION,
          position: "left",
          currentId: "eigyo",                     // menu.json の id と一致で強調
          menuUrl: COMMON_BASE + "/menu.json",
          localItems: [
            { label: "設定をリセット", icon: "", onClick: () => resetSettings() }
          ],
        });
        resolve();
      };
      s.onerror = (e) => { menuReady = null; reject(e); };
      document.head.appendChild(s);
    });
  }
  menuReady.then(() => SlideMenu.open()).catch(() => {
    menuReady = null;
    uiAlert("メニューの読み込みに失敗しました。通信環境をご確認ください。");
  });
}

/* ============================================================
   共通モジュール（api.js）との連携設定
   ------------------------------------------------------------
   ・wbsシート名／営業報告シート名を明示
   ・タスク追加／備考編集の後、案件編集モーダルが開いていれば
     タスク一覧（ミニカンバン）と件数バッジを更新する
   ============================================================ */
window.ApiConfig = window.ApiConfig || {};
window.ApiConfig.wbsSheet = "wbs";
window.ApiConfig.eigyoSheet = SHEET_NAME;
window.ApiConfig.onTaskAdded = async () => {
  await loadWbsTaskCounts();          // 一覧・カンバンのバッジを最新化
  if (editingRec) refreshEdTaskPanel();
  renderCurrentPane();
};
window.ApiConfig.onNoteSaved = () => { if (editingRec) refreshEdTaskPanel(); };

/* ============================================================
   起動
   ============================================================ */
if (window.Office) {
  Office.onReady(() => whenDomReady(init));
} else {
  window.addEventListener("DOMContentLoaded", () => init());
}

/* ============================================================
   DOM の準備を待つ
   ------------------------------------------------------------
   Excel on the web では Office.onReady が DOM の解析より先に
   解決することがある。app.js は <head> で読み込まれるため、
   その場合 init() の 1 行目で version-label が null になり
   TypeError で初期化が丸ごと止まる（ペインが空白のまま）。
   ============================================================ */
function whenDomReady(fn) {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", fn, { once: true });
  } else {
    fn();
  }
}

async function init() {
  const vl = document.getElementById("version-label");
  if (vl) vl.textContent = APP_VERSION;
  restoreFiltersCookie();
  restoreGanttCookie();
  bindStaticUI();
  await loadAll();
  await loadWbsTaskCounts();
  const si = document.getElementById("search-input");
  if (si) si.value = filters.q || "";
  renderFilters();
  updateWeekBtns();
  renderCurrentPane();
}

function bindStaticUI() {
  // 要素が1つ欠けても以降のバインドと初期化を止めない
  const input = document.getElementById("search-input");
  if (input) {
    input.addEventListener("input", () => { filters.q = input.value.trim(); saveFiltersCookie(); renderCurrentPane(); });
  }
  const clearBtn = document.getElementById("search-clear");
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      if (input) input.value = "";
      filters.q = ""; saveFiltersCookie(); renderCurrentPane();
    });
  }
  const ownerSel = document.getElementById("filter-owner");
  if (ownerSel) {
    ownerSel.addEventListener("change", e => {
      filters.owner = e.target.value; saveFiltersCookie(); renderCurrentPane();
    });
  }
  document.addEventListener("click", e => {
    ["status", "client"].forEach(k => {
      if (!e.target.closest("#ms-" + k)) {
        const dd = document.getElementById(k + "-dd");
        if (dd) dd.style.display = "none";
      }
    });
  });
  // 詳細画面の変更検知（委譲）：ステージ切替で再描画されても効くように
  const emodal = document.getElementById("edit-modal");
  if (emodal) {
    emodal.addEventListener("input", markDirty);
    emodal.addEventListener("change", markDirty);
  }
}

function clearFilters() {
  filters = { q: "", status: [], client: [], owner: "", lastWeekOnly: false, thisWeekOnly: false };
  const si0 = document.getElementById("search-input");
  if (si0) si0.value = "";
  saveFiltersCookie();
  renderFilters();
  updateLastWeekBtn();
  renderCurrentPane();
}

/* ============================================================
   Excel 読み書き
   ============================================================ */
async function loadAll() {
  if (!window.Office || !window.Excel) {
    loadDemo();
    document.getElementById("demo-badge").style.display = "";
    return;
  }
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      const hdr = sheet.getRange(`A1:${LAST_COL}1`);
      hdr.load("values");
      await ctx.sync();
      const cur = hdr.values[0];
      const mismatch = SHEET_COLUMNS.some((h, i) => (cur[i] || "").toString().trim() !== h);
      if (mismatch) {
        hdr.values = [SHEET_COLUMNS];
        hdr.format.fill.color = "#44546A";
        hdr.format.font.color = "#FFFFFF";
        hdr.format.font.bold = true;
        await ctx.sync();
      }
      // 使用範囲の行数だけ読む（A2:AI500固定読みを避け、使用範囲の膨張を防ぐ）
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const lastRow = Math.min(Math.max(used.rowCount, 1), MAX_ROWS);
      if (lastRow >= 2) {
        const rng = sheet.getRange(`A2:${LAST_COL}${lastRow}`);
        rng.load("values");
        await ctx.sync();
        records = parseRows(rng.values);
      } else {
        records = [];
      }
    });
    await migrateHoldFlags();
    await migrateWorkHours();
    await ensureCustomerSheet();
    await loadConfidenceRates();
    demoMode = false;
  } catch (e) {
    console.warn("Excel読込に失敗。デモモードで起動します。", e);
    loadDemo();
  }
  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
}

/* ---------- 旧データ移行：状態が「保留」の行を保留フラグ(AL)へ ----------
 * 保留に入る前の状態は追跡できないため、AL列にTRUEを立てるところまでを自動で行う。
 * 状態(E列)は「保留」のまま残すので、本来の状態は手作業で入れ直す。
 * 対象行は一覧・カンバンで「状態要確認」として目印を付ける。 */
async function migrateHoldFlags() {
  const targets = records.filter(r => r.holdLegacy && !r.holdWritten && r.row);
  if (!targets.length) return;
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      targets.forEach(r => { sheet.getRange(`AL${r.row}`).values = [["TRUE"]]; });
      await ctx.sync();
    });
    targets.forEach(r => { r.holdWritten = true; });
    console.info(`保留フラグを移行しました（${targets.length}件）。状態(E列)は手作業で本来の状態に戻してください。`);
  } catch (e) {
    console.warn("保留フラグの移行に失敗しました。", e);
  }
}

/* ---------- 旧データ移行：対応工数（K列→AM列） ----------
 * 保守対応／瑕疵対応／調整は、これまで見積工数（K列）を対応工数として
 * 流用していたが、専用列（AM）へ移行する。AMが未入力の行のみK列の値を
 * コピーする（既にAMに値がある行は上書きしない＝安全な移行）。 */
async function migrateWorkHours() {
  const targets = records.filter(r =>
    ["保守対応", "瑕疵対応", "調整"].includes(r.type) &&
    r.workHours == null && r.hours != null && r.row
  );
  if (!targets.length) return;
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      targets.forEach(r => { sheet.getRange(`AM${r.row}`).values = [[r.hours]]; });
      await ctx.sync();
    });
    targets.forEach(r => { r.workHours = r.hours; });
    console.info(`対応工数（AM列）へ見積工数（K列）の値を移行しました（${targets.length}件）。`);
  } catch (e) {
    console.warn("対応工数の移行に失敗しました。", e);
  }
}

function parseRows(values) {
  const out = [];
  values.forEach((r, i) => {
    if (!r[0] && !r[14]) return;
    const rawStatus = str(r[4]);
    const holdCell = toBool(r[37]);
    const holdLegacy = rawStatus === HOLD;   // 旧データ：状態が「保留」
    out.push({
      row: i + 2,
      id: str(r[0]), client: str(r[1]), no: r[2],
      type: str(r[3]),
      status: normalizeStatus(rawStatus),
      occur: toDate(r[5]), done: toDate(r[6]),
      owner: str(r[7]), contact: str(r[8]), priority: str(r[9]),
      hours: numOrNull(r[10]), amount: numOrNull(r[11]),
      order: str(r[12]), deliver: toDate(r[13]),
      content: str(r[14]), progress: str(r[15]), note: str(r[16]), memo: str(r[17]),
      kind: str(r[18]), stageStart: toDate(r[19]), basis: str(r[20]), deal: str(r[21]),
      confirm: str(r[22]), book: toDate(r[23]), finalHours: numOrNull(r[24]),
      finalAmount: numOrNull(r[25]), terms: str(r[26]),
      reporter: str(r[27]),
      quoteDone: toDate(r[28]), considerDone: toDate(r[29]),
      dealDone: toDate(r[30]), confirmDone: toDate(r[31]),
      orderDone: toDate(r[32]), workStart: toDate(r[33]), dueDate: toDate(r[34]),
      lastUpdate: toDate(r[35]),
      quoteLimit: toDate(r[36]),
      hold: holdCell || holdLegacy,
      holdLegacy,                    // 状態(E列)の手直しが必要な行
      holdWritten: holdCell,         // AL列に既にTRUEが入っているか
      workHours: numOrNull(r[38]),   // AM: 対応工数（人日）
      acceptHours: numOrNull(r[39]), // AN: 受託工数（人日）
    });
  });
  return out;
}
function toBool(v) {
  if (v === true) return true;
  const s = str(v).toUpperCase();
  return s === "TRUE" || s === "1" || s === "○" || s === "YES";
}
function normalizeStatus(s) {
  if (!s) return "新規";
  return LEGACY_STATUS[s] || s;
}

async function ensureCustomerSheet() {
  await Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let ws = sheets.items.find(s => s.name === CUST_SHEET);
    if (!ws) {
      const ns = sheets.add(CUST_SHEET);
      ns.getRange("A1:D1").values = [["顧客コード", "取引先名", "窓口", "備考"]];
      ns.getRange("A1:D1").format.fill.color = "#44546A";
      ns.getRange("A1:D1").format.font.color = "#FFFFFF";
      const seed = {};
      records.forEach(r => {
        const code = (r.id || "").split("-")[0];
        if (code && r.client && !seed[code]) seed[code] = r.client;
      });
      const rows = Object.entries(seed).map(([code, name]) => [code, name, "", ""]);
      if (rows.length) ns.getRange(`A2:D${rows.length + 1}`).values = rows;
      await ctx.sync();
    }
    const rng = ctx.workbook.worksheets.getItem(CUST_SHEET).getRange("A2:D200");
    rng.load("values");
    await ctx.sync();
    customers = rng.values
      .map((r, i) => ({ row: i + 2, code: str(r[0]), name: str(r[1]), contact: str(r[2]), note: str(r[3]) }))
      .filter(c => c.code && c.name);
  });
}

/* ---------- 確度ランク係数（「確度設定」シート） ----------
 * シート構成: A列=ランク名（濃厚/五分五分/薄め）, B列=係数（0〜1、%書式）
 * 優先度（高/中/低）はランク名に対応づけて重み付けする。未設定・空欄は「薄め」扱い。 */
const PRIORITY_TO_RANK = { "高": "濃厚", "中": "五分五分", "低": "薄め" };
const CONFIDENCE_SHEET = "確度設定";
let confidenceRates = { "濃厚": 1, "五分五分": 0.5, "薄め": 0 }; // シートが無い/読めない場合のデフォルト
function rankOfPriority(p) { return PRIORITY_TO_RANK[p] || "薄め"; }
function rateOfPriority(p) { return confidenceRates[rankOfPriority(p)] ?? 0; }
async function loadConfidenceRates() {
  try {
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const ws = sheets.items.find(s => s.name === CONFIDENCE_SHEET);
      if (!ws) return; // 無ければデフォルト値のまま
      const used = ws.getUsedRange(true);
      used.load("values");
      await ctx.sync();
      const rows = used.values.slice(1); // 1行目は見出し
      rows.forEach(r => {
        const label = str(r[0]);
        const pct = numOrNull(r[1]);
        if (label && pct != null) confidenceRates[label] = pct;
      });
    });
  } catch (e) {
    console.warn("確度設定の読み込みに失敗。デフォルト値を使用します。", e);
  }
}

function nextCaseId(code) {
  let max = 0;
  records.forEach(r => {
    if (r.id && r.id.startsWith(code + "-")) {
      const n = parseInt(r.id.slice(code.length + 1), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return `${code}-${String(max + 1).padStart(2, "0")}`;
}

function recToRow(rec) {
  return [[
    rec.id, rec.client, rec.no ?? "", rec.type, rec.status,
    toSerial(rec.occur), toSerial(rec.done),
    rec.owner ?? "", rec.contact ?? "", rec.priority ?? "",
    rec.hours ?? "", rec.amount ?? "", rec.order ?? "", toSerial(rec.deliver),
    rec.content ?? "", rec.progress ?? "", rec.note ?? "", rec.memo ?? "",
    rec.kind ?? "", toSerial(rec.stageStart), rec.basis ?? "", rec.deal ?? "",
    rec.confirm ?? "", toSerial(rec.book), rec.finalHours ?? "", rec.finalAmount ?? "",
    rec.terms ?? "",
    rec.reporter ?? "",
    toSerial(rec.quoteDone), toSerial(rec.considerDone),
    toSerial(rec.dealDone), toSerial(rec.confirmDone),
    toSerial(rec.orderDone), toSerial(rec.workStart), toSerial(rec.dueDate),
    toSerial(rec.lastUpdate),
    toSerial(rec.quoteLimit),
    rec.hold ? "TRUE" : "",
    rec.workHours ?? "", rec.acceptHours ?? "",
  ]];
}

async function writeRecord(rec) {
  rec.lastUpdate = new Date(); // 保存の都度、最終更新日を自動打刻
  if (demoMode) {
    const i = records.findIndex(r => r.id === rec.id);
    if (i >= 0) records[i] = rec; else { rec.row = 0; records.push(rec); }
    return;
  }
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
    let row = rec.row;
    if (!row) {
      // 既存レコードの最終行の次に追記（500行スキャンで使用範囲を広げない）
      let maxRow = 1;
      records.forEach(r => { if (r.row && r.row > maxRow) maxRow = r.row; });
      row = maxRow + 1;
      rec.row = row;
    }
    const rng = sheet.getRange(`A${row}:${LAST_COL}${row}`);
    rng.values = recToRow(rec);
    ["F", "G", "N", "T", "X", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK"].forEach(c =>
      sheet.getRange(`${c}${row}`).numberFormat = [["yyyy/m/d"]]);
    ["L", "Z"].forEach(c => sheet.getRange(`${c}${row}`).numberFormat = [["#,##0"]]);
    // 折り返しを無効化し行高さを固定（複数行に広がらないように）
    rng.format.wrapText = false;
    sheet.getRange(`${row}:${row}`).format.rowHeight = 18;
    await ctx.sync();
  });
  const i = records.findIndex(r => r.row === rec.row);
  if (i >= 0) records[i] = rec; else records.push(rec);
}

/* ドラッグによる日程変更専用の軽量書き込み（AH:開始日/workStart, AI:完了予定日/dueDateのみ）
 * writeRecord()はA:AI全35セル＋書式を毎回書き直すため、ドラッグ確定のたびに
 * 重いExcel.run()バッチが走る。ここでは変更のあった2セルだけを書き込む。 */
async function writeScheduleDates(rec) {
  if (demoMode) {
    const i = records.findIndex(r => r.row === rec.row);
    if (i >= 0) records[i] = rec;
    return;
  }
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
    const row = rec.row;
    const ah = sheet.getRange(`AH${row}`);
    const ai = sheet.getRange(`AI${row}`);
    ah.values = [[toSerial(rec.workStart)]];
    ai.values = [[toSerial(rec.dueDate)]];
    ah.numberFormat = [["yyyy/m/d"]];
    ai.numberFormat = [["yyyy/m/d"]];
    await ctx.sync();
  });
  const i = records.findIndex(r => r.row === rec.row);
  if (i >= 0) records[i] = rec;
}

async function writeCustomer(cust) {
  if (demoMode) { customers.push(cust); return; }
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(CUST_SHEET);
    const row = customers.length + 2;
    sheet.getRange(`A${row}:D${row}`).values = [[cust.code, cust.name, cust.contact ?? "", cust.note ?? ""]];
    await ctx.sync();
    cust.row = row;
  });
  customers.push(cust);
}

/* ============================================================
   ユーティリティ
   ============================================================ */
function str(v) { return v == null ? "" : String(v).trim(); }
function numOrNull(v) { return (v === "" || v == null) ? null : Number(v); }
function toDate(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400000));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function toSerial(d) { return d ? Math.round(d.getTime() / 86400000) + 25569 : ""; }
function fmtDate(d) { return d ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : ""; }
function md(d) { return d ? `${d.getMonth() + 1}/${d.getDate()}` : ""; }
function fmtDateInput(d) {
  if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function fromDateInput(s) { return s ? new Date(s + "T00:00:00") : null; }
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function monthKey(d) { return monthKeyYM(d.getFullYear(), d.getMonth() + 1); }
function monthKeyYM(y, m) { return `${y}/${String(m).padStart(2, "0")}`; }
function withTax(n) { return n == null ? null : Math.round(n * (1 + TAX_RATE)); }

function allStatusesOf(type) {
  const wf = WORKFLOWS[type];
  return wf ? [...wf.steps, ...wf.terminals, HOLD] : [];
}
function isTerminal(rec) {
  return rec.status === "失注" || rec.status === "完了";
}

/* 状態ラベル：保留中は「保留（元の状態）」、ワークフロー完了後は「状態（m/d）」 */
function statusLabel(rec) {
  if (isHold(rec)) {
    return rec.holdLegacy && rec.status === HOLD ? "保留（状態要確認）" : `保留（${rec.status}）`;
  }
  if (isTerminal(rec) && rec.done) return `${rec.status}（${md(rec.done)}）`;
  return rec.status;
}

function allowedTransitions(rec) {
  const wf = WORKFLOWS[rec.type];
  if (!wf) return [];
  const chain = wf.steps;
  const cur = rec.status;
  const res = [];
  /* 保留中は「保留解除（元の状態へ戻る）」のみ。進めるのは解除後。 */
  if (isHold(rec)) return chain.includes(cur) ? [cur] : [chain[0]];
  const idx = chain.indexOf(cur);
  if (idx >= 0) {
    if (idx + 1 < chain.length) res.push(chain[idx + 1]);
    else wf.terminals.forEach(t => res.push(t));
    if (idx > 0) res.push(chain[idx - 1]);
    res.push(HOLD);
  } else if (wf.terminals.includes(cur)) {
    res.push(chain[chain.length - 1]);
  }
  return res;
}
function isValidTransition(rec, to) { return allowedTransitions(rec).includes(to); }

function applyStatus(rec, to) {
  rec.status = to;
  const wf = WORKFLOWS[rec.type];
  if (to === "失注") { rec.order = "失注"; if (!rec.done) rec.done = new Date(); }
  else if (to === "受注") { rec.order = "受注"; }
  else if (to === "完了") { if (!rec.done) rec.done = new Date(); }
  else { rec.done = null; }
  if (wf && wf.steps.includes(to) && to !== "新規" && !rec.stageStart) rec.stageStart = new Date();
}

/* ============================================================
   タブ制御
   ============================================================ */
function switchTab(tab) {
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  ["list", "kanban", "sched", "agg"].forEach(t => {
    document.getElementById("pane-" + t).style.display = (t === tab) ? "" : "none";
  });
  document.getElementById("filter-bar").style.display =
    (tab === "list" || tab === "kanban") ? "" : "none";
  const isKanban = tab === "kanban";
  document.getElementById("ms-status").style.display = isKanban ? "none" : "";
  renderCurrentPane();
}
function activeTab() { return document.querySelector(".tab.active").dataset.tab; }
function renderCurrentPane() {
  const t = activeTab();
  if (t === "list") renderList();
  else if (t === "kanban") renderKanban();
  else if (t === "sched") renderSched();
  else if (t === "agg") renderAgg();
}

/* ============================================================
   フィルタ
   ============================================================ */
function renderFilters() {
  const clients = [...new Set(records.map(r => r.client).filter(Boolean))];
  renderMulti("status", [...new Set(records.map(r => effectiveStatus(r)).filter(s => s && s !== "削除"))], "状態");
  renderMulti("client", clients, "取引先");
  fillSelect("filter-owner", ["（担当者: 全て）", ...allOwners()], filters.owner);
}
function allOwners() {
  return [...new Set(records.flatMap(r => [...splitOwners(r.owner), ...splitOwners(r.reporter)]))];
}
function splitOwners(s) { return str(s).split(/[、,\s]+/).filter(Boolean); }
function fillSelect(id, options, selected) {
  const el = document.getElementById(id);
  el.innerHTML = options.map((o, i) =>
    `<option value="${i === 0 ? "" : esc(o)}"${o === selected ? " selected" : ""}>${esc(o)}</option>`).join("");
}

/* 汎用マルチセレクト（status/client 共通） */
const MS_LABEL = { status: "状態", client: "取引先" };
function renderMulti(key, options, label) {
  filters[key] = filters[key].filter(v => options.includes(v));
  const btn = document.getElementById(`ms-${key}-btn`);
  btn.textContent = filters[key].length ? `${label}: ${filters[key].length}件選択 ▾` : `（${label}: 全て）▾`;
  const dd = document.getElementById(`${key}-dd`);
  const pill = key === "status";
  dd.innerHTML = options.map(o => `
    <label class="ms-item">
      <input type="checkbox" value="${esc(o)}" ${filters[key].includes(o) ? "checked" : ""}
        onchange="onMsCheck('${key}',this)">
      ${pill ? `<span class="status-pill st-${esc(o)}">${esc(o)}</span>` : `<span>${esc(o)}</span>`}
    </label>`).join("") +
    `<button class="ms-clear" onclick="clearMsFilter('${key}')">選択解除</button>`;
}
function toggleMsDD(ev, key) {
  ev.stopPropagation();
  ["status", "client"].forEach(k => {
    const dd = document.getElementById(`${k}-dd`);
    if (dd) dd.style.display = (k === key && dd.style.display === "none") ? "" : "none";
  });
}
function onMsCheck(key, cb) {
  if (cb.checked) { if (!filters[key].includes(cb.value)) filters[key].push(cb.value); }
  else filters[key] = filters[key].filter(v => v !== cb.value);
  const label = MS_LABEL[key];
  document.getElementById(`ms-${key}-btn`).textContent =
    filters[key].length ? `${label}: ${filters[key].length}件選択 ▾` : `（${label}: 全て）▾`;
  saveFiltersCookie();
  renderCurrentPane();
}
function clearMsFilter(key) {
  filters[key] = [];
  saveFiltersCookie();
  renderFilters();
  renderCurrentPane();
}

/* 削除を除いた全レコード（集計はこれを使う） */
function activeRecords() { return records.filter(r => r.status !== "削除"); }

function filteredRecords(opts) {
  const skipStatus = !!(opts && opts.ignoreStatusFilter);
  return records.filter(r => {
    if (r.status === "削除") return false;   // 削除済みは表示しない
    if (!skipStatus && filters.status.length && !filters.status.includes(effectiveStatus(r))) return false;
    if (filters.client.length && !filters.client.includes(r.client)) return false;
    if (filters.owner && !splitOwners(r.owner).includes(filters.owner)) return false;
    if (filters.q) {
      const q = filters.q.toLowerCase();
      const hay = [r.id, r.client, r.content, r.progress, r.note, r.contact].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

/* ============================================================
   一覧（優先度を左端に）
   ============================================================ */
function renderList() {
  const cont = document.getElementById("list-container");
  const recs = filteredRecords();
  if (!recs.length) { cont.innerHTML = `<div class="empty-note">条件に一致する案件がありません</div>`; return; }
  let html = "";
  TYPES.forEach(type => {
    const all = recs.filter(r => r.type === type);
    if (!all.length) return;
    /* 保留は各種別グループの末尾に寄せる（本来の状態は状態ラベルに併記） */
    const group = [...all.filter(r => !isHold(r)), ...all.filter(r => isHold(r))];
    const holdCnt = all.filter(r => isHold(r)).length;
    const collapsed = collapsedTypes.has(type);
    html += `<div class="list-group lg-${type}${collapsed ? " collapsed" : ""}">
      <div class="list-group-head" onclick="toggleTypeGroup('${esc(type)}')">
        <span class="lg-chevron">▾</span>${esc(type)} <span class="cnt">${all.length}件</span>
        ${holdCnt ? `<span class="cnt cnt-hold">保留 ${holdCnt}</span>` : ""}
      </div>
      <table class="list-table">
        <tr><th>優先度</th><th>ID</th><th>取引先</th><th>状態</th><th>内容</th><th>担当</th><th>WBS</th><th>発生日</th><th>金額</th></tr>
        ${group.map(r => `
        <tr data-id="${esc(r.id)}" class="${r.id === selectedId ? "row-selected" : ""}${isHold(r) ? " row-hold" : ""}${isDimmedByWeek(r) ? " row-dim" : ""}"
            oncontextmenu="onRowContext(event,'${esc(r.id)}')"
            onclick="onRowClick('${esc(r.id)}')" ondblclick="openEditModal('${esc(r.id)}')">
          <td class="c">${r.priority ? `<span class="pri pri-${esc(r.priority)}">${esc(r.priority)}</span>` : ""}</td>
          <td class="muted">${esc(r.id)}</td>
          <td>${esc(r.client)}</td>
          <td><span class="status-pill st-${esc(effectiveStatus(r))}">${esc(statusLabel(r))}</span></td>
          <td>${esc(shorten(r.content, 34))}</td>
          <td>${esc(r.owner)}</td>
          <td class="c">${wbsBadgeHtml(r.id)}</td>
          <td class="muted">${fmtDate(r.occur)}</td>
          <td class="r">${dispAmount(r)}</td>
        </tr>`).join("")}
      </table>
    </div>`;
  });
  cont.innerHTML = html || `<div class="empty-note">案件がありません</div>`;
}
function dispAmount(r) {
  const a = r.finalAmount ?? r.amount;
  return a != null ? Number(a).toLocaleString() : "";
}
function shorten(s, n) { s = str(s).replace(/\n/g, " "); return s.length > n ? s.slice(0, n) + "…" : s; }
function onRowContext(ev, id) { ev.preventDefault(); openEditModal(id); }
/* スケジュール左クリック：行選択＋Excel該当行へジャンプ（排他ハイライト） */
function onGanttSelect(id) {
  selectedId = id;
  document.querySelectorAll("#gantt-wrap .g-row[data-id]").forEach(row =>
    row.classList.toggle("g-row-selected", row.dataset.id === id));
  const rec = records.find(r => r.id === id);
  if (rec) jumpToExcel(rec.row);
}
/* スケジュール右クリック：詳細画面を開く */
function onGanttContext(ev, id) {
  ev.preventDefault();
  openEditModal(id);
}
function onDrillContext(ev, id) {
  ev.preventDefault();
  openEditModal(id);
}
/* 見積一覧（ドリルダウン）共通: 左クリック=Excel該当行へジャンプ＆選択 */
function onDrillRowClick(id) {
  selectedId = id;
  document.querySelectorAll(".drill-row[data-id]").forEach(row =>
    row.classList.toggle("row-selected", row.dataset.id === id));
  const rec = records.find(r => r.id === id);
  if (rec) jumpToExcel(rec.row);
}

/* 左クリック：Excel該当行へジャンプ＆行選択、一覧はハイライト（排他） */
function onRowClick(id) {
  selectedId = id;
  document.querySelectorAll("#list-container tr[data-id]").forEach(tr =>
    tr.classList.toggle("row-selected", tr.dataset.id === id));
  const rec = records.find(r => r.id === id);
  if (rec) jumpToExcel(rec.row);
}
function onCardClick(id) {
  const rec = records.find(r => r.id === id);
  if (rec) jumpToExcel(rec.row);
}

/* Excelの該当行を選択状態にする（デモモードでは何もしない） */
async function jumpToExcel(row) {
  if (demoMode || !window.Office || !window.Excel || !row) return;
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
      sheet.activate();
      const range = sheet.getRange(`A${row}:${LAST_COL}${row}`);
      range.select();
      await ctx.sync();
    });
  } catch (e) {
    console.warn("Excel行選択に失敗:", e);
  }
}

/* ============================================================
   カンバン
   ============================================================ */
function renderKanban() {
  const bar = document.getElementById("kanban-typebar");
  bar.innerHTML = TYPES.map(t =>
    `<button class="${t === currentKanbanType ? "active" : ""}" data-type="${esc(t)}"
       onclick="setKanbanType('${esc(t)}')">${esc(t)}</button>`).join("");
  bar.className = "kanban-typebar type-seg";

  renderStepper(document.getElementById("kanban-stepper"), currentKanbanType, null);

  const board = document.getElementById("board");
  const lanes = allStatusesOf(currentKanbanType);
  const recs = filteredRecords({ ignoreStatusFilter: true }).filter(r => r.type === currentKanbanType);
  const dndLane = ENABLE_KANBAN_DND
    ? `ondragover="onLaneDragOver(event)" ondragleave="onLaneDragLeave(event)" ondrop="onLaneDrop(event)"` : "";
  board.innerHTML = lanes.map(st => {
    /* 保留は状態ではなくフラグ。保留レーンに寄せ、他レーンからは除く */
    const cards = (st === HOLD)
      ? recs.filter(r => isHold(r))
      : recs.filter(r => !isHold(r) && r.status === st);
    return `<div class="lane${st === HOLD ? " lane-hold" : ""}" data-status="${esc(st)}" ${dndLane}>
      <div class="lane-head">${esc(st)}<span class="cnt">${cards.length}</span></div>
      <div class="lane-body">
        ${cards.map(r => `
          <div class="card t-${esc(r.type)}${isHold(r) ? " card-hold" : ""}${isDimmedByWeek(r) ? " dim" : ""}" draggable="${ENABLE_KANBAN_DND}" data-id="${esc(r.id)}"
               ${ENABLE_KANBAN_DND ? `ondragstart="onCardDragStart(event)"` : ""}
               oncontextmenu="onRowContext(event,'${esc(r.id)}')"
               onclick="onCardClick('${esc(r.id)}')" ondblclick="openEditModal('${esc(r.id)}')">
            <div class="cid">${esc(r.id)}｜${esc(r.client)}</div>
            <div class="ctitle">${esc(shorten(r.content, 46))}</div>
            <div class="cmeta">
              <span>${esc(r.owner)}</span>
              ${wbsBadgeHtml(r.id)}
              ${dispAmount(r) ? `<span>${dispAmount(r)}円</span>` : ""}
              ${r.priority ? `<span>優先:${esc(r.priority)}</span>` : ""}
              ${isHold(r) ? `<span class="c-hold">元:${esc(r.holdLegacy && r.status === HOLD ? "要確認" : r.status)}</span>` : ""}
              ${isTerminal(r) && r.done ? `<span>${md(r.done)}完了</span>` : ""}
            </div>
          </div>`).join("")}
      </div>
    </div>`;
  }).join("");
}
function setKanbanType(t) { currentKanbanType = t; renderKanban(); }

function onCardDragStart(ev) { dragId = ev.currentTarget.dataset.id; }
function onLaneDragOver(ev) { ev.preventDefault(); ev.currentTarget.classList.add("drag-over"); }
function onLaneDragLeave(ev) { ev.currentTarget.classList.remove("drag-over"); }
async function onLaneDrop(ev) {
  ev.preventDefault();
  ev.currentTarget.classList.remove("drag-over");
  const to = ev.currentTarget.dataset.status;
  const rec = records.find(r => r.id === dragId);
  dragId = null;
  if (!rec) return;
  /* 保留レーンへ／から：状態は変えず保留フラグだけを切り替える */
  if (to === HOLD) {
    if (isHold(rec)) return;
    rec.hold = true;
    await writeRecord(rec);
    renderFilters(); renderKanban();
    return;
  }
  if (isHold(rec)) {
    if (to !== rec.status) {
      uiAlert(`保留中は状態を進められません。\n保留を解除すると「${rec.status}」から再開できます。`);
      return;
    }
    rec.hold = false;
    await writeRecord(rec);
    renderFilters(); renderKanban();
    return;
  }
  if (rec.status === to) return;
  if (!isValidTransition(rec, to)) {
    uiAlert(`「${rec.status}」から「${to}」へは遷移できません。\nワークフロー: ${workflowLabel(rec.type)}`);
    return;
  }
  if (to === "受注") {
    openEditModal(rec.id, "確認中");
    uiAlert("受注は編集画面の「確認中」タブで結果を選択し、「受注」タブで計上日等を入力してください。");
    return;
  }
  applyStatus(rec, to);
  await writeRecord(rec);
  renderFilters();
  renderKanban();
}
function workflowLabel(type) {
  const wf = WORKFLOWS[type];
  return wf ? [...wf.steps, wf.terminals.join(" or ")].join(" → ") : "";
}

/* ============================================================
   ステッパー
   ============================================================ */
/* ステッパー：ステージ名のピルを ▶ でつなぐ（ステージタブと表記を揃える）
 * 完了／失注に到達したときだけ、末尾にその結果ピルを足す。 */
function renderStepper(el, type, currentStatus, held) {
  const wf = WORKFLOWS[type];
  if (!wf) { el.innerHTML = ""; return; }
  const chain = stageTabsOf(type);                       // 起票／見積中／確認中／…
  const stage = currentStatus === "新規" ? "起票" : currentStatus;
  const idx = chain.indexOf(stage);
  const term = wf.terminals.includes(currentStatus) ? currentStatus : null;
  const pills = chain.map((s, i) => {
    let cls = "step-pill";
    if (currentStatus != null) {
      if (term || (idx >= 0 && i < idx)) cls += " done";
      else if (i === idx) cls += held ? " now held" : " now";
    }
    return `<span class="${cls}">${held && i === idx ? "❙❙ " : ""}${esc(s)}</span>`;
  });
  let html = pills.join(`<span class="step-sep">▶</span>`);
  if (term) {
    html += `<span class="step-sep">▶</span>` +
      `<span class="step-pill ${term === "失注" ? "lose" : "win"}">${esc(term)}</span>`;
  }
  /* 旧データ（状態＝保留）でステージが特定できない場合だけ保留ピルを添える */
  if (held && idx < 0) html += `<span class="step-pill hold-pill">❙❙ 保留中</span>`;
  el.innerHTML = html;
}

/* ============================================================
   新規入力モーダル（起票者を入力）
   ============================================================ */
function openNewModal() {
  renderInputForm();
  document.getElementById("in-msg").textContent = "";
  document.getElementById("new-modal").style.display = "";
}
function closeNewModal() { document.getElementById("new-modal").style.display = "none"; }

function renderInputForm() {
  const sel = document.getElementById("in-client");
  sel.innerHTML = `<option value="">選択してください</option>` +
    customers.map(c => `<option value="${esc(c.code)}">${esc(c.name)}（${esc(c.code)}）</option>`).join("");
  sel.onchange = updateNewId;

  const seg = document.getElementById("in-type-seg");
  seg.innerHTML = TYPES.map(t =>
    `<button data-type="${esc(t)}" class="${t === inputType ? "active" : ""}"
       onclick="setInputType('${esc(t)}')">${esc(t)}</button>`).join("");
  renderStepper(document.getElementById("in-stepper"), inputType, "新規");

  fillOwnerSelect("in-owner");
  document.getElementById("in-occur").value = fmtDateInput(new Date());
  updateNewId();
}
function setInputType(t) {
  inputType = t;
  document.querySelectorAll("#in-type-seg button").forEach(b =>
    b.classList.toggle("active", b.dataset.type === t));
  renderStepper(document.getElementById("in-stepper"), t, "新規");
}
function ownersOptions(selected) {
  const owners = allOwners();
  let html = `<option value=""></option>` +
    owners.map(o => `<option${o === selected ? " selected" : ""}>${esc(o)}</option>`).join("");
  if (selected && !owners.includes(selected)) html += `<option selected>${esc(selected)}</option>`;
  return html;
}
function fillOwnerSelect(id, selected) {
  document.getElementById(id).innerHTML = ownersOptions(selected);
}
function updateNewId() {
  const code = document.getElementById("in-client").value;
  document.getElementById("in-id").value = code ? nextCaseId(code) : "";
}

async function saveNewRecord() {
  const msg = document.getElementById("in-msg");
  msg.className = "save-msg"; msg.textContent = "";
  const code = document.getElementById("in-client").value;
  const content = document.getElementById("in-content").value.trim();
  if (!code) { msg.className = "save-msg err"; msg.textContent = "取引先を選択してください"; return; }
  if (!content) { msg.className = "save-msg err"; msg.textContent = "内容を入力してください"; return; }
  const reporter = document.getElementById("in-owner").value;
  const cust = customers.find(c => c.code === code);
  const rec = {
    row: 0,
    id: nextCaseId(code),
    client: cust.name, no: "",
    type: inputType, status: "新規",
    occur: fromDateInput(document.getElementById("in-occur").value) || new Date(),
    done: null,
    owner: reporter,          // 担当者の初期値 = 起票者
    reporter,                 // 起票者
    contact: cust.contact || "",
    priority: document.getElementById("in-priority").value,
    hours: null, amount: null, order: "", deliver: null,
    content,
    progress: "",
    note: document.getElementById("in-note").value,
    memo: "",
    kind: "", stageStart: null, basis: "", deal: "", confirm: "",
    book: null, finalHours: null, finalAmount: null, terms: "",
    quoteDone: null, considerDone: null, dealDone: null, confirmDone: null,
    quoteLimit: null, hold: false,
  };
  try {
    await writeRecord(rec);
    msg.textContent = `登録しました（${rec.id}）`;
    ["in-content", "in-note"].forEach(id => document.getElementById(id).value = "");
    renderFilters();
    renderCurrentPane();
    updateNewId();
    setTimeout(closeNewModal, 600);
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
  }
}

/* ============================================================
   編集モーダル（ステージタブ式）
   ============================================================ */
function openEditModal(id, forceTab) {
  const rec = records.find(r => r.id === id);
  if (!rec) return;
  editingRec = JSON.parse(JSON.stringify(rec), (k, v) =>
    (["occur","done","deliver","stageStart","book","quoteDone","considerDone","dealDone","confirmDone","orderDone","workStart","dueDate","quoteLimit"].includes(k) && v)
      ? new Date(v) : v);
  editingRec.row = rec.row;
  editDirty = false;
  /* タイトルバー：案件番号／取引先／内容（内容は長い場合は省略表示＋ツールチップ） */
  document.getElementById("ed-title").innerHTML =
    `<span class="mt-id">${esc(rec.id)}</span>` +
    `<span class="mt-client">${esc(rec.client)}</span>` +
    (str(rec.content) ? `<span class="mt-content" title="${esc(rec.content)}">${esc(rec.content)}</span>` : "");
  document.getElementById("ed-id").value = rec.id;
  document.getElementById("ed-client").value = rec.client;
  const tSel = document.getElementById("ed-type");
  // 一度登録した案件は種別変更不可（種別は固定表示）
  tSel.innerHTML = `<option>${esc(rec.type)}</option>`;
  tSel.value = rec.type;
  tSel.disabled = true;
  tSel.classList.add("ro");
  fillOwnerSelect("ed-owner", rec.owner);
  document.getElementById("ed-priority").value = rec.priority;
  document.getElementById("ed-note").value = rec.note;
  document.getElementById("ed-msg").textContent = "";
  // 削除ボタンは既に削除済みの場合は隠す
  const delBtn = document.getElementById("ed-delete-btn");
  if (delBtn) delBtn.style.display = (rec.status === "削除") ? "none" : "";
  currentStageTab = forceTab || defaultStageTab(editingRec);
  refreshEditModal();
  // 編集画面は必ず先に表示する。タスク一覧（WBS連携）は付随機能なので、
  // 初期化に失敗しても編集画面自体は開けるようにする。
  document.getElementById("edit-modal").style.display = "";
  try {
    initEdTaskPanel(rec.id);
  } catch (e) {
    console.warn("タスク一覧の初期化に失敗:", e);
  }
}
function markDirty() { editDirty = true; }

/* ============================================================
   WBSタスク一覧（ミニカンバン。中身は api.js の renderMiniKanban）
   ------------------------------------------------------------
   ・案件番号（＝小分類）が一致するwbsタスクを、案件編集モーダル内に
     未着手／対応中／完了の3レーンで表示する
   ・パネルは既定で閉じておき、「タスク一覧」ボタンで開閉する
   ・件数バッジは開閉に関わらず常に表示する
   ============================================================ */
/* api.js（共通モジュール）が読み込めているか。
   GitHub Pagesの公開前・通信不良・旧index.htmlのままなど、
   読み込めていないケースでも編集画面本体は動かせるようにする。 */
/* ============================================================
   WBS大分類の判定（種別＋状態から一意に決まる）
   ------------------------------------------------------------
     保守対応      → 保守
     瑕疵対応      → 瑕疵
     調整          → 調整
     見積り        → 見積   （受注確定後は 受託）
     プリセールス  → プリセ （受注確定後は 受託）
   受注確定 = 状態が 受注 / 受託中 / 完了（ORDER_CONFIRMED_STATUSES）
   小分類は常に案件番号。
   ============================================================ */
const WBS_CATEGORY_BY_TYPE = {
  "保守対応": "保守",
  "瑕疵対応": "瑕疵",
  "調整":     "調整",
  "見積り":       "見積",
  "プリセールス": "プリセ"
};
const WBS_ORDERED_CATEGORY = "受託";

function wbsCategoryOf(rec) {
  if (!rec) return "";
  // 見積り／プリセールスは受注確定後に「受託」へ切り替わる
  if (QUOTE_TYPES.includes(rec.type) && ORDER_CONFIRMED_STATUSES.includes(rec.status)) {
    return WBS_ORDERED_CATEGORY;
  }
  return WBS_CATEGORY_BY_TYPE[rec.type] || "";
}

/* 大分類がなぜその値になったかの説明（タスク追加モーダルに表示） */
function wbsCategoryReason(rec) {
  const cat = wbsCategoryOf(rec);
  if (!cat) return "";
  return (cat === WBS_ORDERED_CATEGORY)
    ? `種別「${rec.type}」＋状態「${rec.status}」から自動設定`
    : `種別「${rec.type}」から自動設定`;
}

/* ============================================================
   WBSタスク件数（一覧・カンバンのバッジ用）
   ------------------------------------------------------------
   案件ごとに毎回wbsを読むと重いので、起動時に一度だけ
   小分類(B列)→{total, done} を集計してキャッシュする。
   ============================================================ */
let wbsTaskCounts = {};   // { "KM-13": { total: 9, done: 9 }, ... }

async function loadWbsTaskCounts() {
  wbsTaskCounts = {};
  if (demoMode || !window.Excel) return;
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem("wbs");
      const used = sheet.getUsedRange(true);
      used.load(["rowIndex", "rowCount"]);
      await ctx.sync();

      const lastRow = Math.max(used.rowIndex + used.rowCount, 11);
      const range = sheet.getRangeByIndexes(0, 0, lastRow, 26); // A1:Z
      range.load("values");
      await ctx.sync();

      // 担当者（休み）の行数が変動するため、開始行は固定で持たない。
      // api.js の findWbsHeader() を使い、無い場合だけ従来の11行目にフォールバック。
      const dataIdx = (typeof findWbsHeader === "function")
        ? findWbsHeader(range.values).dataIdx
        : 10;

      range.values.slice(dataIdx).forEach(r => {
        if (!r[25] || r[19] === "-") return;                 // Z空 / T="-" は除外
        const key = (r[1] ?? "").toString().trim();          // B列=小分類=案件番号
        if (!key) return;
        if (!wbsTaskCounts[key]) wbsTaskCounts[key] = { total: 0, done: 0 };
        wbsTaskCounts[key].total++;
        if (r[18]) wbsTaskCounts[key].done++;                // S列=実績終了日
      });
    });
  } catch (e) {
    console.warn("WBSタスク件数の集計に失敗:", e);
  }
}

/* 一覧・カンバン用のバッジHTML（完了数/総数）。0件はグレー表示 */
function wbsBadgeHtml(caseId) {
  const c = wbsTaskCounts[caseId];
  if (!c || !c.total) return `<span class="wbs-badge none" title="WBSタスク未登録">0/0</span>`;
  const done = c.done === c.total;
  return `<span class="wbs-badge${done ? " all-done" : ""}" title="WBSタスク 完了${c.done} / 全${c.total}件">${c.done}/${c.total}</span>`;
}

function apiReady() {
  return typeof renderMiniKanban === "function" && typeof matchByCaseId === "function";
}

function initEdTaskPanel(caseId) {
  const caseEl = document.getElementById("ed-task-case");
  const panel = document.getElementById("task-panel");
  const btn = document.getElementById("ed-tasklist-btn");

  // 旧index.htmlのままなど、タスク一覧のDOMが無い場合は何もしない
  if (!panel || !btn) return;

  if (caseEl) caseEl.textContent = caseId;
  panel.classList.add("hidden");
  const closeBtn = panel.querySelector(".close-mini");
  if (closeBtn) closeBtn.textContent = "閉じる ▲";

  if (!apiReady()) {
    btn.classList.add("empty");
    btn.textContent = "タスク一覧（読込不可）";
    btn.disabled = true;
    return;
  }
  btn.disabled = false;
  refreshEdTaskPanel();
}

/* パネルの開閉トグル（開くときだけミニカンバンを再取得・再描画） */
function toggleEdTaskPanel() {
  const panel = document.getElementById("task-panel");
  if (!panel) return;
  const closeBtn = panel.querySelector(".close-mini");
  const willOpen = panel.classList.contains("hidden");
  panel.classList.toggle("hidden");
  if (closeBtn) closeBtn.textContent = willOpen ? "閉じる ▲" : "開く ▼";
  if (willOpen && editingRec) refreshEdTaskPanel();
}

/* ミニカンバンを再取得・再描画し、件数バッジも更新する */
async function refreshEdTaskPanel() {
  if (!editingRec || !apiReady()) return;
  if (!document.getElementById("ed-mk-board")) return;
  const caseId = editingRec.id;
  try {
    const tasks = await renderMiniKanban(
      "ed-mk-board",
      matchByCaseId(caseId),
      { onChanged: () => refreshEdTaskBadge(caseId) }
    );
    // 案件が切り替わっている間に届いた古い結果は無視する
    if (!editingRec || editingRec.id !== caseId) return;
    updateEdTaskBadge(tasks.length);
  } catch (e) {
    console.warn("タスク一覧の取得に失敗:", e);
  }
}

/* ドラッグ操作後など、パネルを開き直さずバッジだけ更新したいとき用 */
async function refreshEdTaskBadge(caseId) {
  if (!apiReady()) return;
  try {
    const tasks = await fetchWbsTasks(matchByCaseId(caseId));
    if (!editingRec || editingRec.id !== caseId) return;
    updateEdTaskBadge(tasks.length);
  } catch (e) {
    console.warn("タスク件数の取得に失敗:", e);
  }
}

function updateEdTaskBadge(n) {
  const btn = document.getElementById("ed-tasklist-btn");
  if (!btn) return;
  btn.classList.toggle("empty", n === 0);
  btn.innerHTML = n > 0
    ? `タスク一覧 <span class="cnt">${n}</span>`
    : `タスク未登録`;
}

/* 「＋タスク追加」：この案件の番号を小分類にあらかじめ選択した状態で開く */
function openTaskAddForCase() {
  if (!editingRec) return;
  if (typeof openTaskAdd !== "function") {
    uiAlert("タスク追加機能を読み込めませんでした。通信環境をご確認ください。");
    return;
  }
  const cat = wbsCategoryOf(editingRec);
  if (!cat) {
    uiAlert(`種別「${editingRec.type}」に対応するWBS大分類が未定義です。`);
    return;
  }
  openTaskAdd({
    lock: true,
    category: cat,
    subCategory: editingRec.id,
    reason: wbsCategoryReason(editingRec)
  });
}

/* ============================================================
   汎用ダイアログ（Office環境では window.confirm/alert 不可）
   ============================================================ */
let dialogResolve = null;
function uiConfirm(message) {
  return new Promise(resolve => {
    dialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "";
    document.getElementById("dialog-ok").textContent = "OK";
    document.getElementById("dialog-modal").style.display = "";
  });
}
function uiAlert(message) {
  return new Promise(resolve => {
    dialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "none";
    document.getElementById("dialog-ok").textContent = "OK";
    document.getElementById("dialog-modal").style.display = "";
  });
}
function dialogRespond(ok) {
  document.getElementById("dialog-modal").style.display = "none";
  const r = dialogResolve;
  dialogResolve = null;
  if (r) r(ok);
}

/* オーバーレイクリック時：変更があれば閉じない */
function tryCloseEditModal() {
  if (editDirty) {
    // 変更あり → 閉じない（誤操作防止）
    return;
  }
  closeEditModal();
}
function closeEditModal() { document.getElementById("edit-modal").style.display = "none"; editingRec = null; editDirty = false; }
/* ✕ボタン用：変更があれば確認してから閉じる */
async function closeEditModalConfirm() {
  if (editDirty) {
    const ok = await uiConfirm("変更内容が保存されていません。閉じてもよろしいですか？");
    if (!ok) return;
  }
  closeEditModal();
}

function activeStageTab(rec) {
  if (isTerminal(rec)) return null;
  const st = rec.status;
  if (st === "確認中" && rec.order === "受注") return "受注";       // 旧データ互換
  if (st === "受注") return "受注";                                 // 受注確定・最終登録（→受託中へ）
  if (st === "受託中") return "受託中";                             // 実行フェーズ
  if (st === "新規" || st === HOLD) return firstStageOf(rec.type);
  if (st === "対応中" || st === "見積中" || st === "検討中") return st;
  if (st === "商談中") return "商談中";
  if (st === "確認中") return "確認中";
  return null;
}
function defaultStageTab(rec) {
  return activeStageTab(rec) || stageTabsOf(rec.type)[stageTabsOf(rec.type).length - 1];
}

/* ステージの完了日（タブ表示用） */
function stageDoneDate(rec, t) {
  if (t === "対応中") return (rec.status === "完了") ? rec.done : null;
  if (t === "見積中") return rec.quoteDone;
  if (t === "検討中") return rec.considerDone;
  if (t === "商談中") return rec.dealDone;
  if (t === "確認中") return rec.confirmDone;
  if (t === "受注") return rec.orderDone;
  if (t === "受託中") return (rec.status === "完了") ? rec.done : null;
  return null;
}

function refreshEditModal() {
  const rec = editingRec;
  renderStepper(document.getElementById("ed-stepper"), rec.type, rec.status, isHold(rec));
  let stLabel = statusLabel(rec);
  if (!isHold(rec) && rec.status === "確認中" && rec.order === "受注") stLabel += "（受注・最終登録待ち）";
  if (isTerminal(rec)) stLabel += "／チケット完了";
  document.getElementById("ed-status").value = stLabel;
  document.getElementById("ed-owner").value = rec.owner;
  /* 保留チェックボックス：見積り／プリセールスのみ表示（終了案件は操作不可） */
  const wrap = document.getElementById("ed-hold-wrap");
  const cb = document.getElementById("ed-hold");
  if (wrap && cb) {
    const show = QUOTE_TYPES.includes(rec.type);
    wrap.style.display = show ? "" : "none";
    cb.checked = isHold(rec);
    cb.disabled = isTerminal(rec);
    wrap.classList.toggle("on", isHold(rec));
  }
  const banner = document.getElementById("ed-hold-note");
  if (banner) {
    banner.style.display = isHold(rec) ? "" : "none";
    banner.innerHTML = rec.holdLegacy && rec.status === HOLD
      ? `保留中です。旧データのため<b>保留前の状態が不明</b>です。状態(E列)を本来の状態に直してください。`
      : `保留中です。本来の状態「<b>${esc(rec.status)}</b>」は保持されているので、保留を解除すればそのまま再開できます。`;
  }
  renderHoursTop();
  renderStageTabs();
  renderStageBody();
}

/* 見積り／プリセールス：対応工数(人日)・受託工数(人日)（担当者行の右下、
   タスク一覧／タスク追加ボタンと同じ並びに表示）
   ・対応工数：「受託中」タブ以外で入力可（受託中はグレーアウト）
   ・受託工数：「受託中」タブでのみ入力可（それ以外はグレーアウト） */
function renderHoursTop() {
  const rec = editingRec;
  const group = document.getElementById("ed-hours-group");
  if (!group) return;
  const isQuote = QUOTE_TYPES.includes(rec.type);
  group.style.display = isQuote ? "" : "none";
  if (!isQuote) return;
  const isAccepted = currentStageTab === "受託中";
  const wh = document.getElementById("ed-workhours-top");
  const ah = document.getElementById("ed-accepthours-top");
  if (wh) { wh.value = rec.workHours ?? ""; wh.disabled = isAccepted; wh.classList.remove("need"); }
  if (ah) { ah.value = rec.acceptHours ?? ""; ah.disabled = !isAccepted; ah.classList.remove("need"); }
  /* 受託工数：「受託中」タブでは常に必須（対応完了時に検証）。
     対応工数：「受注」タブで受託確定チェックが入っている時のみ必須（onOrderDoneToggleが制御）。
     タブを離れたらいったんリセットする。 */
  const ahReq = document.getElementById("ed-accepthours-req");
  if (ahReq) ahReq.style.display = isAccepted ? "" : "none";
  const whReq = document.getElementById("ed-workhours-req");
  if (whReq && currentStageTab !== "受注") whReq.style.display = "none";
}

/* 保留チェックの切替：状態は変えず、保留フラグのみを更新（登録で確定） */
function onToggleHold(cb) {
  if (!editingRec) return;
  editingRec.hold = cb.checked;
  editDirty = true;
  refreshEditModal();
}

function renderStageTabs() {
  const rec = editingRec;
  const tabs = stageTabsOf(rec.type);
  const active = activeStageTab(rec);
  const el = document.getElementById("ed-stage-tabs");
  el.innerHTML = tabs.map(t => {
    const enabled = (t === "起票") || (t === active);
    const isCur = t === currentStageTab;
    const dd = stageDoneDate(rec, t);
    const label = dd ? `${t}（${md(dd)}）` : t;
    return `<button class="stage-tab${isCur ? " current" : ""}${enabled ? "" : " locked"}"
      onclick="setStageTab('${esc(t)}')">${esc(label)}${enabled || t === "起票" || dd ? "" : " 🔒"}</button>`;
  }).join("");
}
function setStageTab(t) { currentStageTab = t; renderHoursTop(); renderStageTabs(); renderStageBody(); }

/* ステージ担当者の変更を担当者欄に即時反映 */
function syncOwner(sel) {
  editingRec.owner = sel.value;
  const ed = document.getElementById("ed-owner");
  if (![...ed.options].some(o => o.value === sel.value)) {
    const opt = document.createElement("option");
    opt.textContent = sel.value;
    ed.appendChild(opt);
  }
  ed.value = sel.value;
}

/* 見積完了チェックのON/OFFで「見積有効期限」の必須表示を切り替える */
function onQuoteDoneToggle() {
  const done = document.getElementById("st-done");
  const req = document.getElementById("st-qlimit-req");
  const opt = document.getElementById("st-qlimit-opt");
  const inp = document.getElementById("st-qlimit");
  if (!req || !opt) return;
  const on = !!(done && done.checked);
  req.style.display = on ? "" : "none";
  opt.style.display = on ? "none" : "";
  if (inp) inp.classList.toggle("need", on && !inp.value);
}

/* 受注→受託中：納品日／最終工数／最終価格／対応工数を、受託確定チェックが
   入っている時だけ「必須」表示にする（登録時の必須チェックはsaveEditRecord側）。 */
function onOrderDoneToggle() {
  const done = document.getElementById("st-done");
  const on = !!(done && done.checked);
  [
    ["st-deliver", "st-deliver-req"],
    ["st-fhours", "st-fhours-req"],
    ["st-famount", "st-famount-req"],
  ].forEach(([inputId, reqId]) => {
    const inp = document.getElementById(inputId);
    const req = document.getElementById(reqId);
    if (req) req.style.display = on ? "" : "none";
    if (inp) inp.classList.toggle("need", on && !inp.value);
  });
  const whInp = document.getElementById("ed-workhours-top");
  const whReq = document.getElementById("ed-workhours-req");
  if (whReq) whReq.style.display = on ? "" : "none";
  if (whInp) whInp.classList.toggle("need", on && !whInp.value);
}

/* 受託中→完了：完了予定日／納品日／計上日／受託工数を、対応完了チェックが
   入っている時だけ「必須」表示にする（登録時の必須チェックはsaveEditRecord側）。 */
function onWorkDoneToggle() {
  const done = document.getElementById("st-done");
  const on = !!(done && done.checked);
  [
    ["st-duedate", "st-duedate-req"],
    ["st-deliver2", "st-deliver2-req"],
    ["st-book2", "st-book2-req"],
  ].forEach(([inputId, reqId]) => {
    const inp = document.getElementById(inputId);
    const req = document.getElementById(reqId);
    if (req) req.style.display = on ? "" : "none";
    if (inp) inp.classList.toggle("need", on && !inp.value);
  });
  const ahInp = document.getElementById("ed-accepthours-top");
  if (ahInp) ahInp.classList.toggle("need", on && !ahInp.value);
}

function renderStageBody() {
  const rec = editingRec;
  const body = document.getElementById("ed-stage-body");
  const active = activeStageTab(rec);
  const t = currentStageTab;
  const dis = (t !== "起票" && t !== active) || (isTerminal(rec) && t !== "起票") ? "disabled" : "";
  const disAll = isTerminal(rec) ? "disabled" : "";
  /* 保留中はどのステージも「完了して次へ進む」操作を止める（解除後に再開） */
  const holdLock = isHold(rec);
  const doneDis = dis || (holdLock ? "disabled" : "");
  const holdMsg = holdLock ? `<span class="hold-lock">保留中は次のステージへ進められません</span>` : "";

  if (t === "起票") {
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-row"><label>発生日</label><input type="date" id="st-occur" value="${fmtDateInput(rec.occur)}" ${disAll}></div>
        <div class="form-row"><label>起票者</label><select id="st-reporter" ${disAll}>${ownersOptions(rec.reporter)}</select></div>
        <div class="form-row"><label>窓口</label><input type="text" id="st-contact" value="${esc(rec.contact)}" ${disAll}></div>
      </div>
      <div class="form-row"><label>問合せ・提案内容</label>
        <textarea id="st-content" rows="4" ${disAll}>${esc(rec.content)}</textarea></div>`;
    return;
  }

  if (t === "対応中" || t === "見積中" || t === "検討中") {
    const isQuote = t === "見積中";
    const isHoshu = rec.type === "保守対応";
    const ownerLabel = t === "対応中" ? "対応担当者" : t === "見積中" ? "見積担当者" : "検討担当者";
    const doneLabel = t === "対応中" ? "対応完了（チケットを完了する）"
      : t === "見積中" ? "見積完了（確認中へ進める）"
      : "検討完了（商談中へ進める）";
    const progressLabel = t === "検討中" ? "対応状況（検討・検証・提案作成など）" : "対応状況";
    const dd = stageDoneDate(rec, t);
    const startInfo = rec.stageStart
      ? `<span class="stage-info">開始日: ${fmtDate(rec.stageStart)}${dd ? `　完了日: ${fmtDate(dd)}` : ""}</span>`
      : `<span class="stage-info">※対応状況を記入して登録すると「${esc(t)}」に遷移し、開始日を記録します</span>`;
    body.innerHTML = `
      ${startInfo}
      <div class="form-row"><label>${esc(ownerLabel)}（変更すると担当者欄も更新）</label>
        <select id="st-owner" ${dis} onchange="syncOwner(this)">${ownersOptions(rec.owner || rec.reporter)}</select></div>
      ${isHoshu ? `
      <div class="form-row"><label>区分 <span class="req">必須</span></label>
        <div class="radio-row">
          <label class="radio"><input type="radio" name="st-kind" value="問合せ" ${rec.kind === "問合せ" ? "checked" : ""} ${dis}>問合せ</label>
          <label class="radio"><input type="radio" name="st-kind" value="改修" ${rec.kind === "改修" ? "checked" : ""} ${dis}>改修</label>
        </div>
      </div>` : ""}
      <div class="form-row"><label>${esc(progressLabel)}</label>
        <textarea id="st-progress" rows="4" ${dis}>${esc(rec.progress)}</textarea></div>
      ${t === "対応中" ? `
      <div class="form-row"><label>対応工数（人日）</label>
        <input type="number" step="0.5" id="st-workhours" value="${rec.workHours ?? ""}" ${dis}></div>` : ""}
      ${isQuote ? `
      <div class="form-grid">
        <div class="form-row"><label>工数（人日）</label><input type="number" step="0.5" id="st-hours" value="${rec.hours ?? ""}" ${dis}></div>
        <div class="form-row"><label>価格（税抜・円）</label><input type="number" step="1000" id="st-amount" value="${rec.amount ?? ""}" ${dis} oninput="updateTaxView()"></div>
      </div>
      <div class="form-row"><label>税込価格（自動計算）</label>
        <input type="text" id="st-tax" readonly class="ro" value="${rec.amount != null ? withTax(rec.amount).toLocaleString() + " 円" : ""}"></div>
      <div class="form-row"><label>根拠</label>
        <textarea id="st-basis" rows="3" ${dis}>${esc(rec.basis)}</textarea></div>
      <div class="form-row"><label>見積有効期限
          <span class="req" id="st-qlimit-req" style="display:none">必須</span>
          <span class="opt-tag" id="st-qlimit-opt">任意（見積完了にする場合は必須）</span></label>
        <input type="date" id="st-qlimit" value="${fmtDateInput(rec.quoteLimit)}" ${dis}></div>` : ""}
      <label class="check-row ${doneDis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${doneDis} ${isQuote ? `onchange="onQuoteDoneToggle()"` : ""}> ${esc(doneLabel)}
        ${holdMsg}
      </label>`;
    if (isQuote) onQuoteDoneToggle();
    return;
  }

  if (t === "商談中") {
    const dd = rec.dealDone;
    body.innerHTML = `
      ${dd ? `<span class="stage-info">商談完了日: ${fmtDate(dd)}</span>` : ""}
      <div class="form-row"><label>窓口</label>
        <input type="text" id="st-dcontact" value="${esc(rec.contact)}" ${dis}></div>
      <div class="form-row"><label>商談状況</label>
        <textarea id="st-deal" rows="5" ${dis}>${esc(rec.deal)}</textarea></div>
      <label class="check-row ${doneDis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${doneDis}> 商談完了（確認中へ進める）
        ${holdMsg}
      </label>`;
    return;
  }

  if (t === "確認中") {
    const dd = rec.confirmDone;
    body.innerHTML = `
      ${dd ? `<span class="stage-info">確認完了日: ${fmtDate(dd)}</span>` : ""}
      <div class="form-row"><label>確認状況</label>
        <textarea id="st-confirm" rows="4" ${dis}>${esc(rec.confirm)}</textarea></div>
      <div class="form-row"><label>結果</label>
        <label class="check-row win-row ${doneDis ? "off" : ""}">
          <input type="checkbox" id="st-win" ${rec.order === "受注" ? "checked" : ""} ${doneDis}
            onchange="if(this.checked)document.getElementById('st-lose').checked=false">
          受注確定を完了にする（確認完了日を記録し、状態を受注にする）
        </label>
        <label class="check-row lose-row ${doneDis ? "off" : ""}">
          <input type="checkbox" id="st-lose" ${doneDis}
            onchange="if(this.checked)document.getElementById('st-win').checked=false">
          失注（確認完了日・完了日を記録し、チケット完了）
        </label>
        ${holdMsg}
      </div>`;
    return;
  }

  if (t === "受注") {
    const base = rec.finalAmount ?? rec.amount;
    const dd = rec.orderDone;
    body.innerHTML = `
      ${dd ? `<span class="stage-info">受注確定日: ${fmtDate(dd)}</span>` : ""}
      <div class="form-grid">
        <div class="form-row"><label>納品日 <span class="req" id="st-deliver-req" style="display:none">必須</span></label>
          <input type="date" id="st-deliver" value="${fmtDateInput(rec.deliver)}" ${dis}></div>
        <div class="form-row"><label>計上日</label><input type="date" id="st-book" value="${fmtDateInput(rec.book)}" ${dis}></div>
        <div class="form-row"><label>最終工数（人日） <span class="req" id="st-fhours-req" style="display:none">必須</span></label>
          <input type="number" step="0.5" id="st-fhours" value="${rec.finalHours ?? rec.hours ?? ""}" ${dis}></div>
        <div class="form-row"><label>最終価格（税抜・円） <span class="req" id="st-famount-req" style="display:none">必須</span></label>
          <input type="number" step="1000" id="st-famount" value="${base ?? ""}" ${dis} oninput="updateTaxView2()"></div>
      </div>
      <div class="form-row"><label>税込価格（自動計算）</label>
        <input type="text" id="st-tax2" readonly class="ro" value="${base != null ? withTax(base).toLocaleString() + " 円" : ""}"></div>
      <div class="form-row"><label>受注条件（必要に応じて）</label>
        <textarea id="st-terms" rows="3" ${dis}>${esc(rec.terms)}</textarea></div>
      <label class="check-row ${doneDis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${doneDis} onchange="onOrderDoneToggle()"> この内容で登録し、受注確定する（状態を受託中にする）
        ${holdMsg}
      </label>`;
    onOrderDoneToggle();
    return;
  }

  if (t === "受託中") {
    body.innerHTML = `
      ${rec.orderDone ? `<span class="stage-info">受注確定日: ${fmtDate(rec.orderDone)}</span>` : ""}
      <div class="form-grid">
        <div class="form-row"><label>開始日 <span class="req">必須</span></label>
          <input type="date" id="st-workstart" value="${fmtDateInput(rec.workStart)}" ${dis}></div>
        <div class="form-row"><label>完了予定日 <span class="req" id="st-duedate-req" style="display:none">必須</span></label>
          <input type="date" id="st-duedate" value="${fmtDateInput(rec.dueDate)}" ${dis}></div>
        <div class="form-row"><label>納品日 <span class="req" id="st-deliver2-req" style="display:none">必須</span></label>
          <input type="date" id="st-deliver2" value="${fmtDateInput(rec.deliver)}" ${dis}></div>
        <div class="form-row"><label>計上日 <span class="req" id="st-book2-req" style="display:none">必須</span></label>
          <input type="date" id="st-book2" value="${fmtDateInput(rec.book)}" ${dis}></div>
      </div>
      <div class="date-order-hint">開始日 ≦ 完了予定日 ≦ 納品日 ≦ 計上日 の順で入力してください</div>
      <label class="check-row ${doneDis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${doneDis} onchange="onWorkDoneToggle()"> 対応完了（完了日を記録し、チケットを完了する）
        ${holdMsg}
      </label>`;
    onWorkDoneToggle();
    return;
  }
  body.innerHTML = "";
}

/* 旧 loadWbsStatus()（受託中タブ内のWBS件数チップ）は廃止。
   案件編集画面 上部の「タスク一覧」（ミニカンバン）に一本化した。 */
function updateTaxView() {
  const v = numOrNull(document.getElementById("st-amount").value);
  document.getElementById("st-tax").value = v != null ? withTax(v).toLocaleString() + " 円" : "";
}
function updateTaxView2() {
  const v = numOrNull(document.getElementById("st-famount").value);
  document.getElementById("st-tax2").value = v != null ? withTax(v).toLocaleString() + " 円" : "";
}

async function saveEditRecord() {
  const rec = editingRec;
  const msg = document.getElementById("ed-msg");
  msg.className = "save-msg"; msg.textContent = "";

  rec.owner = document.getElementById("ed-owner").value;
  rec.priority = document.getElementById("ed-priority").value;
  rec.note = document.getElementById("ed-note").value;
  if (QUOTE_TYPES.includes(rec.type)) {
    const whTop = document.getElementById("ed-workhours-top");
    const ahTop = document.getElementById("ed-accepthours-top");
    if (whTop) rec.workHours = numOrNull(whTop.value);
    if (ahTop) rec.acceptHours = numOrNull(ahTop.value);
  }
  const holdCb = document.getElementById("ed-hold");
  if (holdCb && QUOTE_TYPES.includes(rec.type)) rec.hold = holdCb.checked;
  /* 状態が本来の値に直されたら、旧データ目印は外す */
  if (rec.holdLegacy && rec.status !== HOLD) rec.holdLegacy = false;

  const active = activeStageTab(rec);
  const t = currentStageTab;
  const editable = (t === "起票" && !isTerminal(rec)) || t === active;

  /* 保留中は「完了して次へ進む」系の操作を受け付けない（入力の保存自体は可） */
  if (editable && isHold(rec)) {
    const advancing = ["st-done", "st-win", "st-lose"]
      .some(id => { const el = document.getElementById(id); return el && el.checked; });
    if (advancing) {
      msg.className = "save-msg err";
      msg.textContent = `保留中は次のステージへ進められません。保留を解除すると「${rec.status}」から再開できます`;
      return;
    }
  }

  if (editable) {
    if (t === "起票") {
      rec.occur = fromDateInput(document.getElementById("st-occur").value);
      rec.reporter = document.getElementById("st-reporter").value;
      rec.contact = document.getElementById("st-contact").value;
      rec.content = document.getElementById("st-content").value;
    }
    else if (t === "対応中" || t === "見積中" || t === "検討中") {
      const stageOwner = document.getElementById("st-owner").value;
      if (stageOwner) rec.owner = stageOwner;  // 担当者欄を更新
      const progress = document.getElementById("st-progress").value;
      rec.progress = progress;
      if (rec.type === "保守対応") {
        const k = document.querySelector('input[name="st-kind"]:checked');
        rec.kind = k ? k.value : rec.kind;
      }
      if (t === "見積中") {
        rec.hours = numOrNull(document.getElementById("st-hours").value);
        rec.amount = numOrNull(document.getElementById("st-amount").value);
        rec.basis = document.getElementById("st-basis").value;
        rec.quoteLimit = fromDateInput(document.getElementById("st-qlimit").value);
      }
      if (t === "対応中") {
        const wh = document.getElementById("st-workhours");
        if (wh) rec.workHours = numOrNull(wh.value);   // 対応工数（人日）※AM列。旧K列(見積工数)とは別管理
      }
      if (rec.status === "新規" && progress.trim() && !isHold(rec)) {
        rec.status = t;
        if (!rec.stageStart) rec.stageStart = new Date();
      }
      const doneChk = document.getElementById("st-done");
      if (doneChk && doneChk.checked) {
        if (rec.type === "保守対応" && !rec.kind) {
          msg.className = "save-msg err"; msg.textContent = "保守対応は「問合せ／改修」の区分を選択してください"; return;
        }
        if (rec.status === "新規" && !progress.trim()) {
          msg.className = "save-msg err"; msg.textContent = "対応状況を記入してください"; return;
        }
        if (!rec.stageStart) rec.stageStart = new Date();
        if (t === "対応中") {
          if (rec.workHours == null) {
            msg.className = "save-msg err";
            msg.textContent = "完了にする場合は「対応工数（人日）」を入力してください";
            const inp = document.getElementById("st-workhours");
            if (inp) { inp.classList.add("need"); inp.focus(); }
            return;
          }
          applyStatus(rec, "完了");        // 対応完了日 = 完了日(G)
        }
        else if (t === "見積中") {
          if (rec.amount == null) { msg.className = "save-msg err"; msg.textContent = "価格を入力してください"; return; }
          // 見積完了にする場合のみ、見積有効期限を必須とする
          if (!rec.quoteLimit) {
            msg.className = "save-msg err";
            msg.textContent = "見積完了にする場合は「見積有効期限」を入力してください";
            const inp = document.getElementById("st-qlimit");
            if (inp) { inp.classList.add("need"); inp.focus(); }
            return;
          }
          rec.quoteDone = new Date();                             // 見積完了日
          applyStatus(rec, "確認中");
        }
        else if (t === "検討中") {
          rec.considerDone = new Date();                          // 検討完了日
          applyStatus(rec, "商談中");
        }
      }
    }
    else if (t === "商談中") {
      rec.contact = document.getElementById("st-dcontact").value;
      rec.deal = document.getElementById("st-deal").value;
      const doneChk = document.getElementById("st-done");
      if (doneChk && doneChk.checked) {
        rec.dealDone = new Date();                                // 商談完了日
        applyStatus(rec, "確認中");
      }
    }
    else if (t === "確認中") {
      rec.confirm = document.getElementById("st-confirm").value;
      const win = document.getElementById("st-win");
      const lose = document.getElementById("st-lose");
      if (lose && lose.checked) {
        rec.confirmDone = new Date();                             // 確認完了日
        applyStatus(rec, "失注");                                 // 完了日も更新
      } else if (win && win.checked) {
        rec.confirmDone = new Date();                             // 確認完了日
        rec.order = "受注";
        applyStatus(rec, "受注");                                 // 状態=受注（受注タブへ進む）
      }
    }
    else if (t === "受注") {
      rec.deliver = fromDateInput(document.getElementById("st-deliver").value);
      rec.book = fromDateInput(document.getElementById("st-book").value);
      rec.finalHours = numOrNull(document.getElementById("st-fhours").value);
      rec.finalAmount = numOrNull(document.getElementById("st-famount").value);
      rec.terms = document.getElementById("st-terms").value;
      const doneChk = document.getElementById("st-done");
      if (doneChk && doneChk.checked) {
        /* 受託中へ切り替える場合の必須項目：納品日／最終工数／最終価格／対応工数
           （確認メッセージはタブ切替時ではなく、この登録操作の時にのみ表示する） */
        const need = (id, label) => {
          msg.className = "save-msg err";
          msg.textContent = `受託中に切り替える場合は「${label}」を入力してください`;
          const inp = document.getElementById(id);
          if (inp) { inp.classList.add("need"); inp.focus(); }
        };
        if (!rec.deliver) { need("st-deliver", "納品日"); return; }
        if (rec.finalHours == null) { need("st-fhours", "最終工数（人日）"); return; }
        if (rec.finalAmount == null) { need("st-famount", "最終価格"); return; }
        if (rec.workHours == null) { need("ed-workhours-top", "対応工数（人日）"); return; }
        const ok = await uiConfirm(
          `対応工数（人日）: ${rec.workHours} 人日 で確定します。\nこの内容で登録し、状態を「受託中」にしますか？`);
        if (!ok) {
          msg.className = "save-msg";
          msg.textContent = "登録を中止しました";
          return;
        }
        rec.orderDone = new Date();          // 受注確定日
        applyStatus(rec, "受託中");          // 状態=受託中（受託中タブが活性化）
      }
    }
    else if (t === "受託中") {
      const workStart = fromDateInput(document.getElementById("st-workstart").value);
      const dueDate = fromDateInput(document.getElementById("st-duedate").value);
      const deliver = fromDateInput(document.getElementById("st-deliver2").value);
      const book = fromDateInput(document.getElementById("st-book2").value);
      // 相関チェック: 開始日 ≦ 完了予定日 ≦ 納品日 ≦ 計上日
      const seq = [["開始日", workStart], ["完了予定日", dueDate], ["納品日", deliver], ["計上日", book]]
        .filter(x => x[1]);
      for (let i = 1; i < seq.length; i++) {
        if (seq[i][1] < seq[i - 1][1]) {
          msg.className = "save-msg err";
          msg.textContent = `${seq[i - 1][0]}は${seq[i][0]}以前にしてください（開始日≦完了予定日≦納品日≦計上日）`;
          return;
        }
      }
      rec.workStart = workStart;
      rec.dueDate = dueDate;
      rec.deliver = deliver;
      rec.book = book;
      const doneChk = document.getElementById("st-done");
      if (doneChk && doneChk.checked) {
        /* 完了にする場合の必須項目：開始日／完了予定日／納品日／計上日／受託工数 */
        const need = (id, label) => {
          msg.className = "save-msg err";
          msg.textContent = `完了にする場合は「${label}」を入力してください`;
          const inp = document.getElementById(id);
          if (inp) { inp.classList.add("need"); inp.focus(); }
        };
        if (!workStart) { need("st-workstart", "開始日"); return; }
        if (!dueDate) { need("st-duedate", "完了予定日"); return; }
        if (!deliver) { need("st-deliver2", "納品日"); return; }
        if (!book) { need("st-book2", "計上日"); return; }
        if (rec.acceptHours == null) { need("ed-accepthours-top", "受託工数（人日）"); return; }
        applyStatus(rec, "完了");           // 完了日(G列)が記録されチケット完了
      }
    }
  }

  try {
    await writeRecord(rec);
    editDirty = false;
    renderFilters();
    renderCurrentPane();
    closeEditModal();       // 登録したら画面を閉じる
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
  }
}

/* ---------- 削除 ---------- */
async function deleteRecord() {
  if (!editingRec) return;
  const ok = await uiConfirm(`案件「${editingRec.id}　${editingRec.client}」を削除します。よろしいですか？\n（状態が「削除」となり、一覧・カンバンに表示されなくなります）`);
  if (!ok) return;
  const rec = editingRec;
  rec.status = "削除";
  const msg = document.getElementById("ed-msg");
  msg.className = "save-msg";
  try {
    await writeRecord(rec);
    editDirty = false;
    renderFilters();
    renderCurrentPane();
    closeEditModal();
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "削除に失敗しました: " + e.message;
  }
}

/* ---------- 顧客追加モーダル ---------- */
function openCustomerModal() { document.getElementById("cust-modal").style.display = ""; }
function closeCustomerModal() { document.getElementById("cust-modal").style.display = "none"; }
async function saveCustomer() {
  const msg = document.getElementById("cu-msg");
  msg.className = "save-msg"; msg.textContent = "";
  const code = document.getElementById("cu-code").value.trim().toUpperCase();
  const name = document.getElementById("cu-name").value.trim();
  if (!/^[A-Z]{2,4}$/.test(code)) { msg.className = "save-msg err"; msg.textContent = "顧客コードは英字2〜4文字です"; return; }
  if (customers.some(c => c.code === code)) { msg.className = "save-msg err"; msg.textContent = "そのコードは既に使われています"; return; }
  if (!name) { msg.className = "save-msg err"; msg.textContent = "取引先名を入力してください"; return; }
  try {
    await writeCustomer({
      code, name,
      contact: document.getElementById("cu-contact").value,
      note: document.getElementById("cu-note").value,
    });
    renderInputForm();
    document.getElementById("in-client").value = code;
    updateNewId();
    ["cu-code", "cu-name", "cu-contact", "cu-note"].forEach(id => document.getElementById(id).value = "");
    closeCustomerModal();
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
  }
}

/* ============================================================
   集計（期ベース: 10月〜翌9月）
   ============================================================ */
let currentAgg = "hoshu";
let showHours = true;        // 保守状況の対応工数の表示ON/OFF
let mitsuOpenStatus = null;  // 見積状況で件数展開中の状態
let mitsuOpenPri = null;     // 確度内訳で件数展開中の優先度
let mitsuSubView = "conf";   // 見積状況サブタブ: conf(確度内訳) / status(状態別集計)
function switchMitsuSub(v) {
  mitsuSubView = v;
  document.querySelectorAll(".mitsu-sub .seg").forEach(b => b.classList.toggle("active", b.dataset.sub === v));
  renderAgg();
}
function switchAgg(k) {
  currentAgg = k;
  if (k !== "mitsu") { mitsuOpenStatus = null; mitsuOpenPri = null; }
  document.querySelectorAll(".agg-seg .seg").forEach(b => b.classList.toggle("active", b.dataset.agg === k));
  renderAgg();
}
function toggleHours(cb) { showHours = cb.checked; renderAgg(); }
function shiftTerm(d) { currentTerm += d; renderAgg(); }
function termBarHtml() {
  return `<div class="term-bar">
    <button class="term-btn" onclick="shiftTerm(-1)">◀</button>
    <span class="term-label">${esc(termLabel(currentTerm))}</span>
    <button class="term-btn" onclick="shiftTerm(1)">▶</button>
    <span class="term-bar-sep"></span>
    <label class="hours-toggle" title="保守状況の対応工数の表示を切り替え">
      <input type="checkbox" id="hours-toggle-cb" ${showHours ? "checked" : ""} onchange="toggleHours(this)"> 工数表示
    </label>
  </div>`;
}
function renderAgg() {
  const cont = document.getElementById("agg-container");
  if (currentAgg === "hoshu") cont.innerHTML = termBarHtml() + renderHoshuAgg();
  else if (currentAgg === "mitsu") cont.innerHTML = renderMitsuAgg();
  else cont.innerHTML = termBarHtml() + renderJuchuAgg();
}

/* --- 保守状況 --- */
function renderHoshuAgg() {
  const months = fiscalMonths(currentTerm);
  const target = activeRecords().filter(r => r.type === "保守対応" || r.type === "瑕疵対応");
  const open = target.filter(r => !isTerminal(r)).length;

  /* 発生は「保守（問合せ）／保守（改修）／瑕疵」の3区分に分解して積み上げる */
  const hoshu = target.filter(r => r.type === "保守対応");
  const kashi = target.filter(r => r.type === "瑕疵対応");
  const stack = {
    "保守（問合せ）": countByMonth(hoshu.filter(r => r.kind === "問合せ"), "occur", months),
    "保守（改修）": countByMonth(hoshu.filter(r => r.kind === "改修"), "occur", months),
    "瑕疵": countByMonth(kashi, "occur", months),
  };
  /* 区分未設定の保守は「問合せ」に寄せず、内訳が合うように別途足す */
  const noKind = countByMonth(hoshu.filter(r => r.kind !== "問合せ" && r.kind !== "改修"), "occur", months);
  if (noKind.some(v => v)) stack["保守（区分未設定）"] = noKind;
  const stackColors = {
    "保守（問合せ）": "#2c6e9b", "保守（改修）": "#d9a038", "瑕疵": "#8e5aa8",
    "保守（区分未設定）": "#b9c2c9",
  };
  const doneSeries = countByMonth(target, "done", months);

  // 対応工数（人日）の月次集計：着手日ベース（無ければ完了日→発生日）
  const hours = sumByMonth(target, "workHours", "stageStart", months);
  const totalHoursR = Math.round(hours.reduce((a, v) => a + v, 0) * 10) / 10;

  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${target.length}</div><div class="kl">保守・瑕疵 総件数</div></div>
      <div class="kpi"><div class="kv">${open}</div><div class="kl">未完了件数</div></div>
      ${showHours ? `<div class="kpi"><div class="kv">${totalHoursR}</div><div class="kl">対応工数計（人日）</div></div>` : ""}
    </div>
    <div class="agg-card">
      <h3>保守・瑕疵 月次推移（発生の内訳・完了${showHours ? "＋対応工数" : ""}）</h3>
      ${legendHtml(stackColors, { "完了件数": "line:#548235", ...(showHours ? { "対応工数（人日・右軸）": "area:#ed7d31" } : {}) })}
      <div class="chart-wrap">${hoshuTrendChart(months, stack, stackColors, doneSeries, showHours ? hours : null)}</div>
      <p style="font-size:10px;color:#a9b2ba;margin-top:4px">積み上げ棒＝発生件数（左軸）／実線＝完了件数（左軸）${showHours ? "／面＝対応工数 人日（右軸・着手月ベース）" : ""}</p>
    </div>
    ${renderMaintList(target)}`;
}

/* --- 保守案件一覧（取引先・発生月・区分でフィルタ） --- */
let maintF = { clients: [], month: null, monthTerm: null, kinds: [] };
/* 月ラベル: "2026/07" → "7月" ／ 期の12か月は 10月〜9月の順 */
function monthBtnLabel(key) { return `${Number(key.slice(5))}月`; }
/* 既定は当月（表示中の期に当月が含まれない場合は「すべて」） */
function defaultMonthKey(months) {
  const k = monthKey(new Date());
  return months.includes(k) ? k : "";
}
function pickMaintMonth(v) {
  maintF.month = (maintF.month === v) ? "" : v;
  renderAgg();
}
function toggleMaintChip(which, v) {
  const arr = which === "client" ? maintF.clients : maintF.kinds;
  const i = arr.indexOf(v);
  if (i >= 0) arr.splice(i, 1); else arr.push(v);
  renderAgg();
}
function clearMaintChips(which) {
  if (which === "client") maintF.clients = []; else maintF.kinds = [];
  renderAgg();
}
function maintKindOf(r) {
  if (r.type === "瑕疵対応") return "瑕疵";
  return r.kind === "改修" ? "改修" : r.kind === "問合せ" ? "問合せ" : "区分未設定";
}
function renderMaintList(target) {
  const months = fiscalMonths(currentTerm);
  const clients = [...new Set(target.map(r => r.client).filter(Boolean))];
  const kinds = ["問合せ", "改修", "瑕疵"];
  maintF.clients = maintF.clients.filter(c => clients.includes(c));
  /* 期が変わったら当月を選び直す */
  if (maintF.monthTerm !== currentTerm) {
    maintF.month = defaultMonthKey(months);
    maintF.monthTerm = currentTerm;
  }
  if (maintF.month && !months.includes(maintF.month)) maintF.month = "";

  const rows = target.filter(r =>
    (!maintF.clients.length || maintF.clients.includes(r.client)) &&
    (!maintF.kinds.length || maintF.kinds.includes(maintKindOf(r))) &&
    (!maintF.month || (r.occur && monthKey(r.occur) === maintF.month))
  ).sort((a, b) => (b.occur || 0) - (a.occur || 0));

  const hoursSum = Math.round(rows.reduce((a, r) => a + (Number(r.workHours) || 0), 0) * 10) / 10;
  const cnt = k => rows.filter(r => maintKindOf(r) === k).length;

  const chip = (which, v, on) =>
    `<button class="fchip${on ? " on" : ""}" onclick="toggleMaintChip('${which}','${esc(v)}')">${esc(v)}</button>`;

  return `
    <div class="agg-card">
      <h3>保守案件一覧 <span class="cnt-inline">${rows.length}/${target.length}件</span></h3>
      <div class="agg-filters">
        <div class="af-row"><span class="af-label">取引先</span>
          <div class="fchips">
            <button class="fchip clear${maintF.clients.length ? "" : " on"}" onclick="clearMaintChips('client')">すべて</button>
            ${clients.map(c => chip("client", c, maintF.clients.includes(c))).join("")}
          </div></div>
        <div class="af-row"><span class="af-label">月別（発生月）</span>
          <div class="fchips month-chips">
            <button class="fchip clear${maintF.month ? "" : " on"}" onclick="pickMaintMonth('')">すべて</button>
            ${months.map(m => `<button class="fchip mchip${maintF.month === m ? " on" : ""}"
              title="${m}" onclick="pickMaintMonth('${m}')">${monthBtnLabel(m)}</button>`).join("")}
          </div></div>
        <div class="af-row"><span class="af-label">区分</span>
          <div class="fchips">
            <button class="fchip clear${maintF.kinds.length ? "" : " on"}" onclick="clearMaintChips('kind')">すべて</button>
            ${kinds.map(k => chip("kind", k, maintF.kinds.includes(k))).join("")}
          </div></div>
        <div class="af-meta">絞込結果: <b>${rows.length}件</b>　未完了 <b>${rows.filter(r => !isTerminal(r)).length}件</b>　
          問合せ <b>${cnt("問合せ")}</b>／改修 <b>${cnt("改修")}</b>／瑕疵 <b>${cnt("瑕疵")}</b>　
          対応工数計 <b>${hoursSum}人日</b></div>
      </div>
      <table class="agg-table drill-table">
        <tr><th>ID</th><th>取引先</th><th>内容</th><th>発生月</th><th>区分</th><th>工数</th><th>状態</th></tr>
        ${rows.length ? rows.map(r => `<tr class="drill-row${r.id === selectedId ? " row-selected" : ""}" data-id="${esc(r.id)}"
            onclick="onDrillRowClick('${esc(r.id)}')" oncontextmenu="onDrillContext(event,'${esc(r.id)}')">
          <td>${esc(r.id)}</td><td class="l">${esc(r.client)}</td>
          <td class="l">${esc(shorten(r.content, 26))}</td>
          <td>${r.occur ? monthKey(r.occur) : '<span class="muted">－</span>'}</td>
          <td><span class="kind-tag k-${esc(maintKindOf(r))}">${esc(maintKindOf(r))}</span></td>
          <td class="r">${r.workHours != null && r.workHours !== "" ? r.workHours : '<span class="muted">－</span>'}</td>
          <td><span class="status-pill st-${esc(effectiveStatus(r))}">${esc(statusLabel(r))}</span></td>
        </tr>`).join("") : `<tr><td colspan="7" class="muted">条件に一致する案件がありません</td></tr>`}
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">左クリック：Excelの該当行へジャンプ＆選択　／　右クリック：案件の編集画面を開く</p>
    </div>`;
}
function countByMonth(recs, field, months) {
  const map = Object.fromEntries(months.map(m => [m, 0]));
  recs.forEach(r => {
    const d = r[field];
    if (d && map[monthKey(d)] != null) map[monthKey(d)]++;
  });
  return months.map(m => map[m]);
}

/* --- 見積状況: 新規→検討中→見積中→商談中→確認中（＋保留は1行にまとめる） ---
 * 失注・受注・受託中・完了は集計対象から除外する。 */
const MITSU_ORDER = ["新規", "検討中", "見積中", "商談中", "確認中"];
const MITSU_EXCLUDED = ["失注", ...ORDER_CONFIRMED_STATUSES];
/* 見積一覧（ドリルダウン）用フィルタ */
let quoteF = { clients: [], owners: [] };
function toggleQuoteChip(which, v) {
  const arr = which === "client" ? quoteF.clients : quoteF.owners;
  const i = arr.indexOf(v);
  if (i >= 0) arr.splice(i, 1); else arr.push(v);
  renderAgg();
}
function clearQuoteChips(which) {
  if (which === "client") quoteF.clients = []; else quoteF.owners = [];
  renderAgg();
}

function renderMitsuAgg() {
  /* 集計対象: 見積り・プリセールスのうち 失注／受注／受託中／完了 を除いたもの */
  const pipeline = activeRecords().filter(r =>
    QUOTE_TYPES.includes(r.type) && !MITSU_EXCLUDED.includes(r.status));
  const held = pipeline.filter(r => isHold(r));
  const rows = MITSU_ORDER.map(st => {
    const g = pipeline.filter(r => !isHold(r) && r.status === st);
    return { key: st, st, cnt: g.length, amt: g.reduce((a, r) => a + (r.amount || 0), 0), sub: "" };
  });
  const brk = MITSU_ORDER.filter(st => held.some(r => r.status === st))
    .map(st => `${st} ${held.filter(r => r.status === st).length}`).join("・");
  const unknown = held.filter(r => !MITSU_ORDER.includes(r.status)).length;
  rows.push({
    key: HOLD, st: HOLD, cnt: held.length,
    amt: held.reduce((a, r) => a + (r.amount || 0), 0),
    sub: held.length
      ? `<span class="hold-brk">本来の状態: ${esc(brk || "－")}${unknown ? `・要確認 ${unknown}` : ""}</span>`
      : "",
  });
  const totalCnt = pipeline.length;
  const totalAmt = pipeline.reduce((a, r) => a + (r.amount || 0), 0);
  const pipelineAmt = totalAmt;
  // 加重パイプライン：優先度→確度ランク（確度設定シート）で重み付け
  const weightedAmt = pipeline.reduce((a, r) => a + (r.amount || 0) * rateOfPriority(r.priority), 0);
  const overdue = pipeline.filter(r => isQuoteOverdue(r)).length;

  const priGroups = {};
  pipeline.forEach(r => {
    const p = r.priority || "－";
    if (!priGroups[p]) priGroups[p] = { cnt: 0, amt: 0 };
    priGroups[p].cnt++;
    priGroups[p].amt += (r.amount || 0);
  });
  const priRows = ["高", "中", "低", "－"].filter(p => priGroups[p]).map(p => {
    const g = priGroups[p];
    const rank = rankOfPriority(p === "－" ? "" : p);
    const rate = confidenceRates[rank] ?? 0;
    return { p, rank, cnt: g.cnt, amt: g.amt, rate, weighted: g.amt * rate };
  });

  /* 確度内訳：件数展開中の優先度に対応する見積一覧 */
  let priDrillHtml = "";
  if (mitsuOpenPri) {
    const list = mitsuOpenPri === "__all"
      ? pipeline
      : pipeline.filter(r => (r.priority || "－") === mitsuOpenPri);
    priDrillHtml = quoteListHtml(list,
      mitsuOpenPri === "__all" ? "合計（全件）" : `優先度 ${esc(mitsuOpenPri)}`, "closePriDrill");
  }

  /* 状態別集計：展開中の状態に対応する見積一覧 */
  let drillHtml = "";
  if (mitsuOpenStatus) {
    const list = mitsuOpenStatus === "__all" ? pipeline
      : mitsuOpenStatus === HOLD ? held
        : pipeline.filter(r => !isHold(r) && r.status === mitsuOpenStatus);
    const title = mitsuOpenStatus === "__all" ? "合計（全件）"
      : `<span class="status-pill st-${esc(mitsuOpenStatus)}">${esc(mitsuOpenStatus)}</span>`;
    drillHtml = quoteListHtml(list, title, "closeMitsuDrill");
  }

  const confSection = `
    <div class="agg-card">
      <h3>優先度別 確度内訳</h3>
      <table class="agg-table">
        <tr><th>優先度</th><th>確度ランク</th><th>件数</th><th>金額計</th><th>確度</th><th>加重額</th></tr>
        ${priRows.map(r => `<tr>
          <td>${esc(r.p)}</td><td>${esc(r.rank)}</td>
          <td>${r.cnt > 0
            ? `<button class="cnt-link${mitsuOpenPri === r.p ? " open" : ""}" onclick="openPriDrill('${esc(r.p)}')">${r.cnt}</button>`
            : r.cnt}</td>
          <td class="r">${r.amt.toLocaleString()}円</td>
          <td class="r">${Math.round(r.rate * 100)}%</td>
          <td class="r">${Math.round(r.weighted).toLocaleString()}円</td></tr>`).join("")}
        <tr class="total"><td colspan="2">合計</td>
          <td><button class="cnt-link${mitsuOpenPri === "__all" ? " open" : ""}" onclick="openPriDrill('__all')">${pipeline.length}</button></td>
          <td class="r">${pipelineAmt.toLocaleString()}円</td><td class="r">－</td>
          <td class="r">${Math.round(weightedAmt).toLocaleString()}円</td></tr>
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">※ 確度は「確度設定」シート（優先度: 高＝濃厚／中＝五分五分／低＝薄め）の値を使用。未設定の優先度は「薄め」として計算。件数（合計を含む）をクリックすると下に見積一覧が表示されます。</p>
    </div>
    ${priDrillHtml}`;

  const statusSection = `
    <div class="agg-card">
      <h3>見積り・プリセールス 状態別集計</h3>
      <table class="agg-table">
        <tr><th>状態</th><th>件数</th><th>見積金額合計（税抜）</th></tr>
        ${rows.map(r => `<tr${r.key === HOLD ? ' class="row-hold"' : ""}>
          <td><span class="status-pill st-${esc(r.st)}">${esc(r.st)}</span>${r.sub}</td>
          <td>${r.cnt > 0
            ? `<button class="cnt-link${mitsuOpenStatus === r.key ? " open" : ""}" onclick="openMitsuDrill('${esc(r.key)}')">${r.cnt}</button>`
            : r.cnt}</td>
          <td class="r">${r.amt ? r.amt.toLocaleString() + "円" : "－"}</td></tr>`).join("")}
        <tr class="total"><td>合計</td>
          <td><button class="cnt-link${mitsuOpenStatus === "__all" ? " open" : ""}" onclick="openMitsuDrill('__all')">${totalCnt}</button></td>
          <td class="r">${totalAmt.toLocaleString()}円</td></tr>
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">※ 失注・受注・受託中・完了は集計対象外。保留は1行にまとめ、本来の状態を内訳として併記しています（二重計上なし）。件数（合計を含む）をクリックすると下に見積一覧が表示されます。</p>
    </div>
    ${drillHtml}`;

  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${pipeline.length}</div><div class="kl">進行中案件</div></div>
      <div class="kpi"><div class="kv">${(pipelineAmt / 10000).toLocaleString()}万</div><div class="kl">パイプライン金額</div></div>
      <div class="kpi"><div class="kv">${(weightedAmt / 10000).toLocaleString()}万</div><div class="kl">加重パイプライン（確度反映）</div></div>
      <div class="kpi kpi-hold"><div class="kv">${held.length}</div><div class="kl">うち保留</div></div>
      <div class="kpi kpi-alert"><div class="kv">${overdue}</div><div class="kl">見積有効期限 超過</div></div>
    </div>
    <div class="sched-seg mitsu-sub">
      <button class="seg${mitsuSubView === "conf" ? " active" : ""}" data-sub="conf" onclick="switchMitsuSub('conf')">確度内訳</button>
      <button class="seg${mitsuSubView === "status" ? " active" : ""}" data-sub="status" onclick="switchMitsuSub('status')">状態別集計</button>
    </div>
    ${mitsuSubView === "conf" ? confSection : statusSection}`;
}

/* 見積有効期限の超過判定・残日数 */
function today0() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function isQuoteOverdue(r) { return !!(r.quoteLimit && r.quoteLimit < today0()); }
function quoteLimitDays(r) {
  if (!r.quoteLimit) return null;
  return Math.round((r.quoteLimit - today0()) / 86400000);
}

/* 見積一覧（取引先・担当者フィルタ付き。超過を先頭に寄せて表示） */
function quoteListHtml(list, title, closeFn) {
  const clients = [...new Set(list.map(r => r.client).filter(Boolean))];
  const owners = [...new Set(list.flatMap(r => splitOwners(r.owner)))];
  quoteF.clients = quoteF.clients.filter(c => clients.includes(c));
  quoteF.owners = quoteF.owners.filter(o => owners.includes(o));

  const rows = list.filter(r =>
    (!quoteF.clients.length || quoteF.clients.includes(r.client)) &&
    (!quoteF.owners.length || splitOwners(r.owner).some(o => quoteF.owners.includes(o)))
  ).sort((a, b) => {
    /* ①超過（期限の古い順）→ ②期限内（期限の近い順）→ ③期限未設定 */
    const ra = !a.quoteLimit ? 2 : (isQuoteOverdue(a) ? 0 : 1);
    const rb = !b.quoteLimit ? 2 : (isQuoteOverdue(b) ? 0 : 1);
    if (ra !== rb) return ra - rb;
    if (!a.quoteLimit && !b.quoteLimit) return (b.occur || 0) - (a.occur || 0);
    return a.quoteLimit - b.quoteLimit;
  });

  const overdue = rows.filter(r => isQuoteOverdue(r)).length;
  const noLimit = rows.filter(r => !r.quoteLimit).length;
  const amt = rows.reduce((a, r) => a + (r.amount || 0), 0);
  const chip = (which, v, on) =>
    `<button class="fchip${on ? " on" : ""}" onclick="toggleQuoteChip('${which}','${esc(v)}')">${esc(v)}</button>`;

  return `
    <div class="agg-card drill-card">
      <h3>見積一覧：${title}（${rows.length}件）
        <button class="drill-close" onclick="${closeFn}()">閉じる ✕</button></h3>
      <div class="agg-filters">
        <div class="af-row"><span class="af-label">取引先</span>
          <div class="fchips">
            <button class="fchip clear${quoteF.clients.length ? "" : " on"}" onclick="clearQuoteChips('client')">すべて</button>
            ${clients.map(c => chip("client", c, quoteF.clients.includes(c))).join("")}
          </div></div>
        <div class="af-row"><span class="af-label">担当者</span>
          <div class="fchips">
            <button class="fchip clear${quoteF.owners.length ? "" : " on"}" onclick="clearQuoteChips('owner')">すべて</button>
            ${owners.map(o => chip("owner", o, quoteF.owners.includes(o))).join("")}
          </div></div>
        <div class="af-meta">絞込結果: <b>${rows.length}件</b>　金額合計 <b>${amt.toLocaleString()}円</b>　
          期限超過 <b class="txt-alert">${overdue}件</b>　期限未設定 <b>${noLimit}件</b></div>
      </div>
      <table class="agg-table drill-table">
        <tr><th>ID</th><th>取引先</th><th>種別</th><th>内容</th><th>担当</th><th>状態</th><th>見積有効期限</th><th>期限</th><th>金額</th></tr>
        ${rows.length ? rows.map(r => {
    const od = isQuoteOverdue(r);
    const dleft = quoteLimitDays(r);
    const judge = dleft == null ? `<span class="lim-tag lim-none">未設定</span>`
      : od ? `<span class="lim-tag lim-over">超過${-dleft}日</span>`
        : dleft <= 7 ? `<span class="lim-tag lim-soon">残${dleft}日</span>`
          : `<span class="lim-tag lim-ok">残${dleft}日</span>`;
    return `<tr class="drill-row${r.id === selectedId ? " row-selected" : ""}${od ? " row-overdue" : (isHold(r) ? " row-hold" : "")}"
            data-id="${esc(r.id)}" onclick="onDrillRowClick('${esc(r.id)}')" oncontextmenu="onDrillContext(event,'${esc(r.id)}')">
          <td>${esc(r.id)}</td><td class="l">${esc(r.client)}</td><td>${esc(r.type)}</td>
          <td class="l">${esc(shorten(r.content, 22))}</td><td>${esc(r.owner)}</td>
          <td><span class="status-pill st-${esc(effectiveStatus(r))}">${esc(statusLabel(r))}</span></td>
          <td class="${od ? "lim-cell-over" : ""}">${r.quoteLimit ? fmtDate(r.quoteLimit) : '<span class="muted">－</span>'}</td>
          <td>${judge}</td>
          <td class="r">${r.amount != null ? r.amount.toLocaleString() + "円" : "－"}</td>
        </tr>`;
  }).join("") : `<tr><td colspan="9" class="muted">条件に一致する案件がありません</td></tr>`}
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">並び順：①期限超過（古い順）→ ②期限内（近い順）→ ③期限未設定。左クリック：Excelの該当行へジャンプ＆選択　／　右クリック：案件の編集画面を開く</p>
    </div>`;
}
function openPriDrill(p) { if (mitsuOpenPri !== p) clearQuoteChips2(); mitsuOpenPri = (mitsuOpenPri === p) ? null : p; renderAgg(); }
function closePriDrill() { mitsuOpenPri = null; renderAgg(); }
function openMitsuDrill(st) { if (mitsuOpenStatus !== st) clearQuoteChips2(); mitsuOpenStatus = (mitsuOpenStatus === st) ? null : st; renderAgg(); }
function closeMitsuDrill() { mitsuOpenStatus = null; renderAgg(); }
function clearQuoteChips2() { quoteF.clients = []; quoteF.owners = []; }

/* --- 受注状況: 受注確定（受注区分=受注）の計上日ベース --- */
let juchuF = { clients: [], owners: [], month: null, monthTerm: null };
function pickJuchuMonth(v) {
  juchuF.month = (juchuF.month === v) ? "" : v;
  renderAgg();
}
function toggleJuchuChip(which, v) {
  const arr = which === "client" ? juchuF.clients : juchuF.owners;
  const i = arr.indexOf(v);
  if (i >= 0) arr.splice(i, 1); else arr.push(v);
  renderAgg();
}
function clearJuchuChips(which) {
  if (which === "client") juchuF.clients = []; else juchuF.owners = [];
  renderAgg();
}
function renderJuchuAgg() {
  const months = fiscalMonths(currentTerm);
  const won = activeRecords().filter(r => QUOTE_TYPES.includes(r.type) && ORDER_CONFIRMED_STATUSES.includes(r.status));
  const map = Object.fromEntries(months.map(m => [m, 0]));
  let noBook = 0;
  won.forEach(r => {
    const amt = (r.finalAmount ?? r.amount) || 0;
    if (r.book && map[monthKey(r.book)] != null) map[monthKey(r.book)] += amt;
    else if (!r.book) noBook++;
  });
  const vals = months.map(m => map[m]);
  const total = won.reduce((a, r) => a + ((r.finalAmount ?? r.amount) || 0), 0);
  const colors = { "確定売上": "#4472c4" };

  /* 一覧のフィルタ */
  const clients = [...new Set(won.map(r => r.client).filter(Boolean))];
  const owners = [...new Set(won.flatMap(r => splitOwners(r.owner)))];
  juchuF.clients = juchuF.clients.filter(c => clients.includes(c));
  juchuF.owners = juchuF.owners.filter(o => owners.includes(o));
  if (juchuF.monthTerm !== currentTerm) {
    juchuF.month = defaultMonthKey(months);
    juchuF.monthTerm = currentTerm;
  }
  if (juchuF.month && juchuF.month !== "__none" && !months.includes(juchuF.month)) juchuF.month = "";

  const rows = won.filter(r =>
    (!juchuF.clients.length || juchuF.clients.includes(r.client)) &&
    (!juchuF.owners.length || splitOwners(r.owner).some(o => juchuF.owners.includes(o))) &&
    (!juchuF.month
      ? true
      : juchuF.month === "__none" ? !r.book
        : (r.book && monthKey(r.book) === juchuF.month))
  ).sort((a, b) => (b.book || 0) - (a.book || 0));
  const rowsAmt = rows.reduce((a, r) => a + ((r.finalAmount ?? r.amount) || 0), 0);
  const chip = (which, v, on) =>
    `<button class="fchip${on ? " on" : ""}" onclick="toggleJuchuChip('${which}','${esc(v)}')">${esc(v)}</button>`;

  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${won.length}</div><div class="kl">受注案件数</div></div>
      <div class="kpi"><div class="kv">${(total / 10000).toLocaleString()}万</div><div class="kl">受注金額合計</div></div>
      ${noBook ? `<div class="kpi kpi-alert"><div class="kv">${noBook}</div><div class="kl">計上日未入力</div></div>` : ""}
    </div>
    <div class="agg-card">
      <h3>受注状況（受注確定・計上日ベース 月別売上）</h3>
      ${legendHtml(colors)}
      <div class="chart-wrap">${groupedBarChart(months, { "確定売上": vals }, colors, v => (v / 10000) + "万")}</div>
    </div>
    <div class="agg-card">
      <h3>受注案件一覧 <span class="cnt-inline">${rows.length}/${won.length}件</span></h3>
      <div class="agg-filters">
        <div class="af-row"><span class="af-label">取引先</span>
          <div class="fchips">
            <button class="fchip clear${juchuF.clients.length ? "" : " on"}" onclick="clearJuchuChips('client')">すべて</button>
            ${clients.map(c => chip("client", c, juchuF.clients.includes(c))).join("")}
          </div></div>
        <div class="af-row"><span class="af-label">月別（計上日）</span>
          <div class="fchips month-chips">
            <button class="fchip clear${juchuF.month ? "" : " on"}" onclick="pickJuchuMonth('')">すべて</button>
            ${months.map(m => `<button class="fchip mchip${juchuF.month === m ? " on" : ""}"
              title="${m}" onclick="pickJuchuMonth('${m}')">${monthBtnLabel(m)}</button>`).join("")}
            <button class="fchip mchip-none${juchuF.month === "__none" ? " on" : ""}"
              onclick="pickJuchuMonth('__none')">未入力</button>
          </div></div>
        <div class="af-row"><span class="af-label">担当者</span>
          <div class="fchips">
            <button class="fchip clear${juchuF.owners.length ? "" : " on"}" onclick="clearJuchuChips('owner')">すべて</button>
            ${owners.map(o => chip("owner", o, juchuF.owners.includes(o))).join("")}
          </div></div>
        <div class="af-meta">絞込結果: <b>${rows.length}件</b>　受注金額合計 <b>${rowsAmt.toLocaleString()}円</b></div>
      </div>
      <table class="agg-table drill-table">
        <tr><th>ID</th><th>取引先</th><th>内容</th><th>最終価格</th><th>計上日</th><th>納品日</th><th>担当</th><th>状態</th></tr>
        ${rows.length ? rows.map(r => `<tr class="drill-row${r.id === selectedId ? " row-selected" : ""}" data-id="${esc(r.id)}" onclick="onDrillRowClick('${esc(r.id)}')" oncontextmenu="onDrillContext(event,'${esc(r.id)}')">
          <td>${esc(r.id)}</td><td class="l">${esc(r.client)}</td>
          <td class="l">${esc(shorten(r.content, 20))}</td>
          <td class="r">${(r.finalAmount ?? r.amount) != null ? ((r.finalAmount ?? r.amount)).toLocaleString() + "円" : "－"}</td>
          <td>${r.book ? fmtDate(r.book) : '<span style="color:#c00000">未入力</span>'}</td>
          <td>${r.deliver ? fmtDate(r.deliver) : '<span class="muted">－</span>'}</td>
          <td>${esc(r.owner)}</td>
          <td><span class="status-pill st-${esc(effectiveStatus(r))}">${esc(statusLabel(r))}</span></td></tr>`).join("")
      : `<tr><td colspan="8" class="muted">条件に一致する受注案件がありません</td></tr>`}
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">左クリック：Excelの該当行へジャンプ＆選択　／　右クリック：案件の編集画面を開く</p>
    </div>`;
}

/* --- SVG グループ棒グラフ --- */
function groupedBarChart(labels, series, colors, fmtVal) {
  fmtVal = fmtVal || (v => String(v));
  const names = Object.keys(series);
  const W = Math.max(480, labels.length * 44), H = 190;
  const padL = 30, padB = 26, padT = 10;
  const chartW = W - padL - 8, chartH = H - padT - padB;
  const maxV = Math.max(1, ...names.flatMap(n => series[n]));
  const groupW = chartW / labels.length;
  const barW = Math.min(16, (groupW - 8) / names.length);
  let bars = "", labelsSvg = "", grid = "";
  const gridN = 4;
  for (let g = 0; g <= gridN; g++) {
    const y = padT + chartH - (chartH * g / gridN);
    grid += `<line x1="${padL}" y1="${y}" x2="${W - 4}" y2="${y}" stroke="#eceff2"/>` +
      `<text x="${padL - 4}" y="${y + 3}" font-size="8" text-anchor="end" fill="#999">${fmtVal(Math.round(maxV * g / gridN))}</text>`;
  }
  labels.forEach((lb, i) => {
    const gx = padL + groupW * i + (groupW - barW * names.length) / 2;
    names.forEach((n, j) => {
      const v = series[n][i];
      const h = chartH * v / maxV;
      const x = gx + j * barW, y = padT + chartH - h;
      bars += `<rect x="${x}" y="${y}" width="${barW - 1.5}" height="${h}" fill="${colors[n]}" rx="1.5">
        <title>${lb} ${n}: ${fmtVal(v)}</title></rect>`;
    });
    labelsSvg += `<text x="${padL + groupW * i + groupW / 2}" y="${H - 8}" font-size="8.5"
      text-anchor="middle" fill="#667">${lb.slice(2)}</text>`;
  });
  return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="max-width:100%">
    ${grid}${bars}${labelsSvg}</svg>`;
}
/* --- SVG 複合グラフ（棒＝件数[左軸] ＋ 折れ線＝工数[右軸]） --- */
function comboChart(labels, barSeries, barColors, lineSeries, lineColors, fmtLine) {
  fmtLine = fmtLine || (v => String(v));
  const bNames = Object.keys(barSeries);
  const lNames = Object.keys(lineSeries);
  const W = Math.max(500, labels.length * 46), H = 200;
  const padL = 26, padR = 30, padB = 26, padT = 10;
  const chartW = W - padL - padR, chartH = H - padT - padB;
  const maxBar = Math.max(1, ...bNames.flatMap(n => barSeries[n]));
  const maxLine = Math.max(1, ...lNames.flatMap(n => lineSeries[n]));
  const groupW = chartW / labels.length;
  const barW = Math.min(15, (groupW - 8) / bNames.length);
  const yBar = v => padT + chartH - (chartH * v / maxBar);
  const yLine = v => padT + chartH - (chartH * v / maxLine);
  const xCenter = i => padL + groupW * i + groupW / 2;
  let grid = "", bars = "", lines = "", dots = "", labelsSvg = "", axes = "";
  const gridN = 4;
  for (let g = 0; g <= gridN; g++) {
    const y = padT + chartH - (chartH * g / gridN);
    grid += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#eceff2"/>`;
    axes += `<text x="${padL - 3}" y="${y + 3}" font-size="8" text-anchor="end" fill="#8ba">${Math.round(maxBar * g / gridN)}</text>`;
    axes += `<text x="${W - padR + 3}" y="${y + 3}" font-size="8" text-anchor="start" fill="#d0894f">${fmtLine(Math.round(maxLine * g / gridN * 10) / 10)}</text>`;
  }
  labels.forEach((lb, i) => {
    const gx = padL + groupW * i + (groupW - barW * bNames.length) / 2;
    bNames.forEach((n, j) => {
      const v = barSeries[n][i];
      const h = chartH * v / maxBar;
      bars += `<rect x="${gx + j * barW}" y="${yBar(v)}" width="${barW - 1.5}" height="${h}" fill="${barColors[n]}" rx="1.5"><title>${lb} ${n}: ${v}</title></rect>`;
    });
    labelsSvg += `<text x="${xCenter(i)}" y="${H - 8}" font-size="8.5" text-anchor="middle" fill="#667">${lb.slice(2)}</text>`;
  });
  lNames.forEach(n => {
    const pts = lineSeries[n].map((v, i) => `${xCenter(i)},${yLine(v)}`).join(" ");
    lines += `<polyline points="${pts}" fill="none" stroke="${lineColors[n]}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    lineSeries[n].forEach((v, i) => {
      dots += `<circle cx="${xCenter(i)}" cy="${yLine(v)}" r="2.6" fill="${lineColors[n]}"><title>${labels[i]} ${n}: ${fmtLine(v)}</title></circle>`;
    });
  });
  return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="max-width:100%">
    ${grid}${axes}${bars}${lines}${dots}${labelsSvg}
    <text x="${padL - 3}" y="${padT - 1}" font-size="8" text-anchor="end" fill="#8ba">件</text>
    <text x="${W - padR + 3}" y="${padT - 1}" font-size="8" text-anchor="start" fill="#d0894f">人日</text>
  </svg>`;
}

/* --- SVG 積み上げ棒（発生の区分別）＋ 折れ線（完了）＋ 面（対応工数・右軸） --- */
function hoshuTrendChart(labels, stack, colors, doneSeries, hoursSeries) {
  const names = Object.keys(stack);
  const W = Math.max(520, labels.length * 46), H = 205;
  const padL = 28, padR = hoursSeries ? 34 : 8, padT = 14, padB = 26;
  const chartW = W - padL - padR, chartH = H - padT - padB;
  const totals = labels.map((_, i) => names.reduce((a, n) => a + (stack[n][i] || 0), 0));
  const maxC = Math.max(1, ...totals, ...doneSeries) + 1;
  const maxH = hoursSeries ? Math.max(1, ...hoursSeries) * 1.15 : 1;
  const step = chartW / labels.length;
  const bw = Math.min(20, step * 0.46);
  const cx = i => padL + step * (i + 0.5);
  const yC = v => padT + chartH - chartH * v / maxC;
  const yH = v => padT + chartH - chartH * v / maxH;
  let grid = "", area = "", bars = "", line = "", dots = "", xlab = "";
  for (let g = 0; g <= 4; g++) {
    const y = padT + chartH - chartH * g / 4;
    grid += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#eceff2"/>` +
      `<text x="${padL - 3}" y="${y + 3}" font-size="8" text-anchor="end" fill="#8ba">${Math.round(maxC * g / 4)}</text>`;
    if (hoursSeries) grid += `<text x="${W - padR + 3}" y="${y + 3}" font-size="8" text-anchor="start" fill="#c06e2e">${Math.round(maxH * g / 4 * 10) / 10}</text>`;
  }
  grid += `<text x="${padL - 3}" y="${padT - 4}" font-size="8" text-anchor="end" fill="#8ba">件</text>`;
  if (hoursSeries) grid += `<text x="${W - padR + 3}" y="${padT - 4}" font-size="8" text-anchor="start" fill="#c06e2e">人日</text>`;
  // 工数は背景の面グラフ（棒より先に描いて背面に置く）
  if (hoursSeries) {
    const pts = hoursSeries.map((v, i) => `${cx(i)},${yH(v)}`).join(" ");
    area = `<polygon points="${padL},${padT + chartH} ${pts} ${W - padR},${padT + chartH}" fill="rgba(237,125,49,.20)"/>` +
      `<polyline points="${pts}" fill="none" stroke="rgba(237,125,49,.70)" stroke-width="1.5"/>` +
      hoursSeries.map((v, i) => `<circle cx="${cx(i)}" cy="${yH(v)}" r="2.4" fill="#ed7d31"><title>${labels[i]} 対応工数: ${v}人日</title></circle>`).join("");
  }
  labels.forEach((lb, i) => {
    let acc = 0;
    names.forEach(n => {
      const v = stack[n][i] || 0;
      if (!v) return;
      const top = yC(acc + v), bot = yC(acc);
      bars += `<rect x="${cx(i) - bw / 2}" y="${top}" width="${bw}" height="${bot - top}" fill="${colors[n]}"><title>${lb} ${n}: ${v}件</title></rect>`;
      acc += v;
    });
    if (acc) bars += `<text x="${cx(i)}" y="${yC(acc) - 3}" font-size="8.5" text-anchor="middle" fill="#667" font-weight="bold">${acc}</text>`;
    xlab += `<text x="${cx(i)}" y="${H - 8}" font-size="8.5" text-anchor="middle" fill="#667">${lb.slice(2)}</text>`;
  });
  line = `<polyline points="${doneSeries.map((v, i) => `${cx(i)},${yC(v)}`).join(" ")}" fill="none" stroke="#548235" stroke-width="2" stroke-linejoin="round"/>`;
  dots = doneSeries.map((v, i) => `<circle cx="${cx(i)}" cy="${yC(v)}" r="2.8" fill="#fff" stroke="#548235" stroke-width="2"><title>${labels[i]} 完了: ${v}件</title></circle>`).join("");
  return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="max-width:100%">
    ${grid}${area}${bars}${line}${dots}${xlab}
    <line x1="${padL}" y1="${padT + chartH}" x2="${W - padR}" y2="${padT + chartH}" stroke="#c5d2d8"/></svg>`;
}

function legendHtml(colors, extra) {
  let html = Object.entries(colors).map(([n, c]) =>
    `<span><span class="sw" style="background:${c}"></span>${esc(n)}</span>`).join("");
  if (extra) {
    html += Object.entries(extra).map(([n, spec]) => {
      const [kind, c] = String(spec).split(":");
      const sw = kind === "line"
        ? `<span class="sw sw-line" style="border-top-color:${c}"></span>`
        : kind === "area"
          ? `<span class="sw" style="background:${c}33;border:1px solid ${c}"></span>`
          : `<span class="sw" style="background:${c}"></span>`;
      return `<span>${sw}${esc(n)}</span>`;
    }).join("");
  }
  return `<div class="legend">${html}</div>`;
}

/* --- SVG 折れ線グラフ --- */
function lineChart(labels, series, colors, fmtVal) {
  fmtVal = fmtVal || (v => String(v));
  const names = Object.keys(series);
  const W = Math.max(480, labels.length * 44), H = 190;
  const padL = 34, padB = 26, padT = 10;
  const chartW = W - padL - 8, chartH = H - padT - padB;
  const maxV = Math.max(1, ...names.flatMap(n => series[n]));
  const stepX = labels.length > 1 ? chartW / (labels.length - 1) : 0;
  const xAt = i => padL + stepX * i;
  const yAt = v => padT + chartH - (chartH * v / maxV);
  let grid = "", lines = "", dots = "", labelsSvg = "";
  const gridN = 4;
  for (let g = 0; g <= gridN; g++) {
    const y = padT + chartH - (chartH * g / gridN);
    grid += `<line x1="${padL}" y1="${y}" x2="${W - 4}" y2="${y}" stroke="#eceff2"/>` +
      `<text x="${padL - 4}" y="${y + 3}" font-size="8" text-anchor="end" fill="#999">${fmtVal(Math.round(maxV * g / gridN))}</text>`;
  }
  names.forEach(n => {
    const pts = series[n].map((v, i) => `${xAt(i)},${yAt(v)}`).join(" ");
    lines += `<polyline points="${pts}" fill="none" stroke="${colors[n]}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    series[n].forEach((v, i) => {
      dots += `<circle cx="${xAt(i)}" cy="${yAt(v)}" r="2.6" fill="${colors[n]}"><title>${labels[i]} ${n}: ${fmtVal(v)}</title></circle>`;
    });
  });
  labels.forEach((lb, i) => {
    labelsSvg += `<text x="${xAt(i)}" y="${H - 8}" font-size="8.5" text-anchor="middle" fill="#667">${lb.slice(2)}</text>`;
  });
  return `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="max-width:100%">
    ${grid}${lines}${dots}${labelsSvg}</svg>`;
}

/* 月別に数値フィールドを合計（日付フィールドで月を決定） */
function sumByMonth(recs, valField, dateField, months) {
  const map = Object.fromEntries(months.map(m => [m, 0]));
  recs.forEach(r => {
    const d = r[dateField] || r.done || r.occur;
    const v = Number(r[valField]) || 0;
    if (d && map[monthKey(d)] != null) map[monthKey(d)] += v;
  });
  return months.map(m => Math.round(map[m] * 10) / 10);
}

/* ============================================================
   デモモード
   ============================================================ */
function loadDemo() {
  demoMode = true;
  const d = (y, m, day) => new Date(y, m - 1, day);
  const blank = { kind: "", stageStart: null, basis: "", deal: "", confirm: "", book: null,
    finalHours: null, finalAmount: null, terms: "", reporter: "",
    quoteDone: null, considerDone: null, dealDone: null, confirmDone: null,
    orderDone: null, workStart: null, dueDate: null, lastUpdate: null,
    quoteLimit: null, hold: false, holdLegacy: false, holdWritten: false,
    workHours: null, acceptHours: null };
  const today = new Date();
  const daysAgo = n => { const dd = new Date(today); dd.setDate(dd.getDate() - n); return dd; };
  const daysFromNow = n => { const dd = new Date(today); dd.setDate(dd.getDate() + n); return dd; };
  customers = [
    { row: 2, code: "KM", name: "kakimoto arms", contact: "佐竹様", note: "" },
    { row: 3, code: "HN", name: "ハンター製菓", contact: "鈴木様", note: "" },
    { row: 4, code: "AG", name: "アサヒグラント", contact: "川野様", note: "" },
    { row: 5, code: "EX", name: "エキスプレス", contact: "中道様", note: "" },
  ];
  records = [
    { ...blank, row: 2, id: "KM-01", client: "kakimoto arms", no: 1, type: "見積り", status: "見積中", occur: d(2026, 6, 29), done: null, owner: "小川", reporter: "小川", contact: "佐竹様", priority: "中", hours: 10, amount: null, order: "", deliver: null, content: "ネット予約でフリースタッフを選択できるようにしたい", progress: "調査中", note: "", memo: "", stageStart: d(2026, 7, 1), amount: 480000, quoteLimit: daysAgo(12), workHours: 3 },
    { ...blank, row: 3, id: "KM-02", client: "kakimoto arms", no: 2, type: "見積り", status: "確認中", occur: d(2026, 6, 18), done: null, owner: "小川", reporter: "小川", contact: "佐竹様", priority: "", hours: 8, amount: 600000, order: "受注", deliver: d(2026, 7, 17), content: "ネット予約LINEログイン連携", progress: "60万で提示", note: "", memo: "", stageStart: d(2026, 6, 20), quoteDone: d(2026, 7, 1), confirmDone: d(2026, 7, 8), confirm: "受注の内諾。最終登録待ち", book: d(2026, 7, 31), finalAmount: 600000, finalHours: 8 },
    { ...blank, row: 4, id: "KM-03", client: "kakimoto arms", no: 3, type: "保守対応", status: "完了", occur: d(2026, 7, 2), done: d(2026, 7, 2), owner: "小川", reporter: "小川", contact: "西野様", priority: "", workHours: 0.5, amount: null, order: "", deliver: null, content: "スタッフ指名予約で店舗が正しく選択されない", progress: "外部サイト側の設定が原因", note: "", memo: "", kind: "問合せ", stageStart: d(2026, 7, 2) },
    { ...blank, row: 5, id: "KM-04", client: "kakimoto arms", no: 4, type: "調整", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", reporter: "小川", contact: "佐竹様", priority: "", workHours: null, amount: null, order: "", deliver: null, content: "会社体制変更に伴うご挨拶のスケジュール調整", progress: "日程調整中", note: "", memo: "", stageStart: d(2026, 7, 3) },
    { ...blank, row: 6, id: "KM-05", client: "kakimoto arms", no: 5, type: "保守対応", status: "対応中", occur: d(2026, 7, 7), done: null, owner: "小川", reporter: "紺谷", contact: "中田様", priority: "低", workHours: 2, amount: null, order: "", deliver: null, content: "メンズ予約時の注意事項表示・メール文面変更", progress: "設定変更で対応可能", note: "", memo: "", kind: "改修", stageStart: d(2026, 7, 8) },
    { ...blank, row: 7, id: "HN-01", client: "ハンター製菓", no: 1, type: "瑕疵対応", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", reporter: "小川", contact: "鈴木様", priority: "低", workHours: 1.5, amount: null, order: "", deliver: null, content: "在庫管理伝票一覧画面バグ対応", progress: "修正済み、次回リリースで反映", note: "", memo: "", stageStart: d(2026, 7, 4) },
    { ...blank, row: 8, id: "HN-02", client: "ハンター製菓", no: 2, type: "プリセールス", status: "商談中", occur: d(2026, 7, 6), done: null, owner: "小川", reporter: "小川", contact: "柳澤様", priority: "高", hours: null, amount: 2500000, order: "", deliver: null, content: "原価計算の改修", progress: "提案書作成済み", note: "9月本稼働目標", memo: "", stageStart: d(2026, 7, 7), considerDone: d(2026, 7, 15), deal: "7/22打ち合わせ予定", hold: true },
    { ...blank, row: 9, id: "AG-01", client: "アサヒグラント", no: 1, type: "見積り", status: "確認中", occur: d(2026, 6, 30), done: null, owner: "紺谷", reporter: "紺谷", contact: "川野様", priority: "中", hours: 5, amount: 350000, order: "", deliver: null, content: "インフォマートデータ交換の仕様変更", progress: "再見積提出済み", note: "", memo: "", stageStart: d(2026, 7, 1), quoteDone: d(2026, 7, 5), basis: "設計2人日＋実装2人日＋試験1人日", quoteLimit: daysFromNow(5) },
    { ...blank, row: 10, id: "EX-01", client: "エキスプレス", no: 1, type: "見積り", status: "新規", occur: d(2026, 7, 6), done: null, owner: "紺谷", reporter: "紺谷", contact: "中道様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "削除した請求書を参照できる機能の見積", progress: "", note: "", memo: "" },
    { ...blank, row: 11, id: "HN-03", client: "ハンター製菓", no: 3, type: "プリセールス", status: "新規", occur: d(2026, 7, 9), done: null, owner: "小川", reporter: "小川", contact: "", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "加工所日報のモバイル入力の提案", progress: "", note: "", memo: "" },
    { ...blank, row: 12, id: "AG-02", client: "アサヒグラント", no: 2, type: "見積り", status: "受託中", occur: d(2026, 5, 20), done: null, owner: "紺谷", reporter: "紺谷", contact: "川野様", priority: "中", hours: 6, amount: 480000, order: "受注", deliver: d(2026, 6, 30), content: "受注管理の帳票カスタマイズ", progress: "承認いただき受注確定", note: "", memo: "", stageStart: d(2026, 5, 22), quoteDone: d(2026, 5, 28), confirmDone: d(2026, 6, 10), confirm: "正式発注", book: d(2026, 8, 20), finalAmount: 480000, finalHours: 6, orderDone: d(2026, 6, 12), workStart: d(2026, 6, 20), dueDate: d(2026, 7, 31), workHours: 6, acceptHours: 4 },
    { ...blank, row: 13, id: "HN-04", client: "ハンター製菓", no: 4, type: "見積り", status: "保留", occur: d(2026, 5, 12), done: null, owner: "紺谷", reporter: "紺谷", contact: "柳澤様", priority: "中", hours: 3, amount: 220000, order: "", deliver: null, content: "旧データ：状態が保留のまま移行された案件", progress: "先方都合で一旦停止", note: "", memo: "", stageStart: d(2026, 5, 14), hold: true, holdLegacy: true },
    { ...blank, row: 14, id: "IH-02", client: "一広", no: 2, type: "見積り", status: "完了", occur: d(2026, 4, 10), done: d(2026, 6, 5), owner: "小川", reporter: "小川", contact: "宮崎様", priority: "", hours: 4, amount: 300000, order: "受注", deliver: d(2026, 5, 25), content: "取引先マスタ一括登録機能", progress: "対応完了", note: "", memo: "", stageStart: d(2026, 4, 12), quoteDone: d(2026, 4, 18), confirmDone: d(2026, 4, 25), confirm: "正式発注", book: d(2026, 5, 25), finalAmount: 300000, finalHours: 4, orderDone: d(2026, 4, 26), workStart: d(2026, 5, 1), dueDate: d(2026, 5, 20) },
  ];
  const lw = lastWeekRange();
  const midLastWeek = new Date(lw.start); midLastWeek.setDate(midLastWeek.getDate() + 2);
  const setLU = (id, d) => { const r = records.find(x => x.id === id); if (r) r.lastUpdate = d; };
  setLU("KM-01", midLastWeek);
  setLU("KM-02", midLastWeek);
  setLU("HN-01", daysAgo(20));
  setLU("AG-01", daysAgo(30));
}

/* ============================================================
   受託スケジュール（ガントチャート）
   ------------------------------------------------------------
   ・対象: 見積り/プリセールスで状態が 受注/受託中/完了 の案件
   ・バー: 開始=受託開始日(AH)、終了=完了なら完了日(G)、それ以外は納品日(N)
   ・▲: 納品日(N)。バー端ドラッグ=期日変更、本体ドラッグ=期間移動
   ・期(10月〜翌9月)単位。ズーム12/6/3/1ヶ月、背景ドラッグでパン
   ============================================================ */
let ganttTerm = termOfDate(new Date());
let ganttZoom = 12;               // 表示月数 12/6/3/1
let ganttHideDone = false;        // 完了除く（チェックすると完了案件を隠す）
const GANTT_ROW_H = 46;
const DAY_MS = 86400000;

function shiftGanttTerm(d) { ganttTerm += d; saveGanttCookie(); renderSched(); }
function setGanttZoom(m) { ganttZoom = m; saveGanttCookie(); renderSched(); }
function toggleGanttHideDone(cb) { ganttHideDone = cb.checked; saveGanttCookie(); renderSched(); }

/* ============================================================
   ① 案件行をクリックしたときの WBSタスク展開
   ------------------------------------------------------------
   小分類 = その案件ID のタスクを抽出し、1件1行の「線」と
   リストを案件行の直下に差し込む。描画部品は api.js と共有。
   ============================================================ */
let schedExpandedId = null;
let schedExpandTasks = [];
let schedShowLeave = false;
let schedExpandMeta = null;
let schedHolidaysTried = false;

async function toggleSchedExpand(id) {
  if (schedExpandedId === id) {
    schedExpandedId = null;
    schedExpandTasks = [];
    schedExpandMeta = null;
    renderSched();
    return;
  }
  schedExpandedId = id;
  schedExpandTasks = [];
  schedExpandMeta = null;
  if (typeof fetchWbsTasks === "function" && typeof matchByCaseId === "function") {
    try {
      schedExpandTasks = await fetchWbsTasks(matchByCaseId(id));
    } catch (e) {
      console.warn("WBSタスクの取得に失敗:", e);
    }
  }
  if (schedShowLeave && typeof loadLeaveGrid === "function") {
    try { await loadLeaveGrid(); } catch (e) { /* 休みが読めなくてもガントは出す */ }
  }
  try {
    renderSched();
  } catch (e) {
    console.error("スケジュールの再描画に失敗:", e);
    schedExpandedId = null; schedExpandTasks = [];
    renderSched();
  }
}

/* 案件行クリック：選択（既存）＋ タスク展開 */
function onGanttRowClick(id) {
  onGanttSelect(id);
  toggleSchedExpand(id);
}

async function toggleSchedLeave(cb) {
  schedShowLeave = !!(cb && cb.checked);
  if (schedShowLeave && typeof loadLeaveGrid === "function") {
    try { await loadLeaveGrid(); } catch (e) { /* 同上 */ }
  }
  renderSched();
}

function renderSched() {
  const cont = document.getElementById("sched-container");
  if (!cont) return;

  /* 休業日（wbs 8行目）を1度だけ読み込む。読めたらもう一度描き直して影を出す。
     読めなくてもガント自体は描くので、失敗しても表示は壊れない。 */
  if (!schedHolidaysTried && typeof loadHolidays === "function") {
    schedHolidaysTried = true;
    loadHolidays().then(() => renderSched()).catch(() => { /* 影なしで続行 */ });
  }
  cont.innerHTML = ganttHtml();
  setupGantt();
  if (typeof bindTaskLines === "function") bindTaskLines(cont);
}

/* 展開部（線＋リスト）。api.js の共通部品で描く */
function schedExpandHtml(r, g, geo) {
  if (schedExpandedId !== r.id) return "";
  if (typeof taskLineRowsHtml !== "function") return "";
  try {
    return schedExpandHtmlInner(r, g, geo);
  } catch (e) {
    // 展開部だけの失敗でガント全体を壊さない
    console.warn("タスク展開の描画に失敗:", e);
    return `<div class="g-row wsc-listrow"><div class="g-label"></div><div class="g-track">
      <div class="wsc-list"><div class="wsc-empty">タスクの表示に失敗しました。再読み込みしてください。</div></div>
    </div></div>`;
  }
}

function schedExpandHtmlInner(r, g, geo) {

  const tasks = schedExpandTasks.slice().sort((a, b) => {
    const x = a.start ? toDate(a.start) : null;
    const y = b.start ? toDate(b.start) : null;
    if (!x && !y) return 0;
    if (!x) return 1;
    if (!y) return -1;
    return x - y;
  });

  const geom = {
    t0: geo.t0, totalDays: geo.totalDays,
    monthLines: geo.monthLines, todayHtml: geo.todayHtml,
    guide: { start: g.start, end: g.end },
    dueLimit: r.dueDate || null,
    showActual: true,
    showLeave: schedShowLeave,
    showCat: true
  };
  const done = tasks.filter(t => t.actualEnd).length;
  const over = r.dueDate
    ? tasks.filter(t => t.end && toDate(t.end) > r.dueDate).length : 0;

  // リストはガントの外（ズームで横に伸びないように）へ回す
  schedExpandMeta = {
    tasks, showCat: true, bare: true, dueLimit: r.dueDate || null,
    title: `${r.id}　${r.client}　${r.content || ""}`,
    meta: `${r.status} ／ タスク ${tasks.length}件 ／ 完了 ${done}件`
          + (over ? ` ／ 完了予定日を超過 ${over}件` : ""),
    hint: `「＋ タスク追加」で小分類に ${r.id} を入れて登録すると、この行に並びます。`
  };
  return taskLineRowsHtml(tasks, geom);
}

/* ガントの外に出すリスト。ズーム倍率の影響を受けない */
function schedExpandListHtml() {
  if (!schedExpandedId || !schedExpandMeta) return "";
  if (typeof taskListHtml !== "function") return "";
  const m = schedExpandMeta;
  return `<div class="wsc-outlist">${taskListHtml(m.tasks, m)}</div>`;
}

/* --- 期の開始/終了日 --- */
function termStartDate(term) { return new Date(term + 1988, 9, 1); }               // 10/1
function termEndDate(term) { return new Date(term + 1989, 9, 1); }                 // 翌10/1(排他)

/* --- ガント対象レコード --- */
function ganttRecords() {
  return activeRecords()
    .filter(r => QUOTE_TYPES.includes(r.type) && ORDER_CONFIRMED_STATUSES.includes(r.status))
    .filter(r => !(ganttHideDone && r.status === "完了"))
    .map(r => {
      // バー: 開始日(AH=workStart) 〜 完了予定日(AI=dueDate)
      const start = r.workStart || r.orderDone || r.book || r.occur;
      const end = r.dueDate
        || (r.status === "完了" ? r.done : null)
        || (start ? new Date(start.getTime() + 14 * DAY_MS) : null);
      return { rec: r, start, end, provisional: !r.workStart || !r.dueDate };
    })
    .filter(g => g.start && g.end)
    .sort((a, b) => a.start - b.start);
}

function ganttHtml() {
  const t0 = termStartDate(ganttTerm);
  const t1 = termEndDate(ganttTerm);
  const totalDays = Math.round((t1 - t0) / DAY_MS);
  const widthPct = (12 / ganttZoom) * 100;          // 内側の横幅（ビューポート比）

  /* カレンダー（api.js の共通実装。WBSカンバンのスケジュールと同じ見た目）
     内部は常に12ヶ月ぶん描き、ズームは幅の倍率で表すため、
     描く月数は12、密度の基準は ganttZoom を渡す。 */
  let monthCells = "", dateCells = "", monthLines = "", todayHtml = "", todayChip = "";
  let calRows = null;
  if (typeof schedCalendar === "function") {
    const cal = schedCalendar({ t0, t1, totalDays, monthCount: 12, zoom: ganttZoom });
    calRows = cal.rows;
    monthLines = cal.bg;
    todayHtml = cal.todayLine;
    todayChip = cal.todayChip;
  } else {
    // api.js が古い場合のフォールバック（従来の簡易ヘッダー）
    for (let i = 0; i < 12; i++) {
      const d = new Date(t0.getFullYear(), t0.getMonth() + i, 1);
      const next = new Date(t0.getFullYear(), t0.getMonth() + i + 1, 1);
      const left = ((d - t0) / DAY_MS) / totalDays * 100;
      const width = ((next - d) / DAY_MS) / totalDays * 100;
      monthCells += `<div class="g-mcell" style="left:${left}%;width:${width}%">${String(d.getFullYear()).slice(2)}/${String(d.getMonth() + 1).padStart(2, "0")}</div>`;
      if (i > 0) monthLines += `<div class="g-vline" style="left:${left}%"></div>`;
    }
    const today0 = new Date(); today0.setHours(0, 0, 0, 0);
    if (today0 >= t0 && today0 < t1) {
      const lp = ((today0 - t0) / DAY_MS) / totalDays * 100;
      todayHtml = `<div class="g-today" style="left:${lp}%"></div>`;
      todayChip = todayHtml + `<div class="g-today-chip" style="left:${lp}%">今日 ${md(today0)}</div>`;
    }
    calRows = [{ h: 20, html: monthCells + todayChip }];
  }

  const items = ganttRecords();
  const rows = items.map((g, i) => {
    const r = g.rec;
    const l = Math.max(((g.start - t0) / DAY_MS) / totalDays * 100, -8);
    const rEnd = Math.min(((g.end - t0) / DAY_MS + 1) / totalDays * 100, 108);
    const w = Math.max(rEnd - l, 0.7);
    const cls = r.status === "完了" ? "done" : r.status === "受託中" ? "work" : "order";
    const locked = r.status === "完了";
    let deliverHtml = "";
    if (r.deliver && r.deliver >= t0 && r.deliver < t1) {
      const dl = ((r.deliver - t0) / DAY_MS + 0.5) / totalDays * 100;
      deliverHtml = `<div class="g-deliver" style="left:${dl}%">▲${md(r.deliver)}</div>`;
    }
    return `
    <div class="g-row${r.id === selectedId ? " g-row-selected" : ""}" data-id="${esc(r.id)}">
      <div class="g-label" onclick="onGanttRowClick('${esc(r.id)}')"
           oncontextmenu="onGanttContext(event,'${esc(r.id)}')" title="${esc(r.client)}">
        <div class="g-id">${esc(r.id)}</div>
        <div class="g-client">${esc(r.client)}</div>
      </div>
      <div class="g-track" data-row="${i}">
        ${monthLines}${todayHtml}
        <span class="g-date g-date-s" data-i="${i}">${md(g.start)}</span>
        <div class="g-bar ${cls}${g.provisional ? " prov" : ""}${locked ? " locked" : ""}"
             data-i="${i}" style="left:${l}%;width:${w}%" title="${esc(r.content)}">
          ${locked ? "" : `<span class="g-hdl g-hdl-l" data-i="${i}"></span>`}
          <span class="g-title">${esc(r.content)}</span>
          ${locked ? "" : `<span class="g-hdl g-hdl-r" data-i="${i}"></span>`}
        </div>
        <span class="g-date g-date-e" data-i="${i}">${md(g.end)}</span>
        ${deliverHtml}
      </div>
    </div>${schedExpandHtml(r, g, { t0, totalDays, monthLines, todayHtml })}`;
  }).join("");

  return `
  <div class="gantt-toolbar">
    <button class="term-btn" onclick="shiftGanttTerm(-1)">◀</button>
    <span class="term-label">${esc(termLabel(ganttTerm))}</span>
    <button class="term-btn" onclick="shiftGanttTerm(1)">▶</button>
    <span class="term-bar-sep"></span>
    <label class="hours-toggle" title="完了した案件をガントから隠す">
      <input type="checkbox" id="gantt-hidedone-cb" ${ganttHideDone ? "checked" : ""}
        onchange="toggleGanttHideDone(this)"> 完了除く
    </label>
    <label class="hours-toggle" title="展開したタスク行に担当者の休みを重ねる">
      <input type="checkbox" id="gantt-leave-cb" ${schedShowLeave ? "checked" : ""}
        onchange="toggleSchedLeave(this)"> 休み
    </label>
    <span class="g-sp"></span>
    <div class="g-zoom">
      ${[12, 6, 3, 1].map(m =>
        `<button class="${m === ganttZoom ? "active" : ""}" onclick="setGanttZoom(${m})">${m}ヶ月</button>`).join("")}
    </div>
  </div>
  <div class="gantt-wrap" id="gantt-wrap">
    <div class="gantt-inner" style="width:${widthPct}%">
      <div class="g-row g-headrow">
        <div class="g-label g-corner">案件</div>
        <div class="g-headstack">
          ${calRows.map((r, i) => `<div class="g-track g-head cal-row" style="height:${r.h}px">${r.html}${
            i === 0 ? '<span class="g-pan-hint">← ドラッグで期間移動 ／ 案件名クリックでタスク展開 →</span>' : ""
          }</div>`).join("")}
        </div>
      </div>
      ${rows || `<div class="g-empty">受注確定済みの案件がありません（確認中で受注→受注タブで最終登録すると表示されます）</div>`}
    </div>
  </div>
  ${schedExpandListHtml()}
  <div class="gantt-legend">
    <span><i class="g-sw order"></i>受注（開始待ち）</span>
    <span><i class="g-sw work"></i>受託中</span>
    <span><i class="g-sw done"></i>完了</span>
    <span class="g-dv">▲ 納品日</span>
    <span class="g-hint">バー両端＝期日変更／本体＝期間移動（点線は開始日・納品日が未確定の仮表示）</span>
  </div>`;
}

/* --- ドラッグ操作（バー編集＋背景パン） --- */
function setupGantt() {
  const wrap = document.getElementById("gantt-wrap");
  if (!wrap) return;
  const items = ganttRecords();
  const t0 = termStartDate(ganttTerm);
  const totalDays = Math.round((termEndDate(ganttTerm) - t0) / DAY_MS);

  let drag = null;   // {mode:'move'|'l'|'r'|'pan', i, startX, origStart, origEnd, pxPerDay, bar, scrollLeft}

  wrap.addEventListener("pointerdown", e => {
    const hdl = e.target.closest(".g-hdl");
    const bar = e.target.closest(".g-bar");
    const track = e.target.closest(".g-track");
    if (bar && !bar.classList.contains("locked")) {
      const narrow = bar.getBoundingClientRect().width < 30;   // 細いバーはハンドル誤爆防止で移動のみ
      if (hdl && !narrow) {
        startBarDrag(e, hdl.classList.contains("g-hdl-l") ? "l" : "r", Number(hdl.dataset.i), bar);
      } else {
        startBarDrag(e, "move", Number(bar.dataset.i), bar);
      }
    } else if (track || e.target.closest(".g-head")) {
      drag = { mode: "pan", startX: e.clientX, scrollLeft: wrap.scrollLeft };
      wrap.classList.add("panning");
      wrap.setPointerCapture && wrap.setPointerCapture(e.pointerId);
    }
  });

  function startBarDrag(e, mode, i, bar) {
    e.preventDefault();
    const track = bar.parentElement;
    const pxPerDay = track.getBoundingClientRect().width / totalDays;
    drag = {
      mode, i, bar, track, pxPerDay,
      startX: e.clientX,
      origStart: new Date(items[i].start),
      origEnd: new Date(items[i].end),
    };
    bar.classList.add("dragging");
    wrap.setPointerCapture && wrap.setPointerCapture(e.pointerId);
  }

  wrap.addEventListener("pointermove", e => {
    if (!drag) return;
    if (drag.mode === "pan") {
      wrap.scrollLeft = drag.scrollLeft - (e.clientX - drag.startX);
      return;
    }
    const dayDelta = Math.round((e.clientX - drag.startX) / drag.pxPerDay);
    let ns = new Date(drag.origStart), ne = new Date(drag.origEnd);
    if (drag.mode === "move") {
      ns = new Date(ns.getTime() + dayDelta * DAY_MS);
      ne = new Date(ne.getTime() + dayDelta * DAY_MS);
    } else if (drag.mode === "l") {
      ns = new Date(ns.getTime() + dayDelta * DAY_MS);
      if (ns > ne) ns = new Date(ne);
    } else {
      ne = new Date(ne.getTime() + dayDelta * DAY_MS);
      if (ne < ns) ne = new Date(ns);
    }
    drag.curStart = ns; drag.curEnd = ne;
    const l = ((ns - t0) / DAY_MS) / totalDays * 100;
    const w = Math.max(((ne - ns) / DAY_MS + 1) / totalDays * 100, 0.7);
    drag.bar.style.left = l + "%";
    drag.bar.style.width = w + "%";
    const row = drag.track;
    const sEl = row.querySelector(`.g-date-s[data-i="${drag.i}"]`);
    const eEl = row.querySelector(`.g-date-e[data-i="${drag.i}"]`);
    if (sEl) sEl.textContent = md(ns);
    if (eEl) eEl.textContent = md(ne);
  });

  async function finishDrag() {
    if (!drag) return;
    const d = drag; drag = null;
    wrap.classList.remove("panning");
    if (d.mode === "pan") return;
    d.bar.classList.remove("dragging");
    if (!d.curStart && !d.curEnd) { positionGanttDates(); return; }
    const rec = items[d.i].rec;
    // 開始日→AH(workStart)、完了予定日→AI(dueDate)。納品日(N)は連動しない
    rec.workStart = d.curStart || d.origStart;
    rec.dueDate = d.curEnd || d.origEnd;
    try {
      await writeScheduleDates(rec);   // ← 軽量書き込みに変更（AH/AIのみ）
    } catch (err) {
      console.warn("スケジュール保存に失敗:", err);
    }
    renderSched();
  }
  wrap.addEventListener("pointerup", finishDrag);
  wrap.addEventListener("pointercancel", finishDrag);

  positionGanttDates();

  // 初期スクロール: 今日（期外なら最初のバー）を中央付近に
  if (ganttZoom < 12) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let ratio = null;
    if (today >= t0 && today < termEndDate(ganttTerm)) {
      ratio = ((today - t0) / DAY_MS) / totalDays;
    } else if (items.length) {
      ratio = ((items[0].start - t0) / DAY_MS) / totalDays;
    }
    if (ratio != null) {
      const inner = wrap.querySelector(".gantt-inner");
      wrap.scrollLeft = Math.max(inner.scrollWidth * ratio - wrap.clientWidth / 2, 0);
    }
  }
}

/* バー両端のm/dラベルをバー位置に追従させる */
function positionGanttDates() {
  document.querySelectorAll("#gantt-wrap .g-bar").forEach(bar => {
    const i = bar.dataset.i;
    const track = bar.parentElement;
    const s = track.querySelector(`.g-date-s[data-i="${i}"]`);
    const e = track.querySelector(`.g-date-e[data-i="${i}"]`);
    const l = parseFloat(bar.style.left), w = parseFloat(bar.style.width);
    if (s) { s.style.left = `calc(${l}% - 4px)`; }
    if (e) { e.style.left = `calc(${l + w}% + 4px)`; }
  });
}