/* ============================================================
 * 週次状況アドイン weekly.js
 * ------------------------------------------------------------
 * 「営業報告」シートと「wbs」シートから、選択した報告期間
 * （月〜金）に該当する状況記述（m/d・mm/dd・yyyy/mm/dd の
 * 日付で始まる行ブロック）を抽出し、取引先別に表示する。
 *
 * 抽出元:
 *   営業報告: P列(進捗状況)・V列(商談状況)・W列(確認状況)
 *   wbs     : O列(備考)の日付付き記述＋実績開始日/実績終了日
 *
 * 表示と同時に「報告」シートへタンキング（蓄積）する。
 * 同一報告期間の行は洗い替え（削除→書き直し）で重複を防ぐ。
 * ============================================================ */

const APP_VERSION = "rev_20260713_a";
const EIGYO_SHEET = "営業報告";
const WBS_SHEET = "wbs";
const REPORT_SHEET = "報告";
const MAX_ROWS = 1000;

/* ---------- 設定（必要に応じて編集） ---------- */
/* wbs 小分類の除外一覧（社内系など。ここに含まれる小分類は表示しない） */
const EXCLUDE_SUBCATS = ["事業部", "経営", "経理", "総務", "改善", "引継ぎ", "#"];
/* wbs 大分類の除外一覧 */
const EXCLUDE_CATEGORIES = [];
/* wbs 小分類 → 営業報告の取引先名への名寄せ（同一取引先としてまとめる） */
const CLIENT_ALIAS = { "柿本": "kakimoto arms" };
/* true にすると土日の記述も同じ週に含める（表示ラベルは月〜金のまま） */
const INCLUDE_WEEKEND = false;

/* 報告シートの列見出し */
const REPORT_COLUMNS = ["週開始日", "週終了日", "ソース", "取引先", "件名", "担当者", "状態", "状況（週次要約）", "登録日時"];

/* ---------- 状態 ---------- */
let weekStart = mondayOf(new Date());   // 表示中の週の月曜日
let eigyoRows = [];                      // 営業報告の生データ
let wbsRows = [];                        // wbs の生データ
let items = [];                          // 抽出結果（表示中の週）
let demoMode = false;
let filters = { client: "", source: "" };

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
          appName: "週次状況",
          version: APP_VERSION,
          currentId: "weekly",                    // menu.json の id と一致で強調
          menuUrl: COMMON_BASE + "/menu.json",
          localItems: [
            { section: "操作" },
            { label: "再読み込み", icon: "🔄", onClick: () => init() },
            { label: "今週へ移動", icon: "📅", onClick: () => gotoThisWeek() },
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
    alert("メニューの読み込みに失敗しました。通信環境をご確認ください。");
  });
}

/* ============================================================
   起動
   ============================================================ */
if (window.Office) {
  Office.onReady(() => init());
} else {
  window.addEventListener("DOMContentLoaded", () => init());
}

async function init() {
  document.getElementById("version-label").textContent = APP_VERSION;
  bindStaticUI();
  await loadAll();
  await refreshWeek();
}

function bindStaticUI() {
  document.getElementById("filter-client").addEventListener("change", e => {
    filters.client = e.target.value; renderBody();
  });
  document.getElementById("filter-source").addEventListener("change", e => {
    filters.source = e.target.value; renderBody();
  });
}

/* ============================================================
   週の操作
   ============================================================ */
function mondayOf(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const dow = (x.getDay() + 6) % 7; // 月曜=0
  x.setDate(x.getDate() - dow);
  return x;
}
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function weekEndDisp() { return addDays(weekStart, 4); }                       // 表示上の終了日（金）
function weekEndReal() { return addDays(weekStart, INCLUDE_WEEKEND ? 6 : 4); } // 抽出上の終了日

function moveWeek(n) { weekStart = addDays(weekStart, n * 7); refreshWeek(); }
function gotoThisWeek() { weekStart = mondayOf(new Date()); refreshWeek(); }
function jumpToDate(v) {
  if (!v) return;
  weekStart = mondayOf(new Date(v + "T00:00:00"));
  refreshWeek();
}

async function refreshWeek() {
  extractWeek();
  renderWeekBar();
  renderClientFilter();
  renderBody();
  await tankToReport();
}

/* ============================================================
   Excel 読み込み
   ============================================================ */
async function loadAll() {
  if (!window.Office || !window.Excel) {
    loadDemo();
    document.getElementById("demo-badge").style.display = "";
    return;
  }
  try {
    await Excel.run(async ctx => {
      // 営業報告
      const es = ctx.workbook.worksheets.getItem(EIGYO_SHEET);
      const eUsed = es.getUsedRange(true);
      eUsed.load("rowCount");
      // wbs
      const ws = ctx.workbook.worksheets.getItem(WBS_SHEET);
      const wUsed = ws.getUsedRange(true);
      wUsed.load("values");
      await ctx.sync();

      const lastRow = Math.min(Math.max(eUsed.rowCount, 1), MAX_ROWS);
      if (lastRow >= 2) {
        const rng = es.getRange(`A2:AF${lastRow}`);
        rng.load("values");
        await ctx.sync();
        eigyoRows = rng.values;
      } else {
        eigyoRows = [];
      }
      // wbs はカンバンアドインと同じ読み方（11行目以降がデータ）
      wbsRows = (wUsed.values || []).slice(10);
    });
    demoMode = false;
  } catch (e) {
    console.warn("Excel読込に失敗。デモモードで起動します。", e);
    loadDemo();
  }
  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
}

/* ============================================================
   日付付き記述ブロックの抽出
   ------------------------------------------------------------
   「7/8 …」「07/02 …」「2026/07/09 …」「8/28（火）…」のように
   日付で始まる行をブロックの先頭とみなし、次の日付行までを
   1ブロックとして扱う。
   ============================================================ */
const DATE_LINE_RE = /^\s*(?:(\d{4})[\/年])?(\d{1,2})[\/月](\d{1,2})日?(?![\d\/])/;

function parseBlocks(text) {
  if (!text) return [];
  const out = [];
  let cur = null;
  String(text).split(/\r?\n/).forEach(line => {
    const m = line.match(DATE_LINE_RE);
    const mm = m ? Number(m[2]) : 0, dd = m ? Number(m[3]) : 0;
    if (m && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      if (cur) out.push(cur);
      cur = { y: m[1] ? Number(m[1]) : null, m: mm, d: dd, lines: [line.trim()] };
    } else if (cur && line.trim()) {
      cur.lines.push(line.trim());
    }
  });
  if (cur) out.push(cur);
  return out;
}

/* 年なし日付（m/d）の年を報告週から推定する */
function resolveBlockDate(b) {
  if (b.y) return new Date(b.y, b.m - 1, b.d);
  let dt = new Date(weekStart.getFullYear(), b.m - 1, b.d);
  const diff = dt - weekStart;
  const HALF_YEAR = 183 * 86400000;
  if (diff > HALF_YEAR) dt = new Date(dt.getFullYear() - 1, b.m - 1, b.d);
  else if (diff < -HALF_YEAR) dt = new Date(dt.getFullYear() + 1, b.m - 1, b.d);
  return dt;
}

function blocksInWeek(text) {
  const s = weekStart, e = weekEndReal();
  return parseBlocks(text)
    .map(b => ({ date: resolveBlockDate(b), text: b.lines.join("\n") }))
    .filter(b => b.date >= s && b.date <= e)
    .sort((a, b) => b.date - a.date); // 新しい順
}

/* ============================================================
   週次データの抽出
   ============================================================ */
function extractWeek() {
  items = [];

  /* --- 営業報告: P(15)進捗状況 / V(21)商談状況 / W(22)確認状況 --- */
  eigyoRows.forEach(r => {
    if (!r || (!r[0] && !r[14])) return;
    const blocks = [
      ...blocksInWeek(r[15]),
      ...blocksInWeek(r[21]),
      ...blocksInWeek(r[22]),
    ].sort((a, b) => b.date - a.date);
    if (!blocks.length) return;
    items.push({
      source: "営業報告",
      id: str(r[0]),
      client: str(r[1]) || "（取引先未設定）",
      subject: str(r[14]),
      owner: str(r[7]),
      status: str(r[4]),
      summary: dedupeBlocks(blocks).map(b => b.text).join("\n"),
    });
  });

  /* --- wbs: O(14)備考 ＋ 実績開始(17)/実績終了(18) --- */
  wbsRows.forEach(row => {
    if (!row || !row[25] || row[19] === "-") return;           // タイトル無し・除外行
    const cat = str(row[0]), sub = str(row[1]);
    if (EXCLUDE_CATEGORIES.includes(cat) || EXCLUDE_SUBCATS.includes(sub)) return;

    const blocks = blocksInWeek(row[14]);
    const aStart = toDate(row[17]), aEnd = toDate(row[18]);
    const events = [];
    const s = weekStart, e = weekEndReal();
    if (aStart && aStart >= s && aStart <= e) events.push({ date: aStart, text: `${md(aStart)} 着手` });
    if (aEnd && aEnd >= s && aEnd <= e) events.push({ date: aEnd, text: `${md(aEnd)} 完了` });
    const all = [...events, ...blocks].sort((a, b) => b.date - a.date);
    if (!all.length) return;

    items.push({
      source: "WBS",
      id: str(row[24]),
      client: CLIENT_ALIAS[sub] || sub || "（分類未設定）",
      subject: str(row[25]),
      owner: str(row[13]),
      status: aEnd ? "完了" : (aStart ? "対応中" : "未着手"),
      summary: dedupeBlocks(all).map(b => b.text).join("\n"),
    });
  });
}

/* 同一テキストのブロック重複を除去 */
function dedupeBlocks(blocks) {
  const seen = new Set();
  return blocks.filter(b => {
    const k = b.text;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

/* ============================================================
   描画
   ============================================================ */
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
function fmtMD(d) { return `${d.getMonth() + 1}/${d.getDate()}（${WEEKDAYS[d.getDay()]}）`; }

function renderWeekBar() {
  document.getElementById("week-label").textContent =
    `${weekStart.getFullYear()}年 ${fmtMD(weekStart)} 〜 ${fmtMD(weekEndDisp())}`;
  document.getElementById("week-picker").value = fmtDateInput(weekStart);
}

function renderClientFilter() {
  const sel = document.getElementById("filter-client");
  const cur = filters.client;
  const clients = [...new Set(items.map(i => i.client))].sort((a, b) => a.localeCompare(b, "ja"));
  sel.innerHTML = `<option value="">（取引先: 全て）</option>` +
    clients.map(c => `<option value="${esc(c)}"${c === cur ? " selected" : ""}>${esc(c)}</option>`).join("");
  if (cur && !clients.includes(cur)) { filters.client = ""; sel.value = ""; }
}

function renderBody() {
  const list = items.filter(i =>
    (!filters.client || i.client === filters.client) &&
    (!filters.source || i.source === filters.source));

  const clients = [...new Set(list.map(i => i.client))].sort((a, b) => a.localeCompare(b, "ja"));
  document.getElementById("week-summary").textContent =
    list.length ? `取引先 ${clients.length} 件 ／ 案件 ${list.length} 件` : "";

  const box = document.getElementById("weekly-container");
  if (!list.length) {
    box.innerHTML = `<div class="empty-note">この報告期間に該当する状況記述はありません。</div>`;
    return;
  }

  box.innerHTML = clients.map(c => {
    const rows = list.filter(i => i.client === c)
      .sort((a, b) => a.source === b.source ? a.subject.localeCompare(b.subject, "ja") : (a.source === "営業報告" ? -1 : 1));
    return `
      <div class="client-section">
        <div class="client-head">${esc(c)} <span class="client-count">${rows.length}件</span></div>
        ${rows.map(i => `
          <div class="item-card">
            <div class="item-head">
              <span class="src-badge ${i.source === "WBS" ? "src-wbs" : "src-eigyo"}">${i.source}</span>
              <span class="item-subject">${esc(i.subject)}</span>
            </div>
            <div class="item-meta">
              ${i.id ? `<span>${esc(i.id)}</span>` : ""}
              ${i.owner ? `<span>担当: ${esc(i.owner)}</span>` : ""}
              ${i.status ? `<span class="st-chip">${esc(i.status)}</span>` : ""}
            </div>
            <div class="item-summary">${esc(i.summary)}</div>
          </div>`).join("")}
      </div>`;
  }).join("");
}

/* ============================================================
   報告シートへのタンキング（同一週は洗い替え）
   ============================================================ */
async function tankToReport() {
  const stateEl = document.getElementById("save-state");
  if (demoMode) { stateEl.textContent = "DEMO（保存なし）"; return; }
  try {
    const wkStartStr = fmtDate(weekStart);
    const wkEndStr = fmtDate(weekEndDisp());
    const now = new Date();
    const nowStr = `${fmtDate(now)} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
    const newRows = items.map(i =>
      [wkStartStr, wkEndStr, i.source, i.client, i.subject, i.owner, i.status, i.summary, nowStr]);

    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      let sheet = sheets.items.find(s => s.name === REPORT_SHEET)
        ? sheets.getItem(REPORT_SHEET)
        : sheets.add(REPORT_SHEET);

      // 既存データを読み、同一週の行を除外して残す
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("values, rowCount, isNullObject");
      await ctx.sync();

      const wkKey = weekStart.getTime();
      const kept = [];
      ((used.isNullObject ? [] : used.values) || []).forEach((r, idx) => {
        if (idx === 0) return; // ヘッダー
        if (!r[0] && !r[3] && !r[4]) return;
        const d = toDate(r[0]);
        if (d && mondayOf(d).getTime() === wkKey) return; // 同一週 → 洗い替え対象
        kept.push(r.slice(0, REPORT_COLUMNS.length).map(v => v == null ? "" : v));
      });

      // クリアして書き直し（ヘッダー＋既存他週＋今週）
      if (!used.isNullObject) sheet.getUsedRange(true).clear(Excel.ClearApplyTo.contents);
      const hdr = sheet.getRange("A1:I1");
      hdr.values = [REPORT_COLUMNS];
      hdr.format.fill.color = "#44546A";
      hdr.format.font.color = "#FFFFFF";
      hdr.format.font.bold = true;

      const all = [...kept, ...newRows];
      if (all.length) {
        const rng = sheet.getRange(`A2:I${all.length + 1}`);
        rng.values = all;
        rng.format.verticalAlignment = "Top";
        // 週開始日・週終了日の表示形式
        sheet.getRange(`A2:B${all.length + 1}`).numberFormatLocal =
          all.map(() => ["yyyy/mm/dd", "yyyy/mm/dd"]);
      }
      await ctx.sync();
    });
    stateEl.textContent = `報告シートへ保存済み（${newRows.length}件） ✓`;
    stateEl.classList.remove("err");
  } catch (e) {
    console.warn("報告シートへの保存に失敗しました。", e);
    stateEl.textContent = "報告シートへの保存に失敗";
    stateEl.classList.add("err");
  }
}

/* ============================================================
   デモデータ（Excel外で開いたとき用）
   ============================================================ */
function loadDemo() {
  demoMode = true;
  const mon = mondayOf(new Date());
  const d1 = `${mon.getMonth() + 1}/${mon.getDate()}`;
  const wed = addDays(mon, 2);
  const d3 = `${wed.getMonth() + 1}/${wed.getDate()}`;
  eigyoRows = [
    ["KM-01", "kakimoto arms", "", "見積り", "見積中", "", "", "小川", "佐竹様", "中",
      8, 480000, "", "", "ネット予約でフリースタッフを選択できるようにしたい",
      `${d3} 見積作成のための調査を実施\n${d1} 要件確認の打ち合わせ`, "", "", "", "", "", "", "",
      "", null, null, "", "小川", "", "", "", ""],
    ["AG-01", "アサヒグラント", "", "見積り", "見積中", "", "", "紺谷", "川野様", "高",
      5, 280000, "", "", "インフォマートデータ交換の仕様変更",
      `${d1} インフォマートと打ち合わせ日程を調整`, "", "", "", "", "", "", "",
      "", null, null, "", "紺谷", "", "", "", ""],
  ];
  const wbsDemo = new Array(26).fill("");
  wbsDemo[0] = "開発"; wbsDemo[1] = "柿本"; wbsDemo[13] = "小川";
  wbsDemo[14] = `＜状況＞\n${d1} 新サーバーのキッティングを完了`;
  wbsDemo[24] = "開発柿本デモ"; wbsDemo[25] = "青山店サーバーキッティング";
  wbsRows = [wbsDemo];
}

/* ============================================================
   ユーティリティ
   ============================================================ */
function str(v) { return v == null ? "" : String(v).trim(); }
function toDate(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400000));
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function fmtDate(d) { return d ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : ""; }
function md(d) { return d ? `${d.getMonth() + 1}/${d.getDate()}` : ""; }
function fmtDateInput(d) {
  if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
