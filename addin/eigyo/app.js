/* ============================================================
 * 営業報告アドイン app.js（ステージタブ版 rev_c）
 * ------------------------------------------------------------
 * 対象シート: 「営業報告」（1案件1行、ヘッダー行=1行目）
 *   A:ID B:取引先 C:No D:種別 E:状況 F:発生日 G:完了日 H:担当者
 *   I:窓口 J:優先度 K:見積工数 L:見積金額 M:受注区分 N:納品日
 *   O:問合せ・提案内容 P:対応状況（経緯） Q:備考 R:要確認メモ
 *   S:区分(問合せ/改修) T:対応開始日 U:見積根拠 V:商談状況
 *   W:確認状況 X:計上日 Y:最終工数 Z:最終価格 AA:受注条件
 * 顧客マスタ: 「顧客マスタ」シート（無ければ自動作成）
 * ============================================================ */

const APP_VERSION = "rev_20260710_c";
const SHEET_NAME = "営業報告";
const CUST_SHEET = "顧客マスタ";
const MAX_ROWS = 500;
const TAX_RATE = 0.10;

/* カンバンのドラッグ＆ドロップ制御（true にすると有効化） */
const ENABLE_KANBAN_DND = false;

/* ---------- ワークフロー定義 ---------- */
const WORKFLOWS = {
  "保守対応":     { steps: ["新規", "対応中"],                     terminals: ["完了"] },
  "瑕疵対応":     { steps: ["新規", "対応中"],                     terminals: ["完了"] },
  "見積り":       { steps: ["新規", "見積中", "確認中"],           terminals: ["受注", "失注"] },
  "プリセールス": { steps: ["新規", "検討中", "商談中", "確認中"], terminals: ["受注", "失注"] },
  "調整":         { steps: ["新規", "対応中"],                     terminals: ["完了"] },
};
const TYPES = Object.keys(WORKFLOWS);
const HOLD = "保留";
const QUOTE_TYPES = ["見積り", "プリセールス"];

/* ステージタブ定義（編集モーダル） */
function stageTabsOf(type) {
  if (type === "見積り") return ["起票", "見積中", "確認中", "受注"];
  if (type === "プリセールス") return ["起票", "検討中", "商談中", "確認中", "受注"];
  return ["起票", "対応中"];
}
function firstStageOf(type) {
  if (type === "見積り") return "見積中";
  if (type === "プリセールス") return "検討中";
  return "対応中";
}

/* 旧ステータス表記の読み替え */
const LEGACY_STATUS = {
  "未着手": "新規", "検討中※旧": "検討中",
  "作成中": "検討中",
  "見積作成中": "見積中", "見積提出済み": "確認中",
  "調整中": "対応中",
  "完了(受注)": "受注", "完了(失注)": "失注",
};

const EXT_HEADERS = ["区分", "対応開始日", "見積根拠", "商談状況", "確認状況", "計上日", "最終工数", "最終価格", "受注条件"];

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
let filters = { q: "", type: "", status: [], client: "", owner: "" };

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
  renderFilters();
  renderCurrentPane();
}

function bindStaticUI() {
  const input = document.getElementById("search-input");
  input.addEventListener("input", () => { filters.q = input.value.trim(); renderCurrentPane(); });
  document.getElementById("search-clear").addEventListener("click", () => {
    input.value = ""; filters.q = ""; renderCurrentPane();
  });
  ["type", "client", "owner"].forEach(k => {
    document.getElementById("filter-" + k).addEventListener("change", e => {
      filters[k] = e.target.value; renderCurrentPane();
    });
  });
  // 状態マルチセレクトの外側クリックで閉じる
  document.addEventListener("click", e => {
    if (!e.target.closest("#ms-status")) {
      const dd = document.getElementById("status-dd");
      if (dd) dd.style.display = "none";
    }
  });
}

function clearFilters() {
  filters = { q: "", type: "", status: [], client: "", owner: "" };
  document.getElementById("search-input").value = "";
  renderFilters();
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
      const hdr = sheet.getRange("S1:AA1");
      hdr.load("values");
      await ctx.sync();
      if (!hdr.values[0][0]) {
        hdr.values = [EXT_HEADERS];
        hdr.format.fill.color = "#44546A";
        hdr.format.font.color = "#FFFFFF";
      }
      const rng = sheet.getRange(`A2:AA${MAX_ROWS}`);
      rng.load("values");
      await ctx.sync();
      records = parseRows(rng.values);
    });
    await ensureCustomerSheet();
    demoMode = false;
  } catch (e) {
    console.warn("Excel読込に失敗。デモモードで起動します。", e);
    loadDemo();
  }
  document.getElementById("demo-badge").style.display = demoMode ? "" : "none";
}

function parseRows(values) {
  const out = [];
  values.forEach((r, i) => {
    if (!r[0] && !r[14]) return;
    out.push({
      row: i + 2,
      id: str(r[0]), client: str(r[1]), no: r[2],
      type: str(r[3]),
      status: normalizeStatus(str(r[4])),
      occur: toDate(r[5]), done: toDate(r[6]),
      owner: str(r[7]), contact: str(r[8]), priority: str(r[9]),
      hours: numOrNull(r[10]), amount: numOrNull(r[11]),
      order: str(r[12]), deliver: toDate(r[13]),
      content: str(r[14]), progress: str(r[15]), note: str(r[16]), memo: str(r[17]),
      kind: str(r[18]), stageStart: toDate(r[19]), basis: str(r[20]), deal: str(r[21]),
      confirm: str(r[22]), book: toDate(r[23]), finalHours: numOrNull(r[24]),
      finalAmount: numOrNull(r[25]), terms: str(r[26]),
    });
  });
  return out;
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
  ]];
}

async function writeRecord(rec) {
  if (demoMode) {
    const i = records.findIndex(r => r.id === rec.id);
    if (i >= 0) records[i] = rec; else { rec.row = 0; records.push(rec); }
    return;
  }
  await Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(SHEET_NAME);
    let row = rec.row;
    if (!row) {
      const colA = sheet.getRange(`A2:A${MAX_ROWS}`);
      colA.load("values");
      await ctx.sync();
      row = 2;
      for (let i = 0; i < colA.values.length; i++) {
        if (!colA.values[i][0]) { row = i + 2; break; }
        row = i + 3;
      }
      rec.row = row;
    }
    const rng = sheet.getRange(`A${row}:AA${row}`);
    rng.values = recToRow(rec);
    ["F", "G", "N", "T", "X"].forEach(c => sheet.getRange(`${c}${row}`).numberFormat = [["yyyy/m/d"]]);
    ["L", "Z"].forEach(c => sheet.getRange(`${c}${row}`).numberFormat = [["#,##0"]]);
    await ctx.sync();
  });
  const i = records.findIndex(r => r.row === rec.row);
  if (i >= 0) records[i] = rec; else records.push(rec);
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
  if (rec.status === "失注" || rec.status === "完了") return true;
  if (rec.status === "受注") return true; // 受注は受注タブ完了登録後にのみ設定される
  return false;
}

function allowedTransitions(rec) {
  const wf = WORKFLOWS[rec.type];
  if (!wf) return [];
  const chain = wf.steps;
  const cur = rec.status;
  const res = [];
  if (cur === HOLD) { chain.forEach(s => res.push(s)); return res; }
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
  ["list", "kanban", "agg"].forEach(t => {
    document.getElementById("pane-" + t).style.display = (t === tab) ? "" : "none";
  });
  document.getElementById("filter-bar").style.display =
    (tab === "list" || tab === "kanban") ? "" : "none";
  // カンバンでは種別・状態のフィルタは非表示（種別はカンバン内のセグメントで選択）
  const isKanban = tab === "kanban";
  document.getElementById("filter-type").style.display = isKanban ? "none" : "";
  document.getElementById("ms-status").style.display = isKanban ? "none" : "";
  renderCurrentPane();
}
function activeTab() { return document.querySelector(".tab.active").dataset.tab; }
function renderCurrentPane() {
  const t = activeTab();
  if (t === "list") renderList();
  else if (t === "kanban") renderKanban();
  else if (t === "agg") renderAgg();
}

/* ============================================================
   フィルタ（状態は複数選択）
   ============================================================ */
function renderFilters() {
  fillSelect("filter-type", ["（種別: 全て）", ...TYPES], filters.type);
  renderStatusMulti();
  const clients = [...new Set(records.map(r => r.client).filter(Boolean))];
  fillSelect("filter-client", ["（取引先: 全て）", ...clients], filters.client);
  const owners = [...new Set(records.flatMap(r => splitOwners(r.owner)))];
  fillSelect("filter-owner", ["（担当者: 全て）", ...owners], filters.owner);
}
function splitOwners(s) { return str(s).split(/[、,\s]+/).filter(Boolean); }
function fillSelect(id, options, selected) {
  const el = document.getElementById(id);
  el.innerHTML = options.map((o, i) =>
    `<option value="${i === 0 ? "" : esc(o)}"${o === selected ? " selected" : ""}>${esc(o)}</option>`).join("");
}

/* 状態マルチセレクト */
function renderStatusMulti() {
  const sts = [...new Set(records.map(r => r.status).filter(Boolean))];
  filters.status = filters.status.filter(s => sts.includes(s));
  const btn = document.getElementById("ms-status-btn");
  btn.textContent = filters.status.length
    ? `状態: ${filters.status.length}件選択 ▾`
    : "（状態: 全て）▾";
  const dd = document.getElementById("status-dd");
  dd.innerHTML = sts.map(s => `
    <label class="ms-item">
      <input type="checkbox" value="${esc(s)}" ${filters.status.includes(s) ? "checked" : ""}
        onchange="onStatusCheck(this)">
      <span class="status-pill st-${esc(s)}">${esc(s)}</span>
    </label>`).join("") +
    `<button class="ms-clear" onclick="clearStatusFilter()">選択解除</button>`;
}
function toggleStatusDD(ev) {
  ev.stopPropagation();
  const dd = document.getElementById("status-dd");
  dd.style.display = dd.style.display === "none" ? "" : "none";
}
function onStatusCheck(cb) {
  if (cb.checked) { if (!filters.status.includes(cb.value)) filters.status.push(cb.value); }
  else filters.status = filters.status.filter(s => s !== cb.value);
  document.getElementById("ms-status-btn").textContent = filters.status.length
    ? `状態: ${filters.status.length}件選択 ▾` : "（状態: 全て）▾";
  renderCurrentPane();
}
function clearStatusFilter() {
  filters.status = [];
  renderStatusMulti();
  renderCurrentPane();
}

function filteredRecords() {
  return records.filter(r => {
    if (filters.type && r.type !== filters.type) return false;
    if (filters.status.length && !filters.status.includes(r.status)) return false;
    if (filters.client && r.client !== filters.client) return false;
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
   一覧（優先度列を追加）
   ============================================================ */
function renderList() {
  const cont = document.getElementById("list-container");
  const recs = filteredRecords();
  if (!recs.length) { cont.innerHTML = `<div class="empty-note">条件に一致する案件がありません</div>`; return; }
  let html = "";
  TYPES.forEach(type => {
    const group = recs.filter(r => r.type === type);
    if (!group.length) return;
    html += `<div class="list-group lg-${type}">
      <div class="list-group-head">${esc(type)} <span class="cnt">${group.length}件</span></div>
      <table class="list-table">
        <tr><th>ID</th><th>取引先</th><th>状態</th><th>優先度</th><th>内容</th><th>担当</th><th>発生日</th><th>金額</th></tr>
        ${group.map(r => `
        <tr data-id="${esc(r.id)}" oncontextmenu="onRowContext(event,'${esc(r.id)}')" onclick="openEditModal('${esc(r.id)}')">
          <td class="muted">${esc(r.id)}</td>
          <td>${esc(r.client)}</td>
          <td><span class="status-pill st-${esc(r.status)}">${esc(r.status)}</span>${r.status === "確認中" && r.order === "受注" ? '<span class="mini-flag">受注</span>' : ""}</td>
          <td class="c">${r.priority ? `<span class="pri pri-${esc(r.priority)}">${esc(r.priority)}</span>` : ""}</td>
          <td>${esc(shorten(r.content, 36))}</td>
          <td>${esc(r.owner)}</td>
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

/* ============================================================
   カンバン（DnDはENABLE_KANBAN_DNDで制御）
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
  const recs = filteredRecords().filter(r => r.type === currentKanbanType);
  const dndLane = ENABLE_KANBAN_DND
    ? `ondragover="onLaneDragOver(event)" ondragleave="onLaneDragLeave(event)" ondrop="onLaneDrop(event)"` : "";
  board.innerHTML = lanes.map(st => {
    const cards = recs.filter(r => r.status === st);
    return `<div class="lane" data-status="${esc(st)}" ${dndLane}>
      <div class="lane-head">${esc(st)}<span class="cnt">${cards.length}</span></div>
      <div class="lane-body">
        ${cards.map(r => `
          <div class="card t-${esc(r.type)}" draggable="${ENABLE_KANBAN_DND}" data-id="${esc(r.id)}"
               ${ENABLE_KANBAN_DND ? `ondragstart="onCardDragStart(event)"` : ""}
               oncontextmenu="onRowContext(event,'${esc(r.id)}')"
               onclick="openEditModal('${esc(r.id)}')">
            <div class="cid">${esc(r.id)}｜${esc(r.client)}</div>
            <div class="ctitle">${esc(shorten(r.content, 46))}</div>
            <div class="cmeta">
              <span>${esc(r.owner)}</span>
              ${dispAmount(r) ? `<span>${dispAmount(r)}円</span>` : ""}
              ${r.priority ? `<span>優先:${esc(r.priority)}</span>` : ""}
              ${r.status === "確認中" && r.order === "受注" ? `<span class="mini-flag">受注</span>` : ""}
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
  if (!rec || rec.status === to) return;
  if (!isValidTransition(rec, to)) {
    alert(`「${rec.status}」から「${to}」へは遷移できません。\nワークフロー: ${workflowLabel(rec.type)}`);
    return;
  }
  if (to === "受注") {
    openEditModal(rec.id, "確認中");
    alert("受注は編集画面の「確認中」タブで結果を選択し、「受注」タブで計上日等を入力してください。");
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
function renderStepper(el, type, currentStatus) {
  const wf = WORKFLOWS[type];
  if (!wf) { el.innerHTML = ""; return; }
  const idx = wf.steps.indexOf(currentStatus);
  const termReached = wf.terminals.includes(currentStatus);
  let html = "";
  wf.steps.forEach((s, i) => {
    let cls = "step";
    if (currentStatus != null) {
      if (termReached || i < idx) cls += " done";
      else if (i === idx) cls += " current";
    }
    html += `<div class="${cls}"><div class="dot"><div class="circle">${i + 1}</div><div class="lbl">${esc(s)}</div></div>
      <div class="arrow"></div></div>`;
  });
  const parts = wf.terminals.map(t => {
    let cls = "step";
    if (currentStatus === t) cls += t === "失注" ? " terminal-lose current" : " terminal-win current";
    const mark = t === "受注" ? "○" : t === "失注" ? "×" : "✓";
    return `<div class="${cls}"><div class="dot"><div class="circle">${mark}</div><div class="lbl">${esc(t)}</div></div></div>`;
  });
  html += parts.join(`<div class="step"><div class="branch">or</div></div>`);
  if (currentStatus === HOLD) {
    html += `<div class="step current" style="margin-left:8px"><div class="dot"><div class="circle" style="background:#ed7d31;border-color:#ed7d31;color:#fff">||</div><div class="lbl">保留中</div></div></div>`;
  }
  el.innerHTML = html;
}

/* ============================================================
   新規入力モーダル（窓口は商談中タブで入力するため非表示）
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
function fillOwnerSelect(id, selected) {
  const owners = [...new Set(records.flatMap(r => splitOwners(r.owner)))];
  const el = document.getElementById(id);
  el.innerHTML = `<option value=""></option>` +
    owners.map(o => `<option${o === selected ? " selected" : ""}>${esc(o)}</option>`).join("");
  if (selected && !owners.includes(selected)) {
    el.innerHTML += `<option selected>${esc(selected)}</option>`;
  }
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
  const cust = customers.find(c => c.code === code);
  const rec = {
    row: 0,
    id: nextCaseId(code),
    client: cust.name, no: "",
    type: inputType, status: "新規",
    occur: fromDateInput(document.getElementById("in-occur").value) || new Date(),
    done: null,
    owner: document.getElementById("in-owner").value,
    contact: cust.contact || "",
    priority: document.getElementById("in-priority").value,
    hours: null, amount: null, order: "", deliver: null,
    content,
    progress: "",
    note: document.getElementById("in-note").value,
    memo: "",
    kind: "", stageStart: null, basis: "", deal: "", confirm: "",
    book: null, finalHours: null, finalAmount: null, terms: "",
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
    (["occur","done","deliver","stageStart","book"].includes(k) && v) ? new Date(v) : v);
  editingRec.row = rec.row;
  document.getElementById("ed-title").textContent = `${rec.id}　${rec.client}`;
  document.getElementById("ed-id").value = rec.id;
  document.getElementById("ed-client").value = rec.client;
  const tSel = document.getElementById("ed-type");
  tSel.innerHTML = TYPES.map(t => `<option${t === rec.type ? " selected" : ""}>${esc(t)}</option>`).join("");
  tSel.onchange = () => {
    editingRec.type = tSel.value;
    if (!allStatusesOf(editingRec.type).includes(editingRec.status)) editingRec.status = "新規";
    currentStageTab = defaultStageTab(editingRec);
    refreshEditModal();
  };
  fillOwnerSelect("ed-owner", rec.owner);
  document.getElementById("ed-priority").value = rec.priority;
  document.getElementById("ed-note").value = rec.note;
  document.getElementById("ed-msg").textContent = "";
  currentStageTab = forceTab || defaultStageTab(editingRec);
  refreshEditModal();
  document.getElementById("edit-modal").style.display = "";
}
function closeEditModal() { document.getElementById("edit-modal").style.display = "none"; editingRec = null; }

/* 状態から活性なステージタブを決める（確認中＋受注区分=受注 → 受注タブ） */
function activeStageTab(rec) {
  if (isTerminal(rec)) return null;
  const st = rec.status;
  if (st === "確認中" && rec.order === "受注") return "受注";
  if (st === "新規" || st === HOLD) return firstStageOf(rec.type);
  if (st === "対応中" || st === "見積中" || st === "検討中") return st;
  if (st === "商談中") return "商談中";
  if (st === "確認中") return "確認中";
  return null;
}
function defaultStageTab(rec) {
  return activeStageTab(rec) || stageTabsOf(rec.type)[stageTabsOf(rec.type).length - 1];
}

function refreshEditModal() {
  const rec = editingRec;
  renderStepper(document.getElementById("ed-stepper"), rec.type, rec.status);
  let stLabel = rec.status;
  if (rec.status === "確認中" && rec.order === "受注") stLabel += "（受注・最終登録待ち）";
  if (isTerminal(rec)) stLabel += "（チケット完了）";
  document.getElementById("ed-status").value = stLabel;
  renderStageTabs();
  renderStageBody();
}

function renderStageTabs() {
  const rec = editingRec;
  const tabs = stageTabsOf(rec.type);
  const active = activeStageTab(rec);
  const el = document.getElementById("ed-stage-tabs");
  el.innerHTML = tabs.map(t => {
    const enabled = (t === "起票") || (t === active);
    const isCur = t === currentStageTab;
    return `<button class="stage-tab${isCur ? " current" : ""}${enabled ? "" : " locked"}"
      onclick="setStageTab('${esc(t)}')">${esc(t)}${enabled || t === "起票" ? "" : " 🔒"}</button>`;
  }).join("");
}
function setStageTab(t) { currentStageTab = t; renderStageTabs(); renderStageBody(); }

function renderStageBody() {
  const rec = editingRec;
  const body = document.getElementById("ed-stage-body");
  const active = activeStageTab(rec);
  const t = currentStageTab;
  const dis = (t !== "起票" && t !== active) || (isTerminal(rec) && t !== "起票") ? "disabled" : "";
  const disAll = isTerminal(rec) ? "disabled" : "";

  if (t === "起票") {
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-row"><label>発生日</label><input type="date" id="st-occur" value="${fmtDateInput(rec.occur)}" ${disAll}></div>
        <div class="form-row"><label>窓口</label><input type="text" id="st-contact" value="${esc(rec.contact)}" ${disAll}></div>
      </div>
      <div class="form-row"><label>問合せ・提案内容</label>
        <textarea id="st-content" rows="4" ${disAll}>${esc(rec.content)}</textarea></div>`;
    return;
  }

  if (t === "対応中" || t === "見積中" || t === "検討中") {
    const isQuote = t === "見積中";
    const isHoshu = rec.type === "保守対応";
    const doneLabel = t === "対応中" ? "対応完了（チケットを完了する）"
      : t === "見積中" ? "見積完了（確認中へ進める）"
      : "検討完了（商談中へ進める）";
    const progressLabel = t === "検討中" ? "対応状況（検討・検証・提案作成など）" : "対応状況";
    const startInfo = rec.stageStart ? `<span class="stage-info">開始日: ${fmtDate(rec.stageStart)}</span>`
      : `<span class="stage-info">※対応状況を記入して登録すると「${esc(t)}」に遷移し、開始日を記録します</span>`;
    body.innerHTML = `
      ${startInfo}
      ${isHoshu ? `
      <div class="form-row"><label>区分 <span class="req">必須</span></label>
        <div class="radio-row">
          <label class="radio"><input type="radio" name="st-kind" value="問合せ" ${rec.kind === "問合せ" ? "checked" : ""} ${dis}>問合せ</label>
          <label class="radio"><input type="radio" name="st-kind" value="改修" ${rec.kind === "改修" ? "checked" : ""} ${dis}>改修</label>
        </div>
      </div>` : ""}
      <div class="form-row"><label>${esc(progressLabel)}</label>
        <textarea id="st-progress" rows="4" ${dis}>${esc(rec.progress)}</textarea></div>
      ${isQuote ? `
      <div class="form-grid">
        <div class="form-row"><label>工数（人日）</label><input type="number" step="0.5" id="st-hours" value="${rec.hours ?? ""}" ${dis}></div>
        <div class="form-row"><label>価格（税抜・円）</label><input type="number" step="1000" id="st-amount" value="${rec.amount ?? ""}" ${dis} oninput="updateTaxView()"></div>
      </div>
      <div class="form-row"><label>税込価格（自動計算）</label>
        <input type="text" id="st-tax" readonly class="ro" value="${rec.amount != null ? withTax(rec.amount).toLocaleString() + " 円" : ""}"></div>
      <div class="form-row"><label>根拠</label>
        <textarea id="st-basis" rows="3" ${dis}>${esc(rec.basis)}</textarea></div>` : ""}
      <label class="check-row ${dis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${dis}> ${esc(doneLabel)}
      </label>`;
    return;
  }

  if (t === "商談中") {
    body.innerHTML = `
      <div class="form-row"><label>窓口</label>
        <input type="text" id="st-dcontact" value="${esc(rec.contact)}" ${dis}></div>
      <div class="form-row"><label>商談状況</label>
        <textarea id="st-deal" rows="5" ${dis}>${esc(rec.deal)}</textarea></div>
      <label class="check-row ${dis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${dis}> 商談完了（確認中へ進める）
      </label>`;
    return;
  }

  if (t === "確認中") {
    body.innerHTML = `
      <div class="form-row"><label>確認状況</label>
        <textarea id="st-confirm" rows="4" ${dis}>${esc(rec.confirm)}</textarea></div>
      <div class="form-row"><label>結果</label>
        <div class="radio-row">
          <label class="radio win"><input type="radio" name="st-result" value="受注" ${rec.order === "受注" ? "checked" : ""} ${dis}>受注（受注タブで最終登録）</label>
          <label class="radio lose"><input type="radio" name="st-result" value="失注" ${dis}>失注（チケット完了）</label>
        </div>
      </div>
      <span class="stage-info">※受注を選択して登録してもステータスは「確認中」のまま。「受注」タブでの最終登録で受注確定・チケット完了となります。</span>`;
    return;
  }

  if (t === "受注") {
    const base = rec.finalAmount ?? rec.amount;
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-row"><label>納品日</label><input type="date" id="st-deliver" value="${fmtDateInput(rec.deliver)}" ${dis}></div>
        <div class="form-row"><label>計上日 <span class="req">売上集計に使用</span></label><input type="date" id="st-book" value="${fmtDateInput(rec.book)}" ${dis}></div>
        <div class="form-row"><label>最終工数（人日）</label><input type="number" step="0.5" id="st-fhours" value="${rec.finalHours ?? rec.hours ?? ""}" ${dis}></div>
        <div class="form-row"><label>最終価格（税抜・円）</label><input type="number" step="1000" id="st-famount" value="${base ?? ""}" ${dis} oninput="updateTaxView2()"></div>
      </div>
      <div class="form-row"><label>税込価格（自動計算）</label>
        <input type="text" id="st-tax2" readonly class="ro" value="${base != null ? withTax(base).toLocaleString() + " 円" : ""}"></div>
      <div class="form-row"><label>受注条件（必要に応じて）</label>
        <textarea id="st-terms" rows="3" ${dis}>${esc(rec.terms)}</textarea></div>
      <label class="check-row ${dis ? "off" : ""}">
        <input type="checkbox" id="st-done" ${dis}> この内容で登録し、受注確定・チケットを完了する
      </label>`;
    return;
  }
  body.innerHTML = "";
}
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

  const active = activeStageTab(rec);
  const t = currentStageTab;
  const editable = (t === "起票" && !isTerminal(rec)) || t === active;

  if (editable) {
    if (t === "起票") {
      rec.occur = fromDateInput(document.getElementById("st-occur").value);
      rec.contact = document.getElementById("st-contact").value;
      rec.content = document.getElementById("st-content").value;
    }
    else if (t === "対応中" || t === "見積中" || t === "検討中") {
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
      }
      if ((rec.status === "新規" || rec.status === HOLD) && progress.trim()) {
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
        if (t === "対応中") { applyStatus(rec, "完了"); }
        else if (t === "見積中") {
          if (rec.amount == null) { msg.className = "save-msg err"; msg.textContent = "価格を入力してください"; return; }
          applyStatus(rec, "確認中");
        }
        else if (t === "検討中") { applyStatus(rec, "商談中"); }
      }
    }
    else if (t === "商談中") {
      rec.contact = document.getElementById("st-dcontact").value;
      rec.deal = document.getElementById("st-deal").value;
      const doneChk = document.getElementById("st-done");
      if (doneChk && doneChk.checked) applyStatus(rec, "確認中");
    }
    else if (t === "確認中") {
      rec.confirm = document.getElementById("st-confirm").value;
      const res = document.querySelector('input[name="st-result"]:checked');
      if (res) {
        if (res.value === "受注") {
          // ステータスは確認中のまま、受注区分のみ設定 → 受注タブが活性化
          rec.order = "受注";
        } else {
          applyStatus(rec, "失注"); // チケット完了
        }
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
        if (!rec.book) { msg.className = "save-msg err"; msg.textContent = "計上日を入力してください（売上集計に使用します）"; return; }
        if (rec.finalAmount == null) { msg.className = "save-msg err"; msg.textContent = "最終価格を入力してください"; return; }
        applyStatus(rec, "受注");   // 受注確定
        rec.done = new Date();      // チケット完了
      }
    }
  }

  try {
    await writeRecord(rec);
    msg.textContent = "登録しました";
    renderFilters();
    renderCurrentPane();
    currentStageTab = defaultStageTab(rec);
    refreshEditModal();
    if (isTerminal(rec)) setTimeout(closeEditModal, 700);
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
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
function switchAgg(k) {
  currentAgg = k;
  document.querySelectorAll(".agg-seg .seg").forEach(b => b.classList.toggle("active", b.dataset.agg === k));
  renderAgg();
}
function shiftTerm(d) { currentTerm += d; renderAgg(); }
function termBarHtml() {
  return `<div class="term-bar">
    <button class="term-btn" onclick="shiftTerm(-1)">◀</button>
    <span class="term-label">${esc(termLabel(currentTerm))}</span>
    <button class="term-btn" onclick="shiftTerm(1)">▶</button>
  </div>`;
}
function renderAgg() {
  const cont = document.getElementById("agg-container");
  if (currentAgg === "hoshu") cont.innerHTML = termBarHtml() + renderHoshuAgg();
  else if (currentAgg === "mitsu") cont.innerHTML = renderMitsuAgg();
  else cont.innerHTML = termBarHtml() + renderUriageAgg();
}

/* --- 保守状況: 発生と完了の推移（期の12か月） --- */
function renderHoshuAgg() {
  const months = fiscalMonths(currentTerm);
  const target = records.filter(r => r.type === "保守対応" || r.type === "瑕疵対応");
  const series = {
    "発生": countByMonth(target, "occur", months),
    "完了": countByMonth(target, "done", months),
  };
  const colors = { "発生": "#4472c4", "完了": "#548235" };
  const open = target.filter(r => !isTerminal(r)).length;
  const hoshu = target.filter(r => r.type === "保守対応");
  const kashi = target.filter(r => r.type === "瑕疵対応");
  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${target.length}</div><div class="kl">保守・瑕疵 総件数</div></div>
      <div class="kpi"><div class="kv">${open}</div><div class="kl">未完了件数</div></div>
    </div>
    <div class="agg-card">
      <h3>保守対応・瑕疵対応 月次推移（発生・完了）</h3>
      ${legendHtml(colors)}
      <div class="chart-wrap">${groupedBarChart(months, series, colors)}</div>
    </div>
    <div class="agg-card">
      <h3>月別明細（内訳）</h3>
      <table class="agg-table">
        <tr><th>月</th><th>発生 計</th><th>完了 計</th><th>保守 発生</th><th>保守 完了</th><th>瑕疵 発生</th><th>瑕疵 完了</th></tr>
        ${months.map((m, i) => `<tr><td>${m}</td>
          <td><b>${series["発生"][i]}</b></td><td><b>${series["完了"][i]}</b></td>
          <td>${countByMonth(hoshu, "occur", [m])[0]}</td><td>${countByMonth(hoshu, "done", [m])[0]}</td>
          <td>${countByMonth(kashi, "occur", [m])[0]}</td><td>${countByMonth(kashi, "done", [m])[0]}</td></tr>`).join("")}
      </table>
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

/* --- 見積状況: 状態別（新規→見積中→検討中→商談中→確認中→失注→受注） --- */
const MITSU_ORDER = ["新規", "見積中", "検討中", "商談中", "確認中", "失注", "受注"];
function renderMitsuAgg() {
  const target = records.filter(r => QUOTE_TYPES.includes(r.type));
  const rows = MITSU_ORDER.map(st => {
    const g = target.filter(r => r.status === st);
    return { st, cnt: g.length, amt: g.reduce((a, r) => a + ((r.finalAmount ?? r.amount) || 0), 0) };
  });
  const held = target.filter(r => r.status === HOLD);
  if (held.length) rows.push({ st: HOLD, cnt: held.length, amt: held.reduce((a, r) => a + (r.amount || 0), 0) });
  const totalCnt = rows.reduce((a, r) => a + r.cnt, 0);
  const totalAmt = rows.reduce((a, r) => a + r.amt, 0);
  const pipeline = target.filter(r => !["受注", "失注"].includes(r.status));
  const pipelineAmt = pipeline.reduce((a, r) => a + (r.amount || 0), 0);
  const wonAmt = target.filter(r => r.order === "受注").reduce((a, r) => a + ((r.finalAmount ?? r.amount) || 0), 0);
  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${pipeline.length}</div><div class="kl">進行中案件</div></div>
      <div class="kpi"><div class="kv">${(pipelineAmt / 10000).toLocaleString()}万</div><div class="kl">パイプライン金額</div></div>
      <div class="kpi"><div class="kv">${(wonAmt / 10000).toLocaleString()}万</div><div class="kl">受注確定金額</div></div>
    </div>
    <div class="agg-card">
      <h3>見積り・プリセールス 状態別集計</h3>
      <table class="agg-table">
        <tr><th>状態</th><th>件数</th><th>見積金額合計（税抜）</th></tr>
        ${rows.map(r => `<tr>
          <td><span class="status-pill st-${esc(r.st)}">${esc(r.st)}</span></td>
          <td>${r.cnt}</td><td class="r">${r.amt ? r.amt.toLocaleString() + "円" : "－"}</td></tr>`).join("")}
        <tr class="total"><td>合計</td><td>${totalCnt}</td><td class="r">${totalAmt.toLocaleString()}円</td></tr>
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">※ 受注は最終価格、それ以外は見積金額で集計。確認中で受注区分が付いた案件は「確認中」に含まれます。</p>
    </div>`;
}

/* --- 売上予測: 受注確定（受注区分=受注）の計上日ベース、期の12か月 --- */
function renderUriageAgg() {
  const months = fiscalMonths(currentTerm);
  const won = records.filter(r => QUOTE_TYPES.includes(r.type) && r.order === "受注");
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
  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${won.length}</div><div class="kl">受注案件数</div></div>
      <div class="kpi"><div class="kv">${(total / 10000).toLocaleString()}万</div><div class="kl">受注金額合計</div></div>
      ${noBook ? `<div class="kpi"><div class="kv" style="color:#c00000">${noBook}</div><div class="kl">計上日未入力</div></div>` : ""}
    </div>
    <div class="agg-card">
      <h3>月別売上予測（受注確定・計上日ベース）</h3>
      ${legendHtml(colors)}
      <div class="chart-wrap">${groupedBarChart(months, { "確定売上": vals }, colors, v => (v / 10000) + "万")}</div>
    </div>
    <div class="agg-card">
      <h3>受注案件一覧</h3>
      <table class="agg-table">
        <tr><th>ID</th><th>取引先</th><th>内容</th><th>最終価格</th><th>計上日</th><th>納品日</th><th>状態</th></tr>
        ${won.length ? won.map(r => `<tr>
          <td>${esc(r.id)}</td><td class="l">${esc(r.client)}</td>
          <td class="l">${esc(shorten(r.content, 20))}</td>
          <td class="r">${(r.finalAmount ?? r.amount) != null ? ((r.finalAmount ?? r.amount)).toLocaleString() + "円" : "－"}</td>
          <td>${r.book ? fmtDate(r.book) : '<span style="color:#c00000">未入力</span>'}</td>
          <td>${r.deliver ? fmtDate(r.deliver) : '<span class="muted">－</span>'}</td>
          <td><span class="status-pill st-${esc(r.status)}">${esc(r.status)}</span></td></tr>`).join("")
        : `<tr><td colspan="7" class="muted">受注案件はまだありません</td></tr>`}
      </table>
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
function legendHtml(colors) {
  return `<div class="legend">` + Object.entries(colors).map(([n, c]) =>
    `<span><span class="sw" style="background:${c}"></span>${esc(n)}</span>`).join("") + `</div>`;
}

/* ============================================================
   デモモード
   ============================================================ */
function loadDemo() {
  demoMode = true;
  const d = (y, m, day) => new Date(y, m - 1, day);
  const blank = { kind: "", stageStart: null, basis: "", deal: "", confirm: "", book: null, finalHours: null, finalAmount: null, terms: "" };
  customers = [
    { row: 2, code: "KM", name: "kakimoto arms", contact: "佐竹様", note: "" },
    { row: 3, code: "HN", name: "ハンター製菓", contact: "鈴木様", note: "" },
    { row: 4, code: "AG", name: "アサヒグラント", contact: "川野様", note: "" },
    { row: 5, code: "EX", name: "エキスプレス", contact: "中道様", note: "" },
  ];
  records = [
    { ...blank, row: 2, id: "KM-01", client: "kakimoto arms", no: 1, type: "見積り", status: "見積中", occur: d(2026, 6, 29), done: null, owner: "小川", contact: "佐竹様", priority: "中", hours: 10, amount: null, order: "", deliver: null, content: "ネット予約でフリースタッフを選択できるようにしたい", progress: "調査中", note: "", memo: "", stageStart: d(2026, 7, 1) },
    { ...blank, row: 3, id: "KM-02", client: "kakimoto arms", no: 2, type: "見積り", status: "確認中", occur: d(2026, 6, 18), done: null, owner: "小川", contact: "佐竹様", priority: "", hours: 8, amount: 600000, order: "受注", deliver: d(2026, 7, 17), content: "ネット予約LINEログイン連携", progress: "60万で提示", note: "", memo: "", stageStart: d(2026, 6, 20), confirm: "受注の内諾。最終登録待ち", book: d(2026, 7, 31), finalAmount: 600000, finalHours: 8 },
    { ...blank, row: 4, id: "KM-03", client: "kakimoto arms", no: 3, type: "保守対応", status: "完了", occur: d(2026, 7, 2), done: d(2026, 7, 2), owner: "小川", contact: "西野様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "スタッフ指名予約で店舗が正しく選択されない", progress: "外部サイト側の設定が原因", note: "", memo: "", kind: "問合せ", stageStart: d(2026, 7, 2) },
    { ...blank, row: 5, id: "KM-04", client: "kakimoto arms", no: 4, type: "調整", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", contact: "佐竹様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "会社体制変更に伴うご挨拶のスケジュール調整", progress: "日程調整中", note: "", memo: "", stageStart: d(2026, 7, 3) },
    { ...blank, row: 6, id: "KM-05", client: "kakimoto arms", no: 5, type: "保守対応", status: "対応中", occur: d(2026, 7, 7), done: null, owner: "小川", contact: "中田様", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "メンズ予約時の注意事項表示・メール文面変更", progress: "設定変更で対応可能", note: "", memo: "", kind: "改修", stageStart: d(2026, 7, 8) },
    { ...blank, row: 7, id: "HN-01", client: "ハンター製菓", no: 1, type: "瑕疵対応", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", contact: "鈴木様", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "在庫管理伝票一覧画面バグ対応", progress: "修正済み、次回リリースで反映", note: "", memo: "", stageStart: d(2026, 7, 4) },
    { ...blank, row: 8, id: "HN-02", client: "ハンター製菓", no: 2, type: "プリセールス", status: "商談中", occur: d(2026, 7, 6), done: null, owner: "小川", contact: "柳澤様", priority: "高", hours: null, amount: 2500000, order: "", deliver: null, content: "原価計算の改修", progress: "提案書作成済み", note: "9月本稼働目標", memo: "", stageStart: d(2026, 7, 7), deal: "7/22打ち合わせ予定" },
    { ...blank, row: 9, id: "AG-01", client: "アサヒグラント", no: 1, type: "見積り", status: "確認中", occur: d(2026, 6, 30), done: null, owner: "紺谷", contact: "川野様", priority: "中", hours: 5, amount: 350000, order: "", deliver: null, content: "インフォマートデータ交換の仕様変更", progress: "再見積提出済み", note: "", memo: "", stageStart: d(2026, 7, 1), basis: "設計2人日＋実装2人日＋試験1人日" },
    { ...blank, row: 10, id: "EX-01", client: "エキスプレス", no: 1, type: "見積り", status: "新規", occur: d(2026, 7, 6), done: null, owner: "紺谷", contact: "中道様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "削除した請求書を参照できる機能の見積", progress: "", note: "", memo: "" },
    { ...blank, row: 11, id: "HN-03", client: "ハンター製菓", no: 3, type: "プリセールス", status: "新規", occur: d(2026, 7, 9), done: null, owner: "小川", contact: "", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "加工所日報のモバイル入力の提案", progress: "", note: "", memo: "" },
  ];
}
