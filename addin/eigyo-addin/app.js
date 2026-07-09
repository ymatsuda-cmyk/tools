/* ============================================================
 * 営業報告アドイン app.js
 * ------------------------------------------------------------
 * 対象シート: 「営業報告」（1案件1行、ヘッダー行=1行目）
 *   A:ID B:取引先 C:No D:種別 E:状況 F:発生日 G:完了日 H:担当者
 *   I:窓口 J:優先度 K:見積工数 L:見積金額 M:受注区分 N:納品日
 *   O:問合せ・提案内容 P:対応状況（経緯） Q:備考 R:要確認メモ
 * 顧客マスタ: 「顧客マスタ」シート（無ければ自動作成）
 *   A:顧客コード B:取引先名 C:窓口 D:備考
 * Excelが使えない環境ではデモモードで動作（ブラウザ単体テスト用）
 * ============================================================ */

const APP_VERSION = "rev_20260710_a";
const SHEET_NAME = "営業報告";
const CUST_SHEET = "顧客マスタ";
const MAX_ROWS = 500;

/* ---------- ワークフロー定義（④） ----------
 * steps: 通常遷移の並び / terminals: 最終状態（分岐含む） */
const WORKFLOWS = {
  "保守対応":     { steps: ["新規", "対応中"],                     terminals: ["完了"] },
  "瑕疵対応":     { steps: ["新規", "対応中"],                     terminals: ["完了"] },
  "見積り":       { steps: ["新規", "見積中", "確認中"],           terminals: ["受注", "失注"] },
  "プリセールス": { steps: ["新規", "作成中", "商談中", "確認中"], terminals: ["受注", "失注"] },
  "調整":         { steps: ["新規", "対応中"],                     terminals: ["完了"] },
};
const TYPES = Object.keys(WORKFLOWS);
const HOLD = "保留";
const QUOTE_TYPES = ["見積り", "プリセールス"];

/* 旧ステータス表記の読み替え */
const LEGACY_STATUS = {
  "未着手": "新規", "検討中": "新規",
  "見積作成中": "見積中", "見積提出済み": "確認中",
  "調整中": "対応中",
  "完了(受注)": "受注", "完了(失注)": "失注",
};

/* ---------- 状態 ---------- */
let records = [];          // {row, id, client, no, type, status, occur, done, owner, contact, priority, hours, amount, order, deliver, content, progress, note, memo}
let customers = [];        // {row, code, name, contact, note}
let demoMode = false;
let editingRec = null;
let inputType = "保守対応";
let currentKanbanType = "保守対応";
let dragId = null;
let filters = { q: "", type: "", status: "", client: "", owner: "" };

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
  renderInputForm();
  renderFilters();
  renderCurrentPane();
}

function bindStaticUI() {
  const input = document.getElementById("search-input");
  input.addEventListener("input", () => { filters.q = input.value.trim(); renderCurrentPane(); });
  document.getElementById("search-clear").addEventListener("click", () => {
    input.value = ""; filters.q = ""; renderCurrentPane();
  });
  ["type", "status", "client", "owner"].forEach(k => {
    document.getElementById("filter-" + k).addEventListener("change", e => {
      filters[k] = e.target.value; renderCurrentPane();
    });
  });
}

function clearFilters() {
  filters = { q: "", type: "", status: "", client: "", owner: "" };
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
      const rng = sheet.getRange(`A2:R${MAX_ROWS}`);
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
    if (!r[0] && !r[14]) return; // ID・内容が両方空ならスキップ
    out.push({
      row: i + 2,
      id: str(r[0]), client: str(r[1]), no: r[2],
      type: str(r[3]),
      status: normalizeStatus(str(r[4]), str(r[3])),
      occur: toDate(r[5]), done: toDate(r[6]),
      owner: str(r[7]), contact: str(r[8]), priority: str(r[9]),
      hours: numOrNull(r[10]), amount: numOrNull(r[11]),
      order: str(r[12]), deliver: toDate(r[13]),
      content: str(r[14]), progress: str(r[15]), note: str(r[16]), memo: str(r[17]),
    });
  });
  return out;
}

function normalizeStatus(s, type) {
  if (!s) return "新規";
  if (LEGACY_STATUS[s]) s = LEGACY_STATUS[s];
  if (s === "完了" && QUOTE_TYPES.includes(type)) return "受注"; // 旧「完了」見積は暫定的に受注扱い…はせず保留が安全
  return s;
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
      // 既存の営業報告データから種を作る
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

/* 案件番号の採番（⑤ 顧客コード + 通番） */
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
    const rng = sheet.getRange(`A${row}:R${row}`);
    rng.values = recToRow(rec);
    ["F", "G", "N"].forEach(c => sheet.getRange(`${c}${row}`).numberFormat = [["yyyy/m/d"]]);
    sheet.getRange(`L${row}`).numberFormat = [["#,##0"]];
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
function toSerial(d) {
  if (!d) return "";
  return Math.round(d.getTime() / 86400000) + 25569;
}
function fmtDate(d) {
  if (!d) return "";
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}
function fmtDateInput(d) {
  if (!d) return "";
  const m = String(d.getMonth() + 1).padStart(2, "0"), day = String(d.getDate()).padStart(2, "0");
  return `${d.getFullYear()}-${m}-${day}`;
}
function fromDateInput(s) { return s ? new Date(s + "T00:00:00") : null; }
function fmtYen(n) { return n == null ? "" : Number(n).toLocaleString("ja-JP") + "円"; }
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function monthKey(d) { return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}`; }

function allStatusesOf(type) {
  const wf = WORKFLOWS[type];
  if (!wf) return [];
  return [...wf.steps, ...wf.terminals, HOLD];
}

/* 遷移可能な次状態（⑥） */
function allowedTransitions(rec) {
  const wf = WORKFLOWS[rec.type];
  if (!wf) return [];
  const chain = wf.steps;
  const cur = rec.status;
  const res = [];
  if (cur === HOLD) {
    chain.forEach(s => res.push({ to: s, kind: "normal" }));
    return res;
  }
  const idx = chain.indexOf(cur);
  if (idx >= 0) {
    if (idx + 1 < chain.length) {
      res.push({ to: chain[idx + 1], kind: "normal" });
    } else {
      wf.terminals.forEach(t =>
        res.push({ to: t, kind: t === "受注" ? "win" : (t === "失注" ? "lose" : "win") }));
    }
    if (idx > 0) res.push({ to: chain[idx - 1], kind: "back" });
    res.push({ to: HOLD, kind: "hold" });
  } else if (wf.terminals.includes(cur)) {
    res.push({ to: chain[chain.length - 1], kind: "back" }); // 差戻し（再オープン）
  }
  return res;
}

function isValidTransition(rec, to) {
  return allowedTransitions(rec).some(t => t.to === to);
}

/* 状態変更の適用（完了日・受注区分の自動セット） */
function applyStatus(rec, to) {
  rec.status = to;
  const wf = WORKFLOWS[rec.type];
  if (wf && wf.terminals.includes(to)) {
    if (!rec.done) rec.done = new Date();
    if (to === "受注") rec.order = "受注";
    else if (to === "失注") rec.order = "失注";
  } else {
    rec.done = null;
    if (to !== "受注" && to !== "失注") rec.order = "";
  }
}

/* ============================================================
   タブ制御
   ============================================================ */
function switchTab(tab) {
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
  ["input", "list", "kanban", "agg"].forEach(t => {
    document.getElementById("pane-" + t).style.display = (t === tab) ? "" : "none";
  });
  document.getElementById("filter-bar").style.display =
    (tab === "list" || tab === "kanban") ? "" : "none";
  renderCurrentPane();
}
function activeTab() {
  return document.querySelector(".tab.active").dataset.tab;
}
function renderCurrentPane() {
  const t = activeTab();
  if (t === "list") renderList();
  else if (t === "kanban") renderKanban();
  else if (t === "agg") renderAgg();
}

/* ============================================================
   フィルタ（②）
   ============================================================ */
function renderFilters() {
  fillSelect("filter-type", ["（種別: 全て）", ...TYPES], filters.type);
  const sts = [...new Set(records.map(r => r.status).filter(Boolean))];
  fillSelect("filter-status", ["（状態: 全て）", ...sts], filters.status);
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

function filteredRecords() {
  return records.filter(r => {
    if (filters.type && r.type !== filters.type) return false;
    if (filters.status && r.status !== filters.status) return false;
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
   一覧（② 種別別グループ表示、③ 右クリック編集）
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
        <tr><th>ID</th><th>取引先</th><th>状態</th><th>内容</th><th>担当</th><th>発生日</th><th>金額</th></tr>
        ${group.map(r => `
        <tr data-id="${esc(r.id)}" oncontextmenu="onRowContext(event,'${esc(r.id)}')" onclick="openEditModal('${esc(r.id)}')">
          <td class="muted">${esc(r.id)}</td>
          <td>${esc(r.client)}</td>
          <td><span class="status-pill st-${esc(r.status)}">${esc(r.status)}</span></td>
          <td>${esc(shorten(r.content, 40))}</td>
          <td>${esc(r.owner)}</td>
          <td class="muted">${fmtDate(r.occur)}</td>
          <td class="r">${r.amount != null ? esc(Number(r.amount).toLocaleString()) : ""}</td>
        </tr>`).join("")}
      </table>
    </div>`;
  });
  cont.innerHTML = html || `<div class="empty-note">案件がありません</div>`;
}
function shorten(s, n) { s = str(s).replace(/\n/g, " "); return s.length > n ? s.slice(0, n) + "…" : s; }
function onRowContext(ev, id) { ev.preventDefault(); openEditModal(id); }

/* ============================================================
   カンバン（② 状態レーン、③ 右クリック編集、⑥ 遷移制御）
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
  board.innerHTML = lanes.map(st => {
    const cards = recs.filter(r => r.status === st);
    return `<div class="lane" data-status="${esc(st)}"
        ondragover="onLaneDragOver(event)" ondragleave="onLaneDragLeave(event)" ondrop="onLaneDrop(event)">
      <div class="lane-head">${esc(st)}<span class="cnt">${cards.length}</span></div>
      <div class="lane-body">
        ${cards.map(r => `
          <div class="card t-${esc(r.type)}" draggable="true" data-id="${esc(r.id)}"
               ondragstart="onCardDragStart(event)"
               oncontextmenu="onRowContext(event,'${esc(r.id)}')"
               onclick="openEditModal('${esc(r.id)}')">
            <div class="cid">${esc(r.id)}｜${esc(r.client)}</div>
            <div class="ctitle">${esc(shorten(r.content, 46))}</div>
            <div class="cmeta">
              <span>${esc(r.owner)}</span>
              ${r.amount != null ? `<span>${esc(Number(r.amount).toLocaleString())}円</span>` : ""}
              ${r.priority ? `<span>優先:${esc(r.priority)}</span>` : ""}
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
  if (to === "受注" && (!rec.amount || !rec.deliver)) {
    openEditModal(rec.id);
    alert("受注にするには「見積金額」と「納品日」の入力をお願いします。編集画面を開きました。");
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
   ワークフロー可視化ステッパー（④）
   ============================================================ */
function renderStepper(el, type, currentStatus) {
  const wf = WORKFLOWS[type];
  if (!wf) { el.innerHTML = ""; return; }
  const idx = wf.steps.indexOf(currentStatus);
  let html = "";
  wf.steps.forEach((s, i) => {
    let cls = "step";
    if (currentStatus != null) {
      if (i < idx) cls += " done";
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
   新規入力（①⑤）
   ============================================================ */
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
  toggleQuoteFields(document.getElementById("pane-input"), QUOTE_TYPES.includes(inputType));

  fillOwnerSelect("in-owner");
  document.getElementById("in-occur").value = fmtDateInput(new Date());
  updateNewId();
}
function setInputType(t) {
  inputType = t;
  document.querySelectorAll("#in-type-seg button").forEach(b =>
    b.classList.toggle("active", b.dataset.type === t));
  renderStepper(document.getElementById("in-stepper"), t, "新規");
  toggleQuoteFields(document.getElementById("pane-input"), QUOTE_TYPES.includes(t));
}
function toggleQuoteFields(scope, show) {
  scope.querySelectorAll(".quote-only").forEach(el => el.style.display = show ? "" : "none");
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
    client: cust.name,
    no: "",
    type: inputType,
    status: "新規",
    occur: fromDateInput(document.getElementById("in-occur").value) || new Date(),
    done: null,
    owner: document.getElementById("in-owner").value,
    contact: document.getElementById("in-contact").value || cust.contact,
    priority: document.getElementById("in-priority").value,
    hours: numOrNull(document.getElementById("in-hours").value),
    amount: numOrNull(document.getElementById("in-amount").value),
    order: "", deliver: null,
    content,
    progress: document.getElementById("in-progress").value,
    note: document.getElementById("in-note").value,
    memo: "",
  };
  try {
    await writeRecord(rec);
    msg.textContent = `登録しました（${rec.id}）`;
    ["in-content", "in-progress", "in-note", "in-contact", "in-hours", "in-amount"].forEach(id =>
      document.getElementById(id).value = "");
    renderFilters();
    updateNewId();
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
  }
}

/* ============================================================
   編集モーダル（③⑥）
   ============================================================ */
function openEditModal(id) {
  const rec = records.find(r => r.id === id);
  if (!rec) return;
  editingRec = { ...rec };
  document.getElementById("ed-title").textContent = `${rec.id}　${rec.client}`;
  document.getElementById("ed-id").value = rec.id;
  document.getElementById("ed-client").value = rec.client;
  const tSel = document.getElementById("ed-type");
  tSel.innerHTML = TYPES.map(t => `<option${t === rec.type ? " selected" : ""}>${esc(t)}</option>`).join("");
  tSel.onchange = () => {
    editingRec.type = tSel.value;
    if (!allStatusesOf(editingRec.type).includes(editingRec.status)) editingRec.status = "新規";
    refreshEditWorkflow();
  };
  document.getElementById("ed-status").value = rec.status;
  document.getElementById("ed-occur").value = fmtDateInput(rec.occur);
  document.getElementById("ed-done").value = fmtDateInput(rec.done);
  fillOwnerSelect("ed-owner", rec.owner);
  document.getElementById("ed-contact").value = rec.contact;
  document.getElementById("ed-priority").value = rec.priority;
  document.getElementById("ed-hours").value = rec.hours ?? "";
  document.getElementById("ed-amount").value = rec.amount ?? "";
  document.getElementById("ed-deliver").value = fmtDateInput(rec.deliver);
  document.getElementById("ed-content").value = rec.content;
  document.getElementById("ed-progress").value = rec.progress;
  document.getElementById("ed-note").value = rec.note;
  document.getElementById("ed-msg").textContent = "";
  refreshEditWorkflow();
  document.getElementById("edit-modal").style.display = "";
}

function refreshEditWorkflow() {
  renderStepper(document.getElementById("ed-stepper"), editingRec.type, editingRec.status);
  document.getElementById("ed-status").value = editingRec.status;
  toggleQuoteFields(document.getElementById("edit-modal"), QUOTE_TYPES.includes(editingRec.type));
  const bar = document.getElementById("ed-transitions");
  const trans = allowedTransitions(editingRec);
  bar.innerHTML = `<span class="tlabel">次の状態へ：</span>` + (trans.length
    ? trans.map(t => {
        const cls = t.kind === "win" ? "win" : t.kind === "lose" ? "lose" : t.kind === "hold" ? "hold" : t.kind === "back" ? "back" : "";
        const arrow = t.kind === "back" ? "↩ " : "→ ";
        return `<button class="tr-btn ${cls}" onclick="doTransition('${esc(t.to)}')">${arrow}${esc(t.to)}</button>`;
      }).join("")
    : `<span class="tlabel">（遷移可能な状態がありません）</span>`);
}

function doTransition(to) {
  if (to === "受注") {
    const amount = numOrNull(document.getElementById("ed-amount").value);
    const deliver = fromDateInput(document.getElementById("ed-deliver").value);
    if (!amount || !deliver) {
      const m = document.getElementById("ed-msg");
      m.className = "save-msg err";
      m.textContent = "受注にするには「見積金額」と「納品日」を入力してください";
      return;
    }
    editingRec.amount = amount;
    editingRec.deliver = deliver;
  }
  applyStatus(editingRec, to);
  document.getElementById("ed-done").value = fmtDateInput(editingRec.done);
  refreshEditWorkflow();
}

async function saveEditRecord() {
  const msg = document.getElementById("ed-msg");
  msg.className = "save-msg"; msg.textContent = "";
  editingRec.occur = fromDateInput(document.getElementById("ed-occur").value);
  editingRec.done = fromDateInput(document.getElementById("ed-done").value);
  editingRec.owner = document.getElementById("ed-owner").value;
  editingRec.contact = document.getElementById("ed-contact").value;
  editingRec.priority = document.getElementById("ed-priority").value;
  editingRec.hours = numOrNull(document.getElementById("ed-hours").value);
  editingRec.amount = numOrNull(document.getElementById("ed-amount").value);
  editingRec.deliver = fromDateInput(document.getElementById("ed-deliver").value);
  editingRec.content = document.getElementById("ed-content").value;
  editingRec.progress = document.getElementById("ed-progress").value;
  editingRec.note = document.getElementById("ed-note").value;
  try {
    await writeRecord(editingRec);
    msg.textContent = "保存しました";
    renderFilters();
    renderCurrentPane();
    setTimeout(closeEditModal, 500);
  } catch (e) {
    msg.className = "save-msg err"; msg.textContent = "保存に失敗しました: " + e.message;
  }
}
function closeEditModal() { document.getElementById("edit-modal").style.display = "none"; editingRec = null; }

/* ---------- 顧客追加モーダル（⑤） ---------- */
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
   集計（⑦）SVGチャート（外部ライブラリ不使用）
   ============================================================ */
let currentAgg = "hoshu";
function switchAgg(k) {
  currentAgg = k;
  document.querySelectorAll(".agg-seg .seg").forEach(b => b.classList.toggle("active", b.dataset.agg === k));
  renderAgg();
}
function renderAgg() {
  const cont = document.getElementById("agg-container");
  if (currentAgg === "hoshu") cont.innerHTML = renderHoshuAgg();
  else if (currentAgg === "mitsu") cont.innerHTML = renderMitsuAgg();
  else cont.innerHTML = renderUriageAgg();
}

/* 直近12か月の月キー配列 */
function last12Months() {
  const out = [];
  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    out.push(monthKey(d));
  }
  return out;
}

/* --- 保守状況: 保守対応＋瑕疵対応の月次 発生/完了 --- */
function renderHoshuAgg() {
  const months = last12Months();
  const target = records.filter(r => r.type === "保守対応" || r.type === "瑕疵対応");
  const series = {
    "保守 発生": countByMonth(target.filter(r => r.type === "保守対応"), "occur", months),
    "保守 完了": countByMonth(target.filter(r => r.type === "保守対応"), "done", months),
    "瑕疵 発生": countByMonth(target.filter(r => r.type === "瑕疵対応"), "occur", months),
    "瑕疵 完了": countByMonth(target.filter(r => r.type === "瑕疵対応"), "done", months),
  };
  const colors = { "保守 発生": "#548235", "保守 完了": "#a9d08e", "瑕疵 発生": "#c00000", "瑕疵 完了": "#f4a7a7" };
  const open = target.filter(r => !WORKFLOWS[r.type].terminals.includes(r.status)).length;
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
      <h3>月別明細</h3>
      <table class="agg-table">
        <tr><th>月</th><th>保守 発生</th><th>保守 完了</th><th>瑕疵 発生</th><th>瑕疵 完了</th></tr>
        ${months.map((m, i) => `<tr><td>${m}</td>
          <td>${series["保守 発生"][i]}</td><td>${series["保守 完了"][i]}</td>
          <td>${series["瑕疵 発生"][i]}</td><td>${series["瑕疵 完了"][i]}</td></tr>`).join("")}
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

/* --- 見積状況: 見積り＋プリセールスの状態別件数・金額 --- */
function renderMitsuAgg() {
  const target = records.filter(r => QUOTE_TYPES.includes(r.type));
  const statuses = [...new Set([
    ...WORKFLOWS["見積り"].steps, ...WORKFLOWS["プリセールス"].steps,
    "受注", "失注", HOLD,
  ])];
  const rows = statuses.map(st => {
    const g = target.filter(r => r.status === st);
    return { st, cnt: g.length, amt: g.reduce((a, r) => a + (r.amount || 0), 0) };
  }).filter(r => r.cnt > 0);
  const totalCnt = rows.reduce((a, r) => a + r.cnt, 0);
  const totalAmt = rows.reduce((a, r) => a + r.amt, 0);
  const pipeline = target.filter(r => !["受注", "失注"].includes(r.status));
  const pipelineAmt = pipeline.reduce((a, r) => a + (r.amount || 0), 0);
  const wonAmt = target.filter(r => r.status === "受注").reduce((a, r) => a + (r.amount || 0), 0);
  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${pipeline.length}</div><div class="kl">進行中案件</div></div>
      <div class="kpi"><div class="kv">${(pipelineAmt / 10000).toLocaleString()}万</div><div class="kl">パイプライン金額</div></div>
      <div class="kpi"><div class="kv">${(wonAmt / 10000).toLocaleString()}万</div><div class="kl">受注確定金額</div></div>
    </div>
    <div class="agg-card">
      <h3>見積り・プリセールス 状態別集計</h3>
      <table class="agg-table">
        <tr><th>状態</th><th>件数</th><th>見積金額合計</th></tr>
        ${rows.map(r => `<tr>
          <td><span class="status-pill st-${esc(r.st)}">${esc(r.st)}</span></td>
          <td>${r.cnt}</td><td class="r">${r.amt ? r.amt.toLocaleString() + "円" : "－"}</td></tr>`).join("")}
        <tr class="total"><td>合計</td><td>${totalCnt}</td><td class="r">${totalAmt.toLocaleString()}円</td></tr>
      </table>
      <p style="font-size:11px;color:#999;margin-top:6px">※ 見積金額が未入力の案件は金額集計に含まれません。</p>
    </div>`;
}

/* --- 売上予測: 受注確定の納品日ベース 月別売上 --- */
function renderUriageAgg() {
  const months = last12Fwd();
  const won = records.filter(r => QUOTE_TYPES.includes(r.type) && r.status === "受注");
  const map = Object.fromEntries(months.map(m => [m, 0]));
  let noDeliver = 0;
  won.forEach(r => {
    if (r.deliver && map[monthKey(r.deliver)] != null) map[monthKey(r.deliver)] += (r.amount || 0);
    else if (!r.deliver) noDeliver++;
  });
  const vals = months.map(m => map[m]);
  const total = won.reduce((a, r) => a + (r.amount || 0), 0);
  const colors = { "確定売上": "#4472c4" };
  return `
    <div class="kpi-row">
      <div class="kpi"><div class="kv">${won.length}</div><div class="kl">受注案件数</div></div>
      <div class="kpi"><div class="kv">${(total / 10000).toLocaleString()}万</div><div class="kl">受注金額合計</div></div>
      ${noDeliver ? `<div class="kpi"><div class="kv" style="color:#c00000">${noDeliver}</div><div class="kl">納品日未入力</div></div>` : ""}
    </div>
    <div class="agg-card">
      <h3>月別売上予測（受注確定・納品日ベース）</h3>
      ${legendHtml(colors)}
      <div class="chart-wrap">${groupedBarChart(months, { "確定売上": vals }, colors, v => (v / 10000) + "万")}</div>
    </div>
    <div class="agg-card">
      <h3>受注案件一覧</h3>
      <table class="agg-table">
        <tr><th>ID</th><th>取引先</th><th>内容</th><th>金額</th><th>納品日</th></tr>
        ${won.length ? won.map(r => `<tr>
          <td>${esc(r.id)}</td><td class="l">${esc(r.client)}</td>
          <td class="l">${esc(shorten(r.content, 26))}</td>
          <td class="r">${r.amount ? r.amount.toLocaleString() + "円" : "－"}</td>
          <td>${r.deliver ? fmtDate(r.deliver) : '<span style="color:#c00000">未入力</span>'}</td></tr>`).join("")
        : `<tr><td colspan="5" class="muted">受注案件はまだありません</td></tr>`}
      </table>
    </div>`;
}
/* 今月起点で前2か月＋先9か月 */
function last12Fwd() {
  const out = [];
  const now = new Date();
  for (let i = -2; i <= 9; i++) {
    out.push(monthKey(new Date(now.getFullYear(), now.getMonth() + i, 1)));
  }
  return out;
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
  const barW = Math.min(14, (groupW - 8) / names.length);
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
   デモモード（ブラウザ単体テスト用）
   ============================================================ */
function loadDemo() {
  demoMode = true;
  const d = (y, m, day) => new Date(y, m - 1, day);
  customers = [
    { row: 2, code: "KM", name: "kakimoto arms", contact: "佐竹様", note: "" },
    { row: 3, code: "HN", name: "ハンター製菓", contact: "鈴木様", note: "" },
    { row: 4, code: "AG", name: "アサヒグラント", contact: "川野様", note: "" },
    { row: 5, code: "EX", name: "エキスプレス", contact: "中道様", note: "" },
  ];
  records = [
    { row: 2, id: "KM-01", client: "kakimoto arms", no: 1, type: "見積り", status: "見積中", occur: d(2026, 6, 29), done: null, owner: "小川", contact: "佐竹様", priority: "中", hours: 10, amount: null, order: "", deliver: null, content: "ネット予約でフリースタッフを選択できるようにしたい", progress: "調査中", note: "", memo: "" },
    { row: 3, id: "KM-02", client: "kakimoto arms", no: 2, type: "見積り", status: "受注", occur: d(2026, 6, 18), done: d(2026, 7, 3), owner: "小川", contact: "佐竹様", priority: "", hours: null, amount: 600000, order: "受注", deliver: d(2026, 7, 17), content: "ネット予約LINEログイン連携", progress: "60万で受注確定", note: "", memo: "" },
    { row: 4, id: "KM-03", client: "kakimoto arms", no: 3, type: "保守対応", status: "完了", occur: d(2026, 7, 2), done: d(2026, 7, 2), owner: "小川", contact: "西野様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "スタッフ指名予約で店舗が正しく選択されない", progress: "外部サイト側の設定が原因", note: "", memo: "" },
    { row: 5, id: "KM-04", client: "kakimoto arms", no: 4, type: "調整", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", contact: "佐竹様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "会社体制変更に伴うご挨拶のスケジュール調整", progress: "", note: "", memo: "" },
    { row: 6, id: "KM-05", client: "kakimoto arms", no: 5, type: "保守対応", status: "対応中", occur: d(2026, 7, 7), done: null, owner: "小川", contact: "中田様", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "メンズ予約時の注意事項表示・メール文面変更", progress: "設定変更で対応可能", note: "", memo: "" },
    { row: 7, id: "HN-01", client: "ハンター製菓", no: 1, type: "瑕疵対応", status: "対応中", occur: d(2026, 7, 3), done: null, owner: "小川", contact: "鈴木様", priority: "低", hours: null, amount: null, order: "", deliver: null, content: "在庫管理伝票一覧画面バグ対応", progress: "修正済み、次回リリースで反映", note: "", memo: "" },
    { row: 8, id: "HN-02", client: "ハンター製菓", no: 2, type: "プリセールス", status: "商談中", occur: d(2026, 7, 6), done: null, owner: "小川", contact: "柳澤様", priority: "高", hours: null, amount: 2500000, order: "", deliver: null, content: "原価計算の改修", progress: "7/22打ち合わせ予定", note: "9月本稼働目標", memo: "" },
    { row: 9, id: "AG-01", client: "アサヒグラント", no: 1, type: "見積り", status: "確認中", occur: d(2026, 6, 30), done: null, owner: "紺谷", contact: "川野様", priority: "", hours: null, amount: 350000, order: "", deliver: null, content: "インフォマートデータ交換の仕様変更", progress: "再見積中", note: "", memo: "" },
    { row: 10, id: "EX-01", client: "エキスプレス", no: 1, type: "見積り", status: "新規", occur: d(2026, 7, 6), done: null, owner: "紺谷", contact: "中道様", priority: "", hours: null, amount: null, order: "", deliver: null, content: "削除した請求書を参照できる機能の見積", progress: "", note: "", memo: "" },
  ];
}
