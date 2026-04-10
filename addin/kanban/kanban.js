// =====================
// バージョン
// =====================
const APP_VERSION = "rev.20260409_001";

let tasks = [];
let idRowMap = {};
let currentTask = null;

let activeUser = null;
let activeCategory = null;
let activePeriod = "all";

// =====================
Office.onReady(() => init());

async function init() {
  tasks = await loadTasks();
  restoreFilters();
  render();
  initFilterUI();
  initCategoryFilter();
  updateActiveUI();
  setVersion(); // ★追加
}

// =====================
// バージョン表示
// =====================
function setVersion() {
  const el = document.getElementById("version");
  if (el) el.textContent = APP_VERSION;
}

// =====================
// Excel取得
// =====================
async function loadTasks() {
  return await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");
    const range = sheet.getRange("A11:Z1000");

    range.load("values");
    await context.sync();

    return range.values.map((row, i) => {

      const category = row[0];
      const name = row[13];
      const note = row[14];

      const start = row[15];
      const end   = row[16];
      const r     = row[17];
      const s     = row[18];

      const id = row[24];
      const task_name = row[25];

      if (!task_name) return null;

      const rowNumber = i + 11;
      const safeId = id || `row-${rowNumber}`;
      idRowMap[safeId] = rowNumber;

      let status = "todo";
      if (s) status = "done";
      else if (r) status = "doing";

      return {
        id: safeId,
        row: rowNumber,
        category,
        task_name,
        name,
        note,
        plannedStart: start,
        plannedEnd: end,
        actualStart: r,
        actualEnd: s,
        status
      };

    }).filter(Boolean);
  });
}

// =====================
// 描画
// =====================
function render() {

  document.querySelectorAll(".card-list").forEach(e => e.innerHTML = "");

  tasks.sort((a, b) => {
    const da = toDate(a.plannedEnd) || new Date(9999,0,1);
    const db = toDate(b.plannedEnd) || new Date(9999,0,1);
    return da - db;
  });

  tasks.forEach(task => {

    if (!isVisible(task)) return;

    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.task_name;
    card.draggable = true;

    card.onclick = () => openModal(task);

    card.ondragstart = e => {
      e.dataTransfer.setData("id", task.id);
    };

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatRange(task.plannedStart, task.plannedEnd);
    card.appendChild(meta);

    applyDeadlineColor(card, task);

    document.querySelector(`#${task.status} .card-list`)?.appendChild(card);
  });
}

// =====================
// フィルタ
// =====================
function isVisible(task) {

  if ((task.name && task.name.includes("#")) ||
      (task.category && task.category.includes("#"))) return false;

  if (activeUser && task.name !== activeUser) return false;
  if (activeCategory && task.category !== activeCategory) return false;

  if (activePeriod === "all") return true;

  const start = toDate(task.plannedStart);
  const end   = toDate(task.plannedEnd);

  const today = new Date();
  today.setHours(0,0,0,0);

  const day = today.getDay() || 7;
  const thisWeekStart = new Date(today);
  thisWeekStart.setDate(today.getDate() - day + 1);

  const thisWeekEnd = new Date(thisWeekStart);
  thisWeekEnd.setDate(thisWeekStart.getDate() + 6);

  const nextWeekStart = new Date(thisWeekStart);
  nextWeekStart.setDate(thisWeekStart.getDate() + 7);

  const nextWeekEnd = new Date(thisWeekStart);
  nextWeekEnd.setDate(thisWeekStart.getDate() + 13);

  const ts = start || end;
  const te = end || start;

  if (!ts || !te) return true;

  if (activePeriod === "past") return te < thisWeekStart;
  if (activePeriod === "week") return ts <= thisWeekEnd && te >= thisWeekStart;
  if (activePeriod === "nextweek") return ts <= nextWeekEnd && te >= nextWeekStart;
  if (activePeriod === "future") return ts > nextWeekEnd;

  return true;
}

// =====================
// フィルタUI
// =====================
function initFilterUI() {
  const users = [...new Set(tasks.map(t => t.name).filter(Boolean))];
  const el = document.getElementById("user-filters");
  el.innerHTML = "";

  users.forEach(u => {
    if (u.includes("#")) return;

    const b = document.createElement("button");
    b.textContent = u;

    b.onclick = () => {
      activeUser = (activeUser === u) ? null : u;
      updateActiveUI();
      saveFilters();
      render();
    };

    el.appendChild(b);
  });
}

function initCategoryFilter() {
  const cats = [...new Set(tasks.map(t => t.category).filter(Boolean))];
  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    if (c.includes("#")) return;

    const b = document.createElement("button");
    b.textContent = c;

    b.onclick = () => {
      activeCategory = (activeCategory === c) ? null : c;
      updateActiveUI();
      saveFilters();
      render();
    };

    el.appendChild(b);
  });
}

function setPeriod(p) {
  activePeriod = (activePeriod === p) ? "all" : p;
  updateActiveUI();
  saveFilters();
  render();
}

function updateActiveUI() {
  document.querySelectorAll("#user-filters button").forEach(b =>
    b.classList.toggle("active", b.textContent === activeUser)
  );
  document.querySelectorAll("#category-filters button").forEach(b =>
    b.classList.toggle("active", b.textContent === activeCategory)
  );
  document.querySelectorAll("[data-period]").forEach(b =>
    b.classList.toggle("active", b.dataset.period === activePeriod)
  );
}

function saveFilters() {
  localStorage.setItem("kanbanFilter", JSON.stringify({
    user: activeUser,
    category: activeCategory,
    period: activePeriod
  }));
}

function restoreFilters() {
  const d = JSON.parse(localStorage.getItem("kanbanFilter") || "{}");
  activeUser = d.user || null;
  activeCategory = d.category || null;
  activePeriod = d.period || "all";
}

// =====================
// 色
// =====================
function applyDeadlineColor(card, task) {

  if (task.status === "done") {
    card.classList.add("done");
    return;
  }

  const end = toDate(task.plannedEnd);
  if (!end) return;

  const today = new Date();
  today.setHours(0,0,0,0);
  end.setHours(0,0,0,0);

  if (end < today) {
    card.classList.add("overdue");
    return;
  }

  const day = today.getDay() || 7;
  const start = new Date(today);
  start.setDate(today.getDate() - day + 1);

  const endWeek = new Date(start);
  endWeek.setDate(start.getDate() + 6);

  if (end >= start && end <= endWeek) {
    card.classList.add("thisweek");
  }
}

// =====================
// D&D（Excel連携）
// =====================
window.allowDrop = function(e) {
  e.preventDefault();
};

window.drop = async function(e, status) {
  e.preventDefault();

  const id = e.dataTransfer.getData("id");
  const task = tasks.find(t => t.id === id);
  if (!task) return;

  task.status = status;

  const row = idRowMap[task.id];
  const today = new Date();
  const formatted = `${today.getFullYear()}/${today.getMonth()+1}/${today.getDate()}`;

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");

    const startCell = sheet.getRange(`R${row}`);
    const endCell   = sheet.getRange(`S${row}`);

    startCell.load("values");
    await context.sync();

    const startVal = startCell.values[0][0];

    if (status === "doing") {
      startCell.values = [[formatted]];
      endCell.values = [[""]];
    }
    else if (status === "done") {
      if (!startVal) startCell.values = [[formatted]];
      endCell.values = [[formatted]];
    }
    else {
      startCell.values = [[""]];
      endCell.values = [[""]];
    }

    await context.sync();
  });

  render();
};

// =====================
// モーダル
// =====================
function openModal(task) {
  currentTask = task;
  document.getElementById("modal-title").textContent = task.task_name;
  document.getElementById("modal-note").value = task.note || "";
  document.getElementById("modal").classList.remove("hidden");
}

function closeModal() {
  document.getElementById("modal").classList.add("hidden");
}

// 備考保存（行高さ維持）
async function saveNote() {

  const note = document.getElementById("modal-note").value;
  const row = idRowMap[currentTask.id];

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");

    const rowRange = sheet.getRange(`A${row}:Z${row}`);
    rowRange.load("rowHeight");

    const noteCell = sheet.getRange(`O${row}`);

    await context.sync();

    const height = rowRange.rowHeight;

    noteCell.values = [[note]];

    await context.sync();

    rowRange.rowHeight = height;

    await context.sync();
  });

  currentTask.note = note;
  closeModal();
}

// =====================
function toDate(v) {
  if (!v) return null;
  if (typeof v === "number") return new Date((v - 25569) * 86400 * 1000);
  return new Date(v);
}

function formatRange(s, e) {
  const sd = toDate(s);
  const ed = toDate(e);
  const f = d => `${d.getMonth()+1}/${d.getDate()}`;
  if (sd && ed) return `${f(sd)}～${f(ed)}`;
  if (sd) return `${f(sd)}～`;
  if (ed) return `～${f(ed)}`;
  return "";
}