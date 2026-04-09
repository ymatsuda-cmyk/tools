// =====================
// 状態
// =====================
let tasks = [];
let idRowMap = {};
let currentTask = null;
let clickTimer = null;

let activeUser = null;
let activeCategory = null;
let activePeriod = "all";
let includeOverdue = false;

// =====================
// 初期化
// =====================
Office.onReady(() => init());

async function init() {
  tasks = await loadTasks();
  restoreFilters();
  render();
  initFilterUI();
  initCategoryFilter();
  updateActiveUI();
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
      const end = row[16];
      const r = row[17];
      const s = row[18];
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

  tasks.sort((a, b) => new Date(a.plannedEnd) - new Date(b.plannedEnd));

  tasks.forEach(task => {

    if (!isVisible(task)) return;

    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.task_name;
    card.draggable = true;

    card.onclick = () => {
      clickTimer = setTimeout(() => focusRow(task.id), 200);
    };

    card.ondblclick = () => {
      clearTimeout(clickTimer);
      openModal(task);
    };

    card.ondragstart = e => {
      e.dataTransfer.setData("id", task.id);
    };

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatRange(task.plannedStart, task.plannedEnd);
    card.appendChild(meta);

    applyDeadlineColor(card, task.plannedEnd);

    document.querySelector(`#${task.status} .card-list`)?.appendChild(card);
  });
}

// =====================
// フィルタ
// =====================
function isVisible(task) {

  if (activeUser && task.name !== activeUser) return false;
  if (activeCategory && task.category !== activeCategory) return false;

  const end = new Date(task.plannedEnd);
  const now = new Date();
  const diff = (end - now) / 86400000;
  const isOverdue = diff < 0;

  if (activePeriod !== "all") {

    if (includeOverdue && isOverdue) return true;

    if (activePeriod === "week" && diff > 7) return false;
    if (activePeriod === "nextweek" && diff > 14) return false;

    if (activePeriod === "month") {
      if (end.getMonth() !== now.getMonth()) return false;
    }
  }

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

// =====================
// 期間
// =====================
function setPeriod(p) {
  activePeriod = (activePeriod === p) ? "all" : p;
  updateActiveUI();
  saveFilters();
  render();
}

function toggleOverdue(cb) {
  includeOverdue = cb.checked;
  saveFilters();
  render();
}

// =====================
// UI更新
// =====================
function updateActiveUI() {

  document.querySelectorAll("#user-filters button").forEach(b => {
    b.classList.toggle("active", b.textContent === activeUser);
  });

  document.querySelectorAll("#category-filters button").forEach(b => {
    b.classList.toggle("active", b.textContent === activeCategory);
  });

  document.querySelectorAll("[data-period]").forEach(b => {
    b.classList.toggle("active", b.dataset.period === activePeriod);
  });
}

// =====================
// localStorage
// =====================
function saveFilters() {
  localStorage.setItem("kanbanFilter", JSON.stringify({
    user: activeUser,
    category: activeCategory,
    period: activePeriod,
    includeOverdue
  }));
}

function restoreFilters() {
  const d = JSON.parse(localStorage.getItem("kanbanFilter") || "{}");
  activeUser = d.user || null;
  activeCategory = d.category || null;
  activePeriod = d.period || "all";
  includeOverdue = d.includeOverdue || false;
}

// =====================
// D&D
// =====================
function allowDrop(e) { e.preventDefault(); }

async function onDrop(e) {

  e.preventDefault();

  const id = e.dataTransfer.getData("id");
  const task = tasks.find(t => t.id == id);
  if (!task) return;

  task.status = e.currentTarget.id;

  await updateExcelStatus(task);
  render();
}

// =====================
// Excel更新（省略なし）
// =====================
async function updateExcelStatus(task) {

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");
    const row = await getRowById(context, sheet, task.id);

    const r = sheet.getRange(`R${row}`);
    const s = sheet.getRange(`S${row}`);

    r.load("values");
    s.load("values");

    await context.sync();

    const er = r.values[0][0];
    const es = s.values[0][0];

    const today = toExcelDateString(new Date());

    if (task.status === "todo") {
      r.values = [[""]];
      s.values = [[""]];
    }
    if (task.status === "doing") {
      r.values = [[er || today]];
      s.values = [[""]];
    }
    if (task.status === "done") {
      r.values = [[er || today]];
      s.values = [[es || today]];
    }

    await context.sync();
  });
}

// =====================
// その他
// =====================
function applyDeadlineColor(card, endDate) {

  if (!endDate) return;

  const end = new Date(endDate);
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

  const endW = new Date(start);
  endW.setDate(start.getDate() + 6);

  const nextS = new Date(start);
  nextS.setDate(start.getDate() + 7);

  const nextE = new Date(start);
  nextE.setDate(start.getDate() + 13);

  if (end >= start && end <= endW) card.classList.add("thisweek");
  else if (end >= nextS && end <= nextE) card.classList.add("nextweek");
}

function toExcelDateString(d) {
  return `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
}

function formatRange(s, e) {
  const f = d => `${d.getMonth()+1}/${d.getDate()}`;
  const sd = s ? new Date(s) : null;
  const ed = e ? new Date(e) : null;
  if (sd && ed) return `${f(sd)}～${f(ed)}`;
  if (sd) return `${f(sd)}～`;
  if (ed) return `～${f(ed)}`;
  return "";
}