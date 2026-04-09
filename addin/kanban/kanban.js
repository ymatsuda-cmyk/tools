let tasks = [];
let currentTask = null;
let idRowMap = {};
let clickTimer = null;

// フィルタ状態
let activeUser = null;
let activePeriod = "all";
let activeCategory = null;
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
// Excelデータ取得
// =====================
async function loadTasks() {
  return await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");
    const range = sheet.getRange("A11:Z1000");

    range.load("values");
    await context.sync();

    return range.values.map((row, i) => {

      const category = row[0];   // A
      const name = row[13];      // N
      const note = row[14];      // O
      const plannedStart = row[15]; // P
      const plannedEnd = row[16];   // Q
      const actualStart = row[17];  // R
      const actualEnd = row[18];    // S
      const priority = row[19];     // T
      const id = row[24];           // Y
      const task_name = row[25];    // Z

      if (!task_name) return null;

      const rowNumber = i + 11;
      const safeId = id || `row-${rowNumber}`;
      idRowMap[safeId] = rowNumber;

      let status = "todo";
      if (actualEnd) status = "done";
      else if (actualStart) status = "doing";

      return {
        id: safeId,
        row: rowNumber,

        category,
        task_name,
        name,
        note,

        plannedStart,
        plannedEnd,
        actualStart,
        actualEnd,

        priority,
        status
      };

    }).filter(Boolean);
  });
}

// =====================
// 描画
// =====================
function render() {

  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  tasks.sort((a, b) => new Date(a.plannedEnd) - new Date(b.plannedEnd));

  tasks.forEach(task => {

    if (!isVisible(task)) return;

    const card = document.createElement("div");
    card.className = "card";
    card.dataset.id = task.id;
    card.textContent = task.task_name;
    card.draggable = true;

    // クリック → Excelジャンプ
    card.addEventListener("click", () => {
      clickTimer = setTimeout(() => focusRow(task.id), 200);
    });

    // ダブルクリック → モーダル
    card.addEventListener("dblclick", () => {
      clearTimeout(clickTimer);
      openModal(task);
    });

    card.addEventListener("dragstart", e => {
      e.dataTransfer.setData("id", task.id);
    });

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatRange(task.plannedStart, task.plannedEnd);
    card.appendChild(meta);

    applyDeadlineColor(card, task.plannedEnd);

    document.querySelector(`#${task.status} .card-list`)?.appendChild(card);
  });
}

// =====================
// フィルタ判定
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
// 担当者フィルタ
// =====================
function initFilterUI() {

  const users = [...new Set(tasks.map(t => t.name).filter(Boolean))];
  const container = document.getElementById("user-filters");

  container.innerHTML = "";

  users.forEach(u => {

    const btn = document.createElement("button");
    btn.textContent = u;

    btn.onclick = () => {
      activeUser = (activeUser === u) ? null : u;
      updateActiveUI();
      saveFilters();
      render();
    };

    container.appendChild(btn);
  });
}

// =====================
// 大分類フィルタ
// =====================
function initCategoryFilter() {

  const categories = [...new Set(tasks.map(t => t.category).filter(Boolean))];
  const container = document.getElementById("category-filters");

  container.innerHTML = "";

  categories.forEach(c => {

    const btn = document.createElement("button");
    btn.textContent = c;

    btn.onclick = () => {
      activeCategory = (activeCategory === c) ? null : c;
      updateActiveUI();
      saveFilters();
      render();
    };

    container.appendChild(btn);
  });
}

// =====================
// 期間フィルタ
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
// UI状態反映
// =====================
function updateActiveUI() {

  document.querySelectorAll("#user-filters button").forEach(btn => {
    btn.classList.toggle("active", btn.textContent === activeUser);
  });

  document.querySelectorAll("#category-filters button").forEach(btn => {
    btn.classList.toggle("active", btn.textContent === activeCategory);
  });

  document.querySelectorAll("[data-period]").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.period === activePeriod);
  });
}

// =====================
// localStorage
// =====================
function saveFilters() {
  localStorage.setItem("kanbanFilter", JSON.stringify({
    user: activeUser,
    period: activePeriod,
    category: activeCategory,
    includeOverdue
  }));
}

function restoreFilters() {
  const data = JSON.parse(localStorage.getItem("kanbanFilter") || "{}");

  activeUser = data.user || null;
  activePeriod = data.period || "all";
  activeCategory = data.category || null;
  includeOverdue = data.includeOverdue || false;
}

// =====================
// Excelジャンプ
// =====================
async function focusRow(id) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    const row = await getRowById(context, sheet, id);
    sheet.getRange(`N${row}:Z${row}`).select();
    await context.sync();
  });
}

// =====================
// ID検索
// =====================
async function getRowById(context, sheet, id) {

  if (idRowMap[id]) return idRowMap[id];

  const range = sheet.getRange("Y11:Y1000");
  range.load("values");
  await context.sync();

  for (let i = 0; i < range.values.length; i++) {
    if (range.values[i][0] == id) {
      const row = i + 11;
      idRowMap[id] = row;
      return row;
    }
  }

  throw new Error("ID not found");
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
// Excel更新
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

async function saveNote() {

  const note = document.getElementById("modal-note").value;

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");
    const row = await getRowById(context, sheet, currentTask.id);

    sheet.getRange(`O${row}`).values = [[note]];

    await context.sync();
  });

  closeModal();
}

// =====================
// 日付
// =====================
function applyDeadlineColor(card, endDate) {

  if (!endDate) return;

  const end = new Date(endDate);
  const now = new Date();
  const diff = (end - now) / 86400000;

  if (diff < 0) card.classList.add("overdue");
  else if (diff <= 7) card.classList.add("thisweek");
  else if (diff <= 14) card.classList.add("nextweek");
}

function toExcelDateString(d) {
  return `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
}

function formatRange(start, end) {

  const toDate = v => {
    if (!v) return null;
    if (typeof v === "number") return new Date((v - 25569) * 86400 * 1000);
    return new Date(v);
  };

  const s = toDate(start);
  const e = toDate(end);

  const f = d => `${d.getMonth()+1}/${d.getDate()}`;

  if (s && e) return `${f(s)}～${f(e)}`;
  if (s) return `${f(s)}～`;
  if (e) return `～${f(e)}`;
  return "";
}