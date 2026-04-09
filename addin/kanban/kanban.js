let tasks = [];
let currentTask = null;
let idRowMap = {};

// フィルタ状態
let selectedUsers = new Set();
let selectedPeriod = "all";

// =========================
// 初期化
// =========================
Office.onReady(() => init());

async function init() {
  loadFilterState(); // ←復元
  tasks = await loadTasks();

  initUserFilter();
  initPeriodFilter();
  render();
}

// =========================
// localStorage
// =========================
function saveFilterState() {
  localStorage.setItem("kanban_users", JSON.stringify([...selectedUsers]));
  localStorage.setItem("kanban_period", selectedPeriod);
}

function loadFilterState() {
  const users = JSON.parse(localStorage.getItem("kanban_users") || "[]");
  selectedUsers = new Set(users);

  selectedPeriod = localStorage.getItem("kanban_period") || "all";
}

// =========================
// Reload
// =========================
async function reloadTasks() {
  tasks = await loadTasks();
  initUserFilter();
  initPeriodFilter();
  render();
}

// =========================
// 日付
// =========================
function toDate(v) {
  if (!v) return null;
  if (typeof v === "number") {
    return new Date((v - 25569) * 86400 * 1000);
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function toExcelDateString(date) {
  return `${date.getFullYear()}/${date.getMonth()+1}/${date.getDate()}`;
}

// =========================
// データ取得
// =========================
async function loadTasks() {
  return await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("N11:Z1000");

    range.load("values");
    await context.sync();

    idRowMap = {};

    return range.values.map((row, i) => {

      const name = row[0];
      const note = row[1];
      const plannedEnd = row[3];
      const actualStart = row[4];
      const actualEnd = row[5];
      const id = row[11];
      const task_name = row[12];

      if (!task_name) return null;

      const rowNum = i + 11;
      const safeId = id || `row-${rowNum}`;
      idRowMap[safeId] = rowNum;

      let status = "todo";
      if (actualEnd) status = "done";
      else if (actualStart) status = "doing";

      return {
        id: safeId,
        row: rowNum,
        task_name,
        name,
        note,
        plannedEnd,
        actualStart,
        actualEnd,
        status
      };

    }).filter(Boolean);
  });
}

// =========================
// フィルタUI
// =========================
function initUserFilter() {

  const container = document.getElementById("user-filter");
  container.innerHTML = "";

  const users = [...new Set(tasks.map(t => t.name).filter(Boolean))];

  users.forEach(user => {

    const label = document.createElement("span");
    label.textContent = user;
    label.className = "filter-label";

    if (selectedUsers.has(user)) {
      label.classList.add("active");
    }

    label.onclick = () => {
      if (selectedUsers.has(user)) {
        selectedUsers.delete(user);
      } else {
        selectedUsers.add(user);
      }
      saveFilterState();
      initUserFilter();
      render();
    };

    container.appendChild(label);
  });
}

function initPeriodFilter() {

  const container = document.getElementById("period-filter");
  container.innerHTML = "";

  const periods = [
    { key: "all", label: "全期間" },
    { key: "thisWeek", label: "今週中" },
    { key: "nextWeek", label: "来週中" },
    { key: "thisMonth", label: "今月中" }
  ];

  periods.forEach(p => {

    const label = document.createElement("span");
    label.textContent = p.label;
    label.className = "filter-label";

    if (selectedPeriod === p.key) {
      label.classList.add("active-period");
    }

    label.onclick = () => {
      selectedPeriod = p.key;
      saveFilterState();
      initPeriodFilter();
      render();
    };

    container.appendChild(label);
  });
}

// =========================
// フィルタ判定
// =========================
function isInPeriod(task) {

  if (selectedPeriod === "all") return true;

  const end = toDate(task.plannedEnd);
  if (!end) return false;

  const now = new Date();

  const endOfWeek = new Date(now);
  endOfWeek.setDate(now.getDate() + (6 - now.getDay()));

  const endOfNextWeek = new Date(endOfWeek);
  endOfNextWeek.setDate(endOfWeek.getDate() + 7);

  const endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  const isOverdue = end < now;

  if (selectedPeriod === "thisWeek") {
    return end <= endOfWeek || isOverdue;
  }

  if (selectedPeriod === "nextWeek") {
    return end <= endOfNextWeek || isOverdue;
  }

  if (selectedPeriod === "thisMonth") {
    return end <= endOfMonth || isOverdue;
  }

  return true;
}

// =========================
// 描画
// =========================
function render() {

  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  tasks.forEach(task => {

    if (selectedUsers.size > 0 && !selectedUsers.has(task.name)) return;
    if (!isInPeriod(task)) return;

    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.task_name;
    card.draggable = true;

    card.addEventListener("dragstart", e => {
      e.dataTransfer.setData("text/plain", task.id);
    });

    document
      .querySelector(`#${task.status} .card-list`)
      ?.appendChild(card);
  });
}

// =========================
// ドラッグ
// =========================
function allowDrop(e) {
  e.preventDefault();
}

async function onDrop(e) {
  e.preventDefault();

  const id = e.dataTransfer.getData("text/plain");
  const lane = e.currentTarget.closest(".lane").id;

  const task = tasks.find(t => t.id == id);
  if (!task) return;

  task.status = lane;
  await updateExcel(task);

  tasks = await loadTasks();
  render();
}

// =========================
// Excel更新
// =========================
async function updateExcel(task) {

  const row = idRowMap[task.id];

  await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();

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