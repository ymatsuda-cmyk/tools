// ===== kanban.js FINAL =====

let tasks = [];
let currentTask = null;
let idRowMap = {};

document.addEventListener("keydown", e => {
  if (e.key === "Escape") closeModal();
});

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("modal")?.addEventListener("click", e => {
    if (e.target.id === "modal") closeModal();
  });
});

Office.onReady(() => {
  init();
});

function isOfficeAvailable() {
  return typeof Office !== "undefined";
}

async function init() {
  if (isOfficeAvailable()) {
    tasks = await loadTasks();
  } else {
    tasks = [
      { id:1, row:2, task_name:"タスクA", status:"todo", order:1 },
      { id:2, row:3, task_name:"タスクB", status:"doing", order:2 },
      { id:3, row:4, task_name:"タスクC", status:"done", order:3 }
    ];
  }

  render();
  initFilter();
}

async function loadTasks() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("N11:Z1000");

    range.load("values");
    await context.sync();

    const rows = range.values;
    idRowMap = {};

    const tasks = rows.map((row, i) => {
      const name = row[0];
      const note = row[1];
      const plannedStart = row[2];
      const plannedEnd = row[3];
      const actualStart = row[4];
      const actualEnd = row[5];
      const id = row[11];
      const task_name = row[12];

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
        task_name,
        name,
        note,
        plannedStart,
        plannedEnd,
        actualStart,
        actualEnd,
        status,
        order: i
      };
    }).filter(Boolean);

    return tasks;
  });
}

function render() {
  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  const filterUser = document.getElementById("filter-user")?.value;

  tasks.sort((a, b) => a.order - b.order);

  tasks.forEach(task => {
    if (filterUser && task.name !== filterUser) return;

    const card = document.createElement("div");
    card.dataset.id = task.id;
    card.className = "card";
    card.textContent = task.task_name;
    card.draggable = true;

    card.addEventListener("dragstart", e => onDragStart(e, task.id));
    card.addEventListener("click", () => openModal(task));

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatRange(task.plannedStart, task.plannedEnd);

    // 期限ハイライト
    if (task.plannedEnd) {
      const now = new Date();
      const end = new Date(task.plannedEnd);
      const diff = (end - now) / (1000 * 60 * 60 * 24);

      if (diff < 0) card.classList.add("overdue");
      else if (diff <= 7) card.classList.add("thisweek");
      else if (diff <= 14) card.classList.add("nextweek");
    }

    card.appendChild(meta);

    const column = document.querySelector(`#${task.status} .card-list`);
    if (column) column.appendChild(card);
  });
}

function formatRange(start, end) {
  const f = d => d ? `${d.getMonth()+1}/${d.getDate()}` : "";

  const toDate = v => {
    if (!v) return null;
    if (typeof v === "number") return new Date((v - 25569) * 86400 * 1000);
    return new Date(v);
  };

  const s = toDate(start);
  const e = toDate(end);

  if (s && e) return `${f(s)}～${f(e)}`;
  if (s) return `${f(s)}～`;
  if (e) return `～${f(e)}`;
  return "";
}

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
  const newNote = document.getElementById("modal-note").value;
  const row = idRowMap[currentTask.id];
  if (!row) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(`O${row}`).values = [[newNote]];
    await context.sync();
  });

  closeModal();
}

function allowDrop(e) {
  e.preventDefault();
}

function onDragStart(e, taskId) {
  e.dataTransfer.setData("text/plain", taskId);
}

async function onDrop(e) {
  e.preventDefault();

  const taskId = e.dataTransfer.getData("text/plain");
  const newStatus = e.currentTarget.id;

  const task = tasks.find(t => t.id == taskId);
  if (!task) return;

  task.status = newStatus;

  await updateExcelStatus(task);

  tasks = await loadTasks();
  render();
}

async function updateExcelStatus(task) {
  const row = idRowMap[task.id];
  if (!row) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const rCell = sheet.getRange(`R${row}`);
    const sCell = sheet.getRange(`S${row}`);

    rCell.load("values");
    sCell.load("values");
    await context.sync();

    const existingR = rCell.values[0][0];
    const existingS = sCell.values[0][0];
    const today = new Date();

    if (task.status === "todo") {
      rCell.values = [[""]];
      sCell.values = [[""]];
    }

    if (task.status === "doing") {
      rCell.values = [[existingR || today]];
      sCell.values = [[""]];
    }

    if (task.status === "done") {
      rCell.values = [[existingR || today]];
      sCell.values = [[existingS || today]];
    }

    await context.sync();
  });
}

function initFilter() {
  const select = document.getElementById("filter-user");
  if (!select) return;

  const users = [...new Set(tasks.map(t => t.name).filter(Boolean))];
  select.innerHTML = '<option value="">全員</option>';

  users.forEach(u => {
    const opt = document.createElement("option");
    opt.value = u;
    opt.textContent = u;
    select.appendChild(opt);
  });
}

function applyFilter() {
  render();
}