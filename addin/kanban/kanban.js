let tasks = [];
let currentTask = null;
let idRowMap = {};

// =========================
// 初期化
// =========================
Office.onReady(() => init());

async function init() {
  tasks = await loadTasks();
  render();
  initFilter();
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
      const plannedStart = row[2];
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
        plannedStart,
        plannedEnd,
        actualStart,
        actualEnd,
        status,
        order: i
      };

    }).filter(Boolean);
  });
}

// =========================
// 描画
// =========================
function render() {

  document.querySelectorAll(".card-list")
    .forEach(el => el.innerHTML = "");

  const user = document.getElementById("filter-user").value;

  tasks.forEach(task => {

    if (user && task.name !== user) return;

    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.task_name;
    card.draggable = true;
    card.dataset.id = task.id;

    card.addEventListener("dragstart", e => {
      e.dataTransfer.setData("text/plain", task.id);
    });

    card.addEventListener("click", () => openModal(task));

    // 日付表示
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatRange(task.plannedStart, task.plannedEnd);
    card.appendChild(meta);

    // 期限色
    if (task.plannedEnd) {
      const now = new Date();
      const end = new Date(task.plannedEnd);
      const diff = (end - now) / (1000*60*60*24);

      if (diff < 0) card.classList.add("overdue");
      else if (diff <= 7) card.classList.add("thisweek");
      else if (diff <= 14) card.classList.add("nextweek");
    }

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
    const today = new Date();

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

// =========================
// モーダル
// =========================
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
  const row = idRowMap[currentTask.id];

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(`O${row}`).values = [[note]];
    await context.sync();
  });

  closeModal();
}

// =========================
// フィルタ
// =========================
function initFilter() {

  const select = document.getElementById("filter-user");

  const users = [...new Set(tasks.map(t => t.name).filter(Boolean))];

  select.innerHTML = '<option value="">全員</option>';

  users.forEach(u => {
    const o = document.createElement("option");
    o.value = u;
    o.textContent = u;
    select.appendChild(o);
  });
}

function applyFilter() {
  render();
}

// =========================
// 日付表示
// =========================
function formatRange(s, e) {

  const f = d => d ? `${d.getMonth()+1}/${d.getDate()}` : "";

  const toDate = v => {
    if (!v) return null;
    if (typeof v === "number")
      return new Date((v - 25569) * 86400 * 1000);
    return new Date(v);
  };

  const sd = toDate(s);
  const ed = toDate(e);

  if (sd && ed) return `${f(sd)}～${f(ed)}`;
  if (sd) return `${f(sd)}～`;
  if (ed) return `～${f(ed)}`;
  return "";
}