// ===== JSバージョン（GitHub Actionsで上書き）=====
const APP_VERSION = "rev_20260410_xxxxxx";

// ===== 状態管理 =====
let allTasks = [];
let currentDraggedId = null;

// フィルタ
let selectedUser = null;
let selectedCategory = null;
let selectedPeriod = "all";

// ===== 初期化 =====
Office.onReady(() => {
  init();
});

async function init() {
  await loadExcelData();
  renderFilters();
  renderBoard();
}

// ===== Excel取得 =====
async function loadExcelData() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");

    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const rows = range.values;

    allTasks = rows.slice(1).map((row, i) => ({
      id: row[24],             // Y列（主キー）
      category: row[0],        // A列
      title: row[1],
      user: row[2],
      status: row[3],
      start: row[15],          // 予定開始
      end: row[16],            // 予定終了
      actualStart: row[17],    // 実績開始
      actualEnd: row[18],      // 実績終了
      note: row[10],
      rowIndex: i + 2
    }));
  });
}

// ===== フィルタUI =====
function renderFilters() {
  renderUserFilter();
  renderCategoryFilter();
}

// 担当者（排他）
function renderUserFilter() {
  const users = [...new Set(allTasks.map(t => t.user).filter(v => v && v !== "#"))];

  const el = document.getElementById("user-filters");
  el.innerHTML = "";

  users.forEach(u => {
    const btn = document.createElement("button");
    btn.textContent = u;

    btn.onclick = () => {
      selectedUser = (selectedUser === u) ? null : u;
      renderBoard();
    };

    el.appendChild(btn);
  });
}

// 分類（排他）
function renderCategoryFilter() {
  const cats = [...new Set(allTasks.map(t => t.category).filter(v => v && v !== "#"))];

  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    const btn = document.createElement("button");
    btn.textContent = c;

    btn.onclick = () => {
      selectedCategory = (selectedCategory === c) ? null : c;
      renderBoard();
    };

    el.appendChild(btn);
  });
}

// 期間
function setPeriod(p) {
  selectedPeriod = p;
  renderBoard();
}

// ===== フィルタ判定 =====
function isMatch(task) {
  if (selectedUser && task.user !== selectedUser) return false;
  if (selectedCategory && task.category !== selectedCategory) return false;

  if (selectedPeriod === "all") return true;

  const today = new Date();
  const startOfWeek = getMonday(today);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(endOfWeek.getDate() + 6);

  const nextWeekStart = new Date(startOfWeek);
  nextWeekStart.setDate(nextWeekStart.getDate() + 7);

  const nextWeekEnd = new Date(nextWeekStart);
  nextWeekEnd.setDate(nextWeekEnd.getDate() + 6);

  const end = new Date(task.end);

  switch (selectedPeriod) {
    case "past":
      return end < startOfWeek;

    case "week":
      return end >= startOfWeek && end <= endOfWeek;

    case "nextweek":
      return end >= nextWeekStart && end <= nextWeekEnd;

    case "future":
      return end > nextWeekEnd;

    default:
      return true;
  }
}

// ===== ボード描画 =====
function renderBoard() {
  const lanes = ["todo", "doing", "done"];

  lanes.forEach(l => {
    document.querySelector(`#${l} .card-list`).innerHTML = "";
  });

  const filtered = allTasks
    .filter(isMatch)
    .sort((a, b) => new Date(a.end) - new Date(b.end));

  filtered.forEach(task => {
    const card = createCard(task);
    const lane = getLane(task.status);
    document.querySelector(`#${lane} .card-list`).appendChild(card);
  });
}

// ===== カード作成 =====
function createCard(task) {
  const div = document.createElement("div");
  div.className = "card";
  div.draggable = true;

  div.innerHTML = `
    <div>${task.title}</div>
    <small>${task.user || ""}</small>
  `;

  // 色
  applyColor(div, task);

  // ドラッグ
  div.ondragstart = () => currentDraggedId = task.id;

  // クリック（Excelジャンプ）
  div.onclick = () => jumpToExcel(task.rowIndex);

  // ダブルクリック（モーダル）
  div.ondblclick = () => openModal(task);

  return div;
}

// ===== 色判定（週ベース）=====
function applyColor(el, task) {
  if (task.status === "完了") {
    el.style.border = "2px solid #555";
    return;
  }

  const today = new Date();
  const startOfWeek = getMonday(today);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(endOfWeek.getDate() + 6);

  const nextWeekStart = new Date(startOfWeek);
  nextWeekStart.setDate(nextWeekStart.getDate() + 7);

  const nextWeekEnd = new Date(nextWeekStart);
  nextWeekEnd.setDate(nextWeekEnd.getDate() + 6);

  const end = new Date(task.end);

  if (end < today) {
    el.style.border = "2px solid red";
  } else if (end >= startOfWeek && end <= endOfWeek) {
    el.style.border = "2px solid green"; // 今週
  }
}

// ===== ドラッグ =====
function allowDrop(e) {
  e.preventDefault();
}

function drop(e, status) {
  e.preventDefault();

  const task = allTasks.find(t => t.id === currentDraggedId);
  if (!task) return;

  updateStatus(task, status);
}

// ===== ステータス更新 =====
async function updateStatus(task, status) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");

    const row = task.rowIndex;

    sheet.getRange(`D${row}`).values = [[status]];

    const today = new Date().toISOString().split("T")[0].replace(/-/g, "/");

    if (status === "対応中") {
      sheet.getRange(`R${row}`).values = [[today]];
      sheet.getRange(`S${row}`).values = [[""]];
    }

    if (status === "完了") {
      if (!task.actualStart) {
        sheet.getRange(`R${row}`).values = [[today]];
      }
      sheet.getRange(`S${row}`).values = [[today]];
    }

    await context.sync();
  });

  init();
}

// ===== Excelジャンプ =====
async function jumpToExcel(row) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    sheet.activate();
    const range = sheet.getRange(`A${row}`);
    range.select();
    await context.sync();
  });
}

// ===== モーダル =====
let currentTask = null;

function openModal(task) {
  currentTask = task;

  document.getElementById("modal-title").textContent = task.title;
  document.getElementById("modal-note").value = task.note || "";

  document.getElementById("modal").classList.remove("hidden");
}

function closeModal() {
  document.getElementById("modal").classList.add("hidden");
}

// 備考保存（列幅維持）
async function saveNote() {
  const note = document.getElementById("modal-note").value;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    const cell = sheet.getRange(`K${currentTask.rowIndex}`);

    cell.values = [[note]];

    await context.sync();
  });

  closeModal();
  init();
}

// ===== util =====
function getLane(status) {
  if (status === "未着手") return "todo";
  if (status === "対応中") return "doing";
  return "done";
}

function getMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}