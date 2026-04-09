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

      const category = row[0];
      const name = row[13];
      const note = row[14];
      const plannedStart = row[15];
      const plannedEnd = row[16];
      const actualStart = row[17];
      const actualEnd = row[18];
      const priority = row[19];
      const id = row[24];
      const task_name = row[25];

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

    card.addEventListener("click", () => {
      clickTimer = setTimeout(() => focusRow(task.id), 200);
    });

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
// 🎯 ここが修正ポイント（月曜開始）
// =====================
function applyDeadlineColor(card, endDate) {

  if (!endDate) return;

  const end = new Date(endDate);
  const today = new Date();

  today.setHours(0,0,0,0);
  end.setHours(0,0,0,0);

  // 期限切れ
  if (end < today) {
    card.classList.add("overdue");
    return;
  }

  // 月曜開始の週計算
  const day = today.getDay() || 7; // 日曜=7扱い
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - day + 1);

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);

  const startOfNextWeek = new Date(startOfWeek);
  startOfNextWeek.setDate(startOfWeek.getDate() + 7);

  const endOfNextWeek = new Date(startOfWeek);
  endOfNextWeek.setDate(startOfWeek.getDate() + 13);

  if (end >= startOfWeek && end <= endOfWeek) {
    card.classList.add("thisweek");
  }
  else if (end >= startOfNextWeek && end <= endOfNextWeek) {
    card.classList.add("nextweek");
  }
}

// =====================
// その他（省略なし）
// =====================
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