status: "todo" | "doing" | "done"
let tasks = [];
let currentTask = null;
let idRowMap = {};

document.addEventListener("keydown", e => {
  if (e.key === "Escape") closeModal();
});

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("modal").addEventListener("click", e => {
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
    console.warn("⚠ Officeなし → モックデータ");

    tasks = [
      { id:1, row:2, name:"タスクA", status:"todo", order:1 },
      { id:2, row:3, name:"タスクB", status:"doing", order:2 },
      { id:3, row:4, name:"タスクC", status:"done", order:3 }
    ];
  }
  console.log("🔥 tasks:", tasks);
  render();
}

async function loadTasks() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("N11:Z1000");

    range.load("values");
    await context.sync();

    const rows = range.values;

    const tasks = rows
      .map((row, i) => {
        const name = row[0];          // N 担当者
        const note = row[1];          // O 備考
        const plannedStart = row[2];  // P 予定開始日
        const plannedEnd = row[3];    // Q 予定終了日
        const actualStart = row[4];   // R 実際開始日
        const actualEnd = row[5];     // S 実際終了日
        const priority = row[6];      // T 優先度
        const id = row[11];           // Y ID ←主キー
        const task_name = row[12];    // Z タスク名

        // 🔥 空タスク除外
        if (!task_name) return null;

        const rowNumber = i + 11;
        // 🔥 ここでマップ作る
        const safeId = id || `row-${rowNumber}`;
        idRowMap[safeId] = rowNumber;

        // 🔥 ステータス判定（超重要）
        let status = "todo";
        if (actualEnd) {
          status = "done";
        } else if (actualStart) {
          status = "doing";
        }

        if (!id) {
          console.warn("ID missing at row:", i + 11);
        }
        return {
          id: safeId,
          row: rowNumber, // ← Excel行番号（N11開始）

          task_name,
          name,
          note,

          plannedStart,
          plannedEnd,
          actualStart,
          actualEnd,

          priority,

          status,
          order: i
        };
      })
      .filter(Boolean); // null除外

    // 🔥 デバッグ
    console.log("🔥 tasks:", tasks);
    window.tasks = tasks;

    return tasks;
  });
}

function render() {
  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  tasks.sort((a, b) => a.order - b.order);

  tasks.forEach(task => {
    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.task_name;

    card.addEventListener("click", () => openModal(task));

    const meta = document.createElement("div");
    meta.className = "meta";

    meta.textContent = formatRange(
      task.plannedStart,
      task.plannedEnd
    );

    card.appendChild(meta);

    document
      .querySelector(`#${task.status} .card-list`)
      ?.appendChild(card);
  });
}

function formatDate(date) {
  if (!date) return "";
  const d = new Date(date);
  return `${d.getMonth()+1}/${d.getDate()}`;
}

function formatRange(start, end) {
  if (!start && !end) return "";

  const toDate = (v) => {
    if (!v) return null;

    // Excelシリアル対応
    if (typeof v === "number") {
      return new Date((v - 25569) * 86400 * 1000);
    }

    return new Date(v);
  };

  const format = (d) => {
    if (!d) return "";
    return `${d.getMonth() + 1}/${d.getDate()}`;
  };

  const s = toDate(start);
  const e = toDate(end);

  if (s && e) return `${format(s)}～${format(e)}`;
  if (s) return `${format(s)}～`;
  if (e) return `～${format(e)}`;
  return "";
}

function mapStatus(value) {
  switch (value) {
    case 1: return "todo";
    case 2: return "doing";
    case 3: return "done";
    default: return "todo";
  }
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

  if (!row) {
    alert("行が見つかりません");
    return;
  }

  currentTask.note = newNote;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const cell = sheet.getRange(`O${row}`);
    cell.values = [[newNote]];

    await context.sync();
  });

  closeModal();
}

async function findRowById(context, sheet, targetId) {
  const range = sheet.getRange("Y11:Y1000");
  range.load("values");

  await context.sync();

  const values = range.values;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] == targetId) {
      return i + 11; // ← 行番号
    }
  }

  throw new Error("ID not found: " + targetId);
}