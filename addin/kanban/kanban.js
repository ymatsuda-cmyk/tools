status: "todo" | "doing" | "done"
let tasks = [];

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
    const range = sheet.getRange("A2:D100");

    range.load("values");
    await context.sync();

    return range.values.map((row, i) => ({
      id: i + 1,
      row: i + 2,
      name: row[2],
      status: mapStatus(row[3]),  // StatusはExcel上では数値（1,2,3）で管理している想定
      order: i
    }));
  });
}

function render() {
  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  tasks.sort((a, b) => a.order - b.order);

  tasks.forEach(task => {
    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.name;

    card.dataset.row = task.row;
    card.dataset.id = task.id;

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatDate(task.plannedEnd);

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

function mapStatus(value) {
  switch (value) {
    case 1: return "todo";
    case 2: return "doing";
    case 3: return "done";
    default: return "todo";
  }
}