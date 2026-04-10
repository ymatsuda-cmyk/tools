const APP_VERSION = "rev_20260410_11911a8";

let allTasks = [];
let currentDraggedId = null;
let currentTask = null;

let selectedUser = null;
let selectedCategory = null;
let selectedPeriod = "all";

Office.onReady(() => init());

async function init() {
  await loadExcelData();
  renderFilters();
  renderBoard();
  renderPeriodFilter();
}

// ===== データ取得 =====
async function loadExcelData() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    const rows = range.values;

    allTasks = rows.slice(1).map((row, i) => {
      const t = {
        id: row[24],
        category: row[0],
        title: row[25],
        user: row[13],
        start: row[15],
        end: row[16],
        actualStart: row[17],
        actualEnd: row[18],
        note: row[14],
        rowIndex: i + 2
      };
      t.status = getStatus(t);
      return t;
    });
  });
}

// ===== ステータス =====
function getStatus(t) {
  if (t.actualEnd) return "完了";
  if (t.actualStart) return "対応中";
  return "未着手";
}

// ===== フィルタ =====
function renderFilters() {
  renderUserFilter();
  renderCategoryFilter();
}

function renderUserFilter() {
  const users = [...new Set(allTasks.map(t => t.user).filter(v => v && v !== "#"))];
  const el = document.getElementById("user-filters");
  el.innerHTML = "";

  users.forEach(u => {
    const b = document.createElement("button");
    b.textContent = u;

    if (selectedUser === u) b.classList.add("active");

    b.onclick = () => {
      selectedUser = (selectedUser === u) ? null : u;
      renderBoard();
      renderFilters();
    };

    el.appendChild(b);
  });
}

function renderCategoryFilter() {
  const cats = [...new Set(allTasks.map(t => t.category).filter(v => v && v !== "#"))];
  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    const b = document.createElement("button");
    b.textContent = c;

    if (selectedCategory === c) b.classList.add("active");

    b.onclick = () => {
      selectedCategory = (selectedCategory === c) ? null : c;
      renderBoard();
      renderFilters();
    };

    el.appendChild(b);
  });
}

function setPeriod(p) {
  selectedPeriod = (selectedPeriod === p) ? "all" : p;
  renderBoard();
  renderPeriodFilter();
}

function renderPeriodFilter() {
  document.querySelectorAll("[data-period]").forEach(b => {
    if (b.dataset.period === selectedPeriod) {
      b.classList.add("active");
    } else {
      b.classList.remove("active");
    }
  });
}

function isMatch(t) {
  if (selectedUser && t.user !== selectedUser) return false;
  if (selectedCategory && t.category !== selectedCategory) return false;
  return true;
}

// ===== 描画 =====
function renderBoard() {
  ["todo","doing","done"].forEach(l =>
    document.querySelector(`#${l} .card-list`).innerHTML = ""
  );

  const filtered = allTasks.filter(isMatch);

  const normal = filtered
    .filter(t => t.status !== "完了")
    .sort((a,b)=>new Date(a.end)-new Date(b.end));

  const done = filtered
    .filter(t => t.status === "完了")
    .sort((a,b)=>new Date(b.actualEnd)-new Date(a.actualEnd));

  [...normal, ...done].forEach(t=>{
    const lane = getLane(t.status);
    document.querySelector(`#${lane} .card-list`).appendChild(createCard(t));
  });
}

// ===== カード =====
function createCard(t) {
  const d = document.createElement("div");
  d.className = "card";
  d.draggable = true;

  const row1 = document.createElement("div");
  row1.className = "card-row1";

  const left = document.createElement("span");
  const right = document.createElement("span");

  right.textContent = t.user || "";

  if (t.status === "未着手") {
    left.textContent = `${t.start}～${t.end}`;
  } else if (t.status === "対応中") {
    left.textContent = `${t.start}～${t.end} → ${t.actualStart}～`;
  } else {
    left.textContent = `${t.start}～${t.end} → ${t.actualStart}～${t.actualEnd}`;
  }

  row1.appendChild(left);
  row1.appendChild(right);

  const row2 = document.createElement("div");
  row2.textContent = t.title;

  d.appendChild(row1);
  d.appendChild(row2);

  applyColor(d, t);

  d.ondragstart = () => currentDraggedId = t.id;
  d.onclick = () => jumpToExcel(t.rowIndex);
  d.oncontextmenu = (e) => {
    e.preventDefault();
    openModal(t);
  };

  return d;
}

// ===== 色 =====
function applyColor(el, t) {
  if (t.status === "完了") {
    el.style.border = "2px solid #555";
    return;
  }

  if (!t.end) return;

  const end = new Date(t.end);
  const today = new Date();
  today.setHours(0,0,0,0);

  const monday = getMonday(new Date());
  const nextMonday = addDays(monday, 7);

  if (end < today) {
    el.style.border = "2px solid red";
    return;
  }

  if (end >= monday && end < nextMonday) {
    el.style.border = "2px solid green";
    return;
  }

  el.style.border = "1px solid #ccc";
}

// ===== Excel操作 =====
async function jumpToExcel(row){
  await Excel.run(async (ctx)=>{
    const s = ctx.workbook.worksheets.getItem("wbs");
    s.activate();
    s.getRange(`A${row}:Z${row}`).select();
    await ctx.sync();
  });
}

// ===== util =====
function getLane(s){
  if(s==="未着手") return "todo";
  if(s==="対応中") return "doing";
  return "done";
}

function getMonday(d){
  const t=new Date(d);
  const day=t.getDay();
  const diff=t.getDate()-day+(day===0?-6:1);
  return new Date(t.setDate(diff));
}

function addDays(d,n){
  const t=new Date(d);
  t.setDate(t.getDate()+n);
  return t;
}