const APP_VERSION = "rev_20260410_98d877f";

// window.APP_VERSIONも設定してindex.htmlから参照可能にする
window.APP_VERSION = APP_VERSION;

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
  
  // バージョン表示を更新
  if (typeof updateVersionDisplay === 'function') {
    updateVersionDisplay();
  }
}

// ===== Excel日付変換 =====
function excelDateToJS(value) {
  if (!value) return null;
  if (typeof value === "number") {
    return new Date((value - 25569) * 86400 * 1000);
  }
  return new Date(value);
}

function fmt(v) {
  const d = excelDateToJS(v);
  if (!d || isNaN(d)) return "";
  return `${d.getMonth()+1}/${d.getDate()}`;
}

// ===== データ取得 =====
async function loadExcelData() {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    const rows = range.values;

    allTasks = rows.slice(10).map((row, i) => {
      if (!row[25] || row[19] === "-") return null;

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
        rowIndex: i + 11,

        isNoSchedule: !row[15] && !row[16],
        isStar: row[14] && row[14].toString().startsWith('★')  // 備考の先頭に★があるかチェック
      };

      t.status = getStatus(t);
      return t;
    }).filter(x => x);
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
  const users = [...new Set(
    allTasks
      .map(t => t.user)
      .filter(v => v && v !== "#")
  )];

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
  const cats = [...new Set(
    allTasks
      .map(t => t.category)
      .filter(v => v && v !== "#")
  )];

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
    b.classList.toggle("active", b.dataset.period === selectedPeriod);
  });
}

// ===== 描画 =====
function renderBoard() {
  ["todo","doing","done"].forEach(l =>
    document.querySelector(`#${l} .card-list`).innerHTML = ""
  );

  const filtered = allTasks.filter(isMatch);

  const normal = filtered
    .filter(t => t.status !== "完了")
    .sort((a, b) => {
      // スター付きを優先
      if (a.isStar && !b.isStar) return -1;
      if (!a.isStar && b.isStar) return 1;
      // 同じスター状態なら期限日順
      return excelDateToJS(a.end) - excelDateToJS(b.end);
    });

  const done = filtered
    .filter(t => t.status === "完了")
    .sort((a,b)=>excelDateToJS(b.actualEnd)-excelDateToJS(a.actualEnd));

  [...normal, ...done].forEach(t=>{
    const lane = getLane(t.status);
    document.querySelector(`#${lane} .card-list`).appendChild(createCard(t));
  });

  setupDnD();
}

// ===== カード =====
function createCard(t) {
  const d = document.createElement("div");
  d.className = "card";
  d.draggable = true;

  d.addEventListener("dragstart", (e) => {
    currentDraggedId = t.id;
    e.dataTransfer.setData("text/plain", t.id);
    d.classList.add("dragging");
  });

  d.addEventListener("dragend", () => {
    d.classList.remove("dragging");
  });

  d.addEventListener("click", (e) => {
    if (e.button !== 0) return;
    jumpToExcel(t.rowIndex);
  });

  d.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    e.stopPropagation();
    openModal(t);
  });

  const row1 = document.createElement("div");
  row1.className = "card-row1";

  const left = document.createElement("span");
  const right = document.createElement("span");
  const star = document.createElement("span");
  
  star.className = "star-icon";
  star.textContent = t.isStar ? "★" : "☆";
  star.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    toggleStar(t);
  });

  right.textContent = t.user || "";

  // ★ ここ修正（重要）
  if (t.isNoSchedule) {
    left.textContent = "TODO";
  } else if (t.status === "未着手") {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)}`;
  } else if (t.status === "対応中") {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)} → ${fmt(t.actualStart)}～`;
  } else {
    left.textContent = `${fmt(t.start)}～${fmt(t.end)} → ${fmt(t.actualStart)}～${fmt(t.actualEnd)}`;
  }

  row1.appendChild(left);
  row1.appendChild(right);
  row1.appendChild(star);

  const row2 = document.createElement("div");
  row2.textContent = t.title;

  d.appendChild(row1);
  d.appendChild(row2);

  // スター状態に応じてカードスタイルを適用 
  if (t.isStar) {
    d.classList.add("starred");
  }

  applyColor(d, t);

  return d;
}

// ===== 色 =====
function applyColor(el, t) {
  if (t.status === "完了") {
    el.style.border = "2px solid #333";
    return;
  }

  const start = excelDateToJS(t.start);
  const end = excelDateToJS(t.end);

  if (!start || !end) return;

  const today = new Date();
  today.setHours(0,0,0,0);

  // ★ 遅延
  if (end < today) {
    el.style.border = "2px solid red";
    return;
  }

  // ★ 期間内（←ここが今回のポイント）
  if (start <= today && end >= today) {
    el.style.border = "2px solid green";
    return;
  }

  el.style.border = "1px solid #ccc";
}

// ===== DnD =====
function setupDnD() {
  ["todo","doing","done"].forEach(id=>{
    const lane = document.getElementById(id);

    lane.ondragover = (e)=>e.preventDefault();

    lane.ondrop = (e)=>{
      e.preventDefault();
      const t = allTasks.find(x=>x.id===currentDraggedId);
      if (t) updateStatus(t, id);
    };
  });
}

// ===== Excel =====
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

// JavaScriptのDateをExcelシリアル値に変換
function dateToExcelSerial(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return "";
  
  // Excel epoch: 1900年1月1日
  const excelEpoch = new Date(1900, 0, 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  
  // 日数差を計算
  const daysDiff = Math.floor((date - excelEpoch) / msPerDay);
  
  // Excelの1900年うるう年バグを考慮（1900年3月1日以降は+1）
  return daysDiff + (date >= new Date(1900, 2, 1) ? 2 : 1);
}

async function updateStatus(task, lane) {
  let actualStart = task.actualStart;
  let actualEnd = task.actualEnd;

  if (lane === "todo") {
    actualStart = "";
    actualEnd = "";
  }

  if (lane === "doing") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = "";
  }

  if (lane === "done") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = new Date();
    
    // 完了時は★→☆に変更
    if (task.isStar) {
      task.isStar = false;
    }
  }

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;

    const startCell = sheet.getRange(`R${row}`);
    const endCell = sheet.getRange(`S${row}`);

    // Date型をExcelシリアル値に変換して設定
    startCell.values = [[dateToExcelSerial(actualStart)]];
    endCell.values = [[dateToExcelSerial(actualEnd)]];

    // 表示形式をm/d に設定
    startCell.numberFormat = [["m/d"]];
    endCell.numberFormat = [["m/d"]];

    // 完了時に備考から★を削除
    if (lane === "done" && task.note && task.note.includes('★')) {
      let newNote = (task.note || "").replace(/★/g, "");
      const noteCell = sheet.getRange(`O${row}`);
      noteCell.values = [[newNote]];
      noteCell.format.wrapText = false;
      task.note = newNote; // タスクの備考も更新
    }

    await ctx.sync();
  });

  await init();
}

function isValidDate(v) {
  return v instanceof Date && !isNaN(v);
}

// ===== スター切り替え =====
async function toggleStar(task) {
  // スター状態を切り替え
  task.isStar = !task.isStar;
  
  // 備考を更新
  let newNote = task.note || "";
  
  if (task.isStar) {
    // ★を★に変更：先頭に★を付与
    if (!newNote.startsWith('★')) {
      newNote = '★' + newNote;
    }
  } else {
    // ★を☆に変更：備考から★を完全に削除（""に置換）
    newNote = newNote.replace(/★/g, "");
  }
  
  // Excelに備考を更新
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;
    const cell = sheet.getRange(`O${row}`);
    
    cell.values = [[newNote]];
    cell.format.wrapText = false;
    
    await ctx.sync();
  });
  
  // タスクの備考を更新
  task.note = newNote;
  
  // 画面を再描画
  renderBoard();
}

function openModal(task) {
  currentTask = task;

  document.getElementById("modal-title").textContent = task.title;
  document.getElementById("modal-note").value = task.note || "";

  document.getElementById("modal").classList.remove("hidden");
}

function closeModal() {
  document.getElementById("modal").classList.add("hidden");
}

async function saveNote() {
  const note = document.getElementById("modal-note").value;

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = currentTask.rowIndex;

    const cell = sheet.getRange(`O${row}`);

    cell.values = [[note]];

    // ★これ追加
    cell.format.wrapText = false;

    // ★行高さ固定（例：20）
    const entireRow = sheet.getRange(`${row}:${row}`);
    entireRow.format.rowHeight = 20;

    await ctx.sync();
  });

  closeModal();
  await init();
}

function isMatch(t) {

  // 担当者
  if (selectedUser && t.user !== selectedUser) return false;

  // 分類
  if (selectedCategory && t.category !== selectedCategory) return false;

  // ★ 本日フィルタ：スター付きのみ表示
  if (selectedPeriod === "today") {
    return t.isStar;
  }

  // ★ 日付なし（TODO）
  if (t.isNoSchedule) {
    return selectedPeriod === "all";
  }

  const start = excelDateToJS(t.start);
  const end = excelDateToJS(t.end);

  if (!start || !end) return false;

  const today = new Date();
  today.setHours(0,0,0,0);

  const monday = getMonday(today);
  const sunday = addDays(monday, 6);
  const nextMonday = addDays(monday, 7);
  const nextSunday = addDays(monday, 13);

  switch (selectedPeriod) {

    case "past":
      return end < monday;

    case "week":
      return (start <= sunday && end >= monday);

    case "nextweek":
      return (start <= nextSunday && end >= nextMonday);

    case "future":
      return start > nextSunday;

    case "all":
    default:
      return true;
  }
}
