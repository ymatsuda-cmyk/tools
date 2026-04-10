// ===== JSバージョン（自動更新）=====
const APP_VERSION = "rev_20260410_xxxxxx";

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
}

// ===== データ取得 =====
async function loadExcelData() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const rows = range.values;

    allTasks = rows.slice(1).map((row, i) => {
      const task = {
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

      task.status = getStatus(task);
      return task;
    });
  });
}

// ===== ステータス判定 =====
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
    b.onclick = () => {
      selectedUser = (selectedUser === u) ? null : u;
      renderBoard();
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
    b.onclick = () => {
      selectedCategory = (selectedCategory === c) ? null : c;
      renderBoard();
    };
    el.appendChild(b);
  });
}

function setPeriod(p) {
  selectedPeriod = p;
  renderBoard();
}

function isMatch(t) {
  if (selectedUser && t.user !== selectedUser) return false;
  if (selectedCategory && t.category !== selectedCategory) return false;

  if (selectedPeriod === "all") return true;

  const end = new Date(t.end);
  const today = new Date();
  const monday = getMonday(today);

  const nextWeek = new Date(monday);
  nextWeek.setDate(nextWeek.getDate() + 7);

  switch (selectedPeriod) {
    case "past": return end < monday;
    case "week": return end >= monday && end < nextWeek;
    case "nextweek": return end >= nextWeek && end < addDays(nextWeek, 7);
    case "future": return end >= addDays(nextWeek, 7);
  }
  return true;
}

// ===== 描画 =====
function renderBoard() {
  ["todo","doing","done"].forEach(l =>
    document.querySelector(`#${l} .card-list`).innerHTML = ""
  );

  allTasks
    .filter(isMatch)
    .sort((a,b)=>new Date(a.end)-new Date(b.end))
    .forEach(t=>{
      const lane = getLane(t.status);
      document.querySelector(`#${lane} .card-list`).appendChild(createCard(t));
    });
}

function createCard(t) {
  const d = document.createElement("div");
  d.className = "card";
  d.draggable = true;

  d.innerHTML = `<div>${t.title}</div><small>${t.user||""}</small>`;

  applyColor(d, t);

  d.ondragstart = () => currentDraggedId = t.id;

  // 左クリック → Excel
  d.onclick = () => jumpToExcel(t.rowIndex);

  // 右クリック → モーダル
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

  const end = new Date(t.end);
  const monday = getMonday(new Date());
  const next = addDays(monday,7);

  if (end < new Date()) el.style.border = "2px solid red";
  else if (end >= monday && end < next) el.style.border = "2px solid green";
}

// ===== DnD =====
function allowDrop(e){e.preventDefault();}

function drop(e, lane){
  e.preventDefault();
  const t = allTasks.find(x=>x.id===currentDraggedId);
  if(t) updateStatus(t,lane);
}

// ===== Excel更新 =====
async function updateStatus(t, lane){
  await Excel.run(async (ctx)=>{
    const s = ctx.workbook.worksheets.getItem("wbs");
    const r = t.rowIndex;
    const today = new Date().toISOString().split("T")[0].replace(/-/g,"/");

    if(lane==="doing"){
      s.getRange(`R${r}`).values=[[today]];
      s.getRange(`S${r}`).values=[[""]];
    }
    if(lane==="done"){
      if(!t.actualStart) s.getRange(`R${r}`).values=[[today]];
      s.getRange(`S${r}`).values=[[today]];
    }
    if(lane==="todo"){
      s.getRange(`R${r}`).values=[[""]];
      s.getRange(`S${r}`).values=[[""]];
    }

    await ctx.sync();
  });
  init();
}

// ===== Excelジャンプ =====
async function jumpToExcel(row){
  await Excel.run(async (ctx)=>{
    const s = ctx.workbook.worksheets.getItem("wbs");
    s.activate();
    s.getRange(`A${row}:Z${row}`).select();
    await ctx.sync();
  });
}

// ===== モーダル =====
function openModal(t){
  currentTask = t;
  document.getElementById("modal-title").textContent = t.title;
  document.getElementById("modal-note").value = t.note||"";
  document.getElementById("modal").classList.remove("hidden");
}

function closeModal(){
  document.getElementById("modal").classList.add("hidden");
}

// 行高さ固定
async function saveNote(){
  const note = document.getElementById("modal-note").value;

  await Excel.run(async (ctx)=>{
    const s = ctx.workbook.worksheets.getItem("wbs");
    const r = currentTask.rowIndex;

    const rowRange = s.getRange(`A${r}:Z${r}`);
    rowRange.load("rowHeight");
    await ctx.sync();

    const h = rowRange.rowHeight;

    const cell = s.getRange(`O${r}`);
    cell.values=[[note]];
    cell.format.wrapText=false;

    await ctx.sync();

    rowRange.rowHeight = h;
    await ctx.sync();
  });

  closeModal();
  init();
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