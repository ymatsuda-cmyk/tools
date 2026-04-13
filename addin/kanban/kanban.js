const APP_VERSION = "rev_20260413_c9a2dcd";

// window.APP_VERSIONも設定してindex.htmlから参照可能にする
window.APP_VERSION = APP_VERSION;

let allTasks = [];
let currentDraggedId = null;
let currentTask = null;

let selectedUser = null;
let selectedCategory = null;
let selectedPeriod = "all";

Office.onReady(() => {
  // 保存されたサイズを復元
  restoreSavedSize();
  
  // 保存されたフィルター設定を復元
  restoreSavedFilters();
  
  // サイズ変更の監視を開始
  setupSizeMonitoring();
  
  init();
});

// ===== サイズ記憶機能 =====
function restoreSavedSize() {
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    if (savedSize) {
      const size = JSON.parse(savedSize);
      
      // 最小サイズの制限
      const minWidth = 400;
      const minHeight = 300;
      const width = Math.max(size.width || 400, minWidth);
      const height = Math.max(size.height || 600, minHeight);
      
      // DOM要素のサイズを設定
      document.documentElement.style.minWidth = width + "px";
      document.body.style.minWidth = width + "px";
      
      // Office APIを使用してタスクペインのサイズを設定（可能な場合）
      if (Office.context.requirements.isSetSupported('TaskPaneApp', '1.1')) {
        try {
          Office.addin.setTaskpaneSize(width, height);
        } catch (e) {
          console.log("TaskPane resize not supported:", e);
        }
      }
      
      // 親ウィンドウへのサイズヒント
      if (window.parent && window.parent.postMessage) {
        window.parent.postMessage({
          type: 'resize',
          width: width,
          height: height
        }, '*');
      }
      
      console.log(`Restored size: ${width}x${height}`);
    } else {
      // デフォルトサイズを設定
      setDefaultSize();
    }
  } catch (e) {
    console.log("Size restoration error:", e);
    setDefaultSize();
  }
}

function setDefaultSize() {
  const defaultWidth = 400;
  const defaultHeight = 600;
  
  document.documentElement.style.minWidth = defaultWidth + "px";
  document.body.style.minWidth = defaultWidth + "px";
  
  if (window.parent && window.parent.postMessage) {
    window.parent.postMessage({
      type: 'resize',
      width: defaultWidth,
      height: defaultHeight
    }, '*');
  }
}

function setupSizeMonitoring() {
  let saveTimeout;
  
  // ResizeObserverでサイズ変更を監視
  if (window.ResizeObserver) {
    const resizeObserver = new ResizeObserver(entries => {
      for (const entry of entries) {
        const { width, height } = entry.contentRect;
        
        // デバウンス処理（連続した変更を制限）
        clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
          saveSizeToStorage(Math.round(width), Math.round(height));
        }, 500);
      }
    });
    
    // body要素を監視
    resizeObserver.observe(document.body);
  } else {
    // ResizeObserverが利用できない場合はwindowのresizeイベントを使用
    let lastWidth = window.innerWidth;
    let lastHeight = window.innerHeight;
    
    window.addEventListener('resize', () => {
      clearTimeout(saveTimeout);
      saveTimeout = setTimeout(() => {
        const currentWidth = window.innerWidth; 
        const currentHeight = window.innerHeight;
        
        if (currentWidth !== lastWidth || currentHeight !== lastHeight) {
          saveSizeToStorage(currentWidth, currentHeight);
          lastWidth = currentWidth;
          lastHeight = currentHeight;
        }
      }, 500);
    });
  }
}

function saveSizeToStorage(width, height) {
  try {
    const sizeData = {
      width: width,
      height: height,
      timestamp: Date.now()
    };
    
    localStorage.setItem('kanban-taskpane-size', JSON.stringify(sizeData));
    console.log(`Saved size: ${width}x${height}`);
  } catch (e) {
    console.log("Size saving error:", e);
  }
}

// サイズ設定をリセットする関数
function resetSize() {
  try {
    localStorage.removeItem('kanban-taskpane-size');
    localStorage.removeItem('kanban-filters');
    console.log("Size and filter settings reset");
    
    // ページをリロードして新しい設定を適用
    window.location.reload();
  } catch (e) {
    console.log("Reset error:", e);
  }
}

// ===== フィルター記憶機能 =====
function restoreSavedFilters() {
  try {
    const savedFilters = localStorage.getItem('kanban-filters');
    if (savedFilters) {
      const filters = JSON.parse(savedFilters);
      
      selectedUser = filters.user || null;
      selectedCategory = filters.category || null;
      selectedPeriod = filters.period || "all";
      
      console.log('Restored filters:', filters);
    }
  } catch (e) {
    console.log("Filter restoration error:", e);
    // エラー時はデフォルト値を維持
    selectedUser = null;
    selectedCategory = null; 
    selectedPeriod = "all";
  }
}

function saveFilters() {
  try {
    const filterData = {
      user: selectedUser,
      category: selectedCategory,
      period: selectedPeriod,
      timestamp: Date.now()
    };
    
    localStorage.setItem('kanban-filters', JSON.stringify(filterData));
    console.log('Saved filters:', filterData);
  } catch (e) {
    console.log("Filter saving error:", e);
  }
}

async function init() {
  // 保存されたサイズまたはデフォルトサイズを適用
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    let minWidth = 400;
    
    if (savedSize) {
      const size = JSON.parse(savedSize);
      minWidth = Math.max(size.width || 400, 400);
    }
    
    document.documentElement.style.minWidth = minWidth + "px";
    document.body.style.minWidth = minWidth + "px";
  } catch (e) {
    // エラー時はデフォルトサイズ
    document.documentElement.style.minWidth = "400px";
    document.body.style.minWidth = "400px";
  }
  
  // ボードコンテナの幅も確実に設定
  const boardEl = document.getElementById("board");
  if (boardEl) {
    boardEl.style.minWidth = "350px";
    boardEl.style.width = "100%";
  }

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
        classification: row[1],  // B列の分類を追加
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
      saveFilters(); // フィルタ設定を保存
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
      saveFilters(); // フィルタ設定を保存
      renderBoard();
      renderFilters();
    };

    el.appendChild(b);
  });
}

function setPeriod(p) {
  selectedPeriod = (selectedPeriod === p) ? "all" : p;
  saveFilters(); // フィルタ設定を保存
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
  const rightGroup = document.createElement("span");
  rightGroup.className = "right-group";
  
  const user = document.createElement("span");
  user.className = "user-name";
  user.textContent = t.user || "";

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

  // 右グループにユーザー名を追加
  rightGroup.appendChild(user);
  
  // 完了状態以外にのみスターアイコンを追加
  if (t.status !== "完了") {
    const star = document.createElement("span");
    star.className = "star-icon";
    star.textContent = t.isStar ? "★" : "☆";
    
    // ★の場合は金色クラスを追加
    if (t.isStar) {
      star.classList.add("filled");
    }
    
    star.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleStar(t);
    });
    rightGroup.appendChild(star);
  }

  // 日付情報と右グループ（ユーザー名+スター）を追加
  row1.appendChild(left);
  row1.appendChild(rightGroup);

  // タイトル行（タイトル + 分類）
  const row2 = document.createElement("div");
  row2.className = "card-title-row";
  
  const titleSpan = document.createElement("span");
  titleSpan.className = "card-title";
  titleSpan.textContent = t.title;
  
  const classificationSpan = document.createElement("span");
  classificationSpan.className = "card-classification";
  if (t.classification && t.classification.trim() !== "") {
    classificationSpan.textContent = `<${t.classification}>`;
  }
  
  row2.appendChild(titleSpan);
  row2.appendChild(classificationSpan);

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

  const startRaw = excelDateToJS(t.start);
  const endRaw = excelDateToJS(t.end);

  if (!startRaw || !endRaw) return;

  // 時刻情報を除去して日付のみで比較
  const start = new Date(startRaw);
  start.setHours(0,0,0,0);
  
  const end = new Date(endRaw);
  end.setHours(0,0,0,0);
  
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
    s.getRange(`${row}:${row}`).select();
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

  const modal = document.getElementById("modal");
  modal.classList.remove("hidden");
  
  // Escキーでモーダルを閉じる
  const handleEscKey = (event) => {
    if (event.key === 'Escape') {
      closeModal();
    }
  };
  
  // モーダル外クリックで閉じる
  const handleOverlayClick = (event) => {
    if (event.target === modal) {
      closeModal();
    }
  };
  
  // モーダルコンテンツ内のクリックでイベント伝播を止める
  const modalContent = modal.querySelector('.modal-content');
  const handleContentClick = (event) => {
    event.stopPropagation();
  };
  
  // イベントリスナーを追加
  document.addEventListener('keydown', handleEscKey);
  modal.addEventListener('click', handleOverlayClick);
  modalContent.addEventListener('click', handleContentClick);
  
  // クリーンアップ関数をモーダルに保存
  modal._cleanup = () => {
    document.removeEventListener('keydown', handleEscKey);
    modal.removeEventListener('click', handleOverlayClick);
    modalContent.removeEventListener('click', handleContentClick);
  };
  
  // テキストエリアにフォーカス
  setTimeout(() => {
    document.getElementById("modal-note").focus();
  }, 100);
}

function closeModal() {
  const modal = document.getElementById("modal");
  modal.classList.add("hidden");
  
  // イベントリスナーをクリーンアップ
  if (modal._cleanup) {
    modal._cleanup();
    modal._cleanup = null;
  }
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
    return selectedPeriod === "all" || selectedPeriod === "todo";
  }

  // ★ TODOフィルタが選択されている場合、日付ありのタスクは除外
  if (selectedPeriod === "todo") {
    return false;
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