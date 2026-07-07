<<<<<<< HEAD
/* ============================================================
 * kanban.js — Excel Kanban（新UI版）
 * ------------------------------------------------------------
 * Excel連携ロジック（列定義・ステータス判定・DnD更新・
 * 備考編集・スター・フィルタ保存）は旧版を踏襲。
 * UI描画をチップフィルタ／セグメント／色レールカードに刷新し、
 * 検索（タスク名・備考・分類）と共通スライドメニューを追加。
 *
 * レイアウトのペイン追従はCSS(flex)に一本化したため、
 * 旧版のJSによるレーン幅・高さ計算処理は廃止。
 * ============================================================ */
=======
const APP_VERSION = "rev_20260707_7d5a7bd";
>>>>>>> 2156db890a14e886ff9b9d6916f69333d42df12b

const APP_VERSION = "rev_20260707_7d5a7bd";
window.APP_VERSION = APP_VERSION;

let allTasks = [];
let currentDraggedId = null;
let currentTask = null;

let selectedUser = null;
let selectedCategory = null;
let selectedSubCategory = null;
let selectedPeriod = "all";
<<<<<<< HEAD
let showHeld = true;
let searchQuery = "";
=======
let showHeld = true; // 保留表示フラグ
let officeReady = false;
let isExcelHost = false;
>>>>>>> 2156db890a14e886ff9b9d6916f69333d42df12b

<<<<<<< HEAD
Office.onReady(() => {
=======
// レーン高さ調整のデバウンス用変数
let heightAdjustTimeout = null;
let isAdjustingHeights = false;

Office.onReady((info) => {
  officeReady = true;
  isExcelHost = info && info.host === Office.HostType.Excel;

  if (!isExcelHost) {
    console.log("Excel以外のホスト、またはOffice外の実行環境のため初期化をスキップします。");
    return;
  }

  // 保存されたサイズを復元
  restoreSavedSize();
  
  // 保存されたフィルター設定を復元
>>>>>>> 2156db890a14e886ff9b9d6916f69333d42df12b
  restoreSavedFilters();
  restoreHeldDisplay();
  bindStaticUI();
  init();
});

<<<<<<< HEAD
/* ============================================================
   初期化
   ============================================================ */
async function init() {
  await loadExcelData();
  renderFilters();
  renderPeriodSegment();
  renderBoard();
=======
function canUseExcelApi() {
  return officeReady && isExcelHost && typeof Excel !== "undefined";
}

async function reloadKanban() {
  if (!canUseExcelApi()) {
    console.log("Excel APIが利用できないため再読み込みをスキップしました。");
    return;
  }

  await init();

  if (typeof updateVersionDisplay === "function") {
    updateVersionDisplay();
  }
}

// ===== サイズ記憶機能 =====
function restoreSavedSize() {
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    if (savedSize) {
      const size = JSON.parse(savedSize);
      
// 最小サイズの制限（デスクトップ版対応で極小に設定）
  const minWidth = 120;
  const minHeight = 300;
  const width = Math.max(size.width || 120, minWidth);
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
>>>>>>> 2156db890a14e886ff9b9d6916f69333d42df12b

  const v = document.getElementById("version-label");
  if (v) v.textContent = APP_VERSION;
}

/* 静的UIのイベント（初回のみ） */
function bindStaticUI() {
  // 検索
  const input = document.getElementById("search-input");
  const clearBtn = document.getElementById("search-clear");

  input.addEventListener("input", () => {
    searchQuery = input.value.trim();
    renderBoard();
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      input.value = "";
      searchQuery = "";
      renderBoard();
    }
  });
  clearBtn.addEventListener("click", () => {
    input.value = "";
    searchQuery = "";
    renderBoard();
    input.focus();
  });

  // 期間セグメント
  document.querySelectorAll("#seg-period button").forEach(b => {
    b.addEventListener("click", () => setPeriod(b.dataset.p));
  });

  // ドロップダウンの外側クリックで閉じる
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".chip") && !e.target.closest(".dropdown")) {
      closeAllDropdowns();
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeAllDropdowns();
  });

  // 保留トグルの初期表示
  document.getElementById("held-toggle").classList.toggle("on", showHeld);
}

/* ============================================================
   設定の保存・復元（localStorage）
   ============================================================ */
function restoreSavedFilters() {
  try {
    const saved = localStorage.getItem("kanban-filters");
    if (saved) {
      const f = JSON.parse(saved);
      selectedUser = f.user || null;
      selectedCategory = f.category || null;
      selectedSubCategory = f.subCategory || null;
      selectedPeriod = f.period || "all";
    }
  } catch (e) {
    selectedUser = null;
    selectedCategory = null;
    selectedSubCategory = null;
    selectedPeriod = "all";
  }
}

function saveFilters() {
  try {
    localStorage.setItem("kanban-filters", JSON.stringify({
      user: selectedUser,
      category: selectedCategory,
      subCategory: selectedSubCategory,
      period: selectedPeriod,
      timestamp: Date.now()
    }));
  } catch (e) { /* noop */ }
}

<<<<<<< HEAD
function restoreHeldDisplay() {
  const saved = localStorage.getItem("kanban-show-held");
  showHeld = saved !== null ? saved === "true" : true;
}
=======
async function init() {
  if (!canUseExcelApi()) {
    console.log("Excel APIが利用可能になるまで待機中です。");
    return;
  }

  // 保存されたサイズまたはデフォルトサイズを適用
  try {
    const savedSize = localStorage.getItem('kanban-taskpane-size');
    let minWidth = 200;
    
    if (savedSize) {
      const size = JSON.parse(savedSize);
      minWidth = Math.max(size.width || 200, 200);
    }
    
    document.documentElement.style.minWidth = minWidth + "px";
    document.body.style.minWidth = minWidth + "px";
  } catch (e) {
    // エラー時はデフォルトサイズ
    document.documentElement.style.minWidth = "200px";
    document.body.style.minWidth = "200px";
  }
  
  // ボードコンテナをペイン幅に完全追従させる
  const boardEl = document.getElementById("board");
  if (boardEl) {
    boardEl.style.width = "100%"; // ペイン幅に完全追従
    boardEl.style.maxWidth = "100%";
    boardEl.style.minWidth = "200px";
    boardEl.style.boxSizing = "border-box";
  }
>>>>>>> 2156db890a14e886ff9b9d6916f69333d42df12b

function resetSettings() {
  try {
    localStorage.removeItem("kanban-filters");
    localStorage.removeItem("kanban-show-held");
    localStorage.removeItem("kanban-taskpane-size"); // 旧版の残骸も掃除
    window.location.reload();
  } catch (e) { /* noop */ }
}

/* ============================================================
   Excel日付変換
   ============================================================ */
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
  return `${d.getMonth() + 1}/${d.getDate()}`;
}

/* ============================================================
   データ取得（列定義は旧版と同一）
   A:分類 B:小分類 N:担当 O:備考 P:開始 Q:終了
   R:実績開始 S:実績終了 T:除外("-") Y:ID Z:タイトル
   ============================================================ */
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
        classification: row[1],
        title: row[25],
        user: row[13],
        start: row[15],
        end: row[16],
        actualStart: row[17],
        actualEnd: row[18],
        note: row[14],
        rowIndex: i + 11,

        isNoSchedule: !row[15] && !row[16],
        isStar: row[14] && row[14].toString().startsWith("★")
      };

      t.status = getStatus(t);
      return t;
    }).filter(x => x);
  });
}

/* ============================================================
   ステータス
   ============================================================ */
function getStatus(t) {
  if (t.actualEnd) return "完了";
  if (t.actualStart) return "対応中";
  return "未着手";
}

/* ============================================================
   フィルタUI（チップ＋ドロップダウン）
   ============================================================ */
function renderFilters() {
  renderUserDropdown();
  renderCategoryDropdown();
  renderSubCategoryDropdown();
  updateChips();
}

function toggleDropdown(id, chip) {
  const dd = document.getElementById(id);
  const wasOpen = dd.classList.contains("open");
  closeAllDropdowns();
  if (!wasOpen) {
    dd.classList.add("open");
    // チップの真下に配置
    const rect = chip.getBoundingClientRect();
    const barRect = chip.closest(".filter-bar").getBoundingClientRect();
    let left = rect.left - barRect.left;
    dd.style.left = left + "px";
    // 右端はみ出し補正
    requestAnimationFrame(() => {
      const ddRect = dd.getBoundingClientRect();
      const over = ddRect.right - (barRect.right - 4);
      if (over > 0) dd.style.left = Math.max(4, left - over) + "px";
    });
  }
}

function closeAllDropdowns() {
  document.querySelectorAll(".dropdown").forEach(d => d.classList.remove("open"));
}

/* チップの表示テキストを選択状態に同期 */
function updateChips() {
  const userChip = document.getElementById("chip-user");
  if (selectedUser) {
    userChip.classList.add("selected");
    userChip.innerHTML =
      `担当: ${escapeHtml(selectedUser)} <span class="clear" onclick="clearUserFilter(event)">✕</span>`;
  } else {
    userChip.classList.remove("selected");
    userChip.innerHTML = `担当者 <span class="caret"></span>`;
  }

  const catChip = document.getElementById("chip-cat");
  if (selectedCategory) {
    const label = selectedSubCategory
      ? `${selectedCategory} / ${selectedSubCategory}`
      : selectedCategory;
    catChip.classList.add("selected");
    catChip.innerHTML =
      `分類: ${escapeHtml(label)} <span class="clear" onclick="clearCategoryFilter(event)">✕</span>`;
  } else {
    catChip.classList.remove("selected");
    catChip.innerHTML = `分類 <span class="caret"></span>`;
  }
}

function clearUserFilter(e) {
  e.stopPropagation();
  selectedUser = null;
  saveFilters();
  renderFilters();
  renderBoard();
}

function clearCategoryFilter(e) {
  e.stopPropagation();
  selectedCategory = null;
  selectedSubCategory = null;
  saveFilters();
  renderFilters();
  renderBoard();
}

function renderUserDropdown() {
  const users = [...new Set(
    allTasks.map(t => t.user).filter(v => v && v !== "#")
  )];

  const el = document.getElementById("user-filters");
  el.innerHTML = "";

  users.forEach(u => {
    const b = document.createElement("button");
    b.className = "dd-item" + (selectedUser === u ? " on" : "");

    const av = document.createElement("span");
    av.className = "avatar";
    av.style.background = userColor(u);
    av.textContent = String(u).charAt(0);
    b.appendChild(av);
    b.appendChild(document.createTextNode(u));

    b.onclick = () => {
      selectedUser = (selectedUser === u) ? null : u;
      saveFilters();
      renderFilters();
      renderBoard();
      closeAllDropdowns();
    };

    el.appendChild(b);
  });
}

function renderCategoryDropdown() {
  const cats = [...new Set(
    allTasks.map(t => t.category).filter(v => v && v !== "#")
  )];

  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    const b = document.createElement("button");
    b.className = "dd-item" + (selectedCategory === c ? " on" : "");
    b.textContent = c;

    b.onclick = () => {
      selectedCategory = (selectedCategory === c) ? null : c;
      selectedSubCategory = null; // 大分類変更時は小分類をリセット
      saveFilters();
      renderFilters();
      renderBoard();
      // 小分類がある場合はドロップダウンを開いたままにする
      const hasSub = selectedCategory &&
        allTasks.some(t => t.category === selectedCategory &&
          t.classification && String(t.classification).trim() !== "" &&
          t.classification !== "#");
      if (!hasSub) closeAllDropdowns();
    };

    el.appendChild(b);
  });
}

function renderSubCategoryDropdown() {
  const section = document.getElementById("sub-category-section");
  const el = document.getElementById("sub-category-filters");
  if (!section || !el) return;

  if (!selectedCategory) {
    section.style.display = "none";
    return;
  }

  const subCats = [...new Set(
    allTasks
      .filter(t => t.category === selectedCategory)
      .map(t => t.classification)
      .filter(v => v && v !== "#" && v.toString().trim() !== "")
  )];

  if (subCats.length === 0) {
    section.style.display = "none";
    return;
  }

  section.style.display = "block";
  el.innerHTML = "";

  subCats.forEach(s => {
    const b = document.createElement("button");
    b.className = "dd-item" + (selectedSubCategory === s ? " on" : "");
    b.textContent = s;

    b.onclick = () => {
      selectedSubCategory = (selectedSubCategory === s) ? null : s;
      saveFilters();
      renderFilters();
      renderBoard();
      closeAllDropdowns();
    };

    el.appendChild(b);
  });
}

/* 担当者名から一意な色を生成 */
function userColor(name) {
  let h = 0;
  const s = String(name);
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) % 360;
  }
  return `hsl(${h}, 48%, 48%)`;
}

/* ============================================================
   期間フィルタ（セグメント）
   ============================================================ */
function setPeriod(p) {
  selectedPeriod = (selectedPeriod === p) ? "all" : p;
  saveFilters();
  renderPeriodSegment();
  renderBoard();
}

function renderPeriodSegment() {
  document.querySelectorAll("#seg-period button").forEach(b => {
    b.classList.toggle("active", b.dataset.p === selectedPeriod);
  });
}

/* ============================================================
   保留表示切替
   ============================================================ */
function toggleHeldDisplay(e) {
  if (e) e.preventDefault();
  showHeld = !showHeld;
  localStorage.setItem("kanban-show-held", showHeld);
  document.getElementById("held-toggle").classList.toggle("on", showHeld);
  renderBoard();
}

/* ============================================================
   描画
   ============================================================ */
function renderBoard() {
  ["todo", "held", "doing", "done"].forEach(l => {
    const lane = document.querySelector(`#${l} .card-list`);
    if (lane) lane.innerHTML = "";
  });

  // 保留レーンの表示/非表示
  const heldLane = document.getElementById("held");
  if (heldLane) heldLane.style.display = showHeld ? "" : "none";

  const filtered = allTasks.filter(isMatch);

  const normal = filtered
    .filter(t => t.status !== "完了")
    .sort((a, b) => {
      if (a.isStar && !b.isStar) return -1;
      if (!a.isStar && b.isStar) return 1;
      return excelDateToJS(a.end) - excelDateToJS(b.end);
    });

  const done = filtered
    .filter(t => t.status === "完了")
    .sort((a, b) => excelDateToJS(b.actualEnd) - excelDateToJS(a.actualEnd));

  [...normal, ...done].forEach(t => {
    const lane = getLane(t);
    document.querySelector(`#${lane} .card-list`).appendChild(createCard(t));
  });

  // 空レーン表示と件数バッジ
  ["todo", "held", "doing", "done"].forEach(l => {
    const laneEl = document.getElementById(l);
    const list = laneEl.querySelector(".card-list");
    const n = list.children.length;
    laneEl.querySelector(".count").textContent = n;
    if (n === 0) {
      const em = document.createElement("div");
      em.className = "empty";
      em.textContent = "なし";
      list.appendChild(em);
    }
  });

  // 検索ヒット件数
  const box = document.getElementById("search-box");
  const hits = document.getElementById("search-hits");
  box.classList.toggle("has-value", searchQuery.length > 0);
  hits.textContent = searchQuery ? `${filtered.length}件` : "";

  setupDnD();
}

/* ============================================================
   カード生成
   ============================================================ */
function createCard(t) {
  const d = document.createElement("div");
  d.className = "card";
  d.draggable = true;

  // 色レール（旧applyColorの枠線色に相当）
  const lane = getLane(t);
  if (t.status === "完了") {
    d.classList.add("is-done");
  } else if (lane === "held") {
    d.classList.add("is-held");
  } else {
    const startRaw = excelDateToJS(t.start);
    const endRaw = excelDateToJS(t.end);
    if (startRaw && endRaw) {
      const start = new Date(startRaw); start.setHours(0, 0, 0, 0);
      const end = new Date(endRaw);     end.setHours(0, 0, 0, 0);
      const today = new Date();         today.setHours(0, 0, 0, 0);
      if (end < today) d.classList.add("is-delay");
      else if (start <= today && end >= today) d.classList.add("is-active");
    }
  }
  if (t.isStar) d.classList.add("starred");

  // DnD
  d.addEventListener("dragstart", (e) => {
    currentDraggedId = t.id;
    e.dataTransfer.setData("text/plain", t.id);
    d.classList.add("dragging");
  });
  d.addEventListener("dragend", () => d.classList.remove("dragging"));

  // 左クリック：Excelへジャンプ
  d.addEventListener("click", (e) => {
    if (e.button !== 0) return;
    jumpToExcel(t.rowIndex);
  });

  // 右クリック：備考編集
  d.addEventListener("contextmenu", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    await openModal(t);
  });

  /* --- 1行目：日付＋担当＋スター --- */
  const meta = document.createElement("div");
  meta.className = "card-meta";

  const dates = document.createElement("span");
  dates.className = "card-dates";

  if (t.isNoSchedule) {
    const badge = document.createElement("span");
    badge.className = "badge-todo";
    badge.textContent = "TODO";
    meta.appendChild(badge);
  } else {
    dates.innerHTML =
      `${fmt(t.start)} <span class="arrow">→</span> ${fmt(t.end)}`;
    if (d.classList.contains("is-delay")) dates.classList.add("delay");
  }
  meta.appendChild(dates);

  if (t.user) {
    const av = document.createElement("span");
    av.className = "card-user";
    av.style.background = userColor(t.user);
    av.textContent = String(t.user).charAt(0);
    av.title = t.user;
    meta.appendChild(av);
  }

  // 完了以外にスターを表示
  if (t.status !== "完了") {
    const star = document.createElement("button");
    star.className = "card-star" + (t.isStar ? " on" : "");
    star.textContent = t.isStar ? "★" : "☆";
    star.title = "本日の優先タスク";
    star.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleStar(t);
    });
    meta.appendChild(star);
  }

  d.appendChild(meta);

  /* --- 2行目：タイトル＋分類 --- */
  const row2 = document.createElement("div");
  row2.className = "card-title-row";

  const titleSpan = document.createElement("span");
  titleSpan.className = "card-title";
  titleSpan.innerHTML = highlight(String(t.title), searchQuery);

  row2.appendChild(titleSpan);

  if (t.classification && String(t.classification).trim() !== "") {
    const cls = document.createElement("span");
    cls.className = "card-cls";
    cls.innerHTML = highlight(String(t.classification), searchQuery);
    row2.appendChild(cls);
  }
  d.appendChild(row2);

  /* --- 実績日 --- */
  if (t.status === "対応中") {
    const ac = document.createElement("div");
    ac.className = "card-actual";
    ac.textContent = `実績 ${fmt(t.actualStart)} 〜`;
    d.appendChild(ac);
  } else if (t.status === "完了") {
    const ac = document.createElement("div");
    ac.className = "card-actual";
    ac.textContent = `実績 ${fmt(t.actualStart)} 〜 ${fmt(t.actualEnd)}`;
    d.appendChild(ac);
  }

  /* --- 備考プレビュー（検索が備考にヒットした時のみ） --- */
  const note = (t.note || "").toString();
  if (searchQuery && note.toLowerCase().includes(searchQuery.toLowerCase())) {
    const np = document.createElement("div");
    np.className = "card-note-hit";
    np.innerHTML = "📝 " + highlight(note, searchQuery);
    d.appendChild(np);
    d.classList.add("show-note");
  }

  return d;
}

/* ============================================================
   検索ハイライト
   ============================================================ */
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function highlight(text, q) {
  if (!q) return escapeHtml(text);
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return escapeHtml(text).replace(
    new RegExp(`(${esc})`, "gi"),
    '<mark class="hit">$1</mark>'
  );
}

/* ============================================================
   DnD
   ============================================================ */
function setupDnD() {
  ["todo", "held", "doing", "done"].forEach(id => {
    const lane = document.getElementById(id);
    const list = lane.querySelector(".card-list");

    lane.ondragover = (e) => {
      e.preventDefault();
      list.classList.add("drop-target");
    };
    lane.ondragleave = () => list.classList.remove("drop-target");
    lane.ondrop = (e) => {
      e.preventDefault();
      list.classList.remove("drop-target");
      const t = allTasks.find(x => x.id === currentDraggedId);
      if (t) updateStatus(t, id);
    };
  });
}

/* ============================================================
   Excel操作
   ============================================================ */
async function jumpToExcel(row) {
  await Excel.run(async (ctx) => {
    const s = ctx.workbook.worksheets.getItem("wbs");
    s.activate();
    s.getRange(`${row}:${row}`).select();
    await ctx.sync();
  });
}

/* ============================================================
   util
   ============================================================ */
function getLane(task) {
  // 備考欄に▲がある場合は保留レーン
  if (task.note && task.note.toString().includes("▲")) {
    return "held";
  }
  const s = task.status;
  if (s === "未着手") return "todo";
  if (s === "保留") return "held";
  if (s === "対応中") return "doing";
  return "done";
}

function getMonday(d) {
  const t = new Date(d);
  const day = t.getDay();
  const diff = t.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(t.setDate(diff));
}

function addDays(d, n) {
  const t = new Date(d);
  t.setDate(t.getDate() + n);
  return t;
}

function dateToExcelSerial(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return "";
  const excelEpoch = new Date(1900, 0, 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  const daysDiff = Math.floor((date - excelEpoch) / msPerDay);
  return daysDiff + (date >= new Date(1900, 2, 1) ? 2 : 1);
}

function isValidDate(v) {
  return v instanceof Date && !isNaN(v);
}

/* ============================================================
   ステータス更新（DnD時、旧版と同一ロジック）
   ============================================================ */
async function updateStatus(task, lane) {
  let actualStart = task.actualStart;
  let actualEnd = task.actualEnd;

  if (lane === "todo") {
    actualStart = "";
    actualEnd = "";

    if (task.note && task.note.toString().includes("▲")) {
      task.note = task.note.toString().replace(/▲/g, "△");
      await updateTaskStatus(task, "未着手");
    }
  }

  if (lane === "held") {
    // 完了から保留に移動した場合のみ実績完了日をクリア
    if (task.status === "完了") {
      actualEnd = "";
    }

    let newNote = ensureStatusSymbols((task.note || "").toString());
    if (newNote.includes("△")) {
      newNote = newNote.replace(/△/g, "▲");
    } else if (!newNote.includes("▲")) {
      const lines = newNote.split("\n");
      lines[0] = lines[0].replace(/△/, "") + "▲";
      newNote = lines.join("\n");
    }
    task.note = newNote;

    await updateTaskStatus(task, "保留");
  }

  if (lane === "doing") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = "";

    if (task.note && task.note.toString().includes("▲")) {
      task.note = task.note.toString().replace(/▲/g, "△");
      await updateTaskStatus(task, "対応中");
    }
  }

  if (lane === "done") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = new Date();

    if (task.isStar) task.isStar = false;

    if (task.note && task.note.toString().includes("▲")) {
      task.note = task.note.toString().replace(/▲/g, "△");
      await updateTaskStatus(task, "完了");
    }
  }

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const row = task.rowIndex;

    const startCell = sheet.getRange(`R${row}`);
    const endCell = sheet.getRange(`S${row}`);

    startCell.values = [[dateToExcelSerial(actualStart)]];
    endCell.values = [[dateToExcelSerial(actualEnd)]];

    startCell.numberFormat = [["m/d"]];
    endCell.numberFormat = [["m/d"]];

    if ((lane === "done" || lane === "doing" || lane === "held") && task.note !== undefined) {
      const noteCell = sheet.getRange(`O${row}`);
      noteCell.values = [[task.note]];
      noteCell.format.wrapText = false;
    }

    // 完了時に備考から★を削除
    if (lane === "done" && task.note && task.note.toString().includes("★")) {
      const newNote = task.note.toString().replace(/★/g, "");
      const noteCell = sheet.getRange(`O${row}`);
      noteCell.values = [[newNote]];
      noteCell.format.wrapText = false;
      task.note = newNote;
    }

    await ctx.sync();
  });

  await init();
}

/* ステータス文字列（H列）更新 */
async function updateTaskStatus(task, newStatus) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const statusCell = sheet.getRange(`H${task.rowIndex}`);
    statusCell.values = [[newStatus]];
    await ctx.sync();
  });
  task.status = newStatus;
}

/* ============================================================
   スター切り替え
   ============================================================ */
async function toggleStar(task) {
  task.isStar = !task.isStar;

  let newNote = (task.note || "").toString();
  if (task.isStar) {
    if (!newNote.startsWith("★")) newNote = "★" + newNote;
  } else {
    newNote = newNote.replace(/★/g, "");
  }

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const cell = sheet.getRange(`O${task.rowIndex}`);
    cell.values = [[newNote]];
    cell.format.wrapText = false;
    await ctx.sync();
  });

  task.note = newNote;
  renderBoard();
}

/* ============================================================
   ステータス記号管理
   ============================================================ */
function ensureStatusSymbols(noteText) {
  if (!noteText) noteText = "";
  const lines = noteText.split("\n");
  let firstLine = lines[0] || "";

  if (!firstLine.includes("★") && !firstLine.includes("☆")) {
    firstLine = "☆" + firstLine;
  }
  if (!firstLine.includes("▲") && !firstLine.includes("△")) {
    firstLine = firstLine + "△";
  }

  lines[0] = firstLine;
  return lines.join("\n");
}

/* ============================================================
   備考編集モーダル（旧版と同一ロジック）
   ============================================================ */
async function openModal(task) {
  currentTask = task;

  // O列から最新の備考内容を取得
  let originalNote = "";
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem("wbs");
      const noteCell = sheet.getRange(`O${task.rowIndex}`);
      noteCell.load("values");
      await ctx.sync();
      originalNote = (noteCell.values[0][0] || "").toString();
    });
  } catch (error) {
    originalNote = (task.note || "").toString();
  }

  let displayNote = originalNote;

  if (!displayNote.trim()) {
    displayNote = "☆△\n＜タスク＞\n＜状況＞";
  } else {
    displayNote = ensureStatusSymbols(displayNote);
    const lines = displayNote.split("\n");
    if (lines.length < 2 || (lines.length === 2 && !lines[1].trim())) {
      displayNote = displayNote.trimEnd() + "\n＜タスク＞\n＜状況＞";
    }
  }

  document.getElementById("modal-title").textContent = task.title;
  document.getElementById("modal-note").value = displayNote;

  const modal = document.getElementById("modal");
  modal.classList.remove("hidden");

  const handleEscKey = (event) => {
    if (event.key === "Escape") closeModal();
  };
  const handleOverlayClick = (event) => {
    if (event.target === modal) {
      const currentNote = document.getElementById("modal-note").value;
      if (currentNote === displayNote) closeModal();
    }
  };
  const modalContent = modal.querySelector(".modal-content");
  const handleContentClick = (event) => event.stopPropagation();

  document.addEventListener("keydown", handleEscKey);
  modal.addEventListener("click", handleOverlayClick);
  modalContent.addEventListener("click", handleContentClick);

  modal._cleanup = () => {
    document.removeEventListener("keydown", handleEscKey);
    modal.removeEventListener("click", handleOverlayClick);
    modalContent.removeEventListener("click", handleContentClick);
  };

  setTimeout(() => document.getElementById("modal-note").focus(), 100);
}

function closeModal() {
  const modal = document.getElementById("modal");
  modal.classList.add("hidden");
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
    cell.format.wrapText = false;

    const entireRow = sheet.getRange(`${row}:${row}`);
    entireRow.format.rowHeight = 20;

    await ctx.sync();
  });

  if (currentTask) {
    currentTask.note = note;
    currentTask.isStar = note.startsWith("★");
  }

  closeModal();
  renderBoard();
}

/* ============================================================
   フィルタ判定（検索を追加）
   ============================================================ */
function isMatch(t) {

  // ★ 検索（タスク名・備考・大分類・小分類を横断）
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    const hit =
      (t.title || "").toString().toLowerCase().includes(q) ||
      (t.note || "").toString().toLowerCase().includes(q) ||
      (t.category || "").toString().toLowerCase().includes(q) ||
      (t.classification || "").toString().toLowerCase().includes(q);
    if (!hit) return false;
  }

  // 担当者
  if (selectedUser && t.user !== selectedUser) return false;

  // 分類（大分類）
  if (selectedCategory && t.category !== selectedCategory) return false;

  // 小分類
  if (selectedSubCategory && t.classification !== selectedSubCategory) return false;

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
  today.setHours(0, 0, 0, 0);

  const monday = getMonday(today);
  const sunday = addDays(monday, 6);
  const nextMonday = addDays(monday, 7);
  const nextSunday = addDays(monday, 13);

  switch (selectedPeriod) {
    case "past":     return end < monday;
    case "week":     return (start <= sunday && end >= monday);
    case "nextweek": return (start <= nextSunday && end >= nextMonday);
    case "future":   return start > nextSunday;
    case "all":
    default:         return true;
  }
}

/* ============================================================
   共通スライドメニュー（遅延ロード）
   ------------------------------------------------------------
   メニュー項目（名前・URL）は tools/common/menu.json で
   一元管理。menu.json を編集すれば全アプリに反映される。
   ============================================================ */
const COMMON_BASE = "https://ymatsuda-cmyk.github.io/tools/common";

let menuReady = null;

function openMenu(btn) {
  if (!menuReady) {
    // 初回クリック時にだけ slide-menu.js を読み込む
    if (btn) btn.disabled = true;
    menuReady = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = COMMON_BASE + "/slide-menu.js";
      s.onload = () => {
        SlideMenu.init({
          appName: "Excel Kanban",
          version: APP_VERSION,
          position: "right",
          width: 250,
          theme: { accent: "#0E7A5F" },
          footer: "© RightArm",
          currentId: "kanban",                       // menu.json のidと一致で強調表示
          menuUrl: COMMON_BASE + "/menu.json",       // ★ メニュー定義はJSONで一元管理
          localItems: [                              // このアプリ固有の操作
            { section: "操作" },
            { label: "再読み込み", icon: "🔄", onClick: () => init() },
            { label: "設定をリセット", icon: "🧹", onClick: () => resetSettings() }
          ]
        });
        resolve();
      };
      s.onerror = () => {
        menuReady = null; // 失敗時は次回リトライ可能に
        reject(new Error("slide-menu.js load failed"));
      };
      document.head.appendChild(s);
    });
  }

  menuReady
    .then(() => {
      if (btn) btn.disabled = false;
      SlideMenu.open();
    })
    .catch(() => {
      if (btn) btn.disabled = false;
      console.warn("メニューを読み込めませんでした");
    });
}
