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

const APP_VERSION = "rev_20260720_b46ccb1";
window.APP_VERSION = APP_VERSION;

let allTasks = [];
let currentDraggedId = null;
let currentTask = null;

let selectedUser = null;
let selectedCategory = null;
let selectedSubCategory = null;
let selectedPeriod = "all";
let showHeld = true;
let searchQuery = "";

if (window.Office && Office.onReady) {
  Office.onReady(() => {
    restoreSavedFilters();
    restoreHeldDisplay();
    bindStaticUI();
    init();
  });
} else {
  // ブラウザ直接表示（開発確認用）: Excel連携なしでUIのみ初期化
  window.addEventListener("DOMContentLoaded", () => {
    restoreSavedFilters();
    restoreHeldDisplay();
    bindStaticUI();
    const v = document.getElementById("version-label");
    if (v) v.textContent = APP_VERSION + " (no-office)";
  });
}

/* ============================================================
   初期化
   ============================================================ */
async function init() {
  await loadExcelData();
  renderFilters();
  renderPeriodSegment();
  renderBoard();

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
      selectedUsers = Array.isArray(f.users) ? f.users : (f.user ? [f.user] : []);
      selectedCategories = Array.isArray(f.categories) ? f.categories : (f.category ? [f.category] : []);
      selectedSubCategories = Array.isArray(f.subCategories) ? f.subCategories : (f.subCategory ? [f.subCategory] : []);
      selectedPeriod = f.period || "all";
    }
  } catch (e) {
    selectedUsers = [];
    selectedCategories = [];
    selectedSubCategories = [];
    selectedPeriod = "all";
  }
}

function saveFilters() {
  try {
    localStorage.setItem("kanban-filters", JSON.stringify({
      users: selectedUsers,
      categories: selectedCategories,
      subCategories: selectedSubCategories,
      period: selectedPeriod,
      timestamp: Date.now()
    }));
  } catch (e) { /* noop */ }
}

function restoreHeldDisplay() {
  const saved = localStorage.getItem("kanban-show-held");
  showHeld = saved !== null ? saved === "true" : true;
}

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
  if (selectedUsers.length) {
    userChip.classList.add("selected");
    const label = selectedUsers.length === 1 ? escapeHtml(selectedUsers[0]) : `${selectedUsers.length}件選択`;
    userChip.innerHTML =
      `担当: ${label} <span class="clear" onclick="clearUserFilter(event)">✕</span>`;
  } else {
    userChip.classList.remove("selected");
    userChip.innerHTML = `担当者 <span class="caret"></span>`;
  }

  const catChip = document.getElementById("chip-cat");
  if (selectedCategories.length) {
    catChip.classList.add("selected");
    const label = selectedCategories.length === 1 ? escapeHtml(selectedCategories[0]) : `${selectedCategories.length}件選択`;
    catChip.innerHTML =
      `分類: ${label} <span class="clear" onclick="clearCategoryFilter(event)">✕</span>`;
  } else {
    catChip.classList.remove("selected");
    catChip.innerHTML = `分類 <span class="caret"></span>`;
  }

  const subCatChip = document.getElementById("chip-subcat");
  if (subCatChip) {
    if (selectedSubCategories.length) {
      subCatChip.classList.add("selected");
      const label = selectedSubCategories.length === 1 ? escapeHtml(selectedSubCategories[0]) : `${selectedSubCategories.length}件選択`;
      subCatChip.innerHTML =
        `小分類: ${label} <span class="clear" onclick="clearSubCategoryFilter(event)">✕</span>`;
    } else {
      subCatChip.classList.remove("selected");
      subCatChip.innerHTML = `小分類 <span class="caret"></span>`;
    }
  }
}

function clearUserFilter(e) {
  e.stopPropagation();
  selectedUsers = [];
  saveFilters();
  renderFilters();
  renderBoard();
}

function clearCategoryFilter(e) {
  e.stopPropagation();
  selectedCategories = [];
  selectedSubCategories = [];
  saveFilters();
  renderFilters();
  renderBoard();
}

function clearSubCategoryFilter(e) {
  e.stopPropagation();
  selectedSubCategories = [];
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
    const b = document.createElement("label");
    b.className = "dd-item" + (selectedUsers.includes(u) ? " on" : "");

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedUsers.includes(u);
    cb.className = "dd-check";

    const av = document.createElement("span");
    av.className = "avatar";
    av.style.background = userColor(u);
    av.textContent = String(u).charAt(0);

    b.appendChild(cb);
    b.appendChild(av);
    b.appendChild(document.createTextNode(u));

    cb.addEventListener("change", () => {
      selectedUsers = cb.checked
        ? [...selectedUsers, u]
        : selectedUsers.filter(x => x !== u);
      b.classList.toggle("on", cb.checked);
      saveFilters();
      updateChips();
      renderBoard();
    });

    el.appendChild(b);
  });

  if (users.length) {
    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "dd-clear";
    clearBtn.textContent = "選択解除";
    clearBtn.onclick = () => clearUserFilter({ stopPropagation() {} });
    el.appendChild(clearBtn);
  }
}

function renderCategoryDropdown() {
  const cats = [...new Set(
    allTasks.map(t => t.category).filter(v => v && v !== "#")
  )];

  const el = document.getElementById("category-filters");
  el.innerHTML = "";

  cats.forEach(c => {
    const b = document.createElement("label");
    b.className = "dd-item" + (selectedCategories.includes(c) ? " on" : "");

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedCategories.includes(c);
    cb.className = "dd-check";

    b.appendChild(cb);
    b.appendChild(document.createTextNode(c));

    cb.addEventListener("change", () => {
      selectedCategories = cb.checked
        ? [...selectedCategories, c]
        : selectedCategories.filter(x => x !== c);
      b.classList.toggle("on", cb.checked);

      // 選択解除された大分類配下の小分類はフィルタから外す
      selectedSubCategories = selectedSubCategories.filter(s =>
        allTasks.some(t =>
          t.classification === s &&
          (!selectedCategories.length || selectedCategories.includes(t.category))
        )
      );

      saveFilters();
      renderFilters();
      renderBoard();
    });

    el.appendChild(b);
  });

  if (cats.length) {
    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "dd-clear";
    clearBtn.textContent = "選択解除";
    clearBtn.onclick = () => clearCategoryFilter({ stopPropagation() {} });
    el.appendChild(clearBtn);
  }
}

function renderSubCategoryDropdown() {
  const el = document.getElementById("sub-category-filters");
  if (!el) return;

  const subCats = [...new Set(
    allTasks
      .filter(t => !selectedCategories.length || selectedCategories.includes(t.category))
      .map(t => t.classification)
      .filter(v => v && v !== "#" && v.toString().trim() !== "")
  )];

  el.innerHTML = "";

  if (subCats.length === 0) {
    const empty = document.createElement("div");
    empty.className = "dd-item";
    empty.textContent = "小分類なし";
    empty.style.opacity = "0.6";
    empty.style.cursor = "default";
    el.appendChild(empty);
    return;
  }

  subCats.forEach(s => {
    const b = document.createElement("label");
    b.className = "dd-item" + (selectedSubCategories.includes(s) ? " on" : "");

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedSubCategories.includes(s);
    cb.className = "dd-check";

    b.appendChild(cb);
    b.appendChild(document.createTextNode(s));

    cb.addEventListener("change", () => {
      selectedSubCategories = cb.checked
        ? [...selectedSubCategories, s]
        : selectedSubCategories.filter(x => x !== s);
      b.classList.toggle("on", cb.checked);
      saveFilters();
      updateChips();
      renderBoard();
    });

    el.appendChild(b);
  });

  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "dd-clear";
  clearBtn.textContent = "選択解除";
  clearBtn.onclick = () => clearSubCategoryFilter({ stopPropagation() {} });
  el.appendChild(clearBtn);
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
  renderSubtaskKanban();

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
   サブタスクカンバン（備考モーダル内・案1: 3レーン）
   ------------------------------------------------------------
   備考テキスト内の行頭記号でサブタスクの状態を表す:
     □ 未着手 / ◎ 対応中 / ■ 完了
   カンバンのドラッグ移動・備考テキスト編集を双方向同期する。
   ============================================================ */
const SUB_MARKS = { "□": "todo", "◎": "doing", "■": "done" };
const SUB_LANES = [
  { key: "todo", mark: "□", label: "未着手", cls: "" },
  { key: "doing", mark: "◎", label: "対応中", cls: "doing" },
  { key: "done", mark: "■", label: "完了", cls: "done" },
];

function parseSubtasks(note) {
  const tasks = [];
  (note || "").split(/\r?\n/).forEach((line, idx) => {
    const m = line.match(/^\s*([□◎■])\s?(.*)$/);
    if (m) tasks.push({ lane: SUB_MARKS[m[1]], title: m[2].trim(), line: idx });
  });
  return tasks;
}
function subtasksToNote(note, tasks) {
  const lines = (note || "").split(/\r?\n/);
  const byLine = {};
  tasks.forEach((t) => { byLine[t.line] = t; });
  const kept = [];
  lines.forEach((line, idx) => {
    if (/^\s*[□◎■]/.test(line)) {
      const t = byLine[idx];
      if (t) kept.push(`${subMark(t.lane)} ${t.title}`);
    } else {
      kept.push(line);
    }
  });
  return kept.join("\n");
}
function subMark(lane) { return lane === "doing" ? "◎" : lane === "done" ? "■" : "□"; }

function renderSubtaskKanban() {
  const host = document.getElementById("subtask-kanban");
  if (!host) return;
  const note = document.getElementById("modal-note").value;
  const tasks = parseSubtasks(note);
  const lanesHtml = SUB_LANES.map((L) => {
    const cards = tasks.filter((t) => t.lane === L.key);
    return `
      <div class="sk-lane ${L.cls}" data-lane="${L.key}">
        <div class="sk-lane-head">${L.mark} ${L.label} <span class="sk-cnt">${cards.length}</span></div>
        <div class="sk-lane-body" data-lane="${L.key}">
          ${cards.map((c) => `<div class="sk-card ${L.cls}" draggable="true" data-line="${c.line}">${escapeHtml(c.title || "（無題）")}</div>`).join("")}
        </div>
      </div>`;
  }).join("");
  host.innerHTML = `
    <div class="sk-board">${lanesHtml}</div>
    <div class="sk-add">
      <input type="text" id="sk-new" placeholder="サブタスク名を入力してEnterまたは＋"
        onkeydown="if(event.key==='Enter'){event.preventDefault();addSubtask()}">
      <button type="button" onclick="addSubtask()">＋追加</button>
    </div>`;
  setupSubtaskDnd();
}

function setupSubtaskDnd() {
  const host = document.getElementById("subtask-kanban");
  let dragLine = null;
  host.querySelectorAll(".sk-card").forEach((card) => {
    card.addEventListener("dragstart", (e) => {
      dragLine = Number(card.dataset.line);
      card.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });
    card.addEventListener("dragend", () => card.classList.remove("dragging"));
  });
  host.querySelectorAll(".sk-lane-body").forEach((body) => {
    body.addEventListener("dragover", (e) => { e.preventDefault(); body.classList.add("over"); });
    body.addEventListener("dragleave", () => body.classList.remove("over"));
    body.addEventListener("drop", (e) => {
      e.preventDefault();
      body.classList.remove("over");
      if (dragLine == null) return;
      moveSubtask(dragLine, body.dataset.lane);
      dragLine = null;
    });
  });
}

/* カード移動 → 備考テキストを書き換え、テキストエリアも即時更新（双方向同期） */
function moveSubtask(line, newLane) {
  const ta = document.getElementById("modal-note");
  const tasks = parseSubtasks(ta.value);
  const t = tasks.find((x) => x.line === line);
  if (!t || t.lane === newLane) return;
  t.lane = newLane;
  ta.value = subtasksToNote(ta.value, tasks);
  renderSubtaskKanban();
}

function addSubtask() {
  const ta = document.getElementById("modal-note");
  const input = document.getElementById("sk-new");
  const title = (input.value || "").trim();
  if (!title) return;
  ta.value = insertIntoTaskSection(ta.value, `□ ${title}`);
  input.value = "";
  renderSubtaskKanban();
}

/* ＜タスク＞セクション内（次のセクション見出しの直前、無ければ末尾）に1行挿入する。
   ＜タスク＞見出し自体が無い場合は、従来どおり末尾へ追記する。 */
function insertIntoTaskSection(note, newLine) {
  const lines = (note || "").split(/\r?\n/);
  const startIdx = lines.findIndex(l => l.trim() === "＜タスク＞");
  if (startIdx === -1) {
    const trimmed = (note || "").replace(/\s+$/, "");
    return (trimmed ? trimmed + "\n" : "") + newLine;
  }
  let endIdx = lines.length;
  for (let i = startIdx + 1; i < lines.length; i++) {
    if (/^＜.*＞$/.test(lines[i].trim())) { endIdx = i; break; }
  }
  lines.splice(endIdx, 0, newLine);
  return lines.join("\n");
}

/* 備考テキストエリアが編集されたらカンバンを再描画（逆方向の同期） */
function onModalNoteEdited() {
  renderSubtaskKanban();
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
  if (selectedUsers.length && !selectedUsers.includes(t.user)) return false;

  // 分類（大分類）
  if (selectedCategories.length && !selectedCategories.includes(t.category)) return false;

  // 小分類
  if (selectedSubCategories.length && !selectedSubCategories.includes(t.classification)) return false;

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

/* ============================================================
   汎用ダイアログ（Office環境では window.confirm/alert 不可）
   ============================================================ */
let dialogResolve = null;
function uiConfirm(message) {
  return new Promise(resolve => {
    dialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "";
    document.getElementById("dialog-modal").classList.remove("hidden");
  });
}
function uiAlert(message) {
  return new Promise(resolve => {
    dialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "none";
    document.getElementById("dialog-modal").classList.remove("hidden");
  });
}
function dialogRespond(ok) {
  document.getElementById("dialog-modal").classList.add("hidden");
  const r = dialogResolve;
  dialogResolve = null;
  if (r) r(ok);
}

/* ============================================================
   タスク追加
   ------------------------------------------------------------
   ・wbsシートの名前定義「タスク範囲」内の選択行に行挿入して追加
   ・T〜FY列の数式は隣接行からコピーして埋める
   ・値の書込み: A=大分類, B=小分類, E=タスク名, N=担当者,
     P=予定開始日, Q=予定終了日
   ・大分類「受注」の場合、小分類は営業報告シートの
     状態=受注/受託中 の案件番号から選択
   ============================================================ */
const EIGYO_SHEET = "営業報告";
const ORDER_CATEGORY = "受注";
const TASK_RANGE_NAME = "タスク範囲";

function openTaskAdd() {
  // 大分類: 既存タスクの大分類 ＋ 受注（無ければ追加）
  const cats = [...new Set(allTasks.map(t => t.category).filter(v => v && v !== "#"))];
  if (!cats.includes(ORDER_CATEGORY)) cats.push(ORDER_CATEGORY);
  const catSel = document.getElementById("ta-cat");
  catSel.innerHTML = cats.map(c => `<option>${escapeHtml(String(c))}</option>`).join("");

  // 担当者: 既存タスクの担当者
  const users = [...new Set(allTasks.map(t => t.user).filter(v => v && v !== "#"))];
  const userSel = document.getElementById("ta-user");
  userSel.innerHTML = `<option value=""></option>` + users.map(u => `<option>${escapeHtml(String(u))}</option>`).join("");

  // 入力初期化
  document.getElementById("ta-subcat").value = "";
  document.getElementById("ta-title").value = "";
  document.getElementById("ta-start").value = "";
  document.getElementById("ta-end").value = "";
  const msg = document.getElementById("ta-msg");
  msg.className = "task-msg"; msg.textContent = "";

  onTaCatChange();

  // wbsシートが表示されていない場合はアクティブにする
  activateWbs();

  document.getElementById("task-modal").classList.remove("hidden");
}
function closeTaskAdd() { document.getElementById("task-modal").classList.add("hidden"); }

async function activateWbs() {
  if (!window.Excel) return;
  try {
    await Excel.run(async ctx => {
      const active = ctx.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await ctx.sync();
      if (active.name !== "wbs") {
        ctx.workbook.worksheets.getItem("wbs").activate();
        await ctx.sync();
      }
    });
  } catch (e) {
    console.warn("wbsシートのアクティブ化に失敗:", e);
  }
}

/* 大分類の変更：受注なら小分類を案件番号セレクトに切替。
   それ以外は、その大分類で使われている既存の小分類をデータリスト（候補）として提示しつつ、
   自由入力でも新しい小分類を追加できるようにする。 */
async function onTaCatChange() {
  const cat = document.getElementById("ta-cat").value;
  const txt = document.getElementById("ta-subcat");
  const sel = document.getElementById("ta-subcat-sel");
  const dl = document.getElementById("ta-subcat-list");
  if (cat === ORDER_CATEGORY) {
    txt.style.display = "none";
    sel.style.display = "";
    sel.innerHTML = `<option value="">読込中…</option>`;
    const ids = await loadOrderCaseIds();
    sel.innerHTML = ids.length
      ? ids.map(x => `<option value="${escapeHtml(x.id)}">${escapeHtml(x.id)}　${escapeHtml(x.client)}</option>`).join("")
      : `<option value="">（対象案件がありません）</option>`;
  } else {
    txt.style.display = "";
    sel.style.display = "none";
    txt.value = "";
    const subs = [...new Set(
      allTasks
        .filter(t => t.category === cat)
        .map(t => t.classification)
        .filter(v => v && String(v).trim() !== "" && v !== "#")
    )];
    dl.innerHTML = subs.map(s => `<option value="${escapeHtml(String(s))}"></option>`).join("");
  }
}

/* 営業報告シートから 状態=受注/受託中 の案件番号を取得 */
async function loadOrderCaseIds() {
  if (!window.Excel) return [];
  try {
    let out = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(EIGYO_SHEET);
      const used = sheet.getUsedRange(true);
      used.load("rowCount");
      await ctx.sync();
      const last = Math.max(used.rowCount, 2);
      const rng = sheet.getRange(`A2:E${last}`);
      rng.load("values");
      await ctx.sync();
      rng.values.forEach(r => {
        const id = (r[0] ?? "").toString().trim();
        const client = (r[1] ?? "").toString().trim();
        const st = (r[4] ?? "").toString().trim();
        if (id && (st === "受注" || st === "受託中")) out.push({ id, client });
      });
    });
    return out;
  } catch (e) {
    console.warn("営業報告シートの読込に失敗:", e);
    return [];
  }
}

/* OK：選択行がタスク範囲内かを検証し、行挿入してタスクを書き込む */
async function saveTaskAdd() {
  const msg = document.getElementById("ta-msg");
  msg.className = "task-msg"; msg.textContent = "";

  const cat = document.getElementById("ta-cat").value;
  const sub = (cat === ORDER_CATEGORY)
    ? document.getElementById("ta-subcat-sel").value
    : document.getElementById("ta-subcat").value.trim();
  const title = document.getElementById("ta-title").value.trim();
  const user = document.getElementById("ta-user").value;
  const start = document.getElementById("ta-start").value;
  const end = document.getElementById("ta-end").value;

  if (!title) { msg.className = "task-msg err"; msg.textContent = "タスク名を入力してください"; return; }
  if (cat === ORDER_CATEGORY && !sub) { msg.className = "task-msg err"; msg.textContent = "案件番号を選択してください"; return; }
  if (!window.Excel) { msg.className = "task-msg err"; msg.textContent = "Excel環境でのみ追加できます"; return; }

  try {
    let inserted = -1;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem("wbs");

      // 選択セルの行
      const selected = ctx.workbook.getSelectedRange();
      selected.load(["rowIndex", "worksheet/name"]);

      // 名前定義「タスク範囲」（ブック→wbsシートの順で検索）
      let nameItem = ctx.workbook.names.getItemOrNullObject(TASK_RANGE_NAME);
      let sheetNameItem = sheet.names.getItemOrNullObject(TASK_RANGE_NAME);
      await ctx.sync();
      if (nameItem.isNullObject && sheetNameItem.isNullObject) {
        throw new Error(`名前定義「${TASK_RANGE_NAME}」が見つかりません。wbsシートに行挿入可能な範囲を「${TASK_RANGE_NAME}」として名前定義してください。`);
      }
      const rangeObj = (!nameItem.isNullObject ? nameItem : sheetNameItem).getRange();
      rangeObj.load(["rowIndex", "rowCount", "worksheet/name"]);
      await ctx.sync();

      if (selected.worksheet.name !== "wbs") {
        throw new Error("wbsシート上で挿入したい行を選択してください。");
      }
      const selRow = selected.rowIndex + 1;             // 1-based
      const rangeTop = rangeObj.rowIndex + 1;
      const rangeBottom = rangeObj.rowIndex + rangeObj.rowCount;
      if (selRow < rangeTop || selRow > rangeBottom) {
        throw new Error(`選択行（${selRow}行目）は「${TASK_RANGE_NAME}」（${rangeTop}〜${rangeBottom}行目）の外です。範囲内の行を選択してください。`);
      }

      // 行挿入（選択行の位置に。既存行は下へ）
      sheet.getRange(`${selRow}:${selRow}`).insert(Excel.InsertShiftDirection.down);
      await ctx.sync();

      // T〜FY列の数式を隣接行からコピー（挿入行の上、先頭行の場合は下からコピー）
      const srcRow = (selRow > rangeTop) ? selRow - 1 : selRow + 1;
      const dst = sheet.getRange(`T${selRow}:FY${selRow}`);
      dst.copyFrom(sheet.getRange(`T${srcRow}:FY${srcRow}`), Excel.RangeCopyType.formulas);

      // 値の書込み
      sheet.getRange(`A${selRow}`).values = [[cat]];
      sheet.getRange(`B${selRow}`).values = [[sub]];
      sheet.getRange(`E${selRow}`).values = [[title]];
      sheet.getRange(`N${selRow}`).values = [[user]];
      if (start) {
        const c = sheet.getRange(`P${selRow}`);
        c.values = [[dateToExcelSerial(new Date(start + "T00:00:00"))]];
        c.numberFormat = [["m/d"]];
      }
      if (end) {
        const c = sheet.getRange(`Q${selRow}`);
        c.values = [[dateToExcelSerial(new Date(end + "T00:00:00"))]];
        c.numberFormat = [["m/d"]];
      }
      await ctx.sync();
      inserted = selRow;
    });

    closeTaskAdd();
    await uiAlert(`${inserted}行目にタスクを追加しました。`);
    await init();   // 再読込してボードへ反映
  } catch (e) {
    msg.className = "task-msg err";
    msg.textContent = e.message || "タスクの追加に失敗しました";
  }
}
