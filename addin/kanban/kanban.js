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

const APP_VERSION = "rev_20260905_f803709";
window.APP_VERSION = APP_VERSION;

/* ============================================================
   シート構造の定点
   ------------------------------------------------------------
   担当者（休み）の行数が変動するため、タスクの開始行を固定で
   持たない。A列に HEADER_KEYWORD がある行をヘッダー行とし、
   その2行下からタスクとして読む。
   優先度は C列。★ に限らず「空欄でなければ優先」と判定する。
   ============================================================ */
const HEADER_KEYWORD = "大分類";   // A列でヘッダー行を探すキーワード
const PRIORITY_MARK  = "★";        // 優先ボタンでC列に書く文字
let wbsHeaderRow = 9;              // 見つかったヘッダー行（1基点）

/** ヘッダー行の探索。通常は api.js の findWbsHeader() を使い、
    これはそれが読み込まれていない場合のフォールバック。 */
function localFindWbsHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 60); i++) {
    const a = rows[i] && rows[i][0];
    if (a != null && a.toString().trim() === HEADER_KEYWORD) {
      return { headerRow: i + 1, dataIdx: i + 2 };
    }
  }
  return { headerRow: 9, dataIdx: 10 };
}

/** C列の値が「優先」かどうか。空欄でなければ優先。 */
function isPriority(v) {
  return v != null && v.toString().trim() !== "";
}

/** カードに出す優先マーク。★以外の手入力値（A・高 など）はそのまま見せる。 */
function priorityMark(v) {
  const t = (v == null ? "" : v.toString()).trim();
  if (!t) return "☆";
  return t.length <= 2 ? t : PRIORITY_MARK;
}

let allTasks = [];
let currentDraggedId = null;
// currentTask は api.js（タスク詳細/備考編集モーダル）が保持する

let selectedUsers = [];
let selectedCategories = [];
let selectedSubCategories = [];
let selectedPeriod = "all";
/* カンバン専用：更新日付(U列)による実績フィルタ  "" | "this" | "last" */
let weekFilter = "";
let showHeld = true;
let showAllDone = false;          // 完了全て（OFF時は直近のみ表示）
let searchQuery = "";

/* ===== 集計タブの状態 ===== */
let currentTab = "board";     // "board" | "sched" | "agg"
let aggSubTab  = "status";    // "status"（対応状況） | "delay"（遅延）
let aggAxis    = "total";     // "total"（総件数） | "cum"（累計） | "day"（当日）
let aggPanelOpen = false;     // 分類フィルタ行の開閉
let selectedDelayUsers = [];  // 遅延タブ専用の担当者フィルタ（対応状況側とは独立）

/* 集計タブ専用のフィルタ（カンバンタブとは完全に分離して管理する）
   カンバン側の検索・担当者・分類を変えても集計には影響しない。 */
let aggUsers = [];
let aggCategories = [];
let aggSubCategories = [];

/* 遅延タブ：KPIクリックで表示するリストの選択 */
let delaySel = "overdue";     // "overdue" | "soon" | "idle" | "held"

/* 遅延タブ：期限接近とみなす残日数 */
const DUE_SOON_DAYS = 3;

/* 集計の達成率警告しきい値（%）。累計・当日にのみ適用 */
const AGG_WARN_RATE = 80;

/* 「完了全て」OFF時に完了レーンへ残す日数（実績完了日 基準） */
const DONE_VISIBLE_DAYS = 15;

/* フィルタ保存フォーマットのバージョン
   v1: "today" = スター付きのみ
   v2: "star"  = スター付きのみ / "today" = 本日（期間内＋遅延＋対応中） */
const FILTER_SCHEMA_VERSION = 2;

// api.js（タスク追加・タスク詳細/備考編集モーダル）からの再描画フック
window.ApiConfig = window.ApiConfig || {};
window.ApiConfig.wbsSheet = "wbs";
window.ApiConfig.eigyoSheet = "営業報告";
// 優先度は C列に移したので、備考の先頭に ☆ を書かない
window.ApiConfig.noteStarSymbol = false;
window.ApiConfig.onLeaveChanged = () => { if (currentTab === "sched") renderScheduleTab(); };
window.ApiConfig.onNoteSaved = () => {
  // api.js の saveNote() は備考先頭の★で isStar を更新するが、
  // 優先度の正は C列なので上書きされた値を戻してから再描画する。
  allTasks.forEach((t) => { t.isStar = isPriority(t.priority); });
  renderBoard();
};
window.ApiConfig.onTaskAdded = () => init();

/* ============================================================
   DOM の準備を待つ
   ------------------------------------------------------------
   Excel on the web では Office.onReady が DOM の解析より先に
   解決することがある。kanban.js は <head> で読み込まれるため、
   その場合 document.getElementById() が null を返し、
   bindStaticUI() の中で TypeError になって初期化が丸ごと止まる。
   （症状：ペインが空白のまま／フィルタが効かない）
   ============================================================ */
function whenDomReady(fn) {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", fn, { once: true });
  } else {
    fn();
  }
}

if (window.Office && Office.onReady) {
  Office.onReady(() => {
    whenDomReady(() => {
      restoreSavedFilters();
      restoreHeldDisplay();
      restoreAllDoneDisplay();
      restoreTabState();
      bindStaticUI();
      init();
    });
  });
} else {
  // ブラウザ直接表示（開発確認用）: Excel連携なしでUIのみ初期化
  window.addEventListener("DOMContentLoaded", () => {
    restoreSavedFilters();
    restoreHeldDisplay();
    restoreAllDoneDisplay();
    restoreTabState();
    bindStaticUI();
    const v = document.getElementById("version-label");
    if (v) v.textContent = APP_VERSION + " (no-office)";
  });
}

/* ============================================================
   初期化
   ============================================================ */
async function init() {
  if (typeof schedRestore === "function") schedRestore();
  await loadExcelData();
  applyTabState();
  renderFilters();
  renderPeriodSegment();
  renderWeekButtons();
  renderBoard();
  renderAggViews();   // 集計はカンバンのフィルタと独立。データ再読込時のみ更新
  if (currentTab === "sched") renderScheduleTab();

  const v = document.getElementById("version-label");
  if (v) v.textContent = APP_VERSION;
}

/* 静的UIのイベント（初回のみ） */
function bindStaticUI() {
  // 検索
  const input = document.getElementById("search-input");
  const clearBtn = document.getElementById("search-clear");

  // 要素が無くても初期化チェーンを止めない（1つの null で全機能が死ぬのを防ぐ）
  if (input) {
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
  }
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      if (input) input.value = "";
      searchQuery = "";
      renderBoard();
      if (input) input.focus();
    });
  }

  // 期間セグメント
  document.querySelectorAll("#seg-period button").forEach(b => {
    b.addEventListener("click", () => setPeriod(b.dataset.p));
  });

  // メインタブ（カンバン／集計）
  document.querySelectorAll("#main-tabs button").forEach(b => {
    b.addEventListener("click", () => switchTab(b.dataset.tab));
  });

  // 集計のサブタブ（対応状況／遅延）
  document.querySelectorAll("#agg-subtabs button").forEach(b => {
    b.addEventListener("click", () => switchAggSub(b.dataset.s));
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

  // 画面サイズ変更時に一覧の高さを追従させる
  let sizeTimer = null;
  window.addEventListener("resize", () => {
    clearTimeout(sizeTimer);
    sizeTimer = setTimeout(sizeAggLists, 120);
  });

  // トグルの初期表示
  const heldToggle = document.getElementById("held-toggle");
  if (heldToggle) heldToggle.classList.toggle("on", showHeld);
  const allDoneToggle = document.getElementById("alldone-toggle");
  if (allDoneToggle) allDoneToggle.classList.toggle("on", showAllDone);
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
      weekFilter = f.weekFilter || "";
      aggUsers = Array.isArray(f.aggUsers) ? f.aggUsers : [];
      aggCategories = Array.isArray(f.aggCategories) ? f.aggCategories : [];
      aggSubCategories = Array.isArray(f.aggSubCategories) ? f.aggSubCategories : [];

      // v1では「本日★」の内部値が "today" だったため "star" に読み替える
      if ((f.v || 1) < 2 && selectedPeriod === "today") selectedPeriod = "star";
    }
  } catch (e) {
    selectedUsers = [];
    selectedCategories = [];
    selectedSubCategories = [];
    selectedPeriod = "all";
    weekFilter = "";
    aggUsers = [];
    aggCategories = [];
    aggSubCategories = [];
  }
}

function saveFilters() {
  try {
    localStorage.setItem("kanban-filters", JSON.stringify({
      users: selectedUsers,
      categories: selectedCategories,
      subCategories: selectedSubCategories,
      period: selectedPeriod,
      weekFilter: weekFilter,
      aggUsers: aggUsers,
      aggCategories: aggCategories,
      aggSubCategories: aggSubCategories,
      v: FILTER_SCHEMA_VERSION,
      timestamp: Date.now()
    }));
  } catch (e) { /* noop */ }
}

function restoreHeldDisplay() {
  const saved = localStorage.getItem("kanban-show-held");
  showHeld = saved !== null ? saved === "true" : true;
}

function restoreAllDoneDisplay() {
  const saved = localStorage.getItem("kanban-show-all-done");
  showAllDone = saved !== null ? saved === "true" : false;  // 既定はOFF
}

function resetSettings() {
  try {
    localStorage.removeItem("kanban-filters");
    localStorage.removeItem("kanban-show-held");
    localStorage.removeItem("kanban-show-all-done");
    localStorage.removeItem("kanban-delay-users");
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
    // ゴースト使用範囲（例: A1:XFD240 = 約393万セル）を拾って
    // Excel on the web の 5MB ペイロード制限を超えるのを防ぐため、
    // 行数だけ先に取得し、必要な A:Z 列のみを読む。
    const used = sheet.getUsedRange(true);
    used.load(["rowIndex", "rowCount"]);
    await ctx.sync();

    const lastRow = Math.max(used.rowIndex + used.rowCount, 11);
    const range = sheet.getRangeByIndexes(0, 0, lastRow, 26); // A1:Z{lastRow}
    range.load("values");
    await ctx.sync();

    const rows = range.values;

    // ヘッダー行（A列＝「大分類」）を探す。担当者の行が増減しても追従する。
    // 判定は api.js の findWbsHeader() に一元化し、3ファイルで挙動をそろえる。
    const hdr = (typeof findWbsHeader === "function")
      ? findWbsHeader(rows)
      : localFindWbsHeader(rows);
    const dataIdx = hdr.dataIdx;
    wbsHeaderRow = hdr.headerRow;

    allTasks = rows.slice(dataIdx).map((row, i) => {
      if (!row[25] || row[19] === "-") return null;

      const t = {
        id: row[24],
        category: row[0],
        classification: row[1],
        priority: row[2],          // C列＝優先度
        title: row[25],
        user: row[13],
        start: row[15],
        end: row[16],
        actualStart: row[17],
        actualEnd: row[18],
        updatedAt: row[20],        // U列=更新日付（今週/先週実績の判定に使用）
        note: row[14],
        rowIndex: dataIdx + i + 1,

        isNoSchedule: !row[15] && !row[16],
        isStar: isPriority(row[2])
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

/* 対応中（実績開始済み・未完了）判定
   status文字列はH列更新で「保留」等に書き換わるため実績日で判定する */
function isInProgress(t) {
  return !!t.actualStart && !t.actualEnd;
}

/* 時刻を落とした日付を返す（不正値はnull） */
function toMidnight(d) {
  if (!d || isNaN(d)) return null;
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

/* 完了タスクが直近 DONE_VISIBLE_DAYS 日以内か
   実績完了日が読めない完了タスクは常に表示する */
function isRecentDone(t) {
  const ae = toMidnight(excelDateToJS(t.actualEnd));
  if (!ae) return true;

  const limit = toMidnight(new Date());
  limit.setDate(limit.getDate() - DONE_VISIBLE_DAYS);
  return ae >= limit;
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
    // 分類（大分類）が未選択なら小分類は非活性
    const subEnabled = selectedCategories.length > 0;
    subCatChip.disabled = !subEnabled;
    subCatChip.title = subEnabled ? "" : "分類を選択すると使用できます";
    if (!subEnabled) {
      const dd = document.getElementById("dd-subcat");
      if (dd) dd.classList.remove("open");
      subCatChip.classList.remove("selected");
      subCatChip.innerHTML = `小分類 <span class="caret"></span>`;
      return;
    }
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

      // 大分類が未選択になったら小分類フィルタは解除（チップが非活性になるため）
      if (!selectedCategories.length) {
        selectedSubCategories = [];
      } else {
        // 選択解除された大分類配下の小分類はフィルタから外す
        selectedSubCategories = selectedSubCategories.filter(s =>
          allTasks.some(t =>
            t.classification === s && selectedCategories.includes(t.category)
          )
        );
      }

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
  const el = document.getElementById("held-toggle");
  if (el) el.classList.toggle("on", showHeld);
  renderBoard();
}

/* ============================================================
   完了全て 表示切替
   OFF: 実績完了日が DONE_VISIBLE_DAYS 日以上前の完了タスクを隠す
   ============================================================ */
function toggleAllDone(e) {
  if (e) e.preventDefault();
  showAllDone = !showAllDone;
  localStorage.setItem("kanban-show-all-done", showAllDone);
  const el = document.getElementById("alldone-toggle");
  if (el) el.classList.toggle("on", showAllDone);
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
    jumpToWbsRow(t.rowIndex);
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

  // サブタスク数（③）。0件は出さない。親が完了なのに残っていればオレンジ
  if (typeof subtaskBadgeHtml === "function") {
    const badge = subtaskBadgeHtml(t.note, { parentDone: t.status === "完了" });
    if (badge) {
      const wrap = document.createElement("span");
      wrap.innerHTML = badge;
      if (wrap.firstChild) meta.appendChild(wrap.firstChild);
    }
  }

  // 完了以外にスターを表示
  if (t.status !== "完了") {
    const star = document.createElement("button");
    star.className = "card-star" + (t.isStar ? " on" : "");
    star.textContent = t.isStar ? priorityMark(t.priority) : "☆";
    star.title = t.isStar
      ? `優先（C列: ${(t.priority == null ? "" : t.priority).toString().trim()}）`
      : "本日の優先タスク";
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
   ------------------------------------------------------------
   escapeHtml は api.js（共通モジュール）で定義されている
   ============================================================ */
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
   ------------------------------------------------------------
   jumpToWbsRow は api.js（共通モジュール）で定義されている。
   ※営業報告アドイン側は自前の jumpToExcel（営業報告シート用）を
     持っているため、名前が衝突しないよう api.js 側は
     jumpToWbsRow という名前にしてある。
   ============================================================ */

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

// dateToExcelSerial は api.js（共通モジュール）で定義されている

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
  }

  if (lane === "doing") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = "";

    if (task.note && task.note.toString().includes("▲")) {
      task.note = task.note.toString().replace(/▲/g, "△");
    }
  }

  if (lane === "done") {
    if (!isValidDate(actualStart)) actualStart = new Date();
    actualEnd = new Date();

    if (task.isStar) task.isStar = false;
    task.priority = "";

    if (task.note && task.note.toString().includes("▲")) {
      task.note = task.note.toString().replace(/▲/g, "△");
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

    // 完了時は優先度（C列）を空にする（備考は触らない）
    if (lane === "done") {
      sheet.getRange(`C${row}`).values = [[""]];
    }

    stampWbsUpdate(sheet, row);
    await ctx.sync();
  });

  await init();
}

/* H列（状態文字列）への書き込みは廃止した。
   状態は 実績開始/完了日（R・S列）と備考の ▲ だけで決まる。
   H列は人が手で書く自由欄として残す。 */

/* ============================================================
   スター切り替え
   ============================================================ */
async function toggleStar(task) {
  const on = !task.isStar;
  task.isStar = on;
  task.priority = on ? PRIORITY_MARK : "";

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const cell = sheet.getRange(`C${task.rowIndex}`);
    cell.values = [[task.priority]];
    stampWbsUpdate(sheet, task.rowIndex);
    await ctx.sync();
  });

  renderBoard();
}

/* ============================================================
   タスク詳細/備考編集モーダル（サブタスクカンバン含む）
   ------------------------------------------------------------
   openModal / closeModal / saveNote / サブタスクカンバン関連は
   api.js（共通モジュール）に移動した
   ============================================================ */

/* ============================================================
   実績フィルタ（更新日付 U列 ベース。カンバンタブ専用）
   ------------------------------------------------------------
   営業報告アドインの「今週実績／先週実績」と同じ考え方で、
   月曜はじまりの週に更新されたタスクだけを表示する。
   先週と今週は排他（一方をONにすると他方はOFF）。
   ============================================================ */
function weekRangeOf(which) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const day = now.getDay();                       // 0=日
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const thisMonday = new Date(now); thisMonday.setDate(now.getDate() + diffToMonday);

  if (which === "this") {
    const sunday = new Date(thisMonday); sunday.setDate(thisMonday.getDate() + 6);
    sunday.setHours(23, 59, 59, 999);
    return { start: thisMonday, end: sunday };
  }
  const lastMonday = new Date(thisMonday); lastMonday.setDate(thisMonday.getDate() - 7);
  const lastSunday = new Date(thisMonday); lastSunday.setDate(thisMonday.getDate() - 1);
  lastSunday.setHours(23, 59, 59, 999);
  return { start: lastMonday, end: lastSunday };
}

function matchesWeekFilter(t) {
  if (!weekFilter) return true;
  const d = excelDateToJS(t.updatedAt);
  if (!d || isNaN(d)) return false;               // 更新日付なしは対象外
  const { start, end } = weekRangeOf(weekFilter);
  return d >= start && d <= end;
}

function setWeekFilter(w) {
  weekFilter = (weekFilter === w) ? "" : w;       // 同じボタンで解除
  saveFilters();
  renderWeekButtons();
  renderBoard();
}

function renderWeekButtons() {
  const t = document.getElementById("btn-thisweek");
  if (t) t.classList.toggle("on", weekFilter === "this");
  const l = document.getElementById("btn-lastweek");
  if (l) l.classList.toggle("on", weekFilter === "last");
}

/* ============================================================
   フィルタ判定（検索を追加）
   ============================================================ */
function isMatch(t) {

  // ★ 実績フィルタ（更新日付）
  if (!matchesWeekFilter(t)) return false;


  // ★ 検索（タスク名・備考）
  if (!matchesSearch(t)) return false;

  // 担当者
  if (selectedUsers.length && !selectedUsers.includes(t.user)) return false;

  // 分類（大分類）
  if (selectedCategories.length && !selectedCategories.includes(t.category)) return false;

  // 小分類
  if (selectedSubCategories.length && !selectedSubCategories.includes(t.classification)) return false;

  // ★ 完了の表示範囲：「完了全て」OFFなら実績完了日が古い完了タスクを隠す
  if (t.status === "完了" && !showAllDone && !isRecentDone(t)) return false;

  // ★ ★フィルタ：スター付きのみ表示
  if (selectedPeriod === "star") {
    return t.isStar;
  }

  // ★ 本日フィルタ：期間内＋遅延＋対応中（完了は対象外）
  if (selectedPeriod === "today") {
    if (t.status === "完了") return false;
    if (isInProgress(t)) return true;        // 対応中は常に表示
    if (t.isNoSchedule) return false;

    const s = toMidnight(excelDateToJS(t.start));
    const e = toMidnight(excelDateToJS(t.end));
    if (!s || !e) return false;

    const today = toMidnight(new Date());
    return e < today || (s <= today && e >= today);  // 遅延 または 期間内
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
          appName: "WBS Kanban",
          version: APP_VERSION,
          position: "left",
          width: 250,
          theme: { accent: "#0E7A5F" },
          footer: "© RightArm",
          currentId: "kanban",                       // menu.json のidと一致で強調表示
          menuUrl: COMMON_BASE + "/menu.json",       // ★ メニュー定義はJSONで一元管理
          localItems: [                              // このアプリ固有の操作
            { section: "操作" },
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
   汎用ダイアログ（uiConfirm/uiAlert/dialogRespond）と
   タスク追加（openTaskAdd/saveTaskAdd 等）は
   api.js（共通モジュール）に移動した
   ============================================================ */


/* ============================================================
   集計タブ
   ------------------------------------------------------------
   ・対応状況：KPI（総数/累計予定/累計実績/当日予定/当日実績）＋
     件数比の積み上げ横棒（全体・大分類別・小分類別・担当者別）＋タスク一覧
   ・遅延：遅延／期限接近／未着手放置／保留の一覧

   指標定義
     累計予定 = 予定終了日(Q) が今日以前
     累計実績 = 実績完了日(S) が今日以前
     当日予定 = 予定終了日(Q) が今日
     当日実績 = 実績完了日(S) が今日

   集計側のフィルタは「検索・担当者・大分類・小分類」のみを適用する。
   期間セグメント／保留トグル／完了全ては board 用途なので集計には効かせない。
   ============================================================ */

/* ===== タブ切替 ===== */
function switchTab(tab) {
  currentTab = tab;
  try { localStorage.setItem("kanban-tab", tab); } catch (e) { /* 保存できなくても継続 */ }
  applyTabState();
  if (tab === "sched") renderScheduleTab();
  else renderBoard();
}

/* ============================================================
   ① スケジュールタブ
   ------------------------------------------------------------
   描画は api.js の renderSchedule()。大分類 → 小分類 → タスクの
   3段で、小分類をクリックするとタスクの線とリストが開く。
   ============================================================ */
function renderScheduleTab() {
  const host = document.getElementById("sched-container");
  if (!host) return;
  if (typeof renderSchedule !== "function") {
    host.innerHTML = '<div style="padding:14px;font-size:11px;color:#93A1AF">'
      + 'api.js が読み込まれていないためスケジュールを表示できません。</div>';
    return;
  }
  renderSchedule(host).catch((e) => {
    console.warn("スケジュールの描画に失敗:", e);
    host.innerHTML = '<div style="padding:14px;font-size:11px;color:#B4262B">'
      + 'スケジュールの描画に失敗しました: ' + escapeHtml(String(e && e.message || e)) + "</div>";
  });
}

/* ④ 個人予定（api.js のモーダル） */
function openPersonalSchedule() {
  if (typeof openMyLeave !== "function") return;
  openMyLeave(null);
}

function switchAggSub(sub) {
  aggSubTab = sub;
  try { localStorage.setItem("kanban-agg-sub", sub); } catch (e) { /* 同上 */ }
  applyTabState();
  renderAggViews();
}

function setAggAxis(axis) {
  aggAxis = axis;
  try { localStorage.setItem("kanban-agg-axis", axis); } catch (e) { /* 同上 */ }
  applyTabState();
  renderAggViews();
}

function restoreTabState() {
  try {
    const t = localStorage.getItem("kanban-tab");
    if (t === "board" || t === "sched" || t === "agg") currentTab = t;
    const s = localStorage.getItem("kanban-agg-sub");
    if (s === "status" || s === "delay") aggSubTab = s;
    const a = localStorage.getItem("kanban-agg-axis");
    if (a === "total" || a === "cum" || a === "day") aggAxis = a;
    const d = localStorage.getItem("kanban-delay-sel");
    if (["overdue", "soon", "idle", "held"].includes(d)) delaySel = d;
    aggPanelOpen = localStorage.getItem("kanban-agg-panel") === "true";
    try {
      const du = JSON.parse(localStorage.getItem("kanban-delay-users") || "[]");
      if (Array.isArray(du)) selectedDelayUsers = du;
    } catch (e) { /* 既定値のまま継続 */ }
  } catch (e) {
    console.log("Tab state restoration error:", e);
  }
}

/* 表示中タブに応じて DOM の表示状態を揃える */
function applyTabState() {
  const isAgg = currentTab === "agg";
  const isSched = currentTab === "sched";

  const board = document.getElementById("board");
  if (board) board.classList.toggle("hidden", isAgg || isSched);

  const agg = document.getElementById("agg-view");
  if (agg) agg.classList.toggle("hidden", !isAgg);

  const sched = document.getElementById("sched-view");
  if (sched) sched.classList.toggle("hidden", !isSched);

  // 集計・スケジュールでは期間・保留・完了全ては使わないため隠す
  const bar = document.getElementById("filter-bar");
  if (bar) bar.classList.toggle("agg-mode", isAgg || isSched);

  document.querySelectorAll("#main-tabs button").forEach(b =>
    b.classList.toggle("on", b.dataset.tab === currentTab));
  document.querySelectorAll("#agg-subtabs button").forEach(b =>
    b.classList.toggle("on", b.dataset.s === aggSubTab));

  const st = document.getElementById("agg-status");
  if (st) st.classList.toggle("hidden", aggSubTab !== "status");
  const dl = document.getElementById("agg-delay");
  if (dl) dl.classList.toggle("hidden", aggSubTab !== "delay");
}

/* 一覧の高さを画面下端までに収める（超える分はスクロール） */
function sizeAggLists() {
  document.querySelectorAll(".agg-list-scroll").forEach(el => {
    const top = el.getBoundingClientRect().top;
    const h = Math.max(window.innerHeight - top - 6, 120);
    el.style.maxHeight = h + "px";
  });
}

/* ===== 集計対象の判定（検索・担当者・大分類・小分類のみ） ===== */
/* 集計タブの絞り込み。カンバンタブの検索・フィルタ・週フィルタは一切適用しない */
function aggMatch(t) {
  if (aggUsers.length && !aggUsers.includes(t.user)) return false;
  if (aggCategories.length && !aggCategories.includes(t.category)) return false;
  if (aggSubCategories.length && !aggSubCategories.includes(t.classification)) return false;
  return true;
}

/* ===== 集計計算 ===== */
function aggregate(rows) {
  const today = toMidnight(new Date());
  const a = { total: 0, todo: 0, doing: 0, done: 0, pc: 0, ac: 0, pd: 0, ad: 0 };

  rows.forEach(t => {
    a.total++;

    if (t.actualEnd) a.done++;
    else if (t.actualStart) a.doing++;
    else a.todo++;

    const end = toMidnight(excelDateToJS(t.end));
    if (end) {
      if (end <= today) a.pc++;
      if (end.getTime() === today.getTime()) a.pd++;
    }

    const ae = toMidnight(excelDateToJS(t.actualEnd));
    if (ae) {
      if (ae <= today) a.ac++;
      if (ae.getTime() === today.getTime()) a.ad++;
    }
  });

  return a;
}

/* 指定キーでグループ化して集計（値が空のものは「未設定」に寄せる） */
function aggregateBy(rows, key) {
  const map = new Map();
  rows.forEach(t => {
    const raw = t[key];
    const name = (raw === null || raw === undefined || String(raw).trim() === "")
      ? "未設定" : String(raw);
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(t);
  });

  // 並び順はカンバンのチップと同じ「シート出現順」に合わせる
  const order = sheetOrder(key);
  const idx = n => {
    const i = order.indexOf(n);
    return i < 0 ? order.length + 1 : i;      // 未設定は末尾
  };

  return [...map.entries()]
    .map(([name, list]) => Object.assign({ name }, aggregate(list)))
    .sort((x, y) => idx(x.name) - idx(y.name));
}

/* 指定列の値をシートの出現順に並べた配列（カンバンのドロップダウンと同じ基準） */
function sheetOrder(key) {
  return [...new Set(allTasks.map(t => t[key]).filter(v => v && v !== "#"))];
}

/* ===== 軸ごとの描画定義 ===== */
const AGG_AXIS_DEF = {
  total: {
    legend: [["seg-done", "完了"], ["seg-doing", "対応中"], ["seg-todo", "未着手"]],
    base: r => r.total,
    segs: r => [
      ["seg-done", r.done, "完" + r.done],
      ["seg-doing", r.doing, "中" + r.doing],
      ["seg-todo", r.todo, "未" + r.todo]
    ],
    meta: r => `${r.done} / ${r.doing} / ${r.todo}`,
    rate: r => r.total ? Math.round(r.done / r.total * 100) : null,
    warn: false,     // 完了率は進行中に低くなるのが自然なため色分けしない
    hint: "総件数：棒＝総件数、内訳は 完了 → 対応中 → 未着手"
  },
  cum: {
    legend: [["seg-act", "累計実績"], ["seg-gap", "予定との差"]],
    base: r => r.pc,
    segs: r => [["seg-act", r.ac, String(r.ac)], ["seg-gap", Math.max(r.pc - r.ac, 0), ""]],
    meta: r => `${r.ac} / ${r.pc}`,
    rate: r => r.pc ? Math.round(r.ac / r.pc * 100) : null,
    warn: true,
    hint: "累計：棒＝累計予定、青＝累計実績、グレー＝未達（達成率80%未満は赤字）"
  },
  day: {
    legend: [["seg-act", "当日実績"], ["seg-gap", "予定との差"]],
    base: r => r.pd,
    segs: r => [["seg-act", r.ad, String(r.ad)], ["seg-gap", Math.max(r.pd - r.ad, 0), ""]],
    meta: r => `${r.ad} / ${r.pd}`,
    rate: r => r.pd ? Math.round(r.ad / r.pd * 100) : null,
    warn: true,
    hint: "当日：棒＝当日予定、青＝当日実績、グレー＝未達（達成率80%未満は赤字）"
  }
};

/* ===== 集計ビューの描画エントリ ===== */
function renderAggViews() {
  if (!document.getElementById("agg-view")) return;
  renderAggStatus(allTasks.filter(aggMatch));
  renderAggDelay(allTasks);          // 遅延はフィルタ対象外（常に全件）
}

/* ===== 対応状況タブ ===== */
function renderAggStatus(rows) {
  const cfg = AGG_AXIS_DEF[aggAxis] || AGG_AXIS_DEF.total;
  const overall = aggregate(rows);

  // 担当者・分類フィルタパネル
  renderAggFilterPanel();

  // KPI（クリックでグラフ軸を切り替える）
  const kpiEl = document.getElementById("agg-kpis");
  if (kpiEl) {
    const defs = [
      ["総数", overall.total, "", "total"],
      ["累計予定", overall.pc, "plan", "cum"],
      ["累計実績", overall.ac, "act", "cum"],
      ["当日予定", overall.pd, "plan", "day"],
      ["当日実績", overall.ad, "act", "day"]
    ];
    kpiEl.innerHTML = defs.map(([label, val, cls, ax]) =>
      `<button class="kpi ${cls} clickable ${aggAxis === ax ? "sel" : ""}" data-ax="${ax}">
         <span class="kpi-k">${escapeHtml(label)}</span>
         <span class="kpi-v">${val}</span>
       </button>`).join("");

    kpiEl.querySelectorAll("[data-ax]").forEach(b => {
      b.addEventListener("click", () => setAggAxis(b.dataset.ax));
    });
  }

  const hintEl = document.getElementById("agg-axis-hint");
  if (hintEl) hintEl.textContent = cfg.hint;

  // 凡例
  const lg = document.getElementById("agg-legend");
  if (lg) {
    lg.innerHTML = cfg.legend
      .map(([c, t]) => `<span><i class="sw ${c}"></i>${escapeHtml(t)}</span>`).join("") +
      `<span class="legend-hint">棒の長さ＝件数比</span>`;
  }

  // 積み上げ横棒（パンくず式ドリルダウン）
  const barsEl = document.getElementById("agg-bars");
  if (barsEl) {
    barsEl.innerHTML = aggDrillHtml(rows, cfg, overall);
    bindAggDrill(barsEl);
  }

  // タスク一覧
  const listEl = document.getElementById("agg-list");
  const cntEl = document.getElementById("agg-list-count");
  if (cntEl) cntEl.textContent = `${rows.length}件`;

  if (listEl) {
    if (!rows.length) {
      listEl.innerHTML = `<div class="agg-empty">該当するタスクがありません</div>`;
    } else {
      const sorted = rows.slice().sort(aggListSort);
      listEl.innerHTML = buildTable(sorted,
        [COL_CAT, COL_SUB, COL_TITLE, COL_USER, COL_PLAN, COL_ACT, COL_STAT]);
      bindAggRowJump(listEl);
    }
  }

  sizeAggLists();
}

/* 一覧の並び：遅延 → 対応中 → 未着手 → 完了、同区分内は予定終了日順 */
function aggListSort(a, b) {
  const rank = t => {
    if (t.actualEnd) return 3;
    if (isOverdue(t)) return 0;
    if (isInProgress(t)) return 1;
    return 2;
  };
  const d = rank(a) - rank(b);
  if (d !== 0) return d;

  const ea = toMidnight(excelDateToJS(a.end));
  const eb = toMidnight(excelDateToJS(b.end));
  if (!ea && !eb) return 0;
  if (!ea) return 1;
  if (!eb) return -1;
  return ea - eb;
}

function aggBarRow(r, cfg, max, showLabels, opts) {
  const o = opts || {};
  const base = cfg.base(r);
  const width = max > 0 && base > 0 ? Math.max(base / max * 100, 4) : 0;
  const rate = cfg.rate(r);

  let cls = "rate";
  if (rate === null) cls = "rate none";
  else if (cfg.warn) cls = rate < AGG_WARN_RATE ? "rate warn" : "rate ok";

  const segs = cfg.segs(r).filter(s => s[1] > 0);
  const bar = segs.length
    ? `<div class="bar" style="width:${width}%">` +
      segs.map(([c, v, lbl]) =>
        `<i class="${c}" style="flex:${v}">${showLabels ? escapeHtml(lbl) : ""}</i>`).join("") +
      `</div>`
    : `<span class="bar-none">—</span>`;

  return `<div class="agg-row ${o.cls || ""} ${o.drill ? "clickable" : ""}"${o.drill ? ` data-drill="${escapeHtml(o.drill)}"` : ""}>
    <span class="agg-name" title="${escapeHtml(r.name)}">${escapeHtml(r.name)}${o.caret ? `<i class="caret">${o.caret}</i>` : ""}</span>
    <span class="agg-track">${bar}</span>
    <span class="agg-meta">${escapeHtml(cfg.meta(r))}</span>
    <span class="agg-${cls}">${rate === null ? "—" : rate + "%"}</span>
  </div>`;
}

/* ============================================================
   ドリルダウン（パンくず式）
   ------------------------------------------------------------
   階層は「フィルタの選択状態」から導出する。専用の状態を持たない
   ことで、チップ操作とドリルダウンが常に一致する。
     レベル0: 大分類別（大分類の単一選択なし）
     レベル1: 小分類別（大分類を1つだけ選択中）
     レベル2: 担当者別（大分類・小分類を1つずつ選択中）
   ============================================================ */
function aggLevel() {
  const c = aggCategories.length;
  const s = aggSubCategories.length;
  if (c === 1 && s === 1) return 2;
  if (c === 1) return 1;
  return 0;
}

const AGG_LEVELS = [
  { key: "category",       title: "大分類別",  drill: "cat"  },
  { key: "classification", title: "小分類別",  drill: "sub"  },
  { key: "user",           title: "担当者別",  drill: "user" }
];

function aggDrillHtml(rows, cfg, overall) {
  const level = aggLevel();
  const def = AGG_LEVELS[level];

  // パンくず
  let crumb = `<button data-drill="root">全体</button>`;
  if (level >= 1) {
    const cat = selectedCategories[0];
    crumb += `<span class="sepa">›</span>` + (level >= 2
      ? `<button data-drill="cat:${escapeHtml(cat)}">${escapeHtml(cat)}</button>`
      : `<span class="cur">${escapeHtml(cat)}</span>`);
  }
  if (level >= 2) {
    crumb += `<span class="sepa">›</span><span class="cur">${escapeHtml(selectedSubCategories[0])}</span>`;
  }
  crumb += `<span class="lvl">表示中：${def.title}</span>`;

  const headName = level === 0 ? "全体"
    : (level === 1 ? selectedCategories[0] : selectedSubCategories[0]);

  let html = `<div class="agg-crumb">${crumb}</div>`;
  html += aggBarRow(Object.assign({ name: headName }, overall), cfg, cfg.base(overall), true, { cls: "head" });

  if (!rows.length) {
    return html + `<div class="agg-empty">該当するタスクがありません</div>`;
  }

  const children = aggregateBy(rows, def.key);
  const max = Math.max(...children.map(cfg.base));

  html += `<div class="agg-sec-head">${def.title}</div>`;
  html += children.map(r => {
    // 「未設定」は絞り込みキーにできないためドリル不可
    const canDrill = r.name !== "未設定";
    return aggBarRow(r, cfg, max, false, {
      drill: canDrill ? `${def.drill}:${r.name}` : null,
      caret: canDrill ? (level < 2 ? "›" : "＋") : ""
    });
  }).join("");

  html += `<div class="agg-hint">${level < 2
    ? "行をクリックすると1段深く掘り、フィルタにも反映されます。"
    : "行をクリックすると担当者フィルタを切り替えます。"}</div>`;

  return html;
}

/* ドリルダウン操作。フィルタ状態を直接書き換えることで
   チップ・カンバン・集計・タスク一覧のすべてが同じ条件で揃う。 */
function aggDrill(kind, name) {
  if (kind === "root") {
    aggCategories = [];
    aggSubCategories = [];
  } else if (kind === "cat") {
    aggCategories = [name];
    aggSubCategories = [];
  } else if (kind === "sub") {
    aggSubCategories = [name];
  } else if (kind === "user") {
    // 同じ担当者を再クリックしたら解除
    aggUsers = (aggUsers.length === 1 && aggUsers[0] === name) ? [] : [name];
  }

  saveFilters();
  renderAggViews();   // カンバン側は再描画しない（フィルタが独立しているため）
}

function bindAggDrill(scope) {
  scope.querySelectorAll("[data-drill]").forEach(el => {
    el.addEventListener("click", () => {
      const v = el.dataset.drill;
      const i = v.indexOf(":");
      if (i < 0) aggDrill(v, null);
      else aggDrill(v.slice(0, i), v.slice(i + 1));
    });
  });
}

/* ============================================================
   一覧テーブル共通
   ------------------------------------------------------------
   ・検索対象はタスク名と備考のみ
   ・タスク名以外の列は「最大文字数」に合わせた固定幅、
     余りをタスク名列が受け取る（table-layout:fixed + colgroup）
   ============================================================ */
/* タスク名列に最低限確保する幅（px） */
const TITLE_MIN_PX = 110;

function matchesSearch(t) {
  if (!searchQuery) return true;
  const q = searchQuery.toLowerCase();
  return (t.title || "").toString().toLowerCase().includes(q) ||
         (t.note  || "").toString().toLowerCase().includes(q);
}

/* 全角を2、半角を1として表示幅を数える（ch単位に対応させるため） */
function dispLen(v) {
  const str = String(v == null ? "" : v);
  let n = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    n += (c >= 0x1100 && (c <= 0x115F || c === 0x2192 || c === 0x2605 || c === 0x2606 ||
          (c >= 0x2E80 && c <= 0xA4CF) || (c >= 0xAC00 && c <= 0xD7A3) ||
          (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFE30 && c <= 0xFE6F) ||
          (c >= 0xFF00 && c <= 0xFF60) || (c >= 0xFFE0 && c <= 0xFFE6))) ? 2 : 1;
  }
  return n;
}

/* cols: [{ label, get, cls, flex }]  flex:true の列が残り幅を受け取る */
function buildTable(rows, cols, extraClass) {
  const widths = cols.map(c => {
    if (c.flex) return null;
    const max = rows.reduce((m, r) => Math.max(m, dispLen(c.get(r))), dispLen(c.label));
    // 余白2ch分を加算し、極端な長さは頭打ちにする
    return Math.min(Math.max(max + 2, 5), 24);
  });

  const colgroup = widths
    .map(w => w === null ? `<col>` : `<col style="width:${w}ch">`).join("");

  // 固定列の合計＋タスク名の最低幅を下限にする。
  // 狭いペインではタスク名が潰れる代わりに横スクロールさせる。
  const fixedCh = widths.reduce((a, w) => a + (w || 0), 0);
  const minWidth = `calc(${fixedCh}ch + ${TITLE_MIN_PX}px)`;

  const head = cols.map(c => `<th class="${c.cls || ""}">${escapeHtml(c.label)}</th>`).join("");

  const body = rows.map(r =>
    `<tr data-row="${r.rowIndex}">` +
    cols.map(c => {
      const v = c.get(r);
      const cell = c.html ? c.html(r) : escapeHtml(v);
      return `<td class="${c.cls || ""}">${cell}</td>`;
    }).join("") +
    `</tr>`).join("");

  return `<div class="agg-list-scroll">
    <table class="agg-table sticky ${extraClass || ""}" style="min-width:${minWidth}">
      <colgroup>${colgroup}</colgroup>
      <thead><tr>${head}</tr></thead>
      <tbody>${body}</tbody>
    </table>
  </div>`;
}

/* 一覧の共通列（タスク名は flex で残り幅） */
function planText(t) { return t.isNoSchedule ? "\u2014" : `${fmt(t.start)}\u2192${fmt(t.end)}`; }
function actText(t) {
  if (!t.actualStart) return "\u2014";
  return `${fmt(t.actualStart)}\u2192${t.actualEnd ? fmt(t.actualEnd) : ""}`;
}
function statusText(t) {
  if (t.actualEnd) return "\u5b8c\u4e86";
  if (isOverdue(t)) return "\u9045\u5ef6";
  if (isInProgress(t)) return "\u5bfe\u5fdc\u4e2d";
  return "\u672a\u7740\u624b";
}
function statusPill(t) {
  const label = statusText(t);
  const cls = label === "完了" ? "p-done"
    : label === "遅延" ? "p-late"
    : label === "対応中" ? "p-doing" : "p-todo";
  return `<span class="pill ${cls}">${label}</span>`;
}

const COL_TITLE = { label: "タスク", cls: "c-title", flex: true, get: t => t.title || "" };
const COL_CAT   = { label: "大分類", cls: "c-cat", get: t => t.category || "" };
const COL_SUB   = { label: "小分類", cls: "c-cat", get: t => t.classification || "" };
const COL_USER  = { label: "担当", cls: "c-user", get: t => t.user || "" };
const COL_PLAN  = { label: "予定", cls: "c-date", get: planText,
                    html: t => `<span class="${isOverdue(t) ? "late" : ""}">${escapeHtml(planText(t))}</span>` };
const COL_ACT   = { label: "実績", cls: "c-date", get: actText };
const COL_STAT  = { label: "状態", cls: "c-stat", get: statusText, html: statusPill };

/* ============================================================
   集計タブのフィルタパネル
   ------------------------------------------------------------
   担当者は常時1行。分類・小分類は折りたたみ（既定は閉）。
   閉じていても選択中の内容をバッジに出すので絞り込みを見落とさない。
   選択状態は既存の selectedUsers / selectedCategories /
   selectedSubCategories をそのまま使うため、カンバンのチップ・
   パンくず・グラフのどこから操作しても表示が一致する。
   ============================================================ */
function renderAggFilterPanel() {
  const el = document.getElementById("agg-filter");
  if (!el) return;

  const users = sheetOrder("user");
  const cats = sheetOrder("category");
  const cat = aggCategories.length === 1 ? aggCategories[0] : null;
  const sub = aggSubCategories.length === 1 ? aggSubCategories[0] : null;

  const badge = cat ? (sub ? `${cat} › ${sub}` : cat) : "すべて";

  let html = `<div class="agg-fpanel">
    <div class="f-row">
      <span class="f-label">担当者</span>
      <span class="f-chips">${chipsHtml(users, aggUsers, "user")}</span>
    </div>
    <div class="f-row">
      <div style="flex:1;min-width:0">
        <button class="f-toggle" data-panel="toggle">
          <span class="cv">${aggPanelOpen ? "▾" : "▸"}</span>
          <span class="lbl">分類で絞り込む</span>
          <span class="badge ${cat ? "" : "off"}">${escapeHtml(badge)}</span>
          ${cat ? `<span class="clr" data-fkey="cat" data-fval="">解除</span>` : ""}
        </button>`;

  if (aggPanelOpen) {
    const subs = cat
      ? [...new Set(allTasks.filter(t => t.category === cat)
          .map(t => t.classification).filter(v => v && v !== "#"))]
      : [];

    html += `<div class="f-sub">
      <div class="f-row inner">
        <span class="f-label sub">分類</span>
        <span class="f-chips">${chipsHtml(cats, aggCategories, "cat")}</span>
      </div>
      <div class="f-row inner">
        <span class="f-label sub">小分類</span>
        <span class="f-chips">${cat
          ? (subs.length ? chipsHtml(subs, aggSubCategories, "sub")
                         : `<span class="f-chip dis">小分類なし</span>`)
          : `<span class="f-chip dis">分類を選択すると表示</span>`}</span>
      </div>
    </div>`;
  }

  html += `</div></div></div>`;
  el.innerHTML = html;

  el.querySelectorAll("[data-panel]").forEach(b => {
    b.addEventListener("click", (e) => {
      if (e.target.closest("[data-fkey]")) return;   // 「解除」は別処理
      aggPanelOpen = !aggPanelOpen;
      try { localStorage.setItem("kanban-agg-panel", aggPanelOpen); } catch (err) { /* 継続 */ }
      renderAggFilterPanel();
      sizeAggLists();
    });
  });

  el.querySelectorAll("[data-fkey]").forEach(b => {
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      applyAggFilter(b.dataset.fkey, b.dataset.fval || null);
    });
  });
}

function chipsHtml(items, selected, key) {
  const all = `<button class="f-chip ${selected.length ? "" : "on"}" data-fkey="${key}" data-fval="">すべて</button>`;
  return all + items.map(v =>
    `<button class="f-chip ${selected.includes(v) ? "on" : ""}" data-fkey="${key}" data-fval="${escapeHtml(v)}">${escapeHtml(v)}</button>`
  ).join("");
}

/* delay-user キーは renderDelayFilterPanel が個別にクリック処理するため、
   applyAggFilter 側のグローバル委譲とは衝突しない（別要素・別リスナー）。 */

/* パネルのチップは単一選択（同じ値の再クリックで解除） */
function applyAggFilter(key, value) {
  const set = (cur) => (!value || (cur.length === 1 && cur[0] === value)) ? [] : [value];

  if (key === "user") {
    aggUsers = set(aggUsers);
  } else if (key === "cat") {
    aggCategories = set(aggCategories);
    aggSubCategories = [];
  } else if (key === "sub") {
    aggSubCategories = set(aggSubCategories);
  }

  saveFilters();
  renderAggViews();   // カンバン側は再描画しない（フィルタが独立しているため）
}

/* 遅延タブ専用の担当者フィルタパネル（分類は対象外） */
function renderDelayFilterPanel() {
  const el = document.getElementById("delay-filter");
  if (!el) return;

  const users = sheetOrder("user");
  el.innerHTML = `<div class="agg-fpanel">
    <div class="f-row">
      <span class="f-label">担当者</span>
      <span class="f-chips">${chipsHtml(users, selectedDelayUsers, "delay-user")}</span>
    </div>
  </div>`;

  el.querySelectorAll("[data-fkey]").forEach(b => {
    b.addEventListener("click", () => {
      const v = b.dataset.fval || null;
      selectedDelayUsers = (!v || (selectedDelayUsers.length === 1 && selectedDelayUsers[0] === v)) ? [] : [v];
      try { localStorage.setItem("kanban-delay-users", JSON.stringify(selectedDelayUsers)); } catch (e) { /* 継続 */ }
      renderAggDelay(allTasks);
    });
  });
}

/* ===== 遅延タブ ===== */
function isOverdue(t) {
  if (t.actualEnd) return false;              // 完了は対象外
  const end = toMidnight(excelDateToJS(t.end));
  if (!end) return false;
  return end < toMidnight(new Date());
}

function daysBetween(from, to) {
  return Math.round((to - from) / 86400000);
}

/* 予定終了日の超過日数（遅延でなければ null） */
function overdueDays(t) {
  const end = toMidnight(excelDateToJS(t.end));
  if (!end) return null;
  const d = daysBetween(end, toMidnight(new Date()));
  return d > 0 ? d : null;
}

/* 予定終了までの残日数（過ぎていれば負値） */
function daysToDue(t) {
  const end = toMidnight(excelDateToJS(t.end));
  if (!end) return null;
  return daysBetween(toMidnight(new Date()), end);
}

/* 未着手のまま予定開始日を過ぎている日数 */
function idleDays(t) {
  if (t.actualStart || t.actualEnd) return null;
  const s = toMidnight(excelDateToJS(t.start));
  if (!s) return null;
  const d = daysBetween(s, toMidnight(new Date()));
  return d > 0 ? d : null;
}

function isHeld(t) {
  return !!(t.note && String(t.note).includes("▲"));
}

function renderAggDelay(allRows) {
  renderDelayFilterPanel();

  // 遅延タブは「担当者」のみで絞り込む（分類・検索は対象外）
  const rows = selectedDelayUsers.length
    ? allRows.filter(t => selectedDelayUsers.includes(t.user))
    : allRows;

  const groups = {
    overdue: {
      label: "遅延", kpiCls: "bad",
      title: "遅延タスク（超過日数順）",
      valueHead: "超過",
      rows: rows.filter(isOverdue)
        .sort((x, y) => (overdueDays(y) || 0) - (overdueDays(x) || 0)),
      value: t => `<span class="d-bad">${overdueDays(t)}日</span>`,
      valueText: t => `${overdueDays(t)}日`
    },
    soon: {
      label: `期限${DUE_SOON_DAYS}日内`, kpiCls: "warn",
      title: `期限接近（${DUE_SOON_DAYS}日以内）`,
      valueHead: "期限",
      rows: rows.filter(t => {
        if (t.actualEnd) return false;
        const d = daysToDue(t);
        return d !== null && d >= 0 && d <= DUE_SOON_DAYS;
      }).sort((x, y) => (daysToDue(x) || 0) - (daysToDue(y) || 0)),
      value: t => { const d = daysToDue(t); return `<span class="d-warn">${d === 0 ? "本日" : "残" + d + "日"}</span>`; },
      valueText: t => { const d = daysToDue(t); return d === 0 ? "本日" : "残" + d + "日"; }
    },
    idle: {
      label: "未着手放置", kpiCls: "warn",
      title: "未着手のまま放置（予定開始日を経過）",
      valueHead: "放置",
      rows: rows.filter(t => idleDays(t) !== null)
        .sort((x, y) => idleDays(y) - idleDays(x)),
      value: t => `<span class="d-warn">${idleDays(t)}日</span>`,
      valueText: t => `${idleDays(t)}日`
    },
    held: {
      label: "保留", kpiCls: "warn",
      title: "保留中（備考に▲）",
      valueHead: "状況",
      rows: rows.filter(t => isHeld(t) && !t.actualEnd)
        .sort((x, y) => (overdueDays(y) || 0) - (overdueDays(x) || 0)),
      value: t => { const d = overdueDays(t); return d ? `<span class="d-bad">${d}日超過</span>` : `<span class="d-none">—</span>`; },
      valueText: t => { const d = overdueDays(t); return d ? `${d}日超過` : "—"; }
    }
  };

  const order = ["overdue", "soon", "idle", "held"];
  if (!order.includes(delaySel)) delaySel = "overdue";

  // KPI（クリックで下の一覧を切り替え）
  const kpiEl = document.getElementById("delay-kpis");
  if (kpiEl) {
    kpiEl.innerHTML = order.map(k => {
      const g = groups[k];
      return `<button class="kpi ${g.kpiCls} clickable ${k === delaySel ? "sel" : ""}" data-delay="${k}">
        <span class="kpi-k">${escapeHtml(g.label)}</span>
        <span class="kpi-v">${g.rows.length}</span>
      </button>`;
    }).join("");

    kpiEl.querySelectorAll("[data-delay]").forEach(btn => {
      btn.addEventListener("click", () => {
        delaySel = btn.dataset.delay;
        try { localStorage.setItem("kanban-delay-sel", delaySel); } catch (e) { /* 継続 */ }
        renderAggDelay(rows);
      });
    });
  }

  // 選択されたグループの一覧のみ表示
  const body = document.getElementById("delay-body");
  if (!body) return;

  const g = groups[delaySel];
  let html = `<div class="agg-sec-head">${escapeHtml(g.title)} <span class="cnt">${g.rows.length}件</span></div>`;

  if (!g.rows.length) {
    html += `<div class="agg-empty">該当なし</div>`;
    body.innerHTML = html;
    return;
  }

  const colValue = {
    label: g.valueHead, cls: "c-val",
    get: g.valueText, html: g.value
  };
  const colDue = {
    label: "予定終了", cls: "c-date",
    get: t => t.isNoSchedule ? "—" : fmt(t.end)
  };

  html += buildTable(g.rows, [COL_CAT, COL_SUB, COL_TITLE, COL_USER, colDue, colValue]);
  body.innerHTML = html;

  bindAggRowJump(body);
  sizeAggLists();
}

/* 行クリックでExcelの該当行へジャンプ／右クリックで備考モーダル */
function bindAggRowJump(scope) {
  scope.querySelectorAll("tr[data-row]").forEach(tr => {
    const rowIndex = Number(tr.dataset.row);

    tr.addEventListener("click", () => {
      jumpToWbsRow(rowIndex).catch(e => console.log("jump error:", e));
    });

    tr.addEventListener("contextmenu", async (e) => {
      e.preventDefault();
      const t = allTasks.find(x => x.rowIndex === rowIndex);
      if (t) await openModal(t);
    });
  });
}