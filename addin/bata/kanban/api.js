/* ============================================================
 * api.js — 共通モジュール（WBS連携）
 * ------------------------------------------------------------
 * カンバンアドインと営業報告アドインの両方から呼び出す共通部品。
 *
 *   ① タスク追加モーダル（wbsシートの「タスク範囲」への行挿入）
 *   ④ タスク詳細／備考編集モーダル（サブタスクカンバン込み）
 *   ⑤ ミニカンバン（案件配下タスクの3レーン表示・D&D対応）
 *   ＋ 汎用ダイアログ（uiAlert/uiConfirm）、escapeHtml、日付変換など
 *     どちらのアドインからも使う小さなユーティリティ
 *
 * 使い方（呼び出し側 index.html）
 * ------------------------------------------------------------
 *   <link rel="stylesheet" href="../common/api.css">
 *   <script src="../common/api.js"></script>
 *   <script src="kanban.js"></script>   ← api.js より後に読み込む
 *
 * モーダルのDOM（#modal, #task-modal, #dialog-modal）は
 * このファイルが自動で <body> に追加するため、呼び出し側の
 * index.html に書く必要はない（すでに存在する場合は何もしない）。
 *
 * 呼び出し側から使うグローバル関数
 * ------------------------------------------------------------
 *   openModal(task)   … task = { rowIndex, title, note, isStar? }
 *                        タスク詳細／備考編集モーダルを開く。
 *   openTaskAdd(preset?) … タスク追加モーダルを開く。
 *                        大分類/担当者候補は本ファイルがwbsシートから
 *                        直接読み込むため、呼び出し側でallTasksなどを
 *                        用意する必要はない。
 *                        preset = { category, caseId } を渡すと、
 *                        大分類と案件番号（受注時の小分類候補）を
 *                        あらかじめ選択した状態で開ける。
 *   renderMiniKanban(container, matchFn, opts)
 *                     … container（要素 or id文字列）に、matchFnに
 *                       一致するwbsタスクの未着手/対応中/完了の
 *                       3レーン・ミニカンバンを描画する。
 *                       ドラッグ＆ドロップで実績日をwbsへ書き込み、
 *                       完了後は自動で再描画する。
 *                       例）
 *                         renderMiniKanban(
 *                           "mini-kanban-AG03",
 *                           matchByCaseId("AG-03"),
 *                           { onChanged: () => refreshBadge() }
 *                         );
 *   matchByCaseId(caseId)
 *                     … 小分類（B列）が案件番号と一致するかを見る
 *                       renderMiniKanban 用の絞り込み関数を作るヘルパー。
 *   jumpToWbsRow(row) … wbsシートの指定行をアクティブ化して選択する。
 *                       （呼び出し側アプリが自前の "jumpToExcel" を
 *                       持っていて名前が衝突するケースがあるため、
 *                       あえて別名にしてある）
 *
 * 連携フック（呼び出し側が必要に応じて設定する）
 * ------------------------------------------------------------
 *   window.ApiConfig.onNoteSaved  = fn   … 備考保存後に呼ばれる（軽量再描画用）
 *   window.ApiConfig.onTaskAdded  = fn   … タスク追加後に呼ばれる（全体再読込用）
 *   window.ApiConfig.wbsSheet     = "wbs"       （既定値）
 *   window.ApiConfig.eigyoSheet   = "営業報告"   （既定値）
 *   window.ApiConfig.orderCategory= "受注"       （既定値）
 *   window.ApiConfig.taskRangeName= "タスク範囲" （既定値）
 * ============================================================ */

/* ------------------------------------------------------------
   以降は IIFE で包み、内部変数をグローバルへ漏らさない。
   api.js と呼び出し側アプリ（app.js / kanban.js）は通常スクリプトとして
   グローバルスコープを共有するため、同名のトップレベル let / const が
   両方にあると SyntaxError となり、後から読み込まれる側のスクリプトが
   丸ごと実行されなくなる（実際に dialogResolve の衝突で営業報告の
   app.js が全く動かなくなる不具合が発生した）。
   公開が必要な関数だけを末尾で window に明示的に載せる。
   ------------------------------------------------------------ */
(function () {
"use strict";

window.ApiConfig = Object.assign({
  wbsSheet: "wbs",
  eigyoSheet: "営業報告",
  orderCategory: "受注",
  taskRangeName: "タスク範囲",
  onNoteSaved: null,
  onTaskAdded: null,
  onLeaveChanged: null,

  /* 優先度は C列に移したので備考の先頭に ☆ を付ける必要はない。
     既存の呼び出し側（バグ管理など）を壊さないため既定は true。 */
  noteStarSymbol: true,

  /* 休み予定：wbs シート上部の 担当者 × 日 マトリクス */
  leaveGrid: {
    nameCol: "N", firstRow: 3,
    dateCol1: "AB", dateColEnd: "FY",
    typeRow: 2, typeCol1: "P", typeColEnd: "AA",
    baseDateCell: "R1",
    scanRows: 60
  },
  /* 種別の先頭1文字 → 表示スタイルと稼働係数。増えたら1行足すだけ。 */
  leaveStyles: {
    "夏": { style: "full",   work: 0,   mark: "夏" },
    "年": { style: "full",   work: 0,   mark: "年" },
    "A":  { style: "am",     work: 0.5, mark: "AM" },
    "P":  { style: "pm",     work: 0.5, mark: "PM" },
    "リ": { style: "remote", work: 1,   mark: "R" },
    "不": { style: "away",   work: 0,   mark: "不" },
    "○": { style: "mark",   work: 1,   mark: "○" },
    "*":  { style: "full",   work: 0 }
  }
}, window.ApiConfig || {});

function cfg() { return window.ApiConfig; }

/* ============================================================
   汎用ユーティリティ（escapeHtml / 日付変換）
   ============================================================ */
function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

/* wbsシートの更新管理列
   U = カンバン/営業報告からの更新日付（V・Wは予備。行挿入時にクリアする） */
const WBS_UPDATE_COL = "U";
const WBS_CLEAR_COLS = "U:W";

/* wbsへ書き込むすべての経路から呼ぶ。既存の Excel.run 内で使うこと
   （単体で sync はしない）。 */
/* ============================================================
   wbs シートのヘッダー行を探す（共通）
   ------------------------------------------------------------
   シート上部に担当者ごとの休み予定マトリクスがあり、その行数が
   変動するため、タスクの開始行を 11 行目固定で持てない。
   A列に「大分類」がある行をヘッダー行とし、その2行下から
   タスクとして読む。見つからないときは従来どおり 11 行目。
   ============================================================ */
const WBS_HEADER_KEYWORD = "大分類";

function findWbsHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 60); i++) {
    const a = rows[i] && rows[i][0];
    if (a != null && a.toString().trim() === WBS_HEADER_KEYWORD) {
      return { headerRow: i + 1, dataIdx: i + 2 };   // headerRow は1基点、dataIdx は0基点
    }
  }
  return { headerRow: 9, dataIdx: 10 };
}

function stampWbsUpdate(sheet, row) {
  try {
    const c = sheet.getRange(`${WBS_UPDATE_COL}${row}`);
    c.values = [[dateToExcelSerial(new Date())]];
    c.numberFormat = [["yyyy/m/d"]];
  } catch (e) {
    console.warn("更新日付の打刻に失敗:", e);
  }
}

function dateToExcelSerial(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return "";
  const excelEpoch = new Date(1900, 0, 1);
  const msPerDay = 24 * 60 * 60 * 1000;
  const daysDiff = Math.floor((date - excelEpoch) / msPerDay);
  return daysDiff + (date >= new Date(1900, 2, 1) ? 2 : 1);
}

/* ============================================================
   汎用ダイアログ（Office環境では window.confirm/alert 不可）
   ------------------------------------------------------------
   ※ 変数名は "wapi" プレフィックス必須。
     api.js と呼び出し側アプリはどちらも通常スクリプト（module ではない）
     としてグローバルスコープを共有するため、同名のトップレベル let を
     両方で宣言すると SyntaxError になり、後から読み込まれる側の
     スクリプトが丸ごと実行されなくなる。
     （実際に営業報告アドインの dialogResolve と衝突して app.js が
       全く動かなくなる不具合が発生した）

     一方 uiConfirm / uiAlert / dialogRespond は関数宣言なので、
     呼び出し側が同名の関数を持つ場合は後勝ちで上書きされる。
     これは意図した動作で、営業報告では同アプリ独自のダイアログ
     （#dialog-modal を style.display で開閉する実装）が使われ、
     カンバンでは下記の api.js 版が使われる。
   ============================================================ */
let wapiDialogResolve = null;
function uiConfirm(message) {
  return new Promise(resolve => {
    wapiDialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "";
    document.getElementById("dialog-modal").classList.remove("wapi-hidden");
  });
}
function uiAlert(message) {
  return new Promise(resolve => {
    wapiDialogResolve = resolve;
    document.getElementById("dialog-msg").textContent = message;
    document.getElementById("dialog-cancel").style.display = "none";
    document.getElementById("dialog-modal").classList.remove("wapi-hidden");
  });
}
function dialogRespond(ok) {
  document.getElementById("dialog-modal").classList.add("wapi-hidden");
  const r = wapiDialogResolve;
  wapiDialogResolve = null;
  if (r) r(ok);
}

/* api.js 内部から確認ダイアログを出すときは必ずこれを使う。
   呼び出し側アプリが独自の uiAlert を持っていればそちらを優先する
   （営業報告は #dialog-modal を style.display で開閉する独自実装のため、
     api.js 版の classList 操作では表示できない）。 */
function apiAlert(message) {
  if (typeof window.uiAlert === "function" && window.uiAlert !== uiAlert) {
    return window.uiAlert(message);
  }
  return uiAlert(message);
}

/* ============================================================
   ④ タスク詳細／備考編集モーダル
   ------------------------------------------------------------
   ・O列の備考をタスクごとに編集し保存する
   ・備考テキスト内の □未着手 ◎対応中 ■完了 でサブタスクカンバンを表示
   ・呼び出し側は task = { rowIndex, title, note } を渡すだけでよい
   ============================================================ */
let currentTask = null;

function ensureStatusSymbols(noteText) {
  if (!noteText) noteText = "";
  const lines = noteText.split("\n");
  let firstLine = lines[0] || "";
  if (cfg().noteStarSymbol !== false
      && !firstLine.includes("★") && !firstLine.includes("☆")) firstLine = "☆" + firstLine;
  if (!firstLine.includes("▲") && !firstLine.includes("△")) firstLine = firstLine + "△";
  lines[0] = firstLine;
  return lines.join("\n");
}

async function openModal(task) {
  currentTask = task;

  // O列から最新の備考内容を取得
  let originalNote = "";
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
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
    displayNote = (cfg().noteStarSymbol === false ? "△" : "☆△") + "\n＜タスク＞\n＜状況＞";
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
  modal.classList.remove("wapi-hidden");

  const handleEscKey = (event) => {
    if (event.key === "Escape") closeModal();
  };
  const handleOverlayClick = (event) => {
    if (event.target === modal) {
      const currentNote = document.getElementById("modal-note").value;
      if (currentNote === displayNote) closeModal();
    }
  };
  const modalContent = modal.querySelector(".wapi-modal-content");
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
  modal.classList.add("wapi-hidden");
  if (modal._cleanup) {
    modal._cleanup();
    modal._cleanup = null;
  }
}

async function saveNote() {
  const note = document.getElementById("modal-note").value;

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
    const row = currentTask.rowIndex;

    const cell = sheet.getRange(`O${row}`);
    cell.values = [[note]];
    cell.format.wrapText = false;

    const entireRow = sheet.getRange(`${row}:${row}`);
    entireRow.format.rowHeight = 20;

    stampWbsUpdate(sheet, row);
    await ctx.sync();
  });

  if (currentTask) {
    currentTask.note = note;
    // 優先度の正は C列。noteStarSymbol:false のときは備考で上書きしない
    if (cfg().noteStarSymbol !== false) currentTask.isStar = note.startsWith("★");
  }

  closeModal();
  if (typeof cfg().onNoteSaved === "function") cfg().onNoteSaved();
}

/* ===== サブタスクカンバン（備考モーダル内） ===== */
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

function onModalNoteEdited() {
  renderSubtaskKanban();
}

/* ============================================================
   ① タスク追加モーダル
   ------------------------------------------------------------
   ・wbsシートの名前定義「タスク範囲」内の選択行に行挿入して追加
   ・T〜FY列の数式は隣接行からコピーして埋める
   ・値の書込み: A=大分類, B=小分類, E=タスク名, N=担当者,
     P=予定開始日, Q=予定終了日
   ・大分類「受注」の場合、小分類は営業報告シートの
     状態=受注/受託中 の案件番号から選択
   ・大分類/担当者/既存小分類の候補は、呼び出し側の変数に頼らず
     このファイルがwbsシートから直接読み込む（＝営業報告からも
     そのまま呼び出せる）
   ============================================================ */

/* wbsシートから大分類・担当者・大分類ごとの小分類候補を集計 */
async function loadWbsMeta() {
  const meta = { categories: [], users: [], subcatsByCategory: {} };
  if (!window.Excel) return meta;
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
      const used = sheet.getUsedRange(true);
      used.load(["rowIndex", "rowCount"]);
      await ctx.sync();

      const lastRow = Math.max(used.rowIndex + used.rowCount, 11);
      const range = sheet.getRangeByIndexes(0, 0, lastRow, 26); // A1:Z{lastRow}
      range.load("values");
      await ctx.sync();

      const catSet = new Set();
      const userSet = new Set();
      const subMap = {};

      const { dataIdx } = findWbsHeader(range.values);
      range.values.slice(dataIdx).forEach((row) => {
        if (!row[25] || row[19] === "-") return; // Z空 or T="-" は除外

        const cat = row[0], sub = row[1], user = row[13];
        if (cat && cat !== "#") {
          catSet.add(cat);
          if (sub && String(sub).trim() !== "" && sub !== "#") {
            if (!subMap[cat]) subMap[cat] = new Set();
            subMap[cat].add(sub);
          }
        }
        if (user && user !== "#") userSet.add(user);
      });

      meta.categories = [...catSet];
      meta.users = [...userSet];
      Object.keys(subMap).forEach(c => { meta.subcatsByCategory[c] = [...subMap[c]]; });
    });
  } catch (e) {
    console.warn("wbsシートの候補読込に失敗:", e);
  }
  return meta;
}

let taskAddMeta = { categories: [], users: [], subcatsByCategory: {} };
/* 大分類・小分類を固定表示にする場合の指定（null なら通常の編集可能モード） */
let taskAddLock = null;

/* preset（任意）: { category, caseId }
   … 呼び出し元（例：営業報告の案件編集画面）から、大分類と
     案件番号（＝小分類、大分類が「受注」の場合の候補）を
     あらかじめ選択した状態でモーダルを開きたいときに使う。
     例）openTaskAdd({ category: "受注", caseId: "AG-03" }) */
async function openTaskAdd(preset) {
  preset = preset || {};
  taskAddMeta = await loadWbsMeta();

  // ロック指定（営業報告の案件編集画面から呼ばれた場合）
  taskAddLock = preset.lock
    ? { category: String(preset.category ?? ""), subCategory: String(preset.subCategory ?? preset.caseId ?? ""),
        reason: preset.reason || "" }
    : null;

  const catSel = document.getElementById("ta-cat");
  const txt = document.getElementById("ta-subcat");
  const sel = document.getElementById("ta-subcat-sel");
  const catHint = document.getElementById("ta-cat-hint");
  const subHint = document.getElementById("ta-subcat-hint");

  // 担当者: wbs既存の担当者
  const userSel = document.getElementById("ta-user");
  userSel.innerHTML = `<option value=""></option>` + taskAddMeta.users.map(u => `<option>${escapeHtml(String(u))}</option>`).join("");

  // 入力初期化
  document.getElementById("ta-title").value = "";
  document.getElementById("ta-start").value = "";
  document.getElementById("ta-end").value = "";
  const msg = document.getElementById("ta-msg");
  msg.className = "task-msg"; msg.textContent = "";

  if (taskAddLock) {
    /* 大分類・小分類は呼び出し元が決定済み。誤って別案件のタスクを
       作らないよう、値を固定して編集不可にする。 */
    catSel.innerHTML = `<option>${escapeHtml(taskAddLock.category)}</option>`;
    catSel.value = taskAddLock.category;
    catSel.disabled = true;
    catSel.classList.add("ta-locked");

    sel.style.display = "none";
    txt.style.display = "";
    txt.value = taskAddLock.subCategory;
    txt.readOnly = true;
    txt.classList.add("ta-locked");

    if (catHint) catHint.textContent = taskAddLock.reason || "";
    if (subHint) subHint.textContent = "案件番号を自動設定";
  } else {
    // 大分類: wbs既存の大分類 ＋ 受注（無ければ追加）
    const cats = [...taskAddMeta.categories];
    if (!cats.includes(cfg().orderCategory)) cats.push(cfg().orderCategory);
    catSel.innerHTML = cats.map(c => `<option>${escapeHtml(String(c))}</option>`).join("");
    if (preset.category && cats.includes(preset.category)) catSel.value = preset.category;
    catSel.disabled = false;
    catSel.classList.remove("ta-locked");

    txt.value = "";
    txt.readOnly = false;
    txt.classList.remove("ta-locked");

    if (catHint) catHint.textContent = "";
    if (subHint) subHint.textContent = "";

    await onTaCatChange();

    // 受注案件番号の事前選択（該当する候補があれば）
    if (preset.caseId && catSel.value === cfg().orderCategory) {
      if ([...sel.options].some(o => o.value === preset.caseId)) sel.value = preset.caseId;
    }
  }

  // wbsシートが表示されていない場合はアクティブにする
  activateWbs();

  document.getElementById("task-modal").classList.remove("wapi-hidden");
}
function closeTaskAdd() { document.getElementById("task-modal").classList.add("wapi-hidden"); }

async function activateWbs() {
  if (!window.Excel) return;
  try {
    await Excel.run(async ctx => {
      const active = ctx.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await ctx.sync();
      if (active.name !== cfg().wbsSheet) {
        ctx.workbook.worksheets.getItem(cfg().wbsSheet).activate();
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
  if (cat === cfg().orderCategory) {
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
    const subs = taskAddMeta.subcatsByCategory[cat] || [];
    dl.innerHTML = subs.map(s => `<option value="${escapeHtml(String(s))}"></option>`).join("");
  }
}

/* 営業報告シートから 状態=受注/受託中 の案件番号を取得 */
async function loadOrderCaseIds() {
  if (!window.Excel) return [];
  try {
    let out = [];
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().eigyoSheet);
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

  const cat = taskAddLock ? taskAddLock.category : document.getElementById("ta-cat").value;
  const sub = taskAddLock
    ? taskAddLock.subCategory
    : ((cat === cfg().orderCategory)
        ? document.getElementById("ta-subcat-sel").value
        : document.getElementById("ta-subcat").value.trim());
  const title = document.getElementById("ta-title").value.trim();
  const user = document.getElementById("ta-user").value;
  const start = document.getElementById("ta-start").value;
  const end = document.getElementById("ta-end").value;

  if (!title) { msg.className = "task-msg err"; msg.textContent = "タスク名を入力してください"; return; }
  if (!taskAddLock && cat === cfg().orderCategory && !sub) { msg.className = "task-msg err"; msg.textContent = "案件番号を選択してください"; return; }
  if (!window.Excel) { msg.className = "task-msg err"; msg.textContent = "Excel環境でのみ追加できます"; return; }

  try {
    let inserted = -1;
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);

      // 選択セルの行
      const selected = ctx.workbook.getSelectedRange();
      selected.load(["rowIndex", "worksheet/name"]);

      // 名前定義「タスク範囲」（ブック→wbsシートの順で検索）
      let nameItem = ctx.workbook.names.getItemOrNullObject(cfg().taskRangeName);
      let sheetNameItem = sheet.names.getItemOrNullObject(cfg().taskRangeName);
      await ctx.sync();
      if (nameItem.isNullObject && sheetNameItem.isNullObject) {
        throw new Error(`名前定義「${cfg().taskRangeName}」が見つかりません。wbsシートに行挿入可能な範囲を「${cfg().taskRangeName}」として名前定義してください。`);
      }
      const rangeObj = (!nameItem.isNullObject ? nameItem : sheetNameItem).getRange();
      rangeObj.load(["rowIndex", "rowCount", "worksheet/name"]);
      await ctx.sync();

      if (selected.worksheet.name !== cfg().wbsSheet) {
        throw new Error(`${cfg().wbsSheet}シート上で挿入したい行を選択してください。`);
      }
      const selRow = selected.rowIndex + 1;             // 1-based
      const rangeTop = rangeObj.rowIndex + 1;
      const rangeBottom = rangeObj.rowIndex + rangeObj.rowCount;
      if (selRow < rangeTop || selRow > rangeBottom) {
        throw new Error(`選択行（${selRow}行目）は「${cfg().taskRangeName}」（${rangeTop}〜${rangeBottom}行目）の外です。範囲内の行を選択してください。`);
      }

      // 行挿入（選択行の位置に。既存行は下へ）
      sheet.getRange(`${selRow}:${selRow}`).insert(Excel.InsertShiftDirection.down);
      await ctx.sync();

      // T〜FY列の数式を隣接行からコピー（挿入行の上、先頭行の場合は下からコピー）
      const srcRow = (selRow > rangeTop) ? selRow - 1 : selRow + 1;
      const dst = sheet.getRange(`T${selRow}:FY${selRow}`);
      dst.copyFrom(sheet.getRange(`T${srcRow}:FY${srcRow}`), Excel.RangeCopyType.formulas);

      /* U〜Wは判定式ではなく更新管理用に使うため、数式コピーで入った値を消す。
         数式コピー自体（T〜FY）は今週タスク・遅延判定のため残す必要がある。 */
      sheet.getRange(`${WBS_CLEAR_COLS.split(":")[0]}${selRow}:${WBS_CLEAR_COLS.split(":")[1]}${selRow}`)
           .clear(Excel.ClearApplyTo.contents);

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
      stampWbsUpdate(sheet, selRow);
      await ctx.sync();
      inserted = selRow;
    });

    closeTaskAdd();
    await apiAlert(`${inserted}行目にタスクを追加しました。`);
    if (typeof cfg().onTaskAdded === "function") cfg().onTaskAdded();
  } catch (e) {
    msg.className = "task-msg err";
    msg.textContent = e.message || "タスクの追加に失敗しました";
  }
}

/* ============================================================
   ⑤ ミニカンバン（案件配下タスクの3レーン表示）
   ------------------------------------------------------------
   ・wbsシートを直接読み込み、matchFnに一致するタスクだけを
     未着手／対応中／完了の3レーンで表示する
   ・ドラッグ＆ドロップで実績開始日・実績完了日をwbsへ書き込む
     （保留レーンは扱わない簡易版。カンバン本体の updateStatus と
     同じ考え方で R列/S列を更新する）
   ・カード左クリック：Excelの該当行へジャンプ
     カード右クリック：タスク詳細/備考編集モーダル（openModal）
   ・営業報告・カンバンどちらの画面からも呼び出せる
   ============================================================ */

/* wbsシートの該当行を選択して見える位置までスクロールする。
   シートのアクティブ化と選択は別の sync に分ける必要がある
   （同一バッチだと、非アクティブなシートに対する select が
     無視されて行が選択されないことがある）。 */
async function jumpToWbsRow(row) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
    sheet.activate();
    await ctx.sync();

    const target = sheet.getRange(`A${row}:Z${row}`);
    target.select();
    await ctx.sync();
  });
}

function miniExcelDateToJS(value) {
  if (!value) return null;
  if (typeof value === "number") return new Date((value - 25569) * 86400 * 1000);
  return new Date(value);
}
function miniFmt(v) {
  const d = miniExcelDateToJS(v);
  if (!d || isNaN(d)) return "";
  return `${d.getMonth() + 1}/${d.getDate()}`;
}
function miniStatus(t) {
  if (t.actualEnd) return "done";
  if (t.actualStart) return "doing";
  return "todo";
}

/* 案件番号（小分類=B列）が一致するかを見る、renderMiniKanban 用の絞り込みヘルパー */
function matchByCaseId(caseId) {
  const target = String(caseId ?? "").trim();
  return (t) => String(t.classification ?? "").trim() === target;
}

/* wbsシートを読み込み、matchFnに一致する行だけをタスクとして返す */
async function fetchWbsTasks(matchFn) {
  const out = [];
  if (!window.Excel) return out;
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
      const used = sheet.getUsedRange(true);
      used.load(["rowIndex", "rowCount"]);
      await ctx.sync();

      const lastRow = Math.max(used.rowIndex + used.rowCount, 11);
      const range = sheet.getRangeByIndexes(0, 0, lastRow, 26); // A1:Z{lastRow}
      range.load("values");
      await ctx.sync();

      const { dataIdx } = findWbsHeader(range.values);
      range.values.slice(dataIdx).forEach((row, i) => {
        if (!row[25] || row[19] === "-") return; // Z空 or T="-" は除外

        const t = {
          id: row[24],
          category: row[0],
          classification: row[1],
          title: row[25],
          user: row[13],
          note: row[14],
          start: row[15],
          end: row[16],
          actualStart: row[17],
          actualEnd: row[18],
          updatedAt: row[20],          // U列=更新日付
          rowIndex: dataIdx + i + 1,
          isNoSchedule: !row[15] && !row[16]
        };
        if (!matchFn || matchFn(t)) out.push(t);
      });
    });
  } catch (e) {
    console.warn("wbsタスクの取得に失敗:", e);
  }
  return out;
}

const MINI_LANES = [
  { key: "todo", label: "未着手" },
  { key: "doing", label: "対応中" },
  { key: "done", label: "完了" }
];

/* container: DOM要素 または id文字列
   matchFn:   task => boolean（wbsの行から作ったタスクオブジェクトを判定）
   opts.onChanged: ドラッグでステータスが変わり再描画された後に呼ばれる（任意）
   戻り値: 表示したタスク配列（呼び出し側で件数バッジ表示等に使える） */
async function renderMiniKanban(container, matchFn, opts) {
  const el = typeof container === "string" ? document.getElementById(container) : container;
  if (!el) return [];
  opts = opts || {};
  el.__mkMatchFn = matchFn;
  el.__mkOpts = opts;

  const tasks = await fetchWbsTasks(matchFn);

  /* 大分類フィルタ。分類が1種類しかない案件でも、機能の存在が分かるよう
     常にチップ行を出す（タスクが0件のときだけ隠す）。 */
  const cats = [...new Set(tasks.map(t => String(t.category ?? "").trim()).filter(Boolean))];
  if (el.__mkCat && !cats.includes(el.__mkCat)) el.__mkCat = "";   // 消えた分類の選択は解除
  const selCat = el.__mkCat || "";
  const shown = selCat ? tasks.filter(t => String(t.category ?? "").trim() === selCat) : tasks;

  const filterHtml = cats.length >= 1
    ? `<div class="mk-filter">
         <span class="lbl">大分類</span>
         <button class="mk-fchip ${selCat ? "" : "on"}" data-cat="">すべて ${tasks.length}</button>
         ${cats.map(c => {
           const n = tasks.filter(t => String(t.category ?? "").trim() === c).length;
           return `<button class="mk-fchip ${selCat === c ? "on" : ""}" data-cat="${escapeHtml(c)}">${escapeHtml(c)} ${n}</button>`;
         }).join("")}
       </div>`
    : "";

  const lanesHtml = MINI_LANES.map(L => {
    const cards = shown.filter(t => miniStatus(t) === L.key);
    return `
      <div class="mk-lane" data-lane="${L.key}">
        <div class="mk-lane-head">${escapeHtml(L.label)} <span>${cards.length}</span></div>
        <div class="mk-lane-body" data-lane="${L.key}">
          ${cards.map(t => {
            const overdue = L.key !== "done" && t.end &&
              miniExcelDateToJS(t.end) < new Date(new Date().toDateString());
            const cat = String(t.category ?? "").trim();
            // 分類が複数あるときだけカードにも分類タグを出す
            const catTag = (cats.length >= 2 && cat) ? `<span class="cat">${escapeHtml(cat)}</span>` : "";
            return `<div class="mk-card ${L.key}${overdue ? " overdue" : ""}" draggable="true" data-row="${t.rowIndex}">
              ${catTag}${escapeHtml(t.title || "（無題）")}
              <span class="due">${t.isNoSchedule ? "TODO" : miniFmt(t.end)}${t.user ? "・" + escapeHtml(t.user) : ""}${subtaskBadgeHtml(t.note, { parentDone: !!t.actualEnd })}</span>
            </div>`;
          }).join("")}
        </div>
      </div>`;
  }).join("");

  el.innerHTML = filterHtml + `<div class="mk-board">${lanesHtml}</div>`;

  el.querySelectorAll(".mk-fchip").forEach(b => {
    b.addEventListener("click", () => {
      el.__mkCat = b.dataset.cat || "";
      renderMiniKanban(el, el.__mkMatchFn, el.__mkOpts);
    });
  });

  bindMiniKanbanEvents(el, shown);
  return tasks;   // 件数バッジ用は常に全件を返す（フィルタの影響を受けない）
}

function bindMiniKanbanEvents(el, tasks) {
  let dragRow = null;

  el.querySelectorAll(".mk-card").forEach(card => {
    card.addEventListener("dragstart", (e) => {
      dragRow = Number(card.dataset.row);
      card.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });
    card.addEventListener("dragend", () => card.classList.remove("dragging"));

    card.addEventListener("click", () => {
      jumpToWbsRow(Number(card.dataset.row)).catch(err => console.log("jump error:", err));
    });

    card.addEventListener("contextmenu", async (e) => {
      e.preventDefault();
      const t = tasks.find(x => x.rowIndex === Number(card.dataset.row));
      if (t) await openModal(t);
    });
  });

  el.querySelectorAll(".mk-lane-body").forEach(body => {
    body.addEventListener("dragover", (e) => { e.preventDefault(); body.classList.add("over"); });
    body.addEventListener("dragleave", () => body.classList.remove("over"));
    body.addEventListener("drop", async (e) => {
      e.preventDefault();
      body.classList.remove("over");
      if (dragRow == null) return;
      const t = tasks.find(x => x.rowIndex === dragRow);
      dragRow = null;
      if (!t || miniStatus(t) === body.dataset.lane) return;
      await moveMiniKanbanTask(t, body.dataset.lane, el);
    });
  });
}

/* ドラッグによるステータス変更：実績開始日・実績完了日をwbsのR/S列へ書き込む */
async function moveMiniKanbanTask(task, lane, el) {
  let actualStart = task.actualStart;
  let actualEnd = task.actualEnd;

  if (lane === "todo") {
    actualStart = "";
    actualEnd = "";
  } else if (lane === "doing") {
    if (!(actualStart instanceof Date) || isNaN(actualStart)) actualStart = new Date();
    actualEnd = "";
  } else if (lane === "done") {
    if (!(actualStart instanceof Date) || isNaN(actualStart)) actualStart = new Date();
    actualEnd = new Date();
  }

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
      const row = task.rowIndex;

      const startCell = sheet.getRange(`R${row}`);
      const endCell = sheet.getRange(`S${row}`);
      startCell.values = [[dateToExcelSerial(actualStart)]];
      endCell.values = [[dateToExcelSerial(actualEnd)]];
      startCell.numberFormat = [["m/d"]];
      endCell.numberFormat = [["m/d"]];

      stampWbsUpdate(sheet, row);
      await ctx.sync();
    });
  } catch (e) {
    console.warn("ミニカンバンのステータス更新に失敗:", e);
    await apiAlert("ステータスの更新に失敗しました。もう一度お試しください。");
  }

  await renderMiniKanban(el, el.__mkMatchFn, el.__mkOpts);
  if (el.__mkOpts && typeof el.__mkOpts.onChanged === "function") el.__mkOpts.onChanged();
}

/* ============================================================
   モーダルDOMの自動注入
   ------------------------------------------------------------
   #modal / #task-modal / #dialog-modal のうち、呼び出し側の
   index.htmlにまだ存在しないものだけを個別に注入する。
   （例：営業報告アドインは独自の #dialog-modal を既に持っているため、
    その分だけスキップし、重複IDを作らないようにする）
   ============================================================ */
function ensureApiDom() {
  const pieces = [];

  if (!document.getElementById("modal")) {
    pieces.push(`
    <div id="modal" class="wapi-modal wapi-hidden">
      <div class="wapi-modal-content">
        <h3 id="modal-title"></h3>
        <textarea id="modal-note" oninput="onModalNoteEdited()"></textarea>
        <div class="subtask-section">
          <div class="subtask-label">サブタスク（□未着手 ◎対応中 ■完了）</div>
          <div id="subtask-kanban"></div>
        </div>
        <div class="wapi-modal-actions">
          <button class="wapi-btn-primary" onclick="saveNote()">保存</button>
          <button class="wapi-btn-ghost" onclick="closeModal()">閉じる</button>
          <small>ESC: 閉じる</small>
        </div>
      </div>
    </div>`);
  }

  if (!document.getElementById("task-modal")) {
    pieces.push(`
    <div id="task-modal" class="wapi-modal wapi-hidden">
      <div class="wapi-modal-content wapi-task-modal-content">
        <h3>タスク追加</h3>
        <div class="task-hint" id="task-hint">追加したい行を選択してください。（wbsシート上で挿入位置の行をクリック）</div>
        <div class="task-grid">
          <div class="t-row"><label>大分類</label><select id="ta-cat" onchange="onTaCatChange()"></select>
            <span class="ta-hint" id="ta-cat-hint"></span>
          </div>
          <div class="t-row"><label>小分類</label>
            <input type="text" id="ta-subcat" list="ta-subcat-list" placeholder="既存候補から選択 or 新規入力">
            <datalist id="ta-subcat-list"></datalist>
            <select id="ta-subcat-sel" style="display:none"></select>
            <span class="ta-hint" id="ta-subcat-hint"></span>
          </div>
          <div class="t-row wide"><label>タスク名 <span class="req">必須</span></label><input type="text" id="ta-title"></div>
          <div class="t-row"><label>担当者</label><select id="ta-user"></select></div>
          <div class="t-row"><label>予定開始日</label><input type="date" id="ta-start"></div>
          <div class="t-row"><label>予定終了日</label><input type="date" id="ta-end"></div>
        </div>
        <div class="wapi-modal-actions">
          <button class="wapi-btn-primary" onclick="saveTaskAdd()">OK（選択行に挿入）</button>
          <button class="wapi-btn-ghost" onclick="closeTaskAdd()">キャンセル</button>
          <span class="task-msg" id="ta-msg"></span>
        </div>
      </div>
    </div>`);
  }

  // dialog-modal は呼び出し側アプリが独自の確認ダイアログを
  // 既に持っていることが多いため（例：営業報告）、無い場合のみ注入する。
  if (!document.getElementById("dialog-modal")) {
    pieces.push(`
    <div id="dialog-modal" class="wapi-modal wapi-hidden">
      <div class="wapi-modal-content wapi-dialog-content">
        <div id="dialog-msg" class="wapi-dialog-msg"></div>
        <div class="wapi-modal-actions">
          <button class="wapi-btn-ghost" id="dialog-cancel" onclick="dialogRespond(false)">キャンセル</button>
          <button class="wapi-btn-primary" id="dialog-ok" onclick="dialogRespond(true)">OK</button>
        </div>
      </div>
    </div>`);
  }

  if (!pieces.length) return;

  const wrap = document.createElement("div");
  wrap.innerHTML = pieces.join("");
  while (wrap.firstElementChild) document.body.appendChild(wrap.firstElementChild);
}


/* ============================================================
   ③ サブタスク数（カード・ミニカンバン・スケジュール共通）
   ------------------------------------------------------------
   出典は備考（O列）の □未着手 ◎対応中 ■完了。
   parseSubtasks() を使うので、どの画面でも数え方が一致する。
   ============================================================ */
function subtaskCount(note) {
  const list = parseSubtasks(note);
  const c = { total: list.length, todo: 0, doing: 0, done: 0 };
  list.forEach((t) => { if (c[t.lane] !== undefined) c[t.lane]++; });
  return c;
}

/* opts.showZero    … 0件でもグレーで出す（営業報告の一覧など）
   opts.parentDone  … 親タスクが完了しているか（残っていればオレンジで警告） */
function subtaskBadgeHtml(note, opts) {
  const o = opts || {};
  const c = subtaskCount(note);
  if (!c.total) {
    return o.showZero
      ? '<span class="wapi-sub none" title="サブタスク未登録">0/0</span>'
      : "";
  }
  const all = c.done === c.total;
  const left = !!o.parentDone && !all;
  const cls = left ? " left" : (all ? " all-done" : "");
  const tip = left
    ? `親タスクは完了だが、サブタスクが ${c.total - c.done} 件残っている`
    : `サブタスク 完了${c.done} / 全${c.total}件`;
  return `<span class="wapi-sub${cls}" title="${escapeHtml(tip)}">■${c.done}/${c.total}</span>`;
}

/* ============================================================
   ④ 休み予定（wbs シート上部の 担当者 × 日 マトリクス）
   ------------------------------------------------------------
   ・氏名   … N列（firstRow 〜 ヘッダー行-2、空行はスキップ）
   ・日付   … AB〜FY列。1列＝1日。ヘッダー行+1 に「日」が入る
   ・月     … ヘッダー行に月と「稼働日：n」
   ・種別   … 2行目 P〜AA のヘッダー。値のある列だけ採用
   ・判定   … シートの COUNTIF と同じ前方一致（全角半角は正規化）
   ============================================================ */
function colToIdx(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function idxToCol(idx) {
  let s = "", n = idx + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function toHalfWidth(v) {
  return String(v ?? "").replace(/[Ａ-Ｚａ-ｚ０-９]/g,
    (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).trim();
}
function isoOf(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0")
       + "-" + String(d.getDate()).padStart(2, "0");
}
/* 種別文字列 → { style, work, mark }。先頭1文字の前方一致。 */
function leaveStyleOf(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const map = cfg().leaveStyles || {};
  const keys = [toHalfWidth(raw).charAt(0).toUpperCase(), raw.charAt(0)];
  for (const k of keys) if (k && map[k]) return Object.assign({ label: raw }, map[k]);
  const fb = map["*"] || { style: "full", work: 0 };
  return Object.assign({ label: raw, mark: raw.charAt(0) }, fb);
}

let leaveGridCache = null;

async function loadLeaveGrid(force) {
  if (leaveGridCache && !force) return leaveGridCache;
  const g = cfg().leaveGrid;
  const empty = { users: [], types: [], dates: [], grid: {}, workdays: {}, ok: false, warn: "" };
  if (!window.Excel) return (leaveGridCache = empty);

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
      const c1 = colToIdx(g.dateCol1), c2 = colToIdx(g.dateColEnd);
      const block = sheet.getRangeByIndexes(0, 0, g.scanRows, c2 + 1);   // A1:FY{scanRows}
      const base = sheet.getRange(g.baseDateCell);
      block.load("values");
      base.load("values");
      await ctx.sync();

      const v = block.values;
      const hIdx = findWbsHeader(v).headerRow - 1;      // 0基点
      const dayRow = v[hIdx + 1] || [];
      const monthRow = v[hIdx] || [];

      /* --- 基準日 --- */
      let baseDate = miniExcelDateToJS(base.values[0][0]);
      let warn = "";

      /* --- 日付列 → 日付 --- */
      const dates = [];
      if (baseDate) {
        for (let c = c1; c <= c2; c++) {
          const d = new Date(baseDate.getTime() + (c - c1) * 86400000);
          d.setHours(0, 0, 0, 0);
          dates.push({ col: c, date: d, iso: isoOf(d) });
        }
        // 検算：シートの「日」と一致するか（値がある列だけ見る）
        let hit = 0, miss = 0;
        dates.forEach((x) => {
          const n = parseInt(toHalfWidth(dayRow[x.col]), 10);
          if (!isNaN(n)) { if (n === x.date.getDate()) hit++; else miss++; }
        });
        if (hit === 0 || miss > hit) {
          warn = `基準日セル（${g.baseDateCell}）から求めた日付が ${g.dateCol1}${hIdx + 2} 行の「日」と一致しません。`
               + `ApiConfig.leaveGrid.baseDateCell を実際の番地に合わせてください。`;
        }
      } else {
        warn = `基準日セル（${g.baseDateCell}）が日付として読めませんでした。`;
      }

      /* --- 種別ヘッダー（2行目 P〜AA） --- */
      const types = [];
      const tRow = v[g.typeRow - 1] || [];
      for (let c = colToIdx(g.typeCol1); c <= colToIdx(g.typeColEnd); c++) {
        const t = String(tRow[c] ?? "").trim();
        if (t) types.push(t);
      }

      /* --- 稼働日（ヘッダー行の「稼働日：n」） --- */
      const workdays = {};
      monthRow.forEach((cell, c) => {
        const m = /稼働日[：:]\s*(\d+)/.exec(String(cell ?? ""));
        if (m && c >= c1 && c <= c2) {
          const d = dates.find((x) => x.col === c);
          if (d) workdays[d.date.getFullYear() + "-" + String(d.date.getMonth() + 1).padStart(2, "0")] = +m[1];
        }
      });

      /* --- 担当者と値 --- */
      const nameIdx = colToIdx(g.nameCol);
      const users = [], grid = {};
      for (let r = g.firstRow - 1; r <= hIdx - 2; r++) {
        const nm = String((v[r] || [])[nameIdx] ?? "").trim();
        if (!nm) continue;                              // 空行はスキップ
        users.push({ name: nm, row: r + 1 });
        grid[nm] = {};
        dates.forEach((x) => {
          const cell = (v[r] || [])[x.col];
          const s = String(cell ?? "").trim();
          if (s) grid[nm][x.iso] = s;
        });
      }

      leaveGridCache = { users, types, dates, grid, workdays, headerRow: hIdx + 1, ok: !warn, warn };
    });
  } catch (e) {
    console.warn("休み予定の読み取りに失敗:", e);
    leaveGridCache = Object.assign({}, empty, { warn: String(e && e.message || e) });
  }
  return leaveGridCache;
}

/* 他機能（スケジュールの帯など）から参照する用。要 loadLeaveGrid() 済み */
function leavesOf(user, from, to) {
  const g = leaveGridCache;
  if (!g || !g.grid[user]) return [];
  const a = from ? isoOf(from) : "0000-00-00";
  const b = to ? isoOf(to) : "9999-99-99";
  const out = [];
  Object.keys(g.grid[user]).forEach((iso) => {
    if (iso < a || iso > b) return;
    const st = leaveStyleOf(g.grid[user][iso]);
    if (st) out.push({ iso, date: new Date(iso + "T00:00:00"), value: g.grid[user][iso], style: st });
  });
  return out.sort((x, y) => (x.iso < y.iso ? -1 : 1));
}

/* 1行ぶんのセルに種別を書く。isoList は "YYYY-MM-DD" の配列。type:"" でクリア */
async function setLeave(user, isoList, type) {
  const g = await loadLeaveGrid();
  const u = g.users.find((x) => x.name === user);
  if (!u || !isoList.length) return false;
  const cols = isoList
    .map((iso) => g.dates.find((d) => d.iso === iso))
    .filter(Boolean)
    .map((d) => d.col)
    .sort((a, b) => a - b);
  if (!cols.length) return false;

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
    // 連続する列はまとめて1回で書く
    let s = 0;
    while (s < cols.length) {
      let e = s;
      while (e + 1 < cols.length && cols[e + 1] === cols[e] + 1) e++;
      const n = e - s + 1;
      const addr = `${idxToCol(cols[s])}${u.row}:${idxToCol(cols[e])}${u.row}`;
      sheet.getRange(addr).values = [new Array(n).fill(type)];
      s = e + 1;
    }
    await ctx.sync();
  });
  leaveGridCache = null;                                // 次回は読み直す
  return true;
}

/* ===== 担当者 × 日 マトリクス ===== */
const leaveView = { y: null, m: null, me: null, drag: null };

async function renderLeaveMatrix(container, opts) {
  const host = typeof container === "string" ? document.getElementById(container) : container;
  if (!host) return;
  const o = opts || {};
  const g = await loadLeaveGrid();

  if (leaveView.y == null) {
    const now = new Date();
    leaveView.y = now.getFullYear(); leaveView.m = now.getMonth();
  }
  if (o.me) leaveView.me = o.me;
  if (leaveView.me == null) {
    try { leaveView.me = localStorage.getItem("wbs-leave-me") || null; } catch (e) { /* 無視 */ }
  }
  // 覚えた名前が担当者一覧に無ければ捨てる（氏名変更・行削除に備える）
  if (leaveView.me && !g.users.some((u) => u.name === leaveView.me)) leaveView.me = null;

  const y = leaveView.y, m = leaveView.m;
  const days = new Date(y, m + 1, 0).getDate();
  const WD = ["日", "月", "火", "水", "木", "金", "土"];
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const monthKey = y + "-" + String(m + 1).padStart(2, "0");

  const inSheet = {};
  g.dates.forEach((d) => { if (d.date.getFullYear() === y && d.date.getMonth() === m) inSheet[d.iso] = true; });

  // 誰の行を編集するか決まっていないときは、選ぶまで全行を読み取り専用にする
  const noMe = !leaveView.me;

  let head = `<tr><th class="wlv-name">${y}/${String(m + 1).padStart(2, "0")}</th>`;
  for (let d = 1; d <= days; d++) {
    const dw = new Date(y, m, d).getDay();
    head += `<th class="${dw === 0 ? "sun" : dw === 6 ? "sat" : ""}">${d}<br>${WD[dw]}</th>`;
  }
  g.types.forEach((t, i) => { head += `<th class="wlv-sum${i === 0 ? " wlv-div" : ""}">${escapeHtml(t)}</th>`; });
  head += '<th class="wlv-work wlv-div">稼働</th></tr>';

  let body = "";
  g.users.forEach((u) => {
    const mine = !noMe && u.name === leaveView.me;
    let tr = `<tr class="${mine ? "wlv-me" : ""}" data-user="${escapeHtml(u.name)}">`
           + `<td class="wlv-name">${escapeHtml(u.name)}</td>`;
    const counts = {}; g.types.forEach((t) => { counts[t] = 0; });
    let lost = 0;

    for (let d = 1; d <= days; d++) {
      const dt = new Date(y, m, d);
      const iso = isoOf(dt);
      const dw = dt.getDay();
      const val = (g.grid[u.name] || {})[iso];
      const st = val ? leaveStyleOf(val) : null;
      const cls = [];
      if (dw === 0 || dw === 6) cls.push("wknd");
      if (iso === isoOf(today)) cls.push("wlv-today");
      if (!inSheet[iso]) cls.push("wlv-nocol");
      const editable = mine && inSheet[iso] && !o.readOnly;   // 自分の行だけ編集可
      if (editable) cls.push("wlv-edit");

      if (st) {
        g.types.forEach((t) => {
          if (toHalfWidth(val).charAt(0).toUpperCase() === toHalfWidth(t).charAt(0).toUpperCase()) counts[t]++;
        });
        if (dw !== 0 && dw !== 6) lost += (1 - (st.work || 0));
      }
      tr += `<td class="${cls.join(" ")}" data-iso="${iso}">`
          + (st ? `<span class="wlv-c ${st.style}" title="${escapeHtml(st.label)}">${escapeHtml(st.mark || st.label.charAt(0))}</span>` : "")
          + "</td>";
    }
    g.types.forEach((t, i) => {
      tr += `<td class="wlv-sum${counts[t] ? "" : " z"}${i === 0 ? " wlv-div" : ""}">${counts[t]}</td>`;
    });
    const wd = g.workdays[monthKey];
    const work = wd == null ? "" : String(Math.round((wd - lost) * 10) / 10);
    body += tr + `<td class="wlv-work wlv-div">${work}</td></tr>`;
  });

  host.innerHTML =
    `<div class="wlv-bar">
       <button class="wlv-nav" data-mv="-1">◀</button>
       <span class="wlv-mlabel">${y}/${String(m + 1).padStart(2, "0")}</span>
       <button class="wlv-nav" data-mv="1">▶</button>
       <select class="wlv-me">
         <option value=""${noMe ? " selected" : ""}>（自分を選択）</option>
         ${g.users.map((u) =>
           `<option${u.name === leaveView.me ? " selected" : ""}>${escapeHtml(u.name)}</option>`).join("")}
       </select>
       <span class="wlv-sp"></span>
       <span class="wlv-hint">${o.readOnly ? ""
         : (noMe ? "編集するには、まず自分の名前を選んでください（次回から覚えます）"
                 : "自分の行をクリック／横ドラッグで種別を選択")}</span>
     </div>
     ${g.warn ? `<div class="wlv-warn">${escapeHtml(g.warn)}</div>` : ""}
     <div class="wlv-wrap"><table class="wlv-mx">${"<thead>" + head + "</thead><tbody>" + body + "</tbody>"}</table></div>
     <div class="wlv-lg">
       <span><i class="full"></i>終日</span><span><i class="am"></i>午前休</span>
       <span><i class="pm"></i>午後休</span><span><i class="remote"></i>リモート（稼働あり）</span>
       <span><i class="away"></i>不在</span><span class="wlv-red">赤枠＝本日</span>
     </div>`;

  bindLeaveMatrix(host, o);
}

function bindLeaveMatrix(host, o) {
  host.querySelectorAll(".wlv-nav").forEach((b) => {
    b.addEventListener("click", () => {
      const d = new Date(leaveView.y, leaveView.m + Number(b.dataset.mv), 1);
      leaveView.y = d.getFullYear(); leaveView.m = d.getMonth();
      closeLeavePop(); renderLeaveMatrix(host, o);
    });
  });
  const sel = host.querySelector(".wlv-me");
  if (sel) sel.addEventListener("change", () => {
    leaveView.me = sel.value || null;
    try {
      if (leaveView.me) localStorage.setItem("wbs-leave-me", leaveView.me);
      else localStorage.removeItem("wbs-leave-me");
    } catch (e) { /* 保存できなくても継続 */ }
    closeLeavePop();
    renderLeaveMatrix(host, Object.assign({}, o, { me: undefined }));
  });
  if (o.readOnly) return;

  const tbl = host.querySelector(".wlv-mx");
  const wrap = host.querySelector(".wlv-wrap");
  if (!tbl || !wrap) return;

  const paint = () => {
    tbl.querySelectorAll("td.wlv-range").forEach((td) => td.classList.remove("wlv-range"));
    leaveRangeCells(tbl).forEach((td) => td.classList.add("wlv-range"));
  };
  tbl.addEventListener("mousedown", (e) => {
    const td = e.target.closest("td.wlv-edit");
    if (!td) return;
    e.preventDefault();
    leaveView.drag = { row: td.parentNode, a: td.dataset.iso, b: td.dataset.iso };
    paint();
  });
  tbl.addEventListener("mousemove", (e) => {
    if (!leaveView.drag) return;
    const td = e.target.closest("td.wlv-edit");
    if (!td || td.parentNode !== leaveView.drag.row) return;
    leaveView.drag.b = td.dataset.iso;
    paint();
  });
  document.addEventListener("mouseup", function onUp() {
    document.removeEventListener("mouseup", onUp);
    if (leaveView.drag) openLeavePop(host, wrap, tbl, o);
  });
}

function leaveRangeCells(tbl) {
  const d = leaveView.drag;
  if (!d) return [];
  const a = d.a < d.b ? d.a : d.b, b = d.a < d.b ? d.b : d.a;
  return Array.prototype.slice.call(d.row.querySelectorAll("td[data-iso]"))
    .filter((td) => td.dataset.iso >= a && td.dataset.iso <= b && td.classList.contains("wlv-edit"));
}
function closeLeavePop() {
  const p = document.getElementById("wlv-pop");
  if (p) p.remove();
  document.querySelectorAll("td.wlv-range").forEach((td) => td.classList.remove("wlv-range"));
  leaveView.drag = null;
}
function openLeavePop(host, wrap, tbl, o) {
  const tds = leaveRangeCells(tbl);
  if (!tds.length) { closeLeavePop(); return; }
  closeLeavePopKeepDrag();
  const g = leaveGridCache || { types: [] };
  const first = tds[0], last = tds[tds.length - 1];
  const wr = wrap.getBoundingClientRect(), r = first.getBoundingClientRect();
  const label = tds.length === 1
    ? first.dataset.iso.slice(5).replace("-", "/")
    : `${first.dataset.iso.slice(5).replace("-", "/")} 〜 ${last.dataset.iso.slice(5).replace("-", "/")}（${tds.length}日）`;

  wrap.insertAdjacentHTML("beforeend",
    `<div class="wlv-pop" id="wlv-pop" style="left:${r.left - wr.left + wrap.scrollLeft}px;top:${r.bottom - wr.top + 4}px">
       <div class="wlv-ph">${escapeHtml(leaveView.me)}　${label}</div>
       <div class="wlv-opts">
         ${g.types.map((t) => `<button data-t="${escapeHtml(t)}">${escapeHtml(t)}</button>`).join("")}
         <button class="clr" data-t="">クリア</button>
       </div>
     </div>`);
  const pop = document.getElementById("wlv-pop");
  if ((r.bottom - wr.top + 4) + pop.offsetHeight > wrap.clientHeight) {
    pop.style.top = Math.max(r.top - wr.top - pop.offsetHeight - 4, 2) + "px";
  }
  pop.querySelectorAll("button").forEach((b) => {
    b.addEventListener("click", async () => {
      const isos = tds.map((td) => td.dataset.iso);
      closeLeavePop();
      await setLeave(leaveView.me, isos, b.dataset.t);
      await renderLeaveMatrix(host, o);
      if (typeof cfg().onLeaveChanged === "function") cfg().onLeaveChanged();
    });
  });
}
function closeLeavePopKeepDrag() {
  const p = document.getElementById("wlv-pop");
  if (p) p.remove();
}

/* ===== 「個人予定」モーダル ===== */
async function openMyLeave(user) {
  ensureLeaveModal();
  const back = document.getElementById("wlv-modal");
  back.classList.remove("wapi-hidden");
  await renderLeaveMatrix("wlv-modal-body", { me: user || null });
}
function closeMyLeave() {
  const back = document.getElementById("wlv-modal");
  if (back) back.classList.add("wapi-hidden");
  closeLeavePop();
}
function ensureLeaveModal() {
  if (document.getElementById("wlv-modal")) return;
  const div = document.createElement("div");
  div.id = "wlv-modal";
  div.className = "wapi-modal wapi-hidden";
  div.innerHTML =
    `<div class="wapi-modal-content wlv-modal-content">
       <h3>個人予定<span class="wlv-src">wbs シート上部の 担当者 × 日 マトリクス</span></h3>
       <div id="wlv-modal-body"></div>
       <div class="wapi-modal-actions" style="margin-top:10px">
         <button class="wapi-btn-ghost" id="wlv-close">閉じる</button>
       </div>
     </div>`;
  document.body.appendChild(div);
  div.addEventListener("click", (e) => { if (e.target === div) closeMyLeave(); });
  document.getElementById("wlv-close").addEventListener("click", closeMyLeave);
}

/* ============================================================
   ① スケジュール共通部品
   ------------------------------------------------------------
   ・taskLineRowsHtml … タスク1件1行の「線」。営業報告の既存ガントに
     差し込む用と、下の renderSchedule 用で共有する。
   ・taskListHtml     … 展開時に出すリスト。
   geom = { t0, totalDays, monthLines, todayHtml, guide, labelCls }
   ============================================================ */
function schedPct(d, geom) { return ((d - geom.t0) / 86400000) / geom.totalDays * 100; }
function schedClamp(v) { return Math.max(-8, Math.min(108, v)); }
function schedMd(d) { return (d.getMonth() + 1) + "/" + d.getDate(); }

/* 遅延は「対応中(遅延)」と「未着手(遅延)」に分ける。
   同じ遅延でも手が付いているかどうかで打ち手が違うため、
   集計タブの5区分と表記・色をそろえる。 */
function taskStatusKey(t) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (t.actualEnd) return "done";
  if ((t.note || "").toString().includes("▲")) return "held";
  if (!t.start || !t.end) return "todo";

  const start = miniExcelDateToJS(t.start);
  const end = miniExcelDateToJS(t.end);
  const doing = !!t.actualStart || (start && start <= today);

  if (end && end < today) return doing ? "delaydoing" : "delaytodo";
  return doing ? "active" : "todo";
}
const TASK_STATUS_JA = {
  todo: "未着手", active: "対応中", held: "保留", done: "完了",
  delaydoing: "対応中(遅延)", delaytodo: "未着手(遅延)"
};
/* 遅延件数＝対応中(遅延) ＋ 未着手(遅延) */
function isDelayKey(k) { return k === "delaydoing" || k === "delaytodo"; }

function taskLineRowsHtml(tasks, geom) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let guideHtml = "";
  if (geom.guide && geom.guide.start && geom.guide.end) {
    const gl = schedClamp(schedPct(geom.guide.start, geom));
    const gr = schedClamp(schedPct(new Date(geom.guide.end.getTime() + 86400000), geom));
    guideHtml = `<div class="wsc-guide" style="left:${gl}%;width:${Math.max(gr - gl, 0.4)}%"></div>`;
  }

  return tasks.map((t) => {
    const st = taskStatusKey(t);
    const s = miniExcelDateToJS(t.start), e = miniExcelDateToJS(t.end);
    const pri = (t.priority == null ? "" : String(t.priority)).trim();
    let bar = "", act = "", tick = "", dl = "";

    if (s && e) {
      const l = schedClamp(schedPct(s, geom));
      const r = schedClamp(schedPct(new Date(e.getTime() + 86400000), geom));
      const w = Math.max(r - l, 0.4);
      const over = geom.dueLimit && e > geom.dueLimit;
      bar = `<div class="wsc-t ${st}${over ? " over" : ""}" data-row="${t.rowIndex}"
              style="left:${l}%;width:${w}%" title="${escapeHtml(t.title + "  " + schedMd(s) + "〜" + schedMd(e) + "  " + (t.user || ""))}">
              ${w > 9 ? `<span class="tt">${escapeHtml(t.title)}</span>` : ""}</div>`;
      if (over) tick = `<div class="wsc-over" style="left:${Math.min(r + 0.3, 99)}%">▸</div>`;
      dl = `<div class="wsc-d" style="left:${Math.min(r + (over ? 1.6 : 0.5), 97)}%">${schedMd(e)}</div>`;
      const as = miniExcelDateToJS(t.actualStart);
      if (geom.showActual !== false && as) {
        const ae = t.actualEnd ? miniExcelDateToJS(t.actualEnd) : today;
        const al = schedClamp(schedPct(as, geom));
        const ar = schedClamp(schedPct(new Date(ae.getTime() + 86400000), geom));
        act = `<div class="wsc-act${t.actualEnd ? " done" : ""}" style="left:${al}%;width:${Math.max(ar - al, 0.4)}%"></div>`;
      }
    } else {
      bar = '<div class="wsc-d" style="left:0.5%;top:8px">日付なし（TODO）</div>';
    }

    let leaveHtml = "";
    if (geom.showLeave && t.user) {
      leavesOf(t.user, geom.t0, new Date(geom.t0.getTime() + geom.totalDays * 86400000)).forEach((lv) => {
        if ((lv.style.work || 0) >= 1) return;                 // リモートは帯を出さない
        const f = schedClamp(schedPct(lv.date, geom));
        const to = schedClamp(schedPct(new Date(lv.date.getTime() + 86400000), geom));
        const k = lv.style.style === "away" ? " away" : ((lv.style.work || 0) > 0 ? " h" : "");
        leaveHtml += `<div class="wsc-leave${k}" title="${escapeHtml(lv.value)}" style="left:${f}%;width:${Math.max(to - f, 0.3)}%"></div>`;
      });
    }

    return `
    <div class="g-row wsc-taskrow" data-row="${t.rowIndex}">
      <div class="g-label wsc-tlabel">
        <div class="wsc-tname">${pri ? `<span class="wsc-pri">${escapeHtml(pri)}</span>` : ""}${escapeHtml(t.title)}<span class="wsc-st ${st}">${TASK_STATUS_JA[st]}</span></div>
        <div class="wsc-tmeta">${escapeHtml(t.user || "")}${geom.showCat && t.category ? `<span class="wsc-cattag">${escapeHtml(t.category)}</span>` : ""}${subtaskBadgeHtml(t.note, { parentDone: !!t.actualEnd })}</div>
      </div>
      <div class="g-track wsc-ttrack">
        ${geom.monthLines || ""}${guideHtml}${leaveHtml}${geom.todayHtml || ""}${bar}${act}${tick}${dl}
      </div>
    </div>`;
  }).join("");
}

function taskListHtml(tasks, opts) {
  const o = opts || {};
  // bare:true … ガントの外に置く用。.g-row/.g-track を使わない
  //   （ホスト側の .g-track は高さが固定されており、はみ出すと隠れてしまう）
  const open = o.bare ? "" : '<div class="g-row wsc-listrow"><div class="g-label"></div><div class="g-track">';
  const close = o.bare ? "" : "</div></div>";

  if (!tasks.length) {
    return `${open}<div class="wsc-list"><div class="wsc-empty">紐づく WBS タスクがありません。
      ${o.hint ? escapeHtml(o.hint) : ""}</div></div>${close}`;
  }
  const rows = tasks.map((t) => {
    const st = taskStatusKey(t);
    const s = miniExcelDateToJS(t.start), e = miniExcelDateToJS(t.end);
    const as = miniExcelDateToJS(t.actualStart), ae = miniExcelDateToJS(t.actualEnd);
    const pri = (t.priority == null ? "" : String(t.priority)).trim();
    const over = o.dueLimit && e && e > o.dueLimit;
    return `<tr data-row="${t.rowIndex}" title="クリックで Excel の ${t.rowIndex} 行へジャンプ">
      <td class="${pri ? "" : "dim"}">${pri ? `<span class="wsc-pri">${escapeHtml(pri)}</span>` : "−"}</td>
      <td class="nm">${escapeHtml(t.title)}${over ? ' <span class="wsc-ov">▸超過</span>' : ""}</td>
      ${o.showCat ? `<td class="dim">${escapeHtml(t.category || "")}</td>` : ""}
      <td>${escapeHtml(t.user || "")}</td>
      <td class="${s && e ? "" : "dim"}">${s && e ? schedMd(s) + "〜" + schedMd(e) : "日付なし"}</td>
      <td class="${as ? "" : "dim"}">${as ? schedMd(as) + "〜" + (ae ? schedMd(ae) : "") : "−"}</td>
      <td>${subtaskBadgeHtml(t.note, { parentDone: !!t.actualEnd, showZero: true })}</td>
      <td><span class="wsc-st ${st}">${TASK_STATUS_JA[st]}</span></td>
    </tr>`;
  }).join("");

  return `${open}<div class="wsc-list">
      <div class="wsc-lhead">${escapeHtml(o.title || "")}
        <span class="wsc-lmeta">${escapeHtml(o.meta || "")}</span></div>
      <table class="wsc-tbl"><thead><tr>
        <th>優先</th><th>タスク</th>${o.showCat ? "<th>大分類</th>" : ""}
        <th>担当</th><th>予定</th><th>実績</th><th>サブ</th><th>状態</th>
      </tr></thead><tbody>${rows}</tbody></table>
    </div>${close}`;
}

/* 展開部（線＋リスト）内のクリック・ホバーを配線。container 単位で1回呼ぶ */
function bindTaskLines(container) {
  const host = typeof container === "string" ? document.getElementById(container) : container;
  if (!host || host.dataset.wscBound === "1") return;
  host.dataset.wscBound = "1";
  host.addEventListener("click", (e) => {
    const tr = e.target.closest(".wsc-tbl tbody tr[data-row]");
    if (tr) { jumpToWbsRow(Number(tr.dataset.row)); return; }
    const bar = e.target.closest(".wsc-t[data-row]");
    if (bar) jumpToWbsRow(Number(bar.dataset.row));
  });
  host.addEventListener("mouseover", (e) => {
    const tr = e.target.closest(".wsc-tbl tbody tr[data-row]");
    if (!tr) return;
    const bar = host.querySelector(`.wsc-t[data-row="${tr.dataset.row}"]`);
    if (bar) bar.classList.add("hl");
  });
  host.addEventListener("mouseout", (e) => {
    const tr = e.target.closest(".wsc-tbl tbody tr[data-row]");
    if (!tr) return;
    const bar = host.querySelector(`.wsc-t[data-row="${tr.dataset.row}"]`);
    if (bar) bar.classList.remove("hl");
  });
}

/* ============================================================
   ① スケジュール（WBSカンバン用：大分類 → 小分類 → タスク）
   ============================================================ */
/* open … 展開中のキー。既定は空＝大分類だけの表示。
   タブを開くたびにリセットするため localStorage には保存しない。
   filter … スケジュールタブ専用の絞り込み（カンバン側とは独立）。 */
const schedState = {
  months: 3, offset: 0, open: {}, showLeave: false, showActual: true, hideDone: false,
  filter: { user: "", cat: "", sub: "" }
};

function schedRestore() {
  try {
    const s = JSON.parse(localStorage.getItem("wbs-sched") || "{}");
    ["months", "offset", "showLeave", "showActual", "hideDone"].forEach((k) => {
      if (s[k] !== undefined) schedState[k] = s[k];
    });
    if (s.filter) Object.assign(schedState.filter, s.filter);
  } catch (e) { /* 既定値で続行 */ }
}
function schedSave() {
  try {
    // 展開状態（open）は毎回リセットするので保存しない
    const { open, ...rest } = schedState;
    localStorage.setItem("wbs-sched", JSON.stringify(rest));
  } catch (e) { /* 無視 */ }
}

let schedHost = null;
let schedTasks = [];

async function renderSchedule(container, opts) {
  const host = typeof container === "string" ? document.getElementById(container) : container;
  if (!host) return;
  schedHost = host;
  const o = opts || {};
  if (o.reload !== false || !schedTasks.length) {
    schedTasks = await fetchWbsTasks(null);
    if (schedState.showLeave) await loadLeaveGrid();
  }
  schedState.open = {};        // 開くたびに大分類だけの表示に戻す
  drawSchedule(o);
}

function drawSchedule(o) {
  const host = schedHost;
  if (!host) return;
  const F = schedState.filter;
  const norm2 = (v) => (v == null ? "" : String(v)).trim();
  const tasks = schedTasks.filter((t) => {
    if (schedState.hideDone && t.actualEnd) return false;
    if (F.user && norm2(t.user) !== F.user) return false;
    if (F.cat && norm2(t.category) !== F.cat) return false;
    if (F.sub && norm2(t.classification) !== F.sub) return false;
    return true;
  });
  /* 選択肢は「絞り込み前」の全件から作る（自分で選んだ値が消えないように）。
     小分類は大分類が選ばれているときだけ、その配下に限定して出す。 */
  const uniq = (arr) => [...new Set(arr.filter((x) => x))].sort();
  const optUsers = uniq(schedTasks.map((t) => norm2(t.user)));
  const optCats = uniq(schedTasks.map((t) => norm2(t.category)));
  const optSubs = uniq(schedTasks
    .filter((t) => !F.cat || norm2(t.category) === F.cat)
    .map((t) => norm2(t.classification)));

  /* --- 期間 --- */
  const now = new Date();
  const t0 = new Date(now.getFullYear(), now.getMonth() + schedState.offset, 1);
  const t1 = new Date(t0.getFullYear(), t0.getMonth() + schedState.months, 1);
  const totalDays = Math.round((t1 - t0) / 86400000);
  const geomBase = { t0, totalDays, showLeave: schedState.showLeave, showActual: schedState.showActual };

  /* ---- カレンダー表記 ----
     ・月ラベルは中央寄せで「2026年8月」→以降は「9月」。年は最初と年替わりだけ
       出すので、12ヶ月表示でも横幅を食わず、どの年か迷わない。
     ・月ごとに背景を薄く互い違いにして、棒がどの月にあるか目で追えるようにする。
     ・日付の目盛りは3ヶ月以下のときだけ出す。6・12ヶ月では線が混んで
       かえって読めないため月ラベルに任せる。
     ・「今日」は専用の行に出して、月ラベルに重ならないようにする。 */
  let monthCells = "", monthLines = "", monthBands = "";
  for (let i = 0; i < schedState.months; i++) {
    const d = new Date(t0.getFullYear(), t0.getMonth() + i, 1);
    const n = new Date(t0.getFullYear(), t0.getMonth() + i + 1, 1);
    const l = schedPct(d, geomBase), w = schedPct(n, geomBase) - l;
    const showYear = i === 0 || d.getMonth() === 0;
    const label = showYear ? `${d.getFullYear()}年${d.getMonth() + 1}月` : `${d.getMonth() + 1}月`;
    monthCells += `<div class="g-mcell" style="left:${l}%;width:${w}%">${label}</div>`;
    if (i % 2 === 1) monthBands += `<div class="wsc-band" style="left:${l}%;width:${w}%"></div>`;
    if (i > 0) monthLines += `<div class="g-vline" style="left:${l}%"></div>`;
  }
  let dateCells = "";
  if (schedState.months <= 3) {
    const f = new Date(t0); f.setDate(f.getDate() - ((f.getDay() + 6) % 7));
    for (let d = new Date(f); d < t1; d.setDate(d.getDate() + 7)) {
      if (d < t0) continue;
      dateCells += `<span class="g-dcell" style="left:${schedPct(d, geomBase)}%">${d.getDate()}</span>`;
    }
  }
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let todayHtml = "", todayChip = "";
  if (today >= t0 && today < t1) {
    const lp = schedPct(today, geomBase);
    todayHtml = `<div class="g-today" style="left:${lp}%"></div>`;
    todayChip = `<div class="g-today-chip" style="left:${lp}%">今日 ${schedMd(today)}</div>`;
  }
  const geom = Object.assign({}, geomBase, { monthLines: monthBands + monthLines, todayHtml });

  /* --- 大分類 → 小分類 --- */
  const order = [], map = {};
  tasks.forEach((t) => {
    const k1 = (t.category || "（大分類なし）").toString();
    const k2 = (t.classification || "（小分類なし）").toString();
    if (!map[k1]) { map[k1] = { order: [], sub: {} }; order.push(k1); }
    if (!map[k1].sub[k2]) { map[k1].sub[k2] = []; map[k1].order.push(k2); }
    map[k1].sub[k2].push(t);
  });
  const firstOf = (list) => {
    const ds = list.map((t) => miniExcelDateToJS(t.start)).filter(Boolean).sort((a, b) => a - b);
    return ds[0] || null;
  };
  const byStart = (a, b) => {
    const x = miniExcelDateToJS(a.start), y = miniExcelDateToJS(b.start);
    if (!x && !y) return 0;
    if (!x) return 1;
    if (!y) return -1;
    return x - y;
  };
  const spanOf = (list) => {
    let mn = null, mx = null;
    list.forEach((t) => {
      const s = miniExcelDateToJS(t.start), e = miniExcelDateToJS(t.end);
      if (s && (!mn || s < mn)) mn = s;
      if (e && (!mx || e > mx)) mx = e;
    });
    return { start: mn, end: mx };
  };
  const rollHtml = (sp, ratio) => {
    if (!sp.start || !sp.end) return "";
    const l = schedClamp(schedPct(sp.start, geom));
    const r = schedClamp(schedPct(new Date(sp.end.getTime() + 86400000), geom));
    const w = Math.max(r - l, 0.5);
    return `<div class="wsc-roll" style="left:${l}%;width:${w}%"></div>`
         + `<div class="wsc-rollp" style="left:${l}%;width:${w * ratio}%"></div>`
         + (l > 4 ? `<div class="wsc-sd" style="left:${Math.max(l - 6.5, 0)}%">${schedMd(sp.start)}</div>` : "");
  };
  const delayN = (list) => list.filter((t) => isDelayKey(taskStatusKey(t))).length;

  let rows = "";
  order.forEach((k1) => {
    const g1 = map[k1];
    const open1 = !!schedState.open[k1];       // 既定は閉（大分類だけ表示）
    let all = [];
    g1.order.forEach((k2) => { all = all.concat(g1.sub[k2]); });
    const dn = all.filter((t) => t.actualEnd).length;
    const dl = delayN(all);

    rows += `<div class="g-row wsc-cat" data-cat="${escapeHtml(k1)}">
      <div class="g-label">
        <div class="wsc-k"><span class="wsc-chev">${open1 ? "▼" : "▶"}</span>${escapeHtml(k1)}
          <span class="wsc-cnt">${dn}/${all.length}</span>
          ${dl ? `<span class="wsc-warn">遅延${dl}</span>` : ""}</div>
        <div class="wsc-sub2">小分類 ${g1.order.length}件</div>
      </div>
      <div class="g-track">${monthLines}${todayHtml}${rollHtml(spanOf(all), all.length ? dn / all.length : 0)}</div>
    </div>`;
    if (!open1) return;

    g1.order.sort((a, b) => {
      const fa = firstOf(g1.sub[a]), fb = firstOf(g1.sub[b]);
      if (!fa && !fb) return 0;
      if (!fa) return 1;
      if (!fb) return -1;
      return fa - fb;
    });

    g1.order.forEach((k2) => {
      const list = g1.sub[k2].slice().sort(byStart);
      const key = k1 + "/" + k2;
      const sel = !!schedState.open[key];
      const d2 = list.filter((t) => t.actualEnd).length;
      const dl2 = delayN(list);
      const sp = spanOf(list);

      rows += `<div class="g-row wsc-subrow${sel ? " sel" : ""}" data-sub="${escapeHtml(key)}">
        <div class="g-label">
          <div class="wsc-k"><span class="wsc-chev">${sel ? "▼" : "▶"}</span>${escapeHtml(k2)}
            <span class="wsc-cnt${sel ? " on" : ""}">${d2}/${list.length}</span>
            ${dl2 ? `<span class="wsc-warn">遅延${dl2}</span>` : ""}</div>
        </div>
        <div class="g-track">${monthLines}${todayHtml}${rollHtml(sp, list.length ? d2 / list.length : 0)}</div>
      </div>`;
      if (!sel) return;

      // スケジュールタブではタスク一覧（表）は出さない。線だけを見せる。
      rows += taskLineRowsHtml(list, Object.assign({}, geom, { guide: sp }));
    });
  });

  const sw = (k, label, title) =>
    `<label class="wsc-sw${schedState[k] ? " on" : ""}" data-sw="${k}" title="${escapeHtml(title || "")}">
       <span class="switch"></span>${label}</label>`;

  const selHtml = (key, label, opts, cur) =>
    `<label class="wsc-fl"><span class="fl-k">${label}</span>
       <select data-f="${key}">
         <option value="">すべて</option>
         ${opts.map((o) => `<option value="${escapeHtml(o)}"${o === cur ? " selected" : ""}>${escapeHtml(o)}</option>`).join("")}
       </select></label>`;

  host.innerHTML = `
  <div class="gantt-toolbar wsc-toolbar">
    <button class="term-btn" data-mv="-1">◀</button>
    <span class="term-label">${t0.getFullYear()}/${String(t0.getMonth() + 1).padStart(2, "0")} 〜 ${new Date(t1 - 86400000).getFullYear()}/${String(new Date(t1 - 86400000).getMonth() + 1).padStart(2, "0")}</span>
    <button class="term-btn" data-mv="1">▶</button>
    <span class="term-bar-sep"></span>
    ${sw("hideDone", "完了除く")}
    <span class="g-sp"></span>
    ${sw("showActual", "実績")}
    ${sw("showLeave", "休み", "担当者の休みを帯で重ねる")}
    <div class="g-zoom wsc-zoom">
      ${[12, 6, 3, 1].map((m) => `<button class="${m === schedState.months ? "active" : ""}" data-z="${m}">${m}ヶ月</button>`).join("")}
    </div>
  </div>

  <div class="wsc-filter">
    ${selHtml("user", "担当者", optUsers, F.user)}
    ${selHtml("cat", "大分類", optCats, F.cat)}
    ${selHtml("sub", "小分類", optSubs, F.sub)}
    ${(F.user || F.cat || F.sub) ? '<button class="wsc-clear" data-clear="1">条件クリア</button>' : ""}
    <span class="g-sp"></span>
    <span class="wsc-hint">ドラッグで上下左右に移動／右クリックで詳細</span>
  </div>

  <div class="wsc-wrap">
    <div class="wsc-inner">
      <div class="g-row g-headrow wsc-head">
        <div class="g-label">大分類 / 小分類</div>
        <div class="wsc-headstack">
          <div class="g-track wsc-todayrow">${todayChip}</div>
          <div class="g-track g-head">${monthCells}</div>
          ${dateCells ? `<div class="g-track g-head g-daterow">${dateCells}</div>` : ""}
        </div>
      </div>
      ${rows || '<div class="wsc-empty2">表示できるタスクがありません</div>'}
    </div>
  </div>
  <div class="wsc-legend">
    <span><i class="roll"></i>集計（緑＝完了割合）</span>
    <span><i class="delaydoing"></i>対応中(遅延)</span><span><i class="delaytodo"></i>未着手(遅延)</span>
    <span><i class="active"></i>期間内</span>
    <span><i class="held"></i>保留</span><span><i class="todo"></i>未着手</span>
    <span><i class="done"></i>完了</span><span><i class="act"></i>実績</span>
    <span><i class="leave"></i>休み（AM・PMは薄く／リモートは出しません）</span>
  </div>`;

  bindSchedule(host);
  bindTaskLines(host);
}

function bindSchedule(host) {
  host.querySelectorAll(".wsc-toolbar [data-mv]").forEach((b) => {
    b.addEventListener("click", () => { schedState.offset += Number(b.dataset.mv); schedSave(); drawSchedule(); });
  });
  host.querySelectorAll(".wsc-zoom [data-z]").forEach((b) => {
    b.addEventListener("click", () => { schedState.months = Number(b.dataset.z); schedSave(); drawSchedule(); });
  });
  host.querySelectorAll(".wsc-toolbar [data-sw]").forEach((b) => {
    b.addEventListener("click", async () => {
      const k = b.dataset.sw;
      schedState[k] = !schedState[k];
      schedSave();
      if (k === "showLeave" && schedState.showLeave) await loadLeaveGrid();
      drawSchedule();
    });
  });
  // 大分類・小分類の開閉（open ベース。既定は大分類だけの表示）
  host.querySelectorAll(".wsc-cat").forEach((r) => {
    r.addEventListener("click", () => {
      const k = r.dataset.cat;
      if (schedState.open[k]) {
        // 大分類を閉じるときは、その配下の小分類の展開も畳む
        delete schedState.open[k];
        Object.keys(schedState.open).forEach((x) => {
          if (x.indexOf(k + "/") === 0) delete schedState.open[x];
        });
      } else {
        schedState.open[k] = true;
      }
      drawSchedule();
    });
  });
  host.querySelectorAll(".wsc-subrow").forEach((r) => {
    r.addEventListener("click", () => {
      const k = r.dataset.sub;
      if (schedState.open[k]) delete schedState.open[k];
      else schedState.open[k] = true;
      drawSchedule();
    });
  });

  // フィルタ
  host.querySelectorAll(".wsc-filter select[data-f]").forEach((sel) => {
    sel.addEventListener("change", () => {
      const k = sel.dataset.f;
      schedState.filter[k] = sel.value;
      if (k === "cat") schedState.filter.sub = "";   // 大分類を変えたら小分類は解除
      schedState.open = {};                          // 絞り込み直後は大分類だけに戻す
      schedSave(); drawSchedule();
    });
  });
  const clr = host.querySelector(".wsc-filter [data-clear]");
  if (clr) clr.addEventListener("click", () => {
    schedState.filter = { user: "", cat: "", sub: "" };
    schedState.open = {};
    schedSave(); drawSchedule();
  });

  bindSchedDrag(host);
  bindSchedContext(host);
}

/* ドラッグでガントを上下左右に動かす。
   バーやフィルタの操作を邪魔しないよう、行の余白部分でだけ掴む。 */
function bindSchedDrag(host) {
  const wrap = host.querySelector(".wsc-wrap");
  if (!wrap) return;
  let drag = null;

  wrap.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    // バー・ラベル・ヘッダーからは開始しない（クリック操作を優先）
    if (e.target.closest(".wsc-t, .g-label, .wsc-head, select, button, label")) return;
    drag = {
      x: e.clientX, y: e.clientY,
      sl: wrap.scrollLeft, st: wrap.scrollTop, moved: false
    };
    wrap.classList.add("dragging");
  });
  window.addEventListener("mousemove", (e) => {
    if (!drag) return;
    const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) drag.moved = true;
    wrap.scrollLeft = drag.sl - dx;
    wrap.scrollTop = drag.st - dy;
    if (drag.moved) e.preventDefault();
  });
  window.addEventListener("mouseup", () => {
    if (!drag) return;
    wrap.classList.remove("dragging");
    drag = null;
  });
}

/* 右クリックでタスクの詳細（備考編集モーダル）を開く */
function bindSchedContext(host) {
  host.addEventListener("contextmenu", (e) => {
    const row = e.target.closest(".wsc-taskrow, .wsc-t[data-row]");
    if (!row) return;
    e.preventDefault();
    const rowIndex = Number(row.dataset.row);
    if (!rowIndex) return;
    const t = schedTasks.find((x) => x.rowIndex === rowIndex);
    if (t && typeof openModal === "function") openModal(t);
  });
}

/* ============================================================
   公開API（inline onclick から呼ばれるものを含む）
   ------------------------------------------------------------
   ここに載せたものだけがグローバルから見える。
   新しく外部公開する関数を追加したら、必ずここにも追記すること。
   ============================================================ */
Object.assign(window, {
  // 呼び出し側アプリから使う
  openModal, openTaskAdd, renderMiniKanban, matchByCaseId, fetchWbsTasks,
  jumpToWbsRow, escapeHtml, dateToExcelSerial, ensureStatusSymbols, stampWbsUpdate,
  findWbsHeader, parseSubtasks,
  // ③ サブタスク数
  subtaskCount, subtaskBadgeHtml,
  // ④ 休み予定
  loadLeaveGrid, leavesOf, setLeave, leaveStyleOf, renderLeaveMatrix, openMyLeave, closeMyLeave,
  // ① スケジュール
  renderSchedule, taskLineRowsHtml, taskListHtml, bindTaskLines, taskStatusKey, schedRestore,
  // 注入したモーダルの inline onclick から呼ばれる
  closeModal, saveNote, onModalNoteEdited, addSubtask,
  closeTaskAdd, saveTaskAdd, onTaCatChange,
  dialogRespond
});

/* uiAlert / uiConfirm は呼び出し側が独自実装を持つ場合があるため、
   まだ定義されていないときだけ載せる（営業報告は自前のものを使う）。 */
if (typeof window.uiAlert !== "function") window.uiAlert = uiAlert;
if (typeof window.uiConfirm !== "function") window.uiConfirm = uiConfirm;

if (document.body) ensureApiDom();
else document.addEventListener("DOMContentLoaded", ensureApiDom);

})();
