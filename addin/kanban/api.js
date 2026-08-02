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
  onTaskAdded: null
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
  if (!firstLine.includes("★") && !firstLine.includes("☆")) firstLine = "☆" + firstLine;
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

    await ctx.sync();
  });

  if (currentTask) {
    currentTask.note = note;
    currentTask.isStar = note.startsWith("★");
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

      range.values.slice(10).forEach((row) => {
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

/* preset（任意）: { category, caseId }
   … 呼び出し元（例：営業報告の案件編集画面）から、大分類と
     案件番号（＝小分類、大分類が「受注」の場合の候補）を
     あらかじめ選択した状態でモーダルを開きたいときに使う。
     例）openTaskAdd({ category: "受注", caseId: "AG-03" }) */
async function openTaskAdd(preset) {
  preset = preset || {};
  taskAddMeta = await loadWbsMeta();

  // 大分類: wbs既存の大分類 ＋ 受注（無ければ追加）
  const cats = [...taskAddMeta.categories];
  if (!cats.includes(cfg().orderCategory)) cats.push(cfg().orderCategory);
  const catSel = document.getElementById("ta-cat");
  catSel.innerHTML = cats.map(c => `<option>${escapeHtml(String(c))}</option>`).join("");
  if (preset.category && cats.includes(preset.category)) catSel.value = preset.category;

  // 担当者: wbs既存の担当者
  const userSel = document.getElementById("ta-user");
  userSel.innerHTML = `<option value=""></option>` + taskAddMeta.users.map(u => `<option>${escapeHtml(String(u))}</option>`).join("");

  // 入力初期化
  document.getElementById("ta-subcat").value = "";
  document.getElementById("ta-title").value = "";
  document.getElementById("ta-start").value = "";
  document.getElementById("ta-end").value = "";
  const msg = document.getElementById("ta-msg");
  msg.className = "task-msg"; msg.textContent = "";

  await onTaCatChange();

  // 受注案件番号の事前選択（該当する候補があれば）
  if (preset.caseId && catSel.value === cfg().orderCategory) {
    const sel = document.getElementById("ta-subcat-sel");
    if ([...sel.options].some(o => o.value === preset.caseId)) sel.value = preset.caseId;
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

  const cat = document.getElementById("ta-cat").value;
  const sub = (cat === cfg().orderCategory)
    ? document.getElementById("ta-subcat-sel").value
    : document.getElementById("ta-subcat").value.trim();
  const title = document.getElementById("ta-title").value.trim();
  const user = document.getElementById("ta-user").value;
  const start = document.getElementById("ta-start").value;
  const end = document.getElementById("ta-end").value;

  if (!title) { msg.className = "task-msg err"; msg.textContent = "タスク名を入力してください"; return; }
  if (cat === cfg().orderCategory && !sub) { msg.className = "task-msg err"; msg.textContent = "案件番号を選択してください"; return; }
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

async function jumpToWbsRow(row) {
  await Excel.run(async (ctx) => {
    const s = ctx.workbook.worksheets.getItem(cfg().wbsSheet);
    s.activate();
    s.getRange(`${row}:${row}`).select();
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

      range.values.slice(10).forEach((row, i) => {
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
          rowIndex: i + 11,
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

  const lanesHtml = MINI_LANES.map(L => {
    const cards = tasks.filter(t => miniStatus(t) === L.key);
    return `
      <div class="mk-lane" data-lane="${L.key}">
        <div class="mk-lane-head">${escapeHtml(L.label)} <span>${cards.length}</span></div>
        <div class="mk-lane-body" data-lane="${L.key}">
          ${cards.map(t => {
            const overdue = L.key !== "done" && t.end &&
              miniExcelDateToJS(t.end) < new Date(new Date().toDateString());
            return `<div class="mk-card ${L.key}${overdue ? " overdue" : ""}" draggable="true" data-row="${t.rowIndex}">
              ${escapeHtml(t.title || "（無題）")}
              <span class="due">${t.isNoSchedule ? "TODO" : miniFmt(t.end)}${t.user ? "・" + escapeHtml(t.user) : ""}</span>
            </div>`;
          }).join("")}
        </div>
      </div>`;
  }).join("");

  el.innerHTML = `<div class="mk-board">${lanesHtml}</div>`;
  bindMiniKanbanEvents(el, tasks);
  return tasks;
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
          <div class="t-row"><label>大分類</label><select id="ta-cat" onchange="onTaCatChange()"></select></div>
          <div class="t-row"><label>小分類</label>
            <input type="text" id="ta-subcat" list="ta-subcat-list" placeholder="既存候補から選択 or 新規入力">
            <datalist id="ta-subcat-list"></datalist>
            <select id="ta-subcat-sel" style="display:none"></select>
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
   公開API（inline onclick から呼ばれるものを含む）
   ------------------------------------------------------------
   ここに載せたものだけがグローバルから見える。
   新しく外部公開する関数を追加したら、必ずここにも追記すること。
   ============================================================ */
Object.assign(window, {
  // 呼び出し側アプリから使う
  openModal, openTaskAdd, renderMiniKanban, matchByCaseId, fetchWbsTasks,
  jumpToWbsRow, escapeHtml, dateToExcelSerial, ensureStatusSymbols,
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
