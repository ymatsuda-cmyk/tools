// src/dialog/kanban.js

let tasks = [];
let assignees = [];
let currentEditId = null;

Office.onReady(() => {
  const hash = new URLSearchParams(location.hash.replace("#", ""));
  const dataStr = hash.get("data") || "%5B%5D"; // default []
  const payload = JSON.parse(decodeURIComponent(dataStr));

  tasks = payload.tasks || [];
  assignees = payload.assignees || [];

  renderKanban(tasks);
  setupEditDialog();
});

const STATUS_COLUMNS = ["Todo", "Doing", "Done"];

function renderKanban(taskList) {
  const board = document.getElementById("board");
  board.innerHTML = "";

  STATUS_COLUMNS.forEach(status => {
    const col = document.createElement("div");
    col.className = "column";
    col.dataset.status = status;
    col.addEventListener("dragover", onColumnDragOver);
    col.addEventListener("drop", onColumnDrop);

    const header = document.createElement("div");
    header.className = "column-header";
    header.textContent = status;
    col.appendChild(header);

    const filtered = taskList.filter(t => t.status === status);
    filtered.forEach(task => {
      const card = document.createElement("div");
      card.className = "card";
      card.draggable = true;
      card.dataset.id = task.id;

      card.addEventListener("dragstart", onCardDragStart);
      card.addEventListener("click", () => openEditDialog(task.id));

      card.innerHTML = `
        <div><strong>${escapeHtml(task.title || "")}</strong></div>
        <div>${escapeHtml(task.assignee || "")}</div>
        <div>${formatPlan(task.plannedStart, task.plannedEnd)}</div>
        <div>
          ${task.tagLarge ? `<span class="tag">${escapeHtml(task.tagLarge)}</span>` : ""}
          ${task.tagSmall ? `<span class="tag">${escapeHtml(task.tagSmall)}</span>` : ""}
        </div>
      `;
      col.appendChild(card);
    });

    board.appendChild(col);
  });
}

function formatPlan(start, end) {
  if (!start && !end) return "";
  if (start && !end) return `${start}〜`;
  if (!start && end) return `〜${end}`;
  return `${start}〜${end}`;
}

function onCardDragStart(ev) {
  ev.dataTransfer.setData("text/plain", ev.target.dataset.id);
}
function onColumnDragOver(ev) {
  ev.preventDefault();
}
function onColumnDrop(ev) {
  ev.preventDefault();
  const id = ev.dataTransfer.getData("text/plain");
  const newStatus = ev.currentTarget.dataset.status; // Todo / Doing / Done
  handleMoveCard(id, newStatus);
}

// 実績が入っている場合は上書き確認を行う
function handleMoveCard(id, newStatus) {
  const task = tasks.find(t => String(t.id) === String(id));
  if (!task) return;

  const hasActual =
    (task.actualStart && task.actualStart !== "") ||
    (task.actualEnd && task.actualEnd !== "");

  let forceOverwrite = false;
  if (hasActual) {
    const ok = window.confirm(
      "Excel側の実績開始日／実績終了日に値が設定されています。\n" +
      `このタスクを「${newStatus}」として上書き更新してもよろしいですか？`
    );
    if (!ok) {
      return; // キャンセル
    }
    forceOverwrite = true;
  }

  // Host へ通知
  Office.context.ui.messageParent(JSON.stringify({
    type: "move",
    id,
    status: newStatus,
    forceOverwrite
  }));

  // ローカル状態も更新して再描画
  task.status = newStatus;
  renderKanban(tasks);
}

// ===== 編集ダイアログ =====

function setupEditDialog() {
  const backdrop = document.getElementById("edit-dialog-backdrop");
  const btnCancel = document.getElementById("edit-cancel");
  const btnOk     = document.getElementById("edit-ok");

  btnCancel.addEventListener("click", () => {
    backdrop.style.display = "none";
    currentEditId = null;
  });

  btnOk.addEventListener("click", () => {
    saveEditDialog();
  });
}

function openEditDialog(id) {
  const task = tasks.find(t => String(t.id) === String(id));
  if (!task) return;

  currentEditId = id;

  document.getElementById("edit-title").textContent = task.title || "";

  const sel = document.getElementById("edit-assignee");
  sel.innerHTML = "";
  const emptyOpt = document.createElement("option");
  emptyOpt.value = "";
  emptyOpt.textContent = "";
  sel.appendChild(emptyOpt);

  assignees.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    if (task.assignee === name) {
      opt.selected = true;
    }
    sel.appendChild(opt);
  });

  document.getElementById("edit-planned-start").value = task.plannedStart || "";
  document.getElementById("edit-planned-end").value   = task.plannedEnd || "";
  document.getElementById("edit-note").value          = task.note || "";

  document.getElementById("edit-dialog-backdrop").style.display = "flex";
}

function saveEditDialog() {
  const backdrop = document.getElementById("edit-dialog-backdrop");
  const assignee = document.getElementById("edit-assignee").value;
  const plannedStart = document.getElementById("edit-planned-start").value.trim();
  const plannedEnd   = document.getElementById("edit-planned-end").value.trim();
  const note         = document.getElementById("edit-note").value;

  // 備考の1024文字チェックは maxlength 属性で UI 側も制限済み

  // 予定日の形式チェック（簡易版：m/d のみ許可する等、必要ならここで実装）

  // Host へ通知
  Office.context.ui.messageParent(JSON.stringify({
    type: "edit",
    id: currentEditId,
    assignee,
    plannedStart,
    plannedEnd,
    note
  }));

  // ローカル状態を更新
  const task = tasks.find(t => String(t.id) === String(currentEditId));
  if (task) {
    task.assignee    = assignee;
    task.plannedStart = plannedStart;
    task.plannedEnd   = plannedEnd;
    task.note         = note;
  }

  renderKanban(tasks);

  backdrop.style.display = "none";
  currentEditId = null;
}

// ===== Utility =====

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}