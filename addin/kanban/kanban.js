let currentTasks = [
  { id: 1, name: "ログイン画面", status: "todo", order: 1 },
  { id: 2, name: "API連携", status: "doing", order: 1 },
  { id: 3, name: "テスト", status: "done", order: 1 },
  { id: 4, name: "UI改善", status: "todo", order: 2 }
];

let undoStack = [];
let redoStack = [];

init();

function init() {
  render();
  setupDnD();
}

function snapshot() {
  undoStack.push(JSON.stringify(currentTasks));
  redoStack = [];
}

function undo() {
  if (!undoStack.length) return;
  redoStack.push(JSON.stringify(currentTasks));
  currentTasks = JSON.parse(undoStack.pop());
  render();
}

function redo() {
  if (!redoStack.length) return;
  undoStack.push(JSON.stringify(currentTasks));
  currentTasks = JSON.parse(redoStack.pop());
  render();
}

function render() {
  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  currentTasks
    .sort((a, b) => a.order - b.order)
    .forEach(task => {
      const el = document.createElement("div");
      el.className = "card";
      el.draggable = true;
      el.textContent = task.name;
      el.dataset.id = task.id;

      el.addEventListener("dragstart", e => {
        el.classList.add("dragging");
        e.dataTransfer.setData("id", task.id);
      });

      el.addEventListener("dragend", () => {
        el.classList.remove("dragging");
      });

      document.querySelector(`#${task.status} .card-list`).appendChild(el);
    });
}

function setupDnD() {
  document.querySelectorAll(".lane").forEach(lane => {
    lane.addEventListener("dragover", e => {
      e.preventDefault();
      const container = lane.querySelector(".card-list");
      const after = getAfter(container, e.clientY);
      const dragging = document.querySelector(".dragging");

      if (!after) container.appendChild(dragging);
      else container.insertBefore(dragging, after);
    });

    lane.addEventListener("drop", e => {
      snapshot();

      const id = e.dataTransfer.getData("id");
      const task = currentTasks.find(t => t.id == id);
      task.status = lane.id;

      updateOrder(lane.id);
      render();
    });
  });
}

function getAfter(container, y) {
  const els = [...container.querySelectorAll(".card:not(.dragging)")];

  return els.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;

    if (offset < 0 && offset > closest.offset) {
      return { offset, element: child };
    } else {
      return closest;
    }
  }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function updateOrder(laneId) {
  const cards = document.querySelectorAll(`#${laneId} .card`);

  cards.forEach((card, index) => {
    const task = currentTasks.find(t => t.id == card.dataset.id);
    task.order = index + 1;
  });
}
