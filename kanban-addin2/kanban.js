let tasks = [];
let dragging = null;

Office.onReady(async () => {
  tasks = await loadWbsTasks();
  render();
});

function render() {
  document.querySelectorAll(".lane-body").forEach(l => l.innerHTML = "");

  tasks.forEach(t => {
    const status =
      t.actualEnd ? "完了" :
      t.actualStart ? "対応中" : "未着手";

    const card = document.createElement("div");
    card.className = "card";
    card.draggable = true;
    card.textContent = t.title;

    card.addEventListener("dragstart", () => dragging = t);
    card.addEventListener("dragend", () => dragging = null);

    document
      .querySelector(`.lane[data-status="${status}"] .lane-body`)
      .appendChild(card);
  });
}

document.querySelectorAll(".lane-body").forEach(lane => {
  lane.addEventListener("dragover", e => {
    e.preventDefault();
    lane.parentElement.classList.add("drag-over");
  });

  lane.addEventListener("dragleave", () => {
    lane.parentElement.classList.remove("drag-over");
  });

  lane.addEventListener("drop", async () => {
    lane.parentElement.classList.remove("drag-over");
    if (!dragging) return;

    const newStatus = lane.parentElement.dataset.status;
    await updateStatus(dragging, newStatus);

    tasks = await loadWbsTasks();
    render();
  });
});