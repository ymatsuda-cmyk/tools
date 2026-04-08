let tasks = [];

Office.onReady(async () => {
  tasks = await loadTasks();
  render();
});

function render() {

  document.querySelectorAll(".card-list").forEach(el => el.innerHTML = "");

  // 🔥 並び順反映
  tasks.sort((a, b) => a.order - b.order);

  tasks.forEach(task => {

    const card = document.createElement("div");
    card.className = "card";
    card.textContent = task.name;

    card.dataset.row = task.row;
    card.dataset.id = task.id;

    // 🔥 期限表示（Trello風）
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = formatDate(task.plannedEnd);

    card.appendChild(meta);

    document
      .querySelector(`#${task.status} .card-list`)
      .appendChild(card);
  });
}

function formatDate(date) {
  if (!date) return "";
  const d = new Date(date);
  return `${d.getMonth()+1}/${d.getDate()}`;
}


