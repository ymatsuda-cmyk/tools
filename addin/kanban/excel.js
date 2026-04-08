async function loadTasks() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    const range = sheet.getRange("P2:Z200");
    range.load("values");
    await context.sync();
    return range.values.map((r, i) => {
      let status = "todo";
      if (r[3]) status = "done";
      else if (r[2]) status = "doing";
      return {
        id: i,
        row: i + 2,
        name: r[10],
        order: r[4] || 0,
        status
      };
    });
  });
}

async function saveOrder(laneId) {
  const cards = document.querySelectorAll(`#${laneId} .card`);
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("wbs");
    cards.forEach((card, index) => {
      const row = card.dataset.row;
      sheet.getRange(`T${row}`).values = [[index + 1]];
    });
    await context.sync();
  });
}
