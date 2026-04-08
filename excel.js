async function loadTasks() {
  return Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("wbs");

    // 🔥 必要列だけ取る（高速＆安全）
    const range = sheet.getRange("P2:Z1000");
    range.load("values");

    await context.sync();

    const tasks = [];

    range.values.forEach((row, i) => {

      const plannedStart = row[0]; // P
      const plannedEnd   = row[1]; // Q
      const actualStart  = row[2]; // R
      const actualEnd    = row[3]; // S
      const order        = row[4]; // T
      const name         = row[10]; // Z

      // 空行スキップ
      if (!name) return;

      // 🔥 状態判定（列は作らない）
      let status = "todo";
      if (actualEnd) status = "done";
      else if (actualStart) status = "doing";

      tasks.push({
        id: i,
        row: i + 2,
        name,
        status,
        order: order || 0,
        plannedStart,
        plannedEnd
      });

    });

    return tasks;
  });
}