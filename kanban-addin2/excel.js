const COL_TASK = 25; // Z列 (0-based)
const COL_PLAN_START = 15; // P
const COL_PLAN_END = 16;   // Q
const COL_ACT_START = 17;  // R
const COL_ACT_END = 18;    // S
const COL_ID = 23;         // ID列（実データに合わせて調整）

async function loadWbsTasks() {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    return range.values
      .map((r, i) => ({ row: i, r }))
      .filter(x =>
        x.r[COL_TASK] &&
        x.r[COL_PLAN_START] &&
        x.r[COL_PLAN_END] &&
        !x.r[COL_ACT_END]
      )
      .map(x => ({
        row: x.row,
        id: x.r[COL_ID],
        title: x.r[COL_TASK],
        start: x.r[COL_PLAN_START],
        end: x.r[COL_PLAN_END],
        actualStart: x.r[COL_ACT_START],
        actualEnd: x.r[COL_ACT_END]
      }));
  });
}

async function updateStatus(task, status) {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem("wbs");

    const startCell = sheet.getCell(task.row, COL_ACT_START);
    const endCell   = sheet.getCell(task.row, COL_ACT_END);

    const today = new Date();

    if (status === "未着手") {
      startCell.values = [[""]];
      endCell.values = [[""]];
    }
    if (status === "対応中") {
      if (!task.actualStart) startCell.values = [[today]];
      endCell.values = [[""]];
    }
    if (status === "完了") {
      if (!task.actualStart) startCell.values = [[today]];
      endCell.values = [[today]];
    }

    await ctx.sync();
  });
}