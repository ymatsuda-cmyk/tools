// src/commands/commands.js

let dialog;

// リボンボタンから呼ばれる ExecuteFunction
async function openKanban(event) {
  try {
    const payload = await Excel.run(async (context) => {
      // === 1) WBS テーブルからタスク情報取得 ===
      const wbsSheet = context.workbook.worksheets.getItem("WBS");
      const wbsTable = wbsSheet.tables.getItem("tblWBS");

      const header = wbsTable.getHeaderRowRange();
      const body   = wbsTable.getDataBodyRange();

      header.load("values");
      body.load("values");

      // === 2) コードシートから担当者候補を取得 ===
      const codeSheet = context.workbook.worksheets.getItem("Codes");
      const assigneeTable = codeSheet.tables.getItem("tblAssignee");
      const assigneeBody  = assigneeTable.getDataBodyRange();
      assigneeBody.load("values");

      await context.sync();

      const headers = header.values[0].map(h => String(h).trim());
      const col = (name) => headers.indexOf(name);

      const idCol           = col("ID");
      const titleCol        = col("task");
      const assigneeCol     = col("担当者");
      const plannedStartCol = col("予定開始日");
      const plannedEndCol   = col("予定終了日");
      const actualStartCol  = col("実績開始日");
      const actualEndCol    = col("実績終了日");
      const noteCol         = col("備考");
      const tagLargeCol     = col("大分類");
      const tagSmallCol     = col("小分類");

      // タスク一覧の生成
      const tasks = body.values.map((r) => {
        const actualStart = r[actualStartCol];
        const actualEnd   = r[actualEndCol];

        let status;
        if (!actualStart && !actualEnd) {
          status = "Todo";
        } else if (actualStart && !actualEnd) {
          status = "Doing";
        } else if (actualEnd) {
          status = "Done";
        }

        return {
          id:          r[idCol],
          title:       r[titleCol],
          assignee:    r[assigneeCol],
          plannedStart:r[plannedStartCol],
          plannedEnd:  r[plannedEndCol],
          actualStart,
          actualEnd,
          status,
          note:        r[noteCol],
          tagLarge:    r[tagLargeCol],
          tagSmall:    r[tagSmallCol],
        };
      });

      // 担当者候補（Codes シート） ※1列目に名前が入っている前提
      const assignees = assigneeBody.values
        .map(row => String(row[0]).trim())
        .filter(name => !!name);

      return {
        tasks,
        assignees
      };
    });

    // === 3) Dialog を開く ===
    const url = `https://your-org.github.io/kanban-addin/src/dialog/kanban.html#data=${encodeURIComponent(JSON.stringify(payload))}`;

    Office.context.ui.displayDialogAsync(
      url,
      { height: 90, width: 90 },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) return;
        dialog = asyncResult.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          onDialogMessage
        );
      }
    );
  } finally {
    // ExecuteFunction の必須
    event.completed();
  }
}

// Dialog からのメッセージを受ける
async function onDialogMessage(arg) {
  const msg = JSON.parse(arg.message);

  switch (msg.type) {
    case "move":
      // ステータス変更 → 実績開始日／実績終了日 更新
      await updateActualDatesByStatus(msg.id, msg.status, msg.forceOverwrite);
      break;
    case "edit":
      // 予定開始日／予定終了日／担当者／備考 の更新
      await updateTaskDetails(msg);
      break;
    default:
      break;
  }
}

// status に応じて 実績開始日／実績終了日 を更新
async function updateActualDatesByStatus(id, newStatus, forceOverwrite) {
  await Excel.run(async (context) => {
    const sheet  = context.workbook.worksheets.getItem("WBS");
    const table  = sheet.tables.getItem("tblWBS");
    const header = table.getHeaderRowRange();
    const body   = table.getDataBodyRange();

    header.load("values");
    body.load("values");
    await context.sync();

    const headers = header.values[0].map(h => String(h).trim());
    const col = (name) => headers.indexOf(name);

    const idCol          = col("ID");
    const actualStartCol = col("実績開始日");
    const actualEndCol   = col("実績終了日");

    const values   = body.values;
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(id));
    if (rowIndex < 0) return;

    const rowRange = body.getRow(rowIndex);
    const today = new Date();

    const currentStart = values[rowIndex][actualStartCol];
    const currentEnd   = values[rowIndex][actualEndCol];

    // forceOverwrite が false のとき、既存値があるなら何もしない(保護)も可能
    // ここでは「Dialog 側で確認済み」を前提に、来たら書き換える想定でもOK

    if (newStatus === "Todo") {
      rowRange.getCell(0, actualStartCol).values = [[null]];
      rowRange.getCell(0, actualEndCol).values   = [[null]];

    } else if (newStatus === "Doing") {
      if (!currentStart || forceOverwrite) {
        rowRange.getCell(0, actualStartCol).values = [[today]];
      }
      rowRange.getCell(0, actualEndCol).values = [[null]];

    } else if (newStatus === "Done") {
      if (!currentStart || forceOverwrite) {
        rowRange.getCell(0, actualStartCol).values = [[today]];
      }
      rowRange.getCell(0, actualEndCol).values = [[today]];
    }

    await context.sync();
  });
}

// 予定開始日／予定終了日／担当者／備考 の更新
async function updateTaskDetails(msg) {
  // msg: { id, assignee, plannedStart, plannedEnd, note }
  await Excel.run(async (context) => {
    const sheet  = context.workbook.worksheets.getItem("WBS");
    const table  = sheet.tables.getItem("tblWBS");
    const header = table.getHeaderRowRange();
    const body   = table.getDataBodyRange();

    header.load("values");
    body.load("values");
    await context.sync();

    const headers = header.values[0].map(h => String(h).trim());
    const col = (name) => headers.indexOf(name);

    const idCol           = col("ID");
    const assigneeCol     = col("担当者");
    const plannedStartCol = col("予定開始日");
    const plannedEndCol   = col("予定終了日");
    const noteCol         = col("備考");

    const values   = body.values;
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(msg.id));
    if (rowIndex < 0) return;

    const rowRange = body.getRow(rowIndex);

    rowRange.getCell(0, assigneeCol).values     = [[msg.assignee || ""]];
    rowRange.getCell(0, plannedStartCol).values = [[msg.plannedStart || ""]];
    rowRange.getCell(0, plannedEndCol).values   = [[msg.plannedEnd || ""]];
    rowRange.getCell(0, noteCol).values         = [[msg.note || ""]];

    await context.sync();
  });
}

// ExecuteFunction の関連付け
Office.actions.associate("openKanban", openKanban);