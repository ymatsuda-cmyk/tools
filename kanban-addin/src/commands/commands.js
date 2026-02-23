// src/commands/commands.js

let dialog;

// Office.js が読み込まれたかチェック
Office.onReady((info) => {
  console.log("Office.js loaded successfully");
  console.log("Office host:", info.host);
  console.log("Office platform:", info.platform);
  
  // ExecuteFunction の関連付け
  if (Office.actions && Office.actions.associate) {
    try {
      Office.actions.associate("openKanban", openKanban);
      console.log("openKanban function registered successfully");
    } catch (error) {
      console.error("Failed to register openKanban function:", error);
    }
  } else {
    console.error("Office.actions.associate is not available");
    console.log("Available Office APIs:", Object.keys(Office));
  }
});

// リボンボタンから呼ばれる ExecuteFunction
async function openKanban(event) {
  console.log("=== openKanban function called ===");
  console.log("Event object:", event);
  
  // 基本的なOffice.js APIの可用性チェック
  console.log("Office object available:", typeof Office !== 'undefined');
  console.log("Office.context available:", !!(Office && Office.context));
  console.log("Excel object available:", typeof Excel !== 'undefined');
  console.log("Excel.run available:", !!(Excel && Excel.run));
  
  try {
    // Office.js が利用可能かチェック
    if (!Office || !Office.context) {
      throw new Error("Office.js が利用できません");
    }
    
    if (!Excel || !Excel.run) {
      throw new Error("Excel.js APIが利用できません");
    }

    console.log("Starting Excel.run");
    const payload = await Excel.run(async (context) => {
      console.log("Inside Excel.run");
      
      // === 1) WBS テーブルからタスク情報取得 ===
      let wbsSheet, wbsTable, header, body;
      
      try {
        console.log("Getting WBS sheet");
        wbsSheet = context.workbook.worksheets.getItem("WBS");
      } catch (error) {
        throw new Error("WBSシートが見つかりません: " + error.message);
      }
      
      try {
        console.log("Getting WBS table");
        wbsTable = wbsSheet.tables.getItem("tblWBS");
      } catch (error) {
        throw new Error("tblWBSテーブルが見つかりません: " + error.message);
      }

      header = wbsTable.getHeaderRowRange();
      body   = wbsTable.getDataBodyRange();

      header.load("values");
      body.load("values");

      // === 2) コードシートから担当者候補を取得 ===
      let codeSheet, assigneeTable, assigneeBody;
      
      try {
        console.log("Getting Codes sheet");
        codeSheet = context.workbook.worksheets.getItem("Codes");
      } catch (error) {
        console.log("Codesシートが見つかりません、空の担当者リストを使用します");
        // Codesシートが無くても続行
        codeSheet = null;
      }
      
      if (codeSheet) {
        try {
          assigneeTable = codeSheet.tables.getItem("tblAssignee");
          assigneeBody  = assigneeTable.getDataBodyRange();
          assigneeBody.load("values");
        } catch (error) {
          console.log("tblAssigneeテーブルが見つかりません、空の担当者リストを使用します");
          assigneeBody = null;
        }
      }

      console.log("Syncing context");
      await context.sync();
      
      console.log("Processing header data");
      const headers = header.values[0].map(h => String(h).trim());
      console.log("Headers found:", headers);
      
      const col = (name) => {
        const index = headers.indexOf(name);
        if (index === -1) {
          console.warn(`Column '${name}' not found in headers`);
        }
        return index;
      };

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

      console.log("Processing tasks data");
      // タスク一覧の生成
      const tasks = body.values.map((r, index) => {
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

      console.log(`Found ${tasks.length} tasks`);

      // 担当者候補（Codes シート） ※1列目に名前が入っている前提
      let assignees = [];
      if (assigneeBody && assigneeBody.values) {
        assignees = assigneeBody.values
          .map(row => String(row[0]).trim())
          .filter(name => !!name);
      }
      console.log(`Found ${assignees.length} assignees`);

      return {
        tasks,
        assignees
      };
    });

    console.log("Excel.run completed successfully");
    console.log("Payload:", payload);

    // === 3) Dialog を開く ===
    const url = `https://ymatsuda-cmyk.github.io/tools/kanban-addin/src/dialog/kanban.html#data=${encodeURIComponent(JSON.stringify(payload))}`;
    console.log("Opening dialog with URL:", url);

    Office.context.ui.displayDialogAsync(
      url,
      { height: 90, width: 90 },
      (asyncResult) => {
        console.log("Dialog async result:", asyncResult);
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to open dialog:", asyncResult.error);
          alert("ダイアログを開けませんでした: " + (asyncResult.error ? asyncResult.error.message : "不明なエラー"));
          return;
        }
        console.log("Dialog opened successfully");
        dialog = asyncResult.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          onDialogMessage
        );
      }
    );
  } catch (error) {
    console.error("Error in openKanban:", error);
    alert("エラーが発生しました: " + error.message + "\n\nExcelのWBSシートとtblWBSテーブルが存在するか確認してください。");
  } finally {
    console.log("Completing event");
    // ExecuteFunction の必須
    event.completed();
  }
}

// Dialog からのメッセージを受ける
async function onDialogMessage(arg) {
  console.log("Received message from dialog:", arg.message);
  
  try {
    const msg = JSON.parse(arg.message);
    console.log("Parsed message:", msg);

    switch (msg.type) {
      case "move":
        console.log("Processing move message");
        // ステータス変更 → 実績開始日／実績終了日 更新
        await updateActualDatesByStatus(msg.id, msg.status, msg.forceOverwrite);
        break;
      case "edit":
        console.log("Processing edit message");
        // 予定開始日／予定終了日／担当者／備考 の更新
        await updateTaskDetails(msg);
        break;
      default:
        console.log("Unknown message type:", msg.type);
        break;
    }
  } catch (error) {
    console.error("Error processing dialog message:", error);
    alert("ダイアログからのメッセージ処理でエラーが発生しました: " + error.message);
  }
}

// status に応じて 実績開始日／実績終了日 を更新
async function updateActualDatesByStatus(id, newStatus, forceOverwrite) {
  console.log(`Updating actual dates for task ${id} to status ${newStatus}`);
  
  try {
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
      if (rowIndex < 0) {
        console.error(`Task with ID ${id} not found`);
        return;
      }

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
    console.log("Actual dates updated successfully");
  });
  } catch (error) {
    console.error("Error updating actual dates:", error);
    alert("実績日時の更新でエラーが発生しました: " + error.message);
  }
}

// 予定開始日／予定終了日／担当者／備考 の更新
async function updateTaskDetails(msg) {
  console.log("Updating task details:", msg);
  
  try {
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
    if (rowIndex < 0) {
      console.error(`Task with ID ${msg.id} not found for update`);
      return;
    }

    const rowRange = body.getRow(rowIndex);

    rowRange.getCell(0, assigneeCol).values     = [[msg.assignee || ""]];
    rowRange.getCell(0, plannedStartCol).values = [[msg.plannedStart || ""]];
    rowRange.getCell(0, plannedEndCol).values   = [[msg.plannedEnd || ""]];
    rowRange.getCell(0, noteCol).values         = [[msg.note || ""]];

    await context.sync();
    console.log("Task details updated successfully");
  });
  } catch (error) {
    console.error("Error updating task details:", error);
    alert("タスク詳細の更新でエラーが発生しました: " + error.message);
  }
}