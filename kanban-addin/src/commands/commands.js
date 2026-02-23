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
      
      // まず現在のワークブックの構造を確認
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      
      console.log("Available worksheets:", worksheets.items.map(ws => ws.name));
      
      // === 1) WBS テーブルからタスク情報取得 ===
      let wbsSheet, wbsTable, header, body;
      
      try {
        console.log("Getting WBS sheet");
        wbsSheet = context.workbook.worksheets.getItem("WBS");
      } catch (error) {
        // WBSシートが見つからない場合、現在のアクティブシートを使用するか確認
        console.warn("WBS sheet not found, checking current active sheet");
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        await context.sync();
        
        const sheetName = activeSheet.name;
        console.log("Current active sheet:", sheetName);
        
        // ユーザーに確認を求める（Office Add-in対応）
        console.log(`WBS sheet not found, using current active sheet: ${sheetName}`);
        wbsSheet = activeSheet;
      }
      
      // テーブルの確認
      const tables = wbsSheet.tables;
      tables.load("items/name");
      await context.sync();
      
      console.log("Available tables in sheet:", tables.items.map(t => t.name));
      
      try {
        console.log("Getting WBS table");
        wbsTable = wbsSheet.tables.getItem("tblWBS");
      } catch (error) {
        // tblWBSテーブルが見つからない場合、最初のテーブルを自動使用
        if (tables.items.length > 0) {
          const tableName = tables.items[0].name;
          console.log(`tblWBS not found, using first available table: ${tableName}`);
          wbsTable = tables.items[0];
        } else {
          // テーブルが存在しない場合、使用されている範囲から新しいテーブルを作成
          console.log("No tables found in sheet. Attempting to create a new table from used range.");
          const usedRange = wbsSheet.getUsedRange();
          usedRange.load("address");
          await context.sync();
          console.log("Used range is:", usedRange.address);
          // ヘッダーがあることを前提としてテーブルを作成
          wbsTable = wbsSheet.tables.add(usedRange, true /*hasHeaders*/);
          wbsTable.name = "tblWBS_auto";
          console.log("New table 'tblWBS_auto' created.");
        }
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
        console.log("Codes sheet not found, will use empty assignee list");
        codeSheet = null;
      }
      
      if (codeSheet) {
        try {
          const codesTables = codeSheet.tables;
          codesTables.load("items/name");
          await context.sync();
          
          console.log("Available tables in Codes sheet:", codesTables.items.map(t => t.name));
          
          try {
            assigneeTable = codeSheet.tables.getItem("tblAssignee");
            assigneeBody  = assigneeTable.getDataBodyRange();
            assigneeBody.load("values");
          } catch (error) {
            console.log("tblAssignee table not found in Codes sheet");
            if (codesTables.items.length > 0) {
              console.log("Using first available table in Codes sheet:", codesTables.items[0].name);
              assigneeTable = codesTables.items[0];
              assigneeBody = assigneeTable.getDataBodyRange();
              assigneeBody.load("values");
            } else {
              console.log("No tables found in Codes sheet, attempting to create one.");
              try {
                const usedRange = codeSheet.getUsedRange();
                usedRange.load("address");
                await context.sync();
                if (usedRange.address) {
                  console.log("Used range in Codes sheet is:", usedRange.address);
                  assigneeTable = codeSheet.tables.add(usedRange, true /*hasHeaders*/);
                  assigneeTable.name = "tblAssignee_auto";
                  assigneeBody = assigneeTable.getDataBodyRange();
                  assigneeBody.load("values");
                  console.log("New table 'tblAssignee_auto' created in Codes sheet.");
                } else {
                   console.log("Codes sheet is empty, will use empty assignee list");
                   assigneeBody = null;
                }
              } catch (usedRangeError) {
                console.log("Failed to get used range in Codes sheet:", usedRangeError.message);
                assigneeBody = null;
              }
            }
          }
        } catch (tablesError) {
          console.log("Failed to access tables in Codes sheet:", tablesError.message);
          console.log("Will use empty assignee list");
          assigneeBody = null;
        }
      }

      console.log("Syncing context");
      await context.sync();
      
      console.log("Processing header data");
      const headers = header.values[0].map(h => String(h).trim());
      console.log("Headers found:", headers);
      
      // より柔軟な列名検索関数
      const findCol = (possibleNames) => {
        for (const name of possibleNames) {
          const index = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
          if (index !== -1) {
            console.log(`Found column '${headers[index]}' for pattern '${name}'`);
            return index;
          }
        }
        console.warn(`No column found for patterns: ${possibleNames.join(', ')}`);
        return -1;
      };

      // 複数のパターンで列を検索
      const idCol           = findCol(["ID", "番号", "No", "識別子"]);
      const titleCol        = findCol(["task", "タスク", "作業", "項目", "件名", "内容"]);
      const assigneeCol     = findCol(["担当者", "assignee", "assigned", "担当"]);
      const plannedStartCol = findCol(["予定開始日", "planned start", "start date", "開始予定"]);
      const plannedEndCol   = findCol(["予定終了日", "planned end", "end date", "終了予定"]);
      const actualStartCol  = findCol(["実績開始日", "actual start", "開始実績", "実際開始"]);
      const actualEndCol    = findCol(["実績終了日", "actual end", "終了実績", "実際終了"]);
      const noteCol         = findCol(["備考", "note", "notes", "コメント", "メモ"]);
      const tagLargeCol     = findCol(["大分類", "category", "大カテゴリー", "分類"]);
      const tagSmallCol     = findCol(["小分類", "subcategory", "小カテゴリー", "詳細分類"]);

      console.log("Column mapping:", {
        id: idCol, title: titleCol, assignee: assigneeCol,
        plannedStart: plannedStartCol, plannedEnd: plannedEndCol,
        actualStart: actualStartCol, actualEnd: actualEndCol,
        note: noteCol, tagLarge: tagLargeCol, tagSmall: tagSmallCol
      });

      console.log("Processing tasks data");
      // タスク一覧の生成
      const tasks = body.values.map((r, index) => {
        const actualStart = actualStartCol >= 0 ? r[actualStartCol] : null;
        const actualEnd   = actualEndCol >= 0 ? r[actualEndCol] : null;

        let status;
        if (!actualStart && !actualEnd) {
          status = "Todo";
        } else if (actualStart && !actualEnd) {
          status = "Doing";
        } else if (actualEnd) {
          status = "Done";
        } else {
          status = "Todo"; // デフォルト
        }

        return {
          id:          idCol >= 0 ? r[idCol] : `task-${index + 1}`,
          title:       titleCol >= 0 ? r[titleCol] : `Task ${index + 1}`,
          assignee:    assigneeCol >= 0 ? r[assigneeCol] : "",
          plannedStart: plannedStartCol >= 0 ? r[plannedStartCol] : "",
          plannedEnd:  plannedEndCol >= 0 ? r[plannedEndCol] : "",
          actualStart,
          actualEnd,
          status,
          note:        noteCol >= 0 ? r[noteCol] : "",
          tagLarge:    tagLargeCol >= 0 ? r[tagLargeCol] : "",
          tagSmall:    tagSmallCol >= 0 ? r[tagSmallCol] : "",
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
      { 
        height: 80, 
        width: 80, 
        requireHTTPS: true,
        displayInIframe: false 
      },
      (asyncResult) => {
        console.log("Dialog async result:", asyncResult);
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to open dialog:", asyncResult.error);
          console.error("ダイアログを開けませんでした: " + (asyncResult.error ? asyncResult.error.message : "不明なエラー"));
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
    console.error("エラーが発生しました: " + error.message + "\n\nExcelのWBSシートとtblWBSテーブルが存在するか確認してください。");
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
    console.error("ダイアログからのメッセージ処理でエラーが発生しました: " + error.message);
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
    console.error("実績日時の更新でエラーが発生しました: " + error.message);
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
    console.error("タスク詳細の更新でエラーが発生しました: " + error.message);
  }
}