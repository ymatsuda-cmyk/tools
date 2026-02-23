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
      
      // テーブルの取得または作成
      console.log("Getting or creating table");
      if (tables.items.length > 0) {
        // 最初のテーブルを使用
        wbsTable = tables.items[0];
        console.log("Using existing table:", wbsTable.name);
      } else {
        // テーブルが存在しない場合、使用されている範囲から新しいテーブルを作成
        console.log("No tables found in sheet. Creating a new table from used range.");
        try {
          const usedRange = wbsSheet.getUsedRange();
          usedRange.load("address");
          await context.sync();
          
          if (!usedRange || !usedRange.address) {
            throw new Error("シートにデータが存在しません。少なくともヘッダー行を含むデータを入力してください。");
          }
          
          console.log("Used range is:", usedRange.address);
          // ヘッダーがあることを前提としてテーブルを作成
          wbsTable = wbsSheet.tables.add(usedRange, true /*hasHeaders*/);
          wbsTable.name = "AutoTable_" + Date.now(); // ユニークなテーブル名
          console.log("New table created:", wbsTable.name);
          
          // テーブル作成直後に同期して確実に利用可能にする
          await context.sync();
        } catch (rangeError) {
          console.error("Failed to create table from used range:", rangeError.message);
          throw new Error("シートからテーブルを作成できませんでした。シートにヘッダー行を含むデータが正しく入力されているか確認してください。");
        }
      }

      header = wbsTable.getHeaderRowRange();
      body   = wbsTable.getDataBodyRange();

      header.load("values");
      body.load("values");
      
      // テーブルの値を確実にロードするため、ここで一度同期
      try {
        await context.sync();
      } catch (syncError) {
        console.error("Failed to sync table data:", syncError.message);
        throw new Error("テーブルデータの読み込みに失敗しました。テーブル構造を確認してください。");
      }

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
          
          if (codesTables.items.length > 0) {
            // 最初のテーブルを使用
            assigneeTable = codesTables.items[0];
            assigneeBody = assigneeTable.getDataBodyRange();
            assigneeBody.load("values");
            console.log("Using table in Codes sheet:", assigneeTable.name);
          } else {
              console.log("No tables found in Codes sheet, attempting to create one.");
              try {
                const usedRange = codeSheet.getUsedRange();
                usedRange.load("address");
                await context.sync();
                if (usedRange.address) {
                  console.log("Used range in Codes sheet is:", usedRange.address);
                  assigneeTable = codeSheet.tables.add(usedRange, true /*hasHeaders*/);
                  assigneeTable.name = "AssigneeTable_" + Date.now();
                  assigneeBody = assigneeTable.getDataBodyRange();
                  assigneeBody.load("values");
                  console.log("New table created in Codes sheet:", assigneeTable.name);
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
      let headers = [];
      if (header && header.values && Array.isArray(header.values) && header.values.length > 0 && Array.isArray(header.values[0])) {
        headers = header.values[0].map(h => String(h).trim());
      } else {
        console.error("Header row is missing or invalid. header.values:", header && header.values);
        throw new Error("テーブルのヘッダー行が取得できません。Excelシートとテーブル構造を確認してください。");
      }
      console.log("Headers found:", headers);

      // ボディデータの確認
      if (!body || !body.values || !Array.isArray(body.values)) {
        console.warn("Table body is empty or invalid, creating empty task list");
        return {
          tasks: [],
          assignees: []
        };
      }
      
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
    console.error("エラーが発生しました: " + error.message + "\n\nExcelシートにヘッダー行を含むデータが正しく入力されているか確認してください。");
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
      // WBSシートを柔軟に検索
      let sheet;
      try {
        sheet = context.workbook.worksheets.getItem("WBS");
      } catch (error) {
        // WBSシートが見つからない場合、アクティブシートを使用
        sheet = context.workbook.worksheets.getActiveWorksheet();
      }
      
      // テーブルを取得（最初のテーブルを使用）
      const tables = sheet.tables;
      tables.load("items/name");
      await context.sync();
      
      if (tables.items.length === 0) {
        throw new Error("更新対象のテーブルが見つかりません");
      }
      
      const table = tables.items[0]; // 最初のテーブルを使用
      console.log("Using table for update:", table.name);
      const header = table.getHeaderRowRange();
      const body   = table.getDataBodyRange();

      header.load("values");
      body.load("values");
      await context.sync();

      const headers = header.values[0].map(h => String(h).trim());
      
      // より柔軟な列名検索関数
      const findCol = (possibleNames) => {
        for (const name of possibleNames) {
          const index = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
          if (index !== -1) return index;
        }
        return -1;
      };

      const idCol          = findCol(["ID", "番号", "No", "識別子"]);
      const actualStartCol = findCol(["実績開始日", "actual start", "開始実績", "実際開始"]);
      const actualEndCol   = findCol(["実績終了日", "actual end", "終了実績", "実際終了"]);

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
    // WBSシートを柔軟に検索
    let sheet;
    try {
      sheet = context.workbook.worksheets.getItem("WBS");
    } catch (error) {
      // WBSシートが見つからない場合、アクティブシートを使用
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }
    
    // テーブルを取得（最初のテーブルを使用）
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();
    
    if (tables.items.length === 0) {
      throw new Error("更新対象のテーブルが見つかりません");
    }
    
    const table = tables.items[0]; // 最初のテーブルを使用
    console.log("Using table for update:", table.name);
    const header = table.getHeaderRowRange();
    const body   = table.getDataBodyRange();

    header.load("values");
    body.load("values");
    await context.sync();

    const headers = header.values[0].map(h => String(h).trim());
    
    // より柔軟な列名検索関数
    const findCol = (possibleNames) => {
      for (const name of possibleNames) {
        const index = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
        if (index !== -1) return index;
      }
      return -1;
    };

    const idCol           = findCol(["ID", "番号", "No", "識別子"]);
    const assigneeCol     = findCol(["担当者", "assignee", "assigned", "担当"]);
    const plannedStartCol = findCol(["予定開始日", "planned start", "start date", "開始予定"]);
    const plannedEndCol   = findCol(["予定終了日", "planned end", "end date", "終了予定"]);
    const noteCol         = findCol(["備考", "note", "notes", "コメント", "メモ"]);

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