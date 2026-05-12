/**
 * Google Apps Script Backend for News PWA
 * 
 * Setup:
 * 1. Open a Google Spreadsheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this code and Save.
 * 4. Deploy > New Deployment > Web App.
 * 5. Execute as: Me, Who has access: Anyone.
 * 6. Copy the Web App URL.
 */

function doGet(e) {
  const date = e.parameter.date;
  if (!date) {
    return createJsonResponse({ error: "Missing date parameter" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == date) {
      return createJsonResponse(JSON.parse(data[i][1]));
    }
  }
  
  return createJsonResponse({ error: "No data found for this date" });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const payload = JSON.parse(e.postData.contents);
    const date = payload.date;
    const content = JSON.stringify(payload);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == date) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow > -1) {
      sheet.getRange(foundRow, 2).setValue(content);
    } else {
      sheet.appendRow([date, content]);
    }
    
    return createJsonResponse({ success: true, date: date });
    
  } catch (err) {
    return createJsonResponse({ error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
