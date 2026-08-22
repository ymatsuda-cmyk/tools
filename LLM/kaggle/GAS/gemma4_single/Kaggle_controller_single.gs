// ============================================================
// Kaggle Notebook コントローラー - Google Apps Script（修正版）
// ============================================================
// 【スクリプトプロパティ】
//   KAGGLE_API_KEY    : Kaggle の API キー
//   KAGGLE_KERNEL     : ymatsuda2025/gemma4-12b
//   SPREADSHEET_ID    : 実行ログを書くスプレッドシートの ID
//   NOTEBOOK_FILE_ID  : Drive 上の gemma4-proxy.ipynb のファイル ID
//   CONTROL_TOKEN     : openssl rand -hex 16 で生成
//   PROXY_API_KEY     : openssl rand -hex 16 で生成（CONTROL_TOKEN とは別値）
//   NGROK_DOMAIN      : pregnant-vindicate-deacon.ngrok-free.dev
//   GAS_WEBAPP_URL    : このウェブアプリの /exec URL（デプロイ後に設定）
//
// 【今回の修正】
//   1. トップレベルに残っていた allMessages 参照コードを削除
//      → ReferenceError で全アクションが落ちていた
//   2. 停止フラグを STOP_REQUESTED に一本化
//      → STOP_FLAG と二系統に分かれていて停止が永久に検知されなかった
//   3. 認証不要の ping アクションを追加（疎通確認用）
// ============================================================

var PROPS              = PropertiesService.getScriptProperties();
var RETRY_COUNT        = 3;
var RETRY_INTERVAL_MS  = 3000;
var SHEET_NAME         = "実行ログ";
var WEEKLY_LIMIT_HOURS = 30;          // 参考値。実際は変動枠
var HEARTBEAT_STALE_MS = 3 * 60000;   // 3分途絶で「応答なし」とみなす
var PROP_STOP          = "STOP_REQUESTED";   // 停止フラグはこの1つだけ

// ============================================================
// 共通
// ============================================================

function getKernel() {
  var k = PROPS.getProperty("KAGGLE_KERNEL");
  if (!k) throw new Error("KAGGLE_KERNEL が未設定です");
  return k;
}

function getAuth() {
  var key  = PROPS.getProperty("KAGGLE_API_KEY");
  if (!key) throw new Error("KAGGLE_API_KEY が未設定です");
  var user = getKernel().split("/")[0];
  return "Basic " + Utilities.base64Encode(user + ":" + key);
}

function requireProp(name) {
  var v = PROPS.getProperty(name);
  if (!v) throw new Error(name + " が未設定です");
  return v;
}

function fetchWithRetry(url, options) {
  var lastError = null;
  for (var i = 0; i < RETRY_COUNT; i++) {
    try {
      var res  = UrlFetchApp.fetch(url, options);
      var code = res.getResponseCode();
      if (code >= 500) {
        lastError = { code: code, body: res.getContentText().substring(0, 300) };
        if (i < RETRY_COUNT - 1) Utilities.sleep(RETRY_INTERVAL_MS);
        continue;
      }
      return res;
    } catch (e) {
      lastError = { code: 0, body: e.message };
      if (i < RETRY_COUNT - 1) Utilities.sleep(RETRY_INTERVAL_MS);
    }
  }
  throw new Error("リトライ失敗: " + JSON.stringify(lastError));
}

// ============================================================
// 停止フラグ（STOP_REQUESTED に一本化）
// ============================================================

function setStopRequested(on) {
  if (on) {
    PROPS.setProperty(PROP_STOP, "1");
  } else {
    PROPS.deleteProperty(PROP_STOP);
  }
}

function isStopRequested() {
  return PROPS.getProperty(PROP_STOP) === "1";
}

// ============================================================
// Notebook ソース（プロジェクト内の proxy_py.html から取得してプレースホルダを差し込む）
// ============================================================
function getNotebookSource() {
  var py = HtmlService.createHtmlOutputFromFile('proxy_py').getContent();

  py = py
    .replace('__PROXY_KEY__',     requireProp('PROXY_API_KEY'))
    .replace('__GAS_URL__',       PROPS.getProperty('GAS_WEBAPP_URL') || ScriptApp.getService().getUrl())
    .replace('__CONTROL_TOKEN__', requireProp('CONTROL_TOKEN'))
    .replace('__NGROK_DOMAIN__',  requireProp('NGROK_DOMAIN'));

  var left = py.match(/__[A-Z_]+__/);
  if (left) throw new Error('未置換のプレースホルダ: ' + left[0]);

  var lines = py.split('\n').map(function (l, i, a) {
    return i === a.length - 1 ? l : l + '\n';
  });

  return JSON.stringify({
    cells: [{
      cell_type: 'code', execution_count: null,
      metadata: { trusted: true }, outputs: [], source: lines
    }],
    metadata: {
      kaggle: {
        accelerator: 'gpu_t4_x2',
        dataSources: [
          { datasetId: 'ymatsuda2025/ollama-gemma4-12b', sourceType: 'dataset' },
          { datasetId: 'ymatsuda2025/ngrok-binary',      sourceType: 'dataset' }
        ],
        dockerImageVersionId: 28755,
        isGpuEnabled: true, isInternetEnabled: true,
        language: 'python', sourceType: 'notebook'
      },
      kernelspec: { display_name: 'Python 3', language: 'python', name: 'python3' },
      language_info: { name: 'python', version: '3.12.13' }
    },
    nbformat: 4, nbformat_minor: 4
  });
}

// ============================================================
// スプレッドシート
// ============================================================

function getSheet() {
  var ss    = SpreadsheetApp.openById(requireProp("SPREADSHEET_ID"));
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["起動日時", "停止日時", "起動時間(分)"]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  }
  return sheet;
}

// Kaggle のクォータは土曜朝（UTC）にリセットされるため、土曜 00:00 UTC を週の起点とする
function getWeekStartUtc() {
  var now = new Date();
  var d   = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  var back = (d.getUTCDay() - 6 + 7) % 7;   // 土曜 = 6
  d.setUTCDate(d.getUTCDate() - back);
  return d;
}

function getWeeklyUsedMinutes() {
  var data  = getSheet().getDataRange().getValues();
  var since = getWeekStartUtc();
  var now   = new Date();
  var total = 0;

  for (var i = 1; i < data.length; i++) {

    var startTs = data[i][0];
    var mins    = data[i][2];

    if (!startTs) continue;

    var start = new Date(startTs);

    if (isNaN(start.getTime())) continue;

    // 今週以前は無視
    if (start < since) continue;

    // 正常終了済み
    if (
      mins !== "" &&
      mins !== null &&
      !isNaN(Number(mins))
    ) {

      var m = Number(mins);

      // 12時間超は異常値扱い
      if (m > 720) {

        Logger.log(
          "異常値を除外 row="
          + (i + 1)
          + " mins="
          + m
        );

        continue;
      }

      total += m;
      continue;
    }

    // 未終了行
    // 今週開始以降のものだけ暫定加算
    var runningMin =
      Math.round((now - start) / 60000);

    if (runningMin > 0) {
      total += runningMin;
    }
  }

  return total;
}

// 停止日時が空の最終行を閉じる
function closeOpenRow() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && !data[i][1]) {
      var start = new Date(data[i][0]);
      var stop  = new Date();
      var mins  = Math.round((stop - start) / 60000);
      sheet.getRange(i + 1, 2).setValue(stop.toISOString());
      sheet.getRange(i + 1, 3).setValue(mins);
      return { row: i + 1, durationMin: mins, stoppedAt: stop.toISOString() };
    }
  }
  return null;
}

// ============================================================
// ルーティング
// ============================================================

// 認証不要のアクション（疎通確認用。機密情報は返さない）
var PUBLIC_ACTIONS = { ping: true };

function doGet(e) {
  var p        = (e && e.parameter) || {};
  var callback = p.callback;
  var action   = p.action || "status";
  var result;

  try {
    if (!PUBLIC_ACTIONS[action] && p.token !== requireProp("CONTROL_TOKEN")) {
      result = { success: false, error: "unauthorized" };
    } else if (action === "ping")       { result = { success: true, pong: new Date().toISOString() };
    } else if (action === "status")     { result = handleStatus();
    } else if (action === "start")      { result = handleStart();
    } else if (action === "stop")       { result = handleStop();
    } else if (action === "started")    { result = handleStarted();
    } else if (action === "shouldStop") { result = handleShouldStop();
    } else if (action === "record")     { result = handleRecord();
    } else {
      result = { success: false, error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { success: false, error: String(err && err.message ? err.message : err) };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || "{}"); } catch (err) {}
  return doGet({ parameter: body });
}

// ============================================================
// ハンドラー
// ============================================================

function handleStatus() {
  var kernel = getKernel();
  var parts  = kernel.split("/");
  var url    = "https://www.kaggle.com/api/v1/kernels/status"
    + "?userName="   + encodeURIComponent(parts[0])
    + "&kernelSlug=" + encodeURIComponent(parts[1]);

  var res  = fetchWithRetry(url, {
    method: "GET",
    headers: { "Authorization": getAuth() },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code !== 200) {
    return { success: false, error: "ステータス取得失敗 (HTTP " + code + ")",
             detail: text.substring(0, 300) };
  }

  var data  = JSON.parse(text);
  var hb    = PROPS.getProperty("LAST_HEARTBEAT");
  var alive = hb ? (new Date() - new Date(hb)) < HEARTBEAT_STALE_MS : false;
  var used  = getWeeklyUsedMinutes();

  return {
    success:         true,
    status:          data.status || "unknown",
    proxyAlive:      alive,                     // Notebook が応答しているか
    lastHeartbeat:   hb || null,
    stopRequested:   isStopRequested(),
    baseUrl:         "https://" + requireProp("NGROK_DOMAIN") + "/v1",
    title:           kernel,
    weeklyUsedMin:   used,
    weeklyRemainMin: Math.max(0, WEEKLY_LIMIT_HOURS * 60 - used),
    weekStartUtc:    getWeekStartUtc().toISOString()
  };
}

function handleStart() {
  // 前回の停止要求が残っていると起動直後に自滅するため、必ず先にクリアする
  setStopRequested(false);

  // 二重起動でクォータを溶かさないためのガード
  var st = handleStatus();
  if (st.success && (st.status === "running" || st.status === "queued")) {
    return { success: false, message: "すでに起動中です (" + st.status + ")", status: st.status };
  }
//  if (st.success && st.weeklyRemainMin <= 0) {
//    return { success: false, message: "今週のクォータを使い切っています" };
//  }

  var parts     = getKernel().split("/");
  var startTime = new Date().toISOString();

  var payload = {
    slug:               parts[0] + "/" + parts[1],
    newTitle:           "gemma4 12b",
    text:               getNotebookSource(),
    language:           "python",
    kernelType:         "notebook",
    isPrivate:          true,
    enableGpu:          true,
    enableInternet:     true,
    datasetDataSources: ["ymatsuda2025/ollama-gemma4-12b-v2", "ymatsuda2025/ngrok-binary"]
  };

  var res  = fetchWithRetry("https://www.kaggle.com/api/v1/kernels/push", {
    method: "POST",
    headers: { "Authorization": getAuth(), "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var out  = {};
  try { out = JSON.parse(res.getContentText()); }
  catch (e) { out = { raw: res.getContentText().substring(0, 300) }; }

  if (code !== 200 || out.hasError) {
    return {
      success: false,
      message: "起動失敗: " + (out.error || out.errorNullable || "HTTP " + code),
      detail:  out
    };
  }

  PROPS.deleteProperty("LAST_HEARTBEAT");
  getSheet().appendRow([startTime, "", ""]);

  return {
    success:   true,
    message:   "起動リクエスト送信完了 (v" + out.versionNumber + ")",
    startTime: startTime,
    detail:    out
  };
}

// ── 停止要求（Notebook が次のポーリングで拾って自ら終了する） ──
function handleStop() {
  setStopRequested(true);
  return {
    success: true,
    message: "停止をリクエストしました（30秒以内に停止します）"
  };
}

// ── Notebook からの起動通知 ──
function handleStarted() {
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  setStopRequested(false);
  return { success: true };
}

// ── Notebook からのポーリング（ハートビートを兼ねる） ──
function handleShouldStop() {
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  return { success: true, stop: isStopRequested() };
}

// ── Notebook 終了時の記録（停止日時が空の最終行を閉じる） ──
function handleRecord() {
  setStopRequested(false);
  PROPS.deleteProperty("LAST_HEARTBEAT");

  var closed = closeOpenRow();
  if (!closed) return { success: false, message: "未クローズの行がありません" };
  return { success: true, durationMin: closed.durationMin, stoppedAt: closed.stoppedAt };
}

// ============================================================
// テスト関数（エディタから直接実行する用）
// ============================================================

function testStatus() { Logger.log(JSON.stringify(handleStatus(), null, 2)); }
function testStart()  { Logger.log(JSON.stringify(handleStart(),  null, 2)); }
function testStop()   { Logger.log(JSON.stringify(handleStop(),   null, 2)); }

// 停止フラグの上げ下げが噛み合っているかを単体で確認する
function testStopFlagCycle() {
  setStopRequested(true);
  Logger.log("stop=true にした直後 shouldStop: "
             + JSON.stringify(handleShouldStop()));
  setStopRequested(false);
  Logger.log("stop=false にした直後 shouldStop: "
             + JSON.stringify(handleShouldStop()));
}

// 停止フラグが立ちっぱなしのときに手動で解除する
function clearStopFlag() {
  setStopRequested(false);
  Logger.log("STOP_REQUESTED を解除しました");
}

function testWeekly() {
  Logger.log("週起点(UTC): " + getWeekStartUtc().toISOString());
  Logger.log("今週使用: " + getWeeklyUsedMinutes() + "分");
}

// push する前に Notebook ソースが正しく組み立つか確認する
function testNotebookSource() {
  var src = getNotebookSource();
  Logger.log("length: " + src.length);
  Logger.log("残プレースホルダ: " + (src.match(/__[A-Z_]+__/g) || []).join(", "));
  Logger.log("JSON: " + (JSON.parse(src) ? "OK" : "NG"));
}

// 必須プロパティが揃っているか確認する
function testProps() {
  var required = ["KAGGLE_API_KEY", "KAGGLE_KERNEL", "SPREADSHEET_ID",
                  "NOTEBOOK_FILE_ID", "CONTROL_TOKEN", "PROXY_API_KEY", "NGROK_DOMAIN"];
  required.forEach(function (k) {
    Logger.log(k + ": " + (PROPS.getProperty(k) ? "OK" : "★未設定"));
  });
  Logger.log("GAS_WEBAPP_URL: " + (PROPS.getProperty("GAS_WEBAPP_URL") || "(未設定→getUrl()を使用)"));
}

function checkKeyLength() {
  var k = PROPS.getProperty("PROXY_API_KEY");
  Logger.log("length: " + k.length);
  Logger.log("非ASCII: " + (/[^\x21-\x7E]/.test(k) ? "あり" : "なし"));
  Logger.log("前後の空白: " + (k !== k.trim() ? "あり" : "なし"));
}

// ============================================================
// 参考：トークン数の概算（このスクリプトからは未使用）
// もとのコードではこの下にトップレベルの実行文があり、
// allMessages が未定義のため全リクエストが ReferenceError で落ちていた。
// 関数定義だけ残し、実行文は削除している。
// ============================================================

function estimateTokens(text) {
  var cjk = 0, other = 0;
  for (var i = 0; i < text.length; i++) {
    if (/[\u3000-\u9FFF\uFF00-\uFFEF]/.test(text[i])) cjk++; else other++;
  }
  return Math.ceil(cjk * 0.7 + other * 0.5);   // 安全側に振った係数
}

function checkToken() {
  var t = PropertiesService.getScriptProperties().getProperty('CONTROL_TOKEN');
  Logger.log('length: ' + (t ? t.length : '未設定'));
}

function checkTokenExact() {
  var t = PropertiesService.getScriptProperties().getProperty('CONTROL_TOKEN');
  Logger.log(JSON.stringify(t));
}

function debugWeekly() {

  var data  = getSheet().getDataRange().getValues();
  var since = getWeekStartUtc();
  var now   = new Date();

  Logger.log("週起点: " + since.toISOString());

  for (var i = 1; i < data.length; i++) {

    var startTs = data[i][0];
    var mins    = data[i][2];

    if (!startTs) continue;

    var start = new Date(startTs);

    if (isNaN(start.getTime())) continue;

    if (start < since) continue;

    if (mins) {

      Logger.log(
        "row=" + (i + 1) +
        " FIXED " +
        mins + "分"
      );

    } else {

      var running =
        Math.round((now - start) / 60000);

      Logger.log(
        "row=" + (i + 1) +
        " OPEN " +
        running + "分"
      );
    }
  }
}