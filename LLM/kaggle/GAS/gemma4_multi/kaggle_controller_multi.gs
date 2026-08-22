// ============================================================
// Kaggle マルチエンドポイント コントローラー — Google Apps Script
// ============================================================
// 【スクリプトプロパティ】
//   CONTROL_TOKEN   : openssl rand -hex 16
//   SPREADSHEET_ID  : 実行ログのスプレッドシート
//   GAS_WEBAPP_URL  : デプロイ後の /exec URL
//   ENDPOINTS       : 下記の JSON 配列（1行でよい）
//
// [
//   {"id":"a","label":"Gemma 4 12B","user":"ymatsuda2025",
//    "key":"KaggleのAPIキー","kernel":"gemma4-12b","title":"gemma4 12b",
//    "model":"gemma4:12b","ngrok":"xxx.ngrok-free.dev","proxyKey":"...",
//    "datasets":["ymatsuda2025/ollama-gemma4-12b-v2","ymatsuda2025/ngrok-binary"]},
//   {"id":"b", ... },
//   {"id":"c", ... }
// ]
//
// 【シート】実行ログ: 起動日時 / 停止日時 / 起動時間(分) / エンドポイント
//   4列目が空の既存行は先頭エンドポイントのものとして集計する。
//
// 【今回の変更点（シングル版からの統合）】
//   1. 認証不要の ping アクションを追加（疎通確認用。機密は返さない）
//   2. 週次集計に異常値除外（12時間超の行はログに残して無視）と
//      NaN ガードを追加。堅牢化のみで挙動は変えていない。
//   3. 停止フラグの読み書きを setStopRequested / isStopRequested に集約。
//      キー名は STOP_REQUESTED_<id> の1系統のみ。
// ============================================================

var PROPS              = PropertiesService.getScriptProperties();
var RETRY_COUNT        = 3;
var RETRY_INTERVAL_MS  = 3000;
var SHEET_NAME         = "実行ログ";
var WEEKLY_LIMIT_HOURS = 30;          // 参考値。実際は変動枠
var HEARTBEAT_STALE_MS = 3 * 60000;   // 3分途絶で「応答なし」とみなす
var STATUS_CACHE_SEC   = 10;
var ABNORMAL_MIN       = 720;         // 12時間超の起動時間は記録ミスとして除外

// ============================================================
// エンドポイント定義
// ============================================================

function getEndpoints() {
  var raw = PROPS.getProperty("ENDPOINTS");
  if (!raw) throw new Error("ENDPOINTS が未設定です");
  var list;
  try { list = JSON.parse(raw); }
  catch (e) { throw new Error("ENDPOINTS の JSON が壊れています: " + e.message); }
  if (!list || !list.length) throw new Error("ENDPOINTS が空です");
  return list;
}

function getEndpoint(id) {
  var list = getEndpoints();
  if (!id) return list[0];
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === id) return list[i];
  }
  throw new Error("未知のエンドポイント: " + id);
}

function authOf(e) {
  if (!e.key) throw new Error(e.id + " の key が未設定です");
  return "Basic " + Utilities.base64Encode(e.user + ":" + e.key);
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
      var res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() >= 500) {
        lastError = { code: res.getResponseCode(), body: res.getContentText().substring(0, 200) };
        if (i < RETRY_COUNT - 1) Utilities.sleep(RETRY_INTERVAL_MS);
        continue;
      }
      return res;
    } catch (err) {
      lastError = { code: 0, body: err.message };
      if (i < RETRY_COUNT - 1) Utilities.sleep(RETRY_INTERVAL_MS);
    }
  }
  throw new Error("リトライ失敗: " + JSON.stringify(lastError));
}

// ============================================================
// 停止フラグ（読み書きをここに集約する）
// ============================================================

function stopKey(id) { return "STOP_REQUESTED_" + id; }

function setStopRequested(id, on) {
  if (on) PROPS.setProperty(stopKey(id), "1");
  else PROPS.deleteProperty(stopKey(id));
}

function isStopRequested(id) {
  return PROPS.getProperty(stopKey(id)) === "1";
}

function heartbeatKey(id) { return "LAST_HEARTBEAT_" + id; }

// ============================================================
// Notebook ソース
// ============================================================

function getNotebookSource(e) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get("nbsrc");
  if (!raw) {
    raw = HtmlService.createHtmlOutputFromFile("proxy_py").getContent();
    try { cache.put("nbsrc", raw, 300); } catch (err) {}
  }

  var gasUrl = PROPS.getProperty("GAS_WEBAPP_URL") || ScriptApp.getService().getUrl();

  var py = raw
    .replace("__ENDPOINT_ID__",   e.id)
    .replace("__MODEL__",         e.model)
    .replace("__PROXY_KEY__",     e.proxyKey)
    .replace("__NGROK_DOMAIN__",  e.ngrok)
    .replace("__GAS_URL__",       gasUrl)
    .replace("__CONTROL_TOKEN__", requireProp("CONTROL_TOKEN"));

  var left = py.match(/__[A-Z_]+__/);
  if (left) throw new Error("未置換のプレースホルダ: " + left[0]);

  var lines = py.split("\n").map(function (l, i, a) {
    return i === a.length - 1 ? l : l + "\n";
  });

  return JSON.stringify({
    cells: [{
      cell_type: "code", execution_count: null,
      metadata: { trusted: true }, outputs: [], source: lines
    }],
    metadata: {
      kaggle: {
        accelerator: "gpu_t4_x2",
        dataSources: (e.datasets || []).map(function (d) {
          return { datasetId: d, sourceType: "dataset" };
        }),
        dockerImageVersionId: 28755,
        isGpuEnabled: true, isInternetEnabled: true,
        language: "python", sourceType: "notebook"
      },
      kernelspec: { display_name: "Python 3", language: "python", name: "python3" },
      language_info: { name: "python", version: "3.12.13" }
    },
    nbformat: 4, nbformat_minor: 4
  });
}

// ============================================================
// スプレッドシート
// ============================================================

function getSheet() {
  var ss = SpreadsheetApp.openById(requireProp("SPREADSHEET_ID"));
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["起動日時", "停止日時", "起動時間(分)", "エンドポイント"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  return sheet;
}

// Kaggle のクォータは土曜朝（UTC）にリセットされる
function getWeekStartUtc() {
  var now = new Date();
  var d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() - 6 + 7) % 7));
  return d;
}

// 異常値（12時間超）を除外し、NaN を無視する堅牢化版
function getWeeklyUsedMinutes(id) {
  var data = getSheet().getDataRange().getValues();
  var since = getWeekStartUtc();
  var now = new Date();
  var fallback = getEndpoints()[0].id;   // 4列目が空の旧データ用
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var startTs = data[i][0];
    if (!startTs) continue;

    var rowId = data[i][3] || fallback;
    if (rowId !== id) continue;

    var start = new Date(startTs);
    if (isNaN(start.getTime())) continue;
    if (start < since) continue;

    var mins = data[i][2];

    if (mins !== "" && mins !== null && !isNaN(Number(mins))) {
      var m = Number(mins);
      if (m > ABNORMAL_MIN) {
        Logger.log("異常値を除外 id=" + id + " row=" + (i + 1) + " mins=" + m);
        continue;
      }
      total += m;
      continue;
    }

    // 未終了行（稼働中）は経過分を暫定加算する
    var running = Math.round((now - start) / 60000);
    if (running > 0) total += running;
  }
  return total;
}

function closeOpenRow(id) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var fallback = getEndpoints()[0].id;

  for (var i = data.length - 1; i >= 1; i--) {
    var rowId = data[i][3] || fallback;
    if (data[i][0] && !data[i][1] && rowId === id) {
      var stop = new Date();
      var mins = Math.round((stop - new Date(data[i][0])) / 60000);
      sheet.getRange(i + 1, 2).setValue(stop.toISOString());
      sheet.getRange(i + 1, 3).setValue(mins);
      return { durationMin: mins, stoppedAt: stop.toISOString() };
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
  var p = (e && e.parameter) || {};
  var action = p.action || "statusAll";
  var result;

  try {
    if (!PUBLIC_ACTIONS[action] && p.token !== requireProp("CONTROL_TOKEN")) {
      result = { success: false, error: "unauthorized" };
    } else if (action === "ping")       { result = { success: true, pong: new Date().toISOString() };
    } else if (action === "statusAll")  { result = handleStatusAll();
    } else if (action === "status")     { result = statusOf(getEndpoint(p.id));
    } else if (action === "start")      { result = handleStart(p.id);
    } else if (action === "stop")       { result = handleStop(p.id);
    } else if (action === "started")    { result = handleStarted(p.id);
    } else if (action === "shouldStop") { result = handleShouldStop(p.id);
    } else if (action === "record")     { result = handleRecord(p.id);
    } else {
      result = { success: false, error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { success: false, error: String(err && err.message ? err.message : err) };
  }

  var json = JSON.stringify(result);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + "(" + json + ")")
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

function statusOf(e) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "kst_" + e.id;
  var kaggleStatus = cache.get(cacheKey);

  if (!kaggleStatus) {
    var url = "https://www.kaggle.com/api/v1/kernels/status"
      + "?userName=" + encodeURIComponent(e.user)
      + "&kernelSlug=" + encodeURIComponent(e.kernel);
    var res = fetchWithRetry(url, {
      method: "GET",
      headers: { "Authorization": authOf(e) },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      return {
        id: e.id, label: e.label, model: e.model, success: false,
        error: "ステータス取得失敗 (HTTP " + res.getResponseCode() + ")"
      };
    }
    kaggleStatus = JSON.parse(res.getContentText()).status || "unknown";
    try { cache.put(cacheKey, kaggleStatus, STATUS_CACHE_SEC); } catch (err) {}
  }

  var hb = PROPS.getProperty(heartbeatKey(e.id));
  var used = getWeeklyUsedMinutes(e.id);

  return {
    id: e.id,
    label: e.label,
    model: e.model,
    success: true,
    status: kaggleStatus,
    proxyAlive: hb ? (new Date() - new Date(hb)) < HEARTBEAT_STALE_MS : false,
    lastHeartbeat: hb || null,
    stopRequested: isStopRequested(e.id),
    baseUrl: "https://" + e.ngrok + "/v1",
    weeklyUsedMin: used,
    weeklyRemainMin: Math.max(0, WEEKLY_LIMIT_HOURS * 60 - used),
    weeklyLimitMin: WEEKLY_LIMIT_HOURS * 60
  };
}

function handleStatusAll() {
  var out = getEndpoints().map(function (e) {
    try { return statusOf(e); }
    catch (err) {
      return { id: e.id, label: e.label, model: e.model, success: false, error: err.message };
    }
  });
  return {
    success: true,
    weekStartUtc: getWeekStartUtc().toISOString(),
    endpoints: out
  };
}

function handleStart(id) {
  var e = getEndpoint(id);

  // 前回の停止要求が残っていると起動直後に自滅するため、必ず先にクリアする
  setStopRequested(e.id, false);

  var st = statusOf(e);
  if (st.success && (st.status === "running" || st.status === "queued")) {
    return { success: false, id: e.id, message: "すでに起動中です (" + st.status + ")" };
  }
  if (st.success && st.weeklyRemainMin <= 0) {
    return { success: false, id: e.id, message: "今週のクォータを使い切っています" };
  }

  var payload = {
    slug:               e.user + "/" + e.kernel,
    newTitle:           e.title || e.label,
    text:               getNotebookSource(e),
    language:           "python",
    kernelType:         "notebook",
    isPrivate:          true,
    enableGpu:          true,
    enableInternet:     true,
    datasetDataSources: e.datasets || []
  };

  var res = fetchWithRetry("https://www.kaggle.com/api/v1/kernels/push", {
    method: "POST",
    headers: { "Authorization": authOf(e), "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var out = {};
  try { out = JSON.parse(res.getContentText()); }
  catch (err) { out = { raw: res.getContentText().substring(0, 300) }; }

  if (res.getResponseCode() !== 200 || out.hasError) {
    return {
      success: false, id: e.id,
      message: "起動失敗: " + (out.error || out.errorNullable || "HTTP " + res.getResponseCode()),
      detail: out
    };
  }

  PROPS.deleteProperty(heartbeatKey(e.id));
  CacheService.getScriptCache().remove("kst_" + e.id);
  getSheet().appendRow([new Date().toISOString(), "", "", e.id]);

  return {
    success: true, id: e.id,
    message: "起動リクエスト送信完了 (v" + out.versionNumber + ")"
  };
}

// 停止フラグを立てる。Notebook が次のポーリングで拾って自ら終了する
function handleStop(id) {
  var e = getEndpoint(id);
  setStopRequested(e.id, true);
  return {
    success: true, id: e.id,
    message: "停止をリクエストしました（30秒以内に停止します）",
    fallbackUrl: "https://www.kaggle.com/code/" + e.user + "/" + e.kernel
  };
}

function handleStarted(id) {
  var e = getEndpoint(id);
  PROPS.setProperty(heartbeatKey(e.id), new Date().toISOString());
  setStopRequested(e.id, false);
  return { success: true, id: e.id };
}

// Notebook からのポーリング。ハートビートを兼ねる
function handleShouldStop(id) {
  var e = getEndpoint(id);
  PROPS.setProperty(heartbeatKey(e.id), new Date().toISOString());
  return { success: true, id: e.id, stop: isStopRequested(e.id) };
}

function handleRecord(id) {
  var e = getEndpoint(id);
  setStopRequested(e.id, false);
  PROPS.deleteProperty(heartbeatKey(e.id));
  CacheService.getScriptCache().remove("kst_" + e.id);

  var closed = closeOpenRow(e.id);
  if (!closed) return { success: false, id: e.id, message: "未クローズの行がありません" };
  return { success: true, id: e.id, durationMin: closed.durationMin };
}

// ============================================================
// テスト関数（エディタから直接実行する用）
// ============================================================

function testStatusAll() { Logger.log(JSON.stringify(handleStatusAll(), null, 2)); }
function testPing()      { Logger.log(JSON.stringify({ success: true, pong: new Date().toISOString() })); }

function testEndpoints() {
  getEndpoints().forEach(function (e) {
    var miss = ["id", "label", "user", "key", "kernel", "model", "ngrok", "proxyKey", "datasets"]
      .filter(function (k) { return !e[k]; });
    Logger.log(e.id + " : " + (miss.length ? "★不足 " + miss.join(", ") : "OK"));
  });
}

function testNotebookSource() {
  getEndpoints().forEach(function (e) {
    var src = getNotebookSource(e);
    Logger.log(e.id + " : " + src.length + " bytes, JSON " + (JSON.parse(src) ? "OK" : "NG"));
  });
}

function testProps() {
  ["CONTROL_TOKEN", "SPREADSHEET_ID", "ENDPOINTS"].forEach(function (k) {
    Logger.log(k + ": " + (PROPS.getProperty(k) ? "OK" : "★未設定"));
  });
  Logger.log("GAS_WEBAPP_URL: " + (PROPS.getProperty("GAS_WEBAPP_URL") || "(未設定→getUrl()を使用)"));
}

// 停止フラグの上げ下げが噛み合っているかを単体で確認する
function testStopFlagCycle() {
  var id = getEndpoints()[0].id;
  setStopRequested(id, true);
  Logger.log("stop=true 直後: " + JSON.stringify(handleShouldStop(id)));
  setStopRequested(id, false);
  Logger.log("stop=false 直後: " + JSON.stringify(handleShouldStop(id)));
}

// 停止フラグが立ちっぱなしのときに手動で解除する
function clearStopFlag(id) {
  var target = id || getEndpoints()[0].id;
  setStopRequested(target, false);
  Logger.log(target + " の STOP_REQUESTED を解除しました");
}

function testWeekly() {
  Logger.log("週起点(UTC): " + getWeekStartUtc().toISOString());
  getEndpoints().forEach(function (e) {
    Logger.log(e.id + " 今週使用: " + getWeeklyUsedMinutes(e.id) + "分");
  });
}

function debugWeekly(id) {
  var target = id || getEndpoints()[0].id;
  var data = getSheet().getDataRange().getValues();
  var since = getWeekStartUtc();
  var now = new Date();
  var fallback = getEndpoints()[0].id;

  Logger.log("id=" + target + " 週起点: " + since.toISOString());

  for (var i = 1; i < data.length; i++) {
    var startTs = data[i][0];
    if (!startTs) continue;
    var rowId = data[i][3] || fallback;
    if (rowId !== target) continue;

    var start = new Date(startTs);
    if (isNaN(start.getTime())) continue;
    if (start < since) continue;

    var mins = data[i][2];
    if (mins) {
      Logger.log("row=" + (i + 1) + " FIXED " + mins + "分");
    } else {
      var running = Math.round((now - start) / 60000);
      Logger.log("row=" + (i + 1) + " OPEN " + running + "分");
    }
  }
}
