// ============================================================
// Kaggle Notebook コントローラー - Google Apps Script（単体版・JSON設定対応）
// ============================================================
// このファイルを、エンドポイントの数だけ「別々の GAS プロジェクト」
// にそれぞれデプロイする。コードは共通、スクリプトプロパティだけが
// プロジェクトごとに異なる。
//
// 【推奨: CONFIG プロパティ1つにまとめる】
//   プロパティ名 "CONFIG" に、次の形の JSON を1行で入れる。
//
// {
//   "kaggleUsername": "ymatsuda2025",
//   "kaggleApiKey": "Kaggle の API キー",
//   "kaggleKernel": "ymatsuda2025/gemma4-12b",
//   "kernelTitle": "gemma4 12b",
//   "model": "gemma4:12b",
//   "ngrokDomain": "xxx.ngrok-free.dev",
//   "proxyApiKey": "このエンドポイント専用のキー",
//   "controlToken": "このエンドポイント専用のキー（別値）",
//   "datasets": ["ymatsuda2025/ollama-gemma4-12b-v2", "ymatsuda2025/ngrok-binary"],
//   "gasWebappUrl": "デプロイ後の /exec URL（省略可）"
// }
//
// 【互換: 個別プロパティでも設定できる（CONFIG が無い場合のみ使われる）】
//   KAGGLE_USERNAME / KAGGLE_API_KEY / KAGGLE_KERNEL / KERNEL_TITLE / MODEL /
//   NGROK_DOMAIN / PROXY_API_KEY / CONTROL_TOKEN /
//   DATASETS（JSON配列の文字列）/ GAS_WEBAPP_URL
//
// 【稼働ログ】USAGE_LOG プロパティに JSON で保存する。スプレッドシートは不要。
//   既存シートから移行する場合のみ spreadsheetId を残し migrateFromSheet() を1回実行する。
//
// 【proxy_py.html】マルチ版と同じプレースホルダを使う。
//   ENDPOINT_ID には kaggleKernel の値をそのまま流用する。
// ============================================================

var PROPS              = PropertiesService.getScriptProperties();
var RETRY_COUNT        = 3;
var RETRY_INTERVAL_MS  = 3000;
var SHEET_NAME         = "実行ログ";   // 移行時のみ使用
var WEEKLY_LIMIT_HOURS = 30;
var HEARTBEAT_STALE_MS = 3 * 60000;
var ABNORMAL_MIN       = 720;
var PROP_STOP          = "STOP_REQUESTED";

// ============================================================
// 設定の読み込み（CONFIG 優先、無ければ個別プロパティにフォールバック）
// ============================================================

var CONFIG_CACHE = null;

function loadConfig() {
  if (CONFIG_CACHE) return CONFIG_CACHE;

  var raw = PROPS.getProperty("CONFIG");
  if (raw) {
    var parsed;
    try { parsed = JSON.parse(raw); }
    catch (e) { throw new Error("CONFIG の JSON が壊れています: " + e.message); }

    CONFIG_CACHE = {
      kaggleUsername: parsed.kaggleUsername || "",
      kaggleApiKey:   parsed.kaggleApiKey   || "",
      kaggleKernel:   parsed.kaggleKernel   || "",
      kernelTitle:    parsed.kernelTitle    || "",
      model:          parsed.model          || "",
      ngrokDomain:    parsed.ngrokDomain    || "",
      proxyApiKey:    parsed.proxyApiKey    || "",
      controlToken:   parsed.controlToken   || "",
      spreadsheetId:  parsed.spreadsheetId  || "",
      datasets:       parsed.datasets       || [],
      gasWebappUrl:   parsed.gasWebappUrl   || ""
    };
    return CONFIG_CACHE;
  }

  // 個別プロパティへのフォールバック（旧構成との互換用）
  var datasets = [];
  try { datasets = JSON.parse(PROPS.getProperty("DATASETS") || "[]"); } catch (e) {}

  CONFIG_CACHE = {
    kaggleUsername: PROPS.getProperty("KAGGLE_USERNAME") || "",
    kaggleApiKey:   PROPS.getProperty("KAGGLE_API_KEY")  || "",
    kaggleKernel:   PROPS.getProperty("KAGGLE_KERNEL")   || "",
    kernelTitle:    PROPS.getProperty("KERNEL_TITLE")    || "",
    model:          PROPS.getProperty("MODEL")           || "",
    ngrokDomain:    PROPS.getProperty("NGROK_DOMAIN")    || "",
    proxyApiKey:    PROPS.getProperty("PROXY_API_KEY")   || "",
    controlToken:   PROPS.getProperty("CONTROL_TOKEN")   || "",
    spreadsheetId:  PROPS.getProperty("SPREADSHEET_ID")  || "",
    datasets:       datasets,
    gasWebappUrl:   PROPS.getProperty("GAS_WEBAPP_URL")  || ""
  };
  return CONFIG_CACHE;
}

function cfg(key, label) {
  var v = loadConfig()[key];
  if (!v && v !== 0) throw new Error((label || key) + " が未設定です（CONFIG または個別プロパティを確認）");
  return v;
}

function getKernel() { return cfg("kaggleKernel", "kaggleKernel"); }

function getAuth() {
  var c = loadConfig();
  var user = c.kaggleUsername || getKernel().split("/")[0];
  return "Basic " + Utilities.base64Encode(user + ":" + cfg("kaggleApiKey", "kaggleApiKey"));
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

function setStopRequested(on) {
  if (on) PROPS.setProperty(PROP_STOP, "1");
  else PROPS.deleteProperty(PROP_STOP);
}
function isStopRequested() { return PROPS.getProperty(PROP_STOP) === "1"; }

// ============================================================
// Notebook ソース
// ============================================================

function getNotebookSource() {
  var py = HtmlService.createHtmlOutputFromFile("proxy_py").getContent();
  var c = loadConfig();
  var gasUrl = c.gasWebappUrl || ScriptApp.getService().getUrl();

  py = py
    .replace("__ENDPOINT_ID__",   getKernel())
    .replace("__MODEL__",         cfg("model", "model"))
    .replace("__PROXY_KEY__",     cfg("proxyApiKey", "proxyApiKey"))
    .replace("__NGROK_DOMAIN__",  cfg("ngrokDomain", "ngrokDomain"))
    .replace("__GAS_URL__",       gasUrl)
    .replace("__CONTROL_TOKEN__", cfg("controlToken", "controlToken"));

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
        dataSources: (c.datasets || []).map(function (d) { return { datasetId: d, sourceType: "dataset" }; }),
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
// 稼働ログ（スクリプトプロパティに保存）
// ============================================================
// USAGE_LOG プロパティに JSON 配列で持つ。
//   [{"s":"2026-08-22T01:30:00Z","e":"2026-08-22T03:48:00Z","m":138}, ...]
// プロパティ 1 件の上限は 9KB。1 レコード約 55 バイトなので
// MAX_LOG_ENTRIES で件数を抑え、古いものから捨てる。

var PROP_LOG        = "USAGE_LOG";
var MAX_LOG_ENTRIES = 120;

function isoShort(d) {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function loadLog() {
  var raw = PROPS.getProperty(PROP_LOG);
  if (!raw) return [];
  try {
    var a = JSON.parse(raw);
    return Array.isArray(a) ? a : [];
  } catch (e) {
    Logger.log("USAGE_LOG が壊れています。空として扱います: " + e.message);
    return [];
  }
}

function saveLog(log) {
  if (log.length > MAX_LOG_ENTRIES) log = log.slice(log.length - MAX_LOG_ENTRIES);
  PROPS.setProperty(PROP_LOG, JSON.stringify(log));
}

// notebook と ブラウザが同時に書くことがあるので排他をかける
function withLogLock(fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("ログの排他取得に失敗しました"); }
  try { return fn(); } finally { lock.releaseLock(); }
}

function appendRunStart() {
  return withLogLock(function () {
    var log = loadLog();
    log.push({ s: isoShort(new Date()), e: null, m: null });
    saveLog(log);
    return log.length;
  });
}

/** 停止日時が空の最終エントリを閉じる */
function closeOpenRun() {
  return withLogLock(function () {
    var log = loadLog();
    for (var i = log.length - 1; i >= 0; i--) {
      if (log[i].s && !log[i].e) {
        var stop = new Date();
        var mins = Math.round((stop - new Date(log[i].s)) / 60000);
        log[i].e = isoShort(stop);
        log[i].m = mins;
        saveLog(log);
        return { durationMin: mins, stoppedAt: log[i].e };
      }
    }
    return null;
  });
}

// Kaggle のクォータは土曜朝（UTC）にリセットされる
function getWeekStartUtc() {
  var now = new Date();
  var d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() - 6 + 7) % 7));
  return d;
}

function getWeeklyUsedMinutes() {
  var log = loadLog();
  var since = getWeekStartUtc();
  var now = new Date();
  var total = 0;

  for (var i = 0; i < log.length; i++) {
    var start = new Date(log[i].s);
    if (isNaN(start.getTime()) || start < since) continue;

    if (log[i].m !== null && log[i].m !== undefined && !isNaN(Number(log[i].m))) {
      var m = Number(log[i].m);
      if (m > ABNORMAL_MIN) {
        Logger.log("異常値を除外 start=" + log[i].s + " mins=" + m);
        continue;
      }
      total += m;
      continue;
    }
    // 停止日時が空 = 稼働中。経過分を暫定加算する
    var running = Math.round((now - start) / 60000);
    if (running > 0) total += running;
  }
  return total;
}

/**
 * 全期間の集計。
 *   totalMin   … 全期間の稼働分（異常値は除外）
 *   currentMin … 今回のセッションの経過分（稼働中でなければ null）
 *   startedAt  … 今回の起動日時
 *   runCount   … 記録に残っている起動回数
 * 注意: GAS 経由で起動した分だけが対象。Kaggle の画面から手動で
 * 起動したセッションは記録されないため含まれない。
 */
function getUsageSummary() {
  var log = loadLog();
  var now = new Date();
  var total = 0, currentMin = null, startedAt = null;

  for (var i = 0; i < log.length; i++) {
    var start = new Date(log[i].s);
    if (isNaN(start.getTime())) continue;

    if (log[i].m !== null && log[i].m !== undefined && !isNaN(Number(log[i].m))) {
      var m = Number(log[i].m);
      if (m <= ABNORMAL_MIN) total += m;
      continue;
    }
    var running = Math.round((now - start) / 60000);
    if (running > 0 && running <= ABNORMAL_MIN) {
      total += running;
      currentMin = running;
      startedAt = new Date(log[i].s).toISOString();
    }
  }
  return { totalMin: total, currentMin: currentMin, startedAt: startedAt, runCount: log.length };
}


// ============================================================
// ルーティング
// ============================================================

var PUBLIC_ACTIONS = { ping: true };

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || "status";
  var result;

  try {
    if (!PUBLIC_ACTIONS[action] && p.token !== cfg("controlToken", "controlToken")) {
      result = { success: false, error: "unauthorized" };
    } else if (action === "ping")       { result = { success: true, pong: new Date().toISOString() };
    } else if (action === "status")     { result = handleStatus();
    } else if (action === "start")      { result = handleStart();
    } else if (action === "stop")       { result = handleStop();
    } else if (action === "forceStop")  { result = handleForceStop();
    } else if (action === "started")    { result = handleStarted(p.sessionId, p.modelReady);
    } else if (action === "shouldStop") { result = handleShouldStop();
    } else if (action === "record")     { result = handleRecord();
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

function handleStatus() {
  var kernel = getKernel();
  var parts = kernel.split("/");
  var url = "https://www.kaggle.com/api/v1/kernels/status"
    + "?userName=" + encodeURIComponent(parts[0])
    + "&kernelSlug=" + encodeURIComponent(parts[1]);

  var res = fetchWithRetry(url, {
    method: "GET",
    headers: { "Authorization": getAuth() },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    return { success: false, error: "ステータス取得失敗 (HTTP " + res.getResponseCode() + ")" };
  }

  var data = JSON.parse(res.getContentText());
  var hb = PROPS.getProperty("LAST_HEARTBEAT");
  var used = getWeeklyUsedMinutes();
  var summary = getUsageSummary();
  var c = loadConfig();
  var kaggleStatus = data.status || "unknown";
  var alive = hb ? (new Date() - new Date(hb)) < HEARTBEAT_STALE_MS : false;
  var modelReady = PROPS.getProperty("MODEL_READY") === "1";
  var live = kaggleStatus === "running" || kaggleStatus === "queued";

  return {
    success: true,
    // Kaggle 上は動いているのに notebook が応答しない状態。強制停止の出番。
    zombie: live && !alive && !!hb,
    sessionId: PROPS.getProperty("SESSION_ID") || null,
    modelReady: modelReady,
    label: c.kernelTitle || kernel,
    model: c.model || "",
    status: kaggleStatus,
    proxyAlive: alive,
    lastHeartbeat: hb || null,
    stopRequested: isStopRequested(),
    baseUrl: "https://" + cfg("ngrokDomain", "ngrokDomain") + "/v1",
    weeklyUsedMin: used,
    weeklyRemainMin: Math.max(0, WEEKLY_LIMIT_HOURS * 60 - used),
    weeklyLimitMin: WEEKLY_LIMIT_HOURS * 60,
    weekStartUtc: getWeekStartUtc().toISOString(),
    totalUsedMin: summary.totalMin,
    currentSessionMin: summary.currentMin,
    currentStartedAt: summary.startedAt,
    runCount: summary.runCount
  };
}

function handleStart() {
  setStopRequested(false);

  var st = handleStatus();
  if (st.success && (st.status === "running" || st.status === "queued")) {
    return { success: false, message: "すでに起動中です (" + st.status + ")" };
  }
  if (st.success && st.weeklyRemainMin <= 0) {
    return { success: false, message: "今週のクォータを使い切っています" };
  }

  var c = loadConfig();
  var parts = getKernel().split("/");
  var payload = {
    slug: parts[0] + "/" + parts[1],
    newTitle: c.kernelTitle || parts[1],
    text: getNotebookSource(),
    language: "python",
    kernelType: "notebook",
    isPrivate: true,
    enableGpu: true,
    enableInternet: true,
    datasetDataSources: c.datasets || []
  };

  var res = fetchWithRetry("https://www.kaggle.com/api/v1/kernels/push", {
    method: "POST",
    headers: { "Authorization": getAuth(), "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var out = {};
  try { out = JSON.parse(res.getContentText()); }
  catch (e) { out = { raw: res.getContentText().substring(0, 300) }; }

  if (res.getResponseCode() !== 200 || out.hasError) {
    return {
      success: false,
      message: "起動失敗: " + (out.error || out.errorNullable || "HTTP " + res.getResponseCode()),
      detail: out
    };
  }

  PROPS.deleteProperty("LAST_HEARTBEAT");
  PROPS.deleteProperty("MODEL_READY");
  appendRunStart();

  return { success: true, message: "起動リクエスト送信完了 (v" + out.versionNumber + ")" };
}

/**
 * 停止。ハートビートが生きていれば通常経路（フラグ→notebook が自ら終了）。
 * 途絶していれば待っても無駄なので、その場で強制停止に切り替える。
 */
function handleStop() {
  var hb = PROPS.getProperty("LAST_HEARTBEAT");
  var alive = hb ? (new Date() - new Date(hb)) < HEARTBEAT_STALE_MS : false;

  if (alive) {
    setStopRequested(true);
    return {
      success: true,
      mode: "graceful",
      message: "停止をリクエストしました（30秒以内に停止します）"
    };
  }

  var forced = handleForceStop();
  forced.mode = "forced";
  forced.message = "応答がないため強制停止しました。" + (forced.message || "");
  return forced;
}

/**
 * セッション ID を自力で取得する。
 * notebook からの started 通知が届かなかった場合の保険。
 * kernels/list は自分のカーネル一覧を返すので、そこから
 * 対象カーネルの currentVersionNumber / id を拾う。
 */
function fetchSessionId() {
  var parts = getKernel().split("/");
  var url = "https://www.kaggle.com/api/v1/kernels/list"
    + "?user=" + encodeURIComponent(parts[0])
    + "&search=" + encodeURIComponent(parts[1])
    + "&pageSize=20";

  var res = fetchWithRetry(url, {
    method: "GET",
    headers: { "Authorization": getAuth() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    Logger.log("kernels/list 失敗: HTTP " + res.getResponseCode());
    return null;
  }

  var list = [];
  try { list = JSON.parse(res.getContentText()); } catch (e) { return null; }

  for (var i = 0; i < list.length; i++) {
    var ref = String(list[i].ref || "");
    if (ref === parts[0] + "/" + parts[1]) {
      // 応答に含まれる ID 候補を順に見る。API のバージョンで名前が異なる。
      var cand = list[i].currentVersionNumber || list[i].id || list[i].kernelSessionId;
      Logger.log("kernels/list から取得: " + JSON.stringify(list[i]).substring(0, 400));
      if (cand) return String(cand);
    }
  }
  return null;
}

/**
 * cancel-session API でセッションを即座に落とす。
 * record が走らないため、稼働時間の記録は GAS が肩代わりする。
 */
function handleForceStop() {
  var sid = PROPS.getProperty("SESSION_ID");

  // notebook からの通知が無かった場合は自力で探す
  if (!sid) {
    sid = fetchSessionId();
    if (sid) Logger.log("SESSION_ID を自力取得しました: " + sid);
  }

  if (!sid) {
    return {
      success: false,
      message: "セッション ID が未取得です。Kaggle の画面から手動で停止してください",
      fallbackUrl: "https://www.kaggle.com/code/" + getKernel()
    };
  }

  var res = fetchWithRetry(
    "https://www.kaggle.com/api/v1/kernels/cancel-session/" + encodeURIComponent(sid), {
      method: "POST",
      headers: { "Authorization": getAuth(), "Content-Type": "application/json" },
      payload: "{}",
      muteHttpExceptions: true
    });

  var code = res.getResponseCode();
  var body = res.getContentText().substring(0, 300);

  setStopRequested(false);
  PROPS.deleteProperty("LAST_HEARTBEAT");
  PROPS.deleteProperty("SESSION_ID");

  // record が走らないので、開いたままの行をここで閉じる
  var closed = closeOpenRun();

  if (code !== 200) {
    // cancel-session が効かない場合の最終手段。
    // GPU 無効の空ノートブックを同じスラッグに push すると、
    // 実行中のセッションが置き換えられて終了する。
    Logger.log("cancel-session 失敗 (HTTP " + code + ")。push による置き換えを試みます");
    var replaced = replaceWithIdleNotebook();
    return {
      success: replaced.success,
      message: replaced.success
        ? "cancel が効かないため空ノートブックで置き換えました。数十秒で停止します"
        : "強制停止に失敗 (HTTP " + code + "): " + body,
      fallbackUrl: "https://www.kaggle.com/code/" + getKernel()
    };
  }
  return {
    success: true,
    message: closed ? "稼働 " + closed.durationMin + " 分を記録しました" : "",
    durationMin: closed ? closed.durationMin : null
  };
}

/**
 * GPU を無効にした最小のノートブックを同じスラッグへ push する。
 * 走っている GPU セッションを確実に終わらせるための最終手段。
 * 次回 handleStart() を呼べば正規のノートブックで上書きされるので、
 * この状態が残っても支障はない。
 */
function replaceWithIdleNotebook() {
  var parts = getKernel().split("/");
  var idle = JSON.stringify({
    cells: [{
      cell_type: "code", execution_count: null,
      metadata: { trusted: true }, outputs: [],
      source: ["print('stopped by controller')\n"]
    }],
    metadata: {
      kernelspec: { display_name: "Python 3", language: "python", name: "python3" },
      language_info: { name: "python", version: "3.12.13" }
    },
    nbformat: 4, nbformat_minor: 4
  });

  var res = fetchWithRetry("https://www.kaggle.com/api/v1/kernels/push", {
    method: "POST",
    headers: { "Authorization": getAuth(), "Content-Type": "application/json" },
    payload: JSON.stringify({
      slug: parts[0] + "/" + parts[1],
      newTitle: loadConfig().kernelTitle || parts[1],
      text: idle,
      language: "python",
      kernelType: "notebook",
      isPrivate: true,
      enableGpu: false,        // ここが要点。GPU を落とす
      enableInternet: false,
      datasetDataSources: []
    }),
    muteHttpExceptions: true
  });

  var ok = res.getResponseCode() === 200;
  Logger.log("idle push: HTTP " + res.getResponseCode() + " " + res.getContentText().substring(0, 200));
  return { success: ok };
}

function handleStarted(sessionId, modelReady) {
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  // 強制停止に使う。notebook からしか取得できない値なのでここで預かる。
  if (sessionId) PROPS.setProperty("SESSION_ID", String(sessionId));
  // モデルが実際にチャットへ応答できるかを別プロパティで持つ。
  // proxyAlive（通知が来ているか）とは意味が違うので分けておく。
  PROPS.setProperty("MODEL_READY", modelReady === "true" || modelReady === true ? "1" : "0");
  setStopRequested(false);
  return { success: true };
}

function handleShouldStop() {
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  return { success: true, stop: isStopRequested() };
}

function handleRecord() {
  setStopRequested(false);
  PROPS.deleteProperty("LAST_HEARTBEAT");
  PROPS.deleteProperty("SESSION_ID");
  PROPS.deleteProperty("MODEL_READY");
  var closed = closeOpenRun();
  if (!closed) return { success: false, message: "未クローズの行がありません" };
  return { success: true, durationMin: closed.durationMin };
}


// ============================================================
// 自律監視（1分おきのトリガー）
// ============================================================
// ダッシュボードの画面を誰も開いていなくても、このGAS自身が
// 1分ごとに自分の状態を確認し、次の2つを検出する。
//
//   1. Kaggle側が自動でセッションを終了した（9時間上限など）
//   2. 他の人がKaggleの画面から直接停止した
//
// どちらの場合も notebook 側の handleRecord() 通知は届かないため、
// USAGE_LOG が「開いたまま」になり、週次クォータの計算が
// 実態より多め（＝残量が少なめ）にズレ続けてしまう。
// monitorTick() はその食い違いを検出し、記録を閉じて正しい
// 残量に戻す。

var PROP_LAST_SEEN_STATUS = "LAST_SEEN_STATUS";

function monitorTick() {
  var log = loadLog();
  var hasOpenRun = log.length > 0 && log[log.length - 1].e === null;

  if (!hasOpenRun) {
    // 記録上「稼働中」の行が無いなら、何もすることがない
    return;
  }

  var st;
  try {
    st = handleStatus();
  } catch (err) {
    Logger.log("monitorTick: ステータス取得に失敗: " + err.message);
    return;
  }
  if (!st.success) {
    Logger.log("monitorTick: " + st.error);
    return;
  }

  var wasRunning = st.status === "running" || st.status === "queued";

  if (st.zombie) {
    // Kaggle上は動いているがnotebookが応答しない = 詰まっている。
    // 強制停止して記録を閉じ、次回起動できる状態に戻す。
    Logger.log("monitorTick: zombie検出。強制停止します");
    handleForceStop();
    return;
  }

  if (!wasRunning) {
    // Kaggleが自分で終了した、または他の人が画面から止めた。
    // notebookからのrecord通知が来ていないので、ここで記録を閉じる。
    Logger.log("monitorTick: セッション終了を検知（status=" + st.status + "）。記録を閉じます");
    PROPS.deleteProperty("LAST_HEARTBEAT");
    PROPS.deleteProperty("SESSION_ID");
    PROPS.deleteProperty("MODEL_READY");
    setStopRequested(false);
    closeOpenRun();
  }
}

/** 1分おきのトリガーを設定する。setupTrigger() を一度だけ実行すればよい */
function setupMonitorTrigger() {
  removeMonitorTrigger();
  ScriptApp.newTrigger("monitorTick").timeBased().everyMinutes(1).create();
  Logger.log("1分おきの監視トリガーを設定しました");
}

function removeMonitorTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "monitorTick") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function testMonitorTick() { monitorTick(); Logger.log("実行しました"); }

// ============================================================
// テスト関数
// ============================================================

function testPing()    { Logger.log(JSON.stringify({ success: true, pong: new Date().toISOString() })); }
function testStatus()  { Logger.log(JSON.stringify(handleStatus(), null, 2)); }
function testStart()   { Logger.log(JSON.stringify(handleStart(), null, 2)); }
function testStop()    { Logger.log(JSON.stringify(handleStop(), null, 2)); }
function clearStopFlag() { setStopRequested(false); Logger.log("解除しました"); }
function testForceStop() { Logger.log(JSON.stringify(handleForceStop(), null, 2)); }
function testFetchSessionId() { Logger.log("取得結果: " + fetchSessionId()); }
function testReplaceIdle() { Logger.log(JSON.stringify(replaceWithIdleNotebook(), null, 2)); }
function testSessionId() {
  Logger.log("SESSION_ID: " + (PROPS.getProperty("SESSION_ID") || "(未取得。notebook 起動後に入る)"));
}

// CONFIG（または個別プロパティ）が正しく読めるか確認する
function testConfig() {
  var c = loadConfig();
  var required = ["kaggleUsername", "kaggleApiKey", "kaggleKernel", "model",
                   "ngrokDomain", "proxyApiKey", "controlToken"];
  required.forEach(function (k) {
    Logger.log(k + ": " + (c[k] ? "OK" : "★未設定"));
  });
  Logger.log("datasets: " + JSON.stringify(c.datasets));
  Logger.log("gasWebappUrl: " + (c.gasWebappUrl || "(未設定→getUrl()を使用)"));
  Logger.log("読み込み元: " + (PROPS.getProperty("CONFIG") ? "CONFIG" : "個別プロパティ"));
}

function testNotebookSource() {
  var src = getNotebookSource();
  Logger.log("length: " + src.length);
  Logger.log("残プレースホルダ: " + (src.match(/__[A-Z_]+__/g) || []).join(", "));
  Logger.log("JSON: " + (JSON.parse(src) ? "OK" : "NG"));
}

function testUsage() { Logger.log(JSON.stringify(getUsageSummary(), null, 2)); }

function testLog() {
  var log = loadLog();
  var raw = PROPS.getProperty(PROP_LOG) || "";
  Logger.log("件数: " + log.length + " / " + MAX_LOG_ENTRIES);
  Logger.log("サイズ: " + raw.length + " 文字（上限 9216）");
  Logger.log(JSON.stringify(log.slice(-5), null, 2));
}

/** 稼働ログを全消去する。やり直したいときだけ使う */
function clearLog() {
  PROPS.deleteProperty(PROP_LOG);
  Logger.log("USAGE_LOG を削除しました");
}

/**
 * 既存のスプレッドシートから USAGE_LOG へ一度だけ移行する。
 * CONFIG に spreadsheetId が残っている必要がある。実行後は不要。
 */
function migrateFromSheet() {
  var id = loadConfig().spreadsheetId;
  if (!id) { Logger.log("spreadsheetId が未設定です"); return; }

  var sheet = SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log("シート「" + SHEET_NAME + "」が見つかりません"); return; }

  var data = sheet.getDataRange().getValues();
  var log = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var start = new Date(data[i][0]);
    if (isNaN(start.getTime())) continue;
    var stop = data[i][1] ? new Date(data[i][1]) : null;
    var mins = data[i][2];
    log.push({
      s: isoShort(start),
      e: stop && !isNaN(stop.getTime()) ? isoShort(stop) : null,
      m: (mins !== "" && mins !== null && !isNaN(Number(mins))) ? Number(mins) : null
    });
  }
  saveLog(log);
  Logger.log("移行しました: " + log.length + " 件（上限 " + MAX_LOG_ENTRIES + " 件で切り詰め）");
  Logger.log("サイズ: " + (PROPS.getProperty(PROP_LOG) || "").length + " 文字");
}

function testWeekly() {
  Logger.log("週起点: " + getWeekStartUtc().toISOString());
  Logger.log("今週使用: " + getWeeklyUsedMinutes() + "分");
}

// CONFIG の雛形をログに出す（コピーして値を埋めるとよい）
function printConfigTemplate() {
  Logger.log(JSON.stringify({
    kaggleUsername: "ymatsuda2025",
    kaggleApiKey: "",
    kaggleKernel: "ymatsuda2025/gemma4-12b",
    kernelTitle: "gemma4 12b",
    model: "gemma4:12b",
    ngrokDomain: "xxx.ngrok-free.dev",
    proxyApiKey: "",
    controlToken: "",
    datasets: ["ymatsuda2025/ollama-gemma4-12b-v2", "ymatsuda2025/ngrok-binary"],
    gasWebappUrl: ""
  }, null, 2));
}

function checkForOldSyntax() {
  var src = HtmlService.createHtmlOutputFromFile('proxy_py').getContent();
  var hasOld = src.indexOf('while time.time()') >= 0;
  Logger.log(hasOld ? '★古いコードのままです（while time.time() が残っている）' : '新しいコードです（for attempt in range 方式）');
}

function clearNbCache() {
  CacheService.getScriptCache().remove("nbsrc");
  Logger.log("キャッシュを削除しました");
}

function testNotebookSource() {
  var src = getNotebookSource();
  Logger.log("length: " + src.length);
  Logger.log("残プレースホルダ: " + (src.match(/__[A-Z_]+__/g) || []).join(", "));
  Logger.log("JSON: " + (JSON.parse(src) ? "OK" : "NG"));
}

function handleStarted(sessionId, modelReady) {
  Logger.log("受け取った modelReady: " + JSON.stringify(modelReady) + " (型: " + typeof modelReady + ")");
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  if (sessionId) PROPS.setProperty("SESSION_ID", String(sessionId));
  PROPS.setProperty("MODEL_READY", modelReady === "true" || modelReady === true ? "1" : "0");
  setStopRequested(false);
  return { success: true };
}

function checkStartedRouting() {
  var src = doGet.toString();
  var i = src.indexOf('"started"');
  Logger.log(src.substring(Math.max(0, i - 30), i + 200));
}

function checkStartedCall() {
  var src = getNotebookSource();
  var i = src.indexOf("gas('started'");
  Logger.log(src.substring(i, i + 150));
}