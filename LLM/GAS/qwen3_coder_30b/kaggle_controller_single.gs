// ============================================================
// Kaggle Notebook コントローラー - Google Apps Script（単体版）
// ============================================================
// このファイルは「1エンドポイント = 1つの GAS プロジェクト」の原則に
// 従い、エンドポイントごとに別々のGASプロジェクトへそのままデプロイする。
// コードは全エンドポイント共通で、設定だけが CONFIG スクリプトプロパティで変わる。
//
// 【設定】スクリプトプロパティ "CONFIG" に、次の形の JSON を1行で入れる。
//
// {
//   "kaggleUsername": "matsuda2026",
//   "kaggleApiKey": "Kaggle の API キー",
//   "kaggleKernel": "matsuda2026/qwen3_8-27b",
//   "kernelTitle": "qwen3.8 27b",
//   "model": "qwen3.8:27b",
//   "ngrokDomain": "smartly-overflow-unleash.ngrok-free.dev",
//   "proxyApiKey": "このエンドポイント専用のキー",
//   "controlToken": "このエンドポイント専用のキー（別値）",
//   "datasets": ["matsuda2026/ngrok-binary", "matsuda2026/ollama-qwen3-8-27b"],
//   "gasWebappUrl": "デプロイ後の /exec URL（省略可・省略時は ScriptApp.getService().getUrl() を使用）"
// }
//
// 旧形式（個別プロパティ）は使わない。CONFIG が無い場合はエラーにする。
//
// 【proxy_py.html】このプロジェクトの HTML ファイルとして保存する。
//   プレースホルダ: __PROXY_KEY__ __GAS_URL__ __CONTROL_TOKEN__
//                    __NGROK_DOMAIN__ __ENDPOINT_ID__ __MODEL__
//   ENDPOINT_ID には kaggleKernel の値をそのまま流用する。
// ============================================================

var PROPS              = PropertiesService.getScriptProperties();
var RETRY_COUNT         = 3;
var RETRY_INTERVAL_MS   = 3000;
var WEEKLY_LIMIT_HOURS  = 30;
var HEARTBEAT_STALE_MS  = 3 * 60000;   // 3分ハートビートが無ければ死んでいるとみなす
var BOOT_GRACE_MS       = 15 * 60000;  // 起動リクエストから初回ハートビートまでの猶予（zstd/ollama pull/vision確認/wait_model_ready込み）
var ABNORMAL_MIN        = 720;         // 12時間超の記録は異常値として週間集計から除外
var USAGE_LOG_MAX       = 120;
var PROP_STOP           = "STOP_REQUESTED";
var PROP_STOP_AT        = "STOP_REQUESTED_AT";
var STOP_STUCK_MS       = 5 * 60000;   // 停止指示からこの時間応答がなければ「詰まった停止」とみなす

// ============================================================
// CONFIG の読み込み
// ============================================================

var CONFIG_CACHE = null;

function loadConfig() {
  if (CONFIG_CACHE) return CONFIG_CACHE;
  var raw = PROPS.getProperty("CONFIG");
  if (!raw) throw new Error("CONFIG スクリプトプロパティが未設定です");
  var c;
  try {
    c = JSON.parse(raw);
  } catch (e) {
    throw new Error("CONFIG の JSON が壊れています: " + e.message);
  }
  var required = [
    "kaggleUsername", "kaggleApiKey", "kaggleKernel", "kernelTitle",
    "model", "ngrokDomain", "proxyApiKey", "controlToken"
  ];
  var missing = required.filter(function (k) { return !c[k]; });
  if (missing.length) {
    throw new Error("CONFIG に必須項目がありません: " + missing.join(", "));
  }
  if (!Array.isArray(c.datasets)) c.datasets = [];
  CONFIG_CACHE = c;
  return c;
}

function getAuth() {
  var c = loadConfig();
  return "Basic " + Utilities.base64Encode(c.kaggleUsername + ":" + c.kaggleApiKey);
}

function getGasUrl() {
  var c = loadConfig();
  return c.gasWebappUrl || ScriptApp.getService().getUrl();
}

// ============================================================
// 共通ユーティリティ
// ============================================================

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

/**
 * 停止指示のON/OFFを、時刻付きで記録する。
 * 時刻を持たせることで、「停止を指示したのに一向に反映されない」
 * 詰まった状態を monitorTick() が検出できるようにする。
 */
function setStopRequested(on) {
  if (on) {
    PROPS.setProperty(PROP_STOP, "1");
    PROPS.setProperty(PROP_STOP_AT, new Date().toISOString());
  } else {
    PROPS.deleteProperty(PROP_STOP);
    PROPS.deleteProperty(PROP_STOP_AT);
  }
}
function isStopRequested() { return PROPS.getProperty(PROP_STOP) === "1"; }

// ============================================================
// Notebook ソース生成
// ============================================================

function getNotebookSource() {
  var c = loadConfig();
  var py = HtmlService.createHtmlOutputFromFile("proxy_py").getContent();

  py = py
    .replace(/__ENDPOINT_ID__/g,   c.kaggleKernel)
    .replace(/__MODEL__/g,         c.model)
    .replace(/__PROXY_KEY__/g,     c.proxyApiKey)
    .replace(/__NGROK_DOMAIN__/g,  c.ngrokDomain)
    .replace(/__GAS_URL__/g,       getGasUrl())
    .replace(/__CONTROL_TOKEN__/g, c.controlToken);

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
        accelerator: "gpu_t4_x2", // 表示用のみ。実際のGPU種別は push payload の machineShape で決まる
        dataSources: c.datasets.map(function (d) { return { datasetId: d, sourceType: "dataset" }; }),
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

// 停止用: GPU無効・データセット無しの空ノートブック（最終手段の強制置換に使う）
function getIdleNotebookSource() {
  var c = loadConfig();
  return JSON.stringify({
    cells: [{
      cell_type: "code", execution_count: null,
      metadata: {}, outputs: [], source: ["print('stopped by forceStop')"]
    }],
    metadata: {
      kaggle: {
        accelerator: "none",
        dataSources: [],
        dockerImageVersionId: 28755,
        isGpuEnabled: false, isInternetEnabled: false,
        language: "python", sourceType: "notebook"
      },
      kernelspec: { display_name: "Python 3", language: "python", name: "python3" },
      language_info: { name: "python", version: "3.12.13" }
    },
    nbformat: 4, nbformat_minor: 4
  });
}

// ============================================================
// 稼働ログ（USAGE_LOG プロパティに JSON 配列で保存）
//   [{"s": 開始ISO, "e": 終了ISO|null, "m": 分|null}, ...]
// ============================================================

function getUsageLog() {
  try { return JSON.parse(PROPS.getProperty("USAGE_LOG") || "[]"); }
  catch (e) { return []; }
}

function saveUsageLog(log) {
  if (log.length > USAGE_LOG_MAX) log = log.slice(log.length - USAGE_LOG_MAX);
  PROPS.setProperty("USAGE_LOG", JSON.stringify(log));
}

function appendRunStart() {
  var log = getUsageLog();
  log.push({ s: new Date().toISOString(), e: null, m: null });
  saveUsageLog(log);
}

// 未クローズ行（e:null）を今の時刻で締める。強制停止時にGASが肩代わりする。
function closeOpenRun() {
  var log = getUsageLog();
  for (var i = log.length - 1; i >= 0; i--) {
    if (log[i].e === null || log[i].e === undefined) {
      var stop = new Date();
      var mins = Math.round((stop - new Date(log[i].s)) / 60000);
      log[i].e = stop.toISOString();
      log[i].m = mins;
      saveUsageLog(log);
      return { durationMin: mins, stoppedAt: stop.toISOString() };
    }
  }
  return null;
}

function hasOpenRun() {
  var log = getUsageLog();
  return log.length > 0 && (log[log.length - 1].e === null || log[log.length - 1].e === undefined);
}

// 週の起点は土曜0時UTC
function getWeekStartUtc() {
  var now = new Date();
  var d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() - 6 + 7) % 7));
  return d;
}

function getWeeklyUsedMinutes() {
  var log = getUsageLog();
  var since = getWeekStartUtc();
  var now = new Date();
  var total = 0;

  log.forEach(function (row) {
    var start = new Date(row.s);
    if (isNaN(start.getTime()) || start < since) return;

    if (row.m !== null && row.m !== undefined && row.m !== "") {
      var m = Number(row.m);
      if (m > ABNORMAL_MIN) return; // 異常値は除外（closeStaleRuns で個別修復）
      total += m;
      return;
    }
    // 未クローズ行は「稼働中」とみなし経過時間を暫定加算
    var running = Math.round((now - start) / 60000);
    if (running > 0) total += running;
  });
  return total;
}

// 未クローズ行を手動で締める（クォータが実態と乖離したときの修復用）
function closeStaleRuns() {
  var closed = closeOpenRun();
  Logger.log(closed ? JSON.stringify(closed) : "未クローズの行はありません");
  return closed;
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
    var c = loadConfig();
    if (!PUBLIC_ACTIONS[action] && p.token !== c.controlToken) {
      result = { success: false, error: "unauthorized" };
    } else if (action === "ping")       { result = { success: true, pong: new Date().toISOString() };
    } else if (action === "status")     { result = handleStatus();
    } else if (action === "start")      { result = handleStart();
    } else if (action === "stop")       { result = handleStop();
    } else if (action === "started")    { result = handleStarted(p.sessionId, p.modelReady);
    } else if (action === "shouldStop") { result = handleShouldStop();
    } else if (action === "record")     { result = handleRecord();
    } else if (action === "forceStop")  { result = handleForceStopAction();
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

// 注意: Kaggleの kernels/status API（実行セッション状態の監視）は、公開APIトークンから
// 権限が外されており(Permission 'kernels.get' was denied)、Legacy/新形式どちらのキーでも
// 403になることが確認されている。そのため状態判定はKaggle APIに依存せず、
// このNotebook自身が30秒おきに送ってくるハートビート(shouldStopポーリング)と、
// 起動リクエストを送った時刻(START_REQUESTED_AT)から自前で組み立てる。
// Kaggle側のstatusは「取れれば参考情報として付ける」程度に格下げする。
function handleStatus() {
  var c = loadConfig();
  var now = new Date();
  var hb = PROPS.getProperty("LAST_HEARTBEAT");
  var aliveByHeartbeat = hb ? (now - new Date(hb)) < HEARTBEAT_STALE_MS : false;
  var stopReq = isStopRequested();
  var startedAt = PROPS.getProperty("START_REQUESTED_AT");
  var withinBootGrace = startedAt ? (now - new Date(startedAt)) < BOOT_GRACE_MS : false;

  var status;
  if (aliveByHeartbeat) {
    status = "running";
  } else if (stopReq) {
    status = "stopping";
  } else if (withinBootGrace) {
    status = "booting";
  } else {
    status = "stopped";
  }

  // ハートビートを一度でも受け取った後に途絶した（かつ停止指示もしていない）= 異常終了の疑い
  var zombie = !!hb && !aliveByHeartbeat && !stopReq && !withinBootGrace;

  // 停止を指示したのに、猶予時間を過ぎても反映されない = 詰まっている
  var stopAt = PROPS.getProperty(PROP_STOP_AT);
  var stopStuck = stopReq && stopAt && (now - new Date(stopAt)) > STOP_STUCK_MS;

  var used = getWeeklyUsedMinutes();
  var result = {
    success: true,
    label: c.kernelTitle,
    model: c.model,
    status: status,
    proxyAlive: aliveByHeartbeat,
    modelReady: PROPS.getProperty("MODEL_READY") === "1",
    lastHeartbeat: hb || null,
    zombie: zombie,
    stopRequested: stopReq,
    stopStuck: stopStuck,
    sessionId: PROPS.getProperty("SESSION_ID") || null,
    baseUrl: "https://" + c.ngrokDomain + "/v1",
    weeklyUsedMin: used,
    weeklyRemainMin: Math.max(0, WEEKLY_LIMIT_HOURS * 60 - used),
    weeklyLimitMin: WEEKLY_LIMIT_HOURS * 60,
    weekStartUtc: getWeekStartUtc().toISOString()
  };

  // Kaggle kernels/status は参考情報として試すだけ。失敗しても致命的にしない。
  try {
    var parts = c.kaggleKernel.split("/");
    var url = "https://www.kaggle.com/api/v1/kernels/status"
      + "?userName=" + encodeURIComponent(parts[0])
      + "&kernelSlug=" + encodeURIComponent(parts[1]);
    var res = fetchWithRetry(url, {
      method: "GET",
      headers: { "Authorization": getAuth() },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      result.kaggleStatus = JSON.parse(res.getContentText());
    } else {
      result.kaggleStatusNote = "kernels/status 取得不可 (HTTP " + res.getResponseCode() + ")。権限制限の既知事象のため無視。";
    }
  } catch (e) {
    result.kaggleStatusNote = "kernels/status 取得不可: " + String(e && e.message ? e.message : e);
  }

  return result;
}

function handleStart() {
  var c = loadConfig();
  setStopRequested(false);

  var st = handleStatus();
  if (st.status === "running" || st.status === "booting") {
    return { success: false, message: "すでに起動中/起動処理中です (" + st.status + ")" };
  }
  if (st.weeklyRemainMin <= 0) {
    return { success: false, message: "今週のクォータを使い切っています" };
  }

  var parts = c.kaggleKernel.split("/");
  var payload = {
    slug: c.kaggleKernel,
    newTitle: c.kernelTitle,
    text: getNotebookSource(),
    language: "python",
    kernelType: "notebook",
    isPrivate: true,
    enableGpu: true,
    enableInternet: true,
    machineShape: "NvidiaTeslaT4",
    datasetDataSources: c.datasets
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
  PROPS.deleteProperty("SESSION_ID");
  PROPS.setProperty("START_REQUESTED_AT", new Date().toISOString());
  appendRunStart();

  return { success: true, message: "起動リクエスト送信完了 (v" + out.versionNumber + ")" };
}

// 3段階エスカレーション。
// 1) ハートビートが生きていれば通常停止フラグのみ立てる（notebook側が自ら終了しrecordする）
// 2) ハートビートが途絶していれば、cancel-session APIで強制停止し、GASがcloseOpenRun()で記録を肩代わり
// 3) cancel-sessionが失敗すれば、GPU無効の空ノートブックを同じslugへpushして強制的にセッションを置き換える
function handleStop() {
  var hb = PROPS.getProperty("LAST_HEARTBEAT");
  var alive = hb ? (new Date() - new Date(hb)) < HEARTBEAT_STALE_MS : false;

  if (alive) {
    setStopRequested(true);
    return {
      success: true,
      stage: 1,
      message: "停止をリクエストしました（30秒以内のポーリングで自動停止します）"
    };
  }

  // ハートビート途絶 → 強制停止(stage2)
  var sessionId = PROPS.getProperty("SESSION_ID");
  if (sessionId) {
    try {
      var res = fetchWithRetry(
        "https://www.kaggle.com/api/v1/kernels/cancel-session/" + encodeURIComponent(sessionId),
        { method: "POST", headers: { "Authorization": getAuth() }, muteHttpExceptions: true }
      );
      if (res.getResponseCode() === 200) {
        var closed = closeOpenRun();
        setStopRequested(false);
        PROPS.deleteProperty("LAST_HEARTBEAT");
        PROPS.deleteProperty("SESSION_ID");
        PROPS.deleteProperty("MODEL_READY");
        PROPS.deleteProperty("START_REQUESTED_AT");
        return {
          success: true, stage: 2,
          message: "ハートビートが途絶していたため cancel-session で強制停止しました",
          durationMin: closed ? closed.durationMin : null
        };
      }
    } catch (e) {
      // フォールスルーしてstage3へ
    }
  }

  // stage3: 最終手段
  try {
    handleForceStopInternal();
    return { success: true, stage: 3, message: "cancel-session が失敗したため、空ノートブックで強制置換しました" };
  } catch (e) {
    return { success: false, error: String(e && e.message ? e.message : e) };
  }
}

// action=forceStop（強制停止のみ単体実行）
function handleForceStopAction() {
  handleForceStopInternal();
  return { success: true, message: "強制停止（空ノートブック置換）を実行しました" };
}

function handleForceStopInternal() {
  var c = loadConfig();
  var payload = {
    slug: c.kaggleKernel,
    newTitle: c.kernelTitle + " (idle)",
    text: getIdleNotebookSource(),
    language: "python",
    kernelType: "notebook",
    isPrivate: true,
    enableGpu: false,
    enableInternet: false,
    datasetDataSources: []
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

  closeOpenRun();
  setStopRequested(false);
  PROPS.deleteProperty("LAST_HEARTBEAT");
  PROPS.deleteProperty("SESSION_ID");
  PROPS.deleteProperty("MODEL_READY");
  PROPS.deleteProperty("START_REQUESTED_AT");

  if (res.getResponseCode() !== 200 || out.hasError) {
    throw new Error("強制置換pushに失敗: " + JSON.stringify(out));
  }
  return out;
}

function handleStarted(sessionId, modelReady) {
  PROPS.setProperty("LAST_HEARTBEAT", new Date().toISOString());
  if (sessionId) PROPS.setProperty("SESSION_ID", String(sessionId));
  // Python の bool は 'True'/'true' どちらでも来うるので両方受ける
  var ready = (modelReady === "true" || modelReady === "True" || modelReady === true);
  PROPS.setProperty("MODEL_READY", ready ? "1" : "0");
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
  PROPS.deleteProperty("START_REQUESTED_AT");
  var closed = closeOpenRun();
  if (!closed) return { success: false, message: "未クローズの行がありません" };
  return { success: true, durationMin: closed.durationMin };
}

// ============================================================
// 自律監視（1分おきのトリガー）
// ============================================================
// ダッシュボードの画面を誰も見ていなくても、このGAS自身が1分ごとに
// 自分の状態を確認し、次の3つを検出して記録を正しく保つ。
//
//   1. ハートビートが途絶した（zombie）= Kaggleが自動終了した、
//      または他の人がKaggleの画面から直接止めた
//   2. 起動を指示したのに、猶予時間を過ぎても一度もハートビートが
//      来ないまま「stopped」になっている = 起動に静かに失敗した
//   3. 停止を指示したのに、猶予時間を過ぎても反映されない
//      = 停止処理が詰まっている
//
// いずれも handleForceStopInternal() で記録を閉じ、次回起動できる
// 状態に戻す。

function monitorTick() {
  if (!hasOpenRun()) {
    // 記録上「稼働中」の行が無いなら、確認することがない。
    // ただし実行自体は生きていることが分かるよう1行だけ残す。
    Logger.log("monitorTick: 稼働中の記録なし。スキップ");
    return;
  }

  var st;
  try {
    st = handleStatus();
  } catch (err) {
    Logger.log("monitorTick: ステータス取得に失敗: " + err.message);
    return;
  }

  if (st.zombie) {
    Logger.log("monitorTick: zombie検出（ハートビート途絶）。強制停止します");
    try { handleForceStopInternal(); }
    catch (e) { Logger.log("monitorTick: 強制停止に失敗: " + e.message); }
    return;
  }

  if (st.stopStuck) {
    Logger.log("monitorTick: 停止指示から" + (STOP_STUCK_MS / 60000) + "分応答なし。強制停止します");
    try { handleForceStopInternal(); }
    catch (e) { Logger.log("monitorTick: 強制停止に失敗: " + e.message); }
    return;
  }

  if (st.status === "stopped" && !st.stopRequested) {
    // 起動猶予も過ぎ、ハートビートも一度も来ず、こちらから停止も
    // 指示していないのに記録上は稼働中のまま = 起動が静かに失敗していた
    Logger.log("monitorTick: 起動失敗を検知（一度もハートビートなし）。記録を閉じます");
    closeOpenRun();
    PROPS.deleteProperty("START_REQUESTED_AT");
    return;
  }

  // running / booting / stopping はそのまま様子を見る（ログは残さない）
}

/** 1分おきのトリガーを設定する。一度だけ実行すればよい */
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
// テスト関数（GASエディタから手動実行する）
// ============================================================

function testPing() { Logger.log(JSON.stringify({ success: true, pong: new Date().toISOString() })); }
function testStatus() { Logger.log(JSON.stringify(handleStatus(), null, 2)); }
function testStart() { Logger.log(JSON.stringify(handleStart(), null, 2)); }
function testStop() { Logger.log(JSON.stringify(handleStop(), null, 2)); }
function clearStopFlag() { setStopRequested(false); Logger.log("解除しました"); }

// CONFIG の必須項目チェック
function testProps() {
  var raw = PROPS.getProperty("CONFIG");
  if (!raw) { Logger.log("★ CONFIG が未設定です"); return; }
  var c;
  try { c = JSON.parse(raw); } catch (e) { Logger.log("★ CONFIG の JSON が壊れています: " + e.message); return; }

  ["kaggleUsername", "kaggleApiKey", "kaggleKernel", "kernelTitle", "model",
   "ngrokDomain", "proxyApiKey", "controlToken"]
    .forEach(function (k) {
      Logger.log(k + ": " + (c[k] ? "OK" : "★未設定"));
    });
  Logger.log("datasets: " + JSON.stringify(c.datasets || []));
  Logger.log("gasWebappUrl: " + (c.gasWebappUrl || "(未設定→getUrl()を使用)"));
}

// CONFIG の内容を機密情報をマスクして表示
function testConfig() {
  var c = loadConfig();
  var masked = JSON.parse(JSON.stringify(c));
  ["kaggleApiKey", "proxyApiKey", "controlToken"].forEach(function (k) {
    if (masked[k]) masked[k] = masked[k].substring(0, 4) + "…(" + masked[k].length + "文字)";
  });
  Logger.log(JSON.stringify(masked, null, 2));
}

function testNotebookSource() {
  var src = getNotebookSource();
  Logger.log("length: " + src.length);
  Logger.log("残プレースホルダ: " + (src.match(/__[A-Z_]+__/g) || []).join(", "));
  Logger.log("JSON: " + (JSON.parse(src) ? "OK" : "NG"));
}

function testWeekly() {
  Logger.log("週起点: " + getWeekStartUtc().toISOString());
  Logger.log("今週使用: " + getWeeklyUsedMinutes() + "分");
}

// CONFIG の雛形をログに出す（コピーして値を埋めるとよい）
function printConfigTemplate() {
  Logger.log(JSON.stringify({
    kaggleUsername: "matsuda2026",
    kaggleApiKey: "",
    kaggleKernel: "matsuda2026/qwen3_8-27b",
    kernelTitle: "qwen3.8 27b",
    model: "qwen3.8:27b",
    ngrokDomain: "smartly-overflow-unleash.ngrok-free.dev",
    proxyApiKey: "",
    controlToken: "",
    datasets: ["matsuda2026/ngrok-binary", "matsuda2026/ollama-qwen3-8-27b"],
    gasWebappUrl: ""
  }, null, 2));
}

function inspectRawLog() {
  var raw = PROPS.getProperty("USAGE_LOG")
  Logger.log("型: " + typeof raw)
  Logger.log("長さ: " + (raw ? raw.length : 0))
  Logger.log("先頭200文字: " + (raw ? raw.substring(0, 200) : "(null)"))
  try {
    var parsed = JSON.parse(raw)
    Logger.log("パース後の型: " + typeof parsed + " / Array.isArray: " + Array.isArray(parsed))
  } catch (e) {
    Logger.log("JSONパース失敗: " + e.message)
  }
}

function resetLog() {
  PROPS.setProperty("USAGE_LOG", JSON.stringify([]))
  Logger.log("USAGE_LOG を空配列にリセットしました")
}