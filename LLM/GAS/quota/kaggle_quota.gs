/*************************************************************************
 * kaggle_quota.gs  —  Kaggle アクセラレータ残量マネージャ (GAS)
 *
 * 【設計方針】
 *   Kaggle には残量を返す公開APIが無い（kernels/status は403、Web UI の
 *   quota 表示は認証Cookie必須）。よって GAS 側を残量の権威とし、
 *   ノートブックからの started / heartbeat / stopped 通知を台帳
 *   （スプレッドシート）に積み上げて消費時間を自前計算する。
 *
 *   残量 = limit_hours - Σ(セッション稼働秒 ∩ 現在のクォータ窓)
 *
 *   実測とのズレは quota_calibrate で基準点を打ち直して補正する。
 *
 * 【既存 kaggle_controller_single.gs への組み込み】
 *   1) このファイルをプロジェクトに追加
 *   2) QUOTA_CONFIG を編集（API_TOKEN は必ず変更）
 *   3) 既存 doGet の先頭に下記を追加：
 *        var q = quotaHandle(e); if (q) return q;
 *   4) handleStarted()   内で quotaSessionStart(sessionId, kernel, 'GPU')
 *      handleHeartbeat() 内で quotaHeartbeat(sessionId, 'GPU', kernel)
 *      停止処理          内で quotaSessionEnd(sessionId, 'stopped')
 *   5) installQuotaTriggers() を一度だけ手動実行
 *   6) 【重要】コード保存後は「新しいバージョン」として再デプロイ
 *************************************************************************/

var QUOTA_CONFIG = {
  // 変更用エンドポイント（calibrate / purge / reset）の共有シークレット
  API_TOKEN: 'CHANGE_ME_TO_A_LONG_RANDOM_STRING',

  DEFAULT_ACCEL: 'GPU',

  // heartbeat がこの秒数途絶えたら「セッション終了」と見なして締める
  STALE_HEARTBEAT_SEC: 300,

  // 1セッションの上限（Kaggle の最大実行時間。事故時の暴走加算を防ぐ）
  SESSION_MAX_HOURS: 12,

  // 残りがこれ未満なら can_start = false
  LOW_MARGIN_HOURS: 1.0,

  // 台帳の保持日数（これより古い行は purge で退避）
  RETENTION_DAYS: 90,

  ACCEL: {
    // reset_mode: 'weekly'（毎週定時リセット） | 'rolling'（直近7日ローリング）
    // reset_weekday: 0=Sun ... 6=Sat（UTC基準）
    GPU: { limit_hours: 30, reset_mode: 'weekly', reset_weekday: 6, reset_hour_utc: 0 },
    TPU: { limit_hours: 20, reset_mode: 'weekly', reset_weekday: 6, reset_hour_utc: 0 },
    CPU: { limit_hours: 0,  reset_mode: 'weekly', reset_weekday: 6, reset_hour_utc: 0 } // 0 = 無制限扱い
  }
};

var QUOTA_SHEET_NAME_ = 'usage_log';
var QUOTA_HEADER_ = ['session_id', 'accel', 'kernel', 'started_at',
                     'last_heartbeat_at', 'ended_at', 'seconds', 'status', 'note'];

/* ======================================================================
 * 低レベル: ロック / 台帳
 * ==================================================================== */

function quotaProps_() {
  return PropertiesService.getScriptProperties();
}

function quotaLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('quota: lock timeout');
  try { return fn(); } finally { lock.releaseLock(); }
}

function quotaSpreadsheet_() {
  var props = quotaProps_();
  var id = props.getProperty('QUOTA_SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* 消えていたら再作成 */ }
  }
  var ss = SpreadsheetApp.create('kaggle_quota_ledger');
  props.setProperty('QUOTA_SS_ID', ss.getId());
  return ss;
}

function quotaSheet_() {
  var ss = quotaSpreadsheet_();
  var sh = ss.getSheetByName(QUOTA_SHEET_NAME_);
  if (!sh) {
    sh = ss.insertSheet(QUOTA_SHEET_NAME_);
    sh.getRange(1, 1, 1, QUOTA_HEADER_.length)
      .setValues([QUOTA_HEADER_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function quotaParseDate_(v) {
  if (!v) return null;
  var d = (v instanceof Date) ? v : new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function quotaReadRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, QUOTA_HEADER_.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0] && !r[3]) continue;
    out.push({
      row: i + 2,
      session_id: String(r[0]),
      accel: String(r[1] || QUOTA_CONFIG.DEFAULT_ACCEL).toUpperCase(),
      kernel: String(r[2] || ''),
      started_at: quotaParseDate_(r[3]),
      last_hb: quotaParseDate_(r[4]),
      ended_at: quotaParseDate_(r[5]),
      seconds: Number(r[6] || 0),
      status: String(r[7] || ''),
      note: String(r[8] || '')
    });
  }
  return out;
}

function quotaMaxSessionSec_() {
  return QUOTA_CONFIG.SESSION_MAX_HOURS * 3600;
}

/* ======================================================================
 * セッション記録 (started / heartbeat / stopped)
 * ==================================================================== */

function quotaStartNoLock_(sh, sessionId, kernel, accel, startedAt, now) {
  var sec = Math.max(0, Math.min((now - startedAt) / 1000, quotaMaxSessionSec_()));
  sh.appendRow([
    String(sessionId),
    String(accel || QUOTA_CONFIG.DEFAULT_ACCEL).toUpperCase(),
    String(kernel || ''),
    startedAt.toISOString(),
    now.toISOString(),
    '',
    Math.round(sec),
    'RUNNING',
    ''
  ]);
  return { ok: true, created: true, row: sh.getLastRow() };
}

/**
 * セッション開始を記録。startedAtIso 省略時は現在時刻。
 * 同一 session_id の RUNNING 行があれば重複作成しない。
 */
function quotaSessionStart(sessionId, kernel, accel, startedAtIso) {
  return quotaLock_(function () {
    var sh = quotaSheet_();
    var now = new Date();
    var startedAt = quotaParseDate_(startedAtIso) || now;
    if (startedAt.getTime() > now.getTime()) startedAt = now;

    var rows = quotaReadRows_(sh);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].session_id === String(sessionId) && rows[i].status === 'RUNNING') {
        sh.getRange(rows[i].row, 5).setValue(now.toISOString());
        return { ok: true, dedup: true, row: rows[i].row };
      }
    }
    return quotaStartNoLock_(sh, sessionId, kernel, accel, startedAt, now);
  });
}

/**
 * heartbeat を記録し、稼働秒数を更新。
 * started を取りこぼしていた場合はこの時点でセッションを起こす。
 */
function quotaHeartbeat(sessionId, accel, kernel) {
  return quotaLock_(function () {
    var sh = quotaSheet_();
    var now = new Date();
    var rows = quotaReadRows_(sh);

    var target = null;
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].session_id === String(sessionId)) { target = rows[i]; break; }
    }
    if (!target) {
      return quotaStartNoLock_(sh, sessionId, kernel, accel, now, now);
    }
    if (target.status !== 'RUNNING') {
      // 一度 STALE で締めた後に復活した場合は新規セッション扱い
      return quotaStartNoLock_(sh, sessionId + '-r' + now.getTime(), kernel, accel, now, now);
    }
    var sec = Math.max(0, Math.min((now - target.started_at) / 1000, quotaMaxSessionSec_()));
    sh.getRange(target.row, 5, 1, 3)
      .setValues([[now.toISOString(), '', Math.round(sec)]]);
    return { ok: true, session_id: target.session_id, seconds: Math.round(sec) };
  });
}

/** セッション終了を記録。 */
function quotaSessionEnd(sessionId, note) {
  return quotaLock_(function () {
    var sh = quotaSheet_();
    var now = new Date();
    var rows = quotaReadRows_(sh);
    var closed = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.session_id !== String(sessionId) || r.status !== 'RUNNING') continue;
      var sec = Math.max(0, Math.min((now - r.started_at) / 1000, quotaMaxSessionSec_()));
      sh.getRange(r.row, 6, 1, 4)
        .setValues([[now.toISOString(), Math.round(sec), 'DONE', String(note || '')]]);
      closed++;
    }
    return { ok: true, closed: closed };
  });
}

/**
 * heartbeat 途絶 / 上限超過の RUNNING 行を締める。
 * トリガーから定期実行し、quotaStatus 内でも呼ばれる。
 */
function quotaReapStale_() {
  return quotaLock_(function () {
    var sh = quotaSheet_();
    var now = new Date();
    var rows = quotaReadRows_(sh);
    var maxSec = quotaMaxSessionSec_();
    var reaped = 0;

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.status !== 'RUNNING' || !r.started_at) continue;
      var hb = r.last_hb || r.started_at;
      var idle = (now - hb) / 1000;
      var dur = Math.max(0, (hb - r.started_at) / 1000);

      if (idle > QUOTA_CONFIG.STALE_HEARTBEAT_SEC) {
        // 最後の heartbeat 時点で切る（それ以降は動いていた証拠が無い）
        sh.getRange(r.row, 6, 1, 4).setValues([[
          hb.toISOString(), Math.round(Math.min(dur, maxSec)), 'STALE',
          'heartbeat idle ' + Math.round(idle) + 's'
        ]]);
        reaped++;
      } else if ((now - r.started_at) / 1000 >= maxSec) {
        sh.getRange(r.row, 6, 1, 4).setValues([[
          new Date(r.started_at.getTime() + maxSec * 1000).toISOString(),
          Math.round(maxSec), 'DONE', 'session max reached'
        ]]);
        reaped++;
      }
    }
    return { ok: true, reaped: reaped };
  });
}

/* ======================================================================
 * クォータ窓の計算
 * ==================================================================== */

function quotaWeeklyWindowStart_(now, weekday, hourUtc) {
  var d = new Date(Date.UTC(
    now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(),
    hourUtc || 0, 0, 0, 0));
  var diff = (d.getUTCDay() - weekday + 7) % 7;
  d.setUTCDate(d.getUTCDate() - diff);
  if (d.getTime() > now.getTime()) d.setUTCDate(d.getUTCDate() - 7);
  return d;
}

function quotaWindow_(cfg, now) {
  if (cfg.reset_mode === 'rolling') {
    return { start: new Date(now.getTime() - 7 * 86400000), end: now, next_reset: null };
  }
  var start = quotaWeeklyWindowStart_(now, cfg.reset_weekday, cfg.reset_hour_utc);
  var end = new Date(start.getTime() + 7 * 86400000);
  return { start: start, end: end, next_reset: end };
}

function quotaOverlapSec_(aStart, aEnd, bStart, bEnd) {
  var s = Math.max(aStart.getTime(), bStart.getTime());
  var e = Math.min(aEnd.getTime(), bEnd.getTime());
  return Math.max(0, (e - s) / 1000);
}

/* ======================================================================
 * キャリブレーション（Kaggle UI の実測値で基準を打ち直す）
 * ==================================================================== */

function quotaBaselineKey_(accel) { return 'QUOTA_BASELINE_' + accel; }

function quotaGetBaseline_(accel) {
  var raw = quotaProps_().getProperty(quotaBaselineKey_(accel));
  if (!raw) return { at: null, used_sec: 0 };
  try {
    var o = JSON.parse(raw);
    return { at: quotaParseDate_(o.at), used_sec: Number(o.used_sec || 0) };
  } catch (err) {
    return { at: null, used_sec: 0 };
  }
}

/**
 * 「今この瞬間、Kaggle UI では usedHours 消費済み」と宣言する。
 * 以降の消費はこの基準点からの積み上げになる。
 */
function quotaCalibrate(accel, usedHours) {
  var A = String(accel || QUOTA_CONFIG.DEFAULT_ACCEL).toUpperCase();
  if (!QUOTA_CONFIG.ACCEL[A]) throw new Error('unknown accel: ' + A);
  var used = Number(usedHours);
  if (isNaN(used) || used < 0) throw new Error('used_hours must be a non-negative number');

  var payload = { at: new Date().toISOString(), used_sec: Math.round(used * 3600) };
  quotaProps_().setProperty(quotaBaselineKey_(A), JSON.stringify(payload));
  return { ok: true, accel: A, baseline: payload };
}

function quotaClearBaseline(accel) {
  var A = String(accel || QUOTA_CONFIG.DEFAULT_ACCEL).toUpperCase();
  quotaProps_().deleteProperty(quotaBaselineKey_(A));
  return { ok: true, accel: A, cleared: true };
}

/* ======================================================================
 * 残量の算出（メイン）
 * ==================================================================== */

function quotaStatus(accel) {
  var A = String(accel || QUOTA_CONFIG.DEFAULT_ACCEL).toUpperCase();
  var cfg = QUOTA_CONFIG.ACCEL[A];
  if (!cfg) throw new Error('unknown accel: ' + A);

  quotaReapStale_();

  var now = new Date();
  var win = quotaWindow_(cfg, now);

  // 基準点が現在の窓の中にあるときだけ有効
  var base = quotaGetBaseline_(A);
  var effStart = win.start;
  var usedSec = 0;
  var baselineInfo = null;
  if (base.at && base.at.getTime() > win.start.getTime() && base.at.getTime() <= now.getTime()) {
    effStart = base.at;
    usedSec = base.used_sec;
    baselineInfo = { at: base.at.toISOString(), used_hours: +(base.used_sec / 3600).toFixed(3) };
  }

  var rows = quotaReadRows_(quotaSheet_());
  var running = [];
  var sessions = 0;

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.accel !== A || !r.started_at) continue;
    var end;
    if (r.ended_at) end = r.ended_at;
    else if (r.status === 'RUNNING') end = now;
    else end = (r.last_hb || r.started_at);

    var ov = quotaOverlapSec_(r.started_at, end, effStart, now);
    if (ov > 0) sessions++;
    usedSec += ov;

    if (r.status === 'RUNNING') {
      running.push({
        session_id: r.session_id,
        kernel: r.kernel,
        started_at: r.started_at.toISOString(),
        last_heartbeat_at: r.last_hb ? r.last_hb.toISOString() : null,
        elapsed_sec: Math.round((now - r.started_at) / 1000)
      });
    }
  }

  var limitSec = cfg.limit_hours * 3600;
  var unlimited = !(limitSec > 0);
  var remainSec = unlimited ? null : Math.max(0, limitSec - usedSec);

  return {
    ok: true,
    accel: A,
    unlimited: unlimited,
    limit_hours: cfg.limit_hours,
    used_hours: +(usedSec / 3600).toFixed(3),
    used_sec: Math.round(usedSec),
    remaining_hours: unlimited ? null : +(remainSec / 3600).toFixed(3),
    remaining_sec: unlimited ? null : Math.round(remainSec),
    remaining_pct: unlimited ? null : +(remainSec / limitSec * 100).toFixed(1),
    window: {
      mode: cfg.reset_mode,
      start: win.start.toISOString(),
      end: win.end ? win.end.toISOString() : null,
      next_reset: win.next_reset ? win.next_reset.toISOString() : null,
      next_reset_in_sec: win.next_reset
        ? Math.max(0, Math.round((win.next_reset - now) / 1000)) : null
    },
    baseline: baselineInfo,
    sessions_in_window: sessions,
    running: running,
    can_start: unlimited || (remainSec / 3600) >= QUOTA_CONFIG.LOW_MARGIN_HOURS,
    low_margin_hours: QUOTA_CONFIG.LOW_MARGIN_HOURS,
    source: 'gas-ledger (Kaggle has no public quota API)',
    checked_at: now.toISOString()
  };
}

function quotaStatusAll() {
  var out = { ok: true, checked_at: new Date().toISOString(), accelerators: {} };
  var keys = Object.keys(QUOTA_CONFIG.ACCEL);
  for (var i = 0; i < keys.length; i++) {
    try { out.accelerators[keys[i]] = quotaStatus(keys[i]); }
    catch (err) { out.accelerators[keys[i]] = { ok: false, error: String(err) }; }
  }
  return out;
}

/**
 * 起動前チェック。needHours 分の稼働余地があるか。
 * 既存コントローラの push 前に呼ぶことを推奨。
 */
function quotaCanStart(accel, needHours) {
  var st = quotaStatus(accel);
  var need = Number(needHours || QUOTA_CONFIG.LOW_MARGIN_HOURS);
  var allowed = st.unlimited || st.remaining_hours >= need;
  return {
    ok: true, allowed: allowed, accel: st.accel,
    need_hours: need, remaining_hours: st.remaining_hours,
    reason: allowed ? 'ok' : 'insufficient quota'
  };
}

/* ======================================================================
 * 台帳の整理
 * ==================================================================== */

function quotaPurgeOld_() {
  return quotaLock_(function () {
    var sh = quotaSheet_();
    var cutoff = new Date(Date.now() - QUOTA_CONFIG.RETENTION_DAYS * 86400000);
    var rows = quotaReadRows_(sh);
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i];
      if (r.status === 'RUNNING') continue;
      var ref = r.ended_at || r.last_hb || r.started_at;
      if (ref && ref.getTime() < cutoff.getTime()) { sh.deleteRow(r.row); deleted++; }
    }
    return { ok: true, deleted: deleted };
  });
}

function quotaLog(limit) {
  var n = Math.min(Math.max(parseInt(limit || 30, 10) || 30, 1), 200);
  var rows = quotaReadRows_(quotaSheet_());
  var out = rows.slice(Math.max(0, rows.length - n)).map(function (r) {
    return {
      session_id: r.session_id, accel: r.accel, kernel: r.kernel,
      started_at: r.started_at ? r.started_at.toISOString() : null,
      ended_at: r.ended_at ? r.ended_at.toISOString() : null,
      hours: +(r.seconds / 3600).toFixed(3),
      status: r.status, note: r.note
    };
  }).reverse();
  return { ok: true, count: out.length, rows: out };
}

/* ======================================================================
 * HTTP API
 * ==================================================================== */

function quotaJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function quotaRequireToken_(p) {
  var expected = QUOTA_CONFIG.API_TOKEN;
  if (!expected || expected === 'CHANGE_ME_TO_A_LONG_RANDOM_STRING') {
    throw new Error('QUOTA_CONFIG.API_TOKEN is not configured');
  }
  if (String(p.token || '') !== expected) throw new Error('unauthorized');
}

/**
 * 既存 doGet / doPost の先頭で呼ぶ。
 * 自分の担当 action なら TextOutput を返し、そうでなければ null を返す。
 */
function quotaHandle(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || '');
  if (action.indexOf('quota') !== 0) return null;

  try {
    switch (action) {
      case 'quota':
        return quotaJson_(quotaStatus(p.accel));

      case 'quota_all':
        return quotaJson_(quotaStatusAll());

      case 'quota_can_start':
        return quotaJson_(quotaCanStart(p.accel, p.need_hours));

      case 'quota_log':
        return quotaJson_(quotaLog(p.limit));

      case 'quota_session_start':
        if (!p.session_id) throw new Error('session_id required');
        return quotaJson_(quotaSessionStart(p.session_id, p.kernel, p.accel, p.started_at));

      case 'quota_heartbeat':
        if (!p.session_id) throw new Error('session_id required');
        return quotaJson_(quotaHeartbeat(p.session_id, p.accel, p.kernel));

      case 'quota_end':
        if (!p.session_id) throw new Error('session_id required');
        return quotaJson_(quotaSessionEnd(p.session_id, p.note));

      case 'quota_calibrate':
        quotaRequireToken_(p);
        return quotaJson_(quotaCalibrate(p.accel, p.used_hours));

      case 'quota_clear_baseline':
        quotaRequireToken_(p);
        return quotaJson_(quotaClearBaseline(p.accel));

      case 'quota_reap':
        quotaRequireToken_(p);
        return quotaJson_(quotaReapStale_());

      case 'quota_purge':
        quotaRequireToken_(p);
        return quotaJson_(quotaPurgeOld_());

      case 'quota_ledger_url':
        quotaRequireToken_(p);
        return quotaJson_({ ok: true, url: quotaSpreadsheet_().getUrl() });

      default:
        return quotaJson_({ ok: false, error: 'unknown quota action: ' + action });
    }
  } catch (err) {
    return quotaJson_({ ok: false, error: String(err && err.message || err), action: action });
  }
}

/**
 * このモジュールを単独プロジェクトとして立てる場合のみ有効化。
 * 既存 kaggle_controller_single.gs に同居させる場合はコメントアウトのまま。
 */
// function doGet(e)  { var r = quotaHandle(e); return r || quotaJson_({ ok:false, error:'no action' }); }
// function doPost(e) { return doGet(e); }

/* ======================================================================
 * トリガー / 動作確認
 * ==================================================================== */

function installQuotaTriggers() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    var fn = existing[i].getHandlerFunction();
    if (fn === 'quotaReapStale_' || fn === 'quotaPurgeOld_') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  ScriptApp.newTrigger('quotaReapStale_').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('quotaPurgeOld_').timeBased().everyDays(1).atHour(4).create();
  Logger.log('triggers installed. ledger: ' + quotaSpreadsheet_().getUrl());
}

function quotaSelfTest() {
  var sid = 'selftest-' + Date.now();
  Logger.log('ledger: ' + quotaSpreadsheet_().getUrl());
  Logger.log('start: ' + JSON.stringify(quotaSessionStart(sid, 'matsuda2026/qwen3-coder-30b', 'GPU')));
  Logger.log('hb   : ' + JSON.stringify(quotaHeartbeat(sid, 'GPU')));
  Logger.log('end  : ' + JSON.stringify(quotaSessionEnd(sid, 'selftest')));
  Logger.log('status: ' + JSON.stringify(quotaStatus('GPU'), null, 2));
}
