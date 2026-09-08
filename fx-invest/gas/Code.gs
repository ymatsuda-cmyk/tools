/**
 * ドル円 かんたん投資 — 通知＋売買履歴の保存用 Google Apps Script
 *
 * 役割:
 *  1. アプリを閉じていても「買い時 / 売り時」をメールで知らせる
 *  2. 売買履歴をスプレッドシートに保存し、端末間で同期する
 *
 * 判定: いまのレートが「いつもの値段（過去75営業日の平均）」から
 *       どれだけ離れたかで、買い時 / 売り時 / 様子見を決める。
 *
 * 使い方:
 *  1. Googleドライブ → 新規 → その他 → Google Apps Script でプロジェクトを作る
 *  2. このファイルの中身をすべて貼り付ける
 *  3. 「プロジェクトの設定」→「スクリプト プロパティ」に以下を登録
 *       FX_MAIL_TO    … 通知先メールアドレス
 *       FX_SECRET     … アプリから接続するときの合言葉（同期を使うなら必須）
 *       FX_LEVEL      … easy / normal / hard （省略時 normal）
 *       FX_LOT_JPY    … 1回の金額。メール本文の表示に使う（省略時 1000000）
 *       FX_SHEET_ID   … 保存先のスプレッドシートID（省略時は自動作成）
 *       FX_WEBHOOK_URL… Slack等のWebhook URL（任意。{"text":"..."} を送る）
 *  4. setupFxTrigger を1回実行する（毎日夕方に自動チェック）
 *  5. 端末間で同期する場合は「デプロイ」→「新しいデプロイ」→ 種類：ウェブアプリ
 *     アクセスできるユーザー：全員 でデプロイし、発行された /exec のURLと
 *     FX_SECRET をアプリの設定画面に入力する
 *
 * 動作確認は testFxSignal / setupFxSheet を実行してログを見る。
 */

var FXP = PropertiesService.getScriptProperties();

var FX_SMA_DAYS = 75;
var FX_LEVELS = {
  easy:   { buy: 1.5, sell: 2.5 },
  normal: { buy: 3.0, sell: 5.0 },
  hard:   { buy: 5.0, sell: 8.0 }
};

function fxCfg(key, fallback) {
  var v = FXP.getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

/* ============================================================
   レート取得（Frankfurter / ECB基準・無料・APIキー不要）
   ============================================================ */

function fetchUsdJpySeries(days) {
  var tz = 'UTC';
  var end = new Date();
  var start = new Date(end.getTime() - days * 86400000);
  var range = Utilities.formatDate(start, tz, 'yyyy-MM-dd') + '..' +
              Utilities.formatDate(end, tz, 'yyyy-MM-dd');

  var hosts = ['https://api.frankfurter.app', 'https://api.frankfurter.dev/v1'];
  var lastErr = '';

  for (var i = 0; i < hosts.length; i++) {
    try {
      var res = UrlFetchApp.fetch(hosts[i] + '/' + range + '?from=USD&to=JPY', {
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());

      var rates = JSON.parse(res.getContentText()).rates || {};
      var dates = Object.keys(rates).sort();
      if (!dates.length) throw new Error('データが空です');

      return dates.map(function (d) { return { date: d, rate: rates[d].JPY }; });
    } catch (e) {
      lastErr = String(e.message || e);
    }
  }
  throw new Error('レートを取得できませんでした: ' + lastErr);
}

/* ============================================================
   いまのレート（数分おきに更新される無料API）

   Frankfurter はECBの1日1回の基準レートなので、現在値は別途取得する。
   取得できなかったときは直近の終値で代用する。
   ============================================================ */

function fetchUsdJpyLive() {
  var sources = [
    { url: 'https://api.fxratesapi.com/latest?base=USD&currencies=JPY',
      pick: function (j) { return j.rates && j.rates.JPY; } },
    { url: 'https://open.er-api.com/v6/latest/USD',
      pick: function (j) { return j.rates && j.rates.JPY; } }
  ];

  for (var i = 0; i < sources.length; i++) {
    try {
      var res = UrlFetchApp.fetch(sources[i].url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      var rate = sources[i].pick(JSON.parse(res.getContentText()));
      if (typeof rate === 'number' && rate > 0) return rate;
    } catch (e) { /* 次の取得先を試す */ }
  }
  return null;
}

/* ============================================================
   判定
   ============================================================ */

function fxJudge() {
  var series = fetchUsdJpySeries(200);
  if (series.length < FX_SMA_DAYS) throw new Error('データが足りません');

  var recent = series.slice(series.length - FX_SMA_DAYS);
  var sum = recent.reduce(function (a, x) { return a + x.rate; }, 0);
  var avg = sum / FX_SMA_DAYS;

  var close = series[series.length - 1];
  var live = fetchUsdJpyLive();
  var rate = live === null ? close.rate : live;

  var gap = (rate - avg) / avg * 100;
  var level = FX_LEVELS[fxCfg('FX_LEVEL', 'normal')] || FX_LEVELS.normal;

  var state = 'hold';
  if (gap <= -level.buy) state = 'buy';
  else if (gap >= level.sell) state = 'sell';

  return {
    state: state, gap: gap, rate: rate, avg: avg, live: live !== null,
    prevClose: close.rate, prevDate: close.date,
    date: live === null ? close.date
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

/** 判定と建値の損益をまとめて返す（ダッシュボードのカード用） */
function fxSummary() {
  var j = fxJudge();
  var rows = fxListTrades();

  var profit = 0;
  var sum = 0;
  rows.forEach(function (t) {
    if (!t.rate) return;
    // 売りは建値で売って現在値で買い戻すので、値幅の符号が逆になる
    var diff = t.side === 'sell' ? t.rate - j.rate : j.rate - t.rate;
    profit += FX_POS_UNITS * diff;
    sum += t.rate;
  });

  return {
    signal: j.state, gap: j.gap, rate: j.rate, avg: j.avg,
    live: j.live, date: j.date, prevClose: j.prevClose,
    qty: rows.length,
    avgCost: rows.length ? sum / rows.length : null,
    profit: profit,
    units: rows.length * FX_POS_UNITS,
    trades: rows.length
  };
}

/* ============================================================
   通知
   ============================================================ */

function fxMessage(j) {
  var lot = Number(fxCfg('FX_LOT_JPY', 1000000));
  var word = j.state === 'buy' ? '安い' : '高い';
  var head = j.state === 'buy' ? '【買い時】ドル円' : '【売り時】ドル円';

  var body = [
    (j.state === 'buy' ? '買い時のサインが出ました。' : '売り時のサインが出ました。'),
    '',
    'レート    : ' + j.rate.toFixed(2) + ' 円（' + j.date + (j.live ? ' 時点' : ' 終値') + '）',
    'いつもの値段: ' + j.avg.toFixed(2) + ' 円（過去' + FX_SMA_DAYS + '営業日の平均）',
    'かい離    : ' + Math.abs(j.gap).toFixed(1) + '% ' + word,
    '',
    'アプリで「' + (j.state === 'buy' ? '買う' : '売る') + ' 1」（' +
      Math.round(lot).toLocaleString() + '円ぶん）を記録してください。',
    '',
    '※ これは投資助言ではありません。発注はご自身の判断で行ってください。'
  ].join('\n');

  return { subject: head + ' ' + j.rate.toFixed(2) + '円', body: body };
}

function fxSend(j) {
  var msg = fxMessage(j);

  var to = fxCfg('FX_MAIL_TO');
  if (to) {
    MailApp.sendEmail({ to: to, subject: msg.subject, body: msg.body });
  }

  var hook = fxCfg('FX_WEBHOOK_URL');
  if (hook) {
    UrlFetchApp.fetch(hook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: msg.subject + '\n' + msg.body }),
      muteHttpExceptions: true
    });
  }

  if (!to && !hook) Logger.log('FX_MAIL_TO も FX_WEBHOOK_URL も未設定です');
}

/* ============================================================
   建値の保存（スプレッドシート）

   1行 = 「この値段で買った / 売った」という記録。
   id はアプリ側で付け、同じ id は上書きするので
   複数端末から送っても重複しない。
   ============================================================ */

var FX_HEADER = ['id', 'at', 'date', 'side', 'rate', 'jpy', 'usd', 'pl'];
var FX_POS_UNITS = 10000;   // 1件あたりの数量（アプリの POS_UNITS と揃える）

function fxSheet() {
  var id = fxCfg('FX_SHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('ドル円 かんたん投資 売買履歴');
    FXP.setProperty('FX_SHEET_ID', ss.getId());
  }
  var sh = ss.getSheetByName('trades') || ss.insertSheet('trades');
  if (sh.getLastRow() === 0) sh.appendRow(FX_HEADER);
  return sh;
}

/** シートが日付型に変換してしまった値を文字列に戻す */
function fxCell(v, pattern) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern);
  }
  return String(v);
}

function fxListTrades() {
  var sh = fxSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, FX_HEADER.length).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      return {
        id:   fxCell(r[0], 'yyyy-MM-dd'),
        at:   fxCell(r[1], "yyyy-MM-dd'T'HH:mm:ss.SSSXXX"),
        date: fxCell(r[2], 'yyyy-MM-dd'),
        side: String(r[3]),
        rate: Number(r[4]),
        jpy:  Number(r[5]),
        usd:  Number(r[6]),
        pl:   (r[7] === '' || r[7] === null) ? null : Number(r[7])
      };
    })
    .sort(function (a, b) { return a.at < b.at ? -1 : (a.at > b.at ? 1 : 0); });
}

function fxAddTrade(t) {
  if (!t || !t.id) throw new Error('id のない記録は保存できません');
  if (t.side !== 'buy' && t.side !== 'sell') throw new Error('side が不正です');

  var row = [
    String(t.id),
    String(t.at || new Date().toISOString()),
    String(t.date || ''),
    String(t.side),
    Number(t.rate) || 0,
    Number(t.jpy) || 0,
    Number(t.usd) || 0,
    (t.pl === null || t.pl === undefined || t.pl === '') ? '' : Number(t.pl)
  ];

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = fxSheet();
    var at = fxRowIndex(sh, t.id);
    if (at > 0) sh.getRange(at, 1, 1, FX_HEADER.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return fxListTrades();
}

/** id のある行番号を返す。見つからなければ 0 */
function fxRowIndex(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function fxRemoveTrade(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = fxSheet();
    var at = fxRowIndex(sh, id);
    if (at > 0) sh.deleteRow(at);
  } finally {
    lock.releaseLock();
  }
  return fxListTrades();
}

function fxClearTrades() {
  var sh = fxSheet();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  return [];
}

/* ============================================================
   Web API（アプリから呼ぶ）

   ブラウザの事前確認（preflight）を避けるため、アプリ側は
   Content-Type: text/plain で POST してくる。
   ============================================================ */

function fxJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fxHandle(req) {
  try {
    var secret = fxCfg('FX_SECRET');
    if (!secret) throw new Error('FX_SECRET が未設定です');
    if (String(req.secret || '') !== String(secret)) throw new Error('合言葉が違います');

    if (req.action === 'list')    return fxJson({ ok: true, trades: fxListTrades() });
    if (req.action === 'add')     return fxJson({ ok: true, trades: fxAddTrade(req.trade) });
    if (req.action === 'remove')  return fxJson({ ok: true, trades: fxRemoveTrade(req.id) });
    if (req.action === 'clear')   return fxJson({ ok: true, trades: fxClearTrades() });
    if (req.action === 'signal')  return fxJson({ ok: true, signal: fxJudge() });
    if (req.action === 'summary') return fxJson({ ok: true, summary: fxSummary() });

    return fxJson({ ok: false, error: '不明な action: ' + req.action });
  } catch (err) {
    return fxJson({ ok: false, error: String(err.message || err) });
  }
}

function doGet(e) {
  return fxHandle((e && e.parameter) || {});
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { /* 不正なJSONは空扱い */ }
  return fxHandle(req);
}

/* ============================================================
   定期実行の入口
   ============================================================ */

/**
 * トリガーから呼ばれる。状態が変わったときだけ通知するので、
 * 同じサインで毎日メールが届くことはない。
 */
function checkFxSignal() {
  var j = fxJudge();
  var prev = FXP.getProperty('FX_LAST_STATE') || 'hold';

  if (j.state === prev) return;
  FXP.setProperty('FX_LAST_STATE', j.state);

  if (j.state === 'hold') return;
  fxSend(j);
  Logger.log('通知: ' + j.state + ' / ' + j.rate.toFixed(2));
}

/** 毎日18時台にチェックするトリガーを作る（重複登録はしない） */
function setupFxTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'checkFxSignal';
  });
  if (exists) {
    Logger.log('トリガーは登録済みです');
    return;
  }
  ScriptApp.newTrigger('checkFxSignal').timeBased().atHour(18).everyDays(1).create();
  Logger.log('毎日18時台のトリガーを登録しました');
}

/* ============================================================
   動作確認用
   ============================================================ */

function testFxSignal() {
  var j = fxJudge();
  Logger.log(JSON.stringify(j, null, 2));
  Logger.log(fxMessage(j).body);
}

/** 保存先シートを作成（または確認）してURLをログに出す */
function setupFxSheet() {
  var sh = fxSheet();
  Logger.log('シート: ' + sh.getParent().getUrl());
  Logger.log('保存済みの売買: ' + fxListTrades().length + '件');
}

/** 履歴を全件消す（エディタから手動で実行する用） */
function clearFxTrades() {
  fxClearTrades();
  Logger.log('履歴を消しました');
}

/** サインの状態を忘れさせる。次回のチェックで必ず通知が飛ぶ */
function resetFxState() {
  FXP.deleteProperty('FX_LAST_STATE');
  Logger.log('状態をリセットしました');
}
