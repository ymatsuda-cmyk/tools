/**
 * ドル円 かんたん投資 — 通知用 Google Apps Script
 *
 * アプリを閉じていても「買い時 / 売り時」をメールで知らせるためのスクリプト。
 * ブラウザ側アプリ（fx-invest/index.html）と同じ判定ロジックを使う。
 *
 * 判定: いまのレートが「いつもの値段（過去75営業日の平均）」から
 *       どれだけ離れたかで、買い時 / 売り時 / 様子見を決める。
 *
 * 使い方:
 *  1. Googleドライブ → 新規 → その他 → Google Apps Script でプロジェクトを作る
 *  2. このファイルの中身をすべて貼り付ける
 *  3. 「プロジェクトの設定」→「スクリプト プロパティ」に以下を登録
 *       FX_MAIL_TO    … 通知先メールアドレス（必須）
 *       FX_LEVEL      … easy / normal / hard （省略時 normal）
 *       FX_LOT_JPY    … 1回の金額。メール本文の表示に使う（省略時 1000000）
 *       FX_WEBHOOK_URL… Slack等のWebhook URL（任意。{"text":"..."} を送る）
 *  4. エディタで setupFxTrigger を1回実行する（毎日夕方に自動チェックされる）
 *
 * 動作確認は testFxSignal を実行してログを見る。
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
    date: live === null ? close.date
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
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

/** サインの状態を忘れさせる。次回のチェックで必ず通知が飛ぶ */
function resetFxState() {
  FXP.deleteProperty('FX_LAST_STATE');
  Logger.log('状態をリセットしました');
}
