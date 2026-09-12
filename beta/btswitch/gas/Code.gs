/**
 * イヤホン切替(btswitch)用 GAS
 *
 * 役割:
 *  - 「いまどの端末がイヤホンを使うか」を1か所に持つ。
 *    ブラウザ(index.html)が希望を書き、各PCの常駐エージェントがそれを読んで
 *    自分の端末で接続 / 切断する。Bluetoothはブラウザから他端末を操作できないため、
 *    この「掲示板」を挟む形にしている。
 *  - スマホにはエージェントを置けないので、owner を phone にすると
 *    すべてのPCが切断し、スマホ側が自動で掴める状態を作る。
 *
 * 事前設定 (スクリプトプロパティ):
 *  - ACCESS_TOKEN  画面とエージェントで共有する合言葉(自分で決める)
 *
 * デプロイ:
 *  - 種類: ウェブアプリ / 実行するユーザー: 自分 / アクセス: 全員
 *    (URLを知っていれば誰でも叩けるため ACCESS_TOKEN のチェックを必ず通す)
 */

var STATE_KEY = 'BT_STATE';
var HISTORY_LIMIT = 20;

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    assertToken_(body);

    var result;
    switch (body.action) {
      case 'getState':
        result = getState_();
        break;
      case 'claim':
        result = claim_(body.deviceId, body.by);
        break;
      case 'release':
        result = claim_('', body.by);
        break;
      case 'heartbeat':
        result = heartbeat_(body.deviceId, body.label, body.connected, body.note);
        break;
      case 'forget':
        result = forget_(body.deviceId);
        break;
      default:
        throw new Error('unknown action: ' + body.action);
    }
    return json_({ ok: true, data: result });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

/** エージェントからは GET でも状態を取れるようにしておく(curl で試しやすい) */
function doGet(e) {
  try {
    assertToken_({ token: (e.parameter || {}).token });
    return json_({ ok: true, data: getState_() });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function assertToken_(body) {
  var expected = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
  if (!expected || (body && body.token) !== expected) throw new Error('unauthorized');
}

function load_() {
  var raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  if (!raw) return { owner: '', updatedAt: null, updatedBy: '', devices: {}, history: [] };
  try {
    var s = JSON.parse(raw);
    s.devices = s.devices || {};
    s.history = s.history || [];
    return s;
  } catch (e) {
    return { owner: '', updatedAt: null, updatedBy: '', devices: {}, history: [] };
  }
}

function save_(state) {
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return state;
}

function getState_() {
  return load_();
}

/** 切替の希望を書く。実際の接続 / 切断は各エージェントが自分の端末で行う */
function claim_(deviceId, by) {
  var state = load_();
  var id = String(deviceId || '');
  if (state.owner !== id) {
    state.history.unshift({ owner: id, at: new Date().toISOString(), by: String(by || '') });
    state.history = state.history.slice(0, HISTORY_LIMIT);
  }
  state.owner = id;
  state.updatedAt = new Date().toISOString();
  state.updatedBy = String(by || '');
  return save_(state);
}

/**
 * エージェントの生存報告。ついでに自分の接続状態を書き戻す。
 * 画面はこれを見て「本当につながっているか」を出す。
 */
function heartbeat_(deviceId, label, connected, note) {
  var id = String(deviceId || '');
  if (!id) throw new Error('deviceId は必須です');
  var state = load_();
  state.devices[id] = {
    id: id,
    label: String(label || id),
    connected: Boolean(connected),
    note: String(note || ''),
    lastSeenAt: new Date().toISOString(),
  };
  return save_(state);
}

function forget_(deviceId) {
  var state = load_();
  delete state.devices[String(deviceId || '')];
  return save_(state);
}
