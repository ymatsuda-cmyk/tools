/**
 * 議事録ビューア用 GAS プロキシ
 *
 * 役割:
 *  - Notion API はブラウザから直接叩けない(CORSヘッダーなし)ため、
 *    このスクリプトが仲介する。
 *  - 音声アップロード用に、Drive の resumable upload セッションを発行する
 *    (このスクリプト自身の権限で発行し、実際のバイト送信はブラウザから
 *    Google へ直接行う。利用者は Google にログイン不要)。
 *  - フロントは text/plain で JSON を POST する(プリフライトを発生させないため)。
 *    Content-Type を application/json にすると OPTIONS preflight が飛び、
 *    GAS は OPTIONS に応答できないため必ず失敗する。
 *
 * 事前設定 (スクリプトプロパティ):
 *  - NOTION_TOKEN            Notion Integration のシークレット
 *  - ACCESS_TOKEN            フロントの minutes:config.notionToken と一致させる共有トークン
 *  - AUDIO_INBOX_FOLDER_ID   音声アップロード先の Drive フォルダ ID
 *  - code                    コードと権限の対応表 (JSON、verifyCode_ 参照)
 *
 * マニフェスト (appsscript.json) に Drive スコープが必要:
 *   "oauthScopes": [
 *     "https://www.googleapis.com/auth/script.external_request",
 *     "https://www.googleapis.com/auth/drive"
 *   ]
 *
 *   ※ マニフェストに書くだけでは不十分な場合がある。Apps Script はコードを
 *   静的解析して実際に使うサービスからスコープを決めるため、UrlFetchApp で
 *   Drive API を直接叩くだけのコードだと「Driveを使っている」と認識されず、
 *   ACCESS_TOKEN_SCOPE_INSUFFICIENT (403) になることがある。
 *   その場合は「サービス」→「Drive API」(高度なサービス) を追加し、
 *   touchDriveScope_() のように実際に呼び出すコードを含めること。
 *
 * デプロイ:
 *  - 種類: ウェブアプリ
 *  - 実行するユーザー: 自分
 *  - アクセスできるユーザー: 全員
 *  (URLを知っていれば誰でも叩けるため、ACCESS_TOKEN のチェックを必ず通す)
 */

var NOTION_VERSION = '2022-06-28';

// プロパティ名(Notion側のカラム名とここを一致させること)
var PROP_SUMMARY   = '要約';       // カード用の短い要約 (rich_text)
var PROP_DECISIONS = '決定事項';   // 改行区切り (rich_text)
var PROP_TODOS     = 'ToDo';       // 改行区切り (rich_text)
var PROP_TOPICS    = '論点';       // 改行区切り (rich_text)
var PROP_MODEL     = '要約モデル'; // 生成に使ったモデル名 (rich_text)
var PROP_GENERATED = '要約日時';   // 生成日時、鮮度判定に使用 (date)
var PROP_STATUS    = '状態';       // 進捗ステータス (select)
var STATUS_SUMMARIZED = '要約';    // 要約生成完了時にセットする値
var STATUS_RETRANSCRIBE = '再取得'; // 再文字起こしをMac mini側バッチに依頼する際にセットする値
var STATUS_DELETED = '削除';       // 削除ボタン押下時にセットする値(Notionページ自体は残す)
var PROP_CATEGORY  = 'カテゴリー'; // タグ (multi_select、自由入力可)
var PROP_AGENDA    = '議事';       // 議題ごとの経緯 (rich_text、JSON文字列で格納)
var PROP_MEMO      = 'メモ';       // 自由記述のメモ (rich_text)
var PROP_RAW_COUNT = '原文文字数'; // 文字起こし全文の文字数キャッシュ (number)
var PROP_TITLE     = 'ミーティング名'; // タイトル (title)
var PROP_PERMISSION = '権限';      // 閲覧権限 (multi_select)
var ADMIN_ROLE     = 'xYz';        // 全機能を使える管理者権限

// ============ エントリーポイント ============

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);

    // コード検証だけは共有トークン不要で受ける
    // (初回はまだトークンが手元に無い状態で権限を問い合わせるため)
    if (body.action === 'verifyCode') {
      return jsonOutput_({ ok: true, data: verifyCode_(body.code) });
    }

    assertToken_(body);

    switch (body.action) {
      case 'fetchTranscript':
        result = fetchTranscript_(body.pageId);
        break;
      case 'fetchSummary':
        result = fetchSummary_(body.pageId);
        break;
      case 'saveSummary':
        result = saveSummary_(body.pageId, body.cardSummary, body.detail, body.model, body.rawContextCount);
        break;
      case 'saveTags':
        result = saveTags_(body.pageId, body.tags);
        break;
      case 'saveTitle':
        result = saveTitle_(body.pageId, body.title);
        break;
      case 'saveDetail':
        result = saveDetail_(body.pageId, body.cardSummary, body.detail);
        break;
      case 'requestRetranscribe':
        result = requestRetranscribe_(body.pageId);
        break;
      case 'deleteItem':
        result = deleteItem_(body.pageId);
        break;
      case 'savePermissions':
        result = savePermissions_(body.pageIds, body.permissions, body.mode);
        break;
      case 'saveMemo':
        result = saveMemo_(body.pageId, body.memo);
        break;
      case 'updateRawContextCount':
        result = updateRawContextCount_(body.pageId, body.count);
        break;
      case 'initUpload':
        result = initUpload_(body);
        break;
      case 'writeSidecar':
        result = writeSidecar_(body);
        break;
      default:
        throw new Error('unknown action: ' + body.action);
    }
    return jsonOutput_({ ok: true, data: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function assertToken_(body) {
  var expected = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
  if (!expected || body.token !== expected) {
    throw new Error('unauthorized');
  }
}

// ============ Notion 共通ヘルパー ============

function notionHeaders_() {
  var token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  return {
    Authorization: 'Bearer ' + token,
    'Notion-Version': NOTION_VERSION,
    'Content-Type': 'application/json',
  };
}

function notionFetch_(path, method, payload) {
  var options = {
    method: method || 'get',
    headers: notionHeaders_(),
    muteHttpExceptions: true,
  };
  if (payload) {
    options.payload = JSON.stringify(payload);
  }
  var res = UrlFetchApp.fetch('https://api.notion.com/v1/' + path, options);
  var code = res.getResponseCode();
  var json = JSON.parse(res.getContentText());
  if (code >= 300) {
    throw new Error('Notion API ' + code + ': ' + (json.message || res.getContentText()));
  }
  return json;
}

/** ページ本文のブロックを全件取得する(ページネーション対応) */
function fetchAllBlocks_(blockId) {
  var blocks = [];
  var cursor = null;
  do {
    var path = 'blocks/' + blockId + '/children?page_size=100';
    if (cursor) path += '&start_cursor=' + cursor;
    var res = notionFetch_(path, 'get');
    blocks = blocks.concat(res.results);
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return blocks;
}

/** リッチテキスト配列からプレーンテキストを連結する */
function plainTextOf_(richTextArray) {
  return (richTextArray || []).map(function (t) { return t.plain_text; }).join('');
}

// ============ アクション実装 ============

/**
 * 文字起こし全文を取得する。
 * ページ本文のうち paragraph / heading / bulleted_list_item 等のテキスト系ブロックを
 * 上から連結して返す。updatedAt はページの last_edited_time。
 */
function fetchTranscript_(pageId) {
  var page = notionFetch_('pages/' + pageId, 'get');
  var blocks = fetchAllBlocks_(pageId);

  var lines = [];
  blocks.forEach(function (b) {
    var rt = b[b.type] && b[b.type].rich_text;
    if (rt) {
      var text = plainTextOf_(rt);
      if (text) lines.push(text);
    }
  });

  return {
    text: lines.join('\n'),
    updatedAt: page.last_edited_time,
  };
}

/**
 * 既存の要約を取得する。
 * カード要約・決定事項・ToDo・論点・モデル・生成日時、すべて専用プロパティから読む。
 * "要約日時" が空なら未生成として null 相当を返す(フロント側の判定に使う)。
 */
function fetchSummary_(pageId) {
  var page = notionFetch_('pages/' + pageId, 'get');
  var props = page.properties;

  var generatedAt = props[PROP_GENERATED] && props[PROP_GENERATED].date
    ? props[PROP_GENERATED].date.start
    : null;

  // 生成日時が無ければ未生成扱い(detailは空でも意味を持たせない)
  if (!generatedAt) {
    return {
      cardSummary: null,
      detail: null,
      model: null,
      generatedAt: null,
      updatedAt: page.last_edited_time,
      tags: tagsOf_(props),
      memo: richTextOf_(props, PROP_MEMO),
      rawContextCount: numberOf_(props, PROP_RAW_COUNT),
    };
  }

  return {
    cardSummary: richTextOf_(props, PROP_SUMMARY) || null,
    detail: {
      decisions: splitLines_(richTextOf_(props, PROP_DECISIONS)),
      todos: parseTodos_(richTextOf_(props, PROP_TODOS)),
      topics: splitLines_(richTextOf_(props, PROP_TOPICS)),
      agenda: parseAgenda_(richTextOf_(props, PROP_AGENDA)),
    },
    model: richTextOf_(props, PROP_MODEL) || null,
    generatedAt: generatedAt,
    updatedAt: page.last_edited_time,
    tags: tagsOf_(props),
    memo: richTextOf_(props, PROP_MEMO),
    rawContextCount: numberOf_(props, PROP_RAW_COUNT),
  };
}

/** number プロパティの値を取得する。無ければ0 */
function numberOf_(properties, name) {
  var prop = properties[name];
  return prop && typeof prop.number === 'number' ? prop.number : 0;
}

/**
 * ToDo行をパースする。"[x] 内容" / "[ ] 内容" / "内容"(マーカー無し=未完了) に対応。
 * @returns {{text: string, done: boolean}[]}
 */
function parseTodos_(text) {
  return splitLines_(text).map(function (line) {
    var m = line.match(/^\[([ xX])\]\s?(.*)$/);
    if (m) {
      return { text: m[2], done: m[1].toLowerCase() === 'x' };
    }
    return { text: line, done: false };
  });
}

/** ToDo配列を "[x] 内容" 形式の改行区切り文字列にする */
function serializeTodos_(todos) {
  if (!Array.isArray(todos)) return '';
  return todos.map(function (t) {
    if (typeof t === 'string') return '[ ] ' + t;
    return (t.done ? '[x] ' : '[ ] ') + (t.text || '');
  }).join('\n');
}

/** 議事JSONをパースする。壊れていれば空配列を返す */
function parseAgenda_(text) {
  if (!text) return [];
  try {
    var parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    // 保存時に文字数上限で切られていると閉じ括弧が無く、ここに来る。
    // 黙って空配列を返すと「議事が未登録」に見えて原因が追えないため記録する。
    console.error('議事JSONのパースに失敗しました(長さ ' + text.length + '): ' + e);
    return [];
  }
}

/**
 * 人手による編集内容を保存する。
 * saveSummary_ と違い、要約日時・要約モデル・状態は変更しない
 * (AIが生成した時刻とモデルの記録を、手編集で上書きしないため)。
 */
function saveDetail_(pageId, cardSummary, detail) {
  detail = detail || {};
  var props = {};
  props[PROP_SUMMARY]   = richTextProp_(cardSummary);
  props[PROP_DECISIONS] = richTextProp_(joinLines_(detail.decisions));
  props[PROP_TODOS]     = richTextProp_(serializeTodos_(detail.todos));
  props[PROP_TOPICS]    = richTextProp_(joinLines_(detail.topics));
  props[PROP_AGENDA]    = richTextProp_(detail.agenda ? JSON.stringify(detail.agenda) : '');

  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/**
 * 状態を「再取得」に変更する。実際の音声再取得・再文字起こしはここでは行わず、
 * Mac mini側のバッチ処理がこの状態を見て後続処理を行う想定。
 */
function requestRetranscribe_(pageId) {
  var props = {};
  props[PROP_STATUS] = { select: { name: STATUS_RETRANSCRIBE } };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, status: STATUS_RETRANSCRIBE };
}

/**
 * スクリプトプロパティ "code" からコードと権限の対応を引く。
 * JSON形式: {"コード": "権限", ...}
 * 該当が無い、またはJSON自体が壊れていれば 'err' を返す。
 *   {"xynekgo":"jba","sdajoihn":"kijun","dfkjnga":"xYz"}
 */
function verifyCode_(code) {
  var raw = PropertiesService.getScriptProperties().getProperty('code') || '';
  var input = String(code || '').trim();
  if (!input) return { role: 'err' };

  var map;
  try {
    map = JSON.parse(raw);
  } catch (e) {
    return { role: 'err' };
  }

  var role = map[input];
  if (!role) return { role: 'err' };
  return { role: role, isAdmin: role === ADMIN_ROLE };
}

/** multi_select プロパティから権限名の配列を取り出す */
function permissionsOf_(properties) {
  var prop = properties[PROP_PERMISSION];
  if (!prop || !prop.multi_select) return [];
  return prop.multi_select.map(function (o) { return o.name; });
}

/**
 * 複数ページの権限をまとめて更新する。
 * mode: 'add' 既存に追加 / 'remove' 指定分を除去 / 'replace' 置き換え
 * 6分の実行上限に収めるため、フロント側から適度な件数で分割して呼ぶこと。
 */
function savePermissions_(pageIds, permissions, mode) {
  var ids = Array.isArray(pageIds) ? pageIds : [];
  var names = (Array.isArray(permissions) ? permissions : []).filter(Boolean);
  var op = mode || 'add';
  var updated = 0;
  var errors = [];

  for (var i = 0; i < ids.length; i++) {
    try {
      var next;
      if (op === 'replace') {
        next = names;
      } else {
        var page = notionFetch_('pages/' + ids[i], 'get');
        var current = permissionsOf_(page.properties);
        if (op === 'remove') {
          next = current.filter(function (c) { return names.indexOf(c) === -1; });
        } else {
          next = current.slice();
          names.forEach(function (n) { if (next.indexOf(n) === -1) next.push(n); });
        }
      }

      var props = {};
      props[PROP_PERMISSION] = {
        multi_select: next.map(function (name) { return { name: name }; }),
      };
      notionFetch_('pages/' + ids[i], 'patch', { properties: props });
      updated++;
    } catch (err) {
      errors.push(ids[i] + ': ' + (err.message || err));
    }
  }

  return { updated: updated, failed: errors.length, errors: errors };
}

/** メモ(自由記述)を更新する。要約とは独立した項目なので単独で保存する */
function saveMemo_(pageId, memo) {
  var props = {};
  props[PROP_MEMO] = richTextProp_(memo);
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/**
 * 状態を「削除」に変更する。Notionページ自体は削除しない
 * (アプリの一覧から除外するだけの論理削除)。
 */
function deleteItem_(pageId) {
  var props = {};
  props[PROP_STATUS] = { select: { name: STATUS_DELETED } };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, status: STATUS_DELETED };
}

/** ミーティング名(title プロパティ)を更新する */
function saveTitle_(pageId, title) {
  var props = {};
  props[PROP_TITLE] = {
    title: [{ text: { content: String(title || '').slice(0, 2000) } }],
  };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, title: title };
}

/** multi_select プロパティから選択肢名の配列を取り出す */
function tagsOf_(properties) {
  var prop = properties[PROP_CATEGORY];
  if (!prop || !prop.multi_select) return [];
  return prop.multi_select.map(function (o) { return o.name; });
}

/** タグ(カテゴリー)だけを更新する。要約の生成/再生成とは独立した操作。 */
function saveTags_(pageId, tags) {
  var names = (Array.isArray(tags) ? tags : []).filter(Boolean);
  var props = {};
  props[PROP_CATEGORY] = { multi_select: names.map(function (name) { return { name: name }; }) };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, tags: names };
}

/**
 * 要約を書き戻す。すべて専用プロパティへの patch 一発で完結する
 * (本文ブロックは操作しない)。
 */
function saveSummary_(pageId, cardSummary, detail, model, rawContextCount) {
  detail = detail || {};
  var props = {};
  props[PROP_SUMMARY]   = richTextProp_(cardSummary);
  props[PROP_DECISIONS] = richTextProp_(joinLines_(detail.decisions));
  props[PROP_TODOS]     = richTextProp_(serializeTodos_(detail.todos));
  props[PROP_TOPICS]    = richTextProp_(joinLines_(detail.topics));
  // 議事は「議題ごとに複数の経緯を持つ」入れ子構造のため、行区切りでは表現できない。
  // Notion上での可読性より構造保持を優先し、JSON文字列として保存する。
  props[PROP_AGENDA]    = richTextProp_(detail.agenda ? JSON.stringify(detail.agenda) : '');
  props[PROP_MODEL]     = richTextProp_(model);
  props[PROP_GENERATED] = { date: { start: new Date().toISOString() } };
  props[PROP_STATUS]    = { select: { name: STATUS_SUMMARIZED } };
  if (typeof rawContextCount === 'number' && rawContextCount > 0) {
    props[PROP_RAW_COUNT] = { number: rawContextCount };
  }

  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/** 原文の文字数だけを更新する。要約生成を伴わずコンテキスト数だけ後から確定させる場合に使う */
function updateRawContextCount_(pageId, count) {
  var props = {};
  props[PROP_RAW_COUNT] = { number: Number(count) || 0 };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/** rich_text プロパティ値からプレーンテキストを取り出す(無ければ空文字) */
function richTextOf_(properties, name) {
  var prop = properties[name];
  return prop && prop.rich_text ? plainTextOf_(prop.rich_text) : '';
}

/**
 * rich_text プロパティのペイロードを作る。
 * Notionは1チャンクあたり2000文字までだが、チャンクを複数並べれば
 * 1プロパティに長い文字列を保存できる。議事(agenda)はJSONで
 * 2000文字を超えることがあるため、切り捨てずに分割する。
 */
function richTextProp_(text) {
  var s = String(text || '');
  if (!s) return { rich_text: [] };

  var chunks = [];
  for (var i = 0; i < s.length; i += 2000) {
    chunks.push({ text: { content: s.slice(i, i + 2000) } });
  }
  return { rich_text: chunks };
}

/** 配列を改行区切りの1文字列にする(空配列/未定義は空文字) */
function joinLines_(arr) {
  return Array.isArray(arr) ? arr.filter(Boolean).join('\n') : '';
}

/** 改行区切りの文字列を配列に戻す(空文字は空配列) */
function splitLines_(text) {
  if (!text) return [];
  return text.split('\n').filter(function (line) { return line.length > 0; });
}

// ============ 音声アップロード(ログイン不要) ============

/**
 * Drive API(高度なサービス)を実際に呼び出すことで、Apps Script の静的解析に
 * 「このプロジェクトは Drive を使う」と認識させ、getOAuthToken() が発行する
 * トークンに Drive の書き込みスコープを含めさせる。
 *
 * 事前に「サービス」→「+」→「Drive API」を追加しておくこと
 * (エディタ左側の「サービス」から追加。関数名は Drive でグローバルに使えるようになる)。
 * 戻り値は使わない。呼ぶこと自体に意味がある。
 */
function touchDriveScope_() {
  try {
    Drive.About.get({ fields: 'user' });
  } catch (e) {
    // 高度なサービスが未追加、または一時的なエラーでも致命的ではないので握りつぶす。
    // これが原因で本当にスコープが付かない場合は、初回のトークン発行 403 で気づける。
  }
}

/** デバッグ用。読み取り(About.get)ではなく、実際に書き込み(ファイル作成)を試す。
 *  403 の原因はここ（Files.create）だったので、これが通るかどうかで切り分ける。
 *  成功したら Drive にテスト用の空ファイルができるので、確認後に手動で削除してよい。 */
function testDriveAuth() {
  Logger.log('OAuth token: ' + ScriptApp.getOAuthToken().slice(0, 20) + '...');
  var file = Drive.Files.create({ name: 'drive_auth_test.txt' }, Utilities.newBlob('test'));
  Logger.log('Drive.Files.create 成功: ' + JSON.stringify(file));
}

/**
 * Drive の resumable upload セッションを、このスクリプト自身(管理者)の権限で
 * 発行する。Google の resumable upload は「セッション URL の発行」にだけ認証が要り、
 * 発行された URL への実際のバイト送信(PUT)には認証が要らない仕様になっている。
 * そのため、ブラウザはここで受け取った URL に直接 PUT するだけで済み、
 * Google へのログインが一切不要になる。ファイル本体もこの GAS を経由しない
 * (経由すると doPost 全体で約50MBのペイロード上限に当たるため)。
 */
function initUpload_(body) {
  touchDriveScope_();
  var folderId = PropertiesService.getScriptProperties().getProperty('AUDIO_INBOX_FOLDER_ID');
  if (!folderId) throw new Error('AUDIO_INBOX_FOLDER_ID が未設定です');

  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + token,
        'X-Upload-Content-Type': body.mimeType || 'application/octet-stream',
        'X-Upload-Content-Length': String(body.size || 0),
        // Google はセッション URI への CORS 許可を「セッションを作った最初の
        // POST に Origin ヘッダーがあったか」で決める（バケット/ファイル側の
        // CORS 設定だけでは効かない）。UrlFetchApp はサーバー実行なので
        // Origin を自動では付けない。ブラウザから直接 PUT するために明示する。
        'Origin': 'https://ymatsuda-cmyk.github.io',
      },
      payload: JSON.stringify({ name: body.filename, parents: [folderId] }),
      muteHttpExceptions: true,
    }
  );

  if (res.getResponseCode() >= 300) {
    throw new Error('Drive session failed (' + res.getResponseCode() + '): ' + res.getContentText());
  }

  var headers = res.getAllHeaders();
  var sessionUrl = headers['Location'] || headers['location'];
  if (!sessionUrl) throw new Error('セッション URL が取得できませんでした');

  return { sessionUrl: sessionUrl };
}

/**
 * 打合せ日時などのメタデータをサイドカー JSON として同じフォルダに置く。
 * ファイルサイズが小さいので multipart で一発アップロードする
 * (resumable にする必要がない)。Mac mini 側の drive_inbox.py は
 * このファイルの出現を処理開始の合図として使う。
 */
function writeSidecar_(body) {
  touchDriveScope_();
  var folderId = PropertiesService.getScriptProperties().getProperty('AUDIO_INBOX_FOLDER_ID');
  if (!folderId) throw new Error('AUDIO_INBOX_FOLDER_ID が未設定です');

  var token = ScriptApp.getOAuthToken();
  var boundary = '-------minutesUploader' + Date.now();
  var payload =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({ name: body.name, parents: [folderId] }) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: application/json\r\n\r\n' +
    JSON.stringify(body.meta) + '\r\n' +
    '--' + boundary + '--';

  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      headers: { Authorization: 'Bearer ' + token },
      payload: payload,
      muteHttpExceptions: true,
    }
  );

  if (res.getResponseCode() >= 300) {
    throw new Error('JSON登録に失敗 (' + res.getResponseCode() + '): ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}