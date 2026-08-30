/**
 * 議事録ビューア用 GAS プロキシ
 *
 * 役割:
 *  - Notion API はブラウザから直接叩けない(CORSヘッダーなし)ため、
 *    このスクリプトが仲介する。
 *  - フロントは text/plain で JSON を POST する(プリフライトを発生させないため)。
 *    Content-Type を application/json にすると OPTIONS preflight が飛び、
 *    GAS は OPTIONS に応答できないため必ず失敗する。
 *
 * 事前設定 (スクリプトプロパティ):
 *  - NOTION_TOKEN    Notion Integration のシークレット
 *  - ACCESS_TOKEN    フロントの minutes:config.notionToken と一致させる共有トークン
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
var PROP_CATEGORY  = 'カテゴリー'; // タグ (multi_select、自由入力可)
var PROP_AGENDA    = '議事';       // 議題ごとの経緯 (rich_text、JSON文字列で格納)
var PROP_TITLE     = 'ミーティング名'; // タイトル (title)

// ============ エントリーポイント ============

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    assertToken_(body);

    switch (body.action) {
      case 'fetchTranscript':
        result = fetchTranscript_(body.pageId);
        break;
      case 'fetchSummary':
        result = fetchSummary_(body.pageId);
        break;
      case 'saveSummary':
        result = saveSummary_(body.pageId, body.cardSummary, body.detail, body.model);
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
  };
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
function saveSummary_(pageId, cardSummary, detail, model) {
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

  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/** rich_text プロパティ値からプレーンテキストを取り出す(無ければ空文字) */
function richTextOf_(properties, name) {
  var prop = properties[name];
  return prop && prop.rich_text ? plainTextOf_(prop.rich_text) : '';
}

/** 2000字上限に収めた rich_text プロパティのペイロードを作る */
function richTextProp_(text) {
  return { rich_text: [{ text: { content: String(text || '').slice(0, 2000) } }] };
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