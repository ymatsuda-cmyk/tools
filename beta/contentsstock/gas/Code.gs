/**
 * 文書ナレッジ(contentsstock)用 GAS プロキシ
 *
 * 役割:
 *  - Notion API はブラウザから直接叩けない(CORSヘッダーが無い)ため、
 *    このスクリプトが仲介する。
 *  - フロントは text/plain で JSON を POST する(プリフライトを発生させないため)。
 *    Content-Type を application/json にすると OPTIONS preflight が飛び、
 *    GAS は OPTIONS に応答できないため必ず失敗する。
 *  - 動画ナレッジ(clipstock)の Code.gs と同じ思想。扱うDBが1つだけ、
 *    原文の取り直しが無い(ファイルから1度だけ取り込む)ぶん簡単になっている。
 *
 * 事前設定 (スクリプトプロパティ):
 *  - NOTION_TOKEN    Notion Integration のシークレット
 *  - ACCESS_TOKEN    フロントの contents:config.accessToken と一致させる共有トークン
 *  - code            権限コードと権限の対応 JSON 例: {"dfkjnga":"xYz","abc":"team"}
 *  - CONTENTS_DB_ID  コンテンツDBのID(省略時は下の DEFAULT_DB_ID を使う)
 *  - INBOX_FOLDER_ID 動画の取り込み先フォルダのID(省略時は DEFAULT_INBOX_FOLDER_ID)
 *
 * デプロイ:
 *  - 種類: ウェブアプリ / 実行するユーザー: 自分 / アクセス: 全員
 *    (URLを知っていれば誰でも叩けるため ACCESS_TOKEN のチェックを必ず通す)
 *  - 高度なサービスで Drive API を有効にしておくこと(動画アップロードで使う)
 */

var NOTION_VERSION = '2022-06-28';
var DEFAULT_DB_ID = 'd600e7a535dc83caadf381afe7abea03';
// 動画の取り込み先(inbox)。Mac側の監視スクリプトがここを見て文字起こしする
var DEFAULT_INBOX_FOLDER_ID = '10dNn2zgtWCL4FpyYzam_EayKNtD7mGkz';
// 文字起こし後の動画の置き場。Driveリンクを後から探すときに使う
var DEFAULT_STORE_FOLDER_ID = '16SN7XBWosS7WfbpEPUby4gWPDyAAY_px';

// ---- プロパティ名(Notion側のカラム名とここを一致させること) ----
var PROP_TITLE     = 'タイトル';       // title
var PROP_FILE      = 'ファイル名';     // rich_text  取り込み元のファイル名
var PROP_DRIVE     = 'Driveリンク';     // url        取り込んだ元ファイル(動画はここから再生)
var PROP_KIND      = '種別';           // select     mp4 / mov / pdf / docx ...
var PROP_TAGS      = 'タグ';           // multi_select
var PROP_STATUS    = '状態';           // select
var PROP_SUMMARY   = '要約';           // rich_text  カード用サマリ
var PROP_MINDMAP   = 'マインドマップ'; // rich_text  markmap用マークダウン
var PROP_FIELDS    = '分野別要約';     // rich_text  セクション形式
var PROP_APPLY     = '応用';           // rich_text  セクション形式
var PROP_IDEAS     = '活用アイデア';   // rich_text  セクション形式
var PROP_MEMO      = 'メモ';           // rich_text
var PROP_MODEL     = '要約モデル';     // rich_text
var PROP_GENERATED = '要約日時';       // date
var PROP_RAW_COUNT = '原文文字数';     // number
var PROP_PUBLIC    = '公開';           // checkbox   マインドマップ一覧に並べるか
var PROP_CREATED   = '作成日時';       // created_time

// ---- 状態の値 ----
// 文字起こしのやり直しは無いので「新規」「再取得」に相当する状態は持たない
var STATUS_DONE       = '完了';     // 本文の取り込み済み・要約待ち
var STATUS_SUMMARIZED = '要約済み'; // AI生成完了
var STATUS_EXCLUDED   = '除外';     // 論理削除。Notionページ自体は残す

var ADMIN_ROLE = 'xYz';

// ============ エントリーポイント ============

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);

    // コード検証だけは共有トークン不要で受ける(初回はまだトークンが手元に無い)
    if (body.action === 'verifyCode') {
      return jsonOutput_({ ok: true, data: verifyCode_(body.code) });
    }

    assertToken_(body);

    switch (body.action) {
      case 'listContents':
        result = listContents_();
        break;
      case 'listIdeas':
        result = listIdeas_();
        break;
      case 'fetchTranscript':
        result = fetchTranscript_(body.pageId);
        break;
      case 'fetchDetail':
        result = fetchDetail_(body.pageId);
        break;
      case 'saveGenerated':
        result = saveGenerated_(body.pageId, body.detail, body.model, body.rawCount);
        break;
      case 'saveField':
        result = saveField_(body.pageId, body.field, body.value);
        break;
      case 'saveMemo':
        result = saveMemo_(body.pageId, body.memo);
        break;
      case 'saveTags':
        result = saveTags_(body.pageId, body.tags);
        break;
      case 'mergeTag':
        result = mergeTag_(body.from, body.to);
        break;
      case 'saveTitle':
        result = saveTitle_(body.pageId, body.title);
        break;
      case 'setStatus':
        result = setStatus_(body.pageId, body.status);
        break;
      case 'updateRawCount':
        result = updateRawCount_(body.pageId, body.count);
        break;
      case 'setPublic':
        result = setPublic_(body.pageId, body.isPublic);
        break;
      case 'linkDrive':
        result = linkDrive_(body.pageId, body.url);
        break;
      case 'deleteContent':
        result = deleteContent_(body.pageId);
        break;
      case 'requestRebuild':
        result = requestRebuild_(body.reason);
        break;
      case 'takeRebuildRequest':
        result = takeRebuildRequest_();
        break;
      case 'initUpload':
        result = initUpload_(body);
        break;
      case 'putChunk':
        result = putChunk_(body);
        break;
      case 'writeSidecar':
        result = writeSidecar_(body);
        break;
      default:
        throw new Error('unknown action: ' + body.action);
    }
    return jsonOutput_({ ok: true, data: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String((err && err.message) || err) });
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

function dbId_() {
  return PropertiesService.getScriptProperties().getProperty('CONTENTS_DB_ID') || DEFAULT_DB_ID;
}

function notionToken_() {
  return PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN') || '';
}

// ============ Notion 共通ヘルパー ============

function notionFetch_(path, method, payload) {
  var options = {
    method: method || 'get',
    headers: {
      Authorization: 'Bearer ' + notionToken_(),
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  var res = UrlFetchApp.fetch('https://api.notion.com/v1/' + path, options);
  var code = res.getResponseCode();
  var json = JSON.parse(res.getContentText());
  if (code >= 300) {
    var err = new Error('Notion API ' + code + ': ' + (json.message || res.getContentText()));
    err.statusCode = code;
    throw err;
  }
  return json;
}

function plainTextOf_(richTextArray) {
  return (richTextArray || []).map(function (t) { return t.plain_text; }).join('');
}

function richTextOf_(properties, name) {
  var prop = properties[name];
  return prop && prop.rich_text ? plainTextOf_(prop.rich_text) : '';
}

function numberOf_(properties, name) {
  var prop = properties[name];
  return prop && typeof prop.number === 'number' ? prop.number : 0;
}

function selectOf_(properties, name) {
  var prop = properties[name];
  return (prop && prop.select && prop.select.name) || '';
}

function multiSelectOf_(properties, name) {
  var prop = properties[name];
  if (!prop || !prop.multi_select) return [];
  return prop.multi_select.map(function (o) { return o.name; });
}

function checkboxOf_(properties, name) {
  var prop = properties[name];
  return Boolean(prop && prop.checkbox);
}

function urlOf_(properties, name) {
  var prop = properties[name];
  return (prop && prop.url) || '';
}

function titleOf_(properties, name) {
  var prop = properties[name];
  return prop && prop.title ? plainTextOf_(prop.title) : '';
}

/** title 型のプロパティ名を探す。DBごとに名前が違うことがあるため */
function titlePropName_(properties) {
  for (var key in properties) {
    if (properties[key] && properties[key].type === 'title') return key;
  }
  return PROP_TITLE;
}

function titleAnyOf_(properties) {
  return titleOf_(properties, titlePropName_(properties));
}

function dateOf_(properties, name) {
  var prop = properties[name];
  return (prop && prop.date && prop.date.start) || null;
}

/**
 * rich_text プロパティのペイロードを作る。
 * 1オブジェクトあたり2000字が Notion の上限なので、超える分は
 * 複数オブジェクトに分割して詰める(配列は最大100要素 = 実質20万字)。
 */
function richTextProp_(text) {
  var s = String(text == null ? '' : text);
  if (!s) return { rich_text: [] };
  var chunks = [];
  for (var i = 0; i < s.length && chunks.length < 100; i += 2000) {
    chunks.push({ text: { content: s.slice(i, i + 2000) } });
  }
  return { rich_text: chunks };
}

/** ページ本文のブロックを全件取得する(ページネーション対応) */
function fetchAllBlocks_(blockId) {
  var blocks = [];
  var cursor = null;
  do {
    var path = 'blocks/' + blockId + '/children?page_size=100';
    if (cursor) path += '&start_cursor=' + cursor;
    var res = notionFetch_(path, 'get', null);
    blocks = blocks.concat(res.results);
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return blocks;
}

/** DBの全ページを取得する(ページネーション対応) */
function queryDbAll_(sorts) {
  var pages = [];
  var cursor = null;
  do {
    var payload = { page_size: 100, sorts: sorts };
    if (cursor) payload.start_cursor = cursor;
    var res = notionFetch_('databases/' + dbId_() + '/query', 'post', payload);
    pages = pages.concat(res.results);
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return pages;
}

function allPages_() {
  // 作成日時カラムの有無に依存させたくないので、ページの作成時刻で並べる
  return queryDbAll_([{ timestamp: 'created_time', direction: 'descending' }]);
}

// ============ アクション実装 ============

/**
 * 一覧を返す。カード表示と検索に必要な項目だけを返し、
 * 長文(分野別要約・応用・活用アイデア・マインドマップ)は
 * 有無のフラグだけにしてペイロードを軽く保つ。
 */
function listContents_() {
  var items = allPages_().map(toListItem_);
  items.sort(function (a, b) {
    return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
  return { items: items, fetchedAt: new Date().toISOString() };
}

function toListItem_(page) {
  var p = page.properties;
  return {
    key: page.id,
    source: 'doc',
    title: titleAnyOf_(p) || richTextOf_(p, PROP_FILE) || '(タイトル未設定)',
    file: richTextOf_(p, PROP_FILE),
    driveUrl: urlOf_(p, PROP_DRIVE),
    kind: selectOf_(p, PROP_KIND),
    status: selectOf_(p, PROP_STATUS) || STATUS_DONE,
    tags: multiSelectOf_(p, PROP_TAGS),
    createdAt: (p[PROP_CREATED] && p[PROP_CREATED].created_time) || page.created_time,
    editedAt: page.last_edited_time,
    summary: richTextOf_(p, PROP_SUMMARY),
    model: richTextOf_(p, PROP_MODEL) || null,
    generatedAt: dateOf_(p, PROP_GENERATED),
    rawCount: numberOf_(p, PROP_RAW_COUNT),
    isPublic: checkboxOf_(p, PROP_PUBLIC),
    has: {
      mindmap: Boolean(richTextOf_(p, PROP_MINDMAP)),
      fields: Boolean(richTextOf_(p, PROP_FIELDS)),
      apply: Boolean(richTextOf_(p, PROP_APPLY)),
      ideas: Boolean(richTextOf_(p, PROP_IDEAS)),
      memo: Boolean(richTextOf_(p, PROP_MEMO)),
    },
  };
}

/**
 * 応用・活用アイデアだけを全件返す。「アイデア一覧」画面が使う。
 * 一覧(listContents)に混ぜると毎回のペイロードが重くなるため別アクションにしている。
 */
function listIdeas_() {
  var items = [];
  allPages_().forEach(function (page) {
    var p = page.properties;
    var apply = richTextOf_(p, PROP_APPLY);
    var ideas = richTextOf_(p, PROP_IDEAS);
    if (!apply && !ideas) return;
    items.push({
      key: page.id,
      source: 'doc',
      title: titleAnyOf_(p),
      kind: selectOf_(p, PROP_KIND),
      tags: multiSelectOf_(p, PROP_TAGS),
      status: selectOf_(p, PROP_STATUS),
      apply: apply,
      ideas: ideas,
    });
  });
  return { items: items };
}

/** 原文全文を取得する。ページ本文のテキスト系ブロックを上から連結する */
function fetchTranscript_(pageId) {
  var page = notionFetch_('pages/' + pageId, 'get', null);
  var blocks = fetchAllBlocks_(pageId);

  var lines = [];
  blocks.forEach(function (b) {
    var rt = b[b.type] && b[b.type].rich_text;
    if (rt) {
      var text = plainTextOf_(rt);
      if (text) lines.push(text);
    }
  });

  return { text: lines.join('\n'), updatedAt: page.last_edited_time };
}

/** AI生成物とメモをまとめて取得する(詳細を開いたときの1リクエスト) */
function fetchDetail_(pageId) {
  var page = notionFetch_('pages/' + pageId, 'get', null);
  var p = page.properties;
  return {
    title: titleAnyOf_(p),
    file: richTextOf_(p, PROP_FILE),
    driveUrl: urlOf_(p, PROP_DRIVE),
    kind: selectOf_(p, PROP_KIND),
    status: selectOf_(p, PROP_STATUS),
    tags: multiSelectOf_(p, PROP_TAGS),
    summary: richTextOf_(p, PROP_SUMMARY),
    mindmap: richTextOf_(p, PROP_MINDMAP),
    fields: richTextOf_(p, PROP_FIELDS),
    apply: richTextOf_(p, PROP_APPLY),
    ideas: richTextOf_(p, PROP_IDEAS),
    memo: richTextOf_(p, PROP_MEMO),
    model: richTextOf_(p, PROP_MODEL) || null,
    generatedAt: dateOf_(p, PROP_GENERATED),
    rawCount: numberOf_(p, PROP_RAW_COUNT),
    isPublic: checkboxOf_(p, PROP_PUBLIC),
    updatedAt: page.last_edited_time,
  };
}

/** フロントのフィールド名 -> Notionプロパティ名 */
var FIELD_MAP = {
  summary: PROP_SUMMARY,
  mindmap: PROP_MINDMAP,
  fields: PROP_FIELDS,
  apply: PROP_APPLY,
  ideas: PROP_IDEAS,
};

/**
 * AI生成物を書き戻す。detail に含まれるフィールドだけを更新するので、
 * 「分野別だけ作り直す」のような部分生成にも同じ入口で対応できる。
 */
function saveGenerated_(pageId, detail, model, rawCount) {
  detail = detail || {};
  var props = {};
  Object.keys(FIELD_MAP).forEach(function (k) {
    if (typeof detail[k] === 'string') props[FIELD_MAP[k]] = richTextProp_(detail[k]);
  });
  if (Array.isArray(detail.tags)) {
    props[PROP_TAGS] = {
      multi_select: detail.tags.filter(Boolean).map(function (n) { return { name: String(n).slice(0, 100) }; }),
    };
  }
  if (model) props[PROP_MODEL] = richTextProp_(model);
  props[PROP_GENERATED] = { date: { start: new Date().toISOString() } };
  props[PROP_STATUS] = { select: { name: STATUS_SUMMARIZED } };
  if (typeof rawCount === 'number' && rawCount > 0) props[PROP_RAW_COUNT] = { number: rawCount };

  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/**
 * 人手による1フィールドの編集を保存する。
 * saveGenerated_ と違い、要約日時・モデル・状態は変更しない。
 */
function saveField_(pageId, field, value) {
  var prop = FIELD_MAP[field];
  if (!prop) throw new Error('unknown field: ' + field);
  var props = {};
  props[prop] = richTextProp_(value);
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

function saveMemo_(pageId, memo) {
  var props = {};
  props[PROP_MEMO] = richTextProp_(memo);
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

function saveTags_(pageId, tags) {
  var names = (Array.isArray(tags) ? tags : []).filter(Boolean);
  var props = {};
  props[PROP_TAGS] = {
    multi_select: names.map(function (name) { return { name: String(name).slice(0, 100) }; }),
  };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, tags: names };
}

/**
 * タグを統合する。from が付いている全ページを from -> to に書き換える。
 *
 * 書き換えたページはフィルタ(from を含む)から外れるため、カーソルで先へ進めるのではなく
 * 「0件になるまで先頭から取り直す」形にしている。
 */
function mergeTag_(from, to) {
  var fromName = String(from || '').trim();
  var toName = String(to || '').trim();
  if (!fromName || !toName) throw new Error('from と to は必須です');
  if (fromName === toName) throw new Error('from と to が同じです');

  var updated = 0;
  var errors = [];

  // 6分の実行上限に収めるための保険。10巡(最大1000件)で打ち切る
  for (var round = 0; round < 10; round++) {
    var res;
    try {
      res = notionFetch_('databases/' + dbId_() + '/query', 'post', {
        page_size: 100,
        filter: { property: PROP_TAGS, multi_select: { contains: fromName } },
      });
    } catch (err) {
      errors.push(String((err && err.message) || err));
      break;
    }
    if (!res.results.length) break;

    var before = updated;
    res.results.forEach(function (page) {
      try {
        var current = multiSelectOf_(page.properties, PROP_TAGS);
        var next = [];
        current.forEach(function (t) {
          var name = t === fromName ? toName : t;
          if (next.indexOf(name) === -1) next.push(name);
        });
        var props = {};
        props[PROP_TAGS] = { multi_select: next.map(function (name) { return { name: name }; }) };
        notionFetch_('pages/' + page.id, 'patch', { properties: props });
        updated++;
      } catch (err) {
        errors.push(page.id + ': ' + ((err && err.message) || err));
      }
    });

    // 1件も進まなかったら、同じページで失敗し続けている。無限ループを避けて抜ける
    if (updated === before) break;
  }

  return { updated: updated, failed: errors.length, errors: errors, from: fromName, to: toName };
}

function saveTitle_(pageId, title) {
  var page = notionFetch_('pages/' + pageId, 'get', null);
  var props = {};
  props[titlePropName_(page.properties)] = { title: [{ text: { content: String(title || '').slice(0, 2000) } }] };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, title: title };
}

/**
 * 状態を変更する。取り込みのやり直しは無いので「完了」「要約済み」「除外」だけ。
 * Notionのselectは未登録の選択肢名でもAPI側で自動追加される。
 */
function setStatus_(pageId, status) {
  var allowed = [STATUS_DONE, STATUS_SUMMARIZED, STATUS_EXCLUDED];
  if (allowed.indexOf(status) === -1) throw new Error('unknown status: ' + status);
  var props = {};
  props[PROP_STATUS] = { select: { name: status } };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, status: status };
}

function updateRawCount_(pageId, count) {
  var props = {};
  props[PROP_RAW_COUNT] = { number: Number(count) || 0 };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true };
}

/** マインドマップ一覧に出すかどうか。列が無いDBでは Notion が400を返す */
function setPublic_(pageId, isPublic) {
  var props = {};
  props[PROP_PUBLIC] = { checkbox: Boolean(isPublic) };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, isPublic: Boolean(isPublic) };
}

/**
 * Driveリンクを付け直す。
 *
 * アプリから上げた動画はサイドカーのJSONにファイルIDが入るので取り込み時に入るが、
 * Driveに直接置いた分はIDが分からず空のままになる。リンクが無いと再生も
 * タイムスタンプの飛び先も出せないため、ファイル名から探して埋められるようにしている。
 * url を渡されたときは探さずにそれを使う。
 */
function linkDrive_(pageId, url) {
  if (!pageId) throw new Error('pageId は必須です');
  var found = String(url || '').trim();

  if (!found) {
    var page = notionFetch_('pages/' + pageId, 'get', null);
    var filename = richTextOf_(page.properties, PROP_FILE);
    if (!filename) throw new Error('ファイル名が空なので探せません。URLを直接入力してください');
    var id = findDriveFileId_(filename);
    if (!id) throw new Error('Drive に「' + filename + '」が見つかりませんでした');
    found = 'https://drive.google.com/file/d/' + id + '/view';
  }

  var props = {};
  props[PROP_DRIVE] = { url: found };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, driveUrl: found };
}

/** 取り込み先と保管先の両方を同じ名前で探す */
function findDriveFileId_(filename) {
  var folders = [storeFolderId_(), inboxFolderId_()];
  for (var i = 0; i < folders.length; i++) {
    if (!folders[i]) continue;
    try {
      var files = DriveApp.getFolderById(folders[i]).getFilesByName(filename);
      if (files.hasNext()) return files.next().getId();
    } catch (err) {
      // フォルダがID違い・権限無しのときは次を見る
    }
  }
  return null;
}

function storeFolderId_() {
  return PropertiesService.getScriptProperties().getProperty('CONTENTS_FOLDER_ID') || DEFAULT_STORE_FOLDER_ID;
}

/**
 * ページを削除する。Notion API に完全削除は無いのでアーカイブ(ゴミ箱)になる。
 * 「除外」(setStatus)はページを残す論理削除なので、用途が違う。
 */
function deleteContent_(pageId) {
  if (!pageId) throw new Error('pageId は必須です');
  notionFetch_('pages/' + pageId, 'patch', { archived: true });
  return { deleted: true, pageId: pageId };
}

/**
 * スクリプトプロパティ "code" からコードと権限の対応を引く。
 * JSON形式: {"コード": "権限", ...}  該当が無ければ 'err'。
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

// ============ 一覧JSONの作り直し依頼 ============

/**
 * 一覧JSON(index-doc.json / idea-doc.json)は Mac 側の Python が作る。
 * ブラウザから Mac は叩けないので、ここに「作り直してほしい」という印を置き、
 * Mac 側の常駐スクリプトが takeRebuildRequest で拾って実行する。
 */
var REBUILD_PROP = 'REBUILD_REQUEST';

function requestRebuild_(reason) {
  var props = PropertiesService.getScriptProperties();
  var current = parseRebuild_(props.getProperty(REBUILD_PROP));
  var reasons = current.reasons || [];
  var label = String(reason || 'unknown');
  if (reasons.indexOf(label) === -1) reasons.push(label);
  props.setProperty(REBUILD_PROP, JSON.stringify({
    requestedAt: new Date().toISOString(),
    reasons: reasons.slice(0, 20),
  }));
  return { requested: true };
}

/** Mac 側が呼ぶ。印を返して消す（取りこぼしは次回の定期実行で拾う） */
function takeRebuildRequest_() {
  var props = PropertiesService.getScriptProperties();
  var current = parseRebuild_(props.getProperty(REBUILD_PROP));
  if (current.requestedAt) props.deleteProperty(REBUILD_PROP);
  return { requestedAt: current.requestedAt || null, reasons: current.reasons || [] };
}

function parseRebuild_(raw) {
  if (!raw) return {};
  try {
    return JSON.parse(raw) || {};
  } catch (e) {
    return {};
  }
}

// ============ 動画のアップロード(ログイン不要) ============
//
// ブラウザ -> GAS -> Drive の inbox フォルダ、という中継にしている。
// ブラウザから Drive の resumable セッションURLへ直接PUTするとCORSで弾かれるため、
// バイト列もここを通す。1回の doPost は数MBのチャンクなので50MB制限には当たらない。

function inboxFolderId_() {
  return PropertiesService.getScriptProperties().getProperty('INBOX_FOLDER_ID') || DEFAULT_INBOX_FOLDER_ID;
}

/**
 * Drive API(高度なサービス)を実際に呼ぶことで、Apps Script の静的解析に
 * 「このプロジェクトは Drive を使う」と認識させ、getOAuthToken() の
 * トークンに Drive の書き込みスコープを含めさせる。
 * 事前に「サービス」から Drive API を追加しておくこと。
 */
function touchDriveScope_() {
  try {
    Drive.About.get({ fields: 'user' });
  } catch (e) {
    // 未追加でも致命的ではない。本当にスコープが付かない場合は初回の403で気づける
  }
}

/** アップロード先のセッションURLを、このスクリプトの権限で発行する */
function initUpload_(body) {
  touchDriveScope_();
  var folderId = inboxFolderId_();
  if (!folderId) throw new Error('INBOX_FOLDER_ID が未設定です');

  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': body.mimeType || 'application/octet-stream',
        'X-Upload-Content-Length': String(body.size || 0),
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
  if (!sessionUrl) throw new Error('セッションURLが取得できませんでした');
  return { sessionUrl: sessionUrl };
}

/** 音声・動画のバイト列を中継する。308 はまだ途中、200/201 で完了 */
function putChunk_(body) {
  var bytes = Utilities.base64Decode(body.chunk);
  var start = body.offset;
  var end = start + bytes.length - 1;

  var res = UrlFetchApp.fetch(body.sessionUrl, {
    method: 'put',
    contentType: 'application/octet-stream',
    headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + body.total },
    payload: bytes,
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  if (code === 200 || code === 201) return { done: true, file: JSON.parse(res.getContentText()) };
  if (code === 308) return { done: false };
  throw new Error('chunk upload failed (' + code + '): ' + res.getContentText());
}

/**
 * タイトルやメモなどのメタデータを、同じ名前の .json として同じフォルダに置く。
 * Mac 側の取り込みスクリプトは、この出現を処理開始の合図として使う。
 */
function writeSidecar_(body) {
  touchDriveScope_();
  var folderId = inboxFolderId_();
  if (!folderId) throw new Error('INBOX_FOLDER_ID が未設定です');

  var boundary = '-------contentsUploader' + Date.now();
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
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: payload,
      muteHttpExceptions: true,
    }
  );

  if (res.getResponseCode() >= 300) {
    throw new Error('JSON登録に失敗 (' + res.getResponseCode() + '): ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}
