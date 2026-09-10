/**
 * 動画ナレッジ用 GAS プロキシ
 *
 * 役割:
 *  - Notion API はブラウザから直接叩けない(CORSヘッダーが無い)ため、
 *    このスクリプトが仲介する。
 *  - フロントは text/plain で JSON を POST する(プリフライトを発生させないため)。
 *    Content-Type を application/json にすると OPTIONS preflight が飛び、
 *    GAS は OPTIONS に応答できないため必ず失敗する。
 *  - 議事録アプリ(minutes)の Code.gs と同じ思想。別デプロイにしても、
 *    1つの GAS に両方の action を入れて共用しても動く。
 *
 * 事前設定 (スクリプトプロパティ):
 *  - NOTION_TOKEN   Notion Integration のシークレット
 *  - ACCESS_TOKEN   フロントの videos:config.accessToken と一致させる共有トークン
 *  - code           権限コードと権限の対応 JSON 例: {"dfkjnga":"xYz","abc":"team"}
 *  - VIDEO_DB_ID    動画DBのID(省略時は下の DEFAULT_DB_ID を使う)
 *  - WEB_DB_ID      web記事DBのID(省略時は DEFAULT_WEB_DB_ID。空文字にするとwebを読まない)
 *
 * デプロイ:
 *  - 種類: ウェブアプリ / 実行するユーザー: 自分 / アクセス: 全員
 *    (URLを知っていれば誰でも叩けるため ACCESS_TOKEN のチェックを必ず通す)
 */

var NOTION_VERSION = '2022-06-28';
var DEFAULT_DB_ID = '3630e7a535dc8154ac62d41f7611540f';
var DEFAULT_WEB_DB_ID = '4130e7a535dc83509c9a01cd6ac0a6a7';

// ---- プロパティ名(Notion側のカラム名とここを一致させること) ----
var PROP_TITLE    = '動画タイトル';   // title
var PROP_WEB_TITLE = 'タイトル';       // title  web記事DB側の名前。他のカラム名は動画DBと共通
var PROP_URL      = 'URL';            // url  ※MCPでは userDefined:URL と表示されるが REST では "URL"
var PROP_THUMB    = 'サムネイル';     // url
var PROP_TAGS     = 'タグ';           // multi_select
var PROP_STATUS   = '状態';           // select
var PROP_SUMMARY  = '要約';           // rich_text  カード用サマリ
var PROP_MINDMAP  = 'マインドマップ'; // rich_text  markmap用マークダウン
var PROP_FIELDS   = '分野別要約';     // rich_text  セクション形式(要追加)
var PROP_APPLY    = '応用';           // rich_text  セクション形式(要追加)
var PROP_IDEAS    = '活用アイデア';   // rich_text  セクション形式(要追加)
var PROP_MEMO     = 'メモ';           // rich_text  (要追加)
var PROP_MODEL    = '要約モデル';     // rich_text  (要追加)
var PROP_GENERATED = '要約日時';      // date       (要追加)
var PROP_RAW_COUNT = '原文文字数';    // number     (要追加)
var PROP_CREATED  = '作成日時';       // created_time

// ---- 状態の値 ----
var STATUS_NEW        = '新規';     // 未処理。Mac mini 側バッチが拾う
var STATUS_RUNNING    = '処理中';
var STATUS_DONE       = '完了';     // 文字起こし済み・要約待ち
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
      case 'listVideos':
        result = listVideos_();
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
      case 'deleteVideo':
        result = deleteVideo_(body.pageId);
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
  return PropertiesService.getScriptProperties().getProperty('VIDEO_DB_ID') || DEFAULT_DB_ID;
}

/** web記事DBのID。プロパティを空文字にしておけばwebを読まない運用にできる */
function webDbId_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('WEB_DB_ID');
  return raw === null ? DEFAULT_WEB_DB_ID : String(raw).trim();
}

// ============ Notion 共通ヘルパー ============

function notionFetch_(path, method, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  var options = {
    method: method || 'get',
    headers: {
      Authorization: 'Bearer ' + token,
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
    throw new Error('Notion API ' + code + ': ' + (json.message || res.getContentText()));
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

function urlOf_(properties, name) {
  var prop = properties[name];
  return (prop && prop.url) || '';
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

function titleOf_(properties, name) {
  var prop = properties[name];
  return prop && prop.title ? plainTextOf_(prop.title) : '';
}

/** title 型のプロパティ名を探す。DBごとに名前が違う(動画タイトル / タイトル)ため */
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
 * これにより「2000字で切り捨て」を回避している。
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
    var res = notionFetch_(path, 'get');
    blocks = blocks.concat(res.results);
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return blocks;
}

// ============ アクション実装 ============

/**
 * 一覧を返す。カード表示と検索に必要な項目だけを返し、
 * 長文(分野別要約・応用・活用アイデア・マインドマップ)は
 * 有無のフラグだけにしてペイロードを軽く保つ。
 */
function listVideos_() {
  var items = [];
  eachSourcePage_(function (page, source) {
    var item = toListItem_(page, source);
    if (item) items.push(item);
  });
  return { items: sortByCreatedDesc_(items), fetchedAt: new Date().toISOString() };
}

/**
 * 動画DBとweb記事DBを順に走査する。
 * webDBはあとから足したもので、共有し忘れなどで落ちても動画側は出したいので、
 * webの取得に失敗しても握りつぶす(一覧が空になるより古い動画一覧のほうがまし)。
 */
function eachSourcePage_(fn) {
  queryDbAll_(dbId_(), [{ property: PROP_CREATED, direction: 'descending' }]).forEach(function (page) {
    fn(page, 'video');
  });

  var webId = webDbId_();
  if (!webId) return;
  try {
    // web側は作成日時カラムの有無に依存させたくないので、ページの作成時刻で並べる
    queryDbAll_(webId, [{ timestamp: 'created_time', direction: 'descending' }]).forEach(function (page) {
      fn(page, 'web');
    });
  } catch (err) {
    // 一覧は動画だけで返す
  }
}

/** DBの全ページを取得する(ページネーション対応) */
function queryDbAll_(databaseId, sorts) {
  var pages = [];
  var cursor = null;
  do {
    var payload = { page_size: 100, sorts: sorts };
    if (cursor) payload.start_cursor = cursor;
    var res = notionFetch_('databases/' + databaseId + '/query', 'post', payload);
    pages = pages.concat(res.results);
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);
  return pages;
}

function sortByCreatedDesc_(items) {
  return items.sort(function (a, b) {
    return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
}

/** 一覧の1件。web記事は状態が空欄のものを取り込まない(下書きが混ざるため) */
function toListItem_(page, source) {
  var p = page.properties;
  var status = selectOf_(p, PROP_STATUS);
  if (source === 'web' && !status) return null;

  return {
    key: page.id,
    source: source,
    title: titleAnyOf_(p) || '(タイトル未取得)',
    url: urlOf_(p, PROP_URL),
    thumb: urlOf_(p, PROP_THUMB),
    status: status || STATUS_NEW,
    tags: multiSelectOf_(p, PROP_TAGS),
    createdAt: (p[PROP_CREATED] && p[PROP_CREATED].created_time) || page.created_time,
    editedAt: page.last_edited_time,
    summary: richTextOf_(p, PROP_SUMMARY),
    model: richTextOf_(p, PROP_MODEL) || null,
    generatedAt: dateOf_(p, PROP_GENERATED),
    rawCount: numberOf_(p, PROP_RAW_COUNT),
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
 * 一覧(listVideos)に混ぜると毎回のペイロードが重くなるため別アクションにしている。
 */
function listIdeas_() {
  var items = [];
  eachSourcePage_(function (page, source) {
    var p = page.properties;
    var apply = richTextOf_(p, PROP_APPLY);
    var ideas = richTextOf_(p, PROP_IDEAS);
    if (!apply && !ideas) return;
    var status = selectOf_(p, PROP_STATUS);
    if (source === 'web' && !status) return;
    items.push({
      key: page.id,
      source: source,
      title: titleAnyOf_(p),
      url: urlOf_(p, PROP_URL),
      thumb: urlOf_(p, PROP_THUMB),
      tags: multiSelectOf_(p, PROP_TAGS),
      status: status,
      apply: apply,
      ideas: ideas,
    });
  });

  return { items: items };
}

/**
 * 文字起こし全文を取得する。ページ本文のテキスト系ブロックを上から連結する。
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

  return { text: lines.join('\n'), updatedAt: page.last_edited_time };
}

/** AI生成物とメモをまとめて取得する(詳細を開いたときの1リクエスト) */
function fetchDetail_(pageId) {
  var page = notionFetch_('pages/' + pageId, 'get');
  var p = page.properties;
  return {
    title: titleAnyOf_(p),
    url: urlOf_(p, PROP_URL),
    thumb: urlOf_(p, PROP_THUMB),
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
 * 要約日時・モデル・状態も同時に更新する。
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
 * saveGenerated_ と違い、要約日時・モデル・状態は変更しない
 * (AIが生成した時刻とモデルの記録を手編集で上書きしないため)。
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
 * 現在のタグはサーバ側で読み直すので、フロントの一覧が古くても壊れない。
 *
 * 書き換えたページはフィルタ(from を含む)から外れるため、カーソルで先へ進めるのではなく
 * 「0件になるまで先頭から取り直す」形にしている。カーソルを使うと、削られて詰まった分を
 * 読み飛ばしてしまう。
 */
function mergeTag_(from, to) {
  var fromName = String(from || '').trim();
  var toName = String(to || '').trim();
  if (!fromName || !toName) throw new Error('from と to は必須です');
  if (fromName === toName) throw new Error('from と to が同じです');

  var updated = 0;
  var errors = [];
  var dbIds = [dbId_()];
  var webId = webDbId_();
  if (webId) dbIds.push(webId);

  dbIds.forEach(function (databaseId) {
    // 6分の実行上限に収めるための保険。10巡(最大1000件)で打ち切る
    for (var round = 0; round < 10; round++) {
      var res;
      try {
        res = notionFetch_('databases/' + databaseId + '/query', 'post', {
          page_size: 100,
          filter: { property: PROP_TAGS, multi_select: { contains: fromName } },
        });
      } catch (err) {
        errors.push(databaseId + ': ' + ((err && err.message) || err));
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
          props[PROP_TAGS] = {
            multi_select: next.map(function (name) { return { name: name }; }),
          };
          notionFetch_('pages/' + page.id, 'patch', { properties: props });
          updated++;
        } catch (err) {
          errors.push(page.id + ': ' + ((err && err.message) || err));
        }
      });

      // 1件も進まなかったら、同じページで失敗し続けている。無限ループを避けて抜ける
      if (updated === before) break;
    }
  });

  return { updated: updated, failed: errors.length, errors: errors, from: fromName, to: toName };
}

function saveTitle_(pageId, title) {
  // タイトル欄の名前はDBによって違うので、ページから引く
  var page = notionFetch_('pages/' + pageId, 'get');
  var props = {};
  props[titlePropName_(page.properties)] = { title: [{ text: { content: String(title || '').slice(0, 2000) } }] };
  notionFetch_('pages/' + pageId, 'patch', { properties: props });
  return { saved: true, title: title };
}

/**
 * 状態を変更する。
 *  - 「新規」に戻す = 次回バッチで文字起こしをやり直させる
 *  - 「除外」= 一覧から外す論理削除(Notionページ自体は消さない)
 * Notionのselectは未登録の選択肢名でもAPI側で自動追加されるため、
 * 事前にオプションを作っておく必要はない。
 */
function setStatus_(pageId, status) {
  var allowed = [STATUS_NEW, STATUS_RUNNING, STATUS_DONE, STATUS_SUMMARIZED, STATUS_EXCLUDED];
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

/**
 * ページを削除する。Notion API に完全削除は無いのでアーカイブ(ゴミ箱)になる。
 * 一覧のクエリには出てこなくなり、Notion側では30日間復元できる。
 * 「除外」(setStatus)はページを残す論理削除なので、用途が違う。
 */
function deleteVideo_(pageId) {
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
