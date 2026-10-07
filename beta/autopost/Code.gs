/**
 * PostPilot — X自動ポストアプリ（Google Apps Script 版）
 *
 * 構成：
 *   - データ：このスクリプトに紐づくスプレッドシート（accounts / items / drafts / logs）
 *   - 秘密情報：スクリプトプロパティ（APIキー・Xトークン）
 *   - 画面：index.html（GASのWebアプリとして配信。google.script.run で呼び出すのでCORS不要）
 *   - 定期実行：cronCollect（収集→選定→下書き生成）、cronPost（予約投稿）
 */

// ===== 設定 =====
const X_API = 'https://api.twitter.com';
const DEFAULT_MODEL = 'claude-sonnet-5-5';
const TZ = 'Asia/Tokyo';

const SHEETS = {
  accounts: ['id', 'name', 'handle', 'badge', 'color', 'keywords', 'rss', 'persona', 'target', 'tone', 'ng', 'examples',
             'slots', 'approvalMode', 'includeUrl', 'avoidCrossDup', 'active', 'picksPerRun', 'createdAt'],
  items:    ['id', 'accountId', 'title', 'url', 'source', 'publishedAt', 'summary', 'score', 'reason', 'status', 'createdAt'],
  drafts:   ['id', 'accountId', 'itemId', 'text', 'slotAt', 'status', 'reason', 'warn', 'tweetId', 'postedAt', 'error', 'createdAt'],
  logs:     ['at', 'level', 'accountId', 'message']
};
// items.status : new（未採点）/ scored（採点済み・未使用）/ used（下書き化）/ skipped（見送り）
// drafts.status: generating / pending（承認待ち）/ scheduled（予約）/ posted / failed / rejected

// ===== Webアプリ =====
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PostPilot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== スプレッドシート共通処理 =====
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const created = SpreadsheetApp.create('PostPilot データ');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', created.getId());
  return created;
}

function ensureSheets_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
}

function sheet_(name) { return ss_().getSheetByName(name); }

function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift() || [];
  return values.map(function (row, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = row[j]; });
    return o;
  });
}

function append_(name, obj) {
  const head = SHEETS[name];
  sheet_(name).appendRow(head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
  return obj;
}

function update_(name, id, patch) {
  const head = SHEETS[name];
  const rows = readAll_(name);
  const row = rows.find(function (r) { return String(r.id) === String(id); });
  if (!row) throw new Error(name + ' に ID ' + id + ' が見つかりません');
  const merged = Object.assign({}, row, patch);
  sheet_(name).getRange(row._row, 1, 1, head.length)
    .setValues([head.map(function (h) { return merged[h] === undefined ? '' : merged[h]; })]);
  return merged;
}

function delete_(name, id) {
  const row = readAll_(name).find(function (r) { return String(r.id) === String(id); });
  if (row) sheet_(name).deleteRow(row._row);
}

function uid_() { return Utilities.getUuid().slice(0, 8); }
function nowIso_() { return new Date().toISOString(); }

function log_(level, accountId, message) {
  try {
    append_('logs', { at: nowIso_(), level: level, accountId: accountId || '', message: String(message).slice(0, 1000) });
  } catch (e) { console.error(e); }
}

function props_() { return PropertiesService.getScriptProperties(); }

function bool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }

function account_(id) {
  const a = readAll_('accounts').find(function (r) { return String(r.id) === String(id); });
  if (!a) throw new Error('アカウントが見つかりません：' + id);
  return a;
}

// ===== 画面向けAPI =====
function getBoard() {
  ensureSheets_();
  const accounts = readAll_('accounts').map(publicAccount_);
  const since = Date.now() - 36 * 3600 * 1000;
  const drafts = readAll_('drafts').filter(function (d) {
    if (d.status === 'rejected') return false;
    if (d.status === 'posted' || d.status === 'failed') return new Date(d.postedAt || d.createdAt).getTime() > since;
    return true;
  }).map(function (d) {
    return { id: d.id, accountId: d.accountId, itemId: d.itemId, text: d.text, slotAt: d.slotAt, status: d.status,
             reason: d.reason, warn: d.warn, error: d.error, postedAt: d.postedAt, tweetId: d.tweetId };
  });
  const itemsAll = readAll_('items');
  const itemById = {};
  itemsAll.forEach(function (it) { itemById[it.id] = it; });
  drafts.forEach(function (d) {
    const it = itemById[d.itemId];
    d.sourceTitle = it ? it.title : '';
    d.sourceUrl = it ? it.url : '';
  });
  const candidates = itemsAll
    .filter(function (it) { return it.status === 'scored' || it.status === 'new'; })
    .sort(function (a, b) { return (Number(b.score) || 0) - (Number(a.score) || 0); })
    .slice(0, 60)
    .map(function (it) {
      return { id: it.id, accountId: it.accountId, title: it.title, url: it.url, source: it.source,
               score: it.score, reason: it.reason, status: it.status };
    });
  const lastRun = props_().getProperty('LAST_RUN') || '';
  return { accounts: accounts, drafts: drafts, candidates: candidates, lastRun: lastRun,
           autoOn: triggersOn_(), setupDone: setupDone_() };
}

function publicAccount_(a) {
  const p = props_();
  return {
    id: a.id, name: a.name, handle: a.handle, badge: a.badge, color: a.color,
    keywords: a.keywords, rss: a.rss, persona: a.persona, target: a.target, tone: a.tone, ng: a.ng,
    examples: a.examples, slots: a.slots, approvalMode: a.approvalMode || 'manual',
    includeUrl: bool_(a.includeUrl), avoidCrossDup: a.avoidCrossDup === '' ? true : bool_(a.avoidCrossDup),
    active: a.active === '' ? true : bool_(a.active), picksPerRun: Number(a.picksPerRun) || 2,
    xConnected: !!p.getProperty('X_TOKEN_' + a.id)
  };
}

function approveDraft(id) {
  const d = update_('drafts', id, { status: 'scheduled' });
  return d.id;
}

function unscheduleDraft(id) {
  update_('drafts', id, { status: 'pending' });
  return id;
}

function rejectDraft(id) {
  update_('drafts', id, { status: 'rejected' });
  return id;
}

function updateDraft(id, text, slotAt) {
  const patch = { text: text };
  if (slotAt) patch.slotAt = new Date(slotAt).toISOString();
  patch.warn = warnFor_(text);
  update_('drafts', id, patch);
  return id;
}

function retryDraft(id) {
  update_('drafts', id, { status: 'scheduled', error: '', slotAt: nowIso_() });
  return id;
}

function regenerateDraft(id) {
  const d = readAll_('drafts').find(function (r) { return String(r.id) === String(id); });
  if (!d) throw new Error('下書きが見つかりません');
  const a = account_(d.accountId);
  const it = readAll_('items').find(function (r) { return String(r.id) === String(d.itemId); });
  if (!it) throw new Error('ネタ元が見つかりません');
  const text = writePost_(a, it, d.text);
  update_('drafts', id, { text: text, warn: warnFor_(text) });
  return id;
}

/** ネタ候補から手動で下書きを作る */
function draftFromItem(itemId) {
  const it = readAll_('items').find(function (r) { return String(r.id) === String(itemId); });
  if (!it) throw new Error('ネタが見つかりません');
  const a = account_(it.accountId);
  const slot = nextFreeSlots_(a, 1)[0] || new Date(Date.now() + 3600 * 1000);
  createDraft_(a, it, slot, it.reason || '手動で選択');
  return true;
}

function skipItem(itemId) {
  update_('items', itemId, { status: 'skipped' });
  return true;
}

/** 「今すぐネタを探す」 */
function runNow(accountId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('別の処理が実行中です。少し待ってからもう一度押してください。');
  try {
    const targets = readAll_('accounts').filter(function (a) {
      return (accountId === 'all' || String(a.id) === String(accountId)) && publicAccount_(a).active;
    });
    if (!targets.length) throw new Error('対象のアカウントがありません。設定でアカウントを追加してください。');
    const result = targets.map(function (a) { return runPipelineForAccount_(a); });
    props_().setProperty('LAST_RUN', nowIso_());
    return result;
  } finally {
    lock.releaseLock();
  }
}

// ===== 設定 =====
function getSettings() {
  ensureSheets_();
  const p = props_();
  return {
    accounts: readAll_('accounts').map(publicAccount_),
    global: {
      anthropicKeySet: !!p.getProperty('ANTHROPIC_API_KEY'),
      xKeySet: !!p.getProperty('X_CONSUMER_KEY') && !!p.getProperty('X_CONSUMER_SECRET'),
      model: p.getProperty('MODEL') || DEFAULT_MODEL,
      collectHours: Number(p.getProperty('COLLECT_HOURS')) || 6,
      autoOn: triggersOn_(),
      sheetUrl: ss_().getUrl()
    }
  };
}

function saveGlobal(g) {
  const p = props_();
  if (g.anthropicKey) p.setProperty('ANTHROPIC_API_KEY', g.anthropicKey.trim());
  if (g.xConsumerKey) p.setProperty('X_CONSUMER_KEY', g.xConsumerKey.trim());
  if (g.xConsumerSecret) p.setProperty('X_CONSUMER_SECRET', g.xConsumerSecret.trim());
  if (g.model) p.setProperty('MODEL', g.model.trim());
  if (g.collectHours) p.setProperty('COLLECT_HOURS', String(g.collectHours));
  if (typeof g.autoOn === 'boolean') setupTriggers(g.autoOn);
  return getSettings();
}

function saveAccount(a) {
  ensureSheets_();
  const clean = {
    name: String(a.name || '').trim() || '新しいアカウント',
    badge: String(a.badge || '').trim().slice(0, 2) || String(a.name || '新').slice(0, 1),
    color: a.color || '#1D4FBF',
    keywords: String(a.keywords || ''),
    rss: String(a.rss || ''),
    persona: String(a.persona || ''),
    target: String(a.target || ''),
    tone: a.tone || 'です・ます',
    ng: String(a.ng || ''),
    examples: String(a.examples || ''),
    slots: normalizeSlots_(a.slots),
    approvalMode: a.approvalMode === 'auto' ? 'auto' : 'manual',
    includeUrl: !!a.includeUrl,
    avoidCrossDup: a.avoidCrossDup !== false,
    active: a.active !== false,
    picksPerRun: Math.max(1, Math.min(5, Number(a.picksPerRun) || 2))
  };
  if (a.id) {
    update_('accounts', a.id, clean);
    return a.id;
  }
  const id = uid_();
  append_('accounts', Object.assign({ id: id, handle: '', createdAt: nowIso_() }, clean));
  return id;
}

function deleteAccount(id) {
  delete_('accounts', id);
  const p = props_();
  p.deleteProperty('X_TOKEN_' + id);
  p.deleteProperty('X_SECRET_' + id);
  return true;
}

function normalizeSlots_(s) {
  return String(s || '')
    .split(/[,、\s]+/)
    .map(function (t) { return t.trim(); })
    .filter(function (t) { return /^\d{1,2}:\d{2}$/.test(t); })
    .map(function (t) { const p = t.split(':'); return ('0' + p[0]).slice(-2) + ':' + p[1]; })
    .sort()
    .join(',');
}

function setupDone_() {
  const p = props_();
  return !!p.getProperty('ANTHROPIC_API_KEY');
}

// ===== 定期実行 =====
function setupTriggers(enabled) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const f = t.getHandlerFunction();
    if (f === 'cronCollect' || f === 'cronPost') ScriptApp.deleteTrigger(t);
  });
  if (enabled) {
    const hours = Number(props_().getProperty('COLLECT_HOURS')) || 6;
    ScriptApp.newTrigger('cronCollect').timeBased().everyHours([1, 2, 4, 6, 8, 12].indexOf(hours) >= 0 ? hours : 6).create();
    ScriptApp.newTrigger('cronPost').timeBased().everyMinutes(5).create();
  }
  return triggersOn_();
}

function triggersOn_() {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'cronPost'; });
}

function cronCollect() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    readAll_('accounts').filter(function (a) { return publicAccount_(a).active; }).forEach(function (a) {
      try { runPipelineForAccount_(a); } catch (e) { log_('error', a.id, '定期収集に失敗：' + e.message); }
    });
    props_().setProperty('LAST_RUN', nowIso_());
  } finally {
    lock.releaseLock();
  }
}

function cronPost() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const now = Date.now();
    readAll_('drafts').filter(function (d) {
      return d.status === 'scheduled' && d.slotAt && new Date(d.slotAt).getTime() <= now;
    }).forEach(function (d) {
      try {
        const res = postTweet_(d.accountId, d.text);
        update_('drafts', d.id, { status: 'posted', tweetId: res.id, postedAt: nowIso_(), error: '' });
        log_('info', d.accountId, '投稿しました：' + res.id);
      } catch (e) {
        update_('drafts', d.id, { status: 'failed', error: e.message, postedAt: nowIso_() });
        log_('error', d.accountId, '投稿に失敗：' + e.message);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

// ===== パイプライン：収集 → 選定 → 生成 =====
function runPipelineForAccount_(a) {
  const pa = publicAccount_(a);
  const collected = collect_(a);
  const slots = nextFreeSlots_(a, pa.picksPerRun);
  if (!slots.length) {
    log_('info', a.id, '空き枠がないため下書きは作りませんでした（収集 ' + collected + '件）');
    return { accountId: a.id, collected: collected, drafted: 0 };
  }
  const picks = selectItems_(a, slots.length);
  let drafted = 0;
  picks.forEach(function (p, i) {
    try {
      createDraft_(a, p.item, slots[i], p.reason);
      drafted++;
    } catch (e) {
      log_('error', a.id, '下書き生成に失敗：' + e.message);
    }
  });
  log_('info', a.id, '収集 ' + collected + '件・下書き ' + drafted + '件');
  return { accountId: a.id, collected: collected, drafted: drafted };
}

/** キーワード（Google ニュースRSS）と登録RSSから記事を集める */
function collect_(a) {
  const items = readAll_('items');
  const pa = publicAccount_(a);
  const mine = {};
  const usedElsewhere = {};
  items.forEach(function (it) {
    if (String(it.accountId) === String(a.id)) mine[it.url] = true;
    else if (it.status === 'used') usedElsewhere[it.url] = true;
  });

  const feeds = [];
  String(a.keywords || '').split(/[,、\n]+/).map(function (k) { return k.trim(); }).filter(String).forEach(function (k) {
    feeds.push('https://news.google.com/rss/search?q=' + encodeURIComponent(k + ' when:3d') + '&hl=ja&gl=JP&ceid=JP:ja');
  });
  String(a.rss || '').split(/\n+/).map(function (u) { return u.trim(); }).filter(function (u) { return /^https?:\/\//.test(u); })
    .forEach(function (u) { feeds.push(u); });

  let added = 0;
  const cutoff = Date.now() - 4 * 24 * 3600 * 1000;
  feeds.forEach(function (url) {
    let entries = [];
    try { entries = fetchFeed_(url).slice(0, 10); }
    catch (e) { log_('warn', a.id, 'RSS取得に失敗：' + url + ' / ' + e.message); return; }
    entries.forEach(function (en) {
      if (!en.url || mine[en.url]) return;
      if (pa.avoidCrossDup && usedElsewhere[en.url]) return;
      if (en.publishedAt && new Date(en.publishedAt).getTime() < cutoff) return;
      mine[en.url] = true;
      append_('items', {
        id: uid_(), accountId: a.id, title: en.title, url: en.url, source: en.source,
        publishedAt: en.publishedAt, summary: en.summary, score: '', reason: '', status: 'new', createdAt: nowIso_()
      });
      added++;
    });
  });
  return added;
}

function fetchFeed_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() >= 300) throw new Error('HTTP ' + res.getResponseCode());
  const doc = XmlService.parse(res.getContentText());
  const root = doc.getRootElement();
  const out = [];
  const channel = root.getChild('channel');
  if (channel) {
    channel.getChildren('item').forEach(function (it) {
      const src = it.getChild('source');
      out.push({
        title: text_(it, 'title'),
        url: text_(it, 'link'),
        source: src ? src.getText() : '',
        publishedAt: toIso_(text_(it, 'pubDate')),
        summary: stripHtml_(text_(it, 'description')).slice(0, 400)
      });
    });
    return out;
  }
  const atom = root.getNamespace();
  root.getChildren('entry', atom).forEach(function (en) {
    const link = en.getChild('link', atom);
    out.push({
      title: (en.getChild('title', atom) || { getText: function () { return ''; } }).getText(),
      url: link ? (link.getAttribute('href') ? link.getAttribute('href').getValue() : link.getText()) : '',
      source: '',
      publishedAt: toIso_((en.getChild('updated', atom) || en.getChild('published', atom) || { getText: function () { return ''; } }).getText()),
      summary: stripHtml_((en.getChild('summary', atom) || en.getChild('content', atom) || { getText: function () { return ''; } }).getText()).slice(0, 400)
    });
  });
  return out;
}

function text_(el, name) { const c = el.getChild(name); return c ? c.getText() : ''; }
function toIso_(s) { const d = new Date(s); return isNaN(d.getTime()) ? '' : d.toISOString(); }
function stripHtml_(s) { return String(s || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim(); }

/** AIが未採点のネタを採点し、上位を選ぶ */
function selectItems_(a, n) {
  const items = readAll_('items').filter(function (it) {
    return String(it.accountId) === String(a.id) && (it.status === 'new' || it.status === 'scored');
  }).slice(-40);
  if (!items.length) return [];

  const recent = readAll_('drafts')
    .filter(function (d) { return String(d.accountId) === String(a.id) && d.status !== 'rejected'; })
    .slice(-15).map(function (d) { return '- ' + String(d.text).slice(0, 80); }).join('\n');

  const list = items.map(function (it, i) {
    return i + '. ' + it.title + '（' + (it.source || '出典不明') + '）' + (it.summary ? ' — ' + it.summary.slice(0, 120) : '');
  }).join('\n');

  const system = 'あなたはSNS運用担当の編集者です。与えられた記事候補から、指定アカウントのX投稿ネタとして良いものを採点します。出力はJSONのみ。前置きやコードブロックは不要です。';
  const user = [
    '## アカウント', accountBrief_(a),
    '## 最近の投稿（重複を避けること）', recent || '（なし）',
    '## 記事候補', list,
    '## 指示',
    '各候補を0〜100で採点してください。基準：新しさ、ターゲットへの刺さりやすさ、ペルソナとの相性、最近の投稿と被らないこと、NGルールに触れないこと。',
    'JSON形式：{"scores":[{"idx":0,"score":80,"reason":"40字以内の理由"}]}'
  ].join('\n');

  const json = parseJson_(callClaude_(system, user, 2000));
  const scores = (json && json.scores) || [];
  scores.forEach(function (s) {
    const it = items[s.idx];
    if (it) update_('items', it.id, { score: s.score, reason: String(s.reason || '').slice(0, 120), status: 'scored' });
  });

  return scores
    .filter(function (s) { return items[s.idx] && Number(s.score) >= 60; })
    .sort(function (x, y) { return y.score - x.score; })
    .slice(0, n)
    .map(function (s) { return { item: items[s.idx], reason: s.reason }; });
}

function createDraft_(a, item, slot, reason) {
  const pa = publicAccount_(a);
  const id = uid_();
  append_('drafts', { id: id, accountId: a.id, itemId: item.id, text: '', slotAt: slot.toISOString(), status: 'generating',
                      reason: reason || '', createdAt: nowIso_() });
  update_('items', item.id, { status: 'used' });
  try {
    const text = writePost_(a, item, '');
    update_('drafts', id, { text: text, warn: warnFor_(text), status: pa.approvalMode === 'auto' ? 'scheduled' : 'pending' });
  } catch (e) {
    update_('drafts', id, { status: 'failed', error: '生成に失敗：' + e.message });
    throw e;
  }
  return id;
}

/** 投稿文を作る（文字数チェックつき・最大2回） */
function writePost_(a, item, previous) {
  const pa = publicAccount_(a);
  const urlPart = pa.includeUrl ? 24 : 0; // URLは23文字＋改行扱い
  const maxWeight = 280 - urlPart;
  const system = 'あなたはXの投稿を書くプロのライターです。指定されたペルソナとして、ターゲット読者に届く投稿を日本語で1つ書きます。出力はJSONのみ。';
  let lastErr = '';
  for (let attempt = 0; attempt < 2; attempt++) {
    const user = [
      '## アカウント', accountBrief_(a),
      '## ネタ元', '題名：' + item.title, '媒体：' + (item.source || '不明'), '概要：' + (item.summary || '（なし）'),
      previous ? '## 前回の案（これとは違う切り口で）\n' + previous : '',
      '## ルール',
      '- 全角' + Math.floor(maxWeight / 2) + '文字以内（ハッシュタグ含む）',
      '- 記事にない数字・事実を作らない。断定しすぎない',
      '- ハッシュタグは1〜2個まで',
      '- URLは書かない',
      lastErr ? '- 前回は長すぎました。もっと短く。' : '',
      'JSON形式：{"text":"投稿文"}'
    ].filter(String).join('\n');
    const json = parseJson_(callClaude_(system, user, 800));
    let text = json && json.text ? String(json.text).trim() : '';
    if (!text) { lastErr = 'empty'; continue; }
    if (weightedLength_(text) > maxWeight) { lastErr = 'long'; continue; }
    if (pa.includeUrl && item.url) text += '\n' + item.url;
    return text;
  }
  throw new Error('文字数内の投稿文を作れませんでした');
}

function accountBrief_(a) {
  return [
    'アカウント名：' + a.name,
    'ペルソナ（誰として書くか）：' + (a.persona || '指定なし'),
    'ターゲット（誰に届けるか）：' + (a.target || '指定なし'),
    '口調：' + (a.tone || 'です・ます'),
    'NGルール：' + (a.ng || 'なし'),
    a.examples ? 'お手本の投稿：\n' + a.examples : ''
  ].filter(String).join('\n');
}

function warnFor_(text) {
  const w = [];
  if (/[0-9０-９]+\s*(%|％|倍|万|億|円|人|件)/.test(text)) w.push('数値表現あり');
  return w.join(',');
}

/** Xの文字数カウント（日本語などは2、英数字は1、上限280） */
function weightedLength_(text) {
  let n = 0;
  for (const ch of String(text)) {
    const c = ch.codePointAt(0);
    const light = (c <= 0x10FF) || (c >= 0x2000 && c <= 0x200D) || (c >= 0x2010 && c <= 0x201F) || (c >= 0x2032 && c <= 0x2037);
    n += light ? 1 : 2;
  }
  return n;
}

/** 次の空き投稿枠（最大3日先まで） */
function nextFreeSlots_(a, n) {
  const slots = normalizeSlots_(a.slots).split(',').filter(String);
  if (!slots.length) return [];
  const taken = {};
  readAll_('drafts').forEach(function (d) {
    if (String(d.accountId) === String(a.id) && ['generating', 'pending', 'scheduled', 'posted'].indexOf(d.status) >= 0 && d.slotAt) {
      taken[new Date(d.slotAt).getTime()] = true;
    }
  });
  const out = [];
  const now = new Date();
  const minTime = now.getTime() + 20 * 60 * 1000; // 20分以内の枠は確認が間に合わないので避ける
  for (let day = 0; day < 3 && out.length < n; day++) {
    const base = new Date(now.getFullYear(), now.getMonth(), now.getDate() + day);
    for (let i = 0; i < slots.length && out.length < n; i++) {
      const hm = slots[i].split(':');
      const t = new Date(base.getFullYear(), base.getMonth(), base.getDate(), Number(hm[0]), Number(hm[1]));
      if (t.getTime() > minTime && !taken[t.getTime()]) out.push(t);
    }
  }
  return out;
}

// ===== Claude API =====
function callClaude_(system, user, maxTokens) {
  const p = props_();
  const key = p.getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('Claude APIキーが未設定です（設定画面で登録してください）');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: p.getProperty('MODEL') || DEFAULT_MODEL,
      max_tokens: maxTokens || 1000,
      system: system,
      messages: [{ role: 'user', content: user }]
    }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (code >= 300) throw new Error('Claude API ' + code + '：' + ((body.error && body.error.message) || res.getContentText().slice(0, 200)));
  return (body.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('\n');
}

function parseJson_(s) {
  const t = String(s || '').replace(/```json|```/g, '').trim();
  const start = t.search(/[\[{]/);
  if (start < 0) return null;
  const end = Math.max(t.lastIndexOf('}'), t.lastIndexOf(']'));
  try { return JSON.parse(t.slice(start, end + 1)); } catch (e) { return null; }
}

// ===== X API（OAuth 1.0a・PIN方式で複数アカウントを連携） =====
function xStartAuth(accountId) {
  const p = props_();
  const ck = p.getProperty('X_CONSUMER_KEY'), cs = p.getProperty('X_CONSUMER_SECRET');
  if (!ck || !cs) throw new Error('XのAPI Key / Secret が未設定です');
  const url = X_API + '/oauth/request_token';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: oauthHeader_('POST', url, {}, ck, cs, '', '', { oauth_callback: 'oob' }) },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('連携の開始に失敗：' + res.getContentText().slice(0, 200));
  const q = parseForm_(res.getContentText());
  PropertiesService.getUserProperties().setProperty('XREQ_' + accountId, JSON.stringify(q));
  return X_API + '/oauth/authorize?oauth_token=' + encodeURIComponent(q.oauth_token);
}

function xFinishAuth(accountId, pin) {
  const p = props_();
  const ck = p.getProperty('X_CONSUMER_KEY'), cs = p.getProperty('X_CONSUMER_SECRET');
  const req = JSON.parse(PropertiesService.getUserProperties().getProperty('XREQ_' + accountId) || '{}');
  if (!req.oauth_token) throw new Error('先に「Xと連携」を押してください');
  const url = X_API + '/oauth/access_token';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: oauthHeader_('POST', url, {}, ck, cs, req.oauth_token, req.oauth_token_secret, { oauth_verifier: String(pin).trim() }) },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('PINの確認に失敗：' + res.getContentText().slice(0, 200));
  const q = parseForm_(res.getContentText());
  p.setProperty('X_TOKEN_' + accountId, q.oauth_token);
  p.setProperty('X_SECRET_' + accountId, q.oauth_token_secret);
  PropertiesService.getUserProperties().deleteProperty('XREQ_' + accountId);
  update_('accounts', accountId, { handle: q.screen_name || '' });
  return q.screen_name || '';
}

function xDisconnect(accountId) {
  const p = props_();
  p.deleteProperty('X_TOKEN_' + accountId);
  p.deleteProperty('X_SECRET_' + accountId);
  update_('accounts', accountId, { handle: '' });
  return true;
}

function postTweet_(accountId, text) {
  const p = props_();
  const ck = p.getProperty('X_CONSUMER_KEY'), cs = p.getProperty('X_CONSUMER_SECRET');
  const tk = p.getProperty('X_TOKEN_' + accountId), ts = p.getProperty('X_SECRET_' + accountId);
  if (!tk || !ts) throw new Error('このアカウントはXと未連携です');
  const url = X_API + '/2/tweets';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: oauthHeader_('POST', url, {}, ck, cs, tk, ts, {}) },
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('X API ' + code + '：' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText()).data;
}

function oauthHeader_(method, url, params, ck, cs, token, tokenSecret, extra) {
  const oauth = {
    oauth_consumer_key: ck,
    oauth_nonce: Utilities.getUuid().replace(/-/g, ''),
    oauth_signature_method: 'HMAC-SHA1',
    oauth_timestamp: String(Math.floor(Date.now() / 1000)),
    oauth_version: '1.0'
  };
  if (token) oauth.oauth_token = token;
  Object.keys(extra || {}).forEach(function (k) { oauth[k] = extra[k]; });
  const all = Object.assign({}, params, oauth);
  const paramStr = Object.keys(all).sort().map(function (k) { return pct_(k) + '=' + pct_(all[k]); }).join('&');
  const base = [method.toUpperCase(), pct_(url), pct_(paramStr)].join('&');
  const key = pct_(cs) + '&' + pct_(tokenSecret || '');
  const sig = Utilities.base64Encode(
    Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, base, key, Utilities.Charset.UTF_8));
  oauth.oauth_signature = sig;
  return 'OAuth ' + Object.keys(oauth).sort().map(function (k) { return pct_(k) + '="' + pct_(oauth[k]) + '"'; }).join(', ');
}

function pct_(s) {
  return encodeURIComponent(String(s)).replace(/[!'()*]/g, function (c) { return '%' + c.charCodeAt(0).toString(16).toUpperCase(); });
}

function parseForm_(s) {
  const o = {};
  String(s).split('&').forEach(function (kv) {
    const i = kv.indexOf('=');
    if (i > 0) o[decodeURIComponent(kv.slice(0, i))] = decodeURIComponent(kv.slice(i + 1));
  });
  return o;
}

// ===== 動作確認用（エディタから実行） =====
function testClaude() {
  Logger.log(callClaude_('一言で答えてください。', 'こんにちは', 50));
}

function testWeighted() {
  Logger.log(weightedLength_('日本語140文字テスト abc'));
}
