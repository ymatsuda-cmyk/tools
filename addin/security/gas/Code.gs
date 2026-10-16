/**
 * 情報漏洩対策集アドイン用：元記事の取得・整理（中継）スクリプト
 *
 * アドイン（ブラウザ側）は他社サイトを直接読めないため、このスクリプトが代わりに元記事を取得し、
 * 内容を7項目（発生日・概要・発覚の経緯・リスク・原因・対処・導入業者としての対策）に整理して返します。
 *
 * 【設定】プロジェクトの設定 › スクリプト プロパティ
 *   TOKEN              必須。長いランダムな文字列（アドインの設定にも同じ値を入れる）
 *   ANTHROPIC_API_KEY  任意。あると記事の本文を読んで7項目を整理する。無いと記事の紹介文だけ返す
 *   ANTHROPIC_MODEL    任意。既定は claude-sonnet-5-5
 *   ALLOWED_HOSTS      任意。取得を許可するサイト（カンマ区切り）。既定は cybersecurity-jp.com
 *   DAILY_LIMIT        任意。1日に処理できる記事数の上限。既定は 200
 *
 * 【公開】デプロイ › 新しいデプロイ › 種類「ウェブアプリ」
 *   次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 */
var VERSION = '1.0';
var TZ = 'Asia/Tokyo';
var UA = 'Mozilla/5.0 (compatible; leak-cases-relay/1.0)';
var CATS = ['ext', 'chain', 'ransom', 'acct', 'human', 'other'];
var SEC_KEYS = ['sec_date', 'sec_summary', 'sec_discovery', 'sec_risk', 'sec_cause', 'sec_response', 'sec_vendor'];

function props_() { return PropertiesService.getScriptProperties(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** ブラウザで開いたときの動作確認用（トークン不要・設定内容は返さない） */
function doGet() {
  return json_({ ok: true, service: 'leak-cases-relay', version: VERSION });
}

/** アドインからの呼び出し。Content-Type は text/plain（CORSの事前確認を避けるため） */
function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!tokenOk_(req.token)) return json_({ ok: false, error: 'unauthorized' });
    if (req.action === 'ping') {
      return json_({ ok: true, version: VERSION, ai: !!props_().getProperty('ANTHROPIC_API_KEY'), hosts: allowedHosts_() });
    }
    if (req.action === 'article') return json_(handleArticle_(req));
    return json_({ ok: false, error: '未対応の操作です' });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function tokenOk_(t) {
  var want = props_().getProperty('TOKEN');
  if (!want || want.length < 12 || typeof t !== 'string' || t.length !== want.length) return false;
  var d = 0;
  for (var i = 0; i < want.length; i++) d |= want.charCodeAt(i) ^ t.charCodeAt(i);
  return d === 0;
}

function allowedHosts_() {
  return String(props_().getProperty('ALLOWED_HOSTS') || 'cybersecurity-jp.com')
    .split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
}

/** https かつ許可サイトだけ（IPアドレス・内部ホストは不可） */
function checkUrl_(url) {
  var m = String(url).match(/^https:\/\/([^\/:?#]+)(?::\d+)?(\/|$|\?)/i);
  if (!m) throw new Error('https のURLだけ取得できます');
  var host = m[1].toLowerCase();
  if (/^[\d.]+$/.test(host) || host.indexOf('[') >= 0 || host === 'localhost') throw new Error('このURLは取得できません');
  var ok = allowedHosts_().some(function (h) { return host === h || host.slice(-(h.length + 1)) === '.' + h; });
  if (!ok) throw new Error('取得を許可していないサイトです（' + host + '）。ALLOWED_HOSTS に追加してください');
}

function countToday_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var p = props_();
    var key = 'COUNT_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
    var all = p.getProperties();
    Object.keys(all).forEach(function (k) { if (k.indexOf('COUNT_') === 0 && k !== key) p.deleteProperty(k); });
    var n = Number(all[key] || 0) + 1;
    var max = Number(p.getProperty('DAILY_LIMIT') || 200);
    if (n > max) throw new Error('1日の上限（' + max + '件）に達しました');
    p.setProperty(key, String(n));
  } finally {
    lock.releaseLock();
  }
}

function resolve_(base, loc) {
  if (/^https?:\/\//i.test(loc)) return loc;
  var origin = String(base).match(/^(https?:\/\/[^\/]+)/i)[1];
  if (loc.charAt(0) === '/') return origin + loc;
  return String(base).replace(/[^\/]*$/, '') + loc;
}

/** リダイレクトは自分で追い、移動先も許可サイトかどうか確認する */
function fetchPage_(url) {
  for (var i = 0; i < 4; i++) {
    checkUrl_(url);
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': UA } });
    var code = r.getResponseCode();
    if (code >= 300 && code < 400) {
      var h = r.getHeaders() || {};
      var loc = h['Location'] || h['location'];
      if (!loc) throw new Error('移動先が分かりません');
      url = resolve_(url, loc);
      continue;
    }
    if (code !== 200) throw new Error('元記事を取得できませんでした（HTTP ' + code + '）');
    var bytes = r.getContent();
    if (bytes && bytes.length > 3000000) throw new Error('ページが大きすぎます');
    var ct = String((r.getHeaders() || {})['Content-Type'] || (r.getHeaders() || {})['content-type'] || '');
    var cs = (ct.match(/charset=([\w\-]+)/i) || [])[1];
    if (!cs) cs = (r.getContentText('ISO-8859-1').slice(0, 4096).match(/<meta[^>]+charset=["']?([\w\-]+)/i) || [])[1];
    var html;
    try { html = r.getContentText(cs || 'UTF-8'); } catch (e) { html = r.getContentText('UTF-8'); }
    return { url: url, html: html };
  }
  throw new Error('リダイレクトが多すぎます');
}

var TAIL_MARKERS = ['関連コラム', 'SNSでもご購読', 'この記事を書いた人', '関連記事', 'おすすめ記事'];

function toText_(h) {
  var t = String(h).replace(/<br\s*\/?>|<\/(p|div|li|h[1-6]|tr|td)>/gi, '\n').replace(/<[^>]+>/g, '');
  t = unescape_(t);
  return t.split('\n').map(function (l) { return l.replace(/[ \t\u3000]+/g, ' ').trim(); }).filter(Boolean).join('\n');
}

/** article → main → ページ全体の順に、十分な長さ（200字以上）のものを使う。関連記事・広告の末尾は切る */
function htmlToText_(html, limit) {
  var t = String(html).replace(/<(script|style|noscript|nav|header|footer|aside|form|svg)[^>]*>[\s\S]*?<\/\1>/gi, '');
  t = t.replace(/<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/gi, function (m, inner) { return '\n■ ' + inner.replace(/<[^>]+>/g, '') + '\n'; });
  var text = '';
  var rxs = [/<article[^>]*>([\s\S]*?)<\/article>/i, /<main[^>]*>([\s\S]*?)<\/main>/i];
  for (var i = 0; i < rxs.length && !text; i++) {
    var m = t.match(rxs[i]);
    if (m) { var x = toText_(m[1]); if (x.length >= 200) text = x; }
  }
  if (!text) text = toText_(t);
  var cuts = TAIL_MARKERS.map(function (k) { return text.indexOf(k, 200); }).filter(function (n) { return n >= 0; });
  if (cuts.length) text = text.slice(0, Math.min.apply(null, cuts));
  return text.slice(0, limit || 14000);
}

/** 紹介文（desc）で始まる記事冒頭の段落を、文の途中で切れない形で返す（サイトの紹介文は途中で切れていることが多いため） */
function leadParagraph_(text, desc, maxlen) {
  maxlen = maxlen || 300;
  if (!desc) return '';
  var key = desc.slice(0, 30), lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf(key) >= 0) {
      if (lines[i].length <= maxlen) return lines[i];
      var cut = lines[i].slice(0, maxlen), k = cut.lastIndexOf('。');
      return k >= 20 ? cut.slice(0, k + 1) : cut;
    }
  }
  return '';
}

function unescape_(s) {
  return String(s).replace(/&(#x?[0-9a-f]+|amp|lt|gt|quot|apos|nbsp);/gi, function (m, g) {
    var k = g.toLowerCase();
    if (k === 'amp') return '&'; if (k === 'lt') return '<'; if (k === 'gt') return '>';
    if (k === 'quot') return '"'; if (k === 'apos') return "'"; if (k === 'nbsp') return ' ';
    var n = k.charAt(1) === 'x' ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10);
    return isFinite(n) ? String.fromCharCode(n) : m;
  });
}

function metaDescription_(html) {
  var rx = [/<meta[^>]+property=["']og:description["'][^>]+content=["']([^"']+)/i,
            /<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:description/i,
            /<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)/i];
  for (var i = 0; i < rx.length; i++) {
    var m = html.match(rx[i]);
    if (m) return unescape_(m[1]).replace(/\s+/g, ' ').trim();
  }
  return '';
}

var SYSTEM_PROMPT = 'あなたは情報セキュリティの調査担当です。与えられた記事本文だけを根拠に、情報漏洩事例を整理してください。' +
  '記事に書かれていないことは推測せず、「記事に記載なし」と書きます。' +
  '本文の文章をそのまま写さず、自分の言葉で簡潔にまとめます（引用する場合は短い語句だけ）。' +
  '出力はJSONオブジェクトのみ（説明文やコードブロックは付けない）。';

function userPrompt_(url, title, body) {
  return '次の記事を整理し、下記のキーを持つJSONで答えてください。\n' +
    '- co: 企業・団体名\n- svc: 対象のサービス・システム名（不明なら空文字）\n- title: 事例名（40字以内）\n- sub: 運営会社や委託関係の補足（不明なら空文字）\n' +
    '- cat: ext(外部からの不正アクセス)/chain(委託先・クラウド経由)/ransom(ランサムウェア・マルウェア)/acct(メール・アカウント乗っ取り)/human(誤送信・設定ミス・人的)/other(改ざん・その他) のいずれか\n' +
    '- n: 漏洩した件数（数値。不明なら0）\n- nt: 件数の表記（例：約660万件。不明なら空文字）\n' +
    '- base: 発生日（不明なら検知日、それも不明なら公表日）。YYYY-MM-DD\n- baseType: base が 発生/検知/公表 のどれか\n' +
    '- facts: [["検知","2026年9月25日"],["漏洩","約660万件"]] の形式（最大4組）\n' +
    '- sec_date: 時系列（発生・検知・公表など）。各行を「・」で始める\n- sec_summary: 概要（2〜4行。各行を「・」で始める）\n' +
    '- sec_discovery: 発覚の経緯\n- sec_risk: リスク・影響範囲（漏洩した情報、想定される被害）\n- sec_cause: 原因（判明している原因と、被害を広げた要因）\n' +
    '- sec_response: 対処（暫定対策と恒久対策に分けて書く）\n' +
    '- sec_vendor: システムを導入・保守する業者として取るべき対策を3〜5点（記事の要約ではなく、この事例から導く提案として書く）\n\n' +
    '記事のURL: ' + url + '\n記事のタイトル: ' + title + '\n記事の本文:\n' + body + '\n';
}

function callClaude_(key, prompt) {
  var model = props_().getProperty('ANTHROPIC_MODEL') || 'claude-sonnet-5-5';
  var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: model, max_tokens: 3000, system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }] })
  });
  if (r.getResponseCode() !== 200) throw new Error('AIの呼び出しに失敗しました（HTTP ' + r.getResponseCode() + '）');
  var data = JSON.parse(r.getContentText());
  return (data.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
}

function str_(v, n) { return typeof v === 'string' ? v.trim().slice(0, n || 3000) : ''; }

/** AIの出力を検証して、アドインの事例の形に整える（不正な値は捨てる） */
function normalize_(d, url) {
  if (!str_(d.title, 80) && !SEC_KEYS.some(function (k) { return str_(d[k]); })) throw new Error('整理結果が空でした');
  var o = { co: str_(d.co, 60), svc: str_(d.svc, 60), title: str_(d.title, 80), sub: str_(d.sub, 120),
            nt: str_(d.nt, 40), n: 0, method: 'AI', sources: '元記事|' + url };
  if (CATS.indexOf(d.cat) >= 0) o.cat = d.cat;
  var n = parseInt(d.n, 10); if (isFinite(n) && n > 0) o.n = n;
  if (/^\d{4}-\d{2}-\d{2}$/.test(str_(d.base, 10))) {
    o.base = d.base;
    if (['発生', '検知', '公表'].indexOf(d.baseType) >= 0) o.baseType = d.baseType;
  }
  var facts = [];
  (Array.isArray(d.facts) ? d.facts : []).slice(0, 4).forEach(function (f) {
    if (Array.isArray(f) && f.length === 2 && str_(f[0], 20) && str_(f[1], 60)) {
      facts.push(str_(f[0], 20).replace(/\|/g, '/').replace(/=/g, '＝') + '=' + str_(f[1], 60).replace(/\|/g, '/'));
    }
  });
  o.facts = facts.join('|');
  SEC_KEYS.forEach(function (k) { o[k] = str_(d[k]); });
  return o;
}

function parseJson_(text) {
  var t = String(text).trim().replace(/^```(?:json)?\s*|\s*```$/g, '');
  var m = t.match(/\{[\s\S]*\}/);
  return JSON.parse(m ? m[0] : t);
}

/** AIが使えないときに、サイトの記載をそのまま見せるための本文（冒頭の段落から。上限あり） */
function displayBody_(text, lead) {
  var t = text;
  if (lead) { var i = t.indexOf(lead.slice(0, 30)); if (i > 0) t = t.slice(i); }
  return t.slice(0, 20000);
}

function handleArticle_(req) {
  var url = String(req.url || '').trim();
  checkUrl_(url);
  countToday_();
  var page = fetchPage_(url);
  var desc = metaDescription_(page.html);
  var body = htmlToText_(page.html);
  var key = props_().getProperty('ANTHROPIC_API_KEY');
  var reason = key ? '' : 'GASに ANTHROPIC_API_KEY が未設定です';
  if (key && req.mode !== 'meta') {
    try {
      if (body.length < 80) throw new Error('記事本文を取り出せませんでした');
      var tm = page.html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
      var title = tm ? unescape_(tm[1]).replace(/\s+/g, ' ').trim() : '';
      var out = callClaude_(key, userPrompt_(page.url, title, body));
      return { ok: true, mode: 'ai', result: normalize_(parseJson_(out), page.url) };
    } catch (err) {
      reason = 'AIの整理に失敗（' + String((err && err.message) || err) + '）';
    }
  }
  var lead = leadParagraph_(body, desc);
  var disp = displayBody_(body, lead);
  var text = disp.length >= 80 ? disp : (lead || desc);
  if (!text) return { ok: false, error: reason || '記事の本文を取り出せませんでした' };
  // AIが使えないとき：7項目に無理に分けず、サイトの記載をそのまま本文として返す
  return { ok: true, mode: 'text', reason: reason, warn: reason, result: { body: text, sources: '元記事|' + page.url, method: '本文' } };
}
