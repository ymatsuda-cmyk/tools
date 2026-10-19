/**
 * 共有ボード — GAS サーバー
 *
 * 役割
 *  - push : 端末から届いた操作（線の追加・削除・テキスト更新など）に連番を振ってキャッシュに置く
 *  - poll : 指定の連番以降の操作を返す（2秒前後の間隔で各端末から呼ばれる）
 *  - load : GitHub のスナップショット＋未保存の操作を合成して、ボード全体を返す
 *  - signal : 端末どうしを直接つなぐ（WebRTC）ための接続情報を受け渡す
 *  - 1分ごとのトリガー（flushAll）で、変更のあったボードを GitHub の private リポジトリへ保存
 *
 * スクリプト プロパティ（プロジェクトの設定 → スクリプト プロパティ）
 *  GITHUB_TOKEN  : Fine-grained PAT（対象リポジトリの Contents: Read and write）
 *  GITHUB_REPO   : owner/repo
 *  GITHUB_BRANCH : main（省略時 main）
 *  BOARD_DIR     : boards（省略時 boards）
 *  TEAM_KEY      : 社内で共有する合言葉（画面の設定に入力するもの）
 */

const PROPS = PropertiesService.getScriptProperties();
const CACHE = CacheService.getScriptCache();
const OP_TTL = 21600;          // 操作ログの保持（秒）。キャッシュの上限は6時間
const MAX_OPS_PER_PUSH = 100;
const MAX_POLL = 300;
const PRESENCE_MS = 12000;

/* ---------- 入口 ---------- */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が不正です' });
  }
  return handle_(body);
}

function handle_(q) {
  try {
    const c = conf_();
    if (!c.key || q.key !== c.key) return json_({ ok: false, error: 'unauthorized' });
    switch (q.action) {
      case 'ping':
        return json_({ ok: true, repo: c.repo || '(未設定)', github: !!c.token });
      case 'list':
        return json_(listBoards_(c));
      case 'load':
        return json_(loadBoard_(c, boardId_(q.board), q));
      case 'poll':
        return json_(poll_(boardId_(q.board), Number(q.since) || 0, q));
      case 'push':
        return json_(push_(boardId_(q.board), q.ops, q));
      case 'save':
        return json_(flushBoard_(c, boardId_(q.board)));
      case 'signal':
        return json_(signal_(boardId_(q.board), q));
      default:
        return json_({ ok: false, error: '不明な操作です: ' + q.action });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function conf_() {
  const p = PROPS.getProperties();
  return {
    token: p.GITHUB_TOKEN,
    repo: p.GITHUB_REPO,
    branch: p.GITHUB_BRANCH || 'main',
    dir: p.BOARD_DIR || 'boards',
    key: p.TEAM_KEY
  };
}

function boardId_(b) {
  b = String(b || '').trim();
  if (!b || b.length > 60 || /[\/\\?#%:*"<>|\x00-\x1f]/.test(b) || b.charAt(0) === '.') {
    throw new Error('ボード名に使えない文字が含まれています');
  }
  return b;
}

/* ---------- 操作ログ ---------- */

function opKey_(board, seq) { return 'o:' + board + ':' + seq; }
function getHead_(board) { return Number(PROPS.getProperty('head:' + board) || 0); }

function cleanOp_(op) {
  if (!op || typeof op !== 'object') throw new Error('操作の形式が不正です');
  if (['add', 'del', 'upd', 'clear'].indexOf(op.t) < 0) throw new Error('不明な操作種別です');
  const o = { t: op.t };
  if (op.t !== 'clear') {
    if (!op.el || typeof op.el.id !== 'string' || op.el.id.length > 40) throw new Error('要素IDが不正です');
    o.el = op.t === 'del' ? { id: op.el.id } : op.el;
  }
  return o;
}

function push_(board, ops, q) {
  if (!Array.isArray(ops) || !ops.length) return { ok: true, head: getHead_(board) };
  if (ops.length > MAX_OPS_PER_PUSH) throw new Error('一度に送れる操作は' + MAX_OPS_PER_PUSH + '件までです');
  const by = String(q.user || 'ゲスト').slice(0, 24);
  const cleaned = ops.map(cleanOp_);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let head = getHead_(board);
    const put = {};
    cleaned.forEach(function (op) {
      head++;
      op.seq = head;
      op.by = by;
      const s = JSON.stringify(op);
      if (s.length > 95000) throw new Error('線が長すぎます。分けて描いてください');
      put[opKey_(board, head)] = s;
    });
    CACHE.putAll(put, OP_TTL);            // 先に操作を置き、
    PROPS.setProperty('head:' + board, String(head)); // その後で連番を進める
    return { ok: true, head: head };
  } finally {
    lock.releaseLock();
  }
}

function poll_(board, since, q) {
  const users = touchPresence_(board, q);
  const sig = takeSignals_(board, q.uid);
  const head = getHead_(board);
  if (since >= head) return { ok: true, head: head, ops: [], users: users, sig: sig };
  const to = Math.min(head, since + MAX_POLL);
  const keys = [];
  for (let s = since + 1; s <= to; s++) keys.push(opKey_(board, s));
  const got = CACHE.getAll(keys);
  const ops = [];
  for (let i = 0; i < keys.length; i++) {
    const v = got[keys[i]];
    if (!v) return { ok: true, reset: true, head: head, users: users, sig: sig }; // 欠けがあれば全体を読み直す
    ops.push(JSON.parse(v));
  }
  return { ok: true, head: to, more: to < head, ops: ops, users: users, sig: sig };
}

function applyOp_(m, op) {
  if (op.t === 'clear') { m.clear(); return; }
  const el = op.el;
  if (!el || !el.id) return;
  if (op.t === 'del') m.delete(el.id);
  else if (op.t === 'add') { m.delete(el.id); m.set(el.id, el); }
  else if (op.t === 'upd') { if (m.has(el.id)) m.set(el.id, el); }
}

function buildState_(board, snap, head) {
  const m = new Map();
  (snap.elements || []).forEach(function (el) { m.set(el.id, el); });
  for (let s = (snap.seq || 0) + 1; s <= head; s += 500) {
    const keys = [];
    for (let i = s; i < Math.min(s + 500, head + 1); i++) keys.push(opKey_(board, i));
    const got = CACHE.getAll(keys);
    keys.forEach(function (k) { if (got[k]) applyOp_(m, JSON.parse(got[k])); });
  }
  return Array.from(m.values());
}

/* ---------- 参加者表示 ---------- */

function touchPresence_(board, q) {
  const k = 'pr:' + board;
  let m = {};
  try { m = JSON.parse(CACHE.get(k) || '{}'); } catch (e) { m = {}; }
  const now = Date.now();
  if (q && q.uid) {
    const me = { n: String(q.user || 'ゲスト').slice(0, 24), t: now };
    const x = Number(q.cx), y = Number(q.cy);
    if (q.cx !== undefined && q.cx !== '' && isFinite(x) && isFinite(y)) { me.x = Math.round(x); me.y = Math.round(y); }
    m[String(q.uid).slice(0, 24)] = me;
  }
  Object.keys(m).forEach(function (u) { if (now - m[u].t > PRESENCE_MS) delete m[u]; });
  CACHE.put(k, JSON.stringify(m), 300);
  return Object.keys(m).map(function (u) { return { uid: u, n: m[u].n, x: m[u].x, y: m[u].y }; });
}

/* ---------- 端末どうしの直接接続（シグナリング） ---------- */

function sigKey_(board, to) { return 'sg:' + board + ':' + String(to).slice(0, 24); }

function signal_(board, q) {
  if (!q.uid || !q.to || !q.data) throw new Error('signal の宛先がありません');
  const s = JSON.stringify(q.data);
  if (s.length > 60000) throw new Error('接続情報が大きすぎます');
  const k = sigKey_(board, q.to);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let list = [];
    try { list = JSON.parse(CACHE.get(k) || '[]'); } catch (e) { list = []; }
    list.push({ f: String(q.uid).slice(0, 24), d: q.data, t: Date.now() });
    CACHE.put(k, JSON.stringify(list.slice(-10)), 120);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function takeSignals_(board, uid) {
  if (!uid) return [];
  const k = sigKey_(board, uid);
  if (!CACHE.get(k)) return [];          // 届いていなければロックしない
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let list = [];
    try { list = JSON.parse(CACHE.get(k) || '[]'); } catch (e) { list = []; }
    CACHE.remove(k);
    return list;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- GitHub ---------- */

function ghPath_(c, board) {
  const dir = c.dir.split('/').map(encodeURIComponent).join('/');
  return 'https://api.github.com/repos/' + c.repo + '/contents/' + dir + '/' + encodeURIComponent(board) + '.json';
}

function ghOpt_(c, method, accept, payload) {
  const o = {
    method: method || 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + c.token,
      Accept: accept || 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    }
  };
  if (payload) { o.contentType = 'application/json'; o.payload = JSON.stringify(payload); }
  return o;
}

function needGh_(c) {
  if (!c.token || !c.repo) throw new Error('GITHUB_TOKEN と GITHUB_REPO を設定してください');
}

function ghMeta_(c, board) {
  const r = UrlFetchApp.fetch(ghPath_(c, board) + '?ref=' + encodeURIComponent(c.branch), ghOpt_(c));
  if (r.getResponseCode() === 404) return null;
  if (r.getResponseCode() !== 200) throw new Error('GitHub 読み込み失敗 (' + r.getResponseCode() + ')');
  return JSON.parse(r.getContentText());
}

function ghRaw_(c, board) {
  const r = UrlFetchApp.fetch(ghPath_(c, board) + '?ref=' + encodeURIComponent(c.branch),
    ghOpt_(c, 'get', 'application/vnd.github.raw+json'));
  if (r.getResponseCode() === 404) return null;
  if (r.getResponseCode() !== 200) throw new Error('GitHub 読み込み失敗 (' + r.getResponseCode() + ')');
  return r.getContentText('UTF-8');
}

function ghPut_(c, board, text, sha, message) {
  const body = {
    message: message,
    content: Utilities.base64Encode(text, Utilities.Charset.UTF_8),
    branch: c.branch
  };
  if (sha) body.sha = sha;
  const r = UrlFetchApp.fetch(ghPath_(c, board), ghOpt_(c, 'put', null, body));
  const code = r.getResponseCode();
  if (code === 409 || code === 422) throw new Error('保存が競合しました。少し待ってから再度保存してください');
  if (code !== 200 && code !== 201) throw new Error('GitHub 保存失敗 (' + code + ')');
}

function readSnap_(c, board) {
  const raw = ghRaw_(c, board);
  if (!raw) return { board: board, seq: 0, elements: [] };
  try { return JSON.parse(raw); } catch (e) { throw new Error('ボードのファイルが壊れています: ' + board); }
}

function serialize_(board, seq, elements) {
  // 1要素1行にして、GitHub 上の差分を読みやすくする
  return '{"board":' + JSON.stringify(board) +
    ',"seq":' + seq +
    ',"updated":' + JSON.stringify(new Date().toISOString()) +
    ',"elements":[\n' + elements.map(function (el) { return JSON.stringify(el); }).join(',\n') + '\n]}\n';
}

/* ---------- ボード操作 ---------- */

function listBoards_(c) {
  const names = {};
  if (c.token && c.repo) {
    const dir = c.dir.split('/').map(encodeURIComponent).join('/');
    const r = UrlFetchApp.fetch('https://api.github.com/repos/' + c.repo + '/contents/' + dir +
      '?ref=' + encodeURIComponent(c.branch), ghOpt_(c));
    if (r.getResponseCode() === 200) {
      JSON.parse(r.getContentText()).forEach(function (f) {
        if (f.type === 'file' && /\.json$/.test(f.name)) names[f.name.replace(/\.json$/, '')] = true;
      });
    } else if (r.getResponseCode() !== 404) {
      throw new Error('GitHub 一覧取得失敗 (' + r.getResponseCode() + ')');
    }
  }
  const props = PROPS.getProperties();
  Object.keys(props).forEach(function (k) { if (k.indexOf('head:') === 0) names[k.slice(5)] = true; });
  return { ok: true, boards: Object.keys(names).sort() };
}

function loadBoard_(c, board, q) {
  needGh_(c);
  const snap = readSnap_(c, board);
  const head = Math.max(getHead_(board), snap.seq || 0);
  const elements = buildState_(board, snap, head);
  return { ok: true, board: board, head: head, elements: elements, users: touchPresence_(board, q) };
}

function flushBoard_(c, board) {
  needGh_(c);
  const head = getHead_(board);
  const saved = Number(PROPS.getProperty('saved:' + board) || 0);
  if (head <= saved) return { ok: true, saved: false, seq: saved };

  const flag = 'flushing:' + board;
  if (CACHE.get(flag)) return { ok: true, saved: false, busy: true };
  CACHE.put(flag, '1', 60);
  try {
    const meta = ghMeta_(c, board);
    const snap = meta ? readSnap_(c, board) : { board: board, seq: 0, elements: [] };
    if ((snap.seq || 0) >= head) {
      PROPS.setProperty('saved:' + board, String(snap.seq));
      return { ok: true, saved: false, seq: snap.seq };
    }
    const elements = buildState_(board, snap, head);
    ghPut_(c, board, serialize_(board, head, elements), meta && meta.sha, 'board: ' + board + ' (#' + head + ')');
    PROPS.setProperty('saved:' + board, String(head));
    return { ok: true, saved: true, seq: head };
  } finally {
    CACHE.remove(flag);
  }
}

/** 1分ごとのトリガーから呼ばれる */
function flushAll() {
  const c = conf_();
  if (!c.token || !c.repo) return;
  const props = PROPS.getProperties();
  Object.keys(props).forEach(function (k) {
    if (k.indexOf('head:') !== 0) return;
    const b = k.slice(5);
    if (Number(props[k]) > Number(props['saved:' + b] || 0)) {
      try { flushBoard_(c, b); } catch (e) { console.error(b, e); }
    }
  });
}

/** 初回に一度だけ手動実行：設定チェックとトリガー登録 */
function setup() {
  ['GITHUB_TOKEN', 'GITHUB_REPO', 'TEAM_KEY'].forEach(function (k) {
    if (!PROPS.getProperty(k)) throw new Error('スクリプト プロパティ ' + k + ' が未設定です');
  });
  const c = conf_();
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + c.repo, ghOpt_(c));
  if (r.getResponseCode() !== 200) {
    throw new Error('リポジトリにアクセスできません (' + r.getResponseCode() + ')。トークンの権限と GITHUB_REPO を確認してください');
  }
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'flushAll'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('flushAll').timeBased().everyMinutes(1).create();
  console.log('設定OK：' + c.repo + ' / ' + c.branch + ' / ' + c.dir + '。1分ごとの保存トリガーを登録しました');
}
