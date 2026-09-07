/**
 * リソースダッシュボード用 GAS
 *
 * 役割:
 *  1. Notionデータベース → index.json を生成して GitHub にコミット
 *  2. ダッシュボードからの画像 / mp4 アップロードを GitHub にコミット
 *  3. ダッシュボードから編集した index.json を直接コミット
 *
 * 事前設定（スクリプトプロパティ）:
 *  GITHUB_TOKEN      … Fine-grained PAT（対象リポジトリの Contents: Read and write）
 *  GITHUB_OWNER      … GitHubのユーザー名 or Organization名
 *  GITHUB_REPO       … リポジトリ名
 *  GITHUB_BRANCH     … 省略時 main
 *  NOTION_TOKEN      … Notion internal integration token（ntn_ で始まる）
 *  NOTION_DATABASE_ID… リンク管理用データベースのID
 *  SHARED_SECRET     … ダッシュボードから呼ぶときの共有パスワード（任意の文字列）
 */

const PROP = PropertiesService.getScriptProperties();

function cfg(key, fallback) {
  const v = PROP.getProperty(key);
  return (v === null || v === '') ? fallback : v;
}

/* ============================================================
   GitHub Contents API
   ============================================================ */

/**
 * リポジトリ内のファイルを作成または更新する。
 * 1回のリクエストでコミットまで完了する（pushは不要）。
 */
function commitFile(path, base64Content, message) {
  const owner  = cfg('GITHUB_OWNER');
  const repo   = cfg('GITHUB_REPO');
  const branch = cfg('GITHUB_BRANCH', 'main');
  const token  = cfg('GITHUB_TOKEN');

  if (!owner || !repo || !token) {
    throw new Error('GITHUB_OWNER / GITHUB_REPO / GITHUB_TOKEN が未設定です');
  }

  const base = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json'
  };

  // 既存ファイルを上書きするには sha が必要
  let sha = null;
  const probe = UrlFetchApp.fetch(base + '?ref=' + encodeURIComponent(branch), {
    headers: headers,
    muteHttpExceptions: true
  });
  if (probe.getResponseCode() === 200) {
    sha = JSON.parse(probe.getContentText()).sha;
  }

  const payload = {
    message: message || ('update ' + path),
    content: base64Content,
    branch: branch
  };
  if (sha) payload.sha = sha;

  const res = UrlFetchApp.fetch(base, {
    method: 'put',
    headers: headers,
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error('GitHub API エラー (' + code + '): ' + res.getContentText().slice(0, 400));
  }
  return JSON.parse(res.getContentText());
}

function commitText(path, text, message) {
  const b64 = Utilities.base64Encode(text, Utilities.Charset.UTF_8);
  return commitFile(path, b64, message);
}

/* ============================================================
   Notion → index.json
   ============================================================ */

const RUN_MAP = {
  'Webアプリ': 'web',
  'ダウンロード': 'download',
  'ローカル実行': 'local'
};
const SIZE_MAP = {
  '標準': 'normal',
  '横2列分': 'wide',
  '横2列分・背高': 'large'
};

function plainText(prop) {
  if (!prop) return '';
  if (prop.type === 'title')      return (prop.title || []).map(t => t.plain_text).join('');
  if (prop.type === 'rich_text')  return (prop.rich_text || []).map(t => t.plain_text).join('');
  if (prop.type === 'url')        return prop.url || '';
  if (prop.type === 'select')     return prop.select ? prop.select.name : '';
  if (prop.type === 'number')     return prop.number === null ? '' : String(prop.number);
  if (prop.type === 'checkbox')   return prop.checkbox ? 'true' : 'false';
  return '';
}

function detectType(url) {
  if (/youtube\.com|youtu\.be/.test(url)) return 'youtube';
  if (/office\.com|sharepoint\.com|docs\.google\.com\/spreadsheets|\.xlsx/.test(url)) return 'excel';
  return 'site';
}

/**
 * Notionデータベースの全ページを取得（ページネーション対応）
 */
function fetchNotionPages() {
  const token = cfg('NOTION_TOKEN');
  const dbId  = cfg('NOTION_DATABASE_ID');
  if (!token || !dbId) throw new Error('NOTION_TOKEN / NOTION_DATABASE_ID が未設定です');

  const url = 'https://api.notion.com/v1/databases/' + dbId + '/query';
  let cursor = null;
  const pages = [];

  do {
    const body = { page_size: 100 };
    if (cursor) body.start_cursor = cursor;

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + token,
        'Notion-Version': '2022-06-28'
      },
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      throw new Error('Notion API エラー (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 400));
    }
    const data = JSON.parse(res.getContentText());
    data.results.forEach(p => pages.push(p));
    cursor = data.has_more ? data.next_cursor : null;
  } while (cursor);

  return pages;
}

/**
 * Notionのページ配列を index.json の resources 形式に変換
 */
function buildResources(pages) {
  const rows = [];

  pages.forEach(page => {
    const p = page.properties;
    const published = p['公開'] && p['公開'].type === 'checkbox' ? p['公開'].checkbox : true;
    if (!published) return;

    const title = plainText(p['名前']) || plainText(p['Name']);
    const url   = plainText(p['URL']);
    if (!title || !url) return;   // 必須項目が空の行はスキップ

    const orderRaw = plainText(p['並び順']);
    rows.push({
      _order: orderRaw === '' ? 9999 : Number(orderRaw),
      id: page.id.replace(/-/g, '').slice(0, 12),
      title: title,
      url: url,
      category: plainText(p['カテゴリ']) || '',
      runType: RUN_MAP[plainText(p['実行方法'])] || 'web',
      cardSize: SIZE_MAP[plainText(p['カードサイズ'])] || 'normal',
      thumb: plainText(p['サムネイル']) || null,
      preview: plainText(p['プレビュー']) || null,
      type: detectType(url)
    });
  });

  rows.sort((a, b) => a._order - b._order);
  rows.forEach(r => delete r._order);
  return rows;
}

/**
 * 既存 index.json の config を引き継ぐ（見た目の設定を消さないため）
 */
function fetchExistingConfig() {
  const owner  = cfg('GITHUB_OWNER');
  const repo   = cfg('GITHUB_REPO');
  const branch = cfg('GITHUB_BRANCH', 'main');
  const url = 'https://raw.githubusercontent.com/' + owner + '/' + repo + '/' + branch + '/index.json';

  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(res.getContentText()).config || null;
  } catch (e) {
    return null;
  }
}

const DEFAULT_CONFIG = {
  title: 'マイリソース',
  cardMin: 180,
  thumbH: 110,
  gap: 14,
  radius: 14,
  fade: 350,
  indexPath: 'index.json'
};

/**
 * メイン処理: Notion を読んで index.json をコミット
 * これを時間主導トリガーに設定すれば定期同期になる
 */
function syncFromNotion() {
  const pages = fetchNotionPages();
  const resources = buildResources(pages);
  const config = Object.assign({}, DEFAULT_CONFIG, fetchExistingConfig() || {});

  const json = JSON.stringify({ config: config, resources: resources }, null, 2);
  commitText('index.json', json, 'Notionから同期 (' + resources.length + '件)');

  Logger.log('同期完了: ' + resources.length + '件');
  return resources.length;
}





/* ============================================================
   APIクォータの自己記録カウンタ（Gemini / OpenAI などの無料枠用）

   これらのサービスには「残り回数を問い合わせるAPI」が公式には
   存在しないため、呼び出し側のスクリプトが自己申告した回数を
   ここで積算し、手動設定した1日の上限との差分を残量として返す。

   外部への通信は一切行わない。すべてこのGASのスクリプトプロパティ
   だけで完結する（Kaggleのような別プロジェクトは不要）。
   ============================================================ */

function quotaDateKeyUtc() {
  var d = new Date();
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function quotaPropKey(service, model, metric) {
  var safe = (service + '_' + model).toUpperCase().replace(/[^A-Z0-9]/g, '_');
  return 'QUOTA_' + safe + '_' + metric + '_' + quotaDateKeyUtc();
}

/**
 * 呼び出し側スクリプトから action:'recordApiUsage' で呼ばれる。
 * リクエスト回数とトークン数を両方積算しておく（どちらを見るかは
 * monitors 側の registration で選べるようにするため）。
 */
function recordApiUsage(service, model, requestCount, tokenCount) {
  var reqKey = quotaPropKey(service, model, 'REQ');
  var tokKey = quotaPropKey(service, model, 'TOK');

  var reqNext = (Number(PROP.getProperty(reqKey)) || 0) + (Number(requestCount) || 1);
  var tokNext = (Number(PROP.getProperty(tokKey)) || 0) + (Number(tokenCount) || 0);

  PROP.setProperty(reqKey, String(reqNext));
  PROP.setProperty(tokKey, String(tokNext));

  return { requests: reqNext, tokens: tokNext };
}

/**
 * monitors の type:"quota" 用。ネットワーク通信なしで即座に返る。
 * monitor.metric が "tokens" ならトークン数、それ以外（省略時含む）は
 * リクエスト回数を残量として表示する。
 */
function getQuotaStatus(monitor) {
  var metric = monitor.metric === 'tokens' ? 'TOK' : 'REQ';
  var unit = monitor.metric === 'tokens' ? 'tok' : '回';

  var used = Number(PROP.getProperty(quotaPropKey(monitor.service, monitor.model, metric))) || 0;
  var limit = Number(monitor.dailyLimit) || 0;
  var remaining = Math.max(0, limit - used);

  return {
    id: monitor.id,
    name: monitor.name || monitor.id,
    state: 'tracking',   // 起動/停止の概念が無いことを表す専用ステート
    remaining: { value: remaining, max: limit, unit: unit },
    note: used >= limit && limit > 0 ? '本日の上限に到達した可能性があります' : '',
    updatedAt: new Date().toISOString()
  };
}

/* ============================================================
   Kaggle コントローラー（別デプロイのGAS）への橋渡し

   Kaggle_controller_single.gs は、このダッシュボード用GASとは
   別の独立したGASプロジェクトとしてデプロイする。
   認証は Authorization ヘッダではなく ?token= クエリパラメータ
   で行う仕様のため、専用の関数で対応する。

   monitors の登録で "type": "kaggle" を指定したものだけ、
   この経路を使う。
   ============================================================ */

function fetchKaggleStatus(monitor) {
  var token = monitorToken(monitor.id);
  var url = monitor.endpoint + '?action=status&token=' + encodeURIComponent(token || '');

  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  var data;
  try { data = JSON.parse(res.getContentText()); }
  catch (e) { return { id: monitor.id, state: 'error', error: 'JSON解析に失敗しました' }; }

  if (!data.success) {
    return { id: monitor.id, state: 'error', error: data.error || '取得に失敗しました' };
  }

  // status の値はコントローラーのバージョンにより異なる:
  //   旧版: Kaggle APIの値をそのまま使う（running / queued / stopped など）
  //   新版: ハートビート自前判定（running / booting / stopping / stopped）
  // 両方に対応できるよう、値そのもので分岐する
  var state;
  if (data.zombie) {
    state = 'error';
  } else if (data.status === 'stopping') {
    state = 'stopping';
  } else if (data.status === 'booting') {
    state = 'starting';
  } else if (data.status === 'running' || data.status === 'queued') {
    state = data.proxyAlive ? 'running' : 'starting';
  } else {
    state = 'stopped';
  }

  var maxH = Math.round((data.weeklyLimitMin || 1800) / 60);
  var remH = Math.round(((data.weeklyRemainMin || 0) / 60) * 10) / 10;

  return {
    id: monitor.id,
    name: data.label || monitor.name || monitor.id,
    state: state,
    remaining: { value: remH, max: maxH, unit: 'h' },
    note: data.zombie ? '応答なし。強制停止が必要な可能性があります' : '',
    updatedAt: new Date().toISOString()
  };
}

function sendKaggleControl(monitor, command) {
  var token = monitorToken(monitor.id);
  var action = command === 'start' ? 'start' : 'stop';

  // Kaggle_controller_single.gs の doPost は、受け取ったJSONを
  // そのまま doGet のパラメータとして扱うため、この形で送れば届く
  var res = UrlFetchApp.fetch(monitor.endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ action: action, token: token }),
    muteHttpExceptions: true
  });

  var data;
  try { data = JSON.parse(res.getContentText()); }
  catch (e) { throw new Error('応答の解析に失敗しました'); }

  if (!data.success) {
    throw new Error(data.message || data.error || '操作に失敗しました');
  }
  return fetchKaggleStatus(monitor);
}

/* ============================================================
   稼働状況モニターの中継

   各サービスは以下の共通仕様のAPIを用意する（docs/service-api.md 参照）:
     GET  {endpoint}?action=status   → 稼働状況を返す
     POST {endpoint}  {action:'start'|'stop'}  → 起動/停止を実行

   ブラウザから直接叩くとCORSとトークン露出の問題があるため、
   GASを経由して呼び出す。
   サービスごとのトークンはスクリプトプロパティに
   MONITOR_TOKEN_{監視ID大文字} の形式で登録する。
   例: 監視IDが gpu-server なら MONITOR_TOKEN_GPU_SERVER
   ============================================================ */

function monitorToken(id) {
  const key = 'MONITOR_TOKEN_' + String(id).toUpperCase().replace(/[^A-Z0-9]/g, '_');
  return PROP.getProperty(key);
}

function monitorHeaders(monitor) {
  const headers = { Accept: 'application/json' };
  const token = monitorToken(monitor.id);
  if (token) headers.Authorization = 'Bearer ' + token;
  return headers;
}

/**
 * 1件分の稼働状況を取得する。
 * 失敗しても例外を投げず、error を含むオブジェクトを返す。
 */
function fetchMonitorStatus(monitor) {
  try {
    if (!monitor.endpoint) throw new Error('endpoint が未設定です');

    if (monitor.type === 'kaggle') {
      return fetchKaggleStatus(monitor);
    }
    if (monitor.type === 'quota') {
      return getQuotaStatus(monitor);
    }

    const sep = monitor.endpoint.indexOf('?') >= 0 ? '&' : '?';
    const url = monitor.endpoint + sep + 'action=status';

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: monitorHeaders(monitor),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });

    const code = res.getResponseCode();
    if (code !== 200) {
      return { id: monitor.id, state: 'error', error: 'HTTP ' + code };
    }

    const data = JSON.parse(res.getContentText());
    return {
      id: monitor.id,
      name: data.name || monitor.name || monitor.id,
      state: data.state || 'unknown',
      remaining: data.remaining || null,
      note: data.note || '',
      updatedAt: data.updatedAt || new Date().toISOString()
    };

  } catch (err) {
    return { id: monitor.id, state: 'error', error: String(err.message || err) };
  }
}

/**
 * 起動 / 停止コマンドを送る
 */
function sendMonitorControl(monitor, command) {
  if (command !== 'start' && command !== 'stop') {
    throw new Error('command は start か stop のみです');
  }
  if (!monitor.endpoint) throw new Error('endpoint が未設定です');

  if (monitor.type === 'kaggle') {
    return sendKaggleControl(monitor, command);
  }
  if (monitor.type === 'quota') {
    throw new Error('この監視対象には起動/停止の概念がありません');
  }

  const res = UrlFetchApp.fetch(monitor.endpoint, {
    method: 'post',
    headers: monitorHeaders(monitor),
    contentType: 'application/json',
    payload: JSON.stringify({ action: command }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 202) {
    throw new Error('サービスがエラーを返しました (HTTP ' + code + '): ' +
                    res.getContentText().slice(0, 200));
  }

  const data = JSON.parse(res.getContentText());
  if (data.ok === false) {
    throw new Error(data.error || 'サービス側で処理できませんでした');
  }

  return {
    id: monitor.id,
    name: data.name || monitor.name || monitor.id,
    state: data.state || (command === 'start' ? 'starting' : 'stopping'),
    remaining: data.remaining || null,
    updatedAt: new Date().toISOString()
  };
}

function testMonitor() {
  const monitor = { id: 'sample', name: 'サンプル', endpoint: 'https://example.com/api/resource' };
  Logger.log(JSON.stringify(fetchMonitorStatus(monitor), null, 2));
}

/* ============================================================
   Web API（ダッシュボードから呼ぶ）
   ============================================================ */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // 疎通確認用
  return jsonOut({ ok: true, message: 'GAS is running' });
}

/**
 * ダッシュボードからのPOSTを処理
 *
 * リクエスト例:
 *  { secret:'...', action:'upload',    path:'assets/previews/bot.mp4', contentBase64:'...' }
 *  { secret:'...', action:'saveIndex', json:'{...}' }
 *  { secret:'...', action:'sync' }
 */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    const expected = cfg('SHARED_SECRET');
    if (expected && req.secret !== expected) {
      return jsonOut({ ok: false, error: '認証に失敗しました' });
    }

    if (req.action === 'upload') {
      if (!req.path || !req.contentBase64) throw new Error('path と contentBase64 は必須です');
      if (!/^assets\//.test(req.path)) throw new Error('アップロード先は assets/ 配下にしてください');
      commitFile(req.path, req.contentBase64, 'アップロード: ' + req.path);
      return jsonOut({ ok: true, path: req.path });
    }

    if (req.action === 'saveIndex') {
      if (!req.json) throw new Error('json は必須です');
      JSON.parse(req.json);   // 形式チェック
      commitText('index.json', req.json, 'ダッシュボードから更新');
      return jsonOut({ ok: true });
    }

    if (req.action === 'recordApiUsage') {
      if (!req.service || !req.model) throw new Error('service と model は必須です');
      var total = recordApiUsage(req.service, req.model, req.count, req.tokens);
      return jsonOut({ ok: true, total: total });
    }

    if (req.action === 'monitorStatus') {
      const list = req.monitors || [];
      const results = list.map(function (m) { return fetchMonitorStatus(m); });
      return jsonOut({ ok: true, results: results });
    }

    if (req.action === 'monitorControl') {
      if (!req.monitor) throw new Error('monitor は必須です');
      const snapshot = sendMonitorControl(req.monitor, req.command);
      return jsonOut({ ok: true, snapshot: snapshot });
    }

    if (req.action === 'sync') {
      const n = syncFromNotion();
      return jsonOut({ ok: true, count: n });
    }

    return jsonOut({ ok: false, error: '不明なaction: ' + req.action });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err.message || err) });
  }
}

/* ============================================================
   動作確認用
   ============================================================ */

function testGitHubConnection() {
  const owner = cfg('GITHUB_OWNER');
  const repo  = cfg('GITHUB_REPO');
  const res = UrlFetchApp.fetch('https://api.github.com/repos/' + owner + '/' + repo, {
    headers: { Authorization: 'Bearer ' + cfg('GITHUB_TOKEN') },
    muteHttpExceptions: true
  });
  Logger.log('GitHub: ' + res.getResponseCode());
  Logger.log(res.getContentText().slice(0, 300));
}

function testNotionConnection() {
  const pages = fetchNotionPages();
  Logger.log('取得件数: ' + pages.length);
  Logger.log(JSON.stringify(buildResources(pages), null, 2));
}