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
 *  GITHUB_BASE_PATH  … リポジトリ配下のディレクトリ（例 beta/dashboard）。省略時はリポジトリ直下
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

/** GITHUB_BASE_PATH を前後のスラッシュ無しに正規化する */
function basePath() {
  return cfg('GITHUB_BASE_PATH', '').replace(/^\/+|\/+$/g, '');
}

/**
 * ダッシュボードから見た相対パスを、リポジトリ内の実際のパスに変換する。
 * 例: GITHUB_BASE_PATH='beta/dashboard' のとき
 *     'assets/thumbs/bot.png' → 'beta/dashboard/assets/thumbs/bot.png'
 */
function repoPath(path) {
  const clean = String(path).replace(/^\/+/, '');
  const base = basePath();
  return base ? base + '/' + clean : clean;
}

/**
 * リポジトリ内のファイルを作成または更新する。
 * path は GITHUB_BASE_PATH からの相対パスで渡す。
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

  const full = repoPath(path);
  const base = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + full;
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
    message: message || ('update ' + full),
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
  const url = 'https://raw.githubusercontent.com/' + owner + '/' + repo + '/' + branch + '/' + repoPath('index.json');

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
   手動カウンタ（Gemini / ChatGPT のチャットUI利用枠など）

   ブラウザのチャットUIの利用回数は外部から取得する手段が無いため、
   ダッシュボードのカードから手動で加算・リセットする。
   カウントはスクリプトプロパティに持つので端末をまたいで共有される。

   monitors の登録例:
     { "id":"gemini-pro", "name":"Gemini Pro", "type":"manual",
       "limit":100, "unit":"回", "cycle":"daily" }
     { "id":"chatgpt-free", "name":"ChatGPT (無料)", "type":"manual",
       "limit":10, "unit":"回", "cycle":"rolling", "windowHours":5 }

   cycle: "daily"（既定） / "monthly" / "rolling"（windowHours 時間の
   ローリングウィンドウ。最初の1回を記録した時点から計測を始める）
   ============================================================ */

function manualKey(id, suffix) {
  var safe = String(id).toUpperCase().replace(/[^A-Z0-9]/g, '_');
  return 'MANUAL_' + safe + '_' + suffix;
}

/** daily / monthly の「今の期間」を表す文字列。期間が変われば値も変わる */
function manualPeriodKey(monitor) {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  if (monitor.cycle === 'monthly') return Utilities.formatDate(now, tz, 'yyyy-MM');
  return Utilities.formatDate(now, tz, 'yyyy-MM-dd');
}

/**
 * 保存済みのカウントを読む。期間が切り替わっていれば 0 として扱う
 * （実際の書き込みは記録時にまとめて行う）。
 */
function readManual(monitor) {
  var used = Number(PROP.getProperty(manualKey(monitor.id, 'USED'))) || 0;
  var mark = PROP.getProperty(manualKey(monitor.id, 'MARK')) || '';
  var start = 0;
  var resetAt = null;

  if (monitor.cycle === 'rolling') {
    var span = (Number(monitor.windowHours) || 5) * 3600 * 1000;
    start = Number(mark) || 0;
    if (!start || Date.now() - start >= span) { used = 0; start = 0; }
    else resetAt = new Date(start + span);
  } else if (mark !== manualPeriodKey(monitor)) {
    used = 0;
  }

  return { used: used, start: start, resetAt: resetAt };
}

function manualNote(monitor, state) {
  if (monitor.cycle === 'rolling') {
    if (!state.resetAt) return (Number(monitor.windowHours) || 5) + '時間の枠。1回目の記録から計測開始';
    var left = Math.max(0, state.resetAt.getTime() - Date.now());
    var h = Math.floor(left / 3600000);
    var m = Math.floor((left % 3600000) / 60000);
    return 'あと ' + h + '時間' + m + '分でリセット';
  }
  return monitor.cycle === 'monthly' ? '毎月1日にリセット' : '毎日0時にリセット';
}

function getManualStatus(monitor) {
  var state = readManual(monitor);
  var limit = Number(monitor.limit) || 0;

  return {
    id: monitor.id,
    name: monitor.name || monitor.id,
    state: 'tracking',
    remaining: { value: Math.max(0, limit - state.used), max: limit, unit: monitor.unit || '回' },
    used: state.used,
    note: manualNote(monitor, state),
    updatedAt: new Date().toISOString()
  };
}

/**
 * カードのボタンから呼ばれる。delta は増減量、reset:true なら 0 に戻す。
 */
function recordManualUsage(monitor, delta, reset) {
  var state = readManual(monitor);
  var used = reset ? 0 : Math.max(0, state.used + (Number(delta) || 0));

  PROP.setProperty(manualKey(monitor.id, 'USED'), String(used));

  if (monitor.cycle === 'rolling') {
    // ウィンドウの起点は「0 から増えた瞬間」。以降は延長しない
    if (used === 0) PROP.deleteProperty(manualKey(monitor.id, 'MARK'));
    else if (!state.start) PROP.setProperty(manualKey(monitor.id, 'MARK'), String(Date.now()));
  } else {
    PROP.setProperty(manualKey(monitor.id, 'MARK'), manualPeriodKey(monitor));
  }

  return getManualStatus(monitor);
}

/* ============================================================
   GitHub Copilot のプレミアムリクエスト残量

   GitHub の課金APIから当月の使用量を取得し、契約プランの上限
   （monitor.monthlyLimit）との差分を残量として返す。

     GET /users/{user}/settings/billing/usage?year=&month=

   このAPIは Fine-grained PAT の「Account permissions → Plan:
   Read-only」が必要。index.json のコミットに使っている GITHUB_TOKEN
   にその権限が無い場合は、監視ID用のトークン
   （MONITOR_TOKEN_COPILOT など）を別に登録すればそちらが優先される。

   注意: このエンドポイントは個人アカウントの課金明細のみを返す。
   ・割り当て枠内の利用は明細に載らないため、超過分が出るまで used は 0
   ・組織管理の Copilot Business/Enterprise シートは対象外（組織側の
     請求管理者権限が必要な別エンドポイントにしか記録が無い）
   いずれの場合も、type:"manual" の手動カウンタ（cycle:"monthly"）を使う。

   monitors の登録例:
     { "id":"copilot", "name":"GitHub Copilot", "type":"copilot",
       "monthlyLimit":300 }
   ============================================================ */

function fetchCopilotStatus(monitor) {
  var token = monitorToken(monitor.id) || cfg('GITHUB_TOKEN');
  var user = monitor.user || cfg('GITHUB_OWNER');
  if (!token) {
    return {
      id: monitor.id, state: 'error',
      error: 'スクリプトプロパティ ' + monitorTokenKey(monitor.id) +
             ' または GITHUB_TOKEN を登録してください'
    };
  }
  if (!user) {
    return {
      id: monitor.id, state: 'error',
      error: 'GITHUB_OWNER を登録するか、monitor に user を指定してください'
    };
  }

  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var url = 'https://api.github.com/users/' + encodeURIComponent(user) +
            '/settings/billing/usage' +
            '?year=' + Utilities.formatDate(now, tz, 'yyyy') +
            '&month=' + Number(Utilities.formatDate(now, tz, 'MM'));

  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code === 401) {
    return {
      id: monitor.id, state: 'error',
      error: 'HTTP 401：トークンが無効か期限切れです（' + monitorTokenKey(monitor.id) + ' / GITHUB_TOKEN）'
    };
  }
  if (code === 403 || code === 404) {
    return {
      id: monitor.id, state: 'error',
      error: 'HTTP ' + code + '：トークンに Plan(Read-only) 権限が必要です'
    };
  }
  if (code !== 200) {
    return { id: monitor.id, state: 'error', error: 'HTTP ' + code };
  }

  var data;
  try { data = JSON.parse(res.getContentText()); }
  catch (e) { return { id: monitor.id, state: 'error', error: 'JSON解析に失敗しました' }; }

  // usageItems は日次の明細。プレミアムリクエスト分だけを当月で合計する
  var used = 0;
  var found = false;
  (data.usageItems || []).forEach(function (it) {
    if (String(it.product || '').toLowerCase().indexOf('copilot') < 0) return;
    if (!/premium/i.test(String(it.sku || ''))) return;
    found = true;
    used += Number(it.quantity) || 0;
  });

  var limit = Number(monitor.monthlyLimit) || 300;

  // 個人アカウントの割り当て枠内の利用は課金明細に載らない
  var note;
  if (!found) note = '課金明細にCopilotの記録なし（枠内利用は反映されません）';
  else if (used >= limit) note = '当月の割り当てを使い切っています';
  else note = '毎月1日にリセット';

  return {
    id: monitor.id,
    name: monitor.name || monitor.id,
    state: 'tracking',
    remaining: { value: Math.max(0, limit - used), max: limit, unit: '回' },
    used: used,
    note: note,
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
   ドル円 かんたん投資（fx-invest のGAS）への橋渡し

   fx-invest/gas/Code.gs をウェブアプリとしてデプロイし、
   その /exec を monitor.endpoint に指定する。合言葉（FX_SECRET）は
   MONITOR_TOKEN_{監視ID} に登録する。

   monitors の登録例:
     { "id":"fx", "name":"ドル円 かんたん投資", "type":"fx",
       "endpoint":"https://script.google.com/macros/s/xxx/exec" }
   ============================================================ */

function fetchFxStatus(monitor) {
  var res = UrlFetchApp.fetch(monitor.endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ action: 'summary', secret: monitorToken(monitor.id) || '' }),
    muteHttpExceptions: true
  });

  var data;
  try { data = JSON.parse(res.getContentText()); }
  catch (e) { return { id: monitor.id, state: 'error', error: 'JSON解析に失敗しました' }; }

  if (!data.ok) {
    return { id: monitor.id, state: 'error', error: data.error || '取得に失敗しました' };
  }

  return {
    id: monitor.id,
    name: monitor.name || monitor.id,
    state: 'tracking',
    fx: data.summary,
    updatedAt: new Date().toISOString()
  };
}

/* ============================================================
   ドル円ポジションの損益（endpoint 不要）

   建値と金額だけを登録し、いまのレートとの差から損益を計算する。
   売買履歴は持たず、この1件だけを見る用途。

   monitors の登録例:
     { "id":"usd-150", "name":"ドル円 150.55 買い", "type":"fxpos",
       "side":"buy", "price":150.55, "amount":1000000 }

   side  : "buy"（既定）/ "sell"
   price : 建値（円）
   amount: 建てた金額（円）。省略時 1000000
   ============================================================ */

/** 数分おきに更新される無料API。1分だけキャッシュして共有する */
function dashUsdJpyRate() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('USDJPY');
  if (hit) return Number(hit);

  var urls = [
    'https://api.fxratesapi.com/latest?base=USD&currencies=JPY',
    'https://open.er-api.com/v6/latest/USD'
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var res = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      var rate = (JSON.parse(res.getContentText()).rates || {}).JPY;
      if (typeof rate === 'number' && rate > 0) {
        cache.put('USDJPY', String(rate), 60);
        return rate;
      }
    } catch (e) { /* 次の取得先を試す */ }
  }
  return null;
}

function getFxPositionStatus(monitor) {
  var price = Number(monitor.price);
  if (!price) {
    return { id: monitor.id, state: 'error', error: 'price（建値）を指定してください' };
  }

  var rate = dashUsdJpyRate();
  if (rate === null) {
    return { id: monitor.id, state: 'error', error: 'レートを取得できませんでした' };
  }

  var side = monitor.side === 'sell' ? 'sell' : 'buy';
  var amount = Number(monitor.amount) || 1000000;
  // 売りは建値で売って現在値で買い戻すので、損益の符号が逆になる
  var ratio = side === 'buy' ? (rate / price - 1) : (1 - rate / price);

  return {
    id: monitor.id,
    name: monitor.name || monitor.id,
    state: 'tracking',
    fxpos: {
      side: side, price: price, amount: amount, rate: rate,
      ratio: ratio, profit: amount * ratio,
      cap: Number(monitor.cap) || 100000,
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')
    },
    updatedAt: new Date().toISOString()
  };
}

/* ============================================================
   残量の手動修正

   取得元の値が実態とずれるときのために、差分を
   ADJUST_{監視ID} に保存して以降の取得値に加算する。
   リセット（週次・月次）で取得元が戻ったときは、
   調整値を解除すること。
   ============================================================ */

function adjustKey(id) {
  return 'ADJUST_' + String(id).toUpperCase().replace(/[^A-Z0-9]/g, '_');
}

function applyAdjust(snap, monitor) {
  if (!snap || !snap.remaining) return snap;

  var delta = Number(PROP.getProperty(adjustKey(monitor.id)));
  if (!delta) return snap;

  var max = Number(snap.remaining.max) || 0;
  var next = (Number(snap.remaining.value) || 0) + delta;
  snap.remaining.value = Math.max(0, max ? Math.min(max, next) : next);

  var label = '手動調整 ' + (delta > 0 ? '+' : '') + (Math.round(delta * 10) / 10);
  snap.note = snap.note ? snap.note + ' ／ ' + label : label;
  return snap;
}

/** 表示したい残量を受け取り、取得値との差を調整値として保存する */
function setMonitorAdjust(monitor, value, reset) {
  var key = adjustKey(monitor.id);

  if (reset) {
    PROP.deleteProperty(key);
    return fetchMonitorStatus(monitor);
  }

  var raw = fetchMonitorRaw(monitor);
  if (!raw || !raw.remaining) throw new Error('残量を持たない監視対象です');

  var target = Number(value);
  if (!isFinite(target)) throw new Error('数値を入力してください');

  PROP.setProperty(key, String(target - (Number(raw.remaining.value) || 0)));
  return applyAdjust(raw, monitor);
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

function monitorTokenKey(id) {
  return 'MONITOR_TOKEN_' + String(id).toUpperCase().replace(/[^A-Z0-9]/g, '_');
}

function monitorToken(id) {
  return PROP.getProperty(monitorTokenKey(id));
}

function monitorHeaders(monitor) {
  const headers = { Accept: 'application/json' };
  const token = monitorToken(monitor.id);
  if (token) headers.Authorization = 'Bearer ' + token;
  return headers;
}

/**
 * 1件分の稼働状況を取得する。手動調整を反映した値を返す。
 */
function fetchMonitorStatus(monitor) {
  return applyAdjust(fetchMonitorRaw(monitor), monitor);
}

/**
 * 取得元が返したままの値。
 * 失敗しても例外を投げず、error を含むオブジェクトを返す。
 */
function fetchMonitorRaw(monitor) {
  try {
    // endpoint を持たない種別を先に処理する
    if (monitor.type === 'quota')   return getQuotaStatus(monitor);
    if (monitor.type === 'manual')  return getManualStatus(monitor);
    if (monitor.type === 'copilot') return fetchCopilotStatus(monitor);
    if (monitor.type === 'fxpos')   return getFxPositionStatus(monitor);

    if (!monitor.endpoint) throw new Error('endpoint が未設定です');

    if (monitor.type === 'kaggle') {
      return fetchKaggleStatus(monitor);
    }
    if (monitor.type === 'fx') {
      return fetchFxStatus(monitor);
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
    if (data.ok === false) {
      return { id: monitor.id, state: 'error', error: data.error || 'サービス側でエラーが発生しました' };
    }
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
  if (monitor.type === 'quota' || monitor.type === 'manual' ||
      monitor.type === 'copilot' || monitor.type === 'fx' || monitor.type === 'fxpos') {
    throw new Error('この監視対象には起動/停止の概念がありません');
  }
  if (!monitor.endpoint) throw new Error('endpoint が未設定です');

  if (monitor.type === 'kaggle') {
    return sendKaggleControl(monitor, command);
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
 * リクエスト例（path は GITHUB_BASE_PATH からの相対パス）:
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
      if (req.path.indexOf('..') !== -1) throw new Error('パスに .. は使えません');
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

    if (req.action === 'monitorUsage') {
      if (!req.monitor) throw new Error('monitor は必須です');
      if (req.monitor.type !== 'manual') throw new Error('手動カウンタ以外は記録できません');
      const snapshot = recordManualUsage(req.monitor, req.delta, req.reset === true);
      return jsonOut({ ok: true, snapshot: snapshot });
    }

    if (req.action === 'monitorControl') {
      if (!req.monitor) throw new Error('monitor は必須です');
      const snapshot = sendMonitorControl(req.monitor, req.command);
      return jsonOut({ ok: true, snapshot: snapshot });
    }

    if (req.action === 'monitorAdjust') {
      if (!req.monitor) throw new Error('monitor は必須です');
      const snapshot = setMonitorAdjust(req.monitor, req.value, req.reset === true);
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
  Logger.log('コミット先: ' + owner + '/' + repo + ' の ' + repoPath('index.json'));
  Logger.log(res.getContentText().slice(0, 300));
}

function testNotionConnection() {
  const pages = fetchNotionPages();
  Logger.log('取得件数: ' + pages.length);
  Logger.log(JSON.stringify(buildResources(pages), null, 2));
}

/** Copilot監視の設定確認。登録済みプロパティとAPIの応答をログに出す */
function testCopilotQuota() {
  Logger.log('MONITOR_TOKEN_COPILOT: ' + (PROP.getProperty('MONITOR_TOKEN_COPILOT') ? 'あり' : 'なし'));
  Logger.log('GITHUB_TOKEN: ' + (cfg('GITHUB_TOKEN') ? 'あり' : 'なし'));
  Logger.log('GITHUB_OWNER: ' + (cfg('GITHUB_OWNER') || 'なし'));

  // used が 0 のとき、明細が空なのか sku の判定漏れなのかを切り分ける
  var token = PROP.getProperty('MONITOR_TOKEN_COPILOT') || cfg('GITHUB_TOKEN');
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var url = 'https://api.github.com/users/' + encodeURIComponent(cfg('GITHUB_OWNER')) +
            '/settings/billing/usage?year=' + Utilities.formatDate(now, tz, 'yyyy') +
            '&month=' + Number(Utilities.formatDate(now, tz, 'MM'));
  var raw = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  Logger.log('HTTP ' + raw.getResponseCode());

  // product / sku ごとに集計して、Copilot の明細が含まれるかを確認する
  var items = (JSON.parse(raw.getContentText()).usageItems) || [];
  Logger.log('usageItems 件数: ' + items.length);
  var agg = {};
  items.forEach(function (it) {
    var key = it.product + ' | ' + it.sku + ' | ' + it.unitType;
    agg[key] = (agg[key] || 0) + (Number(it.quantity) || 0);
  });
  Object.keys(agg).forEach(function (k) { Logger.log(k + ' => ' + agg[k]); });

  Logger.log(JSON.stringify(
    fetchCopilotStatus({ id: 'copilot', name: 'GitHub Copilot', monthlyLimit: 300 }), null, 2));
}