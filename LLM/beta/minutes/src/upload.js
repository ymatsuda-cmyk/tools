// src/upload.js — 音声アップロード機能
//
// main.js の内部実装には依存せず、DOM 挿入と localStorage だけで完結させている。
// ヘッダーの #open-upload ボタンでモーダルを開き、タイトル・日時・音声ファイル
// (ドラッグ&ドロップ対応) を入力して Google Drive の audio-inbox へ直接アップロード
// する。Drive API は CORS 対応済みなので GAS を経由しない。
//
// アップロード後、一覧の先頭に「新規」バッジ付きの仮カードを表示する。
// 実際の文字起こしは Mac mini 側の drive_inbox.py が担当し、Notion 登録が
// index.json に反映されて一覧に本物のカードが現れたら、仮カードは自動で消える
// （タイトル文字列の突き合わせで判定している）。

const CFG_KEY = 'minutes:upload';
const PENDING_KEY = 'minutes:pendingUploads';
const SCOPE = 'https://www.googleapis.com/auth/drive.file';
const CHUNK = 8 * 1024 * 1024;
const AUDIO_EXT = ['.m4a', '.mp3', '.wav', '.mp4', '.aac', '.flac', '.ogg'];
const PENDING_TTL_MS = 3 * 60 * 60 * 1000; // 3時間。失敗時にカードが残り続けないための保険

// ---------------------------------------------------------------- storage

function loadCfg() {
  try { return JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); } catch { return {}; }
}
function saveCfg(c) { localStorage.setItem(CFG_KEY, JSON.stringify(c)); }
function loadPending() {
  try { return JSON.parse(localStorage.getItem(PENDING_KEY) || '[]'); } catch { return []; }
}
function savePending(list) { localStorage.setItem(PENDING_KEY, JSON.stringify(list)); }

// ---------------------------------------------------------------- auth

let tokenClient = null;
let cachedToken = null;
let tokenExpiresAt = 0;

/** ユーザー操作起因で呼ぶこと。ポップアップがブロックされる。 */
function getToken(clientId) {
  if (cachedToken && Date.now() < tokenExpiresAt - 60_000) {
    return Promise.resolve(cachedToken);
  }
  if (!window.google?.accounts?.oauth2) {
    return Promise.reject(new Error(
      'Google Identity Services が読み込まれていません。index.html の <head> に GSI の script タグが必要です。'
    ));
  }
  return new Promise((resolve, reject) => {
    if (!tokenClient) {
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPE,
        callback: (res) => {
          if (res.error) return reject(new Error(res.error_description || res.error));
          cachedToken = res.access_token;
          tokenExpiresAt = Date.now() + (res.expires_in ?? 3600) * 1000;
          resolve(cachedToken);
        },
        error_callback: (err) => reject(new Error(err.message || '認証がキャンセルされました')),
      });
    } else {
      tokenClient.callback = (res) => {
        if (res.error) return reject(new Error(res.error_description || res.error));
        cachedToken = res.access_token;
        tokenExpiresAt = Date.now() + (res.expires_in ?? 3600) * 1000;
        resolve(cachedToken);
      };
    }
    tokenClient.requestAccessToken({ prompt: cachedToken ? '' : 'select_account' });
  });
}

// ---------------------------------------------------------------- drive

/** multipart/related で1回POST。5MB程度までの小さいファイル・JSON向け */
async function uploadMultipart(token, meta, blob, mime) {
  const boundary = '----minutesUploader' + Math.random().toString(36).slice(2);
  const body = new Blob([
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
    JSON.stringify(meta),
    `\r\n--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`,
    blob,
    `\r\n--${boundary}--\r\n`,
  ]);
  const res = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body,
    }
  );
  if (!res.ok) throw new Error(`Drive upload failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/** resumable upload。分割送信するので容量上限がなく、途中失敗にも強い */
async function uploadResumable(token, meta, file, onProgress) {
  const init = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Upload-Content-Type': file.type || 'application/octet-stream',
        'X-Upload-Content-Length': String(file.size),
      },
      body: JSON.stringify(meta),
    }
  );
  if (!init.ok) throw new Error(`Drive session failed (${init.status}): ${await init.text()}`);

  const sessionUrl = init.headers.get('Location');
  if (!sessionUrl) {
    // CORS で Location が読めない環境向けのフォールバック
    if (file.size > 5 * 1024 * 1024) {
      throw new Error(
        'resumable セッション URL が取得できず、ファイルが 5MB を超えています。' +
        'ブラウザの拡張機能を無効にして再試行してください。'
      );
    }
    onProgress(0.5);
    const r = await uploadMultipart(token, meta, file, file.type || 'application/octet-stream');
    onProgress(1);
    return r;
  }

  let offset = 0;
  while (offset < file.size) {
    const end = Math.min(offset + CHUNK, file.size);
    const res = await fetch(sessionUrl, {
      method: 'PUT',
      headers: { 'Content-Range': `bytes ${offset}-${end - 1}/${file.size}` },
      body: file.slice(offset, end),
    });

    if (res.status === 308) {
      const range = res.headers.get('Range');
      offset = range ? Number(range.split('-')[1]) + 1 : end;
    } else if (res.ok) {
      onProgress(1);
      return res.json();
    } else {
      throw new Error(`chunk ${offset} failed (${res.status}): ${await res.text()}`);
    }
    onProgress(offset / file.size);
  }
  throw new Error('アップロードが完了しませんでした');
}

// ---------------------------------------------------------------- naming

function pad(n) { return String(n).padStart(2, '0'); }
function toIsoJst(localValue) { return `${localValue}:00+09:00`; }
function stampFrom(localValue) {
  const [d, t] = localValue.split('T');
  return `${d.replace(/-/g, '')}-${t.replace(':', '')}`;
}
function slugify(s) {
  return (s || 'untitled')
    .replace(/[\\/:*?"<>|\s]+/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 60);
}
function extOf(name) {
  const m = name.match(/\.[^.]+$/);
  return m ? m[0].toLowerCase() : '';
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

// ---------------------------------------------------------------- pending list

function renderPending() {
  const host = document.getElementById('list-items');
  if (!host) return;
  const items = loadPending();
  let box = document.getElementById('pending-uploads');

  if (!items.length) {
    box?.remove();
    return;
  }
  if (!box) {
    box = document.createElement('div');
    box.id = 'pending-uploads';
    host.prepend(box);
  }

  box.innerHTML = items.map((p) => `
    <div class="list-item pending-list-item" data-pending-id="${p.id}">
      <div class="list-item-time">${(p.meetingAt || '').slice(11, 16)}</div>
      <div class="list-item-body">
        <div class="list-item-title">${escapeHtml(p.title)}</div>
        <div class="list-item-meta">
          <span class="badge status-新規">新規</span>
          ${p.status === 'uploading' ? ` 送信中 ${p.progress ?? 0}%` : ' 文字起こし待ち'}
        </div>
      </div>
    </div>
  `).join('');
}

function addPending(entry) {
  const items = loadPending();
  items.unshift(entry);
  savePending(items);
  renderPending();
}
function updatePending(id, patch) {
  savePending(loadPending().map((p) => (p.id === id ? { ...p, ...patch } : p)));
  renderPending();
}
function removePending(id) {
  savePending(loadPending().filter((p) => p.id !== id));
  renderPending();
}

/** 一覧に本物のカードが現れたら仮カードを消す。タイトル文字列の一致で判定するので
 *  main.js がどうやって一覧を描画しているかは問わない。 */
function reconcilePending() {
  const host = document.getElementById('list-items');
  if (!host) return;
  const items = loadPending();
  if (!items.length) return;

  const liveTitles = Array.from(
    host.querySelectorAll('.list-item:not(.pending-list-item) .list-item-title')
  ).map((el) => el.textContent.trim());

  let next = items.filter((p) => !liveTitles.includes(p.title.trim()));
  next = next.filter((p) => Date.now() - p.createdAt < PENDING_TTL_MS);

  if (next.length !== items.length) {
    savePending(next);
    renderPending();
  }
}

// ---------------------------------------------------------------- modal

function openModal() {
  const cfg = loadCfg();
  const root = document.getElementById('modal-root');
  if (!root) return;

  const overlay = document.createElement('div');
  overlay.className = 'raw-modal-overlay';
  overlay.innerHTML = `
    <div class="raw-modal" style="width:min(480px,100%)">
      <div class="raw-modal-header">
        <span>音声をアップロード</span>
        <button class="btn-ghost" id="up-close" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
      </div>
      <div class="raw-modal-body">
        <div class="upload-field">
          <label for="up-title">タイトル</label>
          <input type="text" id="up-title" placeholder="GMO自動化テスト 定例">
        </div>
        <div class="upload-field">
          <label for="up-at">打合せ日時</label>
          <input type="datetime-local" id="up-at">
        </div>
        <div class="upload-field">
          <label>音声ファイル</label>
          <div class="upload-drop" id="up-drop" tabindex="0" role="button">
            <strong id="up-drop-name">クリックまたはドラッグして選択</strong>
            <span id="up-drop-size">m4a / mp3 / wav / mp4</span>
          </div>
          <input type="file" id="up-file" accept="audio/*,video/mp4" hidden>
        </div>
        <div class="upload-progress" id="up-bar"><div class="upload-progress-fill" id="up-bar-fill"></div></div>
        <div class="upload-msg" id="up-msg" role="status" aria-live="polite"></div>

        <details id="up-setup" ${cfg.clientId && cfg.folderId ? '' : 'open'} style="margin-top:10px">
          <summary style="cursor:pointer;font-size:11px;color:var(--text-muted)">接続設定</summary>
          <div class="upload-field" style="margin-top:8px">
            <label for="up-cid">OAuth クライアント ID</label>
            <input type="text" id="up-cid" value="${escapeHtml(cfg.clientId || '')}" placeholder="xxxx.apps.googleusercontent.com">
          </div>
          <div class="upload-field">
            <label for="up-fid">audio-inbox のフォルダ ID</label>
            <input type="text" id="up-fid" value="${escapeHtml(cfg.folderId || '')}" placeholder="Drive の URL の /folders/ 以降">
          </div>
        </details>
      </div>
      <div class="raw-modal-footer">
        <button class="btn" id="up-cancel">キャンセル</button>
        <button class="btn" id="up-go" disabled>アップロード</button>
      </div>
    </div>
  `;
  root.appendChild(overlay);

  const $ = (s) => overlay.querySelector(s);
  const drop = $('#up-drop');
  const fileInput = $('#up-file');
  const titleInput = $('#up-title');
  const atInput = $('#up-at');
  const go = $('#up-go');
  const bar = $('#up-bar');
  const barFill = $('#up-bar-fill');
  const msg = $('#up-msg');
  const cidInput = $('#up-cid');
  const fidInput = $('#up-fid');

  let picked = null;

  function say(text, kind = '') {
    msg.textContent = text;
    msg.className = `upload-msg ${kind}`;
  }
  function close() { overlay.remove(); }

  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
  $('#up-close').onclick = close;
  $('#up-cancel').onclick = close;
  document.addEventListener('keydown', function onEsc(e) {
    if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onEsc); }
  });

  function accept(file) {
    if (!file) return;
    const ext = extOf(file.name);
    if (!AUDIO_EXT.includes(ext)) {
      say(`${ext || 'その拡張子'} は扱えません。m4a / mp3 / wav などを選んでください。`, 'error');
      return;
    }
    picked = file;
    const d = new Date(file.lastModified);
    if (!atInput.value) {
      atInput.value = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
    if (!titleInput.value) titleInput.value = file.name.replace(/\.[^.]+$/, '');
    $('#up-drop-name').textContent = file.name;
    $('#up-drop-size').textContent = `${(file.size / 1048576).toFixed(1)} MB`;
    go.disabled = false;
    say('');
  }

  drop.onclick = () => fileInput.click();
  drop.onkeydown = (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  };
  fileInput.onchange = () => accept(fileInput.files[0]);
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('drag-over'); };
  drop.ondragleave = () => drop.classList.remove('drag-over');
  drop.ondrop = (e) => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    accept(e.dataTransfer.files[0]);
  };

  cidInput.onchange = () => saveCfg({ ...loadCfg(), clientId: cidInput.value.trim() });
  fidInput.onchange = () => saveCfg({ ...loadCfg(), folderId: fidInput.value.trim() });

  go.onclick = async () => {
    const c = loadCfg();
    if (!c.clientId || !c.folderId) return say('接続設定にクライアント ID とフォルダ ID を入れてください。', 'error');
    if (!picked) return say('音声ファイルを選んでください。', 'error');
    if (!atInput.value) return say('打合せ日時を入れてください。', 'error');

    go.disabled = true;
    bar.style.display = 'block';
    barFill.style.width = '0%';

    const pendingId = `p${Date.now()}`;
    const title = titleInput.value.trim() || picked.name.replace(/\.[^.]+$/, '');
    const meetingAt = toIsoJst(atInput.value);

    addPending({ id: pendingId, title, meetingAt, status: 'uploading', progress: 0, createdAt: Date.now() });

    try {
      say('Google に接続しています…');
      const token = await getToken(c.clientId);

      const base = `${stampFrom(atInput.value)}_${slugify(title)}_${Math.random().toString(16).slice(2, 6)}`;
      const audioName = base + extOf(picked.name);

      say('音声を送っています…');
      await uploadResumable(token, { name: audioName, parents: [c.folderId] }, picked, (p) => {
        const pct = Math.round(p * 100);
        barFill.style.width = `${pct}%`;
        say(`音声を送っています… ${pct}%`);
        updatePending(pendingId, { progress: pct });
      });

      // 音声の送信完了後に JSON を置く。これが Mac mini 側の処理開始の合図。
      say('メタデータを登録しています…');
      const meta = {
        audio: audioName,
        title,
        meetingAt,
        source: 'minutes-viewer',
        uploadedAt: new Date().toISOString(),
      };
      await uploadMultipart(
        token,
        { name: base + '.json', parents: [c.folderId] },
        new Blob([JSON.stringify(meta, null, 2)], { type: 'application/json' }),
        'application/json'
      );

      updatePending(pendingId, { status: 'queued' });
      say('アップロードしました。文字起こしが終わると一覧に反映されます。', 'ok');
      setTimeout(close, 900);
    } catch (err) {
      console.error(err);
      say(err.message, 'error');
      removePending(pendingId);
      go.disabled = false;
      bar.style.display = 'none';
    }
  };

  titleInput.focus();
}

// ---------------------------------------------------------------- init

export function initUpload() {
  const btn = document.getElementById('open-upload');
  if (btn) btn.onclick = openModal;

  renderPending();
  setInterval(reconcilePending, 20_000);

  // 一覧の再描画タイミングが main.js 側でいつ起きるか分からないので、
  // DOM の変化を見て即座に照合する。
  const host = document.getElementById('list-items');
  if (host) {
    const mo = new MutationObserver(() => reconcilePending());
    mo.observe(host, { childList: true, subtree: true });
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initUpload);
} else {
  initUpload();
}
