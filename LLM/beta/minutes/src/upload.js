// src/upload.js — 音声アップロード機能（ログイン不要版）
//
// main.js の内部実装には依存しないが、設定は共有する。GAS の URL とアクセス
// トークンは、既存の設定画面（main.js / minutes-config.js）で入力済みの
// localStorage "minutes:config" をそのまま読む。ここで別の接続設定を持たない。
//
// 【認証・通信経路について】
// Drive の resumable upload は「セッション URL を発行する最初の1回」だけ認証が要り、
// 発行された URL への実際のデータ送信（PUT）自体には認証が要らない。これを利用して
// 当初は「セッション発行だけ GAS に任せ、バイトはブラウザから直接 Google へ」という
// 構成を試みたが、Drive API v3 のセッション URL はブラウザからの直接 CORS アクセスに
// 対応しておらず（Origin ヘッダーを明示しても改善せず、502 も発生）、断念した。
// そのため、バイト自体も数MB単位のチャンクに分割して GAS 経由で中継する方式にした。
// 1回のリクエストは小さいままなので、GAS の1リクエストあたりのペイロード上限
// （約50MB）には当たらない。利用者は一切 Google にログインする必要がない。
//
// 権限タグについても自己申告のドロップダウンは持たない。設定画面で入力した
// コードが verifyCode_ で検証済みの role なので、それをそのまま使う
// （管理者コードなら権限は付けず、それ以外ならその role をタグとして付与する）。
//
// アップロード後、一覧の先頭に「新規」バッジ付きの仮カードを表示する。
// 実際の文字起こしは Mac mini 側の drive_inbox.py が担当し、Notion 登録が
// index.json に反映されて一覧に本物のカードが現れたら、仮カードは自動で消える
// （タイトルの一致件数がアップロード時点より増えたかで判定。単純な存在判定だと
// 毎週同名の会議で誤って早期に消えてしまうため件数比較にしている）。

const MAIN_CFG_KEY = 'minutes:config'; // main.js / minutes-config.js と同じキー
const ADMIN_ROLE = 'xYz'; // minutes-config.js の ADMIN_ROLE と一致させること
const PENDING_KEY = 'minutes:pendingUploads';
const PUT_CHUNK = 3 * 1024 * 1024; // base64化すると4/3倍になるので、小さめに保つ
const AUDIO_EXT = ['.m4a', '.mp3', '.wav', '.mp4', '.aac', '.flac', '.ogg'];
const PENDING_TTL_MS = 3 * 60 * 60 * 1000; // 3時間。失敗時にカードが残り続けないための保険

// ---------------------------------------------------------------- storage

function loadMainConfig() {
  try { return JSON.parse(localStorage.getItem(MAIN_CFG_KEY) || '{}'); } catch { return {}; }
}
function loadPending() {
  try { return JSON.parse(localStorage.getItem(PENDING_KEY) || '[]'); } catch { return []; }
}
function savePending(list) { localStorage.setItem(PENDING_KEY, JSON.stringify(list)); }

// ---------------------------------------------------------------- GAS

/**
 * GAS 呼び出しの共通部分。
 * text/plain で送るのは、application/json だとブラウザが CORS preflight (OPTIONS) を
 * 先に飛ばしてしまい、GAS の Web App がそれに正しく応答できず失敗するため。
 * 既存の Notion 連携アクションと同じ回避パターン・同じ token フィールドを使う。
 */
async function gasCall(gasUrl, payload) {
  const res = await fetch(gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
  const data = await res.json().catch(() => null);
  if (!res.ok || !data || data.ok === false) {
    throw new Error((data && data.error) || `GAS呼び出しに失敗 (${res.status})`);
  }
  return data.data;
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

/** 音声をチャンクに分割し、1つずつ GAS 経由で Google に中継する。 */
async function uploadToSession(gasUrl, token, sessionUrl, file, onProgress) {
  let offset = 0;
  while (offset < file.size) {
    const end = Math.min(offset + PUT_CHUNK, file.size);
    const buf = await file.slice(offset, end).arrayBuffer();
    const result = await gasCall(gasUrl, {
      token,
      action: 'putChunk',
      sessionUrl,
      offset,
      total: file.size,
      chunk: arrayBufferToBase64(buf),
    });
    offset = end;
    onProgress(offset / file.size);
    if (result.done) return result.file;
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

/** 一覧上の「タイトルが完全一致する項目」の件数を数える。同名の会議が
 *  繰り返し登録される場合があるため、存在有無ではなく件数で見る。 */
function countLiveTitle(host, title) {
  return Array.from(host.querySelectorAll('.list-item:not(.pending-list-item) .list-item-title'))
    .filter((el) => el.textContent.trim() === title).length;
}

function addPending(entry) {
  const host = document.getElementById('list-items');
  // アップロード時点で同名の項目が何件あったかを記録しておく。「存在するか」
  // ではなく「その数より増えたか」で判定しないと、毎週同じ議題名の会議（例:
  // 「GMO定例」）で先週分を「もう届いた」と誤認して仮カードを消してしまう。
  entry.baselineCount = host ? countLiveTitle(host, entry.title) : 0;
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

/** 一覧に本物のカードが現れたら仮カードを消す。タイトルの一致件数が
 *  アップロード時点より増えていたら「本当に届いた」と判断する。 */
function reconcilePending() {
  const host = document.getElementById('list-items');
  if (!host) return;
  const items = loadPending();
  if (!items.length) return;

  let next = items.filter((p) => countLiveTitle(host, p.title) <= (p.baselineCount ?? 0));
  next = next.filter((p) => Date.now() - p.createdAt < PENDING_TTL_MS);

  if (next.length !== items.length) {
    savePending(next);
    renderPending();
  }
}

// ---------------------------------------------------------------- modal

function openModal() {
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

  let picked = null;

  function say(text, kind = '') {
    msg.textContent = text;
    msg.className = `upload-msg ${kind}`;
  }
  function close() { overlay.remove(); }

  // 画面外タップや Escape での誤クローズを防ぐため、閉じる手段は
  // キャンセルボタンと右上の×だけにする。アップロード中の誤操作対策。
  $('#up-close').onclick = close;
  $('#up-cancel').onclick = close;

  const main = loadMainConfig();
  if (!main.gasUrl || !main.notionToken) {
    say('設定画面（歯車アイコン）で GAS の URL とコードを先に登録してください。', 'error');
  }

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

  go.onclick = async () => {
    const cfg = loadMainConfig();
    if (!cfg.gasUrl || !cfg.notionToken) {
      return say('設定画面（歯車アイコン）で GAS の URL とコードを先に登録してください。', 'error');
    }
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
      const base = `${stampFrom(atInput.value)}_${slugify(title)}_${Math.random().toString(16).slice(2, 6)}`;
      const audioName = base + extOf(picked.name);

      say('アップロード先を準備しています…');
      const { sessionUrl } = await gasCall(cfg.gasUrl, {
        token: cfg.notionToken,
        action: 'initUpload',
        filename: audioName,
        mimeType: picked.type || 'application/octet-stream',
        size: picked.size,
      });
      if (!sessionUrl) throw new Error('セッション URL を取得できませんでした');

      say('音声を送っています…');
      await uploadToSession(cfg.gasUrl, cfg.notionToken, sessionUrl, picked, (p) => {
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
      // 設定画面で検証済みの role をそのまま使う。管理者コードなら権限は付けず
      // 後から手動で割り当てる。自己申告のドロップダウンは持たせない
      // （なりすまし防止。role は verifyCode_ を通過した本物の権限）。
      if (cfg.role && cfg.role !== ADMIN_ROLE) meta.permission = [cfg.role];

      await gasCall(cfg.gasUrl, {
        token: cfg.notionToken,
        action: 'writeSidecar',
        name: base + '.json',
        meta,
      });

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
