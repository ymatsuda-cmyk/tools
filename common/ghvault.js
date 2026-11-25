/* GhVault — アプリのデータを暗号化して、GitHub のプライベートリポジトリ（Contents API）に保存する共通部品
 *
 *   ・暗号化：PBKDF2-SHA256（60万回）＋ AES-256-GCM。パスフレーズは GitHub にも送らない
 *   ・接続情報（owner / repo / branch / path / token）とパスフレーズは、この端末の localStorage に保存する。
 *     一度入れれば、次回からは自動で暗号化・復号する
 *   ・他の端末へは common/devicelink.js（QR・6桁コード・リンク）で渡す（package() で包みを作る）
 *
 *   const vault = GhVault.create({ app:'dashboard', defaultPath:'dashboard/state.enc.json' });
 *   vault.configured()                 → 接続情報とパスフレーズが揃っているか
 *   await vault.load()                 → { data, sha } ／ 保存ファイルがまだ無ければ null
 *   await vault.save(data, {force})    → 保存（他の端末が先に更新していれば Error.code === 'CONFLICT'。force:true で上書き）
 *   await vault.package()              → DeviceLink の包み（他の端末へ渡す用）
 *
 *   // 接続の試行（保存はしない）→ 成功したら adopt で端末に保存する
 *   const t = GhVault.scratch(opts, conn, pass);  await t.check();  const r = await t.load();  vault.adopt(t);
 *
 *   Error.code：BADPASS（パスフレーズ違い／改ざん）, CONFLICT, CONFIG, NOCONFIG, NOBRANCH, FORMAT, NOCRYPTO, NODL, HTTP_<status>
 */
(function(){
'use strict';

const ITER = 600000, FORMAT = 'ghvault/v1';
const te = new TextEncoder(), td = new TextDecoder();
const b2a = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
const a2b = s => Uint8Array.from(atob(String(s).replace(/\s/g, '')), c => c.charCodeAt(0));
const APP_RE = /^[a-z0-9-]{1,32}$/, NAME_RE = /^[A-Za-z0-9._-]{1,100}$/, PATH_RE = /^[\w.\/-]+$/;

function err(code, message){ const e = new Error(message); e.code = code; return e; }
const HTTP_MSG = {
  401: 'トークンが無効か期限切れです（401）',
  403: '権限がありません（403）。トークンの対象リポジトリと、Contents の「Read and write」権限を確認してください',
  404: 'リポジトリが見つかりません（404）。オーナー名・リポジトリ名と、トークンの対象リポジトリを確認してください'
};
const httpErr = status => err('HTTP_' + status, HTTP_MSG[status] || `GitHub がエラーを返しました（${status}）`);

function normalizeConn(c, defPath){
  const o = {
    owner:  String(c && c.owner  || '').trim(),
    repo:   String(c && c.repo   || '').trim(),
    branch: String(c && c.branch || '').trim() || 'main',
    path:   String(c && c.path   || '').trim().replace(/^\/+|\/+$/g, '') || defPath,
    token:  String(c && c.token  || '').trim()
  };
  if (!NAME_RE.test(o.owner)) throw err('CONFIG', 'オーナー名が正しくありません');
  if (!NAME_RE.test(o.repo)) throw err('CONFIG', 'リポジトリ名が正しくありません');
  if (!PATH_RE.test(o.branch) || o.branch.includes('..')) throw err('CONFIG', 'ブランチ名が正しくありません');
  if (!PATH_RE.test(o.path) || o.path.includes('..')) throw err('CONFIG', '保存ファイルのパスが正しくありません');
  if (!o.token) throw err('CONFIG', 'アクセストークンを入力してください');
  return o;
}

const ls = {
  get(k){ try{ return localStorage.getItem(k); }catch{ return null; } },
  set(k, v){ try{ localStorage.setItem(k, v); return true; }catch{ return false; } },
  del(k){ try{ localStorage.removeItem(k); }catch{} }
};

function make(opt, persist, init){
  const app = opt && opt.app;
  if (!APP_RE.test(app || '')) throw err('CONFIG', 'app は英小文字・数字・ハイフンで指定してください');
  const defPath = opt.defaultPath || `${app}/data.enc.json`;
  const K_CONN = `${app}.ghvault.conn`, K_PASS = `${app}.ghvault.pass`;
  let conn = null, pass = '', sha = null, salt = null, keys = new Map();

  if (init){ conn = normalizeConn(init.conn, defPath); pass = String(init.pass || ''); }
  else if (persist){
    try{ conn = normalizeConn(JSON.parse(ls.get(K_CONN) || 'null'), defPath); }catch{ conn = null; }
    pass = ls.get(K_PASS) || '';
  }

  /* ---------- 暗号 ---------- */
  const aad = () => te.encode(`ghvault:${app}`);
  async function keyFor(s, iter){
    const id = s + ':' + iter;
    if (keys.has(id)) return keys.get(id);
    if (!(globalThis.crypto && crypto.subtle)) throw err('NOCRYPTO', 'この環境では暗号化を使えません（https か localhost で開いてください）');
    const base = await crypto.subtle.importKey('raw', te.encode(pass), 'PBKDF2', false, ['deriveKey']);
    const k = await crypto.subtle.deriveKey({ name:'PBKDF2', salt:a2b(s), iterations:iter, hash:'SHA-256' },
      base, { name:'AES-GCM', length:256 }, false, ['encrypt', 'decrypt']);
    keys.set(id, k);
    return k;
  }
  async function seal(data){
    if (!salt) salt = b2a(crypto.getRandomValues(new Uint8Array(16)));
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const k = await keyFor(salt, ITER);
    const ct = new Uint8Array(await crypto.subtle.encrypt({ name:'AES-GCM', iv, additionalData:aad() }, k, te.encode(JSON.stringify(data))));
    return { format:FORMAT, app, kdf:{ name:'PBKDF2', hash:'SHA-256', iterations:ITER, salt },
             cipher:{ name:'AES-GCM', iv:b2a(iv) }, data:b2a(ct) };
  }
  async function unseal(env){
    if (!env || env.format !== FORMAT || !env.kdf || !env.cipher) throw err('FORMAT', '暗号化ファイルの形式が違います');
    if (env.app !== app) throw err('FORMAT', '別のアプリのデータファイルです');
    const iter = Number(env.kdf.iterations);
    if (!(iter >= 100000 && iter <= 5000000)) throw err('FORMAT', '暗号化ファイルの設定値が正しくありません');
    const k = await keyFor(env.kdf.salt, iter);
    let pt;
    try{ pt = await crypto.subtle.decrypt({ name:'AES-GCM', iv:a2b(env.cipher.iv), additionalData:aad() }, k, a2b(env.data)); }
    catch{ throw err('BADPASS', 'パスフレーズが違うか、データが壊れています'); }
    salt = env.kdf.salt;      // 以後の保存でも同じ salt を使い、鍵の再計算（約1秒）を避ける
    return JSON.parse(td.decode(pt));
  }

  /* ---------- GitHub ---------- */
  const needConn = () => { if (!conn) throw err('NOCONFIG', 'GitHub の接続情報が未設定です'); };
  const need = () => { needConn(); if (!pass) throw err('NOCONFIG', 'パスフレーズが未設定です'); };
  const encPath = p => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');
  const repoURL = () => `https://api.github.com/repos/${encodeURIComponent(conn.owner)}/${encodeURIComponent(conn.repo)}`;
  const fileURL = () => `${repoURL()}/contents/${encPath(conn.path)}`;
  function gh(url, o){
    o = o || {};
    const headers = { 'Accept':'application/vnd.github+json', 'X-GitHub-Api-Version':'2022-11-28', 'Authorization':`Bearer ${conn.token}` };
    if (o.body) headers['Content-Type'] = 'application/json';
    return fetch(url, { ...o, cache:'no-store', headers });
  }

  /** リポジトリとブランチに届くか確認する。private が false なら公開リポジトリ（暗号化はされるが非推奨） */
  async function check(){
    needConn();
    let r = await gh(repoURL());
    if (!r.ok) throw httpErr(r.status);
    const repo = await r.json();
    if (repo.permissions && repo.permissions.push === false) throw err('HTTP_403', 'このトークンには書き込み権限がありません。Contents の「Read and write」を付けてください');
    r = await gh(`${repoURL()}/branches/${encPath(conn.branch)}`);
    if (r.status === 404) throw err('NOBRANCH', `ブランチ「${conn.branch}」が見つかりません`);
    if (!r.ok) throw httpErr(r.status);
    return { private: !!repo.private };
  }

  async function readRaw(skipParse){
    const r = await gh(`${fileURL()}?ref=${encodeURIComponent(conn.branch)}`);
    if (r.status === 404) return null;
    if (!r.ok) throw httpErr(r.status);
    const j = await r.json();
    if (Array.isArray(j)) throw err('FORMAT', '保存ファイルのパスがフォルダを指しています');
    if (skipParse) return { sha:j.sha };
    let b64 = j.content;
    if (!b64 || j.encoding === 'none'){                       // 1MB を超えると content が空になる → blob から取る
      const b = await gh(j.git_url);
      if (!b.ok) throw httpErr(b.status);
      b64 = (await b.json()).content;
    }
    let env;
    try{ env = JSON.parse(td.decode(a2b(b64))); }catch{ throw err('FORMAT', '保存ファイルを読み取れません（暗号化ファイルではありません）'); }
    return { sha:j.sha, env };
  }

  async function load(){
    need();
    const f = await readRaw();
    if (!f){ sha = null; return null; }
    const data = await unseal(f.env);
    sha = f.sha;
    return { data, sha };
  }

  async function save(data, o){
    need();
    if (o && o.force){ const f = await readRaw(true); sha = f ? f.sha : null; }
    const env = await seal(data);
    const body = { message:`Update ${conn.path}`, content:b2a(te.encode(JSON.stringify(env, null, 2))), branch:conn.branch };
    if (sha) body.sha = sha;
    const r = await gh(fileURL(), { method:'PUT', body:JSON.stringify(body) });
    if (r.status === 409 || r.status === 422) throw err('CONFLICT', '他の端末で先に更新されています');
    if (!r.ok) throw httpErr(r.status);
    sha = (await r.json()).content.sha;
    return sha;
  }

  /* ---------- 端末への保存・他の端末へ渡す ---------- */
  function adopt(from){                       // scratch で試して成功した接続を、この端末に保存して使い始める
    const s = from._state();
    conn = s.conn; pass = s.pass; sha = s.sha; salt = s.salt; keys = s.keys;
    if (!persist) return true;
    const a = ls.set(K_CONN, JSON.stringify(conn)), b = ls.set(K_PASS, pass);
    return a && b;                            // false ＝ この端末に保存できなかった（プライベートブラウズなど）
  }
  function forget(){
    conn = null; pass = ''; sha = null; salt = null; keys = new Map();
    if (persist){ ls.del(K_CONN); ls.del(K_PASS); }
  }
  async function pack(){
    need();
    const DL = globalThis.DeviceLink;
    if (!DL) throw err('NODL', '受け渡しの部品（common/devicelink.js）を読み込めませんでした。オンラインで開き直してください');
    return DL.seal({ owner:conn.owner, repo:conn.repo, branch:conn.branch, path:conn.path, token:conn.token }, pass,
      { app:opt.dlApp || app, label:`${conn.owner}/${conn.repo}` });
  }

  return {
    configured: () => !!(conn && pass),
    describe: () => conn ? { owner:conn.owner, repo:conn.repo, branch:conn.branch, path:conn.path, tokenSet:!!conn.token } : null,
    check, load, save, adopt, forget, package:pack,
    _state: () => ({ conn, pass, sha, salt, keys })
  };
}

window.GhVault = Object.freeze({
  version: FORMAT,
  create: opt => make(opt, true),
  scratch: (opt, conn, pass) => make(opt, false, { conn, pass })
});
})();
