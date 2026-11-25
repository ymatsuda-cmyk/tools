/**
 * DeviceLink 中継（Google Apps Script）— 全アプリ共通の1デプロイ
 *
 * 置くのは各アプリがパスフレーズで暗号化した「包み」だけ。コードは置き場所を探す番号で、
 * コードが知られても、パスフレーズが無ければ中身は読めない。
 *
 * ・アプリ名（app）ごとに置き場所を分ける
 * ・コードは6桁、5分で消える、1回読まれたら消える
 * ・間違ったコードが10分間に10回あると、そのアプリの受け取りを一時停止する（総当たり対策）
 *
 * デプロイ：［デプロイ］→［新しいデプロイ］→ 種類「ウェブアプリ」
 *   次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *   表示された URL（…/exec）を shared/devicelink.js の RELAY_URL に入れる
 * コードを直したときは［デプロイを管理］→ 編集 → バージョン「新バージョン」で同じ URL のまま更新できる
 */
const TTL = 300;          // コードの有効期間（秒）
const MAX_FAIL = 10;      // この回数間違えると受け取りを止める
const FAIL_WINDOW = 600;  // 間違いを数える期間・止める期間（秒）
const MAX_SIZE = 4000;    // 受け付ける大きさの上限（文字）
const APP_RE = /^[a-z0-9-]{1,32}$/;

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (_) { return out({ error: 'bad-request' }); }
  const app = String(req.app || '');
  if (!APP_RE.test(app)) return out({ error: 'bad-request' });
  const cache = CacheService.getScriptCache();
  const key = code => 'c:' + app + ':' + code;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    if (req.action === 'put') {
      const data = String(req.data || '');
      if (!data || data.length > MAX_SIZE || data.split('.')[1] !== app) return out({ error: 'bad-request' });
      for (let i = 0; i < 20; i++) {
        const code = newCode();
        if (cache.get(key(code))) continue;
        cache.put(key(code), data, TTL);
        return out({ code: code, ttl: TTL });
      }
      return out({ error: 'busy' });
    }
    if (req.action === 'get') {
      const fk = 'f:' + app;
      const fails = Number(cache.get(fk) || 0);
      if (fails >= MAX_FAIL) return out({ error: 'locked' });
      const code = String(req.code || '').replace(/\D/g, '');
      const data = code.length === 6 ? cache.get(key(code)) : null;
      if (!data) {
        cache.put(fk, String(fails + 1), FAIL_WINDOW);
        return out({ error: 'not-found' });
      }
      cache.remove(key(code));   // 1回きり
      return out({ data: data });
    }
    if (req.action === 'cancel') {
      const code = String(req.code || '').replace(/\D/g, '');
      if (code.length === 6) cache.remove(key(code));
      return out({ ok: true });
    }
    return out({ error: 'bad-request' });
  } finally {
    lock.releaseLock();
  }
}

// 予測しにくい6桁（UUID v4 の乱数部分から作る）
function newCode() {
  const hex = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  return String(parseInt(hex, 16) % 1000000).padStart(6, '0');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
