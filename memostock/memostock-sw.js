// MemoStock をオフラインでも開けるようにするサービスワーカー
// secure-memo.html と同じフォルダに置いてください。メモの中身（GitHub・JSONBin との通信）はキャッシュしません。
const CACHE = 'memostock-v1';
const LIB_HOSTS = ['cdn.jsdelivr.net', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.add('./secure-memo.html')).catch(() => {}));
});
self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const u = new URL(req.url);
  // アプリ本体：まずネットから取り、つながらないときは保存しておいた版を使う
  if (u.origin === location.origin && /secure-memo\.html$|\/$/.test(u.pathname)){
    e.respondWith((async () => {
      try{
        const r = await fetch(req, { cache:'no-store' });
        if (r.ok){ const c = await caches.open(CACHE); c.put('./secure-memo.html', r.clone()); }
        return r;
      }catch{
        return (await caches.match('./secure-memo.html')) || Response.error();
      }
    })());
    return;
  }
  // マインドマップ・手書きボードのライブラリとフォント：保存しておいた版をすぐ返し、裏で更新
  if (LIB_HOSTS.includes(u.hostname)){
    e.respondWith((async () => {
      const c = await caches.open(CACHE);
      const hit = await c.match(req);
      const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => null);
      return hit || (await net) || Response.error();
    })());
  }
});
