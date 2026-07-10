const CACHE_NAME = 'kyuyodaicho-v2';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    Promise.all([
      self.clients.claim(),
      caches.keys().then((names) => Promise.all(
        names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )),
    ])
  );
});

// シンプルなネットワーク優先＋キャッシュ保存（オフライン時のフォールバック）
// ただし、このアプリ自身のファイル（同一オリジン）だけを対象にする。
// GAS（script.google.com）やDriveの公開ドキュメントへの通信は、
// サービスワーカーが一切手を出さず、ブラウザにそのまま任せる。
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
