// 華森行程通報 Service Worker v5
const CACHE = 'huasen-v5';

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(c =>
      c.add(new Request(self.registration.scope, {cache:'reload'})).catch(()=>{})
    )
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
      .then(() => clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.includes('script.google.com') ||
      e.request.url.includes('fonts.googleapis.com') ||
      e.request.url.includes('fonts.gstatic.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached => {
      const net = fetch(e.request).then(res => {
        if (res && res.status === 200 && res.type !== 'opaque') {
          caches.open(CACHE).then(c => c.put(e.request, res.clone()));
        }
        return res;
      }).catch(() => null);
      return cached || net || new Response('Offline', {status:503});
    })
  );
});

self.addEventListener('message', e => {
  if (e.data?.type === 'SKIP_WAITING') self.skipWaiting();
});
