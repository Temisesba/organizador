// Service Worker de Organizador — caché primero, red en segundo plano
// (mismo patrón que Gastos/sw.js): si hay copia guardada se sirve de
// inmediato y la red actualiza la caché en silencio para la próxima vez.
const CACHE_NAME = 'organizador-shell-v1';
const SHELL_FILE = 'organizador.html';

self.addEventListener('install', () => { self.skipWaiting(); });

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(nombres.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const esMismoOrigen = new URL(req.url).origin === self.location.origin;
  if (req.method !== 'GET' || !esMismoOrigen) return;

  event.respondWith(
    caches.match(req).then((cacheado) => {
      const enRed = fetch(req)
        .then((res) => {
          const copia = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copia));
          return res;
        })
        .catch(() => cacheado || caches.match(SHELL_FILE));
      return cacheado || enRed;
    })
  );
});
