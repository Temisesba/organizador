// Service Worker de Organizador — red primero, caché como respaldo sin
// conexión. Con "caché primero" (como estaba antes, v1) siempre veías la
// copia de la carga ANTERIOR — cada cambio nuevo tardaba una recarga
// extra en notarse, y a veces daba la impresión de que algo "seguía sin
// funcionar" cuando en realidad ya estaba arreglado, solo que el
// navegador no lo había pedido todavía. Con la app abierta con internet
// (el caso normal) siempre se pide la versión más nueva; solo si de
// verdad no hay conexión se usa lo último que se guardó.
const CACHE_NAME = 'organizador-shell-v2';
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
    fetch(req)
      .then((res) => {
        const copia = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copia));
        return res;
      })
      .catch(() => caches.match(req).then((cacheado) => cacheado || caches.match(SHELL_FILE)))
  );
});
