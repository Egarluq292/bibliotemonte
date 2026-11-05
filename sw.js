// Service worker de la Biblioteca. Sube VERSION en cada actualización.
const VERSION = 'biblioteca-v24';
const BASICOS = ['./', './index.html', './manifest.webmanifest', './img/biblioteca.webp', './img/icono-192.png', './img/icono-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(BASICOS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Las reservas y comentarios nunca se guardan en caché: siempre en directo.
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  // La página: primero la red (para recibir actualizaciones); si no hay conexión, la copia guardada.
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request, { cache: 'no-store' }).then(r => { const c = r.clone(); caches.open(VERSION).then(x => x.put('./index.html', c)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
