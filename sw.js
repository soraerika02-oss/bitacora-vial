// Service worker: guarda la app en el celular para que abra sin internet.
// Cuando publique una versión nueva de index.html, cambie el número de CACHE.
const CACHE = 'bitacora-vial-v1';
const ARCHIVOS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return; // Google Sheets va directo a la red
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then(enCache => {
      const red = fetch(req).then(resp => {
        if (resp && resp.ok) { const copia = resp.clone(); caches.open(CACHE).then(c => c.put(req, copia)); }
        return resp;
      }).catch(() => enCache || (req.mode === 'navigate' ? caches.match('./index.html') : undefined));
      return enCache || red;
    })
  );
});
