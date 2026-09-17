/* Dinero — service worker: guarda la app para que abra sin internet.
   Sube el número de CACHE cada vez que publiques una versión nueva. */

var CACHE = 'dinero-v2';
var ARCHIVOS = ['./', './index.html', './manifest.json', './icon.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(ARCHIVOS); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (llaves) {
      return Promise.all(llaves.map(function (k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  if (e.request.url.indexOf('script.google.com') > -1 || e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request).then(function (r) {
      if (r && r.status === 200 && r.type === 'basic') {
        var copia = r.clone();
        caches.open(CACHE).then(function (c) { c.put(e.request, copia); });
      }
      return r;
    }).catch(function () {
      return caches.match(e.request).then(function (r) {
        return r || caches.match('./index.html');
      });
    })
  );
});
