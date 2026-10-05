/* LexiRef service worker — makes the app work fully offline. Bump CACHE on every release. */
var CACHE = 'lexiref-v1.0.0';
var SHELL = [
  './', './index.html', './lexiref-core.js', './manifest.json', './privacy.html',
  './icons/icon-16.png', './icons/icon-32.png', './icons/icon-48.png', './icons/icon-64.png', './icons/icon-128.png',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) { return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  // Never cache metadata APIs (Crossref / OpenAlex) or Office.js.
  if (url.origin !== location.origin) return;
  // App shell: cache first, then network (and refresh the cache in the background).
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(function (cached) {
    var fetched = fetch(e.request).then(function (res) {
      if (res && res.ok) caches.open(CACHE).then(function (c) { c.put(e.request, res.clone()); });
      return res;
    }).catch(function () { return cached; });
    return cached || fetched;
  }));
});
self.addEventListener('message', function (e) { if (e.data === 'skipWaiting') self.skipWaiting(); });
