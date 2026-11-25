/* 在庫確認：オフライン用 Service Worker（任意）
 * index.html と同じフォルダに置くと、一度開いたページを端末に保存し、
 * 通信できないときも開けるようになります。無くても index.html は単体で動きます。
 * 表示は保存済みの版を先に使い、裏で新しい版を取りに行きます（次回の起動で反映）。 */
var CACHE = 'zaiko-v1';

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(['./']); }).catch(function () {}));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('message', function (e) {
  if (e.data && e.data.type === 'cache' && typeof e.data.url === 'string') {
    var url = new URL(e.data.url);
    if (url.origin === self.location.origin) {
      e.waitUntil(caches.open(CACHE).then(function (c) { return c.add(url.href); }).catch(function () {}));
    }
  }
});

/* 通信できるときは最新版を取り、保存し直す。通信できない・4秒以上かかるときは保存済みの版を使う */
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  var network = fetch(req, { cache: 'no-store' }).then(function (res) {
    if (res && res.ok) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); });
    }
    return res;
  });
  var fallback = function () {
    return caches.match(req, { ignoreSearch: true }).then(function (cached) {
      return cached || new Response('オフラインのため開けません。通信できる状態で一度開いてください。', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
    });
  };
  e.respondWith(new Promise(function (resolve) {
    var done = false;
    var timer = setTimeout(function () {
      caches.match(req, { ignoreSearch: true }).then(function (cached) { if (cached && !done) { done = true; resolve(cached); } });
    }, 4000);
    network.then(function (res) {
      if (done) return;
      done = true; clearTimeout(timer);
      if (res && (res.ok || res.status < 500)) resolve(res); else fallback().then(resolve);
    }, function () {
      if (done) return;
      done = true; clearTimeout(timer); fallback().then(resolve);
    });
  }));
});
