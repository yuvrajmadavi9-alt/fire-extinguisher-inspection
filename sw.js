const CACHE_NAME = "fire-inspection-v1";

const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json"
];

self.addEventListener("install", function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        return cache.addAll(APP_SHELL);
      })
      .then(function() {
        return self.skipWaiting();
      })
  );
});

self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.map(function(key) {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

self.addEventListener("fetch", function(event) {

  if (event.request.url.indexOf("script.google.com") !== -1) {
    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then(function(cached) {
        return cached || fetch(event.request)
          .catch(function() {
            return caches.match("./index.html");
          });
      })
  );
});
