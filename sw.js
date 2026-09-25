const APP_CACHE_VERSION = "2026.08.21.1";
const APP_CACHE_NAME = "family-app-" + APP_CACHE_VERSION;
const OFFLINE_URL = "offline.html?v=2026.08.21.1";
const STATIC_ASSETS = [
  OFFLINE_URL,
  "manifest.webmanifest?v=2026.08.21.1",
  "pwa-icon-192.png?v=2026.08.21.1",
  "pwa-icon-512.png?v=2026.08.21.1",
  "pwa-maskable-512.png?v=2026.08.21.1"
];

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(APP_CACHE_NAME)
      .then((cache) => cache.addAll(STATIC_ASSETS))
      .then(clearOldCaches)
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    clearOldCaches().then(() => self.clients.claim())
  );
});

self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(new Request(request, { cache: "no-store" }))
        .catch(() => caches.match(OFFLINE_URL))
    );
    return;
  }

  event.respondWith(
    fetch(new Request(request, { cache: "no-store" }))
      .catch(() => caches.match(request, { ignoreSearch: true }))
  );
});

function clearOldCaches() {
  return caches.keys().then((keys) => {
    return Promise.all(keys.map((key) => {
      if (key === APP_CACHE_NAME) return Promise.resolve();
      return caches.delete(key);
    }));
  });
}
