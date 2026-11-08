// Cash Flow — keeps the app saved on your device so it opens with no signal.
// Bump VERSION whenever index.html changes so phones pick up the new version.
const VERSION = "cashflow-v40";
const FILES = ["./", "./index.html", "./manifest.webmanifest", "./icon-180.png", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", e => {
  // cache: "reload" skips the browser's own short-term cache, so a new version never saves an old copy
  e.waitUntil(caches.open(VERSION)
    .then(c => c.addAll(FILES.map(f => new Request(f, {cache: "reload"}))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

// App files: show the saved copy instantly, refresh it in the background when online.
// Requests to your Apps Script (your sheet data) are never touched here; the app saves those itself.
self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== self.location.origin) return;
  const key = req.mode === "navigate" ? "./index.html" : req;
  // the fresh copy is fetched in the background; waitUntil keeps the phone from cutting it off before it's saved
  const fresh = caches.open(VERSION).then(cache => fetch(req, {cache: "no-cache"}).then(res => { if (res && res.ok) return cache.put(key, res.clone()).then(() => res); return res; })).catch(() => null);
  e.waitUntil(fresh);
  e.respondWith(caches.open(VERSION).then(async cache => {
    const saved = await cache.match(key, {ignoreSearch: true});
    return saved || (await fresh) || new Response("Offline", {status: 503});
  }));
});
