/* Service worker mínimo: hace la app instalable y guarda la carcasa en caché.
   La app real (Apps Script) necesita conexión, no se cachea. */
const CACHE = "entrenos-v1";
const ASSETS = [
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(ASSETS); }));
  self.skipWaiting();
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; })
                            .map(function (k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (e) {
  const url = new URL(e.request.url);
  // Solo servimos desde caché los archivos propios de la carcasa (mismo origen).
  if (url.origin === self.location.origin) {
    e.respondWith(
      caches.match(e.request).then(function (r) { return r || fetch(e.request); })
    );
  }
  // Todo lo demás (Apps Script, Google, Chart.js…) va directo a la red.
});
