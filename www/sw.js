// Guarda la app en el dispositivo para que abra sin internet.
// Al publicar una versión nueva, sube el número de VERSION.
const VERSION = 'empaque-v8-47-0';
const FILES = [
  "./",
  "index.html",
  "app.css",
  "marco.css",
  "app.js",
  "manuales.js",
  "config.js",
  "fonts.css",
  "manifest.webmanifest",
  "vendor/exceljs.min.js",
  "vendor/pptxgen.min.js",
  "vendor/qrcode.min.js",
  "vendor/webp/webp_enc.js",
  "vendor/webp/webp_enc.wasm",
  "xl.js",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "img/logo-vl.jpg",
  "img/logo-ec.jpg",
  "img/logo-vlusa.jpg",
  "img/logo-jc.jpg",
  "img/ico-vl.png",
  "img/ico-ec.png",
  "img/lemon-bg.jpg",
  "img/trailer.jpg",
  "img/tractor.jpg",
  "audio/inicio-sesion.mp3",
  "fonts/barlow-condensed-latin-500-normal.woff2",
  "fonts/barlow-condensed-latin-600-normal.woff2",
  "fonts/barlow-condensed-latin-700-normal.woff2",
  "fonts/barlow-latin-400-normal.woff2",
  "fonts/barlow-latin-500-normal.woff2",
  "fonts/barlow-latin-600-normal.woff2",
  "fonts/ibm-plex-mono-latin-400-normal.woff2",
  "fonts/ibm-plex-mono-latin-500-normal.woff2",
  "fonts/ibm-plex-mono-latin-600-normal.woff2",
  "fonts/figtree-latin-400-normal.woff2",
  "fonts/figtree-latin-500-normal.woff2",
  "fonts/figtree-latin-600-normal.woff2",
  "fonts/figtree-latin-700-normal.woff2"
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  if (new URL(r.url).origin !== location.origin) return; // la base de datos nunca se guarda aquí
  e.respondWith(caches.match(r).then(hit => {
    const red = fetch(r).then(res => {
      if (res && res.ok) { const copia = res.clone(); caches.open(VERSION).then(c => c.put(r, copia)); }
      return res;
    }).catch(() => hit);
    return hit || red;
  }));
});
