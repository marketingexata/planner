/**
 * REELSIGHT — Service Worker
 * Membuat aplikasi bisa di-install (PWA) dan tetap bisa dibuka saat offline.
 * Data insight tetap tersimpan lewat localStorage (lihat app.js), bukan lewat SW ini.
 *
 * PENTING: setiap kali file app di-update, naikkan angka CACHE_VERSION di bawah
 * supaya pengguna lama mendapat versi terbaru (bukan versi cache basi).
 */

const CACHE_VERSION = "v4";
const CACHE_NAME = `reelsight-${CACHE_VERSION}`;

// File "inti" aplikasi — wajib bisa dibuka walau offline
const APP_SHELL = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./config.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./icons/apple-touch-icon.png"
];

/* ---------- INSTALL: precache app shell ---------- */
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

/* ---------- ACTIVATE: buang cache versi lama ---------- */
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key.startsWith("reelsight-") && key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

/* ---------- FETCH: strategi berbeda tergantung jenis request ---------- */
self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Jangan pernah cache panggilan ke backend Apps Script (data harus selalu fresh / kirim real-time)
  if (url.hostname.includes("script.google.com")) {
    event.respondWith(fetch(req).catch(() => cachedFallback(req)));
    return;
  }

  // Hanya tangani GET; biarkan POST/PUT dsb lewat langsung ke jaringan
  if (req.method !== "GET") return;

  // Navigasi halaman (buka app / refresh) -> network-first, fallback ke cache saat offline
  if (req.mode === "navigate") {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put("./index.html", clone));
          return res;
        })
        .catch(() => caches.match("./index.html"))
    );
    return;
  }

  // Aset sendiri (css/js/icon/manifest) -> cache-first, lalu update cache di belakang layar
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => {
            const clone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
            return res;
          })
          .catch(() => cached);
        return cached || network;
      })
    );
    return;
  }

  // Sumber eksternal (Tesseract.js CDN, Google Fonts) -> stale-while-revalidate
  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          return res;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});

function cachedFallback(req) {
  return caches.match(req).then((cached) => {
    if (cached) return cached;
    return new Response(
      JSON.stringify({ ok: false, error: "Sedang offline — coba lagi saat tersambung internet." }),
      { headers: { "Content-Type": "application/json" } }
    );
  });
}
