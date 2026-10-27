/* AgendaGuru service worker: membuat aplikasi bisa dipasang dan tetap terbuka saat sinyal lemah.
 * - Halaman & file aplikasi: ambil versi terbaru dari internet, pakai salinan tersimpan bila offline.
 * - Pustaka CDN (Bootstrap, ikon, grafik): pakai salinan tersimpan agar cepat.
 * - Panggilan ke Apps Script (data) tidak pernah disimpan di sini. */
const CACHE = 'agendaguru-v1';
const SHELL = ['./', './index.html', './api.js', './app.js', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
    e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => { }));
    self.skipWaiting();
});

self.addEventListener('activate', e => {
    e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
        .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
    const req = e.request;
    if (req.method !== 'GET') return;
    const url = new URL(req.url);
    if (/script\.google(usercontent)?\.com$/.test(url.hostname)) return;

    if (url.origin === self.location.origin) {
        // Network-first; file dengan ?v=... disimpan tanpa query agar tetap tersedia offline
        const key = req.mode === 'navigate' ? './index.html' : url.pathname;
        e.respondWith(fetch(req).then(res => {
            if (res && res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(key, copy)); }
            return res;
        }).catch(() => caches.match(key).then(r => r || caches.match(req))));
        return;
    }

    if (/^(cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.(googleapis|gstatic)\.com|unpkg\.com)$/.test(url.hostname)) {
        e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
            if (res && (res.ok || res.type === 'opaque')) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
            return res;
        })));
    }
});
