/* Агенція логіки — офлайн-кеш для уроків 21–22, 8 клас.
   Прогрес і персональні дані НЕ зберігаються в цьому кеші:
   вони залишаються тільки в localStorage браузера учня. */
const VERSION = 'agency-logic-8-offline-v20261127-4';
const FILES = [
  './',
  './index.html',
  './logic-true-false-8/',
  './logic-true-false-8/index.html',
  './logic-boolean-code-8/',
  './logic-boolean-code-8/index.html',
  './progress/',
  './progress/index.html',
  './offline.js',
  './playtime.js'
];
const resources = FILES.map(p => new URL(p, self.registration.scope).href);

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(VERSION);
    await cache.addAll(resources);
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names
      .filter(name => name.startsWith('agency-logic-8-offline-') && name !== VERSION)
      .map(name => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Ніколи не кешувати Google Apps Script, результати чи зовнішні ресурси.
  if (url.origin !== self.location.origin ||
      !url.pathname.startsWith(new URL(self.registration.scope).pathname)) return;

  // Знаходимо сторінку з кешу за точною адресою, інакше пробуємо інтернет.
  // Працює без мережі після одноразового встановлення всього набору.
  event.respondWith((async () => {
    const cache = await caches.open(VERSION);
    const exact = await cache.match(req, {ignoreSearch: true});
    if (exact) return exact;
    try { return await fetch(req); }
    catch (e) {
      if (req.mode === 'navigate') {
        const path = url.pathname.replace(/\/index\.html$/, '/');
        const fallback = await cache.match(url.origin + path);
        if (fallback) return fallback;
      }
      throw e;
    }
  })());
});
