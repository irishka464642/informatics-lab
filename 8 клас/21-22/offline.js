/* Спільний індикатор офлайн-готовності для меню та двох ігор. */
(() => {
  'use strict';
  const status = document.getElementById('offlineState');
  if (!status) return;
  const swURL = new URL('sw.js', document.currentScript.src);
  const scopeURL = new URL('./', swURL);
  const files = [
    './', './index.html',
    './logic-true-false-8/', './logic-true-false-8/index.html',
    './logic-boolean-code-8/', './logic-boolean-code-8/index.html',
    './offline.js'
  ];
  let confirmed = false;

  function update(message, kind) {
    status.textContent = message;
    status.dataset.state = kind;
  }
  function show() {
    if (!confirmed) {
      update(navigator.onLine ?
        '⏳ Готуємо ігри для роботи без інтернету…' :
        '⚠️ Офлайн-кеш ще не готовий. Підключись до інтернету й відкрий меню ігор.', 'pending');
    } else if (navigator.onLine) {
      update('✅ Обидві гри збережено на телефоні. Можна працювати без інтернету.', 'ready');
    } else {
      update('📴 Офлайн-режим: гра працює. Прогрес зберігається на телефоні.', 'offline');
    }
  }
  async function readyInCache() {
    if (!('caches' in window)) return false;
    const names = await caches.keys();
    const relevant = names.filter(name => name.startsWith('agency-logic-8-offline-'));
    for (const name of relevant.reverse()) {
      const cache = await caches.open(name);
      const checks = await Promise.all(
        files.map(path => cache.match(new URL(path, scopeURL).href)));
      if (checks.every(Boolean)) return true;
    }
    return false;
  }
  async function init() {
    if (!('serviceWorker' in navigator) || !('caches' in window)) {
      update('⚠️ Цей браузер не підтримує надійне офлайн-кешування. Відкрий Chrome або Safari.', 'error');
      return;
    }
    show();
    try {
      const registration = await navigator.serviceWorker.register(swURL.href,
        {scope: scopeURL.href, updateViaCache: 'none'});
      // Service worker install waits until ALL game pages have been cached.
      await navigator.serviceWorker.ready;
      confirmed = await readyInCache();
      if (!confirmed && navigator.onLine) {
        try { await registration.update(); } catch {}
        confirmed = await readyInCache();
      }
      if (!confirmed) {
        update('⚠️ Не всі сторінки збережено. Підключись до інтернету та онови сторінку.', 'error');
      } else show();
    } catch (e) {
      update('⚠️ Не вдалося зберегти ігри для офлайн-роботи. Спробуй ще раз з інтернетом.', 'error');
    }
  }
  window.addEventListener('online', () => {
    show();
    if (!confirmed) init();
  });
  window.addEventListener('offline', show);
  init();
})();
