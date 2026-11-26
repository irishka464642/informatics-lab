/* Спільний індикатор офлайн-готовності для меню, двох ігор і картки гравця. */
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
    './progress/', './progress/index.html',
    './offline.js', './playtime.js'
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
      update('✅ Обидві гри та картка прогресу збережені на телефоні. Можна працювати без інтернету.', 'ready');
    } else {
      update('📴 Офлайн-режим: ігри й картка доступні. Прогрес зберігається на телефоні.', 'offline');
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
      // Якщо інтернет зник, не намагаємося повторно завантажити sw.js:
      // використовуємо вже встановлений service worker з телефону.
      let registration = await navigator.serviceWorker.getRegistration(scopeURL.href);
      // Інші ігри цього репозиторію можуть мати власний SW з ширшою областю.
      // Нам потрібна реєстрація саме для цього офлайн-комплекту.
      if (registration && registration.scope !== scopeURL.href) registration = null;
      if (!registration) {
        if (!navigator.onLine) {
          update('⚠️ Спочатку відкрий меню один раз з інтернетом і дочекайся ✅.', 'error');
          return;
        }
        registration = await navigator.serviceWorker.register(swURL.href,
          {scope: scopeURL.href, updateViaCache: 'none'});
      } else if (navigator.onLine) {
        // Перевірка оновлення не перешкоджає роботі старого офлайн-кешу.
        registration.update().catch(() => {});
      }
      // Service worker install waits until ALL game pages have been cached.
      await navigator.serviceWorker.ready;
      confirmed = await readyInCache();
      if (!confirmed && navigator.onLine) {
        try { await registration.update(); } catch {}
        // Оновлення SW і завантаження всіх сторінок асинхронні.
        // Зелений статус показуємо лише коли весь комплект є у кеші.
        for (let i=0; i<24 && !confirmed; i++) {
          confirmed = await readyInCache();
          if (!confirmed) await new Promise(resolve => setTimeout(resolve, 500));
        }
      }
      if (!confirmed) {
        update('⚠️ Картка або гра ще не збережена. Підключись до інтернету, онови меню й дочекайся зеленого повідомлення.', 'error');
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
