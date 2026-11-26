/* Облік активного часу роботи у двох іграх. Тільки локально, без мережі. */
(() => {
  'use strict';
  const game = document.currentScript?.dataset.game;
  if (!['theory','practice'].includes(game)) return;
  const KEY = 'logic-grade8-playtime-v1';
  let last = Date.now();
  const read = () => {
    try {
      const v=JSON.parse(localStorage.getItem(KEY)||'{}');
      return v && typeof v === 'object' ? v : {};
    } catch { return {}; }
  };
  function record() {
    const now=Date.now();
    const elapsed=Math.max(0,now-last);
    last=now;
    if (document.hidden || elapsed<1000) return;
    // Враховуємо лише час активного перегляду гри, без довгого простою у фоні.
    const seconds=Math.min(30,Math.floor(elapsed/1000));
    if (seconds<=0) return;
    try {
      const state=read();
      state[game]=Math.max(0,Math.trunc(Number(state[game])||0))+seconds;
      localStorage.setItem(KEY,JSON.stringify(state));
    } catch {}
  }
  // При приховуванні сторінки спочатку фіксуємо останні секунди.
  document.addEventListener('visibilitychange',() => {
    if (document.hidden) {
      const hidden=document.hidden;
      const current=Date.now();
      const seconds=Math.min(30,Math.floor(Math.max(0,current-last)/1000));
      last=current;
      if (hidden && seconds>0) {
        try {
          const state=read();
          state[game]=Math.max(0,Math.trunc(Number(state[game])||0))+seconds;
          localStorage.setItem(KEY,JSON.stringify(state));
        }catch{}
      }
    } else last=Date.now();
  });
  window.addEventListener('pagehide',record);
  window.addEventListener('pageshow',() => {last=Date.now();});
  setInterval(record,10000);
})();