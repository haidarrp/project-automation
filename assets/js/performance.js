(function () {
  'use strict';
  if (!('serviceWorker' in navigator) || !/^https?:$/.test(location.protocol)) return;
  const register = () => navigator.serviceWorker.register('./sw.js').catch((error) => console.warn('Service worker tidak dapat didaftarkan:', error));
  const schedule = () => {
    if (typeof window.requestIdleCallback === 'function') window.requestIdleCallback(register, { timeout: 1500 });
    else setTimeout(register, 1000);
  };
  if (document.readyState === 'complete') schedule();
  else window.addEventListener('load', schedule, { once: true });
})();
