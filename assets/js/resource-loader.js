(function () {
  'use strict';

  const RESOURCES = Object.freeze({
    xlsx: {
      src: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
      ready: () => Boolean(window.XLSX)
    },
    exceljs: {
      src: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
      ready: () => Boolean(window.ExcelJS)
    },
    jszip: {
      src: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
      ready: () => Boolean(window.JSZip)
    },
    filesaver: {
      src: 'https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js',
      ready: () => typeof window.saveAs === 'function'
    }
  });

  const pending = new Map();

  function load(name) {
    const resource = RESOURCES[name];
    if (!resource) return Promise.reject(new Error(`Resource tidak dikenal: ${name}`));
    if (resource.ready()) return Promise.resolve();
    if (pending.has(name)) return pending.get(name);

    const promise = new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[data-lazy-resource="${name}"]`);
      if (existing) {
        existing.addEventListener('load', () => resource.ready() ? resolve() : reject(new Error(`Resource ${name} gagal diinisialisasi.`)), { once: true });
        existing.addEventListener('error', () => reject(new Error(`Resource ${name} gagal dimuat.`)), { once: true });
        return;
      }

      const script = document.createElement('script');
      script.src = resource.src;
      script.async = true;
      script.dataset.lazyResource = name;
      script.crossOrigin = 'anonymous';
      script.addEventListener('load', () => {
        if (resource.ready()) resolve();
        else reject(new Error(`Resource ${name} selesai dimuat tetapi tidak tersedia.`));
      }, { once: true });
      script.addEventListener('error', () => reject(new Error(`Resource ${name} gagal dimuat. Periksa koneksi internet.`)), { once: true });
      document.head.appendChild(script);
    }).finally(() => {
      if (!resource.ready()) pending.delete(name);
    });

    pending.set(name, promise);
    return promise;
  }

  async function ensure(names) {
    const list = Array.isArray(names) ? names : [names];
    await Promise.all(list.filter(Boolean).map(load));
  }

  function warm(names) {
    const list = Array.isArray(names) ? names : [names];
    const schedule = window.requestIdleCallback || ((callback) => setTimeout(callback, 50));
    schedule(() => {
      list.filter(Boolean).forEach((name) => load(name).catch(() => {}));
    }, { timeout: 1500 });
  }

  window.ResourceLoader = Object.freeze({ ensure, warm, load });
})();
