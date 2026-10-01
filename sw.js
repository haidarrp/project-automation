const STATIC_CACHE = 'pusdatin-static-v11';

self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(
    keys.filter((key) => key.startsWith('pusdatin-static-') && key !== STATIC_CACHE).map((key) => caches.delete(key))
  )).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (!['script', 'style', 'image', 'font'].includes(request.destination)) return;

  // Berkas autentikasi/konfigurasi Microsoft tidak boleh dilayani dari cache lama.
  // Redirect bridge MSAL harus selalu konsisten dengan bundle yang sedang dideploy.
  if (url.pathname.includes('/assets/vendor/msal-') || url.pathname.endsWith('/assets/js/microsoft-config.js')) return;

  event.respondWith(caches.open(STATIC_CACHE).then(async (cache) => {
    const cached = await cache.match(request);
    const network = fetch(request).then((response) => {
      if (response && response.ok) cache.put(request, response.clone()).catch(() => {});
      return response;
    }).catch(() => cached);
    return cached || network;
  }));
});
