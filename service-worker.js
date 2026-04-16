const CACHE_NAME = 'terapia-pwa-v1002';
const APP_ASSETS = [
  './',
  './index.html',
  './offline.html',
  './style.css',
  './app.js',
  './manifest.webmanifest',
  './icons/favicon-v2.svg',
  './icons/favicon-16-v2.png',
  './icons/favicon-32-v2.png',
  './icons/favicon-48-v2.png',
  './icons/favicon-192-v2.png',
  './icons/icon-192-v2.png',
  './icons/icon-512-v2.png',
  './icons/apple-touch-icon.png',
  './screenshots/mobile.png',
  './screenshots/desktop.png'
];
const APP_SHELL_PATHS = new Set(APP_ASSETS.map((asset) => new URL(asset, self.location.href).pathname));

async function networkFirst(request) {
  const cache = await caches.open(CACHE_NAME);

  try {
    const networkResponse = await fetch(request);

    if (networkResponse && networkResponse.ok) {
      cache.put(request, networkResponse.clone());
    }

    return networkResponse;
  } catch (error) {
    const cachedResponse = await caches.match(request);
    if (cachedResponse) {
      return cachedResponse;
    }

    if (request.mode === 'navigate') {
      return caches.match('./offline.html');
    }

    return Response.error();
  }
}

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') {
    return;
  }

  const requestUrl = new URL(event.request.url);
  const isSameOrigin = requestUrl.origin === self.location.origin;
  const isNavigate = event.request.mode === 'navigate';
  const isAppShell = APP_SHELL_PATHS.has(requestUrl.pathname);

  // Cache solo asset interni o navigazione
  if (isSameOrigin && (isNavigate || isAppShell)) {
    event.respondWith(networkFirst(event.request));
  }
  // Tutte le altre richieste (Firebase, Font, CDN) vanno dirette in rete 
  // e NON vengono salvate dal service worker (hanno già la loro persistenza)
});

// Listener per messaggi dall'app (es. pulizia forzata)
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  
  if (event.data && event.data.type === 'CLEAR_CACHE') {
    event.waitUntil(
      caches.keys().then((keys) => 
        Promise.all(keys.map((key) => caches.delete(key)))
      )
    );
  }
});
