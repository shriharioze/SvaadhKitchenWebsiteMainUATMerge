const CACHE_NAME = 'svaadh-cache-v2';
const urlsToCache = [
  '/',
  '/index.html',
  '/order.html',
  '/favicon.svg',
  '/images/svaadh-kitchen-hero.PNG',
  '/i18n.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        return cache.addAll(urlsToCache);
      })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(name => name !== CACHE_NAME).map(name => caches.delete(name))
      );
    })
  );
  self.clients.claim();
});

// Network-first strategy — only for same-origin static assets.
// NEVER intercept API calls (script.google.com) — causes "null response" failures.
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  
  const url = event.request.url;
  // Skip all cross-origin / API requests — let browser handle them natively
  if (url.includes('script.google.com') || url.includes('macros/s/') 
      || !url.startsWith(self.location.origin)) {
    return; // don't call event.respondWith — browser handles it
  }
  
  event.respondWith(
    fetch(event.request)
      .catch(() => caches.match(event.request))
      .then(response => response || fetch(event.request))
  );
});
