const CACHE_NAME = 'svaadh-cache-v1';
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

// A simple network-first strategy, falling back to cache
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  
  event.respondWith(
    fetch(event.request)
      .catch(() => caches.match(event.request))
  );
});
