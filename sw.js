// sw.js - Basic Caching Service Worker
// Timestamp: 2025-05-24T14:55:00EDT
// Summary: Improved offline refresh behavior by explicitly serving cached index.html for failed navigation fetches.

const CACHE_NAME = 'fsm-dashboard-cache-v1'; // Keep or update cache name as needed
// List of files to cache immediately upon installation (Relative Paths)
const urlsToCache = [
  '.',             // Represents index.html within /fsm_dashboard/
  'index.html',    // Relative to sw.js
  'style.css',     // Relative to sw.js
  'script.js',     // Relative to sw.js
  'manifest.json', // Relative to sw.js
  'icons/icon.svg',// Relative to sw.js (assumes icons folder inside fsm_dashboard)
  // External libraries remain absolute
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@latest/dist/chart.umd.js'
];

// Install event: Cache the core assets
self.addEventListener('install', event => {
  console.log('[Service Worker] Install');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Caching app shell');
        const promises = urlsToCache.map(url => {
            // For relative URLs, fetch respects the service worker's path
            return fetch(url, { cache: 'reload' }) // 'reload' bypasses HTTP cache for fresh copies
                .then(response => {
                    if (!response.ok) {
                        // Log error but don't let one failed resource stop others
                        console.error(`[Service Worker] Request failed for ${url} during cache: ${response.statusText}`);
                        return Promise.resolve(); 
                    }
                    // Check if response type is valid for caching
                    if(response.status === 200 || response.type === 'basic' || response.type === 'cors') {
                         return cache.put(url, response);
                    } else {
                         console.warn(`[Service Worker] Skipping caching for ${url} due to response status/type: ${response.status} / ${response.type}`);
                         return Promise.resolve();
                    }
                }).catch(fetchError => {
                    console.error(`[Service Worker] Fetch error for ${url} during cache:`, fetchError);
                    return Promise.resolve(); // Continue caching other files
                });
        });
        return Promise.all(promises);
      })
      .then(() => {
        console.log('[Service Worker] App shell caching attempted, skipping waiting.');
        return self.skipWaiting(); // Activate the new service worker immediately
      })
      .catch(error => {
         console.error('[Service Worker] Caching failed during install setup:', error);
      })
  );
});

// Activate event: Clean up old caches
self.addEventListener('activate', event => {
  console.log('[Service Worker] Activate');
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (!cacheWhitelist.includes(cacheName)) {
            console.log('[Service Worker] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
         console.log('[Service Worker] Claiming clients.');
         return self.clients.claim(); // Ensure the new service worker controls all clients
    })
  );
});

// Fetch event: Serve from cache first, fall back to network
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') { return; } // Only handle GET requests

  event.respondWith(
    caches.match(event.request)
      .then(cachedResponse => {
        // Cache hit - return response
        if (cachedResponse) {
          // console.log('[Service Worker] Serving from cache:', event.request.url);
          return cachedResponse;
        }

        // Not in cache - try network
        // console.log('[Service Worker] Not in cache, fetching from network:', event.request.url);
        return fetch(event.request).then(
          networkResponse => {
            // Optional: Cache dynamically fetched resources if needed
            // Example: Cache new successful GET requests from your origin
            // if (networkResponse && networkResponse.status === 200 && event.request.url.startsWith(self.location.origin)) {
            //   const responseToCache = networkResponse.clone();
            //   caches.open(CACHE_NAME)
            //     .then(cache => {
            //       console.log('[Service Worker] Caching new resource:', event.request.url);
            //       cache.put(event.request, responseToCache);
            //     });
            // }
            return networkResponse;
          }
        ).catch(error => {
          console.error('[Service Worker] Network fetch failed for:', event.request.url, error);
          
          // IMPORTANT: For navigation requests (like refreshing index.html), 
          // try to serve the cached index.html as a fallback.
          if (event.request.mode === 'navigate') {
            console.log('[Service Worker] Fetch failed for navigation, attempting to serve cached index.html');
            return caches.match('index.html'); // Or a specific offline.html page if you have one
          }
          
          // For other failed requests (assets etc.), they will result in a network error
          // if not found in cache and network fails. This is often the desired behavior
          // for non-critical assets or to indicate that a specific non-cached resource is unavailable.
          return undefined; 
        });
      })
  );
});
