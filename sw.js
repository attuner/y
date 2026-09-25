const CACHE_NAME = "yaadys-cache-v3";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "https://cdn.tailwindcss.com",
  "https://unpkg.com/lucide@latest",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
  "https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap"
];

// Install Event: Cache App Shell
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(APP_SHELL);
    }).then(() => self.skipWaiting())
  );
});

// Activate Event: Evict Old Caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch Event: Bypass API calls, Stale-While-Revalidate for UI Assets
self.addEventListener("fetch", (event) => {
  const requestUrl = new URL(event.request.url);
  
  // Skip caching non-GET requests (e.g. POST to Apps Script)
  if (event.request.method !== "GET") {
    return;
  }
  
  // Google Apps Script API calls: Always bypass cache directly to network
  if (
    requestUrl.hostname.includes("script.google.com") ||
    requestUrl.hostname.includes("script.googleusercontent.com")
  ) {
    event.respondWith(fetch(event.request));
    return;
  }
  
  // Static Assets & Shell: Stale-While-Revalidate with status checks
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const fetchPromise = fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === "basic") {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(() => cachedResponse);
      
      return cachedResponse || fetchPromise;
    })
  );
});