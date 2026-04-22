const VERSION = "budget-card-view-v2";
const SHELL_CACHE = `${VERSION}-shell`;
const RUNTIME_CACHE = `${VERSION}-runtime`;
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./vendor/capacitor.js",
  "./vendor/synapse.js",
  "./vendor/capacitor-filesystem.js",
  "./vendor/capacitor-share.js",
  "./vendor/xlsx.full.min.js",
  "./manifest.json",
  "./offline.html",
  "./apple-touch-icon.png",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-app.svg",
  "./icon-192.svg",
  "./icon-512.svg",
];
const RUNTIME_ASSETS = [];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(SHELL_CACHE);
    await cache.addAll(APP_SHELL);

    for (const url of RUNTIME_ASSETS) {
      try {
        await cache.add(new Request(url, { mode: "no-cors" }));
      } catch (error) {
        console.error("Runtime cache warmup failed:", error);
      }
    }

    await self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const cacheNames = await caches.keys();

    await Promise.all(
      cacheNames
        .filter((cacheName) => ![SHELL_CACHE, RUNTIME_CACHE].includes(cacheName))
        .map((cacheName) => caches.delete(cacheName))
    );

    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const { request } = event;

  if (request.method !== "GET") {
    return;
  }

  if (request.mode === "navigate") {
    event.respondWith(handleNavigation(request));
    return;
  }

  const url = new URL(request.url);
  if (url.origin === self.location.origin || RUNTIME_ASSETS.includes(url.href)) {
    event.respondWith(staleWhileRevalidate(request));
  }
});

async function handleNavigation(request) {
  try {
    const response = await fetch(request);
    const cache = await caches.open(RUNTIME_CACHE);
    await cache.put("./index.html", response.clone());
    return response;
  } catch (error) {
    return (
      (await caches.match(request, { ignoreSearch: true })) ||
      (await caches.match("./index.html")) ||
      (await caches.match("./offline.html"))
    );
  }
}

async function staleWhileRevalidate(request) {
  const cachedResponse = await caches.match(request, { ignoreSearch: true });

  const networkResponsePromise = fetch(request)
    .then(async (response) => {
      if (response && (response.ok || response.type === "opaque")) {
        const cache = await caches.open(RUNTIME_CACHE);
        await cache.put(request, response.clone());
      }

      return response;
    })
    .catch(() => null);

  return cachedResponse || (await networkResponsePromise) || Response.error();
}
