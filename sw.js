const CACHE_NOME = "philo-propostas-v2";
const ARQUIVOS_SHELL = [
  "./",
  "./index.html",
  "./app.js",
  "./calc-engine.js",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./logo-header.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NOME).then((cache) => cache.addAll(ARQUIVOS_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((chaves) =>
      Promise.all(chaves.filter((c) => c !== CACHE_NOME).map((c) => caches.delete(c)))
    )
  );
  self.clients.claim();
});

// Estratégia: shell do app (HTML/JS/CSS/ícones) via cache-first;
// chamadas de API (NASA POWER, Nominatim, Apps Script) sempre direto da rede,
// porque são dados que mudam e não fazem sentido cacheados.
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  const ehApiExterna =
    url.hostname.includes("power.larc.nasa.gov") ||
    url.hostname.includes("nominatim.openstreetmap.org") ||
    url.hostname.includes("script.google.com");

  if (ehApiExterna) return; // deixa passar direto pra rede

  event.respondWith(
    caches.match(event.request).then((resp) => resp || fetch(event.request))
  );
});
