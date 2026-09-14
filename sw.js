const CACHE_NOME = "philo-propostas-v4";
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
    ).then(() => self.clients.claim())
  );
});

// Estratégia: NETWORK-FIRST para o shell do app (HTML/JS) — sempre busca a
// versão mais nova da rede primeiro, e só usa o cache como fallback se
// estiver offline. Isso evita ficar preso em versão antiga do código.
// Chamadas de API (NASA POWER, Nominatim, Apps Script) sempre direto da rede,
// sem cache, porque são dados que mudam.
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  const ehApiExterna =
    url.hostname.includes("power.larc.nasa.gov") ||
    url.hostname.includes("nominatim.openstreetmap.org") ||
    url.hostname.includes("script.google.com");

  if (ehApiExterna) return; // deixa passar direto pra rede

  event.respondWith(
    fetch(event.request)
      .then((resp) => {
        const copia = resp.clone();
        caches.open(CACHE_NOME).then((cache) => cache.put(event.request, copia));
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
