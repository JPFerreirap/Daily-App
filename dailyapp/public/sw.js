// Service worker de Daily App.
// Solo guarda en caché los archivos de la propia app (HTML, JS, CSS, íconos)
// para que abra rápido y pueda abrir sin conexión. NUNCA toca las llamadas a
// Supabase ni a otras APIs (son de otro origen): los datos siempre vienen de
// la nube, así no se muestran datos viejos ni se interfiere con el guardado.
const CACHE = "daily-app-v1";
const ARCHIVOS_BASE = ["/", "/manifest.webmanifest", "/icon-192.png", "/icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARCHIVOS_BASE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((claves) => Promise.all(claves.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // Supabase, fuentes, CDN: directo a la red

  // Páginas: primero la red (siempre la versión más nueva), y si no hay
  // conexión, la última copia guardada.
  if (req.mode === "navigate") {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copia = res.clone();
          caches.open(CACHE).then((c) => c.put("/", copia));
          return res;
        })
        .catch(() => caches.match("/"))
    );
    return;
  }

  // Archivos estáticos (nombres con hash de Vite): primero caché, y se van
  // guardando a medida que se usan.
  event.respondWith(
    caches.match(req).then(
      (enCache) =>
        enCache ||
        fetch(req).then((res) => {
          if (res.ok) {
            const copia = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copia));
          }
          return res;
        })
    )
  );
});
