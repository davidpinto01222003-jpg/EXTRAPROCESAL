/* Service worker: es lo que hace que esto se pueda INSTALAR como app y
   que abra al instante aunque el PC este apagado.

   Regla: la cascara de la app (HTML, CSS, JS, iconos) se guarda en el
   telefono; los DATOS (todo lo que empieza por /api/) nunca se guardan,
   porque mostrar postulaciones viejas como si fueran de ahora seria
   peor que no mostrar nada. */

const CACHE = "empleo-v1";
const CASCARA = [
  "/",
  "/estatico/estilo.css",
  "/estatico/app.js",
  "/estatico/icono-192.png",
  "/estatico/icono-512.png",
  "/manifest.json",
];

self.addEventListener("install", (evento) => {
  evento.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(CASCARA))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (evento) => {
  evento.waitUntil(
    caches.keys()
      .then((nombres) => Promise.all(
        nombres.filter((n) => n !== CACHE).map((n) => caches.delete(n))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (evento) => {
  const url = new URL(evento.request.url);

  if (evento.request.method !== "GET" || url.origin !== self.location.origin) return;

  // Los datos siempre frescos del PC. Si no hay PC, que falle y la app
  // lo diga ("PC apagado"), en vez de mentir con datos viejos.
  if (url.pathname.startsWith("/api/")) return;

  evento.respondWith(
    caches.match(evento.request, { ignoreSearch: true }).then((guardado) => {
      const red = fetch(evento.request).then((respuesta) => {
        if (respuesta && respuesta.status === 200) {
          const copia = respuesta.clone();
          caches.open(CACHE).then((cache) => cache.put(evento.request, copia));
        }
        return respuesta;
      });
      return guardado || red;
    }).catch(() => caches.match("/"))
  );
});
