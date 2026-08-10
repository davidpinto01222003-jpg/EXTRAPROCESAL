/* Service worker: guarda la cascara de la app (HTML, CSS, JS, iconos) en
   el telefono para que abra al instante. Los DATOS (todo lo que empieza
   por /api/) nunca se guardan: mostrar postulaciones viejas como si
   fueran de ahora seria peor que no mostrar nada.

   OJO -- CUANDO SI Y CUANDO NO SE USA ESTO
   ----------------------------------------
   Los navegadores solo permiten service workers en conexiones que
   consideran seguras (https, o localhost). Como esta app se sirve por
   http en tu red local (http://192.168.x.x:8777), el celular
   sencillamente NO lo registra -- `navigator.serviceWorker` ni siquiera
   existe ahi, y app.js ya lo tiene previsto.

   No se pierde nada importante: la app igual se agrega a la pantalla de
   inicio y funciona igual, y como TODO lo que muestra viene del PC, sin
   el PC prendido no habria nada que ver de todos modos.

   Este archivo si trabaja cuando abres la app en el propio PC
   (http://127.0.0.1:8777), que cuenta como conexion segura. */

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
