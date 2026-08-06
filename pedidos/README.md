# Gestor de Pedidos

App para tomar pedidos tienda por tienda y saber, al final del día, **cuánto hay
que subir al camión de cada producto**. Funciona en el teléfono, sin internet, y
genera facturas diarias y proyecciones semanales y mensuales.

No necesita servidor, ni cuenta, ni instalación de programas: es una página web
que el teléfono guarda como si fuera una app.

---

## Qué hace

| Pantalla | Para qué sirve |
|---|---|
| **Pedidos** | Registrar un pedido en la puerta de la tienda: nombre del cliente y cantidades. Ves el total del día. |
| **Cargue** | La respuesta a "¿qué subo al camión?": suma todos los pedidos y te da el total por producto, agrupado por categoría, con casillas para ir marcando lo que ya cargaste. |
| **Reportes** | Cuánto vendiste por día, qué productos se mueven más, qué cliente compra más, y la **proyección** de la próxima semana y el próximo mes. Exporta a Excel (CSV). |
| **Catálogo** | Tus productos (nombre, categoría, precio, unidad) y tu libreta de clientes. |
| **Ajustes** | Nombre del negocio para la factura, moneda y respaldos. |

Cada pedido genera una **factura** que se puede compartir por WhatsApp o imprimir
(desde ahí también se guarda como PDF).

---

## Cómo ponerla en el teléfono

### Opción A — publicarla con GitHub Pages (recomendada)

Así queda en una dirección web fija, y se instala como app.

1. En GitHub, entra al repositorio → **Settings** → **Pages**.
2. En *Build and deployment* → *Source*, elige **Deploy from a branch**.
3. Escoge la rama (`main` una vez esté integrado) y la carpeta `/ (root)`. Guarda.
4. A los pocos minutos la app queda en:
   `https://<tu-usuario>.github.io/<repositorio>/pedidos/`
5. Abre esa dirección **en el teléfono** y agrégala a la pantalla de inicio:
   - **Android (Chrome):** menú ⋮ → *Agregar a pantalla principal*.
   - **iPhone (Safari):** botón Compartir → *Agregar a inicio*.

Desde ese momento abre como una app normal, a pantalla completa y **sin
internet**: la primera visita deja todo guardado en el teléfono.

### Opción B — sin publicar nada

Copia la carpeta `pedidos/` al teléfono o al computador y abre `index.html` con
el navegador. Todo funciona igual salvo el modo sin conexión instalado (que
requiere que la página venga de una dirección `https://`).

---

## Primeros pasos

1. Abre **Catálogo** → *Cargar catálogo de ejemplo* (papas, cheetos, palomitas,
   gaseosas…) y edítalo: cambia nombres y precios por los tuyos, borra lo que no
   vendas y agrega lo que falte.
2. En **Ajustes**, escribe el nombre de tu negocio y tu teléfono: eso sale en la
   factura.
3. Ya puedes tomar pedidos. En cada tienda: **+ Nuevo pedido** → escribe el
   nombre de la tienda → busca cada producto y toca **+** las veces que pidan →
   **Guardar**. La tienda queda guardada en la libreta para la próxima visita.
4. Al terminar el recorrido, entra a **Cargue**: ahí está la lista de qué subir
   al camión. Puedes compartirla por WhatsApp o imprimirla.

---

## Dónde quedan los datos (importante)

Todo se guarda **únicamente en el teléfono donde lo usas** (almacenamiento del
navegador). No se sube a ningún lado: nadie más ve tus ventas, y por eso mismo
**nadie más las respalda**.

Si borras los datos del navegador, desinstalas la app o cambias de teléfono, se
pierde todo. Por eso:

> **Ajustes → Exportar respaldo** de vez en cuando (por ejemplo, cada fin de
> semana). Descarga un archivo `.json` que puedes mandarte por correo o guardar
> en Drive. Para recuperarlo: **Ajustes → Importar respaldo**.

Si quieres llevar la contabilidad en Excel, usa **Reportes → Exportar CSV**.

---

## Cómo se calcula la proyección

Toma los últimos días de historial (14, 28 o 56, tú eliges), suma lo vendido y
lo divide entre los días transcurridos desde el primer pedido de esa ventana.
Ese promedio diario se multiplica por 7 (semana) y por 30 (mes).

Es un promedio, no una bola de cristal: **no** sabe de festivos, quincenas ni
temporadas. Con menos de una semana de historial la app te lo advierte, porque
esos números todavía no significan gran cosa. A partir de dos o tres semanas
registrando pedidos empieza a ser útil.

---

## Detalles técnicos

- HTML, CSS y JavaScript puros. Sin dependencias, sin compilación.
- `index.html` · `app.css` · `app.js` — la aplicación.
- `sw.js` — service worker (modo sin conexión). Al cambiar la app, sube el
  número de `VERSION` para que los teléfonos descarguen la versión nueva.
- `manifest.webmanifest`, `icon*.png/svg` — para que se instale como app.
- Los datos viven en `localStorage`, bajo la clave `gestor_pedidos_v1`.
