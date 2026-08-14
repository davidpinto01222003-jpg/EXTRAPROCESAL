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
| **Pedidos** | Registrar un pedido en la puerta de la tienda, y ver la relación de pedidos por **día, semana, mes o año**. |
| **Cargue** | La respuesta a "¿qué subo al camión?": suma todos los pedidos y te da el total por producto, agrupado por categoría, con casillas para ir marcando lo que ya cargaste. |
| **Reportes** | Cuánto vendiste por día, qué productos se mueven más, qué cliente compra más, y la **proyección** de la próxima semana y el próximo mes. Exporta a Excel (CSV). |
| **Catálogo** | Tus productos (nombre, categoría, precio, unidad) y tus tiendas, cada una con **su historial de facturas**. |
| **Ajustes** | Nombre y foto del negocio, moneda y respaldos. |

### Tu empresa

La primera vez que abres la app te pide el **nombre de la empresa** y una **foto
de perfil** (tu logo o una foto). A partir de ahí:

- se ven al abrir la app, en la pantalla de entrada;
- el logo queda en la esquina de la barra superior (tócalo para ir a Ajustes);
- **encabezan todas las facturas y la cuenta del día**.

Se cambian cuando quieras en **Ajustes → Mi negocio**. La foto se achica sola
antes de guardarse, para no llenar la memoria del teléfono.

### Día, semana, mes y año

Tanto **Pedidos** como **Reportes** tienen el mismo selector arriba: *Día ·
Semana · Mes · Año*, con flechas **‹ ›** para moverte al periodo anterior o
siguiente, y un botón *Hoy* para volver. Así puedes mirar la semana pasada, un
mes de hace medio año o el año completo.

Lo que ves en cada uno, en Pedidos:

- **Día** — la lista de pedidos de esa jornada, uno por uno.
- **Semana** y **Mes** — los pedidos agrupados por día, con el subtotal de cada
  día en el encabezado. Toca el encabezado de un día para abrir ese día.
- **Año** — mes por mes, con pedidos, unidades y valor de cada uno. Toca un mes
  para abrirlo.

Arriba de todo, cuatro cifras del periodo: pedidos (y cuántas tiendas),
unidades, valor y **cuánto queda por cobrar**.

En Reportes, el gráfico se adapta: barras por día en semana y mes, **barras por
mes** cuando miras un año.

### Los tres documentos

1. **Factura de cada pedido** — se guarda individualmente, con su número
   consecutivo, cliente, productos y total. Se comparte por WhatsApp o se
   imprime (desde ahí sale el PDF).
2. **Cuenta del día / relación del periodo** — el cierre, con el botón azul de
   la pantalla de Pedidos. Sale del periodo que estés viendo:
   - en **Día**, la *cuenta del día*: cada pedido con su número y su estado;
   - en **Semana**, **Mes** o **Año**, la *relación*: cuánto compró cada tienda
     y cuánto te debe, la evolución día a día (o mes a mes en el año) y el
     detalle por producto.

   Todos terminan con el corte de caja: cobrado contra por cobrar. Se comparten
   por WhatsApp o se imprimen en PDF.
3. **Historial por tienda** — en *Catálogo → Clientes*, toca una tienda: total
   comprado, cuánto te debe, lo que más lleva, y **la lista de todas sus
   facturas**, cada una abrible. Desde ahí puedes tomarle un pedido nuevo con el
   nombre ya puesto, mandarle su estado de cuenta por WhatsApp o imprimir todas
   sus facturas juntas.

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

1. Al abrirla por primera vez, pon el **nombre de tu empresa** y tu **foto**.
   Eso es lo que verán tus clientes en las facturas.
2. Carga tus productos, por cualquiera de estas vías (**Catálogo → Productos**):
   - *Importar lista* — pega tu lista o abre un `.txt` / `.csv`. Ver abajo.
   - *Cargar catálogo de ejemplo* — papas, cheetos, palomitas, gaseosas… y lo
     editas: cambias nombres y precios, borras lo que no vendas.
   - *+ Nuevo producto* — uno por uno.
3. Ya puedes tomar pedidos. En cada tienda: **+ Nuevo pedido** → escribe el
   nombre de la tienda → busca cada producto y toca **+** las veces que pidan →
   **Guardar**. La tienda queda guardada en la libreta para la próxima visita.
4. Al terminar el recorrido, entra a **Cargue**: ahí está la lista de qué subir
   al camión. Puedes compartirla por WhatsApp o imprimirla.
5. Al cerrar el día, **Pedidos → Cuenta del día**: el resumen de la jornada,
   listo para compartir o guardar en PDF.

---

## Importar tu lista de productos

**Catálogo → Productos → Importar lista.** Pega el texto o abre un archivo
`.txt` o `.csv`. Antes de importar nada te muestra una tabla — *"Así lo
entendí"* — para que revises que quedó bien.

Entiende la lista escrita como uno la escribiría a mano. Los renglones en
MAYÚSCULAS o terminados en `:` los toma como categoría de lo que viene debajo:

```
PAPAS
Papas de mayonesa 1800
Papas limón 1.800

GASEOSAS:
Gaseosa personal 2.500
Agua 600 ml 1500
```

También lee tablas de Excel guardadas como CSV, con o sin fila de títulos:

```
Producto;Categoria;Precio;Unidad
Cheetos;Snacks;1.500;paquete
Doritos;Snacks;2.000;paquete
```

Detalles que resuelve solo:

- **Precios en formato colombiano**: `1.800`, `1,800`, `$ 1.800` y `1800` son
  todos mil ochocientos. `12.500,50` son doce mil quinientos con cincuenta.
- **Separadores**: punto y coma, tabulación o coma seguida de espacio. Una coma
  pegada a un número (`1,800`) la trata como parte del precio, no como columna.
- **Productos que ya tienes**: no los repite, les actualiza el precio. Puedes
  usar la importación cada vez que te cambien la lista de precios.
- **Renglones de cierre** (`TOTAL`, `IVA`, `Subtotal`) los descarta.
- Un producto sin precio entra en 0 y te avisa cuántos quedaron así.

Con *Reemplazar todo* borras el catálogo y dejas solo lo de la lista. **Tus
pedidos ya hechos nunca se tocan**: cada factura guarda el nombre y el precio
que tenía el producto ese día.

### Desde una foto

La app no lee fotos, y es a propósito: los lectores de texto que se pueden
meter dentro de una página web pesan varios megas y se equivocan bastante. Tu
teléfono ya trae uno mucho mejor.

1. Tómale la foto a la lista de precios.
2. Ábrela con **Google Fotos** y toca **Lente** (o usa **Google Lens**
   directamente). En iPhone, abre la foto y toca el icono de texto.
3. Selecciona el texto reconocido y cópialo.
4. En la app: *Importar lista* → pega → revisa la tabla → *Importar*.

El paso de revisar es importante: ningún lector de fotos es perfecto, y ahí ves
lo que quedó mal antes de que entre al catálogo.

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
- Los datos viven en `localStorage`, bajo la clave `gestor_pedidos_v1`. La foto
  de perfil se guarda ahí mismo como data URI JPEG (máximo 420 px de lado).
