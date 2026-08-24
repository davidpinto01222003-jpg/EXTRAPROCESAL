# Derechos de petición — prescripción de multas de tránsito

Escritos suscritos **a título personal** por cada peticionario (no en calidad de
apoderado), sobre el membrete de Jorge David Caicedo Pinto.

| Archivo | Peticionario | Autoridad | Comparendos |
|---|---|---|---|
| `DP-PRESCRIPCION-BUCARAMANGA.docx` | Rafael López Gutiérrez | Dirección de Tránsito de Bucaramanga | 3 |
| `DP-PRESCRIPCION-CURITI.docx` | Rafael López Gutiérrez | Secretaría de Tránsito de Curití | 1 |
| `DP-PRESCRIPCION-EL-PLAYON.docx` | Rafael López Gutiérrez | Secretaría de Tránsito de El Playón | 1 |
| `DP-PRESCRIPCION-FLORIDABLANCA.docx` | Brahayan Camilo Rosso Vacca | Dirección de Tránsito de Floridablanca | 1 |

Siguen el formato del modelo de Aguachica: petición escueta, sin tabla, sin
jurisprudencia, sin petición subsidiaria y **sin mencionar las resoluciones de
cobro coactivo** ni afirmar nada sobre notificaciones. Sólo se relacionan número
de comparendo, fecha y lugar.

## Membrete

Tomado del PDF de cotización: logo `logo.png` (200,25 × 66,75 pt, el tamaño
exacto del original), regla dorada `#B8962E`, y pie centrado con teléfono y
correo en gris `#777777`. Va en el encabezado/pie de Word, así que se repite en
todas las páginas. Es constante: no cambia con el peticionario.

## Campos por completar antes de radicar

En `datos.js` (objetos `RAFAEL` y `BRAHAYAN`) o directamente en Word:

- `domicilio` — párrafo de encabezamiento.
- `ciudadNotif`, `direccion`, `cel` — sección NOTIFICACIÓN.
- `emailNotif` — Rafael tiene el suyo; Brahayan usa el del membrete porque no se
  suministró uno propio.
- `fechaLarga` — hoy dice "agosto de 2026" (el modelo usa mes y año, sin día).

## Regenerar

```bash
npm install docx
node generar.js      # escribe en ./salida/
```

Para agregar otro peticionario: definir un objeto nuevo junto a `RAFAEL` y
`BRAHAYAN`, y agregar el destino en `DESTINOS` con `pet:` apuntando a él.

## Datos de los comparendos

Están todos en `datos.js`. Los campos `resolucion`, `fechaResolucion`,
`coactivo`, `fechaCoactivo` y `placa` quedan cargados pero **no se imprimen** en
esta versión del escrito; están ahí por si se necesitan más adelante.
