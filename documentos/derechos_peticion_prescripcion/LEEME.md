# Derechos de petición — prescripción de multas de tránsito

Tres derechos de petición para **Rafael López Gutiérrez** (C.C. 1.098.650.278),
suscritos **a título personal** (no en calidad de apoderado), sobre el membrete
de Jorge David Caicedo Pinto.

| Archivo | Autoridad | Comparendos |
|---|---|---|
| `DP-PRESCRIPCION-BUCARAMANGA.docx` | Dirección de Tránsito de Bucaramanga | 3 |
| `DP-PRESCRIPCION-CURITI.docx` | Secretaría de Tránsito de Curití | 1 |
| `DP-PRESCRIPCION-EL-PLAYON.docx` | Secretaría de Tránsito de El Playón | 1 |

Siguen el formato del modelo de Aguachica: petición escueta, sin tabla, sin
jurisprudencia, sin petición subsidiaria y **sin mencionar las resoluciones de
cobro coactivo** ni afirmar nada sobre notificaciones. Sólo se relacionan número
de comparendo, fecha y lugar.

## Membrete

Tomado del PDF de cotización: logo `logo.png` (200,25 × 66,75 pt, el tamaño
exacto del original), regla dorada `#B8962E`, y pie centrado con teléfono y
correo en gris `#777777`. Va en el encabezado/pie de Word, así que se repite en
todas las páginas.

## Campos por completar antes de radicar

En `datos.js`, o directamente en Word:

- `domicilio` — domicilio de Rafael (párrafo de encabezamiento).
- `ciudadNotif`, `direccion`, `cel` — sección NOTIFICACIÓN.
- `fechaLarga` — hoy dice "agosto de 2026" (el modelo usa mes y año, sin día).

## Regenerar

```bash
npm install docx
node generar.js      # escribe en ./salida/
```

## Datos de los comparendos

Los cinco comparendos están en `datos.js`. Los campos `resolucion`,
`fechaResolucion`, `coactivo`, `fechaCoactivo` y `placa` quedan cargados pero
**no se imprimen** en esta versión del escrito; están ahí por si se necesitan
más adelante.
