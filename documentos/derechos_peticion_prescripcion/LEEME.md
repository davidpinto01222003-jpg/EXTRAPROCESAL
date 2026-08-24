# Derechos de petición — prescripción de multas de tránsito

Tres derechos de petición para **Rafael López Gutiérrez** (C.C. 1.098.650.278),
presentados **a título personal** (no como apoderado), sobre membrete de abogado.

| Archivo | Autoridad | Comparendos |
|---|---|---|
| `DP-PRESCRIPCION-BUCARAMANGA.docx` | Dirección de Tránsito de Bucaramanga | 3 |
| `DP-PRESCRIPCION-CURITI.docx` | Secretaría de Tránsito de Curití | 1 |
| `DP-PRESCRIPCION-EL-PLAYON.docx` | Secretaría de Tránsito de El Playón | 1 |

## Antes de firmar y radicar — completar los campos en blanco

1. **Membrete** (encabezado del documento, se repite en todas las páginas):
   número de T.P. y celular. Si se tiene el membrete en imagen, se reemplaza
   el encabezado por la imagen desde Word (Insertar → Encabezado).
2. **Dirección física de notificación** (sección NOTIFICACIÓN).
3. **Celular** (sección NOTIFICACIÓN).
4. Verificar la **fecha** del documento y la ciudad de expedición de la cédula.

## Regenerar los documentos

Los datos de los cinco comparendos están en `datos.js`; el armado del texto en
`generar.js`. Para regenerar tras editar los datos:

```bash
npm install docx
node generar.js      # escribe en ./salida/
```

## Nota sobre el argumento jurídico

En los cinco casos han transcurrido más de tres (3) años desde la **ocurrencia
del hecho** (art. 159 Ley 769/2002, mod. art. 206 Decreto Ley 019/2012).

A diferencia del formato base, aquí las resoluciones de cobro coactivo se
expidieron *dentro* de los tres años. Por eso la petición se apoya en que la
prescripción **sólo se interrumpe con la NOTIFICACIÓN del mandamiento de pago**,
no con la expedición de actos internos, y en que no consta tal notificación.
Se incluye una **petición subsidiaria** que pide copia del mandamiento de pago y
de su constancia de notificación, y, en subsidio, la prescripción de la acción de
cobro (art. 817 E.T. / Ley 1066 de 2006).
