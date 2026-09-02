# Constitución de NOVA BIENESTAR INTEGRAL S.A.S.

Expediente de simulación para la actividad evaluativa de Sociedades Comerciales
(constitución de una Sociedad por Acciones Simplificada, Ley 1258 de 2008),
elaborado a partir del *Levantamiento de información para la constitución de
sociedad por acciones simplificada*.

- **Sociedad:** NOVA BIENESTAR INTEGRAL S.A.S.
- **Domicilio:** Medellín (Antioquia) — Carrera 37 n.º 10A-58, oficina 804
- **Cámara competente:** Cámara de Comercio de Medellín para Antioquia
- **Fecha del acto constitutivo:** 1.º de octubre de 2026
- **Capital:** autorizado $1.200.000.000 / suscrito y pagado $600.000.000 / valor nominal $1.000

## Documentos elaborados (`docx/` editables, `pdf/` para impresión)

| N.º | Documento | Función |
|-----|-----------|---------|
| 01 | Documento privado de constitución (estatutos) | Crea la sociedad y fija su régimen. 42 artículos y 6 transitorios. |
| 02 | Cartas de aceptación de cargos | Aceptación del representante legal principal y del suplente (art. 42 D.L. 2150/1995). |
| 03 | Certificación de capital y declaración Ley 1780 de 2016 | Acredita el pago del capital y sustenta la exención de matrícula. |
| 04 | Guía de formularios RUES y liquidación de derechos | Campo por campo del RUT, la carátula única empresarial, los anexos y el formulario de otras entidades, más la liquidación de costos. |
| 05 | Memorando jurídico justificativo | Explica y defiende cada documento y cada decisión estatutaria. |
| 06 | Solicitud de inscripción de libros y libro de registro de accionistas | Registro de los libros del art. 175 del D.L. 019/2012. |
| 07 | Formularios del RUES — hojas de transcripción | Carátula única (2 páginas), anexo de persona jurídica, anexo de establecimiento y formulario de otras entidades, campo por campo. |
| 08 | Formulario 001 del RUT — hoja de transcripción | Ruta del trámite en línea ante la DIAN y análisis de las responsabilidades tributarias. |
| 09 | Títulos de acciones y comprobantes de aporte | Títulos 001 y 002 con las menciones del art. 401 C.Co. y los soportes contables del capital. |
| 10 | Presentación personal, reconocimiento notarial y poder | Las dos vías de autenticación del art. 40 C.Co. y el poder especial para radicar. |
| 11 | Impuesto de registro y planilla de radicación | Liquidación del impuesto departamental y lista de chequeo para la ventanilla. |
| 12 | Acta n.º 1 de la asamblea general de accionistas | Reunión universal de puesta en marcha y folio de apertura del libro de actas. |
| 13 | Solicitudes ante otras entidades | Uso del suelo, REPS y habilitación en salud, RIT de Medellín, y cuadro de trámites restantes. |
| 14 | Memorando complementario | Explica y defiende cada documento de la segunda parte, con la ruta cronológica consolidada. |
| 15 | Informe final de auditoría y estado del expediente | Resultado de la verificación, inventario, lo que falta y dónde conseguirlo. |

## Formatos oficiales que hay que descargar

Estos formatos no pueden ser elaborados por el solicitante: deben obtenerse del
formato preimpreso de la Cámara o del portal de la entidad correspondiente. Los documentos 07 y 08 son
**hojas de transcripción**: fijan el valor de cada campo, pero no sustituyen el
formato oficial.

| Formato | Dónde se obtiene |
|---------|------------------|
| Carátula única empresarial (2 páginas) | camaramedellin.com.co → servicios registrales → registro mercantil → «Formatos registro mercantil»; también rues.org.co → «Formatos CAE» (archivo de Medellín) |
| Anexo 1 — establecimientos de comercio | Mismo origen. Único anexo aplicable; el anexo 2 es del Registro Único de Proponentes |
| Formulario adicional de registro con otras entidades | Taquilla de la Cámara o su sede virtual |
| Instructivo del formulario RUES | Mismo sitio de la Cámara, junto a los formularios |
| Formulario 001 — RUT | Portal transaccional de la DIAN, opción «inscripción en el RUT por Cámara de Comercio». No existe versión descargable en blanco |
| Liquidación del impuesto de registro | Portal tributario de la Gobernación de Antioquia, o liquidación de la Cámara si opera el convenio de recaudo |

No existe un anexo separado de persona jurídica: la carátula única sirve a
persona natural y jurídica, y los datos societarios (capital, objeto, duración,
administradores) los toma la Cámara del documento privado de constitución.

La ruta cronológica completa de los 19 pasos está en la sección III del
documento 14.

## Cómo regenerar los documentos

Las fuentes en `fuentes/*.txt` usan un marcado ligero que `fuentes/build_docx.py`
convierte a Word:

```bash
pip install python-docx
python3 fuentes/build_docx.py fuentes/01_documento_privado_constitucion.txt \
        docx/01_Documento_privado_de_constitucion_NOVA_BIENESTAR_INTEGRAL_SAS.docx
```

Para producir los PDF: `soffice --headless --convert-to pdf --outdir pdf docx/*.docx`

## Estado

Expediente auditado. La verificación cruzada de cifras, referencias y
operatividad aritmética de las cláusulas está documentada en el documento 15,
junto con el inventario, lo que falta y dónde conseguirlo.

**Lo que se radica en la Cámara:** documentos 01, 02, 03, 06 y 10, más los
formularios oficiales y el comprobante del impuesto de registro.
**Instrumentos de trabajo** (no se entregan): 04, 07, 08 y 11.
**Sustentación académica:** 05, 14 y 15.
