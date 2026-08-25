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

## Actuaciones que restan (segunda mitad del expediente)

Formularios preimpresos y trámites que solo pueden materializarse ante las
entidades: pre-RUT (formulario 001 DIAN), carátula única empresarial y anexos,
formulario adicional de registro con otras entidades, presentación personal o
autenticación notarial, pago del impuesto de registro ante la Gobernación de
Antioquia, copias ampliadas de cédulas, uso del suelo y concepto sanitario,
inscripción en el REPS y habilitación de servicios de salud, e inscripción en el
RIT de Medellín. El detalle está en el documento 04 y en la sección VII del 05.

## Cómo regenerar los documentos

Las fuentes en `fuentes/*.txt` usan un marcado ligero que `fuentes/build_docx.py`
convierte a Word:

```bash
pip install python-docx
python3 fuentes/build_docx.py fuentes/01_documento_privado_constitucion.txt \
        docx/01_Documento_privado_de_constitucion_NOVA_BIENESTAR_INTEGRAL_SAS.docx
```

Para producir los PDF: `soffice --headless --convert-to pdf --outdir pdf docx/*.docx`
