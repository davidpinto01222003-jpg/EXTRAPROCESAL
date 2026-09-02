# Formularios oficiales diligenciados

Formatos oficiales de la Cámara de Comercio de Medellín para Antioquia,
diligenciados con los datos de **NOVA BIENESTAR INTEGRAL S.A.S.**

El diseño de cada formato queda **intacto**: solo se escriben valores en los
campos del formulario. En los formatos con campos rellenables (AcroForm) el
valor queda además guardado como dato del campo, de modo que el archivo sigue
siendo editable en Acrobat o en el navegador.

| Formato | Origen | Método |
|---|---|---|
| `Formulario_RUES_caratula_DILIGENCIADO.pdf` | Carátula única empresarial, hojas 1 y 2 | 338 campos AcroForm |
| `Anexo1_RUES_establecimiento_DILIGENCIADO.pdf` | Anexo 1 — establecimientos de comercio | 282 campos AcroForm |
| `Anexo_responsabilidades_tributarias_DILIGENCIADO.pdf` | Anexo de responsabilidades tributarias (DIAN) | 28 campos AcroForm |
| `Formato7_inscripcion_libros_DILIGENCIADO.pdf` | Formato 7 — inscripción de libros | 73 campos AcroForm |
| `Formato_grupo_etnico_DILIGENCIADO.pdf` | Anexo de enfoque diferencial étnico | PDF plano: texto estampado |

Los originales sin diligenciar quedan en esta misma carpeta.

## Campos que se dejaron en blanco a propósito

**Datos que no existen todavía:** NIT y dígito de verificación, número de
matrícula mercantil, código postal, barrio, y la fecha de radicación. Se
obtienen durante el trámite.

**Casillas que exigen el instructivo del RUES:** «tipo general de
organización», «tipo específico de organización» y «código del estado actual de
la persona jurídica». El formato remite expresamente a las instrucciones, que
no se incluyeron entre los archivos; conviene tomar los códigos de allí.

**Anexo de enfoque diferencial étnico:** solo se diligenció la razón social.
Las tres preguntas sobre participación de personas de grupos étnicos en el
capital, en la planta y en cargos directivos piden datos personales
autorreportados que el levantamiento de información no contiene y que no
pueden inferirse. Debe responderlas el cliente.

## Decisiones que conviene revisar antes de firmar

- **Grupo NIIF: 2.** Con activos de $600.000.000 y nueve trabajadores la
  sociedad roza el límite del grupo 3. Confirmar con el contador.
- **Sede administrativa en arriendo** y **local ajeno** en el Anexo 1: se
  asumió que la oficina 804 es arrendada.
- **Empresa familiar: no.** El levantamiento no reporta parentesco entre los
  accionistas.
- **Responsabilidades tributarias: 05, 07, 14, 42, 48 y 55**, que es la
  combinación que el propio anexo de la Cámara señala como común para una
  S.A.S. No se incluye la 52 porque el anexo lo prohíbe expresamente:
  la facturación electrónica requiere habilitación previa y se agrega después,
  por actualización del RUT.

## Cómo se regeneran

```bash
pip install python-docx pypdf
python3 -c "from llenar import llenar; from datos_rues import caratula, CASILLAS; \
  llenar('Formulario_RUES.pdf','diligenciados/Formulario_RUES_caratula_DILIGENCIADO.pdf', \
         valores=caratula(), radios=CASILLAS)"
```

`llenar.py` escribe los valores y estampa las marcas de verificación como una
X trazada sobre el contenido de la página, porque varios visores no dibujan la
apariencia original de esas casillas. `datos_rues.py` contiene los valores del
caso.
