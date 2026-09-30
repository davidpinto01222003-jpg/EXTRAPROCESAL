# Trabajo de grado Suárez y Rivera: cambios frente al informe Compilatio (28-09-2026)

## Diagnóstico del informe (35 % de textos sospechosos)

| Indicador | % | Qué lo causa |
|---|---|---|
| Detección de IA | 27 % | Unas 4.000 de 15.015 palabras marcadas. Se concentran en el Resumen, el Abstract, la metodología, la lectura jurídica de los casos, el régimen civil y las seis conclusiones (en algunas está marcado el 90-100 % del párrafo). |
| Similitudes | 6 % | Ninguna fuente pasa del 1 %. Casi todo son la portada y la plantilla de la USB, la bibliografía, los nombres de leyes y citas textuales de sentencias. Hay algunas citas literales sin comillas (art. 1603 y art. 1616 del C.C.) y citas largas de fuentes argentinas y uruguayas. |
| Idiomas no reconocidos | 3 % | Palabras en inglés (Abstract, *Keywords*, *Football Spectators Act*), números de expediente, URL de la portada y celdas de tablas. No hay caracteres raros ni texto oculto en el Word (se revisó). |

## Qué se cambió

**Párrafos marcados como IA (reescritos, 30 párrafos):** Resumen, Abstract, introducción (plantea la hipótesis y los tres planos de análisis), limitaciones, estrategia de casos, ejes de resultados, lectura jurídica de los casos 1, 2 y 3, síntesis de los casos, marco normativo, SC2111-2021 y su desenlace, contrato de espectáculo (arts. 1603, 1604 y 1616), criterios de causa extraña, Dimayor/FCF, efectos económicos, Inglaterra, consumo de sustancias, sociología del deporte, carnetización y las conclusiones 01 a 06.

Técnica aplicada:
- Oraciones de largo variable, con oraciones cortas intercaladas.
- Primera persona del plural en las partes de interpretación ("revisamos", "a nuestro juicio").
- Se quitaron enumeraciones anunciadas ("tres consecuencias: la primera… la segunda…"), conectores de manual ("conviene", "esto es", "de ahí", "en consecuencia") y cierres tipo moraleja.
- Se mantuvieron los títulos en negrita y los rótulos en cursiva ("Lo demostrado", "Interpretación de los autores", "Propuesta").

**Similitud:**
- Art. 1603 y art. 1616 del C.C.: el texto literal quedó entre comillas.
- Castaño Pérez et al. (2014), p. 79 y p. 84: las citas literales se volvieron paráfrasis con número de página.
- Tribuna Segura (Res. 843/2018) y el plazo uruguayo del 31 de marzo: el texto transcrito se volvió paráfrasis con la misma cita.
- Uruguay, 80 % de reducción de detenciones: paráfrasis con la misma cita.

**Corrección de fondo:** en el párrafo de la SC2111-2021 decía que Quiroz Monsalvo fue "ponente ocho semanas después de la SC2905-2021". Por las fechas (2 de junio y 29 de julio de 2021), él fue ponente *de* la SC2905-2021 ocho semanas después. Quedó así.

## Qué NO se tocó
- Datos, cifras, fechas, radicados, número de sentencias y referencias.
- Citas textuales de la Corte Suprema (SC2905-2021 y sentencia de 2011): se dejaron literales y entre comillas.
- Tablas, bibliografía, portada y demás párrafos no marcados.

## Recomendaciones antes de volver a subir a Compilatio
1. Revisen que la redacción en primera persona del plural ("nosotros") sea aceptada por su asesor; si no, es fácil cambiarla a "los autores".
2. Si la plataforma lo permite, excluyan del análisis la portada, la página de la biblioteca y la bibliografía. Eso baja la similitud sin tocar el texto.
3. Suban el .docx directamente en vez de una conversión a .txt, para que Compilatio reconozca mejor las comillas y las tablas.

## Segunda pasada (v3)

Se reescribieron 11 párrafos más, que Compilatio había marcado en parte o que el simulador seguía puntuando alto:
- Limitaciones de la información estadística.
- Lectura jurídica del caso 3.
- Vías penal, contravencional y administrativa (incluye las SP3573-2022 y SP289-2023).
- Cifras y encuadre jurídico.
- Deberes de clubes y organizadores.
- Tercero sin vínculo contractual.
- Conducta de terceros y causa extraña.
- Estado, Policía y comisiones (Consejo de Estado y C-065 de 2021).
- Pérdidas económicas (Resolución 095 de 2024).
- Síntesis de la jurisprudencia.
- Introducción a las conclusiones.

No cambió ningún dato, cita, radicado ni referencia.

### Estimación con el simulador (`simulador_compilatio/`)

| Versión | IA | Similitud | Idiomas | Total |
|---|---|---|---|---|
| Original (Compilatio real) | 27 % | 6 % | 3 % | 35 % |
| Original (simulador) | 28,9 % | 4,4 % | 3,0 % | 36,3 % |
| v2 (simulador) | 26,3 % | 4,2 % | 3,0 % | 33,5 % |
| v3 (simulador) | 23,2 % | 4,2 % | 3,0 % | 30,4 % |

El simulador es aproximado (AUC ≈ 0,67). El número real solo lo da Compilatio.

## Tercera pasada (v4)

Se reescribieron 17 párrafos más, los que el simulador seguía marcando en la v3:
- Introducción sobre muertes y registros.
- Estrategia de casos, respuesta institucional y ejes de resultados.
- Lecturas jurídicas de los casos 2 y 3.
- Vías de responsabilidad del hincha.
- Contrato de espectáculo y terceros sin contrato.
- Sentencia de 2011: se dejó intacta la cita textual.
- Causa extraña y experiencias internacionales.
- Introducción a las conclusiones y conclusiones 02 a 06.

Técnica: oraciones más cortas y en lenguaje más llano, menos dos puntos y punto y coma, preguntas retóricas puntuales y primera persona del plural. No cambió ningún dato, cita, radicado ni referencia.

| Versión | IA | Similitud | Idiomas | Total |
|---|---|---|---|---|
| v4 (simulador) | **16,9 %** | 4,2 % | 3,0 % | **24,1 %** |

## Nueva medición con detectores de Hugging Face

El simulador ahora usa, además de los rasgos de estilo, tres detectores de IA publicados en Hugging Face: un XLM-RoBERTa afinado en español (`CradeyMH/detector-ia-espanol`), el modelo de AuTexTification 2023 (`pandrei7/autextification-upb-mtl`) y Binoculars con Qwen2.5-0.5B. Se aplican al párrafo completo, porque con una sola oración casi no aportan. Detalles en `simulador_compilatio/README.md`.

Se midió de nuevo el AUC agrupando por párrafo, para que las oraciones de un mismo párrafo no queden a la vez en entrenamiento y en prueba. Con esa medición, el modelo anterior (solo estilo) tenía un AUC de **0,63**, no de 0,67. El nuevo llega a **0,71**.

| Versión | IA (simulador anterior) | IA (simulador con detectores HF) | Similitud | Idiomas | Total (con detectores HF) |
|---|---|---|---|---|---|
| Original (Compilatio real: 27 %) | 28,9 % | 26,9 % | 4,4 % | 3,0 % | 34,3 % |
| v2 | 26,3 % | 27,2 % | 4,2 % | 3,0 % | 34,3 % |
| v3 | 23,2 % | 27,2 % | 4,2 % | 3,0 % | 34,4 % |
| v4 | 16,9 % | **24,4 %** | 4,2 % | 3,0 % | **31,5 %** |

**Qué significa:** el 16,9 % que daba el simulador anterior para la v4 era demasiado optimista. Las reescrituras de la v2 a la v4 atacaron justo los rasgos que ese modelo medía (oraciones largas, comas, conectores de manual), así que el número bajaba por construcción. El modelo con detectores, que distingue mejor lo que marcó Compilatio, estima que la v4 sigue alrededor del **24 %** de IA.

Otro hallazgo: en este trabajo, los detectores públicos puntúan al revés que Compilatio. Los párrafos que Compilatio marcó son los que esos detectores consideran más humanos. Por eso, pasar un párrafo por un detector gratuito y ver "humano" no garantiza nada frente a Compilatio.

Párrafos de la v4 que el simulador sigue marcando casi completos: conclusiones 03, 04 y 06; efectos económicos ("El segundo efecto cuesta más medirlo…"); experiencias internacionales ("Lo que muestran los otros países…"); identificación de asistentes; estrategia de casos; hipótesis y tres planos de la introducción. El detalle frase por frase está en `simulador_compilatio/reporte_simulado_v4.html`.
