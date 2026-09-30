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

## Resultado real de la v4 y cuarta pasada (v5)

Compilatio (Studium) analizó la v4 el 29-09-2026: **26 %** en total, con **22 %** de IA, **4 %** de similitudes y **0 %** de idiomas no reconocidos. El simulador anterior había dicho 16,9 % de IA, y el de detectores de Hugging Face, 24,4 %.

### Qué aprendimos del informe de la v4
- Con el nuevo informe se reentrenó el simulador (`simulador_compilatio/`) usando los **dos** informes. Ahora usa gradient boosting con 29 rasgos y memoria de las marcas reales. Sus aciertos: AUC 0,76 en validación cruzada y 0,74 al predecir las marcas de la v4 entrenando solo con el original. Con la v4 da 26,1 %, frente al 26 % real.
- Los detectores de Hugging Face no sirvieron para predecir la v4 (AUC 0,57) y quedaron como opción.
- **La técnica de la v2 a la v4 fue contraproducente.** Las oraciones reescritas salieron más marcadas (37 %) que las originales (32 %). Lo que más pesa en las marcas es que el párrafo tenga muchas oraciones y mucha variación de largo: frases cortísimas ("Son tres.", "Resultado: tribunas a medio llenar.") intercaladas con otras largas y preguntas retóricas.

### Qué se cambió en la v5
Se reescribieron, en prosa académica corriente, los párrafos con más texto marcado en el informe de la v4. Las oraciones son de largo parejo, no hay frases sueltas de dos o tres palabras ni preguntas retóricas, y los párrafos muy largos se dividieron en dos.
- Resumen y Abstract (las oraciones marcadas).
- Hipótesis, segunda estrategia (Castaño Pérez et al.) y tercera estrategia (los tres casos, dividida en dos párrafos).
- Lectura jurídica del caso 1 (dividida en dos párrafos).
- Marco normativo ("En el papel…"), jurisprudencia ("La jurisprudencia apunta…") e Inglaterra/Hillsborough.
- Situaciones de causa extraña ("Con esos criterios en la mano…", dividida en dos párrafos).
- Efectos económicos ("El segundo efecto…", dividida en dos párrafos).
- Introducción a las conclusiones y conclusiones 01 a 06. Cada conclusión quedó en tres párrafos: *Lo demostrado*, *Interpretación de los autores* y *Propuesta*. En la 04 y la 06 solo cambió esa división; el texto es el mismo.

Se comprobó automáticamente que la v5 conserva todos los números, años, citas, artículos, radicados y sentencias de la v4.

### Ortografía
Se revisó todo el documento con LanguageTool (español e inglés), ejecutado en local. El texto estaba limpio; se corrigió:
- Coma antes de "sino" en "no estamos ante violencia entre hinchas, sino ante crimen organizado" y en la conclusión 02.
- "resolución por resolución, por fallas de seguridad…" (faltaba la coma; se leía como una repetición).

Se dejaron como están, porque son correctos o porque no son texto de los autores:
- Nombres propios y siglas: Dimayor, Pécaut, Monsalvo, Uribe Aramburo, AUF, FCF.
- Términos regionales o técnicos: barrismo, cortopunzantes, contravencional, sacol, dick.
- "sólo" dentro de la cita textual de la Corte Suprema.
- "C.Co." (Código de Comercio) en la Tabla 2.
- "APA 7ma ed." en la ficha de la biblioteca de la portada. Si quieren corregirla, la forma normativa es "7.ª ed.".

### Estimación con el simulador reentrenado

| Versión | IA | Similitud | Idiomas | Total |
|---|---|---|---|---|
| v4 (Compilatio real) | 22 % | 4 % | 0 % | 26 % |
| v4 (simulador) | 22,1 % | 4,0 % | 0,0 % | 26,1 % |
| v5 (simulador) | **9,7-12,8 %** | 4,0 % | 0,0 % | **13,7-16,8 %** |

El rango de IA va de la estimación con umbral a la estimación con valor esperado (más pesimista). Hay un escenario aún más pesimista: que Compilatio marcara las oraciones nuevas de la v5 en la misma proporción que las reescritas de la v4 (37 %). Aun así, el total quedaría cerca del **18 %**, por debajo del 24 % buscado. El detalle por frase está en `simulador_compilatio/reporte_simulado_v5.html`. El número real solo lo da Compilatio.

## Prueba de los párrafos reescritos y quinta pasada (v6)

Los 18 párrafos reescritos en la v5 se pasaron solos por Compilatio (Studium, 3.481 palabras): **38 % de IA**. El simulador esperaba cerca de un 10 %. El resultado fue el escenario pesimista que se había anticipado (37 %). Con esas marcas reales, la v5 completa quedaría en unos 16,8 % de IA y 20,8 % en total.

**Qué pasó y qué no.** Pasaron limpios los párrafos anclados en datos concretos:
- el Abstract,
- la segunda estrategia,
- el marco normativo,
- Inglaterra/Hillsborough,
- la primera parte del caso 1,
- las conclusiones 02, 04 y 06 y varias propuestas.

Siguieron marcados los de argumentación general:
- la hipótesis,
- la tercera estrategia,
- la parte civil del caso 1,
- la zona gris y la causa extraña,
- los efectos económicos,
- la jurisprudencia,
- las conclusiones 01, 03 y 05.

**Qué se cambió en la v6.** Se reescribieron solo las frases que siguieron marcadas, anclándolas en datos que ya estaban en el trabajo:
- Andrés Carvajal (Cali, febrero de 2020).
- El Barón Rojo Sur y la veintena de homicidios.
- Laferrere y los quince detenidos en 2018.
- La Resolución 095 de 2024 y la multa de trece millones de pesos al Junior.
- El Decreto 1007 de 2012.
- Las sentencias C-065 de 2021, SC2905-2021 y SC2111-2021.

Además se reescribieron las frases marcadas en la v4 de cinco párrafos que no se habían tocado: lectura jurídica del caso 2, "Vistos juntos, los tres casos…", el asistente con boleta, el desenlace de la SC2111-2021 y los deberes de la Dimayor y la FCF.

No se agregó ningún dato nuevo: todo lo mencionado ya estaba en otra parte del trabajo. Se comprobó que siguen todas las cifras, citas y radicados. La ortografía de lo nuevo se revisó con LanguageTool y se corrigió una coma ("el tercero agrede, pero el organizador…").

| Versión | IA | Similitud | Total |
|---|---|---|---|
| v5 completa (con las marcas reales de los párrafos) | ~16,8 % | 4 % | ~20,8 % |
| v6 (simulador, con el informe de los párrafos incluido) | ~12,7 % | 4 % | **~16,7 %** |

Esta vez el simulador ya es pesimista con las frases nuevas: supone que alrededor del 60 % saldrán marcadas. `Parrafos reescritos en v6.docx` (22 párrafos, 3.283 palabras) sirve para probar la ronda por separado, igual que con la v5.
