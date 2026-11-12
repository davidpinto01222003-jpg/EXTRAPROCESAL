# Diagnóstico de estilo y reescritura (v4)

`Trabajo lectoescritura v4.docx` se genera con `python version4.py "Trabajo lectoescritura v3.docx" "Trabajo lectoescritura v4.docx"`.

## Qué tenía la v3

El largo promedio de oración era razonable (23 palabras). Lo que daba la sensación de texto cortado y robótico era la estructura:

- **Dos puntos de "revelación"** ("La idea es sencilla: …"): 52. Junto con "Por eso" (16) y "entonces," (8), le daban al texto una cadencia de manual.
- **Advertencias repetidas**: "no permite diagnosticar dislexia", "no se puede generalizar" y similares aparecían 14 veces, en casi todos los párrafos del análisis. Ese exceso de equilibrio hacía que ninguna afirmación quedara en pie.
- **La misma explicación una y otra vez**: por qué se usan pseudopalabras se explicaba 9 veces.
- **Los resultados de cada estudiante, dos veces en prosa**: en Resultados y en Análisis.
- **Ideas cortadas**: 109 de 206 párrafos tenían una o dos oraciones, muchas veces una idea suelta sin desarrollo ("La distinción no es menor.", "Esto importa porque…").
- **Las nueve reseñas de antecedentes seguían el mismo molde**: resumen, cita, "Desde una mirada crítica…" y "Para concluir / En suma…", con un cierre que repetía lo ya dicho.
- **Aperturas previsibles** ("La tercera actividad fue especialmente útil para…") y **cierres obvios** ("la lectoescritura es un proceso complejo", "esa distinción le da solidez a la investigación").
- **El marco teórico repetía los antecedentes** y su introducción anticipaba punto por punto los subtítulos.

## Qué se hizo

- **Resultados**: se dejan los datos de cada actividad, sin interpretarlos, y los perfiles individuales pasan a la **Tabla 1** (formato APA).
- **Análisis**: queda la interpretación frente a la teoría, una vez por componente, y las limitaciones van reunidas en un solo párrafo.
- **Antecedentes**: cada reseña tiene ahora una estructura propia según lo que aporta el artículo, y se quitaron los cierres de relleno.
- **Marco teórico**: sin repeticiones de los antecedentes; cada subtítulo dice algo distinto.
- **Resumen y Abstract**: en un solo párrafo, como pide APA.
- **Dedicatoria y agradecimientos**: el mismo contenido, con ritmo menos entrecortado.
- **Sin cambios**: datos, citas, referencias, objetivos, preguntas de investigación y la carta "Un alto en el camino".

## Medición (prosa de Resumen a Conclusiones, sin la carta ni las tablas)

| Medida | v3 | v4 |
|---|---|---|
| Palabras de prosa (sin carta ni tablas) | 11687 | 5612 |
| Párrafos de prosa | 206 | 80 |
| Oraciones | 502 | 182 |
| Largo medio de oración | 23.3 | 30.8 |
| Variación del largo (CV) | 0.44 | 0.47 |
| Oraciones de ≤8 palabras | 32 | 20 |
| Párrafos de 1 o 2 oraciones | 109 | 49 |
| dos puntos de revelación | 52 | 5 |
| "Por eso" | 16 | 3 |
| "entonces," | 8 | 0 |
| advertencias repetidas | 14 | 3 |
| "no es... sino" | 3 | 1 |
| "Sin embargo/No obstante" | 11 | 0 |
| explicación de por qué pseudopalabras | 9 | 3 |
| "proceso complejo" | 4 | 0 |

El texto quedó con oraciones más largas y encadenadas, aunque variadas (solo 5 pasan de 55 palabras, y en general porque llevan citas).

**Ojo con la extensión.** La prosa bajó de unas 11.700 a unas 5.600 palabras, porque casi la mitad eran repeticiones (sobre todo en Resultados y Análisis). Si el programa exige un mínimo de páginas, conviene ampliar con contenido nuevo, no con relleno. Por ejemplo:
- una propuesta pedagógica concreta derivada de los hallazgos (actividades para aislar fonemas, que es la dificultad principal);
- una discusión más extensa con cada antecedente;
- la descripción de la institución y del grupo cuando la autora tenga esos datos.

## Turnitin
- Del texto que Turnitin marcó en el original, en la v4 sigue igual el **7 %**, frente al 15 % en la v3. Si Turnitin solo volviera a marcar eso, saldría alrededor de **3 %**.
- El simulador de estilo da **~31 %**, pero es poco confiable (AUC 0,61) y castiga los párrafos largos y con citas, así que tómalo como el peor caso.
- El número real solo lo da Turnitin.


## v5: ampliación con contenido nuevo

`Trabajo lectoescritura v5.docx` se genera con `python version5.py "Trabajo lectoescritura v4.docx" "Trabajo lectoescritura v5.docx"`. Recupera extensión con contenido que antes no estaba en el trabajo, sin volver a los patrones anteriores.

**Contenido nuevo que sale de los datos de la transcripción**
- **Tipología de los 41 errores de lectura (Tabla 2).** El 43,9 % afecta grupos consonánticos con l o r: 10 casos ocurren en las 13 pseudopalabras que tienen grupo y 8 en palabras sin grupo, donde el estudiante lo creó.
- **Ningún estudiante confundió b con d ni p con q,** aunque la lista tenía esas letras. Esto refuerza que las dificultades son fonológicas y no visoespaciales.
- **Dictado: errores fonológicos frente a ortográficos.** "Sutel" por "Zutel" (5 estudiantes), "Bentil" por "Ventil" (3) y "Cuamir" por "Quamir" (2) suenan igual en el español de Colombia. Los cambios que sí alteran el sonido son "Pamil" por "Pamir" (4), "Ventir" (3), "Tralu" (3) y "Neclon" (3).
- **Descripción del instrumento** (30 pseudopalabras bisílabas, 13 con grupo consonántico; dos listas de dictado), del procedimiento y de los criterios de análisis.

**Contenido nuevo con fuentes verificadas**
- **Planteamiento:** PISA 2022, Colombia 409 en lectura frente a 476 de la OCDE (OECD, 2023).
- **Marco teórico:**
  - el modelo de doble ruta (Coltheart et al., 2001) y su uso en el PROLEC-R (Cuetos et al., 2007);
  - los niveles de la conciencia fonológica (Jiménez González & Ortiz González, 1998);
  - las sílabas con grupos consonánticos (Jiménez González & Jiménez Rodríguez, 1999);
  - la ortografía del español de América (RAE & ASALE, 2010).
- **Discusión:** relación con la comprensión lectora (Pisco-Román & Bailón-Panta, 2023) y con las intervenciones de Domínguez Vázquez (2023) y Rivera Cintrón y Batiz Cartagena (2024).
- **Propuesta pedagógica** por perfiles (de la sílaba al fonema, sílabas trabadas, escritura y ortografía), basada en la Tabla 1.

**La autora debe revisar las fuentes nuevas antes de entregar.** Estas seis referencias las agregué yo y ella no las consultó: Coltheart et al. (2001), Cuetos et al. (2007), Jiménez González y Jiménez Rodríguez (1999), Jiménez González y Ortiz González (1998), OECD (2023) y RAE y ASALE (2010). De Jiménez González y Ortiz González conviene confirmar el año de la edición que consulte: el catálogo indica 1998, pero hay otras ediciones.

| Medida | v3 | v5 |
|---|---|---|
| Palabras de prosa (sin carta ni tablas) | 11687 | 7885 |
| Párrafos de prosa | 206 | 101 |
| Oraciones | 502 | 254 |
| Largo medio de oración | 23.3 | 31.0 |
| Variación del largo (CV) | 0.44 | 0.49 |
| Oraciones de ≤8 palabras | 32 | 25 |
| Párrafos de 1 o 2 oraciones | 109 | 50 |
| dos puntos de revelación | 52 | 4 |
| "Por eso" | 16 | 4 |
| "entonces," | 8 | 0 |
| advertencias repetidas | 14 | 3 |
| "no es... sino" | 3 | 1 |
| "Sin embargo/No obstante" | 11 | 0 |
| explicación de por qué pseudopalabras | 9 | 3 |
| "proceso complejo" | 4 | 0 |

La prosa queda en unas 7.900 palabras (la v4 tenía 5.600 y la v3 11.700, de las cuales casi la mitad eran repeticiones). Del texto que Turnitin marcó en el original sigue igual el 8 %.

## v6: formato APA 7 y discusión ampliada

`Trabajo lectoescritura v6.docx` se genera con `python version6.py "Trabajo lectoescritura v5.docx" "Trabajo lectoescritura v6.docx"`.

La Discusión pasa de 5 a 10 párrafos. Los cinco nuevos tratan temas que el trabajo no había discutido y usan solo datos de la transcripción y fuentes ya verificadas:
- La dislexia como un continuo (Snowling et al., 2020), frente a la distribución de 0 a 9 errores.
- El grado y la edad de los participantes, que no constan, y los grupos consonánticos (Jiménez González y Jiménez Rodríguez, 1999).
- El problema de medición de las intervenciones revisadas (Rivera Cintrón y Batiz Cartagena; Domínguez Vázquez).
- La formación docente y el conocimiento fonológico necesario para corregir el dictado ("Sutel" frente a "Pamil").
- La falta de baremos de la prueba propia frente al PROLEC-R (Cuetos et al., 2007).

También se quitaron los últimos "por eso" (había 4), un "no pretende… sino" y dos dos puntos que anunciaban la idea.

| Medida | v5 | v6 |
|---|---|---|
| Palabras de prosa | 7.885 | 8.531 |
| Largo medio de oración | 31,0 | 30,9 |
| Variación del largo (CV) | 0,49 | 0,49 |
| "Por eso" | 4 | 0 |
| "no es… sino" | 1 | 0 |
| Dos puntos de revelación | 4 | 1 (título de libro) |
| Simulador Turnitin: texto marcado que sigue igual | 8 % | 8 % |
