# Trabajo de LinkedIn: por qué Turnitin marcó 71 % de IA y qué se cambió

Trabajo: *In-Depth LinkedIn Optimization With Marketing Management Techniques: Canadian Logistics Industry* (MRKT-621, Marketing Management). Informe de Turnitin del 12 de noviembre de 2026: **71 % detectado como IA**.

**Versión nueva: `Trabajo LinkedIn revisado v2.docx`.**

## 1. Qué marcó Turnitin

Extraje del PDF del informe cada palabra con su resaltado y comparé los párrafos marcados con los que no se marcaron. Se marcó el 71 % de la prosa (3.867 de 5.439 palabras). Los párrafos que se salvaron tienen algo en común: están llenos de cifras y citas.

| Rasgo (promedio por párrafo) | Marcados | No marcados |
|---|---|---|
| Dígitos por cada 100 palabras | 6,8 | 21,9 |
| Citas | 0,55 | 1,06 |
| "I will" / "I would" | 0,91 | 0,06 |
| Palabras en primera persona por cada 100 | 4,8 | 2,3 |

Turnitin dejó pasar los datos y marcó la parte de razonamiento y de plan: los apartados de STP, las 4P, la revisión del perfil, el plan de acción, la conclusión y los apéndices.

## 2. Los patrones

1. **Oraciones enormes y todas iguales.** La oración media tenía **44,6 palabras**. De 122 oraciones, 75 pasaban de 40 palabras y solo una tenía menos de 15. Casi todos los párrafos eran una o dos oraciones larguísimas. Por ejemplo, la introducción completa era una sola oración de 63 palabras. Ese ritmo parejo es lo que más delata un texto generado o pasado por un parafraseador. El nombre del archivo ("revisado") sugiere que el texto ya había pasado por una revisión de ese tipo. Unir oraciones con comas sube la detección en vez de bajarla.
2. **Cadenas de subordinadas con los mismos conectores.** Hay 37 *which*, 19 *while*, 13 *since*, 13 *because* y 10 *although*. A eso se suman fórmulas fijas como *which is why*, *for that reason*, *so that*, *in light of*, *on that basis* e *in practical terms*. En promedio había 0,69 conectores de este tipo por oración.
3. **El plan en futuro, repetido.** *I will* / *I would* aparecían 47 veces, casi siempre encadenados con *so that* o *whenever*.
4. **Frases que anuncian en vez de decir.** Algunos ejemplos:
   - "Expressed as a formal positioning statement, my proposition reads as follows:"
   - "Figure 6 captures how each of them approaches…"
   - "with Table 8 summarizing the decision that resulted from each element"
   - "For the purpose of this assignment"
5. **Vocabulario comodín.** *relevant* aparece 19 veces y *evidence* 15. También abundan *practical*, *useful*, *specific* y *accurate*.
6. **Valoraciones cautelosas y equilibradas.** Por ejemplo: "can be described as cautiously positive", "should not be read as proof of", "although their value depends on…".
7. **Tablas redactadas como prosa.** Se marcaron las celdas escritas como frases completas, en las Tablas 4, 9, 10, A1 y C1. Las celdas cortas, como nombres, cifras y códigos, no se marcaron.

## 3. Qué se cambió

- **Reescritura de todos los párrafos marcados.** Las oraciones largas se partieron, se alternan oraciones cortas y largas, y se quitaron conectores, anuncios y palabras comodín. **No cambió ningún dato, cifra, cita ni referencia**, y tampoco el contenido de cada apartado.
- **Párrafos no marcados.** Se dejaron como estaban, porque Turnitin no los marcó. Solo se corrigieron los guiones partidos ("business-to- business", "U.S.- linked", "Canada- only").
- **Celdas de las Tablas 4, 9, 10, A1 y C1.** Pasaron a estilo de nota, más breve, con la misma información.
- **Texto propuesto para el About y el borrador de publicación.** Se reescribieron en primera persona, con un tono natural de LinkedIn.
- **Prompt de ChatGPT en el Apéndice C.** Se dejó igual, porque es el registro de lo que se usó. Solo cambió la frase que lo presenta.
- **Formato APA 7, trabajo de estudiante:**
  - Portada.
  - Número de página arriba a la derecha.
  - Times New Roman 12, doble espacio y márgenes de 1".
  - Títulos de nivel 1, 2 y 3. Los subtítulos por año (2021–2025), "Interpretation of the Historical Trend" y los de la revisión del perfil pasan a nivel 3, en negrita y cursiva.
  - Tablas con solo líneas horizontales y encabezado repetido si la tabla cambia de página.
  - Notas con *Note.* en cursiva.
  - Referencias con sangría francesa.
  - Apéndices en página nueva.

| Medida (prosa) | Entrega | v2 |
|---|---|---|
| Palabras | 5.439 | 5.245 |
| Oraciones | 122 | 233 |
| Largo medio de oración | 44,6 | 22,5 |
| Oraciones de más de 40 palabras | 75 | 23 (casi todas en párrafos de datos que no se marcaron) |
| Oraciones de menos de 15 palabras | 1 | 67 |
| Variación del largo de las oraciones (CV) | 0,34 | 0,60 |
| Conectores subordinantes por oración | 0,69 | 0,12 |
| "I will" / "I would" | 47 | 33 |
| relevant / evidence / specific / practical | 40 | 22 |
| **Texto marcado que sigue igual** (secuencias de 5 palabras) | 100 % | **19 %** |

El 19 % que queda corresponde sobre todo a cifras con su cita, nombres propios, el prompt del Apéndice C y celdas de tablas. Para el inglés no hay un simulador calibrado como el del trabajo en español, así que el porcentaje real solo lo dará Turnitin.

## 4. Cómo se hizo

El trabajo se reconstruyó a partir del PDF del informe, porque no estaba el Word original:
- `extraccion/` contiene los scripts que sacan el texto, el resaltado, las tablas, las referencias y los recortes de las figuras.
- `datos/` y `figuras/` guardan lo extraído.
- `contenido.py` tiene el texto reescrito.
- `construir.py` arma el Word.
- `comparar.py` calcula las medidas.

```
python construir.py "Trabajo LinkedIn revisado v2.docx"
python comparar.py "Trabajo LinkedIn revisado v2.docx"
```

## 5. Lo que debe revisar la autora

1. **Portada**: siguen los marcadores [Department, Institution], [Instructor Name] y [Assignment Due Date].
2. **Figuras**: los gráficos y las capturas de LinkedIn se recortaron del PDF. Si tiene los archivos originales, conviene reemplazarlos para ganar nitidez.
3. **Texto de las entradas propuestas para Garage y GLT y del About**: se reescribió la redacción. Antes de pegarlo en LinkedIn, debe confirmar que describe con exactitud lo que hizo.
4. Si tiene el Word original, se pueden aplicar los mismos cambios sobre ese archivo para conservar su formato.
