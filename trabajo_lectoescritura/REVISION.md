# Revisión del trabajo de grado: ¿Cuáles son los procesos de enseñanza en la lectoescritura?

Isabela Escobar Marroquín, Licenciatura en Español e Inglés (UPB). Normas APA 7.ª edición.
Informe de partida: Turnitin, 28-09-2026, **34 % de IA**.

**Versión final: `Trabajo lectoescritura v6.docx`**: la v5 con formato APA 7 corregido y la discusión ampliada (sección 6). La v5 es la v4 (reescritura de estilo sobre la v3) ampliada con contenido nuevo. Ver `DIAGNOSTICO_ESTILO.md`. La v3 contiene las correcciones de contenido descritas abajo. Se genera con `python version3.py original.docx "Trabajo lectoescritura v3.docx"`, que aplica los cambios de `reescribir.py` (v2) y las correcciones de contenido de la v3.

## 1. Estado

| Aspecto | Original | v3 |
|---|---|---|
| Comentarios de revisor o IA pegados en el texto ("debe corregirse", "el documento señala…") | Sí | **No** |
| Metodología coherente con lo que se hizo (10 estudiantes, prueba propia) | No | **Sí** |
| Datos de los estudiantes verificados contra la transcripción | Con errores | **Sí** |
| Citas en el texto con referencia en la lista | Faltaban ~27 | **Todas** |
| Referencias APA completas (revista, volumen, páginas, DOI) | Incompletas | **Completas** (salvo dos, sección 5) |
| Preguntas de investigación respondidas en las conclusiones | 2 de 4 | **4 de 4** |
| Estructura (análisis en Anexos, niveles de título) | Desordenada | **Corregida** |
| Texto que Turnitin marcó y sigue igual | 100 % | **15 %** |

**Estimación de Turnitin para la v3: entre ~5 % y ~30 %.** La cota baja supone que Turnitin solo vuelve a marcar el 15 % del texto marcado que sigue igual, que son sobre todo títulos de artículos, términos técnicos y las frases de la prueba en el anexo. La cota alta la da el simulador de estilo, que es pesimista y poco preciso (AUC 0,61). El número real solo lo da Turnitin.

## 2. Correcciones de contenido en la v3

### Metodología
- **Tipo de estudio:** decía "grupo focal/caso" y "un único participante". Ahora dice **estudio de casos múltiples descriptivo con diez estudiantes**.
- **Participantes:** se agregó el apartado, con lo que el trabajo permite afirmar: diez estudiantes de básica, evaluados individualmente e identificados del 1 al 10.
- **Instrumento:** decía que "la única técnica fue el PROLEC-R", pero los resultados describen una prueba propia de cinco actividades. Ahora dice que es **una prueba diagnóstica de pseudopalabras diseñada para el estudio, con base en la tarea de pseudopalabras del PROLEC-R**. Las cinco actividades se listan, y lo mismo se ajustó en el marco teórico y en el enfoque.
- **Método:** se agregó la revisión documental (artículos de 2018 a 2024), que el resumen mencionaba pero la metodología no describía.
- **Alcance:** decía "del estudiante", en singular; ahora habla de los estudiantes.

### Datos de los diez estudiantes (verificados uno por uno contra las 22 tablas de la transcripción)

| Afirmación en el original | Lo que dicen los datos |
|---|---|
| "Trapi" aparece como error de lectura (estudiante 3) | No existe en ninguna transcripción. El estudiante 3 solo cambió "Zepa" y "Cuome" |
| "Lused" como error del estudiante 3 | "Lused" es la pseudopalabra original |
| El estudiante 4 "presentó modificaciones en lectura" | Leyó las 30 pseudopalabras sin cambios |
| El estudiante 5 "presentó transformaciones en lectura" | Cambió una sola ("Plerox") |
| "Tru-nal" "especialmente en los estudiantes 2, 3, 4, 5, 7, 8 y 9" | Lo dieron los diez |
| "“Pamil” por “Pamil”" | Era "“Pamil” por “Pamir”" |
| El estudiante 10 sin datos fonológicos | Respondió bien todas las preguntas fonológicas |
| Leían mejor "esas mismas palabras" dentro de las frases | Las pseudopalabras de las frases no eran las que alteraron al leer; "teniq" y "bapo" también se leyeron bien sueltas. Se matizó |
| Un estudiante "altera una misma pseudopalabra al leerla y al escribirla" | Las listas de lectura y de dictado eran distintas. Se reformuló como coincidencia en el tipo de error |
| Asociación grafema-fonema: "unir pseudopalabras con su sonido inicial" | La transcripción no registra respuestas; ahora se dice explícitamente |

Se agregaron datos concretos que antes faltaban:
- 9 de 10 estudiantes cambiaron alguna pseudopalabra al leer, con un rango de 0 a 9 cambios.
- Los 10 transformaron alguna al dictado.
- Sonidos: 5 de 10 aislaron bien el inicial de "gropel", 4 el final de "finod" y 4 el interno de "lumep".
- Solo 3 de 10 acertaron todas las preguntas fonológicas.
- En las frases, 8 de 10 dijeron "un bapo" en lugar de "el bapo".

### Citas y referencias
- **Citas no verificables eliminadas**, con la frase reescrita para no afirmar lo que esas fuentes no respaldan: Medina (2020), Cameron (2021), Santana et al. (2021), Snow et al. (2018), Dickinson et al. (2019), Downer et al. (2017), Suárez-Álvarez y Fernández-Alonso (2018), González-Castro y Núñez (2016), Creswell (2014), UNESCO (2017), Save the Children (2018), Casillas Alvarado y Ramírez Martinell (2018).
- **IBM (2013)** como fuente de los efectos de la pandemia (imposible por la fecha) → **López Rivas (2024)**.
- **Duarte-Hernández y Pérez-Mendoza (2020)** existe, pero es un estudio de lateralidad en niños de 2 a 5 años con el test de Harris. Se mantiene donde habla de lateralidad. Donde se le atribuían las confusiones visomotoras y los errores de lectura, la cita se pasó a Aguirre-Medrano y González-López (2021), que sí trata eso.
- **Gutiérrez y Díez (2018)** estaba mal usado como respaldo del "estudio de caso". Se quitó de la metodología y se dejó en la crítica al antecedente, con el nombre completo (Gutiérrez Fresneda y Díez Mediavilla) y su tema real.
- **Citas secundarias en formato APA ("como se citó en")**: Manchado, Vargas Franco, Molina, Carlino et al. y Zambrano y Aragón de Moreno, vía Giraldo Gaviria y Caro Lopera; McLuhan y Birkerts, vía Caamaño Tomás; Nyathi et al., vía Morales Londoño; Vargas et al., vía Rivera Cintrón y Batiz Cartagena; Arnau, vía Núñez Peña (que estaba en la lista sin citarse).
- **Citas textuales sin comillas** (Pisco-Román y Bailón-Panta, pp. 331, 333 y 335; López Rivas, pp. 2 y 3; Rivera Cintrón y Batiz Cartagena, pp. 8 y 9) → **paráfrasis con número de página**.
- **Apellidos**: "Tomás" → **Caamaño Tomás**; "Londoño" → **Morales Londoño**; "Gaviria y Lopera" → **Giraldo Gaviria y Caro Lopera**; "González-López, C. M." → **González-López, M.** ("Dra. C" es un título, no una inicial).
- **"Este autor analiza…"** (López Rivas) no nombraba al autor; corregido.
- **Lista de referencias** reconstruida en orden alfabético, con cursivas APA, y titulada **"Referencias"**. Se agregaron revista, volumen, número, páginas y DOI verificados de:
  - Aguirre-Medrano y González-López (*Santiago*, 156)
  - Pisco-Román y Bailón-Panta (*593 Digital Publisher CEIT*, 8(1-1), DOI)
  - Giraldo Gaviria y Caro Lopera (DOI)
  - Morales Londoño (DOI)
  - López Rivas (DOI)
  - Tinta Aruquipa (año corregido a **2020**)
  - Se agregaron Duarte-Hernández y Pérez-Mendoza (2020), Gutiérrez Fresneda y Díez Mediavilla (2018), Snowling et al. (2020) y Naciones Unidas (2015).

### Estructura y coherencia
- El análisis largo, que estaba bajo el título **"Anexos"**, ahora se llama **"Análisis de resultados"**. La transcripción pasó a **"Anexo A. Transcripción de la prueba"**, después de las referencias, como pide APA.
- Niveles de título unificados: Metodología pasa a Título 1; Alcance, Método y Técnica a Título 2; subtemas del marco teórico a Título 2; antecedentes, con el título de cada artículo y su cita, a Título 3. Los títulos vacíos dejaron de ser títulos.
- Se eliminó la lista de "siete hallazgos", que estaba repetida en Resultados, en Análisis y en Conclusiones.
- **Conclusiones:**
  - Nuevo párrafo que responde las preguntas 1 y 3 (qué estrategias funcionan) con los estudios revisados.
  - Respuesta a la pregunta 4 (relación entre los errores de lectura y de escritura).
  - Explicación de cómo se cumplió cada objetivo específico.
- **Resumen y Abstract**: incluyen ahora los resultados de la prueba (239 y 207 palabras; el límite APA es 250).
- **Siglas**: se agregaron ENS, MEN, ONU, PISA, PROLEC-R, PSC y WISC IV.
- **Carta "Un alto en el camino"**: solo ortografía ("estás", "tú", "sé", "para dónde", "sino también").
- El Word **actualiza la tabla de contenido al abrirlo**: hay que aceptar el aviso.

## 3. Patrones de Turnitin vs. Compilatio

El detalle está en `PATRONES_TURNITIN_COMPILATIO.md`. En resumen:
- **Turnitin** castiga la prosa genérica: conectores de manual, verbos comodín ("permite", "evidencia", "constituye"), adjetivos de relleno ("fundamental", "esencial") y textos formulaicos.
- **Compilatio** castiga más la estructura: oraciones largas, dos puntos y punto y coma.
- **A los dos los baja** lo concreto: cifras, ejemplos y respuestas textuales. Por eso la v3 metió los datos reales de los estudiantes en resultados, análisis, discusión y conclusiones.

## 4. Lo que solo la autora puede completar o confirmar

1. **Participantes**: grado, edad, institución, ciudad y fecha de aplicación. También el **consentimiento informado** de los acudientes, porque son menores: si existe, conviene mencionarlo en Participantes y adjuntarlo como Anexo B.
2. **Estudiantes 7 y 8**: sus respuestas de lectura son **idénticas palabra por palabra**, incluidos los nueve errores. Hay que confirmar que no fue un error al copiar la tabla.
3. **"Zolid"**: los diez estudiantes leyeron "Zolid", pero esa palabra no está en la lista de estímulos (que tiene "Solip"). Confirmen cuál fue la palabra presentada.
4. **Listas de dictado**: a los estudiantes 1 y 2 se les dictó una lista y a los demás otra (la tabla del dictado tiene dos bloques). Confirmen que fue así.
5. **Rivera Cintrón y Batiz Cartagena (2024)**: no pude verificar páginas ni DOI. Si los tienen, agréguenlos. Lo mismo con el DOI de Gutiérrez Fresneda y Díez Mediavilla (2018).
6. **Duarte-Hernández y Pérez-Mendoza (2020)**: verifiquen que el dato de "lateralidad definida hacia los 6 años" sale de ese estudio. Si no, cambien la cita.
7. **Arnau (1995)**: verifiquen que la definición de diseño cuasiexperimental la tomaron de Núñez Peña (2011). Si la tomaron de otra fuente, citen esa.

## 6. Formato APA 7 (v6)

Se genera con `python version6.py "Trabajo lectoescritura v5.docx" "Trabajo lectoescritura v6.docx"`. **La organización no cambia**: las secciones, los niveles de título y el lugar de cada elemento son los mismos de la v5. Solo cambia el formato.

| Aspecto | v5 | v6 |
|---|---|---|
| Márgenes | 2,5 cm | **2,54 cm** en todas las secciones |
| Interlineado del cuerpo | Mezcla de 1,5, 1,08, sencillo y doble | **Doble**, sin espacio antes ni después |
| Alineación | Justificada | **Izquierda** |
| Sangría de primera línea | 1,25 cm | **1,27 cm**. Sin sangría en el resumen, el abstract, las tablas, los rótulos y las notas |
| Separación con párrafos vacíos | 142 | **0**. Los saltos de página se conservan como "salto de página anterior" y cada título de nivel 1 empieza en página nueva |
| Títulos | Fuente de tema, interlineado 1,5 | Times New Roman 12 en negro. Nivel 1 centrado en negrita, nivel 2 a la izquierda en negrita, nivel 3 en negrita y cursiva, sin punto final |
| Tabla de contenido | Texto viejo del original ("Transcipción de prueba.", "Anexos", "Citas y referencias") | **Regenerada** con los títulos y las páginas actuales, con enlaces |
| Numeración de tablas | La Tabla 2 se citaba antes que la 1 | **Numeradas en el orden en que se citan** |
| Tablas | Centradas. Las del anexo tenían cuadrícula completa, sangría y doble espacio | **Alineadas al margen, sin sangría**, ancho completo, solo líneas horizontales, 12 pt con interlineado sencillo, filas que no se parten entre páginas |
| Tablas del anexo | Sin número ni título ("Lo que dijo:", "Escribió:") | **Tabla A1 a A22**, con número en negrita y título en cursiva, en el mismo lugar |
| Fotografías | 24 imágenes flotantes, sin número ni título, encima del texto de "Análisis de resultados" | **En el mismo lugar**, en línea, como **Figuras 1 a 3** (aplicación, primera y segunda página de las hojas de respuesta) con título y nota. El Procedimiento las cita |
| Espacios | Dobles en el anexo, "( q  )", tabuladores para alinear | Corregidos, sin cambiar ninguna respuesta |
| Referencias | Interlineado 1,5 | **Doble**, con sangría francesa de 1,27 cm |

Pendiente para la autora:
1. **Portada**: dice "Danny Jean Paul Mejía,  en Literatura", con dos espacios. Parece que falta el título del orientador (por ejemplo "Magíster"). No lo cambié porque está en un control de contenido de la plantilla de la UPB.
2. **Fotografías de menores**: en las figuras se ven brazos y manos, no rostros. Confirmen que el consentimiento de los acudientes cubre las fotos.
3. **Actividad de unir**: las hojas de respuesta (Figura 3) muestran las líneas con que los estudiantes unieron cada pseudopalabra con su sonido. El trabajo dice que esa parte "no quedó registrada en la transcripción y no se analizó". Pueden agregarla desde las hojas si quieren.
4. **Grado y edad** de los participantes: la nueva discusión explica por qué hacen falta para valorar los errores en sílabas trabadas.
5. Las páginas del índice se calcularon con LibreOffice; si Word muestra un aviso para actualizar campos al abrir, acéptenlo y quedarán exactas.
