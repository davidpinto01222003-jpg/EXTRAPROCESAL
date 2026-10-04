# Revisión del trabajo de grado: ¿Cuáles son los procesos de enseñanza en la lectoescritura?

Isabela Escobar Marroquín, Licenciatura en Español e Inglés (UPB). Normas APA 7.ª edición.
Informe de partida: Turnitin, 28-09-2026, **34 % de IA** (17.710 palabras, 40 bloques marcados). No hay informe de similitud.

## 1. ¿Puede pasar?

**Así como está, no lo recomiendo.** Hay tres riesgos, de mayor a menor:

1. **Frases que parecen comentarios de un revisor (o de una IA) pegados en el texto.** En el análisis de la transcripción había frases como *"lo cual constituye una inconsistencia de identificación que debe revisarse en el documento original… Esta duplicación debe corregirse"*, *"Esto representa una oportunidad para fortalecer el trabajo… se debe señalar…"* y *"En el documento se plantea…"*, *"El trabajo señala expresamente…"*. Un jurado las detecta enseguida. Una de ellas además es falsa: dice que hay dos "Estudiante 8", pero la transcripción va del 1 al 10 sin repetir. **Corregido en la v2.**
2. **34 % de IA en Turnitin.** Muchas universidades piden revisión manual por encima de 20 %. Ese umbral es una referencia, no la regla de la UPB: confírmenlo con el programa. **Reescrito en la v2** (ver punto 3).
3. **APA y coherencia metodológica.** Hay problemas que un evaluador puede señalar aunque el % salga bien (ver punto 4). Algunos los corregí; otros requieren una decisión de la autora.

## 2. Qué patrones tuvo en cuenta Turnitin, comparados con los de Compilatio

Turnitin no marca palabras sueltas: marca **bloques enteros de prosa** (párrafos completos o varios seguidos). Comparando lo marcado con lo no marcado en este trabajo, y lo marcado por Compilatio en el trabajo de Suárez y Rivera, el detalle está en `PATRONES_TURNITIN_COMPILATIO.md`:

| Patrón | Turnitin | Compilatio |
|---|---|---|
| Conectores de apertura ("Asimismo", "Por consiguiente", "En este sentido", "Sin embargo") | **Pesa**: 7,5 vs 5,5 por mil palabras | Casi no aparecen |
| Verbos comodín ("permite", "evidencia", "constituye", "destaca", "resalta") | **Pesa**: 13,4 vs 10,4 | Neutro |
| Adjetivos de relleno ("fundamental", "esencial", "relevante", "significativo", "importante") | **Pesa**: 7,0 vs 4,1 | Neutro |
| Vocabulario genérico de revisión ("factores", "intervenciones", "brechas", "contextos", "enfoques", "fortalecer") | **Pesa** | Pesa el vocabulario genérico del tema |
| Primera persona ("planteamos", "abordamos") | **No protege**: aparece más en lo marcado | No protege |
| Dos puntos y punto y coma | No pesa | **Pesa**: 10,1 vs 6,6 |
| Citas con autor y año dentro de la frase | Protege un poco | **Protege**: 4,9 vs 9,5 |
| Datos concretos (respuestas textuales de los estudiantes, ejemplos) | **Protege**: casi nunca se marcan | Protege |
| Textos formulaicos (resumen, abstract, objetivos, dedicatoria con frases hechas) | **Se marcan casi siempre** | Se marcan |

En pocas palabras: **Turnitin castiga la prosa genérica y "redonda"**, con frases intercambiables, conectores de manual y adjetivos de relleno, sin datos ni ejemplos. **Compilatio castiga más la estructura**, con oraciones largas, dos puntos y punto y coma, y párrafos uniformes. Lo que sirve para los dos es lo mismo: concretar (ejemplos, cifras, quién dijo qué), quitar muletillas y variar el ritmo. La primera persona sola no basta.

## 3. Qué se cambió en la v2 (`Trabajo lectoescritura v2.docx`)

- **Reescritos 146 párrafos**: los que marcó Turnitin y algunos parcialmente marcados, para dar margen. Incluye dedicatoria, agradecimientos, resumen, abstract, introducción, preguntas, planteamiento, antecedentes, justificación, objetivos, marco teórico, resultados, análisis de la transcripción, discusión y conclusiones.
- **No se tocaron**: datos, respuestas de los estudiantes, pseudopalabras, citas textuales, transcripciones de la prueba, tablas ni la carta "Un alto en el camino" (solo su última línea).
- **Se quitaron** los comentarios tipo revisor y la frase falsa sobre el "Estudiante 8" duplicado.
- **Erratas corregidas**: "dyslexia continue" → *"dyslexia is still…"* (abstract), "descenlace", "El este trabajo", "Teniendo en cuento", "donde el cual", "Molina en (2017)", "cambios los curriculares", "se asemejad", "Transcipción".
- **APA corregido en el texto**: "Londoño (2018)" → **Morales Londoño (2018)**; "Gaviria y Lopera (2022)" → **Giraldo Gaviria y Caro Lopera (2022)**; año agregado a Aguirre-Medrano y González-López (2021) y a Duarte-Hernández y Pérez-Mendoza (2020) donde faltaba. Se quitó la negrita de citas y de términos dentro de párrafos reescritos.
- **Referencias**: se eliminó la entrada duplicada de Morales Londoño (2018).
- **Pregunta 2**: decía "¿Cómo se manifiestan *dichas* dificultades…?" sin haber nombrado ninguna dificultad antes. Quedó *"las dificultades de lectoescritura"*.

### Estimación

| Medida | Original | v2 |
|---|---|---|
| Texto que Turnitin marcó y sigue igual | 100 % | **18 %** (sobre todo títulos de artículos, pseudopalabras y frases de la prueba) |
| Si Turnitin solo volviera a marcar ese texto | 34 % | **~6 %** (cota baja) |
| Simulador de estilo calibrado con este informe | 35,2 % | 30,5 % (cota alta, modelo débil: AUC 0,61) |

El resultado real debería quedar entre esas dos cifras. **Solo Turnitin da el número definitivo.** Pásenlo de nuevo y súbanme el informe: con él recalibro el simulador y hago otra pasada sobre lo que siga marcado.

## 4. Pendientes que debe decidir la autora (no los cambié porque tocan el contenido)

1. **Metodología contradictoria con los resultados.**
   - "Tipo de estudio" dice *"se analiza detalladamente el desempeño lector de un **único participante**"*, pero la prueba se aplicó a **diez** estudiantes.
   - "Alcance" habla de *"las dificultades lectoras **del estudiante**"*, en singular.
   - Dice *"estudio de grupo focal/caso"*: hay que escoger uno. Un grupo focal es una entrevista grupal, y aquí no la hubo.
   - Dice que *"la única técnica… fue… PROLEC-R"*, pero los resultados describen una **prueba propia de pseudopalabras** con cinco actividades. Hay que aclarar si la prueba se basó en el PROLEC-R o si se aplicó el PROLEC-R completo.
2. **Faltan en la lista de referencias unas 25 obras citadas en el texto**: Duarte-Hernández y Pérez-Mendoza (2020), Medina (2020), Cameron (2021), Santana et al. (2021), Gutiérrez y Díez (2018), Snowling et al. (2020), Manchado (2009), Vargas Franco (2020), Molina (2017), Carlino et al. (2013), Zambrano y Aragón de Moreno (2015), McLuhan (1962, 1969), Birkerts (1994), Casillas Alvarado y Ramírez Martinell (2018), Nyathi et al. (2011) / Silva (2014), Vargas et al. (2019), Arnau (1995), Naciones Unidas (2015), UNESCO (2017), Save the Children (2018), Snow et al. (2018), Dickinson et al. (2019), Suárez-Álvarez y Fernández-Alonso (2018), Downer et al. (2017), González-Castro y Núñez (2016), Creswell (2014) e IBM (2013). Si son citas tomadas de los artículos reseñados, en APA 7 se escriben *"(Snow et al., 2018, como se citó en Domínguez Vázquez, 2023)"* y solo va a la lista la fuente que sí leyeron. **Núñez Peña (2011)** está en la lista y no se cita en el texto.
3. **IBM (2013)** aparece como fuente de *"la pandemia profundizó las brechas"*. Es imposible por la fecha: hay que cambiarla por una fuente sobre la pandemia (por ejemplo, López Rivas, 2024, que ya está en la lista).
4. **Referencias incompletas o mal formateadas**:
   - Faltan revista, volumen, páginas o DOI en Aguirre-Medrano y González-López (2021) y en Pisco-Román y Bailón-Panta (2023).
   - "Caamaño Tomás, Alejandro." y "López Rivas, Oscar Hugo." llevan el nombre completo; en APA va solo la inicial: *Caamaño Tomás, A.*
   - Tinta Aruquipa se cita como 2020 en el texto y como 2021 en la lista.
   - Falta DOI o URL en casi todas.
5. **Citas textuales sin comillas** (párrafos con "(p. 331)", "(p.2)", "(p.3)" de Pisco-Román y Bailón-Panta y de López Rivas): si son textuales y de menos de 40 palabras, van entre comillas dentro del párrafo. Si no lo son, hay que quitar el número de página.
6. **Estructura**: el análisis largo de la transcripción quedó bajo el título **"Anexos"**, después de las transcripciones. Debería ir en **Resultados** o en **Análisis de resultados**, y en Anexos solo las transcripciones. Los niveles de título también están desordenados: "Alcance", "Método" y "Técnica" están como Título 3, 4 y 5 aunque son del mismo nivel.
7. **Resumen**: habla casi solo de la revisión documental. Convendría agregar una o dos frases con lo que arrojó la prueba de pseudopalabras.
8. La portada de Turnitin dice **80 páginas**, y ustedes me dijeron 36. La diferencia probablemente sale de los anexos y las transcripciones; confirmen el límite de páginas del programa.
