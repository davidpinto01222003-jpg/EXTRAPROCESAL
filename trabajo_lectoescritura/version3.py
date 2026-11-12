# Versión 3: parte de los cambios de reescribir.py (v2) y corrige el contenido
# para que el trabajo quede coherente: metodología, datos de los estudiantes,
# citas y referencias APA, estructura (análisis y anexos) y preguntas de investigación.
# Uso: python version3.py original.docx "Trabajo lectoescritura v3.docx"
import sys
from copy import deepcopy

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import reescribir as R

FULL, LEAD, ERRATAS = dict(R.FULL), dict(R.LEAD), {k: list(v) for k, v in R.ERRATAS.items()}
BORRAR = list(R.BORRAR)
ESTILO = {}  # índice -> estilo de título

GG = "como se citó en Giraldo Gaviria & Caro Lopera, 2022"
CT = "como se citó en Caamaño Tomás, 2021"

# ---------------------------------------------------------------- Resumen y Abstract
FULL[94] = "La revisión muestra que la dislexia sigue siendo poco comprendida en las escuelas y que la lateralidad cruzada, una conciencia fonológica débil y la poca formación de los docentes en este tema hacen más difícil atender a tiempo los problemas de lectura. La pandemia, además, amplió las brechas en los contextos vulnerables. En la prueba, nueve de los diez estudiantes cambiaron al menos una pseudopalabra al leer, todos lo hicieron al escribir al dictado y la conciencia fonológica fue desigual: aislar un sonido les costó más que dividir en sílabas o reconocer terminaciones semejantes, dos tareas que los diez resolvieron bien."
FULL[95] = "Los estudios revisados reportan buenos resultados con intervenciones multisensoriales, actividades lúdicas y enfoques inclusivos. Concluimos que hacen falta diagnósticos tempranos, estrategias pensadas para cada contexto, capacitación docente continua y políticas que aseguren a todos el acceso a la cultura escrita."
FULL[101] = "The review shows that dyslexia is still poorly understood in schools, and that crossed laterality, weak phonological awareness and limited teacher preparation make it harder to address reading problems in time. The pandemic also widened the gaps in vulnerable settings. In the test, nine of the ten students changed at least one pseudoword when reading, all of them did so when writing from dictation, and phonological awareness was uneven: isolating a single sound was harder than splitting a word into syllables or recognizing similar endings, two tasks that all ten students solved correctly."
FULL[102] = "The studies reviewed report good results with multisensory interventions, play-based activities and inclusive approaches. We conclude that early diagnosis, context-sensitive strategies, ongoing teacher training and policies that guarantee everyone access to written culture are needed."

# ---------------------------------------------------------------- Carta "Un alto en el camino" (solo ortografía)
ERRATAS[159] = [("que nos has tenido", "que has tenido"), ("que estas creando", "que estás creando")]
ERRATAS[160] = [("estas brillando", "estás brillando"), ("para donde vas", "para dónde vas"), ("evoca que tu eres", "evoca que tú eres")]
ERRATAS[161] = [("creciendo no solo mental si no que personalmente también", "creciendo no solo mentalmente, sino también como persona")]
ERRATAS[163] = [("y si algunas olvidas algo", "y si alguna vez olvidas algo")]
ERRATAS[164] = [("se que lograrás", "sé que lograrás")]

# ---------------------------------------------------------------- Antecedentes: títulos de cada artículo reseñado
TITULOS = {
    175: "Problemas de dislexia en educación básica: una problemática para la lectoescritura (Aguirre-Medrano & González-López, 2021)",
    181: "Alfabetización académica: una alternativa para repensar la formación inicial docente en las escuelas normales superiores de Colombia (Giraldo Gaviria & Caro Lopera, 2022)",
    189: "Consideraciones acerca de la educación, lectoescritura y aprendizaje en los tiempos de pandemia (Caamaño Tomás, 2021)",
    195: "La lectoescritura como elemento fundamental en el proceso de enseñanza aprendizaje de los estudiantes de Básica Media (Pisco-Román & Bailón-Panta, 2023)",
    201: "La calidad de los procesos de enseñanza de la lectoescritura en Colombia: entre concepciones y prácticas (Morales Londoño, 2018)",
    205: "Neuroaprendizaje basado en la lectoescritura y estrategias lúdicas en el currículo Artes del Lenguaje (Rivera Cintrón & Batiz Cartagena, 2024)",
    210: "Proceso de enseñanza aprendizaje de la escritura a partir de la lectura de la realidad (Tinta Aruquipa, 2020)",
    213: "Enseñanza de la lectoescritura en la escuela primaria en tiempos de pandemia. Estudios de casos (López Rivas, 2024)",
    219: "Una propuesta de intervención para el desarrollo de las habilidades de lectoescritura en alumnado de primer ciclo de primaria con riesgo de exclusión social (Domínguez Vázquez, 2023)",
}
for i, t in TITULOS.items():
    FULL[i] = t
    ESTILO[i] = "Heading 3"

# Aguirre-Medrano y González-López: se quitan citas que no están en la lista y no se pudieron verificar
FULL[177] = "En el caso estudiado, la paciente tiene una lateralidad cruzada mal afirmada, disgrafía fonológica y dificultades para identificar letras y signos de puntuación. Duarte-Hernández y Pérez-Mendoza (2020), que evaluaron con el test de Harris a niños de 2 a 5 años, señalan que la lateralidad suele quedar definida hacia los 6 años. Aunque su coeficiente intelectual está por encima del promedio (CI = 115), la niña lee y escribe con lentitud. Para las autoras, esto indica que la lateralidad pesa en el desarrollo de la lectura y la escritura de los niños con dislexia, y que la escuela debería prestarle más atención."
FULL[178] = "El estudio recomienda intervenciones psicopedagógicas especializadas que trabajen la conciencia fonológica, la discriminación auditiva y la visomotricidad, porque eso podría mejorar tanto la escritura como la lectura. Para las autoras, trabajar esas áreas es el camino para que los niños con dislexia superen sus dificultades."
FULL[179] = "Desde una mirada crítica, el análisis del caso es detallado, pero se limita a una sola paciente, lo que impide generalizar los hallazgos. Además, aunque se usaron pruebas diagnósticas reconocidas, no hubo seguimiento en el tiempo para medir el efecto de las intervenciones en la lectura y la escritura, algo importante si se tiene en cuenta que la conciencia fonológica y la escritura se desarrollan por etapas (Gutiérrez Fresneda & Díez Mediavilla, 2018). Tampoco se discute cómo influyen la escuela y la familia en la dislexia, un aspecto que podría ser decisivo en la evolución de estos trastornos (Snowling et al., 2020)."

# Giraldo Gaviria y Caro Lopera: los estudios que ellos revisan se citan como fuente secundaria
ERRATAS[183] = [("Manchado (2009)", f"Manchado (2009, {GG})")]
ERRATAS[184] = [("Vargas Franco (2020)", f"Vargas Franco (2020, {GG})")]
ERRATAS[185] = [("Molina en (2017)", f"Molina (2017, {GG})"), ("Carlino, Iglesias y Laxalt (2013)", f"Carlino et al. (2013, {GG})")]
ERRATAS[186] = [("cambios los curriculares", "cambios curriculares"), ("Zambrano y Aragón de Moreno (2015)", f"Zambrano y Aragón de Moreno (2015, {GG})")]
ERRATAS[187] = [("creo que el artículo subraya", "consideramos que el artículo subraya"), ("Sin embargo, esta no profundiza", "Sin embargo, no profundiza")]

# Caamaño Tomás: su apellido completo y McLuhan/Birkerts como fuente secundaria
FULL[190] = f"El autor plantea cómo la pandemia de COVID-19 transformó radicalmente la educación, impulsó la enseñanza en línea y modificó los procesos de lectoescritura y aprendizaje. Para explicarlo establece un paralelo con La galaxia Gutenberg: génesis del Homo Typographicus, de Marshall McLuhan (1962, {CT}), obra en la que el teórico canadiense identifica tres momentos clave en la transmisión del conocimiento: el alfabeto fonético, la imprenta y la era electrónica."
FULL[191] = "Según el ensayo, adaptarse a la educación digital no dependerá tanto de los planes de estudio o de las políticas educativas como de la preparación técnica de estudiantes y docentes. El autor describe las barreras tecnológicas y pedagógicas de ese tránsito y muestra cómo la enseñanza en línea cambió la dinámica de las clases al llevar la lectura y la escritura a entornos digitales, algo que, a su juicio, pide enfoques pedagógicos nuevos."
FULL[192] = f"Caamaño Tomás lee el impacto de la pandemia con los lentes de McLuhan, para quien la aldea global es un espacio donde la tecnología moldea la percepción y el aprendizaje. Según McLuhan, los entornos digitales devuelven a la comunicación rasgos de la oralidad: al reintroducir lo auditivo y lo visual, se parecen a las prácticas comunicativas anteriores a la escritura fonética (McLuhan, 1962, {CT}). El autor advierte, de todos modos, que con la virtualidad se perdieron elementos pragmáticos propios de la clase presencial, como los gestos y la interacción en el aula."
FULL[193] = f"El ensayo también se fija en cómo el aprendizaje se volvió más individual en los entornos digitales. El encierro y la dependencia de las clases en línea rompieron en parte la relación entre docentes y alumnos, y con ello se resintieron la colaboración y la participación. McLuhan ya advertía que los medios electrónicos no solo amplían el acceso a la información, sino que cambian la forma en que una cultura se percibe y se transmite, y eso obliga a repensar la educación (McLuhan, 1969, {CT})."
FULL[194] = f"Al cierre, Caamaño Tomás piensa en el futuro de la educación y de la lectoescritura en la era digital y retoma a Sven Birkerts (1994, {CT}), autor de Elegía a Gutenberg, quien advierte que la tecnología puede transformar la conciencia humana y la cultura. El ensayo no propone soluciones definitivas, pero sí insiste en desarrollar competencias de lectura y escritura para los entornos digitales y en enseñar a usar la tecnología con sentido crítico."

# Pisco-Román y Bailón-Panta: las citas sin comillas pasan a paráfrasis con página
FULL[197] = "Los autores parten de que la lectura y la escritura están entre los procesos cognitivos y sociales más complejos que desarrolla el ser humano, y de que ambos ocurren al mismo tiempo (Pisco-Román & Bailón-Panta, 2023, p. 331)."
FULL[198] = "Por eso consideran la lectoescritura la base del aprendizaje escolar: en ella se desarrollan habilidades cognitivas, comunicativas y lingüísticas (Pisco-Román & Bailón-Panta, 2023, p. 333)."
FULL[199] = "Según ellos, quien lee y escribe bien puede organizar sus ideas, comprender textos complejos, pensar de manera crítica y expresarse con claridad y coherencia (Pisco-Román & Bailón-Panta, 2023, p. 335)."
# Morales Londoño
FULL[203] = "Para hablar de calidad, el autor retoma a Nyathi et al. (2011, como se citó en Morales Londoño, 2018), para quienes la calidad de la educación es un concepto multidimensional, multinivel y dinámico, ligado al contexto del modelo educativo, a la misión y los objetivos de cada institución y a los estándares propios de cada programa, disciplina, institución o sistema."
# Rivera Cintrón y Batiz Cartagena
FULL[207] = "Los autores retoman a Vargas et al. (2019, como se citó en Rivera Cintrón & Batiz Cartagena, 2024, p. 8), para quienes la lectura y la escritura son una muestra de conectividad intelectual y neuronal, porque están entre los aprendizajes más complejos que realizan las personas."
FULL[208] = "Entre sus hallazgos destaca la recomendación de fortalecer la lectoescritura con estrategias lúdicas, que los autores consideran el camino más agradable y divertido para que un niño aprenda (Rivera Cintrón & Batiz Cartagena, 2024, p. 9)."
# Tinta Aruquipa
ERRATAS[211] = [("El artículo de Moisés Rubén Tinta Aruquipa (2020) aborda", "Tinta Aruquipa (2020) aborda")]
# López Rivas: faltaba el autor y las citas no tenían comillas
FULL[214] = "López Rivas (2024) analiza cómo se enseñó a leer y escribir durante la pandemia en Guatemala, a partir de estudios de caso en contextos urbanos y rurales. El estudio muestra las desigualdades educativas, la falta de acceso a la tecnología y el peso que tuvieron las familias."
FULL[215] = "Según el autor, la pandemia de COVID-19 desordenó la organización de las clases y afectó los procesos de enseñanza y aprendizaje. Uno de los más golpeados fue la enseñanza de la lectoescritura, porque exige un acompañamiento cercano del docente (López Rivas, 2024, p. 2)."
FULL[216] = "Las escuelas tuvieron que pasar a la modalidad a distancia, y eso obligó a los docentes a inventar estrategias pedagógicas nuevas para llegar a sus estudiantes y asegurar que siguieran aprendiendo (López Rivas, 2024, p. 3)."
FULL[217] = "Uno de los hallazgos principales es que, pese a las dificultades, los docentes mostraron resiliencia y creatividad para mantener el vínculo con sus estudiantes y buscar alternativas para seguir trabajando la lectura y la escritura (López Rivas, 2024, p. 3)."
# Domínguez Vázquez: se quitan citas que no están en la lista y no se pudieron verificar
FULL[220] = "Domínguez Vázquez (2023) presenta una intervención psicopedagógica para mejorar la lectoescritura de estudiantes de segundo de primaria en riesgo de exclusión social. Trabaja con un enfoque de investigación-acción y un diseño cuasiexperimental, es decir, uno en el que los participantes no se asignan a los grupos al azar (Arnau, 1995, como se citó en Núñez Peña, 2011). La autora reporta mejoras importantes en comprensión lectora, escritura, conciencia fonológica y motivación hacia la lectura y la escritura."
FULL[221] = "En lo teórico, el estudio se enmarca en el Objetivo de Desarrollo Sostenible de la ONU sobre educación inclusiva y de calidad (Naciones Unidas, 2015), y responde a una preocupación más amplia: cómo mejorar la alfabetización en entornos desfavorecidos."
FULL[222] = "Lo más interesante es cómo está armado el programa. Incluye actividades multisensoriales y lúdicas, como lectura interactiva, reconocimiento de fonemas y escritura creativa. Esa mezcla recoge enfoques actuales de alfabetización inicial, y nos parece uno de los aciertos del trabajo."
FULL[223] = "El artículo no oculta sus dificultades, y la mayor fue el alto ausentismo en el grupo intervenido. Es un problema serio, porque un estudiante que falta con frecuencia pierde la continuidad que necesita para consolidar competencias básicas. Como solución, la autora propone contar con personal especializado, entre ellos profesores de servicios a la comunidad (PSC)."
FULL[224] = "La investigación mezcla observación cualitativa con mediciones cuantitativas del progreso en cinco áreas: comprensión lectora, escritura, reconocimiento de palabras, conciencia fonológica y motivación. Nos parece un acierto medir a la vez lo cognitivo y lo afectivo, porque en la alfabetización infantil ambos van de la mano."

# ---------------------------------------------------------------- Marco teórico
for i in (246, 252, 255, 260, 263, 265, 268):
    ESTILO[i] = "Heading 1" if i == 246 else "Heading 2"
FULL[248] = "Entendemos la lectoescritura como un proceso cognitivo y pedagógico complejo. Leer no es solo reconocer letras: el lector decodifica, comprende, sostiene en la memoria lo que va leyendo y construye un significado. En ese proceso la conciencia fonológica tiene un papel central, porque es lo que le permite al niño separar los sonidos del habla y unirlos con las letras que los representan. Aguirre-Medrano y González-López (2021) muestran que, cuando esta habilidad falla, como ocurre en la dislexia o en los problemas de discriminación auditiva, aparecen errores típicos de la ruta fonológica, por ejemplo confundir la b con la d. Para observar esa ruta por separado sirven las pseudopalabras. Una palabra inventada como “trunal” no significa nada, así que el lector no la puede reconocer de memoria y tiene que leerla letra por letra, aplicando las reglas de conversión grafema-fonema."
FULL[249] = "A esa mirada cognitiva le sumamos otra, pedagógica y de contexto. Domínguez Vázquez (2023) y López Rivas (2024) muestran que la alfabetización depende mucho del entorno: de si el niño vive en condiciones vulnerables, de si tuvo conexión durante la pandemia y de cuántos libros hay a su alcance. Por eso hacen falta métodos flexibles, ajustados a cada lugar. Los docentes pesan igual. Morales Londoño (2018) y Giraldo Gaviria y Caro Lopera (2022) critican las prácticas tradicionales y piden que los maestros reciban una mejor formación en alfabetización académica, para que puedan diseñar intervenciones a la medida. Con estas dos miradas armamos el marco del trabajo. La idea que lo atraviesa es sencilla: primero hay que saber con precisión qué le cuesta al estudiante, y una prueba de pseudopalabras ayuda a eso; después vienen las estrategias, que deben responder tanto a sus necesidades como a su contexto."
FULL[257] = "Pruebas como el PROLEC-R y las de discriminación fonológica permiten detectar los errores típicos de la ruta fonológica: omisiones, sustituciones, inversiones y segmentaciones incorrectas (Aguirre-Medrano & González-López, 2021)."
FULL[266] = "Aprender a leer no depende solo de procesos internos. El contexto familiar, las oportunidades de acceder a textos, las metodologías de los docentes y las condiciones socioculturales influyen, y mucho. La pandemia amplió las brechas, sobre todo para los estudiantes con poco acceso a recursos tecnológicos, a quienes les costó más sostener su proceso de alfabetización (López Rivas, 2024)."
FULL[270] = "En la misma línea, Aguirre-Medrano y González-López (2021) relacionan la lateralidad mal afirmada con confusiones entre grafemas como b-d y p-q. Estos errores, junto con omisiones, sustituciones e inversiones, suelen salir a la luz cuando el estudiante tiene que leer algo que no puede resolver con la memoria visual o léxica."
FULL[271] = "La batería PROLEC-R, en la que nos basamos para diseñar la prueba, incluye la lectura de pseudopalabras justamente para evaluar esta ruta y observar qué tanto domina el estudiante las reglas fonológicas y con cuánta precisión decodifica."

# ---------------------------------------------------------------- Metodología
ESTILO.update({285: "Heading 1", 287: "Heading 2", 289: "Heading 2", 291: "Heading 2", 293: "Heading 2", 295: "Heading 2"})
FULL[285] = "Metodología"
FULL[291] = "Alcance de la investigación"
FULL[293] = "Método"
FULL[295] = "Técnica e instrumento de recolección de datos"
FULL[288] = "Como el estudio buscaba analizar a fondo las dificultades de lectoescritura de diez estudiantes, optamos por un enfoque cualitativo. Nos interesaba comprender cómo lee y escribe cada uno y qué tipo de errores comete, e interpretar los resultados a la luz de sus condiciones particulares."
FULL[290] = "El trabajo es un estudio de casos múltiples de tipo descriptivo: analizamos por separado el desempeño de cada uno de los diez estudiantes y luego comparamos los resultados entre sí. Este diseño no busca generalizar. Lo que ofrece es información detallada sobre procesos concretos de lectura y escritura, y eso encaja con el carácter diagnóstico del trabajo y con sus objetivos."
FULL[292] = "El alcance es descriptivo. Buscamos caracterizar las dificultades de lectura y escritura de los estudiantes a partir de los errores, patrones y respuestas registrados en las cinco actividades de la prueba. No pretendemos explicar causas ni establecer relaciones estadísticas, sino describir cómo funcionan la ruta fonológica, la decodificación y la correspondencia grafema-fonema en cada caso."
FULL[294] = "El método fue evaluativo-descriptivo y tuvo dos partes. La primera fue una revisión documental de artículos publicados entre 2018 y 2024 en América Latina, España y Estados Unidos sobre dislexia, lateralidad, formación docente, pandemia y estrategias de enseñanza de la lectoescritura; sus resultados se presentan en los antecedentes y en el marco teórico. La segunda fue la aplicación de una prueba diagnóstica de pseudopalabras que diseñamos para este estudio, tomando como base la tarea de lectura de pseudopalabras de la batería PROLEC-R. Las respuestas se analizaron con los criterios de la propia prueba y con los referentes teóricos del trabajo."
FULL[296] = "La técnica de recolección fue la evaluación individual. A cada estudiante se le aplicó la prueba diagnóstica de pseudopalabras, organizada en cinco actividades:"
FULL[297] = "lectura en voz alta de pseudopalabras; escritura al dictado de pseudopalabras; comprensión y conciencia fonológica (sonido inicial, final e interno, terminaciones semejantes y segmentación silábica); fluidez lectora con pseudopalabras en frases, y asociación grafema-fonema."
FULL[298] = "Las respuestas de cada estudiante se registraron por escrito (Anexo A) y después se analizaron una por una, comparándolas con los estímulos originales."
FULL[299] = "Usamos pseudopalabras porque eliminan el apoyo del vocabulario memorizado y obligan a aplicar de forma estricta las reglas grafema-fonema."
NUEVOS_TRAS = {  # índice -> [(texto, estilo o None)] que se insertan después de ese párrafo
    290: [("Participantes", "Heading 2"),
          ("Participaron diez estudiantes de educación básica. A cada uno se le aplicó la prueba de forma individual, y en el trabajo se identifican con números, del Estudiante 1 al Estudiante 10, tal como aparecen en la transcripción (Anexo A).", None)],
}

# ---------------------------------------------------------------- Resultados (datos verificados contra las tablas de la transcripción)
ESTILO[315] = "Heading 1"
FULL[315] = "Resultados"
FULL[318] = "Lectura de pseudopalabras: En la primera actividad los estudiantes leyeron en voz alta una lista de treinta pseudopalabras. Nueve de los diez cambiaron al menos una. Entre las modificaciones aparecen “Zepa”, “Zeman” y “Zapa” por “Zepan”, “Trampin” por “Trapin”, “Digan” por “Dijan”, “Blused” por “Lused”, “Zolil” por “Zonil”, “Flanit” por “Franit” y “Florir” por “Folir”. Estos cambios muestran que, ante estímulos desconocidos, no todos los estudiantes logran una correspondencia precisa entre los grafemas que ven y los sonidos que deben producir."
FULL[320] = "Escritura al dictado de pseudopalabras: En la segunda actividad, los estudiantes debían escuchar pseudopalabras y escribirlas. Los diez escribieron algunas de forma distinta al estímulo original. Se observan producciones como “Lisa” por “Plisa”, “Bieco” por “Bieclo”, “Grac” por “Gruac”, “Tranflopiar” por “Tramkopliar”, “Caudre” por “Cuadre”, “Fragno” por “Fradno” y “Pamil” por “Pamir”, entre otras modificaciones en la composición de los estímulos."
FULL[324] = "Para el sonido inicial de “gropel”, cinco estudiantes dijeron “G” y otros cinco “Gro”, es decir, un segmento más largo. Para el sonido final de “finod”, solo cuatro respondieron “D”; los demás dijeron “Nod”, “Not”, “Od” o incluso “F”, que es el sonido inicial. Con el sonido interno de “lumep” pasó algo parecido: cuatro respondieron “M” y el resto dio respuestas como “Met”, “Ume”, “Me” y “Um”."
FULL[326] = "En contraste, la segmentación silábica fue muy estable: los diez estudiantes dividieron “trunal” como “Tru-nal”. También fue consistente el reconocimiento de terminaciones: todos señalaron “zopar” y “perar” como las pseudopalabras que terminan igual."
FULL[328] = "Fluidez lectora en frases: La cuarta actividad consistió en leer pseudopalabras dentro de oraciones: “El drumo saltó sobre la mesa”, “María compró un teniq azul” y “El perro jugaba con el bapo en el parque”. Todos los estudiantes pronunciaron bien las pseudopalabras dentro de las frases. El cambio más frecuente estuvo en el artículo: ocho de los diez dijeron “un bapo” en lugar de “el bapo”. El estudiante 1, además, omitió el artículo inicial de la primera frase y cambió “El perro” por “Mi perro”."
FULL[330] = "Asociación grafema–fonema: La quinta actividad buscaba ver cómo relacionan los estudiantes las letras con los sonidos iniciales. Cinco estudiantes (1, 2, 5, 8 y 10) identificaron el sonido inicial de “zunel” como /Z/; los otros cinco respondieron con la sílaba “Zu”. Con “quarim” la variación fue mayor: solo los estudiantes 8 y 10 respondieron “q”; seis dieron segmentos como “qu”, “qua” o “cua”, y dos se alejaron más, con “au” (estudiante 7) y “ua” (estudiante 9). La última parte de la actividad, unir pseudopalabras con su sonido inicial, no quedó registrada en la transcripción y por eso no se analizó."
FULL[333] = "Los registros también muestran diferencias entre los estudiantes. El estudiante 1 cambió cinco pseudopalabras al leer y varias al escribir; en las tareas fonológicas acertó todo menos el sonido final de “finod”, para el que respondió “F”. El estudiante 2 también cambió cinco pseudopalabras al leer y tuvo numerosas transformaciones en la escritura al dictado, pero respondió bien todas las preguntas fonológicas."
FULL[334] = "El estudiante 3 cambió solo dos pseudopalabras al leer, pero tuvo más imprecisiones en la identificación de sonidos, con respuestas como “Gro”, “Not” y “Met”, aunque segmentó bien “Tru-nal”. El estudiante 4 leyó toda la lista sin cambios y respondió con precisión las tareas fonológicas; sus dificultades aparecieron al escribir al dictado. El estudiante 5 cambió una sola pseudopalabra al leer (“Plerox” por “Pleros”), tuvo transformaciones al escribir y respondió bien casi todas las tareas fonológicas."
FULL[335] = "En los estudiantes 6 a 10 también se combinan dificultades y fortalezas. Los estudiantes 7 y 8 fueron los que más cambiaron al leer, con nueve pseudopalabras cada uno; el 9 y el 10 cambiaron cuatro, y el 6, dos. El estudiante 10, en cambio, respondió bien todas las preguntas fonológicas, mientras que los estudiantes 6, 7, 8 y 9 dieron respuestas imprecisas en varias de ellas. No hay, entonces, un único perfil de ejecución en el grupo."

# ---------------------------------------------------------------- Análisis de resultados (antes estaba bajo "Anexos")
ESTILO[667] = "Heading 1"
FULL[667] = "Análisis de resultados"
FULL[671] = "Al analizar la transcripción (Anexo A) encontramos que lo que hicieron los estudiantes coincide en buena medida con los referentes teóricos, sobre todo en conciencia fonológica, ruta fonológica y decodificación. Nueve de los diez cambiaron alguna pseudopalabra al leer, los diez lo hicieron al escribir y solo tres respondieron bien todas las preguntas de conciencia fonológica. Eso indica que varios estudiantes todavía no consolidan del todo los mecanismos para convertir grafemas en fonemas con precisión, tal como lo anticipaba el marco teórico, según el cual estas dificultades se manifiestan como sustituciones, omisiones y respuestas fonológicas imprecisas. Con todo, los resultados no alcanzan para diagnosticar dislexia ni para confirmar lateralidad cruzada, porque eso requiere evaluaciones específicas. La prueba funciona como un instrumento descriptivo que ayuda a detectar posibles dificultades."
FULL[681] = "Al leer, los estudiantes modificaron pseudopalabras en cantidades muy distintas. El estudiante 4 leyó toda la lista sin cambios y el 5 alteró una sola (“Plerox” por “Pleros”). Los estudiantes 1 y 2 cambiaron cinco cada uno; el 2, por ejemplo, dijo “Zepa” en lugar de “Zepan” y “Consiul” en lugar de “Zonil”. En el estudiante 3 aparecen solo “Zepa” y “Cuome”. Los casos más visibles son los de los estudiantes 7 y 8, con nueve cambios cada uno: “Zapa”, “Trampin”, “Digan”, “Blused”, “Zolil”, “Flanit”, “Florir”, “Tutim” y “Blumex”."
FULL[686] = "La segunda actividad, la escritura al dictado de pseudopalabras, permitió contrastar el desempeño lector con la producción escrita. Fue especialmente útil porque mostró qué pasaba cuando el estudiante no tenía un estímulo escrito delante y debía escuchar una pseudopalabra y convertirla en escritura."
FULL[687] = "Las modificaciones aparecen en todos los registros. El estudiante 4, por ejemplo, escribió “Troben Pamin Flusma”, “Luma Crepol Quatrio”, “Fradno Sutel Fialdo” y “Spanoi Nimpets Qurodq”, y el estudiante 3, “Troven Pamir Flusma”, “Luma Drepol Cuatrio”, “Fracno Sutel Fialdo” y “Espanoil Nitpel Luad”."
FULL[688] = "El estudiante 5 produjo, entre otras, “Troben Paminl Flusma”, “Luma Guepan Guatrio”, “Fragno Sutel Fialdo” y “Spanoil Mipet Quarado”. En los estudiantes 1 y 2, a quienes se les dictó la otra lista de la prueba, aparecen “Lisa” por “Plisa”, “Grac” por “Gruac” y “Tranklopiar” o “Tranflopiar” por “Tramkopliar”."
FULL[693] = "La tercera actividad, de conciencia y comprensión fonológica, fue especialmente útil para contrastar los resultados con el concepto de conciencia fonológica del marco teórico. Los estudiantes debían identificar sonidos iniciales, finales e internos, reconocer terminaciones semejantes y dividir una pseudopalabra en sílabas."
FULL[694] = "Los resultados cambiaron según la tarea. Para el sonido inicial de “gropel”, cinco estudiantes respondieron “G” (1, 2, 4, 5 y 10) y cinco “Gro” (3, 6, 7, 8 y 9). La diferencia importa, porque la pregunta pedía el sonido inicial y no el segmento completo. Una respuesta como “Gro” indica que el estudiante reconoce el comienzo de la palabra, pero no llega a aislar el fonema que se le pidió."
FULL[695] = "Con el sonido final de “finod” pasó algo parecido. Solo cuatro estudiantes respondieron “D” (2, 4, 5 y 10); los demás dieron “Nod” (7, 8 y 9), “Not” (3), “Od” (6) o “F” (1), que es el sonido inicial y no el final.\nCon el sonido interno de “lumep”, la respuesta esperada, “M”, también apareció en cuatro estudiantes (1, 2, 4 y 10); los otros produjeron formas más largas, como “Met”, “Ume”, “Um” o “Me”.\nDe aquí sale una distinción importante. Los estudiantes parecen más seguros cuando tienen que reconocer una estructura global o separar en sílabas, y vacilan más cuando deben aislar un sonido concreto dentro de una pseudopalabra."
FULL[696] = "Es uno de los hallazgos más interesantes de la prueba. La segmentación de “trunal” como “Tru-nal” aparece en los diez registros.\nNo sería correcto, entonces, hablar de una dificultad general en toda la conciencia fonológica. Los datos muestran un desempeño desigual según el nivel de procesamiento que pide cada tarea: los estudiantes resuelven bien algunas, en particular la segmentación silábica y el reconocimiento de terminaciones parecidas, y les cuesta más aislar sonidos específicos."
FULL[698] = "La actividad de escoger, entre “zopar”, “tamir” y “perar”, las pseudopalabras que terminaban igual también dio un resultado estable: los diez estudiantes eligieron “zopar” y “perar”.\nEsto importa porque muestra que su desempeño no es bajo en todas las dimensiones de la conciencia fonológica. Cuando la tarea es comparar, reconocen sin problema semejanzas sonoras globales."
FULL[701] = "La cuarta actividad, de fluidez lectora con pseudopalabras en frases, mostraba qué pasaba cuando esas palabras inventadas aparecían dentro de oraciones con sentido: “El drumo saltó sobre la mesa”, “María compró un teniq azul” y “El perro jugaba con el bapo en el parque”. Según la transcripción, todos pronunciaron bien las pseudopalabras. Los cambios fueron menores y casi siempre en el artículo: ocho estudiantes dijeron “un bapo” en lugar de “el bapo”, y el estudiante 1 omitió el artículo de la primera frase y dijo “Mi perro” en lugar de “El perro”."
FULL[702] = "Los estudiantes 9 y 10 leyeron las tres frases exactamente como estaban escritas. Los demás las leyeron con relativa continuidad, con los cambios menores ya descritos."
FULL[707] = "La quinta actividad, por último, analizaba de manera directa la relación entre los estímulos escritos y los sonidos iniciales, es decir, la correspondencia grafema-fonema, uno de los conceptos centrales del trabajo. La parte final, unir pseudopalabras con su sonido inicial, no quedó registrada en la transcripción, así que el análisis se limita a las dos primeras preguntas."
FULL[708] = "Cinco estudiantes identificaron bien “zunel” con /Z/; los otros cinco respondieron “Zu”, señal de que no aislaron el fonema sino la sílaba. Con “quarim” la variación fue mayor. El estudiante 9 contestó “ua” y el estudiante 7, “au”, ambos además con “/Zu/” para “zunel”. Los estudiantes 8 y 10, en cambio, identificaron la “Z” y la “q” con precisión."
FULL[712] = "El estudiante 1 cambia cinco pseudopalabras al leer (“Zeman”, “Consiul”, “Pelos”, “Fenir” y “Tumid”) y varias al escribir. En las tareas fonológicas acierta el sonido inicial, el interno, la terminación semejante y la segmentación, pero responde “F” como sonido final de “finod”. En las frases omite el artículo inicial de la primera y cambia “El perro” por “Mi perro”."
FULL[713] = "El estudiante 2 cambia cinco pseudopalabras al leer, entre ellas “Zepa” y “Consiul”, y tiene numerosas transformaciones en la escritura al dictado. Sin embargo, responde bien todas las tareas fonológicas: sonido inicial, final e interno, terminación semejante y segmentación silábica."
FULL[714] = "El estudiante 3 cambia solo dos pseudopalabras al leer (“Zepa” y “Cuome”), pero muestra más dificultad para identificar sonidos: responde “Gro” en lugar de aislar el sonido inicial, “Not” ante el sonido final de “finod” y “Met” ante el sonido interno de “lumep”, aunque segmenta bien “Tru-nal”."
FULL[715] = "El estudiante 4 lee toda la lista de pseudopalabras sin cambios y es preciso en las tareas fonológicas: responde “G”, “D” y “M”, reconoce “zopar–perar” y segmenta “trunal” como “Tru-nal”. Sus dificultades aparecen en la escritura al dictado y en la asociación grafema-fonema, donde responde “Zu” para “zunel”."
FULL[716] = "El estudiante 5 cambia una sola pseudopalabra al leer (“Plerox”), tiene transformaciones al escribir y una respuesta imprecisa en el sonido interno de “lumep”, donde responde “Ume”. Identifica bien el sonido inicial, el final, la semejanza final y la segmentación silábica."
FULL[717] = "El estudiante 6 cambia dos pseudopalabras al leer (“Used” y “Blumex”) y varias al escribir, y responde “Gro”, “Od” y “Um” en las tareas fonológicas; aun así, identifica la terminación semejante y segmenta bien. Este patrón vuelve a mostrar que la dificultad no se distribuye de igual manera en todas las habilidades."
FULL[718] = "El estudiante 7 es, con el 8, quien más modifica la lectura: cambia nueve pseudopalabras, como “Zapa”, “Trampin”, “Digan”, “Blused”, “Zolil”, “Flanit”, “Florir” y “Tutim”. En la comprensión fonológica responde “Gro”, “Nod” y “Me”, pero identifica la semejanza final y segmenta bien; en la asociación grafema-fonema responde “/Zu/” y “au”."
FULL[719] = "El estudiante 8 presenta en la lectura las mismas nueve modificaciones que el estudiante 7, además de cambios en la escritura y respuestas como “Gro”, “Nod” y “Me”. Conserva la respuesta correcta en semejanza final y segmentación silábica, e identifica con precisión /Z/ y “q” en la asociación grafema-fonema."
FULL[721] = "El estudiante 9 cambia cuatro pseudopalabras al leer (“Blused”, “Florir”, “Cuome” y “Brilus”), tiene transformaciones en la escritura y responde de forma imprecisa “Gro”, “Nod” y “Ume” en las tareas fonológicas, y “/Zu/” y “ua” en la asociación grafema-fonema. Las frases, en cambio, las lee tal como están escritas."
FULL[722] = "El estudiante 10 cambia cuatro pseudopalabras al leer (“Blused”, “Flanit”, “Florir” y “Coume”), pero responde bien todas las preguntas fonológicas (“Con la G”, “Con la D”, “Con la M”), reconoce la terminación semejante, segmenta “Tru-nal”, identifica /Z/ y “q” y lee las frases sin cambios."
FULL[723] = "Vistos los diez registros, no hay un único perfil de ejecución. Algunos estudiantes tienen más modificaciones en la lectura, otros en la escritura; unos responden de forma imprecisa en las tareas fonológicas y otros las resuelven bien. Esta diversidad importa, porque muestra que la dificultad lectora no es un fenómeno uniforme."
FULL[728] = "También hay correspondencia con lo que señalan Aguirre-Medrano y González-López (2021) sobre la relación entre lateralidad, confusiones entre grafemas y errores de lectura y escritura."
FULL[737] = "Al contrastar los resultados con los objetivos, la prueba dialoga sobre todo con el primer objetivo específico. Los fundamentos sobre conciencia fonológica, decodificación y rutas lectoras que identificamos en el marco teórico se pudieron observar en la práctica: en la lectura y la escritura de pseudopalabras, en la identificación de sonidos y en la correspondencia grafema-fonema. El tercer objetivo, diseñar una prueba diagnóstica centrada en la ruta fonológica, también se cumplió: la diseñamos con pseudopalabras precisamente para que los estudiantes no se apoyaran en palabras conocidas, y la aplicamos a los diez participantes."
FULL[738] = "El segundo objetivo, describir los factores socioculturales, económicos y pedagógicos, se cumplió desde la revisión documental, en los antecedentes y el marco teórico, pero no desde la prueba, que no los midió. Por eso el trabajo tiene más evidencia empírica sobre los componentes fonológicos y de decodificación que sobre los socioculturales, pedagógicos, de lateralidad o de pandemia."
FULL.pop(739, None)
BORRAR += list(range(739, 747))  # lista de hallazgos repetida (ya está en Resultados y en Conclusiones)

# ---------------------------------------------------------------- Discusión y conclusiones con datos concretos
FULL[750] = "Según el marco teórico, las pseudopalabras sirven para evaluar la ruta fonológica porque obligan a usar las reglas que conectan grafemas y fonemas. Y fueron justamente las pseudopalabras las que dejaron ver la mayor cantidad de cambios: nueve de los diez estudiantes alteraron alguna al leer y todos al escribir al dictado."
FULL[752] = "La comparación también mostró fortalezas. Los diez estudiantes separaron bien “trunal” en sílabas y reconocieron las pseudopalabras con terminaciones parecidas, lo que indica que sí tienen habilidades fonológicas. Lo adecuado, por tanto, no es decir que “no tienen conciencia fonológica”, sino que tienen dificultades puntuales en ciertas operaciones, sobre todo en aislar un fonema."
ESTILO[765] = "Heading 1"
FULL[765] = "Conclusiones"
FULL[766] = "El análisis de la prueba de pseudopalabras nos lleva a concluir que los estudiantes evaluados tienen distintos niveles de consolidación en la decodificación, la conciencia fonológica y la correspondencia grafema-fonema. Nueve de los diez cambiaron alguna pseudopalabra al leer, y esos cambios muestran que, frente a estímulos desconocidos, a varios les cuesta mantener una correspondencia precisa entre letras y sonidos."
FULL[767] = "Leer pseudopalabras permitió observar de cerca los procesos ligados a la ruta fonológica, porque con estos estímulos casi no hay manera de apoyarse en palabras conocidas. Las sustituciones, omisiones y transformaciones que registramos son pistas útiles para detectar posibles dificultades de decodificación. Como fueron muy desiguales, desde ningún cambio en el estudiante 4 hasta nueve en los estudiantes 7 y 8, no se puede hablar de un único perfil lector en el grupo."
FULL[768] = "En la escritura al dictado, todos los estudiantes transformaron algunas pseudopalabras, lo que muestra que las dificultades también aparecen cuando hay que convertir lo que se oye en algo escrito. Sobre la cuarta pregunta de investigación, las omisiones y sustituciones aparecen tanto al leer como al escribir, lo que sugiere un origen común en la conversión grafema-fonema. Como las listas de lectura y de dictado no eran las mismas, la prueba no permite comparar palabra por palabra, y esa comparación queda como tarea para una próxima aplicación."
FULL[769] = "La conciencia fonológica mostró un desempeño desigual. Solo tres estudiantes respondieron bien todas las preguntas; los demás confundieron el sonido con la sílaba (“Gro” por “G”) o dieron segmentos más largos. En cambio, los diez segmentaron bien en sílabas y reconocieron las terminaciones parecidas. No hay, entonces, una ausencia general de conciencia fonológica, sino dificultades en ciertas operaciones de análisis sonoro, sobre todo en aislar un fonema."
NUEVOS_TRAS[775] = [("Sobre las preguntas que miraban a la enseñanza, la revisión documental deja algunas respuestas. Los estudios revisados coinciden en que dan mejores resultados las intervenciones que trabajan de forma explícita la conciencia fonológica y la discriminación auditiva (Aguirre-Medrano & González-López, 2021), las que combinan actividades multisensoriales y lúdicas con estudiantes en riesgo de exclusión social (Domínguez Vázquez, 2023) y las que usan el juego para motivar a los estudiantes (Rivera Cintrón & Batiz Cartagena, 2024). Ninguna de ellas funciona sin docentes bien formados (Giraldo Gaviria & Caro Lopera, 2022; Morales Londoño, 2018), y la pandemia mostró que el acompañamiento de las familias y el acceso a recursos pesan tanto como el método (López Rivas, 2024).", None)]

# ---------------------------------------------------------------- Párrafos largos divididos (ritmo más variado)
FULL[248] = "Entendemos la lectoescritura como un proceso cognitivo y pedagógico complejo. Leer no es solo reconocer letras: el lector decodifica, comprende, sostiene en la memoria lo que va leyendo y construye un significado. En ese proceso la conciencia fonológica tiene un papel central, porque es lo que le permite al niño separar los sonidos del habla y unirlos con las letras que los representan. Cuando esa habilidad falla, como ocurre en la dislexia o en los problemas de discriminación auditiva, aparecen errores típicos de la ruta fonológica; Aguirre-Medrano y González-López (2021) describen, por ejemplo, la confusión entre la b y la d."
NUEVOS_TRAS[248] = [("Para observar esa ruta por separado sirven las pseudopalabras. Una palabra inventada como “trunal” no significa nada. El lector no la puede reconocer de memoria y tiene que leerla letra por letra, aplicando las reglas de conversión grafema-fonema.", None)]
FULL[249] = "A esa mirada cognitiva le sumamos otra, pedagógica y de contexto. Domínguez Vázquez (2023) y López Rivas (2024) muestran que la alfabetización depende mucho del entorno: de si el niño vive en condiciones vulnerables, de si tuvo conexión durante la pandemia y de cuántos libros hay a su alcance. Por eso hacen falta métodos flexibles, ajustados a cada lugar. Los docentes pesan igual. Morales Londoño (2018) y Giraldo Gaviria y Caro Lopera (2022) critican las prácticas tradicionales y piden que los maestros reciban una mejor formación en alfabetización académica, para que puedan diseñar intervenciones a la medida."
NUEVOS_TRAS[249] = [("Con estas dos miradas armamos el marco del trabajo. La idea que lo atraviesa es sencilla: primero hay que saber con precisión qué le cuesta al estudiante, y una prueba de pseudopalabras ayuda a eso. Después vienen las estrategias, que deben responder tanto a sus necesidades como a su contexto.", None)]
FULL[671] = "Al analizar la transcripción (Anexo A) encontramos que lo que hicieron los estudiantes coincide en buena medida con los referentes teóricos, sobre todo en conciencia fonológica, ruta fonológica y decodificación. Nueve de los diez cambiaron alguna pseudopalabra al leer, los diez lo hicieron al escribir y solo tres respondieron bien todas las preguntas de conciencia fonológica. Eso indica que varios estudiantes todavía no consolidan del todo los mecanismos para convertir grafemas en fonemas con precisión, tal como lo anticipaba el marco teórico."
NUEVOS_TRAS[671] = [("Con todo, los resultados no alcanzan para diagnosticar dislexia ni para confirmar lateralidad cruzada, porque eso requiere evaluaciones específicas. La prueba funciona como un instrumento descriptivo que ayuda a detectar posibles dificultades.", None)]
NUEVOS_TRAS[775] = [
    ("Sobre las preguntas que miraban a la enseñanza, la revisión documental deja algunas respuestas. Dan mejores resultados las intervenciones que trabajan de forma explícita la conciencia fonológica y la discriminación auditiva (Aguirre-Medrano & González-López, 2021). También funcionan las que combinan actividades multisensoriales y lúdicas, como la que Domínguez Vázquez (2023) aplicó con estudiantes en riesgo de exclusión social, y las que usan el juego para motivar a los niños (Rivera Cintrón & Batiz Cartagena, 2024).", None),
    ("Ninguna de estas estrategias funciona sin docentes bien formados (Giraldo Gaviria & Caro Lopera, 2022; Morales Londoño, 2018). Y la pandemia dejó claro que el acompañamiento de las familias y el acceso a recursos pesan tanto como el método (López Rivas, 2024).", None),
]

# ---------------------------------------------------------------- Coherencia entre lectura aislada y en frases
# Las listas de lectura y dictado eran distintas, y las pseudopalabras de las frases (drumo, teniq, bapo)
# no son las mismas que los estudiantes alteraron al leer: se ajustan las afirmaciones que suponían lo contrario.
FULL[677] = "También vimos que el desempeño de un mismo estudiante podía cambiar de una actividad a otra. En la lista aislada aparecieron más modificaciones, mientras que las frases con pseudopalabras se leyeron de forma bastante estable. Esto hace pensar que el contexto de la tarea influye en la ejecución, aunque la transcripción no registra tiempos, velocidad ni número exacto de errores que permitan compararlo estadísticamente."
FULL[690] = "Al comparar las dos tareas surge una observación. Las mismas clases de error, sobre todo omisiones y sustituciones de letras, aparecen al leer y al escribir. Eso hace pensar que en algunos estudiantes la representación fonológica todavía no es estable, o que les cuesta pasar del código escrito al sonoro y viceversa. Como las listas de lectura y de dictado eran distintas, no podemos comparar palabra por palabra; la hipótesis, de todos modos, es coherente con nuestro marco teórico, que pone la conciencia fonológica y la conversión grafema-fonema en el centro de la alfabetización."
FULL[703] = "Aquí aparece un contraste con la lectura aislada, aunque hay que matizarlo. Las pseudopalabras de las frases (“drumo”, “teniq” y “bapo”) eran cortas y sencillas, y “teniq” y “bapo” también se leyeron bien en la lista. Lo que sí se ve es que, dentro de las oraciones, ningún estudiante alteró una pseudopalabra."
FULL[706] = "El resultado sugiere que la lectura depende de lo que cada actividad exige. Leer una pseudopalabra sola obliga a decodificar de forma directa, mientras que dentro de una frase el lector cuenta con pistas del contexto."
FULL[711] = "A partir de estos resultados, el análisis conjunto de los diez registros deja ver un patrón general y también diferencias individuales."
FULL[753] = "Este matiz hace la discusión más precisa y menos determinista. A eso se suma que el desempeño cambia según la tarea: hay estudiantes que alteran varias pseudopalabras en la lista y, sin embargo, leen sin tropiezos las frases con pseudopalabras. El contexto de la tarea parece influir en la ejecución, aunque para comprobarlo habría que usar las mismas pseudopalabras en las dos tareas."
FULL[769] = "La conciencia fonológica mostró un desempeño desigual. Solo tres estudiantes respondieron bien todas las preguntas; los demás confundieron el sonido con la sílaba (“Gro” por “G”), dieron segmentos más largos o, como el estudiante 1, señalaron otro sonido. En cambio, los diez segmentaron bien en sílabas y reconocieron las terminaciones parecidas. No hay, entonces, una ausencia general de conciencia fonológica, sino dificultades en ciertas operaciones de análisis sonoro, sobre todo en aislar un fonema."
FULL[771] = "El desempeño, además, cambia según la tarea. Las frases con pseudopalabras se leyeron con relativa estabilidad, mientras que la lista aislada produjo más cambios, lo que sugiere que el contexto lingüístico ayuda a leer. Para comprobarlo habría que usar las mismas pseudopalabras en ambas tareas. Evaluar la lectoescritura exige, por eso, mirar distintas situaciones y niveles de procesamiento."
FULL[773] = "Los resultados no permiten confirmar, por sí solos, la presencia de dislexia. Y aunque la lateralidad y los factores perceptivos forman parte del marco teórico, la prueba no incluyó una evaluación de lateralidad, así que no sería correcto concluir que los participantes tienen lateralidad cruzada. Para eso se necesitan procedimientos e instrumentos específicos."
FULL[774] = "Tampoco pueden darse por demostrados como causas los factores socioculturales, económicos, pedagógicos o los efectos de la pandemia, porque la prueba no los evaluó. Sirven para situar el problema desde la revisión documental; relacionarlos con el desempeño de cada estudiante exigiría otros instrumentos."
FULL[776] = "Por último, la evaluación de la lectoescritura debería hacerse temprano, de forma sistemática y mirando varias habilidades a la vez. Detectar dificultades concretas ayuda a orientar el acompañamiento pedagógico, siempre que los resultados se lean dentro del alcance real de los instrumentos y se complementen, cuando haga falta, con otras formas de evaluación."

# ---------------------------------------------------------------- Siglas
SIGLAS = [
    "APA\t\t\tAmerican Psychological Association",
    "ENS\t\t\tEscuelas Normales Superiores",
    "MEN\t\t\tMinisterio de Educación Nacional",
    "ONU\t\t\tOrganización de las Naciones Unidas",
    "PISA\t\t\tPrograma para la Evaluación Internacional de Estudiantes",
    "PROLEC-R\t\tBatería de Evaluación de los Procesos Lectores, Revisada",
    "PSC\t\t\tProfesores de servicios a la comunidad",
    "UPB\t\t\tUniversidad Pontificia Bolivariana",
    "WISC IV\t\tEscala de Inteligencia de Wechsler para Niños, IV",
]

# ---------------------------------------------------------------- Referencias (APA 7, verificadas)
REFS = [
    [("Aguirre-Medrano, A., & González-López, M. (2021). Problemas de dislexia en educación básica: Una problemática para la lectoescritura. ", False), ("Santiago", True), (", (156), 135–148. https://santiago.uo.edu.cu/index.php/stgo/article/view/5436", False)],
    [("Caamaño Tomás, A. (2021). Consideraciones acerca de la educación, lectoescritura y aprendizaje en los tiempos de pandemia. ", False), ("Fuentes Humanísticas, 33", True), ("(63), 49–57.", False)],
    [("Carlino, P. (2005). ", False), ("Escribir, leer y aprender en la universidad: Una introducción a la alfabetización académica", True), (". Fondo de Cultura Económica.", False)],
    [("Domínguez Vázquez, B. C. (2023). Una propuesta de intervención para el desarrollo de las habilidades de lectoescritura en alumnado de primer ciclo de primaria con riesgo de exclusión social. ", False), ("REIDOCREA, 12", True), ("(38), 507–518.", False)],
    [("Duarte-Hernández, F. J., & Pérez-Mendoza, N. B. (2020). Identificar la lateralidad en niños de 2 a 5 años del Instituto de Recreación y Deportes de Tunja (IRDET) aplicando el test de Harris. ", False), ("Revista Digital: Actividad Física y Deporte, 6", True), ("(2), 118–144.", False)],
    [("Giraldo Gaviria, D. M., & Caro Lopera, M. Á. (2022). Alfabetización académica: Una alternativa para repensar la formación inicial docente en las escuelas normales superiores de Colombia. ", False), ("Zona Próxima", True), (", (37), 53–79. https://doi.org/10.14482/zp.37.378.129", False)],
    [("Gutiérrez Fresneda, R., & Díez Mediavilla, A. (2018). Conciencia fonológica y desarrollo evolutivo de la escritura en las primeras edades. ", False), ("Educación XX1, 21", True), ("(1), 395–415.", False)],
    [("López Rivas, O. H. (2024). Enseñanza de la lectoescritura en la escuela primaria en tiempos de pandemia. Estudios de casos. ", False), ("Revista Historia de la Educación Latinoamericana, 26", True), ("(44), 181–203. https://doi.org/10.19053/uptc.01227238.18254", False)],
    [("Morales Londoño, F. de J. (2018). La calidad de los procesos de enseñanza de la lectoescritura en Colombia: Entre concepciones y prácticas. ", False), ("Horizontes. Revista de Investigación en Ciencias de la Educación, 2", True), ("(8), 258–272. https://doi.org/10.33996/revistahorizontes.v2i8.61", False)],
    [("Naciones Unidas. (2015). ", False), ("Transformar nuestro mundo: La Agenda 2030 para el Desarrollo Sostenible", True), (" (A/RES/70/1). https://undocs.org/es/A/RES/70/1", False)],
    [("Núñez Peña, M. (2011). ", False), ("Diseños de investigación en Psicología", True), (". Universitat de Barcelona. https://hdl.handle.net/2445/20322", False)],
    [("Pisco-Román, J., & Bailón-Panta, A. (2023). La lectoescritura como elemento fundamental en el proceso de enseñanza aprendizaje de los estudiantes de Básica Media. ", False), ("593 Digital Publisher CEIT, 8", True), ("(1-1), 328–347. https://doi.org/10.33386/593dp.2023.1-1.1658", False)],
    [("Rivera Cintrón, A. E., & Batiz Cartagena, M. (2024). Neuroaprendizaje basado en la lectoescritura y estrategias lúdicas en el currículo Artes del Lenguaje. ", False), ("HETS Online Journal, 15", True), ("(1).", False)],
    [("Snowling, M. J., Hulme, C., & Nation, K. (2020). Defining and understanding dyslexia: Past, present and future. ", False), ("Oxford Review of Education, 46", True), ("(4), 501–513. https://doi.org/10.1080/03054985.2020.1765756", False)],
    [("Tinta Aruquipa, M. R. (2020). Proceso de enseñanza aprendizaje de la escritura a partir de la lectura de la realidad. ", False), ("Horizontes. Revista de Investigación en Ciencias de la Educación, 4", True), ("(16), 553–568. https://doi.org/10.33996/revistahorizontes.v4i16.137", False)],
]


def poner_runs(p, segmentos):
    """Deja el párrafo con un run por segmento (texto, cursiva), copiando el formato del primer run."""
    base = p.runs[0]._r
    rpr = deepcopy(base.rPr) if base.rPr is not None else None
    for hijo in list(p._p):  # runs, hipervínculos y marcadores; se conserva el formato de párrafo
        if hijo.tag != qn("w:pPr"):
            p._p.remove(hijo)
    for texto, cursiva in segmentos:
        run = p.add_run(texto)
        if rpr is not None:
            run._r.insert(0, deepcopy(rpr))
            R.limpiar_formato(run)
        run.italic = True if cursiva else None


def insertar_despues(p, texto, estilo=None, molde=None):
    nuevo = deepcopy((molde or p)._p)
    p._p.addnext(nuevo)
    q = Paragraph(nuevo, p._parent)
    R.reemplazar(q, texto)
    if estilo:
        q.style = estilo
    return q


def main(src, dst):
    d = docx.Document(src)
    P = d.paragraphs

    for i, t in FULL.items():
        R.reemplazar(P[i], t)
    for i, t in LEAD.items():
        R.reemplazar_tras_negrita(P[i], t)
    for i, cambios in ERRATAS.items():
        for viejo, nuevo in cambios:
            R.errata(P[i], viejo, nuevo)
    for i, estilo in ESTILO.items():
        if P[i].text != P[i].text.strip():
            R.reemplazar(P[i], P[i].text.strip())
        P[i].style = estilo

    # párrafos nuevos (participantes y respuesta a las preguntas sobre enseñanza)
    for i, nuevos in NUEVOS_TRAS.items():
        ancla = P[i]
        for texto, estilo in nuevos:
            molde = P[289] if estilo else P[i]
            ancla = insertar_despues(ancla, texto, estilo, molde)

    # siglas
    P[76].text = SIGLAS[0]
    P[77].text = SIGLAS[-1]
    ancla = P[76]
    for s in SIGLAS[1:-1]:
        ancla = insertar_despues(ancla, s, molde=P[76])

    # referencias: título APA, entradas completas y en orden alfabético
    P[803].text = "Referencias"
    refs = [P[i] for i in range(804, 815)]
    ancla = refs[-1]
    for k, segs in enumerate(REFS):
        if k < len(refs):
            p = refs[k]
        else:
            nuevo = deepcopy(refs[0]._p)
            ancla._p.addnext(nuevo)
            p = Paragraph(nuevo, refs[0]._parent)
        poner_runs(p, segs)
        ancla = p
    ultima_ref = ancla

    # la transcripción pasa a Anexo A, después de las referencias
    P[339].text = "Anexo A. Transcripción de la prueba"
    P[339].style = "Heading 1"
    P[339].paragraph_format.page_break_before = True
    bloque, el = [], P[339]._p
    while el is not None and el is not P[667]._p:
        bloque.append(el)
        el = el.getnext()
    destino = ultima_ref._p
    for el in bloque:
        destino.addnext(el)
        destino = el

    for i in sorted(set(BORRAR), reverse=True):
        P[i]._p.getparent().remove(P[i]._p)

    # títulos vacíos (espaciadores) pasan a Normal para que no salgan en la tabla de contenido
    for p in d.paragraphs:
        if p.style.name.startswith("Heading") and not p.text.strip():
            p.style = "Normal"

    # que Word actualice la tabla de contenido al abrir
    ajustes = d.settings.element
    if ajustes.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        ajustes.append(uf)
    d.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
