# Versión 5: amplía la v4 con contenido nuevo que sale de los datos de la transcripción y de
# fuentes verificadas: contexto PISA, modelo de doble ruta, sílabas trabadas, errores ortográficos
# frente a fonológicos, descripción del instrumento, procedimiento y criterios de análisis,
# tipología de errores (Tabla 2), discusión con los antecedentes y una propuesta pedagógica.
# Uso: python version5.py "Trabajo lectoescritura v4.docx" "Trabajo lectoescritura v5.docx"
import sys
from copy import deepcopy

import docx
from docx.text.paragraph import Paragraph

import reescribir as R
import version3 as V3
import version4 as V4

RAE = "Real Academia Española & Asociación de Academias de la Lengua Española, 2010"

# (comienzo del párrafo ancla, [(texto, estilo)])  estilo None = párrafo normal como el ancla
INSERTAR = [
    ("En muchos contextos educativos", [
        ("Colombia es un buen ejemplo. En la prueba PISA de 2022, los estudiantes colombianos de 15 años obtuvieron 409 puntos en lectura, frente a un promedio de 476 en los países de la OCDE, el mismo resultado que el país había alcanzado en 2018 (OECD, 2023). PISA evalúa la comprensión de textos en adolescentes y no la decodificación inicial, pero una brecha tan estable indica que las dificultades no se corrigen solas con el paso de los años escolares y que vale la pena revisar cómo se consolidan las bases de la lectura en los primeros grados.", None),
    ]),
    ("La lectoescritura reúne habilidades cognitivas", [
        ("Las dos rutas de la lectura", "Heading 2"),
        ("La psicología cognitiva explica la lectura en voz alta mediante dos vías. Por la ruta léxica, el lector reconoce la palabra completa porque ya la tiene almacenada y accede de inmediato a su pronunciación y a su significado; por la ruta fonológica, convierte cada grafema en su fonema y ensambla la pronunciación paso a paso. El modelo de doble ruta en cascada formalizó esta distinción y mostró que en el lector competente ambas vías trabajan en paralelo, aunque cada una se vuelve indispensable para un tipo de estímulo distinto, las palabras irregulares o muy frecuentes en el caso de la ruta léxica y las palabras desconocidas o inventadas en el de la fonológica (Coltheart et al., 2001).", None),
        ("En español, cuya ortografía es bastante transparente, la ruta fonológica pesa mucho durante el aprendizaje inicial, porque casi cualquier palabra puede leerse bien aplicando las reglas de conversión. El PROLEC-R aprovecha esa característica para evaluar las dos rutas por separado, con la lectura de palabras para la léxica y la de pseudopalabras para la fonológica (Cuetos et al., 2007). Un niño que lee sin tropiezos las palabras de su entorno y se equivoca con pseudopalabras de estructura parecida tiene, desde esta perspectiva, una ruta fonológica que todavía no se ha consolidado, aunque su lectura cotidiana parezca fluida.", None),
    ]),
    ("La conciencia fonológica es la capacidad", [
        ("La conciencia fonológica abarca varios niveles que se adquieren en un orden bastante estable. Primero aparece la capacidad de manipular sílabas, después la de reconocer unidades intrasilábicas, como la rima que permite advertir que dos palabras terminan igual, y por último la conciencia fonémica, que exige aislar y manipular cada sonido y que se desarrolla sobre todo cuando el niño aprende a leer y escribir (Jiménez González & Ortiz González, 1998). Esa gradación explica que un estudiante pueda dividir una palabra en sílabas sin dificultad y, al mismo tiempo, no logre decir cuál es su primer sonido.", None),
        ("No todas las sílabas presentan la misma dificultad. Las que comienzan con un grupo consonántico, como bla, fre o pli, exigen distinguir dos consonantes seguidas antes de la vocal, y los niños tardan más en representarlas con precisión. Jiménez González y Jiménez Rodríguez (1999) observaron en estudiantes de primero a tercero de primaria que estas sílabas producen errores tanto cuando los niños escriben como cuando deben juzgar si una letra forma parte de una palabra, y que esos errores dependen del nivel de conciencia fonémica, con una relación que se debilita hacia tercer grado a medida que esta habilidad madura.", None),
    ]),
    ("Los errores que aparecen en estas tareas", [
        (f"Al analizar la escritura hay que separar, además, los errores fonológicos de los ortográficos. En el español de América, la s, la z y la c ante e o i representan el mismo sonido, y la b y la v corresponden a un único fonema ({RAE}). Un estudiante colombiano que escribe “Sutel” al oír “Zutel” produce una forma que suena exactamente igual al estímulo; lo que desconoce es una convención que el dictado de una palabra inventada no le permite deducir, y contar esa grafía como error fonológico inflaría el diagnóstico.", None),
    ]),
    ("En la primera actividad, el estudiante lee", [
        ("Procedimiento", "Heading 2"),
        ("La prueba se aplicó a cada estudiante por separado. Quien la aplicó leyó en voz alta cada consigna, anotó lo que el estudiante decía en las actividades orales y transcribió lo que escribía en el dictado, de modo que el Anexo A conserva las respuestas en el mismo orden de la prueba. No se registraron tiempos de lectura ni número de pausas.", None),
        ("Criterios de análisis", "Heading 2"),
        ("Se consideró error toda diferencia entre la respuesta y el estímulo, con dos salvedades. En la escritura no se contaron como errores fonológicos las grafías que representan el mismo sonido en el español de Colombia, como s por z, b por v o cu por qu. En la lectura no se tuvo en cuenta “Zolid”, que los diez estudiantes leyeron de la misma manera y que por eso se tomó como la forma presentada. Los errores de lectura se agruparon después según afectaran grupos consonánticos, la consonante final u otras consonantes, la lectura de qu o el orden de las letras, para identificar patrones comunes al grupo.", None),
    ]),
    ("Lectura de pseudopalabras.", [
        ("Agrupados por tipo, los 41 cambios registrados en la lectura muestran un patrón claro (Tabla 2). Casi la mitad afectó grupos consonánticos con l o r, ya fuera porque el estudiante los redujo (“Fenir” por “Flenir”), cambió una líquida por otra (“Flanit” por “Franit”, “Blumex” por “Brumex”) o creó uno donde no lo había (“Blused” por “Lused”, “Florir” por “Folir”). Les siguen las sustituciones y adiciones de otras consonantes y la omisión o el cambio de la consonante final, mientras que la lectura de “qu” como si fuera “cu” y las inversiones fueron menos frecuentes. Ningún estudiante confundió b con d ni p con q.", None),
        ("__TABLA2__", None),
    ]),
    ("Escritura al dictado.", [
        ("Una parte de esas diferencias no altera el sonido. Cinco de los ocho estudiantes que recibieron la segunda lista escribieron “Sutel” por “Zutel”, tres “Bentil” por “Ventil” y dos “Cuamir” por “Quamir”, formas que en el español de Colombia se pronuncian igual que el estímulo. Entre los cambios que sí modifican el sonido destacan otra vez las líquidas y las consonantes finales, con “Pamil” por “Pamir” en cuatro estudiantes y “Ventir” por “Ventil”, “Tralu” por “Tralum” y “Neclon” por “Neclom” en tres cada uno. En la primera lista aparecen además reducciones de grupos consonánticos, como “Lisa” por “Plisa” y “Bieco” por “Bieclo”, junto con transformaciones como “Grac” por “Gruac”, “Caudre” por “Cuadre” o “Tranflopiar” por “Tramkopliar”.", None),
    ]),
    ("Los resultados coinciden en buena medida", [
        ("El tipo de error dice más que su cantidad. Que casi la mitad de los cambios se concentre en grupos consonánticos con l o r coincide con lo que Jiménez González y Jiménez Rodríguez (1999) describen para la escritura en los primeros grados, aunque aquí el fenómeno aparece en la lectura y en las dos direcciones. Diez de esos cambios ocurrieron en las trece pseudopalabras que tenían grupo consonántico, y los otros ocho en palabras que no lo tenían, donde algunos estudiantes lo crearon, como si la secuencia de consonante, líquida y vocal funcionara como un molde que se aplica de más. El dato tiene valor práctico, porque señala un contenido preciso, el de las sílabas trabadas, sobre el cual concentrar el trabajo de conciencia fonémica.", None),
    ]),
    ("Los resultados son coherentes con lo que la literatura", [
        ("Los datos permiten también leer de otra manera el hallazgo de Pisco-Román y Bailón-Panta (2023), para quienes la comprensión lectora es la dificultad principal de los estudiantes de básica media. La prueba de este trabajo no evaluó la comprensión, pero muestra que en un nivel previo, el de la conversión grafema-fonema, persisten imprecisiones. Es razonable pensar que un lector que todavía debe esforzarse para descifrar ciertas sílabas dispone de menos atención para entender lo que lee, y que parte de los problemas de comprensión que reportan esos estudios se origine en una decodificación que no llegó a automatizarse. Si así fuera, atenderlos solo con estrategias de comprensión dejaría intacta una de sus causas.", None),
        ("Las intervenciones revisadas encajan con este diagnóstico en distinta medida. El programa de Domínguez Vázquez (2023) incluye reconocimiento de fonemas y actividades multisensoriales, es decir, trabaja la operación en la que más estudiantes fallaron, mientras que la propuesta de Rivera Cintrón y Batiz Cartagena (2024) apuesta por el juego y la motivación, que mejoran la disposición hacia la lectura pero no aseguran por sí solos el análisis fonémico que estos datos reclaman. Un trabajo explícito sobre los fonemas presentado en formato lúdico reúne lo mejor de ambas orientaciones y es la base de la propuesta que sigue.", None),
    ]),
    ("La dificultad más clara está en aislar fonemas", [
        ("El análisis por tipo de error precisa ese diagnóstico. Casi la mitad de los cambios de lectura se produjo en grupos consonánticos con l o r y ningún estudiante confundió b con d ni p con q, de modo que las dificultades del grupo parecen más fonológicas que visoespaciales y tienen en las sílabas trabadas un foco concreto de intervención. En la escritura, una parte de las diferencias resultó ser ortográfica y no fonológica, lo que obliga a separar ambos tipos de error antes de valorar el desempeño de un estudiante.", None),
    ]),
]

# Propuesta pedagógica: se inserta antes del título "Conclusiones"
PROPUESTA = [
    ("Propuesta pedagógica", "Heading 1"),
    ("La secuencia que se propone para el grupo evaluado parte de lo que los estudiantes ya dominan, la sílaba, para llegar a lo que todavía les cuesta, el fonema aislado y las sílabas trabadas, y separa desde el comienzo los errores de sonido de los de ortografía, porque requieren tratamientos distintos. La Tabla 1 permite asignar a cada estudiante el componente en el que más necesita apoyo.", None),
    ("Del análisis silábico al fonema", "Heading 2"),
    ("Como los diez estudiantes segmentan bien en sílabas, el trabajo puede empezar allí y avanzar hacia unidades menores. Una actividad sencilla consiste en decir una palabra conocida, dividirla en sílabas con palmadas y luego representar cada sonido de la primera sílaba con una ficha, de manera que “sol” se convierte en tres fichas y “gro” también; después el estudiante quita, cambia o añade una ficha y pronuncia la palabra que resulta. Ese paso de la sílaba al fonema es el que no lograron los estudiantes 3, 6, 7, 8 y 9 cuando respondieron “Gro” o “Nod”, y la combinación de objetos, movimientos y sonidos retoma el carácter multisensorial que dio resultado en el programa de Domínguez Vázquez (2023).", None),
    ("Las palabras reales deben preceder a las inventadas. El análisis se practica primero con los nombres de los propios estudiantes y con palabras de su entorno, que les resultan familiares y les permiten concentrarse en los sonidos, y solo cuando la operación está clara se pasa a las pseudopalabras, que obligan a aplicarla sin el apoyo del significado y muestran si el estudiante aisló de verdad el sonido o lo recordó. Esta progresión acompaña el orden en que se desarrollan la conciencia fonológica y la escritura (Gutiérrez Fresneda & Díez Mediavilla, 2018).", None),
    ("Sílabas trabadas", "Heading 2"),
    ("Para los grupos consonánticos, que concentraron casi la mitad de los errores de lectura, sirven los pares de palabras que solo se diferencian por la líquida, como “pato” y “plato”, “fan” y “flan”, “cara” y “clara” o “boca” y “broca”. El estudiante escucha y lee ambas, señala cuál tiene un sonido de más y lo marca con una ficha de otro color, de modo que la segunda consonante, la que tiende a omitir o a cambiar, se vuelve visible. Quienes crearon grupos donde no existían, los estudiantes 7, 8, 9 y 10, necesitan también el ejercicio contrario, leer pseudopalabras como “lused” o “folir” y decidir si llevan o no una consonante antes de la l. Como Jiménez González y Jiménez Rodríguez (1999) vinculan estos errores con el nivel de conciencia fonémica, el trabajo debe ser explícito y no limitarse a la lectura repetida de las mismas sílabas.", None),
    ("Escritura y ortografía", "Heading 2"),
    (f"La escritura requiere dos tratamientos diferentes. Los cambios que alteran el sonido, como “Pamil” por “Pamir” o “Tralu” por “Tralum”, pueden trabajarse con dictados breves en los que otra persona lee en voz alta lo que el estudiante escribió, para que él mismo escuche la diferencia con la palabra dictada. Las grafías que no cambian el sonido, como s y z o b y v, pertenecen a la ortografía ({RAE}) y se enseñan con palabras reales, familias de palabras y reglas, ya que en una palabra inventada no hay manera de deducir cuál corresponde.", None),
    ("Todas estas actividades admiten formatos de juego, como cartas, loterías de sonidos o retos por equipos, en la línea de Rivera Cintrón y Batiz Cartagena (2024), siempre que el juego acompañe el análisis del fonema y no lo sustituya. Para valorar su efecto, la prueba puede aplicarse de nuevo después de un periodo de trabajo, esta vez con las mismas pseudopalabras en la lectura y en el dictado y con registro del tiempo, lo que permitiría comparar los dos momentos y las dos tareas en condiciones equivalentes.", None),
]

# Ajustes de frases existentes (texto exacto -> nuevo)
AJUSTES = [
    ("La técnica fue la evaluación individual. Cada estudiante resolvió las cinco actividades de la prueba, cuyas respuestas se registraron por escrito (Anexo A) y se compararon después, una por una, con los estímulos originales:", "La técnica fue la evaluación individual con la prueba diagnóstica de pseudopalabras, cuyas respuestas se registraron por escrito (Anexo A) y se compararon después, una por una, con los estímulos originales."),
    ("__LISTA__", 'En la primera actividad, el estudiante lee en voz alta treinta pseudopalabras bisílabas, trece de ellas con grupo consonántico, como “bapo”, “trapin” o “blisur”, después de oír la consigna de que no tienen significado pero deben leerse como palabras normales. En la segunda escribe las pseudopalabras que se le dictan, con la indicación de escribirlas tal como las escucha; se usaron dos listas, una para los estudiantes 1 y 2 y otra para los demás. La tercera reúne cinco preguntas de conciencia fonológica sobre “gropel”, “finod”, “lumep”, “zopar”, “tamir”, “perar” y “trunal”, que piden el sonido inicial, el final y el interno, la pareja que termina igual y la división en sílabas. La cuarta consiste en leer tres oraciones con una pseudopalabra cada una, y la quinta pide el sonido inicial de “zunel”, el grafema con que comienza “quarim” y la unión de tres pseudopalabras con su sonido inicial.'),
    ("donde cinco estudiantes respondieron “Zu” y no /Z/ para “zunel”.", "donde cinco estudiantes respondieron “Zu” y no /Z/ para “zunel”. El patrón coincide con la secuencia de niveles descrita en el marco teórico, ya que los diez dominan el nivel silábico y el de la rima y las dificultades aparecen en el fonémico, el último en consolidarse (Jiménez González & Ortiz González, 1998)."),
    ("a partir de la tarea de pseudopalabras del PROLEC-R, y en analizar", "a partir de la tarea de pseudopalabras del PROLEC-R (Cuetos et al., 2007), y en analizar"),
    ("Escritura al dictado. Los diez estudiantes escribieron alguna pseudopalabra de forma distinta al estímulo, con producciones como “Lisa” por “Plisa”, “Bieco” por “Bieclo”, “Grac” por “Gruac”, “Tranflopiar” por “Tramkopliar”, “Caudre” por “Cuadre”, “Fragno” por “Fradno” o “Pamil” por “Pamir”. A los estudiantes",
     "Escritura al dictado. Los diez estudiantes escribieron alguna pseudopalabra de forma distinta al estímulo. A los estudiantes"),
    ("Tampoco permiten afirmar lateralidad cruzada, porque la prueba no incluyó tareas como el Test de Harris que usaron Aguirre-Medrano y González-López (2021) ni evaluó las confusiones entre b-d y p-q que estas autoras asocian con ella.",
     "Tampoco permiten afirmar ni descartar lateralidad cruzada, porque la prueba no incluyó tareas como el Test de Harris que usaron Aguirre-Medrano y González-López (2021). Las confusiones entre b-d y p-q que estas autoras asocian con ella, en cambio, no aparecieron en la lectura, aunque la lista incluía estímulos como “bapo”, “druma”, “blaper” o “diral”, lo que refuerza la idea de que las dificultades de este grupo son de orden fonológico más que visoespacial."),
    ("y en dejar una base para diseñar estrategias diferenciadas de conciencia fonológica, decodificación y escritura que convendría aplicar desde los primeros grados.",
     "y en la propuesta de trabajo por perfiles que se deriva de esos resultados, pensada para aplicarse desde los primeros grados."),
    ("tareas que los diez resolvieron bien. Los estudios revisados reportan buenos resultados con intervenciones multisensoriales, lúdicas e inclusivas, y el conjunto del trabajo apunta a la necesidad de diagnosticar temprano, adaptar las estrategias a cada contexto y formar mejor a los docentes.",
     "tareas que los diez resolvieron bien, y casi la mitad de los errores de lectura afectó grupos consonánticos con l o r. A partir de estos datos y de las intervenciones revisadas, que reportan buenos resultados con enfoques multisensoriales y lúdicos, se propone una secuencia de trabajo que va de la sílaba al fonema y a las sílabas trabadas."),
    ("tasks that all ten solved correctly. The studies reviewed report good results with multisensory, play-based and inclusive interventions, and the work as a whole points to the need for early diagnosis, strategies adapted to each context and better teacher preparation.",
     "tasks that all ten solved correctly, and almost half of the reading errors involved consonant clusters with l or r. Drawing on these data and on the interventions reviewed, which report good results with multisensory and play-based approaches, the study proposes a teaching sequence that moves from the syllable to the phoneme and to consonant clusters."),
]

TABLA2 = {
    "numero": "Tabla 2",
    "titulo": "Tipos de cambio en la lectura de pseudopalabras",
    "encabezado": ["Tipo de cambio", "Ejemplos (estímulo)", "Casos", "%"],
    "filas": [
        ["Alteración de grupos consonánticos con l o r", "Blused (Lused), Florir (Folir), Blumex (Brumex), Flanit (Franit), Fenir (Flenir), Pelos (Pleros), Brilus (Blisur)", "18", "43,9"],
        ["Sustitución o adición de otra consonante", "Digan (Dijan), Tutim (Tudim), Trampin (Trapin), Zeman (Zepan)", "7", "17,1"],
        ["Omisión o cambio de la consonante final", "Zepa y Zapa (Zepan), Plerox (Pleros)", "5", "12,2"],
        ["Lectura de “qu” como “cu”", "Cuome, Coume (Quome)", "3", "7,3"],
        ["Otras alteraciones de la l", "Zolil (Zonil), Used (Lused)", "3", "7,3"],
        ["Inversión de letras", "Tumid (Tudim)", "2", "4,9"],
        ["Otros cambios", "Consiul (Zonil), Quomi (Quome)", "3", "7,3"],
        ["Total", "", "41", "100"],
    ],
    "nota": "Cada caso corresponde a una pseudopalabra alterada por un estudiante. “Brilus” y “Coume” incluyen además una inversión. Elaboración propia a partir de la transcripción (Anexo A).",
}

REFS_NUEVAS = [  # (se inserta después de la referencia que empieza con..., segmentos)
    ("Carlino, P. (2005)", [("Coltheart, M., Rastle, K., Perry, C., Langdon, R., & Ziegler, J. (2001). DRC: A dual route cascaded model of visual word recognition and reading aloud. ", False), ("Psychological Review, 108", True), ("(1), 204–256. https://doi.org/10.1037/0033-295X.108.1.204", False)]),
    ("Coltheart, M.", [("Cuetos, F., Rodríguez, B., Ruano, E., & Arribas, D. (2007). ", False), ("PROLEC-R. Batería de evaluación de los procesos lectores, revisada", True), (". TEA Ediciones.", False)]),
    ("Gutiérrez Fresneda", [("Jiménez González, J. E., & Jiménez Rodríguez, R. (1999). Errores en la escritura de sílabas con grupos consonánticos: Un estudio transversal. ", False), ("Psicothema, 11", True), ("(1), 125–135.", False)]),
    ("Jiménez González, J. E., & Jiménez Rodríguez", [("Jiménez González, J. E., & Ortiz González, M. R. (1998). ", False), ("Conciencia fonológica y aprendizaje de la lectura: Teoría, evaluación e intervención", True), (". Síntesis.", False)]),
    ("Núñez Peña", [("OECD. (2023). ", False), ("PISA 2022 results (Volume I): The state of learning and equity in education", True), (". OECD Publishing. https://doi.org/10.1787/53f23881-en", False)]),
    ("Pisco-Román, J.", [("Real Academia Española & Asociación de Academias de la Lengua Española. (2010). ", False), ("Ortografía de la lengua española", True), (". Espasa.", False)]),
]


def buscar(P, inicio):
    hallados = [p for p in P if p.text.strip().startswith(inicio)]
    assert len(hallados) == 1, (inicio, len(hallados))
    return hallados[0]


def main(src, dst):
    d = docx.Document(src)
    P = list(d.paragraphs)
    molde_h = {"Heading 1": buscar(P, "Discusión"), "Heading 2": buscar(P, "Participantes")}

    for viejo, nuevo in AJUSTES:
        if viejo == "__LISTA__":
            R.reemplazar(buscar(P, "lectura en voz alta de pseudopalabras;"), nuevo)
            continue
        p = next(p for p in P if viejo in p.text)
        R.reemplazar(p, p.text.replace(viejo, nuevo))

    for inicio, nuevos in INSERTAR:
        ancla = buscar(P, inicio)
        molde_p = ancla
        for texto, estilo in nuevos:
            if texto == "__TABLA2__":
                V4.TABLA = TABLA2
                ancla = V4.insertar_tabla(d, ancla, molde_p)
                continue
            ancla = V3.insertar_despues(ancla, texto, estilo, molde_h.get(estilo, molde_p))

    # propuesta pedagógica antes de las conclusiones
    conclusiones = buscar(P, "Conclusiones")
    ancla = conclusiones._p.getprevious()
    ancla = Paragraph(ancla, conclusiones._parent)
    molde_p = buscar(P, "La variedad de perfiles")
    for texto, estilo in PROPUESTA:
        ancla = V3.insertar_despues(ancla, texto, estilo, molde_h.get(estilo, molde_p))

    # sigla OCDE
    V3.insertar_despues(buscar(P, "ONU\t"), "OCDE\t\t\tOrganización para la Cooperación y el Desarrollo Económicos", None, buscar(P, "ONU\t"))

    # referencias nuevas, en su lugar alfabético
    for despues_de, segmentos in REFS_NUEVAS:
        ref = [p for p in d.paragraphs if p.text.startswith(despues_de)]
        assert len(ref) == 1, despues_de
        nuevo = deepcopy(ref[0]._p)
        ref[0]._p.addnext(nuevo)
        V3.poner_runs(Paragraph(nuevo, ref[0]._parent), segmentos)
    d.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
