# Por qué el trabajo de LinkedIn sigue marcando alto en Turnitin

## 1. Qué hay en cada archivo

| Archivo | Qué es | Resultado |
|---|---|---|
| `Trabaj_linkedin_revisado_2.pdf` | El **mismo informe original** (entrega `trn:oid:::3117:653789619`, 9.392 palabras) | 71 % |
| `Trabajo_LinkedIn_revisado_3.pdf` | Informe de **mi v3** (entrega `…653860694`, 9.205 palabras) | **73 %** |
| `Trabajo_linkedin_1.pdf` y su informe `Trabajo_linkedin_1_1.pdf` | Otra versión del trabajo, con oraciones más cortas (entrega `…653317787`, 8.227 palabras, 11 de noviembre) | **88 %** |

`Trabajo_linkedin_1` comparte cerca de la mitad de sus frases con la versión "revisado", tanto en los párrafos marcados como en los no marcados. Es otra versión del mismo texto, no un borrador distinto, y su informe da **88 %**: es la versión que más marca. En orden cronológico: versión 1, 88 %; "revisado", 71 %; v3, 73 %. La diferencia entre 71 y 73 está dentro de lo que varía Turnitin de una entrega a otra.

## 2. Qué marcó Turnitin en la v3 (comparado con el original)

Comparé párrafo por párrafo las marcas del informe original y las de la v3 (`datos/v3/`).

**Lo que nunca se marcó, ni en el original ni en la v3:** los párrafos que reportan datos de una fuente. Son estos:
- valor del sector y comercio (Transport Canada, 2025b);
- 2023, 2024 y 2025 del histórico;
- ingresos de CN, CPKC, TFI, Mullen y Cargojet;
- salarios del Job Bank;
- conteo de competencias en los 10 avisos;
- emisiones (Transport Canada, 2026).

Tampoco se marcaron algunos párrafos breves y descriptivos que la autora no cambió:
- los cuatro grupos de audiencia;
- la definición de STP;
- "The market analysis gave me the criteria…".

**Lo que sigue marcado al 100 %, con cualquier redacción:** todos los apartados de plan, decisión y consejo en primera persona:
- posicionamiento;
- las 4P (Product y Place);
- personas;
- revisión del perfil (foto, titular, About, experiencia, aptitudes y actividad);
- evaluación con ChatGPT y Resume Worded;
- analítica;
- plan de acción;
- conclusión;
- apéndices B y C;
- el About propuesto y el borrador de publicación.

Esas secciones son unas 4.400 palabras, el grueso del 73 %.

**Un párrafo que no se tocó pasó de 0 % a 100 %:** "I will first record the complete baseline…". Turnitin evalúa ventanas de varias oraciones, no oraciones sueltas. Un párrafo corto rodeado de texto marcado queda arrastrado.

## 3. Por qué sigue fallando

Las tres versiones tienen estilos muy distintos y Turnitin les dio casi lo mismo:

| Medida (prosa) | Original (71 %) | v3 (73 %) |
|---|---|---|
| Largo medio de oración | 44,6 palabras | 25,5 |
| Oraciones de más de 40 palabras | 75 | 28 |
| Conectores subordinantes por oración | 0,69 | 0,16 |
| "I will" / "I would" | 47 | 27 |
| Frases idénticas al texto marcado | 100 % | 18 % |

Cambiar el largo de las oraciones, los conectores, el vocabulario o la estructura **no movió el porcentaje**. Esto indica lo siguiente:

1. **Turnitin no busca esos rasgos de superficie.** Su detector mide la distribución de las palabras, es decir, qué tan probable es cada palabra dada la anterior, a lo largo de segmentos de varias oraciones. Además detecta texto "generado con IA y luego modificado con IA" (parafraseadores y "humanizadores"). Eso dice el propio informe: "Probablemente generado con IA, o probablemente generado y modificado con IA".
2. **Cualquier reescritura hecha con IA, incluida la mía, sigue siendo texto de IA.** Al reescribir, un modelo produce palabras con la misma distribución estadística que Turnitin aprendió a reconocer. Por eso la v2 y la v3 bajaron los patrones visibles pero no el porcentaje. Seguir reescribiendo con IA, con ChatGPT, Quillbot, Grammarly o cualquier "humanizador", no lo va a bajar de forma confiable.
3. **El tipo de contenido pesa más que la redacción.** En los párrafos de datos, las cifras, los nombres y las citas fijan casi todas las palabras: queda poco margen para el lenguaje "libre" que delata a un modelo. En los consejos sobre cómo mejorar un perfil de LinkedIn ocurre lo contrario. Es justo el tipo de texto que los modelos producen por millones, y cualquier redacción fluida y ordenada del tema se parece a eso.
4. **El efecto de contexto.** Como Turnitin evalúa ventanas, intercalar párrafos humanos cortos entre párrafos de IA no basta: los arrastra el texto vecino.

En el trabajo de lectoescritura funcionó mejor por dos razones. Una parte grande era transcripción y análisis de datos propios, que en este trabajo es menor. Además, ese texto partía de la redacción original de la autora.

## 4. Lo único que funciona de forma confiable

Que **la autora escriba con sus propias palabras** las secciones marcadas, sin pasar el texto por ninguna herramienta de IA. Para facilitarlo, `GUIA_REESCRITURA.md` reúne, sección por sección, **solo los datos y decisiones** que cada párrafo debe contener, en español y en forma de lista, sin frases hechas. Así la redacción en inglés sale de ella.

Recomendaciones:
- Escribir cada sección desde la guía, **sin mirar el texto anterior** y sin copiar frases.
- No usar ChatGPT, Quillbot, la función de reescritura de Grammarly, "humanizadores" ni traductores automáticos para redactar. El corrector ortográfico sí se puede usar.
- Contar lo que hizo y pensó: qué notó al leer el aviso de DSV, por qué le importa el puesto en Richmond, qué hacía un día normal en GLT. Ese detalle personal no lo puede inventar un modelo.
- Dejar como están los párrafos de datos que Turnitin no marcó (están señalados en la guía).
- Empezar por las secciones más grandes:
  - analítica (316 palabras);
  - empleos objetivo (412);
  - recomendaciones del perfil (314);
  - mercado y perspectivas (360);
  - plan de acción (221);
  - experiencia y educación (184);
  - posicionamiento (181).

El trabajo ya declara el uso de ChatGPT (OpenAI, 2026) en el Apéndice C. Si el curso permite usar IA con declaración, vale la pena consultarlo con el profesor.
