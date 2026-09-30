# Simulador aproximado de Compilatio

No existe un clon open source de Compilatio. Su motor de IA y su base de fuentes son privados. Por eso este simulador **aprende de los informes reales** de Compilatio qué frases del trabajo marcó, y con eso estima el porcentaje de versiones nuevas del mismo trabajo. Hoy está entrenado con tres informes:

| Versión | Informe | IA | Similitud | Idiomas | Total |
|---|---|---|---|---|---|
| Original | `datos/reporte_compilatio.pdf` (Compilatio Magister+) | 27 % | 6 % | 3 % | 35 % |
| v4 | `datos/reporte_compilatio_v4.pdf` (Compilatio Studium) | 22 % | 4 % | 0 % | 26 % |
| Solo los 18 párrafos reescritos en la v5 (3.481 palabras) | `datos/reporte_compilatio_parrafos_v5.pdf` (Compilatio Studium) | 38 % | <1 % | 0 % | 39 % |

## Uso
```
pip install python-docx pdfplumber scikit-learn numpy
python simulador.py entrenar datos/original.docx datos/reporte_compilatio.pdf datos/v4.docx datos/reporte_compilatio_v4.pdf \
                             datos/parrafos_reescritos_v5.docx datos/reporte_compilatio_parrafos_v5.pdf
python simulador.py evaluar "../documentos_academicos/Trabajo de grado Suarez y Rivera CORREGIDO v5.docx" --html reporte.html
```
`entrenar` recibe pares (trabajo .docx, informe .pdf) en orden cronológico. Cuando llegue un informe nuevo, basta con agregar el par al final. Los porcentajes del resumen se leen solos del PDF. También sirven informes de fragmentos sueltos: entran al entrenamiento y a la memoria, pero los porcentajes de referencia siempre salen del último informe del trabajo completo.

## Cómo estima
- **IA**, oración por oración, combinando dos fuentes:
  - **Memoria:** si una oración está igual en el último trabajo analizado, hereda la marca real que le puso Compilatio. Entre el original y la v4, el 92 % de las oraciones que no cambiaron conservaron su marca.
  - **Modelo** (gradient boosting, `HistGradientBoostingClassifier`) para las oraciones nuevas o reescritas. Usa 29 rasgos. Unos son de la oración: largo, comas, conectores, nominalizaciones en -ción/-miento/-idad, citas (autor, año) y otros. Otros son del párrafo: cuántas oraciones tiene, cuánto varía su largo y en qué parte del documento está. El umbral se calibra para reproducir el porcentaje de los informes.
- **Similitud:** de los fragmentos que Compilatio encontró en fuentes externas en el último informe, revisa cuántos siguen en el texto fuera de comillas y escala el porcentaje de ese informe.
- **Idiomas no reconocidos:** se asume igual al último informe.

## Qué tanto acierta
| Prueba | AUC |
|---|---|
| Validación cruzada agrupada por párrafo, con los dos informes | 0,76 |
| Entrenar solo con el original y predecir las marcas reales de la v4 | 0,74 |
| El modelo anterior (solo estilo, 16 rasgos, regresión logística) en esa misma prueba | 0,63 |
| Detectores de Hugging Face en esa misma prueba | 0,57 |

Con memoria y modelo juntos, el simulador da 26,1 % para la v4 (22,1 % IA + 4,0 % similitud); Compilatio dio 26 %.

**Sobre el nivel del porcentaje.** El modelo ordena bien qué oraciones son más sospechosas, pero el nivel absoluto es menos seguro cuando cambia el estilo. Entrenado solo con el original, predecía 36,9 % de IA para la v4, y el real fue 22 %. Por eso, para una versión nueva conviene mirar dos números: la estimación con umbral y la estimación con valor esperado (la suma de las probabilidades), que resulta más pesimista.

**Advertencia sobre reescrituras guiadas por el simulador.** Para la v5 se reescribió pensando en lo que el modelo castigaba, y ese mismo modelo predijo mal cuáles frases iban a pasar: AUC 0,56 contra el informe de esos párrafos. Esperaba cerca de un 10 % marcado y Compilatio marcó el 38 %. Con ese informe ya incluido, el AUC en validación cruzada es 0,72. Por eso las marcas reales pesan más que la predicción, y conviene probar cada ronda de reescritura en Compilatio.

## Lo que enseñaron los informes
- Las oraciones reescritas en la v4 salieron **más** marcadas (37 %) que las del original (32 %). Partir el texto en frases muy cortas intercaladas con largas, y agregar preguntas retóricas, no ayuda: lo que más pesa en las marcas es que el párrafo tenga muchas oraciones y mucha variación de largo.
- Las oraciones con cita (autor, año) salen menos marcadas, y las cargadas de nominalizaciones (-ción, -miento, -idad) salen más.
- En el informe de los párrafos reescritos pasaron limpios los anclados en datos concretos: el Abstract, la segunda estrategia (570 integrantes, p. 82), el marco normativo (leyes y años), Inglaterra (fechas y cifras del Home Office) y la primera parte del caso 1. Siguieron marcados los de argumentación general sin datos: la zona gris y la causa extraña, los efectos económicos, la jurisprudencia, la hipótesis y las interpretaciones de las conclusiones 03 y 05.

## Detectores de Hugging Face (opcional)
`detectores_hf.py` calcula puntajes de tres detectores: `CradeyMH/detector-ia-espanol`, `pandrei7/autextification-upb-mtl` y Binoculars con `Qwen/Qwen2.5-0.5B`/`-Instruct`. Se activan con `entrenar ... --hf`, y requieren `torch`, `transformers`, `safetensors` y `huggingface_hub` y unos 4 GB de descarga. Dentro del original parecían ayudar (AUC 0,71), pero predijeron mal las marcas reales de la v4 (0,57), así que el modelo por defecto no los usa. Además, puntúan al revés que Compilatio: los párrafos que Compilatio marcó son los que esos detectores consideran más humanos.

## Limitaciones
- Es un **indicador**, no el resultado oficial.
- El primer informe salió de Compilatio Magister+ y el segundo de Compilatio Studium. Si el motor de IA difiere entre versiones, parte de la diferencia entre informes puede venir de ahí.
- Solo vale para este trabajo. Para otro documento hay que entrenarlo con sus propios informes.
