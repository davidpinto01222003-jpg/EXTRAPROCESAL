# Simulador aproximado de Compilatio

No existe un clon open source de Compilatio. Su motor de IA y su base de fuentes son privados. Por eso este simulador **aprende de los informes reales** de Compilatio qué frases del trabajo marcó, y con eso estima el porcentaje de versiones nuevas del mismo trabajo. Hoy está entrenado con dos informes:

| Versión | Informe | IA | Similitud | Idiomas | Total |
|---|---|---|---|---|---|
| Original | `datos/reporte_compilatio.pdf` (Compilatio Magister+) | 27 % | 6 % | 3 % | 35 % |
| v4 | `datos/reporte_compilatio_v4.pdf` (Compilatio Studium) | 22 % | 4 % | 0 % | 26 % |

## Uso
```
pip install python-docx pdfplumber scikit-learn numpy
python simulador.py entrenar datos/original.docx datos/reporte_compilatio.pdf datos/v4.docx datos/reporte_compilatio_v4.pdf
python simulador.py evaluar "../documentos_academicos/Trabajo de grado Suarez y Rivera CORREGIDO v5.docx" --html reporte.html
```
`entrenar` recibe pares (trabajo .docx, informe .pdf) en orden cronológico. Cuando llegue un informe nuevo, basta con agregar el par al final. Los porcentajes del resumen se leen solos del PDF.

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

## Lo que enseñaron los dos informes
- Las oraciones reescritas en la v4 salieron **más** marcadas (37 %) que las del original (32 %). Partir el texto en frases muy cortas intercaladas con largas, y agregar preguntas retóricas, no ayuda: lo que más pesa en las marcas es que el párrafo tenga muchas oraciones y mucha variación de largo.
- Las oraciones con cita (autor, año) salen menos marcadas, y las cargadas de nominalizaciones (-ción, -miento, -idad) salen más.

## Detectores de Hugging Face (opcional)
`detectores_hf.py` calcula puntajes de tres detectores: `CradeyMH/detector-ia-espanol`, `pandrei7/autextification-upb-mtl` y Binoculars con `Qwen/Qwen2.5-0.5B`/`-Instruct`. Se activan con `entrenar ... --hf`, y requieren `torch`, `transformers`, `safetensors` y `huggingface_hub` y unos 4 GB de descarga. Dentro del original parecían ayudar (AUC 0,71), pero predijeron mal las marcas reales de la v4 (0,57), así que el modelo por defecto no los usa. Además, puntúan al revés que Compilatio: los párrafos que Compilatio marcó son los que esos detectores consideran más humanos.

## Limitaciones
- Es un **indicador**, no el resultado oficial.
- El primer informe salió de Compilatio Magister+ y el segundo de Compilatio Studium. Si el motor de IA difiere entre versiones, parte de la diferencia entre informes puede venir de ahí.
- Solo vale para este trabajo. Para otro documento hay que entrenarlo con sus propios informes.
