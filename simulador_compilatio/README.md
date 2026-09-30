# Simulador aproximado de Compilatio

No existe un clon open source de Compilatio. Su motor de IA y su base de fuentes son privados. Por eso este simulador **aprende del informe real** de Compilatio (`datos/reporte_compilatio.pdf`) qué frases del trabajo original marcó, y con eso estima el porcentaje de versiones nuevas del mismo trabajo. Desde la versión actual, además de rasgos de estilo, usa como rasgos los puntajes de tres detectores de IA publicados en Hugging Face.

## Uso
```
pip install python-docx pdfplumber scikit-learn numpy
pip install torch --index-url https://download.pytorch.org/whl/cpu
pip install transformers safetensors huggingface_hub
python simulador.py entrenar datos/original.docx datos/reporte_compilatio.pdf   # ya entrenado: modelo.pkl
python simulador.py evaluar "../documentos_academicos/Trabajo de grado Suarez y Rivera CORREGIDO v4.docx" --html reporte.html
```
La primera evaluación de un texto nuevo descarga unos 4 GB de modelos y tarda algunos minutos en CPU. Los puntajes quedan en `cache_hf.json`, así que un párrafo que no cambió no se vuelve a calcular. `entrenar --sin-hf` entrena el modelo anterior, que solo usa rasgos de estilo.

## Cómo estima
- **IA:** regresión logística por oración. Se entrena con las 343 oraciones del original, 116 de ellas subrayadas en azul por Compilatio, y se calibra para reproducir el 27 % real. Usa dos grupos de rasgos:
  - *Estilo* (16): largo, comas, conectores de manual, vocabulario genérico, marcas de voz personal, variación del largo de las oraciones y otros.
  - *Detectores de Hugging Face* (4), calculados sobre el **párrafo** de cada oración (`detectores_hf.py`):
    - `CradeyMH/detector-ia-espanol`: XLM-RoBERTa large afinado para detectar IA en español.
    - `pandrei7/autextification-upb-mtl`: modelo de la competencia AuTexTification 2023 (inglés y español).
    - Binoculars (Hans et al., 2024) con `Qwen/Qwen2.5-0.5B` y `Qwen/Qwen2.5-0.5B-Instruct`, más la perplejidad del observador.
- **Similitud:** toma los fragmentos que Compilatio encontró en fuentes externas y revisa cuántos siguen en el texto fuera de comillas. Escala el 6 % original según esa proporción.
- **Idiomas no reconocidos:** se asume igual al informe (3 %).

## Qué tanto acierta
Validación cruzada de 5 particiones **agrupada por párrafo** (las oraciones de un mismo párrafo nunca quedan a la vez en entrenamiento y en prueba), promedio de 10 repeticiones:

| Rasgos | AUC |
|---|---|
| Solo estilo (modelo anterior) | 0,63 |
| Detectores HF por oración | 0,57 |
| Detectores HF por párrafo | 0,71 |
| **Estilo + detectores HF por párrafo (modelo actual)** | **0,71** |

El AUC de 0,67 que se había informado antes estaba inflado: mezclaba oraciones del mismo párrafo entre entrenamiento y prueba.

## Limitaciones (importante)
- Es un **indicador**, no el resultado oficial. Con un AUC de 0,71 sirve para ordenar qué párrafos revisar primero, no para predecir el número exacto.
- **Los detectores, por sí solos, puntúan al revés que Compilatio en este trabajo.** Los párrafos que Compilatio marcó son los que XLM-R y AuTexTification consideran *más humanos* (AUC 0,29-0,45 usados directamente). El modelo aprende esa relación con coeficiente negativo. Por eso, que un detector público diga "humano" no significa que Compilatio vaya a decir lo mismo. Esos detectores están entrenados con textos genéricos, no con prosa jurídica académica.
- Tiende a marcar párrafos largos y densos aunque ya estén reescritos (por ejemplo, las conclusiones 03, 04 y 06).
- Solo vale para este trabajo. Para otro documento hay que entrenarlo con su propio informe de Compilatio.
