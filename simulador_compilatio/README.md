# Simulador aproximado de Compilatio

No existe un clon open source de Compilatio. Su motor de IA y su base de fuentes son privados, y desde este entorno tampoco se pueden descargar detectores de Hugging Face porque la red lo bloquea. Por eso este simulador **aprende del informe real** de Compilatio (`datos/reporte_compilatio.pdf`) qué frases del trabajo original marcó, y con eso estima el porcentaje de versiones nuevas del mismo trabajo.

## Uso
```
pip install python-docx pdfplumber scikit-learn numpy
python simulador.py entrenar datos/original.docx datos/reporte_compilatio.pdf   # ya entrenado: modelo.pkl
python simulador.py evaluar "../documentos_academicos/Trabajo de grado Suarez y Rivera CORREGIDO v3.docx" --html reporte.html
```

## Cómo estima
- **IA:** regresión logística por oración con rasgos de estilo: largo, comas, conectores de manual, vocabulario genérico, marcas de voz personal, variación del largo de las oraciones y otros. Se entrena con las 343 oraciones del original, 116 de ellas subrayadas en azul por Compilatio, y se calibra para reproducir el 27 % real.
- **Similitud:** toma los fragmentos que Compilatio encontró en fuentes externas y revisa cuántos siguen en el texto fuera de comillas. Escala el 6 % original según esa proporción.
- **Idiomas no reconocidos:** se asume igual al informe (3 %).

## Limitaciones (importante)
- Es un **indicador**, no el resultado oficial. En validación cruzada distingue frases marcadas y no marcadas con un AUC de aproximadamente 0,67: sirve para ordenar qué párrafos revisar primero, no para predecir el número exacto.
- Tiende a marcar párrafos largos y con muchas comas aunque ya estén reescritos (por ejemplo, las conclusiones 01-06).
- Solo vale para este trabajo. Para otro documento hay que entrenarlo con su propio informe de Compilatio.

## Turnitin (`turnitin.py`)

Funciona igual que el simulador de Compilatio, pero calibrado con el informe de IA de Turnitin del trabajo de lectoescritura (`../trabajo_lectoescritura/reporte_turnitin.pdf`). Toma como marcado el texto resaltado en celeste.

```
python turnitin.py entrenar ../trabajo_lectoescritura/original.docx ../trabajo_lectoescritura/reporte_turnitin.pdf
python turnitin.py evaluar "../trabajo_lectoescritura/Trabajo lectoescritura v2.docx" --html reporte.html
python turnitin.py patrones   # tabla de patrones Turnitin vs. Compilatio
```

`evaluar` da dos cifras:
- **Modelo de estilo** (AUC ≈ 0,61): es la cota alta. Cambia poco cuando se reescribe, porque aprendió rasgos muy generales.
- **Texto marcado que sigue igual**: qué parte de lo que Turnitin marcó sigue textual en la versión nueva. Si Turnitin solo volviera a marcar eso, saldría el valor de la cota baja.
