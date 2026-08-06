# Proyecto de grado: índice Blue/Red en la Bahía de Cartagena

Trabajo de apoyo al anteproyecto de maestría de Omar Enrique Gamboa Posada.
Nada de esto se relaciona con los scripts jurídicos del resto del repositorio.

## Entregable

`Informe_avance_tareas_tutora.docx` — informe de avance sobre las ocho tareas que la
dirección del proyecto dejó señaladas en el apartado de tratamiento estadístico del
anteproyecto. Incluye el diagnóstico de hasta dónde llegaba el anteproyecto en cada
tarea, la ejecución real sobre los datos entregados y la redacción propuesta para
corregir el documento principal. El texto se escribió sin usar guiones.

## Contenido

| Carpeta | Qué hay |
|---|---|
| `fuentes/` | Anteproyecto en Word, compendio de la base de datos, archivo de datos por época y presentación teórica sobre espectros del fitoplancton |
| `analisis/` | Rutinas de Python que construyen la base, verifican el índice y corren la estadística; salidas completas en texto plano y bases derivadas en CSV; script de Node que arma el documento |
| `figuras/` | Las tres figuras del informe |

Los libros de espectros de absorción (`ap`, `ad`, `aphy`, `aCDOM`) están en
`../datos/espectros_cdom/`.

## Cómo reproducir

```bash
cd analisis
python3 build_data.py     # base_unificada.csv desde el compendio y los espectros
python3 verifica.py       # control de calidad y reproducibilidad del índice
python3 analisis.py       # estadística completa, deja salida.txt
python3 figuras.py        # figuras finales
node gen_doc.js           # arma el .docx
```

Dependencias: pandas, numpy, scipy, matplotlib, openpyxl y el paquete npm `docx`.
Los scripts esperan los archivos originales en la ruta de subida; ajustar la constante
`U` en `build_data.py` para apuntarlos a `../fuentes/` y `../../datos/espectros_cdom/`.

## Resultados principales del avance

- 131 registros en cinco campañas; 19 excluidos por control de calidad (14,5 por ciento),
  casi todos en estaciones de turbidez alta donde la absorción del fitoplancton se obtiene
  por diferencia entre señales muy grandes y parecidas.
- El índice del compendio se reproduce exactamente desde los espectros con el par
  aphy(440)/aphy(676). Las columnas rotuladas "443-677" están calculadas con 676 nm.
- Cambiar el par de longitudes de onda reclasifica hasta el 16,9 por ciento de las estaciones.
- No hay diferencia significativa del índice entre época seca y época lluviosa (p = 0,126)
  ni asociación entre clase de tamaño y época (p = 0,806). La hipótesis del anteproyecto
  no se sostiene con estos datos.
- Sí hay diferencia entre campañas (Kruskal Wallis, p = 0,001) y entre zonas de influencia
  del Canal del Dique definidas por salinidad (Mann Whitney, p = 0,004).
- El modelo de regresión con temperatura, salinidad y turbidez explica menos del 10 por
  ciento de la variabilidad del índice; faltan clorofila a y nutrientes.
