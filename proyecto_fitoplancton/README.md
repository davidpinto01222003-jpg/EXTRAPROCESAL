# Proyecto de grado: índice Blue/Red en la Bahía de Cartagena

Trabajo de apoyo al anteproyecto de maestría de Omar Enrique Gamboa Posada.
Nada de esto se relaciona con los scripts jurídicos del resto del repositorio.

## Entregables

`Informe_avance_tareas_tutora.docx` — informe de avance sobre las ocho tareas que la
dirección del proyecto dejó señaladas en el apartado de tratamiento estadístico del
anteproyecto. Incluye el diagnóstico de hasta dónde llegaba el anteproyecto en cada
tarea, la ejecución real sobre los datos entregados y la redacción propuesta para
corregir el documento principal. El texto se escribió sin usar guiones.

`Anteproyecto_integrado_Omar_Gamboa.docx` — el anteproyecto original con el avance ya
incorporado en sus secciones correspondientes. Conserva estilos, campos SEQ, índices
automáticos y citas de Zotero del documento de origen. Los puntos de integración se
listan más abajo. Al abrirlo en Word conviene seleccionar todo y actualizar campos con
F9 para regenerar la tabla de contenido y las listas de figuras y tablas.

### Dónde quedó integrado cada aporte

| Sección del anteproyecto | Qué se integró |
|---|---|
| Resumen y Abstract | Corrección de 675 a 676 nm y cierre con los tres resultados preliminares |
| Planteamiento del problema, Hipótesis | Hipótesis reformulada hacia el gradiente estuarino y la variabilidad interanual |
| Metodología, Descripción de la base de datos | Párrafo de consolidación, Tabla 1 de inventario y nota sobre las fechas de las campañas de 2023 |
| Metodología, Cálculo del índice | Corrección a 440 y 676 nm, párrafo del par elegido y Tabla 2 de sensibilidad |
| Metodología, subsección nueva | Control de calidad de los datos, con cuatro criterios y Tabla 3 |
| Metodología, tratamiento estadístico | Diseño de pruebas corregido, entorno de trabajo y definición operativa de zonas. Se retiró el resaltado del bloque |
| Metodología, Modelado descriptivo | Criterio de respaldo si el ajuste no alcanza un R cuadrado de 0,4 |
| Capítulo nuevo antes del cronograma | Resultados Preliminares, con Tablas 5 a 11 y Figuras 5 a 7 |
| Cronograma | Subsección de estado de avance |
| Resultados esperados | Párrafo que separa lo ya obtenido de lo pendiente |

Las tablas existentes se renumeraron en consecuencia: coordenadas de temperatura
superficial pasa a Tabla 4, cronograma a Tabla 12 y presupuesto a Tabla 13.

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
node gen_doc.js           # arma el informe de avance
python3 integrar_en_anteproyecto.py   # integra el avance en el anteproyecto original
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
