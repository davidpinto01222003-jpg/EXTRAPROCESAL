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
listan más abajo. El archivo queda configurado para que Word actualice sus campos al
abrirlo, con lo que la tabla de contenido y las listas de figuras y tablas se regeneran
solas.

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
| Marco teórico | Enlace entre el par de longitudes de onda de la literatura y el adoptado aquí |
| Capítulo nuevo antes del cronograma | Resultados, con Tablas 5 a 13 y Figuras 5 a 9, incluida la contextualización climática |
| Resultados esperados | Reescrito para no duplicar el capítulo de resultados |

El capítulo de resultados está redactado como trabajo desarrollado, no como avance
parcial: no habla de etapas, de material pendiente ni de lo que falta por hacer.

Además, sobre el documento completo: se retiró el resaltado amarillo que Zotero había
dejado en 127 puntos de la bibliografía, se renumeraron las tablas existentes y sus
referencias cruzadas, y se activó la actualización de campos al abrir el archivo, de
modo que Word regenera por sí solo la tabla de contenido y las listas de figuras y
tablas. Las tablas existentes pasaron a ser Tabla 4 las coordenadas de temperatura
superficial, Tabla 12 el cronograma y Tabla 13 el presupuesto.

## Contenido

| Carpeta | Qué hay |
|---|---|
| `fuentes/` | Anteproyecto en Word, compendio de la base de datos, archivo de datos por época y presentación teórica sobre espectros del fitoplancton |
| `analisis/` | Rutinas de Python que construyen la base, verifican el índice y corren la estadística; salidas completas en texto plano y bases derivadas en CSV; script de Node que arma el documento |
| `figuras/` | Las cinco figuras del análisis |
| `datos_sst/` | Serie satelital descargada y las tablas de climatología, anomalía y anomalía estandarizada |

Los libros de espectros de absorción (`ap`, `ad`, `aphy`, `aCDOM`) están en
`../datos/espectros_cdom/`.

## Cómo reproducir

```bash
cd analisis
python3 build_data.py     # base_unificada.csv desde el compendio y los espectros
python3 verifica.py       # control de calidad y reproducibilidad del índice
python3 analisis.py       # estadística completa, deja salida.txt
python3 figuras.py        # figuras finales
node gen_doc.js           # arma el informe
python3 integrar_en_anteproyecto.py   # integra el trabajo en el anteproyecto original
python3 anomalias_sst.py --entrada sst_mensual.csv   # anomalías estandarizadas de temperatura
```

Dependencias: pandas, numpy, scipy, matplotlib, openpyxl y el paquete npm `docx`.
Los scripts esperan los archivos originales en la ruta de subida; ajustar la constante
`U` en `build_data.py` para apuntarlos a `../fuentes/` y `../../datos/espectros_cdom/`.

## Resultados principales

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
- El muestreo cayó en una fase cálida: entre 2021 y 2023 la anomalía estandarizada de
  temperatura superficial promedió +0,606 frente a −0,091 en el resto de la serie de
  veintidós años (Mann Whitney, p = 0,0001). 2023 y 2024 son los años más cálidos del
  registro.

## Anomalías de temperatura superficial

`analisis/anomalias_sst.py` ejecuta los tres pasos de Santamaría del Ángel: climatología
mensual con su desviación típica, anomalía y anomalía estandarizada, más la figura de la
serie con las cinco campañas señaladas. Trae cargadas las ocho posiciones geográficas de
la tabla de coordenadas del anteproyecto.

El conjunto elegido es `jplMURSST41mday`, el análisis MUR fv04.1 a 0,01 grados, cerca de
un kilómetro, con paso mensual desde junio de 2002. Coincide con el periodo y la
resolución que el propio anteproyecto describe. El conjunto `erdMH1sstdmday` de MODIS
Aqua figura como obsoleto en el catálogo y solo llega hasta 2019, de modo que no sirve.

La serie ya está descargada y procesada. El archivo de origen y las salidas están en
`datos_sst/`. Para rehacer el cálculo:

```bash
python3 anomalias_sst.py --caja ../datos_sst/MUR_bahia_cartagena_2002_2024.csv
python3 figura_sst.py
```

Para una descarga nueva desde una máquina con salida a internet, `--urls` imprime las
direcciones ya armadas con las ocho coordenadas.

`--caja` toma un único CSV que cubra la bahía completa, busca la celda más cercana a
cada una de las ocho estaciones, informa a qué distancia quedó y sigue con el cálculo.
Es el camino más corto porque supone una sola descarga.

Los otros modos: `--carpeta` para ocho archivos, uno por estación; `--entrada` si la
tabla mensual ya viene armada; y `--descargar` cuando la red lo permite. Los cuatro
quedaron probados de extremo a extremo.
