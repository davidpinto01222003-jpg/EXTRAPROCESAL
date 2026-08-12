# Variables ópticas recuperadas de los libros de espectros

Estos datos no estaban en el compendio ni en el archivo por época. Salieron de bloques
laterales de los propios libros de espectros, que en una primera lectura pasaron
inadvertidos.

| Archivo | Qué contiene | Alcance |
|---|---|---|
| `seca2021_con_clorofila.csv` | Clorofila a, aphy a 440 nm y coeficiente de absorción específico, junto con el índice y las variables ambientales | Campaña de época seca de 2021, 31 estaciones con clorofila, 19 tras el control de calidad |
| `aportes_443.csv` | Absorción de detritus, fitoplancton y CDOM a 443 nm, en valor absoluto y en porcentaje | Cuatro campañas con libro de espectros |
| `scdom_seca2021.csv` | Pendiente espectral del CDOM, del ajuste exponencial | Campaña de época seca de 2021, 32 estaciones |

## Dónde estaba cada cosa

- La clorofila a está en la hoja `aphy`, en un bloque a la derecha del principal, en la
  fila rotulada `[ ] Clorofila a =`. La etiqueta aparece en los cuatro libros, pero solo
  el de la campaña de 2021 trae valores.
- La hoja `aphy-Especifico` es el cociente entre el aphy y esa clorofila, lo que sirvió
  para confirmar el dato.
- El reparto de la absorción está en la hoja `ax(443)`, en tres bloques: valores
  absolutos, porcentajes con el aCDOM medido y porcentajes con el aCDOM modelado.
- La pendiente del CDOM está en la hoja `aCDOM` del libro de 2021, en el bloque del
  ajuste `f(x) = a*exp(b*x)`, donde `b` es el Scdom.

## Advertencia de calidad

En el libro de la campaña de época lluviosa de 2022, la fila del aCDOM a 443 nm de la
hoja `ax(443)` repite los valores del aphy en las 29 estaciones. Es un error de fórmula
del archivo de origen y por eso esa campaña queda fuera del reparto de la absorción.

## Lo que se verificó que no existe

Se recorrieron todas las celdas de las siete hojas de los siete libros, incluidas filas
y columnas ocultas y nombres definidos. No hay nitratos, ni fosfatos, ni silicatos, ni
sólidos suspendidos totales, ni coordenadas geográficas, ni fechas de campaña.
