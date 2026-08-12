const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
        Table, TableRow, TableCell, WidthType, ShadingType, convertInchesToTwip,
        Footer, PageNumber } = require('docx');

const F = 'Times New Roman';
const AZUL = '1F3864';
const GRIS = 'F4F6F9';

const p = (t, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after === undefined ? 140 : o.after, line: o.line || 300 },
  children: [new TextRun({ text: t, font: F, size: o.size || 24, bold: !!o.bold,
                           italics: !!o.i, color: o.color || '000000' })]
});
const h1 = (t) => new Paragraph({
  spacing: { before: 360, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL } },
  children: [new TextRun({ text: t, font: F, size: 28, bold: true, color: AZUL })]
});
const rot = (etiqueta, texto) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 130, line: 300 },
  children: [
    new TextRun({ text: etiqueta + '  ', font: F, size: 24, bold: true, color: AZUL }),
    new TextRun({ text: texto, font: F, size: 24 })
  ]
});
const li = (t) => new Paragraph({
  spacing: { after: 110, line: 300 },
  indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.2) },
  children: [new TextRun({ text: '•  ' + t, font: F, size: 24 })]
});
const celda = (t, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: { top: 70, bottom: 70, left: 100, right: 100 },
  children: [new Paragraph({ spacing: { after: 0, line: 260 },
    children: [new TextRun({ text: t, font: F, size: 21, bold: !!o.b,
                             color: o.color || '000000' })] })]
});
const tabla = (anchos, filas) => new Table({
  columnWidths: anchos,
  width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
    left: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
    right: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' }
  },
  rows: filas.map((fila, i) => new TableRow({
    children: fila.map((c, j) => celda(c, {
      w: anchos[j], b: i === 0, fill: i === 0 ? AZUL : (i % 2 === 0 ? GRIS : null),
      color: i === 0 ? 'FFFFFF' : '000000'
    }))
  }))
});

const c = [];

// portada breve
c.push(new Paragraph({ spacing: { after: 400 }, children: [] }));
c.push(p('CÓMO SE CUMPLIÓ CADA PUNTO', { size: 34, bold: true, align: AlignmentType.CENTER, color: AZUL, after: 100 }));
c.push(p('Lo que solicitó la dirección del proyecto y los pasos seguidos en cada tarea',
         { size: 26, bold: true, align: AlignmentType.CENTER, after: 320 }));
c.push(p('Uso del índice Blue/Red para la caracterización espacio temporal del fitoplancton en la Bahía de Cartagena y su relación con variables ambientales y climáticas',
         { size: 22, i: true, align: AlignmentType.CENTER, after: 300 }));
c.push(p('Omar Enrique Gamboa Posada', { size: 24, bold: true, align: AlignmentType.CENTER, after: 60 }));
c.push(p('Maestría en Oceanografía. Escuela Naval de Cadetes Almirante Padilla', { size: 22, align: AlignmentType.CENTER, after: 60 }));
c.push(p('Agosto de 2026', { size: 22, align: AlignmentType.CENTER, after: 320 }));
c.push(new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 6, color: AZUL } },
                       spacing: { after: 200 }, children: [] }));
c.push(p('Este documento recorre las ocho tareas señaladas sobre el apartado de tratamiento estadístico del anteproyecto. Cada una se presenta con lo que se pidió, el procedimiento seguido paso a paso y el resultado obtenido.',
         { size: 21, i: true, align: AlignmentType.CENTER }));
c.push(new Paragraph({ pageBreakBefore: true, children: [] }));

// resumen de estado
c.push(h1('Estado general'));
c.push(tabla([700, 4600, 3800], [
  ['Nº', 'Tarea', 'Estado'],
  ['1', 'Organizar y depurar la base de datos', 'Cumplida'],
  ['2', 'Calcular y verificar el índice Blue/Red', 'Cumplida'],
  ['3', 'Análisis exploratorio de datos', 'Cumplida'],
  ['4', 'Análisis estadístico bivariado', 'Cumplida'],
  ['5', 'Análisis multivariado', 'Cumplida'],
  ['6', 'Pruebas de hipótesis', 'Cumplida, con el diseño corregido'],
  ['7', 'Modelo descriptivo y validación', 'Ajustada. La partición no se ejecutó, con justificación'],
  ['8', 'Anomalías estandarizadas de temperatura', 'Cumplida']
]));

// ---- tarea 1
c.push(h1('1. Organizar y depurar la base de datos'));
c.push(rot('Se pidió:', 'organizar las bases de datos in situ del monitoreo y dejarlas listas para el análisis.'));
c.push(rot('Procedimiento:', 'tomé el compendio, que traía las cinco campañas repartidas en bloques dentro de dos hojas, y las uní en una sola tabla de 131 estaciones, con una fila por estación y columnas de época, año, código, temperatura, salinidad, turbidez e índice. Normalicé los códigos de estación para que todos tuvieran el mismo formato. Después revisé la integridad registro por registro: busqué celdas con error de cálculo, valores fuera de rango físico y ceros que en realidad eran datos ausentes.'));
c.push(rot('Resultado:', 'encontré 11 estaciones sin índice por errores de división en la hoja de origen, 8 con valores imposibles, entre ellos uno de 9,52 y otro de 10,20, una temperatura de 38,7 grados que ningún agua superficial de la bahía alcanza, y tres ceros que eran ausencias. La base de trabajo quedó en 112 registros válidos, es decir, se descartó el 14,5 por ciento. Los criterios de exclusión quedaron escritos en una subsección nueva de la metodología.'));

// ---- tarea 2
c.push(h1('2. Calcular y verificar el índice Blue/Red'));
c.push(rot('Se pidió:', 'aplicar el índice sobre los coeficientes de absorción del fitoplancton y clasificar por tamaños según Ciotti y Bricaud.'));
c.push(rot('Procedimiento:', 'en lugar de dar por bueno el índice que traía el compendio, lo recalculé desde el origen. Abrí los cuatro libros de espectros, extraje el valor de la absorción del fitoplancton en cada longitud de onda y realicé el cociente. Comparé mi resultado con la columna del compendio estación por estación. Después probé cinco combinaciones distintas de longitudes de onda y conté cuántas estaciones cambiaban de clase de tamaño con cada una.'));
c.push(rot('Resultado:', 'los valores coinciden de forma exacta, lo que confirma que la base está bien construida. Detecté que las columnas rotuladas como cociente entre 443 y 677 en tres campañas están calculadas en realidad con 676, un error de rótulo que debe corregirse en el archivo. Y comprobé que la elección del par no es un detalle menor: el par que figuraba en mi metodología, 443 y 675, reclasifica 14 de 83 estaciones, casi una de cada seis. Por eso fijé un único par, 440 y 676, y lo justifiqué en el texto.'));

// ---- tarea 3
c.push(h1('3. Análisis exploratorio de datos'));
c.push(rot('Se pidió:', 'histogramas, diagramas de caja y gráficos de dispersión, separados por época climática y por zonas de influencia del Canal del Dique.'));
c.push(rot('Procedimiento:', 'calculé la estadística descriptiva por campaña y construí las tres clases de gráfica solicitadas. Diagramas de caja del índice por campaña con los umbrales de clasificación marcados, gráfico de dispersión del índice contra la salinidad con la turbidez representada en color, e histogramas del índice, la salinidad y el logaritmo de la turbidez, en dos filas, una separando por época climática y otra por zona.'));
c.push(rot('Resultado:', 'la variabilidad entre campañas de una misma época resultó del mismo orden que la variabilidad entre épocas, hecho que condiciona toda la lectura posterior. La turbidez mostró una asimetría de 4,52, que baja a menos 0,51 al transformarla a logaritmo, y por eso se usa transformada en los análisis multivariados. Además, 66 de las 112 estaciones quedan por debajo del umbral del microfitoplancton.'));

// ---- tarea 4
c.push(h1('4. Análisis estadístico bivariado'));
c.push(rot('Se pidió:', 'coeficientes de correlación de Pearson o de Spearman, según el cumplimiento del supuesto de normalidad.'));
c.push(rot('Procedimiento:', 'lo primero fue resolver la elección que quedaba abierta. Apliqué la prueba de Shapiro Wilk a las cuatro variables y la prueba de Levene para la homogeneidad de varianzas entre épocas. Con ese resultado quedó decidido el uso de pruebas no paramétricas. Después calculé las correlaciones de Spearman sobre el conjunto completo y, por separado, dentro de cada época climática.'));
c.push(rot('Resultado:', 'la temperatura es compatible con la normalidad, con probabilidad de 0,298, pero la salinidad, la turbidez y el índice no lo son. El índice desciende cuando sube la temperatura, con rho de menos 0,273 y probabilidad de 0,004, y asciende con la salinidad, con rho de 0,240. Con la turbidez no hay relación significativa. El hallazgo más fino es que la relación con la salinidad desaparece dentro de la época seca y se convierte en la más fuerte del conjunto en la época lluviosa, con rho de 0,463: el gradiente de agua dulce actúa sobre la estructura de tamaños sobre todo cuando el canal descarga con mayor intensidad.'));

// ---- tarea 5
c.push(h1('5. Análisis multivariado'));
c.push(rot('Se pidió:', 'análisis de componentes principales y análisis de conglomerados sobre las estaciones de muestreo.'));
c.push(rot('Procedimiento:', 'estandaricé las cuatro variables, con la turbidez en logaritmo decimal, y apliqué componentes principales sobre las 109 estaciones con información completa. Después agrupé las estaciones con el método de Ward y calculé la correlación cofenética para verificar que el dendrograma representara bien las distancias originales.'));
c.push(rot('Resultado:', 'los dos primeros componentes reúnen el 69,4 por ciento de la varianza. El primero es un eje estuarino que opone la salinidad a la temperatura y la turbidez, y en el que el índice casi no participa. El segundo sí ordena el índice, en oposición a la turbidez. Los conglomerados separan tres ambientes: aguas marinas poco alteradas, pluma activa del canal y bahía interior. Este último grupo, el más numeroso con 52 estaciones, concentra el predominio del microfitoplancton con 44 de ellas. Ninguno de los tres grupos es exclusivo de una época climática.'));

// ---- tarea 6
c.push(h1('6. Pruebas de hipótesis'));
c.push(rot('Se pidió:', 'un análisis de varianza de medidas repetidas o la prueba de Kruskal Wallis para comparar las clases de tamaño entre épocas climáticas y entre zonas.'));
c.push(rot('Procedimiento:', 'antes de correr nada corregí el diseño. Las estaciones no se repiten de forma completa entre campañas, que van de 32 a 15, de modo que no se trata de medidas repetidas sino de muestras independientes desbalanceadas. Apliqué la prueba de Mann Whitney para las comparaciones entre dos grupos, la de Kruskal Wallis para las cinco campañas con corrección de Bonferroni en las comparaciones por pares, y la de chi cuadrado para la asociación entre clase de tamaño y época. Definí además la zona de alta influencia del canal como aquella con salinidad inferior a 25, porque la base entregada no incluye las coordenadas de las estaciones.'));
c.push(rot('Resultado:', 'entre la época seca y la lluviosa no hay diferencia significativa, con probabilidad de 0,126, y tampoco hay asociación entre la clase de tamaño y la época, con 0,806. Sí hay diferencia entre campañas, con 0,001, incluso entre dos campañas de la misma época seca. Y sí la hay entre zonas, con 0,004, con células de mayor tamaño donde el agua es menos salada. Estos resultados obligaron a reformular la hipótesis del anteproyecto: lo que ordena la estructura de tamaños es el gradiente del canal y la variabilidad interanual, no la alternancia entre épocas.'));

// ---- tarea 7
c.push(h1('7. Modelo descriptivo y validación'));
c.push(rot('Se pidió:', 'un modelo empírico descriptivo, con partición de 75 por ciento para entrenamiento y 25 para validación interna, evaluado con residuos, coeficiente de determinación y error cuadrático medio.'));
c.push(rot('Procedimiento:', 'ajusté una regresión múltiple del logaritmo del índice sobre la temperatura, la salinidad y el logaritmo de la turbidez, con las 109 estaciones completas, y evalué la calidad del ajuste antes de dividir el conjunto.'));
c.push(rot('Resultado:', 'el modelo explica menos del diez por ciento de la variabilidad del índice, con un coeficiente de determinación de 0,096, y solo la temperatura resulta significativa. Ante ese resultado decidí no ejecutar la partición, porque entrenar y validar sobre un ajuste tan bajo produce indicadores sin significado. Dejé escrito en la metodología el criterio de respaldo: si con el conjunto ampliado de variables el ajuste no alcanza un coeficiente de 0,4, se optará por un modelo de clasificación de la clase de tamaño dominante, con validación cruzada y matriz de confusión. Esta es una decisión que someto a consideración de la dirección del proyecto.'));

// ---- tarea 8
c.push(h1('8. Anomalías estandarizadas de temperatura superficial'));
c.push(rot('Se pidió:', 'construir la climatología mensual de la temperatura superficial del mar y su transformación en anomalías estandarizadas, para el periodo comprendido entre junio de 2002 y septiembre de 2024, en las ocho posiciones geográficas de la tabla de coordenadas.'));
c.push(rot('Procedimiento:', 'busqué el producto en el catálogo del servidor y descarté el de MODIS Aqua, que figura como obsoleto y se corta en 2019. Elegí el análisis MUR, versión fv04.1, con resolución de 0,01 grados, que corresponde a la resolución cercana a un kilómetro que describe mi metodología y cuya serie arranca en junio de 2002. Descargué un recuadro que cubre la bahía completa, verifiqué qué celdas correspondían a agua y cuáles estaban enmascaradas como tierra, asigné a cada estación su celda y apliqué los tres pasos del método: climatología de cada mes con su desviación típica, anomalía como diferencia frente a esa climatología, y estandarización dividiendo por la desviación típica del mes.'));
c.push(rot('Resultado:', 'se procesaron 267 compuestas mensuales y las ocho posiciones coincidieron con celdas de agua, sin necesidad de desplazarlas. La climatología tiene su mínimo en marzo, con 27,27 grados Celsius, y su máximo en septiembre, con 29,78. La serie se divide en dos regímenes, con predominio de anomalías negativas hasta 2010 y positivas desde 2015. El resultado de mayor consecuencia para el proyecto es que el muestreo se realizó en una fase cálida: entre 2021 y 2023 la anomalía estandarizada promedió 0,606, frente a menos 0,091 en el resto de la serie, con una probabilidad de 0,0001. Las cinco campañas no retratan una bahía en condiciones climáticas medias.'));
c.push(tabla([2400, 1900, 2100, 1600, 2100], [
  ['Campaña', 'Mes', 'Temperatura observada', 'Anomalía', 'Anomalía estandarizada'],
  ['Época seca 2021', 'abril de 2021', '28,22', '+0,35', '+0,82'],
  ['Época seca 2022', 'marzo de 2022', '27,62', '+0,35', '+0,61'],
  ['Época lluviosa 2022', 'octubre de 2022', '29,49', '−0,08', '−0,21'],
  ['Campaña de 2023', 'diciembre de 2023', '29,32', '+0,52', '+1,17']
]));

// ---- formato
c.push(h1('Formato del documento'));
c.push(p('Sobre el documento completo se aplicaron las normas APA en su séptima edición. Los márgenes quedaron en 2,54 centímetros por los cuatro lados, el interlineado en doble para el cuerpo, los pies y la bibliografía, y la alineación a la izquierda con sangría de primera línea de 1,27 centímetros. Los títulos de nivel uno quedaron centrados y en negrita, y 32 subtítulos que estaban como negrita suelta pasaron a ser títulos de nivel dos, con lo que ahora aparecen en la tabla de contenido. Las tablas quedaron con número en negrita, título en cursiva debajo, únicamente líneas horizontales, sin sombreado y con su nota al pie. Las figuras llevan número y título encima de la imagen y la fuente como nota debajo. La bibliografía quedó con sangría francesa y el número de página arriba a la derecha.'));

// ---- limites
c.push(h1('Límites que conviene declarar'));
c.push(p('Tres asuntos quedan abiertos y es preferible plantearlos de forma explícita.'));
c.push(li('Los análisis se realizaron con tres de las siete variables ambientales que enumera la metodología. La clorofila a, los nutrientes y los sólidos suspendidos totales no venían en la base entregada, y esa ausencia es la causa directa de que el modelo descriptivo no alcance un ajuste útil. Gestionar esas variables ante el Centro de Investigaciones Oceanográficas e Hidrográficas del Caribe es la acción de mayor prioridad.'));
c.push(li('La partición de 75 y 25 por ciento no se ejecutó, por la razón expuesta en la tarea siete.'));
c.push(li('Quedan dos verificaciones sobre los datos: cuál de las dos campañas de 2023 corresponde a junio y cuál a diciembre, porque la numeración consecutiva de las muestras contradice las etiquetas de la base, y las coordenadas de cada estación, que por ahora se sustituyen por el criterio de salinidad para separar las zonas. Además, el producto satelital no incluye compuesta mensual para junio de 2023, lo cual es un vacío de la fuente.'));

const doc = new Document({
  creator: 'Omar Enrique Gamboa Posada',
  title: 'Cómo se cumplió cada punto solicitado por la dirección del proyecto',
  description: 'Procedimiento seguido en cada una de las ocho tareas del apartado de tratamiento estadístico',
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
                          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES],
                               font: F, size: 18, color: '666666' })] })] }) },
    children: c
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('Como_se_cumplio_cada_punto.docx', b);
  console.log('documento generado', b.length, 'bytes');
});
