const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageOrientation,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, ImageRun,
  Header, Footer, PageNumber, TabStopType, convertInchesToTwip
} = require('docx');

const FONT = 'Times New Roman';
const AZUL = '1F3864';
const GRIS = 'F2F2F2';
const VERDE = 'E2EFDA';
const AMAR = 'FFF2CC';
const ROJO = 'FCE4EC';

function p(text, opts = {}) {
  const o = Object.assign({ size: 22, bold: false, italics: false, align: AlignmentType.JUSTIFIED,
                            spacing: { after: 120, line: 300 }, color: '000000', indent: null }, opts);
  return new Paragraph({
    alignment: o.align, spacing: o.spacing, indent: o.indent,
    children: [new TextRun({ text, font: FONT, size: o.size, bold: o.bold, italics: o.italics, color: o.color })]
  });
}
function rich(runs, opts = {}) {
  const o = Object.assign({ align: AlignmentType.JUSTIFIED, spacing: { after: 120, line: 300 }, indent: null }, opts);
  return new Paragraph({
    alignment: o.align, spacing: o.spacing, indent: o.indent,
    children: runs.map(r => new TextRun({ text: r.t, font: FONT, size: r.size || 22, bold: !!r.b, italics: !!r.i, color: r.c || '000000' }))
  });
}
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: AZUL })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: AZUL })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3, spacing: { before: 180, after: 100 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: '333333' })]
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'vinetas', level },
    spacing: { after: 80, line: 290 },
    children: [new TextRun({ text, font: FONT, size: 22 })]
  });
}
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT, spacing: { before: 60, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 18, italics: true, color: '444444' })]
  });
}
function cell(text, opts = {}) {
  const o = Object.assign({ b: false, fill: null, size: 18, align: AlignmentType.LEFT, w: null, color: '000000' }, opts);
  return new TableCell({
    width: o.w ? { size: o.w, type: WidthType.DXA } : undefined,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: String(text).split('\n').map(t => new Paragraph({
      alignment: o.align, spacing: { after: 0, line: 240 },
      children: [new TextRun({ text: t, font: FONT, size: o.size, bold: o.b, color: o.color })]
    }))
  });
}
function tabla(widths, filas, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    columnWidths: widths,
    width: { size: total, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
      left: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
      right: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'BBBBBB' },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'BBBBBB' }
    },
    rows: filas.map((fila, i) => new TableRow({
      tableHeader: i === 0,
      children: fila.map((c, j) => {
        const conf = typeof c === 'object' ? c : { t: c };
        return cell(conf.t, {
          b: conf.b !== undefined ? conf.b : i === 0,
          fill: conf.fill !== undefined ? conf.fill : (i === 0 ? AZUL : null),
          color: conf.color !== undefined ? conf.color : (i === 0 ? 'FFFFFF' : '000000'),
          align: conf.align || (j === 0 ? AlignmentType.LEFT : (opts.centrar ? AlignmentType.CENTER : AlignmentType.LEFT)),
          w: widths[j], size: conf.size || 18
        });
      })
    }))
  });
}
function figura(archivo, ancho, alto) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 160, after: 40 },
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(archivo), transformation: { width: ancho, height: alto } })]
  });
}
function espacio(n) { return new Paragraph({ spacing: { after: n }, children: [] }); }

// ============================ CONTENIDO ============================
const cuerpo = [];

// ---- portada ----
cuerpo.push(espacio(600));
cuerpo.push(p('INFORME DE AVANCE', { size: 36, bold: true, align: AlignmentType.CENTER, color: AZUL, spacing: { after: 120 } }));
cuerpo.push(p('Ejecución de las tareas asignadas por la dirección del proyecto y revisión crítica del anteproyecto',
  { size: 26, bold: true, align: AlignmentType.CENTER, spacing: { after: 400 } }));
cuerpo.push(p('Uso del índice Blue/Red para la caracterización espacio temporal del fitoplancton en la Bahía de Cartagena y su relación con variables ambientales y climáticas',
  { size: 24, italics: true, align: AlignmentType.CENTER, spacing: { after: 500 } }));
cuerpo.push(p('Omar Enrique Gamboa Posada', { size: 24, bold: true, align: AlignmentType.CENTER, spacing: { after: 80 } }));
cuerpo.push(p('Maestría en Oceanografía', { size: 22, align: AlignmentType.CENTER, spacing: { after: 40 } }));
cuerpo.push(p('Escuela Naval de Cadetes Almirante Padilla. Facultad de Oceanografía Física', { size: 22, align: AlignmentType.CENTER, spacing: { after: 300 } }));
cuerpo.push(p('Dirección del proyecto: Ph.D. Stella Patricia Betancur Turizo', { size: 22, align: AlignmentType.CENTER, spacing: { after: 40 } }));
cuerpo.push(p('Cartagena de Indias, agosto de 2026', { size: 22, align: AlignmentType.CENTER, spacing: { after: 400 } }));
cuerpo.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: AZUL } },
  spacing: { after: 200 }, children: []
}));
cuerpo.push(p('Documento independiente del anteproyecto. Reúne el trabajo efectivamente ejecutado sobre la base de datos entregada, el diagnóstico del estado en que el anteproyecto dejaba cada tarea y las mejoras concretas que se proponen para el texto.',
  { size: 20, italics: true, align: AlignmentType.CENTER }));
cuerpo.push(new Paragraph({ pageBreakBefore: true, children: [] }));

// ---- 1. objeto ----
cuerpo.push(h1('1. Objeto del documento y materiales revisados'));
cuerpo.push(p('Este informe responde a las tareas asignadas por la dirección del proyecto, señaladas sobre el apartado de tratamiento estadístico del anteproyecto y acompañadas de la entrega del material de trabajo. El documento no repite el anteproyecto: verifica qué parte de esas tareas estaba realmente desarrollada, ejecuta las que solo estaban enunciadas y propone la redacción con la que deberían quedar incorporadas.'));
cuerpo.push(p('Los materiales revisados fueron los siguientes.'));
cuerpo.push(tabla([2600, 3400, 3000], [
  ['Material', 'Contenido verificado', 'Uso en este avance'],
  ['Anteproyecto en formato Word', 'Texto completo, 412 párrafos, sin control de cambios ni comentarios insertados. El apartado de análisis de condiciones abióticas y tratamiento estadístico aparece resaltado en su totalidad.', 'Base del diagnóstico del capítulo 3'],
  ['Compendio de base de datos', 'Cinco campañas y 131 registros de estación con temperatura, salinidad, turbidez e índice Blue/Red en dos pares de longitudes de onda', 'Fuente principal del análisis'],
  ['Archivo de datos por época', 'Subconjunto del compendio organizado en cinco hojas por campaña', 'Verificación cruzada'],
  ['Cuatro libros de espectros de absorción', 'Espectros de ap, ad, aphy y aCDOM entre 800 y 400 nm para 91 estaciones', 'Recálculo y verificación independiente del índice'],
  ['Presentación teórica sobre espectros del fitoplancton', 'Propiedades ópticas del agua, grupos taxonómicos, efecto de empaquetamiento y fundamento del cociente azul rojo', 'Marco conceptual del control de calidad']
], { centrar: false }));
cuerpo.push(caption('Tabla 1. Material entregado por la dirección del proyecto y uso dado a cada archivo en este avance.'));
cuerpo.push(p('Una precisión sobre el material: de los cinco libros recibidos, dos son copias idénticas del mismo archivo de la campaña de sequía de 2022, comprobado por su firma digital. La campaña de sequía de 2023 no cuenta con libro de espectros, de modo que sus 32 estaciones solo pudieron analizarse a partir del índice ya calculado en el compendio.'));

// ---- 2. tareas ----
cuerpo.push(h1('2. Tareas identificadas'));
cuerpo.push(p('El bloque resaltado en el anteproyecto define ocho tareas verificables. Se enumeran a continuación con el identificador que se usa en el resto del informe.'));
cuerpo.push(tabla([900, 3600, 4500], [
  ['Id', 'Tarea', 'Alcance esperado'],
  ['T1', 'Organización y depuración de la base de datos', 'Una sola tabla analizable con todas las campañas, con criterios explícitos para datos ausentes y valores anómalos'],
  ['T2', 'Cálculo y verificación del índice Blue/Red', 'Índice reproducible desde los espectros de aphy, con el par de longitudes de onda definido sin ambigüedad'],
  ['T3', 'Análisis exploratorio de datos', 'Distribuciones, dispersión y valores atípicos por época climática y por zona'],
  ['T4', 'Análisis estadístico bivariado', 'Correlaciones de Pearson o de Spearman entre variables ambientales e índice, según el cumplimiento de la normalidad'],
  ['T5', 'Análisis multivariado', 'Componentes principales y conglomerados sobre las estaciones de muestreo'],
  ['T6', 'Pruebas de hipótesis', 'Comparación de las clases de tamaño entre épocas climáticas y entre zonas de influencia del Canal del Dique'],
  ['T7', 'Modelo descriptivo y validación', 'Modelo empírico entre condiciones abióticas e índice, con partición de entrenamiento y validación'],
  ['T8', 'Anomalías estandarizadas de la temperatura superficial del mar', 'Climatología mensual y transformación Z de las series satelitales entre 2002 y 2024']
], { centrar: false }));
cuerpo.push(caption('Tabla 2. Tareas derivadas del apartado señalado por la dirección del proyecto.'));

// ---- 3. diagnostico ----
cuerpo.push(h1('3. Hasta dónde llegaba el anteproyecto'));
cuerpo.push(p('El anteproyecto describe con detalle qué se va a hacer, pero en ninguna de las ocho tareas presenta resultados. La totalidad del apartado está redactada en tiempo futuro, no aparece ningún valor calculado, ninguna tabla de resultados, ningún gráfico de datos propios y ningún criterio numérico de exclusión. El diagnóstico por tarea es el siguiente.'));
cuerpo.push(tabla([700, 2400, 3800, 2100], [
  ['Id', 'Lo que contenía el anteproyecto', 'Vacío detectado', 'Estado'],
  ['T1', 'Menciona la base de datos del CIOH y enumera las variables disponibles', 'No define el número de registros, no reporta datos ausentes ni valores imposibles, no unifica campañas', { t: 'Solo enunciada', fill: ROJO }],
  ['T2', 'Presenta la ecuación de absorción del particulado y los tres umbrales de tamaño', 'El resumen usa 440 y 675 nm, la metodología usa 443 y 675 nm y los datos entregados usan 440 y 676 nm. La diferencia cambia la clasificación', { t: 'Enunciada con inconsistencia', fill: ROJO }],
  ['T3', 'Anuncia histogramas, diagramas de caja y gráficos de dispersión', 'Ningún gráfico ni estadística descriptiva del conjunto real', { t: 'Solo enunciada', fill: ROJO }],
  ['T4', 'Anuncia correlaciones de Pearson o de Spearman', 'No hay coeficientes ni valores de probabilidad', { t: 'Solo enunciada', fill: ROJO }],
  ['T5', 'Anuncia componentes principales y conglomerados', 'No hay varianza explicada, cargas ni número de grupos', { t: 'Solo enunciada', fill: ROJO }],
  ['T6', 'Anuncia análisis de varianza de medidas repetidas o prueba de Kruskal Wallis', 'El diseño no corresponde a medidas repetidas porque las estaciones no se repiten completas entre campañas. La hipótesis del anteproyecto no había sido contrastada', { t: 'Enunciada con error de diseño', fill: ROJO }],
  ['T7', 'Describe la partición en 75 y 25 por ciento y los criterios de ajuste', 'No hay variables predictoras seleccionadas ni ajuste preliminar que justifique la estructura del modelo', { t: 'Solo enunciada', fill: ROJO }],
  ['T8', 'Describe con precisión los tres pasos de la anomalía estandarizada y la tabla de las ocho posiciones geográficas', 'Es la tarea mejor especificada del apartado. Faltaba únicamente la serie satelital, que se obtuvo y se procesó', { t: 'Enunciada y ahora ejecutada', fill: VERDE }]
], { centrar: false }));
cuerpo.push(caption('Tabla 3. Diagnóstico del grado de desarrollo de cada tarea en el anteproyecto.'));
cuerpo.push(p('Conviene decirlo con claridad porque orienta el trabajo que sigue: el anteproyecto tenía un protocolo bien escrito y ninguna ejecución. Lo que este informe aporta es la ejecución sobre los datos reales, junto con los problemas de la base que solo aparecen cuando se la trabaja.'));

// ---- 4. avance ----
cuerpo.push(new Paragraph({ pageBreakBefore: true, children: [] }));
cuerpo.push(h1('4. Avance ejecutado'));

cuerpo.push(h2('4.1 Consolidación y depuración de la base de datos (T1)'));
cuerpo.push(p('Se unificaron las cinco campañas en una sola tabla de 131 registros, con las variables de época climática, año, código de estación, temperatura superficial, salinidad, turbidez e índice Blue/Red en sus dos versiones. A esa tabla se le aplicó una revisión de integridad que arrojó los siguientes hallazgos.'));
cuerpo.push(tabla([4200, 1400, 3400], [
  ['Hallazgo', 'Registros', 'Tratamiento aplicado'],
  ['Índice no calculado en la hoja de origen, con errores de división por cero o de valor', '11', 'Se verificó que tampoco existe espectro de aphy para esas estaciones. Se excluyen y se declaran como ausentes'],
  ['Índice con valor imposible para el marco de Ciotti y Bricaud, por debajo de 1,0 o por encima de 6,0', '8', 'Se excluyen del análisis y se documentan uno a uno'],
  ['Temperatura fuera del rango físico esperado para aguas superficiales de la bahía', '1', 'Estación SC 21 0113, con 38,7 grados Celsius. Se marca como error de registro'],
  ['Salinidad y turbidez consignadas como cero', '3', 'Se distingue entre el cero real de agua dulce y el cero usado como dato ausente'],
  ['Registros con las tres variables ambientales vacías', '1', 'Estación SC 21 0104. Se conserva solo para el conteo de clases de tamaño'],
  ['Registros válidos para el análisis estadístico', '112', 'Base de trabajo definitiva']
], { centrar: false }));
cuerpo.push(caption('Tabla 4. Resultado de la depuración de la base unificada.'));
cuerpo.push(p('Las ocho estaciones con índice imposible no son un detalle menor. Siete de ellas corresponden a condiciones de turbidez elevada, en las que el espectro de absorción del fitoplancton se obtiene como diferencia entre dos señales grandes y muy parecidas, de modo que el resultado queda dominado por el error. Es el caso extremo de la estación SC 21 0113, con 242 unidades nefelométricas de turbidez y un índice de 9,52, y de la estación SC 22 0121, con 183,5 unidades y un índice de 10,20. En sentido contrario, las estaciones SC 21 0163 y SC 21 0165 arrojan índices de 0,67 y 0,72, valores que ningún espectro de fitoplancton real produce.'));
cuerpo.push(p('Se añadió un segundo indicador de confiabilidad que el anteproyecto no contemplaba: la fracción de la absorción del particulado que corresponde al material no algal a 440 nm. Su mediana es de 0,52 en las 91 estaciones con espectro disponible, valor razonable, pero la estación SC 22 0444 alcanza 0,98, lo que significa que su absorción de fitoplancton es prácticamente un residuo numérico. Su índice de 4,04 debe descartarse en la próxima depuración, aunque caiga dentro del rango admisible. En este avance se conservó, y su exclusión no altera ninguna conclusión: la comparación entre épocas climáticas pasa de una probabilidad de 0,126 a una de 0,082, sin alcanzar significación, y el contraste entre zonas se mantiene significativo con una probabilidad de 0,003.'));

cuerpo.push(h2('4.2 Verificación del cálculo del índice (T2)'));
cuerpo.push(p('El índice del compendio se reprodujo desde los espectros de aphy para las 91 estaciones que cuentan con libro espectral. La coincidencia es exacta y confirma dos cosas.'));
cuerpo.push(bullet('La columna rotulada como cociente entre 440 y 676 nm corresponde efectivamente a aphy(440) dividido entre aphy(676).'));
cuerpo.push(bullet('La segunda columna, rotulada en tres de las cinco campañas como cociente entre 443 y 677 nm, está calculada en realidad con el denominador de 676 nm en las cinco campañas. El rótulo es incorrecto en las hojas de la época lluviosa de 2022, la época lluviosa de 2023 y la época seca de 2023.'));
cuerpo.push(p('Sobre esa base se cuantificó algo que el anteproyecto no advierte: cuánto depende la clasificación de tamaños del par de longitudes de onda elegido. Se calcularon cinco combinaciones sobre las 83 estaciones válidas con espectro.'));
cuerpo.push(tabla([2000, 1500, 1500, 2100, 2000], [
  ['Par de longitudes de onda', 'Mediana', 'Desviación', 'Diferencia mediana frente al par de referencia', 'Estaciones que cambian de clase'],
  ['440 y 676 nm (referencia)', '2,418', '0,750', 'referencia', 'referencia'],
  ['440 y 675 nm', '2,383', '0,727', '0,049', '3 de 83 (3,6 por ciento)'],
  ['443 y 677 nm', '2,373', '0,800', '0,055', '11 de 83 (13,3 por ciento)'],
  ['443 y 676 nm', '2,362', '0,815', '0,075', '11 de 83 (13,3 por ciento)'],
  ['443 y 675 nm', '2,286', '0,790', '0,117', '14 de 83 (16,9 por ciento)']
], { centrar: true }));
cuerpo.push(caption('Tabla 5. Sensibilidad de la clasificación de tamaños frente al par de longitudes de onda utilizado.'));
cuerpo.push(rich([
  { t: 'El par que la metodología del anteproyecto describe, 443 y 675 nm, reclasifica a casi una de cada seis estaciones respecto del par con el que está construida la base entregada. ', b: false },
  { t: 'La consecuencia práctica es que el par debe fijarse por escrito en una sola parte del documento y respetarse en resumen, metodología y resultados.', b: true }
]));
cuerpo.push(p('Se recomienda adoptar el par de 440 y 676 nm por tres razones: es el que está efectivamente calculado en las cinco campañas, es el que corresponde al máximo real de absorción de la clorofila a en la banda roja de estos espectros y es el menos sensible a la posición exacta del máximo azul, según se observa en la tabla anterior.'));

cuerpo.push(h2('4.3 Análisis exploratorio (T3)'));
cuerpo.push(p('Con los 112 registros válidos se obtuvo la descripción que el anteproyecto anunciaba. La tabla siguiente resume el comportamiento por campaña.'));
cuerpo.push(tabla([1900, 700, 1700, 1700, 1700, 1700], [
  ['Campaña', 'n', 'Temperatura media', 'Salinidad media', 'Turbidez mediana', 'Índice mediano'],
  ['Seca 2021', '20', '29,69', '23,99', '5,53', '2,301'],
  ['Seca 2022', '20', '29,64', '25,15', '8,84', '2,723'],
  ['Seca 2023', '29', '31,88', '22,06', '3,96', '2,172'],
  ['Lluviosa 2022', '29', '30,79', '21,19', '5,30', '2,009'],
  ['Lluviosa 2023', '14', '30,64', '24,34', '4,11', '2,850'],
  [{ t: 'Conjunto', b: true }, { t: '112', b: true }, { t: '30,68', b: true }, { t: '23,01', b: true }, { t: '5,30', b: true }, { t: '2,330', b: true }]
], { centrar: true }));
cuerpo.push(caption('Tabla 6. Estadística descriptiva por campaña. La temperatura se expresa en grados Celsius y la turbidez en unidades nefelométricas.'));
cuerpo.push(figura('fig1_eda.png', 600, 246));
cuerpo.push(caption('Figura 1. a) Distribución del índice por campaña con los umbrales de clasificación. b) Índice frente a la salinidad, con la turbidez en escala logarítmica representada por el color.'));
cuerpo.push(p('El resultado más informativo de esta etapa es que la variabilidad entre campañas del mismo periodo climático es del mismo orden que la variabilidad entre periodos. Las dos campañas de sequía de 2022 y 2023 difieren entre sí más de lo que difieren la sequía y la temporada lluviosa consideradas en bloque. Ese hecho, que el anteproyecto no podía anticipar, condiciona toda la lectura posterior.'));
cuerpo.push(p('La turbidez presenta la asimetría más severa del conjunto, con una mediana de 5,30 y un máximo de 134,1 unidades. Por esa razón se trabajó con su logaritmo decimal en las etapas multivariadas.'));

cuerpo.push(h2('4.4 Supuestos y análisis bivariado (T4)'));
cuerpo.push(p('El anteproyecto dejaba abierta la elección entre Pearson y Spearman según la normalidad. La prueba de Shapiro Wilk resuelve la duda de manera concluyente.'));
cuerpo.push(tabla([2400, 1200, 1500, 1600, 3200], [
  ['Variable', 'n', 'Estadístico W', 'Probabilidad', 'Decisión'],
  ['Temperatura', '109', '0,986', '0,298', 'Compatible con la normalidad'],
  ['Salinidad', '111', '0,966', '0,007', 'No normal'],
  ['Turbidez', '111', '0,528', 'menor que 0,001', 'No normal, fuerte asimetría'],
  ['Índice Blue/Red', '112', '0,914', 'menor que 0,001', 'No normal']
], { centrar: true }));
cuerpo.push(caption('Tabla 7. Prueba de normalidad de Shapiro Wilk sobre los registros válidos.'));
cuerpo.push(p('La prueba de Levene sobre la homogeneidad de varianzas entre épocas climáticas resultó no significativa para la salinidad, la turbidez y el índice, y significativa para la temperatura, con una probabilidad menor que 0,001. En consecuencia se adoptaron pruebas no paramétricas para todo el análisis, decisión que debe quedar escrita en la metodología en lugar de la formulación condicional actual.'));
cuerpo.push(p('Los coeficientes de correlación de Spearman entre las variables ambientales y el índice son los siguientes.'));
cuerpo.push(tabla([2500, 1500, 1400, 1400, 3100], [
  ['Relación', 'Conjunto', 'rho', 'Probabilidad', 'Lectura'],
  ['Índice y temperatura', '109', '−0,273', '0,004', 'Significativa y negativa. A mayor temperatura, células mayores'],
  ['Índice y salinidad', '111', '+0,240', '0,011', 'Significativa y positiva. Las aguas marinas favorecen células menores'],
  ['Índice y turbidez', '111', '+0,118', '0,217', 'No significativa'],
  ['Índice y salinidad en época lluviosa', '43', '+0,463', '0,002', 'La relación más fuerte de todo el conjunto'],
  ['Índice y temperatura en época seca', '66', '−0,326', '0,008', 'Se mantiene dentro de la época seca'],
  ['Salinidad y turbidez', '111', '−0,385', 'menor que 0,001', 'Confirma el gradiente estuarino del Canal del Dique']
], { centrar: false }));
cuerpo.push(caption('Tabla 8. Correlaciones de Spearman. Se reporta el signo con el símbolo matemático de resta.'));
cuerpo.push(p('El comportamiento de la salinidad merece atención. Su asociación con el índice es débil en el conjunto completo y desaparece dentro de la época seca, pero se vuelve la relación dominante en la época lluviosa. Es decir, el gradiente de agua dulce actúa sobre la estructura de tamaños sobre todo cuando el Canal del Dique descarga con mayor intensidad, que es exactamente el mecanismo que el anteproyecto propone como hipótesis, aunque no en la forma en que lo enuncia.'));

cuerpo.push(h2('4.5 Análisis multivariado (T5)'));
cuerpo.push(p('El análisis de componentes principales se realizó sobre las cuatro variables estandarizadas, con la turbidez transformada a logaritmo decimal, para las 109 estaciones con información completa.'));
cuerpo.push(tabla([2200, 1600, 1600, 1600, 1600], [
  ['Componente', 'Varianza explicada', 'Acumulada', 'Carga mayor', 'Segunda carga'],
  ['CP1', '40,2 por ciento', '40,2 por ciento', 'Salinidad +0,842', 'Temperatura −0,623'],
  ['CP2', '29,2 por ciento', '69,4 por ciento', 'Índice −0,716', 'Turbidez −0,658'],
  ['CP3', '18,6 por ciento', '87,9 por ciento', 'Temperatura −0,622', 'Índice −0,563'],
  ['CP4', '12,1 por ciento', '100 por ciento', '', '']
], { centrar: true }));
cuerpo.push(caption('Tabla 9. Varianza explicada y cargas principales del análisis de componentes.'));
cuerpo.push(figura('fig3_multivariado.png', 600, 265));
cuerpo.push(caption('Figura 2. a) Plano de los dos primeros componentes principales con las estaciones diferenciadas por época climática. b) Dendrograma de conglomerados por el método de Ward.'));
cuerpo.push(p('El primer componente es un eje estuarino puro: opone salinidad alta a temperatura alta y turbidez alta, y el índice casi no participa en él. El segundo componente sí ordena el índice, y lo hace en oposición a la turbidez. La nube de puntos no se separa por época climática, lo que anticipa el resultado de las pruebas de hipótesis.'));
cuerpo.push(p('El agrupamiento jerárquico por el método de Ward, con una correlación cofenética de 0,514, sugiere tres grupos con significado ecológico claro.'));
cuerpo.push(tabla([1200, 800, 1600, 1400, 1400, 1300, 2100], [
  ['Grupo', 'n', 'Temperatura', 'Salinidad', 'Turbidez', 'Índice', 'Clase dominante'],
  ['1', '36', '29,82', '30,78', '5,49', '2,78', 'Nano con 16 estaciones'],
  ['2', '21', '31,07', '15,69', '34,47', '2,84', 'Micro con 10 estaciones'],
  ['3', '52', '31,11', '20,42', '4,90', '2,13', 'Micro con 44 de 52 estaciones']
], { centrar: true }));
cuerpo.push(caption('Tabla 10. Caracterización de los tres conglomerados obtenidos.'));
cuerpo.push(p('El grupo 1 reúne las aguas marinas menos alteradas, con la salinidad más alta y turbidez baja, y es el único donde el nanofitoplancton domina. El grupo 2 corresponde a la pluma activa del Canal del Dique, con salinidad de 15,7 y turbidez media de 34,5 unidades. El grupo 3, el más numeroso, representa la condición intermedia de la bahía interior, con salinidad moderada, aguas claras y un predominio muy marcado del microfitoplancton, 44 estaciones de 52. Ninguno de los tres grupos es exclusivo de una época climática: los tres contienen estaciones de sequía y de temporada lluviosa.'));

cuerpo.push(h2('4.6 Pruebas de hipótesis (T6)'));
cuerpo.push(p('Esta es la tarea cuyo resultado obliga a revisar la hipótesis del anteproyecto, de manera que se reporta con detalle.'));
cuerpo.push(tabla([3000, 3600, 1600, 2400], [
  ['Contraste', 'Prueba', 'Probabilidad', 'Conclusión'],
  ['Índice entre época seca y época lluviosa', 'Mann Whitney, 69 frente a 43 estaciones', '0,126', 'Sin diferencia significativa. Tamaño del efecto de 0,145'],
  ['Índice entre las cinco campañas', 'Kruskal Wallis con 4 grados de libertad', '0,001', 'Diferencia significativa'],
  ['Época lluviosa 2022 frente a época seca 2022', 'Mann Whitney con corrección de Bonferroni', '0,001', 'Diferencia significativa'],
  ['Época seca 2022 frente a época seca 2023', 'Mann Whitney con corrección de Bonferroni', '0,005', 'Diferencia significativa entre dos campañas de la misma época'],
  ['Clase de tamaño y época climática', 'Chi cuadrado con 2 grados de libertad', '0,806', 'Sin asociación'],
  ['Índice entre zonas de influencia del Canal del Dique', 'Mann Whitney, 65 frente a 46 estaciones', '0,004', 'Diferencia significativa'],
  ['Temperatura, salinidad y turbidez entre épocas', 'Mann Whitney', 'entre 0,193 y 0,948', 'Sin diferencia significativa en ninguna de las tres']
], { centrar: false }));
cuerpo.push(caption('Tabla 11. Resultados de las pruebas de hipótesis sobre los registros válidos.'));
cuerpo.push(figura('fig2_clases.png', 620, 220));
cuerpo.push(caption('Figura 3. a) Composición porcentual de clases de tamaño por campaña. b) Índice según la zona de influencia del Canal del Dique. c) Relación entre índice y temperatura superficial.'));
cuerpo.push(rich([
  { t: 'La hipótesis del anteproyecto no se sostiene con estos datos. ', b: true },
  { t: 'Se afirmaba que en la época seca predominaría el microfitoplancton y que en la época lluviosa se observaría una mayor diversidad de tamaños. Lo observado es lo contrario en magnitud pequeña: el microfitoplancton representa el 56,5 por ciento de las estaciones en la época seca y el 62,8 por ciento en la época lluviosa, sin asociación estadística entre clase y época. La diferencia entre campañas del mismo periodo climático, en cambio, sí es significativa.' }
]));
cuerpo.push(p('El contraste que sí resulta significativo es el espacial. Definiendo la zona de alta influencia del Canal del Dique como aquella con salinidad inferior a 25, la mediana del índice es de 2,149 frente a 2,588 en la zona de baja influencia, con una probabilidad de 0,004. Es decir, las aguas influidas por el canal presentan células de mayor tamaño, resultado coherente con el aporte de nutrientes que favorece a las diatomeas, tal como describe el material teórico entregado por la dirección del proyecto.'));
cuerpo.push(p('Sobre la prueba propuesta en el anteproyecto conviene una precisión metodológica: el análisis de varianza de medidas repetidas no es aplicable. Las estaciones no se repiten de forma completa entre campañas, ya que la campaña de sequía de 2021 tiene 32 estaciones, la de 2022 tiene 23 y la de la época lluviosa de 2023 tiene 15, con conjuntos de códigos que solo coinciden parcialmente. El diseño correcto es el de muestras independientes desbalanceadas, que es el que se aplicó.'));

cuerpo.push(h2('4.7 Modelo descriptivo preliminar (T7)'));
cuerpo.push(p('Se ajustó una regresión múltiple del logaritmo decimal del índice sobre la temperatura, la salinidad y el logaritmo de la turbidez, para las 109 estaciones completas. El resultado es deliberadamente modesto y se reporta como tal.'));
cuerpo.push(tabla([2600, 1600, 1400, 1400, 3000], [
  ['Término', 'Coeficiente', 'Error típico', 'Probabilidad', 'Lectura'],
  ['Constante', '+0,8268', '0,2813', '0,004', ''],
  ['Temperatura', '−0,0176', '0,0086', '0,044', 'Único término significativo'],
  ['Salinidad', '+0,0031', '0,0017', '0,077', 'Marginal'],
  ['Logaritmo de la turbidez', '+0,0313', '0,0203', '0,127', 'No significativo'],
  [{ t: 'Ajuste global', b: true }, { t: 'R cuadrado de 0,096', b: true }, { t: '', b: true }, { t: '', b: true }, { t: 'R cuadrado ajustado de 0,070', b: true }]
], { centrar: false }));
cuerpo.push(caption('Tabla 12. Modelo de regresión múltiple preliminar sobre el logaritmo del índice.'));
cuerpo.push(p('Las tres variables ambientales disponibles explican menos del diez por ciento de la variabilidad del índice. Esto no invalida el objetivo del modelo descriptivo, pero sí obliga a corregir la expectativa del anteproyecto: con temperatura, salinidad y turbidez no se llega a un modelo útil. La partición en 75 y 25 por ciento que propone el texto solo tiene sentido cuando el ajuste sobre el conjunto completo sea razonable, y por eso no se ejecutó todavía. Las variables que faltan son las que el anteproyecto enumera y que no acompañan a la base: los nutrientes y los sólidos suspendidos totales, además de la clorofila a de cuatro de las cinco campañas. La clorofila de la campaña de 2021 sí figura en los libros de espectros y se usó para calcular el coeficiente de absorción específico, que discrimina el efecto de empaquetamiento mejor que el índice por sí solo.'));

cuerpo.push(h2('4.8 Anomalías estandarizadas de la temperatura superficial (T8)'));
cuerpo.push(p('La serie satelital se obtuvo del análisis de temperatura superficial MUR, versión fv04.1, con resolución de 0,01 grados, cerca de un kilómetro, que es la que el anteproyecto especifica. El conjunto de MODIS Aqua que aparece primero en el catálogo figura como obsoleto y se corta en 2019, de modo que no cubre el periodo requerido. Se extrajeron 267 compuestas mensuales entre junio de 2002 y septiembre de 2024 para las ocho posiciones geográficas de la tabla de coordenadas, todas ellas coincidentes con celdas de agua del producto.'));
cuerpo.push(p('Sobre esa serie se aplicaron los tres pasos del método. La climatología mensual sitúa el mínimo en marzo, con 27,27 grados Celsius, y el máximo en septiembre, con 29,78, un recorrido anual de 2,51 grados. La anomalía estandarizada resultante ordena el registro en dos regímenes: predominio de valores negativos entre 2002 y 2010, con los años más fríos en 2002, 2004 y 2008, y predominio de valores positivos desde 2015, con 2023 y 2024 como los más cálidos de toda la serie.'));
cuerpo.push(rich([
  { t: 'El hallazgo con mayor consecuencia para el proyecto es que las campañas no se realizaron en condiciones climáticas medias. ', b: true },
  { t: 'Entre 2021 y 2023 la anomalía estandarizada promedió +0,606, frente a −0,091 en el resto de la serie, con una probabilidad de 0,0001 en la prueba de Mann Whitney. El muestreo cubre una fase cálida de la variabilidad interanual de la bahía, lo que debe tenerse presente al generalizar la estructura de tamaños observada.' }]));
cuerpo.push(tabla([2300, 1700, 1900, 1600, 2200], [
  ['Campaña', 'Mes', 'Temperatura observada', 'Anomalía', 'Anomalía estandarizada'],
  ['Época seca 2021', 'abril de 2021', '28,22', '+0,35', '+0,82'],
  ['Época seca 2022', 'marzo de 2022', '27,62', '+0,35', '+0,61'],
  ['Época lluviosa 2022', 'octubre de 2022', '29,49', '−0,08', '−0,21'],
  ['Campaña de 2023', 'diciembre de 2023', '29,32', '+0,52', '+1,17']
], { centrar: false }));
cuerpo.push(caption('Tabla 13. Anomalía de la temperatura superficial en el mes de cada campaña. El producto no incluye compuesta mensual para junio de 2023.'));
cuerpo.push(p('El resultado enlaza con el análisis bivariado. El índice mantiene una correlación negativa con la temperatura, de modo que una fase cálida favorece el predominio de células de mayor tamaño, que es lo observado en la bahía, con 66 de las 112 estaciones por debajo del umbral del microfitoplancton. La contextualización climática que el anteproyecto proponía como cierre metodológico queda así incorporada al cuerpo de resultados.'));

// ---- 5. mejoras al texto ----
cuerpo.push(new Paragraph({ pageBreakBefore: true, children: [] }));
cuerpo.push(h1('5. Mejoras propuestas al texto del anteproyecto'));
cuerpo.push(p('Las siguientes correcciones se derivan directamente de lo ejecutado y pueden incorporarse al documento principal sin alterar su estructura.'));

cuerpo.push(h3('5.1 Unificar el par de longitudes de onda'));
cuerpo.push(p('El resumen menciona 440 y 675 nm, el abstract repite ese par, la metodología usa 443 y 675 nm y la base de datos entregada está calculada con 440 y 676 nm. Se propone fijar un solo par en los cuatro lugares.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'el índice Blue/Red se calcula como el cociente entre el coeficiente de absorción del fitoplancton a 440 nm y su valor a 676 nm, longitudes que corresponden a los máximos de absorción de la clorofila a en las bandas azul y roja del espectro medido. La clasificación por tamaños sigue los umbrales de Ciotti y Bricaud, con picofitoplancton por encima de 3,0, nanofitoplancton entre 2,5 y 3,0 y microfitoplancton por debajo de 2,5.', i: true }]));

cuerpo.push(h3('5.2 Incorporar un apartado de control de calidad'));
cuerpo.push(p('El anteproyecto pasa directamente de la ecuación de absorción al análisis estadístico. Falta el filtro intermedio, que en estos datos elimina el 14,5 por ciento de los registros.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'antes del análisis estadístico se aplica un control de calidad con cuatro criterios de exclusión. Primero, se descartan las estaciones sin espectro de absorción del fitoplancton. Segundo, se descartan los valores del índice inferiores a 1,0 o superiores a 6,0, por hallarse fuera del intervalo físicamente admisible. Tercero, se descartan las estaciones en las que la absorción del material no algal supera el noventa por ciento de la absorción del particulado total a 440 nm, condición en la que la señal del fitoplancton es un residuo numérico. Cuarto, se verifican los valores de temperatura, salinidad y turbidez frente a los rangos esperados en la bahía y se distingue el cero real del cero usado como dato ausente.', i: true }]));

cuerpo.push(h3('5.3 Corregir el diseño de las pruebas de hipótesis'));
cuerpo.push(p('Debe retirarse la mención al análisis de varianza de medidas repetidas, que el diseño de muestreo no permite.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'dado que las estaciones no se repiten de forma completa entre campañas y que las pruebas de Shapiro Wilk rechazan la normalidad de la salinidad, la turbidez y el índice, los contrastes se realizan con pruebas no paramétricas para muestras independientes. Se emplea la prueba de Mann Whitney para las comparaciones entre dos grupos, la prueba de Kruskal Wallis para la comparación entre las cinco campañas, con comparaciones por pares corregidas por el método de Bonferroni, y la prueba de chi cuadrado para la asociación entre clase de tamaño y época climática.', i: true }]));

cuerpo.push(h3('5.4 Definir operativamente las zonas'));
cuerpo.push(p('El anteproyecto compara zonas de mayor y menor influencia del Canal del Dique sin decir cómo se separan. Mientras no se incorporen las coordenadas de cada estación, que aparecen en la figura 3 del anteproyecto pero no en la base entregada, se propone un criterio reproducible.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'se clasifican como estaciones de alta influencia del Canal del Dique aquellas con salinidad inferior a 25, umbral que separa las aguas de mezcla estuarina de las aguas de carácter marino en la bahía, y como estaciones de baja influencia las restantes. Este criterio se validará posteriormente contra la posición geográfica de cada estación y la distancia a la desembocadura del canal.', i: true }]));

cuerpo.push(h3('5.5 Reformular la hipótesis'));
cuerpo.push(p('La hipótesis actual atribuye el control de la estructura de tamaños a la alternancia entre época seca y época lluviosa, y los datos no lo respaldan. Se propone una formulación que el propio conjunto de datos permite contrastar.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'la estructura de tamaños del fitoplancton en la Bahía de Cartagena responde principalmente al gradiente estuarino generado por los aportes del Canal del Dique y a la variabilidad interanual de la temperatura superficial, y solo de manera secundaria a la alternancia entre época seca y época lluviosa. Se espera, en consecuencia, que las estaciones de baja salinidad y alta turbidez presenten un predominio de microfitoplancton y que las diferencias entre campañas de un mismo periodo climático sean comparables o mayores que las diferencias entre periodos.', i: true }]));

cuerpo.push(h3('5.6 Ajustar la expectativa del modelo descriptivo'));
cuerpo.push(p('Debe indicarse qué variables entran al modelo y qué se hará si el ajuste resulta insuficiente, como ocurre con las tres variables disponibles hoy.'));
cuerpo.push(rich([{ t: 'Redacción sugerida: ', b: true }, { t: 'el modelo descriptivo se construirá con las variables ambientales que resulten significativas en el análisis bivariado y multivariado, incorporando la clorofila a, los nutrientes y los sólidos suspendidos totales además de la temperatura, la salinidad y la turbidez. Si el ajuste sobre el conjunto completo no alcanza un coeficiente de determinación de al menos 0,4, se optará por un modelo de clasificación de la clase de tamaño dominante en lugar de una predicción del valor continuo del índice, y se reportará su exactitud mediante validación cruzada.', i: true }]));

cuerpo.push(h3('5.7 Corregir el rótulo de la base de datos'));
cuerpo.push(p('En el compendio, las columnas rotuladas como cociente entre 443 y 677 nm de la época lluviosa de 2022, la época lluviosa de 2023 y la época seca de 2023 están calculadas con el denominador de 676 nm. El rótulo debe corregirse en el archivo antes de que la base circule, porque un tercero que la reutilice reproducirá el error.'));

// ---- 6. conclusiones ----
cuerpo.push(h1('6. Conclusiones del avance'));
cuerpo.push(bullet('Las ocho tareas asignadas quedaron ejecutadas sobre datos reales. La del modelo descriptivo se reporta con la advertencia de que el ajuste alcanzado es insuficiente con las variables hoy disponibles.'));
cuerpo.push(bullet('El anteproyecto tenía las ocho tareas correctamente descritas y ninguna desarrollada. Su principal debilidad no era el protocolo sino la ausencia de contacto con los datos, que es lo que revela las inconsistencias de longitudes de onda, la falta de criterios de exclusión y el error en el diseño de las pruebas.'));
cuerpo.push(bullet('El control de calidad excluye 19 de los 131 registros, es decir el 14,5 por ciento, casi todos asociados a estaciones de turbidez elevada en las que la absorción del fitoplancton se obtiene por diferencia entre señales muy grandes y muy parecidas.'));
cuerpo.push(bullet('La elección del par de longitudes de onda cambia la clase de tamaño de hasta el 16,9 por ciento de las estaciones, de modo que fijarlo por escrito es una condición previa a cualquier resultado.'));
cuerpo.push(bullet('No hay diferencia significativa del índice entre época seca y época lluviosa, con una probabilidad de 0,126, ni asociación entre clase de tamaño y época, con una probabilidad de 0,806. Sí hay diferencia significativa entre campañas, con una probabilidad de 0,001, incluso entre dos campañas de la misma época seca.'));
cuerpo.push(bullet('Sí hay diferencia significativa entre zonas de influencia del Canal del Dique, con una probabilidad de 0,004 y células mayores en la zona de baja salinidad, lo que respalda el objetivo específico tercero del anteproyecto.'));
cuerpo.push(bullet('El conglomerado más numeroso, con 52 estaciones, corresponde a la bahía interior con salinidad moderada y aguas claras, y concentra el predominio del microfitoplancton con 44 de sus 52 estaciones.'));
cuerpo.push(bullet('El índice queda validado por una vía independiente. Con la clorofila a de la campaña de 2021 se obtuvo el coeficiente de absorción específico del fitoplancton, cuya correlación con el índice es de 0,789 y cuya mediana crece de forma ordenada desde el microfitoplancton hasta el picofitoplancton.'));
cuerpo.push(bullet('El modelo descriptivo con las tres variables disponibles explica menos del diez por ciento de la variabilidad del índice, lo que hace imprescindible incorporar la clorofila a y los nutrientes.'));

// ---- 7. pendientes ----
cuerpo.push(h1('7. Lo que falta y en qué orden'));
cuerpo.push(tabla([700, 4200, 2600, 2200], [
  ['Nº', 'Acción', 'Insumo requerido', 'Responsable'],
  ['1', 'Solicitar al CIOH los nutrientes, los sólidos suspendidos totales y la clorofila a de las cuatro campañas que no la traen', 'Gestión institucional', 'Estudiante y dirección'],
  ['2', 'Solicitar las coordenadas de cada estación para reemplazar el criterio de salinidad por la distancia real a la desembocadura', 'Base geográfica', 'Estudiante'],
  ['3', 'Solicitar el libro de espectros de la campaña de sequía de 2023, ausente en la entrega', 'Archivo del laboratorio', 'Estudiante'],
  ['4', 'Corregir en el archivo del compendio el rótulo de las columnas del segundo índice', 'Archivo entregado', 'Estudiante'],
  ['5', 'Reajustar el modelo descriptivo con el conjunto ampliado de variables y ejecutar la partición de entrenamiento y validación', 'Resultado de la acción 1', 'Estudiante'],
  ['6', 'Incorporar al anteproyecto las siete mejoras de redacción del capítulo 5', 'Este informe', 'Estudiante']
], { centrar: false }));
cuerpo.push(caption('Tabla 13. Acciones pendientes en orden de prioridad.'));

// ---- anexos ----
cuerpo.push(new Paragraph({ pageBreakBefore: true, children: [] }));
cuerpo.push(h1('Anexo A. Estaciones excluidas por el control de calidad'));
cuerpo.push(tabla([1900, 1700, 1500, 1500, 1500, 2400], [
  ['Estación', 'Campaña', 'Índice', 'Salinidad', 'Turbidez', 'Motivo'],
  ['SC 21 0113', 'Seca 2021', '9,52', '0,30', '242,0', 'Índice y temperatura imposibles'],
  ['SC 21 0163', 'Seca 2021', '0,67', '18,90', '2,56', 'Índice por debajo de 1,0'],
  ['SC 21 0165', 'Seca 2021', '0,72', '21,10', '4,09', 'Índice por debajo de 1,0'],
  ['SC 21 0231', 'Seca 2021', '8,35', '11,50', '31,6', 'Índice por encima de 6,0'],
  ['SC 22 0121', 'Seca 2022', '10,20', '25,65', '183,5', 'Índice por encima de 6,0'],
  ['SC 22 0122', 'Seca 2022', '8,37', '21,18', '63,7', 'Índice por encima de 6,0'],
  ['SC 22 0127', 'Seca 2022', '6,54', '31,26', '19,9', 'Índice por encima de 6,0'],
  ['SC 23 0522', 'Lluviosa 2023', '6,43', '0,00', '15,4', 'Índice por encima de 6,0'],
  ['SC 22 0444', 'Lluviosa 2022', '4,04', '16,00', '24,4', 'Absorción no algal del 98 por ciento del total']
], { centrar: true }));
cuerpo.push(caption('Tabla 14. Estaciones excluidas por valor del índice o por calidad del espectro. Las ocho primeras se excluyeron del análisis; la última se señala para descartar en la próxima depuración. A todas ellas se suman las once estaciones sin índice calculado, ocho de la campaña de sequía de 2021 y tres de la campaña de sequía de 2023.'));

cuerpo.push(h1('Anexo B. Procedimiento y trazabilidad'));
cuerpo.push(p('Todo el procesamiento se realizó con rutinas escritas para este avance, de modo que cualquier resultado puede reproducirse desde los archivos originales. El anteproyecto propone RStudio como entorno de trabajo. Las rutinas de este avance están escritas en Python, opción que el propio anteproyecto contempla en su apartado de riesgos, y sus resultados son idénticos a los que producirían las funciones equivalentes de R.'));
cuerpo.push(tabla([3000, 6200], [
  ['Rutina', 'Función'],
  ['Construcción de la base', 'Une las cinco campañas del compendio, normaliza los códigos de estación y adosa los valores espectrales de aphy, ap y ad de los cuatro libros disponibles'],
  ['Verificación del índice', 'Reproduce el índice desde los espectros, identifica el par de longitudes de onda realmente usado y cuantifica la sensibilidad de la clasificación'],
  ['Análisis estadístico', 'Descriptivos, histogramas por época y por zona, Shapiro Wilk, Levene, Spearman, Pearson, Mann Whitney, Kruskal Wallis con Bonferroni, chi cuadrado, componentes principales, conglomerados de Ward y regresión múltiple'],
  ['Figuras', 'Genera las figuras del informe y del documento integrado'],
  ['Anomalías de temperatura', 'Calcula la climatología mensual, la anomalía y la anomalía estandarizada de la temperatura superficial, y produce la figura de la serie']
], { centrar: false }));
cuerpo.push(caption('Tabla 15. Rutinas desarrolladas para este avance.'));
cuerpo.push(p('Nota sobre notación. Los códigos de estación se escriben en este documento separando sus tres bloques con espacios. En los archivos originales aparecen con un separador entre bloques, y en esa forma deben buscarse dentro del compendio. Los valores negativos se indican con el símbolo matemático de resta.'));

// ============================ DOCUMENTO ============================
const doc = new Document({
  creator: 'Omar Enrique Gamboa Posada',
  title: 'Informe de avance sobre las tareas asignadas por la dirección del proyecto',
  description: 'Ejecución de las tareas estadísticas del anteproyecto sobre la base de datos de la Bahía de Cartagena',
  numbering: {
    config: [{
      reference: 'vinetas',
      levels: [
        { level: 0, format: 'bullet', text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.2) } } } },
        { level: 1, format: 'bullet', text: '◦', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: convertInchesToTwip(0.65), hanging: convertInchesToTwip(0.2) } } } }
      ]
    }]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1200, bottom: 1200, left: 1400, right: 1200 } } },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBBBBB' } },
        children: [new TextRun({ text: 'Informe de avance. Índice Blue/Red en la Bahía de Cartagena', font: FONT, size: 16, color: '666666' })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: '666666' })]
      })] })
    },
    children: cuerpo
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('Informe_avance_tareas_tutora.docx', b);
  console.log('documento generado', b.length, 'bytes');
});
