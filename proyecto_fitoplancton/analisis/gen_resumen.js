const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
        Table, TableRow, TableCell, WidthType, ShadingType, convertInchesToTwip,
        Header, Footer, PageNumber } = require('docx');

const F = 'Times New Roman';
const AZUL = '1F3864';

const p = (t, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: { after: o.after === undefined ? 160 : o.after, line: o.line || 300 },
  children: [new TextRun({ text: t, font: F, size: o.size || 24, bold: !!o.bold,
                           italics: !!o.i, color: o.color || '000000' })]
});
const h = (t) => new Paragraph({
  spacing: { before: 320, after: 140 },
  children: [new TextRun({ text: t, font: F, size: 28, bold: true, color: AZUL })]
});
const li = (t, marca = '•') => new Paragraph({
  spacing: { after: 110, line: 300 },
  indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.2) },
  children: [new TextRun({ text: `${marca}  ${t}`, font: F, size: 24 })]
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
      w: anchos[j], b: i === 0, fill: i === 0 ? AZUL : null,
      color: i === 0 ? 'FFFFFF' : '000000'
    }))
  }))
});

const c = [];

c.push(p('EN PALABRAS SENCILLAS', { size: 30, bold: true, align: AlignmentType.CENTER, color: AZUL, after: 80 }));
c.push(p('Qué se hizo con el proyecto y qué falta', { size: 26, bold: true, align: AlignmentType.CENTER, after: 200 }));
c.push(p('Índice Blue/Red en la Bahía de Cartagena. Omar Enrique Gamboa Posada. Agosto de 2026',
         { size: 21, align: AlignmentType.CENTER, i: true, after: 240 }));
c.push(new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 6, color: AZUL } },
                       spacing: { after: 240 }, children: [] }));

c.push(h('1. Qué pidió la profesora'));
c.push(p('Ella marcó en tu anteproyecto la parte que habla del tratamiento estadístico y pidió que dejara de ser una promesa y pasara a ser trabajo hecho. En resumen, pedía ocho cosas:'));
c.push(li('Ordenar y limpiar la base de datos.'));
c.push(li('Calcular bien el índice Blue/Red y comprobarlo.'));
c.push(li('Mirar cómo se distribuyen los datos, con gráficas.'));
c.push(li('Ver qué variables del agua se relacionan con el índice.'));
c.push(li('Hacer los análisis que agrupan estaciones parecidas.'));
c.push(li('Probar si hay diferencias reales entre épocas y entre zonas.'));
c.push(li('Armar un modelo que explique el tamaño del fitoplancton.'));
c.push(li('Calcular las anomalías de temperatura del mar por satélite.'));

c.push(h('2. Cómo estaba el documento antes'));
c.push(p('Las ocho cosas estaban muy bien explicadas, pero ninguna estaba hecha. Todo el capítulo estaba escrito en futuro: se hará, se aplicará, se calculará. No había ni un solo número calculado, ni una tabla de resultados, ni una gráfica con tus datos.'));
c.push(p('Eso no es un defecto grave en un anteproyecto. Lo que sí trajo problemas es que, al no haber tocado nunca los datos, el texto tenía tres cosas mal que solo se ven cuando uno se sienta a trabajar con ellos.'));

c.push(h('3. Qué se hizo'));
c.push(p('Las ocho tareas están ejecutadas con tus datos reales y escritas dentro del anteproyecto. Este es el resumen:'));
c.push(tabla([2200, 6900], [
  ['Tarea', 'Qué se hizo'],
  ['Ordenar la base', 'Se juntaron las cinco campañas en una sola tabla de 131 estaciones y se revisó una por una'],
  ['Revisar el índice', 'Se volvió a calcular desde los espectros originales y dio exactamente igual, lo que confirma que la base está bien construida'],
  ['Mirar los datos', 'Cajas, dispersión e histogramas, separados por época y por zona'],
  ['Relacionar variables', 'Correlaciones con temperatura, salinidad y turbidez, con su prueba de significancia'],
  ['Agrupar estaciones', 'Componentes principales y conglomerados, que separaron la bahía en tres ambientes'],
  ['Probar diferencias', 'Comparaciones entre épocas, entre campañas y entre zonas'],
  ['Modelo', 'Se ajustó y se reportó con honestidad que todavía no sirve, y se explica por qué'],
  ['Temperatura satelital', 'Se procesaron 267 meses de imágenes entre 2002 y 2024 y se calcularon las anomalías']
]));

c.push(h('4. Los cinco hallazgos que importan'));
c.push(p('Estos son los resultados que cambian algo. Vale la pena que los tengas claros para la reunión.'));
c.push(p('Primero. Había un lío con las longitudes de onda.', { bold: true, after: 60 }));
c.push(p('El resumen de tu documento decía 440 y 675 nanómetros, la metodología decía 443 y 675, y la base de datos estaba hecha con 440 y 676. Comprobé que esa elección no es un detalle: cambia la clasificación de tamaño de casi una de cada seis estaciones. Ahora todo el documento usa un solo par y lo justifica.'));
c.push(p('Segundo. Hay datos que no se pueden usar.', { bold: true, after: 60 }));
c.push(p('De las 131 estaciones, 19 quedaron fuera. Casi todas son sitios de agua muy turbia, donde la señal del fitoplancton se calcula restando dos números grandes y muy parecidos, y el resultado se vuelve basura. Hay una estación con un índice de 9,5 y otra con 10,2, valores imposibles. El documento ahora explica con qué criterios se descartan.'));
c.push(p('Tercero. La hipótesis del anteproyecto no se sostiene.', { bold: true, after: 60 }));
c.push(p('Decía que en la época seca predomina el fitoplancton grande y que en la lluviosa habría más variedad. Los datos dicen que entre época seca y época lluviosa no hay diferencia, y que la relación entre tamaño y época tampoco existe. Lo que sí hay son diferencias entre campañas, incluso entre dos campañas de la misma época seca. La hipótesis quedó reescrita.'));
c.push(p('Cuarto. Lo que sí manda es el Canal del Dique.', { bold: true, after: 60 }));
c.push(p('Donde el agua es menos salada, por la descarga del canal, las células son más grandes, y la diferencia es estadísticamente sólida. Ese resultado respalda directamente tu tercer objetivo específico.'));
c.push(p('Quinto. Muestreaste en años calientes.', { bold: true, after: 60 }));
c.push(p('Al procesar veintidós años de temperatura satelital se ve que entre 2021 y 2023 la bahía estuvo claramente más caliente que su promedio histórico, y que 2023 y 2024 son los dos años más calientes de todo el registro. Es decir, tus cinco campañas no retratan una bahía en condiciones normales, sino en una fase cálida. Eso hay que decirlo al momento de generalizar.'));

c.push(h('5. Qué falta'));
c.push(p('Tres cosas. Ninguna depende de más análisis, todas dependen de conseguir información.'));
c.push(p('Falta 1. Las variables que nunca llegaron.', { bold: true, after: 60 }));
c.push(p('Tu metodología nombra siete variables del agua: temperatura, salinidad, turbidez, sólidos suspendidos, nitratos, fosfatos y silicatos. En la base que te entregaron solo venían las tres primeras. Todo se analizó con esas tres, pero con tres variables el modelo explica menos del diez por ciento de lo que pasa. Hay que pedirle al CIOH la clorofila a y los nutrientes. Esto es lo más importante de la lista.'));
c.push(p('Falta 2. El modelo con validación.', { bold: true, after: 60 }));
c.push(p('La profesora pedía partir los datos en dos grupos, uno para entrenar y otro para validar. No lo hice, y fue una decisión consciente: con un ajuste tan bajo, esa partición produce números que no significan nada y que ella tendría razón en cuestionar. En el documento quedó escrito qué hacer cuando lleguen las variables que faltan. Pregúntale si aun así quiere que se haga.'));
c.push(p('Falta 3. Dos verificaciones tuyas.', { bold: true, after: 60 }));
c.push(li('Cuál de las dos campañas de 2023 es la de junio y cuál la de diciembre. La numeración de las muestras contradice las etiquetas de la base, y si están cambiadas, los resultados por época cambian.'));
c.push(li('Las coordenadas de cada estación. Mientras no estén, la separación entre zona influida y no influida por el canal se hace con la salinidad, que funciona bien pero es un sustituto.'));

c.push(h('6. Qué hacer ahora, en orden'));
c.push(tabla([700, 5200, 3200], [
  ['Nº', 'Acción', 'Quién'],
  ['1', 'Pedir al CIOH clorofila a, nutrientes y sólidos suspendidos', 'Tú, con apoyo de la profesora'],
  ['2', 'Confirmar la fecha real de cada campaña de 2023', 'Tú, con el registro de campo'],
  ['3', 'Pedir las coordenadas de las estaciones', 'Tú'],
  ['4', 'Pedir el libro de espectros de una campaña de 2023 que no llegó', 'Tú'],
  ['5', 'Rehacer el modelo con todas las variables', 'Cuando llegue lo del punto 1'],
  ['6', 'Corregir el rótulo equivocado en el archivo de la base', 'Tú, es un minuto']
]));

c.push(h('7. Qué archivos tienes'));
c.push(li('El anteproyecto con todo integrado y con normas APA aplicadas. Es el que le entregas a la profesora.'));
c.push(li('Un informe técnico aparte, más detallado, por si ella quiere ver el detalle de cada prueba estadística.'));
c.push(li('Este resumen, que es el que te sirve a ti para tener claro dónde estás parado.'));
c.push(p('Al abrir el anteproyecto en Word, si te pregunta si quieres actualizar los campos, dile que sí. Eso arma solo la tabla de contenido y las listas de figuras y tablas.',
         { i: true }));

const doc = new Document({
  creator: 'Omar Enrique Gamboa Posada',
  title: 'En palabras sencillas. Qué se hizo con el proyecto y qué falta',
  description: 'Resumen en lenguaje llano del estado del proyecto de grado',
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
                          margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ['Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES],
                               font: F, size: 18, color: '666666' })] })] }) },
    children: c
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('Resumen_sencillo_estado_del_proyecto.docx', b);
  console.log('resumen generado', b.length, 'bytes');
});
