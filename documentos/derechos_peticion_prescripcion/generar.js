const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, Header, Footer,
  PageNumber, PageOrientation, VerticalAlign,
} = require("docx");
const { PETICIONARIO: P, DESTINOS } = require("./datos.js");

const FUENTE = "Arial";
const T = 22;          // 11 pt
const CONTENT_W = 9360; // 12240 - 1440*2

const ORDINALES = ["PRIMERO", "SEGUNDO", "TERCERO", "CUARTO", "QUINTO", "SEXTO"];

// ---------- helpers ----------
const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || T, bold: !!o.bold, italics: !!o.italics, underline: o.underline ? {} : undefined, color: o.color });

const par = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after === undefined ? 160 : o.after, before: o.before || 0, line: o.line || 276 },
  indent: o.indent,
  border: o.border,
});

const txt = (text, o = {}) => par(run(text, o), o);
const vacio = (after = 0) => new Paragraph({ children: [], spacing: { after } });

// Título de sección centrado, negrita, mayúsculas
const seccion = (t) => par(run(t, { bold: true }), { align: AlignmentType.CENTER, before: 260, after: 180 });

// Cita normativa: sangrada e itálica
const cita = (t) => par(run(t, { italics: true, size: 21 }), {
  indent: { left: 720, right: 400 }, after: 130, line: 260,
});

// ---------- membrete (encabezado) ----------
const membrete = () => new Header({
  children: [
    par(run(P.nombre, { bold: true, size: 30 }), { align: AlignmentType.CENTER, after: 0, line: 240 }),
    par(run("A B O G A D O", { size: 19 }), { align: AlignmentType.CENTER, after: 40, line: 240 }),
    par(run(`T.P. No. ${P.tp} del C. S. de la J.   |   Cel. ${P.cel}   |   ${P.email}`, { size: 16 }), {
      align: AlignmentType.CENTER, after: 200, line: 240,
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, space: 6, color: "1F3864" } },
    }),
  ],
});

const pie = () => new Footer({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120 },
      children: [new TextRun({ children: ["Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES], font: FUENTE, size: 16, color: "666666" })],
    }),
  ],
});

// ---------- tabla de comparendos ----------
const celda = (texto, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.head ? { type: ShadingType.CLEAR, fill: "1F3864", color: "auto" } : undefined,
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 60, bottom: 60, left: 90, right: 90 },
  children: [par(run(texto, { bold: !!o.head, size: 17, color: o.head ? "FFFFFF" : undefined }), {
    align: o.align || AlignmentType.LEFT, after: 0, line: 240,
  })],
});

const COLS = [2280, 1240, 700, 3140, 2000];

const tablaComparendos = (comparendos) => new Table({
  columnWidths: COLS,
  width: { size: CONTENT_W, type: WidthType.DXA },
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        celda("No. de comparendo", { w: COLS[0], head: true, align: AlignmentType.CENTER }),
        celda("Fecha del hecho", { w: COLS[1], head: true, align: AlignmentType.CENTER }),
        celda("Cód.", { w: COLS[2], head: true, align: AlignmentType.CENTER }),
        celda("Lugar de los hechos", { w: COLS[3], head: true, align: AlignmentType.CENTER }),
        celda("Resolución sancionatoria", { w: COLS[4], head: true, align: AlignmentType.CENTER }),
      ],
    }),
    ...comparendos.map((c) => new TableRow({
      children: [
        celda(c.numero, { w: COLS[0] }),
        celda(c.fechaCorta, { w: COLS[1], align: AlignmentType.CENTER }),
        celda(c.codigo, { w: COLS[2], align: AlignmentType.CENTER }),
        celda(c.lugar, { w: COLS[3] }),
        celda(`No. ${c.resolucion} del ${c.fechaResolucion}`, { w: COLS[4] }),
      ],
    })),
  ],
});

// ---------- cuerpo ----------
function construirCuerpo(d) {
  const n = d.comparendos.length;
  const plural = n > 1;
  const hijos = [];

  // Fecha
  hijos.push(txt(`${P.ciudadFirma}, ${P.fechaLarga}`, { align: AlignmentType.RIGHT, after: 320 }));

  // Destinatario
  hijos.push(txt("Señores", { bold: true, after: 0 }));
  hijos.push(txt(d.entidad, { bold: true, after: 0 }));
  hijos.push(txt(d.dependencia, { after: 0, size: 21 }));
  hijos.push(txt(d.ciudad, { after: 320 }));

  // Referencia
  hijos.push(par([
    run("REF: ", { bold: true }),
    run("DERECHO DE PETICIÓN – SOLICITUD DE DECLARATORIA OFICIOSA DE PRESCRIPCIÓN DE ", { bold: true }),
    run(plural ? "LAS SANCIONES" : "LA SANCIÓN", { bold: true }),
    run(" POR ", { bold: true }),
    run(plural ? "INFRACCIONES" : "INFRACCIÓN", { bold: true }),
    run(" DE TRÁNSITO (Art. 23 C.P.; Art. 159 de la Ley 769 de 2002, modificado por el Art. 206 del Decreto Ley 019 de 2012).", { bold: true }),
  ], { after: 300 }));

  hijos.push(txt("Respetados señores:", { after: 220 }));

  // Encabezamiento / legitimación — a título personal
  hijos.push(par([
    run(P.nombre, { bold: true }),
    run(`, mayor de edad, identificado con Cédula de Ciudadanía No. ${P.cc} expedida en ${P.ccExp}, `),
    run("actuando en nombre propio y a título estrictamente personal", { bold: true }),
    run(" —esto es, en mi condición de ciudadano y presunto infractor registrado, y no en calidad de apoderado judicial ni extrajudicial de tercero alguno—, en ejercicio del "),
    run("DERECHO CONSTITUCIONAL FUNDAMENTAL DE PETICIÓN", { bold: true }),
    run(" consagrado en el artículo 23 de la Constitución Política y desarrollado en los artículos 13 y siguientes de la Ley 1437 de 2011 (CPACA), sustituidos por la Ley 1755 de 2015, así como en el Decreto 2150 de 1995 y demás disposiciones concordantes y pertinentes, me dirijo respetuosamente a esa Entidad para los efectos del inciso 2.º del artículo 206 del Decreto Ley 019 de 2012, que modificó el artículo 159 de la Ley 769 de 2002 “Código Nacional de Tránsito”, con el fin de formular la siguiente:"),
  ], { after: 240 }));

  hijos.push(seccion("PETICIÓN"));

  // Peticiones, una por comparendo
  d.comparendos.forEach((c, i) => {
    const partes = [
      run(`${ORDINALES[i]}. `, { bold: true }),
      run("OFICIOSAMENTE DECLARE LA PRESCRIPCIÓN", { bold: true }),
      run(" de la sanción que me fuera impuesta con ocasión de infracción de tránsito, según Orden de Comparendo No. "),
      run(c.numero, { bold: true }),
      run(` de fecha ${c.fechaLarga}, siendo las ${c.hora}, en el lugar ${c.lugar}, por el código de infracción ${c.codigo} («${c.infraccion.replace(/\.$/, "")}»)`),
    ];
    if (c.placa) partes.push(run(`, respecto del vehículo de placa ${c.placa}`));
    partes.push(run(`, sancionada mediante Resolución No. ${c.resolucion} del ${c.fechaResolucion}, `));
    partes.push(run("TODA VEZ QUE", { bold: true }));
    partes.push(run(" a la fecha han transcurrido "));
    partes.push(run("más de tres (3) años contados a partir de la ocurrencia del hecho", { bold: true }));
    partes.push(run(", sin que se me haya notificado legalmente mandamiento de pago alguno que hubiera interrumpido el término de prescripción. En consecuencia, solicito que se ordene el archivo definitivo del expediente y se actualicen las bases de datos correspondientes del "));
    partes.push(run("SIMIT", { bold: true }));
    partes.push(run(", el "));
    partes.push(run("RUNT", { bold: true }));
    partes.push(run(" y todas aquellas en las que figure como deudor de esta sanción."));
    hijos.push(par(partes, { indent: { left: 340 }, after: 200 }));
  });

  // Tabla resumen
  hijos.push(txt(plural
    ? "Para mayor claridad de esa Entidad, se relacionan los comparendos objeto de la presente petición:"
    : "Para mayor claridad de esa Entidad, se relaciona el comparendo objeto de la presente petición:",
    { before: 120, after: 160 }));
  hijos.push(tablaComparendos(d.comparendos));
  hijos.push(vacio(200));

  // Sustento: no interrupción
  hijos.push(seccion("SOBRE LA NO INTERRUPCIÓN DEL TÉRMINO DE PRESCRIPCIÓN"));

  hijos.push(par([
    run("Lo anterior encuentra sustento en que se cumplen a cabalidad los requisitos legales aplicables, en virtud del inciso segundo del artículo 159 del Código Nacional de Tránsito, a saber:"),
  ], { after: 160 }));

  hijos.push(cita("“La ejecución de las sanciones que se impongan por violación de las normas de tránsito, estará a cargo de las autoridades de tránsito de la jurisdicción donde se cometió el hecho, quienes estarán investidas de jurisdicción coactiva para el cobro, cuando ello fuere necesario."));
  hijos.push(cita("Las sanciones impuestas por infracciones a las normas de tránsito prescribirán en tres (3) años contados a partir de la ocurrencia del hecho; la prescripción deberá ser declarada de oficio y se interrumpirá con la notificación del mandamiento de pago. La autoridad de tránsito no podrá iniciar el cobro coactivo de sanciones respecto de las cuales se encuentren configurados los supuestos necesarios para declarar su prescripción”. (Subrayado y negrilla fuera del texto original)"));

  // Párrafo clave: coactivo no interrumpe
  const conCoactivo = d.comparendos.filter((c) => c.coactivo);
  if (conCoactivo.length) {
    const lista = conCoactivo.map((c) => `No. ${c.coactivo} del ${c.fechaCoactivo} (comparendo ${c.numero})`).join("; ");
    hijos.push(par([
      run("Advierto desde ya que, según la información publicada por esa misma Entidad en sus canales de consulta, "),
      run(conCoactivo.length > 1 ? "figuran resoluciones de cobro coactivo " : "figura una resolución de cobro coactivo "),
      run(lista + ". "),
      run("Tal circunstancia no altera en modo alguno la configuración de la prescripción", { bold: true }),
      run(", por cuanto la norma transcrita es inequívoca al señalar que el término "),
      run("únicamente se interrumpe con la NOTIFICACIÓN del mandamiento de pago", { bold: true, underline: true }),
      run(", y no con la simple expedición de actos administrativos de trámite, ni con el reparto interno del expediente al área de cobro coactivo, ni con el registro del deudor en las bases de datos del SIMIT o del RUNT. La interrupción es un acto de comunicación al administrado, no un acto interno de la administración."),
    ], { before: 140, after: 180 }));
  }

  hijos.push(par([
    run("En el presente caso "),
    run("jamás he sido notificado —ni personal, ni por aviso, ni por ningún otro medio legalmente previsto— de mandamiento de pago alguno", { bold: true }),
    run(` respecto ${plural ? "de los comparendos relacionados" : "del comparendo relacionado"}, de modo que el término de tres (3) años corrió de manera ininterrumpida desde la fecha de ocurrencia de los hechos y se encuentra holgadamente vencido. Al haberse consolidado la prescripción, esa Entidad tiene el `),
    run("deber imperativo y oficioso", { bold: true }),
    run(" de declararla, sin que le sea dado supeditarla a solicitud de parte, y le está expresamente prohibido iniciar o continuar el cobro coactivo."),
  ], { after: 200 }));

  // Fundamentos de derecho
  hijos.push(seccion("FUNDAMENTOS DE DERECHO DE LA PETICIÓN"));

  hijos.push(txt("Artículo 23 de la Constitución Política de Colombia:", { bold: true, after: 120 }));
  hijos.push(cita("“Toda persona tiene derecho a presentar peticiones respetuosas a las autoridades por motivos de interés general o particular y a obtener pronta resolución. El legislador podrá reglamentar su ejercicio ante organizaciones privadas para garantizar los derechos fundamentales”."));

  hijos.push(txt("Artículo 52 de la Ley 1437 de 2011:", { bold: true, before: 120, after: 120 }));
  hijos.push(cita("“Caducidad de la facultad sancionatoria. Salvo lo dispuesto en leyes especiales, la facultad que tienen las autoridades para imponer sanciones caduca a los tres (3) años de ocurrido el hecho, la conducta u omisión que pudiere ocasionarlas, término dentro del cual el acto administrativo que impone la sanción debe haber sido expedido y notificado”. (Subrayado y negrilla fuera del texto original)"));

  hijos.push(txt("Artículo 206 del Decreto Ley 019 de 2012:", { bold: true, before: 120, after: 120 }));
  hijos.push(cita("“La ejecución de las sanciones que se impongan por violación de las normas de tránsito, estará a cargo de las autoridades de tránsito de la jurisdicción donde se cometió el hecho, quienes estarán investidas de jurisdicción coactiva para el cobro, cuando ello fuere necesario."));
  hijos.push(cita("Las sanciones impuestas por infracciones a las normas de tránsito prescribirán en tres (3) años contados a partir de la ocurrencia del hecho; la prescripción deberá ser declarada de oficio y se interrumpirá con la notificación del mandamiento de pago. La autoridad de tránsito no podrá iniciar el cobro coactivo de sanciones respecto de las cuales se encuentren configurados los supuestos necesarios para declarar su prescripción."));
  hijos.push(cita("Las autoridades de tránsito deberán establecer públicamente a más tardar en el mes de enero de cada año, planes y programas destinados al cobro de dichas sanciones y dentro de este mismo periodo rendirán cuentas públicas sobre la ejecución de los mismos”."));

  hijos.push(txt("Artículo 14 de la Ley 1437 de 2011 (sustituido por la Ley 1755 de 2015):", { bold: true, before: 120, after: 120 }));
  hijos.push(cita("“Salvo norma legal especial y so pena de sanción disciplinaria, toda petición deberá resolverse dentro de los quince (15) días siguientes a su recepción”. (Subrayado y negrilla fuera del texto original)"));

  hijos.push(txt("SENTENCIA DE TUTELA 03248 DEL 11 DE FEBRERO DE 2016 DEL CONSEJO DE ESTADO, SECCIÓN PRIMERA, que en uno de sus apartes estableció:", { bold: true, before: 160, after: 120 }));
  hijos.push(cita("“El cobro de las multas de tránsito corresponde, de conformidad con el artículo 159 de la Ley 769 de 2002, modificado por el artículo 26 de la Ley 1383 de 2010 y por el artículo 206 del Decreto Ley 019 de 2012, ‘estará a cargo de las autoridades de tránsito de la jurisdicción donde se cometió el hecho, quienes estarán investidas de jurisdicción coactiva para el cobro, cuando ello fuere necesario’."));
  hijos.push(cita("Por su parte, en relación con la prescripción de las sanciones impuestas por infracciones de las normas de tránsito, según la norma referida, éstas lo harán ‘en tres (3) años contados a partir de la ocurrencia del hecho; la prescripción deberá ser declarada de oficio y se interrumpirá con la notificación del mandamiento de pago. La autoridad de tránsito no podrá iniciar el cobro coactivo de sanciones respecto de las cuales se encuentren configurados los supuestos necesarios para declarar su prescripción’”."));

  hijos.push(txt("SENTENCIA DEL 16 DE OCTUBRE DE 2015 DEL TRIBUNAL ADMINISTRATIVO DE SANTANDER, que en uno de sus apartes estableció:", { bold: true, before: 160, after: 120 }));
  hijos.push(cita("“Para el caso que nos ocupa, es preciso traer a colación el art. 159 de la Ley 769 de 2002 modificado por el artículo 206 del Decreto 19 del 10 de enero de 2012, por cuanto, pese a que no fue señalado expresamente por el accionante como norma incumplida, se advierte de su lectura que este se encuentra directamente relacionado con el artículo 818 del Estatuto Tributario, pues faculta a la autoridad de tránsito de la jurisdicción correspondiente para exigir el cobro producto de sanción a través del proceso coactivo y si esto no se hace dentro del término de tres (3) años siguientes a la ocurrencia del hecho, se configurará la prescripción de la acción de cobro”."));
  hijos.push(cita("“(…) Se observa entonces de esta norma, un deber imperativo en cabeza de la autoridad de tránsito, según el cual debe declarar de oficio la prescripción de los comparendos por infracción a las normas en los cuales haya transcurrido un término mayor a 3 años”."));

  // Petición subsidiaria
  hijos.push(seccion("PETICIÓN SUBSIDIARIA"));
  hijos.push(par([
    run("En el evento —que desde ya controvierto— de que esa Entidad considere que el término de prescripción fue interrumpido, solicito respetuosamente que, junto con la respuesta de fondo, se me expida a mi costa "),
    run("copia íntegra y legible", { bold: true }),
    run(" de los siguientes documentos:"),
  ], { after: 160 }));

  const subs = [
    "(i) El o los mandamientos de pago proferidos respecto " + (plural ? "de los comparendos relacionados" : "del comparendo relacionado") + ", con su respectivo número y fecha.",
    "(ii) La constancia de notificación de dichos mandamientos de pago (acta de notificación personal, citación, aviso, guía y certificación de entrega de la empresa de correo), con indicación expresa de la fecha en la que la notificación se surtió y quedó ejecutoriada.",
    "(iii) La constancia de notificación de las resoluciones sancionatorias señaladas en la tabla precedente, así como de la audiencia pública de que trata el artículo 136 de la Ley 769 de 2002.",
    "(iv) El estado actual del expediente administrativo y de cobro coactivo, con el detalle de las actuaciones surtidas y sus fechas.",
  ];
  subs.forEach((s) => hijos.push(par(run(s), { indent: { left: 340 }, after: 120 })));

  hijos.push(par([
    run("Y que, en tal evento, se declare igualmente la "),
    run("prescripción de la acción de cobro", { bold: true }),
    run(" de conformidad con el artículo 817 del Estatuto Tributario, aplicable por remisión del artículo 5.º de la Ley 1066 de 2006, habida cuenta de que también se encuentra vencido el término de cinco (5) años allí previsto, contado desde la ejecutoria de los actos administrativos que sirven de título ejecutivo."),
  ], { before: 140, after: 200 }));

  // Cierre
  hijos.push(par([
    run("De conformidad con todo lo esbozado anteriormente, "),
    run(plural
      ? "SOLICITO DE MANERA MUY RESPETUOSA QUE SE DECLARE DE OFICIO LA PRESCRIPCIÓN DE LOS COMPARENDOS INICIALMENTE RESEÑADOS"
      : "SOLICITO DE MANERA MUY RESPETUOSA QUE SE DECLARE DE OFICIO LA PRESCRIPCIÓN DEL COMPARENDO INICIALMENTE RESEÑADO",
      { bold: true }),
    run(", se ordene el archivo del expediente, se elimine el registro de la deuda del SIMIT, del RUNT y de cualquier otra base de datos, y de no accederse a lo pedido, me sean informadas debidamente las razones de hecho y de derecho en que se funde la negativa, junto con los recursos que proceden contra dicha decisión."),
  ], { before: 140, after: 220 }));

  // Notificación
  hijos.push(seccion("NOTIFICACIÓN"));
  hijos.push(par([
    run("Recibiré notificaciones y respuestas en el correo electrónico "),
    run(P.email, { bold: true }),
    run(`, medio que autorizo expresamente para efectos de la notificación electrónica prevista en el artículo 56 de la Ley 1437 de 2011, y/o en la dirección ${P.direccion} de ${P.ciudadFirma}, Santander. Celular: ${P.cel}.`),
  ], { after: 200 }));

  hijos.push(txt("Agradezco de antemano la atención prestada y la respuesta de fondo dentro del término legal.", { after: 320 }));

  // Firma
  hijos.push(txt("Cordialmente,", { after: 700 }));
  hijos.push(txt("_______________________________________", { after: 40, align: AlignmentType.LEFT }));
  hijos.push(txt(P.nombre, { bold: true, after: 0, align: AlignmentType.LEFT }));
  hijos.push(txt(`C.C. No. ${P.cc} de ${P.ccExp}`, { after: 0, align: AlignmentType.LEFT }));
  hijos.push(txt(P.email, { after: 0, align: AlignmentType.LEFT }));

  return hijos;
}

// ---------- generar los tres documentos ----------
const salida = path.join(__dirname, "salida");
fs.mkdirSync(salida, { recursive: true });

(async () => {
  for (const d of DESTINOS) {
    const doc = new Document({
      creator: P.nombre,
      title: `Derecho de Petición – Prescripción – ${d.nombreCorto}`,
      description: "Derecho de petición para declaratoria oficiosa de prescripción de sanciones por infracciones de tránsito",
      styles: { default: { document: { run: { font: FUENTE, size: T } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 620 },
          },
        },
        headers: { default: membrete() },
        footers: { default: pie() },
        children: construirCuerpo(d),
      }],
    });
    const buf = await Packer.toBuffer(doc);
    const f = path.join(salida, `${d.archivo}.docx`);
    fs.writeFileSync(f, buf);
    console.log("OK ->", f, `(${d.comparendos.length} comparendo(s))`);
  }
})();
