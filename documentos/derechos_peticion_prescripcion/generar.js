const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  Header, Footer, ImageRun, PageOrientation, LevelFormat, convertInchesToTwip,
} = require("docx");
const { PETICIONARIO: P, DESTINOS } = require("./datos.js");

const FUENTE = "Arial";
const T = 22; // 11 pt

const run = (text, o = {}) => new TextRun({
  text, font: FUENTE, size: o.size || T,
  bold: !!o.bold, italics: !!o.italics, color: o.color,
});

const par = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after === undefined ? 180 : o.after, before: o.before || 0, line: o.line || 276 },
  indent: o.indent,
  border: o.border,
  numbering: o.numbering,
});

const txt = (text, o = {}) => par(run(text, o), o);

// Cita normativa: sangrada
const cita = (t) => par(run(t, { size: 21 }), {
  indent: { left: 720, right: 400 }, after: 140, line: 264,
});

// ---------- membrete real (logo JDCP + regla dorada) ----------
const ORO = "B8962E";
const LOGO = fs.readFileSync(path.join(__dirname, "logo.png"));

const membrete = () => new Header({
  children: [
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 60, line: 240 },
      children: [new ImageRun({
        type: "png",
        data: LOGO,
        // 200,25 x 66,75 pt  ->  267 x 89 px a 96 dpi (tamaño exacto del PDF original)
        transformation: { width: 267, height: 89 },
      })],
    }),
    new Paragraph({
      spacing: { after: 0, line: 240 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 2, color: ORO } },
      children: [],
    }),
  ],
});

const pie = () => new Footer({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 240 },
      border: { top: { style: BorderStyle.SINGLE, size: 6, space: 6, color: ORO } },
      children: [
        // los glifos van en Segoe UI Symbol, igual que en el membrete original
        new TextRun({ text: "\u260E", font: "Segoe UI Symbol", size: 17, color: "777777" }),
        new TextRun({ text: "  " + P.telMembrete + "      ", font: FUENTE, size: 17, color: "777777" }),
        new TextRun({ text: "\u2709", font: "Segoe UI Symbol", size: 17, color: "777777" }),
        new TextRun({ text: "  " + P.emailMembrete, font: FUENTE, size: 17, color: "777777" }),
      ],
    }),
  ],
});

// ---------- cuerpo (modelo Aguachica) ----------
function construirCuerpo(d) {
  const n = d.comparendos.length;
  const pl = n > 1;
  const h = [];

  h.push(txt(`${P.ciudadFirma}, ${P.fechaLarga}`, { align: AlignmentType.LEFT, after: 360 }));

  h.push(txt("Señores", { after: 0 }));
  h.push(txt(d.entidad, { bold: true, after: 0 }));
  h.push(txt(d.ciudad, { after: 360 }));

  h.push(par([run("REF: ", { bold: true }), run("DERECHO DE PETICIÓN", { bold: true })], { after: 300 }));

  h.push(par([
    run(P.nombre, { bold: true }),
    run(` identificado con Cédula de Ciudadanía Número ${P.cc}, actuando en nombre propio, domiciliado en ${P.domicilio}, en atención a las previsiones que consagra el `),
    run("DERECHO CONSTITUCIONAL FUNDAMENTAL DE PETICIÓN", { bold: true }),
    run(", contenidas en el Artículo 23 de la Constitución Política, desarrolladas en los Artículos 5, 6, 17, 31, 32 del Código Contencioso Administrativo, así como en el Decreto 2150 de 1995, y demás disposiciones concordantes / pertinentes, me dirijo a esa Entidad, para los efectos del Inciso 2º del Artículo 206 del Decreto Ley 019 de 2012 que modificó el Artículo 159 de la Ley 769 de 2002 “Código Nacional de Tránsito” para que:"),
  ], { after: 240 }));

  h.push(par([
    run("OFICIOSAMENTE", { bold: true }),
    run(" declare la "),
    run("PRESCRIPCIÓN", { bold: true }),
    run(pl
      ? " de las sanciones que me fueran impuestas con ocasión de las Infracciones de Tránsito relacionadas a continuación:"
      : " de la sanción que me fuera impuesta con ocasión de la Infracción de Tránsito relacionada a continuación:"),
  ], { indent: { left: 340 }, after: 200 }));

  d.comparendos.forEach((c) => {
    h.push(par(
      run(`Comparendo No ${c.numero} de fecha ${c.fechaCorta} en el lugar ${c.lugar}.`),
      { numbering: { reference: "vinetas", level: 0 }, indent: { left: 900, hanging: 260 }, after: 140 },
    ));
  });

  h.push(txt(
    `Consecuencia de lo anterior, se actualicen las bases de datos correspondientes de SIMIT, RUNT, así como todas aquellas donde aparezca como deudor de ${pl ? "estas sanciones" : "esta sanción"}.`,
    { before: 120, after: 220 },
  ));

  h.push(txt("Lo anterior en razón a que se cumple con los requisitos legales aplicables, en virtud del inciso segundo del artículo 159 del Código Nacional de Tránsito, a saber:", { after: 180 }));

  h.push(cita("“La ejecución de las sanciones que se impongan por violación de las normas de tránsito, estará a cargo de las autoridades de tránsito de la jurisdicción donde se cometió el hecho, quienes estarán investidas de jurisdicción coactiva para el cobro, cuando ello fuere necesario."));
  h.push(par([
    run("Las sanciones impuestas por infracciones a las normas de tránsito prescribirán en tres (3) años contados a partir de la ocurrencia del hecho; ", { bold: true, size: 21 }),
    run("la prescripción deberá ser declarada de oficio y se interrumpirá con la notificación del mandamiento de pago. La autoridad de tránsito no podrá iniciar el cobro coactivo de sanciones respecto de las cuales se encuentren configurados los supuestos necesarios para declarar su prescripción”", { size: 21 }),
  ], { indent: { left: 720, right: 400 }, after: 260, line: 264 }));

  // ----- fundamentos -----
  h.push(par(run("FUNDAMENTOS DE DERECHO DE LA PETICIÓN", { bold: true }), {
    align: AlignmentType.CENTER, before: 280, after: 240,
  }));

  h.push(txt("Artículo 52 de la Ley 1437 de 2011:", { bold: true, after: 140 }));
  h.push(par([
    run("Caducidad de la facultad sancionatoria. Salvo lo dispuesto en leyes especiales, la facultad que tienen las autoridades para imponer sanciones ", { size: 21 }),
    run("caduca a los tres (3) años de ocurrido el hecho, la conducta u omisión que pudiere ocasionarlas", { bold: true, size: 21 }),
    run(", término dentro del cual el acto administrativo que impone la sanción debe haber sido expedido y notificado. (Subrayado y en negrilla fuera del texto original)", { size: 21 }),
  ], { indent: { left: 720, right: 400 }, after: 240, line: 264 }));

  h.push(txt("Artículo 206 del Decreto 19 de 2012:", { bold: true, after: 140 }));
  h.push(cita("“La ejecución de las sanciones que se impongan por violación de las normas de tránsito, estará a cargo de las autoridades de tránsito de la jurisdicción donde se cometió el hecho, quienes estarán investidas de jurisdicción coactiva para el cobro, cuando ello fuere necesario."));
  h.push(par([
    run("Las sanciones impuestas por infracciones a las normas de tránsito prescribirán en tres (3) años contados a partir de la ocurrencia del hecho; ", { bold: true, size: 21 }),
    run("la prescripción deberá ser declarada de oficio y se interrumpirá con la notificación del mandamiento de pago. La autoridad de tránsito no podrá iniciar el cobro coactivo de sanciones respecto de las cuales se encuentren configurados los supuestos necesarios para declarar su prescripción.", { size: 21 }),
  ], { indent: { left: 720, right: 400 }, after: 140, line: 264 }));
  h.push(cita("Las autoridades de tránsito deberán establecer públicamente a más tardar en el mes de enero de cada año, planes y programas destinados al cobro de dichas sanciones y dentro de este mismo periodo rendirán cuentas públicas sobre la ejecución de los mismos.”."));

  h.push(txt("Ley 1437 de 2011, Artículo 14:", { bold: true, before: 140, after: 140 }));
  h.push(par([
    run("“Salvo norma legal especial y so pena de sanción disciplinaria, toda petición deberá resolverse dentro de los ", { size: 21 }),
    run("quince (15) días siguientes a su recepción", { bold: true, size: 21 }),
    run(". (Subrayado y en negrilla fuera del texto original)”", { size: 21 }),
  ], { indent: { left: 720, right: 400 }, after: 280, line: 264 }));

  // ----- cierre -----
  h.push(par([
    run("De conformidad con todo lo esbozado anteriormente, "),
    run(pl
      ? "SOLICITO DE MANERA MUY RESPETUOSA DECLAREN DE OFICIO LA PRESCRIPCIÓN DE LOS COMPARENDOS INICIALMENTE RESEÑADOS"
      : "SOLICITO DE MANERA MUY RESPETUOSA DECLAREN DE OFICIO LA PRESCRIPCIÓN DEL COMPARENDO INICIALMENTE RESEÑADO",
      { bold: true }),
    run(" y de no hacerlo me sean informadas debidamente las razones de hecho y de derecho."),
  ], { after: 300 }));

  // ----- notificación -----
  h.push(par(run("NOTIFICACIÓN", { bold: true }), { align: AlignmentType.CENTER, before: 200, after: 220 }));
  h.push(txt("Notificaciones por correo certificado por favor dirigirlas a:", { after: 140 }));
  h.push(par([run("Ciudad: ", { bold: true }), run(P.ciudadNotif)], { after: 40 }));
  h.push(par([run("Dirección: ", { bold: true }), run(P.direccion)], { after: 40 }));
  h.push(par([run("Teléfono: ", { bold: true }), run(P.cel)], { after: 180 }));
  h.push(par([
    run("Y las notificaciones por correo electrónico a la dirección "),
    run(P.email, { bold: true }),
  ], { after: 420 }));

  // ----- firma -----
  h.push(txt("Cordialmente,", { after: 700 }));
  h.push(txt(P.nombre, { bold: true, after: 0, align: AlignmentType.LEFT }));
  h.push(txt(`Cédula de ciudadanía No. ${P.cc}`, { after: 0, align: AlignmentType.LEFT }));

  return h;
}

// ---------- generar ----------
const salida = path.join(__dirname, "salida");
fs.mkdirSync(salida, { recursive: true });

const numeracion = {
  config: [{
    reference: "vinetas",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: convertInchesToTwip(0.6), hanging: convertInchesToTwip(0.18) } } },
    }],
  }],
};

(async () => {
  for (const d of DESTINOS) {
    const doc = new Document({
      creator: P.nombre,
      title: `Derecho de Petición – Prescripción – ${d.nombreCorto}`,
      description: "Derecho de petición para declaratoria oficiosa de prescripción de sanciones por infracciones de tránsito",
      styles: { default: { document: { run: { font: FUENTE, size: T } } } },
      numbering: numeracion,
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
            margin: { top: 2620, right: 1100, bottom: 1300, left: 1100, header: 700, footer: 900 },
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
    console.log("OK ->", path.basename(f), `(${d.comparendos.length} comparendo(s))`);
  }
})();
