# Reconstruye el trabajo de LinkedIn en Word, con formato APA 7 (trabajo de estudiante) y con el
# texto reescrito de contenido.py. La fuente es el PDF del informe de Turnitin: datos/ guarda el
# texto, las tablas y las referencias extraídos con los scripts de extraccion/, y figuras/ los
# gráficos y capturas recortados de las páginas.
# Uso: python construir.py "Trabajo LinkedIn revisado v2.docx"
import json
import re
import sys
from pathlib import Path

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

import contenido as C
import v3_textos

C.NUEVO.update(v3_textos.V3)  # tercera pasada

AQUI = Path(__file__).parent
P = json.load(open(AQUI / "datos/parrafos.json"))
T = json.load(open(AQUI / "datos/tablas.json"))
REFS = json.load(open(AQUI / "datos/refs.json"))

L1 = {"Research and Analysis", "Segmentation, Targeting, and Positioning (STP)",
      "Marketing Mix: The 4Ps of Personal Branding", "LinkedIn Profile Review",
      "LinkedIn Analytics and Search Keywords", "Action Plan", "Conclusion"}
L3 = {"2021: Recovery Exposed Supply-chain Weaknesses", "2022: Strong Market Value, Persistent Congestion",
      "2023: Weaker Goods Demand and Labour Disruption", "2024: Recovery Differed by Corridor",
      "2025: Trade Reorientation Reshaped Freight Flows", "Interpretation of the Historical Trend",
      "Photo and Banner", "Headline and About", "Proposed About Text", "Experience and Education",
      "Skills, Languages and Featured", "Activity and Recommendations"}

# Tablas: (índice del rótulo "Table N", índice del título, último índice de la tabla, índice de la nota o None)
TABLAS = {16: ("Table 1", 23, 24), 51: ("Table 2", 61, 62), 66: ("Table 3", 79, None),
          91: ("Table 4", 119, 120), 130: ("Table 5", 152, 153), 158: ("Table 6", 173, None),
          192: ("Table 7", 198, 199), 207: ("Table 8", 217, None), 260: ("Table 9", 281, None),
          291: ("Table 10", 315, None), 404: ("Table A1", 453, 454), 456: ("Table A2", 470, None),
          471: ("Table A3", 491, None), 519: ("Table C1", 542, None), 544: ("Table C2", 547, None)}
FIGURAS = {25: ("1", 27), 42: ("2", 44), 81: ("3", 83), 85: ("4", 87), 126: ("5", 128),
           178: ("6", 180), 495: ("B1", None), 497: ("B2", None), 499: ("B3", None),
           501: ("B4", None), 503: ("B5", None), 505: ("B6", None), 507: ("B7", None),
           509: ("B8", None)}
ABOUT = (232, 241)
POST = (283, 288)
PROMPT = 514
REFERENCIAS = (321, 400)


def limpio(t):
    t = t.replace("[[", "").replace("]]", "")
    t = re.sub(r"(?<=[A-Za-z.])- (?=[a-z])", "-", t)  # guiones partidos al final de línea
    return t.strip()


def celdas(nombre):
    t = T[nombre]
    head, rows = t["head"], t["rows"]
    if nombre == "Table 4":
        rows = C.TABLA4
    elif nombre == "Table 8":
        head = ["Element", "Decision", "How it will be evidenced"]
    elif nombre == "Table 9":
        rows = C.TABLA9
    elif nombre == "Table 10":
        head = [["Metric", "First day", "Last day", "Change"], ["", "November 11, 2026", "Date: __________", ""]]
        rows = [r + ["", ""] for r in C.TABLA10]
    elif nombre == "Table A1":
        rows = C.TABLA_A1
    elif nombre == "Table A2":
        head = ["ID", "Link to job description", "Source type"]
        rows = [[r[0][0], r[0][2:], r[1]] for r in rows[:10]]
    elif nombre == "Table C1":
        rows = C.TABLA_C1
    elif nombre == "Table C2":
        rows = [[""] * 6 for _ in range(3)]
    return head if isinstance(head[0], list) else [head], rows


class Doc:
    def __init__(self):
        self.d = docx.Document()
        d = self.d
        st = d.styles["Normal"]
        st.font.name = "Times New Roman"
        st.font.size = Pt(12)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        pf = st.paragraph_format
        pf.line_spacing = 2.0
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        for n in ("Heading 1", "Heading 2", "Heading 3"):
            h = d.styles[n]
            h.font.name = "Times New Roman"
            h.font.size = Pt(12)
            h.font.bold = True
            h.font.italic = n == "Heading 3"
            h.font.color.rgb = RGBColor(0, 0, 0)
            rf = h.element.rPr.find(qn("w:rFonts"))
            for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                rf.attrib.pop(qn(k), None)
            rf.set(qn("w:ascii"), "Times New Roman")
            rf.set(qn("w:hAnsi"), "Times New Roman")
            hp = h.paragraph_format
            hp.line_spacing = 2.0
            hp.space_before = Pt(0)
            hp.space_after = Pt(0)
            hp.keep_with_next = True
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER if n == "Heading 1" else WD_ALIGN_PARAGRAPH.LEFT
        s = d.sections[0]
        s.page_width, s.page_height = Inches(8.5), Inches(11)
        for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            setattr(s, m, Inches(1))
        s.header_distance = Inches(0.5)
        hp = s.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = hp.add_run()
        for tipo, txt in (("begin", None), (None, " PAGE "), ("end", None)):
            if tipo:
                e = OxmlElement("w:fldChar")
                e.set(qn("w:fldCharType"), tipo)
            else:
                e = OxmlElement("w:instrText")
                e.set(qn("xml:space"), "preserve")
                e.text = txt
            r._r.append(e)

    def parrafo(self, texto="", sangria=True, alinear=WD_ALIGN_PARAGRAPH.LEFT, negrita=None, cursiva=None,
                izquierda=None, salto=False):
        p = self.d.add_paragraph()
        p.alignment = alinear
        pf = p.paragraph_format
        pf.first_line_indent = Inches(0.5) if sangria else Inches(0)
        if izquierda is not None:
            pf.left_indent = Inches(izquierda)
        if salto:
            pf.page_break_before = True
        if texto:
            r = p.add_run(texto)
            r.bold, r.italic = negrita, cursiva
        return p

    def titulo(self, texto, nivel, salto=False):
        p = self.d.add_paragraph(texto, style="Heading %d" % nivel)
        if salto:
            p.paragraph_format.page_break_before = True
        return p

    def nota(self, cuerpo):
        p = self.parrafo(sangria=False)
        p.add_run("Note.").italic = True
        p.add_run(" " + cuerpo)

    def rotulo(self, numero, titulo):
        p = self.parrafo(numero, sangria=False, negrita=True)
        p.paragraph_format.keep_with_next = True
        p = self.parrafo(titulo, sangria=False, cursiva=True)
        p.paragraph_format.keep_with_next = True

    def tabla(self, encabezados, filas):
        ncol = len(encabezados[0])
        t = self.d.add_table(rows=len(encabezados) + len(filas), cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        tblPr = t._tbl.tblPr
        bordes = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = OxmlElement("w:" + lado)
            b.set(qn("w:val"), "single" if lado in ("top", "bottom") else "nil")
            if lado in ("top", "bottom"):
                b.set(qn("w:sz"), "8")
                b.set(qn("w:color"), "000000")
            bordes.append(b)
        tblPr.append(bordes)
        w = OxmlElement("w:tblW")
        w.set(qn("w:w"), "5000")
        w.set(qn("w:type"), "pct")
        for e in tblPr.findall(qn("w:tblW")):
            tblPr.remove(e)
        tblPr.insert(1, w)
        for i, fila in enumerate(encabezados + filas):
            row = t.rows[i]
            trPr = row._tr.get_or_add_trPr()
            trPr.append(OxmlElement("w:cantSplit"))
            if i < len(encabezados):
                trPr.append(OxmlElement("w:tblHeader"))
            for j, texto in enumerate(fila):
                c = row.cells[j]
                p = c.paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.first_line_indent = Inches(0)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                if i < len(encabezados):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for k, linea in enumerate(texto.split("\n")):
                    r = p.add_run(linea)
                    r.font.size = Pt(11)
                    if k < len(texto.split("\n")) - 1:
                        r.add_break(WD_BREAK.LINE)
                if i == len(encabezados) - 1:
                    tcPr = c._tc.get_or_add_tcPr()
                    tb = OxmlElement("w:tcBorders")
                    b = OxmlElement("w:bottom")
                    for k, v in (("val", "single"), ("sz", "8"), ("color", "000000")):
                        b.set(qn("w:" + k), v)
                    tb.append(b)
                    tcPr.append(tb)
        # ancho de cada columna según la extensión de su texto
        pesos = []
        for j in range(ncol):
            textos = [f[j] for f in encabezados + filas if f[j]]
            medio = sum(len(x) for x in textos) / max(1, len(textos))
            largo = max((len(w) for x in textos for w in x.split()), default=4)
            pesos.append(max(medio, largo * 1.6, 8))
        total = sum(pesos)
        anchos = [6.5 * p / total for p in pesos]
        # mínimo para columnas de textos cortos (códigos, cifras), para que no se partan
        minimos = [min(1.6, max([len(x) for f in encabezados + filas for x in f[j].split("\n")
                                 if x and len(x) <= 25] or [0]) * 0.075 + 0.1) for j in range(ncol)]
        fijos = [j for j in range(ncol) if anchos[j] < minimos[j]]
        if fijos:
            resto = 6.5 - sum(minimos[j] for j in fijos)
            libres = sum(anchos[j] for j in range(ncol) if j not in fijos)
            anchos = [minimos[j] if j in fijos else anchos[j] * resto / libres for j in range(ncol)]
        grid = t._tbl.tblGrid
        for j, g in enumerate(grid.findall(qn("w:gridCol"))):
            g.set(qn("w:w"), str(int(1440 * anchos[j])))
        for row in t.rows:
            for j, c in enumerate(row.cells):
                c.width = Inches(anchos[j])
        if len(filas) <= 6:  # las tablas cortas no se parten entre páginas
            for row in t.rows[:-1]:
                for c in row.cells:
                    for p in c.paragraphs:
                        p.paragraph_format.keep_with_next = True
        return t

    def figura(self, nombre, con_nota=True):
        ruta = AQUI / "figuras" / ("figura_%s.png" % nombre)
        from PIL import Image
        ancho, alto = Image.open(ruta).size
        w = min(6.0, 6.0)
        h = w * alto / ancho
        if h > 4.6:
            w = 4.6 * ancho / alto
        p = self.parrafo(sangria=False)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = con_nota
        p.add_run().add_picture(str(ruta), width=Inches(w))
        return p


def main(destino):
    D = Doc()
    # Portada (APA 7, trabajo de estudiante)
    for _ in range(3):
        D.parrafo(sangria=False)
    D.parrafo(limpio(P[0]["text"][3:]), sangria=False, alinear=WD_ALIGN_PARAGRAPH.CENTER, negrita=True)
    D.parrafo(limpio(P[1]["text"][3:]), sangria=False, alinear=WD_ALIGN_PARAGRAPH.CENTER, negrita=True)
    D.parrafo(sangria=False)
    for i in range(2, 7):
        D.parrafo(limpio(P[i]["text"]), sangria=False, alinear=WD_ALIGN_PARAGRAPH.CENTER)
    # Primera página del texto: título en negrita y centrado
    D.parrafo(limpio(P[7]["text"][3:]), sangria=False, alinear=WD_ALIGN_PARAGRAPH.CENTER, negrita=True,
              salto=True)

    i = 8
    salto_pendiente = False
    while i < len(P):
        t = P[i]["text"]
        if i in TABLAS:
            nombre, fin, nota = TABLAS[i]
            D.rotulo(nombre, limpio(P[i + 1]["text"]))
            enc, filas = celdas(nombre)
            D.tabla(enc, filas)
            if nota is not None:
                texto_nota = C.NUEVO.get(nota) or limpio(P[nota]["text"])[len("Note. "):]
                D.nota(texto_nota)
            else:
                D.parrafo(sangria=False).paragraph_format.line_spacing = 1.0
            i = (nota if nota is not None else fin) + 1
            continue
        if i in FIGURAS:
            nombre, nota = FIGURAS[i]
            D.rotulo("Figure " + nombre, limpio(P[i + 1]["text"]))
            D.figura(nombre, con_nota=nota is not None)
            if nota is not None:
                D.nota(C.NUEVO.get(nota) or limpio(P[nota]["text"])[len("Note. "):])
            i = (nota if nota is not None else i + 1) + 1
            continue
        if i == ABOUT[0]:
            for k, par in enumerate(C.ABOUT):  # cita en bloque
                D.parrafo(par, sangria=k > 0, izquierda=0.5)
            i = ABOUT[1] + 1
            continue
        if i == POST[0]:
            for k, par in enumerate(C.POST):
                D.parrafo(par, sangria=k > 0, izquierda=0.5)
            i = POST[1] + 1
            continue
        if i == PROMPT:
            D.parrafo(C.PROMPT_INTRO)
            cita = limpio(t)
            cita = cita[cita.index("“"):]
            D.parrafo(cita.strip("“”"), sangria=False, izquierda=0.5)
            i += 1
            continue
        if i == REFERENCIAS[0]:
            D.titulo("References", 1, salto=True)
            for runs in REFS:
                p = D.parrafo(sangria=False)
                p.paragraph_format.left_indent = Inches(0.5)
                p.paragraph_format.first_line_indent = Inches(-0.5)
                for texto, it in runs:
                    p.add_run(texto).italic = it or None
            i = REFERENCIAS[1] + 1
            continue
        if t.startswith("## "):
            h = limpio(t[3:])
            if re.fullmatch(r"Appendix [A-C]", h):
                D.titulo(h, 1, salto=True)
                D.titulo(limpio(P[i + 1]["text"][3:]), 1)
                i += 2
                continue
            nivel = 1 if h in L1 else 3 if h in L3 else 2
            D.titulo(h, nivel, salto=h == "Conclusion")
            i += 1
            continue
        nuevo = C.NUEVO.get(i, limpio(t))
        for par in (nuevo if isinstance(nuevo, list) else [nuevo]):
            D.parrafo(par)
        i += 1
    D.d.save(destino)


if __name__ == "__main__":
    main(sys.argv[1])
