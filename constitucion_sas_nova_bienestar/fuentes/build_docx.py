# -*- coding: utf-8 -*-
"""Generador de documentos Word (.docx) a partir de un marcado ligero.

Uso:  python3 build_docx.py fuente.txt salida.docx
Las dos primeras lineas del archivo fuente pueden ser directivas:
    !HEADER texto del encabezado
    !FOOTER texto del pie (antes del numero de pagina)
"""
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "Times New Roman"


# ---------------------------------------------------------------- utilidades
def _set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        conf = kwargs.get(edge)
        if conf is None:
            continue
        el = OxmlElement("w:{}".format(edge))
        for key in ("sz", "val", "color", "space"):
            if key in conf:
                el.set(qn("w:{}".format(key)), str(conf[key]))
        borders.append(el)
    tcPr.append(borders)


def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _field(paragraph, instr):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr_el)
    run._r.append(fld_end)


def _bottom_border(paragraph, sz=6, color="000000"):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


INLINE = re.compile(r"(\*\*.+?\*\*|__.+?__|\*[^*]+?\*)", re.S)


def add_runs(paragraph, text, size=None, base_bold=False, base_italic=False):
    for chunk in INLINE.split(text):
        if not chunk:
            continue
        bold, italic, underline = base_bold, base_italic, False
        body = chunk
        if chunk.startswith("**") and chunk.endswith("**"):
            bold, body = True, chunk[2:-2]
        elif chunk.startswith("__") and chunk.endswith("__"):
            underline, body = True, chunk[2:-2]
        elif chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            italic, body = True, chunk[1:-1]
        run = paragraph.add_run(body)
        run.bold = bold
        run.italic = italic
        run.underline = underline
        run.font.name = FONT
        if size:
            run.font.size = Pt(size)
    return paragraph


# ------------------------------------------------------------------ documento
def new_document(header_text=None, footer_text=None, margins=(3.0, 2.5, 2.5, 3.0)):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = style.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15

    sec = doc.sections[0]
    sec.page_width = Cm(21.59)
    sec.page_height = Cm(27.94)
    top, right, bottom, left = margins
    sec.top_margin = Cm(top)
    sec.right_margin = Cm(right)
    sec.bottom_margin = Cm(bottom)
    sec.left_margin = Cm(left)
    sec.header_distance = Cm(1.5)
    sec.footer_distance = Cm(1.5)

    if header_text:
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = hp.add_run(header_text)
        r.font.name = FONT
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        r.italic = True
        _bottom_border(hp, sz=4, color="AAAAAA")

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if footer_text:
        r = fp.add_run(footer_text + "  |  ")
        r.font.name = FONT
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    r = fp.add_run("Página ")
    r.font.name = FONT
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    _field(fp, " PAGE ")
    r = fp.add_run(" de ")
    r.font.name = FONT
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    _field(fp, " NUMPAGES ")
    for run in fp.runs:
        run.font.name = FONT
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    return doc


def build(src_path, out_path):
    with open(src_path, encoding="utf-8") as fh:
        lines = fh.read().split("\n")

    header = footer = None
    while lines and lines[0].startswith("!"):
        directive = lines.pop(0)
        tag, _, value = directive.partition(" ")
        if tag == "!HEADER":
            header = value.strip()
        elif tag == "!FOOTER":
            footer = value.strip()

    doc = new_document(header, footer)
    pending_table = None      # (widths, rows)
    pending_widths = None

    def flush_table():
        nonlocal pending_table, pending_widths
        if not pending_table:
            return
        rows = pending_table
        ncols = max(len(r[1]) for r in rows)
        table = doc.add_table(rows=0, cols=ncols)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        widths = pending_widths or [100.0 / ncols] * ncols
        col_cm = [Cm((21.59 - 5.5) * w / 100.0) for w in widths]
        table.autofit = False
        # Fijar tblW y tblGrid: LibreOffice y Word respetan la rejilla antes
        # que el ancho individual de cada celda.
        tblPr = table._tbl.tblPr
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"), str(int(sum(w.twips for w in col_cm))))
        tblW.set(qn("w:type"), "dxa")
        tblPr.append(tblW)
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)
        grid = table._tbl.find(qn("w:tblGrid"))
        if grid is not None:
            table._tbl.remove(grid)
        grid = OxmlElement("w:tblGrid")
        for width in col_cm:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(int(width.twips)))
            grid.append(gc)
        tblPr.addnext(grid)
        for kind, cells in rows:
            row = table.add_row()
            for idx in range(ncols):
                cell = row.cells[idx]
                cell.width = col_cm[idx]
                text = cells[idx] if idx < len(cells) else ""
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if kind == "TH"
                               else WD_ALIGN_PARAGRAPH.LEFT)
                add_runs(p, text, size=10, base_bold=(kind == "TH"))
                if kind == "TH":
                    _shade(cell, "D9D9D9")
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(4)
        pending_table = None
        pending_widths = None

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        tag, _, body = line.partition(" ")
        body = body.strip()

        if tag not in ("#TH", "#TR", "#TW"):
            flush_table()

        if tag == "#T":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            add_runs(p, body.upper(), size=14, base_bold=True)
        elif tag == "#ST":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(14)
            add_runs(p, body, size=12, base_bold=True)
        elif tag == "#1":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(16)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.keep_with_next = True
            add_runs(p, body.upper(), size=12, base_bold=True)
        elif tag == "#2":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            add_runs(p, body, size=12, base_bold=True)
        elif tag == "#3":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = True
            add_runs(p, body, size=12, base_bold=True, base_italic=True)
        elif tag == "#P":
            p = doc.add_paragraph()
            add_runs(p, body)
        elif tag == "#C":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_runs(p, body)
        elif tag == "#R":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_runs(p, body)
        elif tag == "#I":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1.2)
            p.paragraph_format.right_indent = Cm(1.0)
            p.paragraph_format.space_after = Pt(8)
            add_runs(p, body, size=11, base_italic=True)
        elif tag == "#N":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1.0)
            p.paragraph_format.first_line_indent = Cm(-1.0)
            add_runs(p, body)
        elif tag == "#L":
            p = doc.add_paragraph(style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(1.0)
            p.paragraph_format.space_after = Pt(4)
            add_runs(p, body)
        elif tag == "#SP":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
        elif tag == "#HR":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(8)
            _bottom_border(p, sz=6)
        elif tag == "#PB":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif tag == "#TW":
            pending_widths = [float(x) for x in body.split(",")]
        elif tag in ("#TH", "#TR"):
            if pending_table is None:
                pending_table = []
            cells = [c.strip() for c in body.split("|")]
            pending_table.append((tag[1:], cells))
        elif tag == "#SIG":
            parts = [x.strip() for x in body.split("|")]
            for _ in range(2):
                sp = doc.add_paragraph()
                sp.paragraph_format.space_after = Pt(0)
            line_p = doc.add_paragraph()
            line_p.paragraph_format.space_after = Pt(0)
            line_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            add_runs(line_p, "________________________________________")
            for txt, bold in zip(parts, [True] + [False] * (len(parts) - 1)):
                q = doc.add_paragraph()
                q.paragraph_format.space_after = Pt(0)
                q.alignment = WD_ALIGN_PARAGRAPH.LEFT
                add_runs(q, txt, size=11, base_bold=bold)
        elif tag == "#SIG2":
            left_parts, right_parts = body.split("||")
            lp = [x.strip() for x in left_parts.split("|")]
            rp = [x.strip() for x in right_parts.split("|")]
            for _ in range(2):
                sp = doc.add_paragraph()
                sp.paragraph_format.space_after = Pt(0)
            table = doc.add_table(rows=1, cols=2)
            table.style = "Normal Table"
            table.autofit = False
            for idx, parts in enumerate((lp, rp)):
                cell = table.rows[0].cells[idx]
                cell.width = Cm(8.0)
                first = cell.paragraphs[0]
                first.paragraph_format.space_after = Pt(0)
                add_runs(first, "____________________________")
                for j, txt in enumerate(parts):
                    q = cell.add_paragraph()
                    q.paragraph_format.space_after = Pt(0)
                    add_runs(q, txt, size=10, base_bold=(j == 0))
            doc.add_paragraph()
        else:
            p = doc.add_paragraph()
            add_runs(p, line)

    flush_table()
    doc.save(out_path)
    print("OK ->", out_path)


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
