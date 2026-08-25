# -*- coding: utf-8 -*-
"""Liquidacion de intereses moratorios en Excel: una sola hoja, mes a mes,
con la tasa de mora publicada para cada mes de vigencia."""

from datetime import date
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

F = "Arial"
AZUL = Font(name=F, size=10, color="0000FF")
NEGRO = Font(name=F, size=10)
NEGRITA = Font(name=F, size=10, bold=True)
TITULO = Font(name=F, size=14, bold=True)
SUB = Font(name=F, size=11, bold=True)
ENC = Font(name=F, size=10, bold=True, color="FFFFFF")

FILL_ENC = PatternFill("solid", fgColor="1F3864")
FILL_EDIT = PatternFill("solid", fgColor="FFFF00")
FILL_TOT = PatternFill("solid", fgColor="D9E1F2")
FILL_AVISO = PatternFill("solid", fgColor="FCE4D6")
BORDE = Border(*[Side(style="thin", color="8EA9DB")] * 4)

PESOS = '$#,##0;($#,##0);-'
PESOS_C = '$#,##0.00;($#,##0.00);-'
PCT = '0.00%'
PCT6 = '0.000000%'
FECHA = 'DD/MM/YYYY'

# Tasa de mora maxima legal publicada para cada mes (1,5 veces el interes
# bancario corriente de consumo y ordinario certificado por la Superfinanciera).
MORA = {
    (2026, 4): (0.2676, "Res. 0517 de 2026"),
    (2026, 5): (0.2817, "Res. 0662 de 2026"),
    (2026, 6): (0.2879, "Res. 0823 de 2026"),
    (2026, 7): (0.2879, "Res. 0965 de 2026"),
    (2026, 8): (0.2966, "Res. 1139 de 2026"),
    (2026, 9): (0.2966, "PENDIENTE de publicacion: se proyecta con la tasa de agosto/2026"),
}
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def sig_mes(f):
    return date(f.year + (f.month == 12), f.month % 12 + 1, 1)


def tramos(ini, fin):
    out, cur = [], ini
    while cur < fin:
        corte = min(sig_mes(cur), fin)
        out.append((cur, corte))
        cur = corte
    return out


def construir(ruta, capital, ini, fin, titulo):
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = "Liquidacion"

    for i, w in enumerate([24, 14, 14, 9, 16, 17, 22, 40], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws["A1"] = "LIQUIDACION DE INTERESES MORATORIOS"
    ws["A1"].font = TITULO
    ws["A2"] = titulo
    ws["A2"].font = SUB
    ws["A3"] = "Tasa de mora maxima legal publicada para cada mes de vigencia (art. 884 C.Co.)"
    ws["A3"].font = Font(name=F, size=10, italic=True)

    ws["A5"] = "DATOS DE LA OBLIGACION"
    ws["A5"].font = SUB
    for i, (et, val, fmt) in enumerate([
        ("Capital (base de liquidacion)", capital, PESOS),
        ("Fecha inicial de mora", ini, FECHA),
        ("Fecha final de liquidacion", fin, FECHA),
    ], start=6):
        ws.cell(row=i, column=1, value=et).font = NEGRITA
        c = ws.cell(row=i, column=2, value=val)
        c.font, c.fill, c.number_format, c.border = AZUL, FILL_EDIT, fmt, BORDE
    ws.cell(row=9, column=1, value="Total dias de mora").font = NEGRITA
    c = ws.cell(row=9, column=2, value="=B8-B7")
    c.font, c.number_format, c.border = NEGRO, '#,##0', BORDE

    enc = ["Mes de vigencia", "Desde", "Hasta", "Dias", "Tasa mora E.A.",
           "Tasa diaria equiv.", "Intereses del periodo", "Fuente / observacion"]
    fe = 11
    for c_, t in enumerate(enc, start=1):
        c = ws.cell(row=fe, column=c_, value=t)
        c.font, c.fill = ENC, FILL_ENC
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fe].height = 30

    fila = fe + 1
    for a, b in tramos(ini, fin):
        tasa, fuente = MORA[(a.year, a.month)]
        ws.cell(row=fila, column=1, value=f"{MESES[a.month]} {a.year}").font = NEGRITA
        c = ws.cell(row=fila, column=2, value=a); c.font, c.number_format = AZUL, FECHA
        c = ws.cell(row=fila, column=3, value=b); c.font, c.number_format = AZUL, FECHA
        c = ws.cell(row=fila, column=4, value=f"=C{fila}-B{fila}"); c.font, c.number_format = NEGRO, '#,##0'
        c = ws.cell(row=fila, column=5, value=tasa)
        c.font, c.fill, c.number_format = AZUL, FILL_EDIT, PCT
        c = ws.cell(row=fila, column=6, value=f"=(1+E{fila})^(1/365)-1")
        c.font, c.number_format = NEGRO, PCT6
        c = ws.cell(row=fila, column=7, value=f"=$B$6*((1+E{fila})^(D{fila}/365)-1)")
        c.font, c.number_format = NEGRO, PESOS_C
        c = ws.cell(row=fila, column=8, value=fuente)
        c.font = NEGRO
        c.alignment = Alignment(wrap_text=True, vertical="center")
        if "PENDIENTE" in fuente:
            for col in range(1, 9):
                ws.cell(row=fila, column=col).fill = FILL_AVISO
            ws.cell(row=fila, column=5).fill = FILL_EDIT
            ws.cell(row=fila, column=5).comment = Comment(
                "Al 25/08/2026 la Superfinanciera aun no habia publicado la tasa de "
                "septiembre/2026. Se proyecta con la de agosto (29,66 % E.A.). Reemplace "
                "esta celda cuando se publique y el libro se recalcula solo.", "Liquidacion")
        for col in range(1, 9):
            ws.cell(row=fila, column=col).border = BORDE
        fila += 1
    ult = fila - 1

    ft = ult + 1
    ws.cell(row=ft, column=1, value="TOTAL INTERESES DE MORA").font = NEGRITA
    c = ws.cell(row=ft, column=4, value=f"=SUM(D{fe+1}:D{ult})"); c.font, c.number_format = NEGRITA, '#,##0'
    c = ws.cell(row=ft, column=7, value=f"=SUM(G{fe+1}:G{ult})"); c.font, c.number_format = NEGRITA, PESOS_C
    ws.cell(row=ft + 1, column=1, value="Capital").font = NEGRITA
    c = ws.cell(row=ft + 1, column=7, value="=$B$6"); c.font, c.number_format = NEGRO, PESOS_C
    ws.cell(row=ft + 2, column=1, value="TOTAL A PAGAR (capital + intereses)").font = Font(name=F, size=11, bold=True)
    c = ws.cell(row=ft + 2, column=7, value=f"=G{ft}+G{ft+1}")
    c.font, c.number_format = Font(name=F, size=11, bold=True), PESOS_C
    for f_ in (ft, ft + 2):
        for col in range(1, 9):
            ws.cell(row=f_, column=col).fill = FILL_TOT
            ws.cell(row=f_, column=col).border = BORDE

    notas = [
        "Intereses de cada tramo = Capital x [ (1 + tasa mora E.A.) ^ (dias / 365) - 1 ].",
        "Los dias se cuentan como diferencia de fechas: se cuenta el dia inicial y no el final, "
        "de modo que el cambio de mes no duplica ningun dia.",
        "Los intereses se liquidan sobre el capital insoluto; no se capitalizan (no hay anatocismo).",
        "Celdas amarillas: unicas editables (capital, fechas y tasa de cada mes). El resto son formulas.",
    ]
    for i, t in enumerate(notas):
        ws.cell(row=ft + 4 + i, column=1, value="- " + t).font = Font(name=F, size=9, italic=True)

    ws.freeze_panes = f"A{fe+1}"
    wb.save(ruta)
    print("Generado:", ruta)


if __name__ == "__main__":
    base = os.path.dirname(os.path.abspath(__file__))
    construir(os.path.join(base, "Liquidacion_152646000_30abr2026_a_28may2026.xlsx"),
              152_646_000, date(2026, 4, 30), date(2026, 5, 28),
              "Capital $152.646.000 - mora del 30/04/2026 al 28/05/2026")
    construir(os.path.join(base, "Liquidacion_610584000_30jul2026_a_30sep2026.xlsx"),
              610_584_000, date(2026, 7, 30), date(2026, 9, 30),
              "Capital $610.584.000 - mora del 30/07/2026 al 30/09/2026")
