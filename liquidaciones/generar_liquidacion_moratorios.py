# -*- coding: utf-8 -*-
"""
Genera las liquidaciones de intereses moratorios en Excel aplicando, mes a mes,
la tasa de mora maxima legal (1.5 x Interes Bancario Corriente certificado por
la Superintendencia Financiera de Colombia para consumo y ordinario).

Los libros quedan con formulas vivas: si se corrige una tasa o una fecha en la
hoja "Tasas", toda la liquidacion se recalcula sola.
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FUENTE = "Arial"
AZUL = Font(name=FUENTE, size=10, color="0000FF")          # dato duro / editable
NEGRO = Font(name=FUENTE, size=10)                          # formula
VERDE = Font(name=FUENTE, size=10, color="008000")          # enlace a otra hoja
TITULO = Font(name=FUENTE, size=14, bold=True)
SUBTITULO = Font(name=FUENTE, size=11, bold=True)
ENCABEZADO = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
NEGRITA = Font(name=FUENTE, size=10, bold=True)

FILL_ENCABEZADO = PatternFill("solid", fgColor="1F3864")
FILL_EDITABLE = PatternFill("solid", fgColor="FFFF00")
FILL_TOTAL = PatternFill("solid", fgColor="D9E1F2")
FILL_AVISO = PatternFill("solid", fgColor="FCE4D6")

BORDE = Border(*[Side(style="thin", color="8EA9DB")] * 4)

PESOS = '$#,##0;($#,##0);-'
PESOS_CENT = '$#,##0.00;($#,##0.00);-'
PCT = '0.00%'
PCT6 = '0.000000%'
FECHA = 'DD/MM/YYYY'

# IBC (consumo y ordinario) certificado por la Superfinanciera - vigencia 2026.
# La tasa de mora maxima legal es 1.5 x IBC (art. 884 C.Co.), que coincide con
# la tasa de usura certificada para la misma modalidad.
IBC_2026 = {
    (2026, 4): (0.1784, "Resolucion 0517 de 2026 (27-mar-2026)"),
    (2026, 5): (0.1878, "Resolucion 0662 de 2026"),
    (2026, 6): (0.1919, "Resolucion 0823 de 2026"),
    (2026, 7): (0.1919, "Resolucion 0965 de 2026 (30-jun-2026)"),
    (2026, 8): (0.1977, "Resolucion 1139 de 2026"),
    (2026, 9): (0.1977, "PENDIENTE de certificacion - se proyecta con el IBC de agosto/2026"),
}

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def primer_dia_siguiente_mes(f: date) -> date:
    return date(f.year + (f.month == 12), f.month % 12 + 1, 1)


def tramos(inicio: date, fin: date):
    """Parte el periodo [inicio, fin) en tramos que no cruzan fin de mes."""
    out, cursor = [], inicio
    while cursor < fin:
        corte = min(primer_dia_siguiente_mes(cursor), fin)
        out.append((cursor, corte))
        cursor = corte
    return out


def ancho(ws, anchos):
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def hoja_tasas(wb, usados):
    ws = wb.create_sheet("Tasas")
    ancho(ws, [22, 16, 18, 52])

    ws["A1"] = "Tasas certificadas - Superintendencia Financiera de Colombia"
    ws["A1"].font = TITULO
    ws["A2"] = "Credito de consumo y ordinario. Tasa de mora maxima legal = 1,5 x IBC (art. 884 C.Co.)."
    ws["A2"].font = Font(name=FUENTE, size=10, italic=True)

    encabezados = ["Mes de vigencia", "IBC E.A.", "Tasa mora E.A.", "Acto administrativo / observacion"]
    for c, texto in enumerate(encabezados, start=1):
        cel = ws.cell(row=4, column=c, value=texto)
        cel.font = ENCABEZADO
        cel.fill = FILL_ENCABEZADO
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = BORDE

    fila = 5
    refs = {}
    for clave in usados:
        ibc, fuente = IBC_2026[clave]
        anio, mes = clave
        ws.cell(row=fila, column=1, value=f"{MESES[mes]} {anio}").font = NEGRITA
        c_ibc = ws.cell(row=fila, column=2, value=ibc)
        c_ibc.font = AZUL
        c_ibc.number_format = PCT
        c_ibc.fill = FILL_EDITABLE
        c_mora = ws.cell(row=fila, column=3, value=f"=B{fila}*$B$1000")
        c_mora.font = NEGRO
        c_mora.number_format = PCT
        c_fte = ws.cell(row=fila, column=4, value=fuente)
        c_fte.font = NEGRO
        c_fte.alignment = Alignment(wrap_text=True, vertical="center")
        if "PENDIENTE" in fuente:
            for col in range(1, 5):
                ws.cell(row=fila, column=col).fill = FILL_AVISO
            c_ibc.comment = Comment(
                "A la fecha de elaboracion (25/08/2026) la Superfinanciera aun no habia "
                "certificado el IBC de septiembre de 2026. Se proyecta con el IBC de agosto "
                "(19,77 % E.A.). Reemplace esta celda por la tasa certificada cuando se publique "
                "y el libro se recalcula solo.", "Liquidacion")
        for col in range(1, 5):
            ws.cell(row=fila, column=col).border = BORDE
        refs[clave] = fila
        fila += 1

    ws["A1000"] = "Factor legal de mora (1,5 veces el IBC)"
    ws["A1000"].font = NEGRITA
    ws["B1000"] = 1.5
    ws["B1000"].font = AZUL
    ws["B1000"].fill = FILL_EDITABLE
    ws["B1000"].number_format = '0.00"x"'
    ws["C1000"] = "Art. 884 Codigo de Comercio / art. 305 Codigo Penal (usura)."
    ws["C1000"].font = Font(name=FUENTE, size=10, italic=True)

    ws.cell(row=fila + 1, column=1, value="Celdas amarillas: unicas editables.").font = Font(
        name=FUENTE, size=9, italic=True, color="0000FF")
    return ws, refs


def hoja_metodologia(wb, capital, inicio, fin, obligacion):
    ws = wb.create_sheet("Metodologia")
    ancho(ws, [100])
    ws["A1"] = "Metodologia de la liquidacion"
    ws["A1"].font = TITULO

    lineas = [
        "",
        "1. Obligacion liquidada",
        f"   Capital: {capital:,.0f} pesos.".replace(",", "."),
        f"   Periodo de mora: {inicio.strftime('%d/%m/%Y')} a {fin.strftime('%d/%m/%Y')} "
        f"({(fin - inicio).days} dias).",
        f"   {obligacion}",
        "",
        "2. Tasa aplicada",
        "   Se aplica, mes a mes, la tasa de mora maxima legal: 1,5 veces el Interes Bancario",
        "   Corriente para credito de consumo y ordinario certificado por la Superintendencia",
        "   Financiera de Colombia para cada mes de vigencia (art. 884 del Codigo de Comercio).",
        "   Esa tasa coincide con la tasa de usura certificada para la misma modalidad, de modo",
        "   que la liquidacion no excede el limite del art. 305 del Codigo Penal.",
        "",
        "3. Conteo de dias",
        "   El periodo se parte en tramos que no cruzan fin de mes, para que a cada dia se le",
        "   aplique la tasa vigente en su propio mes. Cada tramo se cuenta como diferencia de",
        "   fechas (fecha final menos fecha inicial): el dia inicial se cuenta y el dia final",
        "   no, de modo que ningun dia queda contado dos veces en el paso de un mes al otro.",
        "",
        "4. Formula",
        "   Intereses del tramo = Capital x [ (1 + tasa mora E.A.) ^ (dias / 365) - 1 ]",
        "   Es la conversion de la tasa efectiva anual a su equivalente por el numero exacto de",
        "   dias del tramo. Los intereses se calculan siempre sobre el capital insoluto: no se",
        "   capitalizan intereses (prohibicion de anatocismo, art. 2235 C.C. y art. 886 C.Co.).",
        "",
        "5. Codigo de colores",
        "   Azul: dato tomado de fuente externa o suministrado por el usuario (editable).",
        "   Amarillo: celda de entrada que el usuario puede modificar.",
        "   Negro: resultado calculado por formula. Verde: enlace a otra hoja del mismo libro.",
        "",
        "6. Advertencia",
        "   Esta liquidacion es un calculo aritmetico. La tasa efectivamente aplicable depende",
        "   del titulo, del pacto de las partes y de lo ordenado en el mandamiento de pago; si el",
        "   titulo pacto una tasa inferior, debe reemplazarse en la hoja Tasas.",
    ]
    for i, texto in enumerate(lineas, start=2):
        cel = ws.cell(row=i, column=1, value=texto)
        cel.font = SUBTITULO if texto[:2] in ("1.", "2.", "3.", "4.", "5.", "6.") else Font(name=FUENTE, size=10)
    return ws


def construir(ruta, capital, inicio, fin, titulo, obligacion):
    wb = Workbook()
    # Excel recalcula todas las formulas al abrir el archivo.
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = "Liquidacion"

    usados = []
    for ini_t, fin_t in tramos(inicio, fin):
        clave = (ini_t.year, ini_t.month)
        if clave not in usados:
            usados.append(clave)

    ancho(ws, [22, 14, 14, 9, 13, 15, 16, 20, 34])

    ws["A1"] = "LIQUIDACION DE INTERESES MORATORIOS"
    ws["A1"].font = TITULO
    ws["A2"] = titulo
    ws["A2"].font = SUBTITULO
    ws["A3"] = "Tasa de mora maxima legal (1,5 x IBC certificado por la Superfinanciera), mes a mes"
    ws["A3"].font = Font(name=FUENTE, size=10, italic=True)

    # --- Datos de la obligacion -------------------------------------------
    ws["A5"] = "DATOS DE LA OBLIGACION"
    ws["A5"].font = SUBTITULO

    datos = [
        ("Capital (base de liquidacion)", capital, PESOS),
        ("Fecha inicial de mora", inicio, FECHA),
        ("Fecha final de liquidacion", fin, FECHA),
    ]
    for i, (etiqueta, valor, formato) in enumerate(datos, start=6):
        ws.cell(row=i, column=1, value=etiqueta).font = NEGRITA
        cel = ws.cell(row=i, column=2, value=valor)
        cel.font = AZUL
        cel.fill = FILL_EDITABLE
        cel.number_format = formato
        cel.border = BORDE
    ws.cell(row=9, column=1, value="Total dias de mora").font = NEGRITA
    cel = ws.cell(row=9, column=2, value="=B8-B7")
    cel.font = NEGRO
    cel.number_format = '#,##0'
    cel.border = BORDE
    ws["C6"] = obligacion
    ws["C6"].font = Font(name=FUENTE, size=9, italic=True)

    # --- Tabla de tramos ---------------------------------------------------
    fila_enc = 11
    encabezados = ["Mes de vigencia", "Desde", "Hasta", "Dias", "IBC E.A.",
                   "Tasa mora E.A.", "Tasa diaria equiv.", "Intereses del periodo",
                   "Acto administrativo / observacion"]
    for c, texto in enumerate(encabezados, start=1):
        cel = ws.cell(row=fila_enc, column=c, value=texto)
        cel.font = ENCABEZADO
        cel.fill = FILL_ENCABEZADO
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = BORDE
    ws.row_dimensions[fila_enc].height = 32

    ws_tasas, refs = None, None
    lista = tramos(inicio, fin)
    fila = fila_enc + 1
    primera = fila
    for ini_t, fin_t in lista:
        clave = (ini_t.year, ini_t.month)
        f_tasa = 5 + usados.index(clave)
        ws.cell(row=fila, column=1, value=f"{MESES[ini_t.month]} {ini_t.year}").font = NEGRITA
        c = ws.cell(row=fila, column=2, value=ini_t); c.font = AZUL; c.number_format = FECHA
        c = ws.cell(row=fila, column=3, value=fin_t); c.font = AZUL; c.number_format = FECHA
        c = ws.cell(row=fila, column=4, value=f"=C{fila}-B{fila}"); c.number_format = '#,##0'; c.font = NEGRO
        c = ws.cell(row=fila, column=5, value=f"=Tasas!B{f_tasa}"); c.number_format = PCT; c.font = VERDE
        c = ws.cell(row=fila, column=6, value=f"=Tasas!C{f_tasa}"); c.number_format = PCT; c.font = VERDE
        c = ws.cell(row=fila, column=7, value=f"=(1+F{fila})^(1/365)-1"); c.number_format = PCT6; c.font = NEGRO
        c = ws.cell(row=fila, column=8, value=f"=$B$6*((1+F{fila})^(D{fila}/365)-1)")
        c.number_format = PESOS_CENT; c.font = NEGRO
        c = ws.cell(row=fila, column=9, value=f"=Tasas!D{f_tasa}")
        c.font = VERDE; c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in range(1, 10):
            ws.cell(row=fila, column=col).border = BORDE
        fila += 1
    ultima = fila - 1

    # --- Totales -----------------------------------------------------------
    f_tot = ultima + 1
    ws.cell(row=f_tot, column=1, value="TOTAL INTERESES DE MORA").font = NEGRITA
    c = ws.cell(row=f_tot, column=4, value=f"=SUM(D{primera}:D{ultima})")
    c.number_format = '#,##0'; c.font = NEGRITA
    c = ws.cell(row=f_tot, column=8, value=f"=SUM(H{primera}:H{ultima})")
    c.number_format = PESOS_CENT; c.font = NEGRITA
    for col in range(1, 10):
        ws.cell(row=f_tot, column=col).fill = FILL_TOTAL
        ws.cell(row=f_tot, column=col).border = BORDE

    f_cap = f_tot + 1
    ws.cell(row=f_cap, column=1, value="Capital").font = NEGRITA
    c = ws.cell(row=f_cap, column=8, value="=$B$6"); c.number_format = PESOS_CENT; c.font = NEGRO

    f_gran = f_cap + 1
    ws.cell(row=f_gran, column=1, value="TOTAL A PAGAR (capital + intereses)").font = Font(
        name=FUENTE, size=11, bold=True)
    c = ws.cell(row=f_gran, column=8, value=f"=H{f_tot}+H{f_cap}")
    c.number_format = PESOS_CENT
    c.font = Font(name=FUENTE, size=11, bold=True)
    for col in range(1, 10):
        ws.cell(row=f_gran, column=col).fill = FILL_TOTAL
        ws.cell(row=f_gran, column=col).border = BORDE

    # --- Notas -------------------------------------------------------------
    f_nota = f_gran + 2
    notas = [
        "Formula de cada tramo: Capital x [ (1 + tasa mora E.A.) ^ (dias / 365) - 1 ].",
        "Los dias se cuentan como diferencia de fechas: se cuenta el dia inicial y no el final, "
        "de modo que el cambio de mes no duplica ningun dia.",
        "Los intereses se calculan sobre el capital insoluto; no se capitalizan (no hay anatocismo).",
        "Celdas azules o amarillas: datos editables. Celdas verdes: traidas de la hoja Tasas.",
        "Detalle de tasas y fuentes en la hoja Tasas; supuestos y conteo en la hoja Metodologia.",
    ]
    for i, texto in enumerate(notas):
        cel = ws.cell(row=f_nota + i, column=1, value=("- " + texto))
        cel.font = Font(name=FUENTE, size=9, italic=True)

    ws.freeze_panes = f"A{fila_enc + 1}"

    hoja_tasas(wb, usados)
    hoja_metodologia(wb, capital, inicio, fin, obligacion)
    wb.save(ruta)
    print("Generado:", ruta)


if __name__ == "__main__":
    import os
    base = os.path.dirname(os.path.abspath(__file__))

    construir(
        os.path.join(base, "Liquidacion_152646000_30abr2026_a_28may2026.xlsx"),
        capital=152_646_000,
        inicio=date(2026, 4, 30),
        fin=date(2026, 5, 28),
        titulo="Capital $152.646.000 - mora del 30/04/2026 al 28/05/2026",
        obligacion="Capital y fechas suministrados por el usuario.",
    )
    construir(
        os.path.join(base, "Liquidacion_610584000_30jul2026_a_30sep2026.xlsx"),
        capital=610_584_000,
        inicio=date(2026, 7, 30),
        fin=date(2026, 9, 30),
        titulo="Capital $610.584.000 - mora del 30/07/2026 al 30/09/2026",
        obligacion="Capital y fechas suministrados por el usuario. El IBC de septiembre/2026 aun no estaba certificado.",
    )
