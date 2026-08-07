# -*- coding: utf-8 -*-
"""Integra el avance en el anteproyecto original, respetando estilos y campos de Word."""
import os, shutil, re
from xml.etree import ElementTree as ET

BASE = "/tmp/claude-0/-home-user-EXTRAPROCESAL/85bb4587-828c-5654-b9e1-8dcea31c6b7f/scratchpad/work"
SRC = f"{BASE}/unpacked2"
FIG = f"{BASE}/an"
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
_crudo = open(f"{SRC}/word/document.xml", encoding="utf-8").read()
_i = _crudo.find('<w:document'); _j = _crudo.find('>', _i) + 1
ETIQUETA_RAIZ = _crudo[_i:_j]
for _p, _u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', ETIQUETA_RAIZ):
    ET.register_namespace(_p, _u)
for pfx, uri in [('a', 'http://schemas.openxmlformats.org/drawingml/2006/main'),
                 ('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')]:
    ET.register_namespace(pfx, uri)

tree = ET.parse(f"{SRC}/word/document.xml")
root = tree.getroot()
body = root.find(W + 'body')

def esc(t):
    return (t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))

def parse(xml):
    return list(ET.fromstring(f"<root {NS}>{xml}</root>"))

TABS = '<w:tabs><w:tab w:val="left" w:pos="426"/><w:tab w:val="left" w:pos="993"/></w:tabs>'

def par(text, bold=False, italic=False):
    rpr = '<w:rPr>' + ('<w:b/><w:bCs/>' if bold else '') + ('<w:i/><w:iCs/>' if italic else '') + '</w:rPr>'
    return parse(f'<w:p><w:pPr>{TABS}</w:pPr><w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def rich_par(pieces):
    """pieces: lista de (texto, negrita)"""
    runs = "".join(
        f'<w:r><w:rPr>{"<w:b/><w:bCs/>" if b else ""}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'
        for t, b in pieces)
    return parse(f'<w:p><w:pPr>{TABS}</w:pPr>{runs}</w:p>')

def subtitulo(text):
    return parse(f'<w:p><w:pPr>{TABS}<w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
                 f'<w:r><w:rPr><w:rStyle w:val="Textoennegrita"/><w:bCs w:val="0"/></w:rPr>'
                 f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def titulo1(text):
    return parse(f'<w:p><w:pPr><w:pStyle w:val="Ttulo1"/>'
                 f'<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFill="background1"/></w:pPr>'
                 f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def vineta(text):
    return parse(f'<w:p><w:pPr><w:pStyle w:val="Prrafodelista"/>'
                 f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="25"/></w:numPr>{TABS}'
                 f'<w:jc w:val="both"/></w:pPr>'
                 f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>')

def leyenda(tipo, numero, text):
    """Pie de figura o tabla con campo SEQ, igual que los del documento original."""
    fmt = '<w:i w:val="0"/><w:iCs w:val="0"/><w:color w:val="000000" w:themeColor="text1"/>'
    return parse(
        f'<w:p><w:pPr><w:pStyle w:val="Descripcin"/><w:keepNext/>'
        f'<w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:rPr>{fmt}</w:rPr><w:t xml:space="preserve">{tipo} </w:t>'
        f'<w:fldChar w:fldCharType="begin"/>'
        f'<w:instrText xml:space="preserve"> SEQ {tipo} \\* ARABIC </w:instrText>'
        f'<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr>{fmt}<w:noProof/></w:rPr><w:t>{numero}</w:t>'
        f'<w:fldChar w:fldCharType="end"/></w:r>'
        f'<w:r><w:rPr>{fmt}</w:rPr><w:t xml:space="preserve">. {esc(text)}</w:t></w:r></w:p>')

def tabla(anchos, filas):
    grid = "".join(f'<w:gridCol w:w="{a}"/>' for a in anchos)
    trs = []
    for i, fila in enumerate(filas):
        tcs = []
        for j, celda in enumerate(fila):
            neg = '<w:b/><w:bCs/>' if i == 0 else ''
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="DCE6F1"/>' if i == 0 else ''
            jc = '<w:jc w:val="center"/>' if j > 0 else ''
            tcs.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{anchos[j]}" w:type="dxa"/>{shd}</w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>'
                f'<w:ind w:firstLine="0"/>{jc}<w:rPr>{neg}<w:sz w:val="20"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr>{neg}<w:sz w:val="20"/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(str(celda))}</w:t></w:r></w:p></w:tc>')
        head = '<w:trPr><w:tblHeader/></w:trPr>' if i == 0 else ''
        trs.append(f'<w:tr>{head}{"".join(tcs)}</w:tr>')
    total = sum(anchos)
    return parse(
        f'<w:tbl><w:tblPr><w:tblStyle w:val="Tablaconcuadrcula"/>'
        f'<w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="center"/>'
        f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" '
        f'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{"".join(trs)}</w:tbl>')

# ---------------- imagenes ----------------
EMU = 914400
sect = body.findall(W + 'sectPr')[-1]
pg = sect.find(W + 'pgSz'); mar = sect.find(W + 'pgMar')
ancho_util = int(pg.get(W + 'w')) - int(mar.get(W + 'left')) - int(mar.get(W + 'right'))
ANCHO_EMU = int(ancho_util / 1440 * EMU)
print("ancho util:", ancho_util, "twips =", round(ancho_util / 1440, 2), "pulgadas")

rels_path = f"{SRC}/word/_rels/document.xml.rels"
rels = open(rels_path).read()
next_id = max(int(m) for m in re.findall(r'Id="rId(\d+)"', rels)) + 1
imgs = {}
for i, (clave, archivo, prop) in enumerate([
        ("fig1", "fig1_eda.png", 9 / 3.7),
        ("fig2", "fig2_clases.png", 11 / 3.9),
        ("fig3", "fig3_multivariado.png", 9.5 / 4.2)]):
    rid = f"rId{next_id + i}"
    destino = f"image{10 + i}.png"
    shutil.copy(f"{FIG}/{archivo}", f"{SRC}/word/media/{destino}")
    rels = rels.replace('</Relationships>',
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
        f'2006/relationships/image" Target="media/{destino}"/></Relationships>')
    imgs[clave] = (rid, ANCHO_EMU, int(ANCHO_EMU / prop), 100 + i)
open(rels_path, "w").write(rels)

def figura(clave):
    rid, cx, cy, pid = imgs[clave]
    return parse(
        f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>'
        f'<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{pid}" name="Imagen {pid}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="Imagen {pid}"/>'
        f'<pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')

# ---------------- utilidades de edicion de texto ----------------
def textos(p):
    return [t for t in p.iter(W + 't')]

def reemplazar_en_parrafo(idx, viejo, nuevo):
    p = body[idx]
    ts = textos(p)
    completo = "".join(t.text or "" for t in ts)
    assert viejo in completo, f"no encontrado en {idx}: {viejo[:60]}"
    # caso simple: dentro de un solo run
    for t in ts:
        if t.text and viejo in t.text:
            t.text = t.text.replace(viejo, nuevo)
            return
    # caso partido: se concentra en el primero y se vacian los demas
    ts[0].text = completo.replace(viejo, nuevo)
    for t in ts[1:]:
        t.text = ""

def reemplazar_por_runs(idx, textos_por_run):
    """Asigna un texto a cada run del parrafo, conservando el formato propio de cada uno."""
    p = body[idx]
    runs = p.findall(W + 'r')
    for k, r in enumerate(runs):
        ts = [t for t in r.iter(W + 't')]
        if not ts:
            continue
        ts[0].text = textos_por_run[k] if k < len(textos_por_run) else ""
        ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        for t in ts[1:]:
            t.text = ""

def anexar_a_parrafo(idx, texto):
    p = body[idx]
    ts = textos(p)
    ts[-1].text = (ts[-1].text or "") + texto

def quitar_resaltado(idx):
    p = body[idx]
    for hl in list(p.iter(W + 'highlight')):
        for padre in p.iter():
            if hl in list(padre):
                padre.remove(hl)

def insertar(idx, elementos):
    for k, el in enumerate(elementos):
        body.insert(idx + k, el)

# ============================================================
#   CONTENIDO DE LA INTEGRACION
# ============================================================
BLOQUES = []   # (indice_de_insercion, [elementos])

# ---------- A. Resultados preliminares (capitulo nuevo, antes del cronograma) ----------
cap = []
cap += titulo1('Resultados')
cap += par('Este capítulo presenta los resultados obtenidos sobre la base de datos del Centro de '
           'Investigaciones Oceanográficas e Hidrográficas del Caribe, correspondiente a las cinco campañas '
           'realizadas entre 2021 y 2023 en la Bahía de Cartagena. Todos los contrastes se realizaron sobre '
           'los 112 registros que superaron el control de calidad descrito en la metodología, con el índice '
           'calculado como el cociente entre la absorción del fitoplancton a 440 y a 676 nm.')

cap += subtitulo('Comportamiento general del índice y de las variables ambientales')
cap += par('El índice presenta una mediana general de 2,330 y un recorrido entre 1,18 y 4,84. La turbidez '
           'es la variable con mayor asimetría, con una mediana de 5,30 unidades nefelométricas y un máximo '
           'de 134,1, razón por la cual se trabajó con su logaritmo decimal en los análisis multivariados. '
           'La descripción por campaña se presenta en la tabla siguiente.')
cap += leyenda('Tabla', 5, 'Estadística descriptiva del índice Blue/Red y de las variables ambientales por '
                           'campaña, sobre los registros que superaron el control de calidad. La temperatura '
                           'se expresa en grados Celsius y la turbidez en unidades nefelométricas.')
cap += tabla([1750, 620, 1450, 1450, 1450, 1450], [
    ['Campaña', 'n', 'Temperatura media', 'Salinidad media', 'Turbidez mediana', 'Índice mediano'],
    ['Época seca 2021', '20', '29,69', '23,99', '5,53', '2,301'],
    ['Época seca 2022', '20', '29,64', '25,15', '8,84', '2,723'],
    ['Época seca 2023', '29', '31,88', '22,06', '3,96', '2,172'],
    ['Época lluviosa 2022', '29', '30,79', '21,19', '5,30', '2,009'],
    ['Época lluviosa 2023', '14', '30,64', '24,34', '4,11', '2,850'],
    ['Conjunto', '112', '30,68', '23,01', '5,30', '2,330']])
cap += figura('fig1')
cap += leyenda('Figura', 5, 'a) Distribución del índice Blue/Red por campaña, con los umbrales de '
                            'clasificación de Ciotti y Bricaud. b) Índice frente a la salinidad, con la '
                            'turbidez en escala logarítmica representada por el color. Fuente propia.')
cap += par('El resultado más informativo es que la variabilidad entre campañas de un mismo '
           'periodo climático es del mismo orden que la variabilidad entre periodos. Las campañas de época '
           'seca de 2022 y 2023 difieren entre sí más de lo que difieren la época seca y la época lluviosa '
           'consideradas en bloque, hecho que condiciona toda la lectura posterior.')

cap += subtitulo('Supuestos estadísticos y correlaciones')
cap += par('La prueba de Shapiro Wilk rechaza la normalidad en tres de las cuatro variables, de modo que '
           'la elección entre Pearson y Spearman queda resuelta a favor de la segunda.')
cap += leyenda('Tabla', 6, 'Prueba de normalidad de Shapiro Wilk sobre los registros válidos.')
cap += tabla([2300, 900, 1500, 1600, 2900], [
    ['Variable', 'n', 'Estadístico W', 'Probabilidad', 'Decisión'],
    ['Temperatura', '109', '0,986', '0,298', 'Compatible con la normalidad'],
    ['Salinidad', '111', '0,966', '0,007', 'No normal'],
    ['Turbidez', '111', '0,528', 'menor que 0,001', 'No normal, fuerte asimetría'],
    ['Índice Blue/Red', '112', '0,914', 'menor que 0,001', 'No normal']])
cap += par('La prueba de Levene entre épocas climáticas resulta no significativa para la salinidad, la '
           'turbidez y el índice, y significativa para la temperatura, con una probabilidad menor que 0,001. '
           'Los coeficientes de correlación de Spearman se resumen a continuación.')
cap += leyenda('Tabla', 7, 'Correlaciones de Spearman entre las variables ambientales y el índice Blue/Red. '
                           'El signo negativo se indica con el símbolo matemático de resta.')
cap += tabla([2600, 900, 1100, 1400, 3200], [
    ['Relación', 'n', 'rho', 'Probabilidad', 'Lectura'],
    ['Índice y temperatura', '109', '−0,273', '0,004', 'Significativa y negativa'],
    ['Índice y salinidad', '111', '+0,240', '0,011', 'Significativa y positiva'],
    ['Índice y turbidez', '111', '+0,118', '0,217', 'No significativa'],
    ['Índice y salinidad en época lluviosa', '43', '+0,463', '0,002', 'La relación más fuerte del conjunto'],
    ['Índice y temperatura en época seca', '66', '−0,326', '0,008', 'Se mantiene dentro de la época seca'],
    ['Salinidad y turbidez', '111', '−0,385', 'menor que 0,001', 'Confirma el gradiente estuarino']])
cap += par('La asociación entre el índice y la salinidad es débil en el conjunto completo y desaparece '
           'dentro de la época seca, pero se convierte en la relación dominante en la época lluviosa. El '
           'gradiente de agua dulce actúa sobre la estructura de tamaños sobre todo cuando el Canal del '
           'Dique descarga con mayor intensidad.')

cap += subtitulo('Diferencias entre épocas, campañas y zonas')
cap += leyenda('Tabla', 8, 'Resultados de las pruebas de hipótesis sobre los registros válidos.')
cap += tabla([2900, 3000, 1300, 2000], [
    ['Contraste', 'Prueba', 'Probabilidad', 'Conclusión'],
    ['Índice entre época seca y época lluviosa', 'Mann Whitney, 69 frente a 43 estaciones', '0,126',
     'Sin diferencia significativa'],
    ['Índice entre las cinco campañas', 'Kruskal Wallis con 4 grados de libertad', '0,001',
     'Diferencia significativa'],
    ['Época lluviosa 2022 frente a época seca 2022', 'Mann Whitney con Bonferroni', '0,001',
     'Diferencia significativa'],
    ['Época seca 2022 frente a época seca 2023', 'Mann Whitney con Bonferroni', '0,005',
     'Diferencia dentro de la misma época'],
    ['Clase de tamaño y época climática', 'Chi cuadrado con 2 grados de libertad', '0,806', 'Sin asociación'],
    ['Índice entre zonas de influencia del Canal del Dique', 'Mann Whitney, 65 frente a 46 estaciones',
     '0,004', 'Diferencia significativa']])
cap += figura('fig2')
cap += leyenda('Figura', 6, 'a) Composición porcentual de clases de tamaño por campaña. b) Índice según la '
                            'zona de influencia del Canal del Dique. c) Relación entre el índice y la '
                            'temperatura superficial. Fuente propia.')
cap += rich_par([
    ('La hipótesis formulada en el planteamiento del problema no se sostiene con estos datos. ', True),
    ('El microfitoplancton representa el 56,5 por ciento de las estaciones en la época seca y el 62,8 por '
     'ciento en la época lluviosa, sin asociación estadística entre clase de tamaño y época climática. La '
     'diferencia entre campañas de un mismo periodo climático, en cambio, sí es significativa. Por esta '
     'razón la hipótesis se reformuló en el capítulo correspondiente.', False)])
cap += par('El contraste que sí resulta significativo es el espacial. En la zona de alta influencia del '
           'Canal del Dique la mediana del índice es de 2,149, frente a 2,588 en la zona de baja influencia, '
           'con una probabilidad de 0,004. Las aguas influidas por el canal presentan células de mayor '
           'tamaño, resultado coherente con el aporte de nutrientes que favorece a las diatomeas y que '
           'respalda el tercer objetivo específico de este trabajo.')

cap += subtitulo('Estructura multivariada')
cap += par('El análisis de componentes principales se realizó sobre las cuatro variables estandarizadas, '
           'con la turbidez transformada a logaritmo decimal, para las 109 estaciones con información '
           'completa. Los dos primeros componentes reúnen el 69,4 por ciento de la varianza.')
cap += leyenda('Tabla', 9, 'Varianza explicada y cargas principales del análisis de componentes.')
cap += tabla([1900, 1900, 1800, 2000, 2000], [
    ['Componente', 'Varianza explicada', 'Acumulada', 'Carga mayor', 'Segunda carga'],
    ['CP1', '40,2 por ciento', '40,2 por ciento', 'Salinidad +0,842', 'Temperatura −0,623'],
    ['CP2', '29,2 por ciento', '69,4 por ciento', 'Índice −0,716', 'Turbidez −0,658'],
    ['CP3', '18,6 por ciento', '87,9 por ciento', 'Temperatura −0,622', 'Índice −0,563'],
    ['CP4', '12,1 por ciento', '100 por ciento', '', '']])
cap += figura('fig3')
cap += leyenda('Figura', 7, 'a) Plano de los dos primeros componentes principales, con las estaciones '
                            'diferenciadas por época climática. b) Dendrograma de conglomerados por el '
                            'método de Ward. Fuente propia.')
cap += par('El primer componente es un eje estuarino que opone la salinidad alta a la temperatura y la '
           'turbidez altas, y el índice casi no participa en él. El segundo componente sí ordena el índice, '
           'en oposición a la turbidez. La nube de puntos no se separa por época climática. El agrupamiento '
           'jerárquico de Ward, con una correlación cofenética de 0,514, sugiere tres grupos.')
cap += leyenda('Tabla', 10, 'Caracterización de los tres conglomerados obtenidos por el método de Ward.')
cap += tabla([1100, 800, 1600, 1400, 1400, 1200, 2100], [
    ['Grupo', 'n', 'Temperatura', 'Salinidad', 'Turbidez', 'Índice', 'Clase dominante'],
    ['1', '36', '29,82', '30,78', '5,49', '2,78', 'Nanofitoplancton, 16 estaciones'],
    ['2', '21', '31,07', '15,69', '34,47', '2,84', 'Microfitoplancton, 10 estaciones'],
    ['3', '52', '31,11', '20,42', '4,90', '2,13', 'Microfitoplancton, 44 de 52 estaciones']])
cap += par('El grupo 1 reúne las aguas marinas menos alteradas y es el único donde domina el '
           'nanofitoplancton. El grupo 2 corresponde a la pluma activa del Canal del Dique, con salinidad '
           'de 15,7 y turbidez media de 34,5 unidades. El grupo 3, el más numeroso, representa la condición '
           'intermedia de la bahía interior, con un predominio muy marcado del microfitoplancton. Ninguno '
           'de los tres grupos es exclusivo de una época climática.')

cap += subtitulo('Modelo descriptivo')
cap += par('Se ajustó una regresión múltiple del logaritmo decimal del índice sobre la temperatura, la '
           'salinidad y el logaritmo de la turbidez, para las 109 estaciones completas.')
cap += leyenda('Tabla', 11, 'Modelo de regresión múltiple sobre el logaritmo del índice.')
cap += tabla([2500, 1600, 1500, 1500, 2900], [
    ['Término', 'Coeficiente', 'Error típico', 'Probabilidad', 'Lectura'],
    ['Constante', '+0,8268', '0,2813', '0,004', ''],
    ['Temperatura', '−0,0176', '0,0086', '0,044', 'Único término significativo'],
    ['Salinidad', '+0,0031', '0,0017', '0,077', 'Marginal'],
    ['Logaritmo de la turbidez', '+0,0313', '0,0203', '0,127', 'No significativo'],
    ['Ajuste global', 'R cuadrado de 0,096', '', '', 'R cuadrado ajustado de 0,070']])
cap += par('La temperatura, la salinidad y la turbidez explican en conjunto menos del diez por ciento de '
           'la variabilidad del índice, y solo la temperatura resulta significativa. El resultado delimita '
           'el alcance del modelo descriptivo: la estructura de tamaños del fitoplancton en la bahía no '
           'queda determinada por las variables físicas del agua, de modo que el modelo se construye sobre '
           'el conjunto ampliado que incorpora la clorofila a, los nutrientes y los sólidos suspendidos '
           'totales.')

cap += subtitulo('Síntesis de los resultados')
cap += vineta('El control de calidad excluye 19 de los 131 registros, es decir el 14,5 por ciento, casi '
              'todos asociados a estaciones de turbidez elevada.')
cap += vineta('La elección del par de longitudes de onda cambia la clase de tamaño de hasta el 16,9 por '
              'ciento de las estaciones, por lo que fijarlo es una condición previa a cualquier resultado.')
cap += vineta('No hay diferencia del índice entre época seca y época lluviosa, pero sí entre campañas y '
              'entre zonas de influencia del Canal del Dique.')
cap += vineta('El conglomerado más numeroso corresponde a la bahía interior y concentra el predominio del '
              'microfitoplancton, con 44 de sus 52 estaciones.')
BLOQUES.append((451, cap))

# ---------- C. Modelado descriptivo: criterio de respaldo ----------
mod = []
mod += par('Criterio de respaldo: el ajuste realizado con temperatura, salinidad y turbidez '
           'alcanza un coeficiente de determinación de 0,096, insuficiente para un modelo predictivo. Por '
           'ello se establece que, si el ajuste sobre el conjunto completo de variables no alcanza un '
           'coeficiente de determinación de al menos 0,4, se optará por un modelo de clasificación de la '
           'clase de tamaño dominante en lugar de una predicción del valor continuo del índice, y su '
           'desempeño se reportará mediante validación cruzada y matriz de confusión.')
BLOQUES.append((447, mod))

# ---------- D. Definicion operativa de zonas (tras el tratamiento estadistico) ----------
zon = []
zon += par('Definición operativa de las zonas: se clasifican como estaciones de alta influencia del Canal '
           'del Dique aquellas con salinidad inferior a 25, umbral que separa las aguas de mezcla estuarina '
           'de las aguas de carácter marino en la bahía, y como estaciones de baja influencia las restantes. '
           'Este criterio se adopta porque la base de datos entregada no incluye las coordenadas de cada '
           'estación, y se validará contra la posición geográfica y la distancia a la desembocadura del '
           'canal cuando esa información esté disponible.')
BLOQUES.append((407, zon))

# ---------- E. Control de calidad de los datos (subseccion nueva) ----------
qc = []
qc += subtitulo('Control de calidad de los datos')
qc += par('Antes del análisis estadístico se aplica un control de calidad, necesario porque en aguas de '
          'turbidez elevada la absorción del fitoplancton se obtiene como diferencia entre dos señales '
          'grandes y muy parecidas, de modo que el resultado queda dominado por el error. Se adoptan cuatro '
          'criterios de exclusión.')
qc += vineta('Se descartan las estaciones sin espectro de absorción del fitoplancton.')
qc += vineta('Se descartan los valores del índice inferiores a 1,0 o superiores a 6,0, por hallarse fuera '
             'del intervalo físicamente admisible para el marco de Ciotti y Bricaud.')
qc += vineta('Se descartan las estaciones en las que la absorción del material no algal supera el noventa '
             'por ciento de la absorción del particulado total a 440 nm, condición en la que la señal del '
             'fitoplancton es un residuo numérico.')
qc += vineta('Se verifican la temperatura, la salinidad y la turbidez frente a los rangos esperados en la '
             'bahía y se distingue el cero real del cero usado como dato ausente.')
qc += leyenda('Tabla', 3, 'Resultado de la aplicación del control de calidad sobre la base consolidada.')
qc += tabla([4400, 1300, 3300], [
    ['Hallazgo', 'Registros', 'Tratamiento'],
    ['Índice no calculado en la hoja de origen, por errores de división por cero o de valor', '11',
     'Sin espectro disponible. Se declaran ausentes'],
    ['Índice fuera del intervalo admisible de 1,0 a 6,0', '8', 'Se excluyen del análisis'],
    ['Temperatura fuera del rango físico esperado', '1', 'Se marca como error de registro'],
    ['Salinidad o turbidez consignadas como cero', '3', 'Se distingue el cero real del dato ausente'],
    ['Registros válidos para el análisis estadístico', '112', 'Base de trabajo definitiva']])
qc += par('La aplicación de estos criterios deja 112 registros válidos de los 131 iniciales. Las estaciones '
          'excluidas por valor imposible del índice corresponden en su mayoría a condiciones de turbidez '
          'muy alta, con casos extremos de 242 y 183,5 unidades nefelométricas.')
BLOQUES.append((398, qc))

# ---------- F. Par de longitudes de onda (tras los umbrales de tamano) ----------
onda = []
onda += par('La base de datos entregada incluye el índice calculado con dos pares de longitudes de onda. '
            'Para verificar la incidencia de esa elección se recalculó el índice desde los espectros de '
            'absorción del fitoplancton en cinco combinaciones, sobre las 83 estaciones válidas con '
            'espectro disponible.')
onda += leyenda('Tabla', 2, 'Sensibilidad de la clasificación de tamaños frente al par de longitudes de '
                            'onda utilizado.')
onda += tabla([2100, 1300, 1400, 2100, 2100], [
    ['Par de longitudes de onda', 'Mediana', 'Desviación', 'Diferencia mediana frente a la referencia',
     'Estaciones que cambian de clase'],
    ['440 y 676 nm, referencia', '2,418', '0,750', 'referencia', 'referencia'],
    ['440 y 675 nm', '2,383', '0,727', '0,049', '3 de 83'],
    ['443 y 677 nm', '2,373', '0,800', '0,055', '11 de 83'],
    ['443 y 676 nm', '2,362', '0,815', '0,075', '11 de 83'],
    ['443 y 675 nm', '2,286', '0,790', '0,117', '14 de 83']])
onda += par('El par de 443 y 675 nm reclasifica al 16,9 por ciento de las estaciones respecto del par con '
            'el que está construida la base entregada. Por esa razón se adopta el par de 440 y 676 nm en '
            'todo el documento: es el que está efectivamente calculado en las cinco campañas, corresponde '
            'al máximo real de absorción de la clorofila a en la banda roja de estos espectros y es el '
            'menos sensible a la posición exacta del máximo azul. Conviene señalar que en el compendio las '
            'columnas rotuladas como cociente entre 443 y 677 nm están calculadas con el denominador de '
            '676 nm, error de rótulo que debe corregirse en el archivo de origen.')
BLOQUES.append((397, onda))

# ---------- G. Inventario de la base consolidada ----------
inv = []
inv += par('La consolidación de las cinco campañas en una sola tabla analizable arroja 131 registros de '
           'estación. No todas cuentan con espectro de absorción del fitoplancton: la campaña de época seca '
           'de 2023 no dispone de libro espectral y ocho estaciones de la campaña de 2021 carecen de '
           'espectro registrado.')
inv += leyenda('Tabla', 1, 'Inventario de la base de datos consolidada por campaña.')
inv += tabla([2400, 1900, 2100, 2700], [
    ['Campaña', 'Estaciones', 'Con espectro de aphy', 'Registros válidos tras el control de calidad'],
    ['Época seca 2021', '32', '24', '20'],
    ['Época seca 2022', '23', '23', '20'],
    ['Época seca 2023', '32', '0', '29'],
    ['Época lluviosa 2022', '29', '29', '29'],
    ['Época lluviosa 2023', '15', '15', '14'],
    ['Total', '131', '91', '112']])
inv += par('La correspondencia entre las dos campañas de 2023 y las fechas del 14 de junio y del 21 de '
           'diciembre se verifica contra el registro de campo, porque la numeración consecutiva de las '
           'muestras sugiere que la campaña rotulada como de época seca es la de junio, mes que pertenece '
           'a la temporada lluviosa.')
BLOQUES.append((382, inv))

# ============================================================
#   EDICIONES SOBRE TEXTO EXISTENTE
# ============================================================
# resumen
reemplazar_en_parrafo(111, 'roja (675 nm)', 'roja (676 nm)')
anexar_a_parrafo(111, ' El análisis de los 131 registros de las cinco campañas, de los cuales 112 '
                      'superaron el control de calidad, muestra que el índice no difiere '
                      'entre la época seca y la época lluviosa, con una probabilidad de 0,126, pero sí '
                      'entre campañas, con una probabilidad de 0,001, y entre las zonas de alta y baja '
                      'influencia del Canal del Dique, con una probabilidad de 0,004.')
# abstract
reemplazar_en_parrafo(126, 'red (675 nm) wavelengths', 'red (676 nm) wavelengths')
anexar_a_parrafo(126, ' The analysis of 131 records from the five campaigns, 112 of which passed quality '
                      'control, shows no difference in the index between the dry and '
                      'the rainy season (p = 0.126), but significant differences among campaigns (p = 0.001) '
                      'and between areas of high and low influence of the Canal del Dique (p = 0.004).')
# hipotesis
reemplazar_en_parrafo(303,
    'Las fluctuaciones estacionales y los aportes de agua turbia del Canal del Dique afectan '
    'significativamente el dominio de tamaños en la comunidad fitoplanctónica de la Bahía de Cartagena, de '
    'modo que en la época seca predomina el microfitoplancton, mientras que en la época lluviosa se observa '
    'una mayor diversidad en el dominio de tamaños debido a la variabilidad en las condiciones abióticas y '
    'los niveles de nutrientes.',
    'La estructura de tamaños del fitoplancton en la Bahía de Cartagena responde principalmente al gradiente '
    'estuarino generado por los aportes del Canal del Dique y a la variabilidad interanual de la temperatura '
    'superficial, y solo de manera secundaria a la alternancia entre la época seca y la época lluviosa. Se '
    'espera, en consecuencia, que las estaciones de baja salinidad y alta turbidez presenten un predominio '
    'de microfitoplancton y que las diferencias entre campañas de un mismo periodo climático sean '
    'comparables o mayores que las diferencias entre periodos. El análisis presentado en el '
    'capítulo de resultados respalda esta formulación.')
# calculo del indice
reemplazar_en_parrafo(391, 'a 443 nm (longitud de onda azul) y a 675 nm (longitud de onda roja)',
                           'a 440 nm (longitud de onda azul) y a 676 nm (longitud de onda roja)')
# pruebas de hipotesis y entorno de trabajo
reemplazar_por_runs(405, [
    'Pruebas de hipótesis: ',
    'dado que las estaciones no se repiten de forma completa entre campañas, el diseño no corresponde a '
    'medidas repetidas sino a muestras independientes desbalanceadas. Se emplea la prueba de Mann Whitney '
    'para las comparaciones entre dos grupos, la prueba de Kruskal Wallis para la comparación entre las '
    'cinco campañas, con comparaciones por pares corregidas por el método de Bonferroni, y la prueba de '
    'chi cuadrado para la asociación entre la clase de tamaño y la época climática. Estas pruebas permiten '
    'identificar diferencias significativas entre épocas, entre campañas y entre zonas de influencia del '
    'Canal del Dique.'])
reemplazar_en_parrafo(406,
    'Todos los análisis serán procesados en RStudio, validando previamente los supuestos estadísticos '
    'mediante las pruebas de normalidad de Shapiro-Wilk y homogeneidad de varianzas de Levene. Este enfoque '
    'integral fortalecerá la robustez de los resultados y permitirá interpretar con mayor precisión las '
    'dinámicas ecológicas en la Bahía de Cartagena.',
    'Los supuestos se validan previamente mediante la prueba de normalidad de Shapiro Wilk y la prueba de '
    'homogeneidad de varianzas de Levene. El procesamiento se realiza en RStudio y en Python, alternativa '
    'ya prevista en el plan de contingencia de este documento, con resultados equivalentes entre ambos '
    'entornos. Este enfoque fortalece la robustez de los resultados y permite interpretar con mayor '
    'precisión las dinámicas ecológicas en la Bahía de Cartagena.')
# quitar el resaltado del bloque senalado por la direccion
for i in range(398, 407):
    quitar_resaltado(i)
# resultados esperados
anexar_a_parrafo(488, ' La caracterización por clases de tamaño de las cinco campañas, la relación del '
                      'índice con la temperatura, la salinidad y la turbidez y la comparación entre las '
                      'zonas de influencia del Canal del Dique se desarrollan en el capítulo de resultados, '
                      'y sobre ellas se apoya la contextualización climática mediante las anomalías '
                      'estandarizadas de la temperatura superficial del mar.')

# ---------- insercion de bloques, de atras hacia adelante ----------
for idx, elementos in sorted(BLOQUES, key=lambda x: -x[0]):
    insertar(idx, elementos)
    print(f"insertados {len(elementos)} elementos en la posición {idx}")

# ---------- renumeracion de los pies existentes ----------
def renumerar_leyenda(texto_ancla, viejo, nuevo):
    for p in body.iter(W + 'p'):
        txt = "".join(t.text or "" for t in p.iter(W + 't'))
        if texto_ancla in txt and txt.strip().startswith('Tabla'):
            for t in p.iter(W + 't'):
                if (t.text or "").strip() == viejo:
                    t.text = nuevo
                    return True
    return False

print("renumerar SST:", renumerar_leyenda('Coordenadas geográficas de las series temporales', '1', '4'))
print("renumerar cronograma:", renumerar_leyenda('Cronograma de actividades del proyecto', '2', '12'))
print("renumerar presupuesto:", renumerar_leyenda('Descripción presupuestal', '3', '13'))

# referencias cruzadas en el texto
def renumerar_referencia(fragmento, viejo, nuevo):
    for p in body.iter(W + 'p'):
        txt = "".join(t.text or "" for t in p.iter(W + 't'))
        if fragmento in txt:
            for t in p.iter(W + 't'):
                if (t.text or "").strip() == viejo:
                    t.text = nuevo
                    return True
    return False

print("ref Tabla 1 en texto:", renumerar_referencia('posición geográfica de cada estación', '1', '4'))
print("ref Tabla 2 en texto:", renumerar_referencia('se propone un cronograma de actividades', '2', '12'))

tree.write(f"{SRC}/word/document.xml", xml_declaration=True, encoding='UTF-8', method='xml')
salida = open(f"{SRC}/word/document.xml", encoding="utf-8").read()
k = salida.find('<w:document'); m = salida.find('>', k) + 1
generada = salida[k:m]
uris_originales = {u for _, u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', ETIQUETA_RAIZ)}
extras = [f'xmlns:{p}="{u}"' for p, u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', generada)
          if u not in uris_originales]
raiz = ETIQUETA_RAIZ[:-1] + (" " + " ".join(extras) if extras else "") + ">"
if extras:
    print("declaraciones anadidas a la raiz:", extras)
salida = salida[:k] + raiz + salida[m:]
open(f"{SRC}/word/document.xml", "w", encoding="utf-8").write(salida)
print("documento.xml escrito, etiqueta raiz restaurada")
