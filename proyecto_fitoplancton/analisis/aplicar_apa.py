# -*- coding: utf-8 -*-
"""Aplica el formato de las normas APA, séptima edición, al documento ya integrado."""
import re, sys
from xml.etree import ElementTree as ET

SRC = "/tmp/claude-0/-home-user-EXTRAPROCESAL/85bb4587-828c-5654-b9e1-8dcea31c6b7f/scratchpad/work/unpacked2"
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ---------------------------------------------------------------- estilos
est = open(f"{SRC}/word/styles.xml", encoding="utf-8").read()

# interlineado doble en el estilo normal, que es la base de casi todos
est = est.replace(
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:qFormat/><w:rsid w:val="008D0435"/><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/>',
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:qFormat/><w:rsid w:val="008D0435"/><w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="auto"/>')

# los pies de tabla y figura van en el mismo cuerpo de letra que el texto, en negro
est = re.sub(
    r'(<w:style w:type="paragraph" w:styleId="Descripcin">.*?)<w:pPr>.*?</w:pPr>.*?<w:rPr>.*?</w:rPr>',
    r'\1<w:pPr><w:spacing w:after="0" w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>'
    r'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    r'<w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
    est, flags=re.S)

# el título de nivel uno no lleva sangría de primera línea
est = est.replace(
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Ttulo1APA">'
    '<w:name w:val="Título 1 APA"/><w:basedOn w:val="Normal"/><w:link w:val="Ttulo1APACar"/>'
    '<w:autoRedefine/><w:rsid w:val="00A04D12"/><w:pPr><w:numPr><w:ilvl w:val="12"/></w:numPr>'
    '<w:ind w:firstLine="720"/>',
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Ttulo1APA">'
    '<w:name w:val="Título 1 APA"/><w:basedOn w:val="Normal"/><w:link w:val="Ttulo1APACar"/>'
    '<w:autoRedefine/><w:rsid w:val="00A04D12"/><w:pPr><w:numPr><w:ilvl w:val="12"/></w:numPr>'
    '<w:ind w:firstLine="0"/>')

open(f"{SRC}/word/styles.xml", "w", encoding="utf-8").write(est)
print("estilos: interlineado doble, pies en cuerpo 12 y título uno sin sangría")

# ---------------------------------------------------------------- documento
crudo = open(f"{SRC}/word/document.xml", encoding="utf-8").read()
i = crudo.find('<w:document'); j = crudo.find('>', i) + 1
RAIZ = crudo[i:j]
for p, u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', RAIZ):
    if not re.fullmatch(r'ns\d+', p):      # ElementTree reserva ese formato de prefijo
        ET.register_namespace(p, u)

tree = ET.parse(f"{SRC}/word/document.xml")
body = tree.getroot().find(W + 'body')

def texto(p):
    return "".join(t.text or "" for t in p.iter(W + 't'))

def en_tabla(p, dentro):
    return id(p) in dentro

# paragrafos que viven dentro de una tabla: conservan su formato compacto
dentro_tabla = {id(p) for tbl in body.iter(W + 'tbl') for p in tbl.iter(W + 'p')}

# 1. alineación a la izquierda y interlineado heredado en el texto corrido
quitados_jc = quitados_sp = 0
for p in body.iter(W + 'p'):
    if en_tabla(p, dentro_tabla):
        continue
    ppr = p.find(W + 'pPr')
    if ppr is None:
        continue
    for jc in ppr.findall(W + 'jc'):
        if jc.get(W + 'val') == 'both':
            ppr.remove(jc); quitados_jc += 1
    for sp in ppr.findall(W + 'spacing'):
        if sp.get(W + 'line') == '360':
            ppr.remove(sp); quitados_sp += 1
print(f"cuerpo: {quitados_jc} justificaciones retiradas, {quitados_sp} interlineados de una y media")

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# 1b. los subtítulos en negrita suelta pasan a ser títulos de nivel dos
MAYUSCULAS = {"MARCO REFERENCIAL": "Marco Referencial", "ESTADO DEL ARTE": "Estado del Arte"}
promovidos = 0
for p in body.iter(W + 'p'):
    if en_tabla(p, dentro_tabla):
        continue
    ppr = p.find(W + 'pPr')
    if ppr is not None and ppr.find(W + 'pStyle') is not None:
        continue
    runs = [r for r in p.findall(W + 'r') if "".join(x.text or "" for x in r.iter(W + 't')).strip()]
    if not runs:
        continue
    t = texto(p).strip()
    if not t or len(t) > 90:
        continue
    def es_negrita(r):
        rs = r.find(f'{W}rPr/{W}rStyle')
        return (rs is not None and rs.get(W + 'val') == 'Textoennegrita') or r.find(f'{W}rPr/{W}b') is not None
    if not all(es_negrita(r) for r in runs):
        continue
    if ppr is None:
        ppr = ET.Element(W + 'pPr'); p.insert(0, ppr)
    for e in list(ppr):
        if e.tag in (W + 'jc', W + 'ind', W + 'tabs', W + 'rPr', W + 'spacing'):
            ppr.remove(e)
    ppr.insert(0, ET.fromstring(f'<root {NS}><w:pStyle w:val="Ttulo2"/></root>')[0])
    for r in runs:
        rpr = r.find(W + 'rPr')
        if rpr is not None:
            for tag in ('rStyle', 'b', 'bCs', 'color', 'sz', 'szCs'):
                for e in rpr.findall(W + tag):
                    rpr.remove(e)
    if t in MAYUSCULAS:
        for x in p.iter(W + 't'):
            if x.text and x.text.strip() == t:
                x.text = MAYUSCULAS[t]
    promovidos += 1
print(f"subtítulos convertidos en títulos de nivel dos: {promovidos}")

# 2. pies de tabla y de figura en formato APA

def nuevo_parrafo(xml):
    return ET.fromstring(f"<root {NS}>{xml}</root>")[0]

def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def parrafo_nota(txt):
    return nuevo_parrafo(
        '<w:p><w:pPr><w:pStyle w:val="Descripcin"/><w:spacing w:after="240" w:line="480" '
        'w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>'
        '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">Nota. </w:t></w:r>'
        f'<w:r><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def partir_titulo(t):
    """Separa el título breve de la nota explicativa."""
    m = re.search(r'(?<=[a-záéíóúñ0-9\)])\.\s+(?=[A-ZÉÁÍÓÚ])', t)
    if m and m.start() > 25:
        return t[:m.start() + 1].strip(), t[m.end():].strip()
    return t.strip(), ""

def es_pie(p):
    instr = "".join(x.text or "" for x in p.iter(W + 'instrText'))
    return 'SEQ Figura' in instr or 'SEQ Tabla' in instr

def tiene_imagen(el):
    return el.tag == W + 'p' and el.find(f'.//{W}drawing') is not None

hijos = list(body)
reemplazo, antes_de, despues_de = {}, {}, {}
n_fig = n_tab = 0

for idx, el in enumerate(hijos):
    if el.tag != W + 'p' or not es_pie(el):
        continue
    runs = el.findall(W + 'r')
    fin = next((k for k, r in enumerate(runs)
                if any(f.get(W + 'fldCharType') == 'end' for f in r.iter(W + 'fldChar'))), None)
    if fin is None:
        continue
    cola = runs[fin + 1:]
    resto = "".join("".join(x.text or "" for x in r.iter(W + 't')) for r in cola)
    resto = resto.lstrip(". ").strip()
    titulo, nota = partir_titulo(resto)

    for r in cola:
        el.remove(r)
    # el número queda en negrita
    for r in runs[:fin + 1]:
        rpr = r.find(W + 'rPr')
        if rpr is None:
            rpr = ET.SubElement(r, W + 'rPr'); r.insert(0, rpr)
        for tag in ('i', 'iCs', 'color'):
            for e in rpr.findall(W + tag):
                rpr.remove(e)
        if rpr.find(W + 'b') is None:
            rpr.insert(0, ET.Element(W + 'b'))
            rpr.insert(1, ET.Element(W + 'bCs'))
    # salto de línea y título en cursiva, dentro del mismo párrafo
    for frag in [f'<w:r><w:br/></w:r>',
                 f'<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">{esc(titulo)}</w:t></w:r>']:
        el.append(ET.fromstring(f"<root {NS}>{frag}</root>")[0])

    es_figura = texto(el).startswith("Figura")
    if es_figura:
        n_fig += 1
        img = None
        if idx + 1 < len(hijos) and tiene_imagen(hijos[idx + 1]):
            img = idx + 1
        elif idx - 1 >= 0 and tiene_imagen(hijos[idx - 1]):
            img = idx - 1
            reemplazo[idx] = []
            antes_de.setdefault(img, []).append(el)
        if img is not None and nota:
            despues_de.setdefault(img, []).append(parrafo_nota(nota))
        elif nota:
            despues_de.setdefault(idx, []).append(parrafo_nota(nota))
    else:
        n_tab += 1
        if nota:
            destino = idx + 1 if idx + 1 < len(hijos) and hijos[idx + 1].tag == W + 'tbl' else idx
            despues_de.setdefault(destino, []).append(parrafo_nota(nota))

salida = []
for idx, el in enumerate(hijos):
    salida.extend(antes_de.get(idx, []))
    if idx in reemplazo:
        salida.extend(reemplazo[idx])
    else:
        salida.append(el)
    salida.extend(despues_de.get(idx, []))
body[:] = salida
print(f"pies en formato APA: {n_fig} figuras y {n_tab} tablas")

def insertar_ordenado(padre, elemento, precedentes):
    """El esquema de Word exige un orden fijo entre los hijos de tblPr y de tcPr."""
    pos = 0
    for k, hijo in enumerate(list(padre)):
        if hijo.tag.replace(W, '') in precedentes:
            pos = k + 1
    padre.insert(pos, elemento)


# 3. tablas sin líneas verticales y sin sombreado
BORDES = ('<w:tblBorders>'
          '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
          '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
          '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
          '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
          '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
          '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
          '</w:tblBorders>')
n_tbl = 0
for tbl in body.iter(W + 'tbl'):
    n_tbl += 1
    pr = tbl.find(W + 'tblPr')
    for b in pr.findall(W + 'tblBorders'):
        pr.remove(b)
    insertar_ordenado(pr, ET.fromstring(f"<root {NS}>{BORDES}</root>")[0],
                      {'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
                       'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd'})
    filas = tbl.findall(W + 'tr')
    for n, tr in enumerate(filas):
        for tc in tr.findall(W + 'tc'):
            tcpr = tc.find(W + 'tcPr')
            if tcpr is None:
                continue
            for sh in tcpr.findall(W + 'shd'):
                tcpr.remove(sh)
            if n == 0:
                for b in tcpr.findall(W + 'tcBorders'):
                    tcpr.remove(b)
                insertar_ordenado(tcpr, ET.fromstring(
                    f'<root {NS}><w:tcBorders><w:bottom w:val="single" w:sz="8" '
                    f'w:space="0" w:color="000000"/></w:tcBorders></root>')[0],
                    {'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge'})
print(f"tablas con líneas horizontales y sin sombreado: {n_tbl}")

tree.write(f"{SRC}/word/document.xml", xml_declaration=True, encoding='UTF-8')
sal = open(f"{SRC}/word/document.xml", encoding="utf-8").read()
k = sal.find('<w:document'); m = sal.find('>', k) + 1
gen = sal[k:m]
uris = {u for _, u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', RAIZ)}
extras = [f'xmlns:{p}="{u}"' for p, u in re.findall(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', gen)
          if u not in uris]
raiz = RAIZ[:-1] + (" " + " ".join(extras) if extras else "") + ">"
open(f"{SRC}/word/document.xml", "w", encoding="utf-8").write(sal[:k] + raiz + sal[m:])
print("documento guardado con formato APA")
