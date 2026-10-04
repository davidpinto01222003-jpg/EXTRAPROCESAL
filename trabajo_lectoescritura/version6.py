# Versión 6: formato APA 7 sobre la v5 y discusión ampliada.
# - Márgenes de 2,54 cm, interlineado doble, alineación a la izquierda, sangría de 1,27 cm y sin
#   espacio antes ni después; títulos en Times New Roman 12 negro (nivel 1 centrado y en negrita,
#   nivel 2 a la izquierda en negrita, nivel 3 en negrita y cursiva); sin párrafos vacíos (los
#   saltos de página se conservan como "salto de página anterior"). No cambia la organización:
#   secciones, niveles de título y lugar de cada elemento son los de la v5.
# - Tablas APA: solo líneas horizontales, sin sangría, alineadas al margen, número en negrita y
#   título en cursiva; Tabla 1 y Tabla 2 numeradas en el orden en que se citan; las tablas de la
#   transcripción numeradas A1 a A22 con su título.
# - Las 24 fotografías flotantes que están debajo de "Análisis de resultados" se quedan en ese
#   lugar, en línea y agrupadas como Figuras 1 a 3, con número, título y nota.
# - La tabla de contenido se regenera con los títulos y las páginas reales (requiere LibreOffice
#   Writer; se omite si no está instalado). Word la vuelve a actualizar al abrir el archivo.
# Uso: python version6.py "Trabajo lectoescritura v5.docx" "Trabajo lectoescritura v6.docx"
import os
import re
import shutil
import subprocess
import sys
import tempfile
from copy import deepcopy

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Emu
from docx.text.paragraph import Paragraph

DISCUSION_INICIO = (
    "Los resultados son coherentes con lo que la literatura revisada plantea sobre el papel de la "
    "conciencia fonológica en el aprendizaje lector (Aguirre-Medrano & González-López, 2021) y "
    "permiten precisarlo para este grupo, en el que las sílabas y las rimas se reconocen sin "
    "dificultad y los problemas aparecen al aislar fonemas y al convertir con exactitud grafemas en "
    "sonidos. La distinción tiene consecuencias pedagógicas, porque un estudiante que segmenta bien "
    "pero no aísla fonemas necesita un trabajo distinto del que requiere quien no reconoce las "
    "sílabas, y tratar ambos casos como “baja conciencia fonológica” llevaría a intervenciones poco "
    "ajustadas.")

# (comienzo del párrafo de la discusión tras el cual se inserta, texto nuevo)
DISCUSION = [
    ("Los resultados son coherentes con lo que la literatura",
     "La distribución de los errores apunta en la misma dirección. Entre el estudiante que leyó las "
     "treinta pseudopalabras sin cambios y los dos que alteraron nueve hay casos con uno, dos, cuatro "
     "y cinco cambios, sin un corte que separe a un grupo con dificultades del resto. Snowling et al. "
     "(2020) describen la dislexia de un modo parecido, como el extremo de un continuo de habilidad "
     "lectora cuyos límites dependen del criterio que se adopte, y recuerdan que esos criterios han "
     "cambiado mucho con el tiempo. Desde esa perspectiva, preguntar cuáles de estos diez estudiantes "
     "tienen dislexia importa menos que saber qué operación concreta falla en cada uno, que es la "
     "información que la prueba aporta y la que sirve para decidir qué apoyo necesita."),
    ("Los datos permiten también leer de otra manera",
     "El análisis de los grupos consonánticos deja abierta una pregunta sobre la edad de los "
     "participantes. Jiménez González y Jiménez Rodríguez (1999) encontraron que los errores en sílabas "
     "trabadas dependen del nivel de conciencia fonémica y que esa relación se debilita hacia tercer "
     "grado, cuando la habilidad ha madurado, de modo que su significado cambia con la edad del niño. "
     "En primer grado, leer “Fenir” en lugar de “Flenir” forma parte de un aprendizaje todavía en "
     "curso, mientras que en cuarto o quinto indicaría un retraso que convendría atender pronto. Como "
     "la transcripción no registra el grado ni la edad de los participantes, los datos describen con "
     "precisión el tipo de error y dejan pendiente su gravedad, y cualquier aplicación "
     "posterior de la prueba debería registrar ambos datos."),
    ("Las intervenciones revisadas encajan",
     "Esas mismas intervenciones tienen en común un problema de medición. Rivera Cintrón y Batiz "
     "Cartagena (2024) no evaluaron con pruebas estandarizadas cuánto avanzó la lectoescritura, y el "
     "ausentismo interrumpió la continuidad del programa de Domínguez Vázquez (2023), lo que hace "
     "difícil saber qué parte de la mejora corresponde a la intervención. Una prueba breve de "
     "pseudopalabras, aplicada de forma individual y con registro de cada error, puede servir como "
     "medida antes y después del trabajo en el aula si se repiten los mismos estímulos. Frente a la "
     "observación general del docente tiene la ventaja de obligar a decodificar, sin que el "
     "estudiante pueda apoyarse en palabras que ya reconoce de memoria."),
    ("La variedad de perfiles refuerza",
     "Los resultados también tocan el debate sobre la formación docente que recogen los "
     "antecedentes. Giraldo Gaviria y Caro Lopera (2022) se ocupan de la escritura académica de los "
     "futuros maestros, y Morales Londoño (2018), de las concepciones equivocadas sobre cómo se "
     "aprende a leer. A esas necesidades los datos añaden una más específica, el conocimiento del "
     "sistema fonológico del español. Un docente que corrige el dictado sin distinguir entre "
     "“Sutel” por “Zutel”, que suena igual, y “Pamil” por “Pamir”, que cambia un fonema, sumaría "
     "ambas formas como errores y atribuiría a cinco estudiantes una dificultad fonológica que en "
     "realidad es ortográfica. Algo parecido le ocurriría a quien tome la respuesta “Gro” como un "
     "acierto parcial, sin advertir que el niño dio la sílaba cuando se le pedía un sonido."),
    ("Los resultados también tocan el debate",
     "La comparación con el PROLEC-R marca otro límite del instrumento. Esa batería ofrece baremos "
     "por curso que indican si un número de errores es alto o normal para la edad del niño (Cuetos "
     "et al., 2007), mientras que la prueba diseñada para este trabajo solo permite comparar a los "
     "diez estudiantes entre sí. Decir que el estudiante 9 tiene dificultades porque alteró cuatro "
     "pseudopalabras supone una referencia que el grupo no ofrece, ya que cuatro errores pueden ser "
     "muchos en un curso y pocos en otro. La prueba sirve para describir el tipo de error y orientar "
     "el trabajo, y para juzgar la magnitud de las dificultades haría falta aplicar también la tarea "
     "de pseudopalabras del PROLEC-R o contrastar los resultados con un grupo de la misma edad."),
]

PROCEDIMIENTO = ("No se registraron tiempos de lectura ni número de pausas.",
                 "No se registraron tiempos de lectura ni número de pausas. Las hojas de respuesta "
                 "y dos fotografías de la aplicación se presentan en las Figuras 1 a 3.")

# Clasificación de las 24 fotografías por el número de su relación (rId) en la v5
APLICACION = {17, 19}
PAGINA1 = {11, 12, 13, 16, 18, 22, 23, 26, 28, 31, 33}
PAGINA2 = {14, 15, 20, 21, 24, 25, 27, 29, 30, 32, 34}
FIGURAS = [
    ("Figura 1", "Aplicación individual de la prueba de pseudopalabras", APLICACION, 6.0,
     "Fotografías tomadas durante la aplicación de la prueba."),
    ("Figura 2", "Primera página de las hojas de respuesta (lectura y dictado de pseudopalabras)",
     PAGINA1, 5.0,
     "La primera página incluye la lista de treinta pseudopalabras para la lectura en voz alta, "
     "el espacio para la escritura al dictado y la primera pregunta de conciencia fonológica."),
    ("Figura 3", "Segunda página de las hojas de respuesta (conciencia fonológica, frases y "
     "asociación grafema-fonema)", PAGINA2, 5.0,
     "La segunda página incluye las demás preguntas de conciencia fonológica, las frases con "
     "pseudopalabras y la asociación grafema-fonema, con la actividad de unir cada pseudopalabra "
     "con su sonido inicial."),
]

# Patrones que quedaban de la v5: "por eso", dos puntos que anuncian y "no pretende... sino"
AJUSTES = [
    ("transmisión del conocimiento: el alfabeto fonético, la imprenta y la era electrónica.",
     "transmisión del conocimiento, que corresponden al alfabeto fonético, la imprenta y la era electrónica."),
    ("El diseño no pretende generalizar, sino obtener información detallada sobre procesos",
     "El diseño busca información detallada sobre procesos"),
    ("concentraron los errores: nueve estudiantes", "concentraron los errores, ya que nueve estudiantes"),
    ("lo auditivo y lo visual y por eso recuperan", "lo auditivo y lo visual, con lo que recuperan"),
    ("y propone por eso contar con", "y para compensarlo propone contar con"),
    ("que por eso no puede reconocerla de memoria y debe", "que al no poder reconocerla de memoria debe"),
    ("de la misma manera y que por eso se tomó como", "de la misma manera, así que se tomó como"),
]

ETIQUETAS = {"Lista de pseudopalabras:", "Lo que dijo:", "Escribió:"}

# Orden de los hijos según el esquema de WordprocessingML
ORDEN_TBLPR = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
               "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"]
ORDEN_TCPR = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
              "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]
ORDEN_TRPR = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
              "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"]

NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def ordenar(padre, orden):
    hijos = list(padre)
    for h in hijos:
        padre.remove(h)
    clave = lambda h: orden.index(h.tag.split("}")[1]) if h.tag.split("}")[1] in orden else 99
    for h in sorted(hijos, key=clave):
        padre.append(h)


def hijo(padre, tag, orden):
    e = padre.find(qn(tag))
    if e is None:
        e = OxmlElement(tag)
        padre.append(e)
        ordenar(padre, orden)
    return e


def texto(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def es_titulo(p_el):
    st = p_el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return st is not None and st.get(qn("w:val")) in ("Ttulo1", "Ttulo2", "Ttulo3")


def poner_texto(p, nuevo):
    """Deja un solo run con el texto, con el formato de carácter del primero sin negrita ni cursiva."""
    runs = p._p.findall(qn("w:r"))
    rpr = None
    if runs and runs[0].find(qn("w:rPr")) is not None:
        rpr = deepcopy(runs[0].find(qn("w:rPr")))
        for tag in ("w:b", "w:bCs", "w:i", "w:iCs", "w:rStyle"):
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)
    for e in list(p._p):
        if e.tag != qn("w:pPr"):
            p._p.remove(e)
    r = p.add_run(nuevo)
    if rpr is not None:
        r._r.insert(0, rpr)
    return r


def formato_cuerpo(p, sangria=Cm(1.27), alinear=WD_ALIGN_PARAGRAPH.LEFT):
    pf = p.paragraph_format
    pf.line_spacing = 2.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = alinear
    if sangria is not None:
        pf.left_indent = Cm(0)
        pf.first_line_indent = sangria


HECHOS = []  # párrafos creados aquí con su formato final


def parrafo_nuevo(d, segmentos, estilo="Párr.APA", sangria=Cm(0), seguir=True):
    p = Paragraph(OxmlElement("w:p"), d._body)
    HECHOS.append(p._p)
    p.style = d.styles[estilo]
    for t, negrita, cursiva in segmentos:
        r = p.add_run(t)
        r.bold = negrita
        r.italic = cursiva
    formato_cuerpo(p, sangria)
    p.paragraph_format.keep_with_next = seguir
    return p


def rotulo(d, numero, titulo):
    """Número en negrita y título en cursiva, en dos líneas, sin sangría (APA 7)."""
    return [parrafo_nuevo(d, [(numero, True, None)]), parrafo_nuevo(d, [(titulo, None, True)])]


def nota(d, cuerpo):
    return parrafo_nuevo(d, [("Nota. ", None, True), (cuerpo, None, None)], seguir=False)


def a_linea(run, alto_cm):
    """Convierte la imagen flotante del run en una imagen en línea con el alto indicado."""
    ancla = run.find(".//{%s}anchor" % NS_WP)
    ext = ancla.find("{%s}extent" % NS_WP)
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    alto = int(Cm(alto_cm))
    ancho = int(cx * alto / cy)
    en_linea = OxmlElement("wp:inline")
    for k in ("distT", "distB", "distL", "distR"):
        en_linea.set(k, "0")
    e = OxmlElement("wp:extent")
    e.set("cx", str(ancho))
    e.set("cy", str(alto))
    en_linea.append(e)
    ee = OxmlElement("wp:effectExtent")
    for k in "ltrb":
        ee.set(k, "0")
    en_linea.append(ee)
    for tag in ("docPr", "cNvGraphicFramePr"):
        en_linea.append(deepcopy(ancla.find("{%s}%s" % (NS_WP, tag))))
    grafico = deepcopy(ancla.find("{%s}graphic" % NS_A))
    for x in grafico.iter("{%s}ext" % NS_A):
        if x.getparent().tag == "{%s}xfrm" % NS_A:
            x.set("cx", str(ancho))
            x.set("cy", str(alto))
    en_linea.append(grafico)
    dibujo = ancla.getparent()
    dibujo.replace(ancla, en_linea)
    return run


def formato_tabla(tbl, propia):
    tblPr = tbl.find(qn("w:tblPr"))
    for tag in ("w:tblStyle", "w:tblpPr", "w:tblInd", "w:tblBorders", "w:jc"):
        for e in tblPr.findall(qn(tag)):
            tblPr.remove(e)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "left")
    tblPr.append(jc)
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    tblPr.append(ind)
    bordes = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement("w:" + lado)
        if lado in ("top", "bottom"):
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), "8")
            b.set(qn("w:color"), "000000")
        else:
            b.set(qn("w:val"), "nil")
        b.set(qn("w:space"), "0")
        bordes.append(b)
    tblPr.append(bordes)
    filas = tbl.findall(qn("w:tr"))
    ncol = len(tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol")))
    tblW = tblPr.find(qn("w:tblW"))
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"), "5000")
    if not propia:
        anchos = [9360 // ncol] * ncol
    elif ncol == 4:
        anchos = [2600, 4560, 1100, 1100]
    else:
        anchos = [1300, 1850, 1450, 1700, 1350, 1710]
    for g, a in zip(tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol")), anchos):
        g.set(qn("w:w"), str(a))
    if propia:
        trPr = filas[0].find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            filas[0].insert(0, trPr)
        hijo(trPr, "w:tblHeader", ORDEN_TRPR)
    ordenar(tblPr, ORDEN_TBLPR)
    # columnas de cifras o respuestas cortas, centradas
    centradas = set()
    if propia:
        for c in range(1, ncol):
            if all(len("".join(f.findall(qn("w:tc"))[c].itertext())) <= 12 for f in filas[1:]):
                centradas.add(c)
    for fila in filas:  # una fila no se parte entre dos páginas
        trPr = fila.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tblPrEx = fila.find(qn("w:tblPrEx"))
            (tblPrEx.addnext if tblPrEx is not None else lambda e: fila.insert(0, e))(trPr)
        hijo(trPr, "w:cantSplit", ORDEN_TRPR)
    for nf, fila in enumerate(filas):
        for nc, tc in enumerate(fila.findall(qn("w:tc"))):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                for e in tcPr.findall(qn("w:tcBorders")):
                    tcPr.remove(e)
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is not None:
                    tcW.set(qn("w:w"), str(anchos[nc]))
                    tcW.set(qn("w:type"), "dxa")
                if propia and nf == 0:
                    tb = hijo(tcPr, "w:tcBorders", ORDEN_TCPR)
                    b = OxmlElement("w:bottom")
                    for k, v in (("val", "single"), ("sz", "8"), ("space", "0"), ("color", "000000")):
                        b.set(qn("w:" + k), v)
                    tb.append(b)
                ordenar(tcPr, ORDEN_TCPR)
            for p_el in tc.findall(qn("w:p")):
                p = Paragraph(p_el, None)
                pf = p.paragraph_format
                pf.line_spacing = 1.0
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.left_indent = Cm(0)
                pf.first_line_indent = Cm(0)
                if propia and (nf == 0 or nc in centradas):
                    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.size = Pt(12)


def buscar(d, inicio):
    hallados = [p for p in d.paragraphs if p.text.strip().startswith(inicio)]
    assert len(hallados) == 1, (inicio, len(hallados))
    return hallados[0]


def main(origen, destino):
    d = docx.Document(origen)
    body = d.element.body

    # 1. Tabla 2 se cita antes que la Tabla 1: se intercambian los números
    for p in d.paragraphs:
        for r in p.runs:
            if "Tabla 1" in r.text or "Tabla 2" in r.text:
                r.text = r.text.replace("Tabla 1", "Tabla §").replace("Tabla 2", "Tabla 1").replace("Tabla §", "Tabla 2")

    for viejo, nuevo in AJUSTES:
        runs = [r for p in d.paragraphs for r in p.runs if viejo in r.text]
        assert len(runs) == 1, viejo
        runs[0].text = runs[0].text.replace(viejo, nuevo)

    # 2. Discusión ampliada y remisión a las figuras
    poner_texto(buscar(d, "Los resultados son coherentes con lo que la literatura"), DISCUSION_INICIO)
    for ancla, nuevo in DISCUSION:
        a = buscar(d, ancla)
        p = deepcopy(a._p)
        a._p.addnext(p)
        poner_texto(Paragraph(p, a._parent), nuevo)
    proc = buscar(d, "La prueba se aplicó a cada estudiante por separado.")
    assert proc.text.strip().endswith(PROCEDIMIENTO[0])
    poner_texto(proc, proc.text.strip().replace(PROCEDIMIENTO[0], PROCEDIMIENTO[1]))

    # 3. Fotografías: se toman de los párrafos flotantes de "Análisis de resultados"
    analisis = buscar(d, "Análisis de resultados")._p
    fotos = {}
    sig = analisis.getnext()
    while sig is not None and sig.tag == qn("w:p") and not texto(sig).strip():
        nxt = sig.getnext()
        for run in sig.findall(qn("w:r")):
            blip = run.find(".//{%s}blip" % NS_A)
            if blip is not None:
                rid = int(blip.get("{%s}embed" % NS_R)[3:])
                fotos[rid] = run
        body.remove(sig)
        sig = nxt
    assert sorted(fotos) == sorted(APLICACION | PAGINA1 | PAGINA2), sorted(fotos)

    # 4. Anexo A: títulos de estudiante y rótulos de tabla A1-A22
    anexo_a = buscar(d, "Anexo A. Transcripción de la prueba")._p
    estudiante, n, en_anexo = None, 0, False
    for el in list(body):
        if el is anexo_a:
            en_anexo = True
        if not en_anexo:
            continue
        if el.tag == qn("w:p"):
            m = re.fullmatch(r"Estudiante (\d+)\.?", texto(el).strip())
            if m:
                estudiante = int(m.group(1))
            continue
        if el.tag != qn("w:tbl"):
            continue
        n += 1
        previo = el.getprevious()
        etiqueta = texto(previo).strip()
        if etiqueta == "Lista de pseudopalabras:":
            titulo = "Pseudopalabras presentadas para la lectura en voz alta"
        elif etiqueta == "Lo que dijo:":
            titulo = "Lectura de pseudopalabras del estudiante %d" % estudiante
        elif etiqueta == "Escribió:":
            titulo = "Escritura al dictado del estudiante %d" % estudiante
        else:
            titulo = "Pseudopalabras presentadas en el dictado"
        nuevos = rotulo(d, "Tabla A%d" % n, titulo)
        for p in nuevos:
            el.addprevious(p._p)
        if etiqueta in ETIQUETAS:
            body.remove(previo)
        formato_tabla(el, propia=False)
    assert n == 22, n

    # 6. Párrafos vacíos desde el Resumen; un salto de página pasa al párrafo siguiente
    resumen = buscar(d, "Resumen")._p
    hijos = list(body)
    desde = hijos.index(resumen)
    saltos, pendiente = [], False
    for el in hijos[desde - 1:]:
        if el.tag != qn("w:p"):
            continue
        if el.find(".//" + qn("w:sectPr")) is not None:
            pendiente = False
            continue
        if texto(el).strip() or el.find(".//" + qn("w:drawing")) is not None:
            if pendiente:
                saltos.append(el)
            pendiente = False
            continue
        a, b = el.getprevious(), el.getnext()
        if a is not None and b is not None and a.tag == b.tag == qn("w:tbl"):
            continue
        if any(br.get(qn("w:type")) == "page" for br in el.iter(qn("w:br"))):
            pendiente = True
        body.remove(el)

    # 7. Formato de párrafo del cuerpo
    referencias = buscar(d, "Referencias")._p
    primeros = {buscar(d, "Este trabajo examina")._p, buscar(d, "This study examines")._p}
    zona = "cuerpo"
    for el in list(body)[list(body).index(resumen):]:
        if el.tag != qn("w:p") or el.find(".//" + qn("w:sectPr")) is not None:
            continue
        if any(el is s for s in saltos):
            Paragraph(el, d._body).paragraph_format.page_break_before = True
        if any(el is h for h in HECHOS):
            continue
        p = Paragraph(el, d._body)
        t = p.text.strip()
        if el is referencias:
            zona = "referencias"
        elif el is anexo_a:
            zona = "anexo"
        for r in p.runs:  # espacios dobles y espacios junto a paréntesis
            nuevo = re.sub(r"[ \t]*\t[ \t]*", "\t", r.text)
            nuevo = re.sub(r"(?<=\S)\t+(?=–)", " ", nuevo)
            nuevo = re.sub(r" {2,}", " ", nuevo)
            nuevo = re.sub(r"\( +", "(", nuevo)
            nuevo = re.sub(r" +\)", ")", nuevo)
            nuevo = re.sub(r"\(/ (\w+) /\)", r"(/\1/)", nuevo)
            if nuevo != r.text:
                r.text = nuevo
        m = re.fullmatch(r"(Frando|Mesur|Tupel)\s*–\s*(/\w/)", t)
        if m:  # opciones de la actividad de unir, alineadas antes con tabuladores
            poner_texto(p, "%s – %s" % m.groups())
            t = p.text
        m = re.fullmatch(r"(¿.+(?:zunel|quarim)\?) *\( *(.*?) *\)", t)
        if m:  # respuestas entre paréntesis sin espacios interiores
            poner_texto(p, "%s (%s)" % (m.group(1), re.sub(r"/ *(\w+) */", r"/\1/", m.group(2))))
            t = p.text
        if p.runs:
            p.runs[-1].text = p.runs[-1].text.rstrip(" ")
        if es_titulo(el):
            pf = p.paragraph_format
            if p.style.name == "Heading 1" and el.getprevious().find(".//" + qn("w:sectPr")) is None:
                pf.page_break_before = True  # cada sección principal en página nueva
            pf.alignment = None
            pf.first_line_indent = None
            pf.left_indent = None
            pf.line_spacing = None
            pf.space_before = None
            pf.space_after = None
            if t.endswith(".") and p.runs:
                p.runs[-1].text = p.runs[-1].text.rstrip(".")
            continue
        if zona == "referencias":
            formato_cuerpo(p, sangria=None)
            p.paragraph_format.left_indent = Cm(1.27)
            p.paragraph_format.first_line_indent = Cm(-1.27)
        elif el.find(".//" + qn("w:drawing")) is not None:
            continue
        elif re.fullmatch(r"(Tabla|Figura) [A-Z]?\d+", t) or t.startswith("Nota.") or el in primeros:
            formato_cuerpo(p, sangria=Cm(0))
            p.paragraph_format.keep_with_next = not t.startswith("Nota.")
            for r in p.runs:
                r.font.size = Pt(12)
        elif el.getnext() is not None and el.getnext().tag == qn("w:tbl") and re.fullmatch(r"Tabla \d+", texto(el.getprevious()).strip()):
            formato_cuerpo(p, sangria=Cm(0))
            p.paragraph_format.keep_with_next = True
            for r in p.runs:
                r.italic = True
        elif p._p.find(qn("w:pPr") + "/" + qn("w:numPr")) is not None:
            formato_cuerpo(p, sangria=None)
        else:
            formato_cuerpo(p)

    # 8. Tablas 1 y 2
    propias = [t for t in body.findall(qn("w:tbl"))][:2]
    for t in propias:
        formato_tabla(t, propia=True)

    # 9. Figuras en el mismo lugar donde estaban las fotografías
    bloques = []
    for numero, titulo, ids, alto, cuerpo in FIGURAS:
        bloques += rotulo(d, numero, titulo)
        img = Paragraph(OxmlElement("w:p"), d._body)
        img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img.paragraph_format.line_spacing = 1.0
        img.paragraph_format.space_before = Pt(0)
        img.paragraph_format.space_after = Pt(6)
        img.paragraph_format.first_line_indent = Cm(0)
        img.paragraph_format.keep_with_next = True
        for i, rid in enumerate(sorted(ids)):
            if i:
                img.add_run(" ")
            img._p.append(a_linea(fotos[rid], alto))
        bloques.append(img)
        bloques.append(nota(d, cuerpo))
    for b in reversed(bloques):
        analisis.addnext(b._p)

    # 10. Estilos y márgenes
    for nombre, centrado, cursiva in (("Heading 1", True, False), ("Heading 2", False, False),
                                      ("Heading 3", False, True)):
        st = d.styles[nombre]
        f = st.font
        f.name = "Times New Roman"
        f.size = Pt(12)
        f.bold = True
        f.italic = cursiva
        f.color.rgb = RGBColor(0, 0, 0)
        rf = st.element.rPr.find(qn("w:rFonts"))
        for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            rf.attrib.pop(qn(k), None)
        rf.set(qn("w:eastAsia"), "Times New Roman")
        rf.set(qn("w:cs"), "Times New Roman")
        szcs = st.element.rPr.find(qn("w:szCs"))
        if szcs is not None:
            szcs.set(qn("w:val"), "24")
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if centrado else WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = 2.0
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.keep_with_next = True
    apa = d.styles["Párr.APA"].paragraph_format
    apa.first_line_indent = Cm(1.27)
    apa.line_spacing = 2.0
    apa.space_before = Pt(0)
    apa.space_after = Pt(0)
    apa.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for s in d.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(2.54)
        s.header_distance = s.footer_distance = Cm(1.27)

    d.save(destino)


def paginas_pdf(ruta):
    """Texto y número impreso de cada página, renderizando el documento con LibreOffice."""
    import pdfplumber
    tmp = tempfile.mkdtemp()
    try:
        shutil.copy(ruta, os.path.join(tmp, "doc.docx"))
        subprocess.run(["soffice", "-env:UserInstallation=file://" + os.path.join(tmp, "perfil"),
                        "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmp,
                        os.path.join(tmp, "doc.docx")], check=True, capture_output=True,
                       env=dict(os.environ, HOME=tmp), timeout=300)
        paginas = []
        with pdfplumber.open(os.path.join(tmp, "doc.pdf")) as pdf:
            for i, pg in enumerate(pdf.pages):
                arriba = [w["text"] for w in pg.extract_words() if w["top"] < 60 and w["text"].isdigit()]
                lineas = [re.sub(r"\s+", " ", l).strip() for l in (pg.extract_text() or "").split("\n")]
                paginas.append((arriba[-1] if arriba else str(i + 1), lineas))
        return paginas
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def entrada_indice(ppr, nivel, texto_titulo, marca, pagina, inicio_campo=None):
    p = OxmlElement("w:p")
    pPr = deepcopy(ppr)
    pPr.find(qn("w:pStyle")).set(qn("w:val"), "TDC%d" % nivel)
    p.append(pPr)
    for r in inicio_campo or []:
        p.append(r)

    def run(oculto=False, t=None, tab=False, fld=None, instr=None):
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        if not oculto:
            st = OxmlElement("w:rStyle")
            st.set(qn("w:val"), "Hipervnculo")
            rpr.append(st)
        rpr.append(OxmlElement("w:noProof"))
        if oculto:
            rpr.append(OxmlElement("w:webHidden"))
        r.append(rpr)
        if t is not None:
            e = OxmlElement("w:t")
            e.text = t
            e.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(e)
        if tab:
            r.append(OxmlElement("w:tab"))
        if fld:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), fld)
            r.append(e)
        if instr:
            e = OxmlElement("w:instrText")
            e.text = instr
            e.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(e)
        return r

    h = OxmlElement("w:hyperlink")
    h.set(qn("w:anchor"), marca)
    h.set(qn("w:history"), "1")
    for r in (run(t=texto_titulo), run(True, tab=True), run(True, fld="begin"),
              run(True, instr=" PAGEREF %s \\h " % marca), run(True, fld="separate"),
              run(True, t=pagina), run(True, fld="end")):
        h.append(r)
    p.append(h)
    return p


def actualizar_indice(ruta):
    """Reescribe el contenido guardado de la tabla de contenido con los títulos y páginas actuales."""
    d = docx.Document(ruta)
    body = d.element.body
    resumen = buscar(d, "Resumen")
    titulos = []
    en_cuerpo = False
    for p in d.paragraphs:
        if p._p is resumen._p:
            en_cuerpo = True
        m = re.fullmatch(r"Heading (\d)", p.style.name)
        if en_cuerpo and m and p.text.strip():
            titulos.append((int(m.group(1)), re.sub(r"\s+", " ", p.text).strip(), p))
    # marcadores en los títulos
    ids = [int(b.get(qn("w:id"))) for b in body.iter(qn("w:bookmarkStart"))]
    siguiente = max(ids + [0]) + 1
    marcas = []
    for k, (_, _, p) in enumerate(titulos):
        nombre = None
        for b in p._p.findall(qn("w:bookmarkStart")):
            if b.get(qn("w:name")).startswith("_Toc9"):
                nombre = b.get(qn("w:name"))
        if nombre is None:
            nombre = "_Toc9%07d" % k
            bs = OxmlElement("w:bookmarkStart")
            bs.set(qn("w:id"), str(siguiente))
            bs.set(qn("w:name"), nombre)
            be = OxmlElement("w:bookmarkEnd")
            be.set(qn("w:id"), str(siguiente))
            siguiente += 1
            ppr = p._p.find(qn("w:pPr"))
            if ppr is not None:
                ppr.addnext(bs)
            else:
                p._p.insert(0, bs)
            p._p.append(be)
        marcas.append(nombre)
    d.save(ruta)
    # páginas
    paginas = paginas_pdf(ruta)
    inicio = next(i for i, (_, ls) in enumerate(paginas) if "Tabla de contenido" in ls)
    i = inicio + 1
    while not any(l == "Resumen" for l in paginas[i][1]):
        i += 1
    numeros = []
    for _, t, _ in titulos:
        clave = t[:40]
        while not any(l.startswith(clave) or (len(t) > 40 and t.startswith(l) and len(l) > 15)
                      for l in paginas[i][1]):
            i += 1
        numeros.append(paginas[i][0])
    # contenido del campo
    sdt = next(s for s in body.iter(qn("w:sdt"))
               if any("TOC" in (x.text or "") for x in s.iter(qn("w:instrText"))))
    contenido = sdt.find(qn("w:sdtContent"))
    ps = contenido.findall(qn("w:p"))
    primera = ps[1]
    ppr = primera.find(qn("w:pPr"))
    campo = []
    for r in primera.findall(qn("w:r")):
        campo.append(r)
        if r.find(qn("w:fldChar")) is not None and r.find(qn("w:fldChar")).get(qn("w:fldCharType")) == "separate":
            break
    for p in ps[1:-1]:
        contenido.remove(p)
    ultimo = ps[-1]
    for k, ((nivel, t, _), marca, num) in enumerate(zip(titulos, marcas, numeros)):
        ultimo.addprevious(entrada_indice(ppr, nivel, t, marca, num, campo if k == 0 else None))
    d.save(ruta)
    return list(zip([t for _, t, _ in titulos], numeros))


if __name__ == "__main__":
    main(*sys.argv[1:3])
    if shutil.which("soffice"):
        actualizar_indice(sys.argv[2])
        for t, n in actualizar_indice(sys.argv[2]):  # segunda pasada por si el índice cambió de largo
            print(n.rjust(3), t[:70])
