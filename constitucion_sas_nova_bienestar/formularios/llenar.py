# -*- coding: utf-8 -*-
"""Diligencia los formularios oficiales sin alterar el formato.

Solo escribe valores en los campos AcroForm existentes; no modifica ni una
línea del diseño original del formato.
"""
import re
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject as _Arr, DecodedStreamObject
from pypdf.generic import ArrayObject, FloatObject, NameObject, TextStringObject


def _normalizar_rects(escritor):
    """Resuelve los /Rect guardados como referencias indirectas.

    Algunos formatos oficiales almacenan las coordenadas de los widgets como
    objetos indirectos; pypdf no puede restarlas al construir la apariencia.
    """
    for pagina in escritor.pages:
        for anot in (pagina.get("/Annots") or []):
            objeto = anot.get_object()
            rect = objeto.get("/Rect")
            if rect is None:
                continue
            objeto[NameObject("/Rect")] = ArrayObject(
                [FloatObject(float(v.get_object())) for v in rect])


def llenar(entrada, salida, valores, radios=None, need_appearances=False,
           fuente=None, tamanos=None):
    lector = PdfReader(entrada)
    escritor = PdfWriter(clone_from=entrada)
    # NeedAppearances hace que algunos visores descarten las apariencias ya
    # existentes -entre ellas las de las casillas de verificación-, así que se
    # deja apagado y se confía en las apariencias que se escriben aquí.
    escritor.set_need_appearances_writer(need_appearances)
    _normalizar_rects(escritor)

    existentes = set()
    for pagina in lector.pages:
        for anot in (pagina.get("/Annots") or []):
            objeto = anot.get_object()
            nombre = objeto.get("/T") or (objeto.get("/Parent") or {}).get("/T")
            if nombre:
                existentes.add(str(nombre))
    desconocidos = [k for k in valores if k not in existentes]
    if desconocidos:
        raise SystemExit("Campos inexistentes: %s" % desconocidos)

    # Ajuste de /DA: algunos formatos apuntan a una Helvetica sin codificación,
    # que carece de vocales acentuadas y de la eñe. Se redirige a la fuente del
    # /DR que sí las trae y, cuando hace falta, se reduce el cuerpo para que el
    # valor quepa en casillas estrechas.
    if fuente or tamanos:
        for pagina in escritor.pages:
            for anot in (pagina.get("/Annots") or []):
                objeto = anot.get_object()
                nombre = str(objeto.get("/T")
                             or (objeto.get("/Parent") or {}).get("/T") or "")
                da = objeto.get("/DA")
                if da is None or nombre not in valores:
                    continue
                nueva = str(da)
                if fuente:
                    nueva = nueva.replace("/" + fuente[0], "/" + fuente[1])
                if tamanos and nombre in tamanos:
                    nueva = re.sub(r"(/\S+)\s+[\d.]+\s+Tf",
                                   r"\1 %s Tf" % tamanos[nombre], nueva)
                objeto[NameObject("/DA")] = TextStringObject(nueva)

    for pagina in escritor.pages:
        propios = {}
        for anot in (pagina.get("/Annots") or []):
            objeto = anot.get_object()
            nombre = objeto.get("/T") or (objeto.get("/Parent") or {}).get("/T")
            if nombre in valores:
                propios[nombre] = valores[nombre]
        if propios:
            escritor.update_page_form_field_values(pagina, propios)

    def _estampar(pagina, rect):
        """Dibuja la X sobre el contenido de la página, no en la apariencia.

        Varias implementaciones no pintan la apariencia de las casillas de
        verificación; trazarla en el contenido garantiza que se vea y se
        imprima en cualquier visor. El valor del campo se conserva aparte.
        """
        x0, y0, x1, y1 = rect
        m = min(x1 - x0, y1 - y0) * 0.22
        grosor = max(0.7, min(x1 - x0, y1 - y0) * 0.10)
        ops = (
            "\nq 0 0 0 RG {g} w 1 J\n"
            "{a} {b} m {c} {d} l S\n"
            "{a} {d} m {c} {b} l S\nQ\n"
        ).format(g=grosor, a=x0 + m, b=y0 + m, c=x1 - m, d=y1 - m)
        nuevo = DecodedStreamObject()
        nuevo.set_data(ops.encode("latin-1"))
        ref = escritor._add_object(nuevo)
        contenido = pagina.get("/Contents")
        if isinstance(contenido, _Arr):
            contenido.append(ref)
        else:
            pagina[NameObject("/Contents")] = _Arr([pagina.raw_get("/Contents"), ref])

    def _marca_vectorial(objeto, estado):
        """Redibuja la marca como una X trazada, no como un glifo.

        Las apariencias originales pintan la marca con un carácter de una
        fuente simbólica que varios visores no resuelven; una X vectorial se
        ve igual en todos y es la marca que el propio formato pide.
        """
        apariencia = objeto.get("/AP", {}).get("/N", {})
        if estado not in getattr(apariencia, "keys", lambda: [])():
            return
        forma = apariencia[estado].get_object()
        caja = [float(v) for v in forma.get("/BBox", [0, 0, 9, 9])]
        ancho, alto = caja[2] - caja[0], caja[3] - caja[1]
        m = min(ancho, alto) * 0.22
        trazo = (
            "q 0 0 0 RG {g} w 1 J\n"
            "{x1} {y1} m {x2} {y2} l S\n"
            "{x1} {y2} m {x2} {y1} l S\n"
            "Q\n"
        ).format(g=max(0.7, min(ancho, alto) * 0.09),
                 x1=caja[0] + m, y1=caja[1] + m,
                 x2=caja[2] - m, y2=caja[3] - m)
        forma.set_data(trazo.encode("latin-1"))

    # Radios y casillas: fijar /V en el campo y /AS en el widget elegido
    for nombre, estado in (radios or {}).items():
        for pagina in escritor.pages:
            for anot in (pagina.get("/Annots") or []):
                objeto = anot.get_object()
                propio = objeto.get("/T") or (objeto.get("/Parent") or {}).get("/T")
                if propio != nombre:
                    continue
                estados = list(objeto.get("/AP", {}).get("/N", {}).keys())
                padre = objeto.get("/Parent")
                destino = padre.get_object() if padre is not None else objeto
                destino[NameObject("/V")] = NameObject(estado)
                elegido = estado if estado in estados else "/Off"
                objeto[NameObject("/AS")] = NameObject(elegido)
                if elegido != "/Off":
                    _estampar(pagina, [float(v) for v in objeto["/Rect"]])

    with open(salida, "wb") as fh:
        escritor.write(fh)
    print("OK ->", salida)


if __name__ == "__main__":
    pass
