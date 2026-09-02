"""
Enriquece la base de IPS: pagina oficial, canal de contratacion y convocatorias.

Que hace
--------
Toma la base que genera ips_area_metropolitana.py y, para cada sociedad,
averigua tres cosas:

  1. CUAL ES SU PAGINA OFICIAL. Primero por el dominio del correo que
     reporto al REPS. Si ese correo es de Gmail o no hay correo, deduce
     dominios candidatos a partir de la razon social (clinicasantamaria.com,
     .com.co, .co, siglas...) y solo acepta uno si la propia pagina
     confirma que es de esa entidad -- porque trae el NIT o el nombre.
     Nunca se queda con un dominio "que parece".

  2. POR DONDE SE LE RADICA UNA PROPUESTA. Lee las paginas publicas de
     contacto, contratacion, proveedores, convocatorias, transparencia y
     notificaciones judiciales, y extrae los correos publicados,
     clasificandolos por area.

  3. QUE CONVOCATORIAS TIENE O HA TENIDO. Por dos vias:
       - SECOP (Colombia Compra Eficiente), buscando por NIT y TAMBIEN por
         nombre de la entidad, porque muchas IPS aparecen con el NIT sin
         digito de verificacion, con otro NIT del grupo, o simplemente mal
         digitado. Los hallazgos por nombre quedan marcados como tales para
         que usted los confirme.
       - La propia pagina de la entidad: enlaces a invitaciones, terminos
         de referencia, licitaciones y pliegos publicados.

Cada dato queda con la URL de donde salio. Lo que no se encuentra queda
VACIO, con la razon escrita al lado. Nunca se rellena con suposiciones.

Buen comportamiento en la red
-----------------------------
Respeta robots.txt, se identifica con un User-Agent propio, pausa entre
peticiones al mismo sitio, limita las paginas por dominio y solo lee
paginas publicas. Nunca envia formularios ni entra a zonas privadas.

Como se usa
-----------
  python ips_enriquecer_contactos.py
  python ips_enriquecer_contactos.py --limite 25        (prueba corta)
  python ips_enriquecer_contactos.py --solo-segmento A
  python ips_enriquecer_contactos.py --sin-deducir-dominio
  python ips_enriquecer_contactos.py --sin-secop

o doble clic en ips_enriquecer_contactos.bat
"""

import argparse
import csv
import json
import os
import re
import sqlite3
import socket
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from html.parser import HTMLParser

CARPETA = "datos_ips"
BASE_POR_DEFECTO = os.path.join(CARPETA, "ips_area_metropolitana.db")

CONTACTO_RESPONSABLE = os.environ.get("CORREO_CONTACTO", "").strip()
AGENTE = (
    "EXTRAPROCESAL-contacto/1.1 (busca el canal de contratacion publicado%s)"
    % ((" ; contacto: " + CONTACTO_RESPONSABLE) if CONTACTO_RESPONSABLE else "")
)

TIEMPO_ESPERA = 20          # segundos por peticion normal
TIEMPO_SONDEO = 8           # segundos al probar un dominio candidato
PAUSA_MISMO_SITIO = 1.0     # segundos entre peticiones al mismo dominio
PAGINAS_POR_SITIO = 10      # tope de paginas leidas por dominio
DOMINIOS_A_PROBAR = 14      # tope de dominios candidatos por sociedad
SITIOS_EN_PARALELO = 6      # dominios distintos a la vez
TAMANO_MAXIMO = 2_000_000

APP_TOKEN = os.environ.get("DATOS_GOV_APP_TOKEN", "").strip()

DATASETS_SECOP = [
    ("SECOP II - Contratos electronicos", "jbjy-vk9h"),
    ("SECOP II - Procesos de contratacion", "p6dx-8zbt"),
    ("SECOP I", "xvdy-vvsk"),
]

PALABRAS_JURIDICAS = (
    "juridic", "abogad", "asesoria legal", "servicios legales",
    "representacion judicial", "defensa judicial", "apoyo legal",
    "asesoria juridica", "cobro juridico", "cobro prejuridico",
    "conciliacion", "litigio", "consultoria juridica",
)

RUTAS_CANDIDATAS = [
    "/contratacion", "/proveedores", "/convocatorias", "/contacto",
    "/transparencia", "/contactenos", "/proveedor", "/licitaciones",
    "/invitaciones", "/trabaje-con-nosotros", "/quienes-somos",
]

PALABRAS_ENLACE = (
    "contratacion", "contrataci", "proveedor", "convocatoria", "licitacion",
    "invitacion publica", "invitacion a cotizar", "contacto", "contactenos",
    "transparencia", "notificaciones judiciales", "juridica",
    "trabaje con nosotros", "compras", "pqr", "terminos de referencia",
)

# Un enlace de una pagina de contratacion que suene a convocatoria concreta.
PALABRAS_CONVOCATORIA = (
    "convocatoria", "invitacion", "licitacion", "terminos de referencia",
    "pliego", "cotizacion", "proceso de seleccion", "solicitud de oferta",
    "concurso de meritos", "adenda",
)

CATEGORIAS = [
    ("Notificaciones judiciales", (
        "notificacionesjudiciales", "notificacionjudicial", "judicial",
        "notificaciones",
    )),
    ("Contratacion / proveedores", (
        "contratacion", "contratos", "proveedores", "proveedor", "compras",
        "licitacion", "licitaciones", "convocatoria", "convocatorias",
        "adquisiciones", "suministros",
    )),
    ("Juridica", (
        "juridica", "juridico", "legal", "abogado", "secretariageneral",
    )),
    ("Gerencia / direccion", (
        "gerencia", "gerente", "direccion", "director", "presidencia",
        "administracion", "administrativa",
    )),
    ("PQRS / atencion al usuario", (
        "pqr", "pqrs", "pqrsf", "atencionalusuario", "servicioalcliente",
        "atencionusuario", "sau", "quejas",
    )),
    ("Talento humano", ("talentohumano", "rrhh", "recursoshumanos", "seleccion",
                        "hojasdevida", "empleo")),
]

PRIORIDAD_CANAL = [
    "Contratacion / proveedores",
    "Juridica",
    "Gerencia / direccion",
    "Notificaciones judiciales",
    "General",
    "PQRS / atencion al usuario",
    "Talento humano",
]

BASURA = (
    "example.com", "dominio.com", "correo.com", "tucorreo", "email.com",
    "sentry.io", "wixpress.com", "godaddy.com", "sentry-next",
    "@2x.png", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js",
)

GRATUITOS = (
    "gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "hotmail.es",
    "outlook.es", "yahoo.es", "live.com", "icloud.com", "protonmail.com",
    "msn.com", "hotmail.co", "gmail.es",
)

# Palabras demasiado comunes en el sector: no sirven para identificar a nadie.
GENERICAS = {
    "ips", "eps", "ese", "sas", "sa", "ltda", "eu", "sca", "scs", "cta",
    "clinica", "clinicas", "centro", "centros", "medico", "medica", "medicos",
    "salud", "servicios", "servicio", "integral", "integrales", "especializada",
    "especializado", "unidad", "unidades", "institucion", "prestadora",
    "fundacion", "corporacion", "asociacion", "cooperativa", "empresa",
    "social", "estado", "colombia", "colombiana", "limitada", "anonima",
    "sociedad", "grupo", "del", "de", "la", "las", "los", "el", "y", "en",
    "para", "por", "con", "sus", "una", "uno",
}

TLDS = [".com", ".com.co", ".co", ".org", ".org.co", ".net", ".edu.co"]
TLDS_PUBLICA = [".gov.co"] + TLDS


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def sin_tildes(texto):
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def clave_columna(nombre):
    return re.sub(r"[^a-z0-9]", "", sin_tildes(nombre))


def palabras_de(razon_social):
    """Parte la razon social en palabras utiles, sin puntuacion."""
    limpio = re.sub(r"[^a-z0-9 ]", " ", sin_tildes(razon_social))
    return [p for p in limpio.split() if p]


def tokens_distintivos(razon_social):
    """Las palabras que de verdad identifican a la entidad."""
    return [p for p in palabras_de(razon_social)
            if p not in GENERICAS and len(p) >= 4]


def dominio_de(texto):
    texto = (texto or "").strip().lower()
    if not texto:
        return ""
    if "@" in texto:
        dominio = texto.rsplit("@", 1)[1]
    else:
        if "://" not in texto:
            texto = "http://" + texto
        dominio = urllib.parse.urlparse(texto).netloc
    dominio = dominio.split(":")[0].strip(". ")
    if dominio in GRATUITOS or "." not in dominio:
        return ""
    return dominio


def categoria_de(correo):
    local = sin_tildes(correo.split("@", 1)[0]).replace("-", "").replace("_", "")
    for nombre, marcas in CATEGORIAS:
        if any(m in local for m in marcas):
            return nombre
    return "General"


# ---------------------------------------------------------------------------
# Descarga de paginas
# ---------------------------------------------------------------------------

class Extractor(HTMLParser):
    """Saca de una pagina los enlaces, el titulo y el texto plano."""

    def __init__(self):
        HTMLParser.__init__(self)
        self.enlaces = []
        self.titulo = ""
        self.textos = []
        self._href = None
        self._buffer = []
        self._ignorar = 0
        self._en_titulo = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._ignorar += 1
        elif tag == "title":
            self._en_titulo = True
        elif tag == "a":
            self._href = dict(attrs).get("href")
            self._buffer = []

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._ignorar:
            self._ignorar -= 1
        elif tag == "title":
            self._en_titulo = False
        elif tag == "a":
            if self._href:
                self.enlaces.append((self._href, " ".join(self._buffer).strip()))
            self._href = None
            self._buffer = []

    def handle_data(self, data):
        if self._en_titulo:
            self.titulo += data
        if self._ignorar:
            return
        self.textos.append(data)
        if self._href is not None:
            self._buffer.append(data)

    def texto(self):
        return " ".join(" ".join(self.textos).split())


PATRON_CORREO = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def bajar(url, espera=TIEMPO_ESPERA):
    """Descarga una pagina. Devuelve (html, url_final) o (None, motivo)."""
    try:
        pedido = urllib.request.Request(url, headers={
            "User-Agent": AGENTE,
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "es-CO,es;q=0.9",
        })
        with urllib.request.urlopen(pedido, timeout=espera) as r:
            tipo = r.headers.get("Content-Type", "")
            if "html" not in tipo and "text" not in tipo:
                return None, "no es html"
            crudo = r.read(TAMANO_MAXIMO)
            codificacion = "utf-8"
            coincide = re.search(r"charset=([\w\-]+)", tipo)
            if coincide:
                codificacion = coincide.group(1)
            try:
                return crudo.decode(codificacion, "replace"), r.geturl()
            except LookupError:
                return crudo.decode("utf-8", "replace"), r.geturl()
    except urllib.error.HTTPError as err:
        return None, "HTTP %s" % err.code
    except Exception as err:
        return None, str(err)[:60]


def permiso_robots(dominio):
    lector = urllib.robotparser.RobotFileParser()
    for esquema in ("https", "http"):
        try:
            lector.set_url("%s://%s/robots.txt" % (esquema, dominio))
            lector.read()
            return lector
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# 1. Localizar la pagina oficial
# ---------------------------------------------------------------------------

def dominios_candidatos(razon_social, es_publica):
    """Dominios plausibles construidos a partir de la razon social."""
    todas = palabras_de(razon_social)
    distintivas = tokens_distintivos(razon_social)
    if not distintivas:
        distintivas = [p for p in todas if len(p) >= 4]
    if not distintivas:
        return []

    raices = []

    def agregar(raiz):
        raiz = re.sub(r"[^a-z0-9]", "", raiz)
        if 4 <= len(raiz) <= 30 and raiz not in raices:
            raices.append(raiz)

    agregar("".join(distintivas))
    if len(distintivas) >= 2:
        agregar("".join(distintivas[:2]))
    agregar(distintivas[0])
    # Nombre completo tal cual, por si el generico es parte de la marca
    # ("clinicabucaramanga", "hospitalsanjuan").
    agregar("".join(todas[:3]))
    if len(todas) >= 3:
        agregar("".join(p[0] for p in todas if p not in ("de", "del", "la", "y")))

    tlds = TLDS_PUBLICA if es_publica else TLDS
    candidatos = []
    for raiz in raices:
        for tld in tlds:
            candidato = raiz + tld
            if candidato not in candidatos:
                candidatos.append(candidato)
    return candidatos[:DOMINIOS_A_PROBAR]


def confirma_pertenencia(html, titulo, razon_social, nit):
    """
    Solo aceptamos un dominio si la pagina misma dice que es de esa entidad:
    trae el NIT, o al menos dos palabras distintivas del nombre.
    """
    texto = sin_tildes(html)
    if nit and len(nit) >= 8:
        # El NIT puede venir con puntos, guiones o el digito de verificacion.
        solo_numeros = re.sub(r"\D", "", texto)
        if nit in solo_numeros:
            return True, "la pagina publica el NIT %s" % nit
    distintivas = tokens_distintivos(razon_social)
    encontradas = [t for t in distintivas if t in texto or t in sin_tildes(titulo)]
    if len(encontradas) >= 2:
        return True, "la pagina nombra a la entidad (%s)" % ", ".join(encontradas[:3])
    if len(encontradas) == 1 and len(encontradas[0]) >= 8:
        return True, "la pagina nombra a la entidad (%s)" % encontradas[0]
    return False, ""


def resuelve(dominio):
    """Evita pedir paginas de dominios que ni siquiera existen."""
    try:
        socket.setdefaulttimeout(5)
        socket.getaddrinfo(dominio, None)
        return True
    except Exception:
        return False


def localizar_sitio(sociedad, deducir):
    """
    Devuelve (dominio, url, metodo, motivo). Cadena vacia si no se pudo
    establecer con certeza.
    """
    razon = sociedad.get("razon_social", "")
    nit = re.sub(r"\D", "", sociedad.get("nit", "") or "")
    es_publica = "public" in sin_tildes(sociedad.get("naturaleza", ""))

    # a) El dominio del correo del REPS. Es la pista mas fuerte.
    dominio = dominio_de(sociedad.get("sitio_web") or sociedad.get("email", ""))
    if dominio:
        return dominio, "", "dominio del correo reportado al REPS", ""

    if not deducir:
        return "", "", "no encontrado", "el correo del REPS no da dominio propio"

    # b) Dominios deducidos del nombre, confirmados por la propia pagina.
    probados = 0
    for candidato in dominios_candidatos(razon, es_publica):
        if not resuelve(candidato):
            continue
        probados += 1
        for esquema in ("https", "http"):
            html, final = bajar("%s://%s/" % (esquema, candidato), TIEMPO_SONDEO)
            if not html:
                continue
            extractor = Extractor()
            try:
                extractor.feed(html)
            except Exception:
                pass
            confirmado, motivo = confirma_pertenencia(
                html, extractor.titulo, razon, nit
            )
            if confirmado:
                return candidato, final, "dominio deducido y confirmado", motivo
            break  # respondio pero no es de esta entidad: no insistir en http

    motivo = ("se probaron %d dominios parecidos y ninguno confirmo ser de la "
              "entidad" % probados) if probados else \
             "no existe ningun dominio parecido al nombre"
    return "", "", "no encontrado", motivo


# ---------------------------------------------------------------------------
# 2 y 3. Recorrer el sitio: correos y convocatorias
# ---------------------------------------------------------------------------

def revisar_sitio(dominio, url_inicial=""):
    resultado = {
        "dominio": dominio,
        "sitio": "",
        "correos": {},
        "pagina_contratacion": "",
        "convocatorias": [],      # (titulo, url)
        "paginas_leidas": 0,
        "nota": "",
    }

    robots = permiso_robots(dominio)

    def permitido(url):
        if robots is None:
            return True
        try:
            return robots.can_fetch(AGENTE, url)
        except Exception:
            return True

    inicio = None
    detalle = ""
    arranques = [url_inicial] if url_inicial else []
    arranques += ["https://%s/" % dominio, "http://%s/" % dominio]
    for url in arranques:
        if not permitido(url):
            resultado["nota"] = "robots.txt no permite leer el sitio"
            return resultado
        html, detalle = bajar(url)
        if html:
            inicio = (html, detalle)
            break
        time.sleep(PAUSA_MISMO_SITIO)
    if not inicio:
        resultado["nota"] = "el sitio no respondio (%s)" % (detalle or "sin detalle")
        return resultado

    html, url_final = inicio
    resultado["sitio"] = url_final
    resultado["paginas_leidas"] = 1
    base = urllib.parse.urlparse(url_final)
    host = base.netloc

    def cosechar(html_pagina, url_pagina, es_contratacion):
        extractor = Extractor()
        try:
            extractor.feed(html_pagina)
        except Exception:
            pass
        for correo in PATRON_CORREO.findall(html_pagina):
            correo = correo.strip(".,;:").lower()
            if any(b in correo for b in BASURA) or len(correo) > 90:
                continue
            resultado["correos"].setdefault(correo, url_pagina)
        # En las paginas de contratacion, cada enlace que suene a proceso
        # concreto se guarda como convocatoria, con su titulo y su URL.
        if es_contratacion:
            for href, etiqueta in extractor.enlaces:
                pista = sin_tildes(etiqueta)
                if not pista or len(pista) < 8:
                    continue
                if any(p in pista for p in PALABRAS_CONVOCATORIA):
                    absoluta = urllib.parse.urljoin(url_pagina, href)
                    titulo = etiqueta.strip()[:180]
                    if not any(t == titulo for t, _ in resultado["convocatorias"]):
                        resultado["convocatorias"].append((titulo, absoluta))
        return extractor

    extractor = cosechar(html, url_final, False)

    candidatas = []
    vistas = {url_final.rstrip("/")}
    for href, etiqueta in extractor.enlaces:
        absoluta = urllib.parse.urljoin(url_final, href)
        partes = urllib.parse.urlparse(absoluta)
        if partes.netloc != host or partes.scheme not in ("http", "https"):
            continue
        limpia = absoluta.split("#")[0].rstrip("/")
        if limpia in vistas:
            continue
        pista = sin_tildes(etiqueta + " " + partes.path)
        if any(p in pista for p in PALABRAS_ENLACE):
            candidatas.append(limpia)
            vistas.add(limpia)

    for ruta in RUTAS_CANDIDATAS:
        posible = "%s://%s%s" % (base.scheme, host, ruta)
        if posible.rstrip("/") not in vistas:
            candidatas.append(posible)
            vistas.add(posible.rstrip("/"))

    for url in candidatas:
        if resultado["paginas_leidas"] >= PAGINAS_POR_SITIO:
            break
        if not permitido(url):
            continue
        time.sleep(PAUSA_MISMO_SITIO)
        html_pagina, final_pagina = bajar(url)
        if not html_pagina:
            continue
        resultado["paginas_leidas"] += 1
        pista = sin_tildes(urllib.parse.urlparse(final_pagina).path)
        es_contratacion = any(p in pista for p in (
            "contratacion", "proveedor", "convocatoria", "licitacion",
            "transparencia", "invitacion",
        ))
        cosechar(html_pagina, final_pagina, es_contratacion)
        if es_contratacion and not resultado["pagina_contratacion"]:
            resultado["pagina_contratacion"] = final_pagina

    if not resultado["correos"] and not resultado["nota"]:
        resultado["nota"] = "el sitio respondio pero no publica correos"
    return resultado


# ---------------------------------------------------------------------------
# SECOP: por NIT y por nombre de entidad
# ---------------------------------------------------------------------------

def _json(url):
    pedido = urllib.request.Request(url, headers={
        "Accept": "application/json", "User-Agent": AGENTE,
    })
    if APP_TOKEN:
        pedido.add_header("X-App-Token", APP_TOKEN)
    with urllib.request.urlopen(pedido, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


CAMPOS_SECOP = {
    "nit_entidad": ["nit_entidad", "nitentidad", "nit_de_la_entidad",
                    "nit_entidad_contratante", "nit"],
    "entidad": ["entidad", "nombre_entidad", "nombre_de_la_entidad",
                "nombreentidad"],
    "objeto": ["objeto_del_contrato", "objeto_a_contratar", "objeto",
               "descripcion_del_proceso", "detalle_del_objeto_a_contratar",
               "descripcion_del_procedimiento"],
    "valor": ["valor_del_contrato", "valor_total_adjudicacion", "valor",
              "precio_base", "cuantia_proceso"],
    "fecha": ["fecha_de_firma", "fecha_de_firma_del_contrato",
              "fecha_de_publicacion_del", "fecha_de_publicacion",
              "fecha_de_inicio_del_contrato"],
    "enlace": ["urlproceso", "url_del_proceso", "enlace",
               "referencia_del_contrato", "referencia_del_proceso"],
    "estado": ["estado_contrato", "estado_del_proceso", "estado"],
    "proveedor": ["proveedor_adjudicado", "nombre_del_proveedor",
                  "razon_social_del_contratista", "contratista"],
}


def preparar_secop():
    listos = []
    for nombre, identificador in DATASETS_SECOP:
        url = "https://www.datos.gov.co/resource/%s.json?$limit=1" % identificador
        try:
            muestra = _json(url)
        except Exception:
            continue
        if not muestra:
            continue
        columnas = {clave_columna(k): k for k in muestra[0].keys()}
        mapa = {}
        for campo, candidatos in CAMPOS_SECOP.items():
            for cand in candidatos:
                clave = clave_columna(cand)
                if clave in columnas:
                    mapa[campo] = columnas[clave]
                    break
        if "objeto" in mapa and ("nit_entidad" in mapa or "entidad" in mapa):
            listos.append((nombre, identificador, mapa))
    return listos


def _consulta(identificador, donde, limite=25):
    url = "https://www.datos.gov.co/resource/%s.json?$limit=%d&$where=%s" % (
        identificador, limite, urllib.parse.quote(donde)
    )
    try:
        return _json(url)
    except Exception:
        return []


def _fila_a_hallazgo(fila, mapa, fuente, metodo):
    return {
        "fuente": fuente,
        "metodo": metodo,
        "entidad": fila.get(mapa.get("entidad", ""), "") or "",
        "nit_en_secop": fila.get(mapa.get("nit_entidad", ""), "") or "",
        "objeto": (fila.get(mapa.get("objeto", ""), "") or "")[:250],
        "valor": fila.get(mapa.get("valor", ""), "") or "",
        "fecha": (str(fila.get(mapa.get("fecha", ""), "") or ""))[:10],
        "estado": fila.get(mapa.get("estado", ""), "") or "",
        "proveedor": fila.get(mapa.get("proveedor", ""), "") or "",
        "enlace": fila.get(mapa.get("enlace", ""), "") or "",
    }


def buscar_en_secop(sociedad, datasets, por_nombre=True):
    """
    Busca la entidad en SECOP por NIT y, si por_nombre, tambien por el nombre
    de la entidad. Devuelve (hallazgos_juridicos, total_procesos, notas).
    """
    nit = re.sub(r"\D", "", sociedad.get("nit", "") or "")
    razon = sociedad.get("razon_social", "")
    distintivas = tokens_distintivos(razon)[:3] if por_nombre else []

    hallazgos = []
    total = 0
    notas = []
    vistos = set()

    filtro_juridico = lambda col: "(%s)" % " OR ".join(
        "lower(%s) like '%%%s%%'" % (col, p) for p in PALABRAS_JURIDICAS
    )

    for fuente, identificador, mapa in datasets:
        col_nit = mapa.get("nit_entidad")
        col_ent = mapa.get("entidad")
        col_obj = mapa["objeto"]

        # --- a) Por NIT exacto ---------------------------------------------
        if nit and col_nit:
            # El NIT puede estar guardado con o sin digito de verificacion.
            variantes = {nit}
            if len(nit) == 10:
                variantes.add(nit[:9])
            condicion_nit = "(%s)" % " OR ".join(
                "%s = '%s'" % (col_nit, v) for v in sorted(variantes)
            )
            for fila in _consulta(identificador,
                                  "%s AND %s" % (condicion_nit, filtro_juridico(col_obj))):
                h = _fila_a_hallazgo(fila, mapa, fuente, "NIT exacto")
                clave = (h["objeto"][:80], h["fecha"])
                if clave not in vistos:
                    vistos.add(clave)
                    hallazgos.append(h)
            conteo = _consulta(identificador, condicion_nit, 1)
            if conteo:
                total += 1  # marca que la entidad si contrata por SECOP

        # --- b) Por nombre de entidad --------------------------------------
        # Muchas IPS quedan en SECOP con otro NIT (el del grupo, o mal
        # digitado). Buscar por nombre las rescata; queda marcado para que
        # usted lo confirme antes de usarlo.
        if col_ent and len(distintivas) >= 1:
            condicion_nombre = " AND ".join(
                "lower(%s) like '%%%s%%'" % (col_ent, t) for t in distintivas
            )
            for fila in _consulta(
                identificador,
                "%s AND %s" % (condicion_nombre, filtro_juridico(col_obj))
            ):
                h = _fila_a_hallazgo(fila, mapa, fuente, "por nombre")
                nit_secop = re.sub(r"\D", "", h["nit_en_secop"])
                if nit and nit_secop and nit_secop[:9] == nit[:9]:
                    h["metodo"] = "NIT exacto"
                elif nit and nit_secop:
                    h["metodo"] = "por nombre (NIT en SECOP: %s -- confirmar)" % \
                                  h["nit_en_secop"]
                clave = (h["objeto"][:80], h["fecha"])
                if clave not in vistos:
                    vistos.add(clave)
                    hallazgos.append(h)
            if not total and _consulta(identificador, condicion_nombre, 1):
                total += 1

    if hallazgos and not nit:
        notas.append("la sociedad no tiene NIT en la base: todo se busco por nombre")

    hallazgos.sort(key=lambda h: (h["metodo"] != "NIT exacto", h["fecha"]),
                   reverse=False)
    hallazgos.sort(key=lambda h: h["fecha"], reverse=True)
    return hallazgos, total, "; ".join(notas)


# ---------------------------------------------------------------------------
# Armado de las casillas nuevas
# ---------------------------------------------------------------------------

def decidir_canal(sociedad, sitio, web, secop, total_secop, nota_secop):
    correos = dict(web.get("correos", {})) if web else {}
    if sociedad.get("email"):
        correos.setdefault(sociedad["email"].lower(), "REPS")

    por_categoria = {}
    for correo, origen in correos.items():
        por_categoria.setdefault(categoria_de(correo), []).append((correo, origen))

    canal = canal_tipo = canal_categoria = ""
    for categoria in PRIORIDAD_CANAL:
        if categoria in por_categoria:
            canal, origen_canal = por_categoria[categoria][0]
            canal_categoria = categoria
            canal_tipo = categoria
            if origen_canal == "REPS":
                canal_tipo += " (dato del REPS, sin confirmar en la web)"
            break

    contratacion = [c for c, _ in por_categoria.get("Contratacion / proveedores", [])]
    juridica = [c for c, _ in por_categoria.get("Juridica", [])]
    judiciales = [c for c, _ in por_categoria.get("Notificaciones judiciales", [])]
    otros = [c for c in correos
             if c not in contratacion + juridica + judiciales
             and c != (sociedad.get("email") or "").lower()]

    por_nit = [h for h in secop if h["metodo"] == "NIT exacto"]
    por_nombre = [h for h in secop if h["metodo"] != "NIT exacto"]
    if secop:
        ultima = secop[0]
        resumen_secop = "%d hallazgo(s) juridico(s) (%d por NIT, %d por nombre). " \
                        "Ultimo %s: %s" % (
            len(secop), len(por_nit), len(por_nombre),
            ultima["fecha"] or "sin fecha", ultima["objeto"][:130]
        )
    else:
        resumen_secop = ""

    convocatorias_web = (web or {}).get("convocatorias", [])
    resumen_web = ""
    if convocatorias_web:
        resumen_web = "%d publicada(s) en su web. Ej.: %s" % (
            len(convocatorias_web), convocatorias_web[0][0][:120]
        )

    evidencia = []
    if sitio.get("url") or sitio.get("dominio"):
        evidencia.append(sitio.get("url") or ("http://" + sitio["dominio"]))
    if web and web.get("pagina_contratacion"):
        evidencia.append(web["pagina_contratacion"])
    for correo, origen in list(correos.items())[:6]:
        if origen and origen != "REPS":
            evidencia.append("%s <- %s" % (correo, origen))
    for _, url in convocatorias_web[:2]:
        evidencia.append(url)
    if secop and secop[0].get("enlace"):
        evidencia.append(str(secop[0]["enlace"]))

    if canal:
        casilla = "%s (%s)" % (canal, canal_categoria)
    elif web and web.get("pagina_contratacion"):
        casilla = "Radicar por %s" % web["pagina_contratacion"]
        canal_tipo = "Formulario / pagina web"
    elif sociedad.get("telefono"):
        casilla = ("Sin correo publicado; llamar al %s y pedir el correo de "
                   "contratacion" % sociedad["telefono"])
        canal_tipo = "Telefono"
    else:
        casilla = ""
        canal_tipo = "No encontrado"

    notas = [n for n in [(web or {}).get("nota", ""), sitio.get("motivo", ""),
                         nota_secop] if n]

    return {
        "canal_propuesta": casilla,
        "canal_tipo": canal_tipo,
        "correos_contratacion": ", ".join(sorted(set(contratacion))),
        "correos_juridica": ", ".join(sorted(set(juridica + judiciales))),
        "otros_correos": ", ".join(sorted(set(otros))[:8]),
        # La URL real con la que respondio el sitio, no una armada a mano.
        "sitio_oficial": (web or {}).get("sitio") or sitio.get("url") or (
            ("http://" + sitio["dominio"]) if sitio.get("dominio") else ""),
        "como_se_hallo_el_sitio": sitio.get("metodo", ""),
        "pagina_contratacion": (web or {}).get("pagina_contratacion", ""),
        "convocatorias_secop": resumen_secop,
        "convocatorias_web": resumen_web,
        "procesos_secop": len(secop),
        "hallazgos_por_nit": len(por_nit),
        "hallazgos_por_nombre": len(por_nombre),
        "contrata_por_secop": "Si" if (total_secop or secop) else "",
        "evidencia": " | ".join(evidencia[:10]),
        "nota_busqueda": "; ".join(notas),
        "fecha_enriquecimiento": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------

COLUMNAS_NUEVAS = [
    ("canal_propuesta", "CANAL PARA ENVIAR PROPUESTA"),
    ("canal_tipo", "Tipo de canal"),
    ("sitio_oficial", "Pagina oficial"),
    ("como_se_hallo_el_sitio", "Como se hallo la pagina"),
    ("correos_contratacion", "Correos de contratacion / proveedores"),
    ("correos_juridica", "Correos juridica / notificaciones judiciales"),
    ("convocatorias_secop", "Convocatorias juridicas en SECOP"),
    ("convocatorias_web", "Convocatorias en su propia web"),
    ("contrata_por_secop", "Contrata por SECOP"),
    ("procesos_secop", "N. hallazgos SECOP"),
    ("hallazgos_por_nit", "  de esos, por NIT"),
    ("hallazgos_por_nombre", "  de esos, por nombre (confirmar)"),
    ("pagina_contratacion", "Pagina de contratacion"),
    ("otros_correos", "Otros correos publicados"),
    ("evidencia", "Evidencia (de donde salio cada dato)"),
    ("nota_busqueda", "Nota de la busqueda"),
    ("fecha_enriquecimiento", "Fecha de la busqueda"),
]


def leer_sociedades(ruta_db, solo_segmento, limite):
    if not os.path.isfile(ruta_db):
        raise SystemExit(
            "\nNo encuentro la base: %s\n\n"
            "Corra primero:  python ips_area_metropolitana.py\n" % ruta_db
        )
    con = sqlite3.connect(ruta_db)
    con.row_factory = sqlite3.Row
    filas = [dict(f) for f in con.execute(
        "SELECT * FROM prestadores ORDER BY puntaje DESC, sedes DESC, razon_social"
    )]
    con.close()
    if solo_segmento:
        letra = solo_segmento.strip().upper()[:1]
        filas = [f for f in filas if f.get("segmento", "").startswith(letra)]
    if limite:
        filas = filas[:limite]
    return filas


def guardar(ruta_db, enriquecidos, convocatorias):
    con = sqlite3.connect(ruta_db)
    cur = con.cursor()
    cur.execute("DROP TABLE IF EXISTS enriquecimiento")
    cur.execute("CREATE TABLE enriquecimiento (clave TEXT PRIMARY KEY, %s)" %
                ", ".join("%s TEXT" % c for c, _ in COLUMNAS_NUEVAS))
    campos = ["clave"] + [c for c, _ in COLUMNAS_NUEVAS]
    cur.executemany(
        "INSERT INTO enriquecimiento VALUES (%s)" % ",".join("?" * len(campos)),
        [tuple(str(r.get(c, "")) for c in campos) for r in enriquecidos],
    )

    cur.execute("DROP TABLE IF EXISTS convocatorias")
    cur.execute("""
        CREATE TABLE convocatorias (
            clave TEXT, razon_social TEXT, nit TEXT, origen TEXT, metodo TEXT,
            entidad_en_fuente TEXT, objeto TEXT, fecha TEXT, valor TEXT,
            estado TEXT, proveedor TEXT, enlace TEXT
        )""")
    cur.executemany(
        "INSERT INTO convocatorias VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        [(c["clave"], c["razon_social"], c["nit"], c["origen"], c["metodo"],
          c["entidad_en_fuente"], c["objeto"], c["fecha"], c["valor"],
          c["estado"], c["proveedor"], c["enlace"]) for c in convocatorias],
    )

    cur.execute("DROP VIEW IF EXISTS v_objetivos_contacto")
    cur.execute("""
        CREATE VIEW v_objetivos_contacto AS
        SELECT p.razon_social, p.nit, p.municipios, p.sedes, p.naturaleza,
               p.segmento, p.puntaje,
               e.canal_propuesta, e.canal_tipo, e.sitio_oficial,
               e.correos_contratacion, e.correos_juridica,
               e.convocatorias_secop, e.convocatorias_web, e.contrata_por_secop
        FROM prestadores p
        LEFT JOIN enriquecimiento e ON e.clave = p.clave
        ORDER BY p.puntaje DESC, p.razon_social""")
    cur.execute("CREATE INDEX ix_convocatorias_clave ON convocatorias(clave)")
    con.commit()
    con.close()


COLUMNAS_CONVOCATORIA = [
    ("razon_social", "IPS"), ("nit", "NIT"), ("origen", "Origen"),
    ("metodo", "Como se encontro"), ("entidad_en_fuente", "Entidad en la fuente"),
    ("objeto", "Objeto / titulo"), ("fecha", "Fecha"), ("valor", "Valor"),
    ("estado", "Estado"), ("proveedor", "Contratista"), ("enlace", "Enlace"),
]


def exportar(carpeta, sociedades, enriquecidos, convocatorias, resumen):
    indice = {r["clave"]: r for r in enriquecidos}
    filas = []
    for sociedad in sociedades:
        fila = dict(sociedad)
        fila.update(indice.get(sociedad["clave"], {}))
        filas.append(fila)

    columnas = [
        ("razon_social", "Razon social"), ("nit", "NIT"),
        ("municipios", "Municipios"), ("sedes", "Sedes"),
        ("naturaleza", "Naturaleza"), ("nivel", "Nivel"),
        ("segmento", "Segmento"),
    ] + COLUMNAS_NUEVAS + [
        ("telefono", "Telefono REPS"), ("email", "Correo REPS"),
        ("direccion", "Direccion"), ("representante", "Representante"),
    ]

    ruta_csv = os.path.join(carpeta, "objetivos_ips_con_canal.csv")
    with open(ruta_csv, "w", encoding="utf-8-sig", newline="") as f:
        escritor = csv.writer(f, delimiter=";")
        escritor.writerow([t for _, t in columnas])
        for fila in filas:
            escritor.writerow([fila.get(c, "") for c, _ in columnas])

    ruta_xlsx = os.path.join(carpeta, "IPS_Area_Metropolitana_CON_CANAL.xlsx")
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  (sin openpyxl no se genera el Excel: pip install openpyxl)")
        return ruta_csv, ""

    libro = Workbook()
    blanco = Font(bold=True, color="FFFFFF")
    azul = PatternFill("solid", fgColor="1F4E78")
    naranja = PatternFill("solid", fgColor="C55A11")

    def encabezar(hoja, columnas_hoja, resaltar=()):
        hoja.append([t for _, t in columnas_hoja])
        for i, celda in enumerate(hoja[1], start=1):
            celda.font = blanco
            celda.fill = naranja if columnas_hoja[i - 1][0] in resaltar else azul
            celda.alignment = Alignment(vertical="center", wrap_text=True)

    def anchos(hoja, columnas_hoja, registros):
        for i, (campo, tit) in enumerate(columnas_hoja, start=1):
            ancho = max(len(tit) + 2, 14)
            for registro in registros[:300]:
                ancho = max(ancho, min(len(str(registro.get(campo, ""))) + 2, 60))
            hoja.column_dimensions[get_column_letter(i)].width = ancho

    hoja = libro.active
    hoja.title = "Objetivos con canal"
    encabezar(hoja, columnas, resaltar=("canal_propuesta",))
    for fila in filas:
        hoja.append([fila.get(c, "") for c, _ in columnas])
    hoja.freeze_panes = "B2"
    hoja.auto_filter.ref = hoja.dimensions
    anchos(hoja, columnas, filas)

    hoja_conv = libro.create_sheet("Convocatorias")
    encabezar(hoja_conv, COLUMNAS_CONVOCATORIA)
    for convocatoria in convocatorias:
        hoja_conv.append([convocatoria.get(c, "") for c, _ in COLUMNAS_CONVOCATORIA])
    hoja_conv.freeze_panes = "A2"
    if convocatorias:
        hoja_conv.auto_filter.ref = hoja_conv.dimensions
    anchos(hoja_conv, COLUMNAS_CONVOCATORIA, convocatorias)

    hoja_resumen = libro.create_sheet("Resumen busqueda")
    encabezar(hoja_resumen, [("a", "Indicador"), ("b", "Valor")])
    for clave, valor in resumen.items():
        hoja_resumen.append([clave, valor])
    hoja_resumen.column_dimensions["A"].width = 54
    hoja_resumen.column_dimensions["B"].width = 62

    libro.save(ruta_xlsx)
    return ruta_csv, ruta_xlsx


# ---------------------------------------------------------------------------
# Programa
# ---------------------------------------------------------------------------

def main():
    global PAUSA_MISMO_SITIO

    parser = argparse.ArgumentParser(
        description="Busca la pagina oficial, el canal de contratacion y las "
                    "convocatorias de cada IPS."
    )
    parser.add_argument("--db", default=BASE_POR_DEFECTO)
    parser.add_argument("--salida", default=CARPETA)
    parser.add_argument("--limite", type=int, default=0,
                        help="Procesar solo las primeras N (para probar)")
    parser.add_argument("--solo-segmento", help="Solo el segmento A, B o C")
    parser.add_argument("--sin-web", action="store_true",
                        help="No visitar los sitios de las IPS")
    parser.add_argument("--sin-secop", action="store_true",
                        help="No consultar SECOP")
    parser.add_argument("--sin-deducir-dominio", action="store_true",
                        help="No intentar adivinar el dominio de las que no "
                             "tienen correo propio en el REPS")
    parser.add_argument("--solo-nit", action="store_true",
                        help="En SECOP, buscar solo por NIT (sin busqueda por "
                             "nombre de entidad)")
    parser.add_argument("--pausa", type=float, default=PAUSA_MISMO_SITIO)
    args = parser.parse_args()

    PAUSA_MISMO_SITIO = max(args.pausa, 0.2)

    sociedades = leer_sociedades(args.db, args.solo_segmento, args.limite)
    print("=" * 68)
    print(" Pagina oficial, canal de contratacion y convocatorias")
    print(" Sociedades a revisar: %d" % len(sociedades))
    print("=" * 68)

    datasets = []
    if not args.sin_secop:
        print("Preparando SECOP...")
        try:
            datasets = preparar_secop()
        except Exception as err:
            print("  no se pudo: %s" % str(err)[:70])
        if datasets:
            print("  datasets disponibles: %s" % ", ".join(d[0] for d in datasets))
        else:
            print("  SECOP no respondio; esas columnas quedaran vacias.")

    # --- Paso 1: localizar la pagina oficial de cada una -------------------
    print("\nLocalizando paginas oficiales...")
    sitios = {}
    deducidos = 0
    with ThreadPoolExecutor(max_workers=SITIOS_EN_PARALELO) as pool:
        futuros = {
            pool.submit(localizar_sitio, s, not args.sin_deducir_dominio): s["clave"]
            for s in sociedades
        }
        hechos = 0
        for futuro in as_completed(futuros):
            clave = futuros[futuro]
            hechos += 1
            try:
                dominio, url, metodo, motivo = futuro.result()
            except Exception as err:
                dominio, url, metodo, motivo = "", "", "no encontrado", str(err)[:60]
            sitios[clave] = {"dominio": dominio, "url": url,
                             "metodo": metodo, "motivo": motivo}
            if metodo == "dominio deducido y confirmado":
                deducidos += 1
            if hechos % 20 == 0 or hechos == len(sociedades):
                print("  %d/%d" % (hechos, len(sociedades)))
    con_sitio = sum(1 for s in sitios.values() if s["dominio"])
    print("  con pagina: %d  (de esas, deducidas y confirmadas: %d)"
          % (con_sitio, deducidos))

    # --- Paso 2: recorrer cada sitio una sola vez --------------------------
    resultados_web = {}
    if not args.sin_web:
        dominios = {}
        for sociedad in sociedades:
            sitio = sitios.get(sociedad["clave"], {})
            if sitio.get("dominio"):
                dominios.setdefault(sitio["dominio"], sitio.get("url", ""))
        if dominios:
            print("\nRevisando %d sitios web..." % len(dominios))
            with ThreadPoolExecutor(max_workers=SITIOS_EN_PARALELO) as pool:
                futuros = {pool.submit(revisar_sitio, d, u): d
                           for d, u in dominios.items()}
                hechos = 0
                for futuro in as_completed(futuros):
                    dominio = futuros[futuro]
                    hechos += 1
                    try:
                        resultados_web[dominio] = futuro.result()
                    except Exception as err:
                        resultados_web[dominio] = {
                            "correos": {}, "convocatorias": [],
                            "nota": "error: %s" % str(err)[:60],
                        }
                    r = resultados_web[dominio]
                    print("  [%d/%d] %-34s %d correo(s), %d convocatoria(s)" % (
                        hechos, len(dominios), dominio[:34],
                        len(r.get("correos", {})), len(r.get("convocatorias", []))))

    # --- Paso 3: SECOP y consolidacion ------------------------------------
    enriquecidos = []
    convocatorias = []
    con_secop = 0
    print("\nConsolidando...")
    for i, sociedad in enumerate(sociedades, start=1):
        sitio = sitios.get(sociedad["clave"], {})
        web = resultados_web.get(sitio.get("dominio", ""), {})
        secop, total_secop, nota_secop = ([], 0, "")
        if datasets:
            try:
                secop, total_secop, nota_secop = buscar_en_secop(
                    sociedad, datasets, por_nombre=not args.solo_nit)
            except Exception as err:
                nota_secop = "SECOP fallo: %s" % str(err)[:50]
            if secop:
                con_secop += 1

        for h in secop:
            convocatorias.append({
                "clave": sociedad["clave"], "razon_social": sociedad["razon_social"],
                "nit": sociedad.get("nit", ""), "origen": h["fuente"],
                "metodo": h["metodo"], "entidad_en_fuente": h["entidad"],
                "objeto": h["objeto"], "fecha": h["fecha"], "valor": str(h["valor"]),
                "estado": h["estado"], "proveedor": h["proveedor"],
                "enlace": str(h["enlace"]),
            })
        for titulo, url in (web or {}).get("convocatorias", []):
            convocatorias.append({
                "clave": sociedad["clave"], "razon_social": sociedad["razon_social"],
                "nit": sociedad.get("nit", ""), "origen": "Pagina de la entidad",
                "metodo": "publicada en su web", "entidad_en_fuente": "",
                "objeto": titulo, "fecha": "", "valor": "", "estado": "",
                "proveedor": "", "enlace": url,
            })

        registro = decidir_canal(sociedad, sitio, web, secop, total_secop, nota_secop)
        registro["clave"] = sociedad["clave"]
        enriquecidos.append(registro)
        if i % 25 == 0:
            print("  %d/%d" % (i, len(sociedades)))

    con_canal = sum(1 for r in enriquecidos if r["canal_propuesta"])
    resumen = {
        "Sociedades revisadas": len(sociedades),
        "Con pagina oficial localizada": con_sitio,
        "  de esas, deducidas del nombre y confirmadas": deducidos,
        "Con canal de propuesta identificado": con_canal,
        "Sin canal (requieren gestion manual)": len(sociedades) - con_canal,
        "Con correo de contratacion / proveedores":
            sum(1 for r in enriquecidos if r["correos_contratacion"]),
        "Con correo juridico o de notificaciones judiciales":
            sum(1 for r in enriquecidos if r["correos_juridica"]),
        "Con hallazgos juridicos en SECOP": con_secop,
        "Convocatorias listadas (SECOP + web)": len(convocatorias),
        "  de esas, halladas por nombre (confirmar)":
            sum(1 for c in convocatorias if "nombre" in c["metodo"]),
        "Sitios web recorridos": len(resultados_web),
        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Fuentes": "Pagina oficial del prestador + SECOP (datos.gov.co) + REPS",
        "Advertencia": "Los hallazgos marcados 'por nombre' pueden ser de otra "
                       "entidad con nombre parecido. Confirmelos antes de usarlos.",
    }

    os.makedirs(args.salida, exist_ok=True)
    guardar(args.db, enriquecidos, convocatorias)
    ruta_csv, ruta_xlsx = exportar(args.salida, sociedades, enriquecidos,
                                   convocatorias, resumen)

    print("\n" + "-" * 68)
    for clave, valor in resumen.items():
        if clave not in ("Fuentes", "Advertencia"):
            print(" %-52s %s" % (clave + ":", valor))
    print("-" * 68)
    if ruta_xlsx:
        print(" Excel : %s" % ruta_xlsx)
    print(" CSV   : %s" % ruta_csv)
    print(" Base  : %s" % args.db)
    print("         tablas enriquecimiento y convocatorias,")
    print("         vista v_objetivos_contacto")
    print("-" * 68)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(130)
