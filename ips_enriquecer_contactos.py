"""
Enriquece la base de IPS con el canal real para enviarles la propuesta.

Que hace
--------
Toma la base que genera ips_area_metropolitana.py y, para cada sociedad,
sale a buscar por donde se le puede radicar de verdad una propuesta de
servicios juridicos. Agrega una casilla nueva -- "Canal para enviar
propuesta" -- mas las columnas de respaldo que la sustentan.

De donde saca la informacion (solo fuentes verificables)
--------------------------------------------------------
1) La pagina web del propio prestador. El dominio sale del correo que el
   prestador reporto al REPS (gerencia@clinicasm.com -> clinicasm.com), o
   de la columna "sitio_web" si usted la llena a mano en el Excel. De ahi
   se leen las paginas de contacto, contratacion, proveedores,
   convocatorias, transparencia y notificaciones judiciales, y se extraen
   los correos que aparezcan publicados.

2) SECOP (Colombia Compra Eficiente), via Datos Abiertos. Se busca por el
   NIT de la IPS si tiene procesos o contratos cuyo objeto mencione
   servicios juridicos, abogados, asesoria legal o representacion
   judicial. Eso responde directo a "convocatoria para abogados": si una
   IPS ya contrato abogados por SECOP, ahi esta el precedente, el valor y
   el enlace del proceso.

Cada dato queda con la URL de donde salio, en la columna "Evidencia". Lo
que no se encuentre queda VACIO -- nunca se rellena con suposiciones.

Buen comportamiento en la red
-----------------------------
Esto visita cientos de sitios de terceros. Por eso: respeta robots.txt,
se identifica con un User-Agent propio, hace pausa entre peticiones al
mismo sitio, limita cuantas paginas lee por dominio, y nunca manda
formularios ni entra a zonas privadas. Solo lee paginas publicas.

Como se usa
-----------
  python ips_enriquecer_contactos.py
  python ips_enriquecer_contactos.py --limite 25          (prueba corta)
  python ips_enriquecer_contactos.py --sin-secop
  python ips_enriquecer_contactos.py --solo-segmento A

o doble clic en ips_enriquecer_contactos.bat
"""

import argparse
import csv
import json
import os
import re
import sqlite3
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
    "EXTRAPROCESAL-contacto/1.0 (busca el canal de contratacion publicado%s)"
    % ((" ; contacto: " + CONTACTO_RESPONSABLE) if CONTACTO_RESPONSABLE else "")
)

TIEMPO_ESPERA = 20          # segundos por peticion
PAUSA_MISMO_SITIO = 1.0     # segundos entre peticiones al mismo dominio
PAGINAS_POR_SITIO = 8       # tope de paginas leidas por dominio
SITIOS_EN_PARALELO = 6      # dominios distintos a la vez
TAMANO_MAXIMO = 2_000_000   # no descargar paginas gigantes

# Datasets de SECOP en datos.gov.co. Se prueban en orden y se usan los que
# respondan; si el Ministerio cambia los identificadores, el programa avisa
# en vez de fallar.
DATASETS_SECOP = [
    ("SECOP II - Contratos electronicos", "jbjy-vk9h"),
    ("SECOP II - Procesos de contratacion", "p6dx-8zbt"),
    ("SECOP I", "xvdy-vvsk"),
]
URL_CATALOGO = "https://www.datos.gov.co/api/catalog/v1"
APP_TOKEN = os.environ.get("DATOS_GOV_APP_TOKEN", "").strip()

# Palabras que delatan una contratacion de servicios juridicos.
PALABRAS_JURIDICAS = (
    "juridic", "abogad", "asesoria legal", "servicios legales",
    "representacion judicial", "defensa judicial", "apoyo legal",
    "asesoria juridica", "cobro juridico", "cobro prejuridico",
    "conciliacion", "litigio",
)

# Paginas que vale la pena mirar dentro de cada sitio, con su peso.
RUTAS_CANDIDATAS = [
    "/contratacion", "/proveedores", "/convocatorias", "/contacto",
    "/transparencia", "/contactenos", "/proveedor", "/licitaciones",
    "/trabaje-con-nosotros", "/quienes-somos",
]

PALABRAS_ENLACE = (
    "contratacion", "contrataci", "proveedor", "convocatoria", "licitacion",
    "invitacion publica", "contacto", "contactenos", "transparencia",
    "notificaciones judiciales", "juridica", "trabaje con nosotros",
    "compras", "pqr",
)

# Clasificacion de un correo segun su parte local. El orden importa: manda
# la primera categoria que coincida.
CATEGORIAS = [
    ("Notificaciones judiciales", (
        "notificacionesjudiciales", "notificacionjudicial", "judicial",
        "notificaciones",
    )),
    ("Contratacion / proveedores", (
        "contratacion", "contratos", "contratacion", "proveedores", "proveedor",
        "compras", "licitacion", "licitaciones", "convocatoria", "convocatorias",
        "adquisiciones", "suministros",
    )),
    ("Juridica", (
        "juridica", "juridico", "legal", "abogado", "secretariageneral",
        "secretaria.general",
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

# Orden en que se prefiere un canal para radicar la propuesta.
PRIORIDAD_CANAL = [
    "Contratacion / proveedores",
    "Juridica",
    "Gerencia / direccion",
    "Notificaciones judiciales",
    "General",
    "PQRS / atencion al usuario",
    "Talento humano",
]

# Correos que no sirven de nada (plantillas, proveedores del sitio web...).
BASURA = (
    "example.com", "dominio.com", "correo.com", "tucorreo", "email.com",
    "sentry.io", "wixpress.com", "godaddy.com", "sentry-next",
    "@2x.png", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js",
)


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


def dominio_de(texto):
    """Saca el dominio de un correo o de una URL."""
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
    # Los correos de dominios gratuitos no dan pista del sitio del prestador.
    gratuitos = ("gmail.com", "hotmail.com", "outlook.com", "yahoo.com",
                 "hotmail.es", "outlook.es", "yahoo.es", "live.com",
                 "icloud.com", "protonmail.com", "msn.com")
    if dominio in gratuitos:
        return ""
    if "." not in dominio:
        return ""
    return dominio


def categoria_de(correo):
    local = sin_tildes(correo.split("@", 1)[0]).replace("-", "").replace("_", "")
    for nombre, marcas in CATEGORIAS:
        if any(m.replace(".", "") in local for m in marcas):
            return nombre
    return "General"


# ---------------------------------------------------------------------------
# Lectura de paginas
# ---------------------------------------------------------------------------

class Extractor(HTMLParser):
    """Saca de una pagina los enlaces y el texto plano."""

    def __init__(self):
        HTMLParser.__init__(self)
        self.enlaces = []      # (href, texto del enlace)
        self.textos = []
        self._href = None
        self._buffer = []
        self._ignorar = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._ignorar += 1
        elif tag == "a":
            self._href = dict(attrs).get("href")
            self._buffer = []

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._ignorar:
            self._ignorar -= 1
        elif tag == "a":
            if self._href:
                self.enlaces.append((self._href, " ".join(self._buffer).strip()))
            self._href = None
            self._buffer = []

    def handle_data(self, data):
        if self._ignorar:
            return
        self.textos.append(data)
        if self._href is not None:
            self._buffer.append(data)

    def texto(self):
        return " ".join(" ".join(self.textos).split())


PATRON_CORREO = re.compile(
    r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
)


def bajar(url):
    """Descarga una pagina. Devuelve (html, url_final) o (None, motivo)."""
    try:
        pedido = urllib.request.Request(url, headers={
            "User-Agent": AGENTE,
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "es-CO,es;q=0.9",
        })
        with urllib.request.urlopen(pedido, timeout=TIEMPO_ESPERA) as r:
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
    """Lee robots.txt. Ante la duda, se permite (asi se comporta un navegador)."""
    lector = urllib.robotparser.RobotFileParser()
    for esquema in ("https", "http"):
        try:
            lector.set_url("%s://%s/robots.txt" % (esquema, dominio))
            lector.read()
            return lector
        except Exception:
            continue
    return None


def revisar_sitio(dominio):
    """
    Recorre las paginas publicas utiles de un dominio y devuelve los correos
    encontrados, con la URL donde aparecio cada uno.
    """
    resultado = {
        "dominio": dominio,
        "sitio": "",
        "correos": {},          # correo -> url donde se vio
        "pagina_contratacion": "",
        "menciona_convocatoria": False,
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

    # 1) La portada, para descubrir el sitio y sus enlaces.
    inicio = None
    for esquema in ("https", "http"):
        url = "%s://%s/" % (esquema, dominio)
        if not permitido(url):
            resultado["nota"] = "robots.txt no permite leer el sitio"
            return resultado
        html, final = bajar(url)
        if html:
            inicio = (html, final)
            break
        time.sleep(PAUSA_MISMO_SITIO)
    if not inicio:
        resultado["nota"] = "el sitio no respondio (%s)" % (final or "sin detalle")
        return resultado

    html, url_final = inicio
    resultado["sitio"] = url_final
    resultado["paginas_leidas"] = 1
    base = urllib.parse.urlparse(url_final)
    host = base.netloc

    def cosechar(html_pagina, url_pagina):
        extractor = Extractor()
        try:
            extractor.feed(html_pagina)
        except Exception:
            pass
        texto = extractor.texto()
        for correo in PATRON_CORREO.findall(html_pagina):
            correo = correo.strip(".,;:").lower()
            if any(b in correo for b in BASURA):
                continue
            if len(correo) > 90:
                continue
            resultado["correos"].setdefault(correo, url_pagina)
        if any(p in sin_tildes(texto) for p in
               ("convocatoria", "invitacion publica", "licitacion",
                "proceso de contratacion")):
            resultado["menciona_convocatoria"] = True
        return extractor

    extractor = cosechar(html, url_final)

    # 2) Enlaces internos que suenan a contratacion, proveedores o contacto.
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

    # 3) Rutas tipicas, por si el menu es de JavaScript y no dejo enlaces.
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
        cosechar(html_pagina, final_pagina)
        pista = sin_tildes(urllib.parse.urlparse(final_pagina).path)
        if not resultado["pagina_contratacion"] and any(
            p in pista for p in ("contratacion", "proveedor", "convocatoria",
                                 "licitacion", "transparencia")
        ):
            resultado["pagina_contratacion"] = final_pagina

    if not resultado["correos"] and not resultado["nota"]:
        resultado["nota"] = "el sitio respondio pero no publica correos"
    return resultado


# ---------------------------------------------------------------------------
# SECOP
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
    "enlace": ["urlproceso", "url_del_proceso", "urlproceso_url", "enlace",
               "referencia_del_contrato", "id_contrato", "referencia_del_proceso"],
    "estado": ["estado_contrato", "estado_del_proceso", "estado"],
}


def preparar_secop():
    """Averigua que datasets de SECOP responden y como se llaman sus columnas."""
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
        if "nit_entidad" in mapa and "objeto" in mapa:
            listos.append((nombre, identificador, mapa))
    return listos


def buscar_en_secop(nit, datasets):
    """Procesos o contratos de esa entidad cuyo objeto sea juridico."""
    if not nit or not datasets:
        return []
    hallazgos = []
    for nombre, identificador, mapa in datasets:
        condiciones = " OR ".join(
            "lower(%s) like '%%%s%%'" % (mapa["objeto"], palabra)
            for palabra in PALABRAS_JURIDICAS
        )
        donde = "%s = '%s' AND (%s)" % (mapa["nit_entidad"], nit, condiciones)
        url = "https://www.datos.gov.co/resource/%s.json?$limit=20&$where=%s" % (
            identificador, urllib.parse.quote(donde)
        )
        try:
            filas = _json(url)
        except Exception:
            continue
        for fila in filas:
            hallazgos.append({
                "fuente": nombre,
                "objeto": (fila.get(mapa.get("objeto", ""), "") or "")[:220],
                "valor": fila.get(mapa.get("valor", ""), ""),
                "fecha": (fila.get(mapa.get("fecha", ""), "") or "")[:10],
                "estado": fila.get(mapa.get("estado", ""), ""),
                "enlace": fila.get(mapa.get("enlace", ""), ""),
            })
    hallazgos.sort(key=lambda h: h["fecha"], reverse=True)
    return hallazgos


# ---------------------------------------------------------------------------
# Armado de la casilla nueva
# ---------------------------------------------------------------------------

def decidir_canal(sociedad, web, secop):
    """
    Devuelve el diccionario con las columnas nuevas. La casilla principal es
    "canal_propuesta": una sola linea con lo mejor que se encontro.
    """
    correos = dict(web.get("correos", {})) if web else {}

    # El correo del REPS tambien cuenta, pero de ultimo: suele ser el
    # administrativo de habilitacion.
    if sociedad.get("email"):
        correos.setdefault(sociedad["email"].lower(), "REPS")

    por_categoria = {}
    for correo, origen in correos.items():
        categoria = categoria_de(correo)
        por_categoria.setdefault(categoria, []).append((correo, origen))

    canal = ""
    canal_tipo = ""
    canal_categoria = ""
    for categoria in PRIORIDAD_CANAL:
        if categoria in por_categoria:
            canal, origen_canal = por_categoria[categoria][0]
            canal_categoria = categoria
            canal_tipo = categoria
            # Se deja claro si el correo salio de la web del prestador o si es
            # el que reporto al REPS, que suele ser el de habilitacion.
            if origen_canal == "REPS":
                canal_tipo += " (dato del REPS, sin confirmar en la web)"
            break

    contratacion = [c for c, _ in por_categoria.get("Contratacion / proveedores", [])]
    juridica = [c for c, _ in por_categoria.get("Juridica", [])]
    judiciales = [c for c, _ in por_categoria.get("Notificaciones judiciales", [])]
    otros = [c for c in correos
             if c not in contratacion + juridica + judiciales
             and c != sociedad.get("email", "").lower()]

    convocatorias = ""
    if secop:
        ultima = secop[0]
        convocatorias = "%d proceso(s) juridico(s) en SECOP; ultimo %s: %s" % (
            len(secop), ultima["fecha"] or "sin fecha", ultima["objeto"][:120]
        )

    evidencia = []
    if web and web.get("sitio"):
        evidencia.append(web["sitio"])
    if web and web.get("pagina_contratacion"):
        evidencia.append(web["pagina_contratacion"])
    for correo, origen in list(correos.items())[:6]:
        if origen and origen != "REPS":
            evidencia.append("%s <- %s" % (correo, origen))
    if secop and secop[0].get("enlace"):
        evidencia.append(str(secop[0]["enlace"]))

    # La casilla que pidio el usuario: una linea accionable.
    if canal:
        casilla = "%s (%s)" % (canal, canal_categoria)
    elif web and web.get("pagina_contratacion"):
        casilla = "Radicar por el formulario de %s" % web["pagina_contratacion"]
        canal_tipo = "Formulario web"
    elif sociedad.get("telefono"):
        casilla = "Sin correo publicado; llamar al %s y pedir el correo de contratacion" % \
                  sociedad["telefono"]
        canal_tipo = "Telefono"
    else:
        casilla = ""
        canal_tipo = "No encontrado"

    return {
        "canal_propuesta": casilla,
        "canal_tipo": canal_tipo,
        "correos_contratacion": ", ".join(sorted(set(contratacion))),
        "correos_juridica": ", ".join(sorted(set(juridica + judiciales))),
        "otros_correos": ", ".join(sorted(set(otros))[:8]),
        "sitio_web": (web or {}).get("sitio", ""),
        "pagina_contratacion": (web or {}).get("pagina_contratacion", ""),
        "convocatorias_juridicas": convocatorias,
        "procesos_secop": len(secop),
        "evidencia": " | ".join(evidencia[:8]),
        "nota_busqueda": (web or {}).get("nota", "no se pudo deducir el dominio"),
        "fecha_enriquecimiento": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ---------------------------------------------------------------------------
# Base de datos
# ---------------------------------------------------------------------------

COLUMNAS_NUEVAS = [
    ("canal_propuesta", "CANAL PARA ENVIAR PROPUESTA"),
    ("canal_tipo", "Tipo de canal"),
    ("correos_contratacion", "Correos de contratacion / proveedores"),
    ("correos_juridica", "Correos juridica / notificaciones judiciales"),
    ("convocatorias_juridicas", "Convocatorias juridicas (SECOP)"),
    ("procesos_secop", "N. procesos juridicos en SECOP"),
    ("pagina_contratacion", "Pagina de contratacion"),
    ("sitio_web", "Sitio web"),
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
    consulta = ("SELECT * FROM prestadores "
                "ORDER BY puntaje DESC, sedes DESC, razon_social")
    filas = [dict(f) for f in con.execute(consulta)]
    con.close()
    if solo_segmento:
        letra = solo_segmento.strip().upper()[:1]
        filas = [f for f in filas if f.get("segmento", "").startswith(letra)]
    if limite:
        filas = filas[:limite]
    return filas


def guardar(ruta_db, enriquecidos):
    con = sqlite3.connect(ruta_db)
    cur = con.cursor()
    cur.execute("DROP TABLE IF EXISTS enriquecimiento")
    columnas = ", ".join("%s TEXT" % campo for campo, _ in COLUMNAS_NUEVAS)
    cur.execute("CREATE TABLE enriquecimiento (clave TEXT PRIMARY KEY, %s)" % columnas)
    campos = ["clave"] + [c for c, _ in COLUMNAS_NUEVAS]
    cur.executemany(
        "INSERT INTO enriquecimiento VALUES (%s)" % ",".join("?" * len(campos)),
        [tuple(str(r.get(c, "")) for c in campos) for r in enriquecidos],
    )
    cur.execute("DROP VIEW IF EXISTS v_objetivos_contacto")
    cur.execute("""
        CREATE VIEW v_objetivos_contacto AS
        SELECT p.razon_social, p.nit, p.municipios, p.sedes, p.naturaleza,
               p.segmento, p.puntaje,
               e.canal_propuesta, e.canal_tipo, e.correos_contratacion,
               e.correos_juridica, e.convocatorias_juridicas, e.sitio_web
        FROM prestadores p
        LEFT JOIN enriquecimiento e ON e.clave = p.clave
        ORDER BY p.puntaje DESC, p.razon_social""")
    con.commit()
    con.close()


def exportar(carpeta, sociedades, enriquecidos, resumen):
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
    hoja = libro.active
    hoja.title = "Objetivos con canal"
    hoja.append([t for _, t in columnas])
    encabezado = Font(bold=True, color="FFFFFF")
    for i, celda in enumerate(hoja[1], start=1):
        celda.font = encabezado
        # La casilla nueva va resaltada en otro color, para que salte a la vista.
        celda.fill = PatternFill(
            "solid",
            fgColor="C55A11" if columnas[i - 1][0] == "canal_propuesta" else "1F4E78",
        )
        celda.alignment = Alignment(vertical="center", wrap_text=True)
    for fila in filas:
        hoja.append([fila.get(c, "") for c, _ in columnas])
    hoja.freeze_panes = "B2"
    hoja.auto_filter.ref = hoja.dimensions
    for i, (campo, tit) in enumerate(columnas, start=1):
        ancho = max(len(tit) + 2, 14)
        for fila in filas[:300]:
            ancho = max(ancho, min(len(str(fila.get(campo, ""))) + 2, 60))
        hoja.column_dimensions[get_column_letter(i)].width = ancho

    hoja_resumen = libro.create_sheet("Resumen busqueda")
    hoja_resumen.append(["Indicador", "Valor"])
    for celda in hoja_resumen[1]:
        celda.font = encabezado
        celda.fill = PatternFill("solid", fgColor="1F4E78")
    for clave, valor in resumen.items():
        hoja_resumen.append([clave, valor])
    hoja_resumen.column_dimensions["A"].width = 52
    hoja_resumen.column_dimensions["B"].width = 60

    libro.save(ruta_xlsx)
    return ruta_csv, ruta_xlsx


# ---------------------------------------------------------------------------
# Programa
# ---------------------------------------------------------------------------

def main():
    global PAUSA_MISMO_SITIO

    parser = argparse.ArgumentParser(
        description="Busca por donde radicarle la propuesta a cada IPS."
    )
    parser.add_argument("--db", default=BASE_POR_DEFECTO,
                        help="Base generada por ips_area_metropolitana.py")
    parser.add_argument("--salida", default=CARPETA, help="Carpeta de salida")
    parser.add_argument("--limite", type=int, default=0,
                        help="Procesar solo las primeras N (para probar)")
    parser.add_argument("--solo-segmento", help="Procesar solo el segmento A, B o C")
    parser.add_argument("--sin-web", action="store_true",
                        help="No visitar los sitios de las IPS")
    parser.add_argument("--sin-secop", action="store_true",
                        help="No consultar SECOP")
    parser.add_argument("--pausa", type=float, default=PAUSA_MISMO_SITIO,
                        help="Segundos de pausa entre paginas del mismo sitio")
    args = parser.parse_args()

    PAUSA_MISMO_SITIO = max(args.pausa, 0.2)

    sociedades = leer_sociedades(args.db, args.solo_segmento, args.limite)
    print("=" * 66)
    print(" Buscando el canal para radicar propuestas")
    print(" Sociedades a revisar: %d" % len(sociedades))
    print("=" * 66)

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
            print("  SECOP no respondio; esa columna quedara vacia.")

    # Un dominio puede repetirse entre sedes/sociedades: se visita una sola vez.
    dominios = {}
    for sociedad in sociedades:
        dominio = dominio_de(sociedad.get("sitio_web") or sociedad.get("email", ""))
        sociedad["_dominio"] = dominio
        if dominio:
            dominios.setdefault(dominio, [])
        if dominio:
            dominios[dominio].append(sociedad["clave"])

    resultados_web = {}
    if not args.sin_web and dominios:
        print("\nRevisando %d sitios web..." % len(dominios))
        with ThreadPoolExecutor(max_workers=SITIOS_EN_PARALELO) as pool:
            futuros = {pool.submit(revisar_sitio, d): d for d in dominios}
            hechos = 0
            for futuro in as_completed(futuros):
                dominio = futuros[futuro]
                hechos += 1
                try:
                    resultados_web[dominio] = futuro.result()
                except Exception as err:
                    resultados_web[dominio] = {
                        "correos": {}, "nota": "error: %s" % str(err)[:60]
                    }
                encontrados = len(resultados_web[dominio].get("correos", {}))
                print("  [%d/%d] %-38s %d correo(s)" % (
                    hechos, len(dominios), dominio[:38], encontrados))

    enriquecidos = []
    con_secop = 0
    for i, sociedad in enumerate(sociedades, start=1):
        web = resultados_web.get(sociedad["_dominio"], {})
        secop = []
        if datasets and sociedad.get("nit"):
            try:
                secop = buscar_en_secop(sociedad["nit"], datasets)
            except Exception:
                secop = []
            if secop:
                con_secop += 1
        registro = decidir_canal(sociedad, web, secop)
        registro["clave"] = sociedad["clave"]
        enriquecidos.append(registro)
        if i % 25 == 0:
            print("  consolidadas %d/%d" % (i, len(sociedades)))

    con_canal = sum(1 for r in enriquecidos if r["canal_propuesta"])
    con_contratacion = sum(1 for r in enriquecidos if r["correos_contratacion"])
    con_juridica = sum(1 for r in enriquecidos if r["correos_juridica"])

    resumen = {
        "Sociedades revisadas": len(sociedades),
        "Con canal identificado": con_canal,
        "Sin canal (requieren gestion manual)": len(sociedades) - con_canal,
        "Con correo de contratacion / proveedores": con_contratacion,
        "Con correo juridico o de notificaciones judiciales": con_juridica,
        "Con convocatorias juridicas en SECOP": con_secop,
        "Sitios web revisados": len(resultados_web),
        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Fuentes": "Sitio web publico del prestador + SECOP (datos.gov.co) + REPS",
        "Advertencia": "Cada dato trae la URL de donde salio en la columna "
                       "Evidencia. Verifiquelo antes de radicar.",
    }

    os.makedirs(args.salida, exist_ok=True)
    guardar(args.db, enriquecidos)
    ruta_csv, ruta_xlsx = exportar(args.salida, sociedades, enriquecidos, resumen)

    print("\n" + "-" * 66)
    for clave, valor in resumen.items():
        if clave not in ("Fuentes", "Advertencia"):
            print(" %-52s %s" % (clave + ":", valor))
    print("-" * 66)
    if ruta_xlsx:
        print(" Excel : %s" % ruta_xlsx)
    print(" CSV   : %s" % ruta_csv)
    print(" Base  : %s  (tabla enriquecimiento, vista v_objetivos_contacto)" % args.db)
    print("-" * 66)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(130)
