"""
Base de datos de IPS (sociedades) del Area Metropolitana de Bucaramanga.

Para que sirve
--------------
Arma un listado depurado y accionable de las SOCIEDADES habilitadas como IPS
(Instituciones Prestadoras de Servicios de Salud) en el Area Metropolitana de
Bucaramanga -- Bucaramanga, Floridablanca, Giron y Piedecuesta -- con sus datos
de contacto, para poder ofrecerles el portafolio de servicios juridicos.

De donde sale la informacion
----------------------------
De la fuente oficial: el REPS (Registro Especial de Prestadores de Servicios de
Salud) del Ministerio de Salud y Proteccion Social. Se puede leer de dos formas:

  1) Por internet (opcion por defecto), desde el portal de Datos Abiertos:
     https://www.datos.gov.co/resource/c36g-9fc2.json
     ("Registro Especial de Prestadores y Sedes de Servicios de Salud")

  2) Desde un archivo que usted mismo descargue (opcion --archivo), util si la
     red de la oficina bloquea datos.gov.co o si prefiere trabajar con el corte
     oficial del REPS:
     https://prestadores.minsalud.gov.co/habilitacion/consultas/habilitados_reps.aspx
     Sirve un .csv, un .xlsx o un .json exportado de ahi.

NO se inventa ni se completa a mano ningun dato: todo lo que sale en la base
viene de la fuente. Si el REPS no trae correo o telefono de un prestador, la
celda queda vacia y el prestador queda marcado como "sin contacto" para que
usted decida como abordarlo.

Que produce
-----------
Dentro de la carpeta datos_ips/ :

  * ips_area_metropolitana.db     -> base de datos SQLite (tablas prestadores
                                     y sedes, mas la vista v_objetivos).
  * IPS_Area_Metropolitana.xlsx   -> libro de Excel con las hojas Objetivos,
                                     Sedes, Resumen y Fuente.
  * objetivos_ips.csv             -> el mismo listado plano, para importar a un
                                     CRM o para una combinacion de
                                     correspondencia.

Como se usa
-----------
  python ips_area_metropolitana.py
  python ips_area_metropolitana.py --archivo "C:\\ruta\\Prestadores.xlsx"
  python ips_area_metropolitana.py --incluir-naturales
  python ips_area_metropolitana.py --municipios "BUCARAMANGA,LEBRIJA"

o simplemente doble clic en ips_area_metropolitana.bat
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
from datetime import datetime

# ---------------------------------------------------------------------------
# Configuracion
# ---------------------------------------------------------------------------

DATASET_ID = "c36g-9fc2"
URL_SOCRATA = "https://www.datos.gov.co/resource/%s.json" % DATASET_ID
URL_DATASET_HUMANO = (
    "https://www.datos.gov.co/Salud-y-Protecci-n-Social/"
    "Registro-Especial-de-Prestadores-y-Sedes-de-Servic/%s" % DATASET_ID
)
URL_REPS_OFICIAL = (
    "https://prestadores.minsalud.gov.co/habilitacion/consultas/habilitados_reps.aspx"
)

# Token opcional de datos.gov.co. No es obligatorio; sin el, la descarga
# igual funciona pero con un limite de peticiones mas bajo.
APP_TOKEN = os.environ.get("DATOS_GOV_APP_TOKEN", "").strip()

PAGINA = 50000          # filas por peticion a la API
REINTENTOS = 4          # reintentos por pagina ante fallos de red
CARPETA_SALIDA = "datos_ips"

# Los cuatro municipios que por ley integran el Area Metropolitana de
# Bucaramanga, con su codigo DANE.
MUNICIPIOS_AMB = {
    "68001": "BUCARAMANGA",
    "68276": "FLORIDABLANCA",
    "68307": "GIRON",
    "68547": "PIEDECUESTA",
}

DEPARTAMENTO = "SANTANDER"

# Clases de prestador que NO son IPS y por lo tanto se descartan.
CLASES_NO_IPS = (
    "profesional independiente",
    "transporte especial",
    "objeto social diferente",
)


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------

def sin_tildes(texto):
    """Deja el texto en minusculas, sin tildes y sin espacios sobrantes."""
    if texto is None:
        return ""
    texto = str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def clave_columna(nombre):
    """Normaliza el nombre de una columna a solo letras y numeros."""
    return re.sub(r"[^a-z0-9]", "", sin_tildes(nombre))


def limpiar(valor):
    """Convierte cualquier celda a un texto limpio ('' si esta vacia)."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip()
    if texto.lower() in ("nan", "none", "null", "-"):
        return ""
    return " ".join(texto.split())


def solo_digitos(texto):
    return re.sub(r"\D", "", texto or "")


# Siglas que deben quedar en mayusculas dentro de una razon social.
SIGLAS = {
    "IPS", "IPSI", "EPS", "ESE", "ESS", "SAS", "SA", "SAU", "LTDA", "EU",
    "SCA", "SCS", "SCI", "UT", "UCI", "ORL", "BIC", "CTA", "ONG", "SOM",
    "SAAM", "SEM", "CAJ", "AC",
}

# Palabras que en un nombre propio van en minuscula, salvo al principio.
MENORES = {"de", "del", "la", "las", "los", "y", "e", "en", "a", "el", "para", "por"}


def titulo(texto):
    """Razon social en formato legible (evita los TODO EN MAYUSCULAS del REPS)."""
    texto = limpiar(texto)
    if not texto:
        return ""
    if not (texto.isupper() or texto.islower()):
        return texto
    palabras = []
    for i, palabra in enumerate(texto.split()):
        limpia = re.sub(r"[.,]", "", palabra).upper()
        if limpia in SIGLAS:
            palabras.append(limpia)
        elif i > 0 and palabra.lower() in MENORES:
            palabras.append(palabra.lower())
        else:
            palabras.append(palabra.lower().capitalize())
    return " ".join(palabras)


# ---------------------------------------------------------------------------
# Deteccion de columnas
#
# El REPS se publica con nombres de columna distintos segun el corte y segun si
# viene de datos.gov.co o del portal del Ministerio. En vez de amarrarnos a un
# nombre fijo, buscamos cada dato por una lista de candidatos.
# ---------------------------------------------------------------------------

CANDIDATOS = {
    "codigo_habilitacion": [
        "codigohabilitacion", "codigo_habilitacion", "codigohabilitacionsede",
        "codigoprestador", "habilitacion", "codigo",
    ],
    "tipo_identificacion": [
        "tipoidentificacion", "tipodocumento", "tipoid", "clasedocumento",
        "tipoidentificacionprestador", "tidonombre", "tido",
    ],
    "nit": [
        "nit", "numerodocumento", "numeroidentificacion", "identificacion",
        "documento", "nitsnit", "nits",
    ],
    "razon_social": [
        "nombreprestador", "razonsocial", "nombre", "prestador",
        "nombrecomercial", "nombreinstitucion", "nombresede",
    ],
    "clase_prestador": [
        "claseprestador", "clase", "clprnombre", "tipoprestador",
        "clasificacionprestador",
    ],
    "tipo_persona": [
        "clasepersona", "tipopersona", "naturalezapersona", "personanatural",
        "personajuridica",
    ],
    "naturaleza": [
        "naturalezajuridica", "naturaleza", "najunombre", "sector",
    ],
    "caracter": ["caracter", "caracterterritorial", "carnombre"],
    "nivel": ["nivel", "nivelatencion", "nivelcomplejidad", "complejidad"],
    "departamento": [
        "departamento", "nombredepartamento", "depanombre", "deptonombre", "depa",
    ],
    "codigo_municipio": [
        "codigomunicipio", "municipiocodigo", "codmunicipio", "munucodigo",
        "codigodane", "divipola",
    ],
    "municipio": [
        "municipio", "nombremunicipio", "muninombre", "municipiosede", "ciudad",
    ],
    "direccion": ["direccion", "direccionsede", "dirección"],
    "telefono": ["telefono", "telefonos", "telefonosede", "telefonocontacto", "tel"],
    "email": [
        "email", "correo", "correoelectronico", "emailsede", "correosede",
        "emailcontacto",
    ],
    "representante": [
        "gerente", "representantelegal", "representante", "nombregerente",
        "responsable",
    ],
    "sede_nombre": ["nombresede", "sede", "nombredelasede", "sedenombre"],
    "sede_numero": ["numerosede", "consecutivosede", "codigosede", "sedenumero"],
    "sede_principal": ["sedeprincipal", "esprincipal", "principal"],
    "fecha_apertura": ["fechaapertura", "fechainicio", "fecharadicacion"],
    "fecha_cierre": ["fechacierre", "fechavencimiento", "fechafin"],
    "estado": ["estado", "habilitado", "estadoprestador", "estadosede"],
}


def mapear_columnas(encabezados):
    """Devuelve {campo_canonico: nombre_real_de_la_columna}."""
    normalizados = {}
    for h in encabezados:
        normalizados.setdefault(clave_columna(h), h)

    mapa = {}
    usadas = set()

    # Primera vuelta: coincidencia exacta (es la mas confiable).
    for campo, candidatos in CANDIDATOS.items():
        for cand in candidatos:
            clave = clave_columna(cand)
            if clave in normalizados and normalizados[clave] not in usadas:
                mapa[campo] = normalizados[clave]
                usadas.add(normalizados[clave])
                break

    # Segunda vuelta: coincidencia parcial, solo para lo que quedo sin mapear.
    for campo, candidatos in CANDIDATOS.items():
        if campo in mapa:
            continue
        for cand in candidatos:
            clave = clave_columna(cand)
            for norm, original in normalizados.items():
                if original in usadas:
                    continue
                if clave and (clave in norm or norm in clave):
                    mapa[campo] = original
                    usadas.add(original)
                    break
            if campo in mapa:
                break
    return mapa


def leer(fila, mapa, campo):
    columna = mapa.get(campo)
    if not columna:
        return ""
    return limpiar(fila.get(columna))


# ---------------------------------------------------------------------------
# Lectura de la fuente
# ---------------------------------------------------------------------------

def _peticion(url):
    pedido = urllib.request.Request(url, headers={
        "Accept": "application/json",
        "User-Agent": "EXTRAPROCESAL/ips_area_metropolitana",
    })
    if APP_TOKEN:
        pedido.add_header("X-App-Token", APP_TOKEN)
    with urllib.request.urlopen(pedido, timeout=120) as respuesta:
        return json.loads(respuesta.read().decode("utf-8"))


def descargar_de_datos_abiertos(municipios):
    """Baja el REPS de datos.gov.co, paginando hasta que no queden filas."""
    print("Fuente: Datos Abiertos Colombia (dataset %s)" % DATASET_ID)
    print("  %s" % URL_DATASET_HUMANO)

    # 1) Una fila de muestra, solo para ver como se llaman las columnas.
    try:
        muestra = _peticion(URL_SOCRATA + "?$limit=1")
    except urllib.error.HTTPError as err:
        raise SystemExit(_mensaje_red("HTTP %s al consultar la API" % err.code))
    except Exception as err:
        raise SystemExit(_mensaje_red(str(err)))

    if not muestra:
        raise SystemExit("La API respondio vacio. Intente mas tarde.")

    mapa = mapear_columnas(list(muestra[0].keys()))
    print("  Columnas detectadas: %d de %d posibles" % (len(mapa), len(CANDIDATOS)))

    # 2) Filtro del lado del servidor, para no bajar el pais entero.
    #    Si el servidor lo rechaza, bajamos todo y filtramos aca.
    filtro = ""
    col_dep = mapa.get("departamento")
    if col_dep:
        filtro = "&$where=" + urllib.parse.quote(
            "upper(%s) like '%%SANTANDER%%'" % col_dep
        )

    filas = _paginar(filtro)
    if filas is None:
        print("  El filtro por departamento no fue aceptado; se baja el registro")
        print("  completo y se filtra localmente (tarda un poco mas).")
        filas = _paginar("")
        if filas is None:
            raise SystemExit("No fue posible descargar el registro.")

    print("  Filas descargadas: %d" % len(filas))
    return filas, mapa, "datos.gov.co / %s" % DATASET_ID


def _paginar(filtro):
    filas = []
    offset = 0
    while True:
        url = "%s?$limit=%d&$offset=%d&$order=:id%s" % (
            URL_SOCRATA, PAGINA, offset, filtro
        )
        lote = None
        for intento in range(REINTENTOS):
            try:
                lote = _peticion(url)
                break
            except urllib.error.HTTPError as err:
                if err.code == 400:
                    return None           # el filtro no sirve para este dataset
                espera = 2 ** (intento + 1)
                print("    HTTP %s; reintento en %ss" % (err.code, espera))
                time.sleep(espera)
            except Exception as err:
                espera = 2 ** (intento + 1)
                print("    %s; reintento en %ss" % (err, espera))
                time.sleep(espera)
        if lote is None:
            raise SystemExit(_mensaje_red("se agotaron los reintentos"))
        if not lote:
            break
        filas.extend(lote)
        print("    %d filas..." % len(filas))
        if len(lote) < PAGINA:
            break
        offset += PAGINA
    return filas


def _mensaje_red(detalle):
    return (
        "\nNo se pudo leer datos.gov.co (%s).\n\n"
        "Si la red de la oficina bloquea ese portal, descargue el registro a mano\n"
        "y vuelva a correr el programa apuntando al archivo:\n\n"
        "  1. Entre a %s\n"
        "  2. Consulte el 'Registro Actual' de Santander y exportelo a Excel.\n"
        "  3. Corra:  python ips_area_metropolitana.py --archivo \"ruta\\archivo.xlsx\"\n"
        % (detalle, URL_REPS_OFICIAL)
    )


def leer_archivo(ruta):
    """Lee el REPS desde un .csv, .xlsx o .json ya descargado."""
    if not os.path.isfile(ruta):
        raise SystemExit("No existe el archivo: %s" % ruta)
    extension = os.path.splitext(ruta)[1].lower()
    print("Fuente: archivo local %s" % ruta)

    if extension in (".csv", ".txt"):
        filas = _leer_csv(ruta)
    elif extension in (".xlsx", ".xlsm"):
        filas = _leer_excel(ruta)
    elif extension == ".json":
        with open(ruta, "r", encoding="utf-8") as f:
            filas = json.load(f)
        if isinstance(filas, dict):
            for valor in filas.values():
                if isinstance(valor, list):
                    filas = valor
                    break
    else:
        raise SystemExit(
            "Formato no soportado (%s). Use .csv, .xlsx o .json." % extension
        )

    if not filas:
        raise SystemExit("El archivo no tiene filas de datos.")
    mapa = mapear_columnas(list(filas[0].keys()))
    print("  Filas leidas: %d" % len(filas))
    print("  Columnas detectadas: %d de %d posibles" % (len(mapa), len(CANDIDATOS)))
    return filas, mapa, os.path.basename(ruta)


def _leer_csv(ruta):
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, "r", encoding=codificacion, newline="") as f:
                inicio = f.read(8192)
                f.seek(0)
                try:
                    dialecto = csv.Sniffer().sniff(inicio, delimiters=";,|\t")
                except Exception:
                    dialecto = csv.excel
                    dialecto.delimiter = ";" if inicio.count(";") > inicio.count(",") else ","
                return list(csv.DictReader(f, dialect=dialecto))
        except UnicodeDecodeError:
            continue
    raise SystemExit("No se pudo leer el CSV (codificacion desconocida).")


def _leer_excel(ruta):
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise SystemExit("Falta openpyxl.  Instale con:  pip install openpyxl")
    libro = load_workbook(ruta, read_only=True, data_only=True)
    hoja = libro.active
    filas = []
    encabezados = None
    for fila in hoja.iter_rows(values_only=True):
        if encabezados is None:
            if fila and any(c is not None and str(c).strip() for c in fila):
                encabezados = [limpiar(c) or ("col%d" % i) for i, c in enumerate(fila)]
            continue
        if not any(c is not None and str(c).strip() for c in fila):
            continue
        filas.append(dict(zip(encabezados, fila)))
    libro.close()
    return filas


# ---------------------------------------------------------------------------
# Clasificacion de cada fila
# ---------------------------------------------------------------------------

def es_del_area(fila, mapa, municipios):
    """True si la sede queda en alguno de los municipios pedidos."""
    codigo = solo_digitos(leer(fila, mapa, "codigo_municipio"))
    if codigo:
        if len(codigo) == 4:          # a veces se pierde el cero inicial
            codigo = "0" + codigo
        if len(codigo) >= 5 and codigo[:5] in municipios:
            return True, municipios[codigo[:5]]

    nombre = sin_tildes(leer(fila, mapa, "municipio"))
    if nombre:
        for nombre_amb in municipios.values():
            if sin_tildes(nombre_amb) == nombre:
                return True, nombre_amb

    # Ultimo recurso: el codigo de habilitacion arranca con el codigo DANE.
    habilitacion = solo_digitos(leer(fila, mapa, "codigo_habilitacion"))
    if len(habilitacion) >= 5 and habilitacion[:5] in municipios:
        return True, municipios[habilitacion[:5]]

    return False, ""


def es_ips(fila, mapa):
    """True si la fila corresponde a una IPS (no a un profesional independiente)."""
    clase = sin_tildes(leer(fila, mapa, "clase_prestador"))
    if not clase:
        # Sin columna de clase no podemos descartar; se deja pasar y se marca.
        return True, "no informada"
    for excluida in CLASES_NO_IPS:
        if excluida in clase:
            return False, clase
    if "ips" in clase or "institucion prestadora" in clase:
        return True, clase
    return False, clase


def es_sociedad(fila, mapa):
    """
    Determina si el prestador es una persona juridica (una sociedad), que es lo
    que interesa para el portafolio. Devuelve (es_sociedad, criterio).
    """
    persona = sin_tildes(leer(fila, mapa, "tipo_persona"))
    if persona:
        if "juridic" in persona:
            return True, "clase de persona: juridica"
        if "natural" in persona:
            return False, "clase de persona: natural"

    tipo_doc = sin_tildes(leer(fila, mapa, "tipo_identificacion"))
    if tipo_doc:
        if "nit" in tipo_doc or tipo_doc in ("ni", "n i"):
            return True, "identificada con NIT"
        if any(x in tipo_doc for x in ("cedula", "cc", "ce", "pasaporte")):
            return False, "identificada con cedula"

    # Sin columna que lo diga, el NIT de las sociedades colombianas es de 9 o
    # 10 digitos y empieza por 8 o 9; las cedulas rara vez cumplen las dos.
    nit = solo_digitos(leer(fila, mapa, "nit"))
    if len(nit) in (9, 10) and nit[0] in ("8", "9"):
        return True, "NIT con formato de sociedad (%s...)" % nit[:3]

    razon = sin_tildes(leer(fila, mapa, "razon_social"))
    marcas = (" sas", " s a s", " ltda", " s a", " sa ", " e s e", " ese ",
              " eu", " scs", " s.a", "fundacion", "cooperativa", "asociacion",
              "corporacion", "clinica", "hospital", "centro medico", "ips ")
    if any(m in " %s " % razon for m in marcas):
        return True, "la razon social indica sociedad"

    return False, "sin evidencia de persona juridica"


# ---------------------------------------------------------------------------
# Consolidacion: de sedes a sociedades
# ---------------------------------------------------------------------------

def identidad(registro):
    """Clave con la que se agrupan las sedes de una misma sociedad."""
    if registro["nit"]:
        return "nit:" + registro["nit"]
    habilitacion = solo_digitos(registro["codigo_habilitacion"])
    if len(habilitacion) >= 10:
        return "hab:" + habilitacion[:10]
    return "nom:" + sin_tildes(registro["razon_social"])


def consolidar(filas, mapa, municipios, incluir_naturales):
    """Recorre las sedes, filtra y las agrupa por sociedad."""
    sociedades = {}
    sedes = []
    descartes = {
        "fuera del area": 0,
        "no es IPS": 0,
        "persona natural": 0,
        "sin razon social": 0,
    }

    for fila in filas:
        dentro, municipio = es_del_area(fila, mapa, municipios)
        if not dentro:
            descartes["fuera del area"] += 1
            continue

        ips, clase = es_ips(fila, mapa)
        if not ips:
            descartes["no es IPS"] += 1
            continue

        razon = leer(fila, mapa, "razon_social")
        if not razon:
            descartes["sin razon social"] += 1
            continue

        sociedad, criterio = es_sociedad(fila, mapa)
        if not sociedad and not incluir_naturales:
            descartes["persona natural"] += 1
            continue

        registro = {
            "codigo_habilitacion": leer(fila, mapa, "codigo_habilitacion"),
            "nit": solo_digitos(leer(fila, mapa, "nit")),
            "razon_social": titulo(razon),
            "razon_social_reps": razon,
            "clase_prestador": clase,
            "es_sociedad": "Si" if sociedad else "No",
            "criterio_sociedad": criterio,
            "naturaleza": leer(fila, mapa, "naturaleza"),
            "caracter": leer(fila, mapa, "caracter"),
            "nivel": leer(fila, mapa, "nivel"),
            "municipio": municipio,
            "departamento": leer(fila, mapa, "departamento") or DEPARTAMENTO,
            "direccion": leer(fila, mapa, "direccion"),
            "telefono": leer(fila, mapa, "telefono"),
            "email": leer(fila, mapa, "email").lower(),
            "representante": titulo(leer(fila, mapa, "representante")),
            "sede_nombre": leer(fila, mapa, "sede_nombre"),
            "sede_numero": leer(fila, mapa, "sede_numero"),
            "estado": leer(fila, mapa, "estado"),
            "fecha_apertura": leer(fila, mapa, "fecha_apertura"),
        }

        clave = identidad(registro)
        registro["clave"] = clave
        sedes.append(registro)

        actual = sociedades.get(clave)
        if actual is None:
            copia = dict(registro)
            copia["municipios"] = {municipio}
            copia["sedes"] = 1
            sociedades[clave] = copia
        else:
            actual["municipios"].add(municipio)
            actual["sedes"] += 1
            # Se completa lo que falte con lo que traiga otra sede.
            for campo in ("nit", "direccion", "telefono", "email",
                          "representante", "naturaleza", "nivel", "caracter"):
                if not actual[campo] and registro[campo]:
                    actual[campo] = registro[campo]

    for sociedad in sociedades.values():
        sociedad["municipios"] = ", ".join(sorted(sociedad["municipios"]))
        sociedad["puntaje"], sociedad["motivos"] = calcular_puntaje(sociedad)
        sociedad["segmento"] = segmento(sociedad["puntaje"])
        sociedad["contactable"] = "Si" if (sociedad["email"] or sociedad["telefono"]) else "No"

    return sociedades, sedes, descartes


# ---------------------------------------------------------------------------
# Priorizacion comercial
#
# Es una guia de orden de trabajo, no una verdad absoluta: solo ordena a quien
# llamar primero. Los criterios quedan escritos en la columna "motivos" para
# que cualquiera pueda revisarlos o discutirlos.
# ---------------------------------------------------------------------------

def calcular_puntaje(sociedad):
    puntaje = 0
    motivos = []

    sedes = sociedad["sedes"]
    if sedes > 1:
        suma = min(sedes, 6) * 5
        puntaje += suma
        motivos.append("%d sedes (+%d)" % (sedes, suma))

    naturaleza = sin_tildes(sociedad["naturaleza"])
    if "privad" in naturaleza:
        puntaje += 20
        motivos.append("privada, contrata abogado externo (+20)")
    elif "mixta" in naturaleza:
        puntaje += 12
        motivos.append("mixta (+12)")
    elif "public" in naturaleza or "estatal" in naturaleza:
        puntaje += 5
        motivos.append("publica, contrata por convocatoria (+5)")

    nivel = solo_digitos(sociedad["nivel"])
    romanos = {"iii": 3, "ii": 2, "i": 1}
    if not nivel:
        for romano, valor in romanos.items():
            if romano in sin_tildes(sociedad["nivel"]).split():
                nivel = str(valor)
                break
    if nivel == "3":
        puntaje += 18
        motivos.append("alta complejidad (+18)")
    elif nivel == "2":
        puntaje += 10
        motivos.append("mediana complejidad (+10)")

    if sociedad["email"]:
        puntaje += 12
        motivos.append("tiene correo (+12)")
    if sociedad["telefono"]:
        puntaje += 8
        motivos.append("tiene telefono (+8)")
    if sociedad["es_sociedad"] == "Si":
        puntaje += 10
        motivos.append("persona juridica confirmada (+10)")

    return min(puntaje, 100), "; ".join(motivos)


def segmento(puntaje):
    if puntaje >= 60:
        return "A - contactar primero"
    if puntaje >= 35:
        return "B - contactar despues"
    return "C - requiere completar datos"


# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------

COLUMNAS_OBJETIVOS = [
    ("razon_social", "Razon social"),
    ("nit", "NIT"),
    ("municipio", "Municipio sede"),
    ("municipios", "Municipios donde opera"),
    ("sedes", "Sedes"),
    ("naturaleza", "Naturaleza"),
    ("caracter", "Caracter"),
    ("nivel", "Nivel"),
    ("clase_prestador", "Clase de prestador"),
    ("direccion", "Direccion"),
    ("telefono", "Telefono"),
    ("email", "Correo"),
    ("representante", "Gerente / representante"),
    ("codigo_habilitacion", "Codigo habilitacion"),
    ("estado", "Estado en REPS"),
    ("es_sociedad", "Es sociedad"),
    ("criterio_sociedad", "Por que se clasifico asi"),
    ("contactable", "Contactable"),
    ("segmento", "Segmento"),
    ("puntaje", "Puntaje"),
    ("motivos", "Motivos del puntaje"),
]

COLUMNAS_SEDES = [
    ("razon_social", "Razon social"),
    ("nit", "NIT"),
    ("sede_nombre", "Sede"),
    ("sede_numero", "N. sede"),
    ("municipio", "Municipio"),
    ("direccion", "Direccion"),
    ("telefono", "Telefono"),
    ("email", "Correo"),
    ("codigo_habilitacion", "Codigo habilitacion"),
    ("estado", "Estado"),
    ("fecha_apertura", "Fecha apertura"),
]


def ordenadas(sociedades):
    return sorted(
        sociedades.values(),
        key=lambda s: (-s["puntaje"], -s["sedes"], sin_tildes(s["razon_social"])),
    )


def guardar_sqlite(ruta, lista, sedes, metadatos):
    if os.path.exists(ruta):
        os.remove(ruta)
    con = sqlite3.connect(ruta)
    cur = con.cursor()

    cur.execute("""
        CREATE TABLE prestadores (
            clave TEXT PRIMARY KEY,
            razon_social TEXT, razon_social_reps TEXT, nit TEXT,
            codigo_habilitacion TEXT, clase_prestador TEXT,
            es_sociedad TEXT, criterio_sociedad TEXT,
            naturaleza TEXT, caracter TEXT, nivel TEXT,
            municipio TEXT, municipios TEXT, departamento TEXT, sedes INTEGER,
            direccion TEXT, telefono TEXT, email TEXT, representante TEXT,
            estado TEXT, contactable TEXT, segmento TEXT,
            puntaje INTEGER, motivos TEXT
        )""")
    cur.execute("""
        CREATE TABLE sedes (
            clave TEXT, razon_social TEXT, nit TEXT,
            sede_nombre TEXT, sede_numero TEXT, municipio TEXT,
            direccion TEXT, telefono TEXT, email TEXT,
            codigo_habilitacion TEXT, estado TEXT, fecha_apertura TEXT
        )""")
    cur.execute("CREATE TABLE fuente (clave TEXT, valor TEXT)")

    campos_p = [
        "clave", "razon_social", "razon_social_reps", "nit", "codigo_habilitacion",
        "clase_prestador", "es_sociedad", "criterio_sociedad", "naturaleza",
        "caracter", "nivel", "municipio", "municipios", "departamento", "sedes",
        "direccion", "telefono", "email", "representante", "estado",
        "contactable", "segmento", "puntaje", "motivos",
    ]
    cur.executemany(
        "INSERT INTO prestadores VALUES (%s)" % ",".join("?" * len(campos_p)),
        [tuple(s.get(c, "") for c in campos_p) for s in lista],
    )

    campos_s = [
        "clave", "razon_social", "nit", "sede_nombre", "sede_numero", "municipio",
        "direccion", "telefono", "email", "codigo_habilitacion", "estado",
        "fecha_apertura",
    ]
    cur.executemany(
        "INSERT INTO sedes VALUES (%s)" % ",".join("?" * len(campos_s)),
        [tuple(s.get(c, "") for c in campos_s) for s in sedes],
    )

    cur.executemany("INSERT INTO fuente VALUES (?,?)", list(metadatos.items()))

    cur.execute("""
        CREATE VIEW v_objetivos AS
        SELECT razon_social, nit, municipios, sedes, naturaleza, nivel,
               telefono, email, direccion, representante, segmento, puntaje
        FROM prestadores
        ORDER BY puntaje DESC, sedes DESC, razon_social""")
    cur.execute("CREATE INDEX ix_prestadores_nit ON prestadores(nit)")
    cur.execute("CREATE INDEX ix_sedes_clave ON sedes(clave)")

    con.commit()
    con.close()


def guardar_csv(ruta, lista):
    with open(ruta, "w", encoding="utf-8-sig", newline="") as f:
        escritor = csv.writer(f, delimiter=";")
        escritor.writerow([titulo_col for _, titulo_col in COLUMNAS_OBJETIVOS])
        for sociedad in lista:
            escritor.writerow([sociedad.get(campo, "") for campo, _ in COLUMNAS_OBJETIVOS])


def guardar_excel(ruta, lista, sedes, metadatos, descartes, municipios):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  (sin openpyxl no se genera el Excel; instale con: pip install openpyxl)")
        return False

    libro = Workbook()
    encabezado = Font(bold=True, color="FFFFFF")
    fondo = PatternFill("solid", fgColor="1F4E78")

    def escribir_hoja(hoja, columnas, registros):
        hoja.append([t for _, t in columnas])
        for celda in hoja[1]:
            celda.font = encabezado
            celda.fill = fondo
            celda.alignment = Alignment(vertical="center")
        for registro in registros:
            hoja.append([registro.get(campo, "") for campo, _ in columnas])
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for i, (campo, tit) in enumerate(columnas, start=1):
            ancho = max(len(tit) + 4, 12)
            for registro in registros[:400]:
                ancho = max(ancho, min(len(str(registro.get(campo, ""))) + 2, 55))
            hoja.column_dimensions[get_column_letter(i)].width = ancho

    hoja = libro.active
    hoja.title = "Objetivos"
    escribir_hoja(hoja, COLUMNAS_OBJETIVOS, lista)

    escribir_hoja(libro.create_sheet("Sedes"), COLUMNAS_SEDES, sedes)

    resumen = libro.create_sheet("Resumen")
    resumen.append(["Indicador", "Valor"])
    for celda in resumen[1]:
        celda.font = encabezado
        celda.fill = fondo
    resumen.append(["Sociedades IPS encontradas", len(lista)])
    resumen.append(["Sedes registradas", len(sedes)])
    resumen.append(["Con correo electronico", sum(1 for s in lista if s["email"])])
    resumen.append(["Con telefono", sum(1 for s in lista if s["telefono"])])
    resumen.append(["Sin ningun contacto", sum(1 for s in lista if s["contactable"] == "No")])
    resumen.append([])
    resumen.append(["Por municipio (sede principal)", ""])
    for nombre in sorted(set(municipios.values())):
        resumen.append([nombre, sum(1 for s in lista if s["municipio"] == nombre)])
    resumen.append([])
    resumen.append(["Por segmento", ""])
    for seg in ("A - contactar primero", "B - contactar despues",
                "C - requiere completar datos"):
        resumen.append([seg, sum(1 for s in lista if s["segmento"] == seg)])
    resumen.append([])
    resumen.append(["Por naturaleza", ""])
    naturalezas = {}
    for s in lista:
        naturalezas[s["naturaleza"] or "(no informada)"] = \
            naturalezas.get(s["naturaleza"] or "(no informada)", 0) + 1
    for nombre, cuenta in sorted(naturalezas.items(), key=lambda x: -x[1]):
        resumen.append([nombre, cuenta])
    resumen.append([])
    resumen.append(["Filas descartadas de la fuente", ""])
    for motivo, cuenta in descartes.items():
        resumen.append([motivo, cuenta])
    resumen.column_dimensions["A"].width = 42
    resumen.column_dimensions["B"].width = 18

    fuente = libro.create_sheet("Fuente")
    fuente.append(["Dato", "Valor"])
    for celda in fuente[1]:
        celda.font = encabezado
        celda.fill = fondo
    for clave, valor in metadatos.items():
        fuente.append([clave, valor])
    fuente.column_dimensions["A"].width = 34
    fuente.column_dimensions["B"].width = 90

    libro.save(ruta)
    return True


# ---------------------------------------------------------------------------
# Programa
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Arma la base de datos de IPS del Area Metropolitana de Bucaramanga."
    )
    parser.add_argument(
        "--archivo",
        help="Ruta a un REPS ya descargado (.csv, .xlsx o .json). "
             "Si no se indica, se baja de datos.gov.co.",
    )
    parser.add_argument(
        "--municipios",
        help="Lista separada por comas para cambiar el area (por defecto los "
             "cuatro del Area Metropolitana de Bucaramanga).",
    )
    parser.add_argument(
        "--incluir-naturales", action="store_true",
        help="Incluir tambien a las personas naturales habilitadas como IPS.",
    )
    parser.add_argument(
        "--salida", default=CARPETA_SALIDA,
        help="Carpeta donde se dejan los archivos (por defecto %s)." % CARPETA_SALIDA,
    )
    args = parser.parse_args()

    municipios = dict(MUNICIPIOS_AMB)
    if args.municipios:
        pedidos = [sin_tildes(m).upper() for m in args.municipios.split(",") if m.strip()]
        municipios = {c: n for c, n in MUNICIPIOS_AMB.items()
                      if sin_tildes(n).upper() in pedidos}
        faltantes = [p for p in pedidos
                     if p not in [sin_tildes(n).upper() for n in MUNICIPIOS_AMB.values()]]
        for nombre in faltantes:
            municipios["?" + nombre] = nombre.upper()

    print("=" * 66)
    print(" IPS del Area Metropolitana de Bucaramanga")
    print(" Municipios: %s" % ", ".join(sorted(set(municipios.values()))))
    print("=" * 66)

    if args.archivo:
        filas, mapa, origen = leer_archivo(args.archivo)
    else:
        filas, mapa, origen = descargar_de_datos_abiertos(municipios)

    faltantes = [c for c in ("razon_social", "municipio", "clase_prestador")
                 if c not in mapa]
    if "razon_social" in faltantes:
        raise SystemExit(
            "\nLa fuente no trae una columna de nombre/razon social reconocible.\n"
            "Columnas encontradas: %s" % ", ".join(list(filas[0].keys())[:25])
        )
    for campo in faltantes:
        print("  AVISO: no se encontro la columna '%s'; se trabaja sin ella." % campo)

    sociedades, sedes, descartes = consolidar(
        filas, mapa, municipios, args.incluir_naturales
    )
    lista = ordenadas(sociedades)

    if not lista:
        print("\nNo se encontro ninguna IPS con esos filtros.")
        print("Descartes: %s" % descartes)
        return 1

    os.makedirs(args.salida, exist_ok=True)
    ruta_db = os.path.join(args.salida, "ips_area_metropolitana.db")
    ruta_xlsx = os.path.join(args.salida, "IPS_Area_Metropolitana.xlsx")
    ruta_csv = os.path.join(args.salida, "objetivos_ips.csv")

    sin_dato = [c for c in CANDIDATOS if c not in mapa]
    metadatos = {
        "Fuente": origen,
        "Registro": "REPS - Registro Especial de Prestadores de Servicios de Salud",
        "Entidad": "Ministerio de Salud y Proteccion Social",
        "Enlace dataset": URL_DATASET_HUMANO,
        "Enlace REPS oficial": URL_REPS_OFICIAL,
        "Fecha de extraccion": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Municipios": ", ".join(sorted(set(municipios.values()))),
        "Filtro clase": "IPS (se excluyen profesional independiente, transporte "
                        "especial y objeto social diferente)",
        "Filtro persona": "solo personas juridicas (sociedades)"
                          if not args.incluir_naturales
                          else "personas juridicas y naturales",
        "Filas leidas de la fuente": len(filas),
        "Sedes que pasaron el filtro": len(sedes),
        "Sociedades resultantes": len(lista),
        "Campos no disponibles en la fuente": ", ".join(sin_dato) or "ninguno",
        "Advertencia": "Los datos se toman tal cual del REPS. Verifique el "
                       "contacto antes de una comunicacion formal.",
    }

    guardar_sqlite(ruta_db, lista, sedes, metadatos)
    guardar_csv(ruta_csv, lista)
    hubo_excel = guardar_excel(ruta_xlsx, lista, sedes, metadatos, descartes, municipios)

    print("\n" + "-" * 66)
    print(" Sociedades IPS encontradas : %d" % len(lista))
    print(" Sedes registradas          : %d" % len(sedes))
    print(" Con correo                 : %d" % sum(1 for s in lista if s["email"]))
    print(" Con telefono               : %d" % sum(1 for s in lista if s["telefono"]))
    for nombre in sorted(set(municipios.values())):
        print("   %-16s : %d" % (nombre, sum(1 for s in lista if s["municipio"] == nombre)))
    print("-" * 66)
    print(" Base de datos : %s" % ruta_db)
    if hubo_excel:
        print(" Excel         : %s" % ruta_xlsx)
    print(" CSV para CRM  : %s" % ruta_csv)
    print("-" * 66)

    print("\n Primeros objetivos del segmento A:")
    for sociedad in lista[:10]:
        print("   %-45s %-14s %s" % (
            sociedad["razon_social"][:45],
            sociedad["municipio"][:14],
            sociedad["email"] or sociedad["telefono"] or "(sin contacto)",
        ))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(130)
