"""
Buscador de convocatorias publicas de servicios juridicos en el SECOP II.

Consulta el conjunto de DATOS ABIERTOS del SECOP II que publica Colombia
Compra Eficiente en datos.gov.co (dataset "SECOP II - Procesos de
Contratacion", identificador p6dx-8zbt) y reporta los procesos que:

  1. Fueron publicados dentro de los ultimos DIAS_ATRAS dias.
  2. Coinciden con alguna PALABRA CLAVE juridica en el nombre/objeto,
     o traen un codigo UNSPSC de la familia de servicios legales.
  3. No contienen ninguna palabra de la lista de exclusion.
  4. Estan en un estado que todavia permite presentarse (no adjudicados,
     ni terminados, ni cancelados).
  5. Opcionalmente: son de ciertos departamentos y superan un valor minimo.

Genera dos cosas:

  - buscar_convocatorias_secop.csv  -> las convocatorias encontradas.
  - buscar_convocatorias_secop_vistas.txt -> memoria de lo ya reportado,
    para que cada ejecucion muestre SOLO lo nuevo.

Uso:

    python buscar_convocatorias_secop.py              (solo lo nuevo)
    python buscar_convocatorias_secop.py --todo       (tambien lo ya visto)
    python buscar_convocatorias_secop.py --diagnostico (muestra los nombres
                                          de campo que devuelve la API, util
                                          si Colombia Compra los cambia)

IMPORTANTE: los datos abiertos se actualizan una vez al dia, asi que un
proceso publicado hoy puede aparecer aqui hasta manana. Para procesos con
cierre corto, la busqueda publica del SECOP II sigue siendo la fuente de
verdad:
https://community.secop.gov.co/Public/Tendering/ContractNoticeManagement/Index?currentLanguage=es-CO

No requiere instalar nada: usa solo la libreria estandar de Python.
"""

import csv
import json
import logging
import os
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

# ============================= CONFIGURACION =============================

# Cuantos dias hacia atras se buscan procesos publicados.
DIAS_ATRAS = 30

# Terminos que se le mandan al buscador de la API (busqueda de texto
# completo). Conviene que sean pocos y amplios; el filtro fino se hace
# despues, en PALABRAS_CLAVE.
TERMINOS_CONSULTA = [
    "juridicos",
    "juridica",
    "abogado",
    "abogados",
    "judicial",
    "cartera",
]

# Filtro fino sobre el nombre/objeto del proceso. Basta que aparezca UNA.
# Se comparan sin tildes y sin distinguir mayusculas.
PALABRAS_CLAVE = [
    "juridic",
    "abogad",
    "legal",
    "judicial",
    "litigi",
    "ejecutiv",
    "cobro",
    "cartera",
    "coactiv",
    "tutela",
    "conciliacion",
    "demanda",
    "defensa judicial",
    "representacion judicial",
    "asesoria juridica",
    "consultoria juridica",
]

# Si el nombre/objeto contiene alguna de estas, el proceso se descarta.
# Sirve para sacar el ruido (medicina legal, revisoria fiscal, etc).
PALABRAS_EXCLUIR = [
    "medicina legal",
    "medicin",
    "obra civil",
    "interventoria de obra",
    "suministro de",
    "mantenimiento de",
    "combustible",
    "papeleria",
    "aseo y cafeteria",
    "vigilancia y seguridad",
]

# Codigos UNSPSC de interes. Basta con el prefijo: "8012" cubre toda la
# familia de servicios legales. Un proceso que traiga uno de estos entra
# aunque su nombre no coincida con ninguna palabra clave.
CODIGOS_UNSPSC = [
    "8012",    # Servicios legales
    "841216",  # Servicios de cobranza / recuperacion de cartera
]

# Estados que se descartan (ya no se puede presentar oferta). Se comparan
# sin tildes y sin distinguir mayusculas, por "contiene".
ESTADOS_DESCARTADOS = [
    "adjudicado",
    "celebrado",
    "terminado",
    "cancelado",
    "descartado",
    "desierto",
    "liquidado",
    "suspendido",
]

# Departamentos de interes. Lista vacia = todo el pais.
# Ejemplo: ["Antioquia", "Bogota"]
DEPARTAMENTOS = []

# Valor minimo del proceso en pesos. 0 = sin minimo.
VALOR_MINIMO = 0

# Modalidades a descartar (vacio = ninguna). Util si no quiere ver, por
# ejemplo, contratacion directa.
MODALIDADES_DESCARTADAS = []

# ---- Parametros tecnicos (normalmente no hay que tocarlos) ----

DATASET = "p6dx-8zbt"
URL_API = "https://www.datos.gov.co/resource/{}.json".format(DATASET)
LIMITE_POR_CONSULTA = 1000
TIMEOUT_SEGUNDOS = 90

# Opcional: token de aplicacion de datos.gov.co. Sin token funciona, pero
# con limite de peticiones mas bajo. Se puede registrar gratis en
# https://www.datos.gov.co/profile/edit/developer_settings
APP_TOKEN = os.environ.get("DATOS_GOV_APP_TOKEN", "")

_BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_LOG = os.path.join(_BASE, "buscar_convocatorias_secop.log")
ARCHIVO_CSV = os.path.join(_BASE, "buscar_convocatorias_secop.csv")
ARCHIVO_VISTAS = os.path.join(_BASE, "buscar_convocatorias_secop_vistas.txt")

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def sin_tildes(texto):
    """Pasa a minusculas y quita tildes/dieresis, para comparar sin sorpresas."""
    if texto is None:
        return ""
    texto = str(texto)
    descompuesto = unicodedata.normalize("NFD", texto)
    limpio = "".join(c for c in descompuesto if unicodedata.category(c) != "Mn")
    return limpio.lower().strip()


def primer_valor(registro, nombres_posibles):
    """
    Devuelve el primer campo no vacio de entre varios nombres posibles.

    Los nombres de columna de datos.gov.co cambian de vez en cuando; asi el
    script sigue funcionando si renombran una columna, en vez de reventar.
    """
    for nombre in nombres_posibles:
        valor = registro.get(nombre)
        if valor not in (None, ""):
            return valor
    return ""


def campo_por_fragmento(registro, fragmentos):
    """Ultimo recurso: busca una clave que CONTENGA alguno de los fragmentos."""
    for clave, valor in registro.items():
        clave_limpia = sin_tildes(clave)
        for fragmento in fragmentos:
            if fragmento in clave_limpia and valor not in (None, ""):
                return valor
    return ""


def texto_url(valor):
    """El campo urlproceso llega como dict {'url': ...} o como texto plano."""
    if isinstance(valor, dict):
        return valor.get("url", "")
    return valor or ""


def consultar_api(parametros):
    url = URL_API + "?" + urllib.parse.urlencode(parametros)
    peticion = urllib.request.Request(url, headers={"Accept": "application/json"})
    if APP_TOKEN:
        peticion.add_header("X-App-Token", APP_TOKEN)
    with urllib.request.urlopen(peticion, timeout=TIMEOUT_SEGUNDOS) as respuesta:
        return json.loads(respuesta.read().decode("utf-8"))


def descargar_procesos(desde_fecha):
    """
    Trae los procesos publicados desde 'desde_fecha' que coincidan con
    alguno de los TERMINOS_CONSULTA. Devuelve una lista de registros sin
    duplicados.

    Si la API rechaza el filtro por fecha (por ejemplo porque cambio el
    nombre de la columna), reintenta sin el filtro y la fecha se descarta
    despues, en memoria.
    """
    encontrados = {}
    marca_fecha = desde_fecha.strftime("%Y-%m-%dT00:00:00.000")

    for termino in TERMINOS_CONSULTA:
        parametros = {
            "$q": termino,
            "$limit": LIMITE_POR_CONSULTA,
            "$where": "fecha_de_publicacion_del > '{}'".format(marca_fecha),
            "$order": "fecha_de_publicacion_del DESC",
        }
        try:
            registros = consultar_api(parametros)
        except urllib.error.HTTPError as error:
            if error.code == 400:
                logging.warning(
                    "  La API rechazo el filtro por fecha para '%s'; "
                    "reintentando sin el (se filtrara despues).",
                    termino,
                )
                try:
                    registros = consultar_api(
                        {"$q": termino, "$limit": LIMITE_POR_CONSULTA}
                    )
                except Exception as error2:  # noqa: BLE001
                    logging.error("  Fallo la consulta de '%s': %s", termino, error2)
                    continue
            else:
                logging.error("  Fallo la consulta de '%s': %s", termino, error)
                continue
        except Exception as error:  # noqa: BLE001
            logging.error("  Fallo la consulta de '%s': %s", termino, error)
            continue

        nuevos = 0
        for registro in registros:
            clave = identificador(registro)
            if clave and clave not in encontrados:
                encontrados[clave] = registro
                nuevos += 1
        logging.info(
            "  '%s': %d registros devueltos (%d nuevos en esta corrida).",
            termino,
            len(registros),
            nuevos,
        )

    return list(encontrados.values())


def identificador(registro):
    valor = primer_valor(
        registro,
        ["id_del_proceso", "referencia_del_proceso", "id_del_portafolio"],
    )
    if not valor:
        valor = campo_por_fragmento(registro, ["id_del_proceso", "referencia"])
    return str(valor).strip()


def texto_buscable(registro):
    partes = [
        primer_valor(registro, ["nombre_del_procedimiento"]),
        primer_valor(
            registro,
            ["descripci_n_del_procedimiento", "descripcion_del_procedimiento"],
        ),
        primer_valor(registro, ["tipo_de_contrato", "subtipo_de_contrato"]),
    ]
    if not any(partes):
        partes = [campo_por_fragmento(registro, ["nombre", "descripcion", "objeto"])]
    return sin_tildes(" | ".join(str(p) for p in partes if p))


def codigos_del_registro(registro):
    crudo = " ".join(
        str(v)
        for v in [
            primer_valor(registro, ["codigo_principal_de_categoria"]),
            primer_valor(registro, ["categorias_adicionales"]),
        ]
        if v
    )
    return re.findall(r"\d{4,10}", crudo)


def valor_numerico(registro):
    crudo = primer_valor(
        registro, ["precio_base", "valor_total_adjudicacion", "valor_estimado"]
    )
    if not crudo:
        crudo = campo_por_fragmento(registro, ["precio", "valor"])
    try:
        return float(str(crudo).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def fecha_publicacion(registro):
    return str(
        primer_valor(
            registro,
            [
                "fecha_de_publicacion_del",
                "fecha_de_publicacion_fase_3",
                "fecha_de_ultima_publicaci",
            ],
        )
    )


def interesa(registro, desde_fecha):
    """Aplica todos los filtros. Devuelve (True, '') o (False, motivo)."""
    texto = texto_buscable(registro)
    codigos = codigos_del_registro(registro)

    for palabra in PALABRAS_EXCLUIR:
        if sin_tildes(palabra) in texto:
            return False, "excluido por '{}'".format(palabra)

    coincide_palabra = any(sin_tildes(p) in texto for p in PALABRAS_CLAVE)
    coincide_codigo = any(
        codigo.startswith(prefijo) for codigo in codigos for prefijo in CODIGOS_UNSPSC
    )
    if not (coincide_palabra or coincide_codigo):
        return False, "sin coincidencia juridica"

    estado = sin_tildes(primer_valor(registro, ["estado_del_procedimiento", "estadoresumen"]))
    for descartado in ESTADOS_DESCARTADOS:
        if sin_tildes(descartado) in estado:
            return False, "estado: {}".format(estado)

    if MODALIDADES_DESCARTADAS:
        modalidad = sin_tildes(primer_valor(registro, ["modalidad_de_contratacion"]))
        for descartada in MODALIDADES_DESCARTADAS:
            if sin_tildes(descartada) in modalidad:
                return False, "modalidad descartada"

    if DEPARTAMENTOS:
        depto = sin_tildes(
            primer_valor(registro, ["departamento_entidad", "departamento_proveedor"])
        )
        if not any(sin_tildes(d) in depto for d in DEPARTAMENTOS):
            return False, "otro departamento"

    if VALOR_MINIMO and valor_numerico(registro) < VALOR_MINIMO:
        return False, "por debajo del valor minimo"

    publicacion = fecha_publicacion(registro)[:10]
    if publicacion:
        try:
            if datetime.strptime(publicacion, "%Y-%m-%d") < desde_fecha:
                return False, "publicado antes de la ventana"
        except ValueError:
            pass

    return True, ""


def cargar_vistas():
    if not os.path.exists(ARCHIVO_VISTAS):
        return set()
    with open(ARCHIVO_VISTAS, "r", encoding="utf-8") as archivo:
        return {linea.strip() for linea in archivo if linea.strip()}


def guardar_vistas(vistas):
    with open(ARCHIVO_VISTAS, "w", encoding="utf-8") as archivo:
        for clave in sorted(vistas):
            archivo.write(clave + "\n")


def fila_csv(registro):
    return {
        "fecha_publicacion": fecha_publicacion(registro)[:10],
        "cierre_ofertas": str(
            primer_valor(
                registro,
                ["fecha_de_recepcion_de", "fecha_de_apertura_de_respuesta"],
            )
        )[:16],
        "entidad": primer_valor(registro, ["entidad", "nombre_de_la_entidad"]),
        "departamento": primer_valor(registro, ["departamento_entidad"]),
        "modalidad": primer_valor(registro, ["modalidad_de_contratacion"]),
        "estado": primer_valor(registro, ["estado_del_procedimiento", "estadoresumen"]),
        "objeto": primer_valor(
            registro,
            [
                "nombre_del_procedimiento",
                "descripci_n_del_procedimiento",
                "descripcion_del_procedimiento",
            ],
        ),
        "valor": valor_numerico(registro),
        "unspsc": primer_valor(registro, ["codigo_principal_de_categoria"]),
        "id_proceso": identificador(registro),
        "enlace": texto_url(registro.get("urlproceso")),
    }


COLUMNAS = [
    "fecha_publicacion",
    "cierre_ofertas",
    "entidad",
    "departamento",
    "modalidad",
    "estado",
    "objeto",
    "valor",
    "unspsc",
    "id_proceso",
    "enlace",
]


def escribir_csv(filas):
    with open(ARCHIVO_CSV, "w", encoding="utf-8-sig", newline="") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=COLUMNAS)
        escritor.writeheader()
        for fila in filas:
            escritor.writerow(fila)


def diagnostico(desde_fecha):
    """Muestra los nombres de campo que devuelve hoy la API."""
    logging.info("Pidiendo un registro de muestra a la API...")
    try:
        registros = consultar_api({"$limit": 1})
    except Exception as error:  # noqa: BLE001
        logging.error("No se pudo consultar la API: %s", error)
        return 1
    if not registros:
        logging.error("La API no devolvio registros.")
        return 1
    logging.info("Campos disponibles en el dataset %s:", DATASET)
    for clave in sorted(registros[0].keys()):
        logging.info("  - %s", clave)
    return 0


def main():
    configurar_logging()
    argumentos = [a.lower() for a in sys.argv[1:]]
    solo_nuevos = "--todo" not in argumentos
    desde_fecha = datetime.now() - timedelta(days=DIAS_ATRAS)
    desde_fecha = desde_fecha.replace(hour=0, minute=0, second=0, microsecond=0)

    if "--diagnostico" in argumentos:
        return diagnostico(desde_fecha)

    logging.info("=" * 70)
    logging.info(
        "Busqueda de convocatorias juridicas en SECOP II  -  %s",
        datetime.now().strftime("%Y-%m-%d %H:%M"),
    )
    logging.info(
        "Ventana: ultimos %d dias (desde %s)%s",
        DIAS_ATRAS,
        desde_fecha.strftime("%Y-%m-%d"),
        "" if not DEPARTAMENTOS else "  |  departamentos: " + ", ".join(DEPARTAMENTOS),
    )
    logging.info("=" * 70)

    registros = descargar_procesos(desde_fecha)
    logging.info("Total de registros unicos traidos: %d", len(registros))

    if not registros:
        logging.warning(
            "No se trajo ningun registro. Revise la conexion a internet, o "
            "ejecute con --diagnostico para ver si cambiaron los campos."
        )
        return 1

    vistas = cargar_vistas() if solo_nuevos else set()
    seleccionados = []
    descartados = 0

    for registro in registros:
        aplica, _motivo = interesa(registro, desde_fecha)
        if not aplica:
            descartados += 1
            continue
        clave = identificador(registro)
        if solo_nuevos and clave in vistas:
            continue
        seleccionados.append(registro)

    seleccionados.sort(key=fecha_publicacion, reverse=True)
    filas = [fila_csv(r) for r in seleccionados]
    escribir_csv(filas)

    logging.info("-" * 70)
    logging.info("Descartados por los filtros: %d", descartados)
    logging.info(
        "Convocatorias %s: %d",
        "NUEVAS" if solo_nuevos else "encontradas",
        len(filas),
    )
    logging.info("-" * 70)

    for fila in filas[:40]:
        logging.info(
            "%s  [%s]  %s\n     %s\n     valor: %s   estado: %s\n     %s",
            fila["fecha_publicacion"],
            fila["modalidad"] or "sin modalidad",
            fila["entidad"],
            (fila["objeto"] or "")[:160],
            "{:,.0f}".format(fila["valor"]) if fila["valor"] else "sin dato",
            fila["estado"] or "sin dato",
            fila["enlace"] or "(sin enlace directo)",
        )
    if len(filas) > 40:
        logging.info("... y %d mas. El listado completo esta en el CSV.", len(filas) - 40)

    if solo_nuevos:
        vistas.update(identificador(r) for r in seleccionados)
        guardar_vistas(vistas)

    logging.info("")
    logging.info("Reporte: %s", ARCHIVO_CSV)
    logging.info("Log:     %s", ARCHIVO_LOG)
    return 0


if __name__ == "__main__":
    sys.exit(main())
