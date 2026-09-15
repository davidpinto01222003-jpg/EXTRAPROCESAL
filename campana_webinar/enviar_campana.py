# -*- coding: utf-8 -*-
"""
Envio INDIVIDUAL y PERSONALIZADO de la invitacion al webinar
"Recuperacion de cartera con EPS intervenidas" (miercoles 23 de
septiembre, 8:00 a. m.), desde contactoclientes@oscal.net.

Cada destinatario recibe su propio correo, dirigido a su institucion.
Nadie ve la direccion de nadie mas: NO se usa copia oculta.

Ordenes disponibles:

    python enviar_campana.py vista-previa
        Arma varios correos de ejemplo con contactos reales de la base
        y los deja como archivos .html para abrirlos en el navegador.
        NO envia nada ni se conecta a internet.

    python enviar_campana.py prueba --para tucorreo@gmail.com
        Envia UN correo real de prueba a la direccion que se indique.

    python enviar_campana.py enviar
        Hace el envio de verdad, respetando el tope diario y las pausas
        entre correo y correo. Se puede interrumpir con Ctrl+C y
        retomar despues: nunca le escribe dos veces a la misma
        direccion.

    python enviar_campana.py estado
        Muestra cuantos correos van enviados, cuantos faltan y que paso
        con los que fallaron.

    python enviar_campana.py revisar-buzon
        Entra al buzon por IMAP, busca rebotes y solicitudes de retiro,
        y los deja anotados para que no se les vuelva a escribir.

La contrasena de aplicacion de Google se pide cada vez y NO se guarda
en ningun archivo.
"""

import argparse
import configparser
import csv
import email
import getpass
import html as _html
import imaplib
import os
import random
import re
import smtplib
import ssl
import sys
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

BASE = Path(__file__).resolve().parent
CARPETA_ESTADO = BASE / "estado"
REGISTRO = CARPETA_ESTADO / "enviados.csv"
EXCLUIDOS = CARPETA_ESTADO / "excluidos.csv"
CARPETA_VISTAS = CARPETA_ESTADO / "vista_previa"

DOMINIO = "oscal.net"

# Nombres que puede tener cada columna en la base del cliente. Se
# comparan sin tildes y en minuscula, asi que "INSTITUCIÓN" tambien
# entra por "institucion".
# Columnas de correo, EN ORDEN DE PREFERENCIA: si una fila trae correo
# de juridica se usa ese; si no, el de contratacion; si no, el canal de
# contacto; y de ultimo el que figura en el REPS. Asi el mensaje llega
# a quien de verdad maneja el tema de cartera.
PRIORIDAD_CORREO = (
    "correos juridica / notificaciones judiciales",
    "correos juridica / notificaciones judi",
    "correos de contratacion / proveedores",
    "canal para enviar propuesta",
    "correo", "email", "e-mail", "mail", "correo electronico",
    "correo_electronico", "direccion de correo",
    "correo reps",
    "otros correos publicados",
)

ALIAS_COLUMNAS = {
    "institucion": ("razon social", "razon social o nombre", "institucion",
                    "entidad", "empresa", "nombre entidad", "nombre", "ips",
                    "cliente", "organizacion", "prestador"),
    "ciudad": ("municipios", "municipio", "ciudad", "ubicacion"),
    "segmento": ("segmento", "prioridad", "grupo", "orden"),
}

PATRON_CORREO = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
# Para pescar la direccion dentro de una celda que trae texto alrededor,
# como "sergioruiz@fcv.org (General)".
PATRON_DENTRO = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# Dominios que se escriben mal con mucha frecuencia.
CORRECCIONES_DOMINIO = {
    "gmail.co": "gmail.com", "gmail.con": "gmail.com", "gmial.com": "gmail.com",
    "gamil.com": "gmail.com", "hotmai.com": "hotmail.com",
    "hotmail.con": "hotmail.com", "outlok.com": "outlook.com",
    "yahoo.co": "yahoo.com",
}


# ---------------------------------------------------------------------
#  Utilidades de texto
# ---------------------------------------------------------------------

def sin_tildes(texto):
    """Devuelve el texto en minuscula y sin tildes, para comparar."""
    if not texto:
        return ""
    plano = unicodedata.normalize("NFKD", str(texto))
    plano = "".join(c for c in plano if not unicodedata.combining(c))
    return plano.strip().lower()


def arreglar_mayusculas(texto):
    """
    Las bases de datos suelen traer los nombres en MAYUSCULA SOSTENIDA,
    que en un correo se lee como un grito. Esto lo pasa a formato
    normal, respetando las siglas cortas (IPS, ESE, EPS, SAS) y las
    palabras de union.
    """
    if not texto:
        return ""
    texto = " ".join(str(texto).split())
    letras = [c for c in texto if c.isalpha()]
    if not letras or not all(c.isupper() for c in letras):
        return texto  # ya viene con formato propio: no se toca

    menores = {"de", "del", "la", "las", "los", "y", "e", "en", "el",
               "para", "por", "con", "a", "al"}
    siglas = {"IPS", "ESE", "E.S.E.", "EPS", "SAS", "S.A.S.", "SA", "S.A.",
              "LTDA", "UT", "ASI", "SAI", "CAJA", "AC"}
    salida = []
    for i, palabra in enumerate(texto.split(" ")):
        limpia = palabra.strip(".,")
        if palabra.upper() in siglas or limpia.upper() in siglas:
            salida.append(palabra.upper())
        elif i > 0 and palabra.lower() in menores:
            salida.append(palabra.lower())
        else:
            salida.append(palabra.capitalize())
    return " ".join(salida)


def limpiar_correo(valor):
    """
    Normaliza una direccion y corrige los dominios mal escritos.

    Si la celda trae texto alrededor -- "sergioruiz@fcv.org (General)",
    o dos correos separados por coma -- se queda con la primera
    direccion que encuentre.
    """
    if not valor:
        return ""
    texto = str(valor).strip()
    if "@" in texto and (" " in texto or "(" in texto or "," in texto
                         or ";" in texto or "/" in texto):
        encontradas = PATRON_DENTRO.findall(texto)
        if encontradas:
            texto = encontradas[0]
    correo = texto.strip().strip("<>").replace(" ", "").lower()
    if correo.startswith("mailto:"):
        correo = correo[7:]
    if "@" in correo:
        usuario, _, dominio = correo.rpartition("@")
        dominio = CORRECCIONES_DOMINIO.get(dominio, dominio)
        correo = usuario + "@" + dominio
    return correo


def correo_valido(correo):
    return bool(correo) and bool(PATRON_CORREO.match(correo)) and len(correo) <= 254


# ---------------------------------------------------------------------
#  Lectura de la base de contactos (.xlsx o .csv, sin dependencias)
# ---------------------------------------------------------------------

def _columna_a_indice(letras):
    indice = 0
    for caracter in letras:
        indice = indice * 26 + (ord(caracter.upper()) - 64)
    return indice - 1


def _leer_xlsx(ruta):
    """
    Lector de Excel propio, para no depender de que este instalado
    openpyxl. Lee la primera hoja del libro.
    """
    espacio = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(ruta) as libro:
        compartidas = []
        if "xl/sharedStrings.xml" in libro.namelist():
            raiz = ET.fromstring(libro.read("xl/sharedStrings.xml"))
            for nodo in raiz.findall(espacio + "si"):
                compartidas.append(
                    "".join(t.text or "" for t in nodo.iter(espacio + "t")))

        hojas = sorted(n for n in libro.namelist()
                       if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
        if not hojas:
            raise ValueError("El archivo de Excel no tiene ninguna hoja.")
        raiz = ET.fromstring(libro.read(hojas[0]))

        filas = []
        for nodo_fila in raiz.iter(espacio + "row"):
            celdas = {}
            for celda in nodo_fila.findall(espacio + "c"):
                referencia = celda.get("r") or ""
                letras = "".join(c for c in referencia if c.isalpha())
                if not letras:
                    continue
                tipo = celda.get("t")
                if tipo == "inlineStr":
                    valor = "".join(t.text or "" for t in celda.iter(espacio + "t"))
                else:
                    nodo_valor = celda.find(espacio + "v")
                    valor = (nodo_valor.text or "") if nodo_valor is not None else ""
                    if tipo == "s" and valor.strip().isdigit():
                        indice = int(valor)
                        valor = compartidas[indice] if indice < len(compartidas) else ""
                celdas[_columna_a_indice(letras)] = valor.strip()
            if celdas:
                ancho = max(celdas) + 1
                filas.append([celdas.get(i, "") for i in range(ancho)])
    return filas


def _leer_csv(ruta):
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            with open(ruta, newline="", encoding=codificacion) as archivo:
                muestra = archivo.read(4096)
                archivo.seek(0)
                try:
                    dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t|")
                except csv.Error:
                    dialecto = csv.excel
                return [fila for fila in csv.reader(archivo, dialecto)]
        except UnicodeDecodeError:
            continue
    raise ValueError("No se pudo leer el archivo CSV con ninguna codificacion.")


def leer_base(ruta):
    """
    Devuelve la lista de contactos crudos, como diccionarios con las
    claves correo / institucion / ciudad.
    """
    ruta = Path(ruta)
    if not ruta.is_absolute():
        ruta = BASE / ruta
    if not ruta.exists():
        raise FileNotFoundError(
            "No encuentro la base de contactos en:\n  %s\n\n"
            "Cree el archivo (puede partir de contactos_ejemplo.csv) o "
            "cambie la ruta en config.ini, seccion [base]." % ruta)

    if ruta.suffix.lower() in (".xlsx", ".xlsm"):
        filas = _leer_xlsx(ruta)
    else:
        filas = _leer_csv(ruta)

    filas = [f for f in filas if any(str(c).strip() for c in f)]
    if not filas:
        raise ValueError("La base de contactos esta vacia.")

    # Localizar la fila de encabezados: la primera que mencione algo
    # que parezca una columna de correo.
    indice_encabezado = None
    for i, fila in enumerate(filas[:10]):
        etiquetas = [sin_tildes(c) for c in fila]
        if any(e in PRIORIDAD_CORREO for e in etiquetas):
            indice_encabezado = i
            break

    if indice_encabezado is None:
        # No hay encabezados reconocibles: se busca, columna por
        # columna, cual contiene mas direcciones de correo.
        mejor, puntaje = None, 0
        for columna in range(max(len(f) for f in filas)):
            cuenta = sum(1 for f in filas
                         if columna < len(f) and PATRON_DENTRO.search(str(f[columna])))
            if cuenta > puntaje:
                mejor, puntaje = columna, cuenta
        if mejor is None:
            raise ValueError("La base no tiene ninguna columna con correos.")
        return [{"correo": f[mejor] if mejor < len(f) else "",
                 "institucion": "", "ciudad": "", "segmento": ""}
                for f in filas]

    encabezado = [sin_tildes(c) for c in filas[indice_encabezado]]

    # Todas las columnas de correo que existan, ordenadas por la
    # preferencia de PRIORIDAD_CORREO.
    columnas_correo = []
    for etiqueta_buscada in PRIORIDAD_CORREO:
        for i, etiqueta in enumerate(encabezado):
            if etiqueta == etiqueta_buscada and i not in columnas_correo:
                columnas_correo.append(i)
    if not columnas_correo:
        raise ValueError("La base no tiene ninguna columna de correo.")

    posicion = {}
    for campo, alias in ALIAS_COLUMNAS.items():
        for etiqueta_buscada in alias:
            encontrada = next((i for i, e in enumerate(encabezado)
                               if e == etiqueta_buscada), None)
            if encontrada is not None:
                posicion[campo] = encontrada
                break

    contactos = []
    for fila in filas[indice_encabezado + 1:]:
        correo = ""
        for i in columnas_correo:
            candidato = limpiar_correo(fila[i] if i < len(fila) else "")
            if correo_valido(candidato):
                correo = candidato
                break
        registro = {"correo": correo}
        for campo in ("institucion", "ciudad", "segmento"):
            i = posicion.get(campo)
            registro[campo] = (fila[i] if i is not None and i < len(fila) else "")
        # "BUCARAMANGA, FLORIDABLANCA, GIRON" -> "Bucaramanga": en el
        # encabezado de una carta va una sola ciudad.
        if registro["ciudad"]:
            registro["ciudad"] = registro["ciudad"].split(",")[0].strip()
        contactos.append(registro)
    return contactos


def preparar_contactos(crudos, excluidos=(), excluir_sin_institucion=False):
    """
    Limpia la base: corrige direcciones, descarta las invalidas, quita
    duplicados y arregla los nombres en mayuscula sostenida.

    Con excluir_sin_institucion en True tambien deja por fuera las
    filas que no traen razon social. En la base de IPS del area
    metropolitana ese vacio no es un error: marca a los clientes
    actuales de la firma, a quienes no hay que invitar en frio.

    Devuelve (contactos_buenos, informe).
    """
    buenos, vistos = [], set()
    informe = {"total": len(crudos), "invalidos": [], "duplicados": 0,
               "excluidos": 0, "sin_institucion": 0, "clientes_actuales": 0,
               "por_segmento": {}}
    excluidos = {limpiar_correo(c) for c in excluidos}

    for registro in crudos:
        correo = limpiar_correo(registro.get("correo"))
        if not correo_valido(correo):
            if correo or registro.get("institucion"):
                informe["invalidos"].append(
                    correo or str(registro.get("correo", ""))[:40])
            continue
        if correo in vistos:
            informe["duplicados"] += 1
            continue
        if correo in excluidos:
            informe["excluidos"] += 1
            vistos.add(correo)
            continue
        if excluir_sin_institucion and not str(registro.get("institucion", "")).strip():
            informe["clientes_actuales"] += 1
            vistos.add(correo)
            continue
        vistos.add(correo)
        buenos.append({
            "correo": correo,
            "institucion": arreglar_mayusculas(registro.get("institucion", "")),
            "ciudad": arreglar_mayusculas(registro.get("ciudad", "")),
            "segmento": (registro.get("segmento") or "").strip(),
        })
        if not buenos[-1]["institucion"]:
            informe["sin_institucion"] += 1

    # Si la base trae una columna de segmento o prioridad ("A - contactar
    # primero", "B - ..."), se respeta ese orden: los A salen antes.
    if any(c["segmento"] for c in buenos):
        buenos.sort(key=lambda c: (c["segmento"] or "zzz").upper())
        informe["por_segmento"] = {}
        for contacto in buenos:
            etiqueta = contacto["segmento"] or "(sin segmento)"
            informe["por_segmento"][etiqueta] = \
                informe["por_segmento"].get(etiqueta, 0) + 1
    return buenos, informe


# ---------------------------------------------------------------------
#  Armado del correo
# ---------------------------------------------------------------------

def bloques_saludo(contacto):
    """
    Devuelve el encabezado en version texto y en version HTML.

    Con institucion:        Sin institucion:
        Señores                 Señores
        Clinica Ejemplo         Ciudad
        Bucaramanga
    """
    ciudad = contacto.get("ciudad") or "Ciudad"
    institucion = contacto.get("institucion") or ""
    lineas = ["Señores"]
    if institucion:
        lineas.append(institucion)
    lineas.append(ciudad)
    texto = "\n".join(lineas)
    marcado = "<br>".join(
        ("<strong>%s</strong>" % _html.escape(l)) if i == 1 and institucion
        else _html.escape(l)
        for i, l in enumerate(lineas))
    return texto, marcado


def cargar_plantillas(configuracion):
    html = (BASE / "plantilla_correo.html").read_text(encoding="utf-8")
    # Los comentarios <!-- --> de la plantilla son notas para quien la
    # edita; se quitan para que no viajen dentro del correo.
    html = re.sub(r"<!--.*?-->", "", html, flags=re.S).lstrip()
    texto = (BASE / "texto_correo.txt").read_text(encoding="utf-8")
    enlace = configuracion["mensaje"]["enlace_inscripcion"].strip()
    html = html.replace("[[ENLACE]]", enlace)
    texto = texto.replace("[[ENLACE]]", enlace)
    return html, texto


def ruta_recurso(configuracion, clave):
    """
    Devuelve la ruta del recurso, o None si no esta.

    Para la imagen no exige que la extension coincida: si en config.ini
    dice .jpg y el archivo guardado es .png (o al reves), lo encuentra
    igual. Asi nadie pierde tiempo por el nombre del archivo.
    """
    valor = configuracion["mensaje"].get(clave, "").strip()
    if not valor:
        return None
    ruta = Path(valor)
    if not ruta.is_absolute():
        ruta = BASE / ruta
    if ruta.exists():
        return ruta
    if clave == "imagen":
        for extension in (".jpg", ".jpeg", ".png", ".gif"):
            alterna = ruta.with_suffix(extension)
            if alterna.exists():
                return alterna
        # Ultimo intento: cualquier imagen dentro de la carpeta.
        carpeta = ruta.parent
        if carpeta.is_dir():
            for encontrada in sorted(carpeta.iterdir()):
                if encontrada.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif"):
                    return encontrada
    return None


def construir_mensaje(contacto, configuracion, plantillas, para=None):
    """Arma el EmailMessage completo para un contacto."""
    html_base, texto_base = plantillas
    saludo_texto, saludo_html = bloques_saludo(contacto)

    texto = texto_base.replace("[[BLOQUE_SALUDO]]", saludo_texto)
    cuerpo = html_base.replace("[[BLOQUE_SALUDO]]", saludo_html)

    mensaje = EmailMessage()
    remitente = configuracion["remitente"]
    mensaje["Subject"] = configuracion["mensaje"]["asunto"].strip()
    mensaje["From"] = formataddr((remitente["nombre"].strip(),
                                  remitente["correo"].strip()))
    mensaje["To"] = para or contacto["correo"]
    responder = remitente.get("responder_a", "").strip() or remitente["correo"].strip()
    mensaje["Reply-To"] = responder
    mensaje["Date"] = formatdate(localtime=True)
    mensaje["Message-ID"] = make_msgid(domain=DOMINIO)
    # Deja que el propio cliente de correo ofrezca la baja.
    mensaje["List-Unsubscribe"] = "<mailto:%s?subject=RETIRAR>" % responder

    # La imagen va incrustada; si el archivo no existe, se quita el
    # bloque completo para que no quede un recuadro roto.
    imagen = ruta_recurso(configuracion, "imagen")
    if imagen:
        identificador = make_msgid(domain=DOMINIO)
        cuerpo = cuerpo.replace("[[BLOQUE_IMAGEN]]", (
            '<tr><td style="padding:0;line-height:0;">'
            '<img src="cid:%s" width="600" alt="Webinar: recuperacion de '
            'cartera con EPS intervenidas. Miercoles 23 de septiembre, '
            '8:00 a. m." style="display:block;width:100%%;max-width:600px;'
            'height:auto;border:0;"></td></tr>' % identificador[1:-1]))
    else:
        cuerpo = cuerpo.replace("[[BLOQUE_IMAGEN]]", "")
        identificador = None

    mensaje.set_content(texto)
    mensaje.add_alternative(cuerpo, subtype="html")

    if imagen:
        parte_html = mensaje.get_payload()[-1]
        datos = imagen.read_bytes()
        subtipo = {".png": "png", ".gif": "gif"}.get(imagen.suffix.lower(), "jpeg")
        parte_html.add_related(datos, maintype="image", subtype=subtipo,
                               cid=identificador)

    adjunto = ruta_recurso(configuracion, "adjunto")
    if adjunto:
        mensaje.add_attachment(adjunto.read_bytes(), maintype="application",
                               subtype="pdf", filename=adjunto.name)
    return mensaje


# ---------------------------------------------------------------------
#  Registro de lo que ya se envio
# ---------------------------------------------------------------------

def asegurar_carpetas():
    CARPETA_ESTADO.mkdir(exist_ok=True)
    if not REGISTRO.exists():
        with open(REGISTRO, "w", newline="", encoding="utf-8") as archivo:
            csv.writer(archivo).writerow(
                ["fecha_hora", "correo", "institucion", "estado", "detalle"])
    if not EXCLUIDOS.exists():
        with open(EXCLUIDOS, "w", newline="", encoding="utf-8") as archivo:
            csv.writer(archivo).writerow(["correo", "motivo", "fecha"])


def cargar_registro():
    """Devuelve {correo: estado} de todo lo que ya se proceso."""
    asegurar_carpetas()
    historial = {}
    with open(REGISTRO, newline="", encoding="utf-8") as archivo:
        for fila in csv.DictReader(archivo):
            historial[fila["correo"]] = fila["estado"]
    return historial


def cargar_excluidos():
    asegurar_carpetas()
    lista = []
    with open(EXCLUIDOS, newline="", encoding="utf-8") as archivo:
        for fila in csv.DictReader(archivo):
            lista.append(fila["correo"])
    return lista


def anotar(correo, institucion, estado, detalle=""):
    with open(REGISTRO, "a", newline="", encoding="utf-8") as archivo:
        csv.writer(archivo).writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            correo, institucion, estado, detalle[:200]])


def anotar_excluido(correo, motivo):
    existentes = set(cargar_excluidos())
    if correo in existentes:
        return False
    with open(EXCLUIDOS, "a", newline="", encoding="utf-8") as archivo:
        csv.writer(archivo).writerow(
            [correo, motivo, datetime.now().strftime("%Y-%m-%d")])
    return True


def enviados_hoy():
    hoy = datetime.now().strftime("%Y-%m-%d")
    total = 0
    with open(REGISTRO, newline="", encoding="utf-8") as archivo:
        for fila in csv.DictReader(archivo):
            if fila["estado"] == "enviado" and fila["fecha_hora"].startswith(hoy):
                total += 1
    return total


# ---------------------------------------------------------------------
#  Conexion con Gmail
# ---------------------------------------------------------------------

def pedir_clave(configuracion):
    variable = configuracion["servidor"].get("variable_contrasena", "").strip()
    if variable and os.environ.get(variable):
        print("  Usando la contrasena de la variable de entorno %s." % variable)
        return os.environ[variable]
    print()
    print("  Escriba la contrasena de aplicacion de Google (16 letras).")
    print("  No se vera nada mientras escribe, y no se guarda en ningun lado.")
    clave = getpass.getpass("  Contrasena: ")
    return clave.replace(" ", "")


def conectar(configuracion, clave, host=None, puerto=None):
    host = host or configuracion["servidor"]["host"].strip()
    puerto = int(puerto or configuracion["servidor"]["puerto"])
    if puerto == 465:
        servidor = smtplib.SMTP_SSL(host, puerto, timeout=60,
                                    context=ssl.create_default_context())
    else:
        servidor = smtplib.SMTP(host, puerto, timeout=60)
        servidor.ehlo()
        if servidor.has_extn("starttls"):
            servidor.starttls(context=ssl.create_default_context())
            servidor.ehlo()
    if clave:
        servidor.login(configuracion["remitente"]["correo"].strip(), clave)
    return servidor


SENALES_DE_TOPE = (
    "daily user sending quota exceeded",
    "quota exceeded",
    "try again later",
    "rate limit",
    "too many",
    "4.7.0",
    "5.4.5",
)


def es_tope_de_google(error):
    texto = str(error).lower()
    return any(senal in texto for senal in SENALES_DE_TOPE)


# ---------------------------------------------------------------------
#  Verificacion de dominios (evita rebotes antes de enviar)
# ---------------------------------------------------------------------

# Proveedores masivos: no hace falta consultarlos, siempre existen.
PROVEEDORES_CONOCIDOS = {
    "gmail.com", "hotmail.com", "hotmail.es", "hotmail.com.co",
    "outlook.com", "outlook.es", "live.com", "yahoo.com", "yahoo.es",
    "yahoo.com.co", "icloud.com", "aol.com", "protonmail.com",
    "gmail.com.co", "msn.com",
}

RESOLVEDORES = ("8.8.8.8", "1.1.1.1", "9.9.9.9")


def _consulta_dns(dominio, tipo):
    """
    Pregunta al DNS por un dominio sin depender de ninguna libreria
    externa. tipo 15 = MX, tipo 1 = A. Devuelve True si hay respuesta.
    """
    import socket as _socket
    import struct as _struct
    cabecera = _struct.pack(">HHHHHH", random.randint(0, 65535), 0x0100,
                            1, 0, 0, 0)
    pregunta = b"".join(bytes([len(p)]) + p.encode()
                        for p in dominio.split(".")) + b"\x00"
    paquete = cabecera + pregunta + _struct.pack(">HH", tipo, 1)
    for resolvedor in RESOLVEDORES:
        try:
            conexion = _socket.socket(_socket.AF_INET, _socket.SOCK_DGRAM)
            conexion.settimeout(4)
            conexion.sendto(paquete, (resolvedor, 53))
            respuesta, _ = conexion.recvfrom(2048)
            conexion.close()
            respuestas = _struct.unpack(">H", respuesta[6:8])[0]
            return respuestas > 0
        except Exception:  # noqa: BLE001
            continue
    raise OSError("sin acceso al DNS")


def dominio_recibe_correo(dominio):
    """
    True si el dominio existe y puede recibir correo. Ante la duda
    devuelve True: mas vale enviar de mas que descartar a un cliente
    bueno por una falla de red.
    """
    if dominio in PROVEEDORES_CONOCIDOS:
        return True
    try:
        if _consulta_dns(dominio, 15):   # MX
            return True
        return _consulta_dns(dominio, 1)  # A, por si recibe sin MX
    except OSError:
        # Sin DNS por UDP (red corporativa, firewall): se intenta por
        # la via normal del sistema operativo.
        import socket as _socket
        try:
            _socket.getaddrinfo(dominio, None)
            return True
        except Exception:  # noqa: BLE001
            return True  # no se pudo comprobar: no se descarta a nadie


def orden_verificar(args, configuracion):
    contactos, _ = resumen_base(configuracion)
    dominios = {}
    for contacto in contactos:
        dominios.setdefault(contacto["correo"].split("@")[1], []).append(contacto)

    por_revisar = sorted(d for d in dominios if d not in PROVEEDORES_CONOCIDOS)
    conocidos = len(dominios) - len(por_revisar)

    print()
    print("  Dominios distintos en la base   %d" % len(dominios))
    print("  De proveedores conocidos        %d  (no hace falta revisarlos)"
          % conocidos)
    print("  Por consultar en el DNS         %d" % len(por_revisar))
    print()

    muertos = []
    en_consola = sys.stdout.isatty()
    for i, dominio in enumerate(por_revisar, 1):
        if en_consola:
            print("\r  consultando %d de %d..." % (i, len(por_revisar)),
                  end="", flush=True)
        elif i % 50 == 0 or i == len(por_revisar):
            print("  consultados %d de %d" % (i, len(por_revisar)))
        if not dominio_recibe_correo(dominio):
            muertos.append(dominio)
    if en_consola:
        print("\r" + " " * 40 + "\r", end="")

    if not muertos:
        print("  Todos los dominios responden. No hay nada que descartar.")
        return

    total_afectados = sum(len(dominios[d]) for d in muertos)
    print("  Dominios que ya no existen o no reciben correo: %d"
          % len(muertos))
    print("  Direcciones afectadas: %d" % total_afectados)
    print()
    for dominio in muertos:
        for contacto in dominios[dominio]:
            print("    %-45s %s" % (contacto["correo"],
                                    contacto["institucion"][:28]))

    print()
    respuesta = input("  Excluir esas direcciones del envio? (si/no): ").strip().lower()
    if respuesta in ("si", "s", "sí"):
        anotados = 0
        for dominio in muertos:
            for contacto in dominios[dominio]:
                if anotar_excluido(contacto["correo"], "el dominio no existe"):
                    anotados += 1
        print("  Excluidas %d direcciones. El envio ya no les escribe." % anotados)
    else:
        print("  No se excluyo ninguna. Se les enviara igual.")


# ---------------------------------------------------------------------
#  Ordenes
# ---------------------------------------------------------------------

def cargar_configuracion():
    ruta = BASE / "config.ini"
    if not ruta.exists():
        sys.exit("No encuentro config.ini junto a este programa.")
    configuracion = configparser.ConfigParser()
    configuracion.read(ruta, encoding="utf-8")
    return configuracion


def resumen_base(configuracion, silencioso=False):
    crudos = leer_base(configuracion["base"]["archivo"])
    sin_nombre_fuera = (configuracion["base"]
                        .get("excluir_sin_institucion", "no")
                        .strip().lower() in ("si", "sí", "s", "true", "1"))
    contactos, informe = preparar_contactos(crudos, cargar_excluidos(),
                                            sin_nombre_fuera)
    if not silencioso:
        print("  Base de contactos: %s" % configuracion["base"]["archivo"])
        print("    filas leidas          %d" % informe["total"])
        print("    direcciones validas   %d" % len(contactos))
        if informe["duplicados"]:
            print("    duplicadas (quitadas) %d" % informe["duplicados"])
        if informe["invalidos"]:
            print("    mal escritas          %d" % len(informe["invalidos"]))
            for muestra in informe["invalidos"][:5]:
                print("        - %s" % muestra)
            if len(informe["invalidos"]) > 5:
                print("        ... y %d mas" % (len(informe["invalidos"]) - 5))
        if informe["excluidos"]:
            print("    excluidas (retiro/rebote) %d" % informe["excluidos"])
        if informe.get("clientes_actuales"):
            print("    clientes actuales     %d  (sin razon social: quedan "
                  "fuera de la campana)" % informe["clientes_actuales"])
        if informe.get("sin_institucion"):
            print("    sin razon social      %d  (saludo generico: "
                  "'Senores / Ciudad')" % informe["sin_institucion"])
        if informe.get("por_segmento"):
            print("    orden de envio por segmento:")
            for etiqueta in sorted(informe["por_segmento"]):
                print("        %-30.30s %d"
                      % (etiqueta, informe["por_segmento"][etiqueta]))
    return contactos, informe


def orden_vista_previa(args, configuracion):
    contactos, _ = resumen_base(configuracion)
    if not contactos:
        sys.exit("\nNo hay contactos validos para mostrar.")

    plantillas = cargar_plantillas(configuracion)
    CARPETA_VISTAS.mkdir(parents=True, exist_ok=True)

    imagen = ruta_recurso(configuracion, "imagen")
    adjunto = ruta_recurso(configuracion, "adjunto")
    print()
    print("  Imagen incrustada: %s" % (imagen.name if imagen else
                                       "NO (falta el archivo, el correo sale sin imagen)"))
    print("  Brochure adjunto:  %s" % (
        "%s (%.0f KB)" % (adjunto.name, adjunto.stat().st_size / 1024)
        if adjunto else "NO"))

    muestra = contactos[:args.cantidad]
    print()
    print("  " + "-" * 62)
    print("  ASI QUEDA EL PRIMER CORREO (version en texto plano)")
    print("  " + "-" * 62)
    print("  Asunto: %s" % configuracion["mensaje"]["asunto"].strip())
    print()
    texto_ejemplo = plantillas[1].replace(
        "[[BLOQUE_SALUDO]]", bloques_saludo(muestra[0])[0])
    for linea in texto_ejemplo.splitlines():
        print("  | " + linea)
    print("  " + "-" * 62)

    print()
    for i, contacto in enumerate(muestra, 1):
        cuerpo = plantillas[0].replace(
            "[[BLOQUE_SALUDO]]", bloques_saludo(contacto)[1])
        if imagen:
            try:
                relativa = os.path.relpath(imagen, CARPETA_VISTAS)
            except ValueError:
                relativa = str(imagen)
            cuerpo = cuerpo.replace("[[BLOQUE_IMAGEN]]", (
                '<tr><td style="padding:0;line-height:0;"><img src="%s" '
                'width="600" style="display:block;width:100%%;max-width:600px;'
                'height:auto;border:0;"></td></tr>' % relativa.replace("\\", "/")))
        else:
            cuerpo = cuerpo.replace("[[BLOQUE_IMAGEN]]", "")
        destino = CARPETA_VISTAS / ("ejemplo_%02d.html" % i)
        destino.write_text(cuerpo, encoding="utf-8")
        print("  %d. %-38s -> %s" % (
            i, contacto["institucion"] or contacto["correo"], destino.name))

    print()
    print("  Abra esos archivos con doble clic para ver el diseno.")
    print("  Carpeta: %s" % CARPETA_VISTAS)
    print()
    print("  No se envio ningun correo.")


def orden_prueba(args, configuracion):
    plantillas = cargar_plantillas(configuracion)
    contacto = {
        "correo": args.para,
        "institucion": args.institucion or "Clínica Ejemplo S.A.S.",
        "ciudad": args.ciudad or "Bucaramanga",
    }
    mensaje = construir_mensaje(contacto, configuracion, plantillas,
                                para=args.para)
    tamano = len(bytes(mensaje)) / 1024

    print("  Correo de prueba")
    print("    para      %s" % args.para)
    print("    dirigido a %s" % contacto["institucion"])
    print("    asunto    %s" % mensaje["Subject"])
    print("    tamano    %.0f KB" % tamano)
    print()

    clave = "" if args.host else pedir_clave(configuracion)
    print()
    print("  Conectando...")
    try:
        with conectar(configuracion, clave, args.host, args.puerto) as servidor:
            servidor.send_message(mensaje)
    except smtplib.SMTPAuthenticationError:
        sys.exit("\n  Google rechazo la contrasena.\n"
                 "  Verifique que sea una CONTRASENA DE APLICACION de 16 letras\n"
                 "  (myaccount.google.com/apppasswords), no la clave normal.")
    except Exception as error:  # noqa: BLE001
        sys.exit("\n  No se pudo enviar: %s" % error)

    print("  Enviado.")
    print()
    print("  Ahora abra ese correo, entre a los tres puntos (arriba a la")
    print("  derecha) y elija 'Mostrar original'. Debe decir PASS en SPF,")
    print("  DKIM y DMARC.")


def orden_enviar(args, configuracion):
    contactos, _ = resumen_base(configuracion)
    historial = cargar_registro()

    pendientes = [c for c in contactos if historial.get(c["correo"]) != "enviado"]
    ya_van = len(contactos) - len(pendientes)

    tope_dia = int(configuracion["envio"]["maximo_por_dia"])
    hechos_hoy = enviados_hoy()
    cupo = max(0, tope_dia - hechos_hoy)
    if args.maximo:
        cupo = min(cupo, args.maximo)
    lote = pendientes[:cupo]

    print()
    print("  Plan de este envio")
    print("    ya enviados antes     %d" % ya_van)
    print("    pendientes            %d" % len(pendientes))
    print("    tope diario           %d  (hoy ya van %d)" % (tope_dia, hechos_hoy))
    print("    se enviaran ahora     %d" % len(lote))

    if not lote:
        if not pendientes:
            print("\n  No queda nadie por contactar. La campana esta completa.")
        else:
            print("\n  Se alcanzo el tope de hoy. Vuelva a ejecutarlo manana,")
            print("  o suba maximo_por_dia en config.ini si ya calento la cuenta.")
        return

    pausa_min = float(configuracion["envio"]["pausa_minima"])
    pausa_max = float(configuracion["envio"]["pausa_maxima"])
    reintentos = int(configuracion["envio"]["reintentos"])
    minutos = len(lote) * (pausa_min + pausa_max) / 2 / 60

    print("    pausa entre correos   %.0f a %.0f segundos" % (pausa_min, pausa_max))
    print("    duracion estimada     %.0f minutos" % minutos)

    adjunto = ruta_recurso(configuracion, "adjunto")
    if adjunto:
        print("    brochure adjunto      %s (%.0f KB)" % (
            adjunto.name, adjunto.stat().st_size / 1024))
    if not ruta_recurso(configuracion, "imagen"):
        print("    AVISO: falta la imagen; los correos saldran sin ella.")

    if args.simulacro:
        print()
        print("  SIMULACRO: se arma cada correo pero no se envia nada.")
    else:
        print()
        respuesta = input("  Escriba ENVIAR en mayuscula para confirmar: ").strip()
        if respuesta != "ENVIAR":
            print("  Cancelado. No se envio nada.")
            return

    plantillas = cargar_plantillas(configuracion)
    clave = "" if (args.simulacro or args.host) else pedir_clave(configuracion)

    servidor = None
    if not args.simulacro:
        print()
        print("  Conectando con %s..." % (args.host or configuracion["servidor"]["host"]))
        try:
            servidor = conectar(configuracion, clave, args.host, args.puerto)
        except smtplib.SMTPAuthenticationError:
            sys.exit("\n  Google rechazo la contrasena. Debe ser una "
                     "CONTRASENA DE APLICACION de 16 letras.")
        except Exception as error:  # noqa: BLE001
            sys.exit("\n  No se pudo conectar: %s" % error)

    enviados = fallidos = 0
    print()
    print("  " + "-" * 62)
    try:
        for i, contacto in enumerate(lote, 1):
            etiqueta = contacto["institucion"] or contacto["correo"]
            marca = "%4d/%d  %-34.34s %-34.34s" % (
                i, len(lote), etiqueta, contacto["correo"])

            try:
                mensaje = construir_mensaje(contacto, configuracion, plantillas)
            except Exception as error:  # noqa: BLE001
                print(marca + "  ERROR AL ARMAR")
                anotar(contacto["correo"], contacto["institucion"],
                       "error", str(error))
                fallidos += 1
                continue

            if args.simulacro:
                print(marca + "  (simulacro)")
                enviados += 1
                continue

            entregado = False
            for intento in range(reintentos + 1):
                try:
                    servidor.send_message(mensaje)
                    entregado = True
                    break
                except smtplib.SMTPRecipientsRefused as error:
                    print(marca + "  RECHAZADO")
                    anotar(contacto["correo"], contacto["institucion"],
                           "rechazado", str(error))
                    anotar_excluido(contacto["correo"], "direccion rechazada")
                    fallidos += 1
                    break
                except (smtplib.SMTPServerDisconnected,
                        smtplib.SMTPConnectError) as error:
                    if intento >= reintentos:
                        raise
                    print(marca + "  reconectando...")
                    time.sleep(5 * (intento + 1))
                    try:
                        servidor.quit()
                    except Exception:  # noqa: BLE001
                        pass
                    servidor = conectar(configuracion, clave, args.host, args.puerto)
                except smtplib.SMTPException as error:
                    if es_tope_de_google(error):
                        print(marca + "  TOPE DE GOOGLE")
                        anotar(contacto["correo"], contacto["institucion"],
                               "pendiente", "tope de Google: %s" % error)
                        raise KeyboardInterrupt("tope")
                    if intento >= reintentos:
                        print(marca + "  ERROR")
                        anotar(contacto["correo"], contacto["institucion"],
                               "error", str(error))
                        fallidos += 1
                        break
                    time.sleep(5 * (intento + 1))

            if entregado:
                print(marca + "  enviado")
                anotar(contacto["correo"], contacto["institucion"], "enviado",
                       mensaje["Message-ID"])
                enviados += 1

            if i < len(lote):
                time.sleep(random.uniform(pausa_min, pausa_max))

    except KeyboardInterrupt as interrupcion:
        if str(interrupcion) == "tope":
            print()
            print("  " + "-" * 62)
            print("  Google no acepta mas correos por hoy. Se detuvo el envio.")
            print("  Vuelva a ejecutar la misma orden manana: retoma donde iba.")
        else:
            print()
            print("  Interrumpido. Lo enviado quedo registrado; puede retomar")
            print("  con la misma orden cuando quiera.")
    finally:
        if servidor is not None:
            try:
                servidor.quit()
            except Exception:  # noqa: BLE001
                pass

    print("  " + "-" * 62)
    print("  enviados en esta corrida  %d" % enviados)
    if fallidos:
        print("  con problemas             %d" % fallidos)
    restantes = len(pendientes) - enviados
    print("  quedan pendientes         %d" % max(0, restantes))


def orden_estado(args, configuracion):
    contactos, _ = resumen_base(configuracion)
    historial = cargar_registro()

    conteos = {}
    for estado in historial.values():
        conteos[estado] = conteos.get(estado, 0) + 1

    print()
    print("  Estado de la campana")
    print("    contactos en la base  %d" % len(contactos))
    for estado in sorted(conteos):
        print("    %-20s  %d" % (estado, conteos[estado]))
    pendientes = sum(1 for c in contactos
                     if historial.get(c["correo"]) != "enviado")
    print("    pendientes            %d" % pendientes)
    print("    enviados hoy          %d de %s" % (
        enviados_hoy(), configuracion["envio"]["maximo_por_dia"]))

    excluidos = cargar_excluidos()
    if excluidos:
        print("    excluidos             %d" % len(excluidos))

    problemas = [(c, e) for c, e in historial.items()
                 if e not in ("enviado",)]
    if problemas:
        print()
        print("  Direcciones con problema (ultimas 10):")
        for correo, estado in problemas[-10:]:
            print("    %-45s %s" % (correo, estado))
    print()
    print("  Detalle completo en: %s" % REGISTRO)


def orden_revisar_buzon(args, configuracion):
    """
    Entra al buzon y busca dos cosas: rebotes (direcciones que no
    existen) y respuestas pidiendo el retiro. Las anota para que no se
    les vuelva a escribir.
    """
    correo_propio = configuracion["remitente"]["correo"].strip()
    clave = pedir_clave(configuracion)
    dias = int(configuracion["buzon"]["dias"])

    print()
    print("  Conectando al buzon de %s..." % correo_propio)
    try:
        buzon = imaplib.IMAP4_SSL(configuracion["buzon"]["host"],
                                  int(configuracion["buzon"]["puerto"]))
        buzon.login(correo_propio, clave)
        buzon.select("INBOX")
    except Exception as error:  # noqa: BLE001
        sys.exit("  No se pudo entrar al buzon: %s" % error)

    desde = (datetime.now().timestamp() - dias * 86400)
    criterio = datetime.fromtimestamp(desde).strftime("%d-%b-%Y")
    estado, datos = buzon.search(None, "SINCE", criterio)
    identificadores = datos[0].split() if estado == "OK" and datos[0] else []

    print("  Revisando %d correos de los ultimos %d dias..."
          % (len(identificadores), dias))

    rebotes, retiros, interesados = [], [], []
    palabras_retiro = ("retirar", "no deseo recibir", "dar de baja",
                       "darme de baja", "unsubscribe", "eliminar mi correo",
                       "no enviar mas")

    for numero in identificadores:
        estado, datos = buzon.fetch(numero, "(RFC822)")
        if estado != "OK" or not datos or not isinstance(datos[0], tuple):
            continue
        mensaje = email.message_from_bytes(datos[0][1])
        remitente = (mensaje.get("From") or "").lower()
        asunto = (mensaje.get("Subject") or "")

        cuerpo = ""
        try:
            if mensaje.is_multipart():
                for parte in mensaje.walk():
                    if parte.get_content_type() == "text/plain":
                        cuerpo += parte.get_payload(decode=True).decode(
                            parte.get_content_charset() or "utf-8", "replace")
            else:
                cuerpo = mensaje.get_payload(decode=True).decode(
                    mensaje.get_content_charset() or "utf-8", "replace")
        except Exception:  # noqa: BLE001
            cuerpo = ""

        es_rebote = ("mailer-daemon" in remitente or "postmaster" in remitente
                     or "delivery status" in asunto.lower()
                     or "undelivered" in asunto.lower()
                     or "no se ha entregado" in asunto.lower())

        if es_rebote:
            for direccion in set(re.findall(
                    r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
                    cuerpo)):
                direccion = direccion.lower()
                if direccion != correo_propio and "google.com" not in direccion:
                    rebotes.append(direccion)
            continue

        texto_completo = (asunto + " " + cuerpo[:600]).lower()
        direccion = re.findall(
            r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", remitente)
        direccion = direccion[0].lower() if direccion else ""

        if any(palabra in texto_completo for palabra in palabras_retiro):
            if direccion:
                retiros.append(direccion)
        elif direccion and direccion != correo_propio:
            interesados.append((direccion, asunto[:60]))

    try:
        buzon.close()
        buzon.logout()
    except Exception:  # noqa: BLE001
        pass

    nuevos = 0
    for direccion in set(rebotes):
        if anotar_excluido(direccion, "rebote"):
            nuevos += 1
    for direccion in set(retiros):
        if anotar_excluido(direccion, "pidio retiro"):
            nuevos += 1

    print()
    print("  Rebotes encontrados      %d" % len(set(rebotes)))
    print("  Solicitudes de retiro    %d" % len(set(retiros)))
    print("  Excluidos nuevos         %d" % nuevos)
    for direccion in sorted(set(retiros)):
        print("      retiro: %s" % direccion)

    if interesados:
        print()
        print("  Respuestas que vale la pena mirar (%d):" % len(interesados))
        for direccion, asunto in interesados[:25]:
            print("      %-42s %s" % (direccion, asunto))

    print()
    print("  Los excluidos quedan en: %s" % EXCLUIDOS)
    print("  El proximo envio ya no les escribe.")


# ---------------------------------------------------------------------

def main():
    analizador = argparse.ArgumentParser(
        description="Envio personalizado de la invitacion al webinar de OSCAL.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ordenes = analizador.add_subparsers(dest="orden")

    p = ordenes.add_parser("vista-previa",
                           help="arma correos de ejemplo sin enviar nada")
    p.add_argument("--cantidad", type=int, default=3)
    p.add_argument("--base", help="usar otra base en vez de la del config.ini")
    p.set_defaults(funcion=orden_vista_previa)

    p = ordenes.add_parser("prueba", help="envia un correo de prueba")
    p.add_argument("--para", required=True)
    p.add_argument("--institucion", default="")
    p.add_argument("--ciudad", default="")
    p.add_argument("--host", help=argparse.SUPPRESS)
    p.add_argument("--puerto", help=argparse.SUPPRESS)
    p.set_defaults(funcion=orden_prueba)

    p = ordenes.add_parser("enviar", help="hace el envio real")
    p.add_argument("--maximo", type=int, default=0,
                   help="tope adicional solo para esta corrida")
    p.add_argument("--simulacro", action="store_true",
                   help="arma todo pero no envia")
    p.add_argument("--base", help="usar otra base en vez de la del config.ini")
    p.add_argument("--host", help=argparse.SUPPRESS)
    p.add_argument("--puerto", help=argparse.SUPPRESS)
    p.set_defaults(funcion=orden_enviar)

    p = ordenes.add_parser("estado", help="muestra el avance de la campana")
    p.add_argument("--base", help="usar otra base en vez de la del config.ini")
    p.set_defaults(funcion=orden_estado)

    p = ordenes.add_parser("verificar",
                           help="descarta dominios que ya no existen")
    p.add_argument("--base", help="usar otra base en vez de la del config.ini")
    p.set_defaults(funcion=orden_verificar)

    p = ordenes.add_parser("revisar-buzon",
                           help="busca rebotes y solicitudes de retiro")
    p.set_defaults(funcion=orden_revisar_buzon)

    args = analizador.parse_args()
    if not args.orden:
        analizador.print_help()
        return

    asegurar_carpetas()
    configuracion = cargar_configuracion()
    if getattr(args, "base", None):
        configuracion["base"]["archivo"] = args.base
    print()
    try:
        args.funcion(args, configuracion)
    except (FileNotFoundError, ValueError) as error:
        sys.exit("\n  %s" % error)
    print()


if __name__ == "__main__":
    main()
