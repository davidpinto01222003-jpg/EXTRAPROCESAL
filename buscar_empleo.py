"""
Buscador y postulador automatico de empleo (Colombia).

Vigila los portales de empleo colombianos (Computrabajo, elempleo,
Magneto365, Servicio Publico de Empleo, LinkedIn...), compara cada
vacante nueva contra TU perfil y postula solo a las que de verdad
encajan contigo -- es decir, a las que tienes opcion real de pasar.

COMO FUNCIONA, EN ORDEN
============================================================
  1. LEE TU PERFIL desde `perfil_laboral.json` (cargos que buscas,
     ciudades, anos de experiencia, nivel educativo, salario minimo,
     palabras que NO quieres ver, y la ruta a tu hoja de vida).
  2. BUSCA en cada portal habilitado, usando tus cargos objetivo como
     terminos de busqueda, con un navegador real (Playwright).
  3. ABRE cada vacante nueva y le lee la descripcion completa.
  4. LA CALIFICA de 0 a 100 contra tu perfil (ver evaluar_vacante):
        - Descarta de una si pide mas experiencia de la que tienes,
          mas estudios de los que tienes, paga menos de tu minimo,
          queda en una ciudad que no trabajas, o trae una palabra de
          tu lista de excluyentes.
        - Suma puntos si el cargo es de los que buscas, si aparecen
          tus palabras clave, si la ciudad es la tuya, etc.
  5. POSTULA solo a las que pasan tu UMBRAL (por defecto 65 puntos):
        - Si la vacante publica un CORREO de contacto -> le manda tu
          hoja de vida adjunta con una carta de presentacion armada
          para ESA vacante. Esto es 100% automatico.
        - Si no, entra al portal (con TU sesion ya iniciada) y llena
          el formulario de postulacion del propio portal.
  6. ANOTA TODO en `datos_empleo/postulaciones.json` y en un Excel,
     para no postular dos veces a lo mismo y para que puedas revisar
     que se mando, a donde y por que.

MODOS (ver MODO en CONFIGURACION)
============================================================
  "simulacion"     -> busca, califica y te muestra a que postularia,
                      pero NO manda nada. Empieza SIEMPRE por aqui:
                      es la forma de ver si tu perfil esta bien
                      calibrado antes de mandar nada real.
  "semiautomatico" -> manda las postulaciones por correo solas, y en
                      los portales hace el clic de "Postularme" (que en
                      la mayoria ya postula de una). Si el portal abre
                      ADEMAS un formulario, no lo llena: deja la vacante
                      como "pendiente_revision" para que la termines tu.
  "automatico"     -> manda todo solo, sin preguntar. Solo envia en
                      portales cuando reconocio el formulario
                      completo; si el formulario pide algo que no
                      entiende (preguntas del empleador, examenes,
                      subir documentos raros), lo deja marcado como
                      "pendiente_revision" para que lo termines tu.

ANTES DE DEJARLO CORRIENDO -- LEE ESTO
============================================================
Los portales de empleo (LinkedIn en particular) prohiben en sus
terminos de uso el acceso automatizado, y pueden suspender la cuenta
que lo haga. Por eso este programa esta armado asi:

  - Los envios por CORREO son la via principal, y no dependen de
    ningun portal: mandar tu hoja de vida al correo que la propia
    oferta publica no infringe nada.
  - En los portales, el programa usa TU navegador con TU sesion (la
    inicias tu mismo, a mano, una sola vez con --login) y solo repite
    los pasos que tu ya estas autorizado a hacer manualmente. No entra
    a APIs ocultas, no evade captchas ni bloqueos, y espera entre
    accion y accion como una persona.
  - LinkedIn viene DESACTIVADO por defecto en FUENTES. Actívalo solo
    si aceptas el riesgo de que te restrinjan la cuenta.

Y un consejo practico, no tecnico: postular a todo lo que se mueva
NO sube tus opciones, las baja (los reclutadores del mismo grupo
empresarial ven las postulaciones repetidas). Por eso el programa trae
un tope diario (MAX_POSTULACIONES_DIA) y un umbral de puntaje. Deja el
umbral alto y el tope bajo; es mejor 8 postulaciones donde encajas que
80 donde no.

USO
============================================================
    python buscar_empleo.py --login        (una sola vez: inicias sesion tu)
    python buscar_empleo.py --simular      (prueba: no manda nada)
    python buscar_empleo.py                (una pasada de verdad)
    python buscar_empleo.py --vigilar      (queda vigilando cada N minutos)
    python buscar_empleo.py --reporte      (exporta el Excel y resume)
    python buscar_empleo.py --diagnostico  (revisa si los portales responden)

En Windows puedes usar los .bat: `buscar_empleo.bat` y `vigilar_empleo.bat`.
"""

import argparse
import datetime
import json
import logging
import mimetypes
import os
import random
import re
import smtplib
import ssl
import sys
import time
import unicodedata
from email.message import EmailMessage
from pathlib import Path

# ============================================================
# CONFIGURACION
# ============================================================

CARPETA_BASE = Path(__file__).resolve().parent

# Tu perfil (copia perfil_laboral.example.json -> perfil_laboral.json).
ARCHIVO_PERFIL = CARPETA_BASE / "perfil_laboral.json"

# Correo desde el que se mandan las postulaciones (copia
# credenciales_empleo.example.txt -> credenciales_empleo.txt).
ARCHIVO_CREDENCIALES = CARPETA_BASE / "credenciales_empleo.txt"

# Todo lo que el programa genera vive aqui (registro, log, Excel).
CARPETA_DATOS = CARPETA_BASE / "datos_empleo"

# Sesiones del navegador (cookies de los portales). Se crea sola con --login.
CARPETA_NAVEGADOR = CARPETA_DATOS / "perfil_navegador"

ARCHIVO_REGISTRO = CARPETA_DATOS / "postulaciones.json"
ARCHIVO_EXCEL = CARPETA_DATOS / "postulaciones.xlsx"
ARCHIVO_LOG = CARPETA_DATOS / "buscar_empleo.log"

# "simulacion" | "semiautomatico" | "automatico"  (ver el docstring de arriba)
MODO = "simulacion"

# Puntaje minimo (0-100) para postular. Sube este numero si te llegan
# ofertas que no te sirven; bajalo si casi nunca postula a nada.
UMBRAL_POSTULACION = 65

# Tope de postulaciones por dia. Protege tu reputacion y tu cuenta.
MAX_POSTULACIONES_DIA = 15

# Cuantas paginas de resultados revisar por termino de busqueda.
PAGINAS_POR_BUSQUEDA = 2

# Segundos de espera entre postulacion y postulacion (se le suma algo
# de azar). No lo bajes: mandar 20 correos en 10 segundos parece spam.
ESPERA_ENTRE_POSTULACIONES = (45, 90)

# Segundos de espera entre vacante y vacante al leer descripciones.
ESPERA_ENTRE_LECTURAS = (3, 7)

# Cada cuanto vuelve a revisar en modo --vigilar.
INTERVALO_VIGILANCIA_MIN = 120

# Al terminar cada pasada, mandarte un correo a TI con el resumen (lo
# que se postulo y lo que te espera). Es la forma de enterarte desde el
# celular sin abrir nada: llega como cualquier notificacion de correo.
# Solo se manda cuando hubo algo que contar. Necesita
# credenciales_empleo.txt configurado.
AVISAR_POR_CORREO = True

# El navegador se ve o no. En --login y --diagnostico siempre se ve.
# Ponlo en True si algun portal te devuelve 0 resultados siempre: varios
# rechazan al navegador cuando corre oculto, y visible si los dejan pasar.
NAVEGADOR_VISIBLE = False

# Cuanto espera a que cargue cada pagina (milisegundos).
TIMEOUT_PAGINA_MS = 45000

# Cuantos dias hacia atras acepta una oferta (las viejas ya estan cerradas).
MAX_DIAS_ANTIGUEDAD = 30

# Solo por seguridad: nunca manda un correo a estas direcciones aunque
# aparezcan en la oferta (son buzones que NO reciben postulaciones).
CORREOS_PROHIBIDOS = (
    "noreply", "no-reply", "notificaciones", "soporte@", "info@computrabajo",
    "servicioalcliente", "atencionalcliente", "webmaster", "postmaster",
)


# ============================================================
# FUENTES (los portales)
# ============================================================
# Cada portal se describe aqui. Si alguno cambia su pagina y deja de
# devolver resultados, NO hay que tocar el codigo: se corrigen la
# direccion o los selectores aqui, y listo. `python buscar_empleo.py
# --diagnostico` te dice exactamente cual dejo de funcionar.
#
#   habilitada        -> si se busca en ese portal
#   plantillas_url    -> VARIAS direcciones candidatas, en orden. Se
#                        prueban de arriba abajo y se usa la PRIMERA que
#                        devuelva ofertas; el log dice cual funciono.
#                        Existe porque los portales cambian la forma de
#                        sus direcciones cada tanto, y asi el programa se
#                        adapta solo en vez de quedarse en 0 resultados.
#                        Comodines: {q}=termino-con-guiones, {qp}=termino
#                        para parametro, {ciudad} y {ciudadp} igual,
#                        {pagina}=numero de pagina.
#   patron_enlace     -> como se reconoce el link de una oferta. Es la
#                        red de seguridad: aunque cambien todas las
#                        clases CSS, los links siguen teniendo esta forma.
#   sel_tarjeta       -> selectores CSS candidatos de cada tarjeta
#   sel_titulo/empresa/ciudad -> selectores dentro de la tarjeta
#   sel_descripcion   -> selectores del texto de la oferta ya abierta
#   necesita_sesion   -> si hay que iniciar sesion (con --login)
#   sel_boton_postular-> textos del boton de postularse en el portal

FUENTES = {
    "computrabajo": {
        "habilitada": True,
        "nombre": "Computrabajo",
        "plantillas_url": [
            "https://co.computrabajo.com/trabajo-de-{q}-en-{ciudad}?p={pagina}",
            "https://co.computrabajo.com/trabajo-de-{q}?p={pagina}",
            "https://co.computrabajo.com/empleos-en-{ciudad}?q={qp}&p={pagina}",
            "https://co.computrabajo.com/empleos?q={qp}&p={pagina}",
        ],
        "patron_enlace": r"/ofertas-de-trabajo/",
        "sel_tarjeta": ["article.box_offer", "article[data-id]", ".js-o-container"],
        "sel_titulo": ["h2 a", "a.js-o-link", "h1 a"],
        "sel_empresa": ["p.dFlex.vMiddle a", ".fs16 a", "p.fs16"],
        "sel_ciudad": ["p.fs16 span", ".fs13", "p.fs13"],
        "sel_descripcion": ["div.box_detail", "div.fs16.t_word_wrap", "#detalle-oferta"],
        "necesita_sesion": True,
        "sel_boton_postular": ["postularme", "postular ahora", "postular", "aplicar"],
    },
    "elempleo": {
        "habilitada": True,
        "nombre": "elempleo.com",
        "plantillas_url": [
            "https://www.elempleo.com/co/ofertas-empleo/?Search={qp}&PageIndex={pagina}",
            "https://www.elempleo.com/co/ofertas-empleo/?Keyword={qp}&PageIndex={pagina}",
            "https://www.elempleo.com/co/ofertas-empleo/{q}",
        ],
        "patron_enlace": r"/co/ofertas-trabajo/",
        "sel_tarjeta": [".result-item", ".offer-item", "article.result"],
        "sel_titulo": ["a.js-o-link", ".title-offer a", "h2 a", "a"],
        "sel_empresa": [".company-name", ".text-ellipsis", "span.company"],
        "sel_ciudad": [".city", ".location", "span.info-city"],
        "sel_descripcion": [".description-block", ".offer-description", "#Description"],
        "necesita_sesion": True,
        "sel_boton_postular": ["aplicar ahora", "aplicar", "postularme", "postular"],
    },
    "magneto": {
        "habilitada": True,
        "nombre": "Magneto365",
        "plantillas_url": [
            "https://www.magneto365.com/co/empleos?search={qp}&page={pagina}",
            "https://www.magneto365.com/co/empleos?q={qp}&page={pagina}",
            "https://www.magneto365.com/co/empleos/{q}",
        ],
        "patron_enlace": r"/co/empleos/",
        "sel_tarjeta": ["article", ".vacancy-card", "li.vacancy"],
        "sel_titulo": ["h2", "h3", "a"],
        "sel_empresa": [".company", "h4", "p.company-name"],
        "sel_ciudad": [".location", ".city", "p.location"],
        "sel_descripcion": [".vacancy-detail", "main", "article"],
        "necesita_sesion": True,
        "sel_boton_postular": ["postularme", "aplicar", "postular"],
    },
    "spe": {
        "habilitada": True,
        "nombre": "Servicio Publico de Empleo",
        "plantillas_url": [
            "https://serviciodeempleo.gov.co/buscar-empleo?keyword={qp}&page={pagina}",
            "https://serviciodeempleo.gov.co/buscar-empleo?palabraClave={qp}&pagina={pagina}",
            "https://serviciodeempleo.gov.co/buscar-empleo?q={qp}",
        ],
        "patron_enlace": r"(oferta|vacante)",
        "sel_tarjeta": ["article", ".card-vacante", ".resultado"],
        "sel_titulo": ["h2", "h3", "a"],
        "sel_empresa": [".empresa", "h4"],
        "sel_ciudad": [".ciudad", ".ubicacion"],
        "sel_descripcion": ["main", ".detalle-vacante"],
        "necesita_sesion": False,
        "sel_boton_postular": ["postularme", "postular", "aplicar"],
    },
    "linkedin": {
        # DESACTIVADO a proposito: LinkedIn prohibe el acceso
        # automatizado y suspende cuentas. Ponlo en True solo si
        # aceptas ese riesgo. Lee la advertencia del docstring.
        "habilitada": False,
        "nombre": "LinkedIn",
        "plantillas_url": [
            "https://www.linkedin.com/jobs/search/?keywords={qp}"
            "&location={ciudadp}%2C%20Colombia&f_TPR=r604800&start={pagina0_25}",
        ],
        "patron_enlace": r"/jobs/view/",
        "sel_tarjeta": ["div.job-card-container", "li.jobs-search-results__list-item", "div.base-card"],
        "sel_titulo": ["a.job-card-list__title", "h3", "a"],
        "sel_empresa": [".job-card-container__primary-description", "h4", ".base-search-card__subtitle"],
        "sel_ciudad": [".job-card-container__metadata-item", ".job-search-card__location"],
        "sel_descripcion": [".jobs-description__content", ".description__text", "article"],
        "necesita_sesion": True,
        "sel_boton_postular": ["solicitud sencilla", "easy apply", "solicitar", "postularme"],
    },
}


# ============================================================
# LOG
# ============================================================

def configurar_log(verboso: bool = True) -> logging.Logger:
    CARPETA_DATOS.mkdir(parents=True, exist_ok=True)
    log = logging.getLogger("empleo")
    if log.handlers:
        return log
    log.setLevel(logging.DEBUG)
    formato = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%Y-%m-%d %H:%M:%S")

    archivo = logging.FileHandler(ARCHIVO_LOG, encoding="utf-8")
    archivo.setFormatter(formato)
    archivo.setLevel(logging.DEBUG)
    log.addHandler(archivo)

    # Cuando se arranca con pythonw.exe (sin ventana, que es como corre al
    # iniciar Windows) no hay consola a donde escribir: ahi solo queda el
    # archivo de log, y hay que no intentar siquiera crear el handler.
    if sys.stdout is not None:
        consola = logging.StreamHandler(sys.stdout)
        consola.setFormatter(logging.Formatter("%(message)s"))
        consola.setLevel(logging.DEBUG if verboso else logging.INFO)
        log.addHandler(consola)
    return log


LOG = configurar_log()


# ============================================================
# TEXTO: normalizacion y lectura de requisitos
# ============================================================

def sin_tildes(texto: str) -> str:
    """'Ingeniería' -> 'ingenieria'. Todo se compara asi, en minuscula."""
    if not texto:
        return ""
    normal = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in normal if not unicodedata.combining(c)).lower().strip()


def limpiar_espacios(texto: str) -> str:
    return re.sub(r"\s+", " ", (texto or "")).strip()


def a_slug(texto: str) -> str:
    """'Auxiliar Jurídico' -> 'auxiliar-juridico' (para URLs de ruta)."""
    base = sin_tildes(texto)
    base = re.sub(r"[^a-z0-9]+", "-", base)
    return base.strip("-")


def contiene_palabra(texto_normalizado: str, termino: str) -> bool:
    """Busca el termino como palabra(s) completa(s), no como pedazo.

    Asi 'contador' no coincide con 'contadora de historias', y sobre
    todo 'sin experiencia' no se confunde con 'experiencia'.
    """
    termino = sin_tildes(termino)
    if not termino:
        return False
    patron = r"(?<![a-z0-9])" + re.escape(termino).replace(r"\ ", r"\s+") + r"(?![a-z0-9])"
    return re.search(patron, texto_normalizado) is not None


NIVELES_EDUCATIVOS = {
    "ninguno": 0,
    "bachiller": 1,
    "tecnico": 2,
    "tecnologo": 3,
    "profesional": 4,
    "especializacion": 5,
    "maestria": 6,
    "doctorado": 7,
}

# Como se nombra cada nivel en las ofertas reales.
SINONIMOS_NIVEL = {
    "bachiller": ["bachiller", "bachillerato", "educacion media"],
    "tecnico": ["tecnico", "tecnica laboral"],
    "tecnologo": ["tecnologo", "tecnologia en"],
    "profesional": ["profesional", "pregrado", "universitario", "titulo universitario", "egresado"],
    "especializacion": ["especializacion", "especialista", "postgrado", "posgrado"],
    "maestria": ["maestria", "magister", "master"],
    "doctorado": ["doctorado", "phd"],
}


def nivel_a_numero(nombre: str) -> int:
    return NIVELES_EDUCATIVOS.get(sin_tildes(nombre), 0)


def experiencia_pedida(texto_normalizado: str):
    """Devuelve los anos de experiencia que pide la oferta, o None.

    Reconoce las formas que de verdad se escriben en las ofertas
    colombianas: '2 anos de experiencia', 'experiencia minima de 3
    anos', 'de 1 a 2 anos', '18 meses de experiencia', y tambien
    'sin experiencia' / 'no requiere experiencia' (que devuelve 0).
    """
    if not texto_normalizado:
        return None

    for frase in ("sin experiencia", "no requiere experiencia", "no se requiere experiencia",
                  "con o sin experiencia", "sin experiencia previa", "primer empleo"):
        if frase in texto_normalizado:
            return 0

    candidatos = []

    # "de 1 a 3 anos" -> se toma el minimo (1), que es lo que exigen.
    for m in re.finditer(r"(?:de\s+)?(\d{1,2})\s*(?:a|-|hasta)\s*(\d{1,2})\s*an?os", texto_normalizado):
        candidatos.append(int(m.group(1)))

    # "2 anos de experiencia" / "experiencia minima de 2 anos"
    for m in re.finditer(r"(\d{1,2})\s*an?os?\b", texto_normalizado):
        ventana = texto_normalizado[max(0, m.start() - 60): m.end() + 60]
        if "experiencia" in ventana:
            candidatos.append(int(m.group(1)))

    # "18 meses de experiencia"
    for m in re.finditer(r"(\d{1,3})\s*meses", texto_normalizado):
        ventana = texto_normalizado[max(0, m.start() - 60): m.end() + 60]
        if "experiencia" in ventana:
            candidatos.append(int(m.group(1)) // 12)

    if not candidatos:
        return None
    # Si la oferta menciona varios numeros, el exigente es el menor:
    # "minimo 1 ano" con "ideal 5 anos" se filtra por el minimo real.
    return min(candidatos)


def nivel_pedido(texto_normalizado: str):
    """Nivel educativo mas alto que se menciona como requisito, o None."""
    encontrado = None
    for nivel, palabras in SINONIMOS_NIVEL.items():
        for palabra in palabras:
            if contiene_palabra(texto_normalizado, palabra):
                valor = nivel_a_numero(nivel)
                if encontrado is None or valor > encontrado:
                    encontrado = valor
    return encontrado


def salario_ofrecido(texto: str):
    """Saca el salario en pesos de un texto ('$ 2.500.000', '2500000').

    Devuelve el numero mas bajo que parezca un salario mensual real
    (entre 800 mil y 60 millones), o None si no se puede saber.
    """
    if not texto:
        return None
    candidatos = []
    for m in re.finditer(r"(\d[\d\.\, ]{5,})", texto):
        crudo = m.group(1)
        solo_digitos = re.sub(r"[^\d]", "", crudo)
        if not solo_digitos:
            continue
        try:
            valor = int(solo_digitos)
        except ValueError:
            continue
        if 800_000 <= valor <= 60_000_000:
            candidatos.append(valor)
    # Formas cortas: "2.5 millones", "3 millones"
    for m in re.finditer(r"(\d{1,2}(?:[\.,]\d{1,2})?)\s*millon", sin_tildes(texto)):
        try:
            candidatos.append(int(float(m.group(1).replace(",", ".")) * 1_000_000))
        except ValueError:
            pass
    if not candidatos:
        return None
    return min(candidatos)


def antiguedad_en_dias(texto: str):
    """'Publicado hace 3 dias' -> 3. 'Hoy'/'ayer' -> 0/1. None si no dice."""
    t = sin_tildes(texto)
    if not t:
        return None
    if re.search(r"\b(hoy|hace un momento|hace \d+ (minutos?|horas?))\b", t):
        return 0
    if "ayer" in t:
        return 1
    m = re.search(r"hace\s+(\d{1,3})\s*dias?", t)
    if m:
        return int(m.group(1))
    m = re.search(r"hace\s+(\d{1,2})\s*(semanas?|meses?)", t)
    if m:
        return int(m.group(1)) * (7 if "semana" in m.group(2) else 30)
    return None


def correos_en_texto(texto: str):
    """Correos a los que SI tiene sentido mandar una hoja de vida."""
    if not texto:
        return []
    crudos = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", texto)
    buenos = []
    for correo in crudos:
        bajo = correo.lower()
        if any(malo in bajo for malo in CORREOS_PROHIBIDOS):
            continue
        if bajo.endswith((".png", ".jpg", ".gif")):
            continue
        if bajo not in buenos:
            buenos.append(bajo)
    return buenos


# ============================================================
# PERFIL Y CREDENCIALES
# ============================================================

PERFIL_POR_DEFECTO = {
    "nombre_completo": "",
    "correo": "",
    "telefono": "",
    "ciudad": "",
    "ruta_hoja_de_vida": "",
    "cargos_objetivo": [],
    "palabras_clave_deseables": [],
    "palabras_excluyentes": [],
    "ciudades": [],
    "acepta_remoto": True,
    "anos_experiencia": 0,
    "nivel_educativo": "bachiller",
    "salario_minimo": 0,
    "resumen_profesional": "",
    "asunto_correo": "Postulacion: {titulo}",
    "carta_presentacion": "",
}


class ErrorConfiguracion(Exception):
    """Falta algo que el usuario tiene que llenar antes de correr esto."""


def reparar_perfil_escrito_a_mano(texto: str):
    """Arregla los tres errores que todo el mundo comete editando el perfil.

    El formato de este archivo (JSON) es quisquilloso de una forma que no
    tiene nada que ver con buscar empleo, y por tres tonterias se niega a
    abrir entero. Antes que dejar al usuario peleando con comas, se
    corrigen aqui:

      1. Rutas de Windows con una sola barra: C:\\Users\\... En este
         formato la barra invertida hay que escribirla dos veces.
      2. Comas que faltan entre los renglones de una lista.
      3. Numeros escritos con puntos de miles: 1.750.000 en vez de 1750000.

    Devuelve (texto_corregido, lista_de_lo_que_se_corrigio). Nunca
    corrige en silencio: quien llama avisa de cada cambio.
    """
    arreglos = []

    # 1. Barras invertidas sueltas dentro de textos. En JSON las unicas
    #    combinaciones validas son \" \\ \/ \b \f \n \r \t y \uXXXX;
    #    cualquier otra hay que duplicarla. Fuera de los textos no hay
    #    barras invertidas, asi que se puede mirar el archivo entero.
    #
    #    La alternancia importa: hay que CONSUMIR las parejas validas
    #    enteras. Si solo se buscara "barra que no inicia escape valido",
    #    la segunda barra de un "\\" bien escrito se leeria como suelta y
    #    se romperia un archivo que estaba bien.
    sueltas = 0

    def _arreglar_barra(m):
        nonlocal sueltas
        if m.group(1) is not None:
            return m.group(0)  # pareja valida: se deja igual
        sueltas += 1
        return "\\\\"

    texto_nuevo = re.sub(r'\\(["\\/bfnrt]|u[0-9a-fA-F]{4})|\\', _arreglar_barra, texto)
    if sueltas:
        arreglos.append(f"{sueltas} barra(s) invertida(s) de una ruta de Windows")
        texto = texto_nuevo

    # 2. Numeros con puntos de miles.
    def _sin_puntos(m):
        return m.group(1) + m.group(2).replace(".", "")

    texto_nuevo, cambios = re.subn(
        r'(:\s*)(\d{1,3}(?:\.\d{3})+)(?=\s*[,}\]\r\n])', _sin_puntos, texto)
    if cambios:
        arreglos.append(f"{cambios} numero(s) escrito(s) con puntos de miles")
        texto = texto_nuevo

    # 3. Comas que faltan al final de un renglon. Se mira el renglon
    #    siguiente: si empieza otro valor, es que falto la coma.
    lineas = texto.split("\n")
    faltantes = 0
    for i, linea in enumerate(lineas[:-1]):
        actual = linea.rstrip()
        if not actual or actual.rstrip().endswith((",", "[", "{", ":")):
            continue
        if not actual.rstrip().endswith(('"', "]", "}")) and \
           not re.search(r"(\d|true|false|null)\s*$", actual):
            continue
        siguiente = next((l.strip() for l in lineas[i + 1:] if l.strip()), "")
        if siguiente.startswith(('"', "{", "[")):
            lineas[i] = actual + ","
            faltantes += 1
    if faltantes:
        arreglos.append(f"{faltantes} coma(s) que faltaban al final de un renglon")
        texto = "\n".join(lineas)

    # 4. La coma de mas antes de cerrar una lista o el archivo. Es el otro
    #    error clasico, el espejo del anterior.
    texto_nuevo, cambios = re.subn(r",(\s*[}\]])", r"\1", texto)
    if cambios:
        arreglos.append(f"{cambios} coma(s) de sobra antes de cerrar una lista")
        texto = texto_nuevo

    return texto, arreglos


def cargar_perfil() -> dict:
    if not ARCHIVO_PERFIL.exists():
        raise ErrorConfiguracion(
            f"No encuentro tu perfil: {ARCHIVO_PERFIL.name}\n"
            f"   Copia 'perfil_laboral.example.json', renombralo a "
            f"'perfil_laboral.json' y llenalo con tus datos."
        )

    crudo = ARCHIVO_PERFIL.read_text(encoding="utf-8")
    try:
        datos = json.loads(crudo)
    except json.JSONDecodeError as error_original:
        datos = _rescatar_perfil(crudo, error_original)

    perfil = dict(PERFIL_POR_DEFECTO)
    perfil.update(datos)

    if not perfil["cargos_objetivo"]:
        raise ErrorConfiguracion(
            "Tu perfil no tiene 'cargos_objetivo'. Sin eso no hay que buscar. "
            "Pon ahi los cargos tal como los publican, ej: "
            '["auxiliar juridico", "asistente administrativo"].'
        )
    if not perfil["nombre_completo"]:
        raise ErrorConfiguracion("Tu perfil no tiene 'nombre_completo'.")

    hoja = perfil.get("ruta_hoja_de_vida", "")
    if hoja:
        ruta = Path(os.path.expandvars(os.path.expanduser(hoja)))
        if not ruta.is_absolute():
            ruta = CARPETA_BASE / ruta
        perfil["_ruta_hoja_de_vida"] = ruta
        if not ruta.exists():
            LOG.warning(
                "AVISO: no encuentro tu hoja de vida en '%s'. Sin ella NO se "
                "puede postular por correo (si por portal).", ruta
            )
    else:
        perfil["_ruta_hoja_de_vida"] = None

    # Se precalcula todo lo que se compara mil veces, ya normalizado.
    perfil["_cargos"] = [sin_tildes(c) for c in perfil["cargos_objetivo"] if c]
    perfil["_deseables"] = [sin_tildes(c) for c in perfil["palabras_clave_deseables"] if c]
    perfil["_excluyentes"] = [sin_tildes(c) for c in perfil["palabras_excluyentes"] if c]
    ciudades = list(perfil["ciudades"]) or ([perfil["ciudad"]] if perfil["ciudad"] else [])
    perfil["_ciudades"] = [sin_tildes(c) for c in ciudades if c]
    perfil["_nivel"] = nivel_a_numero(perfil["nivel_educativo"])
    aplicar_ajustes_del_perfil(perfil)
    return perfil


def _rescatar_perfil(crudo: str, error_original: json.JSONDecodeError) -> dict:
    """Intenta salvar un perfil mal escrito; si no puede, explica donde.

    Cuando logra arreglarlo, GUARDA el archivo ya corregido y deja una
    copia del original en '.roto.bak'. Asi el usuario no tiene que
    arreglar lo mismo cada vez que arranca.
    """
    reparado, arreglos = reparar_perfil_escrito_a_mano(crudo)
    if arreglos:
        try:
            datos = json.loads(reparado)
        except json.JSONDecodeError:
            datos = None
        if datos is not None:
            LOG.warning("")
            LOG.warning("Tu %s tenia errores de escritura. Los corregi:", ARCHIVO_PERFIL.name)
            for arreglo in arreglos:
                LOG.warning("   - %s", arreglo)
            respaldo = ARCHIVO_PERFIL.with_suffix(".roto.bak")
            try:
                respaldo.write_text(crudo, encoding="utf-8")
                ARCHIVO_PERFIL.write_text(reparado, encoding="utf-8")
                LOG.warning("   Ya quedo guardado corregido (tu version original "
                            "quedo en %s).", respaldo.name)
            except OSError as e:
                LOG.warning("   No pude guardar la correccion (%s); sigo con ella "
                            "solo por esta vez.", e)
            LOG.warning("")
            return datos

    # No se pudo: se senala el renglon exacto y se explica en cristiano.
    lineas = crudo.split("\n")
    culpable = lineas[error_original.lineno - 1].strip() if 0 < error_original.lineno <= len(lineas) else ""
    pistas = {
        "Invalid \\escape": "una ruta de Windows con una sola barra invertida "
                            "(hay que escribir C:\\\\Users\\\\... con barras dobles)",
        "Expecting ',' delimiter": "falta una coma al final del renglon anterior",
        "Expecting property name": "sobra una coma en el renglon anterior",
        "Expecting value": "un valor mal escrito: si es un numero va sin puntos "
                           "(1750000, no 1.750.000); si es texto va entre comillas",
        "Unterminated string": "falta la comilla que cierra el texto",
    }
    pista = next((p for clave, p in pistas.items() if clave in error_original.msg),
                 "revisa las comas y las comillas de ese renglon")

    raise ErrorConfiguracion(
        f"Tu {ARCHIVO_PERFIL.name} tiene un error de escritura en la LINEA "
        f"{error_original.lineno}:\n"
        f"      {culpable}\n"
        f"   Lo que parece: {pista}.\n"
        f"   (Regla practica: cada renglon de una lista lleva coma al final "
        f"menos el ultimo, y los numeros van sin puntos.)"
    )


def aplicar_ajustes_del_perfil(perfil: dict):
    """Deja que el perfil mande sobre los ajustes de CONFIGURACION.

    Existe para que la app del telefono pueda cambiar el umbral, el tope
    diario o el modo sin editar este archivo: los guarda en
    perfil_laboral.json y aqui se aplican. Si el perfil no los trae,
    quedan los valores de CONFIGURACION tal cual.
    """
    global UMBRAL_POSTULACION, MAX_POSTULACIONES_DIA, INTERVALO_VIGILANCIA_MIN, MODO

    ajustes = {
        "umbral_postulacion": ("UMBRAL_POSTULACION", int, 0, 100),
        "max_postulaciones_dia": ("MAX_POSTULACIONES_DIA", int, 1, 200),
        "intervalo_vigilancia_min": ("INTERVALO_VIGILANCIA_MIN", int, 15, 1440),
    }
    for clave, (nombre, tipo, minimo, maximo) in ajustes.items():
        if perfil.get(clave) in (None, ""):
            continue
        try:
            valor = tipo(perfil[clave])
        except (TypeError, ValueError):
            LOG.warning("Ignoro '%s' del perfil: '%s' no es un numero.", clave, perfil[clave])
            continue
        if not (minimo <= valor <= maximo):
            LOG.warning("Ignoro '%s' del perfil: %s esta fuera de %s-%s.",
                        clave, valor, minimo, maximo)
            continue
        globals()[nombre] = valor

    modo = str(perfil.get("modo", "")).strip().lower()
    if modo:
        if modo in ("simulacion", "semiautomatico", "automatico"):
            MODO = modo
        else:
            LOG.warning("Ignoro 'modo' del perfil: '%s' no es un modo valido.", modo)


def cargar_credenciales():
    """Lee credenciales_empleo.txt. Devuelve dict o None si no existe."""
    if not ARCHIVO_CREDENCIALES.exists():
        return None
    datos = {}
    with open(ARCHIVO_CREDENCIALES, encoding="utf-8") as f:
        for linea in f:
            if "=" in linea and not linea.strip().startswith("#"):
                clave, _, valor = linea.partition("=")
                datos[clave.strip().upper()] = valor.strip()

    # Google muestra la contrasena de aplicacion en cuatro grupos de
    # cuatro letras ("abcd efgh ijkl mnop") y casi todo el mundo la pega
    # tal cual. Con los espacios adentro, el servidor la rechaza y el
    # error que se ve es "clave incorrecta", que manda a buscar el
    # problema donde no esta. Se quitan aqui, de una vez.
    if datos.get("CORREO_APP_PASSWORD"):
        datos["CORREO_APP_PASSWORD"] = re.sub(r"\s+", "", datos["CORREO_APP_PASSWORD"])

    if not datos.get("CORREO_USUARIO") or not datos.get("CORREO_APP_PASSWORD"):
        LOG.warning(
            "AVISO: %s existe pero le falta CORREO_USUARIO o CORREO_APP_PASSWORD. "
            "No se podran mandar postulaciones por correo.", ARCHIVO_CREDENCIALES.name
        )
        return None
    datos.setdefault("SMTP_SERVIDOR", "smtp.gmail.com")
    datos.setdefault("SMTP_PUERTO", "465")
    return datos


# ============================================================
# REGISTRO (que se vio, que se postulo, cuando)
# ============================================================

class Registro:
    """Memoria del programa: sin esto postularia dos veces a lo mismo."""

    def __init__(self, ruta: Path = ARCHIVO_REGISTRO):
        self.ruta = ruta
        self.datos = {}
        if ruta.exists():
            try:
                with open(ruta, encoding="utf-8") as f:
                    self.datos = json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                LOG.warning("No pude leer %s (%s). Empiezo un registro nuevo.", ruta.name, e)
                self.datos = {}

    def conoce(self, clave: str) -> bool:
        return clave in self.datos

    def estado(self, clave: str) -> str:
        return self.datos.get(clave, {}).get("estado", "")

    def anotar(self, clave: str, **campos):
        entrada = self.datos.setdefault(clave, {})
        entrada.update(campos)
        entrada["actualizado"] = datetime.datetime.now().isoformat(timespec="seconds")
        self.guardar()

    def guardar(self):
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        temporal = self.ruta.with_suffix(".tmp")
        with open(temporal, "w", encoding="utf-8") as f:
            json.dump(self.datos, f, ensure_ascii=False, indent=2)
        os.replace(temporal, self.ruta)

    def postuladas_hoy(self) -> int:
        hoy = datetime.date.today().isoformat()
        return sum(
            1 for e in self.datos.values()
            if e.get("estado") == "postulada" and str(e.get("actualizado", "")).startswith(hoy)
        )

    def ya_postule_a_esta_empresa(self, empresa: str, titulo: str) -> bool:
        """Evita postular dos veces al mismo cargo publicado en dos portales."""
        firma = (sin_tildes(empresa), sin_tildes(titulo))
        if not firma[0] and not firma[1]:
            return False
        for e in self.datos.values():
            if e.get("estado") != "postulada":
                continue
            if (sin_tildes(e.get("empresa", "")), sin_tildes(e.get("titulo", ""))) == firma:
                return True
        return False


# ============================================================
# VACANTE Y EVALUACION
# ============================================================

class Vacante:
    def __init__(self, portal, titulo, url, empresa="", ciudad="", resumen="", fecha=""):
        self.portal = portal
        self.titulo = limpiar_espacios(titulo)
        self.url = url
        self.empresa = limpiar_espacios(empresa)
        self.ciudad = limpiar_espacios(ciudad)
        self.resumen = limpiar_espacios(resumen)
        self.fecha = limpiar_espacios(fecha)
        self.descripcion = ""
        self.correos = []

    @property
    def clave(self) -> str:
        """Identidad estable de la oferta: la URL sin parametros de rastreo."""
        base = self.url.split("?")[0].split("#")[0].rstrip("/")
        return f"{self.portal}:{base}"

    @property
    def texto_completo(self) -> str:
        return sin_tildes(" ".join([self.titulo, self.empresa, self.ciudad,
                                    self.resumen, self.descripcion]))

    def __repr__(self):
        return f"<Vacante {self.portal} {self.titulo!r} @ {self.empresa!r}>"


class Evaluacion:
    def __init__(self):
        self.puntaje = 0
        self.a_favor = []
        self.en_contra = []
        self.descartada = False
        self.motivo_descarte = ""

    def descartar(self, motivo: str):
        self.descartada = True
        self.motivo_descarte = motivo
        self.puntaje = 0

    def sumar(self, puntos: int, motivo: str):
        self.puntaje += puntos
        self.a_favor.append(motivo)

    def restar(self, puntos: int, motivo: str):
        self.puntaje -= puntos
        self.en_contra.append(motivo)

    def resumen(self) -> str:
        if self.descartada:
            return f"DESCARTADA: {self.motivo_descarte}"
        partes = ", ".join(self.a_favor[:4])
        if self.en_contra:
            partes += " | contra: " + ", ".join(self.en_contra[:2])
        return f"{self.puntaje}/100 ({partes})"


def evaluar_vacante(vac: Vacante, perfil: dict) -> Evaluacion:
    """Que tanto encaja esta vacante contigo, de 0 a 100.

    La idea no es "cuanto me gusta" sino "que opcion real tengo de que
    me llamen": por eso lo que primero se mira son los requisitos
    DUROS (experiencia, estudios, ciudad, salario). Si no los cumples,
    se descarta de una -- mandar la hoja de vida ahi solo gasta tu
    cupo del dia.
    """
    ev = Evaluacion()
    texto = vac.texto_completo
    titulo = sin_tildes(vac.titulo)

    # --- 1. Descartes duros -------------------------------------
    for palabra in perfil["_excluyentes"]:
        if contiene_palabra(texto, palabra):
            ev.descartar(f"trae tu palabra excluyente '{palabra}'")
            return ev

    dias = antiguedad_en_dias(vac.fecha)
    if dias is not None and dias > MAX_DIAS_ANTIGUEDAD:
        ev.descartar(f"publicada hace {dias} dias (tope: {MAX_DIAS_ANTIGUEDAD})")
        return ev

    pide_anos = experiencia_pedida(texto)
    mis_anos = int(perfil.get("anos_experiencia") or 0)
    if pide_anos is not None and pide_anos > mis_anos:
        # Un ano de diferencia todavia se puede negociar; mas, no.
        if pide_anos - mis_anos > 1:
            ev.descartar(f"pide {pide_anos} anos de experiencia y tienes {mis_anos}")
            return ev
        ev.restar(10, f"pide {pide_anos} anos y tienes {mis_anos}")

    pide_nivel = nivel_pedido(texto)
    if pide_nivel is not None and pide_nivel > perfil["_nivel"] + 1:
        ev.descartar(
            f"pide un nivel educativo mas alto que el tuyo "
            f"({perfil['nivel_educativo']})"
        )
        return ev

    remota = any(p in texto for p in ("remoto", "teletrabajo", "home office", "desde casa", "hibrido"))
    if perfil["_ciudades"]:
        coincide_ciudad = any(contiene_palabra(texto, c) for c in perfil["_ciudades"])
        if not coincide_ciudad and not (remota and perfil.get("acepta_remoto", True)):
            ev.descartar(f"queda en '{vac.ciudad or 'otra ciudad'}' y no trabajas alli")
            return ev

    minimo = int(perfil.get("salario_minimo") or 0)
    if minimo:
        paga = salario_ofrecido(vac.resumen) or salario_ofrecido(vac.descripcion)
        if paga is not None and paga < minimo:
            ev.descartar(f"ofrece ${paga:,.0f} y tu minimo es ${minimo:,.0f}".replace(",", "."))
            return ev

    # --- 2. Puntos a favor --------------------------------------
    cargo_en_titulo = any(contiene_palabra(titulo, c) for c in perfil["_cargos"])
    cargo_en_texto = any(contiene_palabra(texto, c) for c in perfil["_cargos"])
    if cargo_en_titulo:
        ev.sumar(45, "el cargo del titulo es uno de los que buscas")
    elif cargo_en_texto:
        ev.sumar(20, "tu cargo aparece en la descripcion, no en el titulo")
    else:
        ev.restar(10, "el cargo no es de los que buscas")

    aciertos = [p for p in perfil["_deseables"] if contiene_palabra(texto, p)]
    if aciertos:
        ev.sumar(min(20, 5 * len(aciertos)),
                 f"coincide en {len(aciertos)} palabras tuyas ({', '.join(aciertos[:3])})")

    if perfil["_ciudades"] and any(contiene_palabra(texto, c) for c in perfil["_ciudades"]):
        ev.sumar(15, "esta en tu ciudad")
    elif remota:
        ev.sumar(12, "es remota/hibrida")

    if pide_anos is not None and pide_anos <= mis_anos:
        ev.sumar(15, f"pide {pide_anos} anos y tienes {mis_anos}")
    elif pide_anos is None:
        ev.sumar(5, "no exige anos de experiencia explicitos")

    if pide_nivel is not None and pide_nivel <= perfil["_nivel"]:
        ev.sumar(10, "cumples el nivel educativo")

    if vac.correos:
        ev.sumar(5, "publica correo de contacto (se puede postular directo)")

    if dias is not None and dias <= 3:
        ev.sumar(5, "recien publicada")

    ev.puntaje = max(0, min(100, ev.puntaje))
    return ev


# ============================================================
# NAVEGADOR
# ============================================================

class Navegador:
    """Envoltura de Playwright con la sesion del usuario guardada.

    Se usa un perfil PERSISTENTE (CARPETA_NAVEGADOR) para que las
    sesiones que inicies a mano con --login sigan sirviendo despues.
    """

    def __init__(self, visible: bool = None):
        self.visible = NAVEGADOR_VISIBLE if visible is None else visible
        self._pw = None
        self.contexto = None
        self.pagina = None

    def __enter__(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise ErrorConfiguracion(
                "Falta Playwright. Instalalo con:\n"
                "   pip install -r requirements.txt\n"
                "   playwright install chromium"
            )
        CARPETA_NAVEGADOR.mkdir(parents=True, exist_ok=True)
        self._pw = sync_playwright().start()
        # Varios portales rechazan a los navegadores automatizados, y
        # cuando lo hacen el sintoma es "0 resultados", que se confunde
        # con "no hay ofertas". Estas dos cosas -- ocultar la senal de
        # automatizacion y presentarse con un identificador normal --
        # evitan la mayoria de esos rechazos. No evaden captchas ni
        # bloqueos: solo dejan de anunciar que esto es un programa.
        self.contexto = self._pw.chromium.launch_persistent_context(
            user_data_dir=str(CARPETA_NAVEGADOR),
            headless=not self.visible,
            locale="es-CO",
            timezone_id="America/Bogota",
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
            args=["--disable-blink-features=AutomationControlled"],
        )
        self.contexto.set_default_timeout(TIMEOUT_PAGINA_MS)
        self.pagina = self.contexto.pages[0] if self.contexto.pages else self.contexto.new_page()
        return self

    def __exit__(self, *_):
        try:
            if self.contexto:
                self.contexto.close()
        finally:
            if self._pw:
                self._pw.stop()

    def ir_a(self, url: str) -> bool:
        try:
            self.pagina.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_PAGINA_MS)
            self.pagina.wait_for_timeout(2500)
            return True
        except Exception as e:
            LOG.debug("No pude abrir %s (%s)", url, type(e).__name__)
            return False

    def texto_de(self, selectores) -> str:
        """Primer selector de la lista que devuelva texto. '' si ninguno."""
        for sel in selectores:
            try:
                elemento = self.pagina.query_selector(sel)
                if elemento:
                    texto = limpiar_espacios(elemento.inner_text())
                    if texto:
                        return texto
            except Exception:
                continue
        return ""


def _texto_hijo(tarjeta, selectores) -> str:
    for sel in selectores:
        try:
            hijo = tarjeta.query_selector(sel)
            if hijo:
                texto = limpiar_espacios(hijo.inner_text())
                if texto:
                    return texto
        except Exception:
            continue
    return ""


def construir_url(plantilla: str, termino: str, ciudad: str, pagina: int) -> str:
    """Arma una direccion de busqueda a partir de una plantilla.

    Se ofrecen las dos formas del termino porque los portales no se
    ponen de acuerdo: unos lo quieren dentro de la ruta separado por
    guiones (trabajo-de-auxiliar-juridico) y otros como parametro
    (?q=auxiliar+juridico).
    """
    from urllib.parse import quote_plus
    return plantilla.format(
        q=a_slug(termino),
        qp=quote_plus(termino),
        ciudad=a_slug(ciudad),
        ciudadp=quote_plus(ciudad),
        pagina=pagina,
        pagina0_25=(pagina - 1) * 25,
    )


SENALES_DE_BLOQUEO = (
    "just a moment", "verificando tu navegador", "checking your browser",
    "acceso denegado", "access denied", "unusual traffic", "captcha",
    "no eres un robot", "are you a human", "cloudflare",
)


def parece_bloqueo(nav) -> bool:
    """Distingue 'no hay ofertas' de 'el portal no me dejo entrar'.

    Sin esto, un bloqueo antibot se ve exactamente igual que una
    busqueda sin resultados, y se pierde el tiempo buscando el problema
    en la direccion cuando esta bien.
    """
    try:
        texto = sin_tildes(nav.pagina.title() + " " + nav.pagina.inner_text("body")[:1500])
    except Exception:
        return False
    return any(s in texto for s in SENALES_DE_BLOQUEO)


def elegir_plantilla(nav: Navegador, clave_fuente: str, cfg: dict, termino: str, ciudad: str):
    """Prueba las direcciones candidatas y devuelve la que si trae ofertas.

    Los portales cambian la forma de sus direcciones cada tanto, y
    cuando eso pasa el sintoma es siempre el mismo: 0 resultados en
    todo. En vez de quedarse ahi, se prueban las alternativas conocidas
    y se sigue con la que responda.

    Devuelve (plantilla, vacantes_de_esa_primera_busqueda) o (None, []).
    """
    bloqueado = False
    for plantilla in cfg["plantillas_url"]:
        url = construir_url(plantilla, termino, ciudad, 1)
        if not nav.ir_a(url):
            LOG.debug("   no respondio: %s", url)
            continue

        if parece_bloqueo(nav):
            bloqueado = True
            LOG.debug("   el portal mostro una pantalla de verificacion en %s", url)
            continue

        vacantes = _leer_resultados(nav, clave_fuente, cfg)
        if vacantes:
            return plantilla, vacantes
        LOG.debug("   0 resultados con: %s", url)
        time.sleep(random.uniform(*ESPERA_ENTRE_LECTURAS))

    if bloqueado:
        LOG.warning("   %s no dejo entrar: mostro una pantalla de verificacion "
                    "(captcha o similar).", cfg["nombre"])
        LOG.warning("   Suele pasar con el navegador oculto. Abre buscar_empleo.py "
                    "y pon NAVEGADOR_VISIBLE = True.")
    else:
        LOG.warning("   %s no devolvio ofertas con ninguna de sus %d direcciones.",
                    cfg["nombre"], len(cfg["plantillas_url"]))
        LOG.warning("   Corre 'python buscar_empleo.py --diagnostico' para ver que "
                    "esta mostrando el portal y con que direccion.")
    return None, []


def buscar_en_portal(nav: Navegador, clave_fuente: str, cfg: dict, perfil: dict):
    """Devuelve las vacantes que ese portal muestra para tus cargos."""
    encontradas = {}
    ciudad = perfil.get("ciudad", "")
    terminos = list(perfil["cargos_objetivo"])

    # Con el primer termino se averigua cual de las direcciones sirve; el
    # resto de la busqueda ya va derecho por esa.
    plantilla, primeras = elegir_plantilla(nav, clave_fuente, cfg, terminos[0], ciudad)
    if plantilla is None:
        return []
    for vac in primeras:
        encontradas.setdefault(vac.clave, vac)
    if plantilla is not cfg["plantillas_url"][0]:
        LOG.info("   (usando la direccion alterna: %s)", plantilla)

    for indice, termino in enumerate(terminos):
        for pagina in range(1, PAGINAS_POR_BUSQUEDA + 1):
            if indice == 0 and pagina == 1:
                continue  # esa ya se leyo al elegir la direccion
            url = construir_url(plantilla, termino, ciudad, pagina)
            LOG.debug("   %s: buscando '%s' (pag. %d)", cfg["nombre"], termino, pagina)
            if not nav.ir_a(url):
                LOG.warning("   %s no respondio en %s", cfg["nombre"], url)
                break

            nuevas = _leer_resultados(nav, clave_fuente, cfg)
            for vac in nuevas:
                encontradas.setdefault(vac.clave, vac)

            if not nuevas:
                LOG.debug("   %s: 0 resultados en pag. %d", cfg["nombre"], pagina)
                break
            time.sleep(random.uniform(*ESPERA_ENTRE_LECTURAS))

    return list(encontradas.values())


def _leer_resultados(nav: Navegador, clave_fuente: str, cfg: dict):
    """Saca las tarjetas de resultados de la pagina que este abierta.

    Primero intenta con los selectores del portal (que traen empresa y
    ciudad ya separadas). Si el portal cambio su HTML y ningun selector
    pega, cae al plan B: recoger todos los links que tengan la forma de
    una oferta. Feo pero infalible: los links son lo ultimo que cambia.
    """
    vacantes = []
    vistas = set()

    for sel in cfg["sel_tarjeta"]:
        try:
            tarjetas = nav.pagina.query_selector_all(sel)
        except Exception:
            continue
        if not tarjetas:
            continue

        for tarjeta in tarjetas:
            try:
                enlace = None
                for sel_t in cfg["sel_titulo"]:
                    posible = tarjeta.query_selector(sel_t)
                    if posible and posible.get_attribute("href"):
                        enlace = posible
                        break
                if enlace is None:
                    enlace = tarjeta.query_selector("a[href]")
                if enlace is None:
                    continue

                href = enlace.get_attribute("href") or ""
                url = _absoluta(href, nav.pagina.url)
                if not url or url in vistas:
                    continue
                if not re.search(cfg["patron_enlace"], url):
                    continue
                vistas.add(url)

                titulo = limpiar_espacios(enlace.inner_text()) or _texto_hijo(tarjeta, cfg["sel_titulo"])
                texto_tarjeta = limpiar_espacios(tarjeta.inner_text())
                vacantes.append(Vacante(
                    portal=clave_fuente,
                    titulo=titulo,
                    url=url,
                    empresa=_texto_hijo(tarjeta, cfg["sel_empresa"]),
                    ciudad=_texto_hijo(tarjeta, cfg["sel_ciudad"]),
                    resumen=texto_tarjeta[:800],
                    fecha=texto_tarjeta[:300],
                ))
            except Exception:
                continue

        if vacantes:
            return vacantes

    # Plan B: solo los links.
    try:
        for enlace in nav.pagina.query_selector_all("a[href]"):
            href = enlace.get_attribute("href") or ""
            url = _absoluta(href, nav.pagina.url)
            if not url or url in vistas or not re.search(cfg["patron_enlace"], url):
                continue
            vistas.add(url)
            titulo = limpiar_espacios(enlace.inner_text())
            if len(titulo) < 4:
                continue
            vacantes.append(Vacante(portal=clave_fuente, titulo=titulo, url=url))
    except Exception as e:
        LOG.debug("Plan B de lectura fallo: %s", e)

    if vacantes:
        LOG.debug("   %s: lei %d ofertas con el plan B (revisa los selectores "
                  "en FUENTES si esto se repite)", cfg["nombre"], len(vacantes))
    return vacantes


def _absoluta(href: str, url_actual: str) -> str:
    from urllib.parse import urljoin
    href = (href or "").strip()
    if not href or href.startswith(("javascript:", "mailto:", "#")):
        return ""
    return urljoin(url_actual, href)


def leer_detalle(nav: Navegador, vac: Vacante, cfg: dict) -> bool:
    """Abre la oferta y le lee la descripcion y el correo de contacto."""
    if not nav.ir_a(vac.url):
        return False
    texto = nav.texto_de(cfg["sel_descripcion"])
    if not texto:
        try:
            texto = limpiar_espacios(nav.pagina.inner_text("body"))
        except Exception:
            texto = ""
    vac.descripcion = texto[:12000]

    # El correo puede estar en el texto o escondido en un mailto:.
    encontrados = correos_en_texto(vac.descripcion)
    try:
        for enlace in nav.pagina.query_selector_all("a[href^='mailto:']"):
            href = enlace.get_attribute("href") or ""
            encontrados.extend(correos_en_texto(href.replace("mailto:", "")))
    except Exception:
        pass
    vac.correos = list(dict.fromkeys(encontrados))

    if not vac.empresa:
        vac.empresa = nav.texto_de(cfg["sel_empresa"])
    if not vac.ciudad:
        vac.ciudad = nav.texto_de(cfg["sel_ciudad"])
    return True


# ============================================================
# POSTULACION
# ============================================================

def armar_carta(vac: Vacante, perfil: dict) -> str:
    """La carta de presentacion, ya con los datos de ESA vacante."""
    plantilla = perfil.get("carta_presentacion") or (
        "Buen dia,\n\n"
        "Me permito postularme a la vacante de {titulo}{en_empresa}. "
        "{resumen}\n\n"
        "Adjunto mi hoja de vida para su consideracion. Quedo atento(a) a "
        "cualquier informacion adicional que requieran.\n\n"
        "Cordial saludo,\n"
        "{nombre}\n"
        "{telefono}\n"
        "{correo}"
    )
    valores = {
        "titulo": vac.titulo or "la vacante publicada",
        "empresa": vac.empresa or "su empresa",
        "en_empresa": f" en {vac.empresa}" if vac.empresa else "",
        "ciudad": vac.ciudad or perfil.get("ciudad", ""),
        "portal": FUENTES.get(vac.portal, {}).get("nombre", vac.portal),
        "nombre": perfil.get("nombre_completo", ""),
        "telefono": perfil.get("telefono", ""),
        "correo": perfil.get("correo", ""),
        "resumen": perfil.get("resumen_profesional", ""),
    }
    try:
        return plantilla.format(**valores)
    except KeyError as e:
        LOG.warning("Tu 'carta_presentacion' usa un dato que no existe: %s. "
                    "Se manda sin reemplazar.", e)
        return plantilla


def postular_por_correo(vac: Vacante, perfil: dict, cred: dict, simular: bool):
    """Manda la hoja de vida al correo que publico la oferta.

    Devuelve (exito, detalle). Esta es la via mas limpia: no depende
    de ningun portal ni de ninguna sesion.
    """
    destino = vac.correos[0]
    hoja = perfil.get("_ruta_hoja_de_vida")
    if not hoja or not Path(hoja).exists():
        return False, "no tengo tu hoja de vida en disco"

    asunto = (perfil.get("asunto_correo") or "Postulacion: {titulo}").format(
        titulo=vac.titulo or "vacante",
        empresa=vac.empresa or "",
        nombre=perfil.get("nombre_completo", ""),
    )

    if simular:
        return True, f"[SIMULACION] correo a {destino} - asunto: {asunto}"

    mensaje = EmailMessage()
    mensaje["From"] = cred["CORREO_USUARIO"]
    mensaje["To"] = destino
    mensaje["Subject"] = asunto
    if perfil.get("correo") and perfil["correo"] != cred["CORREO_USUARIO"]:
        mensaje["Reply-To"] = perfil["correo"]
    mensaje.set_content(armar_carta(vac, perfil))

    tipo, _ = mimetypes.guess_type(str(hoja))
    principal, _, secundario = (tipo or "application/octet-stream").partition("/")
    with open(hoja, "rb") as f:
        mensaje.add_attachment(f.read(), maintype=principal, subtype=secundario or "octet-stream",
                               filename=Path(hoja).name)

    exito, detalle = _entregar_correo(cred, mensaje)
    if not exito:
        return False, detalle
    return True, f"hoja de vida enviada a {destino}"


def _entregar_correo(cred: dict, mensaje: EmailMessage):
    """Entrega un correo ya armado por SMTP. Devuelve (exito, detalle).

    Lo usan las postulaciones y los avisos que te llegan al celular, para
    no repetir en dos lados el manejo de puertos y errores.
    """
    contexto = ssl.create_default_context()
    puerto = int(cred.get("SMTP_PUERTO", 465))
    servidor = cred.get("SMTP_SERVIDOR", "smtp.gmail.com")
    try:
        if puerto == 587:
            with smtplib.SMTP(servidor, puerto, timeout=60) as smtp:
                smtp.starttls(context=contexto)
                smtp.login(cred["CORREO_USUARIO"], cred["CORREO_APP_PASSWORD"])
                smtp.send_message(mensaje)
        else:
            with smtplib.SMTP_SSL(servidor, puerto, context=contexto, timeout=60) as smtp:
                smtp.login(cred["CORREO_USUARIO"], cred["CORREO_APP_PASSWORD"])
                smtp.send_message(mensaje)
    except smtplib.SMTPAuthenticationError:
        return False, ("el correo rechazo la clave. Recuerda que debe ser una "
                       "CONTRASENA DE APLICACION, no tu clave normal")
    except Exception as e:
        return False, f"error mandando el correo: {type(e).__name__}: {e}"
    return True, "enviado"


def avisar_al_celular(perfil: dict, cred, postuladas: list, pendientes: list):
    """Te manda a TU correo el resumen de lo que acaba de pasar.

    Es la forma de enterarte sin abrir nada: al celular le llega la
    notificacion del correo como cualquier otra. Solo se manda cuando
    hay algo que contar -- un aviso de "no encontre nada" cada dos horas
    se vuelve ruido y se termina silenciando, que es justo lo que no
    queremos.
    """
    if not AVISAR_POR_CORREO or not cred:
        return
    if not postuladas and not pendientes:
        return

    destino = perfil.get("correo") or cred["CORREO_USUARIO"]

    partes = []
    if postuladas:
        partes.append(f"{len(postuladas)} postulacion(es) enviada(s)")
    if pendientes:
        partes.append(f"{len(pendientes)} te esperan")
    asunto = "Empleo: " + ", ".join(partes)

    lineas = []
    if postuladas:
        lineas.append("YA SE POSTULO A:")
        for vac, detalle in postuladas:
            lineas.append(f"  - {vac.titulo}")
            lineas.append(f"    {vac.empresa or 'empresa no publicada'} - {detalle}")
            lineas.append(f"    {vac.url}")
        lineas.append("")
    if pendientes:
        lineas.append("TE TOCA A TI (el portal no dejo terminar solo):")
        for vac, detalle in pendientes:
            lineas.append(f"  - {vac.titulo}")
            lineas.append(f"    {vac.empresa or 'empresa no publicada'} - {detalle}")
            lineas.append(f"    {vac.url}")
        lineas.append("")
    lineas.append("-- Enviado por tu buscador de empleo, desde tu PC.")

    mensaje = EmailMessage()
    mensaje["From"] = cred["CORREO_USUARIO"]
    mensaje["To"] = destino
    mensaje["Subject"] = asunto
    mensaje.set_content("\n".join(lineas))

    exito, detalle = _entregar_correo(cred, mensaje)
    if exito:
        LOG.info(" Aviso enviado a %s", destino)
    else:
        LOG.warning(" No se pudo mandar el aviso: %s", detalle)


TEXTOS_ENVIAR = ("enviar postulacion", "enviar solicitud", "enviar", "confirmar",
                 "finalizar", "submit", "send application")


def postular_en_portal(nav: Navegador, vac: Vacante, cfg: dict, modo: str):
    """Postula con el boton del propio portal.

    Es deliberadamente conservador: si el formulario pide algo que no
    reconoce (preguntas del empleador, pruebas, subir documentos), NO
    inventa respuestas -- deja la vacante marcada como
    'pendiente_revision' con su link, para que la termines tu en dos
    clics. Una respuesta inventada en un formulario te quema la
    postulacion; dejarla a medias, no.
    """
    if not nav.ir_a(vac.url):
        return "error", "no pude abrir la oferta"

    boton = _buscar_boton(nav, cfg["sel_boton_postular"])
    if boton is None:
        return "pendiente_revision", "no encontre el boton de postularme"

    if modo == "simulacion":
        return "simulada", "[SIMULACION] habria hecho clic en postularme"

    try:
        boton.click()
        nav.pagina.wait_for_timeout(3500)
    except Exception as e:
        return "error", f"no pude hacer clic en postularme: {type(e).__name__}"

    # Muchos portales postulan de una con ese solo clic.
    if _parece_postulado(nav):
        return "postulada", "postulacion registrada en el portal"

    # Si abrio un formulario, se necesita mano humana salvo en automatico.
    if modo != "automatico":
        return "pendiente_revision", ("el portal abrio un formulario; terminalo tu "
                                      "(el link queda en el registro)")

    enviar = _buscar_boton(nav, TEXTOS_ENVIAR)
    if enviar is None:
        return "pendiente_revision", "el formulario pide datos que no reconozco"
    try:
        enviar.click()
        nav.pagina.wait_for_timeout(3500)
    except Exception as e:
        return "error", f"no pude enviar el formulario: {type(e).__name__}"

    if _parece_postulado(nav):
        return "postulada", "postulacion enviada en el portal"
    return "pendiente_revision", "envie el formulario pero el portal no confirmo"


def _buscar_boton(nav: Navegador, textos):
    """Busca un boton/enlace cuyo texto sea uno de `textos`."""
    for texto in textos:
        for etiqueta in ("button", "a", "input[type=submit]", "[role=button]"):
            try:
                elementos = nav.pagina.query_selector_all(etiqueta)
            except Exception:
                continue
            for elemento in elementos:
                try:
                    if not elemento.is_visible():
                        continue
                    contenido = sin_tildes(elemento.inner_text() or
                                           elemento.get_attribute("value") or "")
                    if texto in contenido and len(contenido) < 60:
                        return elemento
                except Exception:
                    continue
    return None


def _parece_postulado(nav: Navegador) -> bool:
    """Reconoce el mensaje de confirmacion tipico de los portales."""
    señales = ("postulacion exitosa", "te has postulado", "ya te postulaste",
               "hoja de vida enviada", "solicitud enviada", "application sent",
               "ya aplicaste", "postulacion enviada", "gracias por postularte",
               "tu solicitud fue enviada")
    try:
        cuerpo = sin_tildes(nav.pagina.inner_text("body"))
    except Exception:
        return False
    return any(s in cuerpo for s in señales)


# ============================================================
# REPORTE EN EXCEL
# ============================================================

COLUMNAS_EXCEL = [
    ("actualizado", "Fecha"),
    ("estado", "Estado"),
    ("puntaje", "Puntaje"),
    ("titulo", "Cargo"),
    ("empresa", "Empresa"),
    ("ciudad", "Ciudad"),
    ("portal", "Portal"),
    ("metodo", "Como se postulo"),
    ("detalle", "Detalle"),
    ("url", "Link"),
]


def exportar_excel(registro: Registro) -> Path:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
    except ImportError:
        raise ErrorConfiguracion("Falta openpyxl (pip install -r requirements.txt)")

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Postulaciones"

    hoja.append([titulo for _, titulo in COLUMNAS_EXCEL])
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")

    entradas = sorted(registro.datos.values(),
                      key=lambda e: str(e.get("actualizado", "")), reverse=True)
    for entrada in entradas:
        hoja.append([entrada.get(campo, "") for campo, _ in COLUMNAS_EXCEL])

    anchos = [18, 18, 8, 42, 28, 18, 14, 16, 45, 60]
    for i, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[chr(64 + i)].width = ancho
    hoja.freeze_panes = "A2"

    ARCHIVO_EXCEL.parent.mkdir(parents=True, exist_ok=True)
    libro.save(ARCHIVO_EXCEL)
    return ARCHIVO_EXCEL


def mostrar_resumen(registro: Registro):
    conteo = {}
    for entrada in registro.datos.values():
        conteo[entrada.get("estado", "?")] = conteo.get(entrada.get("estado", "?"), 0) + 1
    LOG.info("")
    LOG.info("RESUMEN HISTORICO")
    LOG.info("-" * 46)
    for estado in sorted(conteo, key=lambda e: -conteo[e]):
        LOG.info("  %-22s %d", estado, conteo[estado])
    LOG.info("  %-22s %d", "postuladas hoy", registro.postuladas_hoy())
    LOG.info("-" * 46)


# ============================================================
# CICLO PRINCIPAL
# ============================================================

def una_pasada(perfil: dict, cred, registro: Registro, modo: str):
    """Una vuelta completa: buscar -> calificar -> postular."""
    simular = (modo == "simulacion")
    fuentes_activas = {k: v for k, v in FUENTES.items() if v.get("habilitada")}
    if not fuentes_activas:
        LOG.warning("No hay ningun portal habilitado en FUENTES.")
        return

    LOG.info("")
    LOG.info("=" * 62)
    LOG.info(" Revisando %d portales | modo: %s | umbral: %d puntos",
             len(fuentes_activas), modo.upper(), UMBRAL_POSTULACION)
    LOG.info("=" * 62)

    cupo = MAX_POSTULACIONES_DIA - registro.postuladas_hoy()
    if cupo <= 0 and not simular:
        LOG.info("Ya llegaste al tope de %d postulaciones hoy. Sigo manana.",
                 MAX_POSTULACIONES_DIA)
        return

    nuevas = descartadas = 0
    postuladas, pendientes = [], []  # (vacante, detalle), para el aviso al celular

    with Navegador() as nav:
        for clave, cfg in fuentes_activas.items():
            LOG.info("")
            LOG.info(">> %s", cfg["nombre"])
            try:
                vacantes = buscar_en_portal(nav, clave, cfg, perfil)
            except Exception as e:
                LOG.warning("   %s fallo por completo: %s: %s", cfg["nombre"], type(e).__name__, e)
                continue

            frescas = [v for v in vacantes if not registro.conoce(v.clave)]
            LOG.info("   %d ofertas vistas, %d nuevas", len(vacantes), len(frescas))

            for vac in frescas:
                nuevas += 1
                try:
                    leer_detalle(nav, vac, cfg)
                except Exception as e:
                    LOG.debug("   no pude leer %s (%s)", vac.url, type(e).__name__)

                ev = evaluar_vacante(vac, perfil)
                base = dict(portal=cfg["nombre"], titulo=vac.titulo, empresa=vac.empresa,
                            ciudad=vac.ciudad, url=vac.url, puntaje=ev.puntaje)

                if ev.descartada or ev.puntaje < UMBRAL_POSTULACION:
                    descartadas += 1
                    motivo = ev.motivo_descarte or f"puntaje {ev.puntaje} < {UMBRAL_POSTULACION}"
                    LOG.debug("   - %-45s  %s", vac.titulo[:45], motivo)
                    registro.anotar(vac.clave, estado="descartada", metodo="", detalle=motivo, **base)
                    continue

                if registro.ya_postule_a_esta_empresa(vac.empresa, vac.titulo):
                    LOG.info("   = %-45s  ya postulaste a esa misma vacante en otro portal",
                             vac.titulo[:45])
                    registro.anotar(vac.clave, estado="duplicada", metodo="",
                                    detalle="misma vacante ya postulada en otro portal", **base)
                    continue

                if cupo <= 0 and not simular:
                    LOG.info("   ! Cupo diario lleno; %s queda para manana", vac.titulo[:45])
                    registro.anotar(vac.clave, estado="en_espera", metodo="",
                                    detalle="cupo diario lleno", **base)
                    continue

                LOG.info("   + %-45s  %s", vac.titulo[:45], ev.resumen())
                estado, metodo, detalle = _postular(nav, vac, cfg, perfil, cred, modo)
                registro.anotar(vac.clave, estado=estado, metodo=metodo, detalle=detalle, **base)

                if estado == "postulada":
                    postuladas.append((vac, detalle))
                    cupo -= 1
                    LOG.info("     -> POSTULADO por %s: %s", metodo, detalle)
                    time.sleep(random.uniform(*ESPERA_ENTRE_POSTULACIONES))
                elif estado == "pendiente_revision":
                    pendientes.append((vac, detalle))
                    LOG.info("     -> TE TOCA A TI: %s", detalle)
                    LOG.info("        %s", vac.url)
                else:
                    LOG.info("     -> %s: %s", estado, detalle)

                time.sleep(random.uniform(*ESPERA_ENTRE_LECTURAS))

    LOG.info("")
    LOG.info("-" * 62)
    LOG.info(" Vacantes nuevas: %d | descartadas: %d | postuladas: %d | para ti: %d",
             nuevas, descartadas, len(postuladas), len(pendientes))
    if simular:
        LOG.info(" (MODO SIMULACION: no se mando nada de verdad)")
    LOG.info("-" * 62)

    if not simular:
        avisar_al_celular(perfil, cred, postuladas, pendientes)


def _postular(nav, vac, cfg, perfil, cred, modo):
    """Elige la via: correo si la oferta lo publica; si no, el portal."""
    simular = (modo == "simulacion")

    # En simulacion no hace falta tener credenciales para MOSTRAR que se
    # habria mandado el correo: es justo lo que se quiere ver en la prueba.
    if vac.correos and (cred or simular):
        exito, detalle = postular_por_correo(vac, perfil, cred, simular)
        if exito:
            return ("simulada" if simular else "postulada"), "correo", detalle
        LOG.debug("     correo fallo (%s); intento por el portal", detalle)
    elif vac.correos and not cred:
        LOG.debug("     la oferta trae correo pero no configuraste "
                  "credenciales_empleo.txt; intento por el portal")

    estado, detalle = postular_en_portal(nav, vac, cfg, modo)
    return estado, "portal", detalle


def probar_correo(perfil: dict, cred) -> int:
    """Manda un correo de prueba a ti mismo y dice claramente que paso.

    Existe para separar dos problemas que de otro modo se confunden: que
    el correo este mal configurado, y que la busqueda no encuentre nada.
    Si esto llega a tu bandeja, el envio de hojas de vida va a funcionar.
    """
    if not cred:
        LOG.error("")
        LOG.error("No encuentro %s (o le faltan datos).", ARCHIVO_CREDENCIALES.name)
        LOG.error("   Copia '%s.example.txt' -> '%s' y llenalo.",
                  ARCHIVO_CREDENCIALES.stem.replace(".example", ""), ARCHIVO_CREDENCIALES.name)
        return 1

    destino = perfil.get("correo") or cred["CORREO_USUARIO"]
    largo = len(cred["CORREO_APP_PASSWORD"])

    LOG.info("")
    LOG.info("Probando el correo...")
    LOG.info("  Cuenta que envia : %s", cred["CORREO_USUARIO"])
    LOG.info("  Servidor         : %s puerto %s",
             cred.get("SMTP_SERVIDOR"), cred.get("SMTP_PUERTO"))
    LOG.info("  Clave            : %d caracteres", largo)
    if largo != 16:
        LOG.warning("  OJO: una contrasena de aplicacion de Google tiene 16 letras. "
                    "La tuya tiene %d: revisa que no hayas pegado tu clave normal.", largo)
    LOG.info("  Enviando a       : %s", destino)
    LOG.info("")

    mensaje = EmailMessage()
    mensaje["From"] = cred["CORREO_USUARIO"]
    mensaje["To"] = destino
    mensaje["Subject"] = "Prueba de tu buscador de empleo"
    mensaje.set_content(
        "Si estas leyendo esto, el correo quedo bien configurado.\n\n"
        "Es el mismo camino por el que van a salir tus hojas de vida y los "
        "avisos de lo que se vaya postulando.\n\n"
        "-- Tu buscador de empleo."
    )

    hoja = perfil.get("_ruta_hoja_de_vida")
    if hoja and Path(hoja).exists():
        # Se adjunta la hoja de vida de verdad: asi la prueba tambien
        # comprueba que el archivo se lee y que no pesa de mas.
        tipo, _ = mimetypes.guess_type(str(hoja))
        principal, _, secundario = (tipo or "application/octet-stream").partition("/")
        with open(hoja, "rb") as f:
            mensaje.add_attachment(f.read(), maintype=principal,
                                   subtype=secundario or "octet-stream",
                                   filename=Path(hoja).name)
        LOG.info("  (va con tu hoja de vida adjunta, para probarla tambien)")

    exito, detalle = _entregar_correo(cred, mensaje)
    LOG.info("")
    if exito:
        LOG.info("=" * 62)
        LOG.info(" LISTO: revisa la bandeja de %s", destino)
        LOG.info(" Si llego, el correo ya esta bien y no hay que tocarlo mas.")
        LOG.info(" (Mira tambien en Spam la primera vez.)")
        LOG.info("=" * 62)
        return 0

    LOG.error("=" * 62)
    LOG.error(" NO SE PUDO ENVIAR")
    LOG.error(" %s", detalle)
    LOG.error("")
    LOG.error(" Lo que casi siempre lo causa:")
    LOG.error("   - Pegaste tu clave normal de Gmail en vez de una")
    LOG.error("     CONTRASENA DE APLICACION (son 16 letras que Google")
    LOG.error("     genera aparte, en myaccount.google.com/apppasswords).")
    LOG.error("   - La verificacion en 2 pasos no esta activada: sin ella")
    LOG.error("     Google no deja crear contrasenas de aplicacion.")
    LOG.error("   - El correo de CORREO_USUARIO no es el mismo con el que")
    LOG.error("     creaste la contrasena de aplicacion.")
    LOG.error("=" * 62)
    return 1


def vigilar(perfil, cred, registro, modo):
    LOG.info("Vigilancia encendida: reviso cada %d minutos. Ctrl+C para parar.",
             INTERVALO_VIGILANCIA_MIN)
    while True:
        try:
            una_pasada(perfil, cred, registro, modo)
            exportar_excel(registro)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            LOG.exception("Error en la pasada (sigo vigilando): %s", e)
        proxima = datetime.datetime.now() + datetime.timedelta(minutes=INTERVALO_VIGILANCIA_MIN)
        LOG.info("Proxima revision: %s", proxima.strftime("%H:%M"))
        time.sleep(INTERVALO_VIGILANCIA_MIN * 60)


def iniciar_sesiones(perfil):
    """Abre el navegador visible para que inicies sesion tu mismo.

    Las cookies quedan guardadas en CARPETA_NAVEGADOR y se reusan
    despues. El programa NUNCA escribe tus claves de los portales.
    """
    portales = [(k, v) for k, v in FUENTES.items()
                if v.get("habilitada") and v.get("necesita_sesion")]
    if not portales:
        LOG.info("Ningun portal habilitado necesita que inicies sesion.")
        return

    LOG.info("")
    LOG.info("Se va a abrir un navegador. Inicia sesion TU MISMO en cada portal")
    LOG.info("que se abra (usuario y clave los escribes tu; el programa no los")
    LOG.info("guarda ni los ve). Cuando termines en uno, vuelve aca y dale Enter.")
    LOG.info("")

    with Navegador(visible=True) as nav:
        for _, cfg in portales:
            inicio = re.match(r"https?://[^/]+", cfg["plantillas_url"][0]).group(0)
            LOG.info(">> Abriendo %s ...", cfg["nombre"])
            nav.ir_a(inicio)
            input(f"   Inicia sesion en {cfg['nombre']} y dale Enter aqui... ")
    LOG.info("Listo. Tus sesiones quedaron guardadas en %s", CARPETA_NAVEGADOR)


def _links_de_la_pagina(nav: Navegador, limite: int = 8):
    """Que formas de enlace hay en la pagina, de la mas repetida a la menos.

    Es LA pista que hace falta cuando un portal deja de dar resultados:
    si el programa busca enlaces con una forma y el portal ahora usa
    otra, aqui se ve cual es la nueva, sin adivinar.
    """
    try:
        enlaces = nav.pagina.eval_on_selector_all(
            "a[href]", "e => e.map(x => x.getAttribute('href') || '')")
    except Exception:
        return []

    from urllib.parse import urlparse
    conteo = {}
    ejemplos = {}
    for href in enlaces:
        ruta = urlparse(href).path
        if not ruta or ruta == "/":
            continue
        # Se agrupa por los dos primeros tramos: /ofertas-de-trabajo/xxx
        tramos = [t for t in ruta.split("/") if t][:2]
        if not tramos:
            continue
        forma = "/" + "/".join(tramos[:1]) + ("/..." if len(tramos) > 1 else "")
        conteo[forma] = conteo.get(forma, 0) + 1
        ejemplos.setdefault(forma, ruta)

    ordenados = sorted(conteo.items(), key=lambda x: -x[1])[:limite]
    return [(forma, veces, ejemplos[forma]) for forma, veces in ordenados]


def diagnostico(perfil):
    """Revisa portal por portal, direccion por direccion, que esta pasando.

    Cuando un portal cambia su pagina, el sintoma es siempre "0
    resultados" y no se sabe si la direccion quedo mala, si el portal
    bloqueo al programa o si de verdad no hay ofertas. Esto lo separa:
    prueba TODAS las direcciones candidatas, dice cual respondio,
    cuantos enlaces de oferta encontro, y ademas lista que formas de
    enlace trae la pagina -- que es lo que hace falta para corregir
    'patron_enlace' si el portal lo cambio.

    Deja la pagina guardada (.html) y una captura de cada portal, para
    poder mirarlas con calma o mandarlas a quien te ayude.
    """
    carpeta = CARPETA_DATOS / "diagnostico"
    carpeta.mkdir(parents=True, exist_ok=True)
    termino = perfil["cargos_objetivo"][0]
    ciudad = perfil.get("ciudad", "")

    LOG.info("")
    LOG.info("=" * 62)
    LOG.info(" DIAGNOSTICO -- buscando '%s' en %s", termino, ciudad or "todo el pais")
    LOG.info("=" * 62)

    with Navegador(visible=True) as nav:
        for clave, cfg in FUENTES.items():
            LOG.info("")
            LOG.info(">> %s", cfg["nombre"])
            if not cfg.get("habilitada"):
                LOG.info("   DESACTIVADO en FUENTES (no se prueba)")
                continue

            gano = None
            for numero, plantilla in enumerate(cfg["plantillas_url"], start=1):
                url = construir_url(plantilla, termino, ciudad, 1)
                LOG.info("   Direccion %d: %s", numero, url)

                if not nav.ir_a(url):
                    LOG.info("      -> no respondio (sin internet, o direccion mala)")
                    continue

                try:
                    titulo = limpiar_espacios(nav.pagina.title())[:70]
                except Exception:
                    titulo = ""
                LOG.info("      titulo de la pagina: %s", titulo or "(sin titulo)")

                if parece_bloqueo(nav):
                    LOG.info("      -> EL PORTAL PIDIO VERIFICACION (captcha o similar).")
                    LOG.info("         Prueba con NAVEGADOR_VISIBLE = True.")
                    continue

                resultados = _leer_resultados(nav, clave, cfg)
                if resultados:
                    LOG.info("      -> BIEN: %d ofertas. Ejemplo: %s",
                             len(resultados), resultados[0].titulo[:45])
                    gano = url
                    break

                LOG.info("      -> abrio, pero no reconoci ninguna oferta.")
                formas = _links_de_la_pagina(nav)
                if formas:
                    LOG.info("         Los enlaces que trae esta pagina son:")
                    for forma, veces, ejemplo in formas:
                        LOG.info("           %-28s x%-4d  ej: %s", forma, veces, ejemplo[:60])
                    LOG.info("         (el programa busca enlaces que digan '%s')",
                             cfg["patron_enlace"])

            # Se guarda lo ultimo que se vio, para poder revisarlo despues.
            try:
                nav.pagina.screenshot(path=str(carpeta / f"{clave}.png"), full_page=False)
                (carpeta / f"{clave}.html").write_text(nav.pagina.content(), encoding="utf-8")
            except Exception:
                pass

            if gano:
                LOG.info("   RESULTADO: funciona -> %s", gano)
            else:
                LOG.info("   RESULTADO: ninguna direccion sirvio.")
                LOG.info("   Mira %s.png y %s.html en la carpeta de diagnostico.",
                         clave, clave)

    LOG.info("")
    LOG.info("=" * 62)
    LOG.info(" Paginas y capturas guardadas en:")
    LOG.info("   %s", carpeta)
    LOG.info("=" * 62)


def main():
    parser = argparse.ArgumentParser(
        description="Busca empleo en los portales colombianos y postula por ti.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Empieza siempre con --simular hasta que el filtro quede a tu gusto.",
    )
    parser.add_argument("--vigilar", action="store_true",
                        help=f"queda revisando cada {INTERVALO_VIGILANCIA_MIN} minutos")
    parser.add_argument("--simular", action="store_true",
                        help="busca y califica, pero NO manda nada")
    parser.add_argument("--automatico", action="store_true",
                        help="postula sin preguntar (ojo con esto)")
    parser.add_argument("--login", action="store_true",
                        help="abre el navegador para que inicies sesion en los portales")
    parser.add_argument("--diagnostico", action="store_true",
                        help="revisa si los portales siguen respondiendo")
    parser.add_argument("--reporte", action="store_true",
                        help="exporta el Excel de postulaciones y sale")
    parser.add_argument("--probar-correo", action="store_true",
                        help="manda un correo de prueba a ti mismo y sale")
    args = parser.parse_args()

    try:
        perfil = cargar_perfil()
    except ErrorConfiguracion as e:
        LOG.error("")
        LOG.error("FALTA CONFIGURAR ALGO:")
        LOG.error("   %s", e)
        return 1

    registro = Registro()

    if args.reporte:
        ruta = exportar_excel(registro)
        mostrar_resumen(registro)
        LOG.info("Excel actualizado: %s", ruta)
        return 0

    try:
        if args.login:
            iniciar_sesiones(perfil)
            return 0
        if args.diagnostico:
            diagnostico(perfil)
            return 0
    except ErrorConfiguracion as e:
        LOG.error("%s", e)
        return 1

    if args.probar_correo:
        return probar_correo(perfil, cargar_credenciales())

    modo = MODO
    if args.simular:
        modo = "simulacion"
    elif args.automatico:
        modo = "automatico"

    cred = cargar_credenciales()
    if not cred and modo != "simulacion":
        LOG.warning("Sin %s no se puede postular por correo (solo por los "
                    "portales). Ver el archivo .example.", ARCHIVO_CREDENCIALES.name)

    try:
        if args.vigilar:
            vigilar(perfil, cred, registro, modo)
        else:
            una_pasada(perfil, cred, registro, modo)
            exportar_excel(registro)
            mostrar_resumen(registro)
            LOG.info("Detalle completo en: %s", ARCHIVO_EXCEL)
    except KeyboardInterrupt:
        LOG.info("")
        LOG.info("Detenido por ti. Lo que se alcanzo a hacer ya quedo guardado.")
    except ErrorConfiguracion as e:
        LOG.error("%s", e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
