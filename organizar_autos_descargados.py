"""
Organiza los documentos que TU ya bajaste a mano a la carpeta de
Descargas y que TERMINAN un proceso -- NO busca nada en Google Drive ni
en el correo, no necesita credenciales, y no se demora horas: solo mira
lo que ya esta en tu disco.

Por defecto revisa SOLO lo que bajaste HOY (ver DIAS_HACIA_ATRAS).

Para cada archivo de Descargas (PDF/DOCX, sueltos o dentro de un .zip):

 1. Le lee el nombre y el texto y decide si es el documento que TERMINA
    un proceso -- un auto de terminacion, de aceptacion de retiro de la
    demanda, de desistimiento, de archivo, una terminacion por pago,
    etc (mismo criterio que
    clasificar_procesos_ejecutivos.es_auto_terminador: no exige que
    diga literalmente "AUTO"). Lo que no lo parece, se deja quieto.

    Si el NOMBRE del archivo ya dice que es otra cosa (una demanda, un
    memorial, unos anexos, un mandamiento... ver
    _el_nombre_lo_descarta), se descarta aunque el texto de adentro
    mencione que se acepto un retiro o que el proceso termino: un
    "ANEXOS DEMANDA.pdf" cuenta la historia del proceso, no es el
    documento que lo cierra.
 2. Lo empareja con su proceso del Excel de control por RADICADO (con
    o sin guiones), por CUENTA, o por el nombre del DEMANDADO
    (cualquiera de los tres alcanza -- es el mismo emparejamiento que
    usa revisar_correo_pro.py para los correos).
 3. Si ese proceso figura en el Excel como TERMINADO -- por auto, por
    pago, por contrato/prepago, CUALQUIER terminado -- mueve el archivo
    RENOMBRADO con el numero de proceso ("245. TERMINADO POR AUTO.pdf",
    "300. TERMINADO POR PAGO.pdf") a la carpeta comun de ese estado:
    "PROCESOS TERMINADOS POR AUTO", "PROCESOS TERMINADOS POR PAGO",
    etc, dentro de CARPETA_PROCESOS (las mismas que usa
    clasificar_procesos_ejecutivos.py).

Lo que NO se puede decidir solo, no se toca: se deja en Descargas y
queda listado en el reporte ARCHIVO_REPORTE (y en el log) con el
motivo. Son tres casos:
  - No coincide con ningun proceso del Excel.
  - Coincide con VARIOS procesos terminados (no hay forma de saber a
    cual de ellos corresponde sin abrirlo).
  - Coincide con un proceso que en el Excel NO figura como terminado
    (ej. sigue como ACTIVO) -- ahi lo que hay que revisar es el Excel,
    no el archivo. El reporte te dice el numero y el estado.

Nunca borra ni sobreescribe nada: si el nombre de destino ya existe (un
proceso con documento de primera y de segunda instancia), el nuevo se
guarda como "245. TERMINADO POR AUTO_2.pdf". Los .zip no se tocan ni se
borran: si adentro viene el documento, se EXTRAE una copia ya
renombrada y el zip se queda como esta.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo revisa y
te dice que moveria y con que nombre, sin tocar ningun archivo.

Reutiliza la lectura del Excel y las reglas de
clasificar_procesos_ejecutivos.py (misma hoja ACTIVOS, mismo
CARPETA_PROCESOS, mismas carpetas por estado) -- no hay nada que
configurar aparte, salvo la carpeta de Descargas si la tuya no es la de
Windows.
"""

import datetime
import importlib
import io
import logging
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

# Las UNICAS librerias que necesita este script (no watchdog, no
# playwright, no las de Google Drive), como "modulo que se importa" ->
# "paquete que hay que instalar".
#
# OJO con cryptography: NINGUN modulo del proyecto la importa, la usa
# pypdf por dentro y SOLO cuando abre un PDF cifrado con AES. Por eso
# hay que revisarla a proposito aqui: si se esperara a que fallara un
# import, no se instalaria nunca, y esos PDF -- que suelen ser
# justamente los autos que manda el juzgado -- se perderian con un
# "cryptography>=3.1 is required for AES algorithm" en el log.
LIBRERIAS_NECESARIAS = {
    "pypdf": "pypdf",
    "docx": "python-docx",
    "openpyxl": "openpyxl",
    "cryptography": "cryptography",
}

# Modulos del proyecto que se usan (alias -> archivo .py de al lado).
_MODULOS_DEL_PROYECTO = {
    "buscador": "buscar_faltantes_en_drive",
    "base": "clasificar_procesos_ejecutivos",
    "organizador": "procesos_juridicos",
    "cruce_excel": "validar_renombrar_carpetas",
}


def _paquetes_que_faltan():
    faltantes = []
    for modulo, paquete in LIBRERIAS_NECESARIAS.items():
        try:
            importlib.import_module(modulo)
        except ImportError:
            faltantes.append(paquete)
    return faltantes


def _asegurar_librerias():
    """
    Instala lo que falte y vuelve a revisar. Asi el .bat funciona con
    doble clic en un PC nuevo, sin pelear con la terminal. Devuelve lo
    que siguio faltando despues de intentar. Se usa print y no logging
    porque el logging todavia no esta configurado.
    """
    faltantes = _paquetes_que_faltan()
    if not faltantes:
        return []

    print()
    print(f"  Falta(n) la(s) libreria(s): {', '.join(faltantes)}.")
    print("  Instalando (solo pasa la primera vez), espera un momento...")
    print()
    subprocess.run([sys.executable, "-m", "pip", "install", *faltantes], check=False)
    print()
    importlib.invalidate_caches()
    return _paquetes_que_faltan()


_faltantes = _asegurar_librerias()
if _faltantes:
    print()
    print("=" * 70)
    print(f"  NO SE PUDO INSTALAR: {', '.join(_faltantes)}")
    print()
    print("  Abre una terminal EN ESTA MISMA CARPETA y corre a mano:")
    print()
    print("      pip install " + " ".join(_faltantes))
    print()
    print("  Si eso falla, revisa que Python este bien instalado y que")
    print("  tengas conexion a internet.")
    print("=" * 70)
    print()
    raise SystemExit(1)

for _alias, _modulo in _MODULOS_DEL_PROYECTO.items():
    globals()[_alias] = importlib.import_module(_modulo)

# ============================= CONFIGURACION =============================

# Carpeta donde caen tus descargas. Se detecta sola como "Downloads" del
# usuario de Windows actual (la misma que ya usa procesos_juridicos.py);
# cambiala aqui si la tuya esta en otro lado.
CARPETA_DESCARGAS = organizador.CARPETA_DESCARGAS

# Donde van los documentos organizados: la MISMA carpeta que usa
# clasificar_procesos_ejecutivos.py, para que no queden en dos sitios.
# Adentro se crea una carpeta por estado -- "PROCESOS TERMINADOS POR
# AUTO", "PROCESOS TERMINADOS POR PAGO", etc (ver
# clasificar_procesos_ejecutivos.carpeta_agrupada_para).
CARPETA_PROCESOS = base.CARPETA_PROCESOS

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "organizar_autos_descargados.log")

# Reporte de lo que NO se pudo organizar solo (con el motivo de cada uno).
ARCHIVO_REPORTE = os.path.join(os.path.dirname(__file__), "autos_descargados_a_revisar.csv")

# True (por defecto): no mueve ni renombra nada, solo revisa y muestra
# que haria. False: organiza de verdad.
MODO_PRUEBA = True

# True (por defecto): el archivo se MUEVE de Descargas a su carpeta (asi
# Descargas queda limpia y no se vuelve a revisar el mismo archivo en la
# proxima corrida). False: se COPIA, dejando el original en Descargas.
MOVER_EN_VEZ_DE_COPIAR = True

# Cuantos dias contar hacia atras, por fecha del archivo:
#   1 = SOLO lo que bajaste HOY (por defecto)
#   2 = hoy y ayer, 7 = la ultima semana, etc
#   0 = revisar TODO lo que haya, sin importar la fecha
# Una carpeta de Descargas normal acumula años de archivos y leerles el
# texto a todos toma un buen rato, ademas de revolver documentos viejos
# que ya organizaste.
DIAS_HACIA_ATRAS = 1

# True (por defecto): tambien revisa las subcarpetas de Descargas.
REVISAR_SUBCARPETAS = True

# True (por defecto): tambien mira dentro de los .zip que haya en
# Descargas (los expedientes del juzgado suelen llegar asi). El zip
# nunca se borra ni se modifica -- si adentro viene el auto, se extrae
# una copia renombrada.
REVISAR_ZIPS = True

# False (por defecto): NO exige que el documento mencione a ESSA/
# Electrificadora de Santander. A diferencia de lo que se baja solo de
# Drive o del correo (donde esa regla es obligatoria para no traer
# expedientes de otro cliente), aca los archivos los bajaste TU a
# proposito. Ponlo en True si tu carpeta de Descargas mezcla documentos
# de varios clientes.
EXIGIR_MENCION_ESSA = False

# Extensiones a las que se les puede leer el texto para clasificarlas.
EXTENSIONES_REVISABLES = {".pdf", ".docx"}

# Subcarpetas de Descargas que nunca se revisan.
SUBCARPETAS_A_OMITIR = {"procesados", "_tmp_extraccion"}

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8", mode="w"),
            logging.StreamHandler(),
        ],
    )


# ==================== Leer lo que hay en Descargas ====================


def _es_reciente(ruta: Path) -> bool:
    """
    True si el archivo es de los ultimos DIAS_HACIA_ATRAS dias de
    CALENDARIO (1 = solo hoy, 2 = hoy y ayer...; 0 = sin limite). Se
    cuenta por dia, no por horas: algo que bajaste hoy a las 8 de la
    mañana cuenta igual aunque ya sean las 11 de la noche.
    """
    if not DIAS_HACIA_ATRAS:
        return True
    try:
        modificado = datetime.date.fromtimestamp(ruta.stat().st_mtime)
    except OSError:
        return False
    limite = datetime.date.today() - datetime.timedelta(days=DIAS_HACIA_ATRAS - 1)
    return modificado >= limite


def listar_candidatos(carpeta_descargas: Path):
    """
    (archivos, zips) de Descargas dentro de la ventana de DIAS_HACIA_ATRAS:
    'archivos' son los PDF/DOCX sueltos, 'zips' los .zip por revisar (si
    REVISAR_ZIPS). Omite SUBCARPETAS_A_OMITIR.
    """
    archivos, zips = [], []
    patron = "**/*" if REVISAR_SUBCARPETAS else "*"
    for ruta in sorted(carpeta_descargas.glob(patron)):
        if not ruta.is_file():
            continue
        partes = {parte.lower() for parte in ruta.relative_to(carpeta_descargas).parts[:-1]}
        if partes & SUBCARPETAS_A_OMITIR:
            continue
        if not _es_reciente(ruta):
            continue
        extension = ruta.suffix.lower()
        if extension in EXTENSIONES_REVISABLES:
            archivos.append(ruta)
        elif extension == ".zip" and REVISAR_ZIPS:
            zips.append(ruta)
    return archivos, zips


def _texto_de_bytes(nombre: str, contenido: bytes) -> str:
    """Texto de un PDF/DOCX que todavia esta en memoria (ej. dentro de un .zip)."""
    extension = Path(nombre).suffix.lower()
    try:
        if extension == ".pdf" and buscador.PdfReader is not None:
            lector = buscador.PdfReader(io.BytesIO(contenido))
            return "\n".join((pagina.extract_text() or "") for pagina in lector.pages)
        if extension == ".docx" and buscador.docx is not None:
            return "\n".join(p.text for p in buscador.docx.Document(io.BytesIO(contenido)).paragraphs)
    except Exception as error:
        logging.warning("   (no se pudo leer '%s' dentro del zip: %s)", nombre, error)
    return ""


# ==================== Decidir de que proceso es ====================


def _mencion_essa_ok(contenido_normalizado: str) -> bool:
    if not EXIGIR_MENCION_ESSA:
        return True
    return any(buscador._nombre_coincide(contenido_normalizado, t) for t in buscador.TERMINOS_DEMANDANTE_VALIDO)


def _el_nombre_lo_descarta(nombre: str, nombre_normalizado: str) -> bool:
    """
    True si el NOMBRE del archivo ya dice que es un documento procesal
    (demanda, memorial, anexos de demanda, mandamiento, embargo... ver
    base.PALABRAS_PROCESAL_EXCLUIR) y NO se anuncia el mismo como el
    documento que termina el proceso.

    Hace falta porque el contenido engaña: un "02. ANEXOS DEMANDA.pdf"
    puede mencionar adentro "se acepta el retiro de la demanda" y
    hacerse pasar por el auto. El nombre es la señal mas confiable (es
    el mismo orden de prioridad que usa base._decision_solo_por_nombre).

    Ojo con el orden: primero se mira si el NOMBRE mismo ya dice que
    termina el proceso ("AUTO ACEPTA EL RETIRO DE LA DEMANDA.pdf"
    contiene la palabra "DEMANDA", pero es exactamente lo que estamos
    buscando), y solo si no, se aplica el descarte.
    """
    if base.es_auto_terminador(nombre_normalizado):
        return False
    return any(marca in nombre_normalizado for marca in base.PALABRAS_PROCESAL_EXCLUIR)


def coincidencias_por_nivel(contenido_normalizado: str, indices):
    """
    Los procesos que coinciden con este documento, SEPARADOS por que tan
    fuerte es la coincidencia y en ese orden -- lo mismo que mira
    base._procesos_que_coinciden_con_correo, pero sin mezclarlo todo en
    una sola bolsa, para poder desempatar cuando salen varios:

      1. RADICADO COMPLETO, los 23 digitos (con o sin guiones, puntos o
         espacios entre sus grupos -- "68001-40-03-001-2024-00050-00" es
         el mismo numero). Es tan especifico que practicamente no se
         repite: si coincide, es ESE proceso.
      2. RADICADO CORTO: "2024-00234" o "2024-234" (y sus variantes con
         el consecutivo de instancia al final). Casi siempre acierta,
         pero dos juzgados distintos pueden tener el mismo consecutivo
         en el mismo año.
      3. NUMERO DE CUENTA.
      4. NOMBRE DEL DEMANDADO -- el mas debil: un mismo demandado puede
         tener varios procesos.

    Cada proceso aparece UNA sola vez, en el nivel mas fuerte donde haya
    coincidido. Devuelve [(criterio, [procesos]), ...].
    """
    por_radicado, por_cuenta, por_demandado = indices
    exactos, cortos, por_su_cuenta, por_su_demandado = [], [], [], []

    radicados_del_texto = set(cruce_excel._radicados_en_texto(contenido_normalizado))
    for radicado, procesos in por_radicado.items():
        if radicado in radicados_del_texto:
            exactos.extend(procesos)
        elif any(buscador._nombre_coincide(contenido_normalizado, corto) for corto in buscador.radicados_cortos(radicado)):
            cortos.extend(procesos)

    for cuenta, procesos in por_cuenta.items():
        if buscador._nombre_coincide(contenido_normalizado, cuenta):
            por_su_cuenta.extend(procesos)

    encabezado = base._quitar_membrete_juzgado(contenido_normalizado)[:base.VENTANA_DEMANDADO_CARACTERES]
    for palabras, procesos in por_demandado.items():
        if all(base._palabra_demandado_coincide(encabezado, palabra) for palabra in palabras):
            por_su_demandado.extend(procesos)

    niveles = [
        ("el radicado completo (23 digitos)", exactos),
        ("el radicado corto (ej. 2024-00234 o 2024-234)", cortos),
        ("el numero de cuenta", por_su_cuenta),
        ("el nombre del demandado", por_su_demandado),
    ]

    vistos = set()
    resultado = []
    for criterio, procesos in niveles:
        unicos = []
        for proceso in procesos:
            if proceso["nombre_carpeta"] in vistos:
                continue
            vistos.add(proceso["nombre_carpeta"])
            unicos.append(proceso)
        resultado.append((criterio, unicos))
    return resultado


def _detalle(procesos) -> str:
    return ", ".join(f"{p['numero']} ({p['estado']})" for p in procesos)


def decidir(nombre: str, contenido_normalizado: str, indices):
    """
    Decide que hacer con UN documento (venga suelto o de un zip).
    Devuelve (proceso, motivo):
      - (proceso, motivo)  -> es el documento que termina ese proceso,
                              se puede organizar.
      - (None, motivo)     -> no se organiza; 'motivo' explica por que.
      - (None, None)       -> ni siquiera parece el documento que
                              termina un proceso: se ignora en silencio
                              (es la mayoria de lo que hay en Descargas
                              y no tiene nada que ver).

    Cuando coinciden VARIOS procesos, se desempata por la fuerza de la
    coincidencia (ver coincidencias_por_nivel): manda el radicado
    completo de 23 digitos; si no aparece, el radicado corto
    ("2024-00234" o "2024-234"); despues la cuenta; y de ultimo el
    nombre del demandado. Se decide con el PRIMER nivel que traiga
    coincidencias y no se sigue bajando: si el documento trae el
    radicado de un proceso, es de ESE proceso, aunque de casualidad
    mencione la cuenta o el demandado de otro.
    """
    nombre_normalizado = buscador._normalizar_para_comparar(Path(nombre).name)

    if not base.es_auto_terminador(contenido_normalizado):
        return None, None

    if _el_nombre_lo_descarta(nombre, nombre_normalizado):
        return None, None

    if not _mencion_essa_ok(contenido_normalizado):
        return None, "parece el documento que termina un proceso, pero no menciona a ESSA/Electrificadora de Santander"

    for criterio, procesos in coincidencias_por_nivel(contenido_normalizado, indices):
        if not procesos:
            continue  # nada por este criterio: se prueba con el siguiente, mas debil

        terminados = [p for p in procesos if base.es_estado_agrupado(p["estado"])]

        if len(terminados) == 1:
            return terminados[0], f"es el documento que termina el proceso (coincide por {criterio})"

        if len(terminados) > 1:
            return None, (
                f"parece el documento que termina un proceso y coincide por {criterio} con VARIOS procesos "
                f"terminados ({_detalle(terminados)}) -- revisalo a mano"
            )

        # Coincide, pero ninguno de esos figura como terminado. No se
        # sigue bajando de nivel: el criterio mas fuerte ya hablo.
        return None, (
            f"parece el documento que termina un proceso y coincide por {criterio} con {_detalle(procesos)}, "
            "pero en el Excel ninguno de esos figura como TERMINADO -- revisa el Excel"
        )

    return None, "parece el documento que termina un proceso, pero no coincide con el radicado, la cuenta ni el demandado de ningun proceso del Excel"


# ==================== Guardar donde va ====================


def _carpeta_destino(proceso) -> Path:
    """
    Carpeta comun del estado de ESTE proceso -- "PROCESOS TERMINADOS POR
    AUTO", "PROCESOS TERMINADOS POR PAGO", etc (la calcula
    clasificar_procesos_ejecutivos, para que los dos scripts dejen todo
    en el mismo sitio).
    """
    return Path(CARPETA_PROCESOS) / proceso["carpeta_destino"]


def _nombre_final(proceso, nombre_original: str) -> str:
    """"<numero>. <ESTADO PROCESAL>" + la extension del documento -- ej. "245. TERMINADO POR AUTO.pdf"."""
    extension = Path(nombre_original).suffix.lower() or ".pdf"
    return organizador.sanear_nombre(f"{proceso['nombre_archivo']}{extension}")


def _ya_esta_organizado(destino: Path, nombre_final: str, tamano: int) -> bool:
    """
    True si en la carpeta de autos ya hay un archivo de ESTE proceso con
    el mismo tamaño -- es el mismo documento, ya organizado en una
    corrida anterior (solo puede pasar copiando, ver
    MOVER_EN_VEZ_DE_COPIAR: moviendo, el original ya no esta en
    Descargas). Evita llenar la carpeta de "_2", "_3" iguales.
    """
    base_nombre, extension = Path(nombre_final).stem, Path(nombre_final).suffix
    if not destino.exists():
        return False
    for existente in destino.iterdir():
        if not existente.is_file() or existente.suffix.lower() != extension.lower():
            continue
        if existente.stem != base_nombre and not existente.stem.startswith(f"{base_nombre}_"):
            continue
        try:
            if existente.stat().st_size == tamano:
                return True
        except OSError:
            continue
    return False


def guardar_archivo_suelto(ruta: Path, proceso) -> Path:
    """Mueve (o copia, ver MOVER_EN_VEZ_DE_COPIAR) 'ruta' a la carpeta comun de su estado, ya renombrada."""
    destino = _carpeta_destino(proceso)
    base._crear_carpeta(destino)
    ruta_final = buscador._ruta_archivo_libre(destino, _nombre_final(proceso, ruta.name))
    origen_seguro = organizador._ruta_larga_segura(str(ruta))
    destino_seguro = organizador._ruta_larga_segura(str(ruta_final))
    if MOVER_EN_VEZ_DE_COPIAR:
        shutil.move(origen_seguro, destino_seguro)
    else:
        shutil.copy2(origen_seguro, destino_seguro)
    return ruta_final


def extraer_de_zip(archivo_zip: zipfile.ZipFile, nombre_interno: str, proceso) -> Path:
    """Extrae UN documento de dentro de un .zip a la carpeta comun de su estado, ya renombrado. El zip no se toca."""
    destino = _carpeta_destino(proceso)
    base._crear_carpeta(destino)
    ruta_final = buscador._ruta_archivo_libre(destino, _nombre_final(proceso, nombre_interno))
    with open(organizador._ruta_larga_segura(str(ruta_final)), "wb") as f:
        f.write(archivo_zip.read(nombre_interno))
    return ruta_final


# ==================== Orquestacion ====================


def _anotar(pendientes, origen: str, nombre: str, motivo: str):
    pendientes.append((nombre, origen, motivo))
    logging.warning("[Revisar] '%s' (%s): %s", nombre, origen, motivo)


def revisar_archivos_sueltos(archivos, indices, pendientes, resumen):
    for ruta in archivos:
        contenido = base._contenido_normalizado_local(ruta)
        proceso, motivo = decidir(ruta.name, contenido, indices)

        if proceso is None:
            if motivo is None:
                resumen["ignorados"] += 1
            else:
                _anotar(pendientes, str(ruta.parent), ruta.name, motivo)
            continue

        nombre_final = _nombre_final(proceso, ruta.name)
        try:
            tamano = ruta.stat().st_size
        except OSError:
            tamano = -1
        if not MOVER_EN_VEZ_DE_COPIAR and _ya_esta_organizado(_carpeta_destino(proceso), nombre_final, tamano):
            logging.info(
                "[Ya estaba] '%s' -> proceso %s: ya hay una copia igual en '%s'.",
                ruta.name, proceso["numero"], proceso["carpeta_destino"],
            )
            resumen["ya_estaban"] += 1
            continue

        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] '%s' -> proceso %s: %s a '%s' como '%s'.",
                ruta.name, proceso["numero"], "se moveria" if MOVER_EN_VEZ_DE_COPIAR else "se copiaria",
                proceso["carpeta_destino"], nombre_final,
            )
            resumen["organizados"] += 1
            continue

        try:
            ruta_final = guardar_archivo_suelto(ruta, proceso)
        except OSError as error:
            _anotar(pendientes, str(ruta.parent), ruta.name, f"no se pudo mover (¿abierto, o bloqueado por OneDrive/antivirus?): {error}")
            continue
        logging.info("[Organizado] '%s' -> proceso %s: %s", ruta.name, proceso["numero"], ruta_final)
        resumen["organizados"] += 1


def revisar_zips(zips, indices, pendientes, resumen):
    for ruta_zip in zips:
        try:
            with zipfile.ZipFile(organizador._ruta_larga_segura(str(ruta_zip))) as archivo_zip:
                nombres = [
                    n for n in archivo_zip.namelist()
                    if not n.endswith("/") and Path(n).suffix.lower() in EXTENSIONES_REVISABLES
                ]
                for nombre_interno in nombres:
                    texto = _texto_de_bytes(nombre_interno, archivo_zip.read(nombre_interno))
                    contenido = buscador._normalizar_para_comparar(Path(nombre_interno).name + " " + texto)
                    proceso, motivo = decidir(nombre_interno, contenido, indices)

                    if proceso is None:
                        if motivo is None:
                            resumen["ignorados"] += 1
                        else:
                            _anotar(pendientes, f"{ruta_zip.name} (zip)", Path(nombre_interno).name, motivo)
                        continue

                    nombre_final = _nombre_final(proceso, nombre_interno)
                    if _ya_esta_organizado(_carpeta_destino(proceso), nombre_final, archivo_zip.getinfo(nombre_interno).file_size):
                        logging.info(
                            "[Ya estaba] '%s' (dentro de %s) -> proceso %s: ya hay una copia igual en '%s'.",
                            Path(nombre_interno).name, ruta_zip.name, proceso["numero"], proceso["carpeta_destino"],
                        )
                        resumen["ya_estaban"] += 1
                        continue

                    if MODO_PRUEBA:
                        logging.info(
                            "[SIMULACION] '%s' (dentro de %s) -> proceso %s: se extraeria a '%s' como '%s'.",
                            Path(nombre_interno).name, ruta_zip.name, proceso["numero"],
                            proceso["carpeta_destino"], nombre_final,
                        )
                        resumen["organizados"] += 1
                        continue

                    ruta_final = extraer_de_zip(archivo_zip, nombre_interno, proceso)
                    logging.info(
                        "[Organizado] '%s' (dentro de %s) -> proceso %s: %s",
                        Path(nombre_interno).name, ruta_zip.name, proceso["numero"], ruta_final,
                    )
                    resumen["organizados"] += 1
        except (zipfile.BadZipFile, OSError) as error:
            _anotar(pendientes, str(ruta_zip.parent), ruta_zip.name, f"no se pudo abrir el zip: {error}")


def faltan_librerias() -> bool:
    """
    pypdf no rompe el import (los modulos del proyecto lo cargan de forma
    tolerante), pero sin el TODOS los PDF darian texto vacio: no se
    reconoceria ningun auto y el script diria tranquilamente que no hay
    nada que organizar. Mejor parar y decirlo claro.
    """
    if buscador.PdfReader is not None:
        return False
    logging.error("Falta la libreria 'pypdf', que es la que le lee el texto a los PDF.")
    logging.error("Sin ella no se puede reconocer ningun auto. Instalala corriendo:  pip install pypdf")
    logging.error("(o  pip install -r requirements.txt  para instalar de una vez todo lo del proyecto)")
    return True


def procesar():
    if faltan_librerias():
        return

    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    carpeta_descargas = Path(CARPETA_DESCARGAS) if CARPETA_DESCARGAS else None
    if carpeta_descargas is None or not carpeta_descargas.is_dir():
        logging.error("No existe la carpeta de Descargas configurada en CARPETA_DESCARGAS: %r", CARPETA_DESCARGAS)
        return

    procesos = base.leer_procesos_control()
    con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)
    # Se indexan TODOS los procesos (no solo los terminados) para poder
    # avisar cuando un documento corresponde a un proceso que en el
    # Excel todavia figura como ACTIVO/suspendido/etc.
    indices = base._indexar_procesos_para_correo(con_radicado + terminados)

    por_estado = {}
    for proceso in terminados:
        if base.es_estado_agrupado(proceso["estado"]):
            por_estado[proceso["estado"].strip().upper()] = por_estado.get(proceso["estado"].strip().upper(), 0) + 1
    logging.info(
        "Excel: %d proceso(s) en total, %d terminado(s) -- son los unicos que se organizan aqui: %s.",
        len(con_radicado) + len(terminados), sum(por_estado.values()),
        ", ".join(f"{estado} ({cuantos})" for estado, cuantos in sorted(por_estado.items())) or "ninguno",
    )

    archivos, zips = listar_candidatos(carpeta_descargas)
    if DIAS_HACIA_ATRAS == 1:
        ventana = " (solo los de HOY)"
    elif DIAS_HACIA_ATRAS:
        ventana = f" (los de los ultimos {DIAS_HACIA_ATRAS} dias)"
    else:
        ventana = " (todos, sin importar la fecha)"
    logging.info(
        "Descargas (%s): %d archivo(s) PDF/DOCX y %d zip(s) por revisar%s.",
        carpeta_descargas, len(archivos), len(zips), ventana,
    )
    if not archivos and not zips:
        logging.info("No hay nada que revisar en Descargas -- no se hizo ningun cambio.")
        return

    resumen = {"organizados": 0, "ya_estaban": 0, "ignorados": 0}
    pendientes = []

    revisar_archivos_sueltos(archivos, indices, pendientes, resumen)
    revisar_zips(zips, indices, pendientes, resumen)

    logging.info(
        "Resumen: %d documento(s) %s, %d que ya estaban organizados, %d archivo(s) que no terminan ningun "
        "proceso (se ignoran), "
        "%d que hay que revisar a mano.",
        resumen["organizados"], "se organizarian (MODO_PRUEBA activo)" if MODO_PRUEBA else "organizados",
        resumen["ya_estaban"], resumen["ignorados"], len(pendientes),
    )

    if pendientes:
        if cruce_excel._escribir_csv_tolerante(
            ARCHIVO_REPORTE,
            ["Archivo", "Donde esta", "Por que no se organizo"],
            pendientes,
            "Autos a revisar",
        ):
            logging.info("[Reporte] %d archivo(s) por revisar en: %s", len(pendientes), ARCHIVO_REPORTE)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se movio ni renombro nada todavia. Revisa el log y, si esta bien, "
            "pon MODO_PRUEBA = False en %s y vuelve a correrlo.", os.path.basename(__file__),
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
