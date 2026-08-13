"""
Organiza los AUTOS que TU ya bajaste a mano a la carpeta de Descargas
-- NO busca nada en Google Drive ni en el correo, no necesita
credenciales, y no se demora horas: solo mira lo que ya esta en tu
disco.

Para cada archivo de Descargas (PDF/DOCX, sueltos o dentro de un .zip):

 1. Le lee el nombre y el texto y decide si es el AUTO que termina un
    proceso -- un auto de terminacion, de aceptacion de retiro de la
    demanda, de desistimiento, de archivo, etc (mismo criterio que
    clasificar_procesos_ejecutivos.es_auto_terminador: no exige que
    diga literalmente "AUTO"). Lo que no lo parece, se deja quieto.
 2. Lo empareja con su proceso del Excel de control por RADICADO (con
    o sin guiones), por CUENTA, o por el nombre del DEMANDADO
    (cualquiera de los tres alcanza -- es el mismo emparejamiento que
    usa revisar_correo_pro.py para los correos).
 3. Si ese proceso figura en el Excel como "TERMINADO POR AUTO", mueve
    el archivo -- RENOMBRADO con el numero de proceso, ej.
    "245. TERMINADO POR AUTO.pdf" -- a la carpeta
    "PROCESOS TERMINADOS POR AUTO" (la misma que usa
    clasificar_procesos_ejecutivos.py), dentro de CARPETA_PROCESOS.

Lo que NO se puede decidir solo, no se toca: se deja en Descargas y
queda listado en el reporte ARCHIVO_REPORTE (y en el log) con el
motivo. Son tres casos:
  - No coincide con ningun proceso del Excel.
  - Coincide con VARIOS procesos terminados por auto (no hay forma de
    saber a cual de ellos corresponde sin abrirlo).
  - Coincide con un proceso que en el Excel NO figura como "TERMINADO
    POR AUTO" (ej. sigue como ACTIVO, o quedo como TERMINADO POR PAGO)
    -- ahi lo que hay que revisar es el Excel, no el archivo.

Nunca borra ni sobreescribe nada: si el nombre de destino ya existe (un
proceso con auto de primera y de segunda instancia), el nuevo se guarda
como "245. TERMINADO POR AUTO_2.pdf". Los .zip no se tocan ni se
borran: si adentro viene el auto, se EXTRAE una copia ya renombrada y
el zip se queda como esta.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo revisa y
te dice que moveria y con que nombre, sin tocar ningun archivo.

Reutiliza la lectura del Excel y las reglas de
clasificar_procesos_ejecutivos.py (misma hoja ACTIVOS, mismo
CARPETA_PROCESOS, misma carpeta de autos) -- no hay nada que configurar
aparte, salvo la carpeta de Descargas si la tuya no es la de Windows.
"""

import datetime
import io
import logging
import os
import shutil
import zipfile
from pathlib import Path

import buscar_faltantes_en_drive as buscador
import clasificar_procesos_ejecutivos as base
import procesos_juridicos as organizador
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

# Carpeta donde caen tus descargas. Se detecta sola como "Downloads" del
# usuario de Windows actual (la misma que ya usa procesos_juridicos.py);
# cambiala aqui si la tuya esta en otro lado.
CARPETA_DESCARGAS = organizador.CARPETA_DESCARGAS

# Donde van los autos organizados: la MISMA carpeta que usa
# clasificar_procesos_ejecutivos.py, para que no queden en dos sitios.
CARPETA_PROCESOS = base.CARPETA_PROCESOS
CARPETA_TERMINADOS_POR_AUTO = base.CARPETA_TERMINADOS_POR_AUTO

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

# Cuantos dias hacia atras revisar, por fecha de modificacion del
# archivo. Una carpeta de Descargas normal acumula años de archivos, y
# leerles el texto a todos toma un buen rato. 0 = revisar TODO sin
# importar la fecha.
DIAS_HACIA_ATRAS = 60

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
    """True si el archivo se modifico dentro de DIAS_HACIA_ATRAS (0 = sin limite)."""
    if not DIAS_HACIA_ATRAS:
        return True
    try:
        modificado = datetime.datetime.fromtimestamp(ruta.stat().st_mtime)
    except OSError:
        return False
    return (datetime.datetime.now() - modificado).days <= DIAS_HACIA_ATRAS


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


def decidir(nombre: str, contenido_normalizado: str, indices):
    """
    Decide que hacer con UN documento (venga suelto o de un zip).
    Devuelve (proceso, motivo):
      - (proceso, motivo)  -> es el auto de ese proceso, se puede organizar.
      - (None, motivo)     -> no se organiza; 'motivo' explica por que.
      - (None, None)       -> ni siquiera parece un auto: se ignora en
                              silencio (es la mayoria de lo que hay en
                              Descargas y no tiene nada que ver).
    """
    if not base.es_auto_terminador(contenido_normalizado):
        return None, None

    if not _mencion_essa_ok(contenido_normalizado):
        return None, "parece un auto, pero no menciona a ESSA/Electrificadora de Santander"

    coincidencias = base._procesos_que_coinciden_con_correo(contenido_normalizado, indices)
    if not coincidencias:
        return None, "parece un auto, pero no coincide con el radicado, la cuenta ni el demandado de ningun proceso del Excel"

    terminados_por_auto = [p for p in coincidencias if base.es_terminado_por_auto(p["estado"])]

    if not terminados_por_auto:
        detalle = ", ".join(f"{p['numero']} ({p['estado']})" for p in coincidencias)
        return None, f"parece un auto y coincide con {detalle}, pero en el Excel ese proceso no figura como TERMINADO POR AUTO -- revisa el Excel"

    if len(terminados_por_auto) > 1:
        detalle = ", ".join(str(p["numero"]) for p in terminados_por_auto)
        return None, f"parece un auto, pero coincide con VARIOS procesos terminados por auto ({detalle}) -- revisalo a mano"

    return terminados_por_auto[0], "es el auto que termina el proceso"


# ==================== Guardar donde va ====================


def _carpeta_destino() -> Path:
    return Path(CARPETA_PROCESOS) / CARPETA_TERMINADOS_POR_AUTO


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
    """Mueve (o copia, ver MOVER_EN_VEZ_DE_COPIAR) 'ruta' a la carpeta de autos, ya renombrada."""
    destino = _carpeta_destino()
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
    """Extrae UN documento de dentro de un .zip a la carpeta de autos, ya renombrado. El zip no se toca."""
    destino = _carpeta_destino()
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
        if not MOVER_EN_VEZ_DE_COPIAR and _ya_esta_organizado(_carpeta_destino(), nombre_final, tamano):
            logging.info("[Ya estaba] '%s' -> proceso %s: ya hay una copia igual en '%s'.", ruta.name, proceso["numero"], CARPETA_TERMINADOS_POR_AUTO)
            resumen["ya_estaban"] += 1
            continue

        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION] '%s' -> proceso %s: %s a '%s' como '%s'.",
                ruta.name, proceso["numero"], "se moveria" if MOVER_EN_VEZ_DE_COPIAR else "se copiaria",
                CARPETA_TERMINADOS_POR_AUTO, nombre_final,
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
                    if _ya_esta_organizado(_carpeta_destino(), nombre_final, archivo_zip.getinfo(nombre_interno).file_size):
                        logging.info(
                            "[Ya estaba] '%s' (dentro de %s) -> proceso %s: ya hay una copia igual en '%s'.",
                            Path(nombre_interno).name, ruta_zip.name, proceso["numero"], CARPETA_TERMINADOS_POR_AUTO,
                        )
                        resumen["ya_estaban"] += 1
                        continue

                    if MODO_PRUEBA:
                        logging.info(
                            "[SIMULACION] '%s' (dentro de %s) -> proceso %s: se extraeria a '%s' como '%s'.",
                            Path(nombre_interno).name, ruta_zip.name, proceso["numero"],
                            CARPETA_TERMINADOS_POR_AUTO, nombre_final,
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


def procesar():
    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    carpeta_descargas = Path(CARPETA_DESCARGAS) if CARPETA_DESCARGAS else None
    if carpeta_descargas is None or not carpeta_descargas.is_dir():
        logging.error("No existe la carpeta de Descargas configurada en CARPETA_DESCARGAS: %r", CARPETA_DESCARGAS)
        return

    procesos = base.leer_procesos_control()
    con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)
    # Se indexan TODOS los procesos (no solo los terminados por auto)
    # para poder avisar cuando un auto corresponde a un proceso que en
    # el Excel todavia figura como ACTIVO/terminado por pago/etc.
    indices = base._indexar_procesos_para_correo(con_radicado + terminados)
    por_auto = sum(1 for p in terminados if base.es_terminado_por_auto(p["estado"]))
    logging.info(
        "Excel: %d proceso(s) en total, %d de ellos TERMINADO POR AUTO (son los unicos que se organizan aqui).",
        len(con_radicado) + len(terminados), por_auto,
    )

    archivos, zips = listar_candidatos(carpeta_descargas)
    logging.info(
        "Descargas (%s): %d archivo(s) PDF/DOCX y %d zip(s) por revisar%s.",
        carpeta_descargas, len(archivos), len(zips),
        f" (modificados en los ultimos {DIAS_HACIA_ATRAS} dias)" if DIAS_HACIA_ATRAS else "",
    )
    if not archivos and not zips:
        logging.info("No hay nada que revisar en Descargas -- no se hizo ningun cambio.")
        return

    resumen = {"organizados": 0, "ya_estaban": 0, "ignorados": 0}
    pendientes = []

    revisar_archivos_sueltos(archivos, indices, pendientes, resumen)
    revisar_zips(zips, indices, pendientes, resumen)

    logging.info(
        "Resumen: %d auto(s) %s, %d que ya estaban organizados, %d archivo(s) que no eran autos (se ignoran), "
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
