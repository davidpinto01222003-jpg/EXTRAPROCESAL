"""
Script individual y de SOLO LECTURA: filtra del Excel de control (hoja
ACTIVOS -- misma lectura que clasificar_procesos_ejecutivos.py, con el
mismo arrastre hacia abajo de los procesos "acumulados") los procesos
cuyo ESTADO PROCESAL es exactamente "TERMINADO POR AUTO" o "TERMINADO
POR PAGO" (no incluye "TERMINADO POR CONTRATO"/"PREPAGO" ni "NO
INICIO"), y genera una lista en CSV con las PARTES (demandante --
siempre ESSA/Electrificadora de Santander en este proyecto -- y
demandado(s)) y el RADICADO de cada uno.

Un proceso "acumulado" (varias cuentas bajo el mismo radicado) sale en
UNA sola fila del reporte, con todos sus demandados juntos -- igual que
una sola carpeta en clasificar_procesos_ejecutivos.py.

Ademas del reporte detallado (terminados_auto_pago.csv), genera
TAMBIEN un archivo .xlsx listo para importar a un sistema propio:
UNA fila por caso, dos columnas -- "radicado" y "correo del
responsable" (esta ultima se deja SIEMPRE vacia; se llena a mano
despues, o el sistema de destino crea el caso sin asignar si el correo
no pertenece a su firma). Los procesos sin radicado todavia (raro en
terminados por auto/pago, pero puede pasar) NO salen en ese archivo --
no hay como importarlos sin radicado -- y quedan listados aparte en el
log para que los completes a mano.

Este script nunca crea, renombra ni borra ninguna carpeta -- solo lee
el Excel y escribe los reportes.
"""

import logging
import os

from openpyxl import Workbook

import clasificar_procesos_ejecutivos as base
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "listar_terminados_auto_pago.log")
ARCHIVO_REPORTE = os.path.join(os.path.dirname(__file__), "terminados_auto_pago.csv")
ARCHIVO_IMPORTAR = os.path.join(os.path.dirname(__file__), "terminados_auto_pago_importar.xlsx")

# Demandante de todos los procesos de este proyecto -- no existe una
# columna de "demandante" en el Excel porque siempre es la misma parte.
DEMANDANTE = "ELECTRIFICADORA DE SANTANDER S.A. E.S.P. (ESSA)"

# Prefijos de ESTADO PROCESAL a incluir en el reporte -- SOLO terminados
# por auto o por pago (no "TERMINADO POR CONTRATO"/"PREPAGO", ni "NO
# INICIO").
PREFIJOS_ESTADO_A_LISTAR = ("TERMINADO POR AUTO", "TERMINADO POR PAGO")

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


def filtrar_terminados_auto_pago(terminados):
    """De la lista 'terminados' de clasificar_procesos_ejecutivos.clasificar_procesos, solo los de PREFIJOS_ESTADO_A_LISTAR."""
    return [
        proceso for proceso in terminados
        if proceso["estado"].upper().startswith(PREFIJOS_ESTADO_A_LISTAR)
    ]


def escribir_archivo_importar(filtrados):
    """
    Genera ARCHIVO_IMPORTAR: una fila por caso, columnas "radicado" y
    "correo del responsable" (siempre vacia). Solo incluye los procesos
    que YA tienen radicado -- devuelve (filas_escritas, numeros_sin_radicado).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Importar"
    ws.append(["radicado", "correo del responsable"])

    sin_radicado = []
    escritas = 0
    for proceso in filtrados:
        if not proceso["radicado"]:
            sin_radicado.append(proceso["numero"])
            continue
        ws.append([proceso["radicado"], ""])
        escritas += 1

    wb.save(ARCHIVO_IMPORTAR)
    return escritas, sin_radicado


def procesar():
    if not base.RUTA_EXCEL_CONTROL or not os.path.exists(base.RUTA_EXCEL_CONTROL):
        logging.error("No se encontro el Excel configurado en RUTA_EXCEL_CONTROL: %r", base.RUTA_EXCEL_CONTROL)
        return

    procesos = base.leer_procesos_control()
    _con_radicado, terminados, _sin_estado = base.clasificar_procesos(procesos)

    filtrados = filtrar_terminados_auto_pago(terminados)
    logging.info(
        "Excel: %d proceso(s) terminado(s) en total, %d de ellos terminados por auto o por pago.",
        len(terminados), len(filtrados),
    )

    if not filtrados:
        logging.info("No hay ningun proceso terminado por auto o por pago todavia -- no se genera el reporte.")
        return

    filtrados.sort(key=lambda p: p["numero"])
    filas_reporte = [
        (
            proceso["numero"],
            proceso["estado"],
            proceso["radicado"] or "",
            DEMANDANTE,
            ", ".join(proceso["demandados"]),
        )
        for proceso in filtrados
    ]

    if cruce_excel._escribir_csv_tolerante(
        ARCHIVO_REPORTE,
        ["No.", "Estado", "Radicado", "Demandante", "Demandado(s)"],
        filas_reporte,
        "Terminados por auto/pago",
    ):
        logging.info("[Reporte] %d proceso(s) guardados en: %s", len(filtrados), ARCHIVO_REPORTE)

    escritas, sin_radicado = escribir_archivo_importar(filtrados)
    logging.info("[Importar] %d fila(s) (radicado, correo vacio) guardadas en: %s", escritas, ARCHIVO_IMPORTAR)
    if sin_radicado:
        logging.warning(
            "%d proceso(s) terminado(s) por auto/pago todavia NO tienen radicado en el Excel -- no salen en "
            "%s, revisalos a mano: %s",
            len(sin_radicado), ARCHIVO_IMPORTAR, ", ".join(str(n) for n in sin_radicado),
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
