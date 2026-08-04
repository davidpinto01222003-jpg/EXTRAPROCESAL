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

Este script nunca crea, renombra ni borra ninguna carpeta -- solo lee
el Excel y escribe el reporte.
"""

import logging
import os

import clasificar_procesos_ejecutivos as base
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "listar_terminados_auto_pago.log")
ARCHIVO_REPORTE = os.path.join(os.path.dirname(__file__), "terminados_auto_pago.csv")

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


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
