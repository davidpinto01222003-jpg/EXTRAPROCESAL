"""
App de celular para el buscador de empleo.

Levanta en tu PC un servidor pequeno y lo publica en tu red WiFi. Desde
el celular abres la direccion que este programa te muestra, le das
"Agregar a pantalla de inicio", y a partir de ahi se comporta como
cualquier app del telefono: icono propio, pantalla completa, sin barra
de navegador.

POR QUE EL BUSCADOR NO CORRE DENTRO DEL CELULAR
============================================================
Para leer los portales hay que abrir un navegador Chromium de verdad
(Playwright). Ni Android ni iPhone dejan hacer eso en segundo plano, y
las tiendas de apps rechazan las aplicaciones que se postulan solas. Lo
que si funciona -- y es lo que hace esto -- es repartir el trabajo:

    EL PC        busca en los portales y manda las hojas de vida
    EL CELULAR   es el control remoto: ver, decidir y disparar

Consecuencia practica: **el PC tiene que estar encendido** y con este
programa corriendo para que el celular funcione. Si apagas el PC, la app
del celular se queda sin datos nuevos (los ultimos que bajo se siguen
viendo).

DESDE EL CELULAR PUEDES
============================================================
  - Ver todas tus postulaciones, con su puntaje y por que se descarto
    cada oferta.
  - Ver las que quedaron "para ti" (las que el portal no dejo terminar
    solo) y abrirlas de un toque para rematarlas.
  - Disparar una busqueda en el momento, en modo prueba o de verdad.
  - Cambiar el filtro: cargos, ciudades, palabras excluyentes, anos de
    experiencia, salario minimo, umbral y tope diario.
  - Ver el avance en vivo mientras el PC busca.

SEGURIDAD
============================================================
El panel muestra tus datos personales y puede mandar postulaciones a tu
nombre, asi que:

  - Solo escucha en tu red local (tu WiFi). No lo publiques en internet.
  - Pide un CODIGO de 8 caracteres que se genera solo la primera vez y
    queda guardado en datos_empleo/token_movil.txt. La primera vez lo
    escribes en el celular (o abres el link completo que se muestra
    abajo) y no lo vuelve a pedir en ese telefono.
  - Si sospechas que alguien mas lo tiene, borra ese archivo y vuelve a
    arrancar: se genera un codigo nuevo.

USO
============================================================
    python app_movil.py              (y abres en el celular lo que muestre)
    python app_movil.py --puerto 9000
    python app_movil.py --solo-este-pc    (no lo publica en el WiFi)

En Windows: doble clic en `abrir_movil.bat`.
"""

import argparse
import datetime
import json
import logging
import mimetypes
import secrets
import socket
import threading
import traceback
from collections import deque
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import buscar_empleo as app

CARPETA_WEB = Path(__file__).resolve().parent / "movil"
ARCHIVO_TOKEN = app.CARPETA_DATOS / "token_movil.txt"

PUERTO_POR_DEFECTO = 8777

# Campos del perfil que se pueden editar desde el celular. El resto
# (hoja de vida, credenciales, carta) se edita en el PC a proposito:
# son los que romperian todo si se tocan por accidente desde el bus.
CAMPOS_EDITABLES = {
    "cargos_objetivo": list,
    "palabras_clave_deseables": list,
    "palabras_excluyentes": list,
    "ciudades": list,
    "acepta_remoto": bool,
    "anos_experiencia": int,
    "nivel_educativo": str,
    "salario_minimo": int,
    "umbral_postulacion": int,
    "max_postulaciones_dia": int,
    "modo": str,
}


# ============================================================
# TAREA EN SEGUNDO PLANO (la busqueda)
# ============================================================

class Tarea:
    """La busqueda corriendo en el PC, vista desde el celular."""

    def __init__(self):
        self.candado = threading.Lock()
        self.hilo = None
        self.modo = ""
        self.inicio = None
        self.fin = None
        self.error = ""
        self.lineas = deque(maxlen=400)

    @property
    def corriendo(self) -> bool:
        return self.hilo is not None and self.hilo.is_alive()

    def arrancar(self, modo: str) -> bool:
        with self.candado:
            if self.corriendo:
                return False
            self.modo = modo
            self.inicio = datetime.datetime.now()
            self.fin = None
            self.error = ""
            self.lineas.clear()
            self.hilo = threading.Thread(target=self._correr, args=(modo,), daemon=True)
            self.hilo.start()
            return True

    def _correr(self, modo: str):
        try:
            perfil = app.cargar_perfil()
            cred = app.cargar_credenciales()
            registro = app.Registro()
            app.una_pasada(perfil, cred, registro, modo)
            try:
                app.exportar_excel(registro)
            except app.ErrorConfiguracion as e:
                app.LOG.warning("No se pudo generar el Excel: %s", e)
        except Exception as e:
            self.error = f"{type(e).__name__}: {e}"
            app.LOG.error("La busqueda fallo: %s", self.error)
            app.LOG.debug(traceback.format_exc())
        finally:
            self.fin = datetime.datetime.now()

    def estado(self) -> dict:
        return {
            "corriendo": self.corriendo,
            "modo": self.modo,
            "inicio": self.inicio.strftime("%H:%M:%S") if self.inicio else "",
            "fin": self.fin.strftime("%H:%M:%S") if self.fin else "",
            "error": self.error,
            "lineas": list(self.lineas)[-60:],
        }


TAREA = Tarea()


class RecolectorDeLineas(logging.Handler):
    """Copia el log del buscador para poder verlo desde el celular.

    Solo recoge lo que escribe EL HILO DE LA BUSQUEDA. Si copiara todo,
    en el telefono se mezclarian los avisos que genera el propio panel
    cada vez que se refresca, y el avance dejaria de leerse.
    """

    def emit(self, registro):
        try:
            hilo = TAREA.hilo
            if hilo is None or threading.get_ident() != hilo.ident:
                return
            TAREA.lineas.append(self.format(registro))
        except Exception:
            pass


# ============================================================
# DATOS QUE VE EL CELULAR
# ============================================================

ETIQUETAS_ESTADO = {
    "postulada": "Postuladas",
    "pendiente_revision": "Para ti",
    "descartada": "Descartadas",
    "duplicada": "Repetidas",
    "en_espera": "En espera",
    "simulada": "De prueba",
    "error": "Con error",
    "revisada": "Ya la hiciste",
}

# En las tarjetas de conteo va el plural ("3 Postuladas"); en cada
# vacante, el singular ("Postulada"). Es lo unico que las diferencia.
ETIQUETAS_UNA = {
    "postulada": "Postulada",
    "pendiente_revision": "Para ti",
    "descartada": "Descartada",
    "duplicada": "Repetida",
    "en_espera": "En espera",
    "simulada": "De prueba",
    "error": "Con error",
    "revisada": "Ya la hiciste",
}

# Lo que se muestra arriba en la app, en este orden.
ORDEN_TARJETAS = ["pendiente_revision", "postulada", "en_espera", "descartada"]


_ESTADO_PERFIL = {"mtime": False, "ok": False, "problema": ""}


def revisar_perfil():
    """Dice si el perfil sirve, releyendolo solo cuando cambia.

    El celular pregunta por el estado cada pocos segundos. Volver a leer
    y validar el perfil en cada pregunta llenaria el log de avisos
    repetidos, asi que se recuerda el resultado hasta que el archivo se
    modifique de verdad.
    """
    try:
        mtime = app.ARCHIVO_PERFIL.stat().st_mtime
    except OSError:
        mtime = None

    if _ESTADO_PERFIL["mtime"] != mtime:
        try:
            app.cargar_perfil()
            _ESTADO_PERFIL.update(ok=True, problema="")
        except app.ErrorConfiguracion as e:
            _ESTADO_PERFIL.update(ok=False, problema=str(e))
        _ESTADO_PERFIL["mtime"] = mtime
    return _ESTADO_PERFIL["ok"], _ESTADO_PERFIL["problema"]


def resumen_general() -> dict:
    registro = app.Registro()
    conteo = {}
    for entrada in registro.datos.values():
        estado = entrada.get("estado", "?")
        conteo[estado] = conteo.get(estado, 0) + 1

    perfil_ok, problema_perfil = revisar_perfil()

    return {
        "conteo": conteo,
        "tarjetas": [
            {"estado": e, "etiqueta": ETIQUETAS_ESTADO.get(e, e), "total": conteo.get(e, 0)}
            for e in ORDEN_TARJETAS
        ],
        "postuladas_hoy": registro.postuladas_hoy(),
        "tope_diario": app.MAX_POSTULACIONES_DIA,
        "umbral": app.UMBRAL_POSTULACION,
        "modo": app.MODO,
        "total": len(registro.datos),
        "perfil_ok": perfil_ok,
        "problema_perfil": problema_perfil,
        "tarea": TAREA.estado(),
    }


def listar(estado: str = "", busqueda: str = "", limite: int = 60) -> list:
    registro = app.Registro()
    filas = []
    aguja = app.sin_tildes(busqueda)

    for clave, entrada in registro.datos.items():
        if estado and entrada.get("estado") != estado:
            continue
        if aguja:
            texto = app.sin_tildes(" ".join(str(entrada.get(c, "")) for c in
                                            ("titulo", "empresa", "ciudad", "portal", "detalle")))
            if aguja not in texto:
                continue
        fila = {c: entrada.get(c, "") for c in
                ("estado", "puntaje", "titulo", "empresa", "ciudad", "portal",
                 "metodo", "detalle", "url", "actualizado")}
        fila["clave"] = clave
        fila["etiqueta"] = ETIQUETAS_UNA.get(fila["estado"], fila["estado"])
        filas.append(fila)

    filas.sort(key=lambda f: str(f.get("actualizado", "")), reverse=True)
    return filas[:limite]


def perfil_para_movil() -> dict:
    """Solo los campos editables, mas lo que hace falta para explicarlos."""
    try:
        with open(app.ARCHIVO_PERFIL, encoding="utf-8") as f:
            datos = json.load(f)
    except (OSError, json.JSONDecodeError):
        datos = {}

    salida = {}
    for campo, tipo in CAMPOS_EDITABLES.items():
        valor = datos.get(campo)
        if valor is None:
            if campo == "umbral_postulacion":
                valor = app.UMBRAL_POSTULACION
            elif campo == "max_postulaciones_dia":
                valor = app.MAX_POSTULACIONES_DIA
            elif campo == "modo":
                valor = app.MODO
            else:
                valor = [] if tipo is list else ("" if tipo is str else 0)
        salida[campo] = valor
    salida["nombre_completo"] = datos.get("nombre_completo", "")
    return salida


def guardar_perfil(cambios: dict) -> dict:
    """Guarda los campos editables. Todo lo demas del perfil se respeta.

    Se valida aqui, no en el celular: un valor absurdo mandado desde el
    telefono no puede dejar el perfil en un estado que rompa la
    busqueda del PC.
    """
    if not app.ARCHIVO_PERFIL.exists():
        raise app.ErrorConfiguracion(
            "Todavia no existe perfil_laboral.json. Crealo en el PC "
            "(copia perfil_laboral.example.json) y vuelve a intentar."
        )
    with open(app.ARCHIVO_PERFIL, encoding="utf-8") as f:
        perfil = json.load(f)

    aplicados = []
    for campo, tipo in CAMPOS_EDITABLES.items():
        if campo not in cambios:
            continue
        valor = cambios[campo]
        try:
            if tipo is list:
                if isinstance(valor, str):
                    valor = [p.strip() for p in valor.split(",")]
                valor = [str(p).strip() for p in valor if str(p).strip()]
            elif tipo is bool:
                valor = bool(valor)
            elif tipo is int:
                valor = int(valor)
            elif tipo is str:
                valor = str(valor).strip()
        except (TypeError, ValueError):
            continue

        if campo == "umbral_postulacion" and not (0 <= valor <= 100):
            continue
        if campo == "max_postulaciones_dia" and not (1 <= valor <= 200):
            continue
        if campo == "anos_experiencia" and not (0 <= valor <= 60):
            continue
        if campo == "salario_minimo" and not (0 <= valor <= 100_000_000):
            continue
        if campo == "nivel_educativo" and app.sin_tildes(valor) not in app.NIVELES_EDUCATIVOS:
            continue
        if campo == "modo" and valor not in ("simulacion", "semiautomatico", "automatico"):
            continue
        if campo == "cargos_objetivo" and not valor:
            continue  # sin cargos no hay nada que buscar

        perfil[campo] = valor
        aplicados.append(campo)

    temporal = app.ARCHIVO_PERFIL.with_suffix(".tmp")
    with open(temporal, "w", encoding="utf-8") as f:
        json.dump(perfil, f, ensure_ascii=False, indent=2)
    temporal.replace(app.ARCHIVO_PERFIL)

    app.cargar_perfil()  # revalida y reaplica los ajustes en caliente
    return {"guardados": aplicados}


def marcar(clave: str, estado: str) -> dict:
    if estado not in ("revisada", "descartada", "pendiente_revision"):
        raise ValueError("estado no permitido")
    registro = app.Registro()
    if clave not in registro.datos:
        raise ValueError("esa vacante no esta en el registro")
    registro.anotar(clave, estado=estado)
    return {"clave": clave, "estado": estado}


# ============================================================
# SERVIDOR
# ============================================================

def token_de_acceso() -> str:
    """Codigo de emparejamiento. Se crea una vez y se reusa siempre."""
    ARCHIVO_TOKEN.parent.mkdir(parents=True, exist_ok=True)
    if ARCHIVO_TOKEN.exists():
        guardado = ARCHIVO_TOKEN.read_text(encoding="utf-8").strip()
        if guardado:
            return guardado
    nuevo = secrets.token_hex(4).upper()
    ARCHIVO_TOKEN.write_text(nuevo, encoding="utf-8")
    return nuevo


TOKEN = ""


class Manejador(BaseHTTPRequestHandler):
    server_version = "BuscadorEmpleo/1.0"

    # --- utilidades ------------------------------------------
    def log_message(self, formato, *args):
        pass  # el log del navegador no aporta nada; el util es el de app.LOG

    def _responder(self, codigo: int, cuerpo: bytes, tipo: str, extra=None):
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        # El panel no debe poder incrustarse en otra pagina.
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        for clave, valor in (extra or {}):
            self.send_header(clave, valor)
        self.end_headers()
        try:
            self.wfile.write(cuerpo)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, datos, codigo: int = 200):
        self._responder(codigo, json.dumps(datos, ensure_ascii=False).encode("utf-8"),
                        "application/json; charset=utf-8")

    def _autorizado(self, consulta: dict) -> bool:
        if secrets.compare_digest(consulta.get("t", [""])[0].upper().strip(), TOKEN):
            return True
        cabecera = (self.headers.get("X-Token") or "").upper().strip()
        if cabecera and secrets.compare_digest(cabecera, TOKEN):
            return True
        for trozo in (self.headers.get("Cookie") or "").split(";"):
            nombre, _, valor = trozo.strip().partition("=")
            if nombre == "token" and secrets.compare_digest(valor.upper().strip(), TOKEN):
                return True
        return False

    def _cuerpo_json(self) -> dict:
        try:
            largo = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return {}
        if largo <= 0 or largo > 200_000:
            return {}
        try:
            return json.loads(self.rfile.read(largo).decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            return {}

    # --- rutas -----------------------------------------------
    def do_GET(self):
        partes = urlparse(self.path)
        ruta = partes.path
        consulta = parse_qs(partes.query)

        # El service worker debe poder pedirse sin cookie todavia.
        if ruta in ("/sw.js", "/manifest.json") or ruta.startswith("/estatico/"):
            return self._archivo(ruta)

        if not self._autorizado(consulta):
            return self._pedir_codigo()

        if ruta == "/":
            cuerpo = (CARPETA_WEB / "index.html").read_bytes()
            galleta = f"token={TOKEN}; Path=/; Max-Age=31536000; SameSite=Strict"
            return self._responder(200, cuerpo, "text/html; charset=utf-8",
                                   extra=[("Set-Cookie", galleta)])

        if ruta == "/api/estado":
            return self._json(resumen_general())

        if ruta == "/api/lista":
            return self._json({"filas": listar(
                estado=consulta.get("estado", [""])[0],
                busqueda=consulta.get("q", [""])[0],
            )})

        if ruta == "/api/perfil":
            return self._json(perfil_para_movil())

        if ruta == "/api/log":
            try:
                lineas = app.ARCHIVO_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
            except OSError:
                lineas = []
            return self._json({"lineas": lineas[-120:]})

        return self._json({"error": "no existe esa direccion"}, 404)

    def do_POST(self):
        partes = urlparse(self.path)
        if not self._autorizado(parse_qs(partes.query)):
            return self._json({"error": "codigo invalido"}, 401)

        datos = self._cuerpo_json()
        try:
            if partes.path == "/api/buscar":
                modo = datos.get("modo", "simulacion")
                if modo not in ("simulacion", "semiautomatico", "automatico"):
                    return self._json({"error": "modo invalido"}, 400)
                if not TAREA.arrancar(modo):
                    return self._json({"error": "ya hay una busqueda corriendo"}, 409)
                return self._json({"ok": True, "modo": modo})

            if partes.path == "/api/perfil":
                return self._json(guardar_perfil(datos))

            if partes.path == "/api/marcar":
                return self._json(marcar(datos.get("clave", ""), datos.get("estado", "")))

        except (app.ErrorConfiguracion, ValueError) as e:
            return self._json({"error": str(e)}, 400)
        except Exception as e:
            app.LOG.debug(traceback.format_exc())
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)

        return self._json({"error": "no existe esa direccion"}, 404)

    def _archivo(self, ruta: str):
        nombre = ruta.lstrip("/")
        if nombre.startswith("estatico/"):
            nombre = nombre[len("estatico/"):]
        destino = (CARPETA_WEB / nombre).resolve()
        # Nadie sale de la carpeta movil/ por mas ".." que mande.
        if CARPETA_WEB.resolve() not in destino.parents or not destino.is_file():
            return self._json({"error": "no existe ese archivo"}, 404)
        tipo, _ = mimetypes.guess_type(destino.name)
        self._responder(200, destino.read_bytes(), tipo or "application/octet-stream")

    def _pedir_codigo(self):
        pagina = (CARPETA_WEB / "codigo.html").read_bytes()
        self._responder(401, pagina, "text/html; charset=utf-8")


def ip_en_la_red() -> str:
    """La IP del PC dentro del WiFi (la que el celular tiene que marcar)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))  # no manda nada; solo mira que ruta usaria
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


def main():
    global TOKEN

    parser = argparse.ArgumentParser(
        description="Publica el buscador de empleo como app para el celular.")
    parser.add_argument("--puerto", type=int, default=PUERTO_POR_DEFECTO)
    parser.add_argument("--solo-este-pc", action="store_true",
                        help="no lo publica en el WiFi (solo se abre en este PC)")
    args = parser.parse_args()

    recolector = RecolectorDeLineas()
    recolector.setFormatter(logging.Formatter("%(message)s"))
    # INFO y no DEBUG: al telefono va el avance legible, no las
    # trazas tecnicas (esas quedan completas en buscar_empleo.log).
    recolector.setLevel(logging.INFO)
    app.LOG.addHandler(recolector)

    TOKEN = token_de_acceso()
    direccion = "127.0.0.1" if args.solo_este_pc else "0.0.0.0"
    ip = "127.0.0.1" if args.solo_este_pc else ip_en_la_red()
    url = f"http://{ip}:{args.puerto}/?t={TOKEN}"

    servidor = ThreadingHTTPServer((direccion, args.puerto), Manejador)
    servidor.daemon_threads = True

    print()
    print("=" * 62)
    print(" LA APP YA ESTA CORRIENDO")
    print("=" * 62)
    print()
    print(" 1. Conecta el celular al MISMO WiFi que este PC.")
    print(" 2. Abre esta direccion en el navegador del celular:")
    print()
    print(f"       {url}")
    print()
    print(f"    (o entra a http://{ip}:{args.puerto} y escribe el codigo: {TOKEN})")
    print()
    print(" 3. Instalala como app:")
    print("       Android (Chrome):  menu (3 puntos) > 'Instalar aplicacion'")
    print("       iPhone (Safari):   compartir > 'Agregar a pantalla de inicio'")
    print()
    print(" Deja esta ventana ABIERTA: el PC es el que hace la busqueda.")
    print(" Ctrl+C para apagar la app.")
    print("=" * 62)
    print()

    if not app.ARCHIVO_PERFIL.exists():
        print(" AVISO: todavia no existe perfil_laboral.json. La app abre igual,")
        print(" pero no podra buscar hasta que lo crees (ver el README).")
        print()

    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nApagando la app. Lo que ya se postulo quedo guardado.")
    finally:
        servidor.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
