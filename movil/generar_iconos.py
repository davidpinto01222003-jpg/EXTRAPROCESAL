"""
Genera los iconos de la app del celular (icono-192.png e icono-512.png).

Los dibuja pixel por pixel con la libreria estandar de Python -- no hace
falta instalar nada. Solo hay que volver a correrlo si quieres cambiar
el color o el dibujo del icono:

    python movil/generar_iconos.py
"""

import struct
import zlib
from pathlib import Path

CARPETA = Path(__file__).resolve().parent

FONDO = (18, 50, 92)      # el azul de la app (--marca)
FRENTE = (255, 255, 255)  # el maletin


def dibujar(lado: int):
    """Un maletin blanco centrado sobre el azul de la app."""
    p = lado / 100.0  # para escribir las medidas en porcentaje del lado
    filas = []

    # Cuerpo del maletin
    cuerpo = (20 * p, 36 * p, 80 * p, 76 * p)
    # Asa (marco exterior y hueco interior)
    asa = (39 * p, 24 * p, 61 * p, 36 * p)
    hueco = (44 * p, 29 * p, 56 * p, 36 * p)
    # Cierre del centro
    cierre = (46 * p, 52 * p, 54 * p, 60 * p)
    # Franja que cruza el maletin
    franja = (20 * p, 48 * p, 80 * p, 51 * p)

    radio = 4 * p  # esquinas redondeadas del cuerpo

    def dentro(x, y, caja):
        return caja[0] <= x < caja[2] and caja[1] <= y < caja[3]

    def dentro_redondeado(x, y, caja, r):
        if not dentro(x, y, caja):
            return False
        x0, y0, x1, y1 = caja
        cx = min(max(x, x0 + r), x1 - r)
        cy = min(max(y, y0 + r), y1 - r)
        return (x - cx) ** 2 + (y - cy) ** 2 <= r * r or (x0 + r <= x < x1 - r) or (y0 + r <= y < y1 - r)

    for y in range(lado):
        fila = bytearray()
        fila.append(0)  # filtro PNG "sin filtro" al inicio de cada linea
        for x in range(lado):
            color = FONDO
            if dentro_redondeado(x + 0.5, y + 0.5, cuerpo, radio):
                color = FRENTE
            if dentro(x + 0.5, y + 0.5, asa) and not dentro(x + 0.5, y + 0.5, hueco):
                color = FRENTE
            if dentro(x + 0.5, y + 0.5, franja) or dentro(x + 0.5, y + 0.5, cierre):
                color = FONDO
            fila.extend(color)
        filas.append(bytes(fila))
    return b"".join(filas)


def guardar_png(ruta: Path, lado: int, datos: bytes):
    def trozo(tipo: bytes, contenido: bytes) -> bytes:
        cuerpo = tipo + contenido
        return (struct.pack(">I", len(contenido)) + cuerpo
                + struct.pack(">I", zlib.crc32(cuerpo) & 0xFFFFFFFF))

    cabecera = struct.pack(">IIBBBBB", lado, lado, 8, 2, 0, 0, 0)  # 8 bits, RGB
    png = (b"\x89PNG\r\n\x1a\n"
           + trozo(b"IHDR", cabecera)
           + trozo(b"IDAT", zlib.compress(datos, 9))
           + trozo(b"IEND", b""))
    ruta.write_bytes(png)
    return len(png)


def main():
    for lado in (192, 512):
        ruta = CARPETA / f"icono-{lado}.png"
        tamano = guardar_png(ruta, lado, dibujar(lado))
        print(f"  {ruta.name}: {lado}x{lado}, {tamano} bytes")


if __name__ == "__main__":
    print("Generando los iconos de la app...")
    main()
    print("Listo.")
