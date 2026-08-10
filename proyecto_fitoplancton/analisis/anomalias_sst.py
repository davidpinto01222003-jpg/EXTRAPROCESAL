# -*- coding: utf-8 -*-
"""
Anomalías climatológicas estandarizadas de la temperatura superficial del mar.

Implementa los tres pasos descritos en la metodología del anteproyecto, siguiendo a
Santamaría del Ángel et al. (2019):

    Paso 1  climatología mensual   Xm  y  desviación típica mensual  sigma_m
    Paso 2  anomalía               A   =  Xi − Xm
    Paso 3  anomalía estandarizada Z   =  (Xi − Xm) / sigma_m

Entrada
-------
Un CSV con una fila por mes y una columna por estación:

    fecha,A01,A05,Muelle_CIOH,S1,S2,S4,S5,S6
    2002-06,28.91,28.87,29.02,...
    2002-07,29.14,29.10,29.20,...

Se admite cualquier subconjunto de estaciones. Los meses ausentes pueden omitirse o
dejarse vacíos; el cálculo los ignora.

Uso
---
    python3 anomalias_sst.py --entrada sst_mensual.csv
    python3 anomalias_sst.py --descargar          # intenta ERDDAP y guarda el CSV

La descarga usa el servidor ERDDAP de NOAA CoastWatch, que sirve las compuestas
mensuales de MODIS Aqua sin credenciales. Si la red del entorno bloquea ese servidor,
el script lo informa y no escribe nada: en ese caso se descarga el CSV desde una
máquina con salida a internet y se pasa con --entrada.
"""
import argparse, os, sys
import numpy as np
import pandas as pd

# posiciones de la tabla de coordenadas del anteproyecto
ESTACIONES = {
    "A01":         (-75.56, 10.39),
    "A05":         (-75.59, 10.39),
    "Muelle_CIOH": (-75.53, 10.39),
    "S1":          (-75.54, 10.40),
    "S2":          (-75.56, 10.32),
    "S4":          (-75.53, 10.35),
    "S5":          (-75.56, 10.29),
    "S6":          (-75.55, 10.38),
}

# campañas del proyecto, para cruzar la anomalía con cada salida de campo
CAMPANAS = {
    "Época seca 2021":     "2021-04",
    "Época seca 2022":     "2022-03",
    "Época lluviosa 2022": "2022-10",
    "Campaña de junio de 2023":     "2023-06",
    "Campaña de diciembre de 2023": "2023-12",
}

# Conjunto elegido: análisis MUR fv04.1, resolución 0,01 grados, cerca de 1 km, mensual,
# desde junio de 2002, que es justo el periodo y la resolución que describe el anteproyecto.
# El conjunto erdMH1sstdmday que aparece en el buscador está marcado como obsoleto y solo
# llega hasta 2019, de modo que no sirve para esta serie.
BASE_URL = "https://coastwatch.pfeg.noaa.gov/erddap/griddap"
CONJUNTO = "jplMURSST41mday"
VARIABLE = "sst"

SERVIDOR = (BASE_URL + "/{conjunto}.csv"
            "?{var}%5B({inicio}):({fin})%5D%5B({lat}):({lat})%5D%5B({lon}):({lon})%5D")

# una sola descarga que cubre la bahía completa y contiene las ocho estaciones
CAJA = (BASE_URL + "/{conjunto}.csv"
        "?{var}%5B({inicio}):({fin})%5D%5B(10.28):(10.41)%5D%5B(-75.60):(-75.52)%5D")


def descargar(inicio="2002-06-01", fin="2024-09-30", destino="sst_mensual.csv"):
    """Extrae la serie mensual de cada estación desde ERDDAP."""
    import urllib.request
    series = {}
    for nombre, (lon, lat) in ESTACIONES.items():
        url = SERVIDOR.format(conjunto=CONJUNTO, var=VARIABLE, inicio=inicio, fin=fin, lat=lat, lon=lon)
        try:
            with urllib.request.urlopen(url, timeout=120) as r:
                crudo = r.read().decode()
        except Exception as e:
            print(f"No fue posible descargar {nombre}: {e}")
            print("El entorno no alcanza el servidor de datos. Descargue el CSV aparte "
                  "y vuelva a ejecutar con --entrada.")
            return None
        filas = [l.split(",") for l in crudo.strip().split("\n")[2:] if l.strip()]
        s = pd.Series({f[0][:7]: float(f[3]) for f in filas if f[3] not in ("", "NaN")})
        series[nombre] = s
        print(f"  {nombre}: {len(s)} meses")
    df = pd.DataFrame(series)
    df.index.name = "fecha"
    df.to_csv(destino)
    print("serie guardada en", destino)
    return df


def imprimir_urls(inicio="2002-06-01", fin="2024-09-30"):
    """Direcciones de descarga, una por estación, para abrir en el navegador."""
    print("OPCIÓN RECOMENDADA: una sola descarga que cubre toda la bahía.\n")
    print("  " + CAJA.format(conjunto=CONJUNTO, var=VARIABLE, inicio=inicio, fin=fin) + "\n")
    print("Si el servidor responde que la variable no existe, pruebe con analysed_sst:\n")
    print("  " + CAJA.format(conjunto=CONJUNTO, var="analysed_sst", inicio=inicio, fin=fin) + "\n")
    print("Guarde ese archivo como bahia.csv y ejecute:\n")
    print("    python3 anomalias_sst.py --caja bahia.csv\n")
    print("Si rechaza el rango de latitud, invierta los dos valores de ese eje.\n")
    print("OPCIÓN ALTERNATIVA: una descarga por estación, guardadas en una misma carpeta,")
    print("y después  python3 anomalias_sst.py --carpeta <esa carpeta>\n")
    for nombre, (lon, lat) in ESTACIONES.items():
        print(f"{nombre}.csv")
        print("  " + SERVIDOR.format(conjunto=CONJUNTO, var=VARIABLE,
                                     inicio=inicio, fin=fin, lat=lat, lon=lon) + "\n")
    print("Si el conjunto no estuviera disponible, busque en")
    print("https://coastwatch.pfeg.noaa.gov/erddap/search y sustituya el identificador:")
    print("la rutina acepta cualquier CSV de ERDDAP con columnas de tiempo y de temperatura.")


def leer_caja(ruta):
    """Extrae la serie de cada estación desde un único CSV que cubre la bahía."""
    crudo = pd.read_csv(ruta, low_memory=False)
    if len(crudo) and not str(crudo.iloc[0, 0])[:4].isdigit():
        crudo = crudo.iloc[1:]
    col_t = next(c for c in crudo.columns if "time" in c.lower() or "fecha" in c.lower())
    col_lat = next(c for c in crudo.columns if "lat" in c.lower())
    col_lon = next(c for c in crudo.columns if "lon" in c.lower())
    resto = [c for c in crudo.columns if c not in (col_t, col_lat, col_lon)]
    col_v = next((c for c in resto if "sst" in c.lower() or "temp" in c.lower()), resto[-1])

    d = pd.DataFrame({
        "mes": pd.to_datetime(crudo[col_t], errors="coerce").dt.strftime("%Y-%m"),
        "lat": pd.to_numeric(crudo[col_lat], errors="coerce"),
        "lon": pd.to_numeric(crudo[col_lon], errors="coerce"),
        "valor": pd.to_numeric(crudo[col_v], errors="coerce")}).dropna(subset=["mes", "lat", "lon"])
    celdas = d[["lat", "lon"]].drop_duplicates()
    print(f"archivo con {len(celdas)} celdas y {d.mes.nunique()} meses, variable {col_v!r}")

    series = {}
    for nombre, (lon, lat) in ESTACIONES.items():
        dist = (celdas.lat - lat) ** 2 + (celdas.lon - lon) ** 2
        clat, clon = celdas.loc[dist.idxmin(), ["lat", "lon"]]
        sel = d[(d.lat == clat) & (d.lon == clon)].dropna(subset=["valor"])
        if sel.empty:
            print(f"  {nombre}: la celda más cercana no tiene datos, se omite")
            continue
        series[nombre] = sel.groupby("mes").valor.mean()
        km = ((clat - lat) ** 2 + (clon - lon) ** 2) ** .5 * 111
        print(f"  {nombre}: celda ({clat:.3f}, {clon:.3f}), a {km:.1f} km, {len(series[nombre])} meses")
    if not series:
        print("Ninguna estación tuvo datos utilizables en el archivo")
        sys.exit(1)
    df = pd.DataFrame(series).sort_index()
    df.index.name = "fecha"
    df.to_csv("sst_mensual.csv")
    print("tabla combinada guardada en sst_mensual.csv")
    return df


def leer_carpeta(carpeta):
    """Arma la tabla mensual a partir de los CSV que entrega ERDDAP, uno por estación."""
    import glob
    series = {}
    for ruta in sorted(glob.glob(os.path.join(carpeta, "*.csv"))):
        nombre = os.path.splitext(os.path.basename(ruta))[0]
        crudo = pd.read_csv(ruta, low_memory=False)
        # ERDDAP inserta una fila de unidades justo debajo del encabezado
        if len(crudo) and not str(crudo.iloc[0, 0])[:4].isdigit():
            crudo = crudo.iloc[1:]
        col_tiempo = next(c for c in crudo.columns if "time" in c.lower() or "fecha" in c.lower())
        candidatas = [c for c in crudo.columns
                      if c != col_tiempo and c.lower() not in ("latitude", "longitude", "altitude", "depth")]
        col_valor = next((c for c in candidatas if "sst" in c.lower() or "temp" in c.lower()), candidatas[-1])
        serie = pd.DataFrame({
            "mes": pd.to_datetime(crudo[col_tiempo], errors="coerce").dt.strftime("%Y-%m"),
            "valor": pd.to_numeric(crudo[col_valor], errors="coerce")}).dropna()
        series[nombre] = serie.groupby("mes").valor.mean()
        print(f"  {nombre}: {len(series[nombre])} meses, columna {col_valor!r}")
    if not series:
        print("No se encontraron archivos CSV en", carpeta)
        sys.exit(1)
    df = pd.DataFrame(series).sort_index()
    df.index.name = "fecha"
    df.to_csv("sst_mensual.csv")
    print("tabla combinada guardada en sst_mensual.csv")
    return df


def a_celsius(df):
    """Algunos conjuntos entregan la temperatura en kelvin; se convierte si hace falta."""
    if df.stack().median() > 200:
        print("valores en kelvin, se convierten a grados Celsius")
        return df - 273.15
    return df


def anomalias(df):
    """Devuelve climatología, desviación típica, anomalía y anomalía estandarizada."""
    fechas = pd.PeriodIndex(df.index, freq="M")
    mes = fechas.month
    clim = df.groupby(mes).mean()
    sigma = df.groupby(mes).std(ddof=1)
    clim.index.name = sigma.index.name = "mes"

    anom = df.copy().astype(float)
    z = df.copy().astype(float)
    for col in df.columns:
        anom[col] = df[col].values - clim[col].reindex(mes).values
        z[col] = anom[col].values / sigma[col].reindex(mes).values
    return clim, sigma, anom, z


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", help="CSV mensual con una columna por estación")
    ap.add_argument("--carpeta", help="carpeta con un CSV de ERDDAP por estación, nombrado <estación>.csv")
    ap.add_argument("--caja", help="un único CSV de ERDDAP que cubra la bahía; se extrae la celda más cercana a cada estación")
    ap.add_argument("--descargar", action="store_true", help="intentar la descarga desde ERDDAP")
    ap.add_argument("--urls", action="store_true", help="imprimir las direcciones de descarga y salir")
    ap.add_argument("--salida", default="anomalias_sst", help="prefijo de los archivos de salida")
    args = ap.parse_args()

    if args.urls:
        imprimir_urls()
        return
    if args.descargar:
        df = descargar()
        if df is None:
            sys.exit(1)
    elif args.caja:
        df = leer_caja(args.caja)
    elif args.carpeta:
        df = leer_carpeta(args.carpeta)
    elif args.entrada:
        df = pd.read_csv(args.entrada, index_col=0)
    else:
        ap.error("indique --caja, --carpeta, --entrada, --descargar o --urls")

    df = df.apply(pd.to_numeric, errors="coerce")
    print(f"serie de {len(df)} meses y {df.shape[1]} estaciones, "
          f"de {df.index[0]} a {df.index[-1]}")

    df = a_celsius(df)
    clim, sigma, anom, z = anomalias(df)

    print("\nClimatología mensual, promedio de las estaciones")
    resumen = pd.DataFrame({"climatología": clim.mean(axis=1).round(2),
                            "desviación típica": sigma.mean(axis=1).round(2)})
    print(resumen.to_string())

    print("\nAnomalía estandarizada en el mes de cada campaña")
    filas = []
    for nombre, ym in CAMPANAS.items():
        if ym in z.index:
            filas.append({"Campaña": nombre, "Mes": ym,
                          "SST observada": round(float(df.loc[ym].mean()), 2),
                          "Anomalía": round(float(anom.loc[ym].mean()), 2),
                          "Anomalía estandarizada": round(float(z.loc[ym].mean()), 2)})
        else:
            print(f"  aviso: la serie no cubre {ym} ({nombre})")
    if filas:
        print(pd.DataFrame(filas).to_string(index=False))
        pd.DataFrame(filas).to_csv(f"{args.salida}_campanas.csv", index=False)

    clim.round(3).to_csv(f"{args.salida}_climatologia.csv")
    sigma.round(3).to_csv(f"{args.salida}_desviacion.csv")
    anom.round(3).to_csv(f"{args.salida}_anomalia.csv")
    z.round(3).to_csv(f"{args.salida}_estandarizada.csv")

    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        serie = z.mean(axis=1)
        x = np.arange(len(serie))
        fig, ax = plt.subplots(figsize=(11, 3.4))
        ax.bar(x, serie.values, color=np.where(serie.values >= 0, "#c62828", "#1565c0"), width=1.0)
        ax.axhline(0, c="k", lw=.8)
        for lim in (1, -1):
            ax.axhline(lim, c="gray", ls="--", lw=.7)
        paso = max(1, len(serie) // 22)
        ax.set_xticks(x[::paso]); ax.set_xticklabels(serie.index[::paso], rotation=90, fontsize=7)
        ax.set_ylabel("Anomalía estandarizada")
        ax.set_title("Anomalía estandarizada de la temperatura superficial del mar en la Bahía de Cartagena")
        for nombre, ym in CAMPANAS.items():
            if ym in serie.index:
                ax.axvline(list(serie.index).index(ym), c="#2e7d32", lw=1.2, alpha=.8)
        plt.tight_layout(); plt.savefig(f"{args.salida}.png", dpi=220); plt.close()
        print(f"\nfigura guardada en {args.salida}.png")
    except ImportError:
        print("matplotlib no disponible, se omite la figura")

    print("tablas guardadas con el prefijo", args.salida)


if __name__ == "__main__":
    main()
