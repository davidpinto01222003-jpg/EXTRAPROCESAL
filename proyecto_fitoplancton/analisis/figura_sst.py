# -*- coding: utf-8 -*-
"""Figura de la climatología y de la anomalía estandarizada de temperatura superficial."""
import pandas as pd, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams.update({"font.size": 9, "axes.unicode_minus": True})

clim = pd.read_csv("anomalias_sst_climatologia.csv", index_col=0).mean(axis=1)
sigma = pd.read_csv("anomalias_sst_desviacion.csv", index_col=0).mean(axis=1)
z = pd.read_csv("anomalias_sst_estandarizada.csv", index_col=0).mean(axis=1)

CAMPANAS = {"2021-04": "Seca 2021", "2022-03": "Seca 2022",
            "2022-10": "Lluviosa 2022", "2023-12": "Diciembre 2023"}
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

fig, ax = plt.subplots(2, 1, figsize=(11, 6.4),
                       gridspec_kw={"height_ratios": [1, 1.25], "hspace": .42})

# a) climatologia mensual
a = ax[0]
a.errorbar(range(1, 13), clim.values, yerr=sigma.values, marker="o", ms=5,
           color="#1565c0", ecolor="#90a4ae", capsize=3, lw=1.6)
a.set_xticks(range(1, 13)); a.set_xticklabels(MESES)
a.set_ylabel("Temperatura (grados Celsius)")
a.set_title("a) Climatología mensual de la temperatura superficial, 2002 a 2024, "
            "con su desviación típica", fontsize=10)
a.grid(alpha=.25, lw=.6)

# b) serie de anomalia estandarizada
b = ax[1]
x = np.arange(len(z))
b.bar(x, z.values, width=1.0, color=np.where(z.values >= 0, "#c62828", "#1565c0"))
b.axhline(0, c="k", lw=.8)
for lim in (1, -1, 2, -2):
    b.axhline(lim, c="gray", ls="--", lw=.6)
etiquetas = list(z.index)
paso = 12
b.set_xticks(x[::paso]); b.set_xticklabels([e[:4] for e in etiquetas[::paso]], fontsize=8)
b.set_ylabel("Anomalía estandarizada")
b.set_xlim(-1, len(z))
b.set_title("b) Anomalía estandarizada mensual, con el mes de cada campaña señalado", fontsize=10)
for ym, nombre in CAMPANAS.items():
    if ym in etiquetas:
        i = etiquetas.index(ym)
        b.axvline(i, c="#2e7d32", lw=1.3)
        b.text(i, 2.95, nombre, rotation=90, fontsize=6.5, color="#2e7d32",
               ha="right", va="top")
b.set_ylim(-3.2, 3.2)

plt.savefig("fig5_sst.png", dpi=220, bbox_inches="tight")
plt.close()
print("fig5_sst.png generada")
