# -*- coding: utf-8 -*-
"""Figuras finales del informe de avance."""
import pandas as pd, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, dendrogram
from scipy import stats

MINUS = "−"
plt.rcParams.update({"font.size": 9, "axes.unicode_minus": True})

d = pd.read_csv("base_analizada.csv")
dv = d[d.valido]
orden = ["Seca 2021", "Seca 2022", "Seca 2023", "Humeda 2022", "Humeda 2023"]
etiq = ["Seca\n2021", "Seca\n2022", "Seca\n2023", "Lluviosa\n2022", "Lluviosa\n2023"]

# ---------- Figura 1: EDA ----------
fig, ax = plt.subplots(1, 2, figsize=(9, 3.7))
dat = [dv[dv.campana == c].br.dropna().values for c in orden]
bp = ax[0].boxplot(dat, tick_labels=etiq, patch_artist=True, widths=.6)
for b_ in bp["boxes"]:
    b_.set_facecolor("#90b8d0")
ax[0].axhline(3.0, ls="--", c="#c62828", lw=1)
ax[0].axhline(2.5, ls="--", c="#2e7d32", lw=1)
ax[0].set_ylabel("Índice B/R   aphy(440) / aphy(676)")
ax[0].set_title("a) Índice por campaña")
ax[0].text(0.62, 3.06, "umbral picofitoplancton", color="#c62828", fontsize=6.5)
ax[0].text(0.62, 2.32, "umbral microfitoplancton", color="#2e7d32", fontsize=6.5)
sc = ax[1].scatter(dv.salinidad, dv.br, c=np.log10(dv.turbidez.replace(0, 0.05)),
                   cmap="viridis", s=27, edgecolor="k", linewidth=.3)
ax[1].axhline(3.0, ls="--", c="#c62828", lw=1)
ax[1].axhline(2.5, ls="--", c="#2e7d32", lw=1)
ax[1].set_xlabel("Salinidad")
ax[1].set_ylabel("Índice B/R")
ax[1].set_title("b) Índice frente a salinidad y turbidez")
plt.colorbar(sc, ax=ax[1], label="log10 de la turbidez (NTU)")
plt.tight_layout(); plt.savefig("fig1_eda.png", dpi=220); plt.close()

# ---------- Figura 2: clases, zonas, temperatura ----------
fig, ax = plt.subplots(1, 3, figsize=(11, 3.9))
ct = pd.crosstab(dv.campana, dv.clase_qc).reindex(orden)
pct = (ct.T / ct.sum(axis=1) * 100).T[["Micro", "Nano", "Pico"]]
col = {"Micro": "#2e7d32", "Nano": "#f9a825", "Pico": "#c62828"}
nom = {"Micro": "Microfitoplancton", "Nano": "Nanofitoplancton", "Pico": "Picofitoplancton"}
bot = np.zeros(len(pct))
for c in pct.columns:
    ax[0].bar(range(len(pct)), pct[c], bottom=bot, label=nom[c], color=col[c], edgecolor="white")
    for i, (v, b) in enumerate(zip(pct[c], bot)):
        if v > 6:
            ax[0].text(i, b + v / 2, f"{v:.0f}", ha="center", va="center", color="white", fontsize=8)
    bot = bot + pct[c].values
ax[0].set_xticks(range(len(pct))); ax[0].set_xticklabels(etiq, fontsize=7.5)
ax[0].set_ylabel("Porcentaje de estaciones")
ax[0].set_title("a) Composición por clase de tamaño")
ax[0].legend(fontsize=6.5, ncol=3, loc="lower center", bbox_to_anchor=(0.5, -0.30), frameon=False)

datos = [dv[dv.zona == "Alta influencia"].br.dropna(), dv[dv.zona == "Baja influencia"].br.dropna()]
bp = ax[1].boxplot(datos, tick_labels=["Alta influencia\n(salinidad < 25)", "Baja influencia\n(salinidad ≥ 25)"],
                   patch_artist=True, widths=.55)
for b_, c_ in zip(bp["boxes"], ["#8d6e63", "#4fc3f7"]):
    b_.set_facecolor(c_)
ax[1].axhline(3.0, ls="--", c="#c62828", lw=1); ax[1].axhline(2.5, ls="--", c="#2e7d32", lw=1)
ax[1].set_ylabel("Índice B/R")
ax[1].set_title("b) Zonas de influencia del Canal del Dique\n(Mann Whitney, p = 0,004)")

s = dv[["temperatura", "br"]].dropna()
ax[2].scatter(s.temperatura, s.br, s=27, c="#5c6bc0", edgecolor="k", linewidth=.3)
z = np.polyfit(s.temperatura, s.br, 1)
xx = np.linspace(s.temperatura.min(), s.temperatura.max(), 50)
ax[2].plot(xx, np.polyval(z, xx), c="#c62828", lw=1.4)
ax[2].axhline(3.0, ls="--", c="#c62828", lw=.8); ax[2].axhline(2.5, ls="--", c="#2e7d32", lw=.8)
ax[2].set_xlabel("Temperatura superficial (grados Celsius)")
ax[2].set_ylabel("Índice B/R")
ax[2].set_title(f"c) Índice frente a temperatura\n(rho de Spearman = {MINUS}0,273; p = 0,004)")
plt.tight_layout(); plt.savefig("fig2_clases.png", dpi=220); plt.close()

# ---------- Figura 3: multivariado ----------
M = pd.read_csv("base_multivariado.csv")
X = M[["temperatura", "salinidad", "turbidez", "br"]].copy()
X["turbidez"] = np.log10(X.turbidez.replace(0, 0.05))
X.columns = ["Temperatura", "Salinidad", "log10 turbidez", "Índice B/R"]
Z = (X - X.mean()) / X.std(ddof=1)
C = np.cov(Z.T, ddof=1)
val, vec = np.linalg.eigh(C)
i = np.argsort(val)[::-1]; val, vec = val[i], vec[:, i]
cargas = pd.DataFrame(vec[:, :2] * np.sqrt(val[:2]), index=X.columns, columns=["CP1", "CP2"])
sc = Z.values @ vec

fig, ax = plt.subplots(1, 2, figsize=(9.5, 4.2))
colr = {"Seca": "#ef6c00", "Humeda": "#1565c0"}
for ep, g in M.groupby("epoca"):
    m = (M.epoca == ep).values
    ax[0].scatter(sc[m, 0], sc[m, 1], c=colr[ep], label="Seca" if ep == "Seca" else "Lluviosa",
                  s=30, edgecolor="k", linewidth=.3)
for v in X.columns:
    ax[0].arrow(0, 0, cargas.CP1[v] * 2.3, cargas.CP2[v] * 2.3, color="k", head_width=.09, lw=1.1)
    ax[0].text(cargas.CP1[v] * 2.6, cargas.CP2[v] * 2.6, v, fontsize=7.5, ha="center")
ax[0].axhline(0, lw=.5, c="gray"); ax[0].axvline(0, lw=.5, c="gray")
ax[0].set_xlabel(f"CP1 ({val[0]/val.sum()*100:.1f} por ciento)")
ax[0].set_ylabel(f"CP2 ({val[1]/val.sum()*100:.1f} por ciento)")
ax[0].legend(fontsize=8); ax[0].set_title("a) Componentes principales")
lk = linkage(Z.values, method="ward")
dendrogram(lk, ax=ax[1], no_labels=True, color_threshold=lk[-2, 2])
ax[1].set_title("b) Conglomerados de Ward"); ax[1].set_ylabel("Distancia")
plt.tight_layout(); plt.savefig("fig3_multivariado.png", dpi=220); plt.close()
print("figuras finales generadas")

# ---------- Figura de histogramas: distribuciones por época y por zona ----------
dv2 = d[d.valido].copy()
dv2["logturb"] = np.log10(dv2.turbidez.replace(0, 0.05))
paneles = [("br", "Índice B/R", None),
           ("salinidad", "Salinidad", None),
           ("logturb", "Logaritmo decimal de la turbidez", None)]
cortes = [("epoca", {"Seca": ("#ef6c00", "Época seca"), "Humeda": ("#1565c0", "Época lluviosa")}),
          ("zona", {"Alta influencia": ("#8d6e63", "Alta influencia"),
                    "Baja influencia": ("#4fc3f7", "Baja influencia")})]
fig, ax = plt.subplots(2, 3, figsize=(11, 5.6))
for fila, (col, grupos) in enumerate(cortes):
    for columna, (var, etiqueta, _) in enumerate(paneles):
        a = ax[fila][columna]
        datos = dv2.dropna(subset=[var, col])
        bordes = np.histogram_bin_edges(datos[var], bins=12)
        for clave, (color, nombre) in grupos.items():
            x = datos.loc[datos[col] == clave, var]
            a.hist(x, bins=bordes, alpha=.62, color=color, label=f"{nombre} (n = {len(x)})",
                   edgecolor="white", linewidth=.5)
        if var == "br":
            a.axvline(2.5, ls="--", c="#2e7d32", lw=1)
            a.axvline(3.0, ls="--", c="#c62828", lw=1)
        a.set_xlabel(etiqueta)
        if columna == 0:
            a.set_ylabel("Número de estaciones")
        a.legend(fontsize=6.5, frameon=False)
ax[0][1].set_title("Distribución por época climática", fontsize=10)
ax[1][1].set_title("Distribución por zona de influencia del Canal del Dique", fontsize=10)
plt.tight_layout(); plt.savefig("fig4_histogramas.png", dpi=220); plt.close()
print("figura de histogramas generada")
