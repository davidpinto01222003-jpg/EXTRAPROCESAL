"""Analisis estadistico completo: EDA, bivariado, multivariado y pruebas de hipotesis."""
import numpy as np, pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, dendrogram, fcluster
from scipy.spatial.distance import pdist

pd.set_option("display.width", 200)
d = pd.read_csv("base_unificada.csv")

# ---------------- control de calidad ----------------
d["br"] = d.br440_compendio           # indice principal aphy(440)/aphy(676)
d["br443"] = d.br443_compendio        # aphy(443)/aphy(676)

# criterios de exclusion propuestos
d["qc_rango"] = (d.br >= 1.0) & (d.br <= 6.0)
d["qc_turbidez"] = d.turbidez < 60
d["qc_temp"] = d.temperatura.between(26, 34)
d["valido"] = d.qc_rango & d.br.notna()

print("=== 0. CONTROL DE CALIDAD ===")
print("registros totales:", len(d))
print("sin indice (errores de hoja):", d.br.isna().sum())
print("indice fuera de 1 a 6:", int((~d.qc_rango & d.br.notna()).sum()))
print("validos:", int(d.valido.sum()))
print(d.loc[d.br.notna() & ~d.qc_rango, ["campana","estacion","br","temperatura","salinidad","turbidez"]].to_string(index=False))
print("\ntemperatura fuera de 26 a 34 C:", d.loc[d.temperatura.notna() & ~d.qc_temp, ["campana","estacion","temperatura"]].to_string(index=False))

# ---------------- clasificacion de tamanos ----------------
def clase(b):
    if pd.isna(b): return np.nan
    if b > 3.0: return "Pico"
    if b < 2.5: return "Micro"
    return "Nano"
d["clase"] = d.br.map(clase)
d["clase443"] = d.br443.map(clase)
d["clase_qc"] = d.where(d.valido).br.map(clase)

print("\n=== 1. CLASES DE TAMANO (sin filtro, como en el compendio) ===")
t = pd.crosstab(d.campana, d.clase)
print(t)
print("\nporcentajes por campana:")
print((t.T / t.sum(1) * 100).T.round(1))
print("\n=== clases con control de calidad aplicado ===")
tq = pd.crosstab(d.campana, d.clase_qc)
print(tq)
print((tq.T / tq.sum(1) * 100).T.round(1))
print("\npor epoca (con QC):")
te = pd.crosstab(d.epoca, d.clase_qc)
print(te); print((te.T/te.sum(1)*100).T.round(1))

# discrepancia entre los dos pares de longitudes de onda
both = d.dropna(subset=["clase","clase443"])
print("\ncoincidencia de clase entre 440/676 y 443/676: %.1f%% (%d de %d)" %
      (100*(both.clase==both.clase443).mean(), (both.clase==both.clase443).sum(), len(both)))

# ---------------- zonas ----------------
d["zona"] = np.where(d.salinidad < 25, "Alta influencia", "Baja influencia")
d.loc[d.salinidad.isna(), "zona"] = np.nan
print("\nestaciones por zona:", d.zona.value_counts().to_dict())

V = ["temperatura","salinidad","turbidez","br"]
dv = d[d.valido].copy()

# ---------------- 2. EDA ----------------
print("\n=== 2. ESTADISTICA DESCRIPTIVA (registros validos) ===")
desc = dv.groupby("campana")[V].agg(["count","mean","median","std","min","max"]).round(2)
print(desc.to_string())
print("\nglobal:")
print(dv[V].describe().round(2).to_string())
print("\npor epoca:")
print(dv.groupby("epoca")[V].agg(["count","mean","median","std"]).round(2).to_string())

# ---------------- 3. supuestos ----------------
print("\n=== 3. SUPUESTOS ESTADISTICOS ===")
for v in V:
    x = dv[v].dropna()
    W, p = stats.shapiro(x)
    print(f"Shapiro-Wilk {v:12s} n={len(x):3d} W={W:.3f} p={p:.4f} -> {'normal' if p>0.05 else 'NO normal'}")
print()
for v in V:
    g = [g[v].dropna().values for _, g in dv.groupby("epoca")]
    W, p = stats.levene(*g, center="median")
    print(f"Levene (epoca) {v:12s} F={W:.3f} p={p:.4f} -> {'homogenea' if p>0.05 else 'NO homogenea'}")
print()
for v in V:
    g = [g[v].dropna().values for _, g in dv.groupby("campana") if len(g[v].dropna())>2]
    W, p = stats.levene(*g, center="median")
    print(f"Levene (campana) {v:12s} F={W:.3f} p={p:.4f}")

# ---------------- 4. bivariado ----------------
print("\n=== 4. CORRELACIONES DE SPEARMAN (validos) ===")
def corrtab(data, label):
    print(f"\n--- {label} (n={len(data)}) ---")
    for a in ["temperatura","salinidad","turbidez"]:
        s = data[[a,"br"]].dropna()
        rho, p = stats.spearmanr(s[a], s.br)
        r, pp = stats.pearsonr(s[a], s.br)
        print(f"  br vs {a:12s} n={len(s):3d} rho={rho:+.3f} p={p:.4f} | Pearson r={r:+.3f} p={pp:.4f}")
corrtab(dv, "todos los validos")
for camp, s in dv.groupby("epoca"):
    corrtab(s, f"epoca {camp}")
print("\nmatriz de Spearman entre variables ambientales:")
print(dv[V].corr(method="spearman").round(3).to_string())

# correlacion con turbidez en escala logaritmica
s = dv[["turbidez","br"]].dropna()
rho, p = stats.spearmanr(np.log10(s.turbidez.replace(0, 0.01)), s.br)
print(f"\nbr vs log10(turbidez): rho={rho:+.3f} p={p:.4f}")

# ---------------- 5. pruebas de hipotesis ----------------
print("\n=== 5. PRUEBAS DE HIPOTESIS ===")
a = dv[dv.epoca=="Seca"].br.dropna(); b = dv[dv.epoca=="Humeda"].br.dropna()
U, p = stats.mannwhitneyu(a, b, alternative="two-sided")
print(f"Mann-Whitney seca vs humeda (br): n1={len(a)} n2={len(b)} U={U:.1f} p={p:.4f}")
print(f"   medianas: seca={a.median():.3f}  humeda={b.median():.3f}")
# tamano del efecto (r = Z/sqrt(N))
z = (U - len(a)*len(b)/2) / np.sqrt(len(a)*len(b)*(len(a)+len(b)+1)/12)
print(f"   tamano del efecto r={abs(z)/np.sqrt(len(a)+len(b)):.3f}")

groups = [g.br.dropna().values for _, g in dv.groupby("campana")]
H, p = stats.kruskal(*groups)
print(f"\nKruskal-Wallis entre las cinco campanas (br): H={H:.3f} gl={len(groups)-1} p={p:.4f}")
print("   medianas:", dv.groupby("campana").br.median().round(3).to_dict())

# comparaciones por pares con correccion de Bonferroni
camps = sorted(dv.campana.unique())
print("\n   comparaciones por pares (Mann-Whitney, Bonferroni):")
pares = []
for i in range(len(camps)):
    for j in range(i+1, len(camps)):
        x = dv[dv.campana==camps[i]].br.dropna(); y = dv[dv.campana==camps[j]].br.dropna()
        U, pp = stats.mannwhitneyu(x, y, alternative="two-sided")
        pares.append((camps[i], camps[j], pp))
k = len(pares)
for a1,b1,pp in pares:
    print(f"     {a1:12s} vs {b1:12s} p={pp:.4f}  p_ajustada={min(1,pp*k):.4f} {'*' if pp*k<0.05 else ''}")

for v in ["temperatura","salinidad","turbidez"]:
    x = dv[dv.epoca=="Seca"][v].dropna(); y = dv[dv.epoca=="Humeda"][v].dropna()
    U, pp = stats.mannwhitneyu(x, y, alternative="two-sided")
    print(f"\nMann-Whitney seca vs humeda ({v}): p={pp:.4f} medianas {x.median():.2f} / {y.median():.2f}")

# zonas
x = dv[dv.zona=="Alta influencia"].br.dropna(); y = dv[dv.zona=="Baja influencia"].br.dropna()
U, pp = stats.mannwhitneyu(x, y, alternative="two-sided")
print(f"\nMann-Whitney zona alta vs baja influencia (br): n1={len(x)} n2={len(y)} p={pp:.4f}")
print(f"   medianas: alta={x.median():.3f} baja={y.median():.3f}")

# tabla de contingencia clase x epoca
ct = pd.crosstab(dv.epoca, dv.clase_qc)
chi2, p, gl, exp = stats.chi2_contingency(ct)
print(f"\nChi cuadrado clase x epoca: chi2={chi2:.3f} gl={gl} p={p:.4f}")
print(ct)
print("frecuencias esperadas:\n", pd.DataFrame(exp, index=ct.index, columns=ct.columns).round(1))

# ---------------- 6. multivariado ----------------
print("\n=== 6. ANALISIS MULTIVARIADO ===")
M = dv.dropna(subset=["temperatura","salinidad","turbidez","br"]).copy()
X = M[["temperatura","salinidad","turbidez","br"]].copy()
X["turbidez"] = np.log10(X.turbidez.replace(0, 0.05))
Z = (X - X.mean()) / X.std(ddof=1)
C = np.cov(Z.T, ddof=1)
val, vec = np.linalg.eigh(C)
idx = np.argsort(val)[::-1]; val = val[idx]; vec = vec[:, idx]
print("n =", len(M))
print("varianza explicada:", (val/val.sum()*100).round(1))
print("acumulada:", np.cumsum(val/val.sum()*100).round(1))
cargas = pd.DataFrame(vec[:, :3]*np.sqrt(val[:3]), index=X.columns, columns=["CP1","CP2","CP3"])
print("\ncargas (correlacion variable componente):")
print(cargas.round(3).to_string())
scores = Z.values @ vec
M["CP1"], M["CP2"] = scores[:,0], scores[:,1]

# conglomerados
lk = linkage(Z.values, method="ward")
for k in (2,3,4):
    M[f"grupo{k}"] = fcluster(lk, k, criterion="maxclust")
print("\nconglomerados (Ward, k=3): medias por grupo")
print(M.groupby("grupo3")[["temperatura","salinidad","turbidez","br"]].agg(["count","mean"]).round(2).to_string())
print("\ncomposicion por epoca:")
print(pd.crosstab(M.grupo3, M.epoca))
print("\nclase dominante por grupo:")
print(pd.crosstab(M.grupo3, M.clase_qc))
coph = stats.spearmanr(pdist(Z.values), __import__("scipy.cluster.hierarchy", fromlist=["cophenet"]).cophenet(lk))[0]
print(f"\ncorrelacion cofenetica (Spearman): {coph:.3f}")

# ---------------- 7. modelo descriptivo ----------------
print("\n=== 7. MODELO DESCRIPTIVO PRELIMINAR ===")
mm = M.dropna(subset=["temperatura","salinidad","turbidez","br"]).copy()
mm["logturb"] = np.log10(mm.turbidez.replace(0,0.05))
Xd = np.column_stack([np.ones(len(mm)), mm.temperatura, mm.salinidad, mm.logturb])
y = np.log10(mm.br.values)
beta, res, rank, sv = np.linalg.lstsq(Xd, y, rcond=None)
yhat = Xd @ beta
ss_res = ((y-yhat)**2).sum(); ss_tot = ((y-y.mean())**2).sum()
R2 = 1-ss_res/ss_tot
n, p_ = Xd.shape
R2adj = 1-(1-R2)*(n-1)/(n-p_)
se = np.sqrt(np.diag(ss_res/(n-p_)*np.linalg.inv(Xd.T@Xd)))
tval = beta/se
pval = 2*(1-stats.t.cdf(np.abs(tval), n-p_))
print(f"n={n}  R2={R2:.3f}  R2 ajustado={R2adj:.3f}  RMSE(log10)={np.sqrt(ss_res/n):.3f}")
for nom, b_, s_, t_, pp in zip(["intercepto","temperatura","salinidad","log10 turbidez"], beta, se, tval, pval):
    print(f"   {nom:16s} b={b_:+.4f}  ee={s_:.4f}  t={t_:+.2f}  p={pp:.4f}")
sw = stats.shapiro(y-yhat)
print(f"   residuos Shapiro-Wilk W={sw[0]:.3f} p={sw[1]:.4f}")

d.to_csv("base_analizada.csv", index=False)
M.to_csv("base_multivariado.csv", index=False)

# ---------------- figuras ----------------
plt.rcParams.update({"font.size":9, "figure.dpi":150})
fig, ax = plt.subplots(1, 2, figsize=(9,3.6))
orden = ["Seca 2021","Seca 2022","Seca 2023","Humeda 2022","Humeda 2023"]
dat = [dv[dv.campana==c].br.dropna().values for c in orden]
bp = ax[0].boxplot(dat, tick_labels=[c.replace(" ","\n") for c in orden], patch_artist=True)
for b_ in bp["boxes"]: b_.set_facecolor("#8fbcd4")
ax[0].axhline(3.0, ls="--", c="#c0392b", lw=1); ax[0].axhline(2.5, ls="--", c="#27ae60", lw=1)
ax[0].set_ylabel("Indice B/R  aphy(440)/aphy(676)"); ax[0].set_title("a) Indice por campana")
ax[0].text(5.4, 3.05, "pico", color="#c0392b", fontsize=7); ax[0].text(5.4, 2.35, "micro", color="#27ae60", fontsize=7)
sc = ax[1].scatter(dv.salinidad, dv.br, c=np.log10(dv.turbidez.replace(0,0.05)), cmap="viridis", s=26, edgecolor="k", linewidth=.3)
ax[1].axhline(3.0, ls="--", c="#c0392b", lw=1); ax[1].axhline(2.5, ls="--", c="#27ae60", lw=1)
ax[1].set_xlabel("Salinidad"); ax[1].set_ylabel("Indice B/R"); ax[1].set_title("b) Indice y salinidad")
plt.colorbar(sc, ax=ax[1], label="log10 turbidez")
plt.tight_layout(); plt.savefig("fig_eda.png", dpi=200); plt.close()

fig, ax = plt.subplots(1, 2, figsize=(9,4))
col = {"Seca":"#e67e22","Humeda":"#2980b9"}
for ep, g in M.groupby("epoca"):
    ax[0].scatter(g.CP1, g.CP2, c=col[ep], label=ep, s=30, edgecolor="k", linewidth=.3)
for i, v in enumerate(X.columns):
    ax[0].arrow(0,0, cargas.CP1[v]*2.2, cargas.CP2[v]*2.2, color="k", head_width=.08, lw=1)
    ax[0].text(cargas.CP1[v]*2.5, cargas.CP2[v]*2.5, v, fontsize=8)
ax[0].axhline(0, lw=.5, c="gray"); ax[0].axvline(0, lw=.5, c="gray")
ax[0].set_xlabel(f"CP1 ({val[0]/val.sum()*100:.1f}%)"); ax[0].set_ylabel(f"CP2 ({val[1]/val.sum()*100:.1f}%)")
ax[0].legend(); ax[0].set_title("a) Analisis de componentes principales")
dendrogram(lk, ax=ax[1], no_labels=True, color_threshold=lk[-2,2])
ax[1].set_title("b) Conglomerados de Ward"); ax[1].set_ylabel("Distancia")
plt.tight_layout(); plt.savefig("fig_multivariado.png", dpi=200); plt.close()
print("\nfiguras guardadas")
