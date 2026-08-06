import pandas as pd, numpy as np
d = pd.read_csv("base_unificada.csv")
d["br443_676"] = d.aphy443 / d.aphy676
d["br443_677"] = d.aphy443 / d.aphy677
for camp, s in d.groupby("campana"):
    s = s.dropna(subset=["br443_compendio", "aphy443"])
    if not len(s):
        continue
    e676 = np.nanmax(np.abs(s.br443_compendio - s.br443_676))
    e677 = np.nanmax(np.abs(s.br443_compendio - s.br443_677))
    e440 = np.nanmax(np.abs(s.br440_compendio - s.aphy440 / s.aphy676))
    print(f"{camp:14s} n={len(s):3d}  maxdif 443/676={e676:.4f}  443/677={e677:.4f}   440/676={e440:.5f}")

print("\nBanderas de error en el compendio:")
print(d[d.br440_flag.notna() & (d.br440_flag != "")].groupby(["campana", "br440_flag"]).size())

print("\nRecuperables con espectro (error en compendio pero aphy disponible):")
m = d.br440_compendio.isna() & d.aphy440.notna()
print(d.loc[m, ["campana", "estacion", "br440_flag", "aphy440", "aphy676"]].to_string())

print("\naphy676 <= 0 o muy bajo:")
print(d.loc[d.aphy676.notna() & (d.aphy676 <= 0.005), ["campana", "estacion", "aphy440", "aphy676", "br440_compendio", "turbidez"]].to_string())

print("\nB/R fuera de 1 a 6 (implausibles):")
print(d.loc[(d.br440_compendio > 6) | (d.br440_compendio < 1), ["campana", "estacion", "br440_compendio", "temperatura", "salinidad", "turbidez"]].to_string())

print("\nDatos faltantes por variable:")
print(d[["temperatura", "salinidad", "turbidez", "br440_compendio", "br443_compendio"]].isna().sum())
print("\nceros sospechosos:")
for v in ["temperatura", "salinidad", "turbidez"]:
    print(v, list(d.loc[d[v] == 0, "estacion"]))
