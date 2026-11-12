"""Simulador aproximado de Compilatio (IA + similitud) calibrado con un informe real.

No es Compilatio: aprende qué frases marcó Compilatio en un informe real y
estima cómo marcaría una versión nueva del mismo documento.

Uso:
    python simulador.py entrenar datos/original.docx datos/reporte_compilatio.pdf
    python simulador.py evaluar "ruta/al/trabajo.docx" --html reporte.html
"""
import argparse
import html
import json
import pickle
import re
import statistics
import sys
from pathlib import Path

import docx
import numpy as np
import pdfplumber
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import cross_val_predict
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import make_pipeline

AQUI = Path(__file__).parent
MODELO = AQUI / "modelo.pkl"

AZUL_IA = "(0.0, 0.749, 1.0)"
ROSA_IDIOMA = "(1.0, 0.3569, 0.7059)"
BLANCO = "(1.0, 1.0, 1.0)"
GRIS_TABLA = "(0.1169, 0.1608, 0.222)"

CONECTORES = [
    "conviene", "esto es", "de ahí", "en consecuencia", "por su parte", "asimismo",
    "en ese sentido", "cabe", "resulta", "se trata de", "no solo", "sino también",
    "por lo anterior", "en efecto", "a su turno", "en este sentido", "ahora bien",
    "de modo que", "de manera que", "en la medida en que", "por tanto", "así,",
    "en primer término", "la primera", "la segunda", "la tercera", "en síntesis",
    "en suma", "finalmente", "además", "sin embargo", "por el contrario",
    "en particular", "es decir", "a juicio de", "en el plano", "frente a",
]
GENERICAS = [
    "fundamental", "clave", "relevante", "central", "estructural", "integral",
    "sostenible", "efectiva", "eficacia", "articulación", "articulada", "abordar",
    "dimensión", "fenómeno", "ámbito", "marco", "garantizar", "coordinación",
    "consolidado", "significativ", "robust", "precisamente", "evidenci",
]
HUMANAS = [
    "nosotros", "creemos", "revisamos", "encontramos", "nuestro", "nuestra",
    "a nuestro juicio", "es cierto", "eso sí", "o sea", "pero ", "y justo",
    "?", "¿", "pasa", "sale", "queda", "casi",
]


def norm(s):
    return re.sub(r"\s+", " ", s).strip()


def tokens(s):
    return re.findall(r"\w+", s.lower())


def shingles(s, n=3):
    t = tokens(s)
    return {tuple(t[i:i + n]) for i in range(max(0, len(t) - n + 1))}


def oraciones(parrafo):
    partes = re.split(r"(?<=[.;?!])\s+(?=[A-ZÁÉÍÓÚÑ¿“(])", parrafo)
    return [p for p in (norm(x) for x in partes) if p]


def leer_docx(ruta):
    d = docx.Document(ruta)
    pars = [norm(p.text) for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                pars.append(norm(c.text))
    return [p for p in pars if p]


def cuerpo(pars):
    """Párrafos que cuentan para IA: prosa (>=25 palabras), antes de la bibliografía."""
    fuera = False
    out = []
    for i, p in enumerate(pars):
        if re.fullmatch(r"(Referencias|Bibliografía|REFERENCIAS)", p):
            fuera = True
        if not fuera and len(p.split()) >= 25:
            out.append((i, p))
    return out


# ---------------------------------------------------------------- etiquetas

def marcas_pdf(pdf):
    ia, sim, idioma = [], [], []
    with pdfplumber.open(pdf) as doc:
        for pg in doc.pages[5:]:
            words = pg.extract_words()
            for r in pg.rects:
                col = str(r.get("non_stroking_color"))
                if r["x0"] < 60 or col in (BLANCO, GRIS_TABLA):
                    continue
                if col in (AZUL_IA, ROSA_IDIOMA):
                    ws = [w for w in words if 0 <= r["top"] - w["bottom"] <= 6
                          and w["x0"] >= r["x0"] - 2 and w["x1"] <= r["x1"] + 2]
                    (ia if col == AZUL_IA else idioma).append(" ".join(w["text"] for w in ws))
                elif r["bottom"] - r["top"] >= 5:
                    ws = [w for w in words if w["top"] >= r["top"] - 2 and w["bottom"] <= r["bottom"] + 2
                          and w["x0"] >= r["x0"] - 2 and w["x1"] <= r["x1"] + 2]
                    sim.append(" ".join(w["text"] for w in ws))
    return [x for x in ia if x], [x for x in sim if len(x.split()) >= 4], [x for x in idioma if x]


def etiquetar(pars_cuerpo, lineas_ia):
    sh_ia = set()
    for l in lineas_ia:
        sh_ia |= shingles(l)
    X, y, meta = [], [], []
    for i, p in pars_cuerpo:
        for s in oraciones(p):
            sh = shingles(s)
            if not sh:
                continue
            frac = len(sh & sh_ia) / len(sh)
            X.append(rasgos(s, p))
            y.append(1 if frac >= 0.5 else 0)
            meta.append((i, s))
    return np.array(X), np.array(y), meta


# ---------------------------------------------------------------- rasgos

def rasgos(s, p):
    ts = tokens(s)
    n = max(1, len(ts))
    low = s.lower()
    largos_p = [len(tokens(o)) for o in oraciones(p)] or [n]
    cv = statistics.pstdev(largos_p) / (statistics.mean(largos_p) or 1)
    return [
        n,
        np.log1p(n),
        s.count(",") / n,
        s.count(";") + s.count(":"),
        len(re.findall(r"\(\w[^)]*\d{4}[a-z]?\)", s)),
        sum(low.count(c) for c in CONECTORES),
        sum(low.count(g) for g in GENERICAS) / n * 10,
        sum(low.count(h) for h in HUMANAS),
        len(set(ts)) / n,
        np.mean([len(t) for t in ts]) if ts else 0,
        sum(len(t) >= 10 for t in ts) / n,
        len(re.findall(r"\bse \w+", low)) / n * 10,
        len(re.findall(r"\b(y|o)\b", low)) / n * 10,
        cv,
        len(p.split()),
        1 if re.search(r"\d", s) else 0,
    ]


# ---------------------------------------------------------------- comandos

def entrenar(docx_path, pdf_path):
    pars = leer_docx(docx_path)
    pc = cuerpo(pars)
    ia, sim, idioma = marcas_pdf(pdf_path)
    X, y, meta = etiquetar(pc, ia)
    modelo = make_pipeline(StandardScaler(), LogisticRegression(max_iter=2000, class_weight="balanced"))
    prob_cv = cross_val_predict(modelo, X, y, cv=5, method="predict_proba")[:, 1]
    modelo.fit(X, y)
    palabras = np.array([len(m[1].split()) for m in meta])
    total_doc = sum(len(p.split()) for p in pars)
    real_pct = (palabras * y).sum() / total_doc * 100
    # umbral que reproduce el % real en validación cruzada
    mejor = min(np.linspace(0.3, 0.8, 51),
                key=lambda u: abs((palabras * (prob_cv >= u)).sum() / total_doc * 100 - real_pct))
    pred = prob_cv >= mejor
    acierto = (pred == y).mean()
    pickle.dump({"modelo": modelo, "umbral": float(mejor), "sim": sim,
                 "ref_ia_pct": 27.0, "ref_sim_pct": 6.0, "ref_idioma_pct": 3.0,
                 "ref_ia_modelo": float(real_pct)}, open(MODELO, "wb"))
    print(f"Oraciones de entrenamiento: {len(y)} (marcadas IA: {y.sum()})")
    print(f"% IA según etiquetas del informe: {real_pct:.1f} %  (Compilatio: 27 %)")
    print(f"Umbral calibrado: {mejor:.2f} | acierto por oración (val. cruzada): {acierto:.0%}")


def evaluar(docx_path, html_out=None):
    M = pickle.load(open(MODELO, "rb"))
    pars = leer_docx(docx_path)
    total_doc = sum(len(p.split()) for p in pars)
    filas = []
    for i, p in cuerpo(pars):
        for s in oraciones(p):
            prob = M["modelo"].predict_proba(np.array([rasgos(s, p)]))[0, 1]
            filas.append((i, s, prob, prob >= M["umbral"]))
    ia_words = sum(len(s.split()) for _, s, _, f in filas if f)
    ia_pct = ia_words / total_doc * 100 * (M["ref_ia_pct"] / M["ref_ia_modelo"])

    # similitud: fragmentos que Compilatio encontró y siguen fuera de comillas
    texto = " ".join(pars)
    sin_comillas = re.sub(r"“[^”]*”|\"[^\"]*\"", " ", texto)
    sh_doc = shingles(sin_comillas, 5)
    quedan = [f for f in M["sim"] if shingles(f, 5) and len(shingles(f, 5) & sh_doc) / len(shingles(f, 5)) >= 0.6]
    base = [f for f in M["sim"] if shingles(f, 5)]
    sim_pct = M["ref_sim_pct"] * (sum(len(f.split()) for f in quedan) / max(1, sum(len(f.split()) for f in base)))
    idioma_pct = M["ref_idioma_pct"]
    total = ia_pct + sim_pct + idioma_pct

    # por párrafo
    por_par = {}
    for i, s, prob, f in filas:
        a = por_par.setdefault(i, [0, 0, []])
        a[0] += len(s.split()) * f
        a[1] += len(s.split())
        a[2].append((s, prob, f))
    peores = sorted(por_par.items(), key=lambda kv: -kv[1][0])[:15]

    print(f"Palabras: {total_doc}")
    print(f"IA estimada ........... {ia_pct:5.1f} %")
    print(f"Similitud estimada .... {sim_pct:5.1f} %")
    print(f"Idiomas no reconocidos  {idioma_pct:5.1f} %  (se asume igual al informe)")
    print(f"TOTAL estimado ........ {total:5.1f} %")
    print("\nPárrafos con más texto marcado como IA:")
    for i, (w, n, _) in peores:
        if w:
            print(f"  [{i}] {w}/{n} palabras — {pars[i][:80]}…")

    if html_out:
        partes = []
        for i, (w, n, ss) in sorted(por_par.items()):
            frag = " ".join(
                f'<span class="{"ia" if f else ""}" title="prob. IA {prob:.0%}">{html.escape(s)}</span>'
                for s, prob, f in ss)
            partes.append(f'<p><small>[{i}]</small> {frag}</p>')
        Path(html_out).write_text(PLANTILLA.format(
            nombre=html.escape(Path(docx_path).name), total=total, ia=ia_pct, sim=sim_pct,
            idioma=idioma_pct, palabras=total_doc, cuerpo="\n".join(partes)), encoding="utf8")
        print(f"\nReporte HTML: {html_out}")
    return {"ia": ia_pct, "sim": sim_pct, "idioma": idioma_pct, "total": total}


PLANTILLA = """<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Simulación Compilatio</title><style>
:root{{--bg:#fff;--fg:#1d1d1f;--mut:#666;--ia:#cdeefe;--line:#00a6e6;--card:#f5f6f8}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--bg:#16181b;--fg:#e8e8ea;--mut:#9a9aa0;--ia:#0b3a4d;--line:#39c1f5;--card:#212428}}}}
body{{background:var(--bg);color:var(--fg);font:15px/1.6 system-ui,sans-serif;max-width:900px;margin:0 auto;padding:16px}}
.kpis{{display:flex;flex-wrap:wrap;gap:12px;margin:16px 0}}.k{{background:var(--card);border-radius:10px;padding:12px 16px;flex:1;min-width:140px}}
.k b{{display:block;font-size:28px}}.ia{{background:var(--ia);border-bottom:2px solid var(--line)}}small{{color:var(--mut)}}
</style></head><body><h1>Simulación tipo Compilatio</h1>
<p><small>{nombre} · {palabras} palabras · estimación aproximada calibrada con el informe real de Compilatio; no reemplaza el análisis oficial.</small></p>
<div class="kpis"><div class="k">Total estimado<b>{total:.1f} %</b></div><div class="k">IA<b>{ia:.1f} %</b></div>
<div class="k">Similitud<b>{sim:.1f} %</b></div><div class="k">Idiomas no reconocidos<b>{idioma:.1f} %</b></div></div>
<p><span class="ia">Texto resaltado</span> = frase que el simulador marcaría como IA (pasa el cursor para ver la probabilidad).</p>
{cuerpo}</body></html>"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("entrenar"); e.add_argument("docx"); e.add_argument("pdf")
    v = sub.add_parser("evaluar"); v.add_argument("docx"); v.add_argument("--html")
    a = ap.parse_args()
    if a.cmd == "entrenar":
        entrenar(a.docx, a.pdf)
    else:
        evaluar(a.docx, a.html)
