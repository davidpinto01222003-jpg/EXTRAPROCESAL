"""Simulador aproximado de Turnitin (detección de IA) calibrado con un informe real,
más análisis de patrones Turnitin vs. Compilatio.

Uso:
    python turnitin.py entrenar ../trabajo_lectoescritura/original.docx ../trabajo_lectoescritura/reporte_turnitin.pdf
    python turnitin.py evaluar "../trabajo_lectoescritura/Trabajo lectoescritura v2.docx" --html reporte_turnitin_v2.html
    python turnitin.py patrones   # compara qué marcó Turnitin y qué marcó Compilatio
"""
import argparse
import html
import math
import pickle
import re
from collections import Counter
from pathlib import Path

import docx
import numpy as np
import pdfplumber
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import GroupKFold, cross_val_predict
from sklearn.metrics import roc_auc_score
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

import simulador as C  # funciones compartidas con el simulador de Compilatio

AQUI = Path(__file__).parent
MODELO = AQUI / "modelo_turnitin.pkl"
CIAN = "(0.3203125, 0.77734375, 0.85546875)"

# Secciones que Turnitin no evalúa como prosa (transcripción de la prueba, referencias)
INICIO_FUERA = re.compile(r"^(Transcipción|Transcripción) de prueba|^Citas y referencias|^Referencias")
FIN_FUERA = re.compile(r"^Anexos$")


def parrafos(ruta):
    d = docx.Document(ruta)
    return [C.norm(p.text) for p in d.paragraphs]


def cuerpo(pars):
    out, fuera = [], False
    for i, p in enumerate(pars):
        if INICIO_FUERA.search(p):
            fuera = True
            if p.startswith(("Citas", "Referencias")):
                break
        elif FIN_FUERA.search(p):
            fuera = False
            continue
        if not fuera and len(p.split()) >= 8:
            out.append((i, p))
    return out


def lineas_cian(pdf):
    out = []
    with pdfplumber.open(pdf) as doc:
        for pg in doc.pages[2:]:
            words = [w for w in pg.extract_words() if 40 < w["top"] < pg.height - 45]
            for r in pg.rects:
                if str(r.get("non_stroking_color")) != CIAN:
                    continue
                ws = [w for w in words
                      if r["top"] - 1 <= (w["top"] + w["bottom"]) / 2 <= r["bottom"] + 1
                      and r["x0"] - 1 <= (w["x0"] + w["x1"]) / 2 <= r["x1"] + 1]
                if ws:
                    out.append(" ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])))
    return out


def etiquetar(pc, lineas):
    sh = set()
    for l in lineas:
        sh |= C.shingles(l)
    X, y, meta = [], [], []
    for i, p in pc:
        for s in C.oraciones(p):
            ss = C.shingles(s)
            if not ss:
                continue
            X.append(C.rasgos(s, p))
            y.append(1 if len(ss & sh) / len(ss) >= 0.5 else 0)
            meta.append((i, s))
    return np.array(X), np.array(y), meta


def entrenar(docx_path, pdf_path, ref_pct):
    pars = parrafos(docx_path)
    pc = cuerpo(pars)
    X, y, meta = etiquetar(pc, lineas_cian(pdf_path))
    grupos = [m[0] for m in meta]
    modelo = make_pipeline(StandardScaler(), LogisticRegression(max_iter=3000, class_weight="balanced"))
    prob = cross_val_predict(modelo, X, y, cv=GroupKFold(5), groups=grupos, method="predict_proba")[:, 1]
    modelo.fit(X, y)
    palabras = np.array([len(m[1].split()) for m in meta])
    total = palabras.sum()
    real = (palabras * y).sum() / total * 100
    umbral = min(np.linspace(0.3, 0.8, 51), key=lambda u: abs((palabras * (prob >= u)).sum() / total * 100 - real))
    pickle.dump({"modelo": modelo, "umbral": float(umbral), "ref_pct": ref_pct, "ref_modelo": float(real)},
                open(MODELO, "wb"))
    print(f"Oraciones: {len(y)} (marcadas: {y.sum()}) | % marcado en prosa: {real:.1f} % (Turnitin: {ref_pct} %)")
    print(f"AUC validación cruzada por párrafos: {roc_auc_score(y, prob):.2f} | umbral {umbral:.2f}")


TEXTO_ORIGINAL = AQUI.parent / "trabajo_lectoescritura" / "original.docx"
REPORTE_ORIGINAL = AQUI.parent / "trabajo_lectoescritura" / "reporte_turnitin.pdf"


def sin_cambios(docx_path):
    """% de los fragmentos que Turnitin marcó en el original que siguen textuales en la versión nueva."""
    marcadas = set()
    for l in lineas_cian(REPORTE_ORIGINAL):
        marcadas |= C.shingles(l, 4)
    base = marcadas & C.shingles(" ".join(parrafos(TEXTO_ORIGINAL)), 4)
    nueva = C.shingles(" ".join(parrafos(docx_path)), 4)
    return len(base & nueva) / max(1, len(base)) * 100


def evaluar(docx_path, html_out=None):
    M = pickle.load(open(MODELO, "rb"))
    pars = parrafos(docx_path)
    filas = []
    for i, p in cuerpo(pars):
        for s in C.oraciones(p):
            pr = M["modelo"].predict_proba(np.array([C.rasgos(s, p)]))[0, 1]
            filas.append((i, s, pr, pr >= M["umbral"]))
    total = sum(len(s.split()) for _, s, _, _ in filas)
    marc = sum(len(s.split()) for _, s, _, f in filas if f)
    pct = marc / total * 100 * (M["ref_pct"] / M["ref_modelo"])
    print(f"IA estimada (Turnitin) ... {pct:5.1f} %  (modelo de estilo, aproximado)")
    if TEXTO_ORIGINAL.exists() and REPORTE_ORIGINAL.exists():
        sigue = sin_cambios(docx_path)
        print(f"Texto marcado en el informe real que sigue igual: {sigue:.0f} % "
              f"(si Turnitin solo volviera a marcar eso: ~{M['ref_pct'] * sigue / 100:.0f} %)")
    por = {}
    for i, s, pr, f in filas:
        a = por.setdefault(i, [0, 0, []]); a[0] += len(s.split()) * f; a[1] += len(s.split()); a[2].append((s, pr, f))
    for i, (w, n, _) in sorted(por.items(), key=lambda kv: -kv[1][0])[:15]:
        if w:
            print(f"  [{i}] {w}/{n} — {pars[i][:80]}…")
    if html_out:
        cuerpo_html = "\n".join(
            "<p><small>[%d]</small> %s</p>" % (i, " ".join(
                f'<span class="{"ia" if f else ""}" title="prob. IA {pr:.0%}">{html.escape(s)}</span>' for s, pr, f in ss))
            for i, (_, _, ss) in sorted(por.items()))
        Path(html_out).write_text(C.PLANTILLA.format(
            nombre=html.escape(Path(docx_path).name), total=pct, ia=pct, sim=0.0, idioma=0.0,
            palabras=total, cuerpo=cuerpo_html).replace("Simulación tipo Compilatio", "Simulación tipo Turnitin (IA)")
            .replace("informe real de Compilatio", "informe real de Turnitin"), encoding="utf8")
        print(f"Reporte HTML: {html_out}")
    return pct


# ---------------------------------------------------------------- patrones

MARCADORES = {
    "Conectores de apertura (Asimismo, Por consiguiente, En este sentido, Sin embargo…)":
        r"^(asimismo|por consiguiente|en este sentido|sin embargo|no obstante|en consecuencia|por lo tanto|por ello|"
        r"por otro lado|de igual manera|de igual forma|además|finalmente|en síntesis|en conclusión|para concluir|"
        r"ahora bien|por su parte|en efecto|así pues|por otra parte|simultáneamente)\b",
    "Verbos-comodín (permite, evidencia, constituye, resulta, destaca, refleja)":
        r"\b(permit\w+|evidenci\w+|constitu\w+|result[ao]n?\b|destac\w+|reflej\w+|subray\w+|resalt\w+)",
    "Adjetivos de relleno (fundamental, esencial, relevante, significativo, clave, crucial, importante)":
        r"\b(fundamental\w*|esencial\w*|relevante\w*|significativ\w+|clave|crucial\w*|importan\w+|pertinente\w*|determinante\w*)",
    "Metadiscurso sobre el texto (el documento, el trabajo, el marco teórico, la investigación)":
        r"\b(el (presente )?(documento|trabajo|estudio)|el marco teórico|la investigación|este análisis|los resultados)\b",
    "Enumeraciones de tres o más elementos (X, Y y Z)":
        r"\b\w+, \w+(?: \w+)?,? y \w+",
    "Dos puntos o punto y coma": r"[:;]",
    "Primera persona (planteamos, encontramos, creo, nuestro)":
        r"\b(\w+amos|\w+emos|creo|considero|nuestr\w+|nosotr\w+|mi|me)\b",
    "Citas con autor y año": r"\(\D*\d{4}[a-z]?[^)]*\)|\w+ \(\d{4}[a-z]?\)",
}


def tasas(oraciones_):
    palabras = sum(len(s.split()) for s in oraciones_) or 1
    out = {}
    for k, rx in MARCADORES.items():
        n = sum(len(re.findall(rx, s.lower() if k.startswith("Conectores") else s, re.I | re.M)) for s in oraciones_)
        out[k] = n / palabras * 1000
    largos = [len(s.split()) for s in oraciones_] or [0]
    out["Largo medio de oración (palabras)"] = float(np.mean(largos))
    return out


def palabras_distintivas(marcadas, limpias, k=25):
    a = Counter(t for s in marcadas for t in C.tokens(s))
    b = Counter(t for s in limpias for t in C.tokens(s))
    na, nb = sum(a.values()), sum(b.values())
    voc = {w for w in set(a) | set(b) if a[w] + b[w] >= 6 and len(w) > 3}
    alpha = 0.5
    def z(w):
        la = math.log((a[w] + alpha) / (na + alpha * len(voc) - a[w] - alpha))
        lb = math.log((b[w] + alpha) / (nb + alpha * len(voc) - b[w] - alpha))
        return (la - lb) / math.sqrt(1 / (a[w] + alpha) + 1 / (b[w] + alpha))
    return sorted(voc, key=z, reverse=True)[:k]


def patrones(salida):
    # Turnitin: trabajo de lectoescritura
    t_docx = AQUI.parent / "trabajo_lectoescritura" / "original.docx"
    t_pdf = AQUI.parent / "trabajo_lectoescritura" / "reporte_turnitin.pdf"
    _, yt, mt = etiquetar(cuerpo(parrafos(t_docx)), lineas_cian(t_pdf))
    # Compilatio: trabajo de Suárez y Rivera
    pc = C.cuerpo(C.leer_docx(AQUI / "datos" / "original.docx"))
    ia, _, _ = C.marcas_pdf(AQUI / "datos" / "reporte_compilatio.pdf")
    _, yc, mc = C.etiquetar(pc, ia)

    tm = [s for (_, s), f in zip(mt, yt) if f]; tl = [s for (_, s), f in zip(mt, yt) if not f]
    cm = [s for (_, s), f in zip(mc, yc) if f]; cl = [s for (_, s), f in zip(mc, yc) if not f]
    rt_m, rt_l, rc_m, rc_l = tasas(tm), tasas(tl), tasas(cm), tasas(cl)

    lineas = ["# Patrones que marcan Turnitin y Compilatio", "",
              "Comparación de las oraciones **marcadas como IA** contra las **no marcadas** en los dos informes reales que tenemos:",
              "- Turnitin, trabajo de lectoescritura (34 % de IA).",
              "- Compilatio, trabajo de Suárez y Rivera (27 % de IA).", "",
              "Las tasas se dan por cada 1.000 palabras, salvo el largo de oración.", "",
              "| Rasgo | Turnitin: marcadas | Turnitin: no marcadas | Compilatio: marcadas | Compilatio: no marcadas |",
              "|---|---|---|---|---|"]
    for k in rt_m:
        lineas.append(f"| {k} | {rt_m[k]:.1f} | {rt_l[k]:.1f} | {rc_m[k]:.1f} | {rc_l[k]:.1f} |")
    pt, pc_ = palabras_distintivas(tm, tl), palabras_distintivas(cm, cl)
    comunes = [w for w in pt if w in set(palabras_distintivas(cm, cl, 60))]
    lineas += ["", "## Palabras que más aparecen en lo marcado", "",
               f"- **Turnitin:** {', '.join(pt)}",
               f"- **Compilatio:** {', '.join(pc_)}",
               f"- **Comunes a ambos:** {', '.join(comunes) or '—'}", ""]
    Path(salida).write_text("\n".join(lineas) + "\n", encoding="utf8")
    print("\n".join(lineas))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("entrenar"); e.add_argument("docx"); e.add_argument("pdf"); e.add_argument("--pct", type=float, default=34.0)
    v = sub.add_parser("evaluar"); v.add_argument("docx"); v.add_argument("--html")
    p = sub.add_parser("patrones"); p.add_argument("--salida", default=str(AQUI.parent / "trabajo_lectoescritura" / "PATRONES_TURNITIN_COMPILATIO.md"))
    a = ap.parse_args()
    if a.cmd == "entrenar":
        entrenar(a.docx, a.pdf, a.pct)
    elif a.cmd == "evaluar":
        evaluar(a.docx, a.html)
    else:
        patrones(a.salida)
