# Compara la prosa de la entrega (texto extraído del informe de Turnitin) con la versión nueva.
# Uso: python comparar.py "Trabajo LinkedIn revisado v2.docx"
import json
import re
import statistics as st
import sys
from pathlib import Path

import docx

AQUI = Path(__file__).parent
P = json.load(open(AQUI / "datos/parrafos.json"))
CONECTORES = r"\b(while|whereas|since|because|although|which is why|for that reason|so that|therefore|in light of|on that basis|in practical terms|considering that|given that|bearing in mind)\b"


def oraciones(t):
    return [s for s in re.split(r"(?<=[.?!])\s+(?=[A-Z“\"])", t) if len(s.split()) > 3]


def prosa_original():
    out = []
    for p in P:
        t = p["text"]
        if t.startswith("##") or p["x0"] < 100 or p["page"] < 4 or 33 <= p["page"] <= 38:
            continue
        plano = t.replace("[[", "").replace("]]", "")
        if len(plano.split()) >= 25:
            out.append(plano)
    return out


def prosa_nueva(ruta):
    d = docx.Document(ruta)
    out, en_refs = [], False
    for p in d.paragraphs:
        t = p.text.strip()
        if t == "References":
            en_refs = True
        elif t.startswith("Appendix"):
            en_refs = False
        if not en_refs and len(t.split()) >= 25 and not t.startswith("Note."):
            out.append(t)
    return out


def tejas(t, n=5):
    w = re.findall(r"[a-z0-9’']+", t.lower())
    return {" ".join(w[i:i + n]) for i in range(len(w) - n + 1)}


def medidas(pars):
    S = [s for p in pars for s in oraciones(p)]
    L = [len(s.split()) for s in S]
    txt = " ".join(pars)
    return {
        "Palabras de prosa": sum(len(p.split()) for p in pars),
        "Oraciones": len(S),
        "Largo medio de oración (palabras)": round(st.mean(L), 1),
        "Oraciones de más de 40 palabras": sum(l > 40 for l in L),
        "Oraciones de menos de 15 palabras": sum(l < 15 for l in L),
        "Variación del largo (CV)": round(st.pstdev(L) / st.mean(L), 2),
        "Conectores subordinantes por oración": round(len(re.findall(CONECTORES, txt, re.I)) / len(S), 2),
        "“I will / I would”": len(re.findall(r"\bI (will|would)\b", txt)),
        "“relevant / evidence / specific / practical”": len(re.findall(r"\b(relevant|evidence|specific|practical)\b", txt, re.I)),
    }


def main(ruta):
    a, b = prosa_original(), prosa_nueva(ruta)
    ma, mb = medidas(a), medidas(b)
    print("| Medida | Entrega | Versión nueva |\n|---|---|---|")
    for k in ma:
        print(f"| {k} | {ma[k]} | {mb[k]} |")
    marcadas = set()
    for p in P:
        for m in re.findall(r"\[\[(.*?)\]\]", p["text"]):
            marcadas |= tejas(m)
    nuevas = set()
    d = docx.Document(ruta)
    textos = [p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells]
    for t in textos:
        nuevas |= tejas(t)
    igual = len(marcadas & nuevas) / len(marcadas) * 100
    print(f"\nTexto marcado por Turnitin que sigue igual (secuencias de 5 palabras): {igual:.0f} %")


if __name__ == "__main__":
    main(sys.argv[1])
