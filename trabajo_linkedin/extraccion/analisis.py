import json, re, sys, statistics as st
P = json.load(open(sys.argv[1]))
prosa = []
for i, p in enumerate(P):
    t = p["text"]
    if t.startswith("##") or p["x0"] < 100 or p["page"] < 4 or p["page"] >= 33 and p["page"] < 39: continue
    plano = t.replace("[[", "").replace("]]", "")
    if len(plano.split()) < 25: continue
    marc = sum(len(m.split()) for m in re.findall(r"\[\[(.*?)\]\]", t))
    prosa.append((i, plano, marc / len(plano.split())))
def oraciones(t): return [s for s in re.split(r"(?<=[.?!])\s+(?=[A-Z“])", t) if s]
CON = r"\b(because|since|while|whereas|although|so that|which is why|for that reason|in light of|on that basis|in practical terms|therefore|in turn|likewise|considering that|given that|bearing in mind|as well as|rather than|instead of)\b"
def rasgos(t):
    o = oraciones(t); w = t.split(); L = [len(s.split()) for s in o]
    return dict(pal=len(w), orac=len(o), largo=st.mean(L), maxlargo=max(L),
                cifras=len(re.findall(r"\d", t)) / len(w) * 100,
                citas=len(re.findall(r"\(\D*?\d{4}[a-z]?\)|n\.d\.", t)),
                yo=len(re.findall(r"\b(I|my|me)\b", t)) / len(w) * 100,
                iwill=len(re.findall(r"\bI (will|would)\b", t)),
                conect=len(re.findall(CON, t, re.I)) / len(o),
                comas=t.count(",") / len(o),
                which=len(re.findall(r", which\b", t)),
                and_=len(re.findall(r", and\b", t)) / len(o))
for grupo, cond in (("MARCADOS (>=80%)", lambda f: f >= .8), ("NO MARCADOS (<20%)", lambda f: f < .2), ("MIXTOS", lambda f: .2 <= f < .8)):
    sel = [rasgos(t) for i, t, f in prosa if cond(f)]
    if not sel: continue
    print(grupo, len(sel), "párrafos")
    for k in sel[0]:
        print(f"   {k:10s} {st.mean(s[k] for s in sel):7.2f}")
tot = sum(len(t.split()) for _, t, _ in prosa); m = sum(len(t.split()) * f for _, t, f in prosa)
print("prosa palabras", tot, "marcadas", round(m), round(m / tot * 100, 1), "%")
for i, t, f in prosa:
    if f < .2: print("NO", i, t[:110])
