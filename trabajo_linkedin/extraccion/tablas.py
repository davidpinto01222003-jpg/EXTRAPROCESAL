"""Extrae las tablas: palabras entre el título (en cursiva) y la nota, asignadas a columnas por x."""
import json, sys, itertools, re
W = json.load(open(sys.argv[1]))
lines = []
for (pg, top), ws in itertools.groupby(W, key=lambda w: (w["page"], round(w["top"]))):
    ws = sorted(ws, key=lambda w: w["x0"]); lines.append((pg, top, ws))
def tabla(nombre):
    # localizar "Table X" en negrita y su título; filas hasta "Note." o encabezado siguiente de prosa
    i = next(k for k, (pg, top, ws) in enumerate(lines) if " ".join(w["text"] for w in ws) == nombre and "Bold" in ws[0]["font"])
    filas, k = [], i + 2
    while k < len(lines):
        pg, top, ws = lines[k]
        t = " ".join(w["text"] for w in ws)
        if t.startswith("Note.") or (ws[0]["size"] >= 12 and ws[0]["x0"] < 110 and "Bold" in ws[0]["font"] and not t.startswith("Table")) or (ws[0]["size"] >= 12 and ws[0]["x0"] > 100 and ws[0]["x0"] < 112 and len(t) > 60):
            break
        if t.strip().isdigit() and ws[0]["x0"] > 700: k += 1; continue
        filas.append((pg, top, ws)); k += 1
    nota = " ".join(w["text"] for w in lines[k][2]) if k < len(lines) and " ".join(w["text"] for w in lines[k][2]).startswith("Note.") else ""
    return filas, nota
nombre = sys.argv[2]
filas, nota = tabla(nombre)
# columnas: x0 de las palabras de la primera línea de encabezado (negrita) agrupadas por huecos grandes
for pg, top, ws in filas:
    segs, cur = [], [ws[0]]
    for a, b in zip(ws, ws[1:]):
        if b["x0"] - a["x1"] > 8: segs.append(cur); cur = [b]
        else: cur.append(b)
    segs.append(cur)
    print(pg, int(top), " | ".join(f'@{int(s[0]["x0"])} ' + " ".join(w["text"] for w in s) for s in segs), "*" if ws[0]["hl"] else "")
print("NOTA:", nota)

def matriz(filas):
    def segmentos(ws):
        segs, cur = [], [ws[0]]
        for a, b in zip(ws, ws[1:]):
            if b["x0"] - a["x1"] > 8: segs.append(cur); cur = [b]
            else: cur.append(b)
        segs.append(cur); return segs
    datos = [(pg, top, ws) for pg, top, ws in filas if not "Bold" in ws[0]["font"]]
    starts = sorted({int(s[0]["x0"]) for pg, top, ws in datos for s in segmentos(ws)})
    cols = []
    for x in starts:
        if not cols or x - cols[-1] > 40: cols.append(x)
    head = [pg_ws for pg_ws in filas if "Bold" in pg_ws[2][0]["font"]]
    if not cols:
        cols = sorted({int(s[0]["x0"]) for pg, top, ws in head[:1] for s in segmentos(ws)})
    # encabezado: solo la primera aparición (se repite en cada página)
    hpage = head[0][0] if head else None
    hcells = [""] * len(cols)
    for pg, top, ws in head:
        if pg != hpage: continue
        for s in segmentos(ws):
            j = min(range(len(cols)), key=lambda c: abs(cols[c] - s[0]["x0"]) if s[0]["x0"] >= cols[c] - 60 else 999)
            hcells[j] = (hcells[j] + " " + " ".join(w["text"] for w in s)).strip()
    rows, prev = [], None
    for pg, top, ws in datos:
        if prev is None or pg != prev[0] or top - prev[1] > 14:
            rows.append([""] * len(cols))
        for s in segmentos(ws):
            j = max(c for c in range(len(cols)) if cols[c] <= s[0]["x0"] + 2)
            rows[-1][j] = (rows[-1][j] + " " + " ".join(w["text"] for w in s)).strip()
        prev = (pg, top)
    return hcells, rows

if __name__ == "__main__" and len(sys.argv) > 3:
    out = {}
    for n in sys.argv[3:]:
        f, nota = tabla(n)
        h, r = matriz(f)
        out[n] = dict(head=h, rows=r, note=nota)
    json.dump(out, open("tablas.json", "w"), ensure_ascii=False, indent=1)
