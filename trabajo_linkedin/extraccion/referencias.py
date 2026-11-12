import json, sys, itertools
W = json.load(open(sys.argv[1]))
W = [w for w in W if 33 <= w["page"] <= 38 and not (w["text"].isdigit() and w["x0"] > 700)]
lines = []
for (pg, top), ws in itertools.groupby(W, key=lambda w: (w["page"], round(w["top"]))):
    ws = sorted(ws, key=lambda w: w["x0"])
    if " ".join(w["text"] for w in ws) == "References": continue
    lines.append(ws)
entries = []
for ws in lines:
    if ws[0]["x0"] < 80: entries.append([])
    entries[-1].append(ws)
out = []
for e in entries:
    toks = []  # (texto, cursiva)
    for li, ws in enumerate(e):
        for wi, w in enumerate(ws):
            it = "Italic" in w["font"]
            t = w["text"]
            if toks and wi == 0:
                prev = toks[-1][0]
                url = "http" in prev or prev.endswith("/")
                if prev.endswith(("-", "—")) or (url and not prev.endswith(".")) or prev.endswith("/"):
                    toks[-1] = (prev + t, toks[-1][1]); continue
            if t == "." and toks:
                toks[-1] = (toks[-1][0] + ".", toks[-1][1]); continue
            toks.append((t, it))
    # agrupar en runs
    runs = []
    for t, it in toks:
        if runs and runs[-1][1] == it: runs[-1][0] += " " + t
        else: runs.append([(" " if runs else "") + t, it])
    out.append(runs)
json.dump(out, open(sys.argv[2], "w"), ensure_ascii=False, indent=0)
for r in out: print("".join(("_" + t + "_") if it else t for t, it in r))
