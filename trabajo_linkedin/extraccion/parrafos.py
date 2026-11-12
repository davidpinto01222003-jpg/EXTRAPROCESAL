import json, sys, itertools
W = json.load(open(sys.argv[1]))
lines = []
for (pg, top), ws in itertools.groupby(W, key=lambda w: (w["page"], round(w["top"]))):
    ws = sorted(ws, key=lambda w: w["x0"])
    lines.append(dict(page=pg, top=top, x0=ws[0]["x0"], x1=ws[-1]["x1"], ws=ws,
                      bold=all("Bold" in w["font"] for w in ws), size=ws[0]["size"]))
# merge lines with nearly same top (groupby on rounded top may split)
pars = []
cur = None
for l in lines:
    txt = " ".join(w["text"] for w in l["ws"])
    if txt.strip().isdigit() and l["x0"] > 700:  # page number
        continue
    nuevo = (cur is None or l["bold"] or cur["bold"] or l["x0"] > 74 or txt.startswith("Note.")
             or cur.get("x0", 72) > 74 and cur["x0"] < 100)
    if l["bold"] and cur is not None and cur["bold"] and abs(l["top"] - cur["last_top"]) < 2:
        nuevo = False
    if nuevo:
        cur = dict(page=l["page"], bold=l["bold"], words=[], last_x1=0, center=l["x0"] > 130 and l["x1"] < 700 and l["bold"], x0=l["x0"])
        pars.append(cur)
    cur["words"] += l["ws"]
    cur["last_x1"] = l["x1"]; cur["last_top"] = l["top"]; cur["page"] = l["page"]
out = []
for p in pars:
    s, hl = [], False
    for w in p["words"]:
        if w["hl"] and not hl: s.append("[["); hl = True
        if not w["hl"] and hl: s.append("]]"); hl = False
        s.append(w["text"])
    if hl: s.append("]]")
    t = " ".join(s).replace("[[ ", "[[").replace(" ]]", "]]")
    pref = "## " if p["bold"] else ""
    out.append(dict(page=p["page"], text=pref + t, x0=p["x0"]))
json.dump(out, open(sys.argv[2], "w"), ensure_ascii=False, indent=0)
with open(sys.argv[3], "w") as f:
    for i, p in enumerate(out):
        f.write(f"[{i}] p{p['page']} x{int(p['x0'])}: {p['text']}\n\n")
