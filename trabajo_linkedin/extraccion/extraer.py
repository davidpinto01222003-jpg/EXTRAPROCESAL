import sys, pdfplumber, collections, json
pdf = sys.argv[1]
cols = collections.Counter()
out = []
with pdfplumber.open(pdf) as doc:
    for n, pg in enumerate(doc.pages[2:], 3):
        rects = [r for r in pg.rects]
        for r in rects: cols[str(r.get("non_stroking_color"))] += 1
        cyan = [r for r in rects if str(r.get("non_stroking_color")) == "(0.3203125, 0.77734375, 0.85546875)"]
        words = [w for w in pg.extract_words(extra_attrs=["fontname","size"]) if 50 < w["top"] < pg.height - 50]
        for w in words:
            cx, cy = (w["x0"]+w["x1"])/2, (w["top"]+w["bottom"])/2
            w["hl"] = any(r["top"]-1 <= cy <= r["bottom"]+1 and r["x0"]-1 <= cx <= r["x1"]+1 for r in cyan)
            out.append(dict(page=n, top=round(w["top"],1), x0=round(w["x0"],1), x1=round(w["x1"],1), text=w["text"], hl=w["hl"], font=w["fontname"], size=round(w["size"],1)))
print(cols.most_common(8), file=sys.stderr)
json.dump(out, open(sys.argv[2], "w"), ensure_ascii=False)
