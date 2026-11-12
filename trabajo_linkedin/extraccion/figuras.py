import sys, re, pdfplumber, pypdfium2 as pdfium
CIAN = "(0.3203125, 0.77734375, 0.85546875)"
src, outdir = sys.argv[1], sys.argv[2]
doc = pdfium.PdfDocument(src)
res = []
with pdfplumber.open(src) as pdf:
    for n, pg in enumerate(pdf.pages[2:], 3):
        words = pg.extract_words(extra_attrs=["fontname"])
        lines = {}
        for w in words:
            lines.setdefault(round(w["top"]), []).append(w)
        L = sorted((t, sorted(ws, key=lambda w: w["x0"])) for t, ws in lines.items())
        for k, (top, ws) in enumerate(L):
            txt = " ".join(w["text"] for w in ws)
            m = re.fullmatch(r"Figure (B?\d+)", txt)
            if not m or "Bold" not in ws[0]["fontname"]:
                continue
            titulo = L[k + 1]
            y0 = max(w["bottom"] for w in titulo[1]) + 3
            y1 = pg.height - 50
            for t2, ws2 in L[k + 2:]:
                t2txt = " ".join(w["text"] for w in ws2)
                if t2txt.startswith("Note.") or re.fullmatch(r"Figure B?\d+", t2txt) or ("Bold" in ws2[0]["fontname"] and t2 > y0 + 20) or t2 > pg.height - 55:
                    y1 = t2 - 3; break
            objs = [o for o in pg.images + pg.curves + pg.lines + pg.rects
                    if str(o.get("non_stroking_color")) != CIAN and o["top"] >= y0 - 2 and o["bottom"] <= y1 + 2]
            chars = [c for c in pg.chars if c["top"] >= y0 and c["bottom"] <= y1]
            xs0 = [o["x0"] for o in objs] + [c["x0"] for c in chars]
            xs1 = [o["x1"] for o in objs] + [c["x1"] for c in chars]
            ys0 = [o["top"] for o in objs] + [c["top"] for c in chars]
            ys1 = [o["bottom"] for o in objs] + [c["bottom"] for c in chars]
            if not xs0:
                print("SIN OBJETOS", m.group(1), n); continue
            box = (max(0, min(xs0) - 4), max(y0, min(ys0) - 4), min(pg.width, max(xs1) + 4), min(y1, max(ys1) + 4))
            sc = 300 / 72
            img = doc[n - 1].render(scale=sc).to_pil()
            crop = img.crop(tuple(int(v * sc) for v in box))
            nombre = f"{outdir}/figura_{m.group(1)}.png"
            crop.save(nombre)
            res.append((m.group(1), n, [round(b) for b in box], crop.size))
for r in res: print(r)
