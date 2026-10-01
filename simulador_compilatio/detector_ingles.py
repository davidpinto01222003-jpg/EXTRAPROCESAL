"""Detector de IA para textos en inglés con modelos publicados en Hugging Face.

El simulador de Compilatio (simulador.py) aprendió de informes sobre un trabajo
en español y sus rasgos no sirven para inglés. Para textos en inglés se usan,
párrafo por párrafo, tres clasificadores entrenados con texto humano y de IA:
- desklib/ai-text-detector-v1.01 (DeBERTa-v3-large; el más descargado),
- fakespot-ai/roberta-base-ai-text-detection-v1,
- Hello-SimpleAI/chatgpt-detector-roberta (HC3, 2023; detecta sobre todo ChatGPT).

El porcentaje de cada detector es la parte de las palabras del cuerpo que está en
párrafos con probabilidad de IA >= 0,5. No es Compilatio ni Turnitin.

Uso: python detector_ingles.py archivo.docx [otro.docx ...] --html informe.html
"""
import argparse
import html
import sys
from pathlib import Path

import torch
import torch.nn as nn
from transformers import (AutoConfig, AutoModel, AutoModelForSequenceClassification,
                          AutoTokenizer, PreTrainedModel)

sys.path.insert(0, str(Path(__file__).parent))
import simulador as S  # noqa: E402  (lectura de .docx)

DESKLIB = "desklib/ai-text-detector-v1.01"
CLASIFICADORES = ["fakespot-ai/roberta-base-ai-text-detection-v1", "Hello-SimpleAI/chatgpt-detector-roberta"]
NOMBRES = {DESKLIB: "Desklib", CLASIFICADORES[0]: "Fakespot", CLASIFICADORES[1]: "HC3-ChatGPT"}


class DesklibAIDetectionModel(PreTrainedModel):
    """Arquitectura publicada en la ficha del modelo de Desklib."""
    config_class = AutoConfig

    def __init__(self, config):
        super().__init__(config)
        self.model = AutoModel.from_config(config)
        self.classifier = nn.Linear(config.hidden_size, 1)

    def forward(self, input_ids, attention_mask=None):
        h = self.model(input_ids, attention_mask=attention_mask)[0]
        m = attention_mask.unsqueeze(-1).float()
        return self.classifier((h * m).sum(1) / m.sum(1).clamp(min=1e-9))


def cargar():
    mods = {}
    tok = AutoTokenizer.from_pretrained(DESKLIB)
    # transformers 5 no carga esta clase propia con from_pretrained: se arma a mano
    from huggingface_hub import hf_hub_download
    from safetensors.torch import load_file
    mod = DesklibAIDetectionModel(AutoConfig.from_pretrained(DESKLIB))
    faltan = mod.load_state_dict(load_file(hf_hub_download(DESKLIB, "model.safetensors")), strict=False)
    assert not faltan.missing_keys, faltan.missing_keys
    mods[DESKLIB] = (tok, mod.eval())
    for m in CLASIFICADORES:
        mods[m] = (AutoTokenizer.from_pretrained(m), AutoModelForSequenceClassification.from_pretrained(m).eval())
    return mods


def probabilidad(mods, nombre, texto):
    tok, mod = mods[nombre]
    enc = tok(texto, truncation=True, max_length=512, return_tensors="pt")
    with torch.no_grad():
        if nombre == DESKLIB:
            return torch.sigmoid(mod(enc["input_ids"], enc["attention_mask"])).item()
        return mod(**enc).logits.softmax(-1)[0, 1].item()  # etiqueta 1 = IA / ChatGPT


def cuerpo_ingles(pars):
    """Párrafos de prosa (>= 25 palabras) antes de la lista de obras citadas."""
    out = []
    for p in pars:
        if p.strip().lower() in ("works cited", "references", "bibliography"):
            break
        if len(p.split()) >= 25:
            out.append(p)
    return out


def analizar(mods, ruta):
    pars = cuerpo_ingles(S.leer_docx(ruta))
    filas = [(p, {n: probabilidad(mods, n, p) for n in mods}) for p in pars]
    total = sum(len(p.split()) for p in pars) or 1
    pct = {n: sum(len(p.split()) for p, pr in filas if pr[n] >= 0.5) / total * 100 for n in mods}
    prom = {n: sum(pr[n] * len(p.split()) for p, pr in filas) / total * 100 for n in mods}
    return {"archivo": Path(ruta).name, "palabras": total, "filas": filas, "pct": pct, "prom": prom}


def informe_html(resultados, salida):
    bloques = []
    for r in resultados:
        cab = "".join(f"<th>{NOMBRES[n]}</th>" for n in r["pct"])
        kpis = "".join(f'<div class="k">{NOMBRES[n]}<b>{r["pct"][n]:.0f} %</b><small>prob. media {r["prom"][n]:.0f} %</small></div>'
                       for n in r["pct"])
        filas = []
        for p, pr in r["filas"]:
            celdas = "".join(f'<td class="{"alto" if v >= .5 else ""}">{v:.0%}</td>' for v in pr.values())
            filas.append(f"<tr><td>{html.escape(p[:220])}{'…' if len(p) > 220 else ''}</td>{celdas}</tr>")
        bloques.append(f"<h2>{html.escape(r['archivo'])}</h2><p><small>{r['palabras']} palabras analizadas</small></p>"
                       f'<div class="kpis">{kpis}</div><table><tr><th>Párrafo</th>{cab}</tr>{"".join(filas)}</table>')
    Path(salida).write_text(PLANTILLA.format(cuerpo="".join(bloques)), encoding="utf8")


PLANTILLA = """<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Detector IA inglés</title><style>
:root{{--bg:#fff;--fg:#1d1d1f;--mut:#666;--alto:#ffd9d6;--card:#f5f6f8;--line:#ddd}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--bg:#16181b;--fg:#e8e8ea;--mut:#9a9aa0;--alto:#5a2420;--card:#212428;--line:#333}}}}
body{{background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,sans-serif;max-width:960px;margin:0 auto;padding:16px}}
.kpis{{display:flex;flex-wrap:wrap;gap:12px;margin:12px 0}}.k{{background:var(--card);border-radius:10px;padding:10px 14px;flex:1;min-width:140px}}
.k b{{display:block;font-size:26px}}small{{color:var(--mut)}}table{{border-collapse:collapse;width:100%;font-size:13px}}
td,th{{border-bottom:1px solid var(--line);padding:6px;vertical-align:top;text-align:left}}td.alto{{background:var(--alto)}}
</style></head><body><h1>Detección de IA en textos en inglés</h1>
<p><small>Tres detectores públicos de Hugging Face, párrafo por párrafo. El porcentaje es la parte del texto en párrafos con probabilidad de IA &ge; 50 %. Es un indicador; no reproduce Compilatio ni Turnitin.</small></p>
{cuerpo}</body></html>"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", nargs="+")
    ap.add_argument("--html")
    a = ap.parse_args()
    mods = cargar()
    res = [analizar(mods, d) for d in a.docx]
    for r in res:
        print(f"\n{r['archivo']} ({r['palabras']} palabras)")
        for n in r["pct"]:
            print(f"  {NOMBRES[n]:12s} {r['pct'][n]:5.1f} % del texto | prob. media {r['prom'][n]:5.1f} %")
    if a.html:
        informe_html(res, a.html)
        print(f"\nInforme: {a.html}")
