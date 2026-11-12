# Compara detectores de IA públicos contra las marcas reales de Turnitin en la entrega, y usa el
# que mejor las reproduce para evaluar una versión nueva.
# Uso:
#   python detectores.py calibrar                # AUC de cada detector frente a Turnitin
#   python detectores.py evaluar archivo.docx [detector]
import json
import re
import sys
from pathlib import Path

import numpy as np
import torch
import torch.nn as nn
from transformers import AutoConfig, AutoModel, AutoModelForSequenceClassification, AutoTokenizer, PreTrainedModel

AQUI = Path(__file__).parent
DETECTORES = {
    "desklib": "desklib/ai-text-detector-v1.01",
    "fakespot": "fakespot-ai/roberta-base-ai-text-detection-v1",
    "simpleai": "Hello-SimpleAI/chatgpt-detector-roberta",
    "openai": "openai-community/roberta-base-openai-detector",
}
_cache = {}


class DesklibAIDetectionModel(PreTrainedModel):
    """Cabeza de clasificación publicada en la ficha de desklib/ai-text-detector-v1.01."""
    config_class = AutoConfig

    def __init__(self, config):
        super().__init__(config)
        self.model = AutoModel.from_config(config)
        self.classifier = nn.Linear(config.hidden_size, 1)
        self.post_init()

    def forward(self, input_ids, attention_mask=None):
        out = self.model(input_ids, attention_mask=attention_mask)[0]
        m = attention_mask.unsqueeze(-1).expand(out.size()).float()
        pooled = (out * m).sum(1) / m.sum(1).clamp(min=1e-9)
        return self.classifier(pooled)


def cargar(clave):
    if clave not in _cache:
        nombre = DETECTORES[clave]
        tok = AutoTokenizer.from_pretrained(nombre)
        if clave == "desklib":
            m = DesklibAIDetectionModel.from_pretrained(nombre)
        else:
            m = AutoModelForSequenceClassification.from_pretrained(nombre)
        m.eval()
        _cache[clave] = (tok, m)
    return _cache[clave]


@torch.no_grad()
def prob_ia(clave, textos):
    tok, m = cargar(clave)
    out = []
    for t in textos:
        enc = tok(t, return_tensors="pt", truncation=True, max_length=512)
        if clave == "desklib":
            out.append(torch.sigmoid(m(enc["input_ids"], enc["attention_mask"])).item())
            continue
        logits = m(**enc).logits[0]
        p = torch.softmax(logits, -1)
        etiquetas = {v.lower(): k for k, v in m.config.id2label.items()}
        ia = next((etiquetas[k] for k in ("ai", "chatgpt", "fake", "label_1", "ai-generated", "machine") if k in etiquetas), 1)
        out.append(p[ia].item())
    return out


def parrafos_entrega():
    P = json.load(open(AQUI / "datos/parrafos.json"))
    out = []
    for p in P:
        t = p["text"]
        if t.startswith("##") or p["x0"] < 100 or p["page"] < 4 or 33 <= p["page"] <= 38:
            continue
        plano = t.replace("[[", "").replace("]]", "")
        if len(plano.split()) < 20:
            continue
        marc = sum(len(m.split()) for m in re.findall(r"\[\[(.*?)\]\]", t))
        out.append((plano, marc / len(plano.split())))
    return out


def calibrar():
    from sklearn.metrics import roc_auc_score
    pars = parrafos_entrega()
    y = np.array([1 if f >= 0.5 else 0 for _, f in pars])
    w = np.array([len(t.split()) for t, _ in pars])
    print(f"{len(pars)} párrafos, {y.mean():.0%} marcados")
    res = {}
    for clave in DETECTORES:
        try:
            p = np.array(prob_ia(clave, [t for t, _ in pars]))
        except Exception as e:  # modelo no disponible
            print(f"  {clave}: no disponible ({e.__class__.__name__}: {str(e)[:80]})")
            continue
        auc = roc_auc_score(y, p)
        # umbral que reproduce el 71 % ponderado por palabras
        objetivo = (w * y).sum() / w.sum()
        umbral = min(np.linspace(0.01, 0.99, 197), key=lambda u: abs((w * (p >= u)).sum() / w.sum() - objetivo))
        acierto = ((p >= umbral) == y).mean()
        estimado_05 = (w * (p >= 0.5)).sum() / w.sum()
        res[clave] = dict(auc=auc, umbral=float(umbral))
        print(f"  {clave:9s} AUC {auc:.2f} | con umbral 0,5 estima {estimado_05:.0%} | umbral calibrado {umbral:.2f}, acierto {acierto:.0%}")
    json.dump(res, open(AQUI / "datos/detectores.json", "w"), indent=1)


def prosa_docx(ruta):
    import docx
    d = docx.Document(ruta)
    out, fuera = [], False
    for p in d.paragraphs:
        t = p.text.strip()
        if t == "References":
            fuera = True
        elif t.startswith("Appendix"):
            fuera = False
        if not fuera and len(t.split()) >= 20 and not t.startswith("Note."):
            out.append(t)
    return out


def evaluar(ruta, claves=None):
    cal = json.load(open(AQUI / "datos/detectores.json"))
    pars = prosa_docx(ruta)
    w = np.array([len(t.split()) for t in pars])
    claves = claves or list(cal)
    todos = {}
    for clave in claves:
        p = np.array(prob_ia(clave, pars))
        todos[clave] = p
        print(f"{clave:9s} umbral calibrado: {(w * (p >= cal[clave]['umbral'])).sum() / w.sum():.0%} | "
              f"umbral 0,5: {(w * (p >= 0.5)).sum() / w.sum():.0%} | prob. media {np.average(p, weights=w):.2f}")
    return pars, todos


if __name__ == "__main__":
    if sys.argv[1] == "calibrar":
        calibrar()
    else:
        pars, todos = evaluar(sys.argv[2], sys.argv[3:] or None)
        clave = (sys.argv[3:] or list(todos))[0]
        for t, v in sorted(zip(pars, todos[clave]), key=lambda x: -x[1])[:12]:
            print(f"  {v:.2f}  {t[:120]}")
