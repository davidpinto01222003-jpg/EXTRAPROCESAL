"""Puntajes de detectores de IA de Hugging Face, por texto, con caché en disco.

Se usan como rasgos adicionales del simulador (no como veredicto directo):
- xlmr_es: CradeyMH/detector-ia-espanol (XLM-RoBERTa large afinado en español).
- autext: pandrei7/autextification-upb-mtl (AuTexTification 2023, inglés y español).
- binoculars: método sin entrenamiento de Hans et al. (2024) con
  Qwen2.5-0.5B (observador) y Qwen2.5-0.5B-Instruct (ejecutor). Más bajo = más IA;
  aquí se devuelve con el signo cambiado para que más alto = más IA.
- ppl: log-perplejidad del texto según Qwen2.5-0.5B (más baja = más predecible).

La primera ejecución descarga unos 4 GB de modelos y tarda; después todo sale
de `cache_hf.json`, así que reentrenar o evaluar la misma oración es inmediato.
"""
import hashlib
import json
from pathlib import Path

AQUI = Path(__file__).parent
CACHE = AQUI / "cache_hf.json"

XLMR_ES = "CradeyMH/detector-ia-espanol"
AUTEXT = "pandrei7/autextification-upb-mtl"
OBSERVADOR = "Qwen/Qwen2.5-0.5B"
EJECUTOR = "Qwen/Qwen2.5-0.5B-Instruct"

NOMBRES = ["xlmr_es", "autext", "binoculars", "ppl"]

_modelos = {}


def _clave(texto):
    return hashlib.sha1(texto.encode("utf8")).hexdigest()


def _cargar_cache():
    if CACHE.exists():
        return json.loads(CACHE.read_text(encoding="utf8"))
    return {}


def _torch():
    import torch
    torch.set_num_threads(max(1, torch.get_num_threads()))
    return torch


def _xlmr(textos):
    torch = _torch()
    from transformers import AutoModelForSequenceClassification, AutoTokenizer
    if "xlmr" not in _modelos:
        tok = AutoTokenizer.from_pretrained(XLMR_ES)
        mod = AutoModelForSequenceClassification.from_pretrained(XLMR_ES).eval()
        _modelos["xlmr"] = (tok, mod)
    tok, mod = _modelos["xlmr"]
    # etiqueta de IA: la que el config llame así; si no, la 1
    etiquetas = {v.lower(): k for k, v in mod.config.id2label.items()}
    ia = next((k for n, k in etiquetas.items() if n in ("ia", "ai", "generated", "machine", "label_1", "fake")), 1)
    out = []
    for i in range(0, len(textos), 8):
        enc = tok(textos[i:i + 8], padding=True, truncation=True, max_length=512, return_tensors="pt")
        with torch.no_grad():
            p = mod(**enc).logits.softmax(-1)[:, ia]
        out += p.tolist()
    return out


def _autext(textos):
    torch = _torch()
    from huggingface_hub import hf_hub_download
    from safetensors.torch import load_file
    from transformers import AutoConfig, AutoTokenizer
    from transformers.dynamic_module_utils import get_class_from_dynamic_module
    if "autext" not in _modelos:
        tok = AutoTokenizer.from_pretrained(AUTEXT)
        # El modelo carga su codificador dentro de __init__, cosa que
        # AutoModel.from_pretrained de transformers 5 no permite; se arma a mano.
        cfg = AutoConfig.from_pretrained(AUTEXT, trust_remote_code=True)
        cls = get_class_from_dynamic_module("model.AutextificationMTLModel", AUTEXT, trust_remote_code=True)
        mod = cls(cfg)
        mod.load_state_dict(load_file(hf_hub_download(AUTEXT, "model.safetensors")), strict=False)
        mod.eval()
        _modelos["autext"] = (tok, mod)
    tok, mod = _modelos["autext"]
    out = []
    for i in range(0, len(textos), 8):
        enc = tok(textos[i:i + 8], padding=True, truncation=True, max_length=512, return_tensors="pt")
        with torch.no_grad():
            p = mod(enc)["bot_prob"].reshape(-1)
        out += p.tolist()
    return out


def _binoculars(textos):
    torch = _torch()
    from transformers import AutoModelForCausalLM, AutoTokenizer
    if "bino" not in _modelos:
        tok = AutoTokenizer.from_pretrained(OBSERVADOR)
        obs = AutoModelForCausalLM.from_pretrained(OBSERVADOR, torch_dtype=torch.float32).eval()
        eje = AutoModelForCausalLM.from_pretrained(EJECUTOR, torch_dtype=torch.float32).eval()
        _modelos["bino"] = (tok, obs, eje)
    tok, obs, eje = _modelos["bino"]
    bino, ppl = [], []
    for t in textos:
        ids = tok(t, return_tensors="pt", truncation=True, max_length=512).input_ids
        if ids.shape[1] < 3:
            bino.append(0.0); ppl.append(0.0); continue
        with torch.no_grad():
            lo = obs(ids).logits[0, :-1].float()
            le = eje(ids).logits[0, :-1].float()
        y = ids[0, 1:]
        logp_o = lo.log_softmax(-1)
        nll = -logp_o.gather(-1, y[:, None]).squeeze(-1).mean()
        # entropía cruzada entre ejecutor y observador
        xent = -(le.softmax(-1) * logp_o).sum(-1).mean()
        bino.append(float(-(nll / xent)))
        ppl.append(float(nll))
    return bino, ppl


def puntajes(textos, verbose=True):
    """Devuelve una lista de dicts {nombre: puntaje} alineada con `textos`."""
    cache = _cargar_cache()
    pasos = [(["xlmr_es"], _xlmr), (["autext"], _autext), (["binoculars", "ppl"], _binoculars)]
    for nombres, fn in pasos:
        faltan = sorted({t for t in textos if nombres[0] not in cache.get(_clave(t), {})})
        if not faltan:
            continue
        if verbose:
            print(f"Detectores HF: {'/'.join(nombres)} en {len(faltan)} textos nuevos…", flush=True)
        r = fn(faltan)
        cols = r if len(nombres) > 1 else [r]
        for j, t in enumerate(faltan):
            d = cache.setdefault(_clave(t), {})
            for n, col in zip(nombres, cols):
                d[n] = col[j]
        CACHE.write_text(json.dumps(cache), encoding="utf8")
    return [cache[_clave(t)] for t in textos]
