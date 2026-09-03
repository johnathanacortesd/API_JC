"""
Core pipeline helpers for API_JC.

This module is intentionally Streamlit-free so the working product can be
repaired and regression-tested without rewriting the UI, PKL branches or
dossier controls. app.py remains the product surface.
"""
from __future__ import annotations

import datetime
import hashlib
import io
import json
import re
import threading
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from unidecode import unidecode

try:
    from sklearn.metrics.pairwise import cosine_similarity
except Exception:  # pragma: no cover
    cosine_similarity = None


_COUNTER_LOCK = threading.Lock()

OPENAI_MODEL_EMBEDDING = "text-embedding-3-small"
OPENAI_MODEL_CLASIFICACION = "gpt-4.1-nano-2025-04-14"
CHAT_BATCH_SIZE = 30
CHAT_PARALLEL_BATCHES = 4
REQUEST_TIMEOUT_S = 25
MAX_RETRIES = 2
MAX_BUCKET = 80
MAX_PAIRS_TOTAL = 40000
MAX_CMP_CHARS = 400
SIMILARITY_THRESHOLD_TITULOS = 0.92
SIMILARITY_THRESHOLD_TITULOS_BCAST = 0.86
SIMILARITY_THRESHOLD_RESUMEN = 0.86
SIMILARITY_THRESHOLD_SEMANTIC = 0.88
MIN_OVERLAP_GRUPO = 0.30
MAX_PALABRAS_SUBTEMA = 7
MIN_PALABRAS_SUBTEMA = 3

OUTPUT_COLUMNS = [
    "ID Noticia",
    "Fecha",
    "Hora",
    "Medio",
    "Tipo de Medio",
    "Sección - Programa",
    "Región",
    "Título",
    "Tono IA",
    "Tema",
    "Subtema",
    "Link Nota",
    "Resumen - Aclaracion",
    "Link (Streaming - Imagen)",
    "Menciones - Empresa",
    "ID duplicada",
    "Cuerpo Completo",
    "Contexto analizado",
    "Coincidencia marca",
    "Origen coincidencia",
    "Tono",
    "Grupo noticia",
]

KEYMAP = {
    "idnoticia": "ID Noticia",
    "fecha": "Fecha",
    "hora": "Hora",
    "medio": "Medio",
    "tipodemedio": "Tipo de Medio",
    "seccion_programa": "Sección - Programa",
    "region": "Región",
    "titulo": "Título",
    "tonoiai": "Tono IA",
    "tema": "Tema",
    "subtema": "Subtema",
    "link_nota": "Link Nota",
    "resumen": "Resumen - Aclaracion",
    "link_streaming": "Link (Streaming - Imagen)",
    "menciones": "Menciones - Empresa",
    "idduplicada": "ID duplicada",
    "cuerpo": "Cuerpo Completo",
    "contexto": "Contexto analizado",
    "coincidencia": "Coincidencia marca",
    "origen": "Origen coincidencia",
    "tono": "Tono",
    "grupo": "Grupo noticia",
}

TRACKING_PARAMS = {
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "utm_id", "utm_cid", "utm_reader", "utm_name", "utm_social",
    "fbclid", "gclid", "gclsrc", "dclid", "msclkid", "twclid", "yclid",
    "mc_cid", "mc_eid", "igshid", "_ga", "_gl", "ref", "ref_src",
    "feature", "ncid", "cmpid", "s",
}

STOPWORDS_ES = set("""
a ante bajo cabe con contra de desde durante en entre hacia hasta mediante
para por segun sin so sobre tras y o u e la el los las un una unos unas lo
al del se su sus le les mi mis tu tus nuestro nuestros vuestra vuestras este
esta estos estas ese esa esos esas aquel aquella aquellos aquellas que cual
cuales quien quienes cuyo cuya cuyos cuyas como cuando donde cual es son fue
fueron era eran sera seran seria serian he ha han habia hay hubo habra
habria estoy estan estaba estaban estamos estar estare estaria
estuvieron estarian estuvo asi ya mas menos tan tanto cada muy todo toda todos
todas ser haber hacer tener poder deber ir dar ver saber querer llegar pasar
encontrar creer decir poner salir volver seguir llevar sentir cambiar
""".split())

_TRAILING_INCOMPLETE = {
    "de", "del", "la", "el", "los", "las", "un", "una", "unos", "unas", "al",
    "su", "sus", "en", "con", "sin", "por", "para", "sobre", "ante", "bajo",
    "contra", "desde", "entre", "hacia", "hasta", "mediante", "tras", "y", "o",
    "u", "e", "lo", "que", "se", "como", "donde", "cuando", "cual", "cuyo",
    "si", "cómo", "como", "qué", "cuál",
}

_CHANNEL_PREFIXES = {
    "video", "en vivo", "vivo", "streaming", "audio", "podcast", "radio",
    "tv", "television", "televisión",
}

_QUESTION_TAILS = {
    "como", "cómo", "si", "que", "qué", "cual", "cuál", "cuando", "cuándo",
    "donde", "dónde", "por", "para",
}

_CONJUGATED_TAILS = {
    "ascienden", "asciende", "saben", "sabe", "destacan", "destaca",
    "confirman", "confirma", "anuncian", "anuncia", "presentan", "presenta",
    "lanzan", "lanza", "reportan", "reporta", "informan", "informa",
    "advierten", "advierte", "revelan", "revela", "ocurren", "ocurre",
    "registran", "registra", "mueren", "muere", "resultan", "resulta",
    "quedan", "queda", "llegan", "llega", "inician", "inicia",
}

_NEXOS = {
    "de", "del", "para", "sobre", "en", "con", "por", "ante", "hacia",
    "entre", "sin", "al", "las", "los",
}

_ACCIONES_OPUESTAS = [
    ({"aprobacion", "aprueba", "apoyo", "acuerdo", "aval", "respaldo"},
     {"rechazo", "rechaza", "desacuerdo", "oposicion", "critica"}),
    ({"aumento", "crecimiento", "alza", "subida", "incremento"},
     {"caida", "reduccion", "baja", "disminucion", "descenso"}),
    ({"apertura", "inauguracion", "inicio", "lanzamiento", "estreno"},
     {"cierre", "suspension", "fin", "clausura", "cancelacion"}),
    ({"exito", "logro", "triunfo", "premio", "reconocimiento"},
     {"fracaso", "derrota", "problema", "crisis", "sancion"}),
    ({"demanda", "denuncia", "investigacion", "sancion", "multa"},
     {"absolucion", "archivo", "exoneracion"}),
]

_TOKENS_DEBILES = STOPWORDS_ES | {
    "noticia", "noticias", "informe", "informacion", "comunicado", "anuncio",
    "colombia", "pais", "nacional", "regional", "local", "sector",
    "empresa", "empresas", "nuevo", "nueva", "plan", "programa", "proyecto",
    "actividad", "gestion", "tema", "caso",
}


# ---------------------------------------------------------------------------
# Instrumentation
# ---------------------------------------------------------------------------

class ComparisonCounter:
    """Counts blocked pairwise comparisons (must stay << n²)."""

    def __init__(self) -> None:
        self.n = 0

    def add(self, k: int = 1) -> None:
        self.n += int(k)

    def reset(self) -> None:
        self.n = 0


class CallCounter:
    def __init__(self) -> None:
        self.chat = 0
        self.embed = 0
        self.embed_items = 0
        self.chat_items = 0

    def reset(self) -> None:
        self.chat = self.embed = self.embed_items = self.chat_items = 0


class ProgressTracker:
    STAGES = (
        "Limpieza", "Duplicados", "Contexto", "Embedding único",
        "Agrupación", "Tono", "Tema/Subtema", "Excel",
    )

    def __init__(self, on_stage: Optional[Callable[[Dict[str, Any]], None]] = None) -> None:
        self.t0 = time.time()
        self.events: List[Dict[str, Any]] = []
        self.calls = CallCounter()
        self.comparisons = ComparisonCounter()
        self._current = None
        self.on_stage = on_stage

    def stage(self, name: str, extra: str = "") -> Dict[str, Any]:
        elapsed = time.time() - self.t0
        ev = {
            "stage": name,
            "elapsed_s": round(elapsed, 3),
            "chat_calls": self.calls.chat,
            "embed_calls": self.calls.embed,
            "comparisons": self.comparisons.n,
            "extra": extra,
            "label": (
                f"{name} · {elapsed:.1f}s · "
                f"chat={self.calls.chat} · emb={self.calls.embed}"
                + (f" · {extra}" if extra else "")
            ),
        }
        ev["index"] = len(self.events) + 1
        ev["total"] = len(self.STAGES)
        self.events.append(ev)
        self._current = ev
        if self.on_stage is not None:
            try:
                self.on_stage(ev)
            except Exception:
                pass
        return ev

    def summary(self) -> str:
        return " | ".join(e["label"] for e in self.events)


# ---------------------------------------------------------------------------
# BUSCARV (VLOOKUP) — do not rename or drop this branch
# ---------------------------------------------------------------------------

def construir_mapa_buscarv(pares: Iterable) -> Dict[str, str]:
    """Excel BUSCARV / VLOOKUP: primera columna = clave, segunda = valor."""
    mapping: Dict[str, str] = {}
    if pares is None:
        return mapping
    if hasattr(pares, "iloc"):
        df = pares.dropna(how="all")
        keys = df.iloc[:, 0].astype(str).str.lower().str.strip()
        vals = df.iloc[:, 1].astype(str)
        for k, v in zip(keys, vals):
            if k and k != "nan":
                mapping[k] = v
        return mapping
    for item in pares:
        if isinstance(item, dict):
            k = str(item.get("clave", item.get("key", ""))).lower().strip()
            v = item.get("valor", item.get("value", ""))
        else:
            if len(item) < 2:
                continue
            k = str(item[0]).lower().strip()
            v = item[1]
        if k and k != "nan":
            mapping[k] = "" if v is None else str(v)
    return mapping


def aplicar_buscarv(clave: Any, mapping: Dict[str, str], si_no_encuentra: Any = "N/A") -> Any:
    """Exact API_JC lookup: lower+strip key, return mapped value or default."""
    if clave is None:
        return si_no_encuentra
    k = str(clave).lower().strip()
    if not k or k == "nan":
        return si_no_encuentra
    return mapping.get(k, si_no_encuentra)


def aplicar_buscarv_dossier(medios: Sequence[Any], region_map: Dict[str, str],
                            internet_map: Dict[str, str],
                            tipos: Optional[Sequence[Any]] = None
                            ) -> Tuple[List[Any], List[Any]]:
    """
    Same two BUSCARV branches the dossier already uses:
    - Región = BUSCARV(Medio, sheet Regiones)
    - Medio (solo Internet) = BUSCARV(Medio, sheet Internet)
    """
    regiones = [aplicar_buscarv(m, region_map, "N/A") for m in medios]
    medios_out = list(medios)
    if tipos is not None:
        for i, (medio, tipo) in enumerate(zip(medios, tipos)):
            if normalizar_tipo_medio(tipo) == "Internet":
                mapped = aplicar_buscarv(medio, internet_map, None)
                if mapped is not None:
                    medios_out[i] = mapped
    return regiones, medios_out


# ---------------------------------------------------------------------------
# Titles, URLs, hyperlinks
# ---------------------------------------------------------------------------

def titulo_original(valor: Any) -> Any:
    """Output Título must stay the original cell; never split on | : - HTML."""
    if isinstance(valor, dict):
        return valor.get("value", valor)
    return valor


def normalize_title_for_comparison(title: Any) -> str:
    if not isinstance(title, str):
        title = "" if title is None else str(title)
    cleaned = unidecode(title)
    return re.sub(r"\W+", " ", cleaned).lower().strip()


def extraer_hipervinculo_celda(cell: Any) -> Optional[str]:
    """Read the embedded Excel hyperlink target (pandas drops this)."""
    if cell is None:
        return None
    hyper = getattr(cell, "hyperlink", None)
    if hyper is not None:
        target = getattr(hyper, "target", None) or getattr(hyper, "location", None)
        if target:
            return str(target)
    value = getattr(cell, "value", cell)
    if isinstance(value, str) and "HYPERLINK" in value.upper():
        m = re.search(r'HYPERLINK\(\s*"([^"]+)"', value, flags=re.I)
        if m:
            return m.group(1)
    return None


def valor_con_hipervinculo(cell: Any) -> Any:
    url = extraer_hipervinculo_celda(cell)
    value = getattr(cell, "value", cell)
    if url:
        return {"value": value if value not in (None, "") else "Link", "url": url}
    return value


def url_de_celda_link(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, dict):
        return str(val.get("url") or "")
    if isinstance(val, str) and val.startswith("http"):
        return val
    return ""


def normalize_url(url: Any) -> str:
    """Strip tracking/fragment/trailing slash and lowercase host; keep path."""
    if not url:
        return ""
    raw = str(url).strip()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "http://" + raw
    try:
        p = urlparse(raw)
    except Exception:
        return raw.strip().lower().rstrip("/")
    host = (p.hostname or "").lower()
    if host.startswith("www."):
        host = host[4:]
    path = p.path or ""
    if path != "/" and path.endswith("/"):
        path = path.rstrip("/")
    q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True)
         if k.lower() not in TRACKING_PARAMS]
    query = urlencode(q, doseq=True)
    rebuilt = urlunparse((p.scheme.lower() or "http", host, path, "", query, ""))
    rebuilt = re.sub(r"^https?://", "", rebuilt)
    return rebuilt.rstrip("/")


def normalizar_tipo_medio(tipo_raw: Any) -> str:
    if not isinstance(tipo_raw, str):
        tipo_raw = str(tipo_raw or "")
    t = unidecode(tipo_raw.strip().lower())
    return {
        "online": "Internet", "internet": "Internet",
        "diario": "Prensa", "prensa": "Prensa",
        "am": "Radio", "fm": "Radio", "radio": "Radio",
        "aire": "Televisión", "cable": "Televisión", "tv": "Televisión",
        "television": "Televisión", "televisión": "Televisión",
        "revista": "Revistas", "revistas": "Revistas",
    }.get(t, tipo_raw.strip().title() or "Otro")


def normalizar_fecha(val: Any) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    if hasattr(val, "strftime"):
        try:
            return val.strftime("%Y-%m-%d")
        except Exception:
            pass
    s = str(val).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except Exception:
            continue
    try:
        import pandas as pd
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if ts is not None and not (hasattr(ts, "isna") and ts.isna()):
            return ts.strftime("%Y-%m-%d")
    except Exception:
        pass
    return s[:10]


def normalizar_hora(val: Any) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    if hasattr(val, "strftime"):
        try:
            return val.strftime("%H:%M")
        except Exception:
            pass
    s = str(val).strip()
    if not s or s.lower() in {"nan", "none"}:
        return ""
    m = re.search(r"(\d{1,2})[:.](\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def _norm_text(texto: Any) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", unidecode(str(texto or "").lower()))).strip()


def _tokens_distintivos(texto: str, min_len: int = 4) -> set:
    return {
        t for t in _norm_text(texto).split()
        if len(t) >= min_len and t not in _TOKENS_DEBILES and not t.isdigit()
    }


def _overlap_distintivo(a: str, b: str) -> float:
    ta, tb = _tokens_distintivos(a), _tokens_distintivos(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(1, min(len(ta), len(tb)))


def _hay_conflicto_accion(a: str, b: str) -> bool:
    ta, tb = _tokens_distintivos(a, min_len=3), _tokens_distintivos(b, min_len=3)
    for ga, gb in _ACCIONES_OPUESTAS:
        if (ta & ga and tb & gb) or (ta & gb and tb & ga):
            return True
    return False


def _ratio(a: str, b: str, threshold: float = 0.0) -> float:
    """Character SequenceMatcher for SHORT strings (titles), capped in length."""
    from difflib import SequenceMatcher
    if not a or not b:
        return 0.0
    a, b = a[:MAX_CMP_CHARS], b[:MAX_CMP_CHARS]
    sm = SequenceMatcher(None, a, b)
    if threshold > 0 and sm.real_quick_ratio() < threshold:
        return 0.0
    return sm.ratio()


MAX_CMP_WORDS = 80


def _ratio_palabras(a_words: Sequence[str], b_words: Sequence[str], threshold: float = 0.0) -> float:
    """Word-level similarity for resúmenes.

    Character-level ratio() is O(len_a*len_b) and froze the UI on 1.5k-char
    summaries. Word sequences (≤80 tokens) give the same "same text" signal
    at ~1/100 of the cost. A set-overlap gate skips hopeless pairs first.
    """
    from difflib import SequenceMatcher
    if not a_words or not b_words:
        return 0.0
    a_words, b_words = a_words[:MAX_CMP_WORDS], b_words[:MAX_CMP_WORDS]
    if threshold > 0:
        sa, sb = set(a_words), set(b_words)
        upper = 2.0 * len(sa & sb) / max(1, len(a_words) + len(b_words))
        # upper bound of ratio() when every shared token could align
        if upper < threshold * 0.5:
            return 0.0
    return SequenceMatcher(None, a_words, b_words, autojunk=False).ratio()


# ---------------------------------------------------------------------------
# Duplicates (NOT the same as Grupo noticia)
# ---------------------------------------------------------------------------

def _mencion_key(row: Dict[str, Any], km: Dict[str, str]) -> str:
    return _norm_text(row.get(km.get("menciones", "Menciones - Empresa"), ""))


def detectar_duplicados(rows: Sequence[Dict[str, Any]],
                        km: Optional[Dict[str, str]] = None) -> List[Dict[str, Any]]:
    """
    Internet/prensa: duplicate only if the embedded Link Nota URL matches
    (normalized). Same title + different URL/media is NOT a duplicate.

    Radio/TV: similar title AND same Fecha AND same Hora.
    Medio is NOT required — the same syndicated broadcast can air on
    several outlets; requiring Medio would miss those (existing code
    keyed on medio+hora+mencion and ignored title/fecha).
    Mención is kept so the dossier expansion (one row per company) does
    not mark two companies on the same note as duplicates of each other.
    """
    km = km or KEYMAP
    processed = deepcopy(list(rows))
    seen_url: Dict[Tuple[str, str], int] = {}
    seen_bcast: Dict[Tuple[str, str, str, str], int] = {}
    # (mencion, fecha, hora) -> [(titulo_norm, idx)] so similar-title checks
    # only scan the same broadcast slot, never the whole dossier.
    bcast_slot: Dict[Tuple[str, str, str], List[Tuple[str, int]]] = defaultdict(list)

    for i, row in enumerate(processed):
        row.setdefault("is_duplicate", False)
        if row.get("is_duplicate"):
            continue
        tipo = normalizar_tipo_medio(row.get(km.get("tipodemedio", "Tipo de Medio"), ""))
        mencion = _mencion_key(row, km)
        if tipo in ("Internet", "Prensa", "Revistas"):
            url = normalize_url(url_de_celda_link(row.get(km.get("link_nota", "Link Nota"))))
            if not url:
                continue
            key = (url, mencion)
            if key in seen_url:
                orig = processed[seen_url[key]]
                row["is_duplicate"] = True
                row[km.get("idduplicada", "ID duplicada")] = orig.get(km.get("idnoticia", "ID Noticia"), "")
                row["Tono IA"] = "Duplicada"
                row["Tema"] = "-"
                row["Subtema"] = "-"
            else:
                seen_url[key] = i
        elif tipo in ("Radio", "Televisión"):
            titulo = normalize_title_for_comparison(row.get(km.get("titulo", "Título"), ""))
            fecha = normalizar_fecha(row.get(km.get("fecha", "Fecha"), ""))
            hora = normalizar_hora(row.get(km.get("hora", "Hora"), ""))
            if not (titulo and fecha and hora):
                continue
            exact = (mencion, fecha, hora, titulo)
            if exact in seen_bcast:
                orig = processed[seen_bcast[exact]]
                row["is_duplicate"] = True
                row[km.get("idduplicada", "ID duplicada")] = orig.get(km.get("idnoticia", "ID Noticia"), "")
                row["Tono IA"] = "Duplicada"
                row["Tema"] = "-"
                row["Subtema"] = "-"
                continue
            matched = None
            for p_tit, prev_i in bcast_slot.get((mencion, fecha, hora), []):
                contained = (
                    len(titulo) >= 20 and len(p_tit) >= 20
                    and (titulo in p_tit or p_tit in titulo)
                )
                if contained or _ratio(titulo, p_tit, SIMILARITY_THRESHOLD_TITULOS_BCAST) >= SIMILARITY_THRESHOLD_TITULOS_BCAST:
                    matched = prev_i
                    break
            if matched is not None:
                orig = processed[matched]
                row["is_duplicate"] = True
                row[km.get("idduplicada", "ID duplicada")] = orig.get(km.get("idnoticia", "ID Noticia"), "")
                row["Tono IA"] = "Duplicada"
                row["Tema"] = "-"
                row["Subtema"] = "-"
            else:
                seen_bcast[exact] = i
                bcast_slot[(mencion, fecha, hora)].append((titulo, i))
    return processed


# ---------------------------------------------------------------------------
# Contexto analizado
# ---------------------------------------------------------------------------

def _lista_alias(marca: str, aliases=None) -> List[str]:
    nombres: List[str] = []
    if marca:
        nombres.extend(str(marca).split(";"))
    if isinstance(aliases, str):
        nombres.extend(aliases.split(";"))
    else:
        nombres.extend(str(a) for a in (aliases or []))
    out, seen = [], set()
    for n in nombres:
        k = _norm_text(n)
        if k and k not in seen:
            seen.add(k)
            out.append(n.strip())
    return out


def _menciona(texto: str, marca: str, aliases=None) -> bool:
    norm = _norm_text(texto)
    if not norm:
        return False
    for nombre in _lista_alias(marca, aliases):
        kn = _norm_text(nombre)
        if kn and re.search(rf"(?<![a-z0-9]){re.escape(kn)}(?![a-z0-9])", norm):
            return True
        toks = [t for t in kn.split() if len(t) >= 3 and t not in {"de", "del", "la", "el", "los", "las", "y"}]
        if len(toks) >= 2 and sum(t in norm.split() for t in toks) >= max(2, int(np.ceil(len(set(toks)) * 0.6))):
            return True
    return False


def _ventanas_mencion(texto: str, marca: str, aliases=None, ventana: int = 220) -> List[str]:
    if not texto:
        return []
    norm = _norm_text(texto)
    hits = []
    for nombre in _lista_alias(marca, aliases):
        kn = _norm_text(nombre)
        if not kn:
            continue
        for m in re.finditer(rf"(?<![a-z0-9]){re.escape(kn)}(?![a-z0-9])", norm):
            # Approximate character window back onto original text.
            ratio = (m.start() / max(len(norm), 1))
            center = int(ratio * len(texto))
            lo, hi = max(0, center - ventana), min(len(texto), center + ventana)
            fragment = texto[lo:hi].strip()
            if fragment:
                hits.append(fragment)
    if hits:
        return hits
    partes = re.split(r"(?<=[\.\!\?\n])\s+", texto)
    return [p.strip() for p in partes if p.strip() and _menciona(p, marca, aliases)]


def extraer_contexto_analizado(titulo: Any, resumen: Any, marca: str,
                               aliases=None, cuerpo: Any = "") -> Dict[str, str]:
    """
    Coherent Colombian-Spanish paragraph from brand mention windows.
    Título + Resumen first; Cuerpo Completo only if needed.
    Fallback: Resumen, then full Título. Never mutates the original title cell.
    """
    tit = "" if titulo is None else str(titulo)
    res = "" if resumen is None else str(resumen)
    cue = "" if cuerpo is None else str(cuerpo)
    nombres = _lista_alias(marca, aliases)
    title_hit = _menciona(tit, marca, aliases)
    res_hit = _menciona(res, marca, aliases)
    cue_hit = _menciona(cue, marca, aliases) if cue else False

    coincidencia = ""
    source = f"{tit} {res} {cue}"
    source_norm = _norm_text(source)
    for n in nombres:
        if _norm_text(n) and re.search(rf"(?<![a-z0-9]){re.escape(_norm_text(n))}(?![a-z0-9])", source_norm):
            coincidencia = n
            break
    if not coincidencia and (title_hit or res_hit or cue_hit):
        coincidencia = marca or ""

    origen_parts = []
    if title_hit:
        origen_parts.append("Título")
    if res_hit:
        origen_parts.append("Resumen")
    if cue_hit and not (title_hit or res_hit):
        origen_parts.append("Cuerpo Completo")
    origen = ", ".join(origen_parts)

    bloques: List[str] = []
    if title_hit or res_hit:
        bloques.extend(_ventanas_mencion(tit, marca, aliases))
        bloques.extend(_ventanas_mencion(res, marca, aliases))
        if not bloques:
            if title_hit:
                bloques.append(tit.strip())
            if res_hit:
                bloques.append(res.strip())
    elif cue_hit:
        bloques.extend(_ventanas_mencion(cue, marca, aliases)[:3])
    else:
        if res.strip():
            bloques.append(res.strip())
            origen = origen or "Resumen"
        elif tit.strip():
            bloques.append(tit.strip())
            origen = origen or "Título"

    vistos, out = set(), []
    for b in bloques:
        k = _norm_text(b)
        if k and k not in vistos:
            vistos.add(k)
            out.append(re.sub(r"\s+", " ", b).strip())
    contexto = " ".join(out).strip()
    if contexto and contexto[-1] not in ".!?…":
        contexto = contexto.rstrip(" ,;") + "."
    return {
        "contexto": contexto[:1800],
        "coincidencia": coincidencia,
        "origen": origen,
    }


# ---------------------------------------------------------------------------
# Blocked grouping (Grupo noticia) — not n²
# ---------------------------------------------------------------------------

class DSU:
    def __init__(self, n: int) -> None:
        self.p = list(range(n))
        self.rank = [0] * n

    def find(self, i: int) -> int:
        while self.p[i] != i:
            self.p[i] = self.p[self.p[i]]
            i = self.p[i]
        return i

    def union(self, i: int, j: int) -> None:
        ri, rj = self.find(i), self.find(j)
        if ri == rj:
            return
        if self.rank[ri] < self.rank[rj]:
            ri, rj = rj, ri
        self.p[rj] = ri
        if self.rank[ri] == self.rank[rj]:
            self.rank[ri] += 1

    def grupos(self, n: int) -> Dict[int, List[int]]:
        c: Dict[int, List[int]] = defaultdict(list)
        for i in range(n):
            c[self.find(i)].append(i)
        return dict(c)


def _prefix(text: str, words: int) -> str:
    toks = _norm_text(text).split()
    return " ".join(toks[:words])


def _candidate_pairs(n: int, buckets: Dict[str, List[int]],
                     counter: Optional[ComparisonCounter] = None
                     ) -> List[Tuple[int, int]]:
    pares = set()
    # Smallest buckets first: the most specific evidence wins the pair budget.
    for idxs in sorted(buckets.values(), key=len):
        if len(idxs) < 2 or len(idxs) > MAX_BUCKET:
            continue
        orden = sorted(set(idxs))
        for a in range(len(orden)):
            for b in range(a + 1, len(orden)):
                pares.add((orden[a], orden[b]))
        if len(pares) >= MAX_PAIRS_TOTAL:
            break
    if counter:
        counter.add(len(pares))
    return sorted(pares)


def agrupar_noticias_bloqueado(
    titulos: Sequence[Any],
    resumenes: Sequence[Any],
    contextos: Optional[Sequence[Any]] = None,
    embeddings: Optional[Sequence[Any]] = None,
    urls: Optional[Sequence[Any]] = None,
    fechas: Optional[Sequence[Any]] = None,
    horas: Optional[Sequence[Any]] = None,
    counter: Optional[ComparisonCounter] = None,
) -> Dict[int, List[int]]:
    """
    Same-fact grouping by similar title OR similar resumen OR high semantic
    similarity on Contexto analizado. Blocked by title prefix, resumen
    prefix, distinctive tokens, normalized URL and date+time.
    """
    n = len(titulos)
    dsu = DSU(n)
    tit_n = [normalize_title_for_comparison(t) for t in titulos]
    res_n = [_norm_text(r) for r in resumenes]
    ctx_n = [_norm_text(c) for c in (contextos or [""] * n)]
    url_n = [normalize_url(u) for u in (urls or [""] * n)]

    buckets: Dict[str, List[int]] = defaultdict(list)
    for i in range(n):
        tp = _prefix(tit_n[i], 8)
        rp = _prefix(res_n[i] or ctx_n[i], 12)
        if tp:
            buckets[f"t:{tp}"].append(i)
        if rp:
            buckets[f"r:{rp}"].append(i)
        if url_n[i]:
            buckets[f"u:{url_n[i]}"].append(i)
        if fechas is not None and horas is not None:
            fh = f"{normalizar_fecha(fechas[i])}|{normalizar_hora(horas[i])}"
            if fh != "|":
                buckets[f"d:{fh}"].append(i)
        for tok in list(_tokens_distintivos(tit_n[i] or res_n[i] or ctx_n[i]))[:8]:
            buckets[f"k:{tok}"].append(i)

    pares = _candidate_pairs(n, buckets, counter)

    embs = list(embeddings) if embeddings is not None else [None] * n
    # Unit-normalize once; per-pair sklearn cosine_similarity cost ~3ms each
    # (≈30s for 9k pairs). A dot product of unit vectors is the same number.
    unit = [None] * n
    for k, e in enumerate(embs):
        if e is None:
            continue
        v = np.asarray(e, dtype=float).ravel()
        nrm = np.linalg.norm(v)
        unit[k] = v / nrm if nrm > 0 else None
    can_cos = True

    # Per-row precomputation: tokens are built once, not once per pair.
    res_words = [r.split()[:MAX_CMP_WORDS] for r in res_n]
    res_pref = [_prefix(r, 10) for r in res_n]
    texto_row = [" ".join(x for x in (tit_n[i], res_n[i], ctx_n[i]) if x) for i in range(n)]
    tok_conf = [_tokens_distintivos(texto_row[i], min_len=3) for i in range(n)]
    tok_ov = [_tokens_distintivos(ctx_n[i] or texto_row[i]) for i in range(n)]

    def _conflicto(i: int, j: int) -> bool:
        ta, tb = tok_conf[i], tok_conf[j]
        for ga, gb in _ACCIONES_OPUESTAS:
            if (ta & ga and tb & gb) or (ta & gb and tb & ga):
                return True
        return False

    def _overlap(i: int, j: int) -> float:
        ta, tb = tok_ov[i], tok_ov[j]
        if not ta or not tb:
            return 0.0
        return len(ta & tb) / max(1, min(len(ta), len(tb)))

    for i, j in pares:
        if _conflicto(i, j):
            continue
        sim_t = _ratio(tit_n[i], tit_n[j], SIMILARITY_THRESHOLD_TITULOS) if tit_n[i] and tit_n[j] else 0.0
        sim_r = _ratio_palabras(res_words[i], res_words[j], SIMILARITY_THRESHOLD_RESUMEN)
        if sim_r < SIMILARITY_THRESHOLD_RESUMEN and res_pref[i] and res_pref[i] == res_pref[j]:
            # identical opening sentence still counts as the same resumen
            sim_r = max(sim_r, 0.90)
        semantic = 0.0
        if can_cos and unit[i] is not None and unit[j] is not None:
            semantic = float(np.dot(unit[i], unit[j]))
        overlap = _overlap(i, j)
        same_url = bool(url_n[i] and url_n[i] == url_n[j])
        mismo_hecho = (
            sim_t >= SIMILARITY_THRESHOLD_TITULOS
            or sim_r >= SIMILARITY_THRESHOLD_RESUMEN
            or same_url
            or (semantic >= SIMILARITY_THRESHOLD_SEMANTIC and overlap >= MIN_OVERLAP_GRUPO)
        )
        if mismo_hecho:
            dsu.union(i, j)
    return dsu.grupos(n)


def ids_grupo(grupos: Dict[int, List[int]]) -> List[str]:
    out = [""] * sum(len(v) for v in grupos.values())
    for numero, idxs in enumerate(grupos.values(), start=1):
        gid = f"G{numero:05d}"
        for i in idxs:
            out[i] = gid
    return out


# ---------------------------------------------------------------------------
# Tema / Subtema quality
# ---------------------------------------------------------------------------

_BAD_COLLAGE = re.compile(
    r"^(?:\S+\s+){2,}\S+$"
)


def _palabras(etiqueta: str) -> List[str]:
    return [p for p in re.split(r"\s+", (etiqueta or "").strip()) if p]


def validar_subtema(etiqueta: str) -> bool:
    """Reject ungrammatical keyword/title fragments. Noun phrase, 3–7 words."""
    if not etiqueta or not str(etiqueta).strip():
        return False
    et = str(etiqueta).strip().strip(" .;:¡!¿?")
    low = unidecode(et.lower())
    words = _palabras(et)
    if not (MIN_PALABRAS_SUBTEMA <= len(words) <= MAX_PALABRAS_SUBTEMA):
        return False
    last = unidecode(words[-1].lower().rstrip(".,;:!?¿¡"))
    if last in _TRAILING_INCOMPLETE or last in _QUESTION_TAILS:
        return False
    if last in _CONJUGATED_TAILS:
        return False
    if any(unidecode(w.lower().rstrip(".,;:")) in _CONJUGATED_TAILS for w in words):
        return False
    if any(unidecode(w.lower()) in {"como", "cómo"} for w in words):
        return False
    head = " ".join(unidecode(w.lower()) for w in words[:2])
    if unidecode(words[0].lower()) in _CHANNEL_PREFIXES or head in _CHANNEL_PREFIXES:
        return False
    if "|" in et or et.lower().startswith("video"):
        return False
    nexos = [unidecode(w.lower()) for w in words[1:] if unidecode(w.lower()) in _NEXOS]
    content = [w for w in words if unidecode(w.lower()) not in _NEXOS | STOPWORDS_ES]
    if len(content) >= 3 and not nexos:
        return False
    if len(words) <= 4 and not nexos:
        return False
    if "?" in et or "¿" in et:
        return False
    return True


def validar_tema(etiqueta: str) -> bool:
    if not etiqueta or not str(etiqueta).strip():
        return False
    words = _palabras(str(etiqueta))
    if not (2 <= len(words) <= 6):
        return False
    last = unidecode(words[-1].lower().rstrip(".,;:!?¿¡"))
    if last in _TRAILING_INCOMPLETE or last in _CONJUGATED_TAILS or last in _QUESTION_TAILS:
        return False
    if any(unidecode(w.lower()) in _CONJUGATED_TAILS for w in words):
        return False
    return True


def _capitalizar(frase: str) -> str:
    frase = re.sub(r"\s+", " ", (frase or "").strip())
    if not frase:
        return "Cobertura informativa general"
    return frase[0].upper() + frase[1:]


def fallback_subtema(contexto: str, titulo: str = "") -> str:
    """
    Deterministic grammatical noun phrase from a parsed sentence.
    Used when the model/heuristic would otherwise emit a keyword collage.
    """
    texto = " ".join(x for x in (contexto, titulo) if x).strip()
    norm = _norm_text(texto)
    sent = re.split(r"(?<=[\.\!\?])\s+", texto.strip())[0] if texto.strip() else ""
    sent_l = unidecode(sent.lower())

    if re.search(r"\b(terremoto|sismo|temblor)\b", norm) and re.search(
        r"\b(victima|victimas|muerto|muertos|herido|heridos|asciend)", norm
    ):
        return "Balance de víctimas del terremoto"
    if re.search(r"\b(sismo|terremoto|temblor)\b", norm) and re.search(
        r"\b(cali|recomend|saber|prepar|simulacro)\b", norm
    ):
        return "Recomendaciones ante sismos en Cali"
    if re.search(r"\b(terremoto|sismo|temblor)\b", norm):
        return "Cobertura del sismo en Colombia"
    if re.search(r"\b(rescate|solidaridad|topos)\b", norm):
        return "Solidaridad en labores de rescate"
    if re.search(r"\bdiporto\b", norm):
        if re.search(r"\b(jornada|resultado|resultados|fecha)\b", norm):
            return "Resultados de la jornada deportiva"
        return "Cobertura deportiva de Diporto"

    # Noun-phrase from first sentence: drop channel prefixes and verbs.
    raw = re.sub(r"^(video|en vivo|audio|streaming)\s*[\|:\-–—]\s*", "", sent, flags=re.I)
    raw = re.sub(r"<[^>]+>", "", raw)
    toks = [t for t in re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+", raw) if t]
    drop = STOPWORDS_ES | _CHANNEL_PREFIXES | _CONJUGATED_TAILS | _QUESTION_TAILS
    kept = []
    for t in toks:
        tl = unidecode(t.lower())
        if tl in drop or tl in _CONJUGATED_TAILS:
            continue
        if len(tl) < 3:
            continue
        kept.append(t)
        if len(kept) >= 4:
            break
    if len(kept) >= 2:
        frase = f"{kept[0]} de {' '.join(kept[1:3])}"
        frase = _capitalizar(frase)
        if validar_subtema(frase):
            return frase
        frase2 = f"Cobertura de {kept[0].lower()} en {kept[1]}"
        frase2 = _capitalizar(frase2)
        if validar_subtema(frase2):
            return frase2
    if kept:
        frase = f"Cobertura sobre {kept[0].lower()}"
        if validar_subtema(frase):
            return _capitalizar(frase)
    return "Cobertura de información relevante"


def fallback_tema(subtema: str, contexto: str = "") -> str:
    norm = _norm_text(f"{subtema} {contexto}")
    if re.search(r"\b(terremoto|sismo|temblor)\b", norm):
        return "Sismos y emergencias"
    if re.search(r"\b(rescate|solidaridad)\b", norm):
        return "Emergencias y solidaridad"
    if re.search(r"\bdiporto\b", norm) or re.search(r"\bdeporte", norm):
        return "Actualidad deportiva"
    words = [w for w in _palabras(subtema) if unidecode(w.lower()) not in _NEXOS | STOPWORDS_ES]
    if len(words) >= 2:
        return _capitalizar(f"{words[0]} {words[1]}")
    if words:
        return _capitalizar(f"Asuntos de {words[0].lower()}")
    return "Cobertura informativa"


def etiquetar_gramatical(contexto: str, titulo: str = "",
                         chat_fn: Optional[Callable] = None) -> Tuple[str, str]:
    sub = fallback_subtema(contexto, titulo)
    tema = fallback_tema(sub, contexto)
    if validar_subtema(sub) and validar_tema(tema):
        return tema, sub
    if chat_fn is None:
        return (
            tema if validar_tema(tema) else "Cobertura informativa",
            sub if validar_subtema(sub) else "Cobertura de información relevante",
        )
    return tema, sub


# ---------------------------------------------------------------------------
# Batched classification
# ---------------------------------------------------------------------------

def _parse_chat_payload(raw: Any) -> Dict[str, Any]:
    if raw is None:
        return {}
    if isinstance(raw, dict):
        return raw
    text = raw if isinstance(raw, str) else str(raw)
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, flags=re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
    return {}


def clasificar_lotes(
    items: Sequence[Dict[str, Any]],
    marca: str,
    aliases=None,
    chat_fn: Optional[Callable] = None,
    batch_size: int = CHAT_BATCH_SIZE,
    call_counter: Optional[CallCounter] = None,
) -> List[Dict[str, str]]:
    """
    One JSON ChatCompletion per 25–40 items. Never one call per news/group
    unless a single leftover batch remains. Heuristic first; LLM only when
    the noun phrase cannot be formed grammatically.
    """
    results: List[Dict[str, str]] = [{"tono": "Neutro", "tema": "", "subtema": ""} for _ in items]
    pending: List[int] = []
    for i, it in enumerate(items):
        ctx = str(it.get("contexto") or "")
        titulo = str(it.get("titulo") or "")
        tema, sub = etiquetar_gramatical(ctx, titulo, chat_fn=None)
        results[i]["tema"] = tema
        results[i]["subtema"] = sub
        # Tone heuristic is weak; if chat is available we still batch tone+repair.
        if chat_fn is not None and (not validar_subtema(sub) or True):
            # Always send to batch when chat_fn exists so tone is model-based,
            # but reuse heuristic labels if the model fails.
            pending.append(i)
        results[i]["tono"] = it.get("tono") or "Neutro"

    if chat_fn is None or not pending:
        return results

    aliases_str = ", ".join(_lista_alias(marca, aliases)[1:6])

    def _run_batch(chunk_idx: List[int]) -> None:
        payload = []
        for i in chunk_idx:
            it = items[i]
            payload.append({
                "id": i,
                "contexto": str(it.get("contexto") or "")[:900],
                "titulo_interno": normalize_title_for_comparison(it.get("titulo") or "")[:180],
            })
        prompt = (
            f"Eres analista de reputación en Colombia. Marca: '{marca}'"
            + (f" (alias: {aliases_str})" if aliases_str else "")
            + ".\nAnaliza SOLO el campo 'contexto' de cada ítem (Contexto analizado).\n"
            "Devuelve JSON: {\"items\":[{\"id\":0,\"tono\":\"Positivo|Negativo|Neutro\","
            "\"tema\":\"frase nominal general 2-5 palabras\","
            "\"subtema\":\"frase nominal específica 3-7 palabras\"}]}\n"
            "REGLAS DE SUBTEMA: frase nominal completa en español colombiano, "
            "con preposición (de/del/en/para/sobre), 3-7 palabras. "
            "PROHIBIDO: colages de keywords, colas interrogativas ('cómo saber si'), "
            "verbos conjugados sueltos ('ascienden'), prefijos de canal ('Video'), "
            "copiar el titular, fragmentar el título.\n"
            "Ejemplos correctos: 'Balance de víctimas del terremoto', "
            "'Recomendaciones ante sismos en Cali'.\n"
            "TEMA más general que el subtema, coherente.\n"
            "TONO = impacto reputacional DIRECTO sobre la marca, no el tono general.\n"
            f"ÍTEMS:\n{json.dumps(payload, ensure_ascii=False)}"
        )
        if call_counter:
            with _COUNTER_LOCK:
                call_counter.chat += 1
                call_counter.chat_items += len(chunk_idx)
        try:
            raw = chat_fn(prompt)
            data = _parse_chat_payload(raw)
            rows = data.get("items") or data.get("resultados") or []
            by_id = {}
            for row in rows:
                try:
                    by_id[int(row.get("id"))] = row
                except Exception:
                    continue
            for i in chunk_idx:
                row = by_id.get(i, {})
                tono = str(row.get("tono") or results[i]["tono"]).strip().title()
                if tono not in ("Positivo", "Negativo", "Neutro"):
                    tono = "Neutro"
                tema = str(row.get("tema") or results[i]["tema"]).strip()
                sub = str(row.get("subtema") or results[i]["subtema"]).strip()
                if not validar_subtema(sub):
                    sub = fallback_subtema(items[i].get("contexto", ""), items[i].get("titulo", ""))
                if not validar_tema(tema):
                    tema = fallback_tema(sub, items[i].get("contexto", ""))
                results[i] = {"tono": tono, "tema": _capitalizar(tema), "subtema": _capitalizar(sub)}
        except Exception:
            for i in chunk_idx:
                if not validar_subtema(results[i]["subtema"]):
                    results[i]["subtema"] = fallback_subtema(
                        items[i].get("contexto", ""), items[i].get("titulo", "")
                    )
                if not validar_tema(results[i]["tema"]):
                    results[i]["tema"] = fallback_tema(
                        results[i]["subtema"], items[i].get("contexto", "")
                    )

    chunks = [pending[s:s + batch_size] for s in range(0, len(pending), batch_size)]
    workers = max(1, min(CHAT_PARALLEL_BATCHES, len(chunks)))
    if workers == 1:
        for ch in chunks:
            _run_batch(ch)
    else:
        # Bounded concurrency: a handful of batches in flight, never one
        # request per news item and never an unbounded retry storm.
        with ThreadPoolExecutor(max_workers=workers) as ex:
            list(ex.map(_run_batch, chunks))
    return results


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def _mejor_representante(idxs: Sequence[int], contextos: Sequence[str],
                         embeddings: Optional[Sequence[Any]] = None) -> int:
    if not idxs:
        return 0
    con = [i for i in idxs if str(contextos[i] or "").strip()]
    pool = con or list(idxs)
    if embeddings is not None:
        vecs = [(i, embeddings[i]) for i in pool if embeddings[i] is not None]
        if len(vecs) >= 2:
            M = np.array([np.asarray(v, dtype=float).ravel() for _, v in vecs])
            centro = M.mean(axis=0)
            norms = np.linalg.norm(M, axis=1) * (np.linalg.norm(centro) or 1.0)
            sims = (M @ centro) / np.where(norms == 0, 1.0, norms)
            return vecs[int(np.argmax(sims))][0]
    return max(pool, key=lambda i: len(str(contextos[i] or "")))


def process_pipeline(
    rows: Sequence[Dict[str, Any]],
    marca: str,
    aliases=None,
    km: Optional[Dict[str, str]] = None,
    embed_fn: Optional[Callable[[List[str]], List[Any]]] = None,
    chat_fn: Optional[Callable] = None,
    pkl_tono_fn: Optional[Callable] = None,
    pkl_tema_fn: Optional[Callable] = None,
    progress: Optional[ProgressTracker] = None,
    on_stage: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Tuple[List[Dict[str, Any]], ProgressTracker]:
    """
    Canonical API_JC path:
    Limpieza → Duplicados → Contexto → Embedding único → Agrupación
    → Tono / Tema/Subtema (batched, cached by group) → Excel caller.
    Duplicates are decided before the LLM and keep Tono IA = Duplicada.
    """
    km = km or KEYMAP
    progress = progress or ProgressTracker(on_stage=on_stage)
    if on_stage is not None and progress.on_stage is None:
        progress.on_stage = on_stage
    out = [dict(r) for r in rows]
    progress.stage("Limpieza", f"{len(out)} filas")

    out = detectar_duplicados(out, km)
    n_dup = sum(1 for r in out if r.get("is_duplicate"))
    progress.stage("Duplicados", f"{n_dup} duplicadas")

    titulos, resumenes, cuerpos = [], [], []
    for r in out:
        tit = r.get("_titulo_original", r.get(km["titulo"], ""))
        r["_titulo_original"] = titulo_original(tit)
        r[km["titulo"]] = r["_titulo_original"]
        titulos.append(r["_titulo_original"])
        resumenes.append(r.get(km["resumen"], ""))
        cuerpos.append(r.get(km.get("cuerpo", "Cuerpo Completo"), ""))

    for i, r in enumerate(out):
        meta = extraer_contexto_analizado(
            titulos[i], resumenes[i], marca, aliases, cuerpos[i]
        )
        r["Contexto analizado"] = meta["contexto"]
        r["Coincidencia marca"] = meta["coincidencia"]
        r["Origen coincidencia"] = meta["origen"]
    contextos = [r["Contexto analizado"] for r in out]
    progress.stage("Contexto")

    embeddings = None
    texts_emb = [
        (contextos[i] or f"{titulos[i]}. {resumenes[i]}")[:1800]
        for i in range(len(out))
    ]
    if embed_fn is not None:
        embeddings = embed_fn(texts_emb)
        progress.calls.embed += 1
        progress.calls.embed_items += len(texts_emb)
    progress.stage("Embedding único", f"{len(texts_emb)} textos")

    grupos = agrupar_noticias_bloqueado(
        titulos, resumenes, contextos, embeddings,
        urls=[url_de_celda_link(r.get(km["link_nota"])) for r in out],
        fechas=[r.get(km["fecha"]) for r in out],
        horas=[r.get(km["hora"]) for r in out],
        counter=progress.comparisons,
    )
    gids = ids_grupo(grupos)
    for i, r in enumerate(out):
        r["Grupo noticia"] = gids[i]
    progress.stage("Agrupación", f"{len(grupos)} grupos · pares={progress.comparisons.n}")

    reps = []
    gid_of_rep = []
    for gid, idxs in grupos.items():
        ri = _mejor_representante(idxs, contextos, embeddings)
        # Prefer a non-duplicate representative when the group has one.
        nondup = [j for j in idxs if not out[j].get("is_duplicate")]
        if nondup:
            ri = _mejor_representante(nondup, contextos, embeddings)
        reps.append({
            "id": ri,
            "contexto": contextos[ri],
            "titulo": titulos[ri],
        })
        gid_of_rep.append(gid)

    labels = clasificar_lotes(
        reps, marca, aliases, chat_fn=chat_fn,
        call_counter=progress.calls,
    )
    if pkl_tono_fn is not None:
        try:
            tonos = pkl_tono_fn([x["contexto"] for x in reps])
            for i, t in enumerate(tonos or []):
                val = t.get("tono", t) if isinstance(t, dict) else t
                if val:
                    labels[i]["tono"] = str(val)
        except Exception:
            pass
    if pkl_tema_fn is not None:
        try:
            temas = pkl_tema_fn([x["contexto"] for x in reps])
            for i, t in enumerate(temas or []):
                if t:
                    labels[i]["tema"] = str(t)
        except Exception:
            pass

    by_gid = {gid: labels[k] for k, gid in enumerate(gid_of_rep)}
    progress.stage("Tono", f"{progress.calls.chat} llamadas chat")
    progress.stage("Tema/Subtema", f"{len(grupos)} grupos cacheados")

    for gid, idxs in grupos.items():
        lab = by_gid.get(gid) or {"tono": "Neutro", "tema": "Cobertura informativa",
                                  "subtema": "Cobertura de información relevante"}
        for i in idxs:
            if out[i].get("is_duplicate"):
                out[i]["Tono IA"] = "Duplicada"
                out[i]["Tema"] = "-"
                out[i]["Subtema"] = "-"
                continue
            out[i]["Tono IA"] = lab["tono"]
            out[i]["Tema"] = lab["tema"]
            out[i]["Subtema"] = lab["subtema"]
    return out, progress


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def generate_output_excel(rows: Sequence[Dict[str, Any]],
                          km: Optional[Dict[str, str]] = None) -> bytes:
    km = km or KEYMAP
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    ws.append(list(OUTPUT_COLUMNS))
    font_header = Font(bold=True)
    font_hyperlink = Font(color="000000", underline=None)
    align_left = Alignment(horizontal="left")
    for i, _ in enumerate(OUTPUT_COLUMNS, start=1):
        ws.cell(row=1, column=i).font = font_header

    for row in rows:
        out_vals = []
        links = {}
        for ci, h in enumerate(OUTPUT_COLUMNS, start=1):
            if h == "Título":
                val = row.get("_titulo_original", row.get(h))
                val = titulo_original(val)
            else:
                val = row.get(h)
            if h == "Fecha" and val is not None and hasattr(val, "to_pydatetime"):
                val = val.to_pydatetime()
            if isinstance(val, dict) and "url" in val:
                cv = val.get("value", "Link")
                if val.get("url"):
                    links[ci] = val["url"]
                out_vals.append(cv)
            elif isinstance(val, str) and val.startswith("http"):
                out_vals.append("Link")
                links[ci] = val
            else:
                out_vals.append(val)
        ws.append(out_vals)
        current = ws.max_row
        for ci, url in links.items():
            cell = ws.cell(row=current, column=ci)
            cell.hyperlink = url
            cell.font = font_hyperlink
            cell.alignment = align_left
        date_cell = ws.cell(row=current, column=OUTPUT_COLUMNS.index("Fecha") + 1)
        if isinstance(date_cell.value, (datetime.datetime, datetime.date)):
            date_cell.number_format = "DD/MM/YYYY"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_titulo_and_links_from_xlsx(data: bytes) -> List[Dict[str, Any]]:
    wb = load_workbook(io.BytesIO(data), data_only=False)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    rows = []
    for excel_row in ws.iter_rows(min_row=2):
        item = {}
        for h, cell in zip(headers, excel_row):
            if h == "Título":
                item["Título"] = cell.value
            if h in ("Link Nota", "Link (Streaming - Imagen)"):
                item[h] = valor_con_hipervinculo(cell)
            else:
                item.setdefault(h, cell.value)
        rows.append(item)
    return rows
