# COD_CLS_V16.py
# Objetivo: estandarizar estratigrafía por intervalos y generar Resultados_CLS (template V1)
# - Usa Diccionario_estratigrafia_V1.xlsx (LITOLOGIA, INDICADORES, MODIFICADORES, CLASES_TAMANO)
# - Aplica reglas del reglamento (dominancia, "con", M_tipo, etc.)
# - Calcula proporcion_sedimento SOLO con pesos (NO usa % del texto):
#     * si hay D,S,M -> 0.6 / 0.3 / 0.1
#     * si no hay M -> 0.7 / 0.3 (el 0.1 se suma a D)
#
# Requiere: pandas, numpy, openpyxl

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd

# =========================
# CONFIG (por defecto usa la base real de la memoria)
# =========================
DB_PATH = Path("Pozos_memoria_CGC.xlsx")
DICT_PATH = Path("Diccionario_estratigrafia.xlsx")
TEMPLATE_PATH = Path("Template_resultado_CLS.xlsx")
OUTPUT_PATH = Path("Resultados_CLS_FINAL.xlsx")

# =========================
# NORMALIZACIÓN
# =========================
def norm_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = re.sub(r"[^A-Z0-9%.,\sÑ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def pick_col(df: pd.DataFrame, candidates):
    cols = {norm_text(c): c for c in df.columns}
    for cand in candidates:
        k = norm_text(cand)
        if k in cols:
            return cols[k]
    # fuzzy contains
    for cand in candidates:
        k = norm_text(cand)
        for nk, orig in cols.items():
            if k and k in nk:
                return orig
    return None


def insert_after_column(df: pd.DataFrame, col_name: str, after_col: str) -> pd.DataFrame:
    """Inserta/mueve una columna justo después de otra si ambas existen."""
    if col_name not in df.columns or after_col not in df.columns:
        return df
    cols = list(df.columns)
    cols.insert(cols.index(after_col) + 1, cols.pop(cols.index(col_name)))
    return df[cols]

# =========================
# UTIL: matching de términos con plural simple
# =========================
def find_all_terms(text: str, terms):
    hits = []
    for term in terms:
        suffix = r"(S)?" if (len(term) > 3 and not term.endswith("S")) else ""
        pattern = r"(?<![A-ZÑ0-9])" + re.escape(term) + suffix + r"(?![A-ZÑ0-9])"
        for m in re.finditer(pattern, text):
            hits.append((m.start(), m.end(), term))
    hits.sort(key=lambda x: (x[0], -(x[1]-x[0])))
    chosen = []
    for h in hits:
        if chosen and h[0] < chosen[-1][1]:
            # overlap: quedarse con el más largo si empieza igual
            if h[0] == chosen[-1][0] and (h[1]-h[0]) > (chosen[-1][1]-chosen[-1][0]):
                chosen[-1] = h
            continue
        chosen.append(h)
    return chosen

_pct_re = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")
def pct_near(text: str, start: int, end: int):
    """Asigna porcentaje a un término buscando el % MÁS CERCANO al término.
    Esto evita el error típico de repetir el primer porcentaje del intervalo.

    Soporta:
      - '40% ARCILLA'
      - 'ARCILLA 40%'
      - 'ARCILLA (40%)'
      - '40 % ARCILLA'
    """
    if not text:
        return None

    # límites de búsqueda alrededor del término
    left = max(0, start - 60)
    right = min(len(text), end + 60)

    # extraer % con su posición absoluta
    candidates = []
    for m in _pct_re.finditer(text[left:right]):
        val = m.group(1).replace(",", ".")
        try:
            v = float(val)
        except Exception:
            continue
        abs_s = left + m.start()
        abs_e = left + m.end()
        # distancia al término: mínimo a cualquiera de sus extremos
        dist = min(abs(abs_s - start), abs(abs_s - end), abs(abs_e - start), abs(abs_e - end))
        candidates.append((dist, abs_s, abs_e, v))

    if not candidates:
        return None

    # escoger el más cercano; en empate, preferir el que esté más a la izquierda (consistente)
    candidates.sort(key=lambda t: (t[0], t[1]))
    return candidates[0][3]

def pct_after(text: str, end: int):
    # compatibilidad: búsqueda hacia adelante (se mantiene para llamadas antiguas)
    return pct_near(text, end, end)


# =========================
# LOAD DICCIONARIO (V1)
# =========================
def load_dictionary(dict_path: Path):
    lit = pd.read_excel(dict_path, sheet_name="LITOLOGIA").fillna("")
    mods = pd.read_excel(dict_path, sheet_name="MODIFICADORES").fillna("")
    inds = pd.read_excel(dict_path, sheet_name="INDICADORES").fillna("")
    tam = pd.read_excel(dict_path, sheet_name="CLASES_TAMANO").fillna("")

    lit["PAL_N"] = lit["PALABRA"].map(norm_text)
    mods["PAL_N"] = mods["PALABRA"].map(norm_text)
    inds["PAL_N"] = inds["PALABRA"].map(norm_text)
    tam["LAB_N"] = tam["TAMANO_LABEL"].map(norm_text)

    # litologías
    lit_terms = sorted(set(lit["PAL_N"].dropna()), key=len, reverse=True)
    # columnas esperadas en LITOLOGIA (algunas pueden no existir según versión del diccionario)
    lit_cols_base = ["T","D","S","M","CLASE","TAMANO_D","TAMANO_S","TAMANO_M",
                     "REDONDEAMIENTO","COMPACTACION","PLASTICIDAD","ORIGEN","TIPO_ROCA"]
    lit_cols = [c for c in lit_cols_base if c in lit.columns]

    lit_map = (
        lit.drop_duplicates("PAL_N")
           .set_index("PAL_N")[lit_cols]
           .to_dict("index")
    )

    # ruido
    ruido = set(inds.loc[inds["TIPO_INDICADOR"].astype(str).str.upper().eq("RUIDO"), "PAL_N"].dropna())

    # tamaño: MOD_TAM -> orden
    tam_map = (
        mods.loc[mods["TIPO_MODIFICADOR"].astype(str).str.upper().eq("MOD_TAM")]
            .dropna(subset=["PAL_N","TAMANO_ORDEN"])
            .drop_duplicates("PAL_N")
            .set_index("PAL_N")["TAMANO_ORDEN"]
            .to_dict()
    )
    tam_terms = sorted(set(tam_map.keys()), key=len, reverse=True)

    # orden <-> label
    order_to_label = (
        tam.dropna(subset=["TAMANO_ORDEN","LAB_N"])
           .drop_duplicates("TAMANO_ORDEN")
           .set_index("TAMANO_ORDEN")["LAB_N"]
           .to_dict()
    )
    label_to_order = (
        tam.dropna(subset=["TAMANO_ORDEN","LAB_N"])
           .drop_duplicates("LAB_N")
           .set_index("LAB_N")["TAMANO_ORDEN"]
           .to_dict()
    )

    # indicadores (con prioridad)
    ind_rules = []
    for _, row in inds.iterrows():
        pal = norm_text(row.get("PALABRA",""))
        if not pal:
            continue
        tipo = norm_text(row.get("TIPO_INDICADOR",""))
        if tipo == "RUIDO":
            continue
        pr = row.get("PRIORIDAD", 0)
        try:
            pr = int(pr) if str(pr).strip() != "" else 0
        except Exception:
            pr = 0
        ind_rules.append({
            "pal": pal,
            "tipo": tipo,
            "accion": norm_text(row.get("ACCION","")),
            "prioridad": pr,
            "obs": norm_text(row.get("OBSERVACIONES","")),
        })
    ind_rules.sort(key=lambda d: -d["prioridad"])

    # modificadores (con prioridad)
    mod_rules = []
    for _, row in mods.iterrows():
        pal = norm_text(row.get("PALABRA",""))
        if not pal:
            continue
        tipo = norm_text(row.get("TIPO_MODIFICADOR",""))
        afecta = norm_text(row.get("AFECTA_A",""))
        val = norm_text(row.get("VALOR",""))
        pr = row.get("PRIORIDAD", 0)
        try:
            pr = int(pr) if str(pr).strip() != "" else 0
        except Exception:
            pr = 0
        mod_rules.append({
            "pal": pal,
            "tipo": tipo,
            "afecta": afecta,
            "valor": val,
            "prioridad": pr,
            "tamano_orden": row.get("TAMANO_ORDEN",""),
        })
    mod_rules.sort(key=lambda d: -d["prioridad"])

    return lit_terms, lit_map, ruido, tam_terms, tam_map, order_to_label, label_to_order, ind_rules, mod_rules

# =========================
# INDICADORES / MODIFICADORES
# =========================
def detectar_indicadores(text_norm: str, ind_rules):
    s = f" {text_norm} "
    hits = []
    for r in ind_rules:
        pal = r["pal"]
        # plural simple
        suffix = r"(S)?" if (len(pal) > 3 and not pal.endswith("S")) else ""
        if re.search(r"(?<![A-ZÑ0-9])" + re.escape(pal) + suffix + r"(?![A-ZÑ0-9])", s):
            hits.append(r)
    hits.sort(key=lambda d: -d["prioridad"])
    return hits

def aplicar_indicadores(ind_hits):
    # ctx: salidas mínimas y conservadoras
    ctx = {"estructuras": [], "abundancia": None, "force_m_tipo": None}
    for h in ind_hits:
        tipo = h["tipo"]
        pal = h["pal"]
        if tipo in {"ESTRUCTURA","GEOMETRIA"}:
            if pal not in ctx["estructuras"]:
                ctx["estructuras"].append(pal)
        elif tipo in {"CANTIDAD","CANTIDAD_FUERTE","CANTIDAD_MEDIA","CANTIDAD_BAJA"}:
            # tomar la primera (alta prioridad) como "abundancia" principal
            if ctx["abundancia"] is None:
                ctx["abundancia"] = pal
        elif tipo in {"MATRIZ","CEMENTO"}:
            ctx["force_m_tipo"] = "MAT"
    return ctx

def detectar_modificadores(text_norm: str, mod_rules):
    s = f" {text_norm} "
    hits = []
    for r in mod_rules:
        pal = r["pal"]
        suffix = r"(S)?" if (len(pal) > 3 and not pal.endswith("S")) else ""
        if re.search(r"(?<![A-ZÑ0-9])" + re.escape(pal) + suffix + r"(?![A-ZÑ0-9])", s):
            hits.append(r)
    hits.sort(key=lambda d: -d["prioridad"])
    return hits

def aplicar_modificadores(mod_hits):
    ctx = {
        "seleccion": [],
        "forma": [],
        "plasticidad": [],
        "origen": [],
        "color": [],
        "estructura": [],
        "atributo_otro": [],
        "suggest_m_tipo": None,
    }
    for h in mod_hits:
        tipo = h["tipo"]
        val = h["valor"] or h["pal"]
        if tipo == "MOD_SEL":
            if val not in ctx["seleccion"]:
                ctx["seleccion"].append(val)
        elif tipo == "MOD_RED":
            if val not in ctx["forma"]:
                ctx["forma"].append(val)
        elif tipo == "MOD_PLAST":
            if val not in ctx["plasticidad"]:
                ctx["plasticidad"].append(val)
        elif tipo == "MOD_ORIG":
            if val not in ctx["origen"]:
                ctx["origen"].append(val)
        elif tipo == "MOD_COLOR":
            if val not in ctx["color"]:
                ctx["color"].append(val)
        elif tipo in {"MOD_EST","ESTRUCTURA"}:
            if val not in ctx["estructura"]:
                ctx["estructura"].append(val)
            ctx["suggest_m_tipo"] = "ACC"
        elif tipo == "MOD_COMP":
            if val not in ctx["atributo_otro"]:
                ctx["atributo_otro"].append(val)
        # MOD_TAM se maneja aparte para tam_min/tam_max
    return ctx

# =========================
# M_tipo (matriz vs accesorio)
# =========================
_MAT_RE = re.compile(r"\b(MATRIZ|EN\s+MATRIZ|CEMENTO|CEMENTADA|CEMENTADO|MATRICIAL)\b")
_ACC_RE = re.compile(r"\b(LENTE|LENTES|INTERCALACION|INTERCALACIONES|INTERCALADO|INTERCALADA|NIVEL|NIVELES|CAPA|CAPAS|LAMINA|LAMINAS|LAMINACION|BANDA|BANDAS|VETA|VETAS|VENILLA|VENILLAS|PRESENCIA)\b")
_FINE_M = {"S","C","L","F"}
_COARSE_M = {"B","G"}

def determinar_m_tipo(T: str, M: str, texto_norm: str, force: str | None = None) -> str:
    if T != "U" or M == "0":
        return "0"
    if force in {"MAT","ACC"}:
        return force
    s = (texto_norm or "").upper()
    if _MAT_RE.search(s):
        return "MAT"
    if _ACC_RE.search(s):
        return "ACC"
    # heurística por granulometría del código M
    if M in _FINE_M:
        return "MAT"
    if M in _COARSE_M:
        return "ACC"
    return "0"

# =========================
# CLS helpers
# =========================
def split_cls(cls: str) -> dict:
    c = "" if pd.isna(cls) else str(cls).strip().upper()
    c = (c + "0000")[:4]
    return {"T": c[0], "D": c[1], "S": c[2], "M": c[3]}


# =========================
# VALIDACIÓN / SANITIZACIÓN CLS (evita letras sin significado)
# =========================
_ALLOWED = {
    "R": {
        "D": set("UFW0"),   # estado roca
        "S": set("PN0"),    # porosidad inicial
        "M": set("0"),      # no aplica
    },
    "U": {
        # permitimos A y S para "Arena" por compatibilidad de versiones
        "D": set("BGSALCFK0"),
        "S": set("BGSALCFK0"),
        "M": set("BGSALCFK0"),
    },
    "S": {
        "D": set("OGSALCFK0"),
        "S": set("OGSALCFK0"),
        "M": set("OGSALCFK0"),
    },
    "0": {"D": set("0"), "S": set("0"), "M": set("0")},
}

def sanitize_cls_code(code: str) -> str:
    """Fuerza coherencia del CLS por contexto.
    - R: D∈{U,F,W,0}, S∈{P,N,0}, M=0
    - U/S: D/S/M deben pertenecer a su alfabeto permitido
    - Si T inválido -> 0000
    """
    c = "" if pd.isna(code) else str(code).strip().upper()
    c = (c + "0000")[:4]
    T, D, S, M = c[0], c[1], c[2], c[3]
    if T not in {"R","U","S"}:
        return "0000"
    rules = _ALLOWED.get(T, _ALLOWED["0"])
    D = D if D in rules["D"] else "0"
    S = S if S in rules["S"] else "0"
    # Roca: M siempre 0, independiente de lo que venga
    if T == "R":
        M = "0"
    else:
        M = M if M in rules["M"] else "0"
    return f"{T}{D}{S}{M}"

def remove_repeated_materials(code: str) -> str:
    """Evita repeticiones en D/S/M.
    Si un material ya apareció, los siguientes se vuelven 0.
    Aplica principalmente a T=U y T=S (sedimentos/suelos).
    """
    c = ("" if pd.isna(code) else str(code).strip().upper())
    c = (c + "0000")[:4]
    T, D, S, M = c[0], c[1], c[2], c[3]

    # Para roca, no aplica (M=0 por sanitize). Devolver tal cual.
    if T == "R":
        return c

    seen = set()
    if D != "0":
        seen.add(D)

    if S != "0":
        if S in seen:
            S = "0"
        else:
            seen.add(S)

    if M != "0" and M in seen:
        M = "0"

    return f"{T}{D}{S}{M}"


def compact_cls_no_gaps(code: str) -> str:
    """Evita códigos con '0' entre dos letras (p.ej., UG0B).
    Si T != R y S == '0' y M != '0' -> desplaza M a S y deja M=0.
    """
    c = ("" if pd.isna(code) else str(code).strip().upper())
    c = (c + "0000")[:4]
    T, D, S, M = c[0], c[1], c[2], c[3]
    if T == "R":
        return c
    if S == "0" and M != "0":
        S, M = M, "0"
    return f"{T}{D}{S}{M}"

def detect_fines_hint(texto_norm: str) -> str:
    """Detecta presencia de finos en el texto normalizado.
    Prioridad: Arcilla(C) > Limo(L) > Finos(F)
    """
    t = texto_norm or ""
    if "ARCILL" in t:
        return "C"
    if "LIMO" in t or "LIMOS" in t:
        return "L"
    if "FINO" in t or "FINOS" in t:
        return "F"
    return "0"

def rescue_fines_to_matrix_with_hint(code: str, fines_hint: str) -> str:
    """Asegura que finos (C/L/F) no se pierdan cuando hay demasiadas litologías.

    Regla (solo T=U):
      - Si hay hint de finos y no está en D/S/M, forzar M=fines_hint.
      - Si M ya estaba ocupado por un componente grueso, se reemplaza por los finos (prioridad geológica).
    Además, compacta al final para evitar códigos tipo UF0L o UG0C.
    """
    if not code or pd.isna(code):
        return code
    c = str(code).strip().upper()
    c = (c + "0000")[:4]
    T, D, S, M = c[0], c[1], c[2], c[3]

    if T != "U":
        return compact_cls_no_gaps(c)
    if fines_hint not in {"C", "L", "F"}:
        return compact_cls_no_gaps(c)
    if fines_hint in {D, S, M}:
        return compact_cls_no_gaps(c)

    coarse = {"B", "G", "S", "A"}

    # si M está libre -> usarlo
    if M == "0":
        M = fines_hint
        return compact_cls_no_gaps(f"{T}{D}{S}{M}")

    # si ya hay 3 componentes y M es grueso -> priorizar finos como matriz
    if M in coarse and (D in coarse or D == "0") and (S in coarse or S == "0"):
        M = fines_hint
        return compact_cls_no_gaps(f"{T}{D}{S}{M}")

    return compact_cls_no_gaps(c)


# Desc de clase por T (para hoja desglose_codificacion T/D/S/M)
__DESC_CLASE_BY_T = {
    "R": {"T": "Tipo de unidad", "D": "Estado de la roca", "S": "Porosidad inicial", "M": "No aplica"},
    "U": {"T": "Tipo de unidad", "D": "Componente dominante", "S": "Componente secundario", "M": "Matriz"},
    "S": {"T": "Tipo de unidad", "D": "Componente principal del suelo", "S": "Componente secundario", "M": "Componente accesorio"},
    "0": {"T": "Tipo de unidad", "D": "Sin información", "S": "Sin información", "M": "Sin información"},
}


def desc_codigo_por_letra(letra: str, T_ctx: str, clase: str) -> str:
    l = ("" if pd.isna(letra) else str(letra)).strip().upper()

    # Clase T: tipo de unidad
    if clase == "T":
        return {"R":"Roca","U":"Sedimento","S":"Suelo","0":"Sin información"}.get(l, "Sin información")

    # Contexto roca: D=estado (U/F/W), S=porosidad (P/N), M no aplica
    if T_ctx == "R":
        if clase == "D":
            return {"U":"Sana","F":"Fracturada","W":"Meteorizada","0":"Sin información"}.get(l, "Sin información")
        if clase == "S":
            return {"P":"Porosa","N":"No porosa/impermeable","0":"Sin información"}.get(l, "Sin información")
        if clase == "M":
            return "No aplica"
        return "Sin información"

    # Materiales (sedimento/suelo)
    return {
        "B": "Bolones/Bloques",
        "G": "Grava",
        "S": "Arena",
        "A": "Arena",
        "L": "Limo",
        "C": "Arcilla",
        "K": "Arcilla",
        "F": "Finos",
        "O": "Orgánico",
        "0": "Sin información",
    }.get(l, "Sin información")
    # material
    return {
        "B": "Bolones/Bloques",
        "G": "Grava",
        "A": "Arena",
        "L": "Limo",
        "C": "Arcilla",
        "K": "Arcilla",
        "F": "Finos",
        "O": "Orgánico",
        "0": "Sin información",
    }.get(l, "Sin información")

# =========================
# CLASIFICACIÓN (reglas reglamento + diccionario)
# =========================
def classify_interval(text_norm: str, lit_terms, lit_map, ruido, tam_terms, tam_map, order_to_label, label_to_order, ind_rules, mod_rules):
    out = {
        "cls": "0000",
        "m_tipo": "0",
        "porcentaje_arcilla": np.nan,
        "tipo_unidad": "Sin información",
        "litologia_dom": "Sin información",
        "grupo_simplificado": "Sin información",
        "confiabilidad": "BAJA",
        "tam_min": np.nan,
        "tam_max": np.nan,
        "desglose_litos": [],
        "estructuras_txt": "",
        "seleccion_txt": "",
        "forma_txt": "",
        "plasticidad_txt": "",
        "origen_txt": "",
        "tipo_roca_txt": "",
        "color_txt": "",
        "atributo_otro_txt": "",
    }

    s = (text_norm or "").strip()
    if not s:
        return out

    if any(f" {w} " in f" {s} " for w in ruido):
        return out

    # indicadores / modificadores (ordenados por prioridad)
    ind_ctx = aplicar_indicadores(detectar_indicadores(s, ind_rules))
    mod_ctx = aplicar_modificadores(detectar_modificadores(s, mod_rules))

    # estructuras combinadas
    est_all = []
    est_all.extend(ind_ctx.get("estructuras", []))
    est_all.extend(mod_ctx.get("estructura", []))
    est_uniq = []
    seen = set()
    for x in est_all:
        x = x.strip()
        if x and x not in seen:
            est_uniq.append(x); seen.add(x)
    out["estructuras_txt"] = "; ".join(est_uniq)

    out["seleccion_txt"] = "; ".join(mod_ctx.get("seleccion", []))
    out["forma_txt"] = "; ".join(mod_ctx.get("forma", []))
    out["plasticidad_txt"] = "; ".join(mod_ctx.get("plasticidad", []))
    out["origen_txt"] = "; ".join(mod_ctx.get("origen", []))
    out["color_txt"] = "; ".join(mod_ctx.get("color", []))

    # colores heurísticos (si diccionario no los trae)
    if not out["color_txt"]:
        COLOR_TERMS = ["GRIS","NEGRO","NEGRA","BLANCO","BLANCA","ROJO","ROJA","CAFE","AMARILLO","AMARILLA","VERDE","AZUL","MARRON","PARDO","ROSADO","ROSADA","MORADO","MORADA"]
        found = []
        for c in COLOR_TERMS:
            suffix = r"(S)?" if (len(c) > 3 and not c.endswith("S")) else ""
            if re.search(r"(?<![A-ZÑ0-9])" + re.escape(c) + suffix + r"(?![A-ZÑ0-9])", f" {s} "):
                found.append(c)
        # dedup
        seen=set(); f2=[]
        for x in found:
            if x not in seen:
                f2.append(x); seen.add(x)
        out["color_txt"] = "; ".join(f2)

    # atributo_otro: abundancia + compactación/estado, excluyendo tamaños
    size_set = set(tam_terms)
    otros = []
    if ind_ctx.get("abundancia"):
        otros.append(ind_ctx["abundancia"])
    otros.extend(mod_ctx.get("atributo_otro", []))

    def _clean_otro(v):
        parts = [p.strip() for p in str(v).split(";")]
        parts = [p for p in parts if p and p not in size_set]
        if len(parts)==1 and "," in parts[0]:
            sub=[q.strip() for q in parts[0].split(",")]
            sub=[q for q in sub if q and q not in size_set]
            parts=sub
        return "; ".join(parts)

    seen=set(); otros2=[]
    for x in otros:
        x2=_clean_otro(x)
        if x2 and x2 not in seen:
            otros2.append(x2); seen.add(x2)
    out["atributo_otro_txt"] = "; ".join(otros2)

    # tamaño (1) por palabras MOD_TAM
    orders = []
    for term in tam_terms:
        suffix = r"(S)?" if (len(term) > 3 and not term.endswith("S")) else ""
        if re.search(r"(?<![A-ZÑ0-9])" + re.escape(term) + suffix + r"(?![A-ZÑ0-9])", f" {s} "):
            if term in tam_map:
                try:
                    orders.append(float(tam_map[term]))
                except Exception:
                    pass

    # litologías
    hits = find_all_terms(s, lit_terms)
    if not hits:
        # Fallback roca: si el texto menciona 'ROCA' o tipos de roca comunes, clasificar como roca básica
        if re.search(r"\bROCA\b", f" {s} ") or re.search(r"\b(TOBA|GRANITO|GRANODIORITA|BASALTO|ANDESITA|DIORITA|RIOLITA)\b", f" {s} "):
            T = "R"
            # estado de la roca
            if re.search(r"FRACTUR", s):
                D = "F"
            elif re.search(r"(METEORIZ|ALTERAD|DESCOMPUEST)", s):
                D = "W"
            else:
                D = "U"
            # porosidad inicial (si no se menciona, 0)
            if re.search(r"POROS", s):
                S = "P"
            elif re.search(r"(IMPERMEAB|NO\s*POROS|MASIV)", s):
                S = "N"
            else:
                S = "0"
            M = "0"
            cls_raw = sanitize_cls_code(f"{T}{D}{S}{M}")
            cls_clean = remove_repeated_materials(cls_raw)
            cls_clean = compact_cls_no_gaps(cls_clean)
            cls_clean = rescue_fines_to_matrix_with_hint(cls_clean, detect_fines_hint(s))
            cls_clean = compact_cls_no_gaps(cls_clean)
            out["cls"] = cls_clean
            out["tipo_unidad"] = "Roca"
            out["litologia_dom"] = "Roca fracturada" if D == "F" else ("Roca meteorizada" if D == "W" else "Roca sana")
            out["grupo_simplificado"] = "Roca"
            out["confiabilidad"] = "MEDIA"
            return out

        return out

    # tamaños (2) desde LITOLOGIA: TAMANO_D/S/M para litologías detectadas
    def _add_lito_size(term_key: str):
        if term_key not in lit_map:
            return
        row = lit_map[term_key]
        for k in ["TAMANO_D","TAMANO_S","TAMANO_M"]:
            lab = norm_text(row.get(k,""))
            if not lab:
                continue
            if lab in label_to_order:
                orders.append(float(label_to_order[lab]))

    seen_terms=set()
    for st,en,term in hits:
        if term in seen_terms:
            continue
        seen_terms.add(term)
        _add_lito_size(term)

    if orders:
        min_o = min(orders); max_o = max(orders)
        out["tam_min"] = order_to_label.get(min_o, str(min_o)).replace("_"," ")
        out["tam_max"] = order_to_label.get(max_o, str(max_o)).replace("_"," ")

    # construir lista de litologías encontradas (con % a la derecha)
    litos=[]
    for st,en,term in hits:
        if term in lit_map:
            litos.append({"term": term, "pos": st, "pct": pct_near(s, st, en), "codigo": str(lit_map[term].get("D","0"))})

    # fallback bolones/bloques: si aparecen en el texto pero no están en el diccionario, agregarlos como litología 'B'
    if not any(x.get("codigo") == "B" for x in litos):
        m_b = re.search(r"\b(BOLON(?:ES)?|BLOQUE(?:S)?)\b", s)
        if m_b:
            litos.append({"term": m_b.group(1), "pos": m_b.start(), "pct": pct_near(s, m_b.start(), m_b.end()), "codigo": "B"})

    if not litos:
        return out

    # regla "pasa a" / incompleto (confiabilidad)
    incomplete = s.endswith(" CON") or s.endswith(",") or s.endswith(" Y") or s.endswith(" E")
    
    # regla %:
    # - Si hay >=2 porcentajes explícitos, los porcentajes mandan (ordenar por % desc).
    # - Si hay 1 solo porcentaje, NO redefine dominancia (reglamento) -> se respeta orden de aparición.
    # - Si no hay %, izquierda->derecha.
    with_pct = [x for x in litos if x["pct"] is not None]

    if len(with_pct) >= 2:
        # ordenar: primero los que tienen %, por % desc; los que no tienen % quedan al final por posición
        l_sorted = sorted(
            litos,
            key=lambda d: (0, -d["pct"], d["pos"]) if d["pct"] is not None else (1, d["pos"])
        )
        D_term = l_sorted[0]["term"]
        S_term = l_sorted[1]["term"] if len(l_sorted) > 1 else None
        M_term = l_sorted[2]["term"] if len(l_sorted) > 2 else None
    elif len(with_pct) == 1:
        # un solo % no redefine dominancia (reglamento)
        D_term = litos[0]["term"]
        S_term = litos[1]["term"] if len(litos) > 1 else None
        M_term = litos[2]["term"] if len(litos) > 2 else None
    else:
        # izquierda->derecha
        D_term = litos[0]["term"]
        S_term = litos[1]["term"] if len(litos) > 1 else None
        M_term = litos[2]["term"] if len(litos) > 2 else None

    if incomplete:
        S_term = None
        M_term = None
        out["confiabilidad"] = "BAJA"
    else:
        out["confiabilidad"] = "ALTA" if S_term else "MEDIA"

    T = str(lit_map.get(D_term, {}).get("T","0"))
    D = str(lit_map.get(D_term, {}).get("D","0"))
    S = str(lit_map.get(S_term, {}).get("D","0")) if S_term else "0"
    M = str(lit_map.get(M_term, {}).get("D","0")) if M_term else "0"

    # Si es roca (T=R), S corresponde a porosidad (P/N) y M no aplica.
    if T == "R":
        S = str(lit_map.get(D_term, {}).get("S","0"))  # porosidad inicial desde diccionario, si existe
        M = "0"

    # tipo_roca y origen desde diccionario de litología (dominante)
    if D_term and isinstance(lit_map.get(D_term, {}), dict):
        out["tipo_roca_txt"] = str(lit_map.get(D_term, {}).get("TIPO_ROCA", "")).strip()
        # origen de litología dominante si no hay origen de modificadores
        if not out.get("origen_txt"):
            out["origen_txt"] = str(lit_map.get(D_term, {}).get("ORIGEN", "")).strip()

    # Regla finos: si (arcilla/limo) aparecen juntos sin % explícitos, se simplifica a FINOS.
    # Reglas finos vs mezcla:
    # (a) Si hay mezcla de limos+arcillas SIN % y además aparecen arenas/gravas/bolones en el texto,
    #     se fuerza D=F y se conservan hasta 2 componentes gruesos (S/M) según el orden de aparición.
    # (b) Si solo hay finos (limo+arcilla) SIN % y NO hay componentes gruesos, se colapsa a UF00.
    fine_set_mix = {"L", "C", "K"}
    coarse_set_mix = {"B", "G", "S", "A"}

    if T == "U" and len(with_pct) == 0 and (D in fine_set_mix) and (S in fine_set_mix):
        # códigos presentes en TODA la descripción
        all_codes = [(x.get("codigo","0"), x.get("pos", 10**9)) for x in litos]
        has_coarse = any(c in coarse_set_mix for c,_ in all_codes)

        if has_coarse:
            D = "F"
            # elegir componentes gruesos únicos por orden
            coarse_unique = []
            for c,pos in sorted(all_codes, key=lambda t: t[1]):
                if c in coarse_set_mix and c not in coarse_unique:
                    coarse_unique.append(c)
            S = coarse_unique[0] if len(coarse_unique) > 0 else "0"
            M = coarse_unique[1] if len(coarse_unique) > 1 else "0"
        else:
            D, S, M = "F", "0", "0"
    # Regla Suelo: solo T y D; omitir S y M
    if T == "S":
        S = S = str(lit_map.get(S_term, {}).get("D","0")) if S_term else "0"
        M = "0"

    cls_raw = sanitize_cls_code(f"{T}{D}{S}{M}")
    cls_clean = remove_repeated_materials(cls_raw)
    cls_clean = compact_cls_no_gaps(cls_clean)
    cls_clean = rescue_fines_to_matrix_with_hint(cls_clean, detect_fines_hint(s))
    cls_clean = compact_cls_no_gaps(cls_clean)
    out["cls"] = cls_clean
    # Refrescar T/D/S/M tras sanitización (para que m_tipo y grupo_simplificado usen el código final)
    parts_final = split_cls(out["cls"])
    T, D, S, M = parts_final.get("T","0"), parts_final.get("D","0"), parts_final.get("S","0"), parts_final.get("M","0")
    # m_tipo
    force = ind_ctx.get("force_m_tipo")
    if force is None and mod_ctx.get("suggest_m_tipo") == "ACC":
        force = "ACC"
    out["m_tipo"] = determinar_m_tipo(T, M, s, force=force)

    # desglose_litos con roles
    out_l=[]
    if D_term:
        out_l.append({"rol":"DOM","term":D, "codigo":D, "pct": next((x["pct"] for x in litos if x["term"]==D_term), np.nan)})
    if S_term:
        out_l.append({"rol":"SEC","term":S, "codigo":S, "pct": next((x["pct"] for x in litos if x["term"]==S_term), np.nan)})
    if M_term:
        out_l.append({"rol":"M","term":M, "codigo":M, "pct": next((x["pct"] for x in litos if x["term"]==M_term), np.nan)})
    out["desglose_litos"] = out_l

    # tipo_unidad y litologia_dom (interpretativo)
    out["tipo_unidad"] = {"R":"Roca","U":"Sedimento","S":"Suelo","0":"Sin información"}.get(T, "Sin información")
    out["litologia_dom"] = desc_codigo_por_letra(D, T, "D")

    # porcentaje_arcilla (si hay % explícitos de arcilla)
    arc = [x for x in litos if x["codigo"] in {"C","K"} and x["pct"] is not None]
    if arc:
        out["porcentaje_arcilla"] = float(arc[0]["pct"])

    # grupo_simplificado (regla final: "con finos" depende de presencia de finos en S y/o matriz efectiva)
    # - Para sedimentos con D grueso (B/G/A): "con finos" si S ∈ {L,C,F,K} o si (M ∈ {L,C,F,K} y m_tipo=MAT).
    # - Si D es fino (L/C/F/K): "Finos".
    if T == "R":
        out["grupo_simplificado"] = "Roca"
    elif T == "S":
        out["grupo_simplificado"] = "Suelo"
    elif T == "U":
        fine_set = {"L", "C", "F", "K"}  # K se considera arcilla
        m_tipo_val = (out.get("m_tipo") or "0")
        m_tipo_val = ("" if pd.isna(m_tipo_val) else str(m_tipo_val)).strip().upper()

        has_fines = (S in fine_set) or (m_tipo_val == "MAT" and M in fine_set)

        if D == "B":
            out["grupo_simplificado"] = "Bolones/Bloques con finos" if has_fines else "Bolones/Bloques limpios"
        elif D == "G":
            out["grupo_simplificado"] = "Grava con finos" if has_fines else "Grava limpia"
        elif D in {"A","S"}:
            out["grupo_simplificado"] = "Arena con finos" if has_fines else "Arena limpia"
        elif D in fine_set:
            out["grupo_simplificado"] = "Finos"
        else:
            out["grupo_simplificado"] = "Sedimento (otro)"
    else:
        out["grupo_simplificado"] = "Sin información"


    return out

# =========================
# Hoja desglose_codificacion (T/D/S/M) según template
# =========================
def build_desglose_codificacion_tsdm(df_cls: pd.DataFrame, info: pd.DataFrame, template_cols):
    df_cls = df_cls.reset_index(drop=True)
    info = info.reset_index(drop=True)
    rows=[]
    for i in range(len(df_cls)):
        id_pozo = df_cls.loc[i,"id_pozo"]
        id_intervalo = df_cls.loc[i,"id_intervalo"]
        parts = split_cls(df_cls.loc[i,"cls"])
        T_val = parts.get("T","0")
        T_ctx = T_val if T_val in {"R","U","S"} else "0"
        desc_map = __DESC_CLASE_BY_T.get(T_ctx, __DESC_CLASE_BY_T["0"])
        m_tipo = (df_cls.loc[i,"m_tipo"] if "m_tipo" in df_cls.columns else "0")
        m_tipo = ("" if pd.isna(m_tipo) else str(m_tipo)).strip().upper()
        if m_tipo not in {"MAT","ACC"}:
            m_tipo = "0"
        litos_list = info.loc[i,"desglose_litos"] if "desglose_litos" in info.columns else []

        def attrs_for_D(pct=np.nan):
            return {
                "pct_lito": pct,
                "origen": info.loc[i,"origen_txt"] if "origen_txt" in info.columns else "",
                "tipo_roca": info.loc[i,"tipo_roca_txt"] if "tipo_roca_txt" in info.columns else "",
                "tam_min": info.loc[i,"tam_min"] if "tam_min" in info.columns else np.nan,
                "tam_max": info.loc[i,"tam_max"] if "tam_max" in info.columns else np.nan,
                "seleccion": info.loc[i,"seleccion_txt"] if "seleccion_txt" in info.columns else "",
                "forma": info.loc[i,"forma_txt"] if "forma_txt" in info.columns else "",
                "plasticidad": info.loc[i,"plasticidad_txt"] if "plasticidad_txt" in info.columns else "",
                "estructuras": info.loc[i,"estructuras_txt"] if "estructuras_txt" in info.columns else "",
                "color": info.loc[i,"color_txt"] if "color_txt" in info.columns else "",
                "atributo_otro": info.loc[i,"atributo_otro_txt"] if "atributo_otro_txt" in info.columns else "",
            }
        def attrs_blank():
            return {
                "pct_lito": np.nan,
                "origen": "",
                "tipo_roca": info.loc[i,"tipo_roca_txt"] if "tipo_roca_txt" in info.columns else "",
                "tam_min": np.nan,
                "tam_max": np.nan,
                "seleccion": "",
                "forma": "",
                "plasticidad": "",
                "estructuras": "",
                "color": "",
                "atributo_otro": "",
            }

        for clase in ["T","D","S","M"]:
            # Para Suelo: solo T y D (omitir S/M)
            if T_ctx == "S" and clase in {"S","M"}:
                continue
            letra = parts.get(clase,"0")
            if clase != "T" and letra == "0":
                continue

            # pct por rol (si existe)
            pct_val = np.nan
            if isinstance(litos_list, list):
                role_need = {"D":"DOM","S":"SEC","M":"M"}.get(clase)
                if role_need:
                    for d in litos_list:
                        if isinstance(d, dict) and d.get("rol") == role_need:
                            pct_val = d.get("pct", np.nan)
                            break

            desc_clase = desc_map.get(clase, "Sin información")
            if clase == "M":
                if m_tipo == "ACC":
                    desc_clase = "Componente accesorio"
                elif m_tipo == "MAT":
                    desc_clase = "Matriz"

            row = {
                "id_pozo": id_pozo,
                "id_intervalo": id_intervalo,
                "clase_lito": clase,
                "desc_clase": desc_clase,
                "codigo_lito": letra,
                "desc_codigo": desc_codigo_por_letra(letra, T_ctx, clase),
            }
            attrs = attrs_for_D(pct_val) if clase == "D" else attrs_blank()

            # Ajuste de asignación de tamaños:
            # Si el dominante es un fino (arcilla/limo/finos) pero la descripción contiene ARENA con modificadores de tamaño,
            # los tamaños se reportan en el componente ARENA (S/M) y se vacían en D.
            texto_norm = str(df_cls.loc[i,"texto_norm"]) if "texto_norm" in df_cls.columns else ""
            if clase == "D" and letra in {"C","L","F","K"} and "ARENA" in texto_norm:
                attrs["tam_min"] = np.nan
                attrs["tam_max"] = np.nan
            if clase in {"S","M"} and letra in {"S","A"} and "ARENA" in texto_norm:
                # copiar tamaños del intervalo (heurística: los modificadores de tamaño suelen referirse a arena)
                attrs["tam_min"] = info.loc[i,"tam_min"] if "tam_min" in info.columns else np.nan
                attrs["tam_max"] = info.loc[i,"tam_max"] if "tam_max" in info.columns else np.nan

            # Ajuste de asignación de color:
            # Si el texto indica color asociado a un término específico (p.ej. "ARENA GRIS"),
            # asignar el color al componente correspondiente (arena) y no al dominante.
            color_txt = info.loc[i,"color_txt"] if "color_txt" in info.columns else ""
            color_first = ""
            if isinstance(color_txt, str) and color_txt.strip():
                color_first = color_txt.split(";")[0].strip().upper()

            if color_first and texto_norm:
                # Detectar color asociado a ARENA (último sustantivo antes del color)
                arena_color = bool(re.search(rf"\bARENA\b[^\.\n]{{0,30}}\b{re.escape(color_first)}\b", texto_norm))
                # (opcional) también "COLOR GRIS" cerca de ARENA
                arena_color = arena_color or bool(re.search(rf"\bARENA\b[^\.\n]{{0,40}}\bCOLOR\b[^\.\n]{{0,10}}\b{re.escape(color_first)}\b", texto_norm))

                if arena_color:
                    if clase == "D" and letra not in {"A","S"}:
                        attrs["color"] = ""
                    if clase in {"S","M"} and letra in {"A","S"}:
                        attrs["color"] = color_first
            # pct_lito debe reflejar el % por componente (DOM/SEC/M) cuando exista
            attrs["pct_lito"] = pct_val
            row.update(attrs)
            rows.append(row)

    out = pd.DataFrame(rows)
    for c in template_cols:
        if c not in out.columns:
            out[c] = np.nan
    return out[template_cols]

# =========================
# proporcion_sedimento
# =========================


def build_proporcion_sedimento(collar: pd.DataFrame, codificacion_estandar: pd.DataFrame, info: pd.DataFrame, template_cols):
    """
    Calcula proporciones por pozo siguiendo la planilla de referencia (Calculo_proporciones.xlsx).

    Salida mínima (siempre presente, aunque el template no lo tenga):
        id_pozo, utm_n, utm_e, prof_total, total_sed_m, n_intervalos,
        bol_bloq_m, grava_m, arena_m, limo_m, arcilla_m, finos_m, omitido_m,
        bol_bloq_pct, grava_pct, arena_pct, limo_pct, arcilla_pct, finos_pct, omitido_pct

    Reglas:
    - Solo intervalos con tipo_unidad == "SEDIMENTO" aportan a litologías.
    - Intervalos NO sedimentarios (roca/suelo/sin info) se suman completos a omitido_m.
    - NO se usan porcentajes de la descripción (solo pesos).
    - Pesos: si hay D,S,M -> 0.6/0.3/0.1; si hay solo D,S -> 0.7/0.3; si hay solo D -> 1.0.
    - total_sed_m NO incluye omitido_m.
    - % litológicos se calculan sobre total_sed_m.
    - omitido_pct se calcula sobre prof_total.
    """
    # ------------------------------------------------------------
    # SOLO POZOS CON ESTRATIGRAFÍA
    # En el Excel de entrada, el collar trae flag_estratigrafia (0/1).
    # Para la hoja proporcion_sedimento queremos SOLO los pozos con 1.
    # ------------------------------------------------------------
    if "flag_estratigrafia" in collar.columns:
        collar = collar[collar["flag_estratigrafia"].fillna(0).astype(int) == 1].copy()
    if "id_pozo" in collar.columns and "id_pozo" in codificacion_estandar.columns:
        ids_ok = set(collar["id_pozo"].astype(str).str.strip())
        codificacion_estandar = codificacion_estandar[
            codificacion_estandar["id_pozo"].astype(str).str.strip().isin(ids_ok)
        ].copy()

    # columnas mínimas requeridas en salida
    REQUIRED = [
        "id_pozo","expediente","utm_n","utm_e","prof_total","total_sed_m","n_intervalos",
        "bol_bloq_m","grava_m","arena_m","limo_m","arcilla_m","finos_m","omitido_m",
        "bol_bloq_pct","grava_pct","arena_pct","limo_pct","arcilla_pct","finos_pct","omitido_pct"
    ]

    collar0 = collar.copy()
    collar0["id_pozo"] = collar0["id_pozo"].astype(str).str.strip()

    # base de acumulación por pozo
    buckets = ["bol_bloq_m","grava_m","arena_m","limo_m","arcilla_m","finos_m","omitido_m"]
    acc = {}
    n_int = {}

    def _bucket(letter: str) -> str:
        l = ("" if pd.isna(letter) else str(letter)).strip().upper()
        if l == "B":
            return "bol_bloq_m"
        if l == "G":
            return "grava_m"
        # arena: aceptamos S o A según versión
        if l in {"A","S"}:
            return "arena_m"
        if l == "L":
            return "limo_m"
        if l in {"C","K"}:
            return "arcilla_m"
        if l == "F":
            return "finos_m"
        return "omitido_m"

    W_D, W_S, W_M = 0.6, 0.3, 0.1

    def _weights_if_missing(D, S, M):
        wD, wS, wM = W_D, W_S, W_M
        if S in ("0","",None) or pd.isna(S):
            wD += wS; wS = 0.0
        if M in ("0","",None) or pd.isna(M):
            wD += wM; wM = 0.0
        return wD, wS, wM

    # preparar dataframes
    sed_all = codificacion_estandar.copy().reset_index(drop=True)
    info_all = info.copy().reset_index(drop=True)

    sed_all["tipo_unidad"] = sed_all["tipo_unidad"].fillna("").astype(str).str.upper()

    for i in range(len(sed_all)):
        id_pozo = str(sed_all.loc[i, "id_pozo"]).strip()
        esp = pd.to_numeric(sed_all.loc[i, "espesor"], errors="coerce")
        if not np.isfinite(esp) or esp <= 0:
            continue

        acc.setdefault(id_pozo, {b: 0.0 for b in buckets})
        n_int[id_pozo] = n_int.get(id_pozo, 0) + 1

        # si NO es sedimento, todo a omitido
        if sed_all.loc[i, "tipo_unidad"] != "SEDIMENTO":
            acc[id_pozo]["omitido_m"] += float(esp)
            continue


        # --- PROPORCIONES: SOLO PESOS (NO USAR % DESCRIPCION) ---
        # Regla final (memoria):
        # - Si hay D,S,M (M != 0): D=0.6, S=0.3, M=0.1
        # - Si hay solo D,S (M == 0): D=0.7, S=0.3  (el 0.1 se suma al dominante)
        # - Si hay solo D: D=1.0
        parts = split_cls(sed_all.loc[i, "cls"])
        D, S, M = parts.get("D","0"), parts.get("S","0"), parts.get("M","0")

        if (M not in ("0","",None)) and (not pd.isna(M)):
            wD, wS, wM = 0.6, 0.3, 0.1
        else:
            # sin matriz
            wM = 0.0
            if (S not in ("0","",None)) and (not pd.isna(S)):
                wD, wS = 0.7, 0.3
            else:
                wD, wS = 1.0, 0.0

        acc[id_pozo][_bucket(D)] += float(esp) * float(wD)
        if wS > 0:
            acc[id_pozo][_bucket(S)] += float(esp) * float(wS)
        if wM > 0:
            acc[id_pozo][_bucket(M)] += float(esp) * float(wM)

    prop = pd.DataFrame([{"id_pozo": k, **v} for k, v in acc.items()]) if acc else pd.DataFrame(columns=["id_pozo"] + buckets)

    # merge con collar (para utm y prof_total)
    base_cols = [c for c in ["id_pozo","expediente","utm_n","utm_e","prof_total"] if c in collar0.columns]
    base = collar0[base_cols].copy()
    out = base.merge(prop, on="id_pozo", how="left")

    for b in buckets:
        out[b] = out[b].fillna(0.0)

    out["n_intervalos"] = out["id_pozo"].map(n_int).fillna(0).astype(int)

    # prof_total (si NaN, usar suma de espesores observados)
    # (esto evita omitido_pct raros si collar no trae prof_total)
    if "prof_total" in out.columns:
        out["prof_total"] = pd.to_numeric(out["prof_total"], errors="coerce")
    else:
        out["prof_total"] = np.nan

    # total_sed_m NO incluye omitido_m
    out["total_sed_m"] = out[["bol_bloq_m","grava_m","arena_m","limo_m","arcilla_m","finos_m"]].sum(axis=1)

    # si prof_total falta, estimar como total_sed + omitido
    out["prof_total"] = out["prof_total"].where(np.isfinite(out["prof_total"]), out["total_sed_m"] + out["omitido_m"])

    # pct litológicos sobre total_sed_m
    for base_col, pct_col in [
        ("bol_bloq_m","bol_bloq_pct"),
        ("grava_m","grava_pct"),
        ("arena_m","arena_pct"),
        ("limo_m","limo_pct"),
        ("arcilla_m","arcilla_pct"),
        ("finos_m","finos_pct"),
    ]:
        out[pct_col] = np.where(out["total_sed_m"] > 0, out[base_col] / out["total_sed_m"] * 100, np.nan)

    # omitido_pct sobre prof_total
    out["omitido_pct"] = np.where(out["prof_total"] > 0, out["omitido_m"] / out["prof_total"] * 100, np.nan)

    # asegurar columnas requeridas + extras del template
    final_cols = REQUIRED.copy()
    extra_cols = []
    if template_cols is not None:
        for c in list(template_cols):
            if c not in final_cols:
                extra_cols.append(c)
    for c in extra_cols:
        out[c] = out[c] if c in out.columns else np.nan
    final_cols = final_cols + extra_cols

    # crear vacías si faltan
    for c in final_cols:
        if c not in out.columns:
            out[c] = np.nan

    out = out[final_cols]
    out = insert_after_column(out, "expediente", "id_pozo")
    return out

# =========================
# MAIN
# =========================
def main():
    # load diccionario
    lit_terms, lit_map, ruido, tam_terms, tam_map, order_to_label, label_to_order, ind_rules, mod_rules = load_dictionary(DICT_PATH)

    # inputs (autodetección de hojas en Pozos_memoria_CG_V1.xlsx)
    xls = pd.ExcelFile(DB_PATH)

    sh_collar = next((sh for sh in xls.sheet_names if "collar" in sh.lower()), None)
    sh_strat = next((sh for sh in xls.sheet_names if "estratigraf" in sh.lower()), None)
    sh_info  = next((sh for sh in xls.sheet_names if "informacion" in sh.lower()), None)

    if sh_collar is None:
        raise ValueError("No se encontró una hoja tipo 'collar' en el Excel de entrada.")
    if sh_strat is None:
        raise ValueError("No se encontró una hoja de estratigrafía (contenga 'estratigraf') en el Excel de entrada.")

    print(f"Usando hoja collar: {sh_collar} | hoja estratigrafía: {sh_strat}" + (f" | hoja info: {sh_info}" if sh_info else ""))

    collar_raw = pd.read_excel(xls, sheet_name=sh_collar)
    strat_raw = pd.read_excel(xls, sheet_name=sh_strat)
    info_raw = pd.read_excel(xls, sheet_name=sh_info) if sh_info else None

    # expediente desde base original
    c_exp = pick_col(collar_raw, ["expediente","expedientes"])

    # collar (mapeo robusto para Pozos_memoria_CG_V1.xlsx)
    c_id = pick_col(collar_raw, ["id_pozo","id","pozo"])

    # Coordenadas: priorizar verificado, si no expediente
    c_e_ver = pick_col(collar_raw, ["este_verificado","utm_e_verificado","x_verificado","este verificado"])
    c_n_ver = pick_col(collar_raw, ["norte_verificado","utm_n_verificado","y_verificado","norte verificado"])
    c_e_exp = pick_col(collar_raw, ["este_expediente","utm_e_expediente","x_expediente","este expediente"])
    c_n_exp = pick_col(collar_raw, ["norte_expediente","utm_n_expediente","y_expediente","norte expediente"])
    c_e = c_e_ver or c_e_exp or pick_col(collar_raw, ["utm_e","este"])
    c_n = c_n_ver or c_n_exp or pick_col(collar_raw, ["utm_n","norte"])

    # Altitud/cota: priorizar verificada
    c_alt_ver = pick_col(collar_raw, ["cota_verificada","altitud_verificada","cota verificada"])
    c_alt_exp = pick_col(collar_raw, ["cota_expediente","altitud_expediente","cota expediente"])
    c_alt = c_alt_ver or c_alt_exp or pick_col(collar_raw, ["altitud_m","altitud","cota"])

    # Profundidad total y nivel estático vienen mejor desde hoja 'informacion' si existe
    c_prof = None
    c_ne = None
    if info_raw is not None:
        c_prof = pick_col(info_raw, ["profundidad_perforada","prof_total","profundidad_total","profundidad"])
        c_ne   = pick_col(info_raw, ["nivel_estatico_mbnt","nivel_estatico","ne","nivel estatico"])
    if c_prof is None:
        c_prof = pick_col(collar_raw, ["prof_total","profundidad_total","profundidad","profundidad perforada"])
    if c_ne is None:
        c_ne  = pick_col(collar_raw, ["nivel_estatico","ne"])

    c_cp  = pick_col(collar_raw, ["cota_piezometrica","cp","cota piezometrica"])

    if not (c_id and c_e and c_n):
        raise ValueError("Faltan columnas mínimas en 'collar': id/id_pozo + este/norte (verificado o expediente).")

    collar = pd.DataFrame({
        "id_pozo": collar_raw[c_id].astype(str).str.strip(),
        "expediente": collar_raw[c_exp].astype(str).str.strip() if c_exp else "",
        "utm_e": pd.to_numeric(collar_raw[c_e], errors="coerce"),
        "utm_n": pd.to_numeric(collar_raw[c_n], errors="coerce"),
        "altitud_m": pd.to_numeric(collar_raw[c_alt], errors="coerce") if c_alt else np.nan,
        "prof_total": np.nan,
        "nivel_estatico": np.nan,
        "cota_piezometrica": np.nan,
    })

    # Completar prof_total y nivel_estatico desde hoja informacion si está disponible
    if info_raw is not None and c_id:
        i_id = pick_col(info_raw, ["id_pozo","id"])
        if i_id:
            info_tmp = info_raw[[i_id] + [c_prof] if c_prof else [i_id]].copy()
            info_tmp["id_pozo"] = info_raw[i_id].astype(str).str.strip()
            if c_prof and c_prof in info_raw.columns:
                info_tmp["prof_total"] = pd.to_numeric(info_raw[c_prof], errors="coerce")
            if c_ne and c_ne in info_raw.columns:
                info_tmp["nivel_estatico"] = pd.to_numeric(info_raw[c_ne], errors="coerce")
            collar = collar.merge(info_tmp[["id_pozo","prof_total","nivel_estatico"]].drop_duplicates("id_pozo"),
                                  on="id_pozo", how="left", suffixes=("","_i"))
            collar["prof_total"] = collar["prof_total"].fillna(collar.get("prof_total_i"))
            collar["nivel_estatico"] = collar["nivel_estatico"].fillna(collar.get("nivel_estatico_i"))
            collar = collar.drop(columns=[c for c in ["prof_total_i","nivel_estatico_i"] if c in collar.columns])

    # Si aún faltan, intenta tomarlas del collar
    if c_prof and c_prof in collar_raw.columns:
        collar["prof_total"] = collar["prof_total"].fillna(pd.to_numeric(collar_raw[c_prof], errors="coerce"))
    if c_ne and c_ne in collar_raw.columns:
        collar["nivel_estatico"] = collar["nivel_estatico"].fillna(pd.to_numeric(collar_raw[c_ne], errors="coerce"))

    # cota_piezometrica: usar columna si existe, si no calcular
    if c_cp and c_cp in collar_raw.columns:
        collar["cota_piezometrica"] = pd.to_numeric(collar_raw[c_cp], errors="coerce")
    else:
        collar["cota_piezometrica"] = collar["altitud_m"] - collar["nivel_estatico"]


    # Si existe hoja 'informacion', usarla para completar campos del collar (sin sobreescribir si ya existen)
    if info_raw is not None:
        i_id = pick_col(info_raw, ["id_pozo","id","pozo"])
        if i_id:
            info_raw["_id_pozo_"] = info_raw[i_id].astype(str).str.strip()

            i_prof = pick_col(info_raw, ["prof_total","profundidad_total","profundidad","profundidad_perforada","profundidad_habilitada"])
            i_ne  = pick_col(info_raw, ["nivel_estatico","nivel_estatico_mbnt","nivel_estatico_mbsp","ne"])
            i_cp  = pick_col(info_raw, ["cota_piezometrica","cp"])
            i_alt = pick_col(info_raw, ["altitud_m","altitud","cota"])

            merge_cols = {"id_pozo": info_raw["_id_pozo_"]}
            if i_prof: merge_cols["prof_total_info"] = pd.to_numeric(info_raw[i_prof], errors="coerce")
            if i_ne:   merge_cols["nivel_estatico_info"] = pd.to_numeric(info_raw[i_ne], errors="coerce")
            if i_cp:   merge_cols["cota_piezometrica_info"] = pd.to_numeric(info_raw[i_cp], errors="coerce")
            if i_alt:  merge_cols["altitud_m_info"] = pd.to_numeric(info_raw[i_alt], errors="coerce")

            merge_df = pd.DataFrame(merge_cols).drop_duplicates("id_pozo")
            collar = collar.merge(merge_df, on="id_pozo", how="left")

            if "prof_total_info" in collar.columns:
                collar["prof_total"] = collar["prof_total"].fillna(collar["prof_total_info"])
            if "nivel_estatico_info" in collar.columns:
                collar["nivel_estatico"] = collar["nivel_estatico"].fillna(collar["nivel_estatico_info"])
            if "cota_piezometrica_info" in collar.columns:
                collar["cota_piezometrica"] = collar["cota_piezometrica"].fillna(collar["cota_piezometrica_info"])
            if "altitud_m_info" in collar.columns:
                collar["altitud_m"] = collar["altitud_m"].fillna(collar["altitud_m_info"])

            # limpiar auxiliares
            for c in ["prof_total_info","nivel_estatico_info","cota_piezometrica_info","altitud_m_info"]:
                if c in collar.columns:
                    collar = collar.drop(columns=[c])

    # estratigrafía mínimos
    s_id = pick_col(strat_raw, ["id_pozo","id"])
    s_from = pick_col(strat_raw, ["prof_desde","desde","top"])
    s_to = pick_col(strat_raw, ["prof_hasta","hasta","bottom"])
    s_txt = pick_col(strat_raw, ["texto_raw","estratigrafia","descripcion","litologia","texto"])
    if not (s_id and s_from and s_to and s_txt):
        raise ValueError("Faltan columnas mínimas en 'estratigrafia_bruta': id_pozo + desde + hasta + texto.")

    strat = pd.DataFrame({
        "id_pozo": strat_raw[s_id].astype(str).str.strip(),
        "prof_desde": pd.to_numeric(strat_raw[s_from], errors="coerce"),
        "prof_hasta": pd.to_numeric(strat_raw[s_to], errors="coerce"),
        "texto_raw": strat_raw[s_txt],
    }).dropna(subset=["id_pozo","prof_desde","prof_hasta"])
    strat = strat[strat["prof_hasta"] >= strat["prof_desde"]].copy()
    strat["espesor"] = strat["prof_hasta"] - strat["prof_desde"]
    strat["texto_norm"] = strat["texto_raw"].map(norm_text)

    # Quitar duplicados exactos de estratigrafía (evita NI_6275 duplicado por merges/lecturas)
    strat = strat.drop_duplicates(subset=["id_pozo","prof_desde","prof_hasta","texto_norm"]).copy()

    strat = strat.sort_values(["id_pozo","prof_desde","prof_hasta"]).copy()
    strat["n_intervalo"] = strat.groupby("id_pozo").cumcount()+1
    strat["id_intervalo"] = strat["id_pozo"] + "_" + strat["n_intervalo"].astype(int).astype(str).str.zfill(2)

    # prof_total si falta
    prof_from = strat.groupby("id_pozo")["prof_hasta"].max()
    collar["prof_total"] = collar["prof_total"].fillna(collar["id_pozo"].map(prof_from))

    # flag_estratigrafia
    has_strat = strat.dropna(subset=["texto_raw"]).groupby("id_pozo").size()
    collar["flag_estratigrafia"] = collar["id_pozo"].isin(has_strat.index).astype(int)

    # clasificación
    info = strat["texto_norm"].apply(
        lambda t: classify_interval(t, lit_terms, lit_map, ruido, tam_terms, tam_map, order_to_label, label_to_order, ind_rules, mod_rules)
    ).apply(pd.Series)

    codificacion_estandar = pd.concat(
        [
            strat[["id_pozo","id_intervalo","prof_desde","prof_hasta","espesor","texto_norm"]],
            info[["cls","m_tipo","porcentaje_arcilla","tipo_unidad","litologia_dom","grupo_simplificado","confiabilidad"]],
        ],
        axis=1,
    )

    # columnas template
    cols_collar = pd.read_excel(TEMPLATE_PATH, sheet_name="collar").columns
    cols_cod = pd.read_excel(TEMPLATE_PATH, sheet_name="codificacion_estandar").columns
    cols_desg = pd.read_excel(TEMPLATE_PATH, sheet_name="desglose_codificacion").columns
    cols_prop = pd.read_excel(TEMPLATE_PATH, sheet_name="proporcion_sedimento").columns
    dic_sheet = pd.read_excel(TEMPLATE_PATH, sheet_name="Diccionario")

    # alinear collar/codificacion_estandar a template (permitiendo agregar expediente aunque no venga en template)
    collar_cols = list(cols_collar)
    if "expediente" not in collar_cols and "id_pozo" in collar_cols:
        collar_cols.insert(collar_cols.index("id_pozo") + 1, "expediente")
    collar_out = pd.DataFrame({c: collar[c] if c in collar.columns else np.nan for c in collar_cols})
    collar_out = insert_after_column(collar_out, "expediente", "id_pozo")

    cod_out = pd.DataFrame({c: codificacion_estandar[c] if c in codificacion_estandar.columns else np.nan for c in cols_cod})

    desglose_out = build_desglose_codificacion_tsdm(codificacion_estandar, info, template_cols=list(cols_desg))
    # incluir solo pozos con flag_estratigrafia == 1 en desglose
    allowed_ids = set(collar.loc[collar["flag_estratigrafia"] == 1, "id_pozo"].astype(str))
    desglose_out = desglose_out[desglose_out["id_pozo"].astype(str).isin(allowed_ids)].copy()
    prop_out = build_proporcion_sedimento(collar, codificacion_estandar, info, template_cols=list(cols_prop))
    prop_out = insert_after_column(prop_out, "expediente", "id_pozo")

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        collar_out.to_excel(w, sheet_name="collar", index=False)
        cod_out.to_excel(w, sheet_name="codificacion_estandar", index=False)
        desglose_out.to_excel(w, sheet_name="desglose_codificacion", index=False)
        prop_out.to_excel(w, sheet_name="proporcion_sedimento", index=False)
        dic_sheet.to_excel(w, sheet_name="Diccionario", index=False)

    print(f"OK -> {OUTPUT_PATH}")

if __name__ == "__main__":
    main()