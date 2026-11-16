#!/usr/bin/env python3
"""
etl_v2.py — IberPhenol: Excel bruto -> SQLite v2 + CSVs + relatório de validação.

Uso (a partir da raiz do repositório):
    python etl_v2/etl_v2.py
    python etl_v2/etl_v2.py --xlsx "Base dados_CP.xlsx" --old webapp/phenoldb.sqlite --out etl_v2/out

Não altera o Excel nem o banco antigo. O modelo está em etl_v2/schema.sql.

O que muda em relação ao pipeline antigo (restructure_full.py + deduplicate_compounds.py):
  * cada pico é uma linha própria (sem apagar λmax/MS das ocorrências repetidas);
  * a linha 3 de cada aba é lida como dado, não só como metadado;
  * cada coluna de concentração vira uma condição (extrato, preparo, sub-amostra, modo de quantificação);
  * abas "Planilha" (tabelas de picos com TR), abas com o nome na coluna de concentração e
    abas de totais por classe passam a ser lidas;
  * valor, desvio, unidade, base (extrato/peso seco/peso fresco/volume) e flags (tr, nd...) são separados,
    e o texto original fica em raw_text;
  * cada linha guarda a aba e a célula de origem (src_sheet, src_cell).
"""
import argparse
import csv
import datetime
import hashlib
import json
import re
import sqlite3
import sys
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

ETL_VERSION = "0.1.0"
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

try:  # regra de produto já usada no site
    from categorize_products import classify as classify_product
except Exception:  # pragma: no cover
    classify_product = None

try:
    from rdkit import Chem, RDLogger
    from rdkit.Chem import Descriptors
    RDLogger.DisableLog("rdApp.*")
except Exception:  # pragma: no cover
    Chem = None

EVIDENCE_CODES = {
    "REP": "Reportado no artigo ou no arquivo do laboratório (ECO:0000304)",
    "CUR": "Inferido por curador (ECO:0000305)",
    "AUT": "Automático validado por regra (ECO:0000501)",
    "CALC": "Calculado (ECO:0000203)",
    "AUT-NR": "Automático não revisado (ECO:0000213)",
}

# ── utilidades de texto ──────────────────────────────────────────────────────

def cv(v):
    """Valor de célula como texto limpo (espaços colapsados, NBSP vira espaço)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).replace("\xa0", " ").split())


def a1(ci, ri):
    return f"{get_column_letter(ci + 1)}{ri + 1}"


def get(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return cv(row[idx])


def name_key(name):
    """Chave conservadora para agrupar grafias do mesmo nome (sem sinônimos)."""
    s = unicodedata.normalize("NFKC", name or "").replace("\xa0", " ")
    s = re.sub(r"[’‘′`´]", "'", s)
    s = re.sub(r"[‐‑‒–—−]", "-", s)
    s = s.lower()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\(\s*", "(", s)
    s = re.sub(r"\s*\)", ")", s)
    s = re.sub(r"\s*,\s*", ",", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip(" .;,")


TENTATIVE_RE = re.compile(
    r"hexos|pentos|deoxyhex|glucuronide(?!.*\d)|derivative|isomer|unknown|not identified|"
    r"\bn\.?i\.?\b|\bdimer|\btrimer|tetramer|acylated|\bder\.", re.I)


def is_tentative(name):
    return 1 if TENTATIVE_RE.search(name or "") else 0


def hnorm(v):
    return re.sub(r"\s+", "", cv(v).lower()).replace("(m/z)", "")


# ── cabeçalho das abas ───────────────────────────────────────────────────────

HDR = {  # chave: (linha de cabeçalho, regex sobre o texto sem espaços)
    "molecule": ("h1", r"^molecule"),
    "lambda": ("h2", r"max\(nm\)|λmax|lmax"),
    "uv": ("h2", r"espectro|uv-vis"),
    "mw": ("h1", r"^molecularweight"),
    "m_plus": ("h2", r"^\[m\]\+"),
    "m_plus_h": ("h2", r"^\[m\+h\]\+"),
    "m_minus_h": ("h2", r"^\[m-h\]-"),
    "m2_plus_h": ("h2", r"^\[2m\+h\]\+"),
    "m2_minus_h": ("h2", r"^\[2m-h\]-"),
    "m_h_2neg": ("h2", r"^\[m-h\]2-"),
    "m_h_3neg": ("h2", r"^\[m-h\]3-"),
    "exact": ("h2", r"^exact"),
    "adducts": ("h1", r"^ad?ducts"),
    "fragments": ("h1", r"^mainmsfragments"),
    "conc": ("h1", r"^concentration"),
    "equipment": ("h2", r"^equipment"),
    "data_acq": ("h2", r"^dataacquisition"),
    "ionization": ("h2", r"^ionization"),
    "cultivar": ("h1", r"^cultivar"),
    "origin": ("h1", r"origin"),
    "season": ("h1", r"^season"),
    "plant_part": ("h1", r"^plantpart"),
    "reference": ("h1", r"^reference"),
}
FIXED = {"molecule": 2, "lambda": 3, "uv": 4, "mw": 5, "m_plus": 6, "m_plus_h": 7, "m_minus_h": 8,
         "m2_plus_h": 9, "m2_minus_h": 10, "m_h_2neg": 11, "m_h_3neg": 12, "exact": 13,
         "adducts": 14, "fragments": 15, "conc": 16}
META_OFFSET = {"data_acq": 1, "ionization": 2, "cultivar": 3, "origin": 4, "season": 5,
               "plant_part": 6, "reference": 7}
META_KEYS = ["equipment", "data_acq", "ionization", "cultivar", "origin", "season", "plant_part", "reference"]
SAMPLE_KEYS = ["cultivar", "origin", "season", "plant_part"]
QUANT_BLOCKS = ["Reference standards (DAD)", "Reference standards (MS)", "Internal standard",
                "Extinction", "Peak area", "Concentration"]


def find_cols(h1, h2):
    cols = {}
    for key, (which, pat) in HDR.items():
        hdr = h1 if which == "h1" else h2
        for i, v in enumerate(hdr):
            if v is not None and re.search(pat, hnorm(v)):
                cols[key] = i
                break
    for key, idx in FIXED.items():
        cols.setdefault(key, idx)
    # os valores começam logo após os fragmentos MS, mesmo quando o rótulo "Concentrations"
    # da linha 1 está mais à direita (ex.: aba bee bread)
    cols["conc"] = cols["fragments"] + 1
    if "equipment" in cols:
        for key, off in META_OFFSET.items():
            cols.setdefault(key, cols["equipment"] + off)
    return cols


def split_block(label):
    """'Reference standards (DAD) mg/g of extract' -> ('Reference standards (DAD)', 'mg/g of extract')."""
    lab = cv(label)
    for b in QUANT_BLOCKS:
        if lab.lower().startswith(b.lower()):
            return b, lab[len(b):].strip(" .;,:-")
    return lab, ""


# ── leitura das células de concentração ──────────────────────────────────────

FLAG_RE = re.compile(r"^(not\s+detected|n\.?\s?d\.?|n/d|tr\.?|traces?|n\.?\s?q\.?|<\s*lo[qd]|-+|–|—)(?=\s|$|[^a-z])", re.I)
UNIT_RE = re.compile(
    r"(?P<num>\d+(?:[.,]\d+)?)\s*(?:(?:±|\+/-)\s*(?P<sd>\d+(?:[.,]\d+)?)\s*)?"
    r"(?P<unit>(?:mg|µg|μg|ug|ng|g)\s*/\s*(?:100\s*g|100\s*ml|kg|g|ml|l)(?![a-z])"
    r"|%\s*(?:of\s*)?peak\s*area|peak\s*area|%)", re.I)
UNIT_TOKEN_RE = re.compile(r"(?:mg|µg|μg|ug|ng|g)\s*/\s*(?:100\s*)?(?:g|kg|ml|l)(?![a-z])", re.I)
FLAG_ANY_RE = re.compile(r"(?:^|\s)(n\.d\.?|nd|n/d|tr|n\.q\.?|nq)(?=\s|$|[,;)])", re.I)
RANGE_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*[-–]\s*(\d+(?:[.,]\d+)?)\s*(?:mg|µg|μg|ug|g|%)", re.I)
KGY_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*kgy", re.I)
TAG_RE = re.compile(r"\(([a-z0-9]{1,4})\)|(?:^|\s)([a-h])\)", re.I)
SAMPLE_TAG_RE = re.compile(r"\bsample\s*(\d+)\b", re.I)
PREP_WORDS = ["decoction", "infusion", "hydromethanolic", "hydroethanolic", "hydroalcoholic",
              "methanolic", "ethanolic", "aqueous", "lyophilized", "juice"]
PREP_ALIASES = {"lyofilized": "lyophilized", "meoh": "methanolic", "etoh": "ethanolic",
                "metanhol": "methanolic", "methanol": "methanolic", "ethanol": "ethanolic",
                "hydro-methanolic": "hydromethanolic", "hydro-ethanolic": "hydroethanolic"}
EXTRACT_RE = re.compile(r"extract|infusion|decoction|lyophili|lyofili|hydromethanol|hydroethanol|"
                        r"hydroalcohol|methanol|ethanol|meoh|etoh|aqueous|metanhol", re.I)


def num(s):
    return float(s.replace(",", ".")) if s else None


def canon_unit(u):
    u = u.lower().replace("μ", "u").replace("µ", "u")
    u = re.sub(r"\s+", "", u)
    if "peakarea" in u:
        return "%" if u.startswith("%") else "area"
    return {"mg/100g": "mg/100 g", "g/100g": "g/100 g", "ug/100g": "ug/100 g", "ug/ml": "ug/mL",
            "mg/ml": "mg/mL", "mg/l": "mg/L", "ug/l": "ug/L", "ng/ml": "ng/mL", "mg/100ml": "mg/100 mL",
            "ug/100ml": "ug/100 mL"}.get(u, u)


VOLUME_UNITS = ("ug/mL", "mg/mL", "mg/L", "ug/L", "ng/mL", "mg/100 mL", "ug/100 mL")


def preparations(text):
    tl = text.lower()
    found = set()
    for w in PREP_WORDS:
        if re.search(rf"\b{w}\b", tl):
            found.add(w)
    for alias, w in PREP_ALIASES.items():
        if re.search(rf"\b{re.escape(alias)}\b", tl):
            found.add(w)
    # "hydromethanolic" contém "methanolic": fica só o mais específico
    for spec, gen in (("hydromethanolic", "methanolic"), ("hydroethanolic", "ethanolic")):
        if spec in found:
            found.discard(gen)
    return "+".join(sorted(found)) or None


def parse_value(raw):
    """Texto da célula -> leitura normalizada; nada é descartado (raw_text guarda o original)."""
    out = {"value": None, "sd": None, "is_range": 0, "flag": None, "unit_verbatim": None,
           "unit": None, "basis": None, "preparation": None, "tag": None,
           "irradiation_kgy": None, "parse_status": None}
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        out.update(value=float(raw), basis="unspecified", parse_status="no_unit")
        return out
    t = cv(raw)
    tl = t.lower()
    m = KGY_RE.search(t)
    if m:
        out["irradiation_kgy"] = num(m.group(1))
        tl_wo = (t[:m.start()] + " " + t[m.end():]).strip()
    else:
        tl_wo = t
    fm = FLAG_RE.match(tl_wo.strip())
    if not fm and not UNIT_RE.search(tl_wo):
        fm = FLAG_ANY_RE.search(tl_wo)  # "hydromethanolic extract nd mg/g"
    if fm:
        tok = fm.group(1).lower().replace(" ", "")
        out["flag"] = ("nd" if tok.startswith("n") and "d" in tok or tok == "n/d"
                       else "tr" if tok.startswith("tr")
                       else "nq" if tok.startswith("n")
                       else "lt_loq" if tok.startswith("<")
                       else "dash")
        out["parse_status"] = "flag"
    rm = RANGE_RE.search(tl_wo)
    um = UNIT_RE.search(tl_wo)
    if um and not out["flag"]:
        out["unit_verbatim"] = um.group("unit")
        out["unit"] = canon_unit(um.group("unit"))
        out["sd"] = num(um.group("sd"))
        if rm:
            out["is_range"] = 1
            out["parse_status"] = "range"
        else:
            out["value"] = num(um.group("num"))
            out["parse_status"] = "ok"
    elif not out["flag"]:
        nm = re.search(r"\d+(?:[.,]\d+)?", tl_wo)
        ut = UNIT_TOKEN_RE.search(tl_wo)
        if nm and ut:  # unidade separada do número: "8,42 malvidin-3-glucoside (mg/ g of skin)"
            out["value"] = num(nm.group())
            out["unit_verbatim"] = ut.group()
            out["unit"] = canon_unit(ut.group())
            out["parse_status"] = "ok_unit_detached"
        elif nm:
            out["value"] = num(nm.group())
            out["parse_status"] = "no_unit"
        else:
            out["parse_status"] = "unparsed"
    elif um:  # flag com unidade escrita ("tr mg/ g of infusion")
        out["unit_verbatim"] = um.group("unit")
        out["unit"] = canon_unit(um.group("unit"))
    # base de cálculo
    if re.search(r"\bdw\b|dry\s*weight|dried\s*weight|\bd\.w\.", tl):
        out["basis"] = "dw"
    elif re.search(r"\bfw\b|fresh\s*weight|\bf\.w\.", tl):
        out["basis"] = "fw"
    elif out["unit"] in VOLUME_UNITS:
        out["basis"] = "volume"
    elif out["unit"] in ("%", "area") and "peak" in tl:
        out["basis"] = "peak_area"
    elif EXTRACT_RE.search(tl):
        out["basis"] = "extract"
    else:
        out["basis"] = "unspecified"
    out["preparation"] = preparations(t)
    tags = [a or b for a, b in TAG_RE.findall(t)]
    tags = [x for x in tags if x and x.lower() not in ("meoh", "etoh", "herb")]
    st = SAMPLE_TAG_RE.search(t)
    if st:
        tags.append(f"sample {st.group(1)}")
    out["tag"] = tags[-1] if tags else None
    return out


def name_like(text):
    """Texto que é um nome (de composto ou padrão), não um valor."""
    if not text or FLAG_RE.match(text.strip()):
        return False
    if UNIT_RE.search(text) or KGY_RE.search(text) or UNIT_TOKEN_RE.search(text) or FLAG_ANY_RE.search(text):
        return False
    return bool(re.search(r"[A-Za-z]{3,}", text))


def parse_ms2(text):
    pairs = re.findall(r"(\d+(?:\.\d+)?)\s*\(\s*(\d+(?:\.\d+)?)\s*\)", text or "")
    return [[float(a), float(b)] for a, b in pairs]


def parse_lambda(text):
    return [float(x) for x in re.findall(r"(?<![\d.])(\d{3})(?:\.\d+)?(?![\d])", text or "")]


def first_num(text):
    m = re.search(r"\d+(?:\.\d+)?", text or "")
    return float(m.group()) if m else None


# ── referências, táxons, letras de sub-amostra ───────────────────────────────

CITATION_START_RE = re.compile(r"^[A-ZÀ-Ý][\w'’\-]+(?:\s[A-ZÀ-Ý]\.)*(?:,|\s+et\s+al|\s+and\s|\s+&\s)")


def split_citations(raw):
    """Uma célula pode ter várias citações, e uma citação pode ocupar várias linhas:
    uma linha só abre citação nova se começar como 'Sobrenome et al.' / 'Sobrenome, ...' e tiver ano."""
    if raw is None:
        return []
    merged = []
    for line in (cv(x) for x in str(raw).replace("\xa0", " ").split("\n")):
        if not line:
            continue
        starts = CITATION_START_RE.match(line) and re.search(r"(19|20)\d{2}", line)
        if merged and not starts:
            merged[-1] += " " + line
        else:
            merged.append(line)
    out = []
    for m in merged:
        out += re.split(r";\s*(?=[A-ZÀ-Ý][\w'’\-]+(?:,|\s+et\s+al))", m)
    return [cv(p) for p in out if cv(p)]


def is_citation(t):
    if re.search(r"et al|\bdoi\b|10\.\d{4,9}/", t, re.I):
        return 1
    return 1 if (len(t) > 25 and re.search(r"(19|20)\d{2}", t) and re.match(r"^[A-ZÀ-Ý]", t)) else 0


def citation_year(t):
    m = re.search(r"\b(19[5-9]\d|20[0-4]\d)\b", t)
    return int(m.group(1)) if m else None


def citation_doi(t):
    m = re.search(r"10\.\d{4,9}/[^\s,;]+", t)
    return m.group().rstrip(".") if m else None


COMMON_NAMES = {  # nome comum na aba -> leitura (sempre marcada para revisão)
    "cereja": ("Prunus avium", "Prunus", "avium"),
    "mirtilos": ("Vaccinium", "Vaccinium", None),
    "amendoa": ("Prunus dulcis", "Prunus", "dulcis"),
    "wine": ("Vitis vinifera", "Vitis", "vinifera"),
    "bee bread": (None, None, None),
}
QUALIFIER_RE = re.compile(
    r"\b(seeds?|roots?|stems?|grapes?|fruits?|leaves|flowers?|peel|veget\.? parts|vegetative parts|"
    r"in vitro|clone [a-z]|cultivated|wilde?|commerc\w*|herb|bulbs?|petals?)\b", re.I)


def parse_taxon(sheet_name):
    s = cv(sheet_name).replace("_", " ")
    copy = re.search(r"\s*\((\d+)\)$|\s+(\d+)$", s)
    qual = []
    if copy:
        qual.append(f"cópia {copy.group(1) or copy.group(2)}")
        s = s[:copy.start()].strip()
    for m in QUALIFIER_RE.finditer(s):
        qual.append(m.group(1).lower())
    low = s.lower()
    if low in COMMON_NAMES:
        b, g, e = COMMON_NAMES[low]
        return {"binomial": b, "genus": g, "epithet": e, "is_common_name": 1, "needs_review": 1,
                "note": f"nome comum na aba: {s}", "qualifier": "; ".join(qual) or None}
    tokens = s.split()
    genus = epithet = None
    review = 0
    note = None
    if tokens and re.fullmatch(r"[A-Z][a-zë]+", tokens[0]):
        genus = tokens[0]
        rest = tokens[1:]
        if rest and rest[0] in ("x", "×") and len(rest) > 1:
            epithet = f"× {rest[1].lower()}"
        elif rest and re.fullmatch(r"[a-zë-]+", rest[0]) and rest[0] not in ("sp.", "spp."):
            epithet = rest[0]
        elif rest and rest[0] in ("sp.", "spp."):
            review, note = 1, "só gênero"
        else:
            review, note = 1, "epíteto não encontrado"
    elif tokens and re.fullmatch(r"[A-Z]\.", tokens[0]):
        review, note = 1, "gênero abreviado"
        epithet = tokens[1].lower() if len(tokens) > 1 else None
        genus = tokens[0]
    else:
        review, note = 1, "nome não reconhecido"
    binomial = f"{genus} {epithet}" if genus and epithet else genus
    return {"binomial": binomial, "genus": genus, "epithet": epithet, "is_common_name": 0,
            "needs_review": review, "note": note, "qualifier": "; ".join(qual) or None}


def letter_map(text):
    """'(a) judia and (b) longal' / 'a) Amarelo; b) batateiro' -> {'a': 'judia', ...}."""
    out = {}
    for m in re.finditer(r"\(?\b([a-h])\)\s*([^;(),]+?)(?=\s*(?:;|,|\(?\b[a-h]\)|\band\b|$))", text or ""):
        out[m.group(1).lower()] = m.group(2).strip(" .")
    return out


def code_map(text):
    """'Montesinho (BB2), Rio de Onor (BB3)' -> {'BB2': 'Montesinho', ...}."""
    return {m.group(2): re.sub(r"^(and|e)\s+", "", m.group(1).strip(" ,;"))
            for m in re.finditer(r"([^,;()]+?)\s*\(([A-Z]{1,3}\d{0,2})\)", text or "")}


# ── ETL ──────────────────────────────────────────────────────────────────────

class ETL:
    def __init__(self, xlsx, old_db, out_dir):
        self.xlsx, self.old_db, self.out = Path(xlsx), Path(old_db), Path(out_dir)
        self.sheets, self.studies, self.taxa, self.methods = [], {}, {}, {}
        self.samples, self.sample_study, self.conditions = [], set(), {}
        self.peaks, self.measurements, self.issues = [], [], []
        self.compounds = {}  # name_key -> dict

    # registros ------------------------------------------------------------
    def issue(self, severity, code, message, entity=None, entity_id=None, sheet=None, cell=None):
        self.issues.append({"severity": severity, "code": code, "entity": entity,
                            "entity_id": None if entity_id is None else str(entity_id),
                            "message": message, "src_sheet": sheet, "src_cell": cell})

    def study_id(self, citation, sheet, cell):
        if citation not in self.studies:
            self.studies[citation] = {"study_id": len(self.studies) + 1, "citation": citation,
                                      "is_citation": is_citation(citation),
                                      "year": citation_year(citation), "doi": citation_doi(citation),
                                      "src_sheet": sheet, "src_cell": cell}
            if not self.studies[citation]["is_citation"]:
                self.issue("warning", "reference_not_citation",
                           f"Texto na coluna Reference não parece citação: {citation[:80]}",
                           "study", self.studies[citation]["study_id"], sheet, cell)
        return self.studies[citation]["study_id"]

    def taxon_id(self, tx):
        key = tx["binomial"] or f"__{tx['note']}"
        if key not in self.taxa:
            self.taxa[key] = {"taxon_id": len(self.taxa) + 1, "binomial": tx["binomial"],
                              "genus": tx["genus"], "epithet": tx["epithet"],
                              "is_common_name": tx["is_common_name"],
                              "needs_review": tx["needs_review"], "note": tx["note"]}
        return self.taxa[key]["taxon_id"]

    def method_id(self, equipment, acq, ion):
        if not (equipment or acq or ion):
            return None
        key = (equipment or None, acq or None, ion or None)
        if key not in self.methods:
            self.methods[key] = {"method_id": len(self.methods) + 1, "equipment": key[0],
                                 "data_acquisition": key[1], "ionization": key[2]}
        return self.methods[key]["method_id"]

    def compound_for(self, name):
        k = name_key(name)
        if k not in self.compounds:
            self.compounds[k] = {"compound_id": f"IPH-C{len(self.compounds) + 1:05d}", "name_key": k,
                                 "names": Counter(), "classes": Counter(), "subclasses": Counter(),
                                 "samples": set(), "n_peaks": 0}
        c = self.compounds[k]
        c["names"][cv(name)] += 1
        return c

    # leitura ----------------------------------------------------------------
    def run(self):
        wb = openpyxl.load_workbook(self.xlsx, data_only=True, read_only=True)
        all_rows = [(name, [tuple(c.value for c in r) for r in wb[name].iter_rows()]) for name in wb.sheetnames]
        wb.close()
        # abas com conteúdo idêntico: fica a que não tem marcador de cópia no nome ("(2)", " 2", "_2")
        groups = defaultdict(list)
        for i, (name, rows) in enumerate(all_rows):
            sig = hashlib.sha1(json.dumps([[cv(v) for v in r] for r in rows]).encode()).hexdigest()
            groups[sig].append(i)
        canonical = {}
        for idxs in groups.values():
            keep = min(idxs, key=lambda i: (bool(re.search(r"\(\d+\)\s*$|[\s_]\d+\s*$", all_rows[i][0])), i))
            for i in idxs:
                canonical[i] = keep
        peak_tables = []
        for i, (name, rows) in enumerate(all_rows):
            sid = i + 1
            sh = {"sheet_id": sid, "name_verbatim": name, "status": None, "layout": None,
                  "duplicate_of": None, "metadata_mode": None, "n_peaks": 0, "n_measurements": 0,
                  "suggested_sheet": None, "suggestion_score": None, "notes": None}
            self.sheets.append(sh)
            h1 = rows[0] if rows else ()
            if canonical[i] != i:
                keep = canonical[i]
                sh.update(status="duplicate", duplicate_of=keep + 1, notes="conteúdo idêntico a outra aba")
                self.issue("info", "sheet_duplicate",
                           f"Aba idêntica a '{all_rows[keep][0]}'; ignorada", "sheet", sid, name)
                continue
            h1n = [hnorm(v) for v in h1]
            if any(x.startswith(("tentativeidentification", "identification")) for x in h1n) \
                    and any(x.startswith("rt") for x in h1n):
                peak_tables.append((sh, rows, False))
                continue
            if name.lower().startswith("planilha") and len(h1) > 4 and isinstance(h1[0], (int, float)) \
                    and re.search(r"\d+\s*\(\d+", cv(h1[3])) and name_like(cv(h1[4])):
                peak_tables.append((sh, rows, True))  # tabela de picos sem cabeçalho
                continue
            h2 = rows[1] if len(rows) > 1 else ()
            if len(rows) < 3 or not any(hnorm(v).startswith("equipment") for v in h2):
                sh.update(status="skipped", notes="sem cabeçalho reconhecível")
                self.issue("info", "sheet_skipped", "Aba sem o cabeçalho padrão", "sheet", sid, name)
                continue
            self.parse_standard(sh, rows, h1, h2)
        for sh, rows, headerless in peak_tables:
            self.parse_peak_table(sh, rows, headerless)
        self.finish_conditions()
        self.finish_compounds()

    def parse_standard(self, sh, rows, h1, h2):
        name, sid = sh["name_verbatim"], sh["sheet_id"]
        col = find_cols(h1, h2)
        conc_cols = list(range(col["conc"], col["equipment"]))
        blocks, current = {}, ""
        for c in conc_cols:
            lab = cv(h2[c]) if c < len(h2) else ""
            if lab:
                current = lab
            blocks[c] = current
        n = len(rows)

        def has_values(row):
            return any(get(row, c) and not name_like(get(row, c)) for c in conc_cols)

        # modo dos metadados: por aba ou por linha
        meta_rows = [ri for ri in range(2, n) if any(get(rows[ri], col[k]) for k in SAMPLE_KEYS)]
        data_meta_rows = [ri for ri in meta_rows
                          if get(rows[ri], col["molecule"]) or has_values(rows[ri])
                          or any(name_like(get(rows[ri], c)) for c in conc_cols[:1])]
        mode = "row" if len(data_meta_rows) >= 2 else "sheet"
        sh["metadata_mode"] = mode
        tx = parse_taxon(name)
        cult_text = " ".join(get(rows[ri], col["cultivar"]) for ri in range(2, n))
        orig_text = " ".join(get(rows[ri], col["origin"]) for ri in range(2, n))
        letters, codes = letter_map(cult_text), code_map(orig_text)

        # linha 3 sem molécula e sem espectro = rótulos das colunas (ex.: BB1..BBC)
        row3_labels = {}
        r3 = rows[2]
        if not get(r3, col["molecule"]) and not any(get(r3, col[k]) for k in ("lambda", "m_minus_h", "fragments")):
            for c in conc_cols:
                if get(r3, c):
                    row3_labels[c] = get(r3, c)
        self._sheet_ctx = {"sheet_id": sid, "blocks": blocks, "row3": row3_labels, "h1": h1, "h2": h2,
                           "letters": letters, "codes": codes}

        # amostra(s)
        sample_by_ctx = {}

        def make_sample(ctx, inferred, cell):
            key = tuple(ctx.get(k) or "" for k in SAMPLE_KEYS)
            if key in sample_by_ctx:
                return sample_by_ctx[key]
            part = ctx.get("plant_part") or None
            prod = None
            if classify_product:
                prod = classify_product(tx["binomial"] or name, part or "")
            smp = {"sample_id": len(self.samples) + 1, "sheet_id": sid, "taxon_id": self.taxon_id(tx),
                   "name_verbatim": name, "qualifier": tx["qualifier"], "plant_part": part,
                   "cultivar": ctx.get("cultivar") or None, "origin": ctx.get("origin") or None,
                   "season": ctx.get("season") or None, "product_category": prod,
                   "method_id": self.method_id(ctx.get("equipment"), ctx.get("data_acq"), ctx.get("ionization")),
                   "context_inferred": 1 if inferred else 0, "src_sheet": name, "src_cell": cell}
            self.samples.append(smp)
            sample_by_ctx[key] = smp["sample_id"]
            return smp["sample_id"]

        sheet_studies, sheet_ctx, sheet_sample = [], {}, None
        if mode == "sheet":
            for k in META_KEYS:
                vals = []
                for ri in range(2, n):
                    v = get(rows[ri], col[k])
                    if v and v not in vals:
                        vals.append(v)
                sheet_ctx[k] = " | ".join(vals) if k != "reference" else None
            for ri in range(2, n):
                for cit in split_citations(rows[ri][col["reference"]] if col["reference"] < len(rows[ri]) else None):
                    st = self.study_id(cit, name, a1(col["reference"], ri))
                    if self.studies[cit]["is_citation"] and st not in sheet_studies:
                        sheet_studies.append(st)

        cur_cls = cur_sub = ""
        ctx, inferred_keys = {}, set()
        cur_study = None
        last_mol = None
        n_peaks = n_meas = 0
        layouts = Counter()
        for ri in range(2, n):
            row = rows[ri]
            cls, sub, mol = get(row, 0), get(row, 1), get(row, col["molecule"])
            if cls:
                cur_cls, cur_sub = cls, sub
            elif sub:
                cur_sub = sub
            if mode == "row":
                for k in META_KEYS:
                    v = get(row, col[k])
                    if v:
                        if k == "reference":
                            cits = [c for c in split_citations(row[col[k]]) if c]
                            ids = [self.study_id(c, name, a1(col[k], ri)) for c in cits]
                            good = [i for c, i in zip(cits, ids) if self.studies[c]["is_citation"]]
                            cur_study = good[0] if good else cur_study
                        else:
                            ctx[k] = v
                            inferred_keys.discard(k)
                    elif ctx.get(k):
                        inferred_keys.add(k)
            if ri == 2 and row3_labels:
                continue
            cells = [(c, row[c]) for c in conc_cols if c < len(row) and cv(row[c])]
            texts = [(c, cv(v)) for c, v in cells if isinstance(v, str) and name_like(cv(v))]
            values = [(c, v) for c, v in cells if not (isinstance(v, str) and name_like(cv(v)))]
            name_src = "molecule_col" if mol else None
            notes = []
            if not mol and texts and texts[0][0] == conc_cols[0]:
                mol, name_src = texts[0][1], "conc_col_text"
                texts = texts[1:]
            for c, t in texts:
                notes.append(f"{split_block(blocks.get(c, ''))[0] or 'coluna ' + get_column_letter(c + 1)}: {t}")
            if not mol and not values and mode == "row" and last_mol and not cls and not sub and (
                    any(get(row, col[k]) for k in ("lambda", "uv", "m_minus_h", "fragments"))
                    or any(get(row, col[k]) for k in SAMPLE_KEYS + ["reference"])):
                # linha de continuação: mesmo composto da linha anterior, outra fonte ou outro contexto
                mol, name_src = last_mol, "continuation_row"
            if not mol and not values:
                continue
            if mol:
                last_mol = mol
            if mode == "sheet" and sheet_sample is None:  # amostra só nasce quando há dado
                sheet_sample = make_sample(sheet_ctx, False, a1(col["cultivar"], 2))
                for st in sheet_studies:
                    self.sample_study.add((sheet_sample, st))
                if len(sheet_studies) > 1:
                    self.issue("warning", "measurement_multi_study",
                               f"{len(sheet_studies)} estudos na aba; medições sem estudo individual",
                               "sheet", sid, name)
            sample_id = make_sample(ctx, bool(inferred_keys & set(SAMPLE_KEYS)), a1(col["cultivar"], ri)) \
                if mode == "row" else sheet_sample
            if mode == "row" and cur_study:
                self.sample_study.add((sample_id, cur_study))
            study = cur_study if mode == "row" else (sheet_studies[0] if len(sheet_studies) == 1 else None)
            peak_id = None
            if mol:
                comp = self.compound_for(mol)
                comp["classes"][cur_cls] += 1
                comp["subclasses"][cur_sub] += 1
                comp["samples"].add(sample_id)
                comp["n_peaks"] += 1
                lam, mh, frag = get(row, col["lambda"]), get(row, col["m_minus_h"]), get(row, col["fragments"])
                peak_id = len(self.peaks) + 1
                self.peaks.append({
                    "peak_id": peak_id, "sheet_id": sid, "sample_id": sample_id,
                    "compound_id": comp["compound_id"], "name_verbatim": mol, "name_source": name_src,
                    "class_verbatim": cur_cls or None, "subclass_verbatim": cur_sub or None,
                    "class_inherited": 0 if cls else 1, "peak_number": None, "rt_min": None,
                    "lambda_max_text": lam or None, "lambda_max_json": json.dumps(parse_lambda(lam)) if lam else None,
                    "uv_spectrum": get(row, col["uv"]) or None, "mw": get(row, col["mw"]) or None,
                    "m_plus": get(row, col["m_plus"]) or None, "m_plus_h": get(row, col["m_plus_h"]) or None,
                    "m_minus_h": mh or None, "m_minus_h_num": first_num(mh),
                    "m2_plus_h": get(row, col["m2_plus_h"]) or None, "m2_minus_h": get(row, col["m2_minus_h"]) or None,
                    "m_h_2neg": get(row, col["m_h_2neg"]) or None, "m_h_3neg": get(row, col["m_h_3neg"]) or None,
                    "exact_mass": get(row, col["exact"]) or None, "adducts": get(row, col["adducts"]) or None,
                    "ms2_text": frag or None, "ms2_json": json.dumps(parse_ms2(frag)) if frag else None,
                    "notes": "; ".join(notes) or None,
                    "evidence": "REP" if name_src == "molecule_col" else "CUR",
                    "src_sheet": name, "src_cell": a1(conc_cols[0] if name_src == "conc_col_text" else col["molecule"], ri)})
                n_peaks += 1
                layouts["name_in_conc_col" if name_src == "conc_col_text" else "standard"] += 1
            target_class = None
            if not mol:
                target_class = cur_sub or cur_cls or None
                layouts["class_totals"] += 1
                if not target_class:
                    self.issue("warning", "orphan_values", "Valores sem nome de composto nem classe",
                               "sheet", sid, name, a1(values[0][0], ri))
            for c, v in values:
                pv = parse_value(v)
                cond = self.condition_for(sid, c)
                cinfo = self.conditions[(sid, c)]
                evidence = "REP" if peak_id else "CUR"
                if pv["unit"] is None and pv["value"] is not None:
                    if cinfo["hdr_unit"]:  # unidade escrita no cabeçalho da coluna
                        pv.update(unit=cinfo["hdr_unit"], parse_status="ok_unit_from_header")
                        if pv["basis"] == "unspecified" and cinfo["hdr_basis"]:
                            pv["basis"] = cinfo["hdr_basis"]
                        evidence = "CUR"
                    elif (cinfo["quant_approach"] or "").startswith("Peak area"):
                        pv.update(unit="area", basis="peak_area", parse_status="ok_unit_from_header")
                        evidence = "CUR"
                self.measurements.append({
                    "measurement_id": len(self.measurements) + 1, "peak_id": peak_id, "sample_id": sample_id,
                    "condition_id": cond, "study_id": study, "target_class": target_class,
                    "raw_text": cv(v), **pv,
                    "evidence": evidence, "src_sheet": name, "src_cell": a1(c, ri)})
                n_meas += 1
                if pv["parse_status"] in ("no_unit", "unparsed", "range"):
                    self.issue("warning" if pv["parse_status"] != "range" else "info",
                               f"value_{pv['parse_status']}", f"Valor '{cv(v)[:60]}'",
                               "measurement", len(self.measurements), name, a1(c, ri))
        sh["n_peaks"], sh["n_measurements"] = n_peaks, n_meas
        sh["layout"] = layouts.most_common(1)[0][0] if layouts else None
        if n_peaks == 0 and n_meas == 0:
            sh["status"] = "template_empty"
            self.issue("info", "sheet_template_empty", "Aba só com o modelo, sem dados", "sheet", sid, name)
            # descarta a amostra criada para a aba vazia
            self.samples = [s for s in self.samples if s["sheet_id"] != sid]
            self.sample_study = {(a, b) for a, b in self.sample_study if a in {s["sample_id"] for s in self.samples}}
        else:
            sh["status"] = "parsed"
            if layouts.get("name_in_conc_col"):
                self.issue("info", "name_from_conc_col",
                           f"{layouts['name_in_conc_col']} picos com o nome na coluna de concentração",
                           "sheet", sid, name)
            if layouts.get("class_totals"):
                self.issue("info", "class_totals",
                           f"{layouts['class_totals']} linhas de valores por classe, sem composto",
                           "sheet", sid, name)

    def condition_for(self, sid, c):
        key = (sid, c)
        if key not in self.conditions:
            ctx = self._sheet_ctx
            block, extra = split_block(ctx["blocks"].get(c, ""))
            h1extra = cv(ctx["h1"][c]) if c < len(ctx["h1"]) else ""
            if h1extra.lower().startswith("concentration"):
                h1extra = ""
            header_label = "; ".join(x for x in (h1extra, extra) if x) or None
            hdr = parse_value("1 " + header_label) if header_label and UNIT_RE.search("1 " + header_label) else None
            self.conditions[key] = {
                "hdr_unit": hdr["unit"] if hdr else None,
                "hdr_basis": hdr["basis"] if hdr and hdr["basis"] != "unspecified" else None,
                "condition_id": len(self.conditions) + 1, "sheet_id": sid, "col_letter": get_column_letter(c + 1),
                "quant_approach": block or None,
                "header_label": "; ".join(x for x in (h1extra, extra) if x) or None,
                "row3_label": ctx["row3"].get(c), "letters": ctx["letters"], "codes": ctx["codes"],
                "preparation": None, "tag": None, "tag_resolved": None, "label": None}
        return self.conditions[key]["condition_id"]

    def parse_peak_table(self, sh, rows, headerless=False):
        """Abas 'Planilha': tabela de picos (nº, TR, λmax, [M-H]-, MS2, identificação), sem contexto."""
        name, sid = sh["name_verbatim"], sh["sheet_id"]
        if headerless:  # TR, λmax, [M-H]-, MS2, identificação a partir da 1ª linha
            c_rt, c_lam, c_mh, c_ms2, c_id, c_cls, start = 0, 1, 2, 3, 4, None, 0
        else:
            idx = {hnorm(v): i for i, v in enumerate(rows[0]) if v is not None}

            def find(*prefixes):
                for p in prefixes:
                    hit = next((i for k, i in idx.items() if k.startswith(p)), None)
                    if hit is not None:
                        return hit
                return None
            c_rt, c_lam = find("rt"), find("lmax", "λmax", "máx", "max")
            c_mh, c_ms2 = find("molecularion", "[m-h]"), find("ms2")
            c_id, c_cls, start = find("tentativeidentification", "identification"), find("class"), 1
        c_pk = None if headerless else (find("peak") if find("peak") is not None else (0 if c_rt != 0 else None))
        keys, n_peaks = set(), 0
        for ri in range(start, len(rows)):
            row = rows[ri]
            ident = get(row, c_id)
            rt, lam, mh, frag = get(row, c_rt), get(row, c_lam), get(row, c_mh), get(row, c_ms2)
            if not (ident or rt or mh):
                continue
            comp = None
            if ident:
                comp = self.compound_for(ident)
                comp["n_peaks"] += 1
                keys.add(comp["name_key"])
            self.peaks.append({
                "peak_id": len(self.peaks) + 1, "sheet_id": sid, "sample_id": None,
                "compound_id": comp["compound_id"] if comp else None,
                "name_verbatim": ident or "(não identificado)", "name_source": "peak_table",
                "class_verbatim": get(row, c_cls) or None, "subclass_verbatim": None, "class_inherited": None,
                "peak_number": get(row, c_pk) or None, "rt_min": first_num(get(row, c_rt)),
                "lambda_max_text": lam or None, "lambda_max_json": json.dumps(parse_lambda(lam)) if lam else None,
                "uv_spectrum": None, "mw": None, "m_plus": None, "m_plus_h": None,
                "m_minus_h": mh or None, "m_minus_h_num": first_num(mh), "m2_plus_h": None, "m2_minus_h": None,
                "m_h_2neg": None, "m_h_3neg": None, "exact_mass": None, "adducts": None,
                "ms2_text": frag or None, "ms2_json": json.dumps(parse_ms2(frag)) if frag else None,
                "notes": None, "evidence": "REP", "src_sheet": name, "src_cell": a1(c_id, ri)})
            n_peaks += 1
        # sugestão de aba de origem: maior sobreposição de nomes (Jaccard)
        by_sheet = defaultdict(set)
        for p in self.peaks:
            if p["sample_id"] is not None:
                by_sheet[p["sheet_id"]].add(name_key(p["name_verbatim"]))
        best = max(((len(keys & s) / len(keys | s), sh_id) for sh_id, s in by_sheet.items() if s), default=(0, None))
        sh.update(status="peak_table_unassigned", layout="peak_table", n_peaks=n_peaks,
                  suggested_sheet=best[1] if best[0] > 0 else None, suggestion_score=round(best[0], 3),
                  notes="tabela de picos com TR, sem espécie nem referência")
        sug = next((s["name_verbatim"] for s in self.sheets if s["sheet_id"] == best[1]), None)
        self.issue("warning", "peak_table_unassigned",
                   f"{n_peaks} picos com TR sem espécie/estudo; sugestão: '{sug}' (Jaccard {best[0]:.2f})",
                   "sheet", sid, name)

    # pós-processamento ------------------------------------------------------
    def finish_conditions(self):
        by_cond = defaultdict(list)
        for m in self.measurements:
            if m["condition_id"]:
                by_cond[m["condition_id"]].append(m)
        for cond in self.conditions.values():
            ms = by_cond.get(cond["condition_id"], [])
            preps = Counter(m["preparation"] for m in ms if m["preparation"])
            tags = Counter(m["tag"] for m in ms if m["tag"])
            cond["preparation"] = preps.most_common(1)[0][0] if preps else None
            cond["tag"] = tags.most_common(1)[0][0] if tags else None
            resolved = None
            if cond["tag"] and cond["tag"].lower() in cond["letters"]:
                resolved = cond["letters"][cond["tag"].lower()]
            elif cond["row3_label"] and cond["row3_label"] in cond["codes"]:
                resolved = cond["codes"][cond["row3_label"]]
            cond["tag_resolved"] = resolved
            parts = [cond["row3_label"],
                     (f"({cond['tag']})" + (f" {resolved}" if resolved else "")) if cond["tag"] else
                     (resolved if resolved and cond["row3_label"] else None),
                     cond["preparation"], cond["header_label"]]
            if cond["quant_approach"] and cond["quant_approach"] != "Reference standards (DAD)":
                parts.append(cond["quant_approach"])
            cond["label"] = " · ".join(p for p in parts if p) or f"coluna {cond['col_letter']}"

    def finish_compounds(self):
        old = {}
        if self.old_db.exists():
            con = sqlite3.connect(self.old_db)
            for nm, st, cid, smi, iso, can in con.execute(
                    "SELECT molecule_name, status, cid, smiles, isomeric_smiles, canonical_smiles FROM chemical_data"):
                s = next((x for x in (iso, smi, can) if isinstance(x, str) and x.strip()), None)
                if st == "found" and s:
                    old[name_key(nm or "")] = (s, int(cid) if cid else None)
            con.close()
        mh_by_comp = defaultdict(list)
        for p in self.peaks:
            for fld, delta in (("m_minus_h_num", 1.00728), ("m_plus_h", -1.00728), ("m_plus", 0.0)):
                v = p[fld] if fld == "m_minus_h_num" else first_num(p[fld])
                if v:
                    mh_by_comp[p["compound_id"]].append(v + delta)
        for k, c in self.compounds.items():
            c["name_preferred"] = c["names"].most_common(1)[0][0]
            c["class_mode"] = next((x for x, _ in c["classes"].most_common() if x), None)
            c["subclass_mode"] = next((x for x, _ in c["subclasses"].most_common() if x), None)
            c["is_tentative"] = is_tentative(c["name_preferred"])
            c["n_samples"] = len(c["samples"])
            cand = old.get(k)
            c["smiles_candidate"], c["cid_candidate"] = (cand if cand else (None, None))
            c["candidate_source"] = "pipeline antigo: PubChem por nome (AUT-NR)" if cand else None
            c["mono_mass"], c["mass_check"] = None, "no_candidate"
            if cand and Chem:
                mol = Chem.MolFromSmiles(cand[0])
                if mol:
                    mono = Descriptors.ExactMolWt(mol)
                    c["mono_mass"] = round(mono, 4)
                    obs = mh_by_comp.get(c["compound_id"], [])
                    if not obs:
                        c["mass_check"] = "unchecked"
                    elif any(abs(mono - o) <= 0.6 for o in obs):
                        c["mass_check"] = "consistent"
                    else:
                        c["mass_check"] = "inconsistent"
                        self.issue("warning", "structure_mass_inconsistent",
                                   f"{c['name_preferred']}: SMILES candidato {mono:.2f} Da não bate com os íons "
                                   f"observados ({', '.join(f'{o:.0f}' for o in sorted(set(obs))[:4])})",
                                   "compound", c["compound_id"])
            if c["is_tentative"] and cand:
                self.issue("info", "structure_on_tentative_name",
                           f"{c['name_preferred']}: estrutura completa atribuída a um nome tentativo",
                           "compound", c["compound_id"])

    # escrita ----------------------------------------------------------------
    def write(self):
        self.out.mkdir(parents=True, exist_ok=True)
        db = self.out / "iberphenol_v2.sqlite"
        if db.exists():
            db.unlink()
        con = sqlite3.connect(db)
        con.executescript((Path(__file__).parent / "schema.sql").read_text(encoding="utf-8"))
        con.execute("PRAGMA foreign_keys = OFF")  # carga em lote; a integridade é conferida no fim

        def ins(table, rows, cols):
            con.executemany(f"INSERT INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                            [[r.get(c) for c in cols] for r in rows])

        def cols_of(table):
            return [r[1] for r in con.execute(f"PRAGMA table_info({table})")]

        ins("sheet", self.sheets, cols_of("sheet"))
        ins("study", self.studies.values(), cols_of("study"))
        ins("taxon", self.taxa.values(), cols_of("taxon"))
        ins("method", self.methods.values(), cols_of("method"))
        ins("sample", self.samples, cols_of("sample"))
        ins("sample_study", [{"sample_id": a, "study_id": b} for a, b in sorted(self.sample_study)],
            ["sample_id", "study_id"])
        ins("condition", self.conditions.values(), cols_of("condition"))
        ins("compound", self.compounds.values(), cols_of("compound"))
        ins("compound_name", [{"compound_id": c["compound_id"], "name_verbatim": n, "n": k}
                              for c in self.compounds.values() for n, k in c["names"].items()],
            ["compound_id", "name_verbatim", "n"])
        ins("peak", self.peaks, cols_of("peak"))
        ins("measurement", self.measurements, cols_of("measurement"))
        ins("issue", [{"issue_id": i + 1, **x} for i, x in enumerate(self.issues)], cols_of("issue"))
        meta = {
            "etl_version": ETL_VERSION,
            "run_at_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
            "source_file": self.xlsx.name,
            "source_sha256": hashlib.sha256(self.xlsx.read_bytes()).hexdigest(),
            "old_db": self.old_db.name if self.old_db.exists() else None,
            "compound_ids_provisional": "true",
            "evidence_codes": json.dumps(EVIDENCE_CODES, ensure_ascii=False),
        }
        con.executemany("INSERT INTO release (key, value) VALUES (?, ?)", meta.items())
        con.commit()
        fk = con.execute("PRAGMA foreign_key_check").fetchall()
        csv_dir = self.out / "csv"
        csv_dir.mkdir(exist_ok=True)
        for t in ["release", "sheet", "study", "taxon", "method", "sample", "sample_study", "condition",
                  "compound", "compound_name", "peak", "measurement", "issue"]:
            cur = con.execute(f"SELECT * FROM {t}")
            with open(csv_dir / f"{t}.csv", "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow([d[0] for d in cur.description])
                w.writerows(cur)
        self.report(con, fk, meta)
        con.close()
        return db

    def report(self, con, fk, meta):
        q = lambda s: con.execute(s).fetchall()
        one = lambda s: con.execute(s).fetchone()[0]
        L = [f"# Relatório de validação — ETL v2 {meta['etl_version']}", "",
             f"Gerado em {meta['run_at_utc']} a partir de `{meta['source_file']}` "
             f"(SHA-256 `{meta['source_sha256'][:16]}…`). Integridade referencial: "
             f"{'OK' if not fk else f'{len(fk)} violações'}.", ""]
        if self.old_db.exists():
            oc = sqlite3.connect(self.old_db)
            o1 = lambda s: oc.execute(s).fetchone()[0]
            rows_cmp = [
                ("Compostos", o1("SELECT COUNT(*) FROM compounds"), one("SELECT COUNT(*) FROM compound")),
                ("Picos (ocorrências com evidência espectral própria)", o1("SELECT COUNT(*) FROM spectral_data"),
                 one("SELECT COUNT(*) FROM peak")),
                ("Picos com tempo de retenção", 0, one("SELECT COUNT(*) FROM peak WHERE rt_min IS NOT NULL")),
                ("Medições", o1("SELECT COUNT(*) FROM measurements"), one("SELECT COUNT(*) FROM measurement")),
                ("Medições com condição identificada (extrato, cultivar, amostra)", 0,
                 one("SELECT COUNT(*) FROM measurement m JOIN condition c ON c.condition_id=m.condition_id "
                     "WHERE c.label NOT LIKE 'coluna %'")),
                ("Amostras", o1("SELECT COUNT(*) FROM samples"), one("SELECT COUNT(*) FROM sample")),
                ("Pares composto × amostra", o1("SELECT COUNT(*) FROM (SELECT DISTINCT compound_id, sample_id FROM measurements)"),
                 one("SELECT COUNT(*) FROM (SELECT DISTINCT compound_id, sample_id FROM peak WHERE sample_id IS NOT NULL)")),
                ("Compostos sem nenhum vínculo com amostra",
                 o1("SELECT COUNT(*) FROM compounds WHERE id NOT IN (SELECT compound_id FROM measurements WHERE compound_id IS NOT NULL)"),
                 one("SELECT COUNT(*) FROM compound c WHERE NOT EXISTS (SELECT 1 FROM peak p WHERE p.compound_id=c.compound_id AND p.sample_id IS NOT NULL)")),
            ]
            oc.close()
            L += ["## Pipeline antigo × ETL v2", "", "| Indicador | Pipeline antigo | ETL v2 |", "| --- | --- | --- |"]
            L += [f"| {a} | {b} | {c} |" for a, b, c in rows_cmp]
            L += ["", "Compostos sem vínculo no v2 são os que só aparecem nas abas \"Planilha\", ainda sem amostra.", ""]
        L += ["## Abas", "", "| Status | Layout | Abas | Picos | Medições |", "| --- | --- | --- | --- | --- |"]
        for st, lay, n, p, m in q("SELECT status, COALESCE(layout,'—'), COUNT(*), SUM(n_peaks), SUM(n_measurements) "
                                  "FROM sheet GROUP BY 1,2 ORDER BY 3 DESC"):
            L.append(f"| {st} | {lay} | {n} | {p or 0} | {m or 0} |")
        L += ["", "## Contagens", "", "| Entidade | n |", "| --- | --- |"]
        for label, sql in [
            ("Picos (total)", "SELECT COUNT(*) FROM peak"),
            ("Picos com amostra", "SELECT COUNT(*) FROM peak WHERE sample_id IS NOT NULL"),
            ("Picos sem contexto (abas Planilha)", "SELECT COUNT(*) FROM peak WHERE sample_id IS NULL"),
            ("Picos com tempo de retenção", "SELECT COUNT(*) FROM peak WHERE rt_min IS NOT NULL"),
            ("Picos com nome tirado da coluna de concentração", "SELECT COUNT(*) FROM peak WHERE name_source='conc_col_text'"),
            ("Picos de linhas de continuação (mesmo composto, outra fonte)", "SELECT COUNT(*) FROM peak WHERE name_source='continuation_row'"),
            ("Picos sem nenhuma medição", "SELECT COUNT(*) FROM peak p WHERE NOT EXISTS (SELECT 1 FROM measurement m WHERE m.peak_id=p.peak_id)"),
            ("Compostos (chave de nome)", "SELECT COUNT(*) FROM compound"),
            ("Compostos com nome tentativo (heurística)", "SELECT COUNT(*) FROM compound WHERE is_tentative=1"),
            ("Amostras", "SELECT COUNT(*) FROM sample"),
            ("Táxons (leitura inicial)", "SELECT COUNT(*) FROM taxon"),
            ("Táxons para revisão", "SELECT COUNT(*) FROM taxon WHERE needs_review=1"),
            ("Condições (colunas de concentração usadas)", "SELECT COUNT(*) FROM condition"),
            ("Medições", "SELECT COUNT(*) FROM measurement"),
            ("Medições em totais de classe", "SELECT COUNT(*) FROM measurement WHERE peak_id IS NULL"),
            ("Medições com estudo individual", "SELECT COUNT(*) FROM measurement WHERE study_id IS NOT NULL"),
            ("Estudos (citações)", "SELECT COUNT(*) FROM study WHERE is_citation=1"),
            ("Textos na coluna Reference que não são citação", "SELECT COUNT(*) FROM study WHERE is_citation=0"),
            ("Estudos com DOI na própria citação", "SELECT COUNT(*) FROM study WHERE doi IS NOT NULL"),
        ]:
            L.append(f"| {label} | {one(sql)} |")
        L += ["", "## Leitura dos valores", "", "| parse_status | flag | n |", "| --- | --- | --- |"]
        for ps, fl, n in q("SELECT parse_status, COALESCE(flag,'—'), COUNT(*) FROM measurement GROUP BY 1,2 ORDER BY 3 DESC"):
            L.append(f"| {ps} | {fl} | {n} |")
        L += ["", "| Unidade | Base | n |", "| --- | --- | --- |"]
        for u, b, n in q("SELECT COALESCE(unit,'—'), basis, COUNT(*) FROM measurement WHERE value IS NOT NULL "
                         "GROUP BY 1,2 ORDER BY 3 DESC"):
            L.append(f"| {u} | {b} | {n} |")
        L += ["", "## Estruturas herdadas do pipeline antigo", "", "| Checagem de massa | Compostos |", "| --- | --- |"]
        for mc, n in q("SELECT mass_check, COUNT(*) FROM compound GROUP BY 1 ORDER BY 2 DESC"):
            L.append(f"| {mc} | {n} |")
        L += ["", "## Tabelas de picos sem contexto (sugestão de aba)", "", "| Aba | Picos | Sugestão | Jaccard |",
              "| --- | --- | --- | --- |"]
        for a, p, s, j in q("SELECT s.name_verbatim, s.n_peaks, t.name_verbatim, s.suggestion_score FROM sheet s "
                            "LEFT JOIN sheet t ON t.sheet_id=s.suggested_sheet WHERE s.status='peak_table_unassigned'"):
            L.append(f"| {a} | {p} | {s or '—'} | {j} |")
        L += ["", "## Registro de problemas", "", "| Severidade | Código | n |", "| --- | --- | --- |"]
        for sv, code, n in q("SELECT severity, code, COUNT(*) FROM issue GROUP BY 1,2 ORDER BY 1, 3 DESC"):
            L.append(f"| {sv} | {code} | {n} |")
        L += ["", "Detalhe de cada problema, com aba e célula: `csv/issue.csv`.", ""]
        (self.out / "validation_report.md").write_text("\n".join(L), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--xlsx", default=str(ROOT / "Base dados_CP.xlsx"))
    ap.add_argument("--old", default=str(ROOT / "webapp" / "phenoldb.sqlite"))
    ap.add_argument("--out", default=str(ROOT / "etl_v2" / "out"))
    a = ap.parse_args()
    etl = ETL(a.xlsx, a.old, a.out)
    etl.run()
    db = etl.write()
    print(f"ok: {db}")
    print(f"  abas {len(etl.sheets)} · picos {len(etl.peaks)} · medições {len(etl.measurements)} · "
          f"compostos {len(etl.compounds)} · amostras {len(etl.samples)} · problemas {len(etl.issues)}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
