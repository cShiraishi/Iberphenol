#!/usr/bin/env python3
"""
Restructures all sheets in Base dados_CP.xlsx into normalized tables,
then exports a formatted multi-sheet Excel workbook.

Each sheet = one plant species, same 2-row header + row-3 metadata.
Column layout (0-indexed):
  0-1  : class / subclass (hierarchy, propagated downward)
  2    : molecule name
  3    : lambda_max_nm
  4    : uv_vis_spectrum
  5    : molecular_weight
  6-12 : MS ions  [M]+  [M+H]+  [M-H]-  [2M+H]+  [2M-H]-  [M-H]2-  [M-H]3-
  13   : exact mass
  14   : adducts
  15   : ms_fragments
  16 .. equip_col-1 : concentration values (variable width)
  equip_col   : equipment
  equip_col+1 : data_acquisition
  equip_col+2 : ionization_mode
  equip_col+3 : cultivar
  equip_col+4 : origin
  equip_col+5 : season
  equip_col+6 : plant_part
  equip_col+7 : reference   (= last col)
"""

import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUT  = "Base dados_CP.xlsx"
OUTPUT = "PhenolDB_estruturado.xlsx"

CONC_RE = re.compile(r"([\d.,]+)\s*(mg/g|µg/g|μg/g|µg/mL|μg/mL|mg/mL|%|g/kg|mg/kg|ng/g|ng/mL)")

# ── helpers ─────────────────────────────────────────────────────────────────

def cv(val):
    """Clean cell value to string."""
    if val is None:
        return ""
    return " ".join(str(val).split())


def find_col(row_vals, keyword):
    for i, v in enumerate(row_vals):
        if v and keyword.lower() in str(v).lower():
            return i
    return None


def parse_conc(text):
    """Return (numeric_value, unit) or (text, '')."""
    text = text.strip()
    m = CONC_RE.search(text)
    if m:
        num = m.group(1).replace(",", ".")
        return num, m.group(2)
    return text, ""


def split_refs(text):
    """Split 'Ref1; Ref2' into list."""
    parts = re.split(r";\s*(?=[A-Z])", text)
    return [p.strip() for p in parts if p.strip()]


# ── load workbook ────────────────────────────────────────────────────────────

print("Loading workbook…")
wb_in = openpyxl.load_workbook(INPUT, data_only=True, read_only=True)

# ── accumulators ─────────────────────────────────────────────────────────────

species_rows     = []   # {id, scientific_name}
reference_rows   = []   # {id, citation}
method_rows      = []   # {id, equipment, data_acquisition, ionization_mode}
sample_rows      = []   # {id, species_id, cultivar, origin, season, plant_part}
compound_rows    = []   # {id, compound_class, subclass, molecule_name}
spectral_rows    = []   # {id, compound_id, lambda_max, …}
measurement_rows = []   # {id, compound_id, sample_id, method_id, reference_id,
                        #   matrix_label, concentration_value, concentration_unit, raw_text}

species_id   = 0
ref_id       = 0
method_id    = 0
sample_id    = 0
compound_id  = 0
spectral_id  = 0
meas_id      = 0

ref_cache    = {}   # citation -> id
method_cache = {}   # equipment -> id

SKIP_NAMES = {"planilha", "sheet"}

# ── process each sheet ───────────────────────────────────────────────────────

for sh_name in wb_in.sheetnames:
    if any(sh_name.strip().lower().startswith(s) for s in SKIP_NAMES):
        continue

    ws = wb_in[sh_name]
    rows = [tuple(c.value for c in row) for row in ws.iter_rows()]

    if len(rows) < 4:
        continue

    # ── locate key columns ──────────────────────────────────────────────────
    hdr1 = rows[0]
    hdr2 = rows[1]

    equip_col = find_col(hdr2, "Equipment")
    if equip_col is None:
        continue  # can't parse sheet

    ref_col     = equip_col + 7
    part_col    = equip_col + 6
    season_col  = equip_col + 5
    origin_col  = equip_col + 4
    cultivar_col= equip_col + 3
    ioniz_col   = equip_col + 2
    acq_col     = equip_col + 1
    conc_end    = equip_col          # concentrations: cols 16 .. equip_col-1

    def get(row, idx):
        try:
            return cv(row[idx])
        except IndexError:
            return ""

    # ── metadata row (row index 2) ──────────────────────────────────────────
    meta = rows[2]

    equipment  = get(meta, equip_col)
    data_acq   = get(meta, acq_col)
    ioniz      = get(meta, ioniz_col)
    cultivar   = get(meta, cultivar_col)
    origin     = get(meta, origin_col)
    season     = get(meta, season_col)
    plant_part = get(meta, part_col)
    ref_text   = get(meta, ref_col)

    if not ref_text:
        # try scanning last few cols for any text
        for ci in range(len(meta) - 1, equip_col + 2, -1):
            v = get(meta, ci)
            if v:
                ref_text = v
                break

    # ── species ─────────────────────────────────────────────────────────────
    species_id += 1
    species_rows.append({"id": species_id, "scientific_name": sh_name.strip()})

    # ── references ──────────────────────────────────────────────────────────
    this_ref_ids = []
    for citation in split_refs(ref_text):
        if citation not in ref_cache:
            ref_id += 1
            ref_cache[citation] = ref_id
            reference_rows.append({"id": ref_id, "citation": citation})
        this_ref_ids.append(ref_cache[citation])
    default_ref_id = this_ref_ids[0] if this_ref_ids else ""

    # ── method ──────────────────────────────────────────────────────────────
    if equipment not in method_cache:
        method_id += 1
        method_cache[equipment] = method_id
        method_rows.append({
            "id": method_id,
            "equipment": equipment,
            "data_acquisition": data_acq,
            "ionization_mode": ioniz,
        })
    this_method_id = method_cache[equipment]

    # ── sample ──────────────────────────────────────────────────────────────
    sample_id += 1
    sample_rows.append({
        "id": sample_id,
        "species_id": species_id,
        "cultivar_variety": cultivar,
        "origin": origin,
        "season": season,
        "plant_part": plant_part,
    })

    # ── compound data rows ──────────────────────────────────────────────────
    current_class    = cv(meta[0]) if meta[0] else ""
    current_subclass = cv(meta[1]) if meta[1] else ""

    for row in rows[3:]:
        cls = get(row, 0)
        sub = get(row, 1)
        mol = get(row, 2)

        if cls:
            current_class = cls
        if sub:
            current_subclass = sub

        if not mol:
            continue

        compound_id += 1
        compound_rows.append({
            "id":             compound_id,
            "species_id":     species_id,
            "compound_class": current_class,
            "subclass":       current_subclass,
            "molecule_name":  mol,
        })

        spectral_id += 1
        spectral_rows.append({
            "id":              spectral_id,
            "compound_id":     compound_id,
            "lambda_max_nm":   get(row, 3),
            "uv_vis_spectrum": get(row, 4),
            "molecular_weight":get(row, 5),
            "m_plus":          get(row, 6),
            "m_plus_h":        get(row, 7),
            "m_minus_h":       get(row, 8),
            "m2_plus_h":       get(row, 9),
            "m2_minus_h":      get(row, 10),
            "m_minus_h_2neg":  get(row, 11),
            "m_minus_h_3neg":  get(row, 12),
            "exact_mass":      get(row, 13),
            "adducts":         get(row, 14),
            "ms_fragments":    get(row, 15),
        })

        # concentrations: any non-empty value between col 16 and equip_col-1
        for ci in range(16, conc_end):
            raw = get(row, ci)
            if not raw:
                continue
            val, unit = parse_conc(raw)
            meas_id += 1
            measurement_rows.append({
                "id":                  meas_id,
                "compound_id":         compound_id,
                "sample_id":           sample_id,
                "method_id":           this_method_id,
                "reference_id":        default_ref_id,
                "concentration_value": val,
                "concentration_unit":  unit,
                "raw_text":            raw,
            })

wb_in.close()

print(f"  Species:      {len(species_rows)}")
print(f"  References:   {len(reference_rows)}")
print(f"  Methods:      {len(method_rows)}")
print(f"  Samples:      {len(sample_rows)}")
print(f"  Compounds:    {len(compound_rows)}")
print(f"  Spectral:     {len(spectral_rows)}")
print(f"  Measurements: {len(measurement_rows)}")

# ── build output workbook ────────────────────────────────────────────────────

print("\nBuilding output workbook…")

H_FILL  = PatternFill("solid", fgColor="1F4E79")
H_FONT  = Font(bold=True, color="FFFFFF", size=10)
A_FILL  = PatternFill("solid", fgColor="D6E4F0")
BD_SIDE = Side(style="thin", color="CCCCCC")
BD      = Border(left=BD_SIDE, right=BD_SIDE, top=BD_SIDE, bottom=BD_SIDE)

def auto_width(ws_out):
    for col in ws_out.columns:
        best = 0
        letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                best = max(best, len(str(cell.value or "")))
            except Exception:
                pass
        ws_out.column_dimensions[letter].width = min(best + 4, 55)

def write_sheet(wb_out, title, headers, data_rows):
    ws_out = wb_out.create_sheet(title=title)
    # header
    for ci, h in enumerate(headers, 1):
        c = ws_out.cell(row=1, column=ci, value=h)
        c.font = H_FONT
        c.fill = H_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BD
    ws_out.row_dimensions[1].height = 28
    # data
    for ri, row_dict in enumerate(data_rows, 2):
        fill = A_FILL if ri % 2 == 0 else PatternFill()
        for ci, h in enumerate(headers, 1):
            c = ws_out.cell(row=ri, column=ci, value=row_dict.get(h, ""))
            c.fill = fill
            c.border = BD
            c.alignment = Alignment(vertical="top", wrap_text=(ci == len(headers)))
    ws_out.freeze_panes = "A2"
    auto_width(ws_out)
    return ws_out


TABLES = [
    ("compounds",    ["id","species_id","compound_class","subclass","molecule_name"]),
    ("measurements", ["id","compound_id","sample_id","method_id","reference_id",
                      "concentration_value","concentration_unit","raw_text"]),
    ("spectral_data",["id","compound_id","lambda_max_nm","uv_vis_spectrum","molecular_weight",
                      "m_plus","m_plus_h","m_minus_h","m2_plus_h","m2_minus_h",
                      "m_minus_h_2neg","m_minus_h_3neg","exact_mass","adducts","ms_fragments"]),
    ("samples",      ["id","species_id","cultivar_variety","origin","season","plant_part"]),
    ("species",      ["id","scientific_name"]),
    ("methods",      ["id","equipment","data_acquisition","ionization_mode"]),
    ("references",   ["id","citation"]),
]

DATA_MAP = {
    "compounds":     compound_rows,
    "measurements":  measurement_rows,
    "spectral_data": spectral_rows,
    "samples":       sample_rows,
    "species":       species_rows,
    "methods":       method_rows,
    "references":    reference_rows,
}

wb_out = openpyxl.Workbook()
wb_out.remove(wb_out.active)

for title, headers in TABLES:
    write_sheet(wb_out, title, headers, DATA_MAP[title])
    print(f"  Sheet '{title}' written ({len(DATA_MAP[title])} rows)")

wb_out.save(OUTPUT)
print(f"\nSaved: {OUTPUT}")
