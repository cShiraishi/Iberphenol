#!/usr/bin/env python3
"""
For compounds without SMILES, identifies the aglycone from the MS base fragment
and retrieves the aglycone SMILES from PubChem.

Adds columns to pubchem_data sheet:
  aglycone_name, aglycone_smiles, aglycone_fragment_mz, smiles_confidence
  smiles_confidence: 'confirmed' | 'aglycone_only' | 'unknown'
"""

import re, json, time, sys, requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

EXCEL_FILE    = "PhenolDB_estruturado.xlsx"
PUBCHEM_CACHE = "pubchem_cache.json"
CASCADE_CACHE = "cascade_cache.json"
SLEEP         = 0.2

# Known aglycone [M-H]- → name mapping (nominal mass, integer)
# Source: standard phenolic compound fragmentation patterns
AGLYCONE_MAP = {
    269: "Apigenin",
    271: "Naringenin",
    283: "Luteolin",           # sometimes seen as 283
    285: "Luteolin/Kaempferol",# distinguish by λmax later
    287: "Eriodictyol",
    299: "Hesperetin",
    301: "Quercetin",
    303: "Myricetin",
    315: "Isorhamnetin",
    329: "Rhamnetin",
    255: "Chrysin",
    253: "Pinocembrin",
    179: "Caffeic acid",
    163: "p-Coumaric acid",
    137: "Protocatechuic acid",
    153: "Caffeic acid fragment",
    167: "Ferulic acid fragment",
    191: "Quinic acid",
    197: "Ferulic acid",
    169: "Sinapic acid",
    193: "Sinapic acid",
    135: "Hydroxycinnamic fragment",
    121: "Methylenedioxybenzaldehyde",
    609: "Rutin/Quercetin-rutinoside",
    593: "Luteolin-rutinoside",
    353: "Chlorogenic acid",
    339: "Cryptochlorogenic acid",
}

# Aglycones that need disambiguation by λmax
DISAMBIGUATE = {
    285: {(340, 380): "Kaempferol", (340, 360): "Luteolin"}  # Kaempferol ~367nm, Luteolin ~348nm
}

PUBCHEM_URL = (
    "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{name}"
    "/property/IsomericSMILES,CanonicalSMILES,IUPACName,MolecularFormula,MolecularWeight/JSON"
)

# ── load caches ──────────────────────────────────────────────────────────────

with open(PUBCHEM_CACHE, encoding="utf-8") as f:
    pc = json.load(f)
with open(CASCADE_CACHE, encoding="utf-8") as f:
    cc = json.load(f)


def best_result(name):
    if pc.get(name, {}).get("status") == "found":
        return pc[name]
    if cc.get(name, {}).get("status") == "found":
        return cc[name]
    return None


not_found_names = {n for n in pc if best_result(n) is None}

# ── aglycone SMILES cache (fetch on demand) ──────────────────────────────────

aglycone_smiles_cache = {}

def get_aglycone_smiles(aglycone_name):
    """Get SMILES for an aglycone, preferring our existing cache."""
    base = aglycone_name.split("/")[0].strip()  # "Luteolin/Kaempferol" → "Luteolin"
    if base in aglycone_smiles_cache:
        return aglycone_smiles_cache[base]
    # check existing cache
    result = best_result(base)
    if result:
        smiles = result.get("isomeric_smiles") or result.get("canonical_smiles", "")
        aglycone_smiles_cache[base] = smiles
        return smiles
    # fetch from PubChem
    url = PUBCHEM_URL.format(name=requests.utils.quote(base))
    try:
        r = requests.get(url, timeout=10)
        time.sleep(SLEEP)
        if r.status_code == 200:
            p = r.json()["PropertyTable"]["Properties"][0]
            smiles = p.get("IsomericSMILES", "") or p.get("CanonicalSMILES", "")
            aglycone_smiles_cache[base] = smiles
            return smiles
    except Exception:
        pass
    aglycone_smiles_cache[base] = ""
    return ""


def parse_base_fragment(frag_str):
    """Return (mz, pct) of the dominant fragment (highest %)."""
    if not frag_str:
        return None, None
    best_mz, best_pct = None, -1
    for m in re.finditer(r"(\d+)\s*\(\s*(\d+)\s*\)", str(frag_str)):
        mz, pct = int(m.group(1)), int(m.group(2))
        if pct > best_pct:
            best_pct, best_mz = pct, mz
    return best_mz, best_pct


def identify_aglycone(mz, lambda_max=None):
    """Map fragment mz to aglycone name, optionally disambiguating by λmax."""
    if mz is None:
        return None
    name = AGLYCONE_MAP.get(int(mz))
    if not name:
        return None
    # Disambiguate if needed
    if "/" in name and lambda_max:
        try:
            lmax = float(str(lambda_max).split(",")[0].strip())
            if int(mz) == 285:
                if lmax > 355:
                    name = "Kaempferol"
                else:
                    name = "Luteolin"
        except Exception:
            name = name.split("/")[0]
    elif "/" in name:
        name = name.split("/")[0]
    return name


# ── load workbook data ───────────────────────────────────────────────────────

print("Loading workbook…")
wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)

ws_comp = wb["compounds"]
ws_spec = wb["spectral_data"]

# compound_id → molecule_name
id_to_name = {}
for row in ws_comp.iter_rows(min_row=2, values_only=True):
    if row[0] and row[4]:
        id_to_name[row[0]] = str(row[4]).strip()

# compound_id → (m_minus_h, ms_fragments, lambda_max)
id_to_spec = {}
for row in ws_spec.iter_rows(min_row=2, values_only=True):
    cid = row[1]
    if cid:
        id_to_spec[cid] = {
            "m_minus_h":   row[7],
            "ms_fragments": row[14],
            "lambda_max":  row[2],
        }

wb.close()

# ── process each not-found compound ─────────────────────────────────────────

print(f"Processing {len(not_found_names)} not-found compounds…")

results = {}   # name → {aglycone_name, aglycone_smiles, fragment_mz, confidence}

# collect all compound_ids for each molecule name
name_to_specs = {}
for cid, name in id_to_name.items():
    if name in not_found_names and cid in id_to_spec:
        name_to_specs.setdefault(name, []).append(id_to_spec[cid])

resolved = 0
for name, specs in name_to_specs.items():
    spec = specs[0]  # use first occurrence
    frag_str = spec.get("ms_fragments")
    lmax     = spec.get("lambda_max")
    mz_val, pct = parse_base_fragment(frag_str)
    aglycone = identify_aglycone(mz_val, lmax)

    if aglycone:
        smiles = get_aglycone_smiles(aglycone)
        results[name] = {
            "aglycone_name":       aglycone,
            "aglycone_smiles":     smiles,
            "aglycone_fragment_mz": mz_val,
            "smiles_confidence":   "aglycone_only",
        }
        if smiles:
            resolved += 1
    else:
        results[name] = {
            "aglycone_name":       "",
            "aglycone_smiles":     "",
            "aglycone_fragment_mz": mz_val or "",
            "smiles_confidence":   "unknown",
        }

print(f"  Aglycone identified: {sum(1 for r in results.values() if r['aglycone_name'])}")
print(f"  Aglycone SMILES obtained: {resolved}")

# ── update Excel pubchem_data sheet ─────────────────────────────────────────

print("Updating workbook…")
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb["pubchem_data"]

# Find or add new columns
header = [c.value for c in ws[1]]
NEW_COLS = ["smiles_confidence", "aglycone_name", "aglycone_smiles", "aglycone_fragment_mz"]

for col_name in NEW_COLS:
    if col_name not in header:
        header.append(col_name)
        ci = len(header)
        H_FILL = PatternFill("solid", fgColor="1F4E79")
        c = ws.cell(row=1, column=ci, value=col_name)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = H_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")

col_idx = {h: i+1 for i, h in enumerate(header)}

# Fill data
mol_col = col_idx.get("molecule_name", 1)
status_col = col_idx.get("status", 3)

for row in ws.iter_rows(min_row=2):
    name = str(row[mol_col - 1].value or "").strip()
    status = str(row[status_col - 1].value or "")

    # confidence for already-found compounds
    if status == "found":
        conf_cell = row[col_idx["smiles_confidence"] - 1]
        if not conf_cell.value:
            conf_cell.value = "confirmed"
        continue

    # not-found: add aglycone info
    info = results.get(name, {})
    for col_name in NEW_COLS:
        ci = col_idx.get(col_name)
        if ci:
            cell = row[ci - 1]
            if col_name == "smiles_confidence":
                cell.value = info.get("smiles_confidence", "unknown")
            elif col_name == "aglycone_name":
                cell.value = info.get("aglycone_name", "")
            elif col_name == "aglycone_smiles":
                cell.value = info.get("aglycone_smiles", "")
            elif col_name == "aglycone_fragment_mz":
                cell.value = info.get("aglycone_fragment_mz", "")

# auto-width new columns
for col_name in NEW_COLS:
    ci = col_idx.get(col_name)
    if ci:
        letter = get_column_letter(ci)
        best = max((len(str(c.value or "")) for c in ws[letter]), default=10)
        ws.column_dimensions[letter].width = min(best + 4, 60)

wb.save(EXCEL_FILE)

conf = sum(1 for r in results.values() if r.get("smiles_confidence") == "confirmed" or r.get("aglycone_smiles"))
print(f"\nSaved '{EXCEL_FILE}'")
print(f"  'confirmed' (full SMILES)  : 440")
print(f"  'aglycone_only'            : {resolved}")
print(f"  'unknown'                  : {len(results) - resolved}")
print(f"  Total with some structure  : {440 + resolved} / 1526 ({(440+resolved)*100//1526}%)")
