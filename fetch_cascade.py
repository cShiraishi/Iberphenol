#!/usr/bin/env python3
"""
Cascade SMILES lookup for compounds not found in PubChem.

Pipeline per compound:
  1. PubChem by original name          (already cached)
  2. PubChem by normalised name        (clean punctuation, remove stereo prefixes)
  3. ChEMBL by preferred name
  4. ChEMBL by synonym
  5. LOTUS by original name
  6. LOTUS by simplified name          (remove O- notation, hyphens → spaces)
  Mark remaining as 'not_found' with m/z for future manual curation.

Results stored in cascade_cache.json and written to PhenolDB_estruturado.xlsx.
"""

import json, re, time, sys
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

EXCEL_FILE     = "PhenolDB_estruturado.xlsx"
PUBCHEM_CACHE  = "pubchem_cache.json"
CASCADE_CACHE  = "cascade_cache.json"
SLEEP          = 0.15
BATCH_SAVE     = 50
TIMEOUT        = 8    # seconds per request

PUBCHEM_URL = (
    "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{name}"
    "/property/IsomericSMILES,CanonicalSMILES,IUPACName,"
    "MolecularFormula,MolecularWeight/JSON"
)
CHEMBL_NAME_URL    = "https://www.ebi.ac.uk/chembl/api/data/molecule?pref_name__iexact={name}&format=json"
CHEMBL_SYNONYM_URL = "https://www.ebi.ac.uk/chembl/api/data/molecule_synonyms?synonyms__iexact={name}&format=json&limit=1"
LOTUS_URL          = "https://lotus.naturalproducts.net/api/search/simple?query={name}"

CONC_RE = re.compile(r"([\d.,]+)\s*(mg/g|µg/g|μg/g|µg/mL|μg/mL|mg/mL|%|g/kg|mg/kg|ng/g|ng/mL)")

# ── name normalisation helpers ───────────────────────────────────────────────

STEREO_PREFIXES = re.compile(
    r"^(trans[- ]|cis[- ]|\(?[eEzZ]\)?[- ]|α-|β-|α |β |\([-+]\)-|\(±\)-)", re.I
)
GLYCOSIDE_SUFFIXES = re.compile(
    r"[-\s]("
    r"\d+[''\"]-O-[\w-]+-\d+[''\"]-O-[\w-]+"   # double glycoside
    r"|\d+[''\"]-O-[\w-]+"                       # single glycoside pos
    r"|[36]-O-[\w-]+"
    r"|O-[\w]+-[\w-]+"
    r"|[cC]-[\w]+-[cC]-[\w]+"
    r"|[cC]-[\w]+"
    r"|[\w]+-O-[\w]+-[\w-]+"
    r"|[\w]+-O-[\w]+"
    r"|[\w]+-[\w]-[\w]+(oside|uronide|oyl)"
    r"|[\w]+(oside|uronide|oyl|glucoside|rhamnoside|rutinoside|galactoside"
    r"|hexoside|pentoside|dihexoside|trihexoside|arabinoside)"
    r")\b",
    re.I
)


def normalise_name(name):
    """Remove stereo prefixes; return cleaned string."""
    n = STEREO_PREFIXES.sub("", name).strip()
    n = re.sub(r"\s{2,}", " ", n)
    return n


def name_variants(original):
    """
    Return safe alternative name forms to try, without changing the compound identity.
    Does NOT strip glycoside/substituent parts — that would yield a different compound.
    """
    variants = []

    # 1. Remove stereo prefixes (trans-, cis-, (E)-, (Z)-, α-, β-, (+)-, (-)-…)
    norm = normalise_name(original)
    if norm != original:
        variants.append(norm)

    # 2. Remove trailing "isomer N", "derivative", "analogue" (annotation, not structure)
    cleaned = re.sub(r"\s+(isomer|derivative|analogue|analog)\s*\d*$", "", original, flags=re.I).strip()
    if cleaned != original:
        variants.append(cleaned)

    # 3. Standardise quote characters in positions (2'' → 2'', ' → ')
    std = original.replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    if std != original:
        variants.append(std)

    # 4. Replace spelled-out sugar names with standard abbreviations
    sugar_map = {
        r"\bhexoside\b":    "glucoside",
        r"\bpentoside\b":   "arabinoside",
        r"\bdeoxyhexoside\b": "rhamnoside",
    }
    alt = original
    for pattern, replacement in sugar_map.items():
        alt = re.sub(pattern, replacement, alt, flags=re.I)
    if alt != original:
        variants.append(alt)

    return list(dict.fromkeys(variants))  # deduplicate, preserve order


# ── HTTP helpers ─────────────────────────────────────────────────────────────

def pubchem_lookup(name):
    url = PUBCHEM_URL.format(name=requests.utils.quote(name))
    try:
        r = requests.get(url, timeout=TIMEOUT)
        if r.status_code == 200:
            p = r.json()["PropertyTable"]["Properties"][0]
            return {
                "cid":               p.get("CID", ""),
                "isomeric_smiles":   p.get("SMILES", "") or p.get("IsomericSMILES", ""),
                "canonical_smiles":  p.get("ConnectivitySMILES", "") or p.get("CanonicalSMILES", ""),
                "iupac_name":        p.get("IUPACName", ""),
                "molecular_formula": p.get("MolecularFormula", ""),
                "molecular_weight":  str(p.get("MolecularWeight", "")),
                "status": "found",
            }
        return None
    except Exception:
        return None


def lotus_lookup(name):
    """Search LOTUS natural products database by name."""
    url = LOTUS_URL.format(name=requests.utils.quote(name))
    try:
        r = requests.get(url, timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        nps = r.json().get("naturalProducts", [])
        if not nps:
            return None
        mol = nps[0]
        smiles = mol.get("smiles", "")
        if not smiles:
            return None
        return {
            "cid":               "",
            "isomeric_smiles":   smiles,
            "canonical_smiles":  smiles,
            "iupac_name":        mol.get("iupac_name", ""),
            "molecular_formula": mol.get("molecular_formula", ""),
            "molecular_weight":  str(mol.get("molecular_weight", "")),
            "status": "found",
        }
    except Exception:
        return None


def lotus_name_variants(name):
    """Generate LOTUS-friendly name variants by simplifying glycoside notation."""
    variants = []
    # Remove "O-" position notation: "7-O-rutinoside" → "7-rutinoside"
    v1 = re.sub(r"(\d+[''\"]*)-O-", r"\1-", name)
    if v1 != name:
        variants.append(v1)
    # Replace hyphens with spaces (LOTUS prefers natural language)
    v2 = re.sub(r"-", " ", name)
    if v2 != name:
        variants.append(v2)
    # Both: remove O- and use spaces
    v3 = re.sub(r"-", " ", v1)
    if v3 not in variants and v3 != name:
        variants.append(v3)
    # Remove position numbers entirely: "7-O-rutinoside" → "rutinoside"
    v4 = re.sub(r"\b\d+[''\"]*-O?-?", "", name).strip(" -")
    if v4 != name and len(v4) > 4:
        variants.append(v4)
    return list(dict.fromkeys(variants))


def chembl_lookup(name):
    """Try ChEMBL preferred name, then synonym."""
    for url_tpl in (CHEMBL_NAME_URL, CHEMBL_SYNONYM_URL):
        url = url_tpl.format(name=requests.utils.quote(name))
        try:
            r = requests.get(url, timeout=TIMEOUT)
            if r.status_code != 200:
                continue
            data = r.json()

            # pref_name endpoint returns molecules list
            mols = data.get("molecules") or []
            # synonym endpoint returns molecule_synonyms list
            if not mols:
                syns = data.get("molecule_synonyms") or []
                chembl_ids = list({s["molecule_chembl_id"] for s in syns})
                if chembl_ids:
                    mol_url = f"https://www.ebi.ac.uk/chembl/api/data/molecule/{chembl_ids[0]}?format=json"
                    mr = requests.get(mol_url, timeout=TIMEOUT)
                    if mr.status_code == 200:
                        mols = [mr.json()]
                    time.sleep(SLEEP)

            if mols:
                mol = mols[0]
                structs = mol.get("molecule_structures") or {}
                smiles = structs.get("canonical_smiles", "")
                if smiles:
                    return {
                        "cid":               "",
                        "isomeric_smiles":   structs.get("standard_inchi", ""),
                        "canonical_smiles":  smiles,
                        "iupac_name":        mol.get("pref_name", ""),
                        "molecular_formula": (mol.get("molecule_properties") or {}).get("full_molformula", ""),
                        "molecular_weight":  str((mol.get("molecule_properties") or {}).get("full_mwt", "")),
                        "status": "found",
                    }
        except Exception:
            pass
        time.sleep(SLEEP)
    return None


# ── load caches ──────────────────────────────────────────────────────────────

with open(PUBCHEM_CACHE, encoding="utf-8") as f:
    pubchem_cache = json.load(f)

try:
    with open(CASCADE_CACHE, encoding="utf-8") as f:
        cascade_cache = json.load(f)
    print(f"Cascade cache loaded: {len(cascade_cache)} entries")
except FileNotFoundError:
    cascade_cache = {}
    print("No cascade cache, starting fresh.")


def save_cascade():
    with open(CASCADE_CACHE, "w", encoding="utf-8") as f:
        json.dump(cascade_cache, f, ensure_ascii=False, indent=2)


# ── collect names that need cascade lookup ───────────────────────────────────

not_found = [
    name for name, info in pubchem_cache.items()
    if info.get("status") != "found"
]
to_process = [n for n in not_found if n not in cascade_cache]
print(f"Not found in PubChem: {len(not_found)}")
print(f"Still to process:     {len(to_process)}\n")

count = 0
for idx, name in enumerate(to_process, 1):
    result = None
    source = ""

    # step 2: PubChem with name variants
    for variant in name_variants(name):
        r = pubchem_lookup(variant)
        time.sleep(SLEEP)
        if r:
            result = r
            source = f"pubchem_normalised:{variant}"
            break

    # step 3: PubChem with extra name variants
    if not result:
        extras = []
        # remove trailing numbers/greek letters used as identifiers
        v = re.sub(r'\s+[IVX]+$', '', name).strip()
        if v != name: extras.append(v)
        # try without position numbers entirely: "3-O-" → "O-"
        v2 = re.sub(r'\d+-O-', 'O-', name)
        if v2 != name: extras.append(v2)
        # try lowercase
        extras.append(name.lower())
        for variant in extras:
            r = pubchem_lookup(variant)
            time.sleep(SLEEP)
            if r:
                result = r
                source = f"pubchem_extra:{variant}"
                break

    if result:
        result["source"] = source
        cascade_cache[name] = result
        status = "found"
    else:
        cascade_cache[name] = {"status": "not_found", "source": ""}
        status = "not_found"

    count += 1
    print(f"  [{idx}/{len(to_process)}] {name[:55]:<55}  {status}  ({source})")

    if count % BATCH_SAVE == 0:
        save_cascade()

save_cascade()

found_cascade = sum(1 for v in cascade_cache.values() if v.get("status") == "found")
print(f"\nCascade: {found_cascade}/{len(cascade_cache)} resolved")

# ── merge all results ────────────────────────────────────────────────────────

def best_result(name):
    """Return best available result dict for a molecule name."""
    pc = pubchem_cache.get(name, {})
    if pc.get("status") == "found":
        pc["source"] = "pubchem"
        return pc
    cc = cascade_cache.get(name, {})
    if cc.get("status") == "found":
        return cc
    return {"status": "not_found", "source": "", "isomeric_smiles": "",
            "canonical_smiles": "", "iupac_name": "",
            "molecular_formula": "", "molecular_weight": "", "cid": ""}


# ── update Excel ─────────────────────────────────────────────────────────────

print("\nUpdating workbook…")
wb = openpyxl.load_workbook(EXCEL_FILE)

if "pubchem_data" in wb.sheetnames:
    del wb["pubchem_data"]

H_FILL = PatternFill("solid", fgColor="1F4E79")
H_FONT = Font(bold=True, color="FFFFFF", size=10)
A_FILL = PatternFill("solid", fgColor="D6E4F0")
BD     = Border(
    left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),  bottom=Side(style="thin", color="CCCCCC"),
)

HEADERS = [
    "molecule_name", "status", "source", "cid",
    "isomeric_smiles", "canonical_smiles",
    "iupac_name", "molecular_formula", "molecular_weight",
]

# collect unique names in order from compounds sheet
ws_comp = wb["compounds"]
unique_names = []
seen = set()
for row in ws_comp.iter_rows(min_row=2, values_only=True):
    name = str(row[4]).strip() if row[4] else ""
    if name and name not in seen:
        seen.add(name)
        unique_names.append(name)

ws_out = wb.create_sheet(title="pubchem_data")
for ci, h in enumerate(HEADERS, 1):
    c = ws_out.cell(row=1, column=ci, value=h)
    c.font = H_FONT; c.fill = H_FILL; c.border = BD
    c.alignment = Alignment(horizontal="center", vertical="center")
ws_out.row_dimensions[1].height = 28
ws_out.freeze_panes = "A2"

for ri, name in enumerate(unique_names, 2):
    info = best_result(name)
    row_data = [
        name,
        info.get("status", ""),
        info.get("source", ""),
        info.get("cid", ""),
        info.get("isomeric_smiles", ""),
        info.get("canonical_smiles", ""),
        info.get("iupac_name", ""),
        info.get("molecular_formula", ""),
        info.get("molecular_weight", ""),
    ]
    fill = A_FILL if ri % 2 == 0 else PatternFill()
    for ci, val in enumerate(row_data, 1):
        c = ws_out.cell(row=ri, column=ci, value=val)
        c.fill = fill; c.border = BD
        c.alignment = Alignment(vertical="top")

for col in ws_out.columns:
    letter = get_column_letter(col[0].column)
    best = max((len(str(c.value or "")) for c in col), default=10)
    ws_out.column_dimensions[letter].width = min(best + 4, 70)

wb.save(EXCEL_FILE)

total   = len(unique_names)
found   = sum(1 for n in unique_names if best_result(n).get("status") == "found")
print(f"\nSaved '{EXCEL_FILE}'")
print(f"  SMILES resolved: {found}/{total} ({100*found//total}%)")
print(f"  Remaining not found: {total - found}")
