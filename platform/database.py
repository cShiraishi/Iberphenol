import sqlite3
import os
import shutil
from pathlib import Path

BASE = Path(__file__).parent.parent
_BUNDLE_DB = Path(__file__).parent / "phenoldb.sqlite"
EXCEL_PATH = BASE / "PhenolDB_estruturado.xlsx"
CASCADE_CACHE = BASE / "cascade_cache.json"
PUBCHEM_CACHE = BASE / "pubchem_cache.json"

# Updated at startup by init_db()
DB_PATH = _BUNDLE_DB


def _runtime_db_path() -> Path:
    """On Vercel the bundle is read-only; copy the SQLite to /tmp."""
    on_vercel = bool(os.environ.get("VERCEL")) or not os.access(str(_BUNDLE_DB.parent), os.W_OK)
    if on_vercel:
        tmp = Path("/tmp/phenoldb.sqlite")
        if not tmp.exists() and _BUNDLE_DB.exists():
            shutil.copy(str(_BUNDLE_DB), str(tmp))
        return tmp
    return _BUNDLE_DB


def init_db(force: bool = False):
    global DB_PATH
    DB_PATH = _runtime_db_path()
    if DB_PATH.exists() and not force:
        return

    # Heavy imports only when rebuilding the DB from scratch
    import json
    import pandas as pd

    print("Initializing database from Excel...")
    conn = sqlite3.connect(str(DB_PATH))

    xl = pd.ExcelFile(EXCEL_PATH)
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet)
            df.columns = [
                str(c).encode("ascii", "replace").decode("ascii")
                .replace("?", "").replace(" ", "_").replace("/", "_")
                .replace("-", "_").replace("(", "").replace(")", "")
                .strip("_") or f"col_{i}"
                for i, c in enumerate(df.columns)
            ]
            for col in df.select_dtypes(include="object").columns:
                df[col] = df[col].apply(
                    lambda x: x.encode("utf-8", "replace").decode("utf-8") if isinstance(x, str) else x
                )
            safe_name = sheet.lower().replace(" ", "_").replace("-", "_")
            df.to_sql(safe_name, conn, if_exists="replace", index=False)
            print(f"  Loaded sheet '{sheet}' -> table '{safe_name}' ({len(df)} rows)")
        except Exception as e:
            print(f"  Warning: could not load sheet '{sheet}': {e}")

    chem: dict = {}
    if PUBCHEM_CACHE.exists():
        with open(PUBCHEM_CACHE, encoding="utf-8") as f:
            chem.update(json.load(f))
    if CASCADE_CACHE.exists():
        with open(CASCADE_CACHE, encoding="utf-8") as f:
            chem.update(json.load(f))

    rows = []
    for name, data in chem.items():
        smiles = data.get("isomeric_smiles") or data.get("canonical_smiles") or ""
        rows.append({
            "molecule_name": name,
            "status": data.get("status", "unknown"),
            "source": data.get("source", ""),
            "cid": data.get("cid"),
            "isomeric_smiles": data.get("isomeric_smiles", ""),
            "canonical_smiles": data.get("canonical_smiles", ""),
            "smiles": smiles,
            "iupac_name": data.get("iupac_name", ""),
            "molecular_formula": data.get("molecular_formula", ""),
            "molecular_weight": str(data.get("molecular_weight", "")),
        })

    if rows:
        df_chem = pd.DataFrame(rows)
        df_chem.to_sql("chemical_data", conn, if_exists="replace", index=False)
        found = sum(1 for r in rows if r["status"] == "found")
        print(f"  Loaded chemical_data: {len(rows)} entries ({found} found)")

    conn.commit()
    conn.close()
    print("Database ready.")


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn
