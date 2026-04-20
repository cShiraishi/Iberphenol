import sqlite3
import json
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent.parent
DB_PATH = Path(__file__).parent / "phenoldb.sqlite"
EXCEL_PATH = BASE / "PhenolDB_estruturado.xlsx"
CASCADE_CACHE = BASE / "cascade_cache.json"
PUBCHEM_CACHE = BASE / "pubchem_cache.json"


def init_db(force: bool = False):
    if DB_PATH.exists() and not force:
        return

    print("Initializing database from Excel...")
    conn = sqlite3.connect(DB_PATH)

    xl = pd.ExcelFile(EXCEL_PATH)
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet)
            # Sanitize column names: replace non-ASCII and special chars
            df.columns = [
                str(c).encode("ascii", "replace").decode("ascii")
                .replace("?", "").replace(" ", "_").replace("/", "_")
                .replace("-", "_").replace("(", "").replace(")", "")
                .strip("_") or f"col_{i}"
                for i, c in enumerate(df.columns)
            ]
            # Sanitize string values
            for col in df.select_dtypes(include="object").columns:
                df[col] = df[col].apply(
                    lambda x: x.encode("utf-8", "replace").decode("utf-8") if isinstance(x, str) else x
                )
            safe_name = sheet.lower().replace(" ", "_").replace("-", "_")
            df.to_sql(safe_name, conn, if_exists="replace", index=False)
            print(f"  Loaded sheet '{sheet}' -> table '{safe_name}' ({len(df)} rows)")
        except Exception as e:
            print(f"  Warning: could not load sheet '{sheet}': {e}")

    # Build unified chemical_data table from JSON caches
    chem: dict = {}
    if PUBCHEM_CACHE.exists():
        with open(PUBCHEM_CACHE, encoding="utf-8") as f:
            chem.update(json.load(f))
    if CASCADE_CACHE.exists():
        with open(CASCADE_CACHE, encoding="utf-8") as f:
            chem.update(json.load(f))  # cascade overrides pubchem (more complete)

    rows = []
    for name, data in chem.items():
        smiles = data.get("isomeric_smiles") or data.get("canonical_smiles") or ""
        rows.append(
            {
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
            }
        )

    if rows:
        df_chem = pd.DataFrame(rows)
        df_chem.to_sql("chemical_data", conn, if_exists="replace", index=False)
        found = sum(1 for r in rows if r["status"] == "found")
        print(f"  Loaded chemical_data: {len(rows)} entries ({found} found)")

    conn.commit()
    conn.close()
    print("Database ready.")


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn
