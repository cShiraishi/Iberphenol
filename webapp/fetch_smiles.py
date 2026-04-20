"""Fetches SMILES from PubChem for all CIDs in chemical_data and updates the DB."""
import sqlite3, time, urllib.request, json, ssl
from database import DB_PATH, init_db

# Bypass SSL verification (needed when behind a corporate proxy with self-signed cert)
_ctx = ssl.create_default_context()
_ctx.check_hostname = False
_ctx.verify_mode = ssl.CERT_NONE

BATCH = 100  # PubChem CID batch size

def fetch_smiles():
    init_db()
    conn = sqlite3.connect(DB_PATH)

    rows = conn.execute(
        "SELECT molecule_name, cid FROM chemical_data WHERE cid IS NOT NULL AND (smiles='' OR smiles IS NULL)"
    ).fetchall()

    print(f"Fetching SMILES for {len(rows)} compounds...")

    # Group into batches
    batches = [rows[i:i+BATCH] for i in range(0, len(rows), BATCH)]
    updated = 0

    for batch in batches:
        cids = [str(int(r[1])) for r in batch]
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/{','.join(cids)}/property/IsomericSMILES,CanonicalSMILES/JSON"
        try:
            with urllib.request.urlopen(url, timeout=15, context=_ctx) as resp:
                data = json.loads(resp.read())
            cid_map = {str(p["CID"]): p for p in data.get("PropertyTable", {}).get("Properties", [])}
            for name, cid in batch:
                cid_str = str(int(cid))
                if cid_str in cid_map:
                    props = cid_map[cid_str]
                    # PubChem returns "SMILES" (isomeric) and "ConnectivitySMILES" (canonical)
                    iso = props.get("IsomericSMILES") or props.get("SMILES") or ""
                    can = props.get("CanonicalSMILES") or props.get("ConnectivitySMILES") or ""
                    smiles = iso or can
                    if smiles:
                        conn.execute(
                            "UPDATE chemical_data SET smiles=?, isomeric_smiles=?, canonical_smiles=? WHERE molecule_name=?",
                            [smiles, iso, can, name]
                        )
                        updated += 1
            conn.commit()
            print(f"  Batch done ({updated} updated so far)…")
            time.sleep(0.3)
        except Exception as e:
            print(f"  Error fetching batch: {e}")
            time.sleep(2)

    print(f"Done. {updated} compounds now have SMILES.")
    conn.close()

if __name__ == "__main__":
    fetch_smiles()
