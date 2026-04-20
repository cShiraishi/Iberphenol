import sqlite3
import os
import re

db_path = "webapp/phenoldb.sqlite"

def normalize_name(name):
    if not name: return ""
    # Remove extra spaces inside parentheses: ( + ) -> (+)
    name = re.sub(r'\(\s*\+\s*\)', '(+)', name)
    name = re.sub(r'\(\s*\-\s*\)', '(-)', name)
    # Remove double spaces
    name = " ".join(name.split())
    return name.strip()

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print "Starting deduplication process..."

    # 1. Fetch all compounds
    cursor.execute("SELECT id, molecule_name, compound_class, subclass FROM compounds")
    all_compounds = cursor.fetchall()
    
    unique_compounds = {} # normalized_name -> {canonical_id, class, subclass}
    id_mapping = {} # old_id -> new_id

    for cid, name, cls, sub in all_compounds:
        norm = normalize_name(name)
        if norm not in unique_compounds:
            unique_compounds[norm] = {'id': cid, 'class': cls, 'sub': sub}
            id_mapping[cid] = cid
        else:
            id_mapping[cid] = unique_compounds[norm]['id']

    print "Found " + str(len(all_compounds)) + " rows. Unified into " + str(len(unique_compounds)) + " unique compounds."

    # 2. Update linked tables
    print "Updating measurements..."
    for old_id, new_id in id_mapping.items():
        if old_id != new_id:
            cursor.execute("UPDATE measurements SET compound_id = ? WHERE compound_id = ?", (new_id, old_id))
    
    print "Updating spectral_data..."
    for old_id, new_id in id_mapping.items():
        if old_id != new_id:
            # Check if target already has spectral data to avoid conflicts
            cursor.execute("SELECT 1 FROM spectral_data WHERE compound_id = ?", (new_id,))
            exists = cursor.fetchone()
            if not exists:
                cursor.execute("UPDATE spectral_data SET compound_id = ? WHERE compound_id = ?", (new_id, old_id))
            else:
                # If both have spectral data, we might lose one, but usually it's the same info.
                # For safety, just delete the duplicate spectral entry.
                cursor.execute("DELETE FROM spectral_data WHERE compound_id = ?", (old_id,))

    # 3. Rebuild compounds table
    print "Rebuilding compounds table..."
    cursor.execute("CREATE TABLE compounds_new (id INTEGER PRIMARY KEY, compound_class TEXT, subclass TEXT, molecule_name TEXT)")
    
    for norm, data in unique_compounds.items():
        cursor.execute("INSERT INTO compounds_new (id, compound_class, subclass, molecule_name) VALUES (?,?,?,?)",
                       (data['id'], data['class'], data['sub'], norm))
    
    cursor.execute("DROP TABLE compounds")
    cursor.execute("ALTER TABLE compounds_new RENAME TO compounds")

    conn.commit()
    conn.close()
    print "Deduplication complete."
else:
    print "Database not found."
