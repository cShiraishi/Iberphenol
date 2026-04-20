import sqlite3
import os

db_path = "webapp/phenoldb.sqlite"

# Map common locations in the DB to coordinates (Lat, Lng)
GEO_MAP = {
    "Braganca, Portugal": (41.8019, -6.7588),
    "Alfandega da Fe, Portugal": (41.3411, -6.9602),
    "Castelo Branco, Portugal": (39.8191, -7.4913),
    "Portalegre, Portugal": (39.2942, -7.4312),
    "Vila Real, Portugal": (41.2952, -7.7441),
    "Guarda, Portugal": (40.5365, -7.2683),
    "Ervital, Castro Daire, Portugal": (40.9016, -7.9547),
    "Miranda do Douro, Portugal": (41.4936, -6.2736),
    "Macedo de Cavaleiros, Portugal": (41.5379, -6.9534),
    "Vimioso, Portugal": (41.5833, -6.5283),
    "Spain": (40.4637, -3.7492),
    "Portugal": (39.3999, -8.2245),
}

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. Create locations table
    cursor.execute("DROP TABLE IF EXISTS locations")
    cursor.execute("""
        CREATE TABLE locations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            lat REAL,
            lng REAL
        )
    """)

    # 2. Extract unique origins from samples
    cursor.execute("SELECT DISTINCT origin FROM samples WHERE origin IS NOT NULL AND origin != ''")
    origins = cursor.fetchall()
    
    for (org,) in origins:
        # Simplify/clean name for matching
        clean = org.replace(u"\u00e7", "c").replace(u"\u00e2", "a").strip()
        coords = GEO_MAP.get(org, GEO_MAP.get(clean, (None, None)))
        
        cursor.execute("INSERT OR IGNORE INTO locations (name, lat, lng) VALUES (?,?,?)", (org, coords[0], coords[1]))

    # 3. Add location_id to samples
    try:
        cursor.execute("ALTER TABLE samples ADD COLUMN location_id INTEGER REFERENCES locations(id)")
    except:
        pass # Already exists
        
    cursor.execute("""
        UPDATE samples 
        SET location_id = (SELECT id FROM locations WHERE locations.name = samples.origin)
    """)

    conn.commit()
    conn.close()
    print "Location data structured and geocoded."
else:
    print "Database not found."
