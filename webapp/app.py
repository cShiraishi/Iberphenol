import os
import io
import csv
import sys
import sqlite3
from pathlib import Path

# Add the current directory to sys.path to allow importing 'database'
sys.path.append(os.path.dirname(__file__))

from database import init_db, get_db

from contextlib import asynccontextmanager
from fastapi import FastAPI, Query, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse


STATIC_DIR = Path(__file__).parent / "static"


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(title="PhenolDB Research Platform", version="1.0", lifespan=lifespan)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
async def root():
    return FileResponse(STATIC_DIR / "index.html")


# ── helpers ──────────────────────────────────────────────────────────────────

def _rows(conn, sql, params=None):
    cur = conn.execute(sql, params or [])
    return [dict(r) for r in cur.fetchall()]


def _scalar(conn, sql, params=None):
    cur = conn.execute(sql, params or [])
    row = cur.fetchone()
    return row[0] if row else 0


def _table_exists(conn, name):
    r = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", [name]
    ).fetchone()
    return r is not None


def _col_exists(conn, table, col):
    info = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    return any(r[1] == col for r in info)


# ── API ───────────────────────────────────────────────────────────────────────

@app.get("/api/stats")
async def get_stats():
    conn = get_db()
    tables = ["compounds", "species", "measurements", "spectral_data", "methods"]
    stats = {}
    for t in tables:
        try:
            stats[t] = _scalar(conn, f'SELECT COUNT(*) FROM "{t}"')
        except Exception:
            stats[t] = 0

    try:
        stats["compounds_with_structure"] = _scalar(
            conn,
            "SELECT COUNT(*) FROM chemical_data WHERE status='found' AND smiles != '' AND smiles IS NOT NULL",
        )
        stats["compounds_not_found"] = _scalar(
            conn, "SELECT COUNT(*) FROM chemical_data WHERE status='not_found'"
        )
    except Exception:
        stats["compounds_with_structure"] = 0
        stats["compounds_not_found"] = 0

    conn.close()
    return stats


@app.get("/api/stats/detailed")
async def get_stats_detailed():
    conn = get_db()

    class_dist = _rows(conn, """
        SELECT compound_class, COUNT(*) as count
        FROM compounds
        WHERE compound_class IS NOT NULL AND compound_class != ''
        GROUP BY compound_class ORDER BY count DESC LIMIT 15
    """)

    try:
        plant_parts = _rows(conn, """
            SELECT sa.plant_part,
                   COUNT(DISTINCT m.compound_id) as compound_count,
                   COUNT(m.id) as measurement_count
            FROM samples sa
            JOIN measurements m ON m.sample_id = sa.id
            WHERE sa.plant_part IS NOT NULL AND sa.plant_part != ''
            GROUP BY sa.plant_part ORDER BY compound_count DESC LIMIT 12
        """)
    except Exception:
        plant_parts = []

    top_species = _rows(conn, """
        SELECT sp.scientific_name, COUNT(DISTINCT c.id) as compound_count
        FROM species sp
        JOIN compounds c ON c.species_id = sp.id
        GROUP BY sp.id ORDER BY compound_count DESC LIMIT 10
    """)

    try:
        subclass_dist = _rows(conn, """
            SELECT subclass, compound_class, COUNT(*) as count
            FROM compounds
            WHERE subclass IS NOT NULL AND subclass != ''
            GROUP BY subclass ORDER BY count DESC LIMIT 15
        """)
    except Exception:
        subclass_dist = []

    try:
        compounds_with_measurements = _scalar(conn,
            "SELECT COUNT(DISTINCT compound_id) FROM measurements")
    except Exception:
        compounds_with_measurements = 0

    conn.close()
    return {
        "class_distribution": class_dist,
        "plant_parts": plant_parts,
        "top_species": top_species,
        "subclass_distribution": subclass_dist,
        "compounds_with_measurements": compounds_with_measurements,
    }


@app.get("/api/graph/species")
async def get_species_graph():
    conn = get_db()

    species_totals = _rows(conn, """
        SELECT sp.id, sp.scientific_name,
               COUNT(DISTINCT c.id) as compound_count,
               COUNT(DISTINCT c.compound_class) as class_count
        FROM species sp
        JOIN compounds c ON c.species_id = sp.id
        GROUP BY sp.id ORDER BY compound_count DESC
    """)

    species_class_edges = _rows(conn, """
        SELECT sp.id as species_id, sp.scientific_name,
               c.compound_class, COUNT(DISTINCT c.id) as count
        FROM species sp
        JOIN compounds c ON c.species_id = sp.id
        WHERE c.compound_class IS NOT NULL AND c.compound_class != ''
        GROUP BY sp.id, c.compound_class
    """)

    try:
        species_overlap = _rows(conn, """
            SELECT a.species_id as s1, b.species_id as s2,
                   sa.scientific_name as name1, sb.scientific_name as name2,
                   COUNT(DISTINCT a.molecule_name) as shared_compounds
            FROM (SELECT species_id, molecule_name FROM compounds) a
            JOIN (SELECT species_id, molecule_name FROM compounds) b
              ON a.molecule_name = b.molecule_name AND a.species_id < b.species_id
            JOIN species sa ON sa.id = a.species_id
            JOIN species sb ON sb.id = b.species_id
            GROUP BY a.species_id, b.species_id
            HAVING shared_compounds >= 1
            ORDER BY shared_compounds DESC LIMIT 60
        """)
    except Exception:
        species_overlap = []

    conn.close()
    return {
        "species": species_totals,
        "species_class_edges": species_class_edges,
        "species_overlap": species_overlap,
    }


@app.get("/api/species")
async def get_species():
    from collections import defaultdict
    conn = get_db()
    species = _rows(conn, 'SELECT * FROM species ORDER BY scientific_name')

    origins_rows = _rows(conn, """
        SELECT sa.species_id, sa.origin, COUNT(*) as cnt
        FROM samples sa
        WHERE sa.origin IS NOT NULL AND sa.origin != ''
        GROUP BY sa.species_id, sa.origin
        ORDER BY sa.species_id, cnt DESC
    """)

    origins_by_species = defaultdict(list)
    for r in origins_rows:
        if len(origins_by_species[r['species_id']]) < 4:
            origins_by_species[r['species_id']].append(r['origin'])

    for s in species:
        s['origins'] = origins_by_species.get(s['id'], [])

    conn.close()
    return species


@app.get("/api/classes")
async def get_classes():
    conn = get_db()
    rows = _rows(
        conn,
        """SELECT DISTINCT compound_class, subclass
           FROM compounds
           WHERE compound_class IS NOT NULL AND compound_class != ''
           ORDER BY compound_class, subclass""",
    )
    conn.close()
    return rows


@app.get("/api/plant_parts")
async def get_plant_parts():
    conn = get_db()
    rows = _rows(
        conn,
        "SELECT DISTINCT plant_part FROM samples WHERE plant_part IS NOT NULL AND plant_part != '' ORDER BY plant_part",
    )
    conn.close()
    return [r["plant_part"] for r in rows]


@app.get("/api/compounds")
async def get_compounds(
    search: str = Query(default=""),
    species_id: int = Query(default=None),
    compound_class: str = Query(default=""),
    plant_part: str = Query(default=""),
    has_structure: bool = Query(default=False),
    page: int = Query(default=1, ge=1),
    limit: int = Query(default=50, le=200),
):
    conn = get_db()
    conds, params = [], []

    if search:
        conds.append(
            "(c.molecule_name LIKE ? OR c.compound_class LIKE ? OR c.subclass LIKE ? "
            "OR cd.iupac_name LIKE ? OR cd.molecular_formula LIKE ?)"
        )
        params.extend([f"%{search}%"] * 5)
    if species_id:
        conds.append("c.species_id = ?")
        params.append(species_id)
    if compound_class:
        conds.append("c.compound_class = ?")
        params.append(compound_class)
    if plant_part:
        conds.append("sa.plant_part LIKE ?")
        params.append(f"%{plant_part}%")
    if has_structure:
        conds.append("cd.smiles != '' AND cd.smiles IS NOT NULL AND cd.status = 'found'")

    where = ("WHERE " + " AND ".join(conds)) if conds else ""
    offset = (page - 1) * limit

    base = f"""
        FROM compounds c
        LEFT JOIN species sp ON c.species_id = sp.id
        LEFT JOIN chemical_data cd ON c.molecule_name = cd.molecule_name
        LEFT JOIN measurements m ON c.id = m.compound_id
        LEFT JOIN samples sa ON m.sample_id = sa.id
        {where}
    """

    total = _scalar(conn, f"SELECT COUNT(DISTINCT c.id) {base}", params)
    rows = _rows(
        conn,
        f"""SELECT DISTINCT c.id, c.species_id, c.compound_class, c.subclass, c.molecule_name,
                   sp.scientific_name,
                   cd.status, cd.cid, cd.smiles, cd.iupac_name,
                   cd.molecular_formula, cd.molecular_weight, cd.source
            {base}
            ORDER BY c.compound_class, c.molecule_name
            LIMIT ? OFFSET ?""",
        params + [limit, offset],
    )

    conn.close()
    return {"total": total, "page": page, "limit": limit, "compounds": rows}


@app.get("/api/compounds/{compound_id}")
async def get_compound_detail(compound_id: int):
    conn = get_db()

    compound = conn.execute(
        """SELECT c.*, sp.scientific_name,
                  cd.status, cd.cid, cd.smiles, cd.isomeric_smiles, cd.canonical_smiles,
                  cd.iupac_name, cd.molecular_formula, cd.molecular_weight, cd.source
           FROM compounds c
           LEFT JOIN species sp ON c.species_id = sp.id
           LEFT JOIN chemical_data cd ON c.molecule_name = cd.molecule_name
           WHERE c.id = ?""",
        [compound_id],
    ).fetchone()

    if not compound:
        raise HTTPException(status_code=404, detail="Compound not found")

    spectral = conn.execute(
        "SELECT * FROM spectral_data WHERE compound_id = ?", [compound_id]
    ).fetchone()

    # Determine references table name
    ref_table = "references" if _table_exists(conn, "references") else None

    if ref_table:
        measurements = _rows(
            conn,
            f"""SELECT m.*, sa.plant_part, sa.origin, sa.cultivar_variety, sa.season,
                       sp.scientific_name,
                       me.equipment, me.data_acquisition, me.ionization_mode,
                       r.citation
                FROM measurements m
                LEFT JOIN samples sa ON m.sample_id = sa.id
                LEFT JOIN species sp ON sa.species_id = sp.id
                LEFT JOIN methods me ON m.method_id = me.id
                LEFT JOIN "{ref_table}" r ON m.reference_id = r.id
                WHERE m.compound_id = ?
                ORDER BY CAST(m.concentration_value AS REAL) DESC""",
            [compound_id],
        )
    else:
        measurements = _rows(
            conn,
            """SELECT m.*, sa.plant_part, sa.origin, sa.cultivar_variety, sa.season,
                      sp.scientific_name,
                      me.equipment, me.data_acquisition, me.ionization_mode
               FROM measurements m
               LEFT JOIN samples sa ON m.sample_id = sa.id
               LEFT JOIN species sp ON sa.species_id = sp.id
               LEFT JOIN methods me ON m.method_id = me.id
               WHERE m.compound_id = ?
               ORDER BY CAST(m.concentration_value AS REAL) DESC""",
            [compound_id],
        )

    conn.close()
    return {
        "compound": dict(compound),
        "spectral": dict(spectral) if spectral else None,
        "measurements": measurements,
    }


@app.get("/api/species/{species_id}/compounds")
async def get_species_compounds(species_id: int):
    conn = get_db()
    rows = _rows(
        conn,
        """SELECT c.*, cd.smiles, cd.molecular_formula, cd.molecular_weight, cd.cid,
                  COUNT(m.id) as measurement_count,
                  AVG(CAST(m.concentration_value AS REAL)) as avg_concentration,
                  MAX(CAST(m.concentration_value AS REAL)) as max_concentration,
                  m.concentration_unit
           FROM compounds c
           LEFT JOIN chemical_data cd ON c.molecule_name = cd.molecule_name
           LEFT JOIN measurements m ON c.id = m.compound_id
           WHERE c.species_id = ?
           GROUP BY c.id
           ORDER BY c.compound_class, c.molecule_name""",
        [species_id],
    )
    conn.close()
    return rows


_GEOCODE = {
    'bragança': (41.8063, -6.7588, 'Bragança'),
    'braganca': (41.8063, -6.7588, 'Bragança'),
    'montesinho': (41.9583, -6.8500, 'Bragança'),
    'alfândega da fé': (41.3479, -6.9573, 'Alfândega da Fé'),
    'alfandega da fe': (41.3479, -6.9573, 'Alfândega da Fé'),
    'miranda do douro': (41.4953, -6.2769, 'Miranda do Douro'),
    'macedo de cavaleiros': (41.5307, -6.9591, 'Macedo de Cavaleiros'),
    'mirandela': (41.4879, -7.1745, 'Mirandela'),
    'valpaços': (41.6087, -7.2988, 'Valpaços'),
    'valpacos': (41.6087, -7.2988, 'Valpaços'),
    'mogadouro': (41.3379, -6.7157, 'Mogadouro'),
    'vinhais': (41.8333, -7.0000, 'Vinhais'),
    'vila real': (41.3004, -7.7457, 'Vila Real'),
    'alijó': (41.2797, -7.4633, 'Alijó'),
    'alijo': (41.2797, -7.4633, 'Alijó'),
    'figueira de castelo rodrigo': (40.8990, -6.9601, 'Figueira de Castelo Rodrigo'),
    'trás-os-montes': (41.6000, -7.0000, 'Trás-os-Montes'),
    'tras-os-montes': (41.6000, -7.0000, 'Trás-os-Montes'),
    'castro daire': (40.9028, -7.9366, 'Castro Daire'),
    'cinfães': (41.0712, -8.0894, 'Cinfães'),
    'cinfaes': (41.0712, -8.0894, 'Cinfães'),
    'vila nova de gaia': (41.1239, -8.6148, 'Vila Nova de Gaia'),
    'porto': (41.1579, -8.6291, 'Porto'),
    'riba de mouro': (41.9178, -8.2028, 'Riba de Mouro'),
    'minho': (41.8000, -8.3000, 'Minho'),
    'viseu': (40.6566, -7.9120, 'Viseu'),
    'serra da estrela': (40.3500, -7.6167, 'Serra da Estrela'),
    'penacova': (40.2745, -8.2818, 'Penacova'),
    'lousã': (40.1075, -8.2452, 'Lousã'),
    'lousa': (40.1075, -8.2452, 'Lousã'),
    'fundão': (40.1394, -7.5010, 'Fundão'),
    'fundao': (40.1394, -7.5010, 'Fundão'),
    'castelo branco': (39.8239, -7.4912, 'Castelo Branco'),
    'oleiros': (41.6800, -7.1700, 'Oleiros'),
    'alcanede': (39.4147, -8.6636, 'Alcanede'),
    'loures': (38.8298, -9.1686, 'Loures'),
    'oeiras': (38.6972, -9.3039, 'Oeiras'),
    'sintra': (38.8029, -9.3817, 'Sintra'),
    'lisboa': (38.7169, -9.1399, 'Lisboa'),
    'lisbon': (38.7169, -9.1399, 'Lisboa'),
    'azores': (37.7412, -25.6756, 'Açores'),
    'furnas': (37.7676, -25.3174, 'Açores'),
    'são miguel': (37.7746, -25.4644, 'Açores'),
    'sao miguel': (37.7746, -25.4644, 'Açores'),
    'zaragoza': (41.6561, -0.8773, 'Zaragoza'),
    'soria': (41.7640, -2.4650, 'Soria'),
    'murcia': (37.9922, -1.1307, 'Murcia'),
    'santomera': (38.0667, -1.0500, 'Murcia'),
    'valencia': (39.4699, -0.3763, 'Valencia'),
    'spain': (40.4168, -3.7038, 'Espanha'),
    'españa': (40.4168, -3.7038, 'Espanha'),
    'chainat': (15.1835, 100.1265, 'Chainat, Thailand'),
    'thailand': (15.8700, 100.9925, 'Thailand'),
    'thessaly': (39.5500, 22.1167, 'Tessália, Grécia'),
    'greece': (39.0742, 21.8243, 'Grécia'),
    'australia': (-25.2744, 133.7751, 'Austrália'),
}


def _geocode_origin(origin: str):
    o = origin.lower()
    best_key = None
    for key in _GEOCODE:
        if key in o:
            if best_key is None or len(key) > len(best_key):
                best_key = key
    return _GEOCODE[best_key] if best_key else None


@app.get("/api/map")
async def get_map_data():
    conn = get_db()
    rows = _rows(conn, """
        SELECT sa.origin, sp.scientific_name, c.id as compound_id,
               c.molecule_name, c.compound_class
        FROM samples sa
        JOIN measurements m ON m.sample_id = sa.id
        JOIN compounds c ON c.id = m.compound_id
        JOIN species sp ON c.species_id = sp.id
        WHERE sa.origin IS NOT NULL AND sa.origin != ''
    """)
    conn.close()

    from collections import defaultdict
    locs = defaultdict(lambda: {'compounds': set(), 'species': set()})

    for row in rows:
        geo = _geocode_origin(row['origin'])
        if not geo:
            continue
        key = geo[2]
        locs[key]['lat'] = geo[0]
        locs[key]['lon'] = geo[1]
        locs[key]['name'] = geo[2]
        locs[key]['compounds'].add(row['compound_id'])
        locs[key]['species'].add(row['scientific_name'])

    return [
        {
            'name': v['name'],
            'lat': v['lat'],
            'lon': v['lon'],
            'compound_count': len(v['compounds']),
            'species': sorted(v['species']),
            'compound_ids': sorted(v['compounds']),
        }
        for v in sorted(locs.values(), key=lambda x: -len(x['compounds']))
        if 'lat' in v
    ]


@app.get("/api/map/compounds")
async def get_map_compounds(location: str = Query(...)):
    conn = get_db()
    rows = _rows(conn, """
        SELECT DISTINCT c.id, c.molecule_name, c.compound_class, c.subclass,
               sp.scientific_name, cd.molecular_formula, cd.cid,
               sa.plant_part, sa.origin
        FROM samples sa
        JOIN measurements m ON m.sample_id = sa.id
        JOIN compounds c ON c.id = m.compound_id
        JOIN species sp ON c.species_id = sp.id
        LEFT JOIN chemical_data cd ON c.molecule_name = cd.molecule_name
        WHERE sa.origin IS NOT NULL
        ORDER BY c.compound_class, c.molecule_name
    """)
    result = []
    seen = set()
    for row in rows:
        geo = _geocode_origin(row['origin'])
        if geo and geo[2] == location and row['id'] not in seen:
            seen.add(row['id'])
            result.append(dict(row))
    conn.close()
    return result


@app.get("/api/export")
async def export_compounds(
    search: str = Query(default=""),
    species_id: int = Query(default=None),
    compound_class: str = Query(default=""),
):
    conn = get_db()
    conds, params = [], []

    if search:
        conds.append("c.molecule_name LIKE ?")
        params.append(f"%{search}%")
    if species_id:
        conds.append("c.species_id = ?")
        params.append(species_id)
    if compound_class:
        conds.append("c.compound_class = ?")
        params.append(compound_class)

    where = ("WHERE " + " AND ".join(conds)) if conds else ""

    rows = _rows(
        conn,
        f"""SELECT c.molecule_name, c.compound_class, c.subclass, sp.scientific_name,
                   cd.molecular_formula, cd.molecular_weight, cd.iupac_name,
                   cd.smiles, cd.cid,
                   sd.lambda_max_nm, sd.exact_mass, sd.ms_fragments,
                   m.concentration_value, m.concentration_unit,
                   sa.plant_part, sa.origin, sa.season
            FROM compounds c
            LEFT JOIN species sp ON c.species_id = sp.id
            LEFT JOIN chemical_data cd ON c.molecule_name = cd.molecule_name
            LEFT JOIN spectral_data sd ON sd.compound_id = c.id
            LEFT JOIN measurements m ON m.compound_id = c.id
            LEFT JOIN samples sa ON m.sample_id = sa.id
            {where}
            ORDER BY c.compound_class, c.molecule_name""",
        params,
    )
    conn.close()

    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(
        [
            "Molecule Name", "Class", "Subclass", "Species",
            "Molecular Formula", "Molecular Weight", "IUPAC Name",
            "SMILES", "PubChem CID",
            "λmax (nm)", "Exact Mass", "MS Fragments",
            "Concentration Value", "Concentration Unit",
            "Plant Part", "Origin", "Season",
        ]
    )
    for r in rows:
        w.writerow(list(r.values()))

    output.seek(0)
    return StreamingResponse(
        iter([output.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=phenoldb_export.csv"},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
