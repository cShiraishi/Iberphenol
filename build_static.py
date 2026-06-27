"""
build_static.py — gera uma versao 100% estatica do PhenolDB para GitHub Pages.

Estrategia:
  - Usa o proprio backend FastAPI (via TestClient) para gerar JSONs identicos
    as respostas reais da API -> fidelidade total, sem reimplementar SQL.
  - Enriquece a lista de compostos com species_ids / plant_parts, para que a
    filtragem (hoje feita server-side em /api/compounds) funcione no navegador.
  - Copia o frontend + assets para docs/ e injeta um shim de fetch() que
    redireciona /api/... para os arquivos JSON estaticos.

Saida: docs/  (GitHub Pages -> branch master, pasta /docs)
"""
import json
import shutil
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).parent
WEBAPP = ROOT / "webapp"
STATIC_SRC = WEBAPP / "static"
DB_PATH = WEBAPP / "phenoldb.sqlite"

DOCS = ROOT / "docs"
DATA = DOCS / "data"

sys.path.insert(0, str(WEBAPP))
from fastapi.testclient import TestClient
from app import app


SHIM_JS = r"""// data-api-shim.js — intercepta fetch('/api/...') e serve os JSON estaticos.
// Gerado por build_static.py. Permite o PhenolDB rodar sem backend (GitHub Pages).
(function () {
  const orig = window.fetch.bind(window);
  let _all = null;

  function abs(rel) { return new URL(rel, document.baseURI).href; }
  function jsonResp(obj) {
    return new Response(JSON.stringify(obj), {
      status: 200, headers: { "Content-Type": "application/json" },
    });
  }
  function like(hay, needle) {
    return String(hay || "").toLowerCase().includes(needle.toLowerCase());
  }
  async function allCompounds() {
    if (!_all) _all = await (await orig(abs("data/compounds.json"))).json();
    return _all;
  }

  const DIRECT = {
    "/api/species": "data/species.json",
    "/api/classes": "data/classes.json",
    "/api/plant_parts": "data/plant_parts.json",
    "/api/stats": "data/stats.json",
    "/api/stats/detailed": "data/stats_detailed.json",
    "/api/locations": "data/locations.json",
    "/api/graph/species": "data/graph_species.json",
    "/api/export": "data/phenoldb_export.csv",
  };

  window.fetch = async function (input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    const i = url.indexOf("/api/");
    if (i === -1) return orig(input, init);

    const rest = url.slice(i);
    const q = rest.indexOf("?");
    const path = q >= 0 ? rest.slice(0, q) : rest;
    const qs = new URLSearchParams(q >= 0 ? rest.slice(q + 1) : "");

    if (DIRECT[path]) return orig(abs(DIRECT[path]), init);

    let m = path.match(/^\/api\/compounds\/(\d+)$/);
    if (m) return orig(abs("data/compounds/" + m[1] + ".json"), init);

    m = path.match(/^\/api\/species\/(\d+)\/compounds$/);
    if (m) return orig(abs("data/species_compounds/" + m[1] + ".json"), init);

    if (path === "/api/compounds") {
      const all = await allCompounds();
      const search = (qs.get("search") || "").trim();
      const speciesId = qs.get("species_id");
      const cls = qs.get("compound_class") || "";
      const part = qs.get("plant_part") || "";
      const hasStruct = qs.get("has_structure") === "true";
      const page = parseInt(qs.get("page") || "1", 10);
      const limit = parseInt(qs.get("limit") || "50", 10);

      const filtered = all.filter((c) => {
        if (search && !(
          like(c.molecule_name, search) || like(c.compound_class, search) ||
          like(c.subclass, search) || like(c.iupac_name, search) ||
          like(c.molecular_formula, search))) return false;
        if (speciesId && !(c.species_ids || []).includes(parseInt(speciesId, 10))) return false;
        if (cls && c.compound_class !== cls) return false;
        if (part && !(c.plant_parts || []).some((pp) => like(pp, part))) return false;
        if (hasStruct && !(c.smiles && c.status === "found")) return false;
        return true;
      });
      const offset = (page - 1) * limit;
      return jsonResp({
        total: filtered.length, page, limit,
        compounds: filtered.slice(offset, offset + limit),
      });
    }

    return orig(input, init);
  };
})();
"""


def write_json(rel_path, obj):
    p = DATA / rel_path
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")


def enrichment_maps():
    """species_ids e plant_parts por compound.id, lidos direto do SQLite,
    replicando os joins usados pelo filtro de /api/compounds."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    species_by_cid, parts_by_cid = {}, {}
    rows = conn.execute(
        """SELECT DISTINCT c.id AS cid, sa.species_id AS sid, sa.plant_part AS pp
           FROM compounds c
           JOIN measurements m ON c.id = m.compound_id
           JOIN samples sa ON m.sample_id = sa.id"""
    ).fetchall()
    for r in rows:
        if r["sid"] is not None:
            species_by_cid.setdefault(r["cid"], set()).add(r["sid"])
        if r["pp"]:
            parts_by_cid.setdefault(r["cid"], set()).add(r["pp"])
    conn.close()
    return species_by_cid, parts_by_cid


def main():
    if DOCS.exists():
        # limpa apenas o que o build gera (preserva nada extra aqui; docs e dedicada)
        shutil.rmtree(DOCS)
    DATA.mkdir(parents=True, exist_ok=True)

    with TestClient(app) as client:
        # --- endpoints simples (resposta direta) -----------------------------
        simple = {
            "species.json": "/api/species",
            "classes.json": "/api/classes",
            "plant_parts.json": "/api/plant_parts",
            "stats.json": "/api/stats",
            "stats_detailed.json": "/api/stats/detailed",
            "locations.json": "/api/locations",
            "graph_species.json": "/api/graph/species",
        }
        for fname, url in simple.items():
            r = client.get(url)
            r.raise_for_status()
            write_json(fname, r.json())
            print(f"  ok {url}")

        # --- lista completa de compostos (paginando a propria API) -----------
        species_by_cid, parts_by_cid = enrichment_maps()
        all_compounds = []
        page = 1
        while True:
            r = client.get(f"/api/compounds?page={page}&limit=200")
            r.raise_for_status()
            data = r.json()
            batch = data["compounds"]
            if not batch:
                break
            all_compounds.extend(batch)
            if len(all_compounds) >= data["total"]:
                break
            page += 1
        for c in all_compounds:
            c["species_ids"] = sorted(species_by_cid.get(c["id"], []))
            c["plant_parts"] = sorted(parts_by_cid.get(c["id"], []))
        write_json("compounds.json", all_compounds)
        print(f"  ok /api/compounds -> {len(all_compounds)} compostos")

        # --- detalhe por composto -------------------------------------------
        for c in all_compounds:
            r = client.get(f"/api/compounds/{c['id']}")
            if r.status_code == 200:
                write_json(f"compounds/{c['id']}.json", r.json())
        print(f"  ok detalhes de {len(all_compounds)} compostos")

        # --- compostos por especie ------------------------------------------
        species = client.get("/api/species").json()
        for s in species:
            r = client.get(f"/api/species/{s['id']}/compounds")
            if r.status_code == 200:
                write_json(f"species_compounds/{s['id']}.json", r.json())
        print(f"  ok compostos de {len(species)} especies")

        # --- export CSV completo --------------------------------------------
        r = client.get("/api/export")
        if r.status_code == 200:
            (DATA / "phenoldb_export.csv").write_bytes(r.content)
            print("  ok /api/export -> phenoldb_export.csv")

    # --- frontend + assets --------------------------------------------------
    shutil.copytree(STATIC_SRC, DOCS / "static")
    # nojekyll para o GitHub Pages servir pastas/_arquivos sem processar Jekyll
    (DOCS / ".nojekyll").write_text("", encoding="utf-8")

    # index.html transformado
    html = (STATIC_SRC / "index.html").read_text(encoding="utf-8")
    html = html.replace("/static/", "static/")  # paths relativos (project pages)
    html = html.replace(
        "return '/api/export?' + p.toString();",
        "return 'data/phenoldb_export.csv';",
    )
    # injeta o shim antes dos demais scripts
    shim_tag = '<script src="data-api-shim.js"></script>\n'
    html = html.replace("<head>", "<head>\n" + shim_tag, 1)
    (DOCS / "index.html").write_text(html, encoding="utf-8")
    (DOCS / "data-api-shim.js").write_text(SHIM_JS, encoding="utf-8")
    print("  ok index.html (paths relativos + shim + export estatico)")

    print(f"\nPronto. Site estatico em: {DOCS}")


if __name__ == "__main__":
    main()
