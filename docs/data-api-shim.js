// data-api-shim.js — intercepta fetch('/api/...') e serve os JSON estaticos.
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
