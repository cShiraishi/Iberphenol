-- IberPhenol v2 — esquema relacional.
-- Unidade central: o pico anotado (peak). Identificação (peak) e quantificação
-- (measurement) ficam separadas: um pico pode existir sem medição.
-- Toda linha derivada do Excel bruto carrega src_sheet + src_cell (notação A1) e um
-- código de evidência (REP, CUR, AUT, CALC, AUT-NR; ver release.evidence_codes).

PRAGMA foreign_keys = ON;

CREATE TABLE release (
    key   TEXT PRIMARY KEY,
    value TEXT
);

-- Abas do Excel bruto e o que o ETL fez com cada uma.
CREATE TABLE sheet (
    sheet_id        INTEGER PRIMARY KEY,
    name_verbatim   TEXT NOT NULL,
    status          TEXT NOT NULL,   -- parsed | template_empty | duplicate | peak_table_unassigned | skipped
    layout          TEXT,            -- standard | name_in_conc_col | class_totals | peak_table
    duplicate_of    INTEGER REFERENCES sheet(sheet_id),
    metadata_mode   TEXT,            -- sheet (metadados valem para a aba) | row (variam por linha)
    n_peaks         INTEGER DEFAULT 0,
    n_measurements  INTEGER DEFAULT 0,
    suggested_sheet INTEGER REFERENCES sheet(sheet_id),  -- só para tabelas de picos sem contexto
    suggestion_score REAL,
    notes           TEXT
);

-- Artigos de origem (ou textos que estavam na coluna Reference).
CREATE TABLE study (
    study_id      INTEGER PRIMARY KEY,
    citation      TEXT NOT NULL UNIQUE,
    is_citation   INTEGER NOT NULL,  -- 0 = texto na coluna Reference que não é citação
    year          INTEGER,
    doi           TEXT,              -- só quando escrito na própria citação; resto via Crossref
    src_sheet     TEXT,
    src_cell      TEXT
);

-- Nome do organismo como está, com uma leitura inicial; a resolução em WFO/NCBI é uma etapa posterior.
CREATE TABLE taxon (
    taxon_id        INTEGER PRIMARY KEY,
    binomial        TEXT,            -- "Genus epithet" (ou só gênero)
    genus           TEXT,
    epithet         TEXT,
    is_common_name  INTEGER DEFAULT 0,
    needs_review    INTEGER DEFAULT 0,
    note            TEXT,
    UNIQUE (binomial)
);

CREATE TABLE method (
    method_id         INTEGER PRIMARY KEY,
    equipment         TEXT,
    data_acquisition  TEXT,
    ionization        TEXT,
    UNIQUE (equipment, data_acquisition, ionization)
);

-- Amostra = aba (modo sheet) ou aba × contexto de linha (modo row: cultivar/origem/estação variam).
CREATE TABLE sample (
    sample_id         INTEGER PRIMARY KEY,
    sheet_id          INTEGER NOT NULL REFERENCES sheet(sheet_id),
    taxon_id          INTEGER REFERENCES taxon(taxon_id),
    name_verbatim     TEXT,            -- nome da aba
    qualifier         TEXT,            -- órgão, clone, preparo ou cópia tirados do nome da aba
    plant_part        TEXT,
    cultivar          TEXT,
    origin            TEXT,
    season            TEXT,
    product_category  TEXT,
    method_id         INTEGER REFERENCES method(method_id),
    context_inferred  INTEGER DEFAULT 0,  -- 1 = algum metadado veio de linha anterior (carry-down)
    src_sheet         TEXT,
    src_cell          TEXT
);

CREATE TABLE sample_study (
    sample_id  INTEGER NOT NULL REFERENCES sample(sample_id),
    study_id   INTEGER NOT NULL REFERENCES study(study_id),
    PRIMARY KEY (sample_id, study_id)
);

-- Uma condição por coluna de concentração: extrato, preparo, sub-amostra, modo de quantificação.
CREATE TABLE condition (
    condition_id    INTEGER PRIMARY KEY,
    sheet_id        INTEGER NOT NULL REFERENCES sheet(sheet_id),
    col_letter      TEXT NOT NULL,
    quant_approach  TEXT,     -- Reference standards (DAD) | Reference standards (MS) | Internal standard | Extinction | Peak area | Concentration
    header_label    TEXT,     -- texto extra no cabeçalho da coluna
    row3_label      TEXT,     -- rótulo na linha 3 quando ela não tem molécula (ex.: BB1)
    preparation     TEXT,     -- palavras de preparo mais frequentes nas células
    tag             TEXT,     -- marcador mais frequente nas células, ex. (a)
    tag_resolved    TEXT,     -- marcador traduzido pelo campo de cultivar, ex. a -> judia
    label           TEXT,     -- rótulo final legível
    UNIQUE (sheet_id, col_letter)
);

CREATE TABLE compound (
    compound_id        TEXT PRIMARY KEY,  -- IPH-C00001
    name_preferred     TEXT NOT NULL,
    name_key           TEXT NOT NULL UNIQUE,
    n_peaks            INTEGER,
    n_samples          INTEGER,
    class_mode         TEXT,
    subclass_mode      TEXT,
    is_tentative       INTEGER,           -- nome com posição/identidade indefinida (heurística)
    smiles_candidate   TEXT,              -- herdado da busca por nome do pipeline antigo (AUT-NR)
    cid_candidate      INTEGER,
    candidate_source   TEXT,
    mono_mass          REAL,              -- CALC a partir do SMILES candidato
    mass_check         TEXT               -- consistent | inconsistent | unchecked | no_candidate
);

CREATE TABLE compound_name (
    compound_id    TEXT NOT NULL REFERENCES compound(compound_id),
    name_verbatim  TEXT NOT NULL,
    n              INTEGER NOT NULL,
    PRIMARY KEY (compound_id, name_verbatim)
);

-- O pico anotado: uma linha por ocorrência, sem deduplicar a evidência espectral.
CREATE TABLE peak (
    peak_id          INTEGER PRIMARY KEY,
    sheet_id         INTEGER NOT NULL REFERENCES sheet(sheet_id),
    sample_id        INTEGER REFERENCES sample(sample_id),   -- NULL em tabelas de picos sem contexto
    compound_id      TEXT REFERENCES compound(compound_id),
    name_verbatim    TEXT NOT NULL,
    name_source      TEXT,             -- molecule_col | conc_col_text | peak_table
    class_verbatim   TEXT,
    subclass_verbatim TEXT,
    class_inherited  INTEGER,          -- 1 = classe veio de linha anterior
    peak_number      TEXT,
    rt_min           REAL,
    lambda_max_text  TEXT,
    lambda_max_json  TEXT,             -- lista de λ em nm
    uv_spectrum      TEXT,
    mw               TEXT,
    m_plus           TEXT,
    m_plus_h         TEXT,
    m_minus_h        TEXT,
    m_minus_h_num    REAL,
    m2_plus_h        TEXT,
    m2_minus_h       TEXT,
    m_h_2neg         TEXT,
    m_h_3neg         TEXT,
    exact_mass       TEXT,
    adducts          TEXT,
    ms2_text         TEXT,
    ms2_json         TEXT,             -- [[m/z, intensidade relativa], ...]
    notes            TEXT,             -- textos achados nas colunas de concentração (ex.: padrão usado)
    evidence         TEXT NOT NULL DEFAULT 'REP',
    src_sheet        TEXT NOT NULL,
    src_cell         TEXT NOT NULL
);

-- Medição: valor como publicado (raw_text) + leitura normalizada.
CREATE TABLE measurement (
    measurement_id  INTEGER PRIMARY KEY,
    peak_id         INTEGER REFERENCES peak(peak_id),      -- NULL em totais de classe
    sample_id       INTEGER REFERENCES sample(sample_id),
    condition_id    INTEGER REFERENCES condition(condition_id),
    study_id        INTEGER REFERENCES study(study_id),    -- NULL quando a aba tem mais de um estudo
    target_class    TEXT,              -- preenchido em totais de classe (ex.: aba Wine)
    raw_text        TEXT NOT NULL,
    value           REAL,
    sd              REAL,
    is_range        INTEGER DEFAULT 0,
    flag            TEXT,              -- tr | nd | nq | lt_loq | dash
    unit_verbatim   TEXT,
    unit            TEXT,              -- mg/g | ug/g | mg/100 g | g/100 g | mg/kg | g/kg | ug/mL | mg/mL | mg/L | % | area
    basis           TEXT,              -- extract | dw | fw | volume | peak_area | unspecified
    preparation     TEXT,
    tag             TEXT,
    irradiation_kgy REAL,
    parse_status    TEXT,              -- ok | flag | no_unit | unparsed
    evidence        TEXT NOT NULL DEFAULT 'REP',
    src_sheet       TEXT NOT NULL,
    src_cell        TEXT NOT NULL
);

-- Registro de validação: tudo o que o ETL decidiu, corrigiu ou não conseguiu ler.
CREATE TABLE issue (
    issue_id   INTEGER PRIMARY KEY,
    severity   TEXT NOT NULL,          -- error | warning | info
    code       TEXT NOT NULL,
    entity     TEXT,
    entity_id  TEXT,
    message    TEXT,
    src_sheet  TEXT,
    src_cell   TEXT
);

CREATE INDEX ix_peak_compound ON peak(compound_id);
CREATE INDEX ix_peak_sample ON peak(sample_id);
CREATE INDEX ix_meas_peak ON measurement(peak_id);
CREATE INDEX ix_meas_condition ON measurement(condition_id);
CREATE INDEX ix_issue_code ON issue(code);
