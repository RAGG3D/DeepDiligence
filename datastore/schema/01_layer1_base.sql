-- ============================================================================
--  DD DATA CENTER — LAYER 1 : BASE / RAW TABLES  (the system of record)
-- ----------------------------------------------------------------------------
--  These tables hold ONLY raw, hand-curated facts and assumptions:
--    * a drug's real-world total net sales (drug_revenue)
--    * the analyst's split of that drug across cancer indications (drug_indication_split)
--    * epidemiology inputs (indication.incidence_rate)
--    * the reference drug ramps / cost lines used to DERIVE parameters
--    * scalar global assumptions (param_input)
--
--  NOTHING in this layer is computed from another row in this layer.
--  Everything derived (TAM, growth, COGS/Price) lives in Layer 2 as VIEWS.
--
--  Facts are addressed by KEY (drug_id, indication_code, year) — never by a
--  spreadsheet row number. Adding a drug is an INSERT; no row ever shifts.
-- ============================================================================

DROP VIEW  IF EXISTS v_peer_rating;
DROP VIEW  IF EXISTS v_param_cogs_price;
DROP VIEW  IF EXISTS v_param_maturity;
DROP VIEW  IF EXISTS v_param_growth;
DROP VIEW  IF EXISTS v_param_incidence;
DROP VIEW  IF EXISTS v_tam_by_group_year;
DROP VIEW  IF EXISTS v_tam_by_indication_year;
DROP VIEW  IF EXISTS v_drug_indication_revenue;
DROP VIEW  IF EXISTS v_drug_split;

DROP TABLE IF EXISTS peer_metric;
DROP TABLE IF EXISTS peer_drug;
DROP TABLE IF EXISTS drug_indication_split;
DROP TABLE IF EXISTS drug_revenue;
DROP TABLE IF EXISTS reference_drug_sale;
DROP TABLE IF EXISTS param_maturity;
DROP TABLE IF EXISTS param_input;
DROP TABLE IF EXISTS drug;
DROP TABLE IF EXISTS indication;

-- One row per cancer indication. incidence_rate is an INPUT (cases per capita
-- per year) read from the TAM Solid "Parameters" section; it is NOT derived.
CREATE TABLE indication (
    indication_code         VARCHAR PRIMARY KEY,   -- 'NSCLC', 'MM', 'BTC', ...
    incidence_rate          DOUBLE,                -- annual cases / world capita
    incidence_global_annual BIGINT                 -- absolute annual cases (if known)
);

-- One row per drug. tam_group routes a drug to the TAM Solid vs TAM Blood world.
CREATE TABLE drug (
    drug_id     VARCHAR PRIMARY KEY,               -- slug, e.g. 'keytruda'
    drug_name   VARCHAR NOT NULL,                  -- display name
    company     VARCHAR,
    molecule    VARCHAR,
    tam_group   VARCHAR NOT NULL CHECK (tam_group IN ('solid','blood')),
    source_file VARCHAR
);

-- A drug's TOTAL net sales per year ($M). The raw real-world fact.
CREATE TABLE drug_revenue (
    drug_id       VARCHAR NOT NULL REFERENCES drug(drug_id),
    year          INTEGER NOT NULL CHECK (year BETWEEN 2000 AND 2050),
    revenue_usd_m DOUBLE  NOT NULL,
    PRIMARY KEY (drug_id, year)
);

-- How a drug's revenue is split across indications. The split is an analyst
-- ASSUMPTION, kept separate from the raw total above.
--   method='share'     -> use `share` directly (e.g. Keytruda NSCLC = 0.339)
--   method='incidence' -> weight by `incidence_weight`; effective share is
--                         weight / SUM(weight) over the drug (see Layer 2).
CREATE TABLE drug_indication_split (
    drug_id          VARCHAR NOT NULL REFERENCES drug(drug_id),
    indication_code  VARCHAR NOT NULL REFERENCES indication(indication_code),
    method           VARCHAR NOT NULL CHECK (method IN ('share','incidence')),
    share            DOUBLE,
    incidence_weight DOUBLE,
    PRIMARY KEY (drug_id, indication_code)
);

-- Raw reference series used to DERIVE parameters in Layer 2.
--   purpose='growth' -> peer drug sales ramps (Alimta, Tagrisso, ...) by year
--   purpose='cogs'   -> Xpovio cost lines (year=0, scalar)
CREATE TABLE reference_drug_sale (
    purpose VARCHAR NOT NULL CHECK (purpose IN ('growth','cogs')),
    item    VARCHAR NOT NULL,
    year    INTEGER NOT NULL DEFAULT 0,
    value   DOUBLE  NOT NULL,
    PRIMARY KEY (purpose, item, year)
);

-- Scalar global assumptions (world population, population growth, ...).
CREATE TABLE param_input (
    key   VARCHAR PRIMARY KEY,
    value DOUBLE,
    note  VARCHAR
);

-- Maturity curve the revenue model actually uses: ramp factor by years-since-
-- launch (year_offset 1..29) for each tier. This is what the Pipeline sheet's
-- INDEX($F$553:$AH$553, ...) consumed as a row; here it is keyed, not positional.
CREATE TABLE param_maturity (
    tier        VARCHAR NOT NULL CHECK (tier IN ('AVG','BIC','T1')),
    year_offset INTEGER NOT NULL,            -- 1 = launch year
    factor      DOUBLE  NOT NULL,
    PRIMARY KEY (tier, year_offset)
);

-- Peer Views: drug-vs-drug clinical-readout comparison tables, extracted from
-- the Peer Views sheet. One row per drug per section; the BIC/T1/AVG rating is
-- DECODED from cell fill color into explicit text (no more color-as-data).
CREATE TABLE peer_drug (
    section_id INTEGER NOT NULL,             -- section order (disambiguates titles)
    section    VARCHAR NOT NULL,             -- indication / setting title
    col        VARCHAR NOT NULL,             -- source column letter (drug slot)
    drug       VARCHAR,
    ticker     VARCHAR,
    rating     VARCHAR CHECK (rating IN ('BIC','T1','AVG') OR rating IS NULL),
    PRIMARY KEY (section_id, col)
);

-- Long/tidy clinical metrics per drug (ORR, Median PFS/OS, sales, dates, ...).
-- value is text because metrics are mixed-type (numbers, dates, free text).
CREATE TABLE peer_metric (
    section_id INTEGER NOT NULL,
    section    VARCHAR NOT NULL,
    col        VARCHAR NOT NULL,
    metric     VARCHAR NOT NULL,
    value      VARCHAR,
    PRIMARY KEY (section_id, col, metric)
);
