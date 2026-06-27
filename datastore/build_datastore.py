#!/usr/bin/env python3
"""
build_datastore.py — build the DD Data Center (DuckDB) from raw inputs and
publish compact CSVs for Excel Power Query.

Pipeline (idempotent — safe to re-run any time):
    1. (re)create Layer 1 base tables          (schema/01_layer1_base.sql)
    2. ingest the per-drug JSON files           -> Layer 1
    3. seed parameter inputs from params_seed.json
    4. create Layer 2 derived views             (schema/02_layer2_views.sql)
    5. validate + export published CSVs          -> export/ (+ optional Excel dir)

The .duckdb file and the export/ CSVs hold proprietary numbers and are NOT
committed (see .gitignore). Only this code + the schema + the docs are tracked.

Usage:
    python datastore/build_datastore.py
    python datastore/build_datastore.py --excel-dir "/mnt/c/Users/yzsun/Desktop/DD/_datastore"
"""
import argparse
import glob
import json
import os
import re

import duckdb

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SCHEMA_DIR = os.path.join(HERE, "schema")
DB_PATH = os.path.join(HERE, "dd.duckdb")
EXPORT_DIR = os.path.join(HERE, "export")
SEED_PATH = os.path.join(HERE, "seed", "params_seed.json")
LEGACY_PATH = os.path.join(HERE, "seed", "legacy_drugs.json")
BLOOD_PATH = os.path.join(HERE, "seed", "blood_drugs.json")

# Each source dir maps to a tam_group and a default split method.
SOURCES = [
    (os.path.join(REPO, "tam_hl_mm_data"), "blood", "share"),      # richer, take first
    (os.path.join(REPO, "tam_data"),       "solid", "incidence"),
]


def slug(name):
    base = name.split("(")[0].strip().lower()
    return re.sub(r"[^a-z0-9]+", "_", base).strip("_")


def parse_name(full):
    full = full.strip()
    if "(" in full and ")" in full:
        base = full[:full.index("(")].strip()
        inside = full[full.index("(") + 1: full.rindex(")")].strip()
        parts = [p.strip() for p in inside.split(",") if p.strip()]
        company = parts[0] if parts else None
        molecule = ", ".join(parts[1:]) if len(parts) > 1 else None
        return base, company, molecule
    return full, None, None


def load_json_drugs():
    """Return (drugs, revenues, splits, indication_cases, conflicts) deduped by drug_id.

    Sources, in priority order (first occurrence of a drug_id wins):
      1. tam_hl_mm_data/*.json  (blood, richest)
      2. tam_data/*.json        (solid)
      3. seed/legacy_drugs.json (solid, extracted from the live TAM Solid sheet)
    """
    drugs, revenues, splits = {}, [], []
    indication_cases = {}   # indication_code -> max incidence_global_annual seen
    conflicts = []

    def add(d, tam_group, source):
        full = d["drug_name"]
        did = slug(full)
        if did in drugs:
            conflicts.append((did, source))
            return
        base, company, molecule = parse_name(full)
        company = d.get("company") or company          # explicit fields win
        molecule = d.get("molecule") or molecule
        drugs[did] = (did, base, company, molecule, tam_group, source)
        for yr, val in d.get("revenues", {}).items():
            revenues.append((did, int(yr), float(val)))
        for ind in d.get("indications", []):
            code = ind["name"].strip()
            if "share" in ind:
                splits.append((did, code, "share", float(ind["share"]), None))
            else:
                w = ind.get("incidence_global_annual")
                splits.append((did, code, "incidence", None,
                               float(w) if w is not None else None))
                if w is not None:
                    indication_cases[code] = max(indication_cases.get(code, 0), int(w))

    for src_dir, tam_group, _ in SOURCES:
        for path in sorted(glob.glob(os.path.join(src_dir, "*.json"))):
            with open(path) as f:
                add(json.load(f), tam_group, os.path.basename(path))

    if os.path.exists(LEGACY_PATH):                       # legacy solid backfill
        for d in json.load(open(LEGACY_PATH)):
            add(d, "solid", "legacy_drugs.json")

    if os.path.exists(BLOOD_PATH):                        # TAM Blood roster
        for d in json.load(open(BLOOD_PATH)):
            add(d, "blood", "blood_drugs.json")

    return list(drugs.values()), revenues, splits, indication_cases, conflicts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel-dir", default=os.environ.get(
        "DD_EXCEL_DIR", "/mnt/c/Users/yzsun/Desktop/DD/_datastore"),
        help="extra folder to copy published CSVs into (for Power Query)")
    args = ap.parse_args()

    # ---- load inputs -------------------------------------------------------
    drugs, revenues, splits, ind_cases, conflicts = load_json_drugs()
    seed = json.load(open(SEED_PATH)) if os.path.exists(SEED_PATH) else \
        {"incidence": {}, "globals": {}, "reference_growth": {}, "cogs": {}}

    # full indication universe = incidence keys + every indication used in a split
    ind_codes = set(seed.get("incidence", {})) | {s[1] for s in splits}

    # ---- build database ----------------------------------------------------
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    con = duckdb.connect(DB_PATH)
    con.execute(open(os.path.join(SCHEMA_DIR, "01_layer1_base.sql")).read())

    con.executemany(
        "INSERT INTO indication VALUES (?,?,?)",
        [(c, seed.get("incidence", {}).get(c), ind_cases.get(c)) for c in sorted(ind_codes)])
    con.executemany("INSERT INTO drug VALUES (?,?,?,?,?,?)", drugs)
    con.executemany("INSERT INTO drug_revenue VALUES (?,?,?)", revenues)
    con.executemany("INSERT INTO drug_indication_split VALUES (?,?,?,?,?)", splits)

    ref_rows = []
    for item, series in seed.get("reference_growth", {}).items():
        for yr, val in series.items():
            ref_rows.append(("growth", item, int(yr), float(val)))
    for item, val in seed.get("cogs", {}).items():
        if val is not None:
            ref_rows.append(("cogs", item, 0, float(val)))
    if ref_rows:
        con.executemany("INSERT INTO reference_drug_sale VALUES (?,?,?,?)", ref_rows)

    gl = seed.get("globals", {})
    param_rows = [(k, float(v), None) for k, v in gl.items() if v is not None]
    if param_rows:
        con.executemany("INSERT INTO param_input VALUES (?,?,?)", param_rows)

    con.execute(open(os.path.join(SCHEMA_DIR, "02_layer2_views.sql")).read())

    # ---- validate ----------------------------------------------------------
    n_drugs = con.sql("SELECT count(*) FROM drug").fetchone()[0]
    n_ind = con.sql("SELECT count(*) FROM indication").fetchone()[0]
    n_rev = con.sql("SELECT count(*) FROM drug_revenue").fetchone()[0]
    print(f"\n=== DD Data Center built: {DB_PATH} ===")
    print(f"Layer 1: {n_drugs} drugs | {n_ind} indications | {n_rev} revenue rows")
    if conflicts:
        names = ", ".join(sorted({c[0] for c in conflicts})[:12])
        print(f"  (deduped {len(conflicts)} duplicate drug_id(s): {names}"
              f"{'...' if len(conflicts) > 12 else ''})")

    print("\nLayer 2 — TAM by indication (2024), top 12:")
    print(con.sql("""
        SELECT indication_code, round(tam_usd_m,0) AS tam_2024_usd_m
        FROM v_tam_by_indication_year
        WHERE year = 2024 ORDER BY tam_2024_usd_m DESC LIMIT 12
    """).fetchdf().to_string(index=False))

    print("\nLayer 2 — derived parameters:")
    print("  growth tiers:",
          con.sql("SELECT tier, round(growth_factor,4) FROM v_param_growth").fetchall())
    cp = con.sql("SELECT round(cogs_price,4) FROM v_param_cogs_price").fetchone()
    print("  COGS/Price :", cp[0] if cp else None)

    # ---- export published CSVs --------------------------------------------
    exports = {
        "tam_by_indication_year": "v_tam_by_indication_year",
        "tam_by_group_year":      "v_tam_by_group_year",
        "drug_indication_revenue": "v_drug_indication_revenue",
        "param_incidence":        "v_param_incidence",
        "param_growth":           "v_param_growth",
        "param_cogs_price":       "v_param_cogs_price",
        "drug":                   "drug",
        "drug_revenue":           "drug_revenue",
    }
    out_dirs = [EXPORT_DIR]
    if args.excel_dir and os.path.isdir(os.path.dirname(args.excel_dir.rstrip("/"))):
        os.makedirs(args.excel_dir, exist_ok=True)
        out_dirs.append(args.excel_dir)

    for d in out_dirs:
        os.makedirs(d, exist_ok=True)
        for name, rel in exports.items():
            target = os.path.join(d, f"{name}.csv").replace("'", "''")
            con.execute(f"COPY (SELECT * FROM {rel}) TO '{target}' (HEADER, DELIMITER ',')")
    print(f"\nPublished {len(exports)} CSVs to: {', '.join(out_dirs)}")
    con.close()


if __name__ == "__main__":
    main()
