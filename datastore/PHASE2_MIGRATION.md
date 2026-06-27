# Phase 2 — Migrating Excel formulas onto the Data Center

Goal: repoint the workbook's downstream formulas from **row-addressed** references
into `TAM Solid` to **key-addressed** lookups against the published Power Query
tables. After this, adding a drug never shifts a reference again.

## What actually references the data-center tabs

A scan of the live workbook found that **only the `Pipeline` sheet** references
these tabs — **368 formulas, all into `TAM Solid`** (TAM Blood and Peer Views are
consumed by the Python scripts, not by in-sheet formulas). The 368 collapse into
**4 patterns**:

| # | Pattern | Count | What it does |
|---|---|---:|---|
| 1 | `SUMIF` over the drug range | 224 | TAM for an indication in a year |
| 2 | pattern 1 × market share × growth `INDEX` | 64 | revenue = TAM × MS × growth tier |
| 3 | COGS/Price row × revenue | 64 | cost of goods |
| 4 | single year-header cell | 16 | pulls a year label |

## Prerequisite (one-time, manual in Excel)

Load the published folder via Power Query so the tables exist as Excel Tables
(see `POWER_QUERY_SETUP.md`):
`tam_by_indication_year`, `param_growth`, `param_cogs_price`, `peer_rating`, …
This GUI step is the only part that can't be done headlessly.

## The four rewrites

**Pattern 1 — TAM by indication × year** (224 cells)
```excel
# before — depends on the literal range $9:$562 (breaks on every drug insert)
=SUMIF('TAM Solid'!$D$9:$D$562, $C10, 'TAM Solid'!S$9:S$562)

# after — keyed by (indication, year); range-free
=SUMIFS(tam_by_indication_year[tam_usd_m],
        tam_by_indication_year[indication_code], $C10,
        tam_by_indication_year[year],          F$6)     # F$6 = the column's year
```

**Pattern 2 — revenue = TAM × market share × maturity factor** (64 cells)
```excel
# before — maturity pinned to row 553 (T1), a 29-wide row that shifts on insert
... * F11 * INDEX('TAM Solid'!$F$553:$AH$553, <year-offset>)

# after — keyed by (tier, year-offset) into the maturity table
... * F11 * SUMIFS(param_maturity[factor],
                   param_maturity[tier],        "T1",
                   param_maturity[year_offset], <year-offset>)
```
> `param_maturity` is the full curve: factor by years-since-launch (offset 1..29)
> per tier (AVG/BIC/T1), extracted straight from the sheet's rows 551–553. The
> `<year-offset>` is the same launch-relative index the old `INDEX` computed.

**Pattern 3 — COGS/Price × revenue** (64 cells)
```excel
# before — pinned to row 562
='TAM Solid'!$P$562 * F15

# after — single-value param table
=param_cogs_price[cogs_price] * F15
```

**Pattern 4 — year-header cell** (16 cells)
```excel
# before
='TAM Solid'!S6
# after — use the Pipeline sheet's own year header (or a literal)
=F$6
```

## How to apply (safe, on a copy)

1. Copy `DCF {TICKER}.xlsx` → `DCF {TICKER}_datacenter.xlsx`.
2. Do the one-time Power Query load (prerequisite above).
3. Find/replace the 4 patterns (each is mechanical and regex-able in the sheet XML).
4. Run `validate_vs_sheet.py` + spot-check Pipeline outputs against the original.
5. Once reconciled, retire `tam/expand_tam.py` row-shifting and the
   `GROWTH_ROW={551,552,553}` / `COGS_PRICE_ROW=562` constants in
   `generate/generate_pipeline.py` — they no longer have any consumer.

## Why this is the payoff

Every one of the 368 references stops caring about *where* a number sits.
A new drug becomes one JSON edit + `build_datastore.py`; the Pipeline formulas,
keyed by indication/year/tier, keep working untouched. The `+61-row cascade` and
its four-file constant edits are gone.
