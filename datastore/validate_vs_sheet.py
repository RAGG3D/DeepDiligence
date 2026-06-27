#!/usr/bin/env python3
"""
validate_vs_sheet.py — compare the Data Center's computed TAM
(v_tam_by_indication_year) against the live TAM Solid sheet's own SUMIF TAM rows.

This is a fidelity check, not a pass/fail gate. Expect divergence where the
datastore comprehensively includes a drug's minor cross-indication slices (e.g.
Keytruda's 9% melanoma share) that the sheet's hand-built TAM rows omit.

Usage:
    python datastore/validate_vs_sheet.py \
        --xlsx "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"
"""
import argparse
import os

import duckdb
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))

# TAM Solid SUMIF rows -> indication code (the indications the sheet computes)
SHEET_TAM_ROWS = {406: "BTC", 408: "CRC", 410: "NSCLC", 412: "HNSCC", 415: "Melanoma"}
YEAR_COL = {2020: 16, 2021: 17, 2022: 18, 2023: 19, 2024: 20}  # columns P..T


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default="/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
    ap.add_argument("--db", default=os.path.join(HERE, "dd.duckdb"))
    args = ap.parse_args()

    ws = openpyxl.load_workbook(args.xlsx, data_only=True)["TAM Solid"]
    con = duckdb.connect(args.db, read_only=True)

    print(f"{'indication':10} {'year':>5} {'sheet':>11} {'datastore':>11} {'diff%':>8}")
    print("-" * 50)
    for row, ind in SHEET_TAM_ROWS.items():
        for y, c in YEAR_COL.items():
            sv = ws.cell(row, c).value
            if not isinstance(sv, (int, float)):
                continue
            dv = con.execute("SELECT tam_usd_m FROM v_tam_by_indication_year "
                             "WHERE indication_code = ? AND year = ?", [ind, y]).fetchone()
            dv = dv[0] if dv else 0.0
            diff = (dv - sv) / sv * 100 if sv else float("nan")
            print(f"{ind:10} {y:>5} {sv:11.1f} {dv:11.1f} {diff:7.1f}%")
    con.close()


if __name__ == "__main__":
    main()
