#!/usr/bin/env python3
"""
main.py – CLI entry point for the DCF auto-fill automation.

Usage
─────
  python main.py --ticker CMPX
  python main.py --ticker BHVN --years 2020 2021 2022 2023 2024
  python main.py --ticker CMPX --path "/custom/path/DCF CMPX.xlsx"
  python main.py --ticker CMPX --dry-run        # fetch & print without writing
  python main.py --ticker CMPX --unit MM         # force MM USD mode

Example
───────
  cd ~/Investment/auto_dcf
  python main.py --ticker CMPX
"""

import argparse
import logging
import sys
from typing import List

from core.sec_fetcher import SECFetcher
from core.excel_writer import update_workbook

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

DEFAULT_YEARS = [2020, 2021, 2022, 2023, 2024]


# ── CLI ────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Auto-fill DCF Excel model from SEC 10-K XBRL data."
    )
    parser.add_argument(
        "--ticker", required=True,
        help="Stock ticker symbol (e.g. CMPX)"
    )
    parser.add_argument(
        "--years", nargs="+", type=int, default=DEFAULT_YEARS,
        metavar="YEAR",
        help=f"Fiscal years to fetch (default: {DEFAULT_YEARS})"
    )
    parser.add_argument(
        "--path", default=None,
        help="Override Excel file path (default: /mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Fetch and print data without modifying the Excel file."
    )
    parser.add_argument(
        "--cik", default=None,
        help="Override CIK (e.g. 0001935979). Bypasses ticker→CIK auto-resolution."
    )
    parser.add_argument(
        "--unit", choices=["K", "MM"], default=None,
        help="Force reporting unit: K=thousands, MM=millions. Auto-detected if omitted."
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable DEBUG-level logging."
    )
    return parser.parse_args()


# ── Pretty-print helper ────────────────────────────────────────────────────────

def print_summary(
    ticker: str,
    years: List[int],
    financial_data: dict,
    reporting_unit: str = "K",
    currency: str = "USD",
) -> None:
    """Print a formatted table of the fetched storage values."""
    unit_label = f"{reporting_unit} {currency}"
    print(f"\n{'═'*90}")
    print(f"  {ticker} – Fetched Financial Data ({unit_label} storage values)")
    print(f"{'═'*90}")

    year_header = "  ".join(f"{y:>10}" for y in sorted(years))
    print(f"  {'Statement':6} {'Sub':3} {'Line Item':<45}  {year_header}")
    print(f"  {'-'*6} {'-'*3} {'-'*45}  {('-'*10 + '  ') * len(years)}")

    for (col_c, col_b, col_d), year_vals in sorted(
        financial_data.items(), key=lambda x: (x[0][0] or "", x[0][2])
    ):
        vals = "  ".join(
            f"{year_vals.get(y):>10,.0f}" if year_vals.get(y) is not None
            else f"{'–':>10}"
            for y in sorted(years)
        )
        sub = str(col_b) if col_b is not None else ""
        line = col_d[:44]
        print(f"  {col_c or '':6} {sub:3} {line:<45}  {vals}")

    print(f"{'═'*90}\n")


# ── Main ────────────────────────────────────────────────────────────────────────

def main() -> int:
    args = parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    ticker = args.ticker.upper().strip()
    years  = sorted(set(args.years))

    print(f"\n▶ Ticker : {ticker}")
    print(f"▶ Years  : {years}")
    if args.dry_run:
        print("▶ Mode   : DRY-RUN (no Excel changes will be made)")
    print()

    # ── 1. Fetch XBRL data ─────────────────────────────────────────────────────
    logger.info(f"Step 1/3 – Fetching SEC XBRL data for {ticker} …")
    fetcher = SECFetcher()
    try:
        financial_data, rename_map, note_details, reporting_unit, currency = \
            fetcher.build_financial_data(
                ticker, years, cik_override=args.cik,
                unit_override=args.unit,
            )
    except Exception as exc:
        logger.error(f"Failed to fetch SEC data: {exc}")
        return 1

    unit_desc = 'thousands' if reporting_unit == 'K' else 'millions'
    print(f"▶ Unit   : {reporting_unit} {currency} ({unit_desc})")

    # ── 2. Print summary ────────────────────────────────────────────────────────
    print_summary(ticker, years, financial_data, reporting_unit, currency)

    if args.dry_run:
        print("Dry-run complete. No Excel file was modified.\n")
        return 0

    # ── 3. Update workbook ──────────────────────────────────────────────────────
    logger.info(f"Step 2/3 – Updating Excel workbook …")
    try:
        saved_path = update_workbook(
            ticker=ticker,
            financial_data=financial_data,
            years=years,
            excel_path=args.path,
            rename_map=rename_map,
            note_details=note_details,
            reporting_unit=reporting_unit,
        )
    except FileNotFoundError as exc:
        logger.error(str(exc))
        logger.error(
            "Tip: Check that the file exists at the expected path, or use "
            "--path to specify a custom location."
        )
        return 1
    except Exception as exc:
        logger.error(f"Failed to update workbook: {exc}", exc_info=True)
        return 1

    logger.info(f"Step 3/3 – Complete. Workbook saved: {saved_path}")
    print(f"\n✔ Workbook updated: {saved_path}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
