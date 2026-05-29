#!/usr/bin/env python3
"""
main_enhanced.py – Enhanced CLI with Notes-level data extraction.

Usage:
  python main_enhanced.py --ticker CMPX
  python main_enhanced.py --ticker CMPX --dry-run
"""

import argparse
import logging
import sys
from typing import List

from sec_fetcher_enhanced import SECFetcher
from excel_writer_enhanced import update_workbook

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

DEFAULT_YEARS = [2020, 2021, 2022, 2023, 2024]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Enhanced DCF auto-fill with Notes-level details."
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
        help="Override Excel file path"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Fetch and print data without modifying the Excel file."
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable DEBUG-level logging."
    )
    return parser.parse_args()


def print_summary(
    ticker: str,
    years: List[int],
    financial_data: dict,
) -> None:
    """Print a formatted table of the fetched K USD values."""
    print(f"\n{'═'*100}")
    print(f"  {ticker} – Enhanced Financial Data with Notes Details (K USD)")
    print(f"{'═'*100}")

    year_header = "  ".join(f"{y:>10}" for y in sorted(years))
    print(f"  {'Stmt':4} {'Sub':3} {'Line Item':<50}  {year_header}")
    print(f"  {'-'*4} {'-'*3} {'-'*50}  {('-'*10 + '  ') * len(years)}")

    for (col_c, col_b, col_d), year_vals in sorted(
        financial_data.items(), key=lambda x: (x[0][0] or "", x[0][1] or 0, x[0][2])
    ):
        vals = "  ".join(
            f"{year_vals.get(y):>10,.0f}" if year_vals.get(y) is not None
            else f"{'–':>10}"
            for y in sorted(years)
        )
        sub = str(col_b) if col_b is not None else ""
        line = col_d[:49]
        print(f"  {col_c or '':4} {sub:3} {line:<50}  {vals}")

    print(f"{'═'*100}\n")


def main() -> int:
    args = parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    ticker = args.ticker.upper().strip()
    years  = sorted(set(args.years))

    print(f"\n{'='*80}")
    print(f"  DCF Auto-Fill Enhanced – Notes-Level Data Extraction")
    print(f"{'='*80}")
    print(f"▶ Ticker : {ticker}")
    print(f"▶ Years  : {years}")
    if args.dry_run:
        print("▶ Mode   : DRY-RUN (no Excel changes)")
    print()

    # ── 1. Fetch XBRL data ─────────────────────────────────────────────────────
    logger.info(f"Step 1/3 – Fetching enhanced SEC XBRL data for {ticker} …")
    fetcher = SECFetcher()
    try:
        financial_data = fetcher.build_financial_data(ticker, years)
    except Exception as exc:
        logger.error(f"Failed to fetch SEC data: {exc}", exc_info=True)
        return 1

    # ── 2. Print summary ────────────────────────────────────────────────────────
    print_summary(ticker, years, financial_data)

    if args.dry_run:
        print("✓ Dry-run complete. No Excel file was modified.\n")
        return 0

    # ── 3. Update workbook ──────────────────────────────────────────────────────
    logger.info(f"Step 2/3 – Updating Excel workbook with enhanced data …")
    try:
        saved_path = update_workbook(
            ticker=ticker,
            financial_data=financial_data,
            years=years,
            excel_path=args.path,
        )
    except FileNotFoundError as exc:
        logger.error(str(exc))
        return 1
    except Exception as exc:
        logger.error(f"Failed to update workbook: {exc}", exc_info=True)
        return 1

    logger.info(f"Step 3/3 – Complete. Workbook saved: {saved_path}")
    print(f"\n{'='*80}")
    print(f"✓ Workbook updated successfully: {saved_path}")
    print(f"{'='*80}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
