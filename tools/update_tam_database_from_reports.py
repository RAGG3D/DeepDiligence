#!/usr/bin/env python3
"""Update the DD data center TAM overrides from parser-compatible reports."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from datastore.tam_overrides import (  # noqa: E402
    EXCEL_EXPORT_DIR,
    EXPORT_DIR,
    parse_report_tam,
    publish_override_csvs,
    upsert_overrides,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Update datastore TAM overrides from research reports")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--report-dir", required=True)
    parser.add_argument("--no-rebuild", action="store_true",
                        help="Only patch seed/export CSVs; do not rebuild DuckDB")
    args = parser.parse_args()

    report_dir = Path(args.report_dir)
    entries = parse_report_tam(report_dir, args.ticker.upper())
    if not entries:
        raise SystemExit(f"No TAM anchors found in {report_dir}")

    overrides = upsert_overrides(entries)
    print(f"Upserted {len(entries)} TAM override entries into datastore/seed/tam_overrides.json")
    for item in sorted(entries, key=lambda x: (x["source_drug"], x["indication_code"])):
        print(
            f"  {item['source_drug']} / {item['indication_code']}: "
            f"{item['anchors']} ({item['tam_group']})"
        )

    if not args.no_rebuild:
        try:
            subprocess.run(
                [sys.executable, "datastore/build_datastore.py"],
                cwd=REPO,
                check=True,
            )
            return
        except Exception as exc:
            print(f"Datastore rebuild failed; applying CSV overlay fallback: {exc}")

    publish_override_csvs(overrides, (EXPORT_DIR, EXCEL_EXPORT_DIR))
    print(f"Published override CSV rows to {EXPORT_DIR} and {EXCEL_EXPORT_DIR}")


if __name__ == "__main__":
    main()
