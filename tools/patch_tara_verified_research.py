#!/usr/bin/env python3
"""Patch TARA facts that were re-verified against primary sources on 2026-07-11.

This deliberately edits the OOXML package in place so Excel data tables and chart
parts are preserved.  Market-data refresh is handled separately by
``neutralize_bloomberg_artifacts.py``.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.neutralize_bloomberg_artifacts import (
    _apply_explicit_values,
    _read_parts,
    _sheet_zip_paths,
    _validate_xml_parts,
    _write_parts,
)


DEFAULT = Path("/mnt/c/Users/yzsun/Desktop/DD/TARA/DCF TARA.xlsx")


def patch(path: Path, backup: bool = True) -> Path | None:
    backup_path = None
    if backup:
        backup_path = path.with_name(
            f"{path.stem}_pre_verified_research_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        shutil.copy2(path, backup_path)

    parts = _read_parts(path)
    order = list(parts)
    sheets = _sheet_zip_paths(parts)

    peer_updates = {
        "H8": "Registrational Phase 3 THRIVE-3 (seamless 2b/3 design); interim analysis expected 2H 2026",
        "P9": "TEAEs were generally mild to moderate; 1/16 (6%) had a Grade 3 TEAE and related Grade 3 TEAE; no serious TEAEs or adverse events of special interest",
    }
    events_updates = {
        # Place the financing news on the actual announcement/pricing dates and
        # remove the previously misdated/mistyped duplicate rows.
        "U104": "$45 million private placement announced to fund TARA-002 and IV Choline",
        "V104": "Financing",
        "U110": "",
        "V110": "",
        "U348": "",
        "V348": "",
        "U352": "$100 million underwritten public offering priced at $6.25 per share",
        "V352": "Financing",
        "AQ77": "FY2025 results reported $197.9 million unrestricted liquidity and runway into 2028",
    }

    for sheet_name, updates in (("Peer View", peer_updates),
                                ("Historical Events", events_updates)):
        sheet_path = sheets[sheet_name]
        xml = parts[sheet_path].decode("utf-8", "ignore")
        parts[sheet_path] = _apply_explicit_values(xml, updates).encode("utf-8")

    _write_parts(path, parts, order)
    _validate_xml_parts(path)
    return backup_path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", type=Path, default=DEFAULT)
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()
    bak = patch(args.path, backup=not args.no_backup)
    print(f"Patched verified TARA research facts: {args.path}")
    if bak:
        print(f"Backup: {bak}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
