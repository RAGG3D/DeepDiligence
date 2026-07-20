#!/usr/bin/env python3
"""Remove circular market-share carry-forward ranges from Scenarios.

The legacy generator emitted ``MAX($Irow:<prior>row)``.  In the first projected
columns that is a reversed range and Excel expands it across the current cell,
creating a circular dependency.  This utility patches an existing workbook in
place without rewriting any other OOXML structures.
"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path


def patch_workbook(path: Path) -> int:
    with zipfile.ZipFile(path, "r") as zin:
        workbook = zin.read("xl/workbook.xml").decode("utf-8")
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        sheet_match = re.search(
            r'<sheet\b[^>]*\bname="Scenarios"[^>]*\br:id="([^"]+)"', workbook
        )
        if not sheet_match:
            raise RuntimeError("Scenarios sheet not found")
        rid = re.escape(sheet_match.group(1))
        rel_match = re.search(
            rf'<Relationship\b[^>]*\bId="{rid}"[^>]*\bTarget="([^"]+)"', rels
        )
        if not rel_match:
            raise RuntimeError("Scenarios relationship not found")
        target = rel_match.group(1).lstrip("/")
        member = target if target.startswith("xl/") else f"xl/{target}"
        xml = zin.read(member).decode("utf-8")

        # The row number must be identical on both ends so unrelated formulas
        # are untouched.  Column E is the first historical year in Scenarios.
        fixed, count = re.subn(
            r'MAX\(\$I(\d+):\$?([A-Z]{1,3})\1\)',
            lambda m: f"MAX($E{m.group(1)}:${m.group(2)}{m.group(1)})",
            xml,
        )
        # Excel normalizes reversed H/I ranges when it saves a workbook.  If a
        # legacy file has already been opened, repair those normalized forms as
        # well: H formulas become G:I and I formulas become H:I.
        for start_col in ("G", "H"):
            fixed, extra = re.subn(
                rf'MAX\({start_col}(\d+):\$I\1\)',
                lambda m, col=start_col: f"MAX($E{m.group(1)}:${col}{m.group(1)})",
                fixed,
            )
            count += extra
        if not count:
            return 0

        with tempfile.NamedTemporaryFile(
            prefix=path.stem + ".", suffix=".xlsx", dir=path.parent, delete=False
        ) as handle:
            tmp = Path(handle.name)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    zout.writestr(item, fixed if item.filename == member else zin.read(item.filename))
            shutil.move(tmp, path)
        finally:
            tmp.unlink(missing_ok=True)
        return count


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    if not args.path.exists():
        raise SystemExit(f"Workbook not found: {args.path}")
    if not args.no_backup:
        backup = args.path.with_name(
            f"{args.path.stem}_pre_scenario_formula_fix_"
            f"{datetime.now():%Y%m%d_%H%M%S}{args.path.suffix}"
        )
        shutil.copy2(args.path, backup)
        print(f"Backup: {backup}")
    print(f"Scenario carry-forward formulas fixed: {patch_workbook(args.path)}")


if __name__ == "__main__":
    main()
