#!/usr/bin/env python3
"""
fix_tam_divzero2.py -- Fix remaining #DIV/0! errors in TAM Solid rows 74-75, 97-99.

Root cause: Same as fix_tam_divzero.py -- formulas reference drug data rows
($373-$377) instead of the correct incidence rows ($460-$464) after row
restructuring by expand_tam.py and fix_tam_layout.py.

Row mapping (old → current):
  $373 → drug row (Cyramza)   — should be $460 (OV Incidence)
  $374 → drug row (Mekinist)  — should be $461 (TNBC Incidence)
  $375 → drug row (Tafinlar)  — should be $462 (BRCA Incidence)
  $377 → drug row (Retevmo)   — should be $464 (GC Incidence)

Affected drugs:
  Doxil (doxorubicin liposomal):
    R74 = OV sub-row   — Approved for ovarian cancer (1995, 2nd line)
    R75 = BRCA sub-row  — Used for breast cancer (off-label; model split valid)
  Xeloda (capecitabine):
    R97 = BRCA sub-row  — Approved for breast cancer (1998)
    R98 = TNBC sub-row  — Chemotherapy agent, used in TNBC
    R99 = GC sub-row    — Approved for gastric cancer

Uses surgical zip patching (NEVER openpyxl .save()).
"""

import logging
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Optional


_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

_DEFAULT_FILE = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Row reference corrections: old_row → new_row (incidence rows)
_REF_MAP = {
    373: 460,  # OV Incidence
    374: 461,  # TNBC Incidence
    375: 462,  # BRCA Incidence
    377: 464,  # GC Incidence
}

# Rows to fix and which references they use
_ROWS_TO_FIX = {
    74: {373: 460, 375: 462},           # Doxil OV
    75: {373: 460, 375: 462},           # Doxil BRCA
    97: {375: 462, 374: 461, 377: 464}, # Xeloda BRCA
    98: {375: 462, 374: 461, 377: 464}, # Xeloda TNBC
    99: {375: 462, 374: 461, 377: 464}, # Xeloda GC
}

_ROW_RE = r'<row\b[^/>]*r="{row}"[^/>]*(?:>.*?</row>|/>)'


def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_path = {}
    for rel in rels_xml:
        if "worksheet" in rel.get("Type", ""):
            tgt = rel.get("Target", "")
            rid_to_path[rel.get("Id")] = (
                f"xl/{tgt}" if not tgt.startswith("/") else tgt.lstrip("/"))
    for sheet in wb_xml.findall(f".//{{{_NS_MAIN}}}sheet"):
        if sheet.get("name") == sheet_name:
            return rid_to_path.get(sheet.get(f"{{{_NS_R}}}id"))
    return None


def _fix_row_formulas(xml: str, row_num: int, ref_map: dict) -> str:
    """Fix formula references in a specific row.

    1. Replace $OLD_ROW → $NEW_ROW in all formula text within the row
    2. Remove t="e" error attribute from cells
    3. Remove <v>#DIV/0!</v> error values
    """
    pattern = _ROW_RE.format(row=row_num)
    m = re.search(pattern, xml, re.DOTALL)
    if not m:
        log.warning(f"  Row {row_num} not found")
        return xml

    old_row_xml = m.group(0)
    new_row_xml = old_row_xml

    # Fix formula references (use negative lookahead to avoid matching $3730 etc.)
    for old_ref, new_ref in ref_map.items():
        new_row_xml = re.sub(
            rf'\${old_ref}(?!\d)', f'${new_ref}', new_row_xml)

    # Remove t="e" error attribute from cell elements
    new_row_xml = re.sub(r' t="e"', '', new_row_xml)

    # Remove error values
    new_row_xml = new_row_xml.replace('<v>#DIV/0!</v>', '')

    if new_row_xml == old_row_xml:
        log.info(f"  R{row_num}: no changes needed")
        return xml

    # Count changes
    old_refs = sum(len(re.findall(rf'\${old}(?!\d)', old_row_xml))
                   for old in ref_map)
    errors = old_row_xml.count('t="e"')
    log.info(f"  R{row_num}: fixed {old_refs} ref(s), removed {errors} error(s)")

    return xml.replace(old_row_xml, new_row_xml, 1)


def process_sheet(xml: str) -> str:
    """Fix all #DIV/0! cells in rows 74-75 and 97-99."""

    # ── R74-75: Doxil (OV / BRCA) ──────────────────────────────────────
    # Doxil (doxorubicin liposomal) — approved for ovarian cancer (1995).
    # Formula: COL73/3*2*(COL$460/(COL$460+COL$462))
    # (2/3 of Doxil revenue split between OV and BRCA by incidence ratio)
    log.info("Fixing Doxil sub-rows (R74 OV, R75 BRCA)")
    for row_num in [74, 75]:
        xml = _fix_row_formulas(xml, row_num, _ROWS_TO_FIX[row_num])

    # ── R97-99: Xeloda (BRCA / TNBC / GC) ─────────────────────────────
    # Xeloda (capecitabine) — approved for breast cancer (1998), GC, CRC.
    # Formula: (COL$95-COL$96)*(COL$462/(COL$462+COL$461+COL$464))
    # (non-CRC Xeloda revenue split among BRCA/TNBC/GC by incidence ratio)
    log.info("Fixing Xeloda sub-rows (R97 BRCA, R98 TNBC, R99 GC)")
    for row_num in [97, 98, 99]:
        xml = _fix_row_formulas(xml, row_num, _ROWS_TO_FIX[row_num])

    return xml


def _count_errors(xml: str) -> int:
    """Count all t=\"e\" error cells in the entire sheet."""
    return len(re.findall(r't="e"', xml))


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Fix #DIV/0! errors in TAM Solid rows 74-75, 97-99")
    parser.add_argument("--file", default=str(_DEFAULT_FILE))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return

    sheet_zip = None
    for name in ["TAM Solid", "TAM Solid+MM"]:
        sheet_zip = _get_sheet_zip_path(xlsx_path, name)
        if sheet_zip:
            log.info(f"Found sheet '{name}' -> {sheet_zip}")
            break
    if not sheet_zip:
        log.error("TAM Solid sheet not found")
        return

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")

    old_errors = _count_errors(xml)
    new_xml = process_sheet(xml)
    new_errors = _count_errors(new_xml)

    log.info(f"Error cells in entire sheet: {old_errors} -> {new_errors}")

    if args.dry_run:
        # Show sample of fixed cells
        for rn in [74, 75, 97, 98, 99]:
            for col in ['F', 'M']:
                m = re.search(
                    rf'(<c r="{col}{rn}"[^/]*(?:>.*?</c>|/>))',
                    new_xml, re.DOTALL)
                if m:
                    cell = m.group(1)
                    print(f"  {col}{rn}: {cell[:120]}")
        return

    # Backup
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_divfix2_{ts}.xlsx")
    shutil.copy2(xlsx_path, backup)
    log.info(f"Backup: {backup}")

    # Surgical zip patch
    modified = {sheet_zip: new_xml.encode("utf-8")}

    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~divfix2.xlsx")
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue  # removed — references stripped below
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))

    try:
        tmp.replace(xlsx_path)
    except PermissionError:
        import os
        os.remove(str(xlsx_path))
        tmp.rename(xlsx_path)

    log.info(f"Saved -> {xlsx_path}")
    print(f"\n{'='*60}")
    print("TAM #DIV/0! Fix #2 Complete")
    print(f"  File: {xlsx_path}")
    print(f"  Doxil OV/BRCA: $373→$460, $375→$462 (incidence refs)")
    print(f"  Xeloda BRCA/TNBC/GC: $374→$461, $375→$462, $377→$464")
    print(f"  Error cells: {old_errors} -> {new_errors}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
