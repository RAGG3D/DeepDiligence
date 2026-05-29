#!/usr/bin/env python3
"""
fix_tam_divzero.py -- Fix #DIV/0! errors in TAM Solid rows 35-44.

Root cause: Formulas reference $374/$375/$376 (now drug rows after restructuring)
instead of $461/$462/$463 (TNBC/BRCA/BLCA incidence rows).

Fix approach based on clinical facts:
  - Taxotere (docetaxel): IS used in TNBC → fix formula to reference correct incidence rows
  - Evista (raloxifene): NOT used in TNBC (ER+ SERM) → TNBC = 0, BRCA = drug total
  - Femara (letrozole): NOT used in TNBC (aromatase inhibitor) → TNBC = 0, BRCA = drug total

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


def _replace_cell(xml: str, row: int, col: str, new_cell: str) -> str:
    """Replace a single cell in the sheet XML.
    Handles both content cells and self-closing cells.
    """
    # Match content cell: <c r="COLrow" ...>...</c>
    pattern = rf'<c r="{col}{row}"[^/]*>.*?</c>'
    m = re.search(pattern, xml)
    if m:
        old = m.group(0)
        xml = xml.replace(old, new_cell, 1)
        return xml

    # Match self-closing cell: <c r="COLrow" .../>
    pattern = rf'<c r="{col}{row}"[^/]*/>'
    m = re.search(pattern, xml)
    if m:
        old = m.group(0)
        xml = xml.replace(old, new_cell, 1)
        return xml

    log.warning(f"  Cell {col}{row} not found in XML")
    return xml


def process_sheet(xml: str) -> str:
    """Fix all #DIV/0! cells in rows 36-44."""

    changes = 0

    # ── R36: Taxotere TNBC ──────────────────────────────────────────────
    # Taxotere IS used in TNBC (chemo, approved 1996 for breast cancer).
    # Fix: replace $374→$461, $375→$462, $376→$463 in formula.
    # Formula: COL$35*(COL$461/(COL$461+COL$462+COL$463))
    log.info("R36 (Taxotere TNBC): fixing incidence references")
    for col in ['T', 'U', 'V', 'W', 'X']:
        new = (f'<c r="{col}36" s="405">'
               f'<f>{col}$35*({col}$461/({col}$461+{col}$462+{col}$463))</f>'
               f'</c>')
        xml = _replace_cell(xml, 36, col, new)
        changes += 1

    # ── R37-R38: Taxotere BRCA/BLCA ─────────────────────────────────────
    # Shared formula si=12, master at T37, ref=T37:X38.
    # Master formula: T$35*(T462/(T$461+T$462+T$463))
    #   - Numerator T462 (no $) → shifts to T463 for R38 (BLCA). Correct!
    #   - Denominator T$461+T$462+T$463 (with $) → absolute. Correct!
    log.info("R37-38 (Taxotere BRCA/BLCA): fixing shared formula master")
    # Replace master cell T37
    new_master = (
        '<c r="T37" s="405">'
        '<f t="shared" ref="T37:X38" si="12">'
        'T$35*(T462/(T$461+T$462+T$463))</f>'
        '</c>')
    xml = _replace_cell(xml, 37, 'T', new_master)
    changes += 1

    # Fix shared reference cells U37-X37 and T38-X38 (remove t="e" error attr)
    for rn in [37, 38]:
        cols = ['U', 'V', 'W', 'X'] if rn == 37 else ['T', 'U', 'V', 'W', 'X']
        for col in cols:
            new = (f'<c r="{col}{rn}" s="405">'
                   f'<f t="shared" si="12"/></c>')
            xml = _replace_cell(xml, rn, col, new)
            changes += 1

    # ── R40: Evista TNBC ────────────────────────────────────────────────
    # Evista (raloxifene) is a SERM targeting estrogen receptors.
    # TNBC has NO estrogen receptors → Evista has 0 TNBC revenue.
    # Breast cancer risk reduction approved 2007 (ER+ only).
    log.info("R40 (Evista TNBC): setting to 0 (not approved for TNBC)")
    for col in ['T', 'U', 'V', 'W', 'X']:
        new = f'<c r="{col}40" s="405"><v>0</v></c>'
        xml = _replace_cell(xml, 40, col, new)
        changes += 1

    # ── R41: Evista BRCA ────────────────────────────────────────────────
    # 100% of Evista breast cancer revenue is ER+ (BRCA category).
    log.info("R41 (Evista BRCA): setting to drug total (=COL$39)")
    for col in ['T', 'U', 'V', 'W', 'X']:
        new = (f'<c r="{col}41" s="405">'
               f'<f>{col}$39</f></c>')
        xml = _replace_cell(xml, 41, col, new)
        changes += 1

    # ── R43: Femara TNBC ────────────────────────────────────────────────
    # Femara (letrozole) is an aromatase inhibitor targeting ER+ breast cancer.
    # TNBC has NO estrogen receptors → Femara has 0 TNBC revenue.
    # Approved for breast cancer (ER+ only) since 1997.
    log.info("R43 (Femara TNBC): setting to 0 (not approved for TNBC)")
    for col in ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']:
        new = f'<c r="{col}43" s="405"><v>0</v></c>'
        xml = _replace_cell(xml, 43, col, new)
        changes += 1

    # ── R44: Femara BRCA ────────────────────────────────────────────────
    # 100% of Femara revenue is ER+ breast cancer (BRCA category).
    # Master shared formula si=14, ref=F44:R44.
    log.info("R44 (Femara BRCA): setting to drug total (=COL$42)")
    # Master cell F44
    new_master = (
        '<c r="F44" s="405">'
        '<f t="shared" ref="F44:R44" si="14">F$42</f>'
        '</c>')
    xml = _replace_cell(xml, 44, 'F', new_master)
    changes += 1

    # Shared reference cells G44-R44
    for col in ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']:
        new = (f'<c r="{col}44" s="405">'
               f'<f t="shared" si="14"/></c>')
        xml = _replace_cell(xml, 44, col, new)
        changes += 1

    log.info(f"Total cell replacements: {changes}")
    return xml


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Fix #DIV/0! errors in TAM Solid rows 35-44")
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

    new_xml = process_sheet(xml)

    if args.dry_run:
        # Verify: count remaining #DIV/0! in rows 35-44
        old_errs = sum(1 for m in re.finditer(r'<c r="[A-Z]+3[5-9]"[^>]*t="e"', xml))
        old_errs += sum(1 for m in re.finditer(r'<c r="[A-Z]+4[0-4]"[^>]*t="e"', xml))
        new_errs = sum(1 for m in re.finditer(r'<c r="[A-Z]+3[5-9]"[^>]*t="e"', new_xml))
        new_errs += sum(1 for m in re.finditer(r'<c r="[A-Z]+4[0-4]"[^>]*t="e"', new_xml))
        log.info(f"[DRY-RUN] Error cells in R35-R44: {old_errs} -> {new_errs}")

        # Show sample of fixed cells
        for rn in [36, 37, 40, 41, 43, 44]:
            for col in ['F', 'T']:
                m = re.search(rf'(<c r="{col}{rn}"[^/]*(?:>.*?</c>|/>))', new_xml)
                if m:
                    cell = m.group(1)
                    if '<f' in cell or '<v' in cell:
                        print(f"  {col}{rn}: {cell[:100]}")
        return

    # Backup
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_divfix_{ts}.xlsx")
    shutil.copy2(xlsx_path, backup)
    log.info(f"Backup: {backup}")

    # Surgical zip patch
    modified = {sheet_zip: new_xml.encode("utf-8")}

    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~divfix.xlsx")
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
    print("TAM #DIV/0! Fix Complete")
    print(f"  File: {xlsx_path}")
    print(f"  Taxotere TNBC: incidence-based split (approved 1996)")
    print(f"  Evista TNBC: 0 (SERM, ER+ only)")
    print(f"  Femara TNBC: 0 (aromatase inhibitor, ER+ only)")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
