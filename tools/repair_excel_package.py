#!/usr/bin/env python3
"""Repair Excel-hostile OOXML artifacts without changing workbook model cells.

The DCF workflow does a lot of surgical XML patching.  Excel is stricter than
LibreOffice/openpyxl for some optional OOXML objects, especially chart caches:
if a chart series is declared as a numeric cache but contains strings, Excel
will show "We found a problem with some content" and repair the workbook.

This tool keeps the workbook formulas/sheets intact and normalizes optional
objects that are safe to drop or rewrite:
  * chart category numCache values that are strings -> strCache
  * LibreOffice-expanded TABLE(...) formulas -> Excel dataTable markers
  * LibreOffice workbook extLst metadata
  * x14 worksheet conditional-formatting extensions
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import time
import warnings
import zipfile
from pathlib import Path

import openpyxl
from lxml import etree


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NSMAP = {"main": NS_MAIN, "c": NS_CHART}


def _default_path(ticker: str) -> Path:
    return Path(f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx")


def _is_float_text(value: str | None) -> bool:
    if value is None:
        return False
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def _parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(remove_blank_text=False, recover=False)
    return etree.fromstring(data, parser=parser)


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _convert_string_category_caches(root: etree._Element) -> int:
    """Convert invalid c:cat/c:numRef/c:numCache string values to strCache."""
    changed = 0
    for cat in root.xpath(".//c:cat", namespaces=NSMAP):
        num_ref = cat.find(f"{{{NS_CHART}}}numRef")
        if num_ref is None:
            continue
        cache = num_ref.find(f"{{{NS_CHART}}}numCache")
        if cache is None:
            continue
        values = [
            v.text
            for v in cache.findall(f".//{{{NS_CHART}}}pt/{{{NS_CHART}}}v")
        ]
        if not values or all(_is_float_text(v) for v in values):
            continue

        # In a chart category axis, strings belong in strRef/strCache.  Keep
        # the formula and points; only change the declared cache type.
        num_ref.tag = f"{{{NS_CHART}}}strRef"
        cache.tag = f"{{{NS_CHART}}}strCache"
        for format_code in cache.findall(f"{{{NS_CHART}}}formatCode"):
            cache.remove(format_code)
        changed += 1
    return changed


def _remove_extlst(root: etree._Element, namespace: str) -> int:
    removed = 0
    for extlst in list(root.xpath(f".//ns:extLst", namespaces={"ns": namespace})):
        parent = extlst.getparent()
        if parent is not None:
            parent.remove(extlst)
            removed += 1
    return removed


def _repair_valuation_data_tables(root: etree._Element) -> int:
    """Rebuild DCF template data-table formulas that LibreOffice expands badly.

    Excel stores these sensitivity tables as three `t="dataTable"` formula
    markers. LibreOffice can rewrite them as normal `TABLE(...)` formulas in
    every output cell. Those normal formulas are not accepted by Excel during
    workbook load.
    """
    table_cells = []
    for cell in root.xpath(".//main:c[main:f]", namespaces=NSMAP):
        formula = cell.find(f"{{{NS_MAIN}}}f")
        text = (formula.text or "").strip() if formula is not None else ""
        if text.upper().startswith("TABLE("):
            table_cells.append(cell)

    if not table_cells:
        return 0

    table_refs = {cell.get("r") for cell in table_cells}
    for cell in table_cells:
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if formula is not None:
            cell.remove(formula)
        if cell.find(f"{{{NS_MAIN}}}v") is None:
            value = etree.SubElement(cell, f"{{{NS_MAIN}}}v")
            value.text = "0"

    # Canonical data-table markers from the DCF template's VALUATION tab.
    # The waterfall table can expand beyond row 30 when a ticker has more
    # drug×indication cumulative scenarios; detect the last scenario-id row in
    # column I and preserve that expanded data-table range.
    waterfall_last = 30
    for cell in root.xpath('.//main:c[starts-with(@r,"I")]', namespaces=NSMAP):
        ref = cell.get("r", "")
        match = re.fullmatch(r"I(\d+)", ref)
        if not match:
            continue
        row = int(match.group(1))
        if row < 26 or row > 60:
            continue
        value = cell.find(f"{{{NS_MAIN}}}v")
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if value is not None or formula is not None:
            waterfall_last = max(waterfall_last, row)

    # Detect the embedded Catalyst scenario list in B9:B*. The results now live
    # beside it in C9:C*, with terminal growth in C8; VALUATION O:P is retired.
    embedded_catalyst_last = 8
    for cell in root.xpath('.//main:c[starts-with(@r,"B")]', namespaces=NSMAP):
        ref = cell.get("r", "")
        match = re.fullmatch(r"B(\d+)", ref)
        if not match:
            continue
        row = int(match.group(1))
        if row < 9 or row > 520:
            continue
        value = cell.find(f"{{{NS_MAIN}}}v")
        formula = cell.find(f"{{{NS_MAIN}}}f")
        if value is not None or formula is not None:
            embedded_catalyst_last = max(embedded_catalyst_last, row)

    markers = {}
    if table_refs & {"G5", "M5", "J26"}:
        markers.update({
            "G5": {"t": "dataTable", "ref": "G5:J9", "dt2D": "1", "dtr": "1", "r1": "C3", "r2": "C5", "ca": "1"},
            "M5": {"t": "dataTable", "ref": "M5", "dt2D": "1", "dtr": "1", "r1": "C5", "r2": "C3", "ca": "1"},
            "J26": {"t": "dataTable", "ref": f"J26:J{waterfall_last}", "dt2D": "1", "dtr": "1", "r1": "C5", "r2": "C3", "ca": "1"},
        })
    if "C9" in table_refs and embedded_catalyst_last >= 9:
        markers["C9"] = {
            "t": "dataTable", "ref": f"C9:C{embedded_catalyst_last}",
            "dt2D": "1", "dtr": "1", "r1": "C6",
            "r2": "B6", "ca": "1",
        }
    rebuilt = 0
    for ref, attrs in markers.items():
        matches = root.xpath(f'.//main:c[@r="{ref}"]', namespaces=NSMAP)
        if not matches:
            continue
        cell = matches[0]
        existing = cell.find(f"{{{NS_MAIN}}}f")
        if existing is not None:
            cell.remove(existing)
        formula = etree.Element(f"{{{NS_MAIN}}}f", **attrs)
        insert_at = 0
        for idx, child in enumerate(list(cell)):
            if child.tag in {f"{{{NS_MAIN}}}v", f"{{{NS_MAIN}}}is"}:
                insert_at = idx
                break
        cell.insert(insert_at, formula)
        rebuilt += 1

    return len(table_cells) + rebuilt


def _cell_ref_parts(ref: str | None) -> tuple[int, int] | None:
    if not ref:
        return None
    match = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not match:
        return None
    col = 0
    for ch in match.group(1):
        col = col * 26 + ord(ch) - 64
    return int(match.group(2)), col


def _repair_cells_in_wrong_rows(root: etree._Element) -> tuple[int, int]:
    """Move cells back to the row indicated by their own r="A1" address.

    A bad regex patch can turn a sequence like row 27, row 28, row 29 into
    one schema-invalid row 27 containing B28/C29/... cells.  Excel rejects
    that structure even though tolerant readers can load it.  The cell address
    is authoritative, so re-home those cells and sort rows/cells afterward.
    """
    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        return 0, 0

    rows = [child for child in sheet_data if child.tag == f"{{{NS_MAIN}}}row"]
    row_by_num: dict[int, etree._Element] = {}
    for row in rows:
        try:
            row_num = int(row.get("r", ""))
        except ValueError:
            continue
        row_by_num[row_num] = row

    moved = 0
    created = 0
    touched: set[etree._Element] = set()

    for row in list(rows):
        try:
            parent_row_num = int(row.get("r", ""))
        except ValueError:
            continue
        for cell in list(row):
            if cell.tag != f"{{{NS_MAIN}}}c":
                continue
            parts = _cell_ref_parts(cell.get("r"))
            if parts is None:
                continue
            cell_row_num, _ = parts
            if cell_row_num == parent_row_num:
                continue

            row.remove(cell)
            target = row_by_num.get(cell_row_num)
            if target is None:
                target = etree.Element(f"{{{NS_MAIN}}}row", r=str(cell_row_num))
                row_by_num[cell_row_num] = target
                sheet_data.append(target)
                created += 1
            target.append(cell)
            touched.add(row)
            touched.add(target)
            moved += 1

    if not moved:
        return 0, 0

    def row_num(row: etree._Element) -> int:
        try:
            return int(row.get("r", "0"))
        except ValueError:
            return 0

    def cell_col(cell: etree._Element) -> int:
        parts = _cell_ref_parts(cell.get("r"))
        return parts[1] if parts else 0

    for row in touched:
        cells = [child for child in row if child.tag == f"{{{NS_MAIN}}}c"]
        others = [child for child in row if child.tag != f"{{{NS_MAIN}}}c"]
        for child in list(row):
            row.remove(child)
        for cell in sorted(cells, key=cell_col):
            row.append(cell)
        for child in others:
            row.append(child)

    non_rows = [child for child in sheet_data if child.tag != f"{{{NS_MAIN}}}row"]
    for child in list(sheet_data):
        sheet_data.remove(child)
    for row in sorted(row_by_num.values(), key=row_num):
        sheet_data.append(row)
    for child in non_rows:
        sheet_data.append(child)

    return moved, created


def _repair_parts(parts: dict[str, bytes]) -> tuple[dict[str, bytes], dict[str, int]]:
    report = {
        "chart_string_caches": 0,
        "chart_extlst_removed": 0,
        "data_table_formulas_rebuilt": 0,
        "worksheet_extlst_removed": 0,
        "workbook_extlst_removed": 0,
        "worksheet_cells_rehomed": 0,
        "worksheet_rows_created": 0,
    }
    fixed: dict[str, bytes] = {}

    for name, data in parts.items():
        if not name.endswith(".xml"):
            fixed[name] = data
            continue

        if name.startswith("xl/charts/"):
            root = _parse_xml(data)
            report["chart_string_caches"] += _convert_string_category_caches(root)
            report["chart_extlst_removed"] += _remove_extlst(root, NS_CHART)
            fixed[name] = _serialize_xml(root)
            continue

        if name == "xl/workbook.xml":
            root = _parse_xml(data)
            report["workbook_extlst_removed"] += _remove_extlst(root, NS_MAIN)
            fixed[name] = _serialize_xml(root)
            continue

        if name.startswith("xl/worksheets/"):
            root = _parse_xml(data)
            moved, created = _repair_cells_in_wrong_rows(root)
            report["worksheet_cells_rehomed"] += moved
            report["worksheet_rows_created"] += created
            report["data_table_formulas_rebuilt"] += _repair_valuation_data_tables(root)
            report["worksheet_extlst_removed"] += _remove_extlst(root, NS_MAIN)
            fixed[name] = _serialize_xml(root)
            continue

        fixed[name] = data

    return fixed, report


def _read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"zip CRC failure in {bad}")
        return {name: zf.read(name) for name in zf.namelist()}


def _write_parts(path: Path, parts: dict[str, bytes]) -> None:
    tmp = path.with_suffix(".~excelrepair.xlsx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    tmp.replace(path)


def _validate_openpyxl_clean(path: Path) -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        wb = openpyxl.load_workbook(path, read_only=False, data_only=False, keep_links=True)
        wb.close()


def repair(path: Path, make_backup: bool = True) -> dict[str, int]:
    if not path.exists():
        raise FileNotFoundError(path)
    parts = _read_parts(path)
    fixed, report = _repair_parts(parts)

    if make_backup:
        backup = path.with_name(f"{path.stem}_pre_excel_package_repair_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")

    _write_parts(path, fixed)

    # Basic structural and reader checks after write.
    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"post-write zip CRC failure in {bad}")
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                etree.fromstring(zf.read(name))

    _validate_openpyxl_clean(path)
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description="Repair Excel-hostile OOXML artifacts in a DCF workbook")
    ap.add_argument("--ticker", help="Ticker used to infer /mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx")
    ap.add_argument("--path", help="Explicit workbook path")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    if not args.path and not args.ticker:
        ap.error("provide --path or --ticker")

    path = Path(args.path) if args.path else _default_path(args.ticker.upper())
    report = repair(path, make_backup=not args.no_backup)
    print(f"Repaired: {path}")
    for key, value in report.items():
        print(f"  {key}: {value}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
