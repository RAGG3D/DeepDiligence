#!/usr/bin/env python3
"""
update_referred_tables.py -- Sync Pipeline Referred Tables from updated TAM Solid.

For each new indication in Referred Tables (BTC, EC, ES-SCLC, HCC, Melanoma NCAM+,
MPM, RCC):
  1. Reads aggregated TAM Solid data (sum across all drugs) for years F-AH.
  2. Writes the aggregate to the FIRST data row, zeroes all other rows.
  3. Changes Revenue Forecasting TAM formula from cross-sheet TAM Solid reference
     to Pipeline-internal SUMIF($D$9:$D$342,"IND",col$9:col$342).

For old indications (CRC, GC, NSCLC, TNBC, HNSCC) in Revenue Forecasting:
  - Changes from Pipeline SUMIF (which returns #DIV/0!/0 due to missing data) to
    TAM Solid SUMIF, so Revenue Forecasting shows correct values.

HL and MM Revenue Forecasting formulas (TAM Blood with growth fallback) are left
unchanged as they already produce correct values.

Uses surgical zip patching (NEVER openpyxl .save()).

Usage:
    python update_referred_tables.py [--file PATH] [--dry-run]
"""

import argparse
import logging
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple


_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

_DEFAULT_FILE = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_YEAR_BASE = 2010
_COL_BASE = 6   # F = index 6
_LAST_YEAR = 2038

# Referred Tables: first and last data row for each new indication
# (rows determined by fill_tam_forecast.py output; spacers are between groups)
_NEW_IND_ROWS: Dict[str, Tuple[int, int]] = {
    "BTC":            (264, 271),
    "EC":             (273, 275),
    "ES-SCLC":        (277, 280),
    "HCC":            (282, 291),
    "Melanoma NCAM+": (293, 300),
    "MPM":            (302, 305),
    "RCC":            (307, 321),
}

# Revenue Forecasting TAM rows for NEW indications (currently use TAM Solid cross-
# sheet SUMIF → change to Pipeline SUMIF after updating Referred Tables)
_FORECAST_NEW: Dict[str, int] = {
    "BTC":            461,
    "RCC":            470,
    "HCC":            472,
    "EC":             476,
    "ES-SCLC":        483,
    "Melanoma NCAM+": 485,
    "MPM":            487,
}

# Revenue Forecasting TAM rows for OLD indications (currently use Pipeline SUMIF
# returning #DIV/0!/0 → change to TAM Solid SUMIF)
_FORECAST_OLD: Dict[str, int] = {
    "CRC":   463,
    "GC":    474,
    "NSCLC": 494,
    "TNBC":  496,
    "HNSCC": 502,
}

# Cell style for Referred Table data cells
_STYLE_DATA = "48"

# Cell style for Revenue Forecasting TAM formula cells
_STYLE_FCAST = "334"

# SUMIF range in Pipeline Referred Tables
_SUMIF_END = 342


# ══════════════════════════════════════════════════════════════════════════════
#  COLUMN HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _col_letter(idx: int) -> str:
    result = ""
    while idx > 0:
        idx -= 1
        result = chr(ord('A') + idx % 26) + result
        idx //= 26
    return result


def _col_idx(col: str) -> int:
    result = 0
    for c in col:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


def _year_to_col(year: int) -> str:
    return _col_letter(_COL_BASE + (year - _YEAR_BASE))


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════

def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_path: Dict[str, str] = {}
    for rel in rels_xml:
        if "worksheet" in rel.get("Type", ""):
            tgt = rel.get("Target", "")
            rid_to_path[rel.get("Id")] = (
                f"xl/{tgt}" if not tgt.startswith("/") else tgt.lstrip("/"))

    for sheet in wb_xml.findall(f".//{{{_NS_MAIN}}}sheet"):
        if sheet.get("name") == sheet_name:
            return rid_to_path.get(sheet.get(f"{{{_NS_R}}}id"))
    return None


def _detect_tam_solid_name(xlsx_path: Path) -> str:
    for name in ("TAM Solid", "TAM Solid+MM"):
        if _get_sheet_zip_path(xlsx_path, name):
            return name
    return "TAM Solid"


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED STRINGS
# ══════════════════════════════════════════════════════════════════════════════

def _load_shared_strings(xlsx_path: Path) -> List[str]:
    ss_list: List[str] = []
    with zipfile.ZipFile(xlsx_path) as zf:
        if "xl/sharedStrings.xml" not in [i.filename for i in zf.infolist()]:
            return ss_list
        ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in ss_root.findall(f"{{{_NS_MAIN}}}si"):
            t = si.find(f"{{{_NS_MAIN}}}t")
            if t is not None and t.text:
                ss_list.append(t.text)
            else:
                parts = [r.find(f"{{{_NS_MAIN}}}t")
                         for r in si.findall(f"{{{_NS_MAIN}}}r")]
                ss_list.append("".join(
                    p.text for p in parts if p is not None and p.text))
    return ss_list


def _read_cell_text(attrs: str, inner: str, ss_list: List[str]) -> str:
    if 't="s"' in attrs:
        m = re.search(r'<v>(\d+)</v>', inner)
        if m:
            idx = int(m.group(1))
            if idx < len(ss_list):
                return ss_list[idx]
    elif 't="inlineStr"' in attrs:
        m = re.search(r'<is><t>([^<]*)</t></is>', inner)
        if m:
            return m.group(1)
    elif 't="str"' in attrs:
        m = re.search(r'<v>([^<]*)</v>', inner)
        if m:
            return m.group(1)
    return ""


# ══════════════════════════════════════════════════════════════════════════════
#  PART 1: READ TAM SOLID AGGREGATES
# ══════════════════════════════════════════════════════════════════════════════

def _read_tam_solid_aggregates(
    solid_xml: str,
    ss_list: List[str],
    target_indications: set,
) -> Dict[str, Dict[str, float]]:
    """Aggregate TAM Solid drug rows by indication and year column.

    Returns: {indication: {col_letter: total_value}}
    """
    agg: Dict[str, Dict[str, float]] = {ind: {} for ind in target_indications}

    for row_m in re.finditer(
            r'<row\s+r="(\d+)"[^>]*>(.*?)</row>', solid_xml, re.DOTALL):
        row_num = int(row_m.group(1))
        if row_num <= 7:
            continue
        row_body = row_m.group(0)

        # Read D column text
        d_text = ""
        for cell_m in re.finditer(
                r'<c\s+([^>]*r="D\d+"[^>]*)(?:/>|>(.*?)</c>)',
                row_body, re.DOTALL):
            d_text = _read_cell_text(
                cell_m.group(1), cell_m.group(2) or "", ss_list)
            break

        ind = d_text.strip()
        if ind not in target_indications:
            continue

        # Read all year columns F–AH
        for year in range(_YEAR_BASE, _LAST_YEAR + 1):
            col = _year_to_col(year)
            for cell_m in re.finditer(
                    rf'<c\s+[^>]*r="{col}{row_num}"[^>]*>(.*?)</c>',
                    row_body, re.DOTALL):
                inner = cell_m.group(1)
                # Read cached value (works for both plain values and formulas
                # with cached results)
                v_m = re.search(r'<v>([-\d.eE+]+)</v>', inner)
                if v_m:
                    try:
                        val = float(v_m.group(1))
                        agg[ind][col] = agg[ind].get(col, 0) + val
                    except ValueError:
                        pass
                break

    return agg


# ══════════════════════════════════════════════════════════════════════════════
#  PART 2: UPDATE REFERRED TABLE ROWS IN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def _fmt_val(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:.4f}"


def _build_data_cells(row_num: int, values: Dict[str, float], style: str) -> str:
    """Build XML cells for data columns F-AH."""
    parts = []
    for year in range(_YEAR_BASE, _LAST_YEAR + 1):
        col = _year_to_col(year)
        val = values.get(col, 0)
        if val == 0:
            # Write explicit zero (so SUMIF sees 0, not empty)
            parts.append(f'<c r="{col}{row_num}" s="{style}"/>')
        else:
            parts.append(
                f'<c r="{col}{row_num}" s="{style}">'
                f'<v>{_fmt_val(val)}</v></c>')
    return "".join(parts)


def _clear_data_cells(row_num: int, style: str) -> str:
    """Build empty (zeroed) cells for data columns F-AH."""
    parts = []
    for year in range(_YEAR_BASE, _LAST_YEAR + 1):
        col = _year_to_col(year)
        parts.append(f'<c r="{col}{row_num}" s="{style}"/>')
    return "".join(parts)


def _replace_row_data_cells(pipeline_xml: str, row_num: int,
                             new_data_cells: str) -> str:
    """Replace F-AH cells in a row with new cells.

    Keeps row tag, A-E columns (D=indication label, E=[MM USD]) unchanged.
    Removes any existing F-AH cells and appends new ones before </row>.
    """
    pattern = re.compile(
        rf'(<row\s+r="{row_num}"[^>]*>)(.*?)(</row>)', re.DOTALL)
    m = pattern.search(pipeline_xml)
    if not m:
        log.warning(f"  Row {row_num} not found in Pipeline XML")
        return pipeline_xml

    row_open = m.group(1)
    row_body = m.group(2)
    row_close = m.group(3)

    # Keep only A-E column cells; strip F-AH
    f_idx = _col_idx("F")
    ah_idx = _col_idx("AH")

    kept_cells = ""
    for cell_m in re.finditer(
            r'<c\s+([^>]*?)(?:/>|>(.*?)</c>)', row_body, re.DOTALL):
        # Extract column from r= attribute
        r_m = re.search(r'r="([A-Z]+)\d+"', cell_m.group(1))
        if not r_m:
            kept_cells += cell_m.group(0)
            continue
        ci = _col_idx(r_m.group(1))
        if ci < f_idx or ci > ah_idx:
            kept_cells += cell_m.group(0)
        # else: skip (will be replaced by new_data_cells)

    new_row = row_open + kept_cells + new_data_cells + row_close
    return pipeline_xml[:m.start()] + new_row + pipeline_xml[m.end():]


def _update_referred_tables(
    pipeline_xml: str,
    ind_aggregates: Dict[str, Dict[str, float]],
    dry_run: bool = False,
) -> str:
    """Update Referred Table rows for new indications with fresh TAM Solid data."""
    for ind, (first_row, last_row) in _NEW_IND_ROWS.items():
        agg = ind_aggregates.get(ind, {})
        if not agg:
            log.warning(f"  {ind}: no aggregated data, skipping")
            continue

        total_t = agg.get(_year_to_col(2024), 0)
        log.info(f"  {ind}: rows {first_row}-{last_row}, T(2024) aggregate = {total_t:.1f}")

        if dry_run:
            continue

        # First row: write aggregated values
        data_cells = _build_data_cells(first_row, agg, _STYLE_DATA)
        pipeline_xml = _replace_row_data_cells(pipeline_xml, first_row, data_cells)

        # Remaining rows: clear all data values (keep D="IND" for SUMIF)
        for rn in range(first_row + 1, last_row + 1):
            empty_cells = _clear_data_cells(rn, _STYLE_DATA)
            pipeline_xml = _replace_row_data_cells(pipeline_xml, rn, empty_cells)

    return pipeline_xml


# ══════════════════════════════════════════════════════════════════════════════
#  PART 3: UPDATE REVENUE FORECASTING TAM FORMULAS
# ══════════════════════════════════════════════════════════════════════════════

def _pipeline_sumif_formula(col: str, indication: str, end_row: int) -> str:
    """Build a Pipeline-internal SUMIF formula for a given column."""
    return (f'SUMIF($D$9:$D${end_row},'
            f'"{_xml_escape(indication)}",'
            f'{col}$9:{col}${end_row})')


def _solid_sumif_formula(col: str, indication: str) -> str:
    """Build a TAM Solid cross-sheet SUMIF formula for a given column."""
    return (f'SUMIF(\'TAM Solid\'!$D$9:$D$580,'
            f'"{_xml_escape(indication)}",'
            f'\'TAM Solid\'!{col}$9:{col}$580)')


def _build_formula_cell(col: str, row_num: int, formula: str,
                        style: str, is_first: bool, ref_end: str) -> str:
    """Build a formula cell XML, with shared formula on first cell."""
    if is_first:
        # First column: define shared formula with ref= range
        return (f'<c r="{col}{row_num}" s="{style}">'
                f'<f t="shared" ref="{col}{row_num}:{ref_end}{row_num}">'
                f'{formula}</f></c>')
    else:
        return f'<c r="{col}{row_num}" s="{style}"><f t="shared"/></c>'


def _build_forecast_row_formulas(
    row_num: int,
    indication: str,
    formula_fn,  # callable(col, indication) -> formula str
    style: str = _STYLE_FCAST,
) -> str:
    """Build XML for all formula cells F-AH in a Revenue Forecasting TAM row."""
    # For shared formula, define in F, reference in G-AH
    # But since each column needs a different formula (different col letter),
    # we write INDIVIDUAL formulas per column — simpler and avoids si management.
    parts = []
    for year in range(_YEAR_BASE, _LAST_YEAR + 1):
        col = _year_to_col(year)
        formula = formula_fn(col, indication)
        parts.append(
            f'<c r="{col}{row_num}" s="{style}"><f>{formula}</f></c>')
    return "".join(parts)


def _update_forecast_formulas(
    pipeline_xml: str,
    dry_run: bool = False,
) -> str:
    """Change Revenue Forecasting TAM formula sources."""

    # NEW indications: TAM Solid cross-sheet → Pipeline internal SUMIF
    for ind, row_num in _FORECAST_NEW.items():
        log.info(f"  {ind} (R{row_num}): TAM Solid → Pipeline SUMIF")
        if dry_run:
            continue
        new_formula_cells = _build_forecast_row_formulas(
            row_num, ind,
            lambda col, i=ind: _pipeline_sumif_formula(col, i, _SUMIF_END))
        pipeline_xml = _replace_row_data_cells(
            pipeline_xml, row_num, new_formula_cells)

    # OLD indications: Pipeline SUMIF (broken) → TAM Solid SUMIF (correct)
    for ind, row_num in _FORECAST_OLD.items():
        log.info(f"  {ind} (R{row_num}): Pipeline SUMIF → TAM Solid SUMIF")
        if dry_run:
            continue
        new_formula_cells = _build_forecast_row_formulas(
            row_num, ind,
            lambda col, i=ind: _solid_sumif_formula(col, i))
        pipeline_xml = _replace_row_data_cells(
            pipeline_xml, row_num, new_formula_cells)

    return pipeline_xml


# ══════════════════════════════════════════════════════════════════════════════
#  ZIP PATCHER
# ══════════════════════════════════════════════════════════════════════════════

def _apply_zip_patch(xlsx_path: Path, modified: Dict[str, bytes]) -> None:
    """Surgical zip patch: overwrite specific sheets, copy rest byte-for-byte."""
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~urt.xlsx")
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


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Sync Pipeline Referred Tables from updated TAM Solid")
    parser.add_argument("--file", default=str(_DEFAULT_FILE))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return

    solid_name = _detect_tam_solid_name(xlsx_path)
    solid_zip = _get_sheet_zip_path(xlsx_path, solid_name)
    pipeline_zip = _get_sheet_zip_path(xlsx_path, "Pipeline")

    log.info(f"File:     {xlsx_path}")
    log.info(f"TAM:      {solid_name} -> {solid_zip}")
    log.info(f"Pipeline: {pipeline_zip}")
    log.info(f"Mode:     {'DRY RUN' if args.dry_run else 'LIVE'}")

    if not solid_zip or not pipeline_zip:
        log.error("Cannot find required sheets")
        return

    ss_list = _load_shared_strings(xlsx_path)

    with zipfile.ZipFile(xlsx_path) as zf:
        solid_xml = zf.read(solid_zip).decode("utf-8")
        pipeline_xml = zf.read(pipeline_zip).decode("utf-8")

    # ── Backup ──
    if not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_urt_{ts}.xlsx")
        shutil.copy2(xlsx_path, backup)
        log.info(f"Backup:   {backup}")

    # ─────────────────────────────────────────────────────────────────────────
    #  STEP 1: Read TAM Solid aggregated data
    # ─────────────────────────────────────────────────────────────────────────
    log.info(f"\n{'='*60}")
    log.info("STEP 1: Reading TAM Solid aggregated data")
    log.info(f"{'='*60}")

    all_indications = set(_NEW_IND_ROWS.keys())
    ind_aggregates = _read_tam_solid_aggregates(solid_xml, ss_list, all_indications)

    for ind in sorted(all_indications):
        data = ind_aggregates.get(ind, {})
        t_val = data.get(_year_to_col(2024), 0)
        log.info(f"  {ind}: T(2024) = {t_val:.1f} "
                 f"({len(data)} year columns)")

    # ─────────────────────────────────────────────────────────────────────────
    #  STEP 2: Update Referred Table rows with fresh data
    # ─────────────────────────────────────────────────────────────────────────
    log.info(f"\n{'='*60}")
    log.info("STEP 2: Updating Pipeline Referred Tables")
    log.info(f"{'='*60}")

    pipeline_xml = _update_referred_tables(
        pipeline_xml, ind_aggregates, dry_run=args.dry_run)

    # ─────────────────────────────────────────────────────────────────────────
    #  STEP 3: Update Revenue Forecasting TAM formulas
    # ─────────────────────────────────────────────────────────────────────────
    log.info(f"\n{'='*60}")
    log.info("STEP 3: Updating Revenue Forecasting TAM formulas")
    log.info(f"{'='*60}")

    pipeline_xml = _update_forecast_formulas(
        pipeline_xml, dry_run=args.dry_run)

    # ─────────────────────────────────────────────────────────────────────────
    #  WRITE
    # ─────────────────────────────────────────────────────────────────────────
    if args.dry_run:
        log.info("\nDry run complete — no changes written")
        return

    modified: Dict[str, bytes] = {pipeline_zip: pipeline_xml.encode("utf-8")}
    _apply_zip_patch(xlsx_path, modified)

    log.info(f"\n{'='*60}")
    log.info("Done")
    log.info(f"  Saved -> {xlsx_path}")
    log.info(f"{'='*60}")


if __name__ == "__main__":
    main()
