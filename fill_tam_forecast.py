#!/usr/bin/env python3
"""
fill_tam_forecast.py -- Fill TAM Solid/Blood empty forecast cells (2024E-2038E)
and add new indication rows to Pipeline Referred Tables.

Part 1: For each drug data row in TAM Solid with empty cells in T-AH,
        classify the trend and fill with a flat forecast value.
Part 2: Read TAM data for new indications (BTC, EC, ES-SCLC, HCC, etc.)
        and add rows to Pipeline Referred Tables so Revenue Forecasting
        can use internal SUMIF instead of cross-sheet refs.

Uses surgical zip-patching (NEVER openpyxl .save()).

Usage:
    python fill_tam_forecast.py [--dry-run] [--file PATH] [--tam-only] [--pipeline-only]
"""

import argparse
import logging
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple


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

# Year <-> column mapping: F=2010, G=2011, ..., S=2023, T=2024, ..., AH=2038
_YEAR_BASE = 2010
_COL_BASE = 6  # Column F = index 6
_FIRST_FORECAST_YEAR = 2024
_LAST_YEAR = 2038

# Column letter for first forecast col
_FIRST_FORECAST_COL = "T"   # col index 20 = 6 + (2024-2010)
_LAST_COL = "AH"            # col index 34 = 6 + (2038-2010)

# New indications to add to Pipeline Referred Tables
# Maps indication abbreviation -> TAM sheet type ("solid" or "blood")
_NEW_INDICATIONS = {
    "BTC":            "solid",
    "EC":             "solid",
    "ES-SCLC":        "solid",
    "HCC":            "solid",
    "Melanoma NCAM+": "solid",
    "MPM":            "solid",
    "RCC":            "solid",
    "HL":             "blood",
    "MM":             "blood",
}

# TAM Blood has different column layout: data starts at J=2013
_BLOOD_YEAR_BASE = 2013
_BLOOD_COL_BASE = 10  # Column J = index 10


# ══════════════════════════════════════════════════════════════════════════════
#  COLUMN HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _col_letter(col_idx: int) -> str:
    """1-based col index -> letter(s). 1=A, 26=Z, 27=AA."""
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(ord('A') + col_idx % 26) + result
        col_idx //= 26
    return result


def _col_idx(col_letter: str) -> int:
    """Column letter -> 1-based index. A=1, Z=26, AA=27."""
    result = 0
    for c in col_letter:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


def _year_to_col(year: int, year_base: int = _YEAR_BASE,
                 col_base: int = _COL_BASE) -> str:
    """Year -> column letter."""
    return _col_letter(col_base + (year - year_base))


def _col_to_year(col: str, year_base: int = _YEAR_BASE,
                 col_base: int = _COL_BASE) -> int:
    """Column letter -> year."""
    return year_base + (_col_idx(col) - col_base)


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════

def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
    """Find the zip path for a named sheet."""
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


def _detect_tam_solid_name(xlsx_path: Path) -> str:
    """Try 'TAM Solid' then 'TAM Solid+MM'."""
    for name in ("TAM Solid", "TAM Solid+MM"):
        if _get_sheet_zip_path(xlsx_path, name):
            return name
    return "TAM Solid"


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED STRINGS LOADER
# ══════════════════════════════════════════════════════════════════════════════

def _load_shared_strings(xlsx_path: Path) -> List[str]:
    """Load shared strings table from xlsx."""
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


# ══════════════════════════════════════════════════════════════════════════════
#  CELL TEXT READER (handles both inlineStr and shared string)
# ══════════════════════════════════════════════════════════════════════════════

def _read_cell_text(cell_xml: str, ss_list: List[str]) -> str:
    """Extract text from a cell XML fragment (handles t='s' and t='inlineStr')."""
    if 't="s"' in cell_xml:
        m = re.search(r'<v>(\d+)</v>', cell_xml)
        if m:
            idx = int(m.group(1))
            if idx < len(ss_list):
                return ss_list[idx]
    elif 't="inlineStr"' in cell_xml:
        m = re.search(r'<is><t>([^<]*)</t></is>', cell_xml)
        if m:
            return m.group(1)
    elif 't="str"' in cell_xml:
        m = re.search(r'<v>([^<]*)</v>', cell_xml)
        if m:
            return m.group(1)
    return ""


# ══════════════════════════════════════════════════════════════════════════════
#  PART 1: FILL TAM FORECAST CELLS
# ══════════════════════════════════════════════════════════════════════════════

def _parse_row_cells(row_xml: str) -> Dict[str, Tuple[Optional[float], Optional[str], bool]]:
    """Parse all cells in a row XML.

    Returns: {col_letter: (value_or_None, style_or_None, has_formula)}
    """
    cells: Dict[str, Tuple[Optional[float], Optional[str], bool]] = {}
    for m in re.finditer(r'<c\s([^>]*?)/>|<c\s([^>]*?)>(.*?)</c>', row_xml, re.DOTALL):
        if m.group(1) is not None:
            attrs, inner = m.group(1), ""
        else:
            attrs, inner = m.group(2), m.group(3) or ""

        r_m = re.search(r'r="([A-Z]+)(\d+)"', attrs)
        if not r_m:
            continue
        col = r_m.group(1)
        s_m = re.search(r's="(\d+)"', attrs)
        style = s_m.group(1) if s_m else None
        has_formula = "<f" in inner
        val = None
        v_m = re.search(r'<v>([-\d.eE+]+)</v>', inner)
        if v_m and not has_formula:
            try:
                val = float(v_m.group(1))
            except ValueError:
                pass
        elif v_m and has_formula:
            # Formula with cached value — still treat as formula (skip)
            try:
                val = float(v_m.group(1))
            except ValueError:
                pass
        cells[col] = (val, style, has_formula)
    return cells


def _classify_and_forecast(values: Dict[str, float]) -> Optional[float]:
    """Given {col_letter: value}, return forecast value or None to skip.

    Algorithm:
    - Skip if all values are zero or empty
    - If only 1 data point: carry forward
    - If last 3 points are all declining: use average of ALL nonzero values
    - Otherwise: carry forward last known value (flat/conservative)
    """
    nonzero = {k: v for k, v in values.items() if v != 0}
    if not nonzero:
        return None  # all zero, skip

    sorted_cols = sorted(nonzero.keys(), key=_col_idx)
    vals = [nonzero[c] for c in sorted_cols]

    if len(vals) < 2:
        return vals[0]  # single data point, carry forward

    # Check declining: last 3 consecutive YoY all negative
    recent = vals[-3:] if len(vals) >= 3 else vals
    declining = all(recent[i] > recent[i + 1] for i in range(len(recent) - 1))

    if declining:
        return round(sum(vals) / len(vals), 2)  # average of all data years
    else:
        return vals[-1]  # last known value (flat carry-forward)


def _find_parameters_boundary(xml: str, ss_list: List[str]) -> int:
    """Find the row where the Parameters section starts.

    Returns the row number of the 'Parameters' header, or a large number
    if not found (meaning all rows are drug data).
    """
    for row_m in re.finditer(
            r'<row\s+r="(\d+)"[^>]*>(.*?)</row>', xml, re.DOTALL):
        row_num = int(row_m.group(1))
        row_body = row_m.group(0)
        if row_num < 100:
            continue
        # Check D column for "Parameters" text
        for cell_m in re.finditer(
                r'<c\s+([^>]*r="D\d+"[^>]*)(?:/>|>(.*?)</c>)',
                row_body, re.DOTALL):
            attrs = cell_m.group(1)
            inner = cell_m.group(2) or ""
            text = _read_cell_text(f'<c {attrs}>{inner}</c>', ss_list)
            if text.strip().lower() == "parameters":
                log.info(f"  Parameters section starts at R{row_num}")
                return row_num
    return 99999


def _fill_tam_sheet_forecasts(
    xml: str,
    sheet_name: str,
    data_col_start: str,
    data_col_end: str,
    forecast_col_start: str,
    forecast_col_end: str,
    ss_list: Optional[List[str]] = None,
    dry_run: bool = False,
) -> Tuple[str, int]:
    """Fill empty forecast cells in a TAM sheet.

    Scans all rows for drug data rows (numeric values in data columns,
    no formulas). For each such row with empty forecast cells, generates
    a forecast value and writes it. Stops at the Parameters section boundary.

    Returns: (modified_xml, count_of_cells_filled)
    """
    start_idx = _col_idx(data_col_start)
    end_idx = _col_idx(data_col_end)
    fc_start_idx = _col_idx(forecast_col_start)
    fc_end_idx = _col_idx(forecast_col_end)

    # Detect Parameters section boundary
    params_row = _find_parameters_boundary(xml, ss_list or [])

    total_filled = 0
    rows_filled = 0

    # Process each row
    for row_m in re.finditer(
            r'(<row\s+r="(\d+)"[^>]*>)(.*?)(</row>)', xml, re.DOTALL):
        row_tag = row_m.group(1)
        row_num = int(row_m.group(2))
        row_body = row_m.group(3)
        row_end = row_m.group(4)

        if row_num <= 7:
            continue  # Skip header rows
        if row_num >= params_row:
            continue  # Skip Parameters section

        cells = _parse_row_cells(row_m.group(0))

        # Check if this is a drug data row:
        # - Has at least 1 numeric value in data columns (no formula)
        # - Is not a formula-only row (indication breakdown)
        has_data = False
        has_formula_in_data = False
        data_values: Dict[str, float] = {}
        last_data_style: Optional[str] = None

        for ci in range(start_idx, end_idx + 1):
            col = _col_letter(ci)
            if col not in cells:
                continue
            val, style, is_formula = cells[col]
            if is_formula:
                has_formula_in_data = True
                break
            if val is not None:
                has_data = True
                data_values[col] = val
                last_data_style = style

        if has_formula_in_data or not has_data:
            continue  # Skip formula rows and empty rows

        # Find which forecast columns are empty
        empty_forecast_cols: List[str] = []
        for ci in range(fc_start_idx, fc_end_idx + 1):
            col = _col_letter(ci)
            if col in cells:
                val, _, is_formula = cells[col]
                if is_formula or val is not None:
                    continue  # Already has data or formula
            empty_forecast_cols.append(col)

        if not empty_forecast_cols:
            continue  # Already fully populated

        # Classify trend and get forecast value
        forecast_val = _classify_and_forecast(data_values)
        if forecast_val is None:
            continue

        # Determine style to use (inherit from last data column)
        # For TAM Solid: projected values typically use s=404 (parent) or s=405 (child)
        # Use last_data_style as fallback, prefer s=48 or whatever existing style is
        ref_style = last_data_style or "48"

        if dry_run:
            log.info(f"  R{row_num}: forecast={forecast_val:.1f} "
                     f"({len(empty_forecast_cols)} cols, style={ref_style})")
            rows_filled += 1
            total_filled += len(empty_forecast_cols)
            continue

        # Replace existing empty cells or insert new cells in column order
        old_row = row_m.group(0)
        new_row = old_row

        for col in empty_forecast_cols:
            val_str = str(int(forecast_val)) if forecast_val == int(forecast_val) else f"{forecast_val:.2f}"
            new_cell = f'<c r="{col}{row_num}" s="{ref_style}"><v>{val_str}</v></c>'

            # Check if empty styled cell already exists for this column
            # Matches: <c r="T11" s="408"/>  or  <c r="T11" s="408"></c>
            empty_pat = re.compile(
                rf'<c\s[^>]*r="{col}{row_num}"[^>]*/>'
                rf'|<c\s[^>]*r="{col}{row_num}"[^>]*></c>'
            )
            if empty_pat.search(new_row):
                # Replace existing empty cell with valued cell
                new_row = empty_pat.sub(new_cell, new_row, count=1)
            else:
                # Insert new cell in correct column position
                # Find the right spot (before a cell with higher column index)
                target_ci = _col_idx(col)
                inserted = False
                for cm in re.finditer(r'<c\s[^>]*r="([A-Z]+)\d+"', new_row):
                    cell_ci = _col_idx(cm.group(1))
                    if cell_ci > target_ci:
                        pos = cm.start()
                        new_row = new_row[:pos] + new_cell + new_row[pos:]
                        inserted = True
                        break
                if not inserted:
                    # Append before </row>
                    close_pos = new_row.rfind("</row>")
                    if close_pos != -1:
                        new_row = new_row[:close_pos] + new_cell + new_row[close_pos:]

        xml = xml.replace(old_row, new_row, 1)
        rows_filled += 1
        total_filled += len(empty_forecast_cols)

    log.info(f"  {sheet_name}: filled {total_filled} cells in {rows_filled} rows")
    return xml, total_filled


# ══════════════════════════════════════════════════════════════════════════════
#  PART 2: READ TAM DATA FOR NEW INDICATIONS
# ══════════════════════════════════════════════════════════════════════════════

def _read_tam_indication_data(
    xml: str,
    ss_list: List[str],
    indication: str,
    year_base: int,
    col_base: int,
    target_year_base: int = _YEAR_BASE,
    target_col_base: int = _COL_BASE,
) -> List[Tuple[str, Dict[str, float]]]:
    """Read all drug rows for a given indication from a TAM sheet.

    Searches for rows where column D matches the indication abbreviation.
    Returns: [(drug_name, {col_letter_in_pipeline: value}), ...]

    Column mapping: TAM columns are converted to Pipeline column letters
    so the data can be written directly to Pipeline Referred Tables.
    """
    drugs: List[Tuple[str, Dict[str, float]]] = []

    # Find indication section: scan for rows with D column text matching indication
    in_section = False
    current_drug_name = ""

    for row_m in re.finditer(
            r'<row\s+r="(\d+)"[^>]*>(.*?)</row>', xml, re.DOTALL):
        row_num = int(row_m.group(1))
        row_body = row_m.group(0)

        if row_num <= 7:
            continue

        # Read D column text
        d_text = ""
        for cell_m in re.finditer(
                r'<c\s+([^>]*r="D\d+"[^>]*)(?:/>|>(.*?)</c>)',
                row_body, re.DOTALL):
            attrs = cell_m.group(1)
            inner = cell_m.group(2) or ""
            d_text = _read_cell_text(f'<c {attrs}>{inner}</c>', ss_list)
            break

        if not d_text:
            # Empty row or separator — reset section tracking
            if in_section:
                in_section = False
            continue

        # Check if this is the indication section header or a drug name
        d_lower = d_text.lower().strip()
        ind_lower = indication.lower().strip()

        # Check if D column matches the indication abbreviation exactly
        # (this is how SUMIF works — exact match on D column)
        if d_text.strip() == indication:
            # This is an indication data row — read values
            cells = _parse_row_cells(row_body)
            values: Dict[str, float] = {}

            # Read data from TAM columns, map to Pipeline columns
            for year in range(year_base, _LAST_YEAR + 1):
                tam_col = _col_letter(col_base + (year - year_base))
                pipeline_col = _col_letter(target_col_base + (year - target_year_base))

                if tam_col in cells:
                    val, _, is_formula = cells[tam_col]
                    if is_formula:
                        # Read cached value from formula cells too
                        v_m = re.search(
                            rf'<c\s[^>]*r="{tam_col}{row_num}"[^>]*>.*?<v>([-\d.eE+]+)</v>',
                            row_body, re.DOTALL)
                        if v_m:
                            try:
                                val = float(v_m.group(1))
                            except ValueError:
                                val = None
                    if val is not None and val != 0:
                        values[pipeline_col] = val

            if values:
                # Try to find the drug name from the row above or nearby
                drugs.append((d_text, values))

    return drugs


def _find_drug_names_for_indication(
    xml: str,
    ss_list: List[str],
    indication: str,
) -> Dict[int, str]:
    """Find drug name rows that precede indication data rows.

    Returns: {indication_row_num: drug_name}
    """
    result: Dict[int, str] = {}
    prev_text = ""
    prev_row = 0

    for row_m in re.finditer(
            r'<row\s+r="(\d+)"[^>]*>(.*?)</row>', xml, re.DOTALL):
        row_num = int(row_m.group(1))
        row_body = row_m.group(0)

        if row_num <= 7:
            continue

        d_text = ""
        for cell_m in re.finditer(
                r'<c\s+([^>]*r="D\d+"[^>]*)(?:/>|>(.*?)</c>)',
                row_body, re.DOTALL):
            attrs = cell_m.group(1)
            inner = cell_m.group(2) or ""
            d_text = _read_cell_text(f'<c {attrs}>{inner}</c>', ss_list)
            break

        if d_text.strip() == indication:
            # The drug name is typically the row above or the parent drug row
            result[row_num] = prev_text if prev_text else indication
        elif d_text:
            prev_text = d_text.strip()
            prev_row = row_num

    return result


# ══════════════════════════════════════════════════════════════════════════════
#  PART 2: BUILD PIPELINE REFERRED TABLE ROWS
# ══════════════════════════════════════════════════════════════════════════════

def _build_referred_drug_row(
    row: int,
    drug_name: str,
    values: Dict[str, float],
    style: str = "48",
) -> str:
    """Build a drug name row for Pipeline Referred Tables."""
    cells = [f'<row r="{row}">']
    cells.append(f'<c r="D{row}" s="{style}" t="inlineStr">'
                 f'<is><t>{_xml_escape(drug_name)}</t></is></c>')
    cells.append(f'<c r="E{row}" s="52" t="inlineStr">'
                 f'<is><t>[MM USD]</t></is></c>')

    for col, val in sorted(values.items(), key=lambda x: _col_idx(x[0])):
        val_str = str(int(val)) if val == int(val) else f"{val:.2f}"
        cells.append(f'<c r="{col}{row}" s="{style}"><v>{val_str}</v></c>')

    cells.append('</row>')
    return ''.join(cells)


def _build_referred_indication_row(
    row: int,
    indication: str,
    values: Dict[str, float],
    style: str = "48",
) -> str:
    """Build an indication row for Pipeline Referred Tables (for SUMIF matching)."""
    cells = [f'<row r="{row}">']
    cells.append(f'<c r="D{row}" s="{style}" t="inlineStr">'
                 f'<is><t>{_xml_escape(indication)}</t></is></c>')
    cells.append(f'<c r="E{row}" s="52" t="inlineStr">'
                 f'<is><t>[MM USD]</t></is></c>')

    for col, val in sorted(values.items(), key=lambda x: _col_idx(x[0])):
        val_str = str(int(val)) if val == int(val) else f"{val:.2f}"
        cells.append(f'<c r="{col}{row}" s="{style}"><v>{val_str}</v></c>')

    cells.append('</row>')
    return ''.join(cells)


def _build_spacer_row(row: int) -> str:
    """Build an empty spacer row."""
    return f'<row r="{row}"/>'


# ══════════════════════════════════════════════════════════════════════════════
#  PART 2: INSERT ROWS AND SHIFT DOWNSTREAM
# ══════════════════════════════════════════════════════════════════════════════

def _shift_rows(xml: str, from_row: int, delta: int) -> str:
    """Shift all rows >= from_row by delta positions.
    Updates row r= attributes, cell r= references, and formula row references."""
    if delta == 0:
        return xml

    # Shift row r= attributes
    def shift_row_attr(m):
        full = m.group(0)
        rnum = int(m.group(1))
        if rnum >= from_row:
            return full[:m.start(1) - m.start()] + str(rnum + delta) + full[m.end(1) - m.start():]
        return full

    xml = re.sub(r'<row\s+r="(\d+)"', lambda m: f'<row r="{int(m.group(1)) + delta}"'
                 if int(m.group(1)) >= from_row else m.group(), xml)

    # Shift cell r= references
    def shift_cell_ref(m):
        col = m.group(1)
        rnum = int(m.group(2))
        if rnum >= from_row:
            return f'r="{col}{rnum + delta}"'
        return m.group()

    xml = re.sub(r'r="([A-Z]+)(\d+)"', shift_cell_ref, xml)

    # Shift formula references (both absolute $D$263 and relative D263)
    def shift_formula(m):
        formula = m.group(1)

        def shift_ref(ref_m):
            dollar_col = ref_m.group(1)
            col = ref_m.group(2)
            dollar_row = ref_m.group(3)
            row = int(ref_m.group(4))
            if row >= from_row:
                return f'{dollar_col}{col}{dollar_row}{row + delta}'
            return ref_m.group()

        shifted = re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', shift_ref, formula)
        return f'<f>{shifted}</f>'

    xml = re.sub(r'<f>(.*?)</f>', shift_formula, xml, flags=re.DOTALL)

    # Also handle shared formula ref= attributes
    def shift_shared_ref(m):
        ref = m.group(1)

        def shift_ref2(ref_m):
            col = ref_m.group(1)
            row = int(ref_m.group(2))
            if row >= from_row:
                return f'{col}{row + delta}'
            return ref_m.group()

        shifted = re.sub(r'([A-Z]+)(\d+)', shift_ref2, ref)
        return f'ref="{shifted}"'

    xml = re.sub(r'ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"', shift_shared_ref, xml)

    return xml


def _update_sumif_ranges(xml: str, old_max: int, new_max: int) -> str:
    """Update SUMIF range $D$9:$D$263 -> $D$9:$D${new_max} in all formulas."""
    # Pattern: $D$9:$D$263 or F$9:F$263
    old_range_d = f'$D${old_max}'
    new_range_d = f'$D${new_max}'
    xml = xml.replace(old_range_d, new_range_d)

    # Also update data column ranges: {col}$9:{col}${old_max}
    old_suffix = f'${old_max})'
    new_suffix = f'${new_max})'
    # Be more careful here — only replace within formulas
    # Pattern: col$9:col$263)  ->  col$9:col$new_max)
    xml = re.sub(
        rf'(\$9:[A-Z]+\$){old_max}\)',
        rf'\g<1>{new_max})',
        xml
    )
    return xml


def _detect_sumif_range(pipeline_xml: str) -> int:
    """Detect the current SUMIF range end from existing TAM Summary formulas.

    Searches for SUMIF($D$9:$D$NNN, ...) and returns NNN.
    """
    m = re.search(r'SUMIF\(\$D\$9:\$D\$(\d+)', pipeline_xml)
    if m:
        return int(m.group(1))
    return 263  # default


def _add_pipeline_indication_rows(
    pipeline_xml: str,
    tam_solid_xml: str,
    tam_blood_xml: Optional[str],
    ss_list: List[str],
    dry_run: bool = False,
) -> Tuple[str, int, int]:
    """Add new indication rows to Pipeline Referred Tables.

    Reads drug data from TAM sheets for each new indication,
    builds referred table rows, inserts them, and shifts downstream.

    Returns: (modified_xml, rows_added, new_referred_max_row)
    """
    # Detect current SUMIF range from existing formulas
    sumif_max = _detect_sumif_range(pipeline_xml)
    log.info(f"  Pipeline SUMIF range: $D$9:$D${sumif_max}")

    # Insert new rows right after the current SUMIF range end
    # This places them within the extended SUMIF range
    insert_row = sumif_max + 1  # e.g., R264

    # Collect all new rows to insert
    new_rows: List[str] = []

    for indication, sheet_type in _NEW_INDICATIONS.items():
        if sheet_type == "solid":
            source_xml = tam_solid_xml
            year_base = _YEAR_BASE
            col_base = _COL_BASE
        elif sheet_type == "blood" and tam_blood_xml:
            source_xml = tam_blood_xml
            year_base = _BLOOD_YEAR_BASE
            col_base = _BLOOD_COL_BASE
        else:
            log.warning(f"  {indication}: no TAM {sheet_type} data available, skipping")
            continue

        # Read all drug rows for this indication from TAM sheet
        drug_data = _read_tam_indication_data(
            source_xml, ss_list, indication,
            year_base, col_base,
            target_year_base=_YEAR_BASE,
            target_col_base=_COL_BASE,
        )

        if not drug_data:
            log.warning(f"  {indication}: no drug data found in TAM {sheet_type}")
            continue

        # Also get drug names for the indication rows
        drug_names = _find_drug_names_for_indication(source_xml, ss_list, indication)

        # Aggregate all values for the indication (sum across drugs)
        agg_values: Dict[str, float] = {}
        for drug_name, values in drug_data:
            for col, val in values.items():
                agg_values[col] = agg_values.get(col, 0) + val

        cur = insert_row + len(new_rows)

        # Build rows for this indication
        # For each drug, create: drug_name row + indication row
        # Then add spacer
        unique_drugs_seen: Set[str] = set()

        for drug_name_text, values in drug_data:
            # Get the actual drug name (row above the indication row)
            # drug_name_text is the indication text here
            # We need the parent drug name
            pass  # We'll handle this differently

        # Simpler approach: just add indication-aggregated rows
        # Each drug gets one row with the indication label in D (for SUMIF)
        for i, (_, values) in enumerate(drug_data):
            # Indication data row (D = indication abbrev for SUMIF matching)
            new_rows.append(
                _build_referred_indication_row(cur, indication, values))
            cur += 1

        # Spacer
        new_rows.append(_build_spacer_row(cur))
        cur += 1

        log.info(f"  {indication}: {len(drug_data)} data rows from TAM {sheet_type}")

    n_added = len(new_rows)
    if n_added == 0:
        return pipeline_xml, 0, sumif_max

    new_sumif_max = sumif_max + n_added

    if dry_run:
        log.info(f"  [DRY-RUN] Would insert {n_added} rows at R{insert_row}")
        log.info(f"  [DRY-RUN] SUMIF range: $D$9:$D${sumif_max} -> $D$9:$D${new_sumif_max}")
        return pipeline_xml, n_added, new_sumif_max

    # Shift all rows >= insert_row by n_added
    pipeline_xml = _shift_rows(pipeline_xml, insert_row, n_added)

    # Insert new rows before the shifted content
    insert_marker = f'<row r="{insert_row + n_added}"'
    insert_pos = pipeline_xml.find(insert_marker)
    if insert_pos == -1:
        insert_pos = pipeline_xml.find('</sheetData>')

    new_xml = '\n'.join(new_rows) + '\n'
    pipeline_xml = pipeline_xml[:insert_pos] + new_xml + pipeline_xml[insert_pos:]

    # Update SUMIF ranges: $D$9:$D$263 -> $D$9:$D${263+n_added}
    pipeline_xml = _update_sumif_ranges(pipeline_xml, sumif_max, new_sumif_max)

    # Update dimension
    dim_m = re.search(r'<dimension ref="([^"]*)"', pipeline_xml)
    if dim_m:
        old_dim = dim_m.group(1)
        parts = old_dim.split(":")
        if len(parts) == 2:
            old_max_row = int(re.search(r'\d+', parts[1]).group())
            new_dim_max = old_max_row + n_added
            new_dim_col = re.match(r'[A-Z]+', parts[1]).group()
            new_dim = f'{parts[0]}:{new_dim_col}{new_dim_max}'
            pipeline_xml = pipeline_xml.replace(
                f'<dimension ref="{old_dim}"',
                f'<dimension ref="{new_dim}"',
                1)

    log.info(f"  Pipeline: inserted {n_added} rows at R{insert_row}, "
             f"new SUMIF range: $D$9:$D${new_sumif_max}")
    return pipeline_xml, n_added, new_sumif_max


# ══════════════════════════════════════════════════════════════════════════════
#  SURGICAL ZIP PATCH
# ══════════════════════════════════════════════════════════════════════════════

def _apply_zip_patch(xlsx_path: Path, modified: Dict[str, bytes]) -> None:
    """Write modified sheets back to xlsx using surgical zip patching."""
    # Add fullCalcOnLoad
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~tam_forecast.xlsx")
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
        description="Fill TAM forecast cells and add Pipeline Referred Tables rows"
    )
    parser.add_argument("--file", default=str(_DEFAULT_FILE),
                        help=f"Excel file to modify (default: {_DEFAULT_FILE})")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview changes without writing")
    parser.add_argument("--tam-only", action="store_true",
                        help="Only fill TAM forecast cells (skip Pipeline)")
    parser.add_argument("--pipeline-only", action="store_true",
                        help="Only add Pipeline Referred Tables rows (skip TAM)")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return

    # Discover sheet zip paths
    solid_name = _detect_tam_solid_name(xlsx_path)
    solid_zip = _get_sheet_zip_path(xlsx_path, solid_name)
    blood_zip = _get_sheet_zip_path(xlsx_path, "TAM Blood")
    pipeline_zip = _get_sheet_zip_path(xlsx_path, "Pipeline")

    log.info(f"File: {xlsx_path}")
    log.info(f"TAM Solid: {solid_name} -> {solid_zip}")
    log.info(f"TAM Blood: {blood_zip}")
    log.info(f"Pipeline: {pipeline_zip}")
    log.info(f"Mode: {'DRY RUN' if args.dry_run else 'LIVE'}")

    # Load shared strings
    ss_list = _load_shared_strings(xlsx_path)
    log.info(f"Shared strings: {len(ss_list)} entries")

    # Backup
    if not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_forecast_{ts}.xlsx")
        shutil.copy2(xlsx_path, backup)
        log.info(f"Backup: {backup}")

    modified: Dict[str, bytes] = {}

    # ── Read TAM sheets ──
    tam_solid_xml = None
    tam_blood_xml = None

    if solid_zip:
        with zipfile.ZipFile(xlsx_path) as zf:
            tam_solid_xml = zf.read(solid_zip).decode('utf-8')
    if blood_zip:
        with zipfile.ZipFile(xlsx_path) as zf:
            tam_blood_xml = zf.read(blood_zip).decode('utf-8')

    # ═══════════════════════════════════════════════════════════════════
    #  PART 1: Fill TAM Solid forecast cells
    # ═══════════════════════════════════════════════════════════════════

    if not args.pipeline_only:
        log.info(f"\n{'='*60}")
        log.info("PART 1: Fill TAM forecast cells")
        log.info(f"{'='*60}")

        if tam_solid_xml:
            log.info(f"\nProcessing {solid_name}...")
            tam_solid_xml, n_solid = _fill_tam_sheet_forecasts(
                tam_solid_xml, solid_name,
                data_col_start="F", data_col_end="AH",
                forecast_col_start="T", forecast_col_end="AH",
                ss_list=ss_list,
                dry_run=args.dry_run,
            )
            if n_solid > 0 and not args.dry_run:
                modified[solid_zip] = tam_solid_xml.encode('utf-8')
        else:
            log.warning("TAM Solid sheet not found")

        # TAM Blood: data cols J-T (2013-2023), NO forecast columns exist
        # Blood sheet only has 11 year columns. Forecasting requires
        # extending the sheet structure (columns U-AH don't exist).
        # Skip for now — Pipeline cross-sheet refs handle growth projection.
        if tam_blood_xml:
            log.info(f"\nTAM Blood: only has columns J-T (2013-2023)")
            log.info("  No forecast columns to fill (growth handled by Pipeline formulas)")

    # ═══════════════════════════════════════════════════════════════════
    #  PART 2: Add new indication rows to Pipeline Referred Tables
    # ═══════════════════════════════════════════════════════════════════

    if not args.tam_only and pipeline_zip:
        log.info(f"\n{'='*60}")
        log.info("PART 2: Add new indications to Pipeline Referred Tables")
        log.info(f"{'='*60}")

        with zipfile.ZipFile(xlsx_path) as zf:
            pipeline_xml = zf.read(pipeline_zip).decode('utf-8')

        if tam_solid_xml and tam_blood_xml:
            pipeline_xml, n_added, new_max = _add_pipeline_indication_rows(
                pipeline_xml, tam_solid_xml, tam_blood_xml, ss_list,
                dry_run=args.dry_run,
            )
            if n_added > 0 and not args.dry_run:
                modified[pipeline_zip] = pipeline_xml.encode('utf-8')
        elif tam_solid_xml:
            pipeline_xml, n_added, new_max = _add_pipeline_indication_rows(
                pipeline_xml, tam_solid_xml, None, ss_list,
                dry_run=args.dry_run,
            )
            if n_added > 0 and not args.dry_run:
                modified[pipeline_zip] = pipeline_xml.encode('utf-8')
        else:
            log.warning("Cannot add Pipeline rows: TAM data not available")
            n_added = 0
            new_max = 263

    # ═══════════════════════════════════════════════════════════════════
    #  WRITE
    # ═══════════════════════════════════════════════════════════════════

    if not modified:
        log.info("\nNo changes to write")
        return

    if args.dry_run:
        log.info("\nDry run complete — no changes written")
        return

    _apply_zip_patch(xlsx_path, modified)
    log.info(f"\nSaved -> {xlsx_path}")

    # Summary
    print(f"\n{'='*60}")
    print("TAM Forecast Fill Complete")
    print(f"{'='*60}")
    print(f"  File: {xlsx_path}")
    for sheet_zip, xml_bytes in modified.items():
        if sheet_zip != "xl/workbook.xml":
            print(f"  Modified: {sheet_zip}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
