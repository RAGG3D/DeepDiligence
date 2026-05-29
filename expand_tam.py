#!/usr/bin/env python3
"""
expand_tam.py -- Add new drugs to TAM Solid Oncology Drug Market section.

Inserts new drugs after R341 (Keytruda RCC, last drug row in the Oncology
Drug Market module).  Rows R9-R341 are NOT modified.  Rows R342+ shift down
to make room.  New incidence parameters are inserted after R477 (RCC Incidence).

Oncology Drug Market format:
  - Drug name row: blue text (s=399/48), total Revenue values
  - Breakdown rows: black text (s=400/395), share-based formulas

Pipeline cross-sheet references (growth rows, COGS, SUMIF ranges) are
updated to reflect the shifted TAM Solid row numbers.

Uses surgical zip patching (NEVER openpyxl .save()).

Usage:
    python expand_tam.py --json-dir path [--file path] [--dry-run]
"""

import argparse
import json
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

# Drug and incidence insertion points are defined in expand_tam_sheet:
# _DRUG_INSERT_AFTER = 341 (after Keytruda RCC)
# _INC_INSERT_AFTER = 477 (after RCC Incidence)

# Year → column mapping
_YEAR_BASE = 2005
_COL_BASE = 6  # Column F = index 6


def _col_letter(col_idx: int) -> str:
    """1-based col index → letter(s)."""
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(ord('A') + col_idx % 26) + result
        col_idx //= 26
    return result


def _parse_col(col_str: str) -> int:
    """Column letter(s) → 1-based index."""
    idx = 0
    for c in col_str:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx


def _year_to_col(year: int) -> str:
    """Year → TAM column letter."""
    return _col_letter(_COL_BASE + (year - _YEAR_BASE))


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def _strip_formula_cache(xml: str) -> str:
    """Remove cached ERROR values from formula cells to force recalculation."""
    def _fix(m):
        tag = m.group(0)
        if '<f' in tag:
            tag = re.sub(r' t="e"', '', tag)
            tag = re.sub(r'<v>[^<]*</v>', '', tag)
        else:
            open_tag = re.match(r'<c\b[^>]*', tag).group(0)
            open_tag = re.sub(r' t="e"', '', open_tag)
            tag = open_tag + '/>'
        return tag
    return re.sub(r'<c\b[^>]* t="e"[^>]*>.*?</c>', _fix, xml)


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


# ══════════════════════════════════════════════════════════════════════════════
#  YEAR COLUMN DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def detect_year_columns(xml: str) -> Dict[int, str]:
    """Detect year→column mapping from header row (typically row 6).

    Handles both explicit values (e.g., J6=2013) and formula chains
    (e.g., K6=J6+1, L6=K6+1) which are common in TAM Blood.
    """
    year_cols = {}
    # Look for numeric values 2005-2040 in rows 5-8
    for row_m in re.finditer(r'<row\s+r="([5-8])"[^>]*>(.*?)</row>', xml, re.DOTALL):
        for cell_m in re.finditer(
                r'<c\s+r="([A-Z]+)\d+"[^>]*><v>(\d+\.?\d*)</v></c>',
                row_m.group(2)):
            col = cell_m.group(1)
            try:
                val = int(float(cell_m.group(2)))
            except ValueError:
                continue
            if 2005 <= val <= 2040:
                year_cols[val] = col

    # If we found at least one year, check for formula-based year columns
    if year_cols:
        base_year = min(year_cols.keys())
        base_col = year_cols[base_year]
        base_col_idx = 0
        for i, c in enumerate(base_col):
            base_col_idx = base_col_idx * 26 + (ord(c) - ord('A') + 1)

        max_year = max(year_cols.keys())
        max_col_idx = 0
        for c in year_cols[max_year]:
            max_col_idx = max_col_idx * 26 + (ord(c) - ord('A') + 1)

        for offset in range(1, 30):
            next_year = max_year + offset
            next_col_idx = max_col_idx + offset
            next_col = _col_letter(next_col_idx)
            pattern = rf'<c\s+r="{next_col}[5-8]"[^>]*><f>[A-Z]+\d+\+1</f>'
            if re.search(pattern, xml):
                year_cols[next_year] = next_col
            else:
                break

    return year_cols


# ══════════════════════════════════════════════════════════════════════════════
#  STYLE DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def _detect_row_style(xml: str, row: int, col: str = "D") -> Optional[str]:
    """Detect the style (s="...") of a cell in the given row/col."""
    m = re.search(rf'<c\s[^>]*r="{col}{row}"[^>]*?s="(\d+)"', xml)
    if m:
        return m.group(1)
    # Try reversed attr order
    m = re.search(rf'<c\s[^>]*s="(\d+)"[^>]*r="{col}{row}"', xml)
    return m.group(1) if m else None


def _detect_drug_styles(xml: str) -> Tuple[str, str, str, str]:
    """Detect styles for drug name and breakdown rows.

    Returns: (drug_d_style, drug_data_style, brkdn_d_style, brkdn_data_style)
    Fallback: R285 (Erbitux=drug header), R286 (CRC=breakdown).
    """
    drug_d = _detect_row_style(xml, 285, "D") or "399"
    drug_data = _detect_row_style(xml, 285, "F") or "48"
    brkdn_d = _detect_row_style(xml, 286, "D") or "400"
    brkdn_data = _detect_row_style(xml, 286, "F") or "395"
    return drug_d, drug_data, brkdn_d, brkdn_data


# ══════════════════════════════════════════════════════════════════════════════
#  ROW BUILDERS  (Oncology Drug Market format)
# ══════════════════════════════════════════════════════════════════════════════

def _build_drug_name_row(
    row: int, drug_name: str,
    revenues: Dict[int, float], year_cols: Dict[int, str],
    d_style: str = "399", data_style: str = "48",
) -> str:
    """Build drug name row: blue text, total Revenue values.

    D column = drug name (s=399 bold blue).
    Year columns = total revenue numeric values (s=48 blue font).
    """
    cells = [f'<row r="{row}">']
    cells.append(
        f'<c r="D{row}" s="{d_style}" t="inlineStr">'
        f'<is><t>{_xml_escape(drug_name)}</t></is></c>')
    for year, val in sorted(revenues.items()):
        if year in year_cols:
            col = year_cols[year]
            cells.append(f'<c r="{col}{row}" s="{data_style}"><v>{val}</v></c>')
    cells.append('</row>')
    return ''.join(cells)


def _build_breakdown_row(
    row: int, indication: str,
    drug_row: int, incidence_rows: List[int],
    this_incidence_row: int,
    year_cols: Dict[int, str],
    d_style: str = "400", data_style: str = "395",
    share: Optional[float] = None,
) -> str:
    """Build indication breakdown row: black text, formula referencing incidence.

    D column = indication abbreviation (s=400, for Pipeline SUMIF matching).
    Year columns = formula (s=395) based on one of three modes:
      1. share=1.0 (or single indication): direct reference to drug row
      2. share=<fraction>: DRUG_ROW * share (fixed percentage split)
      3. share=None + multiple incidence_rows: incidence-proportional split
    """
    cells = [f'<row r="{row}">']
    cells.append(
        f'<c r="D{row}" s="{d_style}" t="inlineStr">'
        f'<is><t>{_xml_escape(indication)}</t></is></c>')

    for year, col in sorted(year_cols.items()):
        if share is not None:
            if share >= 0.999:
                formula = f'{col}{drug_row}'
            else:
                formula = f'{col}{drug_row}*{share}'
        elif len(incidence_rows) == 1:
            formula = f'{col}{drug_row}'
        else:
            inc_refs = '+'.join(f'{col}{r}' for r in incidence_rows)
            formula = f'{col}{drug_row}*{col}{this_incidence_row}/({inc_refs})'
        cells.append(f'<c r="{col}{row}" s="{data_style}"><f>{formula}</f></c>')

    cells.append('</row>')
    return ''.join(cells)


def _build_incidence_row(
    row: int, indication: str, incidence: int,
    world_pop: int = 8000000000,
    ref_style: str = "48",
) -> str:
    """Build incidence parameter row."""
    cells = [f'<row r="{row}">']

    label = f"{indication} Incidence"
    cells.append(
        f'<c r="D{row}" s="{ref_style}" t="inlineStr">'
        f'<is><t>{_xml_escape(label)}</t></is></c>')

    # Incidence as fraction of world population
    pct = incidence / world_pop if world_pop > 0 else 0
    cells.append(f'<c r="F{row}" s="{ref_style}"><v>{pct}</v></c>')

    cells.append('</row>')
    return ''.join(cells)


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED FORMULA EXPLOSION
# ══════════════════════════════════════════════════════════════════════════════

def _adjust_formula(
    formula: str, master_col: int, master_row: int,
    slave_col: int, slave_row: int,
) -> str:
    """Adjust cell references in formula based on offset from master to slave.

    Respects $ anchoring: $A = column locked, $1 = row locked.
    """
    dcol = slave_col - master_col
    drow = slave_row - master_row

    def _adj(m):
        dc = m.group(1)   # '$' or ''
        col = m.group(2)
        dr = m.group(3)   # '$' or ''
        row = int(m.group(4))
        if not dc:
            col = _col_letter(_parse_col(col) + dcol)
        if not dr:
            row = row + drow
        return f'{dc}{col}{dr}{row}'

    return re.sub(r'(\$?)([A-Z]+)(\$?)(\d+)', _adj, formula)


def _explode_shared_formulas(xml: str) -> str:
    """Convert ALL shared formulas to explicit formulas before shifting.

    Shared masters:  <f t="shared" ref="G10:AH10" si="5">FORMULA</f>
    Shared slaves:   <f t="shared" si="5"/>

    After explosion every cell has an explicit <f>FORMULA</f> with no
    shared attributes.  This prevents any corruption when rows are
    inserted or shifted.
    """
    # ── Step 1: Collect masters  si → (col_idx, row, formula) ──
    # Search for <f> tags with t="shared" and ref=, then look backwards
    # for the enclosing <c r="COLrow"> to determine master cell position.
    masters: Dict[str, Tuple[int, int, str]] = {}
    for m in re.finditer(
            r'<f(?![a-z])\s([^>]*t="shared"[^>]*ref="[^"]*"[^>]*)(?<!/)>(.*?)</f>',
            xml, re.DOTALL):
        fa = m.group(1)
        si_m = re.search(r'si="(\d+)"', fa)
        if not si_m:
            continue
        # Find enclosing <c r="COLrow"> by searching backwards
        c_m = re.search(r'<c\s[^>]*r="([A-Z]+)(\d+)"',
                        xml[max(0, m.start()-200):m.start()])
        if c_m:
            masters[si_m.group(1)] = (
                _parse_col(c_m.group(1)), int(c_m.group(2)), m.group(2))

    if not masters:
        return xml
    log.info(f"  Exploding {len(masters)} shared-formula groups")

    # ── Step 2: Replace masters – strip shared attrs, keep formula ──
    def _fix_master(m):
        fa, body = m.group(1), m.group(2)
        if 't="shared"' not in fa or 'ref="' not in fa:
            return m.group(0)
        fa = re.sub(r'\s*t="shared"', '', fa)
        fa = re.sub(r'\s*ref="[^"]*"', '', fa)
        fa = re.sub(r'\s*si="\d+"', '', fa)
        return f'<f{fa}>{body}</f>'

    # (?<!/) prevents matching self-closing <f .../> as opening tag
    xml = re.sub(r'<f(?![a-z])([^>]*)(?<!/)>(.*?)</f>', _fix_master, xml, flags=re.DOTALL)

    # ── Step 3: Replace slaves – compute explicit formula ──
    #   Match full <c …>…</c> elements that contain a shared-formula slave.
    _SLAVE_RE = re.compile(
        r'(<c\s[^>]*r=")([A-Z]+)(\d+)("[^>]*>)(.*?)(</c>)', re.DOTALL)

    def _fix_slave(m):
        prefix, col_s, row_s, mid, inner, close = m.groups()
        f_m = re.search(r'<f\s([^>]*t="shared"[^>]*)/?>', inner)
        if not f_m:
            return m.group(0)
        fa = f_m.group(1)
        if 'ref="' in fa:          # already-processed master remnant
            return m.group(0)
        si = re.search(r'si="(\d+)"', fa)
        if not si or si.group(1) not in masters:
            return m.group(0)
        mc, mr, mf = masters[si.group(1)]
        formula = _adjust_formula(mf, mc, mr, _parse_col(col_s), int(row_s))
        # Replace self-closing <f .../> or empty <f ...></f> with explicit
        new_inner = re.sub(r'<f\s[^>]*t="shared"[^>]*/>', f'<f>{formula}</f>', inner)
        if new_inner == inner:
            new_inner = re.sub(r'<f\s[^>]*t="shared"[^>]*></f>', f'<f>{formula}</f>', inner)
        return f'{prefix}{col_s}{row_s}{mid}{new_inner}{close}'

    xml = _SLAVE_RE.sub(_fix_slave, xml)
    return xml


# ══════════════════════════════════════════════════════════════════════════════
#  ROW SORTING (ensure ascending order after insertion)
# ══════════════════════════════════════════════════════════════════════════════

def _sort_rows_in_sheetdata(xml: str) -> str:
    """Ensure <row> elements in <sheetData> are in ascending row-number order."""
    sd_start = xml.find('<sheetData')
    sd_end = xml.find('</sheetData>')
    if sd_start == -1 or sd_end == -1:
        return xml
    sd_open_end = xml.index('>', sd_start) + 1
    body = xml[sd_open_end:sd_end]

    # Match both <row ...>...</row> and self-closing <row .../>
    _ROW_RE = re.compile(
        r'<row\s[^>]*r="(\d+)"[^>]*(?:/>|>.*?</row>)', re.DOTALL)
    rows = [(int(m.group(1)), m.group(0)) for m in _ROW_RE.finditer(body)]

    if not rows:
        return xml

    rnums = [r[0] for r in rows]
    if rnums == sorted(rnums):
        return xml

    log.info("  Re-sorting rows in <sheetData>")
    rows.sort(key=lambda x: x[0])
    new_body = '\n'.join(r[1] for r in rows)
    return xml[:sd_open_end] + new_body + '\n' + xml[sd_end:]


# ══════════════════════════════════════════════════════════════════════════════
#  ROW SHIFTING
# ══════════════════════════════════════════════════════════════════════════════

def _shift_rows(xml: str, from_row: int, delta: int) -> str:
    """Shift all rows >= from_row by delta positions.

    Handles:
      - <row r="N"> attributes
      - <c r="XN"> cell address attributes
      - Formula cell references: F490, $F$490, $F490, F$490
    """
    # Shift <row r="N"> attributes
    xml = re.sub(
        r'(<row\s[^>]*r=")(\d+)(")',
        lambda m: f'{m.group(1)}{int(m.group(2)) + delta}{m.group(3)}'
        if int(m.group(2)) >= from_row else m.group(), xml)

    # Shift <c r="XN"> cell address attributes
    xml = re.sub(
        r'(r=")([A-Z]+)(\d+)(")',
        lambda m: f'{m.group(1)}{m.group(2)}{int(m.group(3)) + delta}{m.group(4)}'
        if int(m.group(3)) >= from_row else m.group(), xml)

    # ── Formula reference shifting (handles $) ──
    def _shift_ref(m):
        """Shift a single cell reference: $?COL$?ROW"""
        dc, col, dr, row_s = m.group(1), m.group(2), m.group(3), m.group(4)
        row = int(row_s)
        if row >= from_row:
            row += delta
        return f'{dc}{col}{dr}{row}'

    def _shift_formula(text: str) -> str:
        return re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', _shift_ref, text)

    # Shift formula text in <f ...>FORMULA</f>
    def shift_f_with_body(m):
        attrs, formula = m.group(1), m.group(2)
        if 'ref="' in attrs:
            attrs = re.sub(r'ref="([^"]*)"',
                           lambda rm: f'ref="{_shift_formula(rm.group(1))}"', attrs)
        return f'<f{attrs}>{_shift_formula(formula)}</f>'

    xml = re.sub(r'<f(?![a-z])([^>]*)(?<!/)>(.*?)</f>',
                 shift_f_with_body, xml, flags=re.DOTALL)

    # Shift ref="..." in self-closing shared slaves: <f ... />
    def shift_f_self_closing(m):
        attrs = m.group(1)
        if 'ref="' in attrs:
            attrs = re.sub(r'ref="([^"]*)"',
                           lambda rm: f'ref="{_shift_formula(rm.group(1))}"', attrs)
        return f'<f{attrs}/>'

    xml = re.sub(r'<f(?![a-z])([^>]*)/>', shift_f_self_closing, xml)

    return xml


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN EXPANSION  (Oncology Drug Market format)
# ══════════════════════════════════════════════════════════════════════════════

# Known indication → incidence parameter row mapping (original, unshifted)
_INDICATION_INCIDENCE_ROW: Dict[str, int] = {
    "OV": 462,
    "TNBC": 463,
    "BRCA": 464,
    "BLCA": 465,
    "GC": 466,
    "MCC": 467,
    "GBM": 468,
    "NSCLC": 469,
    "Melanoma": 470,
    "BTC": 471,
    "EC": 472,
    "ES-SCLC": 473,
    "HCC": 474,
    "Melanoma NCAM+": 475,
    "MPM": 476,
    "RCC": 477,
}


_DRUG_INSERT_AFTER = 341   # After Keytruda RCC (last Oncology Drug Market drug row)
_INC_INSERT_AFTER = 477    # After RCC Incidence (last incidence parameter row)


def expand_tam_sheet(
    xml: str,
    drugs_data: List[Dict],
    year_cols: Dict[int, str],
    dry_run: bool = False,
) -> Tuple[str, int, Dict[str, int]]:
    """Insert new drugs into TAM Solid Oncology Drug Market module.

    Drug rows inserted after R341 (Keytruda RCC), within the Oncology Drug
    Market section.  Incidence rows inserted after R477 (RCC Incidence),
    within the Parameters section.

    All existing rows below each insertion point shift down.  Formula
    references are updated accordingly (including $ anchored refs).

    Returns: (modified_xml, total_rows_added, row_shift_info)
        row_shift_info = {
            'n_drug_rows': int,
            'n_inc_rows': int,
            'drug_insert_at': int,   # first new drug row
            'inc_insert_at': int,    # first new incidence row (final numbering)
        }
    """
    if not drugs_data:
        return xml, 0, {}

    # Detect styles from existing rows
    drug_d, drug_data_s, brkdn_d, brkdn_data = _detect_drug_styles(xml)
    log.info(f"  Detected styles: drug D={drug_d} data={drug_data_s}, "
             f"breakdown D={brkdn_d} data={brkdn_data}")

    # ── Count new rows ──
    n_drug_rows = sum(1 + len(d.get('indications', [])) for d in drugs_data)

    new_incidence_list: List[Tuple[str, int]] = []
    seen_inc: set = set()
    for drug in drugs_data:
        for ind in drug.get("indications", []):
            name = ind["name"]
            if name in _INDICATION_INCIDENCE_ROW or name in seen_inc:
                continue
            inc_val = ind.get("incidence_global_annual", 0)
            if inc_val > 0:
                new_incidence_list.append((name, inc_val))
                seen_inc.add(name)
    n_inc_rows = len(new_incidence_list)
    total_new = n_drug_rows + n_inc_rows

    log.info(f"  Will insert {n_drug_rows} drug rows after R{_DRUG_INSERT_AFTER}, "
             f"{n_inc_rows} incidence rows after R{_INC_INSERT_AFTER}")

    if dry_run:
        return xml, total_new, {}

    # ── Phase 1: Explode shared formulas ──
    xml = _explode_shared_formulas(xml)

    # ── Phase 2: Shift rows (two passes) ──
    # Pass 1: make room for drug rows at R342
    drug_insert_at = _DRUG_INSERT_AFTER + 1  # R342
    log.info(f"  Shifting rows >= {drug_insert_at} by +{n_drug_rows}")
    xml = _shift_rows(xml, drug_insert_at, n_drug_rows)

    # Pass 2: make room for incidence rows
    # Original R478 is now R478+n_drug_rows after pass 1
    inc_insert_at = _INC_INSERT_AFTER + n_drug_rows + 1
    log.info(f"  Shifting rows >= {inc_insert_at} by +{n_inc_rows}")
    xml = _shift_rows(xml, inc_insert_at, n_inc_rows)

    # ── Phase 2.5: Extend SUMIF ranges within TAM Solid ──
    # After shifting, SUMIF formulas like $D$9:$D$341 still end at R341
    # (last original drug row, below insertion point so not shifted).
    # Extend them to include new drug rows: $D$9:$D${341+n_drug_rows}.
    last_new_drug_row = _DRUG_INSERT_AFTER + n_drug_rows  # 341+58=399
    n_extended = 0
    def _extend_sumif(m):
        nonlocal n_extended
        tag_attrs, formula = m.group(1), m.group(2)
        orig = formula
        formula = re.sub(
            rf'(\$?[A-Z]{{1,2}}\$9:\$?[A-Z]{{1,2}}\$){_DRUG_INSERT_AFTER}\b',
            rf'\g<1>{last_new_drug_row}',
            formula
        )
        if formula != orig:
            n_extended += 1
            return f'<f{tag_attrs}>{formula}</f>'
        return m.group(0)
    xml = re.sub(
        r'<f(?![a-z])([^>]*)(?<!/)>(.*?)</f>',
        _extend_sumif, xml, flags=re.DOTALL)
    log.info(f"  Extended {n_extended} SUMIF ranges: $9:${_DRUG_INSERT_AFTER} → $9:${last_new_drug_row}")

    # ── Phase 3: Build new drug rows (R342 to R342+n_drug_rows-1) ──
    new_rows_xml: List[str] = []
    cur_row = drug_insert_at

    for drug in drugs_data:
        drug_name = drug["drug_name"]
        revenues = drug.get("revenues", {})
        indications = drug.get("indications", [])

        if not indications:
            log.warning(f"  {drug_name}: no indications, skipping")
            continue

        drug_row = cur_row
        new_rows_xml.append(
            _build_drug_name_row(cur_row, drug_name, revenues, year_cols,
                                 drug_d, drug_data_s))
        cur_row += 1

        for ind in indications:
            ind_name = ind["name"]
            ind_share = ind.get("share")
            if ind_share is not None:
                new_rows_xml.append(
                    _build_breakdown_row(
                        cur_row, ind_name, drug_row, [], 0,
                        year_cols, brkdn_d, brkdn_data, share=ind_share))
            else:
                new_rows_xml.append(
                    _build_breakdown_row(
                        cur_row, ind_name, drug_row, [1],
                        1, year_cols, brkdn_d, brkdn_data, share=1.0))
            cur_row += 1

        n_ind = len(indications)
        log.info(f"  {drug_name}: R{drug_row}-R{cur_row-1} "
                 f"({1 + n_ind} rows, {n_ind} indications)")

    # ── Phase 4: Build new incidence rows ──
    new_inc_map: Dict[str, int] = {}
    inc_cur = inc_insert_at
    for name, inc_val in new_incidence_list:
        new_rows_xml.append(_build_incidence_row(inc_cur, name, inc_val))
        new_inc_map[name] = inc_cur
        log.info(f"  New incidence: {name} → R{inc_cur}")
        inc_cur += 1

    # ── Phase 5: Insert new rows and sort ──
    insert_pos = xml.find('</sheetData>')
    xml = xml[:insert_pos] + '\n'.join(new_rows_xml) + '\n' + xml[insert_pos:]
    xml = _sort_rows_in_sheetdata(xml)

    # Update dimension
    all_rows = [int(m) for m in re.findall(r'<row\s+r="(\d+)"', xml)]
    new_max = max(all_rows)
    xml = re.sub(
        r'(<dimension ref="[^"]*?)(\d+)(")',
        lambda m: f'{m.group(1)}{new_max}{m.group(3)}',
        xml)

    log.info(f"  Inserted {total_new} rows → max R{new_max}")

    shift_info = {
        'n_drug_rows': n_drug_rows,
        'n_inc_rows': n_inc_rows,
        'drug_insert_at': drug_insert_at,
        'inc_insert_at': inc_insert_at,
        'new_max': new_max,
    }
    return xml, total_new, shift_info


# ══════════════════════════════════════════════════════════════════════════════
#  LEGACY EXPANSION (for TAM Blood and generic use)
# ══════════════════════════════════════════════════════════════════════════════

def expand_tam_sheet_legacy(
    xml: str,
    indications_data: List[Dict],
    year_cols: Dict[int, str],
    insert_before_row: int,
    incidence_insert_row: int,
    dry_run: bool = False,
) -> Tuple[str, int]:
    """Add new indication data to a TAM sheet (legacy format).

    Returns: (modified_xml, total_rows_added)
    """
    total_new_rows = 0
    new_rows_xml = []

    for ind_data in indications_data:
        indication = ind_data["indication"]
        drugs = ind_data.get("drugs", [])
        if not drugs:
            log.warning(f"  {indication}: no drugs, skipping")
            continue

        log.info(f"  {indication}: {len(drugs)} drugs")

        # Build drug rows
        cur_row = insert_before_row + total_new_rows
        first_drug = cur_row + 1  # +1 for section header

        # Section header (empty row with indication name)
        new_rows_xml.append(
            f'<row r="{cur_row}"><c r="D{cur_row}" s="21" t="inlineStr">'
            f'<is><t>{_xml_escape(indication)}</t></is></c></row>')
        cur_row += 1

        for drug in drugs:
            name = drug.get("name", "Unknown")
            generic = drug.get("generic", "")
            manufacturer = drug.get("manufacturer", "")
            header = name
            if manufacturer or generic:
                parts = [p for p in (manufacturer, generic) if p]
                header += f" ({', '.join(parts)})"

            revenues = drug.get("revenues_mm_usd", {})

            new_rows_xml.append(
                _build_drug_header_row(cur_row, header))
            cur_row += 1

            new_rows_xml.append(
                _build_drug_row(cur_row, name, indication, revenues, year_cols))
            cur_row += 1

        # Total row
        new_rows_xml.append(
            _build_indication_total_row(
                cur_row, indication, first_drug, cur_row - 1, year_cols))
        cur_row += 1

        # Blank separator
        new_rows_xml.append(f'<row r="{cur_row}"/>')
        cur_row += 1

        rows_for_ind = cur_row - (insert_before_row + total_new_rows)
        total_new_rows += rows_for_ind
        log.info(f"    {indication}: {rows_for_ind} rows "
                 f"(R{insert_before_row + total_new_rows - rows_for_ind}"
                 f"-R{insert_before_row + total_new_rows - 1})")

    if total_new_rows == 0:
        return xml, 0

    if dry_run:
        log.info(f"  [DRY-RUN] Would insert {total_new_rows} rows "
                 f"before R{insert_before_row}")
        return xml, total_new_rows

    # Step 1: Shift existing rows from insert_before_row onward
    xml = _shift_rows(xml, insert_before_row, total_new_rows)

    # Step 2: Insert new rows
    insert_marker = f'<row r="{insert_before_row + total_new_rows}"'
    insert_pos = xml.find(insert_marker)
    if insert_pos == -1:
        insert_pos = xml.find('</sheetData>')

    new_xml = '\n'.join(new_rows_xml)
    xml = xml[:insert_pos] + new_xml + '\n' + xml[insert_pos:]

    # Step 3: Add incidence parameters
    incidence_rows_xml = []
    inc_first = incidence_insert_row + total_new_rows + 1
    inc_row = inc_first - 1
    for ind_data in indications_data:
        indication = ind_data["indication"]
        incidence = ind_data.get("incidence_global_annual", 0)
        if incidence > 0:
            inc_row += 1
            incidence_rows_xml.append(
                _build_incidence_row(inc_row, indication, incidence))

    if incidence_rows_xml:
        n_inc = len(incidence_rows_xml)
        xml = _shift_rows(xml, inc_first, n_inc)
        inc_marker = f'<row r="{inc_first + n_inc}"'
        inc_pos = xml.find(inc_marker)
        if inc_pos == -1:
            inc_pos = xml.find('</sheetData>')
        xml = xml[:inc_pos] + '\n'.join(incidence_rows_xml) + '\n' + xml[inc_pos:]
        total_new_rows += n_inc
        log.info(f"  Added {n_inc} incidence parameter rows")

    # Update dimension
    xml = re.sub(
        r'<dimension ref="([^"]*)"',
        lambda m: f'<dimension ref="{m.group(1).split(":")[0]}:'
                  f'{re.sub(r"[0-9]+", lambda n: str(int(n.group()) + total_new_rows), m.group(1).split(":")[1])}"',
        xml)

    return xml, total_new_rows


# Legacy row builders (kept for expand_tam_sheet_legacy)
def _build_drug_header_row(
    row: int, drug_name: str, ref_style: str = "21",
) -> str:
    escaped_name = _xml_escape(drug_name)
    return (f'<row r="{row}">'
            f'<c r="D{row}" s="{ref_style}" t="inlineStr">'
            f'<is><t>{escaped_name}</t></is></c>'
            f'</row>')


def _build_drug_row(
    row: int, drug_name: str, indication: str,
    revenues: Dict[str, float], year_cols: Dict[int, str],
    ref_style: str = "48",
) -> str:
    cells = [f'<row r="{row}">']
    escaped_ind = _xml_escape(indication)
    cells.append(
        f'<c r="D{row}" s="{ref_style}" t="inlineStr">'
        f'<is><t>{escaped_ind}</t></is></c>')
    for year_str, revenue in revenues.items():
        year = int(year_str)
        if year in year_cols:
            col = year_cols[year]
            cells.append(f'<c r="{col}{row}" s="{ref_style}"><v>{revenue}</v></c>')
    cells.append('</row>')
    return ''.join(cells)


def _build_indication_total_row(
    row: int, indication: str, first_drug_row: int, last_drug_row: int,
    year_cols: Dict[int, str], ref_style: str = "21",
) -> str:
    cells = [f'<row r="{row}">']
    escaped = _xml_escape(f"Total {indication}")
    cells.append(
        f'<c r="D{row}" s="{ref_style}" t="inlineStr">'
        f'<is><t>{escaped}</t></is></c>')
    for year, col in year_cols.items():
        formula = f'SUM({col}{first_drug_row}:{col}{last_drug_row})'
        cells.append(f'<c r="{col}{row}" s="{ref_style}"><f>{formula}</f></c>')
    cells.append('</row>')
    return ''.join(cells)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Add new drugs to TAM Solid Oncology Drug Market section"
    )
    parser.add_argument(
        "--json-dir", required=True,
        help="Directory containing drug JSON files")
    parser.add_argument(
        "--file", default=str(_DEFAULT_FILE),
        help=f"Excel file to modify (default: {_DEFAULT_FILE})")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview changes without writing")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    json_dir = Path(args.json_dir)

    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return
    if not json_dir.exists():
        log.error(f"JSON directory not found: {json_dir}")
        return

    # Load drug JSON files
    drugs_data = []
    for jf in sorted(json_dir.glob("tam_drug_*.json")):
        log.info(f"Loading {jf.name}")
        with open(jf) as f:
            data = json.load(f)
        # Convert JSON string year keys to integers
        if "revenues" in data:
            data["revenues"] = {int(k): v for k, v in data["revenues"].items()}
        drugs_data.append(data)

    if not drugs_data:
        log.error("No drug JSON files found (expected tam_drug_*.json)")
        return

    log.info(f"Drugs to add: {[d['drug_name'] for d in drugs_data]}")

    # Backup
    if not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_tam_expand_{ts}.xlsx")
        shutil.copy2(xlsx_path, backup)
        log.info(f"Backup: {backup}")

    # Find TAM Solid sheet
    solid_zip = None
    for solid_name in ["TAM Solid", "TAM Solid+MM"]:
        solid_zip = _get_sheet_zip_path(xlsx_path, solid_name)
        if solid_zip:
            break

    if not solid_zip:
        log.error("TAM Solid sheet not found")
        return

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(solid_zip).decode('utf-8')
    year_cols = detect_year_columns(xml)
    log.info(f"TAM Solid year columns: {dict(sorted(year_cols.items())[-3:])}...")

    # Expand (insert within Oncology Drug Market module)
    xml, n_added, shift_info = expand_tam_sheet(
        xml, drugs_data, year_cols,
        dry_run=args.dry_run)

    if n_added == 0:
        log.info("No changes to write")
        return

    if args.dry_run:
        log.info("Dry run complete — no changes written")
        return

    # Strip cached ERROR values and fix self-closing <v/> tags
    xml = _strip_formula_cache(xml)
    xml = xml.replace('<v/>', '<v></v>')
    modified = {solid_zip: xml.encode('utf-8')}

    # Surgical zip patch
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        ct = zf.read("[Content_Types].xml").decode("utf-8")
        wr = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    # Strip calcChain references
    ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', '', ct)
    wr = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', wr)
    modified["[Content_Types].xml"] = ct.encode("utf-8")
    modified["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")

    # Update Pipeline cross-sheet references to TAM Solid
    n_drug = shift_info.get('n_drug_rows', 0)
    n_inc = shift_info.get('n_inc_rows', 0)
    new_max = shift_info.get('new_max', 562)

    pipe_zip = _get_sheet_zip_path(xlsx_path, "Pipeline")
    if pipe_zip:
        with zipfile.ZipFile(xlsx_path) as zf:
            pipe_xml = zf.read(pipe_zip).decode("utf-8")

        # Build TAM Solid row shift map for Pipeline cross-refs
        drug_insert_at = _DRUG_INSERT_AFTER + 1  # R342
        inc_orig_boundary = _INC_INSERT_AFTER + 1  # R478

        def tam_shift(orig_row: int) -> int:
            if orig_row < drug_insert_at:
                return orig_row
            elif orig_row < inc_orig_boundary:
                return orig_row + n_drug
            else:
                return orig_row + n_drug + n_inc

        # Helper: shift a single cell ref (handles $)
        def _shift_one_ref(dc, col, dr, row_s):
            row = int(row_s)
            return f"{dc}{col}{dr}{tam_shift(row)}"

        # Match 'TAM Solid'!REF or 'TAM Solid'!REF:REF (range end also on TAM Solid)
        _TAM_REF = re.compile(
            r"('TAM Solid'!)"
            r"(\$?)([A-Z]{1,3})(\$?)(\d+)"       # first cell
            r"(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?"  # optional :second cell
        )

        def _shift_tam_ref(m):
            prefix = m.group(1)
            r1 = _shift_one_ref(m.group(2), m.group(3), m.group(4), m.group(5))
            result = f"{prefix}{r1}"
            if m.group(6) is not None:
                r2 = _shift_one_ref(m.group(6), m.group(7), m.group(8), m.group(9))
                result += f":{r2}"
            return result

        pipe_xml = _TAM_REF.sub(_shift_tam_ref, pipe_xml)

        # Extend SUMIF ranges: shifted end rows → new_max
        # After shifting, old R341→R341 (no shift), old R501→R562
        # SUMIF ranges need to cover R342-R399 (new drug rows),
        # so extend any range ending at R341 or R562 to new_max
        for shifted_end in [str(_DRUG_INSERT_AFTER), str(tam_shift(501))]:
            pipe_xml = re.sub(
                rf"('TAM Solid'!\$D\$9:\$D\$){shifted_end}\b",
                rf"\g<1>{new_max}", pipe_xml)
            pipe_xml = re.sub(
                rf"('TAM Solid'![A-Z]{{1,2}}\$9:[A-Z]{{1,2}}\$){shifted_end}\b",
                rf"\g<1>{new_max}", pipe_xml)

        pipe_xml = _strip_formula_cache(pipe_xml)
        pipe_xml = pipe_xml.replace('<v/>', '<v></v>')
        modified[pipe_zip] = pipe_xml.encode("utf-8")

        # Count updates
        n_tam_refs = len(re.findall(r"'TAM Solid'!", pipe_xml))
        log.info(f"  Pipeline: shifted TAM Solid refs, SUMIF ranges → R{new_max} "
                 f"({n_tam_refs} TAM refs total)")
    else:
        log.warning("  Pipeline sheet not found — refs NOT updated")

    tmp = xlsx_path.with_suffix(".~tam_expand.xlsx")
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                else:
                    data = zin.read(item.filename)
                    # Strip formula error cache + fix <v/> from all worksheets
                    if item.filename.startswith("xl/worksheets/sheet") \
                            and item.filename.endswith(".xml"):
                        s = data.decode("utf-8")
                        s = _strip_formula_cache(s)
                        s = s.replace('<v/>', '<v></v>')
                        data = s.encode("utf-8")
                    zout.writestr(item, data)

    try:
        tmp.replace(xlsx_path)
    except PermissionError:
        import os
        os.remove(str(xlsx_path))
        tmp.rename(xlsx_path)

    log.info(f"Saved → {xlsx_path}")

    # Summary
    print(f"\n{'='*60}")
    print("TAM Expansion Complete")
    print(f"{'='*60}")
    for d in drugs_data:
        name = d["drug_name"]
        n_ind = len(d.get("indications", []))
        rev = d.get("revenues", {})
        rev_str = ", ".join(f"{y}=${v:.0f}M" for y, v in sorted(rev.items()))
        print(f"  {name}: {n_ind} indications ({rev_str})")
    if shift_info:
        print(f"  Drug rows: {shift_info.get('n_drug_rows', 0)}, "
              f"Incidence rows: {shift_info.get('n_inc_rows', 0)}")
        print(f"  New max row: {shift_info.get('new_max', '?')}")
    print(f"  Total rows added: {n_added}")
    print(f"  File: {xlsx_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
