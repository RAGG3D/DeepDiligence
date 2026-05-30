#!/usr/bin/env python3
"""
fix_tam_layout.py -- Fix TAM Solid+MM sheet layout after expand_tam.py.

1. Fix incidence rows R509-R515 (broken 2-cell format → full format matching R500-R508)
2. Move matched drug indication rows from new sections into existing drug sections
3. Remove MPM section entirely (all drugs matched)
4. Renumber rows and remap all formula references

Uses surgical zip patching (NEVER openpyxl .save()).
"""

import argparse
import logging
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


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

# Style indices (from existing sheet)
S_ROW_DEFAULT = "235"   # Row-level default + D column (numFmt 205)
S_EMPTY_A = "326"       # Column A empty cell
S_EMPTY_B = "3"         # Column B empty cell
S_EMPTY_C_INC = "326"   # Column C empty (incidence rows)
S_EMPTY_C_SUB = "529"   # Column C empty (sub-rows)
S_HEADER_E = "52"       # E column header cell ([MM USD] or [%])
S_IND_NAME = "410"      # Indication name (D column, sub-row)
S_DATA = "405"          # Data cell (sub-row values)
S_INC_BLUE = "527"      # Incidence carry-forward (blue font)
S_INC_BLACK = "526"     # Incidence formula cell (black font, P column)

# Shared string indices
SS_MM_USD = "150"       # "[MM USD]"
SS_PCT = "259"          # "[%]"

# Column letters for A-AH
_COLS_F_O = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
_COLS_Q_AH = ["Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
              "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"]

# ══════════════════════════════════════════════════════════════════════════════
#  INCIDENCE DATA (cases in millions for P-column formula)
# ══════════════════════════════════════════════════════════════════════════════

# Each: (row_number, shared_string_index, cases_in_millions_str, f_value)
# f_value is the existing value from the broken F column
INCIDENCE_ROWS = [
    (509, "1101", "0.221",     "2.7625000000000001E-5"),
    (510, "1102", "0.417367",  "5.2170874999999997E-5"),
    (511, "1103", "0.232",     "2.9E-5"),
    (512, "1104", "0.938",     "1.1725E-4"),
    (513, "1105", "0.049",     "6.1249999999999998E-6"),
    (514, "1106", "0.035",     "4.3749999999999996E-6"),
    (515, "1107", "0.388",     "4.85E-5"),
]

# ══════════════════════════════════════════════════════════════════════════════
#  DRUG MATCHING TABLE
#  (data_row, header_row, target_drug_insert_after)
# ══════════════════════════════════════════════════════════════════════════════

# Each entry: (data_row_num, header_row_num, insert_after_row)
# header_row is the drug name row (s=21) immediately above the data row
DRUG_MOVES = [
    # BTC section drugs
    (367, 366, 182),   # Imfinzi BTC → after R182 (Imfinzi sub-row)
    (373, 372, 263),   # Keytruda BTC → after R263 (Keytruda last sub-row... wait)
    (375, 374, 251),   # Opdivo BTC → after R251 (Opdivo last sub-row)
    (377, 376, 169),   # Tafinlar+Mek BTC → after R169 (Tafinlar sub-row)
    (379, 378, 131),   # Krazati BTC → after R131 (Krazati sub-row)
    # EC section drugs
    (384, 383, 263),   # Keytruda EC → after Keytruda last sub-row
    # ES-SCLC section drugs
    (393, 392, 173),   # Tecentriq ES-SCLC → after R173 (Tecentriq sub-row)
    (395, 394, 182),   # Imfinzi ES-SCLC → after R182 (Imfinzi sub-row)
    # HCC section drugs
    (408, 407, 100),   # Stivarga HCC → after R100 (Stivarga sub-row)
    (412, 411, 162),   # Cyramza HCC → after R162 (Cyramza sub-row)
    (414, 413, 263),   # Keytruda HCC → after Keytruda last sub-row
    (416, 415, 251),   # Opdivo HCC → after Opdivo last sub-row
    (418, 417, 173),   # Tecentriq HCC → after Tecentriq sub-row
    (420, 419, 182),   # Imfinzi HCC → after Imfinzi sub-row
    (422, 421, 117),   # Imjudo HCC → after R117 (Imjudo sub-row)
    # Melanoma NCAM+ section drugs
    (427, 426, 263),   # Keytruda Mel → after Keytruda last sub-row
    (429, 428, 251),   # Opdivo Mel → after Opdivo last sub-row
    (431, 430, 157),   # Yervoy Mel → after R157 (Yervoy sub-row)
    (433, 432, 169),   # Tafinlar+Mek Mel → after Tafinlar sub-row
    (435, 434, 215),   # Zelboraf+Cot Mel → after R215 (Zelboraf sub-row)
    (437, 436, 190),   # Braftovi+Mek Mel → after R190 (Braftovi sub-row)
    (439, 438, 207),   # Imlygic Mel → after R207 (Imlygic, no existing sub-rows)
    # MPM section drugs
    (446, 445, 224),   # Alimta MPM → after R224 (Alimta sub-row)
    (448, 447, 251),   # Opdivo MPM → after Opdivo last sub-row
    (450, 449, 157),   # Yervoy MPM → after Yervoy sub-row
    (452, 451, 263),   # Keytruda MPM → after Keytruda last sub-row
    # RCC section drugs
    (457, 456, 263),   # Keytruda RCC → after Keytruda last sub-row
    (459, 458, 251),   # Opdivo RCC → after Opdivo last sub-row
    (461, 460, 157),   # Yervoy RCC → after Yervoy sub-row
    (463, 462, 28),    # Bavencio RCC → after R28 (Bavencio sub-row)
    (483, 482, 149),   # Avastin RCC → after R149 (Avastin sub-row)
]

# Rows to remove: all header+data pairs for matched drugs, plus MPM section
# MPM section rows: R444 (header), R445-R452 (4 drug pairs), R453 (total)
MPM_SECTION_ROWS = set(range(444, 454))  # R444-R453

# SUM formula total rows for each new section
SECTION_TOTALS = {
    380: (362, 379),   # Total BTC: SUM(F362:F379)
    389: (383, 388),   # Total EC: SUM(F383:F388)
    400: (392, 399),   # Total ES-SCLC: SUM(F392:F399)
    423: (403, 422),   # Total HCC: SUM(F403:F422)
    442: (426, 441),   # Total Melanoma NCAM+: SUM(F426:F441)
    453: (445, 452),   # Total MPM: SUM(F445:F452) -- will be removed
    486: (456, 485),   # Total RCC: SUM(F456:F485)
}


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


def _col_letter(col_idx: int) -> str:
    """1-based col index → letter(s)."""
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(ord('A') + col_idx % 26) + result
        col_idx //= 26
    return result


def _col_idx(col: str) -> int:
    """Column letter → 1-based index. A=1, B=2, ..., AH=34."""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx


# ══════════════════════════════════════════════════════════════════════════════
#  PART 1: BUILD FIXED INCIDENCE ROWS
# ══════════════════════════════════════════════════════════════════════════════

def _build_fixed_incidence_row(
    row: int, ss_idx: str, cases_mm: str, f_value: str, si_fo: int, si_qah: int,
) -> str:
    """Build a properly formatted incidence row matching R500-R508 pattern."""
    cells = []
    cells.append(f'<row r="{row}" spans="1:34" s="{S_ROW_DEFAULT}" '
                 f'customFormat="1" x14ac:dyDescent="0.25">')

    # A, B, C columns (empty styled)
    cells.append(f'<c r="A{row}" s="{S_EMPTY_A}"/>')
    cells.append(f'<c r="B{row}" s="{S_EMPTY_B}"/>')
    cells.append(f'<c r="C{row}" s="{S_EMPTY_C_INC}"/>')

    # D column: label (shared string)
    cells.append(f'<c r="D{row}" s="{S_ROW_DEFAULT}" t="s"><v>{ss_idx}</v></c>')

    # E column: [%]
    cells.append(f'<c r="E{row}" s="{S_HEADER_E}" t="s"><v>{SS_PCT}</v></c>')

    # F-O: shared formula carry-back (=G{row}, =H{row}, etc.)
    # F is the master with ref, G-O are shared references
    cells.append(f'<c r="F{row}" s="{S_INC_BLUE}">'
                 f'<f t="shared" ref="F{row}:O{row}" si="{si_fo}">G{row}</f>'
                 f'<v>{f_value}</v></c>')
    for col in _COLS_F_O[1:]:  # G through O
        cells.append(f'<c r="{col}{row}" s="{S_INC_BLUE}">'
                     f'<f t="shared" si="{si_fo}"/>'
                     f'<v>{f_value}</v></c>')

    # P column: source formula (cases_mm / P497)
    cells.append(f'<c r="P{row}" s="{S_INC_BLACK}">'
                 f'<f>{cases_mm}/P497</f>'
                 f'<v>{f_value}</v></c>')

    # Q-AH: shared formula carry-forward (=P{row})
    cells.append(f'<c r="Q{row}" s="{S_INC_BLUE}">'
                 f'<f t="shared" ref="Q{row}:AH{row}" si="{si_qah}">P{row}</f>'
                 f'<v>{f_value}</v></c>')
    for col in _COLS_Q_AH[1:]:  # R through AH
        cells.append(f'<c r="{col}{row}" s="{S_INC_BLUE}">'
                     f'<f t="shared" si="{si_qah}"/>'
                     f'<v>{f_value}</v></c>')

    cells.append('</row>')
    return ''.join(cells)


# ══════════════════════════════════════════════════════════════════════════════
#  PART 2: TRANSFORM DATA ROW TO SUB-ROW FORMAT
# ══════════════════════════════════════════════════════════════════════════════

def _transform_to_subrow(row_xml: str, new_row: int) -> str:
    """Transform a sparse s=48 data row into a full sub-row format.

    Input:  <row r="373" spans="4:34"><c r="D373" s="48" t="s"><v>1045</v></c>
            <c r="K373" s="48"><v>0</v></c>...</row>
    Output: Full sub-row with A/B/C/D/E cells, s=410 for D, s=405 for all data.
    """
    # Extract D-column shared string reference
    d_match = re.search(r'<c r="D\d+"[^>]*t="s"[^>]*><v>(\d+)</v>', row_xml)
    if not d_match:
        d_match = re.search(r'<c r="D\d+"[^>]*t="inlineStr"[^>]*><is><t>([^<]*)</t></is>', row_xml)
        if d_match:
            d_is_inline = True
            d_text = d_match.group(1)
        else:
            log.warning(f"Cannot find D-column in row: {row_xml[:100]}")
            return row_xml
    else:
        d_is_inline = False
        d_ss_idx = d_match.group(1)

    # Extract all data cell values (columns F through AH)
    data_values = {}  # col_letter → value_str
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*><v>([^<]*)</v></c>', row_xml):
        col = cm.group(1)
        if col != "D":
            data_values[col] = cm.group(2)

    # Build new row
    cells = []
    cells.append(f'<row r="{new_row}" spans="1:34" s="{S_ROW_DEFAULT}" '
                 f'customFormat="1" x14ac:dyDescent="0.25">')

    # A, B, C columns
    cells.append(f'<c r="A{new_row}" s="{S_EMPTY_A}"/>')
    cells.append(f'<c r="B{new_row}" s="{S_EMPTY_B}"/>')
    cells.append(f'<c r="C{new_row}" s="{S_EMPTY_C_SUB}"/>')

    # D column: indication name
    if d_is_inline:
        cells.append(f'<c r="D{new_row}" s="{S_IND_NAME}" t="inlineStr">'
                     f'<is><t>{d_text}</t></is></c>')
    else:
        cells.append(f'<c r="D{new_row}" s="{S_IND_NAME}" t="s">'
                     f'<v>{d_ss_idx}</v></c>')

    # E column: [MM USD]
    cells.append(f'<c r="E{new_row}" s="{S_HEADER_E}" t="s">'
                 f'<v>{SS_MM_USD}</v></c>')

    # Data columns F through AH (s=405)
    all_data_cols = (_COLS_F_O + ["P"] + _COLS_Q_AH)
    # Actually need F through AH in order
    for ci in range(6, 35):  # columns 6 (F) through 34 (AH)
        col = _col_letter(ci)
        val = data_values.get(col)
        if val is not None:
            cells.append(f'<c r="{col}{new_row}" s="{S_DATA}"><v>{val}</v></c>')
        else:
            cells.append(f'<c r="{col}{new_row}" s="{S_DATA}"/>')

    cells.append('</row>')
    return ''.join(cells)


# ══════════════════════════════════════════════════════════════════════════════
#  PART 3: FORMULA REMAPPING
# ══════════════════════════════════════════════════════════════════════════════

def _remap_formula(formula: str, row_map: Dict[int, int]) -> str:
    """Update cell references in a formula using old→new row mapping."""
    def replace_ref(m):
        prefix = m.group(1)  # optional $
        col = m.group(2)
        dollar = m.group(3)  # optional $ before row
        row = int(m.group(4))
        new_row = row_map.get(row, row)
        return f'{prefix}{col}{dollar}{new_row}'
    return re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', replace_ref, formula)


def _remap_shared_ref(ref_str: str, row_map: Dict[int, int]) -> str:
    """Remap a shared formula ref attribute like 'F360:S360'."""
    def replace_ref(m):
        col = m.group(1)
        row = int(m.group(2))
        return f'{col}{row_map.get(row, row)}'
    return re.sub(r'([A-Z]{1,3})(\d+)', replace_ref, ref_str)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def _parse_rows(xml: str) -> Tuple[str, List[Tuple[int, str]], str]:
    """Parse sheet XML into prefix, list of (row_num, row_xml), suffix.

    Returns (before_sheetData_content, rows, after_sheetData_content).
    """
    # Find <sheetData> content
    sd_open = xml.find('<sheetData')
    sd_open_end = xml.index('>', sd_open) + 1
    sd_close = xml.find('</sheetData>')

    prefix = xml[:sd_open_end]
    suffix = xml[sd_close:]
    body = xml[sd_open_end:sd_close]

    rows = []
    # Single regex matching both self-closing <row .../> and content <row ...>...</row>
    for m in re.finditer(
            r'<row\b[^>]*r="(\d+)"[^>]*(?:>.*?</row>|/>)', body):
        rows.append((int(m.group(1)), m.group(0)))

    # Sort by row number
    rows.sort(key=lambda x: x[0])
    return prefix, rows, suffix


def _get_row_d_text(row_xml: str, ss_list: Optional[List[str]] = None) -> str:
    """Extract D-column text from a row (for logging)."""
    m = re.search(r'<c r="D\d+"[^>]*t="s"[^>]*><v>(\d+)</v>', row_xml)
    if m and ss_list:
        idx = int(m.group(1))
        return ss_list[idx] if idx < len(ss_list) else f"SS[{idx}]"
    m = re.search(r'<c r="D\d+"[^>]*t="inlineStr"[^>]*><is><t>([^<]*)</t></is>', row_xml)
    if m:
        return m.group(1)
    return "?"


def process_sheet(xml: str) -> str:
    """Main processing: fix incidence, move drugs, renumber, remap."""

    prefix, rows, suffix = _parse_rows(xml)
    row_dict = {rn: rx for rn, rx in rows}

    # ── Step 1: Fix incidence rows R509-R515 ──────────────────────────────
    log.info("Step 1: Fixing incidence rows R509-R515")
    si_counter = 172  # Next available shared formula index (max was 171)
    for orig_row, ss_idx, cases_mm, f_value in INCIDENCE_ROWS:
        si_fo = si_counter
        si_qah = si_counter + 1
        si_counter += 2
        new_row_xml = _build_fixed_incidence_row(
            orig_row, ss_idx, cases_mm, f_value, si_fo, si_qah)
        row_dict[orig_row] = new_row_xml
        log.info(f"  Fixed R{orig_row} (si={si_fo},{si_qah})")

    # ── Step 2: Build move plan ───────────────────────────────────────────
    log.info("Step 2: Building move plan")

    # Collect rows to remove and insertions to make
    rows_to_remove: Set[int] = set()
    # {insert_after_row: [data_row_xml, ...]} — in order of insertion
    insertions: Dict[int, List[str]] = {}

    for data_row, header_row, insert_after in DRUG_MOVES:
        rows_to_remove.add(data_row)
        rows_to_remove.add(header_row)
        if data_row in row_dict:
            data_xml = row_dict[data_row]
            # Transform from sparse s=48 to full sub-row format
            transformed = _transform_to_subrow(data_xml, data_row)
            if insert_after not in insertions:
                insertions[insert_after] = []
            insertions[insert_after].append(transformed)
            log.info(f"  Move R{data_row} (hdr R{header_row}) → after R{insert_after}")
        else:
            log.warning(f"  Data row R{data_row} not found in sheet!")

    # Remove entire MPM section
    rows_to_remove |= MPM_SECTION_ROWS
    log.info(f"  Removing MPM section R444-R453 ({len(MPM_SECTION_ROWS)} rows)")

    # Also remove separator rows between sections if they exist
    for sep in [381, 390, 401, 424, 443, 454]:
        if sep in row_dict:
            # Check if it's an empty/separator row
            rx = row_dict[sep]
            if re.match(r'<row\b[^>]*/>', rx) or len(rx) < 80:
                rows_to_remove.add(sep)

    # Remove section headers for sections that lose ALL drugs
    # MPM header (R444) already in MPM_SECTION_ROWS

    # ── Step 3: Identify remaining data rows in each section ──────────────
    # After removing matched drugs, check if any section is now empty
    # For sections with remaining drugs, keep header + total + remaining drugs
    # MPM is fully removed. Others have unmatched drugs remaining.

    # ── Step 4: Reorder rows ──────────────────────────────────────────────
    log.info("Step 3: Reordering rows")

    # Build new row list: walk through original rows, skip removed,
    # insert after targets
    new_rows = []
    for rn, rx in sorted(row_dict.items()):
        if rn in rows_to_remove:
            continue
        new_rows.append((rn, rx))
        # Check for insertions after this row
        if rn in insertions:
            for data_xml in insertions[rn]:
                # Use a temporary row number (will be renumbered)
                new_rows.append((-1, data_xml))

    # ── Step 5: Renumber rows sequentially ────────────────────────────────
    log.info("Step 4: Renumbering rows")

    # We want to maintain the original row numbers where possible,
    # but inserted rows need new numbers. Use a compaction approach:
    # Walk through and assign sequential numbers, preserving gaps
    # for missing rows (empty Excel rows with no XML element).

    # Strategy: original rows keep their numbers if no shift is needed.
    # But insertions change the count. We need to renumber everything.

    # First pass: determine what the final row numbers should be.
    # Approach: maintain relative order. Rows 1-360 (existing drug sections
    # with insertions) will expand. Rows 361+ (new sections, minus removed)
    # will compact upward.

    # Simpler approach: just assign new row numbers sequentially from 1,
    # but preserve gaps for "missing" rows (rows that had no XML element
    # in the original, like blank rows between sections).

    # Actually the simplest correct approach: assign numbers based on
    # the original row's position, inserting new rows right after their
    # target. We need a mapping from old → new.

    # Let's just renumber everything sequentially. The sheet already has
    # gaps (missing rows 390, 401, etc.) which we don't need to preserve
    # since they're just blank.

    # Build the final ordered list with new sequential row numbers.
    # Start from the first row number in the sheet.
    first_row = new_rows[0][0] if new_rows[0][0] > 0 else 1
    row_map: Dict[int, int] = {}  # old_row → new_row

    # Assign new row numbers
    cur_new_row = first_row
    final_rows: List[Tuple[int, int, str]] = []  # (old_row, new_row, xml)

    for old_rn, rx in new_rows:
        new_rn = cur_new_row
        if old_rn > 0:
            row_map[old_rn] = new_rn
        final_rows.append((old_rn, new_rn, rx))
        cur_new_row += 1

    log.info(f"  Row count: {len(rows)} → {len(final_rows)} "
             f"(removed {len(rows) - len(final_rows) + sum(1 for _, rx in new_rows if _ == -1)})")
    log.info(f"  Last row: {cur_new_row - 1} (was {rows[-1][0]})")

    # ── Step 6: Apply renumbering to each row's XML ──────────────────────
    log.info("Step 5: Applying row renumbering and formula remapping")

    renumbered_rows = []
    for old_rn, new_rn, rx in final_rows:
        rx = _renumber_row(rx, new_rn, row_map)
        renumbered_rows.append(rx)

    # ── Step 7: Fix SUM formulas for section totals ──────────────────────
    # SUM range endpoints may reference removed rows (not in row_map),
    # causing incorrect fallback to original row numbers. Fix explicitly.
    log.info("Step 6: Fixing section total SUM formulas")

    for total_row, (orig_first, orig_last) in SECTION_TOTALS.items():
        if total_row in rows_to_remove:
            continue  # MPM total removed
        if total_row not in row_map:
            continue
        # Find first and last surviving rows in the original SUM range
        surviving = [r for r in range(orig_first, orig_last + 1)
                     if r in row_map]
        if not surviving:
            log.warning(f"  No surviving rows for total R{total_row}")
            continue
        new_first = row_map[surviving[0]]
        new_last = row_map[surviving[-1]]
        new_total = row_map[total_row]

        # Find this row in renumbered_rows and fix SUM formulas
        for i, rx in enumerate(renumbered_rows):
            row_m = re.search(r'<row\b[^>]*r="(\d+)"', rx)
            if row_m and int(row_m.group(1)) == new_total:
                # Replace all SUM(COL_old:COL_old) with correct range
                rx = re.sub(
                    r'SUM\(([A-Z]{1,3})\d+:([A-Z]{1,3})\d+\)',
                    lambda m: f'SUM({m.group(1)}{new_first}:{m.group(2)}{new_last})',
                    rx)
                renumbered_rows[i] = rx
                log.info(f"  R{total_row}→R{new_total}: "
                         f"SUM range → {new_first}:{new_last}")
                break

    # ── Step 8: Rebuild XML ──────────────────────────────────────────────
    log.info("Step 6: Rebuilding XML")

    # Update dimension
    last_row = cur_new_row - 1
    prefix = re.sub(
        r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"',
        lambda m: f'<dimension ref="{m.group(1)}:{m.group(2)}{last_row}"',
        prefix)

    body = '\n'.join(renumbered_rows)
    return prefix + body + suffix


def _renumber_row(rx: str, new_row: int, row_map: Dict[int, int]) -> str:
    """Renumber a single row's XML: row r=, cell r=, formulas, shared refs."""

    # Update row r= attribute
    rx = re.sub(r'(<row\b[^>]*)\br="(\d+)"', f'\\1r="{new_row}"', rx)

    # Update cell references (r="D373" → r="D{new_row}")
    rx = re.sub(r'r="([A-Z]+)\d+"',
                lambda m: f'r="{m.group(1)}{new_row}"', rx)

    # Update formulas with row_map
    def remap_formula_tag(m):
        formula = m.group(1)
        remapped = _remap_formula(formula, row_map)
        return f'<f>{remapped}</f>'

    # Handle formulas WITH attributes (shared, ref, etc.)
    def remap_full_formula_tag(m):
        attrs = m.group(1)
        formula = m.group(2)

        # Remap ref= attribute if present
        ref_m = re.search(r'ref="([^"]*)"', attrs)
        if ref_m:
            new_ref = _remap_shared_ref(ref_m.group(1), row_map)
            attrs = attrs[:ref_m.start(1)] + new_ref + attrs[ref_m.end(1):]

        if formula:
            remapped = _remap_formula(formula, row_map)
            return f'<f {attrs}>{remapped}</f>'
        return f'<f {attrs}/>'

    # First handle <f ...>formula</f> (with attributes)
    rx = re.sub(r'<f\s+([^>]*)>([^<]*)</f>', remap_full_formula_tag, rx)
    # Then handle <f .../>  (self-closing with attributes, no formula body)
    rx = re.sub(r'<f\s+([^/>]*)/>', lambda m: f'<f {m.group(1)}/>', rx)
    # Then handle plain <f>formula</f> (no attributes)
    rx = re.sub(r'<f>([^<]+)</f>', remap_formula_tag, rx)

    return rx


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Fix TAM Solid+MM layout after expand_tam.py")
    parser.add_argument(
        "--file", default=str(_DEFAULT_FILE),
        help=f"Excel file to modify (default: {_DEFAULT_FILE})")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview changes without writing")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return

    # Find TAM Solid sheet
    sheet_zip = None
    for name in ["TAM Solid", "TAM Solid+MM"]:
        sheet_zip = _get_sheet_zip_path(xlsx_path, name)
        if sheet_zip:
            log.info(f"Found sheet '{name}' → {sheet_zip}")
            break
    if not sheet_zip:
        log.error("TAM Solid sheet not found")
        return

    # Read sheet XML
    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")

    log.info(f"Sheet XML: {len(xml):,} bytes")

    # Process
    new_xml = process_sheet(xml)

    if args.dry_run:
        log.info("[DRY-RUN] Would write modified sheet")
        # Print some stats
        old_rows = len(re.findall(r'<row\b', xml))
        new_rows = len(re.findall(r'<row\b', new_xml))
        log.info(f"  Rows: {old_rows} → {new_rows}")
        return

    # Backup
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_tam_fix_{ts}.xlsx")
    shutil.copy2(xlsx_path, backup)
    log.info(f"Backup: {backup}")

    # Surgical zip patch
    modified = {sheet_zip: new_xml.encode("utf-8")}

    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~tam_fix.xlsx")
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

    log.info(f"Saved → {xlsx_path}")

    # Summary
    old_rows = len(re.findall(r'<row\b', xml))
    new_rows_count = len(re.findall(r'<row\b', new_xml))
    print(f"\n{'='*60}")
    print("TAM Layout Fix Complete")
    print(f"{'='*60}")
    print(f"  Rows: {old_rows} → {new_rows_count}")
    print(f"  Drug rows moved: {len(DRUG_MOVES)}")
    print(f"  MPM section removed: {len(MPM_SECTION_ROWS)} rows")
    print(f"  Incidence rows fixed: {len(INCIDENCE_ROWS)}")
    print(f"  File: {xlsx_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
