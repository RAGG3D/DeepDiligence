#!/usr/bin/env python3
"""
generate_pipeline.py -- Fill Pipeline sheet Revenue Forecasting from Gemini research reports.

Parses Gemini reports for pipeline assets (drug names, targets, indications,
market shares, pricing) and generates Revenue Forecasting rows in the Pipeline sheet.

Uses surgical zip patching (NEVER openpyxl .save()).

Usage:
    python generate_pipeline.py --ticker CMPX --company-name "Compass Therapeutics" \
        [--report-dir path] [--pricing-dir path] [--dry-run]
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

# Reuse parsing from generate_scenarios.py
from generate_scenarios import (
    parse_gemini_reports, PipelineAsset,
    _asset_full_name, _xml_escape,
)

_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SHEET_NAME = "Pipeline"

# Revenue Forecasting section layout (actual Pipeline sheet layout)
SECTION_HDR = 5         # Headers area
YEAR_HDR = 7            # Year headers (='TAM Solid'!S6 etc.)
REV_SUM = 8             # "Operating Revenue From Sales" sum row
FIRST_DRUG = 9          # First drug block starts here

# Year ↔ column mapping: F=2010, G=2011, ..., S=2023, T=2024, ..., AH=2038
_YEAR_BASE = 2010
_COL_BASE = 6           # Column F = index 6

# Scenarios sheet year mapping: E=2019, F=2020, ..., X=2038
_SCENARIOS_YEAR_BASE = 2019
_SCENARIOS_COL_BASE = 5  # Column E = index 5

# TAM Solid Parameters: maturity curve rows (Average Maturity = R457)
MATURITY_ROW = {'AVG': 457, 'BIC': 457, 'T1': 457}

# TAM Solid Parameters: growth factor rows (unchanged by append-only expansion)
GROWTH_ROW = {'AVG': 551, 'BIC': 552, 'T1': 553}

# TAM Solid Parameters: COGS/Price row (unchanged by append-only expansion)
COGS_PRICE_ROW = 562

# Indications in Pipeline Referred Tables (rows 9-342) → Pipeline SUMIF
# Updated: BTC/EC/ES-SCLC/HCC/Melanoma NCAM+/MPM/RCC/HL/MM moved from
# _TAM_CROSS_REF to here after fill_tam_forecast.py added their data rows.
_PIPELINE_INDICATIONS = {
    "ESCA", "GC", "OV", "TNBC", "BRCA", "BLCA", "CRC", "NSCLC",
    "HNSCC", "Melanoma", "AML", "GBM", "PRAD",
    "BTC", "EC", "ES-SCLC", "HCC", "Melanoma NCAM+", "MPM", "RCC",
    "HL", "MM",
}

# Cross-sheet SUMIF from TAM sheets (empty — all moved to Referred Tables)
_TAM_CROSS_REF: dict = {}

_TAM_BLOOD_COL_OFFSET = 1  # TAM Blood column = Pipeline column + 1 (for same year)

# Aliases: Scenarios indication name → Peer Views section abbreviation
# Handles cases where Scenarios uses different names than Peer Views sections.
_INDICATION_ALIASES = {
    "Metastatic Melanoma": "Melanoma NCAM+",
    "SCLC": "ES-SCLC",
    "cHL": "HL",
    "GC/GEJ": "GC",
}

# Style IDs (from existing template analysis)
S = {
    'drug_a':     '79',    # A col: "X" marker for drug/TAM/price rows
    'drug_d':     '73',    # D col: drug name (inlineStr)
    'drug_stage': '332',   # S-AH: SUMIFS stage formula
    'drug_hist':  '331',   # F-R: empty stage cells
    'drug_e':     '330',   # E col: empty for drug row
    'tam_c':      '333',   # C col: indication label (blue text, same as tam_d)
    'tam_d':      '333',   # D col: TAM formula
    'tam_e':      '316',   # E col: "[Patients]" or "[MM USD]"
    'tam_data':   '334',   # Data cols: TAM values/formulas
    'ms_a':       '2',     # A col: blank for MS rows
    'ms_c':       '90',    # C col: rating label (same as ms_d)
    'ms_d':       '90',    # D col: MS formula
    'ms_data':    '335',   # Data cols: SUMIFS for MS
    'price_d':    '333',   # D col: price formula
    'price_data': '193',   # Data cols: price values
    'rev_d':      '54',    # D col: revenue formula
    'rev_e':      '325',   # E col: "[MM USD]" for revenue
    'rev_data':   '336',   # Data cols: revenue formula
    'cogs_a':     '337',   # A col for COGS row
    'cogs_d':     '54',    # D col: COGS formula
    'cogs_data':  '336',   # Data cols: COGS formula
    'sum_d':      '70',    # D col: sum header
    'sum_data':   '329',   # Data cols: SUM formula
    'sep_b':      '143',   # B col: separator formula
}


def _normalize_formula_xml(xml: str) -> str:
    """Sanitize formula XML to prevent Excel 'Removed Records: Formula' errors.

    Fixes:
    1. Double-escaped entities (&amp;gt; → &gt;, &amp;lt; → &lt;)
       but preserves bare &amp; (Excel concatenation operator).
    2. &quot; inside <f> text → literal " (Excel doesn't decode &quot;).
    3. <v/> self-closing → <v></v>.
    4. Backslash-escaped quotes (\\' → ') from regex replacements.
    5. Strips cached error values (t="e") from formula cells.
    """
    def _fix_formula(m):
        prefix, formula, suffix = m.group(1), m.group(2), m.group(3)
        # Fix double-escaped (but NOT bare &amp; which is valid)
        formula = re.sub(r'&amp;(gt|lt|quot);', r'&\1;', formula)
        # Fix &quot; → literal "
        formula = formula.replace('&quot;', '"')
        # Fix backslash-escaped quotes
        formula = formula.replace("\\'", "'")
        return prefix + formula + suffix

    xml = re.sub(
        r'(<f(?![a-z])[^>]*(?<!/)>)(.*?)(</f>)',
        _fix_formula, xml, flags=re.DOTALL
    )
    # Fix <v/> self-closing
    xml = xml.replace('<v/>', '<v></v>')
    # Strip cached error values: t="e" cells with formulas
    xml = re.sub(r'(<c [^>]*) t="e"', r'\1', xml)
    return xml


def _col_letter(col_idx: int) -> str:
    """1-based col index → letter(s). 1=A, 26=Z, 27=AA."""
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(ord('A') + col_idx % 26) + result
        col_idx //= 26
    return result


def _year_to_col(year: int) -> str:
    """Year → Pipeline column letter."""
    return _col_letter(_COL_BASE + (year - _YEAR_BASE))


def _scenarios_col_for_year(year: int) -> str:
    """Year → Scenarios sheet column letter.

    Pipeline col S = 2023, but Scenarios col S = 2033.
    This function returns the correct Scenarios column for a given year.
    """
    return _col_letter(_SCENARIOS_COL_BASE + (year - _SCENARIOS_YEAR_BASE))


# ═══════════════════════════════════════════════════════════════════════════════
#  READ DRUG NAMES FROM SCENARIOS SHEET (authoritative source)
# ═══════════════════════════════════════════════════════════════════════════════

def _read_scenarios_drug_info(xlsx_path: Path) -> Dict[str, Tuple[str, List[str]]]:
    """Read exact drug names and indications from the Scenarios sheet.

    Pipeline SUMIFS match on drug name strings against Scenarios!$C:$C.
    Names MUST be identical. This function reads the authoritative names
    directly from Scenarios instead of re-parsing Gemini reports.

    Handles BOTH t="inlineStr" (freshly written) and t="s" (after Excel re-save)
    cell types. Excel converts inlineStr → shared string on save.

    Returns: {drug_prefix: (full_name, [indication1, indication2, ...])}
    e.g. {"CTX-009": ("CTX-009 (DLL4 and VEGF-A, BTC/CRC)", ["BTC", "CRC"])}
    """
    sheet_zip = _get_sheet_zip_path(xlsx_path, "Scenarios")
    if not sheet_zip:
        log.warning("Cannot find Scenarios sheet — will use parsed names")
        return {}

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode('utf-8')
        # Load shared strings (needed after Excel re-save converts inlineStr → t="s")
        ss_list: List[str] = []
        if "xl/sharedStrings.xml" in [i.filename for i in zf.infolist()]:
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

    # Find asset name cells in column C — handle both inlineStr and shared string
    asset_cells: List[Tuple[str, str]] = []  # [(row_str, text), ...]

    # 1) inlineStr cells (freshly written, before Excel re-save)
    for m in re.finditer(
            r'<c\s+r="C(\d+)"[^>]*t="inlineStr"[^>]*>.*?<is><t>([^<]*)</t></is>.*?</c>',
            xml, re.DOTALL):
        asset_cells.append((m.group(1), m.group(2)))

    # 2) shared string cells (after Excel re-save)
    for m in re.finditer(
            r'<c\s+r="C(\d+)"[^>]*t="s"[^>]*>.*?<v>(\d+)</v>.*?</c>',
            xml, re.DOTALL):
        idx = int(m.group(2))
        if idx < len(ss_list):
            asset_cells.append((m.group(1), ss_list[idx]))

    # Find formula cells in column C (market share rows: C10&" BTC Market Share")
    formula_cells = re.findall(
        r'<c\s+r="C(\d+)"[^>]*>.*?<f>(C\d+&amp;" [^<]*)</f>.*?</c>',
        xml, re.DOTALL
    )

    assets_by_row: Dict[str, Tuple[str, str]] = {}  # {row_str: (prefix, full_name)}
    result: Dict[str, Tuple[str, List[str]]] = {}

    for row_str, val in asset_cells:
        row = int(row_str)
        if row > 50:
            continue
        if '(' not in val:
            continue  # Skip non-asset rows like "Base", "Bull", "Bear"

        prefix = val.split('(')[0].strip()
        assets_by_row[row_str] = (prefix, val)
        result[prefix] = (val, [])
        log.info(f"  Scenarios asset C{row_str}: {val}")

    # Parse market share rows to extract indications
    for row_str, formula in formula_cells:
        row = int(row_str)
        if row > 50:
            continue
        # Formula: C10&amp;" BTC Market Share"
        m = re.match(r'C(\d+)&amp;" (.+) Market Share"', formula)
        if m:
            ref_row = m.group(1)
            indication = m.group(2)
            if ref_row in assets_by_row:
                prefix = assets_by_row[ref_row][0]
                if prefix in result:
                    result[prefix][1].append(indication)

    for prefix, (full_name, indications) in result.items():
        log.info(f"  {prefix}: {len(indications)} indications → {indications}")

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  PRICING PARSER
# ═══════════════════════════════════════════════════════════════════════════════

def parse_pricing_reports(report_dir: Path, ticker: str) -> Dict[str, Dict[str, float]]:
    """Parse pricing chapter outputs.

    Returns: {drug_name: {indication: price_per_patient_mm_usd}}
    """
    result: Dict[str, Dict[str, float]] = {}
    pattern = f"{ticker}_*_pricing_*.md"
    files = sorted(report_dir.glob(pattern))
    if not files:
        log.info(f"No pricing reports found matching {pattern}")
        return result

    for f in files:
        log.info(f"  Parsing pricing: {f.name}")
        content = f.read_text(encoding="utf-8")

        # Extract drug name from filename: {TICKER}_{DRUG}_pricing_*.md
        parts = f.stem.split('_')
        drug_parts = []
        for p in parts[1:]:
            if p.lower() == 'pricing':
                break
            drug_parts.append(p)
        drug_name = '-'.join(drug_parts) if drug_parts else "Unknown"

        # Parse per-indication pricing tables
        drug_prices: Dict[str, float] = {}

        # Find indication headers: ### Indication Name
        sections = re.split(r'(?=^###\s+(?!Comparable))', content, flags=re.MULTILINE)
        for section in sections:
            header_m = re.match(r'###\s+(.+?)(?:\n|$)', section)
            if not header_m:
                continue
            ind_name = header_m.group(1).strip()
            if ind_name.lower().startswith("comparable"):
                continue

            # Find "Total Treatment Cost Per Patient" in table
            # Handles: "| $0.204 |", "| 0.204 |", "| $0.275 - $0.330 |" (takes midpoint of range)
            price_m = re.search(
                r'Total Treatment Cost Per Patient\s*\(MM USD\)\s*\|\s*\$?([\d.]+)(?:\s*[-–]\s*\$?([\d.]+))?',
                section, re.IGNORECASE
            )
            if price_m:
                lo = float(price_m.group(1))
                hi = float(price_m.group(2)) if price_m.group(2) else lo
                price = (lo + hi) / 2.0
                drug_prices[ind_name] = price
                log.info(f"    {drug_name} / {ind_name}: ${price:.3f} MM")

        if drug_prices:
            result[drug_name] = drug_prices

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════════

def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
    """Find the zip path for a named sheet."""
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_path: Dict[str, str] = {}
    for rel in rels_xml:
        if "worksheet" in rel.get("Type", ""):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rid_to_path[rid] = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")

    for sheet in wb_xml.findall(f".//{{{_NS_MAIN}}}sheet"):
        name = sheet.get("name", "")
        rid = sheet.get(f"{{{_NS_R}}}id", "")
        if name == sheet_name and rid in rid_to_path:
            return rid_to_path[rid]
    return None


def _read_peer_views_ratings(xlsx_path: Path) -> Dict[str, str]:
    """Read per-indication ratings for column-E drugs from Peer Views fill colors.

    Peer Views sections have fill-color-encoded ratings on the drug name cells:
      theme 9 (green)  = BIC (Best-In-Class)  → maturity row 445
      theme 8 (blue)   = T1  (Tier One)       → maturity row 446
      theme 7 (olive)  = AVG (Average)         → maturity row 444

    Approach:
    1. Find section headers (D column with indication text, A column="X")
    2. For each section, find the column-E drug cell (3-5 rows below header)
    3. Read its fill color → rating

    Returns: {indication_abbrev: "BIC"|"T1"|"AVG"}
    Example: {"BTC": "T1", "NSCLC": "AVG", "HL": "T1", ...}
    """
    pv_path = _get_sheet_zip_path(xlsx_path, "Peer Views")
    if not pv_path:
        log.warning("Peer Views sheet not found — using AVG for all")
        return {}

    with zipfile.ZipFile(xlsx_path) as zf:
        styles_root = ET.fromstring(zf.read("xl/styles.xml"))
        pv_xml = zf.read(pv_path).decode("utf-8")
        # Read shared strings (Peer Views uses t="s" cells)
        ss_list: List[str] = []
        if "xl/sharedStrings.xml" in [i.filename for i in zf.infolist()]:
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{_NS_MAIN}}}si"):
                t = si.find(f"{{{_NS_MAIN}}}t")
                if t is not None and t.text:
                    ss_list.append(t.text)
                else:
                    parts = [r.find(f"{{{_NS_MAIN}}}t")
                             for r in si.findall(f"{{{_NS_MAIN}}}r")]
                    ss_list.append("".join(p.text for p in parts if p is not None and p.text))

    ns = _NS_MAIN

    # ── Build style_idx → rating mapping via fills ──
    fills_el = styles_root.find(f"{{{ns}}}fills")
    fill_theme: Dict[int, int] = {}
    for i, fill in enumerate(fills_el):
        pf = fill.find(f"{{{ns}}}patternFill")
        if pf is not None:
            fg = pf.find(f"{{{ns}}}fgColor")
            if fg is not None and fg.get("theme"):
                fill_theme[i] = int(fg.get("theme"))

    cell_xfs = styles_root.find(f"{{{ns}}}cellXfs")
    xf_fill: Dict[int, int] = {}
    for i, xf in enumerate(cell_xfs):
        fid = xf.get("fillId")
        if fid:
            xf_fill[i] = int(fid)

    THEME_RATING = {9: "BIC", 8: "T1", 7: "AVG"}
    style_rating: Dict[int, str] = {}
    for s_idx, fill_id in xf_fill.items():
        theme = fill_theme.get(fill_id)
        if theme in THEME_RATING:
            style_rating[s_idx] = THEME_RATING[theme]

    # ── Build (row, col) → (text, style) map for ALL cells ──
    # (?<!/) negative lookbehind ensures we don't match self-closing <c ... />
    # as open tags (their /> would be consumed as attributes otherwise).
    cell_map: Dict[Tuple[int, str], Tuple[str, int]] = {}
    for m in re.finditer(r'<c\s+([^>]*?)(?<!/)>(.*?)</c>', pv_xml, re.DOTALL):
        attrs, body = m.group(1), m.group(2)
        r_m = re.search(r'r="([A-Z]+)(\d+)"', attrs)
        if not r_m:
            continue
        col, row = r_m.group(1), int(r_m.group(2))
        s_m = re.search(r's="(\d+)"', attrs)
        style = int(s_m.group(1)) if s_m else -1

        text = ""
        if 't="s"' in attrs:
            v_m = re.search(r'<v>(\d+)</v>', body)
            if v_m and int(v_m.group(1)) < len(ss_list):
                text = ss_list[int(v_m.group(1))]
        elif 't="inlineStr"' in attrs:
            t_m = re.search(r'<is><t>([^<]*)</t></is>', body)
            if t_m:
                text = t_m.group(1)
        elif 't="str"' in attrs:
            v_m = re.search(r'<v>([^<]*)</v>', body)
            if v_m:
                text = v_m.group(1)

        if text:
            cell_map[(row, col)] = (text, style)

    # ── Indication keyword → abbreviation (longer/more specific first) ──
    _IND_KW = [
        ("NSCLC", "NSCLC"), ("Non-Small Cell", "NSCLC"),
        ("Triple-Negative", "TNBC"), ("TNBC", "TNBC"),
        ("Biliary", "BTC"), ("BTC", "BTC"),
        ("Colorectal", "CRC"), ("CRC", "CRC"),
        ("Renal", "RCC"), ("RCC", "RCC"),
        ("Hepatocellular", "HCC"), ("HCC", "HCC"),
        ("Endometrial", "EC"),
        ("Melanoma", "Melanoma NCAM+"),
        ("Mesothelioma", "MPM"), ("MPM", "MPM"),
        ("Hodgkin", "HL"), ("cHL", "HL"),
        ("SCLC", "ES-SCLC"), ("Small Cell", "ES-SCLC"),
        ("Gastric", "GC"), ("Gastroesophageal", "GC"),
        ("Multiple Myeloma", "MM"),
    ]

    # ── Find section headers and extract column-E drug ratings ──
    # Sections: A=X + D has indication text → then E drug cell is 3-5 rows below
    result: Dict[str, str] = {}

    # Gather all section header rows (rows 200-486 where A="X" and D has text)
    section_rows: List[Tuple[int, str]] = []  # (row, indication_abbrev)
    for (row, col), (text, _) in cell_map.items():
        if col != "A" or text != "X" or row < 200:
            continue
        # Check D column for indication
        d_text = cell_map.get((row, "D"), ("", -1))[0]
        if not d_text:
            continue
        # Match indication keyword
        indication = None
        for keyword, ind_abbrev in _IND_KW:
            if keyword in d_text:
                if ind_abbrev == "ES-SCLC" and "NSCLC" in d_text:
                    continue
                indication = ind_abbrev
                break
        if indication:
            section_rows.append((row, indication))

    section_rows.sort()
    log.info(f"  Peer Views sections found: {[(r, i) for r, i in section_rows]}")

    # For each section, find column-E drug cell (only if CMPX is in the section)
    for hdr_row, indication in section_rows:
        # First check if CMPX appears in column E ticker rows
        has_cmpx = False
        for r in range(hdr_row + 1, hdr_row + 9):
            e_cell = cell_map.get((r, "E"))
            if e_cell and "CMPX" in e_cell[0]:
                has_cmpx = True
                break
        if not has_cmpx:
            continue

        # Find the drug name cell (first non-ticker E cell after the header)
        for r in range(hdr_row + 2, hdr_row + 9):
            e_cell = cell_map.get((r, "E"))
            if not e_cell:
                continue
            e_text, e_style = e_cell
            if "Equity" in e_text or "CMPX" in e_text:
                continue
            rating = style_rating.get(e_style, "AVG")
            result[indication] = rating
            log.info(f"  Peer Views: {indication} → {rating} "
                     f"(E{r}={e_text[:30]}, s={e_style})")
            break

    return result


def _detect_tam_sheets(xlsx_path: Path) -> Dict[str, Tuple[str, int]]:
    """Auto-detect TAM sheet names and max row counts.

    Returns: {"solid": ("TAM Solid+MM", 405), "blood": ("TAM Blood", 179)}
    Tries "TAM Solid" first, falls back to "TAM Solid+MM".
    """
    result: Dict[str, Tuple[str, int]] = {}
    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rid_to_path: Dict[str, str] = {}
        for rel in rels_xml:
            if "worksheet" in rel.get("Type", ""):
                rid = rel.get("Id", "")
                target = rel.get("Target", "")
                rid_to_path[rid] = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")

        sheet_map: Dict[str, str] = {}  # name → zip path
        for sheet in wb_xml.findall(f".//{{{_NS_MAIN}}}sheet"):
            name = sheet.get("name", "")
            rid = sheet.get(f"{{{_NS_R}}}id", "")
            if rid in rid_to_path:
                sheet_map[name] = rid_to_path[rid]

        # Detect solid TAM sheet (prefer "TAM Solid", fallback "TAM Solid+MM")
        for candidate in ("TAM Solid", "TAM Solid+MM"):
            if candidate in sheet_map:
                xml_bytes = zf.read(sheet_map[candidate])
                dim_m = re.search(rb'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"', xml_bytes)
                max_row = int(dim_m.group(1)) if dim_m else 500
                result["solid"] = (candidate, max_row)
                log.info(f"  TAM Solid: '{candidate}' (max row {max_row})")
                break

        # Detect blood TAM sheet
        if "TAM Blood" in sheet_map:
            xml_bytes = zf.read(sheet_map["TAM Blood"])
            dim_m = re.search(rb'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"', xml_bytes)
            max_row = int(dim_m.group(1)) if dim_m else 200
            result["blood"] = ("TAM Blood", max_row)
            log.info(f"  TAM Blood: 'TAM Blood' (max row {max_row})")

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  XML ROW BUILDERS
# ═══════════════════════════════════════════════════════════════════════════════

def _tc(addr: str, text: str, style: str) -> str:
    """Text cell (inlineStr)."""
    return f'<c r="{addr}" s="{style}" t="inlineStr"><is><t>{_xml_escape(text)}</t></is></c>'


def _fc(addr: str, formula: str, style: str, cached: str = "") -> str:
    """Formula cell."""
    if cached:
        return f'<c r="{addr}" s="{style}" t="str"><f>{formula}</f><v>{_xml_escape(cached)}</v></c>'
    return f'<c r="{addr}" s="{style}"><f>{formula}</f></c>'


def _nc(addr: str, val: float, style: str) -> str:
    """Numeric cell."""
    return f'<c r="{addr}" s="{style}"><v>{val}</v></c>'


def _ec(addr: str, style: str) -> str:
    """Empty styled cell."""
    return f'<c r="{addr}" s="{style}"/>'


def _build_stage_row(row: int, drug_name: str, full_name: str) -> str:
    """Build drug asset / stage row."""
    cells = [f'<row r="{row}">']
    cells.append(_tc(f"A{row}", "X", S['drug_a']))
    cells.append(_tc(f"D{row}", full_name, S['drug_d']))
    cells.append(_ec(f"E{row}", S['drug_e']))

    # Historical columns (F-R = 2010-2022): empty styled
    for year in range(2010, 2023):
        col = _year_to_col(year)
        cells.append(_ec(f"{col}{row}", S['drug_hist']))

    # Forecast columns (S-AH = 2023-2038): SUMIFS from Scenarios
    for year in range(2023, 2039):
        col = _year_to_col(year)           # Pipeline column
        scen_col = _scenarios_col_for_year(year)  # Scenarios column (10-col offset)
        formula = (f'SUMIFS(Scenarios!{scen_col}:{scen_col},'
                   f'Scenarios!$C:$C,Pipeline!$D{row},'
                   f'Scenarios!$A:$A,Pipeline!$E$2)')
        cells.append(_fc(f"{col}{row}", formula, S['drug_stage']))

    cells.append('</row>')
    return ''.join(cells)


def _build_tam_row(row: int, drug_row: int, label_suffix: str,
                   indication: str,
                   tam_sheet: str = "TAM Solid",
                   col_offset: int = 0,
                   max_tam_row: int = 562) -> str:
    """Build TAM row with direct SUMIF from TAM Solid.

    Formula: SUMIF('TAM Solid'!$D$9:$D$562, "indication", 'TAM Solid'!{col}$9:{col}$562)
    max_tam_row=562 covers drug rows R342-R399 after 58-drug + 3-incidence insertion.
    """
    cells = [f'<row r="{row}">']
    cells.append(_tc(f"A{row}", "X", S['drug_a']))
    # C column: blue-text indication label
    cells.append(_tc(f"C{row}", indication, S['tam_c']))
    formula_d = f'D{drug_row}&amp;" {_xml_escape(label_suffix)}"'
    cells.append(_fc(f"D{row}", formula_d, S['tam_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['tam_e']))

    esc_ind = _xml_escape(indication)
    esc_sheet = _xml_escape(tam_sheet)

    for year in range(2010, 2039):
        col = _year_to_col(year)
        col_idx = _COL_BASE + (year - _YEAR_BASE)
        tam_col = _col_letter(col_idx + col_offset)

        formula = (f"SUMIF('{esc_sheet}'!$D$9:$D${max_tam_row},"
                   f'"{esc_ind}",'
                   f"'{esc_sheet}'!{tam_col}$9:{tam_col}${max_tam_row})")

        cells.append(_fc(f"{col}{row}", formula, S['tam_data']))

    cells.append('</row>')
    return ''.join(cells)


def _build_ms_row(row: int, drug_row: int, label_suffix: str,
                  rating: str = "AVG") -> str:
    """Build Market Share row with SUMIFS from Scenarios.
    rating: "BIC"|"T1"|"AVG" from Peer Views.
    """
    cells = [f'<row r="{row}">']
    cells.append(_ec(f"A{row}", S['ms_a']))
    # C column: rating label (full descriptive text for display)
    _RATING_DISPLAY = {"AVG": "Average Growth", "BIC": "Best-In-Class Growth", "T1": "Tier One Growth"}
    cells.append(_tc(f"C{row}", _RATING_DISPLAY.get(rating, rating), S['ms_c']))
    formula_d = f'D{drug_row}&amp;" {_xml_escape(label_suffix)}"'
    cells.append(_fc(f"D{row}", formula_d, S['ms_d']))
    cells.append(_tc(f"E{row}", "[%]", S['tam_e']))

    # Historical (F-R): empty
    for year in range(2010, 2023):
        col = _year_to_col(year)
        cells.append(_ec(f"{col}{row}", S['ms_data']))

    # Forecast (S-AH): SUMIFS from Scenarios (with column offset correction)
    for year in range(2023, 2039):
        col = _year_to_col(year)           # Pipeline column
        scen_col = _scenarios_col_for_year(year)  # Scenarios column (10-col offset)
        formula = (f'SUMIFS(Scenarios!{scen_col}:{scen_col},'
                   f'Scenarios!$C:$C,Pipeline!$D{row},'
                   f'Scenarios!$A:$A,Pipeline!$E$2)')
        cells.append(_fc(f"{col}{row}", formula, S['ms_data']))

    cells.append('</row>')
    return ''.join(cells)


def _build_price_row(row: int, drug_row: int, price_mm: float) -> str:
    """Build List Price row."""
    cells = [f'<row r="{row}">']
    cells.append(_tc(f"A{row}", "X", S['drug_a']))
    formula_d = f'D{drug_row}&amp;" List Price (Per Patient)"'
    cells.append(_fc(f"D{row}", formula_d, S['price_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['tam_e']))

    # All columns: constant price
    for year in range(2010, 2039):
        col = _year_to_col(year)
        cells.append(_nc(f"{col}{row}", price_mm, S['price_data']))

    cells.append('</row>')
    return ''.join(cells)


def _build_revenue_row(row: int, drug_row: int, stage_row: int,
                       tam_ms_pairs: List[Tuple[int, int]],
                       growth_rows: List[int],
                       tam_solid_name: str = "TAM Solid") -> str:
    """Build Revenue row with per-indication maturity from TAM Solid.

    Revenue = IF(COUNTIF(stage_range,5)>0,
                 TAM1*MS1*MaturityFactor + TAM2*MS2*MaturityFactor + ...,
                 0)

    Maturity factors use INDEX into TAM Solid growth curve rows:
      R551 = Average Growth, R552 = Best-In-Class Growth, R553 = Tier One Growth.
    INDEX range is $F$row:$AH$row (columns F-AH = years 2010-2038).
    Maturity = INDEX(curve, MIN(years_since_approval, 29))
    where years_since_approval = COLUMN(col) - MATCH(5, stage_range) - COLUMN($F$1) + 2
    """
    cells = [f'<row r="{row}">']
    cells.append(_ec(f"A{row}", S['cogs_a']))
    formula_d = f'D{drug_row}&amp;" Revenue"'
    cells.append(_fc(f"D{row}", formula_d, S['rev_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['rev_e']))

    f_col = _year_to_col(2010)   # First data column (F)
    esc_sheet = _xml_escape(tam_solid_name)

    for year in range(2010, 2039):
        col = _year_to_col(year)

        # Build per-indication terms: TAM_i * MS_i * MaturityFactor_i
        terms = []
        for (tam_r, ms_r), g_row in zip(tam_ms_pairs, growth_rows):
            maturity = (
                f"INDEX('{esc_sheet}'!$F${g_row}:$AH${g_row},"
                f"MIN(COLUMN({col}1)-MATCH(5,${f_col}${stage_row}:$U${stage_row},0)"
                f"-COLUMN($F$1)+2,29))"
            )
            terms.append(f'{col}{tam_r}*{col}{ms_r}*{maturity}')

        if len(terms) == 1:
            expr = terms[0]
        else:
            expr = '+'.join(terms)

        formula = (
            f'IF(COUNTIF(${f_col}{stage_row}:{col}{stage_row},5)&gt;0,'
            f'{expr},0)'
        )
        cells.append(_fc(f"{col}{row}", formula, S['rev_data']))

    cells.append('</row>')
    return ''.join(cells)


def _build_cogs_row(row: int, drug_row: int, stage_row: int,
                    rev_row: int,
                    tam_solid_name: str = "TAM Solid") -> str:
    """Build COGS row.

    COGS = IF(COUNTIF(stage_range,5)>0, 'TAM Solid'!$P$562 * Revenue, 0)
    Uses flat COGS/Revenue ratio from TAM Solid Parameters R562 col P (0.37).
    Starts immediately when stage reaches 5 (FDA approval).
    """
    cells = [f'<row r="{row}">']
    cells.append(_ec(f"A{row}", S['cogs_a']))
    formula_d = f'D{drug_row}&amp;" COGS"'
    cells.append(_fc(f"D{row}", formula_d, S['cogs_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['rev_e']))

    f_col = _year_to_col(2010)   # F
    esc_sheet = _xml_escape(tam_solid_name)
    for year in range(2010, 2039):
        col = _year_to_col(year)
        formula = (
            f"IF(COUNTIF(${f_col}{stage_row}:{col}{stage_row},5)&gt;0,"
            f"'{esc_sheet}'!$P${COGS_PRICE_ROW}*{col}{rev_row},0)"
        )
        cells.append(_fc(f"{col}{row}", formula, S['cogs_data']))

    cells.append('</row>')
    return ''.join(cells)


def _build_separator_row(row: int) -> str:
    """Build blank separator row between drug blocks."""
    return f'<row r="{row}"><c r="B{row}" s="{S["sep_b"]}"><f>B{row-1}</f></c></row>'


# ═══════════════════════════════════════════════════════════════════════════════
#  DRUG BLOCK ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════

def build_drug_block(asset: PipelineAsset, start_row: int,
                     pricing: Dict[str, float],
                     default_price: float = 0.1,
                     override_full_name: Optional[str] = None,
                     override_indications: Optional[List[str]] = None,
                     tam_sheets: Optional[Dict[str, Tuple[str, int]]] = None,
                     indication_ratings: Optional[Dict[str, str]] = None,
                     tam_solid_name: str = "TAM Solid",
                     ) -> Tuple[List[str], int]:
    """Build all rows for one drug.

    Returns: (list_of_row_xml_strings, next_available_row)

    Revenue uses per-indication growth factors from TAM Solid Parameters
    (R551 AVG, R552 BIC, R553 T1) based on Peer Views ratings.
    TAM uses SUMIF from Pipeline Referred Tables for existing indications,
    or cross-sheet SUMIF from TAM Solid/Blood for new indications.
    COGS references TAM Solid Parameters COGS/Price (R562).

    indication_ratings: {indication: "BIC"|"T1"|"AVG"}
        from _read_peer_views_ratings(). Defaults to AVG if missing.
    """
    rows: List[str] = []
    cur = start_row
    full_name = override_full_name or _asset_full_name(asset)

    # Row 1: Stage / Drug header
    drug_row = cur
    stage_row = cur
    rows.append(_build_stage_row(cur, asset.name, full_name))
    cur += 1

    # Per-indication TAM + MS rows
    if override_indications:
        indications = override_indications
    else:
        indications = list(asset.market_shares.keys())
        if not indications:
            indications = ["All"]

    ratings_map = indication_ratings or {}
    tam_ms_pairs: List[Tuple[int, int]] = []
    ind_ratings: List[str] = []  # per-indication rating strings

    for ind in indications:
        # Resolve rating for this indication
        ind_rating = ratings_map.get(ind)
        if ind_rating is None:
            alias = _INDICATION_ALIASES.get(ind)
            if alias:
                ind_rating = ratings_map.get(alias)
        if ind_rating is None:
            ind_rating = "AVG"

        # TAM row — determine source (Pipeline Referred Tables or TAM sheet)
        tam_label = f"{ind} TAM" if ind not in ("All", "All Indications Combined") else "TAM"
        tam_row = cur

        cross_ref = _TAM_CROSS_REF.get(ind)
        if cross_ref and tam_sheets and cross_ref in tam_sheets:
            sheet_name, max_row = tam_sheets[cross_ref]
            offset = _TAM_BLOOD_COL_OFFSET if cross_ref == "blood" else 0
            rows.append(_build_tam_row(cur, drug_row, tam_label, ind,
                                       tam_sheet=sheet_name,
                                       col_offset=offset,
                                       max_tam_row=max_row))
            log.info(f"    {ind} TAM → cross-sheet SUMIF from '{sheet_name}'")
        else:
            # Default: direct SUMIF from TAM Solid (not Pipeline internal)
            rows.append(_build_tam_row(cur, drug_row, tam_label, ind,
                                       tam_sheet=tam_solid_name))
            log.info(f"    {ind} TAM → SUMIF from '{tam_solid_name}'")

        cur += 1

        # MS row — with rating label in C column
        ms_label = f"{ind} Market Share" if ind not in ("All", "All Indications Combined") else "Market Share"
        ms_row = cur
        rows.append(_build_ms_row(cur, drug_row, ms_label, rating=ind_rating))
        cur += 1

        tam_ms_pairs.append((tam_row, ms_row))
        ind_ratings.append(ind_rating)

    # Price row — kept for informational display
    price = default_price
    for ind in indications:
        if ind in pricing:
            price = pricing[ind]
            break
        for pk, pv in pricing.items():
            if f"({ind})" in pk or pk.upper().startswith(ind.upper()):
                price = pv
                break
        if price != default_price:
            break
    if price == default_price and pricing:
        price = next(iter(pricing.values()))

    price_row = cur
    rows.append(_build_price_row(cur, drug_row, price))
    cur += 1

    # Per-indication growth factor rows from TAM Solid Parameters
    growth_rows: List[int] = []
    rating_strs: List[str] = []
    for ind, ind_rating in zip(indications, ind_ratings):
        g_row = GROWTH_ROW.get(ind_rating, GROWTH_ROW['AVG'])
        growth_rows.append(g_row)
        rating_strs.append(f"{ind}:{ind_rating}")

    # Revenue row — per-indication growth factors from TAM Solid
    rev_row = cur
    rows.append(_build_revenue_row(cur, drug_row, stage_row,
                                   tam_ms_pairs, growth_rows,
                                   tam_solid_name=tam_solid_name))
    cur += 1

    # COGS row — references TAM Solid COGS/Price (R562)
    rows.append(_build_cogs_row(cur, drug_row, stage_row, rev_row,
                                tam_solid_name=tam_solid_name))
    cur += 1

    # Separator
    rows.append(_build_separator_row(cur))
    cur += 1

    log.info(f"  {asset.name}: rows {start_row}-{cur-1} "
             f"({len(indications)} ind, ratings=[{', '.join(rating_strs)}], "
             f"price=${price:.3f}MM)")
    return rows, cur


def _build_rev_sum_row(row: int, rev_rows: List[int]) -> str:
    """Build the Operating Revenue Sum row (R459) with SUM formula."""
    cells = [f'<row r="{row}">']
    cells.append(_tc(f"D{row}", "Operating Revenue From Sales", S['sum_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['rev_e']))

    # Build SUM formula across all revenue rows
    for year in range(2010, 2039):
        col = _year_to_col(year)
        if rev_rows:
            refs = ','.join(f'{col}{r}' for r in rev_rows)
            formula = f'SUM({refs})'
            cells.append(_fc(f"{col}{row}", formula, S['sum_data']))
        else:
            cells.append(_ec(f"{col}{row}", S['sum_data']))

    cells.append('</row>')
    return ''.join(cells)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def generate_pipeline(
    xlsx_path: Path,
    assets: List[PipelineAsset],
    pricing: Dict[str, Dict[str, float]],
    dry_run: bool = False,
) -> None:
    """Generate Revenue Forecasting section in Pipeline sheet."""

    sheet_zip = _get_sheet_zip_path(xlsx_path, SHEET_NAME)
    if not sheet_zip:
        log.error(f"Cannot find '{SHEET_NAME}' sheet")
        return

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")
        # Validate styles.xml has expected xf count (hardcoded S indices depend on it)
        styles_xml = zf.read("xl/styles.xml").decode("utf-8")
    xf_count_m = re.search(r'<cellXfs\s+count="(\d+)"', styles_xml)
    if xf_count_m:
        xf_count = int(xf_count_m.group(1))
        if xf_count < 700:
            log.error(
                f"styles.xml has {xf_count} cellXf entries (expected ~778+). "
                f"Style indices may be invalid. Check S dict values."
            )
            return
    log.info(f"Read {sheet_zip}: {len(xml):,} chars")

    # ── Step 1: Remove existing drug blocks (rows after REV_SUM) ──
    # Find Revenue Sum row end position
    rev_sum_re = re.search(
        rf'(<row r="{REV_SUM}"[^>]*>.*?</row>)',
        xml, re.DOTALL
    )
    if not rev_sum_re:
        # Try self-closing
        rev_sum_re = re.search(
            rf'(<row r="{REV_SUM}"[^/]*/>\s*)',
            xml
        )

    sheetdata_end = xml.find('</sheetData>')
    if sheetdata_end == -1:
        log.error("Cannot find </sheetData>")
        return

    if rev_sum_re:
        # Keep everything up to and including REV_SUM row, remove everything after
        # until </sheetData>
        insert_point = rev_sum_re.end()
    else:
        # No REV_SUM row found, insert at sheetdata_end
        insert_point = sheetdata_end

    xml_before = xml[:insert_point]
    xml_after = xml[sheetdata_end:]  # includes </sheetData> and beyond

    # ── Step 1b: Read authoritative drug names from Scenarios sheet ──
    scenarios_info = _read_scenarios_drug_info(xlsx_path)

    # ── Step 1c: Detect TAM sheets for cross-sheet references ──
    tam_sheets = _detect_tam_sheets(xlsx_path)
    if tam_sheets:
        log.info(f"TAM sheets detected: {list(tam_sheets.keys())}")
    else:
        log.warning("No TAM sheets found — TAM rows will use Pipeline SUMIF only")

    # ── Step 1d: Read per-indication ratings from Peer Views ──
    indication_ratings = _read_peer_views_ratings(xlsx_path)

    # ── Step 1e: Resolve TAM Solid sheet name for cross-sheet references ──
    tam_solid_name = "TAM Solid"
    if tam_sheets and "solid" in tam_sheets:
        tam_solid_name = tam_sheets["solid"][0]
    log.info(f"  TAM Solid sheet name: '{tam_solid_name}'")

    # ── Step 2: Build new drug blocks ──
    log.info(f"\nBuilding revenue forecasting for {len(assets)} drugs...")
    all_rows: List[str] = []
    rev_rows: List[int] = []  # Track revenue row numbers for sum formula
    cur = FIRST_DRUG

    for asset in assets:
        # Use exact Scenarios name + indications (critical for SUMIFS matching)
        info = scenarios_info.get(asset.name)
        if info:
            sc_full_name, sc_indications = info
            log.info(f"  {asset.name}: using Scenarios name → {sc_full_name}")
            log.info(f"    Scenarios indications: {sc_indications}")
        else:
            sc_full_name = None
            sc_indications = None
            log.warning(f"  {asset.name}: no Scenarios match, using parsed name")

        drug_pricing = pricing.get(asset.name, {})
        block_rows, cur = build_drug_block(
            asset, cur, drug_pricing,
            override_full_name=sc_full_name,
            override_indications=sc_indications or None,
            tam_sheets=tam_sheets,
            indication_ratings=indication_ratings,
            tam_solid_name=tam_solid_name,
        )
        # Revenue row is the second-to-last row before separator
        rev_rows.append(cur - 3)  # COGS=cur-2, rev=cur-3
        all_rows.extend(block_rows)

    # ── Step 3: Rebuild Revenue Sum row with SUM across all drugs ──
    rev_sum_xml = _build_rev_sum_row(REV_SUM, rev_rows)

    # Remove old REV_SUM from xml_before if present
    if rev_sum_re:
        xml_before = xml_before[:rev_sum_re.start()] + rev_sum_xml
    else:
        # Insert before the insert_point
        xml_before = xml_before + rev_sum_xml

    if dry_run:
        log.info(f"\nDry run: would write {len(all_rows)} rows ({FIRST_DRUG}-{cur-1})")
        for asset in assets:
            log.info(f"  {_asset_full_name(asset)}")
        return

    # ── Step 4: Assemble final XML ──
    new_xml = xml_before + '\n' + '\n'.join(all_rows) + '\n' + xml_after

    # Update dimension
    new_max_row = cur - 1
    new_xml = re.sub(
        r'<dimension ref="[^"]*"/>',
        f'<dimension ref="A1:AH{new_max_row}"/>',
        new_xml
    )
    log.info(f"Generated rows {FIRST_DRUG}-{new_max_row} ({len(all_rows)} XML rows)")

    # ── Step 4b: Sanitize formula XML to prevent "Removed Records" ──
    new_xml = _normalize_formula_xml(new_xml)

    # ── Step 5: Surgical zip patch ──
    modified: Dict[str, bytes] = {sheet_zip: new_xml.encode("utf-8")}

    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        ct = zf.read("[Content_Types].xml").decode("utf-8")
        wr = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        log.info("Added fullCalcOnLoad to workbook.xml")
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    # Strip calcChain.xml references
    ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', '', ct)
    wr = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', wr)
    modified["[Content_Types].xml"] = ct.encode("utf-8")
    modified["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~pipeline.xlsx")
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

    log.info(f"Pipeline sheet saved → {xlsx_path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Generate Pipeline sheet Revenue Forecasting from Gemini research"
    )
    parser.add_argument("--ticker", required=True, help="Stock ticker (e.g. CMPX)")
    parser.add_argument("--company-name", required=True, help="Full company name")
    parser.add_argument("--report-dir",
                        help="Directory with per-drug research .md files "
                             "(default: DD/{TICKER}/pipeline_base4/)")
    parser.add_argument("--pricing-dir",
                        help="Directory with pricing .md files (default: same as --report-dir)")
    parser.add_argument("--dcf-file",
                        help="DCF file path (auto-detected if not specified)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview changes without writing")
    args = parser.parse_args()

    # Locate report directory
    if args.report_dir:
        report_dir = Path(args.report_dir)
    else:
        report_dir = Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/pipeline_base4")

    if not report_dir.exists():
        log.error(f"Report directory not found: {report_dir}")
        return

    # Locate DCF file
    if args.dcf_file:
        xlsx_path = Path(args.dcf_file)
    else:
        xlsx_path = Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/DCF {args.ticker}.xlsx")

    if not xlsx_path.exists():
        log.error(f"DCF file not found: {xlsx_path}")
        return

    pricing_dir = Path(args.pricing_dir) if args.pricing_dir else report_dir

    log.info(f"Ticker: {args.ticker}")
    log.info(f"Reports: {report_dir}")
    log.info(f"Pricing: {pricing_dir}")
    log.info(f"DCF: {xlsx_path}")
    log.info(f"Mode: {'DRY RUN' if args.dry_run else 'LIVE'}")

    # ── Step 1: Parse Gemini research reports ──
    log.info(f"\n{'='*60}")
    log.info("STEP 1: Parsing Gemini research reports")
    log.info(f"{'='*60}")

    assets = parse_gemini_reports(report_dir, args.ticker)
    if not assets:
        log.error("No pipeline assets found in reports")
        return

    for asset in assets:
        ind_list = list(asset.market_shares.keys())
        log.info(f"  {asset.name} ({asset.target}): {ind_list}")

    # ── Step 2: Parse pricing data ──
    log.info(f"\n{'='*60}")
    log.info("STEP 2: Parsing pricing data")
    log.info(f"{'='*60}")

    pricing = parse_pricing_reports(pricing_dir, args.ticker)
    if not pricing:
        log.warning("No pricing data found — using default $0.10 MM per patient")

    # ── Step 3: Backup ──
    if not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_pipeline_{ts}.xlsx")
        shutil.copy2(xlsx_path, backup)
        log.info(f"Backup: {backup}")

    # ── Step 4: Generate ──
    log.info(f"\n{'='*60}")
    log.info("STEP 3: Generating Pipeline Revenue Forecasting")
    log.info(f"{'='*60}")

    generate_pipeline(xlsx_path, assets, pricing, args.dry_run)

    # ── Summary ──
    total_ind = sum(len(a.market_shares) or 1 for a in assets)
    priced = sum(1 for a in assets if a.name in pricing)
    print(f"\n{'='*60}")
    print("Pipeline Revenue Forecasting Generated")
    print(f"{'='*60}")
    print(f"  Assets:      {len(assets)}")
    print(f"  Indications: {total_ind}")
    print(f"  With pricing: {priced}/{len(assets)}")
    for asset in assets:
        inds = list(asset.market_shares.keys()) or ["All"]
        prices = pricing.get(asset.name, {})
        price_str = ", ".join(f"{k}=${v:.3f}MM" for k, v in prices.items()) if prices else "default"
        print(f"    {asset.name} ({asset.target}): {'/'.join(inds)} [{price_str}]")
    print(f"  File: {xlsx_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
