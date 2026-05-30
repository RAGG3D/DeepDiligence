#!/usr/bin/env python3
"""
generate_peer_views.py — Restyle Peer Views sections with correct per-drug rating colors.

Copies v1 rows 4-201 directly from DCF Template 2020.xlsx (with style bridging),
then restyles v3/v4/PV sections with per-drug BIC/T1/AVG colors.
Uses surgical zip patching (NEVER openpyxl .save()).

Replaces: patch_peer_views_v4.py, v5, v6

Usage:
    python generate_peer_views.py [--dry-run] [--file path]
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

# Default DCF file
_DEFAULT_XLSX = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
_TEMPLATE_XLSX = Path("/mnt/c/Users/yzsun/Desktop/DD/base/DCF Template 2020.xlsx")
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ═══════════════════════════════════════════════════════════════════════════════
#  STYLE MAPPINGS — verified against styles.xml fillId 16/17/18
# ═══════════════════════════════════════════════════════════════════════════════
#
#  fillId 16 = BIC (green, theme 9 accent6)
#  fillId 17 = AVG (gold, theme 7 accent4)
#  fillId 18 = T1  (blue, theme 8 accent5)

RATING_STYLES: Dict[str, Dict[str, str]] = {
    'BIC': {
        'text': '847', 'drug': '862', 'date': '850',
        'pct': '852', 'sae_pct': '864', 'sale': '854',
        'n': '855', 'pfs': '479',
        'mkt_react': '858', 'sae_text': '856', 'phase': '843',
    },
    'T1': {
        'text': '655', 'drug': '648', 'date': '683',
        'pct': '669', 'sae_pct': '669', 'sale': '709',
        'n': '703', 'pfs': '675',
        'mkt_react': '858', 'sae_text': '856', 'phase': '763',
        # NOTE: mkt_react/sae_text are PLACEHOLDERS (green/BIC).
        # Actual blue IDs created at runtime by _create_rating_styles().
    },
    'AVG': {
        'text': '654', 'drug': '647', 'date': '661',
        'pct': '668', 'sae_pct': '668', 'sale': '708',
        'n': '702', 'pfs': '674',
        'mkt_react': '858', 'sae_text': '856', 'phase': '779',
        # NOTE: mkt_react/sae_text are PLACEHOLDERS (green/BIC).
        # Actual gold IDs created at runtime by _create_rating_styles().
    },
    'SUSPEND': {
        # Gray fill (fillId=4). text/drug/phase use existing s=711/713.
        # date/pct/sae_pct/n/pfs/sale/sae_text/mkt_react are PLACEHOLDERS
        # (s=711) — actual numFmtId-preserving IDs created at runtime.
        'text': '711', 'drug': '713', 'date': '711',
        'pct': '711', 'sae_pct': '711', 'sale': '711',
        'n': '711', 'pfs': '711',
        'mkt_react': '711', 'sae_text': '711', 'phase': '711',
    },
    'NEUTRAL': {
        # No fill (fillId=0) — original unfilled styles.
        # mkt_react is PLACEHOLDER (s=849/text) — actual bold style
        # created at runtime by _create_rating_styles().
        'text': '849', 'drug': '860', 'date': '861',
        'pct': '577', 'sae_pct': '863', 'sale': '857',
        'n': '853', 'pfs': '569',
        'mkt_react': '849', 'sae_text': '863', 'phase': '849',
    },
}

# X-column rating legend labels (added at offsets +3/+4/+5/+6)
X_LABELS: List[Tuple[int, str, str]] = [
    (3, "Best-In-Class", "470"),   # BIC fill, borderId=7
    (4, "Tier One",      "471"),   # T1 fill, borderId=28
    (5, "Average",       "472"),   # AVG fill, borderId=28
    (6, "Suspend",       "473"),   # fillId=4, borderId=8
]

# ═══════════════════════════════════════════════════════════════════════════════
#  ROW TYPE MAPPINGS — offset from section base_row → cell_type
# ═══════════════════════════════════════════════════════════════════════════════

# v4 format (25-row compact sections, rows 273-486)
V4_OFFSETS: Dict[int, str] = {
    2:  'text',       # BBG Equity ticker
    3:  'text',       # Company ticker
    4:  'drug',       # Drug name (bold font)
    5:  'text',       # Innovation
    6:  'text',       # Result
    7:  'text',       # Total Treatment Line
    8:  'n',          # Median Treatment Line (numFmtId=202)
    9:  'pct',        # ≥G3 SAE/Patients (numFmtId=193)
    10: 'sae_text',   # ≥G3 clinical AE (shared s=856)
    11: 'date',       # Date (numFmtId=15)
    12: 'mkt_react',  # Market Reaction -1d (shared s=858)
    13: 'mkt_react',  # Market Reaction +1d (shared s=858)
    14: 'phase',      # Readout Phase (shared s=843)
    15: 'text',       # Treatment Line
    16: 'text',       # Evaluable Patients (numFmtId=0 in v4)
    17: 'pct',        # ORR
    18: 'pct',        # CR
    19: 'pct',        # PR
    20: 'pct',        # DCR
    21: 'pfs',        # Median PFS (numFmtId=203)
    22: 'pfs',        # Median OS
    23: 'sale',       # Latest Sale MM USD (numFmtId=185)
}

# v3 format (33-row detailed sections, CRC/BTC rows 203-272)
# All v3 data cells use numFmtId=0 (General), so we use 'text' type
# to change fill color without altering number format.
V3_DRUG_OFFSET = 4   # Drug name row
V3_DATA_START = 5     # First data row (Innovation)
V3_DATA_END = 32      # Last data row (1st Yr Sale)

# ═══════════════════════════════════════════════════════════════════════════════
#  SECTION DEFINITIONS — drug column → rating
# ═══════════════════════════════════════════════════════════════════════════════

# v1 format (variable-length sections, rows 13-200) — up to 18 drug columns E-V
# Ratings verified cell-by-cell against DCF Template 2020.xlsx sheet20.
# NEUTRAL resets columns previously colored with wrong ratings.
# Uses offset-based restyling (same cell_types as V4_OFFSETS).
V1_SECTIONS: List[Dict] = [
    {"name": "mUC Drug",              "base": 13,  "drugs": {
        "E": "BIC",  "F": "BIC",  "I": "NEUTRAL", "J": "BIC",
        "K": "NEUTRAL", "L": "NEUTRAL", "N": "NEUTRAL", "O": "NEUTRAL",
        "Q": "NEUTRAL", "T": "NEUTRAL", "U": "NEUTRAL"}},
    {"name": "+KEYTRUDA 1L Melanoma", "base": 39,  "drugs": {
        "E": "BIC",  "F": "BIC",  "G": "AVG",  "H": "T1",
        "J": "AVG",  "K": "T1",   "M": "AVG",
        "N": "SUSPEND", "O": "NEUTRAL"}},
    {"name": "+KEYTRUDA 1L NSCLC",    "base": 58,  "drugs": {
        "E": "T1",   "F": "AVG",  "G": "AVG",  "H": "AVG",
        "J": "BIC",  "K": "BIC",  "L": "AVG",  "M": "AVG",
        "N": "SUSPEND"}},
    {"name": "ROS1-Positive NSCLC",   "base": 76,  "drugs": {
        "E": "T1",   "F": "NEUTRAL", "G": "NEUTRAL", "H": "NEUTRAL",
        "I": "BIC",  "J": "NEUTRAL", "K": "NEUTRAL", "L": "NEUTRAL",
        "N": "NEUTRAL", "O": "NEUTRAL", "Q": "NEUTRAL",
        "R": "BIC",  "T": "NEUTRAL", "U": "NEUTRAL", "V": "AVG"}},
    {"name": "ALK-Positive NSCLC",    "base": 100, "drugs": {
        "E": "BIC",  "F": "NEUTRAL", "G": "NEUTRAL", "H": "NEUTRAL",
        "I": "NEUTRAL", "J": "NEUTRAL", "K": "NEUTRAL", "L": "NEUTRAL",
        "N": "T1",   "O": "NEUTRAL", "P": "BIC",
        "Q": "NEUTRAL", "S": "T1",   "T": "NEUTRAL", "U": "NEUTRAL",
        "V": "AVG"}},
    {"name": "EGFR-Positive NSCLC",   "base": 124, "drugs": {
        "E": "BIC",  "F": "AVG",  "G": "T1",   "H": "T1",   "I": "AVG",
        "J": "NEUTRAL", "K": "NEUTRAL", "L": "NEUTRAL",
        "N": "NEUTRAL", "O": "NEUTRAL", "Q": "NEUTRAL",
        "T": "NEUTRAL", "U": "NEUTRAL"}},
    {"name": "+KEYTRUDA 1L HNSCC",    "base": 142, "drugs": {
        "E": "T1",   "F": "BIC",  "G": "AVG",
        "I": "AVG",  "J": "AVG",  "K": "AVG",
        "M": "BIC",  "N": "T1",   "O": "AVG",  "P": "SUSPEND"}},
    {"name": "mCRPC Drug",            "base": 161, "drugs": {
        "E": "SUSPEND", "F": "SUSPEND", "G": "SUSPEND", "H": "NEUTRAL",
        "I": "BIC",  "J": "AVG",  "K": "T1",   "L": "T1",
        "M": "AVG",  "N": "T1",   "O": "BIC",  "P": "BIC",
        "Q": "T1",   "R": "T1",   "S": "AVG",  "T": "AVG",
        "U": "BIC",  "V": "T1"}},
    {"name": "IDH1 GBM Drug",         "base": 183, "drugs": {
        "E": "AVG",  "F": "BIC",  "G": "NEUTRAL", "H": "AVG",
        "I": "NEUTRAL", "J": "SUSPEND",
        "K": "NEUTRAL", "L": "NEUTRAL", "N": "NEUTRAL", "O": "NEUTRAL",
        "Q": "NEUTRAL", "T": "NEUTRAL", "U": "NEUTRAL"}},
]

V4_SECTIONS: List[Dict] = [
    {"name": "RCC 2L Post-IO",               "base": 274, "drugs": {
        "E": "AVG", "F": "T1", "G": "AVG", "H": "T1"}},
    {"name": "HCC 2L",                        "base": 301, "drugs": {
        "E": "AVG", "F": "BIC", "G": "AVG"}},
    {"name": "Endometrial Cancer pMMR",        "base": 328, "drugs": {
        "E": "BIC"}},
    {"name": "Melanoma Post-PD1 (NCAM+)",      "base": 355, "drugs": {
        "E": "T1", "F": "BIC"}},
    {"name": "SCLC 3L+",                      "base": 382, "drugs": {
        "E": "AVG", "F": "BIC", "G": "T1", "H": "BIC"}},
    {"name": "NSCLC 2L Post-CPI",              "base": 409, "drugs": {
        "E": "AVG", "F": "BIC", "G": "T1"}},
    {"name": "TNBC 2L Post-CPI",               "base": 436, "drugs": {
        "E": "AVG", "F": "BIC", "G": "T1"}},
    {"name": "cHL Post-CPI (Salvage)",         "base": 463, "drugs": {
        "E": "T1", "F": "BIC"}},
]

# Peer View (sheet8) — compact BTC/CRC tables with same v4-style offsets
# These have BBG ticker at offset+2 (s=575), sec ticker at +3, drug at +4,
# data at +5 to +23.  Same layout as V4_OFFSETS.
PV_SECTIONS: List[Dict] = [
    {"name": "BTC 2L Drug", "base": 5, "drugs": {
        "E": "T1",  "F": "T1",  "G": "BIC", "H": "T1",
        "I": "AVG", "J": "AVG", "K": "AVG"}},
    {"name": "CRC 3L Drug", "base": 32, "drugs": {
        "E": "T1",  "F": "BIC", "G": "AVG", "H": "AVG"}},
]

# CRC/BTC sections (v3 format, need ticker + X label patches)
V3_SECTIONS: List[Dict] = [
    {"name": "BTC", "base": 203, "drugs": {
        "E": "T1", "F": "T1", "G": "BIC", "H": "AVG",
        "I": "T1", "J": "AVG", "K": "AVG"},
     "tickers": {
        "E": "CMPX US Equity",   # CTX-009 mono
        "F": "CMPX US Equity",   # CTX-009 combo
        "G": "JAZZ US Equity",   # Zanidatamab
        "I": "INCY US Equity",   # Pemigatinib
    }},
    {"name": "CRC", "base": 238, "drugs": {
        "E": "T1", "F": "BIC", "G": "AVG", "H": "AVG"},
     "tickers": {
        "E": "CMPX US Equity",   # CTX-009 / Tovecimig
        "F": "TAK US Equity",    # Fruquintinib (Takeda/HUTCHMED)
        "G": "BAYN GR Equity",   # Regorafenib (Bayer)
    }},
]


# ═══════════════════════════════════════════════════════════════════════════════
#  XML HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _col_idx(col_str: str) -> int:
    """Column letter(s) → 1-based index. A=1, Z=26, AA=27."""
    result = 0
    for c in col_str.upper():
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


def _find_cell(xml: str, addr: str) -> Optional[Tuple[int, int, str, str]]:
    """Find cell by address. Returns (lt, end_pos, open_tag, inner) or None."""
    search = f'r="{addr}"'
    start = 0
    while True:
        pos = xml.find(search, start)
        if pos == -1:
            return None
        lt = xml.rfind("<", 0, pos)
        if lt == -1 or xml[lt + 1] != "c" or xml[lt + 2] not in (" ", "\t", "\n", "/", ">"):
            start = pos + 1
            continue
        tag_end = xml.index(">", lt) + 1
        open_tag = xml[lt:tag_end]
        if xml[tag_end - 2:tag_end] == "/>":
            return (lt, tag_end, open_tag, "")
        c_end = xml.index("</c>", tag_end) + 4
        inner = xml[tag_end:c_end - 4]
        return (lt, c_end, open_tag, inner)


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def _restyle_cell(xml: str, addr: str, new_style: str) -> str:
    """Change s='...' attribute on an existing cell. No-op if cell not found."""
    found = _find_cell(xml, addr)
    if not found:
        return xml
    lt, end_pos, open_tag, inner = found

    # Replace or add style attribute
    if re.search(r's="\d+"', open_tag):
        new_tag = re.sub(r's="\d+"', f's="{new_style}"', open_tag)
    else:
        new_tag = open_tag.replace(f'r="{addr}"', f'r="{addr}" s="{new_style}"')

    if open_tag.endswith("/>"):
        # Self-closing tag
        return xml[:lt] + new_tag + xml[end_pos:]
    # Normal tag: reconstruct with inner content
    return xml[:lt] + new_tag + inner + "</c>" + xml[end_pos:]


def _patch_text_cell(xml: str, addr: str, text: str,
                     style: Optional[str] = None) -> str:
    """Write/replace text cell as inlineStr. Preserves xml[:lt] prefix."""
    escaped = _xml_escape(text)
    found = _find_cell(xml, addr)

    if found:
        lt, end_pos, open_tag, inner = found
        # Clean open tag: remove old type, add inlineStr
        tag = re.sub(r'\s+t="[^"]*"', '', open_tag)
        tag = tag.replace("/>", ">")
        if not tag.endswith(">"):
            tag += ">"
        tag = re.sub(r'(r="[^"]*")', r'\1 t="inlineStr"', tag)
        if style:
            if re.search(r's="[^"]*"', tag):
                tag = re.sub(r's="[^"]*"', f's="{style}"', tag)
            else:
                tag = re.sub(r'(r="[^"]*")', rf'\1 s="{style}"', tag)
        return xml[:lt] + f'{tag}<is><t>{escaped}</t></is></c>' + xml[end_pos:]

    # Cell not found — insert new
    row_num = int("".join(c for c in addr if c.isdigit()))
    s_attr = f' s="{style}"' if style else ""
    new_cell = f'<c r="{addr}"{s_attr} t="inlineStr"><is><t>{escaped}</t></is></c>'
    return _insert_cell(xml, row_num, addr, new_cell)


def _insert_cell(xml: str, row_num: int, addr: str, cell_xml: str) -> str:
    """Insert a cell XML into the correct row in column order."""
    row_search = f'r="{row_num}"'
    row_pos = 0
    while True:
        rp = xml.find(row_search, row_pos)
        if rp == -1:
            log.warning(f"  Row {row_num} not found; cannot insert {addr}")
            return xml
        lt = xml.rfind("<", 0, rp)
        if lt != -1 and xml[lt + 1:lt + 4] == "row":
            break
        row_pos = rp + 1

    row_tag_end = xml.index(">", lt) + 1

    # Self-closing row?
    if xml[row_tag_end - 2:row_tag_end] == "/>":
        return xml[:row_tag_end - 2] + f">{cell_xml}</row>" + xml[row_tag_end:]

    row_end = xml.index("</row>", row_tag_end)
    row_body = xml[row_tag_end:row_end]
    col_letter = "".join(c for c in addr if c.isalpha())
    target_idx = _col_idx(col_letter)

    # Find insertion point in column order
    cells = list(re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"', row_body))
    insert_at = len(row_body)
    for m in cells:
        if _col_idx(m.group(1)) > target_idx:
            insert_at = m.start()
            break

    new_body = row_body[:insert_at] + cell_xml + row_body[insert_at:]
    return xml[:row_tag_end] + new_body + xml[row_end:]


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


# ═══════════════════════════════════════════════════════════════════════════════
#  V1 TEMPLATE COPY — copy rows 4-201 from DCF Template 2020.xlsx
# ═══════════════════════════════════════════════════════════════════════════════

def _font_sig(font_el: ET.Element) -> str:
    """Semantic signature for a <font> element (order-independent)."""
    parts = []
    for child in font_el:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('b', 'i', 'u', 'strike'):
            parts.append(tag)
        elif tag == 'sz':
            parts.append(f'sz={child.get("val")}')
        elif tag == 'name':
            parts.append(f'name={child.get("val")}')
        elif tag == 'family':
            parts.append(f'fam={child.get("val")}')
        elif tag == 'color':
            c = (f'{child.get("theme", "")}/{child.get("tint", "")}/'
                 f'{child.get("rgb", "")}/{child.get("indexed", "")}')
            parts.append(f'color={c}')
        elif tag == 'scheme':
            parts.append(f'scheme={child.get("val")}')
    return '|'.join(sorted(parts))


def _fill_sig(fill_el: ET.Element) -> str:
    """Semantic signature for a <fill> element."""
    ns = _NS_MAIN
    pf = fill_el.find(f'{{{ns}}}patternFill')
    if pf is None:
        return 'none'
    pt = pf.get('patternType', 'none')
    parts = [f'pt={pt}']
    for child in pf:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        attrs = ','.join(f'{k}={v}' for k, v in sorted(child.attrib.items()))
        parts.append(f'{tag}={attrs}')
    return '|'.join(sorted(parts))


def _border_sig(border_el: ET.Element) -> str:
    """Semantic signature for a <border> element."""
    parts = []
    for child in border_el:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        style = child.get('style', '')
        colors = []
        for cc in child:
            colors.append(','.join(f'{k}={v}' for k, v in sorted(cc.attrib.items())))
        parts.append(f'{tag}:{style}:{"_".join(colors)}')
    return '||'.join(sorted(parts))


def _build_style_bridge(
    tmpl_styles_raw: str, cmpx_styles_raw: str,
    needed_style_ids: set,
) -> Tuple[str, Dict[int, int]]:
    """Build a style mapping from template → CMPX, adding missing definitions.

    For each template style ID used in rows 4-201, finds or creates an
    equivalent xf entry in CMPX's styles.xml. Adds missing fonts, fills,
    borders as needed.

    Returns: (modified_cmpx_styles_xml, {tmpl_style_id: cmpx_style_id})
    """
    ns = {'x': _NS_MAIN}

    tmpl_root = ET.fromstring(tmpl_styles_raw)
    cmpx_root = ET.fromstring(cmpx_styles_raw)

    tmpl_fonts = tmpl_root.findall('.//x:fonts/x:font', ns)
    tmpl_fills = tmpl_root.findall('.//x:fills/x:fill', ns)
    tmpl_borders = tmpl_root.findall('.//x:borders/x:border', ns)

    cmpx_fonts = cmpx_root.findall('.//x:fonts/x:font', ns)
    cmpx_fills = cmpx_root.findall('.//x:fills/x:fill', ns)
    cmpx_borders = cmpx_root.findall('.//x:borders/x:border', ns)

    # Build CMPX semantic lookups
    cmpx_font_lk = {}
    for i, f in enumerate(cmpx_fonts):
        s = _font_sig(f)
        if s not in cmpx_font_lk:
            cmpx_font_lk[s] = i

    cmpx_fill_lk = {}
    for i, f in enumerate(cmpx_fills):
        s = _fill_sig(f)
        if s not in cmpx_fill_lk:
            cmpx_fill_lk[s] = i

    cmpx_border_lk = {}
    for i, b in enumerate(cmpx_borders):
        s = _border_sig(b)
        if s not in cmpx_border_lk:
            cmpx_border_lk[s] = i

    # Map each template font/fill/border to CMPX (or create new)
    font_map: Dict[int, int] = {}
    new_font_xml: List[str] = []
    n_cmpx_fonts = len(cmpx_fonts)
    for i, f in enumerate(tmpl_fonts):
        sig = _font_sig(f)
        if sig in cmpx_font_lk:
            font_map[i] = cmpx_font_lk[sig]
        else:
            new_id = n_cmpx_fonts + len(new_font_xml)
            font_map[i] = new_id
            cmpx_font_lk[sig] = new_id
            xml_str = ET.tostring(f, encoding='unicode', short_empty_elements=True)
            xml_str = re.sub(r'\s*xmlns[^"]*"[^"]*"', '', xml_str)
            xml_str = re.sub(r'ns\d+:', '', xml_str)
            new_font_xml.append(xml_str)

    fill_map: Dict[int, int] = {}
    new_fill_xml: List[str] = []
    n_cmpx_fills = len(cmpx_fills)
    for i, f in enumerate(tmpl_fills):
        sig = _fill_sig(f)
        if sig in cmpx_fill_lk:
            fill_map[i] = cmpx_fill_lk[sig]
        else:
            new_id = n_cmpx_fills + len(new_fill_xml)
            fill_map[i] = new_id
            cmpx_fill_lk[sig] = new_id
            xml_str = ET.tostring(f, encoding='unicode', short_empty_elements=True)
            xml_str = re.sub(r'\s*xmlns[^"]*"[^"]*"', '', xml_str)
            xml_str = re.sub(r'ns\d+:', '', xml_str)
            new_fill_xml.append(xml_str)

    border_map: Dict[int, int] = {}
    new_border_xml: List[str] = []
    n_cmpx_borders = len(cmpx_borders)
    for i, b in enumerate(tmpl_borders):
        sig = _border_sig(b)
        if sig in cmpx_border_lk:
            border_map[i] = cmpx_border_lk[sig]
        else:
            new_id = n_cmpx_borders + len(new_border_xml)
            border_map[i] = new_id
            cmpx_border_lk[sig] = new_id
            xml_str = ET.tostring(b, encoding='unicode', short_empty_elements=True)
            xml_str = re.sub(r'\s*xmlns[^"]*"[^"]*"', '', xml_str)
            xml_str = re.sub(r'ns\d+:', '', xml_str)
            new_border_xml.append(xml_str)

    log.info(f"  Style bridge: +{len(new_font_xml)} fonts, "
             f"+{len(new_fill_xml)} fills, +{len(new_border_xml)} borders")

    # Extract template xf entries
    xf_re = re.compile(r'(<xf\b[^>]*(?:/>|>(?:(?!</?xf).)*</xf>))', re.DOTALL)
    tmpl_xfs = xf_re.findall(
        tmpl_styles_raw[tmpl_styles_raw.index('<cellXfs'):
                        tmpl_styles_raw.index('</cellXfs>')])
    cmpx_xfs = xf_re.findall(
        cmpx_styles_raw[cmpx_styles_raw.index('<cellXfs'):
                        cmpx_styles_raw.index('</cellXfs>')])

    # Build CMPX xf semantic lookup (using identity maps for CMPX's own IDs)
    def _xf_key(xf_str, fm, flm, bm):
        nf = re.search(r'numFmtId="(\d+)"', xf_str)
        fi = re.search(r'fontId="(\d+)"', xf_str)
        fl = re.search(r'fillId="(\d+)"', xf_str)
        bi = re.search(r'borderId="(\d+)"', xf_str)
        ha = re.search(r'horizontal="([^"]*)"', xf_str)
        va = re.search(r'vertical="([^"]*)"', xf_str)
        qp = 1 if 'quotePrefix="1"' in xf_str else 0
        prot = 1 if '<protection' in xf_str else 0
        return (
            nf.group(1) if nf else "0",
            fm.get(int(fi.group(1)), 0) if fi else 0,
            flm.get(int(fl.group(1)), 0) if fl else 0,
            bm.get(int(bi.group(1)), 0) if bi else 0,
            ha.group(1) if ha else "",
            va.group(1) if va else "",
            qp, prot,
        )

    identity = {i: i for i in range(200)}
    cmpx_xf_lk: Dict[tuple, int] = {}
    for i, xf in enumerate(cmpx_xfs):
        key = _xf_key(xf, identity, identity, identity)
        if key not in cmpx_xf_lk:
            cmpx_xf_lk[key] = i

    # Map template style IDs → CMPX (finding existing or creating new xf)
    style_map: Dict[int, int] = {}
    new_xf_xml: List[str] = []
    n_cmpx_xfs = len(cmpx_xfs)

    for sid in sorted(needed_style_ids):
        if sid >= len(tmpl_xfs):
            continue
        key = _xf_key(tmpl_xfs[sid], font_map, fill_map, border_map)
        if key in cmpx_xf_lk:
            style_map[sid] = cmpx_xf_lk[key]
        else:
            new_id = n_cmpx_xfs + len(new_xf_xml)
            style_map[sid] = new_id
            cmpx_xf_lk[key] = new_id
            # Remap fontId/fillId/borderId in template xf
            xf_str = tmpl_xfs[sid]
            fi_m = re.search(r'fontId="(\d+)"', xf_str)
            fl_m = re.search(r'fillId="(\d+)"', xf_str)
            bi_m = re.search(r'borderId="(\d+)"', xf_str)
            if fi_m:
                xf_str = xf_str.replace(
                    f'fontId="{fi_m.group(1)}"',
                    f'fontId="{font_map[int(fi_m.group(1))]}"')
            if fl_m:
                xf_str = xf_str.replace(
                    f'fillId="{fl_m.group(1)}"',
                    f'fillId="{fill_map[int(fl_m.group(1))]}"')
            if bi_m:
                xf_str = xf_str.replace(
                    f'borderId="{bi_m.group(1)}"',
                    f'borderId="{border_map[int(bi_m.group(1))]}"')
            new_xf_xml.append(xf_str)

    log.info(f"  Style bridge: {len(style_map)} styles mapped, "
             f"+{len(new_xf_xml)} new xf entries")

    # ── Patch CMPX styles.xml ──
    out = cmpx_styles_raw

    # Add fonts
    if new_font_xml:
        old_count = n_cmpx_fonts
        new_count = old_count + len(new_font_xml)
        out = re.sub(r'<fonts\s+count="\d+"', f'<fonts count="{new_count}"', out)
        out = out.replace('</fonts>', ''.join(new_font_xml) + '</fonts>')

    # Add fills
    if new_fill_xml:
        old_count = n_cmpx_fills
        new_count = old_count + len(new_fill_xml)
        out = re.sub(r'<fills\s+count="\d+"', f'<fills count="{new_count}"', out)
        out = out.replace('</fills>', ''.join(new_fill_xml) + '</fills>')

    # Add borders
    if new_border_xml:
        old_count = n_cmpx_borders
        new_count = old_count + len(new_border_xml)
        out = re.sub(r'<borders\s+count="\d+"', f'<borders count="{new_count}"', out)
        out = out.replace('</borders>', ''.join(new_border_xml) + '</borders>')

    # Add xf entries
    if new_xf_xml:
        new_total = n_cmpx_xfs + len(new_xf_xml)
        out = re.sub(r'<cellXfs\s+count="\d+"', f'<cellXfs count="{new_total}"', out)
        out = out.replace('</cellXfs>', ''.join(new_xf_xml) + '</cellXfs>')

    return out, style_map


def _build_shared_strings_table(tmpl_path: Path) -> Dict[int, str]:
    """Build index→text lookup from template's sharedStrings.xml."""
    with zipfile.ZipFile(tmpl_path) as zf:
        if 'xl/sharedStrings.xml' not in zf.namelist():
            return {}
        ss_xml = zf.read('xl/sharedStrings.xml').decode('utf-8')

    table: Dict[int, str] = {}
    idx = 0
    # Match <si> entries — text in <t> elements (possibly within <r> runs)
    for si_m in re.finditer(r'<si>(.*?)</si>', ss_xml, re.DOTALL):
        inner = si_m.group(1)
        # Concatenate all <t>...</t> within this <si>
        parts = re.findall(r'<t[^>]*>([^<]*)</t>', inner)
        text = ''.join(parts)
        # Unescape XML entities
        text = (text.replace('&amp;', '&').replace('&lt;', '<')
                    .replace('&gt;', '>').replace('&quot;', '"'))
        table[idx] = text
        idx += 1

    log.info(f"  Shared strings: {len(table)} entries loaded from template")
    return table


def copy_template_v1_rows(
    tmpl_path: Path, cmpx_styles_xml: str, cmpx_sheet_xml: str,
) -> Tuple[str, str]:
    """Copy v1 rows (4-201) from template into CMPX Peer Views sheet.

    1. Build shared strings table from template
    2. Build style bridge (font/fill/border/xf mapping)
    3. Extract template rows 4-201, convert shared strings → inlineStr
    4. Remap style IDs using the bridge
    5. Replace CMPX rows ≤202 with converted template rows

    Returns: (modified_cmpx_styles_xml, modified_cmpx_sheet_xml)
    """
    # Find template Peer Views sheet path
    tmpl_pvs_zip = _get_sheet_zip_path(tmpl_path, "Peer Views")
    if not tmpl_pvs_zip:
        log.error("Cannot find 'Peer Views' sheet in template")
        return cmpx_styles_xml, cmpx_sheet_xml

    with zipfile.ZipFile(tmpl_path) as zf:
        tmpl_sheet_xml = zf.read(tmpl_pvs_zip).decode('utf-8')
        tmpl_styles_xml = zf.read('xl/styles.xml').decode('utf-8')

    # Step 1: Shared strings table
    ss_table = _build_shared_strings_table(tmpl_path)

    # Step 2: Collect template style IDs used in rows 4-201
    needed_styles: set = set()
    for row_m in re.finditer(
            r'<row\s+r="(\d+)"[^>]*>(.*?)</row>',
            tmpl_sheet_xml, re.DOTALL):
        rnum = int(row_m.group(1))
        if rnum <= 201:
            for cell_m in re.finditer(r's="(\d+)"', row_m.group(2)):
                needed_styles.add(int(cell_m.group(1)))
            # Also check row-level style
            row_tag = re.match(r'<row\s+[^>]*', tmpl_sheet_xml[row_m.start():])
            if row_tag:
                rs = re.search(r'\bs="(\d+)"', row_tag.group())
                if rs:
                    needed_styles.add(int(rs.group(1)))

    log.info(f"  Template v1 rows use {len(needed_styles)} unique styles")

    # Step 3: Build style bridge
    new_styles_xml, style_map = _build_style_bridge(
        tmpl_styles_xml, cmpx_styles_xml, needed_styles)

    # Step 4: Extract and convert template rows
    converted_rows: List[str] = []
    tmpl_row_re = re.compile(
        r'(<row\s+r="(\d+)"[^>]*>)(.*?)(</row>)', re.DOTALL)

    for m in tmpl_row_re.finditer(tmpl_sheet_xml):
        rnum = int(m.group(2))
        if rnum > 201:
            continue

        row_open = m.group(1)
        row_body = m.group(3)
        row_close = m.group(4)

        # Remap row-level style if present
        row_open = re.sub(
            r'\bs="(\d+)"',
            lambda sm: f's="{style_map.get(int(sm.group(1)), sm.group(1))}"',
            row_open)

        # Process each cell in the row
        new_body = _convert_row_cells(row_body, rnum, ss_table, style_map)
        converted_rows.append(row_open + new_body + row_close)

    log.info(f"  Converted {len(converted_rows)} template rows")

    # Step 5: Remove CMPX rows ≤202, insert converted rows
    # Find sheetData boundaries
    sd_start_m = re.search(r'<sheetData[^>]*>', cmpx_sheet_xml)
    sd_end = cmpx_sheet_xml.find('</sheetData>')
    if not sd_start_m or sd_end == -1:
        log.error("Cannot find <sheetData> in CMPX sheet")
        return new_styles_xml, cmpx_sheet_xml

    sd_start = sd_start_m.end()
    sheet_body = cmpx_sheet_xml[sd_start:sd_end]

    # Split: keep rows > 202, remove rows ≤ 202
    kept_rows: List[str] = []
    for row_m in re.finditer(
            r'(<row\s+r="(\d+)"[^>]*>.*?</row>|<row\s+r="(\d+)"[^/]*/>\s*)',
            sheet_body, re.DOTALL):
        rnum_str = row_m.group(2) or row_m.group(3)
        if rnum_str and int(rnum_str) > 202:
            kept_rows.append(row_m.group())

    # Assemble new sheet body
    new_body = '\n'.join(converted_rows) + '\n' + '\n'.join(kept_rows)
    new_sheet_xml = (cmpx_sheet_xml[:sd_start] + '\n' + new_body + '\n'
                     + cmpx_sheet_xml[sd_end:])

    # Update dimension ref
    new_sheet_xml = re.sub(
        r'<dimension ref="[^"]*"/>',
        lambda dm: dm.group().replace(
            dm.group(),
            f'<dimension ref="A4:AN{max(201, _max_row(kept_rows))}"/>'),
        new_sheet_xml)

    log.info(f"  CMPX sheet: {len(converted_rows)} template rows + "
             f"{len(kept_rows)} existing rows (>202)")
    return new_styles_xml, new_sheet_xml


def _max_row(rows: List[str]) -> int:
    """Find the max row number from a list of row XML strings."""
    mx = 0
    for r in rows:
        m = re.search(r'r="(\d+)"', r)
        if m:
            mx = max(mx, int(m.group(1)))
    return mx


def _convert_row_cells(
    row_body: str, row_num: int,
    ss_table: Dict[int, str],
    style_map: Dict[int, int],
) -> str:
    """Convert cells in a template row: shared strings → inlineStr, remap styles."""
    result = []
    last_end = 0

    # Non-greedy [^>]*? prevents capturing '/' from self-closing '/>' tags
    cell_re = re.compile(
        r'(<c\b[^>]*?)(/>|>(.*?)</c>)', re.DOTALL)

    for m in cell_re.finditer(row_body):
        # Copy any text between cells
        result.append(row_body[last_end:m.start()])
        last_end = m.end()

        open_tag = m.group(1)
        rest = m.group(2)
        inner = m.group(3) or ""

        # Remap style attribute
        open_tag = re.sub(
            r's="(\d+)"',
            lambda sm: f's="{style_map.get(int(sm.group(1)), sm.group(1))}"',
            open_tag)

        # Check if this is a shared string cell (t="s")
        is_ss = 't="s"' in open_tag
        if is_ss:
            # Extract shared string index from <v>N</v>
            v_m = re.search(r'<v>(\d+)</v>', inner)
            if v_m:
                ss_idx = int(v_m.group(1))
                text = ss_table.get(ss_idx, "")
                escaped = _xml_escape(text)
                # Convert to inlineStr
                open_tag = open_tag.replace('t="s"', 't="inlineStr"')
                result.append(f'{open_tag}><is><t>{escaped}</t></is></c>')
                continue

        # Non-shared-string cell: keep as-is (with remapped style)
        result.append(open_tag + rest)

    # Append any trailing content
    result.append(row_body[last_end:])
    return ''.join(result)


# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESSING — v4 SECTIONS (rows 273-486)
# ═══════════════════════════════════════════════════════════════════════════════

def restyle_v4_sections(xml: str, dry_run: bool = False) -> str:
    """Restyle all v4 section cells with correct per-drug rating."""
    total = 0
    for sec in V4_SECTIONS:
        base = sec["base"]
        log.info(f"  v4: {sec['name']} (R{base})")
        for col, rating in sec["drugs"].items():
            styles = RATING_STYLES[rating]
            for offset, cell_type in V4_OFFSETS.items():
                addr = f"{col}{base + offset}"
                new_style = styles[cell_type]
                if not dry_run:
                    xml = _restyle_cell(xml, addr, new_style)
                total += 1
            log.info(f"    col {col} → {rating}")
    log.info(f"  v4 total: {total} cells restyled")
    return xml


# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESSING — v3 SECTIONS (CRC/BTC, rows 203-272)
# ═══════════════════════════════════════════════════════════════════════════════

def restyle_v3_sections(xml: str, dry_run: bool = False) -> str:
    """Restyle CRC/BTC sections + add Bloomberg tickers + X-column labels.

    Also fixes:
    - Title row black cell alignment (adds s=551 cells for all drug columns)
    - Ticker row per-column rating coloring
    """
    total = 0
    for sec in V3_SECTIONS:
        base = sec["base"]
        log.info(f"  v3: {sec['name']} (R{base})")
        drug_cols = sorted(sec["drugs"].keys(), key=_col_idx)
        last_col = drug_cols[-1] if drug_cols else "E"

        # 0. Fix title row black cell alignment: add empty s=551 cells
        #    from E to last drug column (matching v4 title row pattern)
        for col in drug_cols:
            addr = f"{col}{base}"
            found = _find_cell(xml, addr)
            if not found:
                # Cell doesn't exist → insert empty styled cell
                if not dry_run:
                    new_cell = f'<c r="{addr}" s="551"/>'
                    xml = _insert_cell(xml, base, addr, new_cell)
                log.info(f"    title black cell added: {addr}")
            else:
                # Cell exists → ensure it has black style
                if not dry_run:
                    xml = _restyle_cell(xml, addr, "551")
                log.info(f"    title black cell restyled: {addr}")

        # 1. Restyle drug columns
        for col, rating in sec["drugs"].items():
            styles = RATING_STYLES[rating]

            # Ticker row (offset +3): rating text style
            addr = f"{col}{base + 3}"
            if not dry_run:
                xml = _restyle_cell(xml, addr, styles['text'])
            total += 1

            # Drug name row (offset +4): bold font style
            addr = f"{col}{base + V3_DRUG_OFFSET}"
            if not dry_run:
                xml = _restyle_cell(xml, addr, styles['drug'])
            total += 1

            # Data rows (offset +5 to +32): text style (preserves numFmtId=0)
            for offset in range(V3_DATA_START, V3_DATA_END + 1):
                addr = f"{col}{base + offset}"
                if not dry_run:
                    xml = _restyle_cell(xml, addr, styles['text'])
                total += 1

            log.info(f"    col {col} → {rating} (ticker+drug+data)")

        # 2. Patch Bloomberg tickers (text, not style — style set above)
        tickers = sec.get("tickers", {})
        for col, ticker_text in tickers.items():
            addr = f"{col}{base + 3}"  # Ticker row at offset +3
            if not dry_run:
                style = RATING_STYLES[sec["drugs"][col]]['text']
                xml = _patch_text_cell(xml, addr, ticker_text, style)
            log.info(f"    ticker {addr} = '{ticker_text}'")

        # 3. Add X-column rating labels
        for offset, label, style in X_LABELS:
            addr = f"X{base + offset}"
            if not dry_run:
                xml = _patch_text_cell(xml, addr, label, style)
            log.info(f"    X label {addr} = '{label}' (s={style})")

    log.info(f"  v3 total: {total} cells restyled")
    return xml


# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESSING — Peer View sheet (sheet8) — compact BTC/CRC tables
# ═══════════════════════════════════════════════════════════════════════════════

# Peer View offsets: same as V4 but skip offset 2 (BBG ticker row uses s=575)
PV_OFFSETS: Dict[int, str] = {k: v for k, v in V4_OFFSETS.items() if k >= 3}


def restyle_pv_sections(xml: str, dry_run: bool = False) -> str:
    """Restyle Peer View (sheet8) BTC/CRC sections with per-drug rating colors.

    Uses offset-based approach (same as v4) for offsets 3-23.
    Also adds X-column rating labels.
    """
    total = 0
    for sec in PV_SECTIONS:
        base = sec["base"]
        log.info(f"  PV: {sec['name']} (R{base})")
        for col, rating in sec["drugs"].items():
            styles = RATING_STYLES[rating]
            for offset, cell_type in PV_OFFSETS.items():
                addr = f"{col}{base + offset}"
                new_style = styles[cell_type]
                if not dry_run:
                    xml = _restyle_cell(xml, addr, new_style)
                total += 1
            log.info(f"    col {col} → {rating}")

    # Add X-column rating labels for both sections
    for sec in PV_SECTIONS:
        base = sec["base"]
        for offset, label, style in X_LABELS:
            addr = f"X{base + offset}"
            if not dry_run:
                xml = _patch_text_cell(xml, addr, label, style)
            log.info(f"    X label {addr} = '{label}' (s={style})")

    log.info(f"  PV total: {total} cells restyled")
    return xml


# ═══════════════════════════════════════════════════════════════════════════════
#  STYLES.XML PATCHING — create missing sae_text/mkt_react per-rating styles
# ═══════════════════════════════════════════════════════════════════════════════

def _create_rating_styles(styles_xml: str) -> Tuple[str, Dict[str, str]]:
    """Create missing per-rating style variants in styles.xml.

    Creates 12 new xf entries:
    - 4 for T1/AVG sae_text and mkt_react (blue/gold fills)
    - 7 for SUSPEND (gray fillId=4): date, pct, sae_pct, n, pfs, sale,
      sae_text, mkt_react
    - 1 for NEUTRAL (no fillId=0): mkt_react (bold font)

    Returns: (patched_styles_xml, new_ids_dict)
    """
    count_m = re.search(r'<cellXfs\s+count="(\d+)"', styles_xml)
    if not count_m:
        log.error("Cannot find cellXfs count in styles.xml")
        return styles_xml, {}

    cur = int(count_m.group(1))

    # Template: standard cell (fontId=10, borderId=6, center aligned)
    _std = (
        '<xf numFmtId="{nfmt}" fontId="10" fillId="{fill}" borderId="6" '
        'applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )
    # Template: sae_text (fontId=10, borderId=6, protection)
    _sae = (
        '<xf numFmtId="9" fontId="10" fillId="{fill}" borderId="6" '
        'applyAlignment="1" applyProtection="1" pivotButton="0" quotePrefix="0" xfId="0">'
        '<alignment horizontal="center" vertical="center"/>'
        '<protection locked="0" hidden="0"/></xf>'
    )
    # Template: mkt_react (fontId=9/bold, borderId=6, numFmtId=10)
    _mkt = (
        '<xf numFmtId="10" fontId="9" fillId="{fill}" borderId="6" '
        'applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )

    # Build new xf entries in order — IDs are cur+0, cur+1, ...
    new_xfs = []
    ids = {}

    def _add(key: str, xf_str: str):
        ids[key] = str(cur + len(new_xfs))
        new_xfs.append(xf_str)

    # --- T1/AVG sae_text and mkt_react (existing 4) ---
    _add('sae_text_avg',  _sae.format(fill="17"))   # gold
    _add('sae_text_t1',   _sae.format(fill="18"))   # blue
    _add('mkt_react_avg', _mkt.format(fill="17"))    # gold
    _add('mkt_react_t1',  _mkt.format(fill="18"))    # blue

    # --- SUSPEND (fillId=4, gray) ---
    _add('susp_date',     _std.format(nfmt="15",  fill="4"))
    _add('susp_pct',      _std.format(nfmt="193", fill="4"))
    _add('susp_sae_text', _sae.format(fill="4"))
    _add('susp_n',        _std.format(nfmt="202", fill="4"))
    _add('susp_pfs',      _std.format(nfmt="203", fill="4"))
    _add('susp_sale',     _std.format(nfmt="185", fill="4"))
    _add('susp_mkt',      _mkt.format(fill="4"))

    # --- NEUTRAL (fillId=0, no fill) ---
    _add('neut_mkt',      _mkt.format(fill="0"))

    # Insert before </cellXfs>
    styles_xml = styles_xml.replace(
        '</cellXfs>',
        ''.join(new_xfs) + '</cellXfs>'
    )
    new_count = cur + len(new_xfs)
    styles_xml = re.sub(
        r'<cellXfs\s+count="\d+"',
        f'<cellXfs count="{new_count}"',
        styles_xml
    )

    log.info(f"Created {len(new_xfs)} new rating styles in styles.xml "
             f"(IDs s{cur}..s{new_count - 1})")
    return styles_xml, ids


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Restyle Peer View + Peer Views with correct per-drug BIC/T1/AVG rating colors"
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview changes without writing")
    parser.add_argument("--file", default=str(_DEFAULT_XLSX),
                        help=f"DCF Excel file (default: {_DEFAULT_XLSX})")
    args = parser.parse_args()

    xlsx = Path(args.file)
    if not xlsx.exists():
        log.error(f"File not found: {xlsx}")
        return

    log.info(f"File: {xlsx}")
    log.info(f"Mode: {'DRY RUN' if args.dry_run else 'LIVE'}")

    # Discover BOTH sheets
    pv_zip = _get_sheet_zip_path(xlsx, "Peer View")      # sheet8.xml (singular)
    pvs_zip = _get_sheet_zip_path(xlsx, "Peer Views")     # sheet20.xml (plural)

    if not pv_zip:
        log.warning("Cannot find 'Peer View' sheet — skipping")
    else:
        log.info(f"Peer View (singular): {pv_zip}")

    if not pvs_zip:
        log.warning("Cannot find 'Peer Views' sheet — skipping")
    else:
        log.info(f"Peer Views (plural): {pvs_zip}")

    if not pv_zip and not pvs_zip:
        log.error("Neither 'Peer View' nor 'Peer Views' sheet found")
        return

    # Backup
    if not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = xlsx.with_name(f"{xlsx.stem}_pre_peergen_{ts}.xlsx")
        shutil.copy2(xlsx, backup)
        log.info(f"Backup: {backup}")

    modified: Dict[str, bytes] = {}

    # ═══════════════════════════════════════════════════════════════════
    #  Step 1: Copy v1 rows from template (must run first to bridge styles)
    # ═══════════════════════════════════════════════════════════════════
    with zipfile.ZipFile(xlsx) as zf:
        styles_xml = zf.read("xl/styles.xml").decode("utf-8")

    if pvs_zip and _TEMPLATE_XLSX.exists() and not args.dry_run:
        with zipfile.ZipFile(xlsx) as zf:
            pvs_xml = zf.read(pvs_zip).decode("utf-8")
        log.info(f"\nRead Peer Views: {len(pvs_xml):,} chars")

        log.info("\n--- Copying v1 rows (4-201) from template ---")
        styles_xml, pvs_xml = copy_template_v1_rows(
            _TEMPLATE_XLSX, styles_xml, pvs_xml)
    elif pvs_zip:
        with zipfile.ZipFile(xlsx) as zf:
            pvs_xml = zf.read(pvs_zip).decode("utf-8")
        log.info(f"\nRead Peer Views: {len(pvs_xml):,} chars")
        if not _TEMPLATE_XLSX.exists():
            log.warning(f"Template not found: {_TEMPLATE_XLSX} — skipping v1 copy")

    # ═══════════════════════════════════════════════════════════════════
    #  Step 2: Create per-rating styles for v3/v4/PV restyling
    # ═══════════════════════════════════════════════════════════════════
    styles_xml, new_ids = _create_rating_styles(styles_xml)
    if new_ids:
        # --- T1/AVG sae_text and mkt_react ---
        RATING_STYLES['AVG']['sae_text'] = new_ids['sae_text_avg']
        RATING_STYLES['AVG']['mkt_react'] = new_ids['mkt_react_avg']
        RATING_STYLES['T1']['sae_text'] = new_ids['sae_text_t1']
        RATING_STYLES['T1']['mkt_react'] = new_ids['mkt_react_t1']
        # --- SUSPEND: numFmtId-preserving gray styles ---
        RATING_STYLES['SUSPEND']['date'] = new_ids['susp_date']
        RATING_STYLES['SUSPEND']['pct'] = new_ids['susp_pct']
        RATING_STYLES['SUSPEND']['sae_pct'] = new_ids['susp_pct']
        RATING_STYLES['SUSPEND']['sae_text'] = new_ids['susp_sae_text']
        RATING_STYLES['SUSPEND']['n'] = new_ids['susp_n']
        RATING_STYLES['SUSPEND']['pfs'] = new_ids['susp_pfs']
        RATING_STYLES['SUSPEND']['sale'] = new_ids['susp_sale']
        RATING_STYLES['SUSPEND']['mkt_react'] = new_ids['susp_mkt']
        # --- NEUTRAL: mkt_react with no fill ---
        RATING_STYLES['NEUTRAL']['mkt_react'] = new_ids['neut_mkt']
    modified["xl/styles.xml"] = styles_xml.encode("utf-8")

    # ═══════════════════════════════════════════════════════════════════
    #  SHEET 1: "Peer Views" (sheet20.xml) — v3 + v4 sections
    #  (v1 rows already copied from template above)
    # ═══════════════════════════════════════════════════════════════════
    if pvs_zip:
        log.info("\n--- Restyling v4 sections (rows 273-486) ---")
        pvs_xml = restyle_v4_sections(pvs_xml, args.dry_run)

        log.info("\n--- Restyling v3 sections (CRC/BTC) + titles + tickers + X labels ---")
        pvs_xml = restyle_v3_sections(pvs_xml, args.dry_run)

        modified[pvs_zip] = pvs_xml.encode("utf-8")

    # ═══════════════════════════════════════════════════════════════════
    #  SHEET 2: "Peer View" (sheet8.xml) — compact BTC + CRC tables
    # ═══════════════════════════════════════════════════════════════════
    if pv_zip:
        with zipfile.ZipFile(xlsx) as zf:
            pv_xml = zf.read(pv_zip).decode("utf-8")
        log.info(f"\nRead Peer View: {len(pv_xml):,} chars")

        log.info("\n--- Restyling Peer View (sheet8) BTC/CRC ---")
        pv_xml = restyle_pv_sections(pv_xml, args.dry_run)

        modified[pv_zip] = pv_xml.encode("utf-8")

    if args.dry_run:
        log.info("\nDry run complete — no changes written.")
        return

    # ── Surgical zip patch ──
    with zipfile.ZipFile(xlsx) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        log.info("Added fullCalcOnLoad to workbook.xml")
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx.with_suffix(".~peergen.xlsx")
    with zipfile.ZipFile(xlsx, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue  # removed — references stripped below
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))

    try:
        tmp.replace(xlsx)
    except PermissionError:
        import os
        os.remove(str(xlsx))
        tmp.rename(xlsx)

    log.info(f"\nPatched → {xlsx}")

    # ── Summary ──
    v4_drugs = sum(len(s["drugs"]) for s in V4_SECTIONS)
    v3_drugs = sum(len(s["drugs"]) for s in V3_SECTIONS)
    v3_tickers = sum(len(s.get("tickers", {})) for s in V3_SECTIONS)
    pv_drugs = sum(len(s["drugs"]) for s in PV_SECTIONS)
    print(f"\n{'='*60}")
    print("Peer View + Peer Views Generation Complete")
    print(f"{'='*60}")
    if pvs_zip:
        print(f"  Peer Views (sheet20):")
        print(f"    v1 rows 4-201: copied from template")
        print(f"    v4 sections: {len(V4_SECTIONS)} ({v4_drugs} drug columns)")
        print(f"    v3 sections: {len(V3_SECTIONS)} ({v3_drugs} drug columns)")
        print(f"    Bloomberg tickers: {v3_tickers}")
        print(f"    X-column labels: {len(V3_SECTIONS) * len(X_LABELS)}")
    if pv_zip:
        print(f"  Peer View (sheet8):")
        print(f"    sections: {len(PV_SECTIONS)} ({pv_drugs} drug columns)")
    print(f"  File: {xlsx}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
