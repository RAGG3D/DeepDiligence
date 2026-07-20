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
import csv
import json
import logging
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

# Reuse parsing from generate_scenarios.py
from generate.generate_scenarios import (
    parse_gemini_reports, PipelineAsset,
    _asset_full_name, _xml_escape,
)
from core.model_assumptions import filter_assets_by_assumptions
# Datastore TAM: inlined directly into the Pipeline revenue formulas so the
# delivered model carries no visible TAM rows and no TAM Solid/Blood tabs.
from generate.wire_tam import (
    load_datastore_tam,
    interp as _tam_interp,
    upsert_report_tam_to_datastore,
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

# TAM Solid Parameters: growth factor rows (legacy row ids used as curve keys)
GROWTH_ROW = {'AVG': 551, 'BIC': 552, 'T1': 553}

_PLATEAU = [1.056] * 21
MATURITY_CURVES = {
    551: [0.193, 0.332, 0.423, 0.541, 0.669, 0.801, 0.913, 1.0] + _PLATEAU,   # AVG
    552: [0.003, 0.078, 0.175, 0.342, 0.586, 0.795, 0.921, 1.0] + _PLATEAU,   # BIC
    553: [0.046, 0.129, 0.233, 0.410, 0.554, 0.777, 0.930, 1.0] + _PLATEAU,   # T1
}
COGS_RATE = 0.30

# Indications that must source TAM from TAM Blood rather than TAM Solid+MM.
# Keep this explicit: the workbook has separate solid/blood data-center tabs,
# and model outputs should call the right tab even when a new indication was
# appended to both tabs by older workflow runs.
_BLOOD_INDICATIONS = {
    "AML", "MDS", "ALL", "CML", "CLL", "CMML", "MPN", "MF", "MDS/MPN",
    "MM", "HL", "cHL", "NHL", "DLBCL", "FL", "MCL", "MZL", "WM",
}

# Cross-sheet SUMIF overrides for unusual abbreviations. Values are "solid" or
# "blood"; general solid/blood resolution is handled by _resolve_tam_source().
_TAM_CROSS_REF: dict = {}

_TAM_BLOOD_COL_OFFSET = 1  # TAM Blood column = Pipeline column + 1 (for same year)

# Aliases: Scenarios indication name → Peer Views section abbreviation
# Handles cases where Scenarios uses different names than Peer Views sections.
_INDICATION_ALIASES = {
    "SCLC": "ES-SCLC",
    "cHL": "HL",
    "GC/GEJ": "GC",
    "MEL": "Melanoma",
}


def _canonical_indication(indication: str) -> str:
    return _INDICATION_ALIASES.get(indication, indication)


def _resolve_tam_source(indication: str) -> str:
    override = _TAM_CROSS_REF.get(indication)
    if override:
        return override
    canonical = _canonical_indication(indication)
    if canonical in _BLOOD_INDICATIONS or indication in _BLOOD_INDICATIONS:
        return "blood"
    return "solid"

# Style IDs (from existing template analysis)
S = {
    'drug_a':     '79',    # A col: "X" marker for drug/TAM/price rows
    'drug_d':     '73',    # D col: drug name (inlineStr)
    'drug_stage': '333',   # S-AH: SUMIFS stage formula, "Stage x" format
    'drug_hist':  '332',   # F-R: empty stage cells
    'drug_e':     '331',   # E col: empty for drug row
    'tam_c':      '11',    # C col: visible indication label (white bg, blue text)
    'tam_d':      '334',   # D col: TAM formula
    'tam_e':      '316',   # E col: "[Patients]" or "[MM USD]"
    'tam_data':   '336',   # Data cols: TAM values/formulas
    'ms_a':       '2',     # A col: blank for MS rows
    'ms_c':       '11',    # C col: rating label (white bg, blue text)
    'ms_d':       '90',    # D col: MS formula
    'ms_data':    '404',   # Data cols: SUMIFS for MS, integer percent, NO fill
                           #  (xf404 == xf337 numFmt/font/align but fillId=0 so the
                           #   colorScale colors the MS cells instead of the drug-header
                           #   theme8 light-blue of xf337)
    'price_d':    '334',   # D col: price formula
    'price_data': '193',   # Data cols: price values
    'rev_a':      '338',   # A col for Revenue row
    'rev_d':      '54',    # D col: revenue formula
    'rev_e':      '325',   # E col: "[MM USD]" for revenue
    'rev_data':   '340',   # Data cols: revenue formula, bold
    'cogs_a':     '341',   # A col for COGS row
    'cogs_d':     '54',    # D col: COGS formula
    'cogs_data':  '340',   # Data cols: COGS formula, bold
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

    # The same drug repeats across the Absolute/Base/Bull/Bear/Breakdown/Catalyst
    # scenario blocks. Read each drug ONCE from its first (lowest-row = Absolute)
    # occurrence — a first-seen-prefix de-dup that works for any pipeline size
    # (the old `row > 50` cap silently dropped drugs once a big pipeline pushed
    # the Absolute block past row 50).
    for row_str, val in sorted(asset_cells, key=lambda rv: int(rv[0])):
        if '(' not in val:
            continue  # Skip non-asset rows like "Base", "Bull", "Bear"
        prefix = val.split('(')[0].strip()
        if prefix in result:
            continue  # already captured from the Absolute block
        assets_by_row[row_str] = (prefix, val)
        result[prefix] = (val, [])
        log.info(f"  Scenarios asset C{row_str}: {val}")

    # Parse market share rows to extract indications
    # No row cap: `ref_row in assets_by_row` already restricts to the Absolute
    # block's drug rows, so market-share rows in the other scenario blocks (which
    # reference their own rows) are ignored without dropping any indication.
    for row_str, formula in formula_cells:
        # Formula: C10&amp;" BTC Market Share"
        m = re.match(r'C(\d+)&amp;" (.+) Market Share"', formula)
        if m:
            ref_row = m.group(1)
            indication = m.group(2)
            if ref_row in assets_by_row:
                prefix = assets_by_row[ref_row][0]
                if prefix in result and indication not in result[prefix][1]:
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


def _read_peer_views_ratings(xlsx_path: Path, ticker: str = "") -> Dict[str, str]:
    """Read approved peer ratings from the DD data center export.

    Ratings are explicit data in datastore/export/peer_rating.csv.  This avoids
    deriving ratings from worksheet fill colors or from ticker-specific Peer
    Views cells that may be overwritten during model generation.

    Returns: {indication_abbrev: "BIC"|"T1"|"AVG"} for rows whose peer ticker
    matches the current ticker.  User-approved asset-level assumptions still
    take precedence in build_drug_block().
    """
    export_path = Path(__file__).resolve().parents[1] / "datastore" / "export" / "peer_rating.csv"
    if not export_path.exists():
        log.warning(f"Peer rating datastore export not found: {export_path}")
        return {}

    def normalize_ticker(value: str) -> str:
        v = (value or "").upper().strip()
        v = re.sub(r"\s+EQUITY$", "", v)
        v = re.sub(r"\s+US$", "", v)
        return v

    def section_to_indication(section: str) -> Optional[str]:
        text = section or ""
        mapping = [
            ("NSCLC", "NSCLC"), ("Non-Small Cell", "NSCLC"),
            ("Triple-Negative", "TNBC"), ("TNBC", "TNBC"),
            ("Biliary", "BTC"), ("BTC", "BTC"), ("mUC", "BLCA"),
            ("Colorectal", "CRC"), ("CRC", "CRC"),
            ("Renal", "RCC"), ("RCC", "RCC"),
            ("Hepatocellular", "HCC"), ("HCC", "HCC"),
            ("Endometrial", "EC"),
            ("Melanoma", "Melanoma"),
            ("Mesothelioma", "MPM"), ("MPM", "MPM"),
            ("Hodgkin", "HL"), ("cHL", "HL"),
            ("SCLC", "ES-SCLC"), ("Small Cell", "ES-SCLC"),
            ("Gastric", "GC"), ("Gastroesophageal", "GC"),
            ("Multiple Myeloma", "MM"),
        ]
        for keyword, abbrev in mapping:
            if keyword in text:
                if abbrev == "ES-SCLC" and "NSCLC" in text:
                    continue
                return abbrev
        return None

    target = normalize_ticker(ticker)
    result: Dict[str, str] = {}
    with export_path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if target and normalize_ticker(row.get("ticker", "")) != target:
                continue
            rating = _normalize_rating(row.get("rating", ""))
            if not rating:
                continue
            indication = section_to_indication(row.get("section", ""))
            if not indication:
                continue
            result[indication] = rating

    if result:
        log.info(f"  Peer ratings from datastore for {ticker}: {result}")
    else:
        log.info(f"  No datastore peer ratings for {ticker}; using approved assumptions/fallbacks")
    return result


def _normalize_rating(value: str) -> Optional[str]:
    if not value:
        return None
    v = str(value).strip().lower().replace("-", " ")
    if v in {"bic", "best in class", "bestinclass"} or "best" in v:
        return "BIC"
    if v in {"t1", "tier one", "tier 1"} or "tier one" in v or "tier 1" in v or "above average" in v:
        return "T1"
    if v in {"avg", "average", "average growth"} or "average" in v:
        return "AVG"
    return None


def parse_report_ratings(report_dir: Path, ticker: str) -> Dict[str, Dict[str, str]]:
    """Parse per-indication competitive tier from research reports.

    This is deliberately a fallback after approved assumptions/datastore.  It
    prevents new-ticker builds from defaulting every fresh drug to AVG when the
    report already contains a line-matched differentiation assessment.
    """
    out: Dict[str, Dict[str, str]] = {}
    for path in sorted(report_dir.glob(f"{ticker}_*_research_*.md")) + sorted(report_dir.glob(f"{ticker.upper()}_*_research_*.md")):
        parts = path.stem.split("_")
        drug_parts = []
        for part in parts[1:]:
            if part.lower() == "research":
                break
            drug_parts.append(part)
        if not drug_parts:
            continue
        drug = "-".join(drug_parts)
        text = path.read_text(encoding="utf-8", errors="ignore")
        ch3 = re.search(
            r'(?:^|\n)#+ *\*{0,2}\s*(?:Chapter 3|CHAPTER 3)[^\n]*\n(.*?)(?=\n#+ *\*{0,2}\s*(?:Chapter 4|CHAPTER 4)|\Z)',
            text,
            re.S | re.I,
        )
        scope = ch3.group(1) if ch3 else text
        sections = re.split(r'(?=^##+ *\*{0,2}\s*3\.\d+\s)', scope, flags=re.M)
        for section in sections:
            hm = re.match(r'##+ *\*{0,2}\s*3\.\d+\s+(.+?)(?:\n|$)', section)
            if not hm:
                continue
            header = hm.group(1).strip().strip("*").strip()
            im = re.search(r'\(([A-Za-z][A-Z0-9a-z/+\-]+)\)', header)
            indication = im.group(1).strip() if im else re.split(r'\s+[-–—]\s+', header)[0].strip()
            # Prefer explicit assessment/rating lines; fall back to the
            # differentiation subsection body.
            rating_scope = section
            am = re.search(
                r'(?:Assessment|Rating|Uptake Tier)\s*[:\-]\s*([^\n|]+)',
                section,
                re.I,
            )
            if am:
                rating_scope = am.group(1)
            rating = _normalize_rating(rating_scope)
            if not rating and re.search(r'\babove[- ]average\b|\bstrong\b|\bdifferentiated\b', rating_scope, re.I):
                rating = "T1"
            if rating:
                out.setdefault(drug, {})[indication] = rating
                log.info(f"  Report rating {drug}/{indication}: {rating}")
    return out


def parse_report_economic_share(report_dir: Path, ticker: str) -> Dict[str, float]:
    """Parse the per-drug 'Economic share: NN%' field from research reports.

    The per-drug prompts (gemini/opus) now emit a machine-parseable
    ``Economic share: NN%`` line in Chapter 1 for partnered/collaboration
    programs (req2).  This promotes it into the model so a partnered asset is
    booked at its net economic share instead of defaulting to 100%.  Fallback
    only: an explicit assumptions file / hand value always wins.  Returns
    {drug_name: fraction in (0, 1]}.
    """
    out: Dict[str, float] = {}
    globs = (sorted(report_dir.glob(f"{ticker}_*_research_*.md"))
             + sorted(report_dir.glob(f"{ticker.upper()}_*_research_*.md")))
    for path in globs:
        parts = path.stem.split("_")
        drug_parts: List[str] = []
        for part in parts[1:]:
            if part.lower() == "research":
                break
            drug_parts.append(part)
        if not drug_parts:
            continue
        drug = "-".join(drug_parts)
        text = path.read_text(encoding="utf-8", errors="ignore")
        m = re.search(r'Economic share\s*[:\-]\s*(\d+(?:\.\d+)?)\s*%', text, re.I)
        if not m:
            continue
        try:
            val = float(m.group(1))
        except ValueError:
            continue
        if val > 1:
            val /= 100.0
        if 0 < val <= 1:
            out[drug] = val
            log.info(f"  Report economic share {drug}: {val:.0%}")
    return out


def _load_model_assumptions(report_dir: Path, ticker: str,
                            explicit_path: Optional[Path] = None) -> Dict:
    """Load approved GPT/user model assumptions if present.

    Expected JSON:
      {
        "economic_share": {"MP0712": 0.5},
        "ratings": {"MP0712": {"ES-SCLC": "T1"}}
      }
    Ratings may be BIC/T1/AVG or full labels.
    """
    candidates: List[Path] = []
    if explicit_path:
        candidates.append(explicit_path)
    candidates.extend([
        report_dir / f"{ticker}_model_assumptions.json",
        report_dir / f"{ticker.upper()}_model_assumptions.json",
        report_dir.parent / f"{ticker}_model_assumptions.json",
        report_dir.parent / f"{ticker.upper()}_model_assumptions.json",
    ])

    for path in candidates:
        if not path or not path.exists():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            log.warning(f"Could not read assumptions file {path}: {exc}")
            continue

        ratings = data.get("ratings") or {}
        norm_ratings: Dict[str, Dict[str, str]] = {}
        for drug, ind_map in ratings.items():
            if not isinstance(ind_map, dict):
                continue
            norm_ratings[str(drug)] = {}
            for ind, rating in ind_map.items():
                r = _normalize_rating(str(rating))
                if r:
                    norm_ratings[str(drug)][str(ind)] = r
        data["ratings"] = norm_ratings

        econ = {}
        for drug, share in (data.get("economic_share") or {}).items():
            try:
                val = float(share)
                if val > 1:
                    val /= 100.0
                econ[str(drug)] = val
            except Exception:
                log.warning(f"Ignoring invalid economic share for {drug}: {share}")
        data["economic_share"] = econ
        log.info(f"Loaded model assumptions: {path}")
        return data

    return {"ratings": {}, "economic_share": {}}


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

    Formula: SUMIF('TAM Solid'!$D$9:$D${max}, "indication", 'TAM Solid'!{col}$9:{col}${max})
    max_tam_row is normally the detected TAM-sheet dimension (from
    _detect_tam_sheets); the 562 default is only a defensive fallback for
    callers that cannot supply the sheet size.
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


def _maturity_formula(growth_row: int, col: str, stage_row: int, f_col: str) -> str:
    curve = MATURITY_CURVES.get(growth_row, MATURITY_CURVES[GROWTH_ROW["AVG"]])
    values = ",".join(f"{x:.6g}" for x in curve)
    idx = (
        f"MIN(COLUMN({col}1)-MATCH(5,${f_col}${stage_row}:$AH${stage_row},0)"
        f"-COLUMN($F$1)+2,29)"
    )
    return f"CHOOSE({idx},{values})"


def _build_revenue_row(row: int, drug_row: int, stage_row: int,
                       ms_rows: List[int],
                       tam_series_list: List[Dict[int, float]],
                       growth_rows: List[int],
                       economic_share: float = 1.0) -> str:
    """Build Revenue row with per-indication maturity curves.

    Revenue = IF(COUNTIF(stage_range,5)>0,
                 TAM1*MS1*MaturityFactor + TAM2*MS2*MaturityFactor + ...,
                 0)

    TAM_i is INLINED as a per-year literal from the data center and the maturity
    factor is an embedded CHOOSE constant, so the Pipeline needs no TAM rows and
    the model can delete the TAM Solid/Blood tabs without broken references.
    MS_i stays a live Market-Share cell so the forecast remains editable.
    """
    cells = [f'<row r="{row}">']
    cells.append(_ec(f"A{row}", S['rev_a']))
    formula_d = f'D{drug_row}&amp;" Revenue"'
    cells.append(_fc(f"D{row}", formula_d, S['rev_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['rev_e']))

    f_col = _year_to_col(2010)   # First data column (F)
    for year in range(2010, 2039):
        col = _year_to_col(year)

        # Build per-indication terms: TAM_i(literal) * MS_i(cell) * MaturityFactor_i
        terms = []
        for ms_r, tam_series, g_row in zip(ms_rows, tam_series_list, growth_rows):
            maturity = _maturity_formula(g_row, col, stage_row, f_col)
            tam_val = float(tam_series.get(year, 0.0))
            terms.append(f'{tam_val:.6g}*{col}{ms_r}*{maturity}')

        if len(terms) == 1:
            expr = terms[0]
        else:
            expr = '+'.join(terms)
        if economic_share != 1.0:
            expr = f'{economic_share:.6g}*({expr})'

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

    COGS = IF(COUNTIF(stage_range,5)>0, COGS_RATE * Revenue, 0).
    The COGS ratio is embedded so final workbooks can remove TAM database tabs.
    Starts immediately when stage reaches 5 (FDA approval).
    """
    cells = [f'<row r="{row}">']
    cells.append(_ec(f"A{row}", S['cogs_a']))
    formula_d = f'D{drug_row}&amp;" COGS"'
    cells.append(_fc(f"D{row}", formula_d, S['cogs_d']))
    cells.append(_tc(f"E{row}", "[MM USD]", S['rev_e']))

    f_col = _year_to_col(2010)   # F
    for year in range(2010, 2039):
        col = _year_to_col(year)
        formula = (
            f"IF(COUNTIF(${f_col}{stage_row}:{col}{stage_row},5)&gt;0,"
            f"{COGS_RATE:.6g}*{col}{rev_row},0)"
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
                     asset_rating_overrides: Optional[Dict[str, str]] = None,
                     economic_share: float = 1.0,
                     tam_solid_name: str = "TAM Solid",
                     tam_db: Optional[Dict[str, Dict[int, float]]] = None,
                     ) -> Tuple[List[str], int]:
    """Build all rows for one drug.

    Returns: (list_of_row_xml_strings, next_available_row, market_share_row_numbers)

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
    asset_ratings = asset_rating_overrides or {}
    tam_db = tam_db or {}
    ms_rows_list: List[int] = []               # one MS row per indication
    tam_series_list: List[Dict[int, float]] = []  # per-indication {year: TAM $MM}, inlined
    ind_ratings: List[str] = []  # per-indication rating strings

    for ind in indications:
        # Resolve rating for this indication
        ind_rating = asset_ratings.get(ind)
        if ind_rating is None:
            ind_rating = ratings_map.get(ind)
        if ind_rating is None:
            alias = _INDICATION_ALIASES.get(ind)
            if alias:
                ind_rating = asset_ratings.get(alias) or ratings_map.get(alias)
        if ind_rating is None:
            ind_rating = "AVG"
            log.warning(
                f"    {asset.name}/{ind}: no approved rating found; "
                "using AVG fallback"
            )

        # TAM is INLINED into the Revenue formula from the data center — the
        # Pipeline no longer shows a TAM row (TAM lives in the DB, not the model).
        # Resolve exactly like wire_tam so the inlined values are identical to
        # what the old visible TAM row would have carried.
        tam_code = _canonical_indication(str(ind).strip()).upper()
        series = tam_db.get(tam_code) or {}
        if not series:
            log.warning(
                f"    {asset.name}/{ind}: no datastore TAM code — this indication's "
                "revenue contribution is 0 (research/upsert the indication's TAM $MM)"
            )
        tam_series: Dict[int, float] = {}
        for year in range(2010, 2039):
            if series:
                v = series.get(year)
                tam_series[year] = float(v) if v is not None else float(_tam_interp(series, year))
            else:
                tam_series[year] = 0.0
        tam_series_list.append(tam_series)
        log.info(f"    {ind} TAM → inlined from data center "
                 f"(2030≈{tam_series.get(2030, 0):.0f} $MM)")

        # Market Share row — the only per-indication input row now shown
        ms_label = f"{ind} Market Share" if ind not in ("All", "All Indications Combined") else "Market Share"
        ms_row = cur
        rows.append(_build_ms_row(cur, drug_row, ms_label, rating=ind_rating))
        cur += 1

        ms_rows_list.append(ms_row)
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

    # Revenue row — TAM (inlined from DB) × MS × per-indication maturity (CHOOSE)
    rev_row = cur
    rows.append(_build_revenue_row(cur, drug_row, stage_row,
                                   ms_rows_list, tam_series_list, growth_rows,
                                   economic_share=economic_share))
    cur += 1

    # COGS row — inlined COGS ratio × Revenue (no TAM-tab reference)
    rows.append(_build_cogs_row(cur, drug_row, stage_row, rev_row,
                                tam_solid_name=tam_solid_name))
    cur += 1

    # Separator
    rows.append(_build_separator_row(cur))
    cur += 1

    log.info(f"  {asset.name}: rows {start_row}-{cur-1} "
             f"({len(indications)} ind, ratings=[{', '.join(rating_strs)}], "
             f"price=${price:.3f}MM, economic_share={economic_share:.1%})")
    return rows, cur, ms_rows_list


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


def _reanchor_ms_colorscales(xml_after: str, ms_rows: List[int]) -> str:
    """Re-anchor the Pipeline colorScale conditional formatting onto the real
    Market-Share rows.

    The template Pipeline sheet ships its colorScale <conditionalFormatting>
    blocks anchored to the old BCYC MS-row positions (S383:AH383, S390, S392,
    S399, S405, S407 — 6 blocks). Those blocks live after </sheetData>, so the
    splice at generate_pipeline() preserves them verbatim while the regenerated
    MS rows land at new positions (11, 14, 17, ...). Left as-is the colorScales
    paint blank rows 383+ and the real MS rows get no gradient.

    This regenerates exactly one colorScale block per actual MS row
    (sqref="S{ms_row}:AH{ms_row}"), reusing the template's own colorScale rule so
    the 2-color gradient is byte-identical, and drops the stale anchors. The
    B-column expression conditionalFormatting and every other trailing element
    (pageMargins/pageSetup/legacyDrawing/...) are left untouched.
    """
    if not ms_rows:
        return xml_after

    cf_re = re.compile(
        r'<conditionalFormatting\b[^>]*>.*?</conditionalFormatting>',
        re.DOTALL,
    )

    # Reuse the template's colorScale <cfRule> so the exact 2-color scale is kept.
    colorscale_rule = None
    for block in cf_re.findall(xml_after):
        if 'colorScale' in block:
            m = re.search(r'<cfRule\b.*?</cfRule>', block, re.DOTALL)
            if m:
                colorscale_rule = m.group(0)
                break
    if colorscale_rule is None:
        log.warning("No Pipeline colorScale rule found to re-anchor onto MS rows")
        return xml_after

    new_blocks = []
    for i, r in enumerate(sorted(set(ms_rows)), start=1):
        rule = re.sub(r'priority="\d+"', f'priority="{i}"', colorscale_rule, count=1)
        new_blocks.append(
            f'<conditionalFormatting sqref="S{r}:AH{r}">{rule}</conditionalFormatting>'
        )
    joined = ''.join(new_blocks)

    state = {'inserted': False}

    def _sub(m):
        block = m.group(0)
        if 'colorScale' not in block:
            return block  # preserve the B-column expression rule (and any non-colorScale CF)
        if not state['inserted']:
            state['inserted'] = True
            return joined  # first colorScale slot → all real MS-row blocks
        return ''  # drop the remaining stale colorScale anchors

    result = cf_re.sub(_sub, xml_after)
    log.info(f"Re-anchored {len(new_blocks)} colorScale block(s) onto MS rows "
             f"{sorted(set(ms_rows))}")
    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def generate_pipeline(
    xlsx_path: Path,
    assets: List[PipelineAsset],
    pricing: Dict[str, Dict[str, float]],
    ticker: str = "",
    assumptions: Optional[Dict] = None,
    tam_db: Optional[Dict[str, Dict[int, float]]] = None,
    dry_run: bool = False,
) -> None:
    """Generate Revenue Forecasting section in Pipeline sheet.

    tam_db: {INDICATION_CODE: {year: tam_usd_m}} from the data center. TAM is
    inlined directly into each Revenue formula, so the Pipeline shows only the
    revenue-forecast rows (no TAM rows) and needs no TAM Solid/Blood tabs.
    """

    sheet_zip = _get_sheet_zip_path(xlsx_path, SHEET_NAME)
    if not sheet_zip:
        log.error(f"Cannot find '{SHEET_NAME}' sheet")
        return

    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")
    # Validate the actual hardcoded styles, not a historical template-wide XF
    # count.  Cleaned delivery templates legitimately have fewer styles after
    # obsolete tabs are removed, while every Pipeline style can still be valid.
        styles_xml = zf.read("xl/styles.xml").decode("utf-8")
    xf_count_m = re.search(r'<cellXfs\s+count="(\d+)"', styles_xml)
    if xf_count_m:
        xf_count = int(xf_count_m.group(1))
        required_max = max(int(style_id) for style_id in S.values())
        if xf_count <= required_max:
            log.error(
                f"styles.xml has {xf_count} cellXf entries but Pipeline requires "
                f"style index {required_max}. Check S dict values."
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
    indication_ratings = _read_peer_views_ratings(xlsx_path, ticker)
    assumptions = assumptions or {"ratings": {}, "economic_share": {}}
    rating_overrides = assumptions.get("ratings") or {}
    economic_shares = assumptions.get("economic_share") or {}

    # ── Step 1e: Resolve TAM Solid sheet name for cross-sheet references ──
    tam_solid_name = "TAM Solid"
    if tam_sheets and "solid" in tam_sheets:
        tam_solid_name = tam_sheets["solid"][0]
    log.info(f"  TAM Solid sheet name: '{tam_solid_name}'")

    # ── Step 2: Build new drug blocks ──
    log.info(f"\nBuilding revenue forecasting for {len(assets)} drugs...")
    all_rows: List[str] = []
    rev_rows: List[int] = []  # Track revenue row numbers for sum formula
    ms_rows: List[int] = []   # Track Market Share row numbers for colorScale re-anchoring
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
        econ_share = float(economic_shares.get(asset.name, 1.0))
        asset_ratings = rating_overrides.get(asset.name, {})
        block_rows, cur, block_ms_rows = build_drug_block(
            asset, cur, drug_pricing,
            override_full_name=sc_full_name,
            override_indications=sc_indications or None,
            tam_sheets=tam_sheets,
            indication_ratings=indication_ratings,
            asset_rating_overrides=asset_ratings,
            economic_share=econ_share,
            tam_solid_name=tam_solid_name,
            tam_db=tam_db,
        )
        # Revenue row is the second-to-last row before separator
        rev_rows.append(cur - 3)  # COGS=cur-2, rev=cur-3
        ms_rows.extend(block_ms_rows)
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

    # ── Step 3b: Re-anchor Pipeline colorScale CF onto the real MS rows ──
    # The template's colorScale blocks (in xml_after) are anchored to the old
    # BCYC MS rows; repoint them to the regenerated MS rows so the 2-color
    # gradient paints the live MS cells and not stale blank rows.
    xml_after = _reanchor_ms_colorscales(xml_after, ms_rows)

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
    parser.add_argument("--assumptions-file",
                        help="Approved model assumptions JSON with ratings/economic shares")
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
    assumptions = _load_model_assumptions(
        report_dir,
        args.ticker,
        Path(args.assumptions_file) if args.assumptions_file else None,
    )
    report_ratings = parse_report_ratings(report_dir, args.ticker)
    if report_ratings:
        assumptions.setdefault("ratings", {})
        for drug, ind_map in report_ratings.items():
            assumptions["ratings"].setdefault(drug, {})
            for ind, rating in ind_map.items():
                assumptions["ratings"][drug].setdefault(ind, rating)

    # Report-derived economic share is a fallback: an explicit assumptions file
    # already parsed above wins per drug (setdefault), so partnered assets get
    # their researched net % instead of defaulting to 100% (finding 23 / req2).
    report_econ = parse_report_economic_share(report_dir, args.ticker)
    if report_econ:
        assumptions.setdefault("economic_share", {})
        for drug, share in report_econ.items():
            assumptions["economic_share"].setdefault(drug, share)

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

    assets, excluded = filter_assets_by_assumptions(assets, assumptions)
    for drug, indication in excluded:
        log.info(f"  Approved assumptions EXCLUDE {drug}/{indication} (all scenario peaks are zero)")
    if not assets:
        log.error("Approved assumptions excluded every parsed pipeline asset")
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

    # ── Step 3b: Load datastore TAM (inlined into revenue; no visible TAM rows) ──
    # Persist this ticker's researched TAM tables to the data center first, then
    # read them back scoped to the active ticker so another ticker's serviceable
    # markets can never bleed in. The values are inlined into the revenue formula.
    try:
        n_up = upsert_report_tam_to_datastore(report_dir, args.ticker)
        if n_up:
            log.info(f"Upserted {n_up} report TAM tables into datastore before pipeline build")
    except Exception as exc:
        log.warning(f"Report TAM upsert skipped ({exc}); using existing datastore TAM")
    tam_db = load_datastore_tam(args.ticker)
    log.info(f"Datastore TAM: {len(tam_db)} indication code(s) available for inlining")

    # ── Step 4: Generate ──
    log.info(f"\n{'='*60}")
    log.info("STEP 3: Generating Pipeline Revenue Forecasting")
    log.info(f"{'='*60}")

    generate_pipeline(
        xlsx_path, assets, pricing,
        ticker=args.ticker,
        assumptions=assumptions,
        tam_db=tam_db,
        dry_run=args.dry_run,
    )

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
