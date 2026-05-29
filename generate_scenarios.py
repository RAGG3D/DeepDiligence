#!/usr/bin/env python3
"""
generate_scenarios.py -- Generate Scenarios sheet from Gemini research report (FIXED VERSION)

⚠️ CRITICAL: Uses template copying + cell replacement (NOT building from scratch)
Preserves ALL template structure, namespaces, attributes.

Usage:
    python generate_scenarios.py --ticker CMPX --research-file CMPX_gemini_research_*.md
"""

import argparse
import logging
import re
import shutil
import sys
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
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
#  DATA STRUCTURES
# ══════════════════════════════════════════════════════════════════════════════

class CatalystEvent:
    """Represents a single upcoming catalyst event."""
    def __init__(self, name: str, indication: str, date: str):
        self.name = name
        self.indication = indication
        self.date = date
        self.positive_shares = {}  # {year: market_share_pct}
        self.negative_shares = {}  # {year: market_share_pct}
        self.positive_peak = 0.0
        self.negative_peak = 0.0


class PipelineAsset:
    """Represents a single pipeline asset."""
    def __init__(self, name: str, target: str, indications: List[str]):
        self.name = name
        self.target = target
        self.indications = indications
        self.stages = {}  # {stage_num: year}
        self.market_shares = {}  # {indication: {year: market_share_pct}}
        self.bull_shares = {}    # {indication: {year: market_share_pct}}
        self.bear_shares = {}    # {indication: {year: market_share_pct}}
        self.catalysts = []      # [CatalystEvent]
        self.rating = "AVG"      # "BIC", "T1", or "AVG" — determines maturity curve


# ══════════════════════════════════════════════════════════════════════════════
#  PARSE PER-DRUG GEMINI RESEARCH REPORTS (new format: one .md per drug)
# ══════════════════════════════════════════════════════════════════════════════

# Phase name → stage number mapping
_PHASE_TO_STAGE = {
    "phase 1": 1, "phase i": 1, "ph 1": 1, "ph1": 1,
    "phase 2": 2, "phase ii": 2, "ph 2": 2, "ph2": 2,
    "phase 2/3": 3, "phase ii/iii": 3, "ph 2/3": 3,
    "phase 3": 3, "phase iii": 3, "ph 3": 3, "ph3": 3,
    "bla": 4, "bla filing": 4, "nda": 4, "nda filing": 4, "filing": 4,
    "approval": 5, "approved": 5, "launch": 5,
}


def parse_gemini_reports(report_dir: Path, ticker: str) -> List[PipelineAsset]:
    """Parse all per-drug Gemini Deep Research reports in a directory."""
    pattern = f"{ticker}_*_research_*.md"
    files = sorted(report_dir.glob(pattern))
    if not files:
        logger.error(f"No report files matching '{pattern}' in {report_dir}")
        return []
    logger.info(f"Found {len(files)} report files in {report_dir}")
    assets = []
    for f in files:
        asset = _parse_single_drug_report(f, ticker)
        if asset:
            assets.append(asset)
    logger.info(f"Parsed {len(assets)} pipeline assets total")
    return assets


def _parse_single_drug_report(path: Path, ticker: str) -> Optional[PipelineAsset]:
    """Parse a single per-drug research report → PipelineAsset."""
    logger.info(f"Parsing: {path.name}")
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # ── Extract drug name from filename: {TICKER}_{DRUG}_research_*.md ──
    fname = path.stem  # e.g. CMPX_CTX-009_research_20260228_162134
    parts = fname.split('_')
    # Skip ticker prefix, take parts until "research"
    drug_parts = []
    for p in parts[1:]:
        if p.lower() == 'research':
            break
        drug_parts.append(p)
    drug_name = '-'.join(drug_parts) if drug_parts else parts[1] if len(parts) > 1 else "Unknown"
    # Handle underscore-separated drug names like CTX_009 → CTX-009
    if not drug_parts:
        drug_name = "Unknown"
    logger.info(f"  Drug: {drug_name}")

    # ── Extract target from Chapter 1 ──
    target = _extract_target(content)
    logger.info(f"  Target: {target}")

    # ── Parse Chapter 3: per-indication market shares ──
    # Handle all header variants: "# Chapter 3:", "# **Chapter 3:**", "# CHAPTER 3:"
    ch3_match = re.search(
        r'(?:^|\n)#+ *\*{0,2}\s*(?:Chapter 3|CHAPTER 3)[^\n]*\n(.*?)(?=\n#+ *\*{0,2}\s*(?:Chapter 4|CHAPTER 4)|\Z)',
        content, re.DOTALL | re.IGNORECASE
    )
    indications = []
    market_shares = {}
    if ch3_match:
        ch3_text = ch3_match.group(1)
        indications, market_shares = _parse_chapter3_market_shares(ch3_text)
        logger.info(f"  Indications: {indications}")
    else:
        logger.warning(f"  No Chapter 3 found in {path.name}")

    # ── Parse bull/bear market shares from Chapter 3 subsections ──
    bull_shares = {}
    bear_shares = {}
    if ch3_match:
        ch3_text = ch3_match.group(1)
        bull_shares = _parse_scenario_shares(ch3_text, "BULL")
        bear_shares = _parse_scenario_shares(ch3_text, "BEAR")
        if bull_shares:
            logger.info(f"  Bull peaks: {_peaks_summary(bull_shares)}")
        if bear_shares:
            logger.info(f"  Bear peaks: {_peaks_summary(bear_shares)}")

    # ── Parse Chapter 4: stage timeline ──
    ch4_match = re.search(
        r'(?:^|\n)#+ *\*{0,2}\s*(?:Chapter 4|CHAPTER 4)[^\n]*\n(.*?)(?=\n#+ *\*{0,2}\s*(?:Chapter 5|CHAPTER 5)|\Z)',
        content, re.DOTALL | re.IGNORECASE
    )
    stages = {}
    if ch4_match:
        stages = _parse_stage_timeline(ch4_match.group(1))
        logger.info(f"  Stages: {stages}")
    else:
        logger.warning(f"  No Chapter 4 found in {path.name}")

    # ── Parse Chapter 5: catalyst events ──
    ch5_match = re.search(
        r'(?:^|\n)#+ *\*{0,2}\s*(?:Chapter 5|CHAPTER 5)[^\n]*\n(.*?)(?=\n#+ *\*{0,2}\s*(?:Chapter 6|CHAPTER 6)|\Z)',
        content, re.DOTALL | re.IGNORECASE
    )
    catalysts = []
    if ch5_match:
        catalysts = _parse_catalyst_events(ch5_match.group(1), indications)
        logger.info(f"  Catalysts: {len(catalysts)} events")
        for cat in catalysts:
            logger.info(f"    {cat.name} ({cat.indication}): +"
                        f"{cat.positive_peak:.1%} / -{cat.negative_peak:.1%}")

    if not indications and not stages:
        logger.warning(f"  Skipping {drug_name}: no indications or stages found")
        return None

    asset = PipelineAsset(drug_name, target, indications)
    asset.stages = stages
    asset.market_shares = market_shares
    asset.bull_shares = bull_shares
    asset.bear_shares = bear_shares
    asset.catalysts = catalysts
    return asset


def _extract_target(content: str) -> str:
    """Extract drug target from Chapter 1 text."""
    target_text = ""

    # Try "**Target:**" or "**Targets:**" line
    m = re.search(r'\*\*Targets?:\*\*\s*(.+?)(?:\n|\*\*)', content)
    if m:
        target_text = m.group(1).strip().rstrip('.')
    else:
        # Try "**Target**: ..." format (without bold colon)
        m = re.search(r'\*\*Target\*\*:\s*(.+?)(?:\n|$)', content)
        if m:
            target_text = m.group(1).strip().rstrip('.')

    if not target_text:
        return ""

    # Abbreviate common long names BEFORE shortening
    _ABBREVS = [
        ("Programmed Cell Death Protein 1", "PD-1"),
        ("Programmed Death-Ligand 1", "PD-L1"),
        ("Vascular Endothelial Growth Factor A", "VEGF-A"),
        ("Vascular Endothelial Growth Factor", "VEGF"),
        ("Delta-like Ligand 4", "DLL4"),
        ("Tumor Necrosis Factor Receptor Superfamily", "TNFRSF"),
    ]
    for long, short in _ABBREVS:
        target_text = target_text.replace(long, short)

    # If result is still very long (e.g. "Bispecific antibody targeting PD-1 ..."),
    # try to extract just the target names from patterns like "X and Y" or "X x Y"
    if len(target_text) > 40:
        # Look for "targeting X and Y" or "targeting X x Y"
        m2 = re.search(r'targeting\s+(.+)', target_text, re.IGNORECASE)
        if m2:
            target_text = m2.group(1).strip()
        # Still long? Take first sentence
        if len(target_text) > 60:
            for sep in ['.', ';']:
                if sep in target_text:
                    target_text = target_text.split(sep)[0].strip()
                    break

    # Clean up: remove parenthesized duplicates like "DLL4 (DLL4)"
    target_text = re.sub(r'\(([A-Z0-9-]+)\)\s+and\s+\1', r'and \1', target_text)
    target_text = re.sub(r'([A-Z0-9-]+)\s+\(\1\)', r'\1', target_text)

    return target_text


def _parse_chapter3_market_shares(ch3_text: str) -> Tuple[List[str], Dict]:
    """Parse Chapter 3 per-indication sections → (indications, market_shares).

    Each indication section is ## 3.N with a market share subsection containing
    either a year-by-year table or text-only peak/launch info.
    """
    indications = []
    market_shares = {}

    # Split into indication sections: ## 3.1, ## **3.1, ### 3.1, etc.
    sections = re.split(r'(?=^##+ *\*{0,2}\s*3\.\d+\s)', ch3_text, flags=re.MULTILINE)

    for section in sections:
        if not section.strip():
            continue

        # Extract indication name from section header
        # Examples:
        #   ## 3.1 Biliary Tract Cancer (BTC) – 2nd Line (2L)
        #   ## **3.1 Biliary Tract Cancer (BTC) – 2nd Line (2L)**
        #   ### 3.1 Renal Cell Carcinoma (RCC) - 3L+ Setting
        #   ## 3.4 Malignant Melanoma (MM)
        #   ## 3.5 Head and Neck (HNSCC)
        header_match = re.match(
            r'##+ *\*{0,2}\s*3\.(\d+)\s+(.+?)(?:\n|$)', section
        )
        if not header_match:
            continue

        ind_header = header_match.group(2).strip().strip('*').strip()

        # Extract abbreviation from parentheses (e.g. BTC, NSCLC, cHL, GC/GEJ)
        abbrev_match = re.search(r'\(([A-Za-z][A-Z0-9a-z/+-]+)\)', ind_header)
        if abbrev_match:
            ind_name = abbrev_match.group(1)
        else:
            # Strip "Indication: " prefix if present, take part before spaced dash
            clean = re.sub(r'^(?:Indication:\s*)', '', ind_header).strip()
            clean = re.split(r'\s+[-–—]\s+', clean)[0].strip()
            # Remove trailing parenthesized content (e.g. "(NCAM+ Post-PD1)")
            clean = re.sub(r'\s*\([^)]*\)\s*$', '', clean).strip().strip('*').strip()
            ind_name = clean if clean else ind_header.strip('*').strip()

        logger.info(f"  Parsing indication: {ind_name}")

        # Find market share subsection (### 3.N.5 or ### 3.N.6 Market Share / Differentiation)
        # Handle bold variants: ### **3.1.5 Market Share Projection**
        ms_match = re.search(
            r'(?:###+ *\*{0,2}\s*3\.\d+\.\d+\s+)?(?:Differentiation\s*&\s*)?Market Share[^\n]*\n(.*?)(?=\n###+ *\*{0,2}\s*3\.\d+\.\d+|\n##+ *\*{0,2}\s*3\.\d+\s|\n---|\n#+ *\*{0,2}\s*(?:Chapter|CHAPTER)|\Z)',
            section, re.DOTALL | re.IGNORECASE
        )

        if not ms_match:
            # Try broader search: any table with Year | Market Share
            ms_match = re.search(
                r'(\|[^\n]*Year[^\n]*Market Share[^\n]*\n.*?)(?=\n---|\n##|\Z)',
                section, re.DOTALL | re.IGNORECASE
            )

        if ms_match:
            ms_text = ms_match.group(1) if ms_match.lastindex else ms_match.group(0)
            shares = _parse_ms_table(ms_text)

            if not shares:
                # Try text-only extraction
                shares = _parse_ms_text_only(section)

            if shares:
                shares = _interpolate_market_shares(shares)
                market_shares[ind_name] = shares
                indications.append(ind_name)
                logger.info(f"    {ind_name}: {len(shares)} year entries, peak={max(shares.values()):.1%}")
        else:
            # Try text-only extraction from the whole section
            shares = _parse_ms_text_only(section)
            if shares:
                shares = _interpolate_market_shares(shares)
                market_shares[ind_name] = shares
                indications.append(ind_name)
                logger.info(f"    {ind_name}: {len(shares)} year entries (text-only), peak={max(shares.values()):.1%}")
            else:
                logger.warning(f"    No market share data found for {ind_name}")

    return indications, market_shares


def _parse_scenario_shares(ch3_text: str, scenario: str) -> Dict[str, Dict[int, float]]:
    """Parse BULL or BEAR market share projections from Chapter 3 subsections.

    Looks for subsections like:
      ### 3.X.8 Market Share Projection — BULL Case (2024-2038)
      ### 3.X.9 Market Share Projection — BEAR Case (2024-2038)

    Returns: {indication_name: {year: share_pct}}
    """
    shares_by_ind = {}

    # Split into indication sections (## 3.1, ## 3.2, etc.)
    sections = re.split(r'(?=^##+ *\*{0,2}\s*3\.\d+\s)', ch3_text, flags=re.MULTILINE)

    for section in sections:
        if not section.strip():
            continue

        # Get indication name (same logic as base parser)
        header_match = re.match(
            r'##+ *\*{0,2}\s*3\.(\d+)\s+(.+?)(?:\n|$)', section
        )
        if not header_match:
            continue

        ind_header = header_match.group(2).strip().strip('*').strip()
        abbrev_match = re.search(r'\(([A-Za-z][A-Z0-9a-z/+-]+)\)', ind_header)
        if abbrev_match:
            ind_name = abbrev_match.group(1)
        else:
            clean = re.sub(r'^(?:Indication:\s*)', '', ind_header).strip()
            clean = re.split(r'\s+[-–—]\s+', clean)[0].strip()
            clean = re.sub(r'\s*\([^)]*\)\s*$', '', clean).strip().strip('*').strip()
            ind_name = clean if clean else ind_header.strip('*').strip()

        # Find the BULL or BEAR subsection
        pattern = (
            rf'(?:###+ *\*{{0,2}}\s*3\.\d+\.\d+\s+)?'
            rf'Market Share Projection[^\n]*{scenario}[^\n]*\n'
            rf'(.*?)'
            rf'(?=\n###+ *\*{{0,2}}\s*3\.\d+\.\d+|\n##+ *\*{{0,2}}\s*3\.\d+\s|\n---|\n#+ *\*{{0,2}}\s*(?:Chapter|CHAPTER)|\Z)'
        )
        ms_match = re.search(pattern, section, re.DOTALL | re.IGNORECASE)
        if ms_match:
            shares = _parse_ms_table(ms_match.group(1))
            if not shares:
                shares = _parse_ms_text_only(ms_match.group(1))
            if shares:
                shares = _interpolate_market_shares(shares)
                shares_by_ind[ind_name] = shares

    return shares_by_ind


def _peaks_summary(shares_by_ind: Dict[str, Dict[int, float]]) -> str:
    """One-line summary of peaks per indication."""
    parts = []
    for ind, sh in shares_by_ind.items():
        if sh:
            parts.append(f"{ind}={max(sh.values()):.1%}")
    return ", ".join(parts) if parts else "none"


def _parse_catalyst_events(ch5_text: str, base_indications: List[str]) -> List['CatalystEvent']:
    """Parse Chapter 5 catalyst events with positive/negative MS projections.

    Looks for subsections like:
      #### 5.N Per-Catalyst Analysis: [Event Name] — [Indication] [Line]
      ##### 5.N.3 Market Share Projection — Catalyst Positive (2024-2038)
      ##### 5.N.4 Market Share Projection — Catalyst Negative (2024-2038)
    """
    catalysts = []

    # Split into per-catalyst sections: #### 5.1, #### 5.2, etc.
    # But skip the calendar (#### 5.1 Catalyst Calendar) — catalysts start at #### 5.2+
    sections = re.split(
        r'(?=^#{3,5} *\*{0,2}\s*5\.\d+\s+(?:Per-Catalyst|Catalyst Analysis))',
        ch5_text, flags=re.MULTILINE | re.IGNORECASE
    )

    for section in sections:
        if not section.strip():
            continue

        # Extract catalyst name and indication from header
        header_match = re.match(
            r'#{3,5} *\*{0,2}\s*5\.(\d+)\s+(?:Per-Catalyst Analysis:\s*)?(.+?)(?:\n|$)',
            section, re.IGNORECASE
        )
        if not header_match:
            continue

        header_text = header_match.group(2).strip().strip('*').strip()

        # Try "Event Name — Indication Line" format
        parts = re.split(r'\s*[—–-]\s*', header_text, maxsplit=1)
        if len(parts) >= 2:
            cat_name = parts[0].strip()
            cat_indication = parts[1].strip()
        else:
            cat_name = header_text
            cat_indication = base_indications[0] if base_indications else "All"

        # Extract expected date from section text
        date_match = re.search(
            r'(?:Expected|Date|Timing)[:\s]*\*{0,2}([A-Z][a-z]+\s+\d{4}|\d{4}\s+[A-Z]|Q[1-4]\s+\d{4}|\d{4})',
            section, re.IGNORECASE
        )
        cat_date = date_match.group(1).strip() if date_match else ""

        cat = CatalystEvent(cat_name, cat_indication, cat_date)

        # Parse positive outcome MS table (5.N.3)
        pos_match = re.search(
            r'(?:#{4,6}[^\n]*)?Market Share Projection[^\n]*(?:Positive|Catalyst Positive)[^\n]*\n'
            r'(.*?)'
            r'(?=\n#{3,6} *\*{0,2}\s*5\.\d+|'
            r'\n#{4,6}[^\n]*Market Share Projection[^\n]*(?:Negative)|'
            r'\n---|\Z)',
            section, re.DOTALL | re.IGNORECASE
        )
        if pos_match:
            shares = _parse_ms_table(pos_match.group(1))
            if shares:
                shares = _interpolate_market_shares(shares)
                cat.positive_shares = shares
                cat.positive_peak = max(shares.values()) if shares else 0

        # Parse negative outcome MS table (5.N.4)
        neg_match = re.search(
            r'(?:#{4,6}[^\n]*)?Market Share Projection[^\n]*(?:Negative|Catalyst Negative)[^\n]*\n'
            r'(.*?)'
            r'(?=\n#{3,6} *\*{0,2}\s*5\.\d+|\n---|\Z)',
            section, re.DOTALL | re.IGNORECASE
        )
        if neg_match:
            shares = _parse_ms_table(neg_match.group(1))
            if shares:
                shares = _interpolate_market_shares(shares)
                cat.negative_shares = shares
                cat.negative_peak = max(shares.values()) if shares else 0

        # Only add if we got at least some data
        if cat.positive_peak > 0 or cat.negative_peak > 0 or cat_name:
            catalysts.append(cat)

    return catalysts


def _parse_ms_table(ms_text: str) -> Dict[int, float]:
    """Parse a market share table with Year | Market Share columns.

    Handles:
    - Standard rows: | 2027 | 2% |
    - Bold: | **2027** | **2%** |
    - Peak annotation: | 2033 | **35% (Peak)** |
    - Range years: | 2024-31 | 0% |  → apply to all years in range
    - Ellipsis rows: | ... | |  → skip
    """
    shares = {}

    # Find all table rows with year(s) and percentage
    # Pattern: | year_spec | pct_spec | ...
    rows = re.findall(
        r'\|\s*\*{0,2}(\d{4}(?:\s*-\s*\d{2,4})?)\*{0,2}\s*\|\s*\*{0,2}([\d.]+)%',
        ms_text
    )

    for year_spec, pct_str in rows:
        pct = float(pct_str) / 100.0
        year_spec = year_spec.strip()

        if '-' in year_spec:
            # Range: "2024-31" or "2024-2031"
            start_str, end_str = year_spec.split('-')
            start_year = int(start_str.strip())
            end_str = end_str.strip()
            if len(end_str) == 2:
                end_year = (start_year // 100) * 100 + int(end_str)
            else:
                end_year = int(end_str)
            for y in range(start_year, end_year + 1):
                shares[y] = pct
        else:
            shares[int(year_spec)] = pct

    return shares


def _parse_ms_text_only(section_text: str) -> Dict[int, float]:
    """Extract market share from text-only descriptions (no table).

    Looks for patterns like:
    - "Peak **3%** share"  /  "Peak Share: **N%**"  /  "Peak 20%"
    - "Launch 2032"  /  "Launch Year"  /  "launch year. **2030**"
    """
    shares = {}

    # Extract peak share
    # Handles: "Peak Share: 8%", "**Peak Share:** 8%", "Peak **3%**", "peak 20%"
    peak_match = re.search(
        r'[Pp]eak\s*(?:[Ss]hare)?[:\s]*\*{0,2}\s*(\d+(?:\.\d+)?)%?\*{0,2}',
        section_text
    )
    peak_pct = None
    if peak_match:
        peak_pct = float(peak_match.group(1)) / 100.0

    # Extract launch year
    # Handles: "Launch 2032", "2029 Launch.", "**2029 Launch.**", "Launch Year: 2030"
    launch_match = re.search(
        r'(?:\*{0,2}(\d{4})\*{0,2}\s+[Ll]aunch|[Ll]aunch\s*(?:[Yy]ear)?[.:\s]*\*{0,2}(\d{4})\*{0,2})',
        section_text
    )
    launch_year = None
    if launch_match:
        launch_year = int(launch_match.group(1) or launch_match.group(2))

    # If peak found but no launch year, default to 2032
    if peak_pct is not None and launch_year is None:
        launch_year = 2032
        logger.info(f"    Text-only: peak={peak_pct:.1%}, no launch year found → defaulting to {launch_year}")

    if peak_pct is not None and launch_year is not None:
        # Generate ramp: 0 before launch, ramp to peak over 5 years, then hold
        for y in range(2024, launch_year):
            shares[y] = 0.0
        ramp_years = 5
        for i in range(ramp_years + 1):
            y = launch_year + i
            if y <= 2038:
                shares[y] = peak_pct * i / ramp_years
        # Hold peak after ramp
        peak_year = launch_year + ramp_years
        for y in range(peak_year, 2039):
            shares[y] = peak_pct

    return shares


def _interpolate_market_shares(shares: Dict[int, float]) -> Dict[int, float]:
    """Fill in missing years 2024-2038 with linear interpolation.

    - Before first data point: 0%
    - Between data points: linear interpolation
    - After last data point: hold last value
    """
    if not shares:
        return shares

    result = {}
    sorted_years = sorted(shares.keys())
    min_year = sorted_years[0]
    max_year = sorted_years[-1]

    for y in range(2024, 2039):
        if y in shares:
            result[y] = shares[y]
        elif y < min_year:
            result[y] = 0.0
        elif y > max_year:
            result[y] = shares[max_year]
        else:
            # Find surrounding data points for interpolation
            prev_y = max(sy for sy in sorted_years if sy <= y)
            next_y = min(sy for sy in sorted_years if sy >= y)
            if prev_y == next_y:
                result[y] = shares[prev_y]
            else:
                frac = (y - prev_y) / (next_y - prev_y)
                result[y] = shares[prev_y] + frac * (shares[next_y] - shares[prev_y])

    return result


def _parse_stage_timeline(ch4_text: str) -> Dict[int, int]:
    """Parse Chapter 4 stage timeline table → {stage_num: year}.

    Handles:
    - Clean: | **1 (Phase I)** | **2026** |
    - Named phase: | **Phase 2/3 (BTC)** | **2023-2025** |  → stage 3, year 2023
    - Split years: | **4 (BLA Filing)** | 2029 (HL) / 2031 (Solid) | → year 2029
    - Range: | **1 (Phase 1)** | **2019-2025** | → year 2019
    """
    stages = {}

    # Process line by line to avoid cross-row regex issues
    for line in ch4_text.split('\n'):
        line = line.strip()
        if not line.startswith('|'):
            continue

        # Split by pipe, strip each cell
        cells = [c.strip().strip('*').strip() for c in line.split('|')]
        # Remove empty first/last elements from leading/trailing pipes
        cells = [c for c in cells if c]

        if len(cells) < 2:
            continue

        col1, col2 = cells[0], cells[1]

        # Skip header/separator rows
        if ':---' in col1 or col1.lower() in ('stage', 'stage / phase'):
            continue

        # Extract stage number
        stage_num = _extract_stage_number(col1)
        if stage_num is None:
            continue

        # Extract year (take earliest/start year)
        year = _extract_earliest_year(col2)
        if year is None:
            continue

        # Only store the first (earliest) occurrence per stage
        if stage_num not in stages:
            stages[stage_num] = year
            logger.info(f"    Stage {stage_num}: {year}")

    return stages


def _extract_stage_number(text: str) -> Optional[int]:
    """Extract stage number from column text.

    Examples:
      "1 (Phase I)"        → 1
      "Phase 2/3 (BTC)"    → 3
      "4 (BLA Filing)"     → 4
      "5 (Approval)"       → 5
      "BLA Filing (BTC)"   → 4
      "Phase 1"            → 1
      "Phase 2 (BTC)"      → 2
    """
    # Try explicit stage number first: "N (..."
    m = re.match(r'(\d)\s*\(', text)
    if m:
        return int(m.group(1))

    # Try "Stage N"
    m = re.search(r'Stage\s+(\d)', text, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # Try phase name mapping — match longest first to prefer "phase 2/3" over "phase 2"
    text_lower = text.lower().strip()
    # Sort by key length descending so "phase 2/3" matches before "phase 2"
    for phase_name, stage_num in sorted(_PHASE_TO_STAGE.items(), key=lambda x: -len(x[0])):
        if phase_name in text_lower:
            return stage_num

    return None


def _extract_earliest_year(text: str) -> Optional[int]:
    """Extract the earliest year from text.

    Handles: "2026", "2023-2025", "2029 (HL) / 2031 (Solid)", "**2026**"
    """
    # Strip bold markers
    text = text.replace('*', '')

    # Find all 4-digit years (2019-2038 range)
    years = [int(y) for y in re.findall(r'20[1-3]\d', text)]

    # Also handle 2-digit end of range: "2024-26" → 2024
    # The start year is already captured above

    return min(years) if years else None


# ── Legacy single-file parser (kept as fallback) ──

def parse_gemini_report(report_path: Path) -> List[PipelineAsset]:
    """Parse old-format single-file Gemini research report (legacy fallback)."""
    with open(report_path, 'r', encoding='utf-8') as f:
        content = f.read()

    assets = []

    part2_section = re.search(
        r'## Part 2:.*?Market Share Projections(.*?)(?=## Part 3|$)',
        content, re.DOTALL | re.IGNORECASE
    )
    if not part2_section:
        logger.warning("Could not find Part 2 in research report")
        return assets

    part2_text = part2_section.group(1)
    p2_asset_blocks = re.split(r'(?=^###\s)', part2_text, flags=re.MULTILINE)
    p2_asset_blocks = [b for b in p2_asset_blocks if b.strip()]

    p2_assets = {}
    for block in p2_asset_blocks:
        m = re.match(
            r'###\s*\*{0,2}(?:Asset\s*\d*:\s*)(\S+)\s*(?:\(([^)]*)\))?\*{0,2}',
            block.strip()
        )
        if not m:
            continue
        asset_name = m.group(1).strip()
        description = (m.group(2) or "").strip()
        p2_assets[asset_name] = (description, block)

    stage_section = re.search(
        r'## Part 3:.*?Stage Timeline Predictions(.*)',
        content, re.DOTALL | re.IGNORECASE
    )
    if not stage_section:
        logger.warning("Could not find Part 3 in research report")
        return assets

    stage_text = stage_section.group(1)
    p3_asset_blocks = re.findall(
        r'###\s*Asset:\s*([^\n(]+?)(?:\s*\([^)]*\))?\s*\n(.*?)(?=###\s*Asset:|\Z)',
        stage_text, re.DOTALL
    )

    for asset_name_raw, asset_body in p3_asset_blocks:
        name = asset_name_raw.strip().rstrip('*').strip()
        description = p2_assets.get(name, ("", ""))[0]
        target = description.split(',')[0].strip() if description else ""
        indications = []
        asset = PipelineAsset(name, target, indications)

        stage_lines = re.findall(
            r'Stage\s+(\d+)\s*\([^)]*\)\s*:\s*([^\n]+)',
            asset_body, re.IGNORECASE
        )
        for stage_num_str, stage_rest in stage_lines:
            stage_num = int(stage_num_str)
            year_match = re.search(r'(?<!\d)(20[1-4]\d)(?!\d)', stage_rest)
            if year_match:
                asset.stages[stage_num] = int(year_match.group(1))

        if not asset.stages:
            table_stages = re.findall(
                r'\|\s*\*{0,2}Stage\s+(\d+)[^|]*\*{0,2}\s*\|\s*\*{0,2}(\d{4})\*{0,2}',
                asset_body, re.IGNORECASE
            )
            for stage_num_str, year_str in table_stages:
                asset.stages[int(stage_num_str)] = int(year_str)

        if name in p2_assets:
            asset_section = p2_assets[name][1]
            indication_ms_blocks = re.findall(
                r'(?:^-\s*|\*\*)?Market Share Projection\s*\(([^)]+)\)(?:\*\*)?[^\n]*\n(.*?)(?=(?:^-\s*|\*\*)Market Share Projection\s*\(|\n---\n|\n####|\n###|\Z)',
                asset_section, re.DOTALL | re.MULTILINE
            )
            if indication_ms_blocks:
                for ind_name, ms_text in indication_ms_blocks:
                    ind_name = ind_name.strip()
                    year_shares = re.findall(r'\|\s*(\d{4})\s*\|\s*([\d.]+)%', ms_text)
                    if year_shares:
                        asset.market_shares[ind_name] = {}
                        for year_str, share_str in year_shares:
                            asset.market_shares[ind_name][int(year_str)] = float(share_str) / 100.0
                        asset.indications.append(ind_name)
            else:
                year_shares = re.findall(r'\|\s*(\d{4})\s*\|\s*([\d.]+)%', asset_section)
                if year_shares:
                    asset.market_shares["All"] = {}
                    for year_str, share_str in year_shares:
                        asset.market_shares["All"][int(year_str)] = float(share_str) / 100.0

        assets.append(asset)

    return assets


# ══════════════════════════════════════════════════════════════════════════════
#  XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _xml_escape(text: str) -> str:
    """Escape XML special characters."""
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
                .replace("'", "&apos;"))


def _col_letter(col_idx: int) -> str:
    """1-based column index → letter. 1=A, 25=Y."""
    return chr(64 + col_idx)


def _asset_full_name(asset: PipelineAsset) -> str:
    """Build display name like 'CTX-009 (DLL4 and VEGF-A, BTC/CRC)'."""
    ind_list = asset.indications if asset.indications else list(asset.market_shares.keys())
    ind_display = [i for i in ind_list if i not in ("All", "All Indications Combined")]
    if asset.target and ind_display:
        return f"{asset.name} ({asset.target}, {'/'.join(ind_display)})"
    elif asset.target:
        return f"{asset.name} ({asset.target})"
    elif ind_display:
        return f"{asset.name} ({'/'.join(ind_display)})"
    return asset.name


# ══════════════════════════════════════════════════════════════════════════════
#  SCENARIO BLOCK XML BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _spacer_row(row_num: int) -> str:
    """Empty spacer row — only Y column styled."""
    return f'<row r="{row_num}"><c r="Y{row_num}" s="502"/></row>'


def _section_divider_row(row_num: int, label: str) -> str:
    """Section divider with thick bottom border (e.g. 'Break Down')."""
    parts = [f'<row r="{row_num}" ht="15.75" thickBot="1">']
    for ci in range(1, 14):
        c = _col_letter(ci)
        if c == 'C':
            parts.append(f'<c r="C{row_num}" s="57" t="inlineStr"><is><t>{_xml_escape(label)}</t></is></c>')
        else:
            parts.append(f'<c r="{c}{row_num}" s="57"/>')
    for ci in range(14, 25):
        parts.append(f'<c r="{_col_letter(ci)}{row_num}" s="15"/>')
    parts.append(f'<c r="Y{row_num}" s="497"/>')
    parts.append('</row>')
    return ''.join(parts)


def _scenario_header_row(row_num: int, scenario_num: int, scenario_name: str) -> str:
    """Scenario header row: B=number, C=name, D-X empty."""
    parts = [f'<row r="{row_num}">']
    parts.append(f'<c r="A{row_num}" s="69"/>')
    parts.append(f'<c r="B{row_num}" s="433"><v>{scenario_num}</v></c>')
    parts.append(f'<c r="C{row_num}" s="70" t="inlineStr"><is><t>{_xml_escape(scenario_name)}</t></is></c>')
    for ci in range(4, 25):
        parts.append(f'<c r="{_col_letter(ci)}{row_num}" s="71"/>')
    parts.append(f'<c r="Y{row_num}" s="502"/>')
    parts.append('</row>')
    return ''.join(parts)


def _scenario_asset_row(row_num: int, header_row: int, abs_asset_row: int,
                        abs_first: int, abs_last: int, first_ms_row: int) -> str:
    """Scenario asset row with SUMIF formulas referencing Absolute section."""
    parts = [f'<row r="{row_num}">']
    parts.append(f'<c r="A{row_num}" s="434"><f>B{header_row}</f></c>')
    parts.append(f'<c r="B{row_num}" s="52" t="str"><f>C{header_row}</f><v></v></c>')
    parts.append(f'<c r="C{row_num}" s="73"><f>$C${abs_asset_row}</f></c>')
    parts.append(f'<c r="D{row_num}" s="71"/>')
    for ci in range(5, 25):
        c = _col_letter(ci)
        f = (f'IF($Y{first_ms_row}=0,0,'
             f'SUMIF($C${abs_first}:$C${abs_last},$C{row_num},'
             f'{c}${abs_first}:{c}${abs_last}))')
        parts.append(f'<c r="{c}{row_num}" s="74"><f>{f}</f></c>')
    parts.append(f'<c r="Y{row_num}" s="503"/>')
    parts.append('</row>')
    return ''.join(parts)


def _scenario_ms_row(row_num: int, scenario_asset_row: int,
                     indication: str, peak) -> str:
    """Scenario market share row.

    peak: float value (literal) or str formula (e.g. '=$Y$11').
    """
    parts = [f'<row r="{row_num}">']
    parts.append(f'<c r="A{row_num}" s="434"><f>A{scenario_asset_row}</f></c>')
    parts.append(f'<c r="B{row_num}" s="52" t="str"><f>B{scenario_asset_row}</f><v></v></c>')
    # C: MS name formula
    if indication in ("All", "All Indications Combined"):
        cf = f'C{scenario_asset_row}&amp;" Market Share"'
    else:
        cf = f'C{scenario_asset_row}&amp;" {_xml_escape(indication)} Market Share"'
    parts.append(f'<c r="C{row_num}" s="75" t="str"><f>{cf}</f><v></v></c>')
    parts.append(f'<c r="D{row_num}" s="52" t="inlineStr"><is><t>[%]</t></is></c>')
    # E-G: literal 0
    for ci in range(5, 8):
        parts.append(f'<c r="{_col_letter(ci)}{row_num}" s="76"><v>0</v></c>')
    # H-X: IF formulas
    for ci in range(8, 25):
        c = _col_letter(ci)
        p = _col_letter(ci - 1)
        f = f'IF({c}{scenario_asset_row}=5,$Y{row_num},MAX($I{row_num}:{p}{row_num}))'
        parts.append(f'<c r="{c}{row_num}" s="77"><f>{f}</f></c>')
    # Y: peak
    if isinstance(peak, str) and peak.startswith('='):
        parts.append(f'<c r="Y{row_num}" s="505"><f>{peak[1:]}</f></c>')
    else:
        parts.append(f'<c r="Y{row_num}" s="505"><v>{peak}</v></c>')
    parts.append('</row>')
    return ''.join(parts)


# ══════════════════════════════════════════════════════════════════════════════
#  PEAK PROVIDERS  (callable(asset_name, indication) → peak value or formula)
# ══════════════════════════════════════════════════════════════════════════════

def _base_peak_provider(abs_ms_rows: Dict):
    """Y = formula referencing Absolute section's Y column."""
    def provider(asset_name, ind):
        r = abs_ms_rows.get((asset_name, ind))
        return f"=$Y${r}" if r else 0
    return provider


def _literal_peak_provider(overrides: Optional[Dict], abs_ms_peaks: Dict):
    """Literal Y values. Use overrides first, fall back to base peaks."""
    def provider(asset_name, ind):
        if overrides and asset_name in overrides and ind in overrides[asset_name]:
            return overrides[asset_name][ind]
        return abs_ms_peaks.get((asset_name, ind), 0)
    return provider


def _breakdown_peak_provider(included_assets: set, abs_ms_rows: Dict):
    """Formula ref for included assets, 0 for excluded (disabled)."""
    def provider(asset_name, ind):
        if asset_name in included_assets:
            r = abs_ms_rows.get((asset_name, ind))
            return f"=$Y${r}" if r else 0
        return 0
    return provider


def _extract_peak_overrides(assets: Optional[List[PipelineAsset]]) -> Optional[Dict]:
    """Extract {drug_name: {indication: peak_float}} from asset list."""
    if not assets:
        return None
    peaks = {}
    for asset in assets:
        peaks[asset.name] = {}
        for ind, shares in asset.market_shares.items():
            if shares:
                peaks[asset.name][ind] = max(shares.values())
    return peaks


def _extract_unified_peaks(assets: List[PipelineAsset], scenario: str) -> Optional[Dict]:
    """Extract bull/bear peaks from unified asset data.

    scenario: "bull" or "bear"
    Returns: {drug_name: {indication: peak_float}} or None
    """
    peaks = {}
    has_data = False
    for asset in assets:
        shares_dict = asset.bull_shares if scenario == "bull" else asset.bear_shares
        if shares_dict:
            has_data = True
            peaks[asset.name] = {}
            for ind, shares in shares_dict.items():
                if shares:
                    peaks[asset.name][ind] = max(shares.values())
    return peaks if has_data else None


def _catalyst_peak_provider(catalyst_asset_name: str, catalyst: 'CatalystEvent',
                            outcome: str, abs_ms_rows: Dict, abs_ms_peaks: Dict):
    """Peak provider for catalyst scenarios.

    For the catalyst's own asset+indication: use positive/negative peak.
    For all other assets: use base peak (formula ref).
    """
    def provider(asset_name, ind):
        if asset_name == catalyst_asset_name:
            # Check if this indication matches the catalyst
            cat_ind = catalyst.indication
            if ind == cat_ind or cat_ind in ind or ind in cat_ind:
                if outcome == "positive" and catalyst.positive_peak > 0:
                    return catalyst.positive_peak
                elif outcome == "negative" and catalyst.negative_peak > 0:
                    return catalyst.negative_peak
        # Fallback: formula ref to absolute
        r = abs_ms_rows.get((asset_name, ind))
        return f"=$Y${r}" if r else abs_ms_peaks.get((asset_name, ind), 0)
    return provider


# ══════════════════════════════════════════════════════════════════════════════
#  SCENARIO BLOCK ASSEMBLY
# ══════════════════════════════════════════════════════════════════════════════

def _add_scenario_block(new_rows: list, start_row: int,
                        scenario_num: int, scenario_name: str,
                        assets: List[PipelineAsset],
                        abs_first: int, abs_last: int,
                        abs_asset_rows: Dict, abs_ms_rows: Dict,
                        peak_provider) -> int:
    """Build a complete scenario block (header + asset/MS rows).

    Returns: next available row number.
    """
    cur = start_row
    header_row = cur
    new_rows.append(_scenario_header_row(cur, scenario_num, scenario_name))
    cur += 1

    for asset in assets:
        abs_ar = abs_asset_rows[asset.name]
        asset_row = cur
        first_ms = cur + 1

        new_rows.append(_scenario_asset_row(
            cur, header_row, abs_ar, abs_first, abs_last, first_ms))
        cur += 1

        indications = list(asset.market_shares.keys())
        if not indications:
            indications = ["All"]
        for ind in indications:
            peak = peak_provider(asset.name, ind)
            new_rows.append(_scenario_ms_row(cur, asset_row, ind, peak))
            cur += 1

    return cur


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN GENERATION FUNCTION
# ══════════════════════════════════════════════════════════════════════════════

def generate_scenarios_from_template(
    template_path: Path,
    dcf_path: Path,
    assets: List[PipelineAsset],
) -> None:
    """Generate complete Scenarios sheet: Absolute + Base/Bull/Bear + Breakdown + Catalyst."""

    # ── Step 1: Extract existing Scenarios sheet XML ──
    logger.info("Extracting existing Scenarios sheet from DCF file...")
    with zipfile.ZipFile(dcf_path) as zf:
        if "xl/worksheets/sheet7.xml" in zf.namelist():
            existing_xml = zf.read("xl/worksheets/sheet7.xml").decode("utf-8")
        else:
            with zipfile.ZipFile(template_path) as zft:
                existing_xml = zft.read("xl/worksheets/sheet7.xml").decode("utf-8")

    # ── Step 2: Find insertion point (after row 9) ──
    row9_match = re.search(r'<row r="9"[^>]*>.*?</row>', existing_xml, re.DOTALL)
    if not row9_match:
        logger.error("Could not find row 9")
        return
    xml_before = existing_xml[:row9_match.end()]

    sheetdata_end = existing_xml.find('</sheetData>')
    if sheetdata_end == -1:
        logger.error("Could not find </sheetData>")
        return
    xml_after = existing_xml[sheetdata_end:]

    logger.info("Building scenario rows...")
    new_rows = []
    current_row = 10

    # ═══════════════════════════════════════════════════════════════════════
    #  MODULE 0: Absolute Value (Scenario 4)
    # ═══════════════════════════════════════════════════════════════════════
    abs_first = current_row
    abs_asset_rows = {}   # {asset.name: row_num}
    abs_ms_rows = {}      # {(asset.name, indication): row_num}
    abs_ms_peaks = {}     # {(asset.name, indication): peak_float}

    for asset in assets:
        abs_asset_rows[asset.name] = current_row
        full_name = _asset_full_name(asset)

        # ── Asset row ──
        rp = [f'<row r="{current_row}" spans="1:31" s="65" customFormat="1">']
        rp.append(f'<c r="A{current_row}" s="62"><v>4</v></c>')
        rp.append(f'<c r="B{current_row}" s="34" t="inlineStr"><is><t> Absolute</t></is></c>')
        rp.append(f'<c r="C{current_row}" s="63" t="inlineStr"><is><t>{_xml_escape(full_name)}</t></is></c>')
        rp.append(f'<c r="D{current_row}" s="61"/>')
        for ci in range(5, 25):
            year = 2019 + (ci - 5)
            c = _col_letter(ci)
            stage_num = None
            for sn, sy in asset.stages.items():
                if sy == year:
                    stage_num = sn
                    break
            if stage_num:
                rp.append(f'<c r="{c}{current_row}" s="64"><v>{stage_num}</v></c>')
            else:
                rp.append(f'<c r="{c}{current_row}" s="64"/>')
        rp.append(f'<c r="Y{current_row}" s="500"/>')
        if current_row == 10:
            rp.append(f'<c r="AA{current_row}" s="72" t="inlineStr"><is><t>Stage 5 = Approved</t></is></c>')
        rp.append('</row>')
        new_rows.append(''.join(rp))
        asset_row = current_row
        current_row += 1

        # ── Market share rows ──
        if not asset.market_shares:
            asset.market_shares["All"] = {}
        for ind_name, shares_dict in asset.market_shares.items():
            abs_ms_rows[(asset.name, ind_name)] = current_row
            peak_share = max(shares_dict.values()) if shares_dict else 0.01
            abs_ms_peaks[(asset.name, ind_name)] = peak_share

            rp = [f'<row r="{current_row}" spans="1:31">']
            rp.append(f'<c r="A{current_row}" s="62"><v>4</v></c>')
            rp.append(f'<c r="B{current_row}" s="34" t="inlineStr"><is><t> Absolute</t></is></c>')
            if ind_name in ("All Indications Combined", "All"):
                formula = f'C{asset_row}&amp;" Market Share"'
                cached = f'{full_name} Market Share'
            else:
                formula = f'C{asset_row}&amp;" {_xml_escape(ind_name)} Market Share"'
                cached = f'{full_name} {ind_name} Market Share'
            rp.append(f'<c r="C{current_row}" s="66" t="str"><f>{formula}</f><v>{_xml_escape(cached)}</v></c>')
            rp.append(f'<c r="D{current_row}" s="34" t="inlineStr"><is><t>[%]</t></is></c>')
            for ci in range(5, 8):
                c = _col_letter(ci)
                v = shares_dict.get(2019 + (ci - 5), 0)
                rp.append(f'<c r="{c}{current_row}" s="67"><v>{v}</v></c>')
            for ci in range(8, 25):
                c = _col_letter(ci)
                p = _col_letter(ci - 1)
                f = f'IF({c}{asset_row}=5,$Y{current_row},MAX($I{current_row}:{p}{current_row}))'
                v = shares_dict.get(2019 + (ci - 5), 0)
                rp.append(f'<c r="{c}{current_row}" s="68"><f>{f}</f><v>{v}</v></c>')
            rp.append(f'<c r="Y{current_row}" s="501"><v>{peak_share}</v></c>')
            rp.append('</row>')
            new_rows.append(''.join(rp))
            current_row += 1

    abs_last = current_row - 1
    logger.info(f"  Absolute (Scenario 4): rows {abs_first}-{abs_last}")

    # ═══════════════════════════════════════════════════════════════════════
    #  MODULE 1: Base / Bull / Bear Scenarios
    # ═══════════════════════════════════════════════════════════════════════

    # ── Base (Scenario 1): Y peaks = formula ref to Absolute ──
    new_rows.append(_spacer_row(current_row)); current_row += 1
    s = current_row
    current_row = _add_scenario_block(
        new_rows, current_row, 1, "Base", assets,
        abs_first, abs_last, abs_asset_rows, abs_ms_rows,
        _base_peak_provider(abs_ms_rows))
    logger.info(f"  Base (Scenario 1): rows {s}-{current_row - 1}")

    # ── Bull (Scenario 2): peaks from asset.bull_shares or fallback to base ──
    new_rows.append(_spacer_row(current_row)); current_row += 1
    s = current_row
    bull_ov = _extract_unified_peaks(assets, "bull")
    current_row = _add_scenario_block(
        new_rows, current_row, 2, "Bull", assets,
        abs_first, abs_last, abs_asset_rows, abs_ms_rows,
        _literal_peak_provider(bull_ov, abs_ms_peaks))
    logger.info(f"  Bull (Scenario 2): rows {s}-{current_row - 1}")

    # ── Bear (Scenario 3): peaks from asset.bear_shares or fallback to base ──
    new_rows.append(_spacer_row(current_row)); current_row += 1
    s = current_row
    bear_ov = _extract_unified_peaks(assets, "bear")
    current_row = _add_scenario_block(
        new_rows, current_row, 3, "Bear", assets,
        abs_first, abs_last, abs_asset_rows, abs_ms_rows,
        _literal_peak_provider(bear_ov, abs_ms_peaks))
    logger.info(f"  Bear (Scenario 3): rows {s}-{current_row - 1}")

    # ═══════════════════════════════════════════════════════════════════════
    #  MODULE 2: Break Down (cumulative drug addition)
    # ═══════════════════════════════════════════════════════════════════════
    new_rows.append(_spacer_row(current_row)); current_row += 1
    new_rows.append(_section_divider_row(current_row, "Break Down")); current_row += 1

    for k in range(len(assets)):
        new_rows.append(_spacer_row(current_row)); current_row += 1
        name = f"{assets[k].name} Only" if k == 0 else f"+{assets[k].name}"
        s = current_row
        included = {a.name for a in assets[:k + 1]}
        current_row = _add_scenario_block(
            new_rows, current_row, 5 + k, name, assets,
            abs_first, abs_last, abs_asset_rows, abs_ms_rows,
            _breakdown_peak_provider(included, abs_ms_rows))
        logger.info(f"  Breakdown '{name}' (Scenario {5 + k}): rows {s}-{current_row - 1}")

    # ═══════════════════════════════════════════════════════════════════════
    #  MODULE 3: Catalyst Scenarios
    # ═══════════════════════════════════════════════════════════════════════
    all_catalysts = []
    for asset in assets:
        for cat in asset.catalysts:
            all_catalysts.append((asset, cat))

    new_rows.append(_spacer_row(current_row)); current_row += 1
    new_rows.append(_section_divider_row(current_row, "Catalyst Scenarios")); current_row += 1

    cat_num = 5 + len(assets)

    if all_catalysts:
        for asset, cat in all_catalysts:
            # Positive outcome block
            new_rows.append(_spacer_row(current_row)); current_row += 1
            pos_label = f"{cat.name} (+)"
            s = current_row
            pos_peaks = _catalyst_peak_provider(
                asset.name, cat, "positive", abs_ms_rows, abs_ms_peaks)
            current_row = _add_scenario_block(
                new_rows, current_row, cat_num, pos_label, assets,
                abs_first, abs_last, abs_asset_rows, abs_ms_rows,
                pos_peaks)
            logger.info(f"  Catalyst+ '{cat.name}' (Scenario {cat_num}): rows {s}-{current_row - 1}")
            cat_num += 1

            # Negative outcome block
            new_rows.append(_spacer_row(current_row)); current_row += 1
            neg_label = f"{cat.name} (-)"
            s = current_row
            neg_peaks = _catalyst_peak_provider(
                asset.name, cat, "negative", abs_ms_rows, abs_ms_peaks)
            current_row = _add_scenario_block(
                new_rows, current_row, cat_num, neg_label, assets,
                abs_first, abs_last, abs_asset_rows, abs_ms_rows,
                neg_peaks)
            logger.info(f"  Catalyst- '{cat.name}' (Scenario {cat_num}): rows {s}-{current_row - 1}")
            cat_num += 1
    else:
        # Placeholder block if no catalysts parsed
        new_rows.append(_spacer_row(current_row)); current_row += 1
        s = current_row
        current_row = _add_scenario_block(
            new_rows, current_row, cat_num, "", assets,
            abs_first, abs_last, abs_asset_rows, abs_ms_rows,
            _base_peak_provider(abs_ms_rows))
        logger.info(f"  Catalyst placeholder (Scenario {cat_num}): rows {s}-{current_row - 1}")

    # ── Step 3: Assemble final XML ──
    new_max_row = current_row - 1
    final_xml = xml_before + '\n' + '\n'.join(new_rows) + '\n' + xml_after
    final_xml = re.sub(
        r'<dimension ref="A\d+:AE\d+"/>',
        f'<dimension ref="A1:AE{new_max_row}"/>',
        final_xml)
    logger.info(f"Generated rows 10-{new_max_row} ({len(new_rows)} XML rows)")

    # ── Step 4: Patch DCF file ──
    logger.info(f"Patching DCF file: {dcf_path}")
    modified = {"xl/worksheets/sheet7.xml": final_xml.encode("utf-8")}

    with zipfile.ZipFile(dcf_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp_path = dcf_path.with_suffix(".~scenarios.xlsx")
    with zipfile.ZipFile(dcf_path, "r") as zin:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue  # removed — references stripped below
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
    try:
        tmp_path.replace(dcf_path)
    except PermissionError:
        import os
        os.remove(str(dcf_path))
        tmp_path.rename(dcf_path)

    logger.info(f"Scenarios sheet saved to: {dcf_path}")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def _print_summary(assets):
    print(f"\n{'='*70}")
    print(f"Parsed {len(assets)} Pipeline Assets (Unified)")
    print(f"{'='*70}")
    for asset in assets:
        ind_list = asset.indications if asset.indications else list(asset.market_shares.keys())
        peak_shares = {k: max(v.values()) if v else 0 for k, v in asset.market_shares.items()}
        bull_peaks = {k: max(v.values()) if v else 0 for k, v in asset.bull_shares.items()}
        bear_peaks = {k: max(v.values()) if v else 0 for k, v in asset.bear_shares.items()}
        print(f"\n  {asset.name} ({asset.target})")
        print(f"    Indications: {', '.join(ind_list)}")
        print(f"    Stages: {asset.stages}")
        for ind, peak in peak_shares.items():
            b_peak = bull_peaks.get(ind, peak)
            r_peak = bear_peaks.get(ind, peak)
            print(f"    {ind}: base={peak:.1%}, bull={b_peak:.1%}, bear={r_peak:.1%}")
        if asset.catalysts:
            print(f"    Catalysts: {len(asset.catalysts)}")
            for cat in asset.catalysts:
                print(f"      {cat.name} ({cat.indication}): +{cat.positive_peak:.1%} / -{cat.negative_peak:.1%}")
    print(f"{'='*70}\n")


def main():
    parser = argparse.ArgumentParser(description="Generate Scenarios sheet from Gemini research")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--report-dir", help="Directory with per-drug research reports (default: DD/{TICKER}/pipeline_base4/)")
    parser.add_argument("--research-file", help="Legacy: single combined research file")
    parser.add_argument("--dcf-file", help="DCF file path")
    parser.add_argument("--template-file", help="Template file (default: DCF Template 2020.xlsx)")
    parser.add_argument("--dry-run", action="store_true", help="Parse only, don't write Excel")

    args = parser.parse_args()

    dcf_path = Path(args.dcf_file) if args.dcf_file else Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/DCF {args.ticker}.xlsx")
    template_path = Path(args.template_file) if args.template_file else Path("/mnt/c/Users/yzsun/Desktop/DD/base/DCF Template 2020.xlsx")

    # Parse base reports
    if args.research_file:
        research_path = Path(args.research_file)
        if not research_path.exists():
            logger.error(f"Research file not found: {research_path}")
            sys.exit(1)
        logger.info(f"Parsing legacy research file: {research_path}")
        assets = parse_gemini_report(research_path)
    else:
        report_dir = Path(args.report_dir) if args.report_dir else Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/pipeline_base4/")
        if not report_dir.exists():
            logger.error(f"Report directory not found: {report_dir}")
            sys.exit(1)
        logger.info(f"Parsing report directory: {report_dir}")
        assets = parse_gemini_reports(report_dir, args.ticker)

    if not assets:
        logger.error("No assets found")
        sys.exit(1)

    _print_summary(assets)

    if args.dry_run:
        logger.info("Dry run — skipping Excel generation")
        return

    if not dcf_path.exists():
        logger.error(f"DCF file not found: {dcf_path}")
        sys.exit(1)

    if not template_path.exists():
        logger.error(f"Template file not found: {template_path}")
        sys.exit(1)

    # Backup
    backup_path = dcf_path.with_name(f"{dcf_path.stem}_pre_scenarios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(dcf_path, backup_path)
    logger.info(f"Backup: {backup_path}")

    generate_scenarios_from_template(template_path, dcf_path, assets)

    total_cats = sum(len(a.catalysts) for a in assets)
    has_bull = any(a.bull_shares for a in assets)
    has_bear = any(a.bear_shares for a in assets)
    print(f"\n{'='*70}")
    print("Scenarios Sheet Generated")
    print(f"{'='*70}")
    print(f"Ticker: {args.ticker}")
    print(f"Assets: {len(assets)}")
    print(f"Modules: Absolute + Base/Bull/Bear + {len(assets)} Breakdown + {total_cats * 2 or 1} Catalyst")
    print(f"Bull peaks: {'from report' if has_bull else 'same as base'}")
    print(f"Bear peaks: {'from report' if has_bear else 'same as base'}")
    print(f"Catalysts: {total_cats} events ({total_cats * 2} blocks)")
    print(f"File: {dcf_path}")
    print(f"Backup: {backup_path}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
