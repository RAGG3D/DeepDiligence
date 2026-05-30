#!/usr/bin/env python3
"""
generate_financials.py — Generate RBS (and future RIS/RCFS/FSA/VALUATION)
from FY DATA content.

Reads FY DATA to discover all IS/BS/CFS items, then generates clean
financial model sheets — no Reserved rows, all items from FY DATA included.

Usage:
    python generate_financials.py --ticker CMPX [--dry-run]
"""
import argparse
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# ── Sheet zip paths ──────────────────────────────────────────────────
SHEET_RBS  = "xl/worksheets/sheet11.xml"
SHEET_RCFS = "xl/worksheets/sheet12.xml"

# ══════════════════════════════════════════════════════════════════════
# Style constants — extracted from DCF Template 2020.xlsx RBS sheet
# ══════════════════════════════════════════════════════════════════════

# ── A column ──
S_A = 2            # standard spacer
S_A_NET = 79       # summary net rows (NOA, NFA)
S_A_THICK = 186    # thick-blank row
S_A_EQ_SUB = 179   # Stockholders' Equity sub-header

# ── B column ──
S_B = 164          # dedup formula cell
S_B_THICK = 187    # thick-blank row

# ── C column ──
S_C = 11           # standard tag
S_C_SUMMARY = 197  # Summary section header C

# ── D column ──
S_D_LABEL = 75     # data rows, equity capital raise items
S_D_COMPANY = 54   # company name, FY label, year header, driver, summary net
S_D_SECTION = 82   # major section headers
S_D_SUBHEAD = 92   # sub-headers (Operating Assets/Liabilities, Financial Assets)
S_D_SUBHEAD_EQ = 73  # Stockholders' Equity sub-header
S_D_TOTAL = 118    # total rows, summary totals
S_D_TOTAL_EQ = 183 # Total Stockholders' Equity
S_D_CHECK = 184    # check rows, APIC delta
S_D_SHARES = 181   # Shares Outstanding, Shares Issued
S_D_THICK = 188    # thick-blank row

# ── E column ──
S_E = 52           # standard unit [MM USD]
S_E_SECTION = 107  # section header
S_E_SUBHEAD = 46   # sub-header
S_E_CHECK = 34     # check rows
S_E_NET = 119      # summary net rows
S_E_THICK = 189    # thick-blank row

# ── F-W data columns (by row type) ──
# Year header
S_YEAR_H = 56      # historical year
S_YEAR_F = 55      # forecast year

# Section headers
S_SECT_INTER = 165  # Intersheet Forecast Drivers, Operating A&L
S_SECT_FIN = 172    # Financial Assets & Liabilities
S_SECT_EQ = 177     # Equity, Equity Capital Raise, Summary

# Sub-headers
S_SUB_OA = 167      # Operating Assets
S_SUB_OL = 171      # Operating Liabilities
S_SUB_FIN = 86      # Financial Assets, Stockholders' Equity

# Data
S_DATA_H = 168      # operating data historical
S_DATA_F = 169      # operating data forecast
S_FIN_H = 174       # financial/equity historical
S_FIN_F = 175       # financial data forecast
S_EQ_F = 182        # equity forecast (Ord Shares, APIC, Shares Outstanding)

# Totals
S_TOT_OP = 170      # operating totals (all cols)
S_TOT_FIN = 176     # financial/equity totals (all cols)
S_TOT_INV = 43      # Total Investments (all cols)

# Check
S_CHECK = 185       # check rows (all cols)

# Driver
S_DRIVER = 49       # intersheet driver (all cols)

# Summary
S_SUMMARY = 40      # summary total (all cols)
S_SUMMARY_NET = 198 # summary net (all cols)

# Thick-blank F
S_F_THICK = 190

# Equity Capital Raise section
S_SHARES_ISS_H = 191  # Shares Issued hist
S_PRICE_H = 192     # Share Price hist
S_PRICE_F = 193     # Share Price forecast
S_DISCOUNT_H = 194  # Discount hist
S_DISCOUNT_F = 94   # Discount forecast
S_ISSUE_H = 182     # Issue Price hist (= S_EQ_F)
S_ISSUE_F = 191     # Issue Price forecast (= S_SHARES_ISS_H)
S_APIC_MOD = 191    # APIC Modeled (all cols)
S_APIC_REP_H = 174  # APIC Reported hist (= S_FIN_H)
S_APIC_REP_F = 195  # APIC Reported forecast

# ── Constants ────────────────────────────────────────────────────────
HIST_COLS = 4        # F-I (2021-2024) → map to FY DATA G-J
YEAR_COLS = 18       # F-W (2021-2038)
FIRST_YEAR = 2021    # RBS starting year
FY_OFFSET = 1        # RBS F → FY DATA G

# Financial Assets (not in Operating Assets)
FINANCIAL_ASSETS = {"Cash And Cash Equivalents"}
# Equity items
EQUITY_ITEMS = {
    "Ordinary Shares",
    "Additional Paid-In Capital",
    "Accumulated Deficit",
    "Accumulated Other Comprehensive (Loss) Income",
}
# PP&E comes from Schedules, not FY DATA
PPE_ITEM = "Property And Equipment, Net"


def _strip_formula_cache(xml: str) -> str:
    """Remove cached ERROR values from formula cells to force Excel recalculation.

    Cells with t="e" and <f>: strip t="e" and <v> (Excel recalculates).
    Cells with t="e" but NO <f>: convert to empty cell (self-closing <c/>).
    Leaves valid cached values (numbers, strings) intact.
    """
    def _fix_error_cell(m):
        tag = m.group(0)
        if '<f' in tag:
            # Has formula — remove error type and cached value, keep formula
            tag = re.sub(r' t="e"', '', tag)
            tag = re.sub(r'<v>[^<]*</v>', '', tag)
        else:
            # No formula — convert to self-closing empty cell
            open_tag = re.match(r'<c\b[^>]*', tag).group(0)
            open_tag = re.sub(r' t="e"', '', open_tag)
            tag = open_tag + '/>'
        return tag

    xml = re.sub(r'<c\b[^>]* t="e"[^>]*>.*?</c>', _fix_error_cell, xml)
    return xml


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;")
            .replace("'", "&apos;"))


def _formula_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def _col_letter(n: int) -> str:
    """1-indexed → letter. 6→F, 23→W."""
    if n <= 26:
        return chr(64 + n)
    return chr(64 + (n - 1) // 26) + chr(64 + (n - 1) % 26 + 1)


# ── FY DATA Discovery ───────────────────────────────────────────────

def _load_shared_strings(zf):
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    raw = zf.read("xl/sharedStrings.xml").decode("utf-8")
    root = ET.fromstring(raw)
    strings = []
    for si in root.findall(f"{{{ns}}}si"):
        texts = si.findall(f".//{{{ns}}}t")
        strings.append("".join(t.text or "" for t in texts))
    return strings


def _cell_text(xml, col, row, strings):
    m = re.search(rf'<c r="{col}{row}"[^/]*?(?:/>|>.*?</c>)', xml)
    if not m:
        return None
    cell = m.group()
    v = re.search(r'<v>(\d+)</v>', cell)
    if v and 't="s"' in cell:
        return strings[int(v.group(1))]
    if 't="inlineStr"' in cell:
        t = re.search(r'<t>([^<]*)</t>', cell)
        return t.group(1) if t else ""
    return None


def _has_sum_formula(xml, row):
    m = re.search(rf'<c r="F{row}"[^/]*?(?:/>|>.*?</c>)', xml)
    if not m:
        return False
    f = re.search(r'<f[^>]*>([^<]*)</f>', m.group())
    return f and "SUM" in f.group(1).upper() if f else False


def discover_fydata_items(zf, strings):
    """Read FY DATA K USD → classified BS items."""
    xml = zf.read("xl/worksheets/sheet15.xml").decode("utf-8")

    bs_items = []  # [(name, section)]
    section = "asset"

    for row in range(50, 90):
        c = _cell_text(xml, "C", row, strings)
        d = _cell_text(xml, "D", row, strings)

        if c is None and d:
            dl = d.lower()
            if "liabilit" in dl:
                section = "liability"
            elif "equity" in dl or "stockholder" in dl:
                section = "equity"
            continue

        if c != "BS" or not d:
            continue
        if _has_sum_formula(xml, row):
            continue  # skip totals

        bs_items.append((d, section))

    return bs_items


# ══════════════════════════════════════════════════════════════════════
# XML Row Builders — exact match to DCF Template 2020.xlsx formatting
# ══════════════════════════════════════════════════════════════════════

NS = ' x14ac:dyDescent="0.25"'
NS_THICK = ' ht="15.75" thickBot="1" x14ac:dyDescent="0.3"'


def _row_open(r, thick=False, row_style=None):
    s = f' s="{row_style}" customFormat="1"' if row_style else ""
    attrs = NS_THICK if thick else NS
    return f'<row r="{r}" spans="1:23"{s}{attrs}>'


def _c(col, r, s, t=None, text=None, formula=None):
    """Build a single cell XML string."""
    s_attr = f' s="{s}"' if s is not None else ""
    if formula:
        return f'<c r="{col}{r}"{s_attr}><f>{_formula_escape(formula)}</f></c>'
    if text is not None:
        return (f'<c r="{col}{r}"{s_attr} t="inlineStr">'
                f'<is><t>{_xml_escape(text)}</t></is></c>')
    if t:
        return f'<c r="{col}{r}"{s_attr} t="inlineStr"><is><t>{t}</t></is></c>'
    return f'<c r="{col}{r}"{s_attr}/>'


# ── Blank rows ──

def _blank_row_r1(r):
    """R1-style blank: A=2, B=164, E=52."""
    return (f'{_row_open(r)}'
            f'{_c("A", r, S_A)}{_c("B", r, S_B)}{_c("E", r, S_E)}'
            f'</row>')


def _blank_row_empty(r, thick=False):
    """Empty-looking row. Must have at least one cell for Excel compatibility."""
    if thick:
        return f'{_row_open(r, thick=True)}{_c("A", r, None)}</row>'
    return f'{_row_open(r)}{_c("A", r, None)}</row>'


def _blank_row_b_only(r):
    """Blank row with B=164 cell (template R9, R31, R39 pattern)."""
    return f'{_row_open(r)}{_c("A", r, None)}{_c("B", r, S_B)}</row>'


def _thick_blank_row(r):
    """Special thick-blank with full styling (template R48, R59)."""
    return (f'{_row_open(r, thick=True, row_style=51)}'
            f'{_c("A", r, S_A_THICK)}'
            f'{_c("B", r, S_B_THICK)}'
            f'{_c("D", r, S_D_THICK)}'
            f'{_c("E", r, S_E_THICK)}'
            + "".join(f'{_c(_col_letter(6+i), r, S_F_THICK)}' for i in range(YEAR_COLS))
            + '</row>')


# ── Header rows ──

def _company_row(r):
    return (f'{_row_open(r)}'
            f'{_c("A", r, S_A)}{_c("B", r, S_B)}'
            f'{_c("D", r, S_D_COMPANY, formula="\'WELCOME!\'!$B$18")}'
            f'{_c("E", r, S_E)}'
            f'</row>')


def _fy_label_row(r, text="Fiscal Year Ended December 31."):
    return (f'{_row_open(r)}'
            f'{_c("A", r, S_A)}{_c("B", r, S_B)}'
            f'{_c("D", r, S_D_COMPANY, text=text)}'
            f'{_c("E", r, S_E)}'
            f'</row>')


def _year_header_row(r, first_year, n_hist):
    cells = [_c("A", r, S_A), _c("B", r, S_B),
             _c("C", r, S_C, text="DATA"),
             _c("D", r, S_D_COMPANY, text="Fiscal Year"),
             _c("E", r, S_E)]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_YEAR_H if i < n_hist else S_YEAR_F
        if i == 0:
            cells.append(f'<c r="{col}{r}" s="{s}"><v>{first_year}</v></c>')
        else:
            cells.append(_c(col, r, s, formula=f"{_col_letter(5+i)}{r}+1"))
    return f'{_row_open(r, thick=True)}' + "".join(cells) + '</row>'


# ── Section and sub-section headers ──

def _section_header(r, text, f_style):
    """Major section header (thickBot, D=82, E=107)."""
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_SECTION, text=text),
             _c("E", r, S_E_SECTION)]
    for i in range(YEAR_COLS):
        cells.append(_c(_col_letter(6+i), r, f_style))
    return f'{_row_open(r, thick=True)}' + "".join(cells) + '</row>'


def _sub_header(r, text, f_style, a_style=S_A, d_style=S_D_SUBHEAD, has_c=True):
    """Sub-section header (no thickBot, D=92, E=46)."""
    cells = [_c("A", r, a_style), _c("B", r, S_B)]
    if has_c:
        cells.append(_c("C", r, S_C))
    cells += [_c("D", r, d_style, text=text),
              _c("E", r, S_E_SUBHEAD)]
    for i in range(YEAR_COLS):
        cells.append(_c(_col_letter(6+i), r, f_style))
    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── B column dedup formula ──

def _b_formula(r):
    # Note: literal " is fine in <f> element content; only &lt; &gt; &amp; need escaping.
    # &quot; in formula text causes Excel to "remove" the formula on repair.
    # <v></v> (explicit close) instead of <v/> for compatibility.
    f = (f'IF(AND($C{r}&lt;&gt;"",COUNTIFS($D:$D,$D{r},$C:$C,$C{r})&gt;1),'
         f'COUNTIFS($D$1:D{r},$D{r},$C$1:C{r},$C{r}),"")' )
    return f'<c r="B{r}" s="{S_B}" t="str"><f>{f}</f><v></v></c>'


# ── Data rows ──

def _data_row(r, name, c_tag, s_hist, s_fore, n_hist,
              source="FY DATA", source_formula=None):
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text=c_tag),
             _c("D", r, S_D_LABEL, text=name),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = s_hist if i < n_hist else s_fore

        if i < n_hist:
            fy_col = _col_letter(6 + i + FY_OFFSET)
            if source_formula:
                f = source_formula.replace("{COL}", fy_col).replace("{ROW}", str(r))
            else:
                f = (f"SUMIFS('{source}'!{fy_col}:{fy_col},"
                     f"'{source}'!$D:$D,$D{r},"
                     f"'{source}'!$C:$C,$C{r},"
                     f"'{source}'!$B:$B,$B{r})")
            cells.append(_c(col, r, s, formula=f))
        elif i == n_hist:
            cells.append(_c(col, r, s,
                           formula=f"AVERAGE(F{r}:{_col_letter(5+n_hist)}{r})"))
        else:
            cells.append(_c(col, r, s, formula=f"{_col_letter(5+i)}{r}"))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Total rows ──

def _total_row(r, label, formula_template, d_style=S_D_TOTAL, num_style=S_TOT_OP):
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="RBS"),
             _c("D", r, d_style, text=label),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        f = formula_template.replace("{COL}", col).replace("{ROW}", str(r))
        cells.append(_c(col, r, num_style, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Check rows ──

def _check_row(r, label, formula_template):
    cells = [_c("A", r, S_A), _c("B", r, S_B),
             _c("D", r, S_D_CHECK, text=label),
             _c("E", r, S_E_CHECK, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        f = formula_template.replace("{COL}", col).replace("{ROW}", str(r))
        cells.append(_c(col, r, S_CHECK, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Driver row ──

def _driver_row(r, label, c_tag, formula_template):
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text=c_tag),
             _c("D", r, S_D_COMPANY, text=label),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        ris_col = _col_letter(6 + i + 1)  # RIS offset +1
        f = formula_template.replace("{COL}", ris_col).replace("{ROW}", str(r))
        cells.append(_c(col, r, S_DRIVER, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Summary rows ──

def _summary_total_row(r, label, formula_template):
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="RBS"),
             _c("D", r, S_D_TOTAL, text=label),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        f = formula_template.replace("{COL}", col)
        cells.append(_c(col, r, S_SUMMARY, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


def _summary_net_row(r, label, formula_template):
    cells = [_c("A", r, S_A_NET), _b_formula(r),
             _c("C", r, S_C, text="RBS"),
             _c("D", r, S_D_COMPANY, text=label),
             _c("E", r, S_E_NET, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        f = formula_template.replace("{COL}", col)
        cells.append(_c(col, r, S_SUMMARY_NET, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Cash row (forecast from RCFS ending cash) ──────────────────────

def _cash_row(r, name, n_hist):
    """Cash: historical from FY DATA, forecast from RCFS Ending Cash."""
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="BS"),
             _c("D", r, S_D_LABEL, text=name),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_FIN_H if i < n_hist else S_FIN_F
        if i < n_hist:
            fy_col = _col_letter(6 + i + FY_OFFSET)
            f = (f"SUMIFS('FY DATA'!{fy_col}:{fy_col},"
                 f"'FY DATA'!$D:$D,$D{r},"
                 f"'FY DATA'!$C:$C,$C{r},"
                 f"'FY DATA'!$B:$B,$B{r})")
            cells.append(_c(col, r, s, formula=f))
        else:
            # Cash = RCFS Ending Cash (the BS plug)
            f = (f'SUMIFS(RCFS!{col}:{col},'
                 f'RCFS!$C:$C,"RCFS",'
                 f'RCFS!$D:$D,"Ending Cash")')
            cells.append(_c(col, r, s, formula=f))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ── Equity special rows ─────────────────────────────────────────────

def _equity_accum_deficit_row(r, name, n_hist):
    """Accumulated Deficit: same style hist/fore (174/174), forecast = prior + RIS Net Income."""
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="BS"),
             _c("D", r, S_D_LABEL, text=name),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        if i < n_hist:
            fy_col = _col_letter(6 + i + FY_OFFSET)
            f = (f"SUMIFS('FY DATA'!{fy_col}:{fy_col},"
                 f"'FY DATA'!$D:$D,$D{r},"
                 f"'FY DATA'!$C:$C,$C{r},"
                 f"'FY DATA'!$B:$B,$B{r})")
            cells.append(_c(col, r, S_FIN_H, formula=f))
        else:
            prev = _col_letter(5 + i)
            ris_col = _col_letter(6 + i + 1)
            cells.append(_c(col, r, S_FIN_H,
                           formula=f"SUM(RIS!{ris_col}$92,{prev}{r})"))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


def _equity_oci_row(r, name, n_hist):
    """AOCI: same style hist/fore (174/174), forecast = prior + RIS OCI."""
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="BS"),
             _c("D", r, S_D_LABEL, text=name),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        if i < n_hist:
            fy_col = _col_letter(6 + i + FY_OFFSET)
            f = (f"SUMIFS('FY DATA'!{fy_col}:{fy_col},"
                 f"'FY DATA'!$D:$D,$D{r},"
                 f"'FY DATA'!$C:$C,$C{r},"
                 f"'FY DATA'!$B:$B,$B{r})")
            cells.append(_c(col, r, S_FIN_H, formula=f))
        else:
            prev = _col_letter(5 + i)
            ris_col = _col_letter(6 + i + 1)
            cells.append(_c(col, r, S_FIN_H,
                           formula=f"SUM(RIS!{ris_col}$96,{prev}{r})"))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


def _equity_apic_row(r, name, n_hist):
    """APIC: 174/182, forecast = prior + APIC modeled (placeholder __APIC__)."""
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="BS"),
             _c("D", r, S_D_LABEL, text=name),
             _c("E", r, S_E, text="[MM USD]")]

    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_FIN_H if i < n_hist else S_EQ_F
        if i < n_hist:
            fy_col = _col_letter(6 + i + FY_OFFSET)
            f = (f"SUMIFS('FY DATA'!{fy_col}:{fy_col},"
                 f"'FY DATA'!$D:$D,$D{r},"
                 f"'FY DATA'!$C:$C,$C{r},"
                 f"'FY DATA'!$B:$B,$B{r})")
            cells.append(_c(col, r, s, formula=f))
        else:
            prev = _col_letter(5 + i)
            cells.append(_c(col, r, s, formula=f"{prev}{r}+{col}__APIC__"))

    return f'{_row_open(r)}' + "".join(cells) + '</row>'


# ══════════════════════════════════════════════════════════════════════
# RBS Generator
# ══════════════════════════════════════════════════════════════════════

def generate_rbs(bs_items, n_hist=4):
    """Generate RBS sheetData XML from classified BS items."""
    # Classify items
    op_assets = []
    fin_assets = []
    op_liabs = []
    equity = []

    for name, section in bs_items:
        if name in FINANCIAL_ASSETS:
            fin_assets.append(name)
        elif name == PPE_ITEM:
            continue  # comes from Schedules
        elif any(name.startswith(e) for e in EQUITY_ITEMS):
            equity.append(name)
        elif section == "asset":
            op_assets.append(name)
        else:
            op_liabs.append(name)

    rows = []
    r = 1

    # ── R1: blank (A=2, B=164, E=52) ──
    rows.append(_blank_row_r1(r)); r += 1

    # ── R2: Company name ──
    rows.append(_company_row(r)); r += 1

    # ── R3: FY label ──
    rows.append(_fy_label_row(r)); r += 1

    # ── R4: Year headers (thickBot) ──
    rows.append(_year_header_row(r, FIRST_YEAR, n_hist))
    year_row = r; r += 1

    # ── R5: Intersheet Forecast Drivers (thickBot section header) ──
    rows.append(_section_header(r, "Intersheet Forecast Drivers", S_SECT_INTER))
    r += 1

    # ── R6: Net Income driver ──
    rows.append(_driver_row(r, "Net Income (Loss) [After Tax]", "RIS",
                            "SUMIFS(RIS!{COL}:{COL},RIS!$D:$D,$D{ROW},RIS!$C:$C,$C{ROW})"))
    net_income_row = r; r += 1

    # ── R7: empty thickBot (before section) ──
    rows.append(_blank_row_empty(r, thick=True)); r += 1

    # ══ Operating Assets & Liabilities ══
    rows.append(_section_header(r, "Operating Assets & Liabilities", S_SECT_INTER))
    r += 1

    # ── R9: blank with B only ──
    rows.append(_blank_row_b_only(r)); r += 1

    # ── Operating Assets sub-header ──
    rows.append(_sub_header(r, "Operating Assets", S_SUB_OA))
    r += 1

    # PP&E from Schedules
    ppe_row = r
    rows.append(_data_row(r, "PP&E [Net]", "SCHE", S_DATA_H, S_DATA_H, YEAR_COLS,
                          source_formula="SUMIFS(Schedules!{COL}:{COL},Schedules!$C:$C,RBS!$C{ROW},Schedules!$D:$D,RBS!$D{ROW})"))
    r += 1

    # Other operating assets
    oa_first = r
    for name in op_assets:
        rows.append(_data_row(r, name, "BS", S_DATA_H, S_DATA_F, n_hist))
        r += 1
    oa_last = r - 1

    # Total Operating Assets
    tot_oa_row = r
    rows.append(_total_row(r, "Total Operating Assets",
                           f"SUM({{COL}}{ppe_row}:{{COL}}{oa_last})",
                           num_style=S_TOT_OP))
    r += 1

    # blank (empty — no cells at all)
    rows.append(_blank_row_empty(r)); r += 1

    # ── Operating Liabilities sub-header (no C cell) ──
    rows.append(_sub_header(r, "Operating Liabilities", S_SUB_OL, has_c=False))
    r += 1

    ol_first = r
    for name in op_liabs:
        rows.append(_data_row(r, name, "BS", S_DATA_H, S_DATA_F, n_hist))
        r += 1
    ol_last = r - 1

    # Total Operating Liabilities
    tot_ol_row = r
    rows.append(_total_row(r, "Total Operating Liabilities",
                           f'SUMIF($C${ol_first}:$C${ol_last},"BS",{{COL}}{ol_first}:{{COL}}{ol_last})',
                           num_style=S_TOT_OP))
    r += 1

    # blank (empty thickBot — before Financial section)
    rows.append(_blank_row_empty(r, thick=True)); r += 1

    # ══ Financial Assets & Liabilities ══
    rows.append(_section_header(r, "Financial Assets & Liabilities", S_SECT_FIN))
    r += 1

    # blank (B only)
    rows.append(_blank_row_b_only(r)); r += 1

    # Financial Assets sub-header
    rows.append(_sub_header(r, "Financial Assets", S_SUB_FIN))
    r += 1

    fa_first = r
    for name in fin_assets:
        if name in FINANCIAL_ASSETS:
            rows.append(_cash_row(r, name, n_hist))
        else:
            rows.append(_data_row(r, name, "BS", S_FIN_H, S_FIN_F, n_hist))
        r += 1
    fa_last = r - 1

    tot_fa_row = r
    rows.append(_total_row(r, "Total Financial Assets",
                           f"SUM({{COL}}{fa_first}:{{COL}}{fa_last})",
                           num_style=S_TOT_FIN))
    r += 1

    # blank (empty)
    rows.append(_blank_row_empty(r)); r += 1

    # Total Investments (no C cell, D=75, F=43)
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("D", r, S_D_LABEL, text="Total Investments"),
             _c("E", r, S_E, text="[MM USD]")]
    f_tmpl = f"SUM({{COL}}{fa_first}:{{COL}}{fa_last})"
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        cells.append(_c(col, r, S_TOT_INV,
                       formula=f_tmpl.replace("{COL}", col)))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # blank (empty thickBot — before Equity section)
    rows.append(_blank_row_empty(r, thick=True)); r += 1

    # ══ Equity ══
    rows.append(_section_header(r, "Equity", S_SECT_EQ))
    r += 1

    # blank (B only)
    rows.append(_blank_row_b_only(r)); r += 1

    # Stockholders' Equity sub-header (special: A=179, D=73, row_style=54)
    rows.append(_sub_header(r, "Stockholders' Equity", S_SUB_FIN,
                            a_style=S_A_EQ_SUB, d_style=S_D_SUBHEAD_EQ))
    # Override row to add row_style=54
    rows[-1] = rows[-1].replace(f'<row r="{r}" spans="1:23"',
                                f'<row r="{r}" spans="1:23" s="54" customFormat="1"')
    r += 1

    eq_first = r
    for name in equity:
        if "Accumulated Deficit" in name:
            rows.append(_equity_accum_deficit_row(r, name, n_hist))
        elif "Accumulated Other Comprehensive" in name:
            rows.append(_equity_oci_row(r, name, n_hist))
        elif "Additional Paid-In Capital" in name:
            rows.append(_equity_apic_row(r, name, n_hist))
        else:
            # Ordinary Shares: 174/182
            rows.append(_data_row(r, name, "BS", S_FIN_H, S_EQ_F, n_hist))
        r += 1
    eq_last = r - 1

    tot_eq_row = r
    rows.append(_total_row(r, "Total Stockholders' Equity",
                           f"SUM({{COL}}{eq_first}:{{COL}}{eq_last})",
                           d_style=S_D_TOTAL_EQ, num_style=S_TOT_FIN))
    r += 1

    # blank (empty)
    rows.append(_blank_row_empty(r)); r += 1

    # ── Check: Accounting Equation ──
    check1_row = r
    rows.append(_check_row(r,
        "Check: Accounting Equation (Assets = Liabilities + Equity)",
        f"IF(ABS(SUM({{COL}}{tot_oa_row},{{COL}}{tot_fa_row})"
        f"-SUM({{COL}}{tot_ol_row},{{COL}}{tot_eq_row}))>1,"
        f"SUM({{COL}}{tot_oa_row},{{COL}}{tot_fa_row})"
        f"-SUM({{COL}}{tot_ol_row},{{COL}}{tot_eq_row}),0)"))
    r += 1

    # ── Thick blank (special: s=51, full styled cells) ──
    rows.append(_thick_blank_row(r)); r += 1

    # ══ Equity Capital Raise ══
    rows.append(_section_header(r, "Equity Capital Raise", S_SECT_EQ))
    eq_raise_row = r; r += 1

    # Shares Outstanding (row_style=54)
    shares_row = r
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="RBS"),
             _c("D", r, S_D_SHARES, text="Shares Outstanding"),
             _c("E", r, S_E, text="[MM Shares]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_FIN_H if i < n_hist else S_EQ_F
        if i < n_hist:
            cells.append(_c(col, r, s,
                           formula=f"SUMIFS('BBG DAPI'!$AV:$AV,'BBG DAPI'!$AU:$AU,RBS!{col}${year_row})"))
        else:
            prev = _col_letter(5 + i)
            cells.append(_c(col, r, s, formula=f"{prev}{r}+{col}{r+1}"))
    rows.append(f'{_row_open(r, row_style=54)}' + "".join(cells) + '</row>')
    r += 1

    # Shares Issued
    shares_issued_row = r
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_SHARES, text="Shares Issued (Repurchased)"),
             _c("E", r, S_E, text="[MM Shares]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_SHARES_ISS_H if i < n_hist else S_EQ_F
        cells.append(_c(col, r, s))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # Share Price FY End
    price_row = r
    cells = [_c("A", r, S_A), _b_formula(r),
             _c("C", r, S_C, text="RBS"),
             _c("D", r, S_D_LABEL, text="Share Price Fiscal Year End"),
             _c("E", r, S_E, text="[USD/Share]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_PRICE_H if i < n_hist else S_PRICE_F
        if i < n_hist:
            cells.append(_c(col, r, s,
                           formula=f"SUMIFS('BBG DAPI'!$AO:$AO,'BBG DAPI'!$AN:$AN,RBS!{col}${year_row})"))
        elif i == n_hist:
            cells.append(_c(col, r, s, formula="'BBG DAPI'!$I$5"))
        else:
            prev = _col_letter(5 + i)
            cells.append(_c(col, r, s, formula=f"{prev}{r}*(1+VALUATION!$C$19)"))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # Discounted to Market Price
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_LABEL, text="Discounted to Market Price [Approximate]"),
             _c("E", r, S_E, text="[% Share Price]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_DISCOUNT_H if i < n_hist else S_DISCOUNT_F
        cells.append(_c(col, r, s))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # Issue Price
    issue_row = r
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_LABEL, text="Issue (Repurchase) Price"),
             _c("E", r, S_E, text="[USD/Share]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_ISSUE_H if i < n_hist else S_ISSUE_F
        cells.append(_c(col, r, s))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # APIC Modeled
    apic_modeled_row = r
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_LABEL, text="Additional Paid-In Capital [As Modeled]"),
             _c("E", r, S_E, text="[MM USD]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        cells.append(_c(col, r, S_APIC_MOD,
                       formula=f"{col}{shares_issued_row}*{col}{issue_row}"))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # APIC Reported
    apic_reported_row = r
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("C", r, S_C),
             _c("D", r, S_D_LABEL, text="Additional Paid-In Capital [As Reported]"),
             _c("E", r, S_E, text="[MM USD]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        s = S_APIC_REP_H if i < n_hist else S_APIC_REP_F
        cells.append(_c(col, r, s))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # Blank (R57-style: A=2, B=164, E=52, F=43)
    cells = [_c("A", r, S_A), _c("B", r, S_B), _c("E", r, S_E)]
    for i in range(YEAR_COLS):
        cells.append(_c(_col_letter(6+i), r, S_TOT_INV))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # APIC Delta
    cells = [_c("A", r, S_A), _c("B", r, S_B),
             _c("D", r, S_D_CHECK, text="Additional Paid-In Capital [Modeled - Reported]"),
             _c("E", r, S_E_CHECK, text="[MM USD]")]
    for i in range(YEAR_COLS):
        col = _col_letter(6 + i)
        f = (f"IFERROR(IF(ABS({col}{apic_modeled_row}-{col}{apic_reported_row})"
             f">1,{col}{apic_modeled_row}-{col}{apic_reported_row},0),\"\")")
        cells.append(_c(col, r, S_CHECK, formula=f))
    rows.append(f'{_row_open(r)}' + "".join(cells) + '</row>')
    r += 1

    # ── Thick blank (s=51) ──
    rows.append(_thick_blank_row(r)); r += 1

    # ══ Summary And Metrics (row_style=65, C=197) ══
    rows.append(_section_header(r, "Summary And Metrics", S_SECT_EQ))
    # Override row_style=65 and C style to S_C_SUMMARY=197
    rows[-1] = rows[-1].replace(
        f'<row r="{r}" spans="1:23"',
        f'<row r="{r}" spans="1:23" s="65" customFormat="1"')
    rows[-1] = rows[-1].replace(f'<c r="C{r}" s="{S_C}"',
                                f'<c r="C{r}" s="{S_C_SUMMARY}"')
    summary_section_row = r; r += 1

    rows.append(_summary_total_row(r, "Total Operating Assets", f"{{COL}}{tot_oa_row}"))
    r += 1
    rows.append(_summary_total_row(r, "Total Operating Liabilities", f"{{COL}}{tot_ol_row}"))
    r += 1
    noa_row = r
    rows.append(_summary_net_row(r, "Net Operating Assets (NOA)",
                                 f"{{COL}}{r-2}-{{COL}}{r-1}"))
    r += 1
    rows.append(_summary_total_row(r, "Total Financial Assets", f"{{COL}}{tot_fa_row}"))
    r += 1
    nfa_row = r
    rows.append(_summary_net_row(r, "Net Financial Assets (NFA)", f"{{COL}}{r-1}"))
    r += 1
    eq_summary_row = r
    rows.append(_summary_total_row(r, "Total Stockholders' Equity",
                                   f"{{COL}}{tot_eq_row}"))
    # Override D style to S_D_TOTAL_EQ=183
    rows[-1] = rows[-1].replace(f'<c r="D{r}" s="{S_D_TOTAL}"',
                                f'<c r="D{r}" s="{S_D_TOTAL_EQ}"')
    r += 1

    rows.append(_check_row(r,
        "Check: Accounting Equation (Assets - Liabilities = Equity)",
        f"IF(ABS(SUM({{COL}}{noa_row},{{COL}}{nfa_row})-{{COL}}{eq_summary_row})>1,"
        f"SUM({{COL}}{noa_row},{{COL}}{nfa_row})-{{COL}}{eq_summary_row},0)"))
    r += 1

    metadata = {
        "tot_oa_row": tot_oa_row,
        "tot_ol_row": tot_ol_row,
        "tot_fa_row": tot_fa_row,
        "tot_eq_row": tot_eq_row,
        "check1_row": check1_row,
        "shares_row": shares_row,
        "apic_modeled_row": apic_modeled_row,
        "total_rows": r - 1,
    }

    return rows, metadata


# ── Conditional Formatting ────────────────────────────────────────────

def _generate_conditional_formatting(rows_xml, meta):
    """Generate conditional formatting matching template pattern."""
    # Collect rows that have B column cells
    b_rows = []
    for rx in rows_xml:
        m = re.search(r'<c r="B(\d+)"', rx)
        if m:
            b_rows.append(int(m.group(1)))

    if not b_rows:
        return ""

    # Build contiguous ranges
    ranges = []
    start = b_rows[0]
    end = b_rows[0]
    for r in b_rows[1:]:
        if r == end + 1:
            end = r
        else:
            ranges.append((start, end))
            start = r
            end = r
    ranges.append((start, end))

    parts = []
    priority = 1
    for start, end in ranges:
        sqref = f"B{start}:B{end}" if start != end else f"B{start}"
        parts.append(
            f'<conditionalFormatting sqref="{sqref}">'
            f'<cfRule type="expression" priority="{priority}">'
            f'<formula>B{start}&lt;&gt;""</formula>'
            f'</cfRule></conditionalFormatting>')
        priority += 1

    # Check row highlighting (ABS > 0.1 → red)
    tr = meta["total_rows"]
    check_refs = []
    for rx in rows_xml:
        m = re.search(r'<c r="D(\d+)"[^>]*>.*?Check:', rx)
        if m:
            check_refs.append(f"D{m.group(1)}:W{m.group(1)}")
    if check_refs:
        sqref = " ".join(check_refs)
        first_r = check_refs[0].split(":")[0][1:]
        parts.append(
            f'<conditionalFormatting sqref="{sqref}">'
            f'<cfRule type="expression" dxfId="61" priority="{priority}">'
            f'<formula>ABS(D{first_r})&gt;0.1</formula>'
            f'</cfRule></conditionalFormatting>')

    return "".join(parts)


# ── RCFS Ending-Cash formula fix ────────────────────────────────────

def _patch_rcfs_cash_formulas(rcfs_xml: str, n_hist: int) -> str:
    """Fix RCFS R38 Ending Cash so that ALL historical columns read from RBS.

    The RCFS template hardcodes G38=RBS!G33 and H38=RBS!H33 (years 2022-2023).
    When n_hist > 3 (e.g. n_hist=4 includes 2024 at col I), col I38 was still
    using the shared SUM formula, producing wrong beginning-cash for 2025+ and
    therefore a broken RBS balance sheet check.

    Fix: for each historical col beyond H (cols I, J, … up to col for n_hist-1),
    replace the shared-formula cell with an explicit RBS!{col}33 reference and
    promote the first forecast col to be the new shared-formula anchor.
    """
    if n_hist <= 3:
        return rcfs_xml  # G38 and H38 already correct; nothing to fix

    # Columns that need fixing: i=3..n_hist-1 (2024, 2025, … if n_hist > 4)
    # (i=0→F, i=1→G, i=2→H already correct in template)
    fix_cols       = [_col_letter(6 + i) for i in range(3, n_hist)]
    first_fc_col   = _col_letter(6 + n_hist)       # first forecast col (e.g. J)
    last_col       = _col_letter(6 + YEAR_COLS - 1) # W

    # ------------------------------------------------------------------
    # Step 1 – replace the shared-formula ANCHOR cell (first fix col, e.g. I38)
    # Its cell XML looks like:
    #   <c r="I38" s="31"><f t="shared" ref="I38:W38" ca="1" si="6">SUM(I36:I37)</f><v>...</v></c>
    # ------------------------------------------------------------------
    anchor = fix_cols[0]  # 'I' for n_hist=4
    rcfs_xml = re.sub(
        rf'<c r="{anchor}38" s="\d+"><f[^<]*</f><v>[^<]*</v></c>',
        f'<c r="{anchor}38" s="31"><f>RBS!{anchor}33</f><v></v></c>',
        rcfs_xml,
    )
    # Edge case: self-closing <f/>
    rcfs_xml = re.sub(
        rf'<c r="{anchor}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
        f'<c r="{anchor}38" s="31"><f>RBS!{anchor}33</f><v></v></c>',
        rcfs_xml,
    )

    # ------------------------------------------------------------------
    # Step 2 – replace any additional historical-fix cols (n_hist > 4)
    # These were shared instances: <c r="J38" s="31"><f t="shared" ca="1" si="6" /><v>...</v></c>
    # ------------------------------------------------------------------
    for col in fix_cols[1:]:
        rcfs_xml = re.sub(
            rf'<c r="{col}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
            f'<c r="{col}38" s="31"><f>RBS!{col}33</f><v></v></c>',
            rcfs_xml,
        )

    # ------------------------------------------------------------------
    # Step 3 – promote the first forecast col to be the new shared-formula anchor
    # Was: <c r="J38" s="31"><f t="shared" ca="1" si="6" /><v>...</v></c>
    # New: <c r="J38" s="31"><f t="shared" ref="J38:W38" ca="1" si="6">SUM(J36:J37)</f><v></v></c>
    # ------------------------------------------------------------------
    rcfs_xml = re.sub(
        rf'<c r="{first_fc_col}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
        (f'<c r="{first_fc_col}38" s="31">'
         f'<f t="shared" ref="{first_fc_col}38:{last_col}38" ca="1" si="6">'
         f'SUM({first_fc_col}36:{first_fc_col}37)</f><v></v></c>'),
        rcfs_xml,
    )

    return rcfs_xml


# ── Main ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate financial model sheets from FY DATA")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    xlsx = Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/DCF {args.ticker}.xlsx")
    if not xlsx.exists():
        print(f"ERROR: {xlsx} not found")
        return

    # Backup
    if not args.dry_run:
        bak = xlsx.with_suffix(".xlsx.bak_gen")
        shutil.copy2(xlsx, bak)
        print(f"Backup: {bak}")

    with zipfile.ZipFile(xlsx, "r") as zf:
        strings = _load_shared_strings(zf)
        bs_items = discover_fydata_items(zf, strings)

        print(f"FY DATA BS items ({len(bs_items)}):")
        for name, section in bs_items:
            print(f"  [{section:10s}] {name}")

        # Generate RBS
        rbs_rows, meta = generate_rbs(bs_items)

        # Fix APIC row reference
        apic_row = meta["apic_modeled_row"]
        for i, row_xml in enumerate(rbs_rows):
            if "__APIC__" in row_xml:
                rbs_rows[i] = row_xml.replace("__APIC__", str(apic_row))

        print(f"\nRBS: {meta['total_rows']} rows generated")
        print(f"  Total Op Assets row:  R{meta['tot_oa_row']}")
        print(f"  Total Op Liabs row:   R{meta['tot_ol_row']}")
        print(f"  Total Fin Assets row: R{meta['tot_fa_row']}")
        print(f"  Total Equity row:     R{meta['tot_eq_row']}")
        print(f"  Check row:            R{meta['check1_row']}")

        # Read existing RBS XML to get header/footer
        rbs_xml = zf.read(SHEET_RBS).decode("utf-8")
        sd_start = rbs_xml.index("<sheetData")
        sd_tag_end = rbs_xml.index(">", sd_start) + 1
        sd_close = rbs_xml.index("</sheetData>")

        header = rbs_xml[:sd_tag_end]
        # footer starts with </sheetData>...; split into close tag + rest
        after_sd = rbs_xml[sd_close + len("</sheetData>"):]

        # Strip old conditional formatting from after_sd
        after_sd = re.sub(r'<conditionalFormatting[^>]*>.*?</conditionalFormatting>',
                          '', after_sd, flags=re.DOTALL)

        # Generate new conditional formatting
        cond_fmt = _generate_conditional_formatting(rbs_rows, meta)

        # Update dimension ref
        header = re.sub(r'<dimension ref="[^"]*"/>',
                        f'<dimension ref="A1:W{meta["total_rows"]}"/>', header)

        new_rbs = (header + "\n".join(rbs_rows) + "</sheetData>"
                   + cond_fmt + after_sd)

        if args.dry_run:
            print("\n  DRY RUN — not written")
            return

        # Patch RCFS Ending Cash formulas for historical cols beyond H
        rcfs_xml = zf.read(SHEET_RCFS).decode("utf-8")
        rcfs_xml = _patch_rcfs_cash_formulas(rcfs_xml, HIST_COLS)

        # Add fullCalcOnLoad
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wb_xml:
            wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)

        # Strip calcChain references (delete file + strip from CT and rels)
        ct_xml  = zf.read("[Content_Types].xml").decode("utf-8")
        wr_xml  = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        ct_xml  = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', '', ct_xml)
        wr_xml  = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', wr_xml)

        # Zip-patch
        modified = {
            SHEET_RBS:                 new_rbs.encode("utf-8"),
            SHEET_RCFS:                rcfs_xml.encode("utf-8"),
            "xl/workbook.xml":         wb_xml.encode("utf-8"),
            "[Content_Types].xml":     ct_xml.encode("utf-8"),
            "xl/_rels/workbook.xml.rels": wr_xml.encode("utf-8"),
        }

        tmp = xlsx.with_suffix(".~gen.xlsx")
        with zipfile.ZipFile(xlsx, "r") as zin, \
             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue  # delete calcChain; fullCalcOnLoad rebuilds it
                if item.filename in modified:
                    data = modified[item.filename]
                elif (item.filename.startswith("xl/worksheets/sheet")
                      and item.filename.endswith(".xml")):
                    # Strip cached formula errors from ALL worksheets
                    data = _strip_formula_cache(
                        zin.read(item.filename).decode("utf-8")
                    ).encode("utf-8")
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

        tmp.replace(xlsx)
        sz = xlsx.stat().st_size
        print(f"\n  Written {sz:,} bytes → {xlsx}")


if __name__ == "__main__":
    main()
