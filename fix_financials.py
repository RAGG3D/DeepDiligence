#!/usr/bin/env python3
"""
fix_financials.py — Fix DCF CMPX.xlsx RIS + Schedules sheets.

1. Text patches: rename drug names, R&D/G&A ISN items, PP&E BSN items
2. Revenue/COGS formulas: replace broken SUMIFS with INDEX/MATCH on Pipeline
3. Zip-patch in-place (skip calcChain.xml, add fullCalcOnLoad)

Usage:
    python fix_financials.py [--dry-run]
"""
import argparse
import re
import shutil
import zipfile
from pathlib import Path

XLSX = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
SHEET_RIS   = "xl/worksheets/sheet10.xml"
SHEET_RBS   = "xl/worksheets/sheet11.xml"
SHEET_RCFS  = "xl/worksheets/sheet12.xml"
SHEET_SCHED = "xl/worksheets/sheet13.xml"

# Historical year count for CMPX (2021-2024 → 4 cols F,G,H,I)
HIST_COLS   = 4
YEAR_COLS   = 18  # F-W

# ── Drug names (must match Scenarios $C exactly) ──────────────────────
DRUG1 = "CTX-009 (DLL4 and VEGF-A, BTC/CRC)"
DRUG2 = "CTX-10726 (PD-1 x VEGF-A, RCC/HCC/GC/EC)"
DRUG3 = "CTX-471 (CD137 (4-1BB, TNFRSF9), ES-SCLC/Melanoma NCAM+/MPM)"
DRUG4 = "CTX-8371 (PD-1 and PD-L1, NSCLC/TNBC/HL/MM/HNSCC)"

# ── RIS text patches ─────────────────────────────────────────────────
RIS_TEXT = {
    # Pipeline timeline (D8-D11)
    "D8": DRUG1, "D9": DRUG2, "D10": DRUG3, "D11": DRUG4,
    # Revenue (D26-D29)
    "D26": f"{DRUG1} Revenue",
    "D27": f"{DRUG2} Revenue",
    "D28": f"{DRUG3} Revenue",
    "D29": f"{DRUG4} Revenue",
    # COGS (D37-D40)
    "D37": f"{DRUG1} COGS",
    "D38": f"{DRUG2} COGS",
    "D39": f"{DRUG3} COGS",
    "D40": f"{DRUG4} COGS",
    # R&D ISN items (must match FY DATA K USD)
    "D47": "CTX-009",
    "D48": "CTX-471",
    "D49": "CTX-8371",
    "D50": "Other research and development expenses",
    "D52": "Other Research And Development",
    "D54": "Reserved", "D55": "Reserved", "D56": "Reserved", "D58": "Reserved",
    # G&A ISN items
    "D62": "General And Administrative Expenses",
    "D63": "Reserved", "D64": "Reserved", "D66": "Reserved", "D67": "Reserved",
}

# ── Schedules text patches ────────────────────────────────────────────
SCHED_TEXT = {
    # Drug timeline
    "D7": DRUG1, "D8": DRUG2, "D9": DRUG3, "D10": DRUG4,
    # PP&E BSN items
    "D20": "Equipment",
    "D22": "Leasehold improvements",
    "D24": "Software",
    "D26": "Furniture and fixtures",
    "D29": "Less: Accumulated depreciation and amortization",
}

# ── Revenue/COGS formula mapping ─────────────────────────────────────
# RIS row → Pipeline row for Revenue and COGS
#   CTX-009:   Pipeline R15 (rev) / R16 (cogs)
#   CTX-10726: Pipeline R28 / R29
#   CTX-471:   Pipeline R39 / R40  (was R41/R42, verify)
#   CTX-8371:  Pipeline R54 / R55  (was R56/R57, verify)
REV_MAP = {26: 15, 27: 28, 28: 39, 29: 54}  # RIS row → Pipeline revenue row
COGS_MAP = {37: 16, 38: 29, 39: 40, 40: 55}  # RIS row → Pipeline COGS row


# ── Helpers ───────────────────────────────────────────────────────────

def _strip_formula_cache(xml: str) -> str:
    """Remove cached ERROR values from formula cells to force Excel recalculation.

    Cells with t="e" and <f>: strip t="e" and <v> (Excel recalculates).
    Cells with t="e" but NO <f>: convert to empty cell (self-closing <c/>).
    Leaves valid cached values (numbers, strings) intact.
    """
    def _fix_error_cell(m):
        tag = m.group(0)
        if '<f' in tag:
            tag = re.sub(r' t="e"', '', tag)
            tag = re.sub(r'<v>[^<]*</v>', '', tag)
        else:
            open_tag = re.match(r'<c\b[^>]*', tag).group(0)
            open_tag = re.sub(r' t="e"', '', open_tag)
            tag = open_tag + '/>'
        return tag

    xml = re.sub(r'<c\b[^>]* t="e"[^>]*>.*?</c>', _fix_error_cell, xml)
    return xml


def _normalize_formula_xml(xml: str) -> str:
    """Fix formula XML issues that cause Excel's 'Removed Records: Formula' repair.

    1. Double-escaped entities from formulas pre-escaped then run through
       _formula_escape again: &amp;gt; → &gt;, &amp;lt; → &lt;, &amp;quot; → "
    2. &quot; inside <f>: not decoded by Excel — replace with literal ".
    3. <v/> self-closing: replace with <v></v> for compatibility.
    """
    def _fix_formula(m):
        f = m.group(0)
        f = f.replace('&amp;gt;', '&gt;')
        f = f.replace('&amp;lt;', '&lt;')
        f = f.replace('&amp;amp;', '&amp;')
        f = f.replace('&amp;quot;', '"')
        f = f.replace('&quot;', '"')
        return f

    xml = re.sub(r'<f\b[^>]*>.*?</f>', _fix_formula, xml, flags=re.DOTALL)
    xml = xml.replace('<v/>', '<v></v>')
    return xml


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
    )


def _col_letter(n: int) -> str:
    """1-indexed column number → letter(s). 6→'F', 24→'X'."""
    if n <= 26:
        return chr(64 + n)
    return chr(64 + (n - 1) // 26) + chr(64 + (n - 1) % 26 + 1)


def _patch_rcfs_cash_formulas(rcfs_xml: str, n_hist: int = HIST_COLS) -> str:
    """Fix RCFS R38 Ending Cash so ALL historical cols read from RBS!{col}33.

    The template hardcodes G38=RBS!G33 and H38=RBS!H33.  For n_hist=4
    col I (2024) is still historical, but the template used a shared SUM
    formula starting at I38 — causing RBS 2025+ Cash to use wrong base cash.

    Fix: make each historical col beyond H use RBS!{col}33 (explicit formula),
    then promote the first forecast col to be the new shared-formula anchor.
    """
    if n_hist <= 3:
        return rcfs_xml  # G38 and H38 already correct; nothing to fix

    fix_cols      = [_col_letter(6 + i) for i in range(3, n_hist)]
    first_fc_col  = _col_letter(6 + n_hist)
    last_col      = _col_letter(6 + YEAR_COLS - 1)  # W

    # Step 1 — replace the shared-formula ANCHOR (e.g. I38) with explicit RBS ref
    anchor = fix_cols[0]
    rcfs_xml = re.sub(
        rf'<c r="{anchor}38" s="\d+"><f[^<]*</f><v>[^<]*</v></c>',
        f'<c r="{anchor}38" s="31"><f>RBS!{anchor}33</f><v></v></c>',
        rcfs_xml,
    )
    rcfs_xml = re.sub(
        rf'<c r="{anchor}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
        f'<c r="{anchor}38" s="31"><f>RBS!{anchor}33</f><v></v></c>',
        rcfs_xml,
    )

    # Step 2 — replace any additional historical-fix cols (n_hist > 4)
    for col in fix_cols[1:]:
        rcfs_xml = re.sub(
            rf'<c r="{col}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
            f'<c r="{col}38" s="31"><f>RBS!{col}33</f><v></v></c>',
            rcfs_xml,
        )

    # Step 3 — promote first forecast col to new shared-formula anchor
    rcfs_xml = re.sub(
        rf'<c r="{first_fc_col}38" s="\d+"><f[^/]*/><v>[^<]*</v></c>',
        (f'<c r="{first_fc_col}38" s="31">'
         f'<f t="shared" ref="{first_fc_col}38:{last_col}38" ca="1" si="6">'
         f'SUM({first_fc_col}36:{first_fc_col}37)</f><v></v></c>'),
        rcfs_xml,
    )
    return rcfs_xml


def _patch_text_cell(xml: str, addr: str, new_text: str) -> str:
    """Replace cell text via inlineStr conversion (preserves xml[:lt])."""
    new_text = _xml_escape(new_text)
    search = f'r="{addr}"'
    start = 0
    while True:
        pos = xml.find(search, start)
        if pos == -1:
            print(f"  WARNING: Cell {addr} not found; skipping")
            return xml
        lt = xml.rfind("<", 0, pos)
        if lt == -1:
            start = pos + 1
            continue
        if xml[lt + 1] != "c" or xml[lt + 2] not in (" ", "\t", "\n", "/", ">"):
            start = pos + 1
            continue

        tag_end = xml.index(">", lt) + 1

        if xml[tag_end - 2: tag_end] == "/>":
            open_tag = xml[lt:tag_end - 2]
            open_tag = re.sub(r'\s+t="[^"]*"', '', open_tag)
            return (
                xml[:lt]
                + open_tag
                + f' t="inlineStr"><is><t>{new_text}</t></is></c>'
                + xml[tag_end:]
            )

        c_end = xml.index("</c>", tag_end)
        open_tag = xml[lt:tag_end - 1]
        open_tag = re.sub(r'\s+t="[^"]*"', '', open_tag)
        return (
            xml[:lt]
            + open_tag
            + f' t="inlineStr"><is><t>{new_text}</t></is></c>'
            + xml[c_end + 4:]
        )


def _replace_row_cells(xml: str, row: int, new_cells: str) -> str:
    """Replace F-X cells in a row; keep A-E cells intact."""
    pat = re.compile(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', re.DOTALL)
    m = pat.search(xml)
    if not m:
        print(f"  WARNING: row {row} not found")
        return xml

    row_tag = m.group(1)
    row_body = m.group(2)

    # Keep everything before the first F-column cell
    f_start = re.search(rf'<c r="F{row}"', row_body)
    if f_start:
        pre_f = row_body[: f_start.start()]
    else:
        pre_f = row_body

    new_row = row_tag + pre_f + new_cells + "</row>"
    return xml[: m.start()] + new_row + xml[m.end():]


def _patch_cell_formula(xml: str, addr: str, new_formula: str) -> str:
    """Replace a cell's formula, preserving style. Drops cached value."""
    search = f'r="{addr}"'
    start = 0
    while True:
        pos = xml.find(search, start)
        if pos == -1:
            print(f"  WARNING: Cell {addr} not found for formula patch")
            return xml
        lt = xml.rfind("<", 0, pos)
        if lt == -1:
            start = pos + 1
            continue
        if xml[lt + 1] != "c" or xml[lt + 2] not in (" ", "\t", "\n", "/", ">"):
            start = pos + 1
            continue

        tag_end = xml.index(">", lt) + 1
        # Get opening tag, strip t="..." and keep style
        if xml[tag_end - 2: tag_end] == "/>":
            open_tag = xml[lt:tag_end - 2]
        else:
            open_tag = xml[lt:tag_end - 1]
        open_tag = re.sub(r'\s+t="[^"]*"', '', open_tag)
        # Remove shared formula attributes
        open_tag = re.sub(r'\s+t="shared"', '', open_tag)

        # Find end of cell
        if xml[tag_end - 2: tag_end] == "/>":
            cell_end = tag_end
        else:
            c_end = xml.index("</c>", tag_end)
            cell_end = c_end + 4

        # Formula XML escaping: only & < > need escaping, NOT quotes
        escaped = (
            new_formula.replace("&", "&amp;")
                       .replace("<", "&lt;")
                       .replace(">", "&gt;")
        )
        return (
            xml[:lt]
            + open_tag
            + f"><f>{escaped}</f></c>"
            + xml[cell_end:]
        )


_SS_CACHE = []


def _resolve_shared_string(zf, idx: int) -> str:
    """Resolve a shared string index to text. Caches on first call."""
    if not _SS_CACHE:
        import xml.etree.ElementTree as ET
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ss_xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
        root = ET.fromstring(ss_xml)
        for si in root.findall(f"{{{ns}}}si"):
            texts = si.findall(f".//{{{ns}}}t")
            _SS_CACHE.append("".join(t.text or "" for t in texts))
    return _SS_CACHE[idx] if idx < len(_SS_CACHE) else ""


def _get_cell_text(xml: str, col: str, row: int, zf) -> str:
    """Read cell text from sheet XML (handles t='s' and inlineStr)."""
    m = re.search(rf'<c r="{col}{row}"[^/]*?(?:/>|>.*?</c>)', xml)
    if not m:
        return None
    cell = m.group()
    v_m = re.search(r'<v>(\d+)</v>', cell)
    if v_m and 't="s"' in cell:
        return _resolve_shared_string(zf, int(v_m.group(1)))
    if 't="inlineStr"' in cell:
        t_m = re.search(r'<t>([^<]*)</t>', cell)
        return t_m.group(1) if t_m else ""
    return None


def _discover_fydata_bs_items(zf) -> dict:
    """Read FY DATA K USD to discover all BS items and classify them.

    Returns {name: {"section": "asset"|"liability"|"equity", "is_sum": bool}}
    """
    xml = zf.read("xl/worksheets/sheet15.xml").decode("utf-8")
    items = {}
    section = "asset"  # start in asset section

    for row in range(50, 90):
        c = _get_cell_text(xml, "C", row, zf)
        d = _get_cell_text(xml, "D", row, zf)

        # Track section transitions from header rows
        if c is None and d:
            d_lower = d.lower()
            if "liabilit" in d_lower:
                section = "liability"
            elif "equity" in d_lower or "stockholder" in d_lower:
                section = "equity"
            continue

        if c != "BS" or not d:
            continue

        # Check if it's a SUM/total row
        f_m = re.search(rf'<c r="F{row}"[^/]*?(?:/>|>.*?</c>)', xml)
        is_sum = False
        if f_m:
            ff = re.search(r'<f[^>]*>([^<]*)</f>', f_m.group())
            if ff and "SUM" in ff.group(1).upper():
                is_sum = True

        items[d] = {"section": section, "is_sum": is_sum}

    return items


def _read_rbs_d_names(rbs_xml: str, zf) -> dict:
    """Read current D-column names from RBS sheet. Returns {row: name}."""
    names = {}
    for row in list(range(11, 17)) + list(range(20, 28)) + [33] + list(range(41, 45)):
        name = _get_cell_text(rbs_xml, "D", row, zf)
        if name:
            names[row] = name
    return names


def _build_formula_cells(ris_row: int, pipe_row: int, style: str = "174") -> str:
    """Build INDEX/MATCH formula cells for F-X (cols 6-24) in a RIS row."""
    cells = []
    for col_num in range(6, 25):  # F=6 through X=24
        col = _col_letter(col_num)
        addr = f"{col}{ris_row}"
        formula = (
            f"IFERROR(INDEX(Pipeline!$F${pipe_row}:$U${pipe_row},"
            f" MATCH({col}$5, Pipeline!$F$7:$U$7, 0)), 0)"
        )
        cells.append(f'<c r="{addr}" s="{style}"><f>{formula}</f></c>')
    return "".join(cells)


# ── Main ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Fix DCF CMPX financials")
    parser.add_argument("--dry-run", action="store_true", help="Show changes without writing")
    args = parser.parse_args()

    if not XLSX.exists():
        print(f"ERROR: {XLSX} not found")
        return

    # Backup
    bak = XLSX.with_suffix(".xlsx.bak")
    if not args.dry_run:
        shutil.copy2(XLSX, bak)
        print(f"Backup: {bak}")

    modified = {}

    with zipfile.ZipFile(XLSX, "r") as zf:
        # ── RIS ──
        ris_xml = zf.read(SHEET_RIS).decode("utf-8")
        ris_rows_before = len(re.findall(r"<row ", ris_xml))

        # Phase 1: Text patches
        for addr, text in RIS_TEXT.items():
            ris_xml = _patch_text_cell(ris_xml, addr, text)
            print(f"  RIS {addr} → {text[:60]}")

        # Phase 2: Revenue formulas (R26-R29)
        for ris_row, pipe_row in REV_MAP.items():
            new_cells = _build_formula_cells(ris_row, pipe_row)
            ris_xml = _replace_row_cells(ris_xml, ris_row, new_cells)
            print(f"  RIS R{ris_row}: Revenue → INDEX/MATCH Pipeline R{pipe_row}")

        # Phase 2: COGS formulas (R37-R40)
        for ris_row, pipe_row in COGS_MAP.items():
            new_cells = _build_formula_cells(ris_row, pipe_row)
            ris_xml = _replace_row_cells(ris_xml, ris_row, new_cells)
            print(f"  RIS R{ris_row}: COGS → INDEX/MATCH Pipeline R{pipe_row}")

        # Phase 3: Fix R61 (G&A total) — SUM → SUMIF to exclude ratio row R65
        for col_num in range(6, 25):  # F=6 through X=24
            col = _col_letter(col_num)
            new_f = f'SUMIF($C$62:$C$67,"ISN",{col}62:{col}67)'
            ris_xml = _patch_cell_formula(ris_xml, f"{col}61", new_f)
        print("  RIS R61: SUM(F62:F67) → SUMIF (exclude ratio row R65)")

        # Phase 3: Fix R65 (G&A/R&D ratio) — wrap with IFERROR
        for col_num in range(6, 11):  # F-J: individual ratio formulas
            col = _col_letter(col_num)
            new_f = f"IFERROR({col}64/{col}46,0)"
            ris_xml = _patch_cell_formula(ris_xml, f"{col}65", new_f)
        # K65: AVERAGE
        ris_xml = _patch_cell_formula(ris_xml, "K65", "IFERROR(AVERAGE(I65:J65),0)")
        # L65-X65: shared chain K65→L65→...→X65 (already propagates IFERROR result)
        # But the shared formula ref must be replaced with explicit formulas
        for col_num in range(12, 25):  # L=12 through X=24
            col = _col_letter(col_num)
            prev = _col_letter(col_num - 1)
            ris_xml = _patch_cell_formula(ris_xml, f"{col}65", f"{prev}65")
        print("  RIS R65: Wrapped with IFERROR to prevent #DIV/0!")

        ris_rows_after = len(re.findall(r"<row ", ris_xml))
        print(f"  RIS row count: {ris_rows_before} → {ris_rows_after}")
        modified[SHEET_RIS] = ris_xml.encode("utf-8")

        # ── Schedules ──
        sched_xml = zf.read(SHEET_SCHED).decode("utf-8")
        sched_rows_before = len(re.findall(r"<row ", sched_xml))

        for addr, text in SCHED_TEXT.items():
            sched_xml = _patch_text_cell(sched_xml, addr, text)
            print(f"  Sched {addr} → {text[:60]}")

        sched_rows_after = len(re.findall(r"<row ", sched_xml))
        print(f"  Sched row count: {sched_rows_before} → {sched_rows_after}")
        modified[SHEET_SCHED] = sched_xml.encode("utf-8")

        # ── RBS — fix BS item mapping to match FY DATA ──
        rbs_xml = zf.read(SHEET_RBS).decode("utf-8")
        rbs_rows_before = len(re.findall(r"<row ", rbs_xml))

        # Read FY DATA to discover actual BS items
        fydata_bs = _discover_fydata_bs_items(zf)

        # Classify FY DATA BS items into RBS sections
        FINANCIAL_ASSETS = {"Cash And Cash Equivalents"}
        EQUITY_ITEMS = {
            "Ordinary Shares, £0.01 Nominal Value",
            "Additional Paid-In Capital",
            "Accumulated Deficit",
            "Accumulated Other Comprehensive (Loss) Income",
        }
        op_assets = []
        op_liabs = []
        for name, info in fydata_bs.items():
            if info["is_sum"]:
                continue
            if name in FINANCIAL_ASSETS or name in EQUITY_ITEMS:
                continue
            if name == "Property And Equipment, Net":
                continue  # Comes from Schedules
            if info["section"] == "asset":
                op_assets.append(name)
            else:
                op_liabs.append(name)

        print(f"  RBS Op Assets: {op_assets}")
        print(f"  RBS Op Liabs:  {op_liabs}")

        # Read current RBS D-column names
        rbs_current = _read_rbs_d_names(rbs_xml, zf)

        # Patch Operating Assets (R12-R16)
        asset_slots = [12, 13, 14, 15, 16]
        for i, row in enumerate(asset_slots):
            new_name = op_assets[i] if i < len(op_assets) else "Reserved"
            old_name = rbs_current.get(row, "")
            if old_name != new_name:
                rbs_xml = _patch_text_cell(rbs_xml, f"D{row}", new_name)
                print(f"  RBS D{row}: '{old_name}' → '{new_name}'")

        # Patch Operating Liabilities (R20-R27)
        liab_slots = [20, 21, 22, 23, 24, 25, 26, 27]
        for i, row in enumerate(liab_slots):
            new_name = op_liabs[i] if i < len(op_liabs) else "Reserved"
            old_name = rbs_current.get(row, "")
            if old_name != new_name:
                rbs_xml = _patch_text_cell(rbs_xml, f"D{row}", new_name)
                print(f"  RBS D{row}: '{old_name}' → '{new_name}'")

        # Fix formulas for rows that previously referenced RIS
        # (Deferred revenue rows) → must now reference FY DATA
        rbs_ris_rows = set()
        for row in liab_slots:
            old_name = rbs_current.get(row, "")
            if old_name in ("Deferred revenue, Current Portion",
                            "Deferred Revenue, Net Of Current Portion"):
                rbs_ris_rows.add(row)

        if rbs_ris_rows:
            FY_COLS = "GHIJ"  # FY DATA cols for RBS F-I (offset +1)
            for row in rbs_ris_rows:
                # Historical: F-I → SUMIFS from FY DATA
                for col_idx, rbs_col in enumerate("FGHI"):
                    fy_col = FY_COLS[col_idx]
                    formula = (f"SUMIFS('FY DATA'!{fy_col}:{fy_col},"
                               f"'FY DATA'!$D:$D,$D{row},"
                               f"'FY DATA'!$C:$C,$C{row},"
                               f"'FY DATA'!$B:$B,$B{row})")
                    rbs_xml = _patch_cell_formula(rbs_xml, f"{rbs_col}{row}", formula)
                # Forecast: J → AVERAGE, K-W → carry-forward
                rbs_xml = _patch_cell_formula(
                    rbs_xml, f"J{row}", f"AVERAGE(F{row}:I{row})")
                for col_num in range(11, 24):  # K=11..W=23
                    col = _col_letter(col_num)
                    prev = _col_letter(col_num - 1)
                    rbs_xml = _patch_cell_formula(
                        rbs_xml, f"{col}{row}", f"{prev}{row}")
                print(f"  RBS R{row}: formulas → FY DATA SUMIFS + forecast")

        rbs_rows_after = len(re.findall(r"<row ", rbs_xml))
        print(f"  RBS row count: {rbs_rows_before} → {rbs_rows_after}")
        modified[SHEET_RBS] = rbs_xml.encode("utf-8")

        # ── Patch RCFS R38 Ending Cash for historical cols beyond H ──
        rcfs_xml = zf.read(SHEET_RCFS).decode("utf-8")
        rcfs_xml_patched = _patch_rcfs_cash_formulas(rcfs_xml, HIST_COLS)
        if rcfs_xml_patched != rcfs_xml:
            modified[SHEET_RCFS] = rcfs_xml_patched.encode("utf-8")
            print(f"  RCFS R38: fixed Ending Cash for historical cols I..{_col_letter(5+HIST_COLS)}")
        else:
            print(f"  RCFS R38: no changes needed (n_hist={HIST_COLS} <= 3 or already fixed)")

        # ── fullCalcOnLoad ──
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        if "fullCalcOnLoad" not in wb_xml:
            wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
            print("  Added fullCalcOnLoad to workbook.xml")
        modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    if args.dry_run:
        print("\n  DRY RUN — no file written")
        return

    # ── Remove calcChain.xml: strip its references from Content_Types + workbook.rels ──
    # An empty/absent calcChain causes "Catastrophic failure" in Excel's XML parser.
    # Correct fix: delete the file AND remove its <Override>/<Relationship> entries
    # so Excel doesn't look for it at all.
    with zipfile.ZipFile(XLSX, "r") as zf:
        ct = zf.read("[Content_Types].xml").decode("utf-8")
        wr = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', '', ct)
    wr = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', wr)
    modified["[Content_Types].xml"] = ct.encode("utf-8")
    modified["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")
    print("  Removed calcChain.xml references from Content_Types + workbook.rels")

    # ── Zip patch ──
    tmp = XLSX.with_suffix(".~patch.xlsx")
    with zipfile.ZipFile(XLSX, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "xl/calcChain.xml":
                continue  # omitted — references already stripped from CT + rels
            if item.filename in modified:
                raw = modified[item.filename]
                if (item.filename.startswith("xl/worksheets/sheet")
                        and item.filename.endswith(".xml")):
                    s = raw.decode("utf-8")
                    s = _strip_formula_cache(s)
                    s = _normalize_formula_xml(s)
                    raw = s.encode("utf-8")
                data = raw
            elif (item.filename.startswith("xl/worksheets/sheet")
                  and item.filename.endswith(".xml")):
                s = zin.read(item.filename).decode("utf-8")
                s = _strip_formula_cache(s)
                s = _normalize_formula_xml(s)
                data = s.encode("utf-8")
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)

    tmp.replace(XLSX)
    sz = XLSX.stat().st_size
    print(f"\n  Written {sz:,} bytes → {XLSX}")


if __name__ == "__main__":
    main()
