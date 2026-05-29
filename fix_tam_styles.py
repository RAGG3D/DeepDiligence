#!/usr/bin/env python3
"""
fix_tam_styles.py -- Fix formatting of new indication sections in TAM Solid.

1. Re-insert missing EC section header (shift rows >= 402 by +1)
2. Apply proper styles matching the reference section (R356-R391):
   - Section headers: black font + bottom thin border (s=73/46)
   - Drug name rows: black font (s=75)
   - Drug data rows: add E column [MM USD], fill empty cols
   - Total rows: bold black font + top thin border (s=481/482/483)

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
from typing import Dict, Optional


_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

_DEFAULT_FILE = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Row regex — handles both content and self-closing rows
_ROW_RE = r'<row\b[^/>]*r="(\d+)"[^/>]*(?:>.*?</row>|/>)'


def _get_sheet_zip_path(xlsx_path: Path, sheet_name: str) -> Optional[str]:
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


def _col_letter(ci: int) -> str:
    r = ""
    while ci > 0:
        ci -= 1
        r = chr(ord('A') + ci % 26) + r
        ci //= 26
    return r


# ══════════════════════════════════════════════════════════════════════════════
#  D-CELL CONTENT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def _extract_d_content(row_xml: str) -> Optional[str]:
    """Extract D-cell 'tail' — everything between s="..." and </c>.

    Returns a string to be used as:
        f'<c r="D{row}" s="73"{tail}</c>'

    Examples:
        ' t="s"><v>1045</v>'   → <c r="D1" s="73" t="s"><v>1045</v></c>
        '><v>123</v>'          → <c r="D1" s="73"><v>123</v></c>
        ''                     → <c r="D1" s="73"/>  (self-closing)
    """
    # Self-closing D cell (no content)
    if re.search(r'<c r="D\d+" s="\d+"/>', row_xml):
        return ''
    # D cell with content
    m = re.search(r'<c r="D\d+" s="\d+"(.*?)</c>', row_xml, re.DOTALL)
    if m:
        return m.group(1)
    return None


def _d_cell(row: int, style: int, d_content: str) -> str:
    """Build a D column cell with given style and extracted content."""
    if not d_content:
        return f'<c r="D{row}" s="{style}"/>'
    return f'<c r="D{row}" s="{style}"{d_content}</c>'


# ══════════════════════════════════════════════════════════════════════════════
#  ROW BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

def _build_section_header(row: int, d_content: str) -> str:
    """Section header: D=s73 (black, bottom thin), E-AH=s46 (bottom thin)."""
    parts = [f'<row r="{row}" spans="1:34" x14ac:dyDescent="0.25">']
    parts.append(f'<c r="A{row}" s="2"/>')
    parts.append(f'<c r="C{row}" s="11"/>')
    parts.append(_d_cell(row, 73, d_content))
    for ci in range(5, 35):  # E=5 through AH=34
        parts.append(f'<c r="{_col_letter(ci)}{row}" s="46"/>')
    parts.append('</row>')
    return ''.join(parts)


def _build_drug_name_row(row: int, d_content: str) -> str:
    """Drug name row: D=s75 (black font), minimal cells."""
    return (f'<row r="{row}" spans="4:4" x14ac:dyDescent="0.25">'
            + _d_cell(row, 75, d_content)
            + '</row>')


def _build_data_row(row: int, d_content: str, data_cells: Dict[str, str]) -> str:
    """Drug data row: D=s48, E=s52 [MM USD], data s=48, projection s=408."""
    parts = [f'<row r="{row}" spans="1:34" x14ac:dyDescent="0.25">']
    parts.append(f'<c r="A{row}" s="123"/>')
    parts.append(f'<c r="C{row}" s="11"/>')
    parts.append(_d_cell(row, 48, d_content))
    parts.append(f'<c r="E{row}" s="52" t="s"><v>150</v></c>')
    for ci in range(6, 35):  # F=6 through AH=34
        col = _col_letter(ci)
        if col in data_cells:
            parts.append(f'<c r="{col}{row}" s="48"><v>{data_cells[col]}</v></c>')
        elif ci <= 20:  # F-T
            parts.append(f'<c r="{col}{row}" s="48"/>')
        else:  # U-AH
            parts.append(f'<c r="{col}{row}" s="408"/>')
    parts.append('</row>')
    return ''.join(parts)


def _build_total_row(row: int, d_content: str, formula_cells: str) -> str:
    """Total row: row-level s=54, D=s481, E=s482, data=s483."""
    parts = [f'<row r="{row}" spans="1:34" s="54" customFormat="1" x14ac:dyDescent="0.25">']
    parts.append(f'<c r="A{row}" s="528"/>')
    parts.append(f'<c r="B{row}" s="3"/>')
    parts.append(f'<c r="C{row}" s="117"/>')
    parts.append(_d_cell(row, 481, d_content))
    parts.append(f'<c r="E{row}" s="482" t="s"><v>150</v></c>')
    parts.append(formula_cells)
    parts.append('</row>')
    return ''.join(parts)


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _extract_data_values(row_xml: str) -> Dict[str, str]:
    """Extract non-D, non-E column values from a data row."""
    values = {}
    for m in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*><v>([^<]*)</v></c>', row_xml):
        col = m.group(1)
        if col not in ("D", "E"):
            values[col] = m.group(2)
    return values


def _extract_formula_cells(row_xml: str) -> str:
    """Extract formula cells from Total row, restyled to s=483."""
    cells = []
    for m in re.finditer(
            r'<c r="([A-Z]+)(\d+)"[^>]*>(<f[^<]*(?:</f>|/>)(?:<v>[^<]*</v>)?)</c>',
            row_xml):
        col, row, content = m.group(1), m.group(2), m.group(3)
        if col not in ("D", "E"):
            cells.append(f'<c r="{col}{row}" s="483">{content}</c>')
    return ''.join(cells)


def _classify_row(row_xml: str, ss_list: list) -> str:
    """Classify a row as SEC-HDR, DRUG-NAME, DATA, TOTAL, or OTHER."""
    d_style_m = re.search(r'<c r="D\d+"[^>]*s="(\d+)"', row_xml)
    ds = d_style_m.group(1) if d_style_m else "?"

    dt = ""
    dm = re.search(r'<c r="D\d+"[^>]*t="s"[^>]*><v>(\d+)</v>', row_xml)
    if dm:
        idx = int(dm.group(1))
        dt = ss_list[idx] if idx < len(ss_list) else ""

    if dt.startswith("Total "):
        return "TOTAL"
    section_names = {"BTC", "EC", "ES-SCLC", "HCC", "Melanoma NCAM+", "RCC"}
    if ds == "21" and dt in section_names:
        return "SEC-HDR"
    if ds == "21" and "SUM(" not in row_xml:
        return "DRUG-NAME"
    if ds == "48" and "SUM(" not in row_xml:
        return "DATA"
    return "OTHER"


# ══════════════════════════════════════════════════════════════════════════════
#  ROW SHIFTING (for EC header insertion)
# ══════════════════════════════════════════════════════════════════════════════

def _shift_rows_from(xml: str, from_row: int, delta: int) -> str:
    """Shift all rows >= from_row by delta. Updates row/cell refs + formulas."""

    # Step 1: Shift row tag r= attribute
    def shift_row(m):
        rn = int(m.group(1))
        if rn >= from_row:
            full = m.group(0)
            # Replace only the <row r="N"> attribute (first r="N" in string)
            return full.replace(f'r="{rn}"', f'r="{rn + delta}"', 1)
        return m.group(0)

    xml = re.sub(_ROW_RE, shift_row, xml)

    # Step 2: Update cell r= attributes within shifted rows
    def fix_cells(m):
        row_xml = m.group(0)
        rm = re.search(r'<row\b[^/>]*r="(\d+)"', row_xml)
        if not rm:
            return row_xml
        rn = int(rm.group(1))
        if rn < from_row + delta:
            return row_xml  # Not a shifted row
        old_rn = rn - delta
        # Replace r="COL{old}" with r="COL{new}" for cell references
        row_xml = re.sub(
            r'(r="[A-Z]+)' + str(old_rn) + '"',
            lambda cm: cm.group(1) + str(rn) + '"',
            row_xml)
        return row_xml

    xml = re.sub(_ROW_RE, fix_cells, xml)

    # Step 3: Shift formula references in ALL formulas
    def shift_ref(m):
        col, rn = m.group(1), int(m.group(2))
        if rn >= from_row:
            return f'{col}{rn + delta}'
        return m.group(0)

    # Regular formulas: <f>BODY</f>
    def shift_f(m):
        body = m.group(1)
        return '<f>' + re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', shift_ref, body) + '</f>'

    xml = re.sub(r'<f>([^<]+)</f>', shift_f, xml)

    # Shared formulas: <f attrs>BODY</f>
    def shift_fa(m):
        attrs, body = m.group(1), m.group(2)
        # Shift ref= endpoints
        def shift_ref_attr(rm):
            ref = rm.group(1)
            new_ref = re.sub(
                r'([A-Z]{1,3})(\d+)',
                lambda x: (f'{x.group(1)}{int(x.group(2)) + delta}'
                           if int(x.group(2)) >= from_row else x.group(0)),
                ref)
            return f'ref="{new_ref}"'

        attrs = re.sub(r'ref="([^"]*)"', shift_ref_attr, attrs)
        if body:
            body = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)', shift_ref, body)
            return f'<f {attrs}>{body}</f>'
        return f'<f {attrs}/>'

    xml = re.sub(r'<f\s+([^>]*)>([^<]*)</f>', shift_fa, xml)

    # Self-closing shared formulas: <f attrs/> (no ref= to shift, no body)
    # These are handled implicitly — no body refs and ref= only on master cells

    # Update dimension
    xml = re.sub(
        r'(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)"',
        lambda m: f'{m.group(1)}{int(m.group(2)) + delta}"', xml)

    return xml


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PROCESSING
# ══════════════════════════════════════════════════════════════════════════════

def process_sheet(xml: str, ss_list: list) -> str:
    """Insert EC header + apply style fixes to new section rows."""

    # ── Phase 1: Insert missing EC section header ────────────────────────
    # Find Total BTC by shared string lookup
    ss_total_btc = None
    for i, s in enumerate(ss_list):
        if s == "Total BTC":
            ss_total_btc = i
            break

    if ss_total_btc is not None:
        total_btc_m = re.search(
            rf'<row\b[^/>]*r="(\d+)"[^/>]*>.*?<v>{ss_total_btc}</v>.*?</row>', xml)
    else:
        total_btc_m = None

    if total_btc_m:
        insert_after = int(total_btc_m.group(1))
        insert_at = insert_after + 1
        log.info(f"Phase 1: Insert EC header at R{insert_at} "
                 f"(after Total BTC R{insert_after})")

        # Shift all rows >= insert_at by +1
        xml = _shift_rows_from(xml, insert_at, 1)

        # Find shared string index for "EC"
        ec_ss = None
        for i, s in enumerate(ss_list):
            if s == "EC":
                ec_ss = i
                break

        if ec_ss is not None:
            ec_d_content = f' t="s"><v>{ec_ss}</v>'
        else:
            ec_d_content = ' t="inlineStr"><is><t>EC</t></is>'

        ec_header = _build_section_header(insert_at, ec_d_content)

        # Insert after Total BTC row (still at insert_after, not shifted)
        btc_row_m = re.search(
            rf'(<row\b[^/>]*r="{insert_after}"[^/>]*>.*?</row>)', xml)
        if btc_row_m:
            pos = btc_row_m.end()
            xml = xml[:pos] + ec_header + xml[pos:]
            log.info(f"  EC header inserted at R{insert_at}")
        else:
            log.warning("  Could not locate Total BTC row for insertion")
    else:
        log.warning("Phase 1: Total BTC not found, skipping EC header")

    # ── Phase 2: Apply style changes ─────────────────────────────────────
    log.info("Phase 2: Applying style changes")

    rows_to_fix = []
    for m in re.finditer(rf'({_ROW_RE})', xml):
        rn = int(m.group(2))
        rx = m.group(1)
        # Only process new-section rows (after Total Solid Tumor)
        if rn <= 391:
            continue
        typ = _classify_row(rx, ss_list)
        if typ == "OTHER":
            continue
        # Verify it's a new-section row (s=21 with few cells, or s=48 data)
        n_cells = len(re.findall(r'<c r=', rx))
        d_style_m = re.search(r'<c r="D\d+"[^>]*s="(\d+)"', rx)
        ds = d_style_m.group(1) if d_style_m else "?"
        if ds == "21" and n_cells <= 35:
            rows_to_fix.append((rn, rx, typ))
        elif ds == "48" and n_cells <= 15:
            rows_to_fix.append((rn, rx, typ))

    if not rows_to_fix:
        log.warning("  No rows found to fix")
        return xml

    log.info(f"  Found {len(rows_to_fix)} rows to fix "
             f"(R{rows_to_fix[0][0]}-R{rows_to_fix[-1][0]})")

    # Apply fixes (reverse order to preserve string positions)
    for rn, rx, typ in reversed(rows_to_fix):
        d_content = _extract_d_content(rx)
        if d_content is None:
            log.warning(f"  R{rn}: cannot extract D content, skipping")
            continue

        if typ == "SEC-HDR":
            new_rx = _build_section_header(rn, d_content)
            log.info(f"  R{rn}: SEC-HDR -> s=73/46")
        elif typ == "DRUG-NAME":
            new_rx = _build_drug_name_row(rn, d_content)
            log.info(f"  R{rn}: DRUG-NAME -> s=75")
        elif typ == "DATA":
            data_vals = _extract_data_values(rx)
            new_rx = _build_data_row(rn, d_content, data_vals)
            log.info(f"  R{rn}: DATA -> +E col")
        elif typ == "TOTAL":
            formula_cells = _extract_formula_cells(rx)
            new_rx = _build_total_row(rn, d_content, formula_cells)
            log.info(f"  R{rn}: TOTAL -> s=481/483")
        else:
            continue

        xml = xml.replace(rx, new_rx, 1)

    return xml


def main():
    parser = argparse.ArgumentParser(
        description="Fix TAM Solid new section row styles")
    parser.add_argument(
        "--file", default=str(_DEFAULT_FILE),
        help=f"Excel file (default: {_DEFAULT_FILE})")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview changes without writing")
    args = parser.parse_args()

    xlsx_path = Path(args.file)
    if not xlsx_path.exists():
        log.error(f"File not found: {xlsx_path}")
        return

    sheet_zip = None
    for name in ["TAM Solid", "TAM Solid+MM"]:
        sheet_zip = _get_sheet_zip_path(xlsx_path, name)
        if sheet_zip:
            log.info(f"Found sheet '{name}' -> {sheet_zip}")
            break
    if not sheet_zip:
        log.error("TAM Solid sheet not found")
        return

    # Read shared strings
    with zipfile.ZipFile(xlsx_path) as zf:
        xml = zf.read(sheet_zip).decode("utf-8")
        ss_xml = zf.read("xl/sharedStrings.xml").decode("utf-8")

    ss_list = []
    for m in re.finditer(r'<si>(.*?)</si>', ss_xml, re.DOTALL):
        texts = re.findall(r'<t[^>]*>([^<]*)</t>', m.group(1))
        ss_list.append(''.join(texts))

    log.info(f"Sheet XML: {len(xml):,} bytes, {len(ss_list)} shared strings")

    new_xml = process_sheet(xml, ss_list)

    if args.dry_run:
        old_rows = len(re.findall(r'<row\b', xml))
        new_rows = len(re.findall(r'<row\b', new_xml))
        log.info(f"[DRY-RUN] Rows: {old_rows} -> {new_rows}")
        # Show a sample of changes
        for m in re.finditer(_ROW_RE, new_xml):
            rn = int(m.group(1))
            if 392 <= rn <= 450:
                rx = m.group(0)
                ds_m = re.search(r'<c r="D\d+"[^>]*s="(\d+)"', rx)
                ds = ds_m.group(1) if ds_m else "?"
                nc = len(re.findall(r'<c r=', rx))
                has_e = '<c r="E' in rx
                print(f"  R{rn}: D(s={ds}), {nc} cells, E={has_e}")
        return

    # Backup
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = xlsx_path.with_name(f"{xlsx_path.stem}_pre_style_fix_{ts}.xlsx")
    shutil.copy2(xlsx_path, backup)
    log.info(f"Backup: {backup}")

    # Surgical zip patch
    modified = {sheet_zip: new_xml.encode("utf-8")}

    with zipfile.ZipFile(xlsx_path) as zf:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    modified["xl/workbook.xml"] = wb_xml.encode("utf-8")

    tmp = xlsx_path.with_suffix(".~style_fix.xlsx")
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

    log.info(f"Saved -> {xlsx_path}")
    print(f"\n{'='*60}")
    print("TAM Style Fix Complete")
    print(f"  File: {xlsx_path}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
