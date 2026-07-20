#!/usr/bin/env python3
"""Repair FSA/RIS/Schedules formula links that are not pure formatting issues.

The generated model keeps the original template's FSA structure, but several
template links can become stale after ticker-specific pipeline/financial tabs
are generated:
  * RIS has a Cost of Sales label row but no matching C/E/formulas.
  * Schedules depreciation can retain a #REF! from the template PPE rollforward.
  * FSA EBITDA subtracts D&A instead of adding it back.
  * A few FSA ratio first-year formulas and labels are inconsistent.
  * RCFS terminal discount-factor formula can retain a #REF! branch.

This script patches formulas directly in OOXML and never saves with openpyxl.
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _default_path(ticker: str) -> Path:
    return Path(f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx")


def _sheet_zip_path(parts: dict[str, bytes], sheet_name: str) -> str:
    wb = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    relmap = {rel.get("Id"): rel.get("Target") for rel in rels}
    for sheet in wb.findall(f".//{{{NS_MAIN}}}sheet"):
        if sheet.get("name") == sheet_name:
            target = relmap[sheet.get(f"{{{NS_R}}}id")]
            target = target.lstrip("/")
            return target if target.startswith("xl/") else "xl/" + target
    raise RuntimeError(f"Sheet not found: {sheet_name}")


def _read_parts(path: Path) -> tuple[dict[str, bytes], list[str]]:
    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"zip CRC failure in {bad}")
        return {name: zf.read(name) for name in zf.namelist()}, zf.namelist()


def _write_parts(path: Path, parts: dict[str, bytes], order: list[str]) -> None:
    tmp = path.with_suffix(".~fsa.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for name in order:
            if name == "xl/calcChain.xml":
                continue
            if name in parts:
                zout.writestr(name, parts[name])
                written.add(name)
        for name, data in parts.items():
            if name not in written and name != "xl/calcChain.xml":
                zout.writestr(name, data)
    tmp.replace(path)


def _cell_col(addr: str) -> int:
    letters = re.match(r"([A-Z]+)", addr).group(1)
    out = 0
    for ch in letters:
        out = out * 26 + ord(ch) - 64
    return out


def _col_letter(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def _cell_style(xml: str, addr: str, default: str = "0") -> str:
    m = re.search(
        rf'<c\b(?=[^>]*\br="{addr}")[^>]*/>|'
        rf'<c\b(?=[^>]*\br="{addr}")[^>]*>.*?</c>',
        xml,
        re.S,
    )
    if not m:
        return default
    sm = re.search(r'\bs="(\d+)"', m.group(0))
    return sm.group(1) if sm else default


def _replace_cell(xml: str, addr: str, cell_xml: str) -> tuple[str, bool]:
    m = re.search(
        rf'<c\b(?=[^>]*\br="{addr}")[^>]*/>|'
        rf'<c\b(?=[^>]*\br="{addr}")[^>]*>.*?</c>',
        xml,
        re.S,
    )
    if m:
        return xml[:m.start()] + cell_xml + xml[m.end():], True

    row_num = int(re.search(r"\d+", addr).group(0))
    row_m = re.search(rf'(<row\b(?=[^>]*\br="{row_num}")[^>]*>)(.*?)(</row>)', xml, re.S)
    if not row_m:
        return xml, False

    body = row_m.group(2)
    insert_at = len(body)
    target_col = _cell_col(addr)
    for cm in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"', body):
        if _cell_col(cm.group(1)) > target_col:
            insert_at = cm.start()
            break
    new_body = body[:insert_at] + cell_xml + body[insert_at:]
    return xml[:row_m.start(2)] + new_body + xml[row_m.end(2):], True


def _text_cell(addr: str, text: str, style: str) -> str:
    return f'<c r="{addr}" s="{style}" t="inlineStr"><is><t>{html.escape(text, quote=False)}</t></is></c>'


def _formula_cell(addr: str, formula: str, style: str) -> str:
    return f'<c r="{addr}" s="{style}"><f>{html.escape(formula, quote=False)}</f></c>'


def _patch_text(xml: str, addr: str, text: str, default_style: str = "0") -> tuple[str, bool]:
    style = _cell_style(xml, addr, default_style)
    return _replace_cell(xml, addr, _text_cell(addr, text, style))


def _patch_formula(xml: str, addr: str, formula: str, default_style: str = "0") -> tuple[str, bool]:
    style = _cell_style(xml, addr, default_style)
    return _replace_cell(xml, addr, _formula_cell(addr, formula, style))


def _mark_full_calc(parts: dict[str, bytes]) -> None:
    wb_xml = parts["xl/workbook.xml"].decode("utf-8", "ignore")
    if "fullCalcOnLoad" not in wb_xml:
        wb_xml = wb_xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    parts["xl/workbook.xml"] = wb_xml.encode("utf-8")
    ct = parts.get("[Content_Types].xml", b"").decode("utf-8", "ignore")
    if ct:
        ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', "", ct)
        parts["[Content_Types].xml"] = ct.encode("utf-8")
    rels = parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "ignore")
    if rels:
        rels = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", rels)
        parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    parts.pop("xl/calcChain.xml", None)


def _patch_ris(xml: str) -> tuple[str, int]:
    changed = 0
    for addr, text in {"C36": "IS", "E36": "[MM USD]"}.items():
        xml, ok = _patch_text(xml, addr, text)
        changed += int(ok)
    for col_idx in range(_cell_col("F1"), _cell_col("Y1") + 1):
        col = _col_letter(col_idx)
        xml, ok = _patch_formula(xml, f"{col}36", f"SUM({col}37:{col}40)")
        changed += int(ok)
    return xml, changed


def _patch_schedules(xml: str) -> tuple[str, int]:
    """Historical D&A on Schedules row 34, derived generically.

    Expense-nature filers (MOLN) disclose a "Depreciation and amortization" ISN
    note row; program-based filers (BHVN) have no such row and only report a BSN
    "Accumulated Depreciation" stock.  Keying on the literal MOLN label left BHVN
    with D&A = 0 every year.  Instead match any FY DATA D-label containing
    "Deprec": prefer the ISN expense flow when present (MOLN, unchanged) and
    otherwise fall back to the year-over-year increase in the (negative) BSN
    Accumulated Depreciation balance (BHVN).  ISN and BSN are kept on separate
    terms so MOLN's flow is never double-counted with the accumulated stock.
    """
    changed = 0
    first_col = _cell_col("F1")
    for col_idx in range(_cell_col("F1"), _cell_col("J1") + 1):
        col = _col_letter(col_idx)
        prev = _col_letter(col_idx - 1)
        isn_flow = (
            f'SUMIFS(\'FY DATA\'!{col}:{col},\'FY DATA\'!$D:$D,'
            f'"*Deprec*",\'FY DATA\'!$C:$C,"ISN")'
        )
        if col_idx > first_col:
            # Accumulated Depreciation is stored negative, so annual depreciation
            # is prior-year minus current-year (a positive number).
            accdep_yoy = (
                f'MAX(0,'
                f'SUMIFS(\'FY DATA\'!{prev}:{prev},\'FY DATA\'!$D:$D,"*Deprec*",\'FY DATA\'!$C:$C,"BSN")'
                f'-SUMIFS(\'FY DATA\'!{col}:{col},\'FY DATA\'!$D:$D,"*Deprec*",\'FY DATA\'!$C:$C,"BSN"))'
            )
            formula = f'IFERROR(IF({isn_flow}<>0,{isn_flow},{accdep_yoy}),0)'
        else:
            formula = f'IFERROR({isn_flow},0)'
        xml, ok = _patch_formula(xml, f"{col}34", formula)
        changed += int(ok)
    return xml, changed


def _patch_fsa(xml: str) -> tuple[str, int]:
    changed = 0
    xml, ok = _patch_text(xml, "D74", "Forward Price to Sales (P/S) Ratio")
    changed += int(ok)

    for col_idx in range(_cell_col("H1"), _cell_col("Y1") + 1):
        col = _col_letter(col_idx)
        prev = _col_letter(col_idx - 1)
        rbs_col = _col_letter(col_idx - 3)  # existing FSA H->RBS E, I->RBS F offset

        patches = {
            f"{col}22": (
                f'SUMIFS(RBS!{rbs_col}:{rbs_col},RBS!$D:$D,"Debt, Current Portion",RBS!$C:$C,"BS")'
                f'+SUMIFS(RBS!{rbs_col}:{rbs_col},RBS!$D:$D,"Long-Term Debt, Net Of Discount",RBS!$C:$C,"BS")'
            ),
            f"{col}36": f"{col}11+{col}16",
            f"{col}56": f"IFERROR({col}8/AVERAGE({prev}33:{col}33),0)",
            f"{col}57": f"IFERROR({col}9/AVERAGE({prev}19:{col}19),0)",
            f"{col}58": f"IFERROR({col}9/AVERAGE({prev}33:{col}33),0)",
        }
        for addr, formula in patches.items():
            xml, ok = _patch_formula(xml, addr, formula)
            changed += int(ok)
    return xml, changed


def _patch_rcfs(xml: str) -> tuple[str, int]:
    changed = 0
    for col_idx in range(_cell_col("G1"), _cell_col("W1") + 1):
        col = _col_letter(col_idx)
        schedules_col = _col_letter(col_idx + 1)
        xml, ok = _patch_formula(
            xml,
            f"{col}7",
            (
                f'IFERROR(SUMIFS(Schedules!{schedules_col}:{schedules_col},'
                f'Schedules!$D:$D,RCFS!$D7,Schedules!$C:$C,RCFS!$C7),0)'
            ),
        )
        changed += int(ok)
    xml, ok = _patch_formula(
        xml,
        "W67",
        'IF(W$66<>"",IF(YEAR($F$67)=YEAR(W$66),MAX(W$66-$F$67,0)/365.25,V67+1),"")',
    )
    changed += int(ok)
    return xml, changed


def fix_fsa_links(path: Path, make_backup: bool = True) -> dict[str, int]:
    parts, order = _read_parts(path)
    report: dict[str, int] = {}

    for sheet_name, patcher in {
        "RIS": _patch_ris,
        "Schedules": _patch_schedules,
        "FSA": _patch_fsa,
        "RCFS": _patch_rcfs,
    }.items():
        sheet_path = _sheet_zip_path(parts, sheet_name)
        xml = parts[sheet_path].decode("utf-8", "ignore")
        xml, changed = patcher(xml)
        parts[sheet_path] = xml.encode("utf-8")
        report[sheet_name] = changed

    _mark_full_calc(parts)

    if make_backup:
        backup = path.with_name(f"{path.stem}_pre_fsa_links_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")
    _write_parts(path, parts, order)
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ticker")
    parser.add_argument("--path")
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    if args.path:
        path = Path(args.path)
    elif args.ticker:
        path = _default_path(args.ticker)
    else:
        raise SystemExit("--ticker or --path required")
    report = fix_fsa_links(path, make_backup=not args.no_backup)
    print(f"FSA links fixed: {path}")
    for sheet, changed in report.items():
        print(f"  {sheet}: {changed} cells")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
