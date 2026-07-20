#!/usr/bin/env python3
"""Remove misleading strike/dash display formats from generated DCF workbooks.

This patches OOXML directly and changes only workbook styles:
- custom number formats whose zero section is an em dash/hyphen become zero-as-zero
- font strike flags are removed from fonts and conditional-format dxfs

Cell values, formulas, cached values, charts, and sheet XML are left untouched.
"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": NS_MAIN}
Q = lambda tag: f"{{{NS_MAIN}}}{tag}"


def zero_format_from_positive(positive: str) -> str:
    if "%" in positive:
        decimals = re.search(r"0(\.0+)?%", positive)
        return "0" + ((decimals.group(1) if decimals else "") or "") + "%"
    decimals = re.search(r"0(\.0+)", positive)
    return "0" + (decimals.group(1) if decimals else "")


def sanitize_number_format(code: str | None) -> tuple[str | None, bool]:
    if not code:
        return code, False
    parts = code.split(";")
    if len(parts) < 3:
        return code, False
    zero = parts[2].strip()
    # Excel often stores these as \— or quoted "—"; treat ASCII hyphen too.
    if zero not in {r"\—", "—", '"—"', r"\-", "-", '"-"'}:
        return code, False
    parts[2] = zero_format_from_positive(parts[0])
    return ";".join(parts), True


def remove_strike_flags(root: ET._Element) -> int:
    removed = 0
    for strike in root.findall(".//m:strike", NS):
        parent = strike.getparent()
        if parent is not None:
            parent.remove(strike)
            removed += 1
    return removed


def patch_styles(styles_xml: bytes) -> tuple[bytes, int, int]:
    root = ET.fromstring(styles_xml)
    patched_numfmts = 0
    numfmts = root.find("m:numFmts", NS)
    if numfmts is not None:
        for numfmt in numfmts.findall("m:numFmt", NS):
            code = numfmt.get("formatCode")
            new_code, changed = sanitize_number_format(code)
            if changed:
                numfmt.set("formatCode", new_code)
                patched_numfmts += 1
    removed_strikes = remove_strike_flags(root)
    return (
        ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=False),
        patched_numfmts,
        removed_strikes,
    )


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", required=True, type=Path)
    args = ap.parse_args()

    with ZipFile(args.path, "r") as zin:
        patched_styles, numfmts, strikes = patch_styles(zin.read("xl/styles.xml"))
        tmp = Path(tempfile.mkstemp(suffix=".xlsx", dir=str(args.path.parent))[1])
        try:
            with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    data = patched_styles if info.filename == "xl/styles.xml" else zin.read(info.filename)
                    zout.writestr(info, data)
            shutil.copystat(args.path, tmp)
            tmp.replace(args.path)
        finally:
            if tmp.exists():
                tmp.unlink()
    print(f"sanitized_numFmts={numfmts} removed_strike_tags={strikes}")


if __name__ == "__main__":
    main()
