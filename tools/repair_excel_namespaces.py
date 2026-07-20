#!/usr/bin/env python3
"""Repair OOXML namespace prefixes that Excel validates strictly.

ElementTree can rewrite known Office prefixes such as mc/x14ac/xr into ns1/ns2
while leaving mc:Ignorable values unchanged.  Python and LibreOffice may still
parse the XML, but desktop Excel can refuse to open the workbook because the
Ignorable attribute references prefixes that are no longer declared.

This tool keeps cell content/formulas/styles intact and only normalizes package
metadata XML.
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import time
import zipfile
from pathlib import Path


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _default_path(ticker: str) -> Path:
    return Path(f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx")


def _repair_styles(xml: str) -> str:
    replacements = {
        'xmlns:ns1="http://schemas.openxmlformats.org/markup-compatibility/2006"': f'xmlns:mc="{MC_NS}"',
        'xmlns:ns2="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"': 'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"',
        'xmlns:ns3="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"': 'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
        'xmlns:ns4="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"': 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"',
        'xmlns:ns5="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"': 'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"',
        'ns1:Ignorable=': 'mc:Ignorable=',
        'ns2:': 'x14ac:',
        'ns3:': 'xr:',
        'ns4:': 'x14:',
        'ns5:': 'x15:',
    }
    for old, new in replacements.items():
        xml = xml.replace(old, new)
    if 'mc:Ignorable="' in xml and 'xmlns:x16r2=' not in xml:
        xml = xml.replace(
            '<styleSheet ',
            '<styleSheet xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" ',
            1,
        )
    return xml


def _sheet_entries(workbook_xml: str) -> list[tuple[str, str]]:
    entries: list[tuple[str, str]] = []
    for match in re.finditer(r'<(?:\w+:)?sheet\b([^>]*)/>', workbook_xml):
        attrs = match.group(1)
        name_m = re.search(r'\bname="([^"]+)"', attrs)
        rid_m = re.search(r'\br:id="([^"]+)"', attrs)
        if name_m and rid_m:
            entries.append((html.unescape(name_m.group(1)), rid_m.group(1)))
    return entries


def _canonicalize_prefix(xml: str, uri: str, canonical: str) -> str:
    """Rename only aliases bound to ``uri``; preserve every workbook node.

    The prior implementation rebuilt workbook.xml from a tiny skeleton, which
    discarded Excel-native workbook properties, extension lists, defined names
    and revision metadata.  Desktop Excel rejected that lossy result even
    though openpyxl and LibreOffice tolerated it.  Prefix repair must be a
    lexical namespace operation, never a workbook reconstruction.
    """
    aliases = re.findall(rf'xmlns:([A-Za-z_][\w.-]*)="{re.escape(uri)}"', xml)
    if canonical in aliases:
        aliases = [alias for alias in aliases if alias != canonical]
    for alias in aliases:
        if f'xmlns:{canonical}=' in xml:
            # The canonical binding already exists; remove the duplicate alias.
            xml = re.sub(
                rf'\s+xmlns:{re.escape(alias)}="{re.escape(uri)}"', "", xml, count=1
            )
        else:
            xml = xml.replace(f'xmlns:{alias}="{uri}"', f'xmlns:{canonical}="{uri}"', 1)
        xml = xml.replace(f'<{alias}:', f'<{canonical}:')
        xml = xml.replace(f'</{alias}:', f'</{canonical}:')
        xml = re.sub(rf'(?<=\s){re.escape(alias)}:', f'{canonical}:', xml)
        # mc:Ignorable stores prefix names as plain tokens, not QName attrs.
        xml = re.sub(
            r'Ignorable="([^"]*)"',
            lambda m: 'Ignorable="' + " ".join(
                canonical if token == alias else token
                for token in m.group(1).split()
            ) + '"',
            xml,
        )
    return xml


def _clean_workbook_xml(workbook_xml: str) -> str:
    mappings = {
        R_NS: "r",
        MC_NS: "mc",
        "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main": "x15",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision": "xr",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6": "xr6",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10": "xr10",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2": "xr2",
    }
    for uri, canonical in mappings.items():
        workbook_xml = _canonicalize_prefix(workbook_xml, uri, canonical)
    return workbook_xml


def _repair_rels(xml: str) -> str:
    xml = xml.replace("ns0:", "").replace(":ns0", "")
    xml = re.sub(r'(<Relationships\b[^>]*?)\s+xmlns="[^"]*"', r'\1', xml, count=1)
    if not re.search(r'<Relationships\b[^>]*\sxmlns=', xml[:500]):
        xml = xml.replace("<Relationships", f'<Relationships xmlns="{REL_NS}"', 1)
    return xml


def _missing_ignorable_prefixes(xml: str) -> list[str]:
    declared = set(re.findall(r'xmlns:([A-Za-z0-9_]+)=', xml[:10000]))
    missing: list[str] = []
    for value in re.findall(r'Ignorable="([^"]+)"', xml[:10000]):
        missing.extend(prefix for prefix in value.split() if prefix not in declared)
    return sorted(set(missing))


def repair(path: Path, make_backup: bool = True) -> dict[str, int]:
    if make_backup:
        backup = path.with_name(f"{path.stem}_pre_namespace_repair_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")

    with zipfile.ZipFile(path, "r") as zin:
        bad = zin.testzip()
        if bad:
            raise RuntimeError(f"zip CRC failure in {bad}")
        parts = {name: zin.read(name) for name in zin.namelist()}
        order = zin.namelist()

    report = {"workbook": 0, "styles": 0, "rels": 0}
    if "xl/workbook.xml" in parts:
        old = parts["xl/workbook.xml"].decode("utf-8", "ignore")
        new = _clean_workbook_xml(old)
        report["workbook"] = int(new != old)
        parts["xl/workbook.xml"] = new.encode("utf-8")
    if "xl/styles.xml" in parts:
        old = parts["xl/styles.xml"].decode("utf-8", "ignore")
        new = _repair_styles(old)
        report["styles"] = int(new != old)
        parts["xl/styles.xml"] = new.encode("utf-8")
    if "xl/_rels/workbook.xml.rels" in parts:
        old = parts["xl/_rels/workbook.xml.rels"].decode("utf-8", "ignore")
        new = _repair_rels(old)
        report["rels"] = int(new != old)
        parts["xl/_rels/workbook.xml.rels"] = new.encode("utf-8")

    tmp = path.with_suffix(".~namespace.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order:
            if name in parts:
                zout.writestr(name, parts[name])
    tmp.replace(path)

    with zipfile.ZipFile(path, "r") as zf:
        for name in ("xl/workbook.xml", "xl/styles.xml"):
            missing = _missing_ignorable_prefixes(zf.read(name).decode("utf-8", "ignore"))
            if missing:
                raise RuntimeError(f"{name} has undeclared Ignorable prefixes: {missing}")
    print(f"Excel namespace repair: {report}")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Repair Excel-sensitive OOXML namespace prefixes")
    parser.add_argument("--ticker")
    parser.add_argument("--path")
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    path = Path(args.path) if args.path else _default_path(args.ticker)
    if not path.exists():
        raise FileNotFoundError(path)
    repair(path, make_backup=not args.no_backup)


if __name__ == "__main__":
    main()
