#!/usr/bin/env python3
"""Replace post-trim #REF formulas with their pre-trim cached values.

This is a fail-closed recovery for legacy workbooks whose Pipeline still points
at a build-time TAM sheet. Only cells currently containing #REF! are touched,
and every replacement value must exist in the supplied pre-trim workbook.
"""
from __future__ import annotations

import argparse
import html
import re
import zipfile
from pathlib import Path

import openpyxl


def sheet_path(zf: zipfile.ZipFile, name: str) -> str:
    wb = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    rid = re.search(rf'<sheet name="{re.escape(name)}"[^>]*r:id="(rId\d+)"', wb).group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def literal_cell(cell_xml: str, addr: str, value) -> str:
    attrs = re.match(r"<c\b([^>]*)", cell_xml).group(1)
    style = re.search(r'\bs="(\d+)"', attrs)
    style_attr = f' s="{style.group(1)}"' if style else ""
    if value is None:
        return f'<c r="{addr}"{style_attr}/>'
    if isinstance(value, bool):
        return f'<c r="{addr}"{style_attr} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{addr}"{style_attr}><v>{value}</v></c>'
    text = html.escape(str(value), quote=False)
    return f'<c r="{addr}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", required=True)
    ap.add_argument("--pre-trim", required=True)
    ap.add_argument("--sheet", default="Pipeline")
    args = ap.parse_args()
    path, backup = Path(args.path), Path(args.pre_trim)

    source = openpyxl.load_workbook(backup, read_only=True, data_only=True, keep_links=False)
    values = source[args.sheet]
    try:
        with zipfile.ZipFile(path) as zf:
            sp = sheet_path(zf, args.sheet)
            order = zf.namelist()
            parts = {n: zf.read(n) for n in order}
        xml = parts[sp].decode("utf-8", "ignore")
        changed, missing = 0, []

        def repl(match: re.Match[str]) -> str:
            nonlocal changed
            cell = match.group(0)
            fm = re.search(r"<f[^>]*>(.*?)</f>", cell, re.S)
            if not fm or "#REF!" not in html.unescape(fm.group(1)):
                return cell
            addr = re.search(r'\br="([A-Z]+\d+)"', cell).group(1)
            value = values[addr].value
            if value is None:
                missing.append(addr)
                return cell
            changed += 1
            return literal_cell(cell, addr, value)

        xml = re.sub(r'<c\b[^>]*?/>|<c\b[^>]*>.*?</c>', repl, xml, flags=re.S)
        if missing:
            raise SystemExit(f"pre-trim cached values missing for: {', '.join(missing[:30])}")
        if changed == 0:
            raise SystemExit("no #REF! formula cells found to freeze")
        parts[sp] = xml.encode("utf-8")
        tmp = path.with_suffix(".~freeze.xlsx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name in order:
                out.writestr(name, parts[name])
        tmp.replace(path)
        print(f"Frozen {changed} deleted-sheet references from pre-trim cached values → {path}")
    finally:
        source.close()


if __name__ == "__main__":
    main()
