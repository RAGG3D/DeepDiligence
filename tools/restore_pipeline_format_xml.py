#!/usr/bin/env python3
"""Restore MOLN Pipeline formatting from a reference workbook without touching values.

This patches OOXML directly instead of saving through Excel/openpyxl. It merges
the source style records into the target styles.xml, then points only
Pipeline!A1:AH35 at the mapped source styles. Existing cell values/formulas are
left intact.
"""
from __future__ import annotations

import argparse
import copy
import re
import shutil
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": NS_MAIN, "r": NS_REL, "rel": NS_PKG_REL}
Q = lambda tag: f"{{{NS_MAIN}}}{tag}"


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def num_to_col(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def coord(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def row_col_from_coord(cell_ref: str) -> tuple[int, int]:
    m = re.match(r"^([A-Z]+)([0-9]+)$", cell_ref)
    if not m:
        raise ValueError(f"bad cell reference: {cell_ref}")
    return int(m.group(2)), col_to_num(m.group(1))


def sheet_map(path: Path) -> dict[str, str]:
    with ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.get("Id"): rel.get("Target") for rel in rels}
        out: dict[str, str] = {}
        for sh in wb.find("m:sheets", NS):
            rid = sh.get(f"{{{NS_REL}}}id")
            target = rid_to_target[rid]
            if target.startswith("/"):
                part = target.lstrip("/")
            elif target.startswith("xl/"):
                part = target
            else:
                part = "xl/" + target
            out[sh.get("name")] = part
        return out


def children(parent: ET._Element, tag: str) -> list[ET._Element]:
    node = parent.find(f"m:{tag}", NS)
    return list(node) if node is not None else []


def ensure_collection(root: ET._Element, tag: str, before_tag: str | None = None) -> ET._Element:
    node = root.find(f"m:{tag}", NS)
    if node is not None:
        return node
    node = ET.Element(Q(tag))
    if before_tag:
        before = root.find(f"m:{before_tag}", NS)
        if before is not None:
            root.insert(root.index(before), node)
            return node
    root.append(node)
    return node


def update_count(node: ET._Element) -> None:
    node.set("count", str(len(list(node))))


def sanitize_number_format(code: str | None) -> str | None:
    """Keep the visual style but do not render zero values as a dash line."""
    if not code:
        return code
    parts = code.split(";")
    if len(parts) < 3:
        return code
    zero = parts[2].strip()
    if zero not in {r"\—", "—", '"—"', r"\-", "-", '"-"'}:
        return code

    positive = parts[0]
    if "%" in positive:
        decimals = re.search(r"0(\.0+)?%", positive)
        parts[2] = "0" + (decimals.group(1) or "") + "%" if decimals else "0%"
    else:
        decimals = re.search(r"0(\.0+)", positive)
        parts[2] = "0" + (decimals.group(1) if decimals else "")
    return ";".join(parts)


def merge_styles(src_styles: bytes, dst_styles: bytes) -> tuple[bytes, dict[int, int], dict[int, int]]:
    src = ET.fromstring(src_styles)
    dst = ET.fromstring(dst_styles)

    dst_numfmts = ensure_collection(dst, "numFmts", "fonts")
    dst_fonts = ensure_collection(dst, "fonts", "fills")
    dst_fills = ensure_collection(dst, "fills", "borders")
    dst_borders = ensure_collection(dst, "borders", "cellStyleXfs")
    dst_cellxfs = ensure_collection(dst, "cellXfs", "cellStyles")
    dst_dxfs = ensure_collection(dst, "dxfs", "tableStyles")

    src_numfmt_by_id = {
        int(n.get("numFmtId")): sanitize_number_format(n.get("formatCode"))
        for n in children(src, "numFmts")
    }
    dst_numfmt_by_code = {
        sanitize_number_format(n.get("formatCode")): int(n.get("numFmtId"))
        for n in children(dst, "numFmts")
    }
    used_numfmt_ids = [int(n.get("numFmtId")) for n in children(dst, "numFmts")]
    next_numfmt = max([163] + used_numfmt_ids) + 1
    numfmt_map: dict[int, int] = {}
    for src_id, code in src_numfmt_by_id.items():
        if code in dst_numfmt_by_code:
            numfmt_map[src_id] = dst_numfmt_by_code[code]
        else:
            numfmt_map[src_id] = next_numfmt
            node = ET.Element(Q("numFmt"))
            node.set("numFmtId", str(next_numfmt))
            node.set("formatCode", code)
            dst_numfmts.append(node)
            next_numfmt += 1

    font_offset = len(children(dst, "fonts"))
    fill_offset = len(children(dst, "fills"))
    border_offset = len(children(dst, "borders"))
    dxf_offset = len(children(dst, "dxfs"))

    for node in children(src, "fonts"):
        dst_fonts.append(copy.deepcopy(node))
    for node in children(src, "fills"):
        dst_fills.append(copy.deepcopy(node))
    for node in children(src, "borders"):
        dst_borders.append(copy.deepcopy(node))
    for node in children(src, "dxfs"):
        dst_dxfs.append(copy.deepcopy(node))

    style_map: dict[int, int] = {}
    base_xf = len(children(dst, "cellXfs"))
    for idx, xf in enumerate(children(src, "cellXfs")):
        new_xf = copy.deepcopy(xf)
        if new_xf.get("fontId") is not None:
            new_xf.set("fontId", str(int(new_xf.get("fontId")) + font_offset))
        if new_xf.get("fillId") is not None:
            new_xf.set("fillId", str(int(new_xf.get("fillId")) + fill_offset))
        if new_xf.get("borderId") is not None:
            new_xf.set("borderId", str(int(new_xf.get("borderId")) + border_offset))
        if new_xf.get("numFmtId") is not None:
            old_numfmt = int(new_xf.get("numFmtId"))
            if old_numfmt in numfmt_map:
                new_xf.set("numFmtId", str(numfmt_map[old_numfmt]))
        # Cell display does not require preserving the source cellStyleXfs index.
        if new_xf.get("xfId") is not None:
            new_xf.set("xfId", "0")
        dst_cellxfs.append(new_xf)
        style_map[idx] = base_xf + idx

    dxf_map = {idx: dxf_offset + idx for idx, _ in enumerate(children(src, "dxfs"))}

    for tag in ("numFmts", "fonts", "fills", "borders", "cellXfs", "dxfs"):
        node = dst.find(f"m:{tag}", NS)
        if node is not None:
            update_count(node)

    return ET.tostring(dst, xml_declaration=True, encoding="UTF-8", standalone=False), style_map, dxf_map


def remap_style_attr(node: ET._Element, style_map: dict[int, int]) -> None:
    if node.get("s") is not None:
        node.set("s", str(style_map[int(node.get("s"))]))
    if node.get("style") is not None:
        node.set("style", str(style_map[int(node.get("style"))]))


def replace_or_insert(dst_root: ET._Element, src_root: ET._Element, tag: str, before_tag: str | None = None) -> None:
    src_node = src_root.find(f"m:{tag}", NS)
    dst_node = dst_root.find(f"m:{tag}", NS)
    if src_node is None:
        if dst_node is not None:
            dst_root.remove(dst_node)
        return
    new_node = copy.deepcopy(src_node)
    if dst_node is not None:
        dst_root.replace(dst_node, new_node)
    elif before_tag:
        before = dst_root.find(f"m:{before_tag}", NS)
        if before is not None:
            dst_root.insert(dst_root.index(before), new_node)
        else:
            dst_root.append(new_node)
    else:
        dst_root.append(new_node)


def row_map(root: ET._Element) -> dict[int, ET._Element]:
    out = {}
    for row in root.findall(".//m:sheetData/m:row", NS):
        out[int(row.get("r"))] = row
    return out


def cell_map(row: ET._Element) -> dict[str, ET._Element]:
    return {c.get("r"): c for c in row.findall("m:c", NS)}


def get_or_create_row(root: ET._Element, r: int) -> ET._Element:
    sheet_data = root.find("m:sheetData", NS)
    rows = row_map(root)
    if r in rows:
        return rows[r]
    row = ET.Element(Q("row"))
    row.set("r", str(r))
    inserted = False
    for existing in sheet_data.findall("m:row", NS):
        if int(existing.get("r")) > r:
            sheet_data.insert(sheet_data.index(existing), row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    return row


def get_or_create_cell(row: ET._Element, ref: str) -> ET._Element:
    cells = cell_map(row)
    if ref in cells:
        return cells[ref]
    cell = ET.Element(Q("c"))
    cell.set("r", ref)
    _, col = row_col_from_coord(ref)
    inserted = False
    for existing in row.findall("m:c", NS):
        _, existing_col = row_col_from_coord(existing.get("r"))
        if existing_col > col:
            row.insert(row.index(existing), cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)
    return cell


def copy_row_attrs(src_row: ET._Element, dst_row: ET._Element, style_map: dict[int, int]) -> None:
    r = dst_row.get("r")
    for key in list(dst_row.attrib):
        if key != "r":
            del dst_row.attrib[key]
    for key, value in src_row.attrib.items():
        if key != "r":
            dst_row.set(key, value)
    dst_row.set("r", r)
    remap_style_attr(dst_row, style_map)


def patch_pipeline_sheet(src_xml: bytes, dst_xml: bytes, style_map: dict[int, int], dxf_map: dict[int, int]) -> bytes:
    src = ET.fromstring(src_xml)
    dst = ET.fromstring(dst_xml)

    # Copy sheet-level formatting/layout, but keep target dimension and values.
    for tag, before in (
        ("sheetPr", "dimension"),
        ("sheetViews", "sheetFormatPr"),
        ("sheetFormatPr", "cols"),
        ("printOptions", "pageMargins"),
        ("pageMargins", "pageSetup"),
        ("pageSetup", "headerFooter"),
        ("headerFooter", "rowBreaks"),
    ):
        replace_or_insert(dst, src, tag, before)

    # Copy source column definitions. They are formatting/layout only.
    src_cols = src.find("m:cols", NS)
    dst_cols = dst.find("m:cols", NS)
    if src_cols is not None:
        new_cols = copy.deepcopy(src_cols)
        for col in new_cols.findall("m:col", NS):
            remap_style_attr(col, style_map)
        if dst_cols is not None:
            dst.replace(dst_cols, new_cols)
        else:
            sheet_data = dst.find("m:sheetData", NS)
            dst.insert(dst.index(sheet_data), new_cols)

    src_rows = row_map(src)
    for r in range(1, 36):
        if r not in src_rows:
            continue
        dst_row = get_or_create_row(dst, r)
        copy_row_attrs(src_rows[r], dst_row, style_map)
        src_cells = cell_map(src_rows[r])
        for c in range(1, 35):
            ref = coord(r, c)
            src_cell = src_cells.get(ref)
            dst_cell = get_or_create_cell(dst_row, ref)
            if src_cell is not None and src_cell.get("s") is not None:
                dst_cell.set("s", str(style_map[int(src_cell.get("s"))]))
            else:
                dst_cell.attrib.pop("s", None)

    # Rebuild conditional formatting for visible live ranges only.
    for cf in list(dst.findall("m:conditionalFormatting", NS)):
        dst.remove(cf)

    source_expression_rule = None
    source_colorscale_rule = None
    for cf in src.findall("m:conditionalFormatting", NS):
        sqref = cf.get("sqref", "")
        for rule in cf.findall("m:cfRule", NS):
            if "B1:B6" in sqref and rule.get("type") == "expression" and source_expression_rule is None:
                source_expression_rule = copy.deepcopy(rule)
            if rule.get("type") == "colorScale" and source_colorscale_rule is None:
                source_colorscale_rule = copy.deepcopy(rule)

    new_cfs: list[ET._Element] = []
    priority = 1
    if source_expression_rule is not None:
        if source_expression_rule.get("dxfId") is not None:
            source_expression_rule.set("dxfId", str(dxf_map[int(source_expression_rule.get("dxfId"))]))
        source_expression_rule.set("priority", str(priority))
        priority += 1
        cf = ET.Element(Q("conditionalFormatting"))
        cf.set("sqref", "B1:B6")
        cf.append(source_expression_rule)
        new_cfs.append(cf)

    if source_colorscale_rule is not None:
        for row in (11, 18, 20, 27, 29, 31):
            rule = copy.deepcopy(source_colorscale_rule)
            rule.set("priority", str(priority))
            priority += 1
            cf = ET.Element(Q("conditionalFormatting"))
            cf.set("sqref", f"F{row}:AH{row}")
            cf.append(rule)
            new_cfs.append(cf)

    sheet_data = dst.find("m:sheetData", NS)
    insert_at = dst.index(sheet_data) + 1
    for offset, cf in enumerate(new_cfs):
        dst.insert(insert_at + offset, cf)

    return ET.tostring(dst, xml_declaration=True, encoding="UTF-8", standalone=False)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True, type=Path)
    ap.add_argument("--target", required=True, type=Path)
    args = ap.parse_args()

    src_sheet = sheet_map(args.source)["Pipeline"]
    dst_sheet = sheet_map(args.target)["Pipeline"]

    with ZipFile(args.source) as zsrc, ZipFile(args.target) as zdst:
        patched_styles, style_map, dxf_map = merge_styles(zsrc.read("xl/styles.xml"), zdst.read("xl/styles.xml"))
        patched_pipeline = patch_pipeline_sheet(
            zsrc.read(src_sheet),
            zdst.read(dst_sheet),
            style_map,
            dxf_map,
        )

        tmp = Path(tempfile.mkstemp(suffix=".xlsx", dir=str(args.target.parent))[1])
        try:
            with ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
                for info in zdst.infolist():
                    if info.filename == "xl/styles.xml":
                        data = patched_styles
                    elif info.filename == dst_sheet:
                        data = patched_pipeline
                    else:
                        data = zdst.read(info.filename)
                    zout.writestr(info, data)
            shutil.copystat(args.target, tmp)
            tmp.replace(args.target)
        finally:
            if tmp.exists():
                tmp.unlink()

    print(f"Patched Pipeline formatting in {args.target}")


if __name__ == "__main__":
    main()
