#!/usr/bin/env python3
"""Restore Peer View(s) sections from the DD data center exports.

The Peer View / Peer Views tabs are shared peer-comparison data, not
ticker-specific template labels.  This script treats datastore/export as the
system of record and rewrites the matching section cells in the workbook with
explicit peer_drug and peer_metric facts.
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

REPO = Path(__file__).resolve().parents[1]
EXPORT = REPO / "datastore" / "export"

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _xml_escape(value: Any) -> str:
    text = "" if value is None else str(value)
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def _load_csvs() -> Tuple[Dict[str, List[Dict[str, str]]], Dict[Tuple[str, str], Dict[str, str]]]:
    drugs_by_section: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    metrics: Dict[Tuple[str, str], Dict[str, str]] = defaultdict(dict)

    with (EXPORT / "peer_drug.csv").open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            drugs_by_section[row["section"]].append(row)

    with (EXPORT / "peer_metric.csv").open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            metrics[(row["section"], row["col"])][row["metric"]] = row.get("value", "")

    return drugs_by_section, metrics


def _sheet_paths(xlsx_path: Path) -> Dict[str, str]:
    with zipfile.ZipFile(xlsx_path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_path = {}
    for rel in rels:
        if "worksheet" not in rel.get("Type", ""):
            continue
        target = rel.get("Target", "")
        rid_to_path[rel.get("Id", "")] = (
            f"xl/{target}" if not target.startswith("/") else target.lstrip("/")
        )

    out = {}
    for sheet in wb.findall(f".//{{{_NS_MAIN}}}sheet"):
        rid = sheet.get(f"{{{_NS_R}}}id", "")
        if rid in rid_to_path:
            out[sheet.get("name", "")] = rid_to_path[rid]
    return out


def _shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall(f"{{{_NS_MAIN}}}si"):
        texts = []
        for t in si.findall(f".//{{{_NS_MAIN}}}t"):
            texts.append(t.text or "")
        out.append("".join(texts))
    return out


def _cell_map(xml: str, shared_strings: List[str]) -> Dict[Tuple[int, str], Tuple[str, str]]:
    cells: Dict[Tuple[int, str], Tuple[str, str]] = {}
    for m in re.finditer(r'<c\s+([^>]*?)(?<!/)>(.*?)</c>', xml, re.S):
        attrs, body = m.group(1), m.group(2)
        ref = re.search(r'r="([A-Z]+)(\d+)"', attrs)
        if not ref:
            continue
        col, row = ref.group(1), int(ref.group(2))
        style_m = re.search(r's="(\d+)"', attrs)
        style = style_m.group(1) if style_m else ""
        text = ""
        if 't="s"' in attrs:
            v = re.search(r'<v>(\d+)</v>', body)
            if v and int(v.group(1)) < len(shared_strings):
                text = shared_strings[int(v.group(1))]
        elif 't="inlineStr"' in attrs:
            t = re.search(r'<is><t[^>]*>(.*?)</t></is>', body, re.S)
            if t:
                text = re.sub(r"<[^>]+>", "", t.group(1))
        elif 't="str"' in attrs:
            v = re.search(r'<v>(.*?)</v>', body, re.S)
            if v:
                text = v.group(1)
        else:
            v = re.search(r'<v>(.*?)</v>', body, re.S)
            if v:
                text = v.group(1)
        cells[(row, col)] = (text.strip() if isinstance(text, str) else text, style)
    return cells


def _find_cell(xml: str, addr: str) -> Optional[Tuple[int, int, str]]:
    for m in re.finditer(r'<c\s+([^>]*?)(?:/>|>.*?</c>)', xml, re.S):
        if re.search(rf'\br="{re.escape(addr)}"', m.group(1)):
            return m.start(), m.end(), m.group(0)
    return None


def _insert_cell(xml: str, row: int, col: str, cell_xml: str) -> str:
    row_m = re.search(rf'<row\b[^>]*\br="{row}"[^>]*>.*?</row>', xml, re.S)
    if not row_m:
        return xml
    row_xml = row_m.group(0)
    new_col_num = _col_to_num(col)
    insert_at = row_xml.rfind("</row>")
    for cm in re.finditer(r'<c\s+[^>]*\br="([A-Z]+)\d+"', row_xml):
        if _col_to_num(cm.group(1)) > new_col_num:
            insert_at = cm.start()
            break
    new_row = row_xml[:insert_at] + cell_xml + row_xml[insert_at:]
    return xml[:row_m.start()] + new_row + xml[row_m.end():]


def _patch_cell(xml: str, addr: str, value: Any, kind: str, fallback_style: str = "") -> str:
    found = _find_cell(xml, addr)
    style = fallback_style
    if found:
        style_m = re.search(r'\bs="(\d+)"', found[2])
        if style_m:
            style = style_m.group(1)
    style_attr = f' s="{style}"' if style else ""

    if kind == "blank":
        new_cell = f'<c r="{addr}"{style_attr}/>'
    elif kind == "number":
        new_cell = f'<c r="{addr}"{style_attr}><v>{value}</v></c>'
    else:
        new_cell = (
            f'<c r="{addr}"{style_attr} t="inlineStr">'
            f'<is><t>{_xml_escape(value)}</t></is></c>'
        )

    if found:
        return xml[:found[0]] + new_cell + xml[found[1]:]
    row = int(re.search(r'\d+', addr).group(0))
    col = re.search(r'[A-Z]+', addr).group(0)
    return _insert_cell(xml, row, col, new_cell)


def _display_ticker(ticker: str) -> str:
    ticker = (ticker or "").strip()
    if not ticker:
        return ""
    upper = ticker.upper()
    if "EQUITY" in upper or "PRIVATE" in upper or "UNLISTED" in upper:
        return ticker
    if re.fullmatch(r"[A-Z.]{1,5}", upper):
        return f"{upper} US Equity"
    return ticker


def _metric_kind(metric: str, value: str) -> Tuple[str, Any]:
    if value is None or value == "":
        return "blank", ""
    if metric.lower() == "date":
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(value))
        if m:
            dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return "number", (dt - datetime(1899, 12, 30)).days
    text = str(value).strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        return "number", text
    return "text", text


def _discover_sections(
    cell_text: Dict[Tuple[int, str], Tuple[str, str]],
    known_sections: set[str],
) -> List[Tuple[int, int, str]]:
    rows = []
    for (row, col), (text, _) in cell_text.items():
        if col == "A" and text == "X":
            section = cell_text.get((row, "D"), ("", ""))[0]
            if section in known_sections:
                rows.append((row, 0, section))
    rows.sort()
    out = []
    for idx, (row, _, section) in enumerate(rows):
        end = rows[idx + 1][0] - 1 if idx + 1 < len(rows) else row + 80
        out.append((row, end, section))
    return out


def _patch_sheet(xml: str, shared_strings: List[str], sheet_name: str,
                 current_ticker: str) -> Tuple[str, int]:
    drugs_by_section, metrics = _load_csvs()
    cells = _cell_map(xml, shared_strings)
    patches = 0

    if current_ticker:
        for (row, col), (text, style) in list(cells.items()):
            if isinstance(text, str) and current_ticker.upper() in text.upper():
                xml = _patch_cell(xml, f"{col}{row}", "", "blank", style)
                patches += 1

    for header_row, end_row, section in _discover_sections(cells, set(drugs_by_section)):
        section_metrics = {
            metric
            for row in drugs_by_section[section]
            for metric in metrics.get((section, row["col"]), {})
        }
        metric_rows = {
            cells[(row, "D")][0]: row
            for row in range(header_row + 1, end_row + 1)
            if (row, "D") in cells and cells[(row, "D")][0] in section_metrics
        }
        if not metric_rows:
            continue
        first_metric_row = min(metric_rows.values())
        drug_row = first_metric_row - 1
        ticker_row = drug_row - 2
        second_ticker_row = drug_row - 1

        for drug in drugs_by_section[section]:
            col = drug["col"]
            primary = _display_ticker(drug.get("ticker", ""))
            xml = _patch_cell(xml, f"{col}{ticker_row}", primary,
                              "text" if primary else "blank",
                              cells.get((ticker_row, col), ("", ""))[1])
            xml = _patch_cell(xml, f"{col}{drug_row}", drug.get("drug", ""),
                              "text" if drug.get("drug") else "blank",
                              cells.get((drug_row, col), ("", ""))[1])
            patches += 2

            existing_second = cells.get((second_ticker_row, col), ("", ""))[0]
            if current_ticker and current_ticker.upper() in existing_second.upper():
                xml = _patch_cell(xml, f"{col}{second_ticker_row}", "", "blank",
                                  cells.get((second_ticker_row, col), ("", ""))[1])
                patches += 1

            for metric, value in metrics.get((section, col), {}).items():
                row = metric_rows.get(metric)
                if row is None:
                    continue
                kind, typed = _metric_kind(metric, value)
                xml = _patch_cell(xml, f"{col}{row}", typed, kind,
                                  cells.get((row, col), ("", ""))[1])
                patches += 1

    return xml, patches


def restore(path: Path, ticker: str) -> int:
    sheet_paths = _sheet_paths(path)
    targets = {k: v for k, v in sheet_paths.items() if k in {"Peer View", "Peer Views"}}
    if not targets:
        raise RuntimeError("Peer View(s) sheets not found")

    modified: Dict[str, bytes] = {}
    total = 0
    with zipfile.ZipFile(path) as zf:
        shared = _shared_strings(zf)
        for sheet_name, zip_path in targets.items():
            xml = zf.read(zip_path).decode("utf-8")
            new_xml, count = _patch_sheet(xml, shared, sheet_name, ticker)
            if count:
                modified[zip_path] = new_xml.encode("utf-8")
                total += count

    if not modified:
        return 0

    tmp = path.with_suffix(".~datastore_peer.xlsx")
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified:
                    zout.writestr(item, modified[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
    tmp.replace(path)
    return total


def main() -> None:
    ap = argparse.ArgumentParser(description="Restore Peer Views from datastore exports")
    ap.add_argument("--ticker", required=True)
    ap.add_argument("--path", help="Workbook path; default DD/{ticker}/DCF {ticker}.xlsx")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    path = Path(args.path) if args.path else Path(
        f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/DCF {args.ticker}.xlsx"
    )
    if not path.exists():
        raise FileNotFoundError(path)
    if not (EXPORT / "peer_drug.csv").exists() or not (EXPORT / "peer_metric.csv").exists():
        sys.exit(f"Datastore exports missing under {EXPORT}")

    if not args.no_backup:
        backup = path.with_name(f"{path.stem}_pre_datastore_peer_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")

    n = restore(path, args.ticker.upper())
    print(f"Peer View(s) restored from datastore: {n} cell patches")


if __name__ == "__main__":
    main()
