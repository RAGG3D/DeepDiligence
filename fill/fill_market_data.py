#!/usr/bin/env python3
"""
fill_market_data.py — populate the `BBG DAPI` tab with live market data.

The DCF's VALUATION tab pulls WACC + football-field inputs from `BBG DAPI` via
SUMIFS on the column-F label (col I = FINAL value). A template-seeded workbook
still holds the template company's (BCYC) values, so VALUATION computes off the
wrong market data. This fills the automatable fields from yfinance:

    Last Price, Adjusted Beta, Shares Outstanding, Market Cap,
    Consensus PT - Median/High/Low, Lifetime High/Low, Risk Free Rate (^TNX)

Both the OVERRIDE (col H) and FINAL (col I) cells are set to the live value.
Fields with no free source (Implied CDS Spread, Equity Risk Premium) are set to
ticker-neutral defaults (never the template company's figure). Currency note: yfinance prices/shares are in the listing
currency (USD for MOLN on Nasdaq) while the financial model is in the reporting
currency (CHF) — flagged, not converted, in this pass.

    python fill/fill_market_data.py --ticker MOLN
"""
import argparse
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.neutralize_bloomberg_artifacts import _market_data


def sheet_zip_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wbx = zf.read("xl/workbook.xml").decode("utf8", "ignore")
    m = re.search(rf'<sheet name="{re.escape(sheet_name)}"[^>]*r:id="(rId\d+)"', wbx)
    if not m:
        raise ValueError(f"sheet {sheet_name!r} not in workbook.xml")
    rid = m.group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8", "ignore")
    mt = re.search(rf'<Relationship[^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels)
    tgt = mt.group(1).lstrip("/")
    return tgt if tgt.startswith("xl/") else "xl/" + tgt


def label_row_map(zf, path):
    """Return {col-F label text -> row number} from the BBG DAPI sheet."""
    xml = zf.read(path).decode("utf8", "ignore")
    shared = []
    if "xl/sharedStrings.xml" in zf.namelist():
        ss = zf.read("xl/sharedStrings.xml").decode("utf8", "ignore")
        shared = re.findall(r"<si>(.*?)</si>", ss, re.S)
        shared = [re.sub(r"<[^>]+>", "", s) for s in shared]
    out = {}
    for cm in re.finditer(r'<c r="F(\d+)"([^>]*)>(.*?)</c>', xml, re.S):
        row, attrs, body = cm.group(1), cm.group(2), cm.group(3)
        vv = re.search(r"<v>(.*?)</v>", body)
        if not vv:
            continue
        if 't="s"' in attrs:
            txt = shared[int(vv.group(1))] if vv.group(1).isdigit() else ""
        elif 't="inlineStr"' in attrs:
            it = re.search(r"<t[^>]*>(.*?)</t>", body)
            txt = it.group(1) if it else ""
        else:
            txt = vv.group(1)
        if txt:
            out[txt.strip()] = int(row)
    return out


def patch_numeric(xml: str, addr: str, value) -> str:
    """Set cell `addr` to a numeric value, dropping any formula/type, keeping style."""
    def head_of(h):
        return re.sub(r'\s+t="[^"]*"', "", h)   # numeric cell carries no t= attr
    # self-closing <c r=".." .../>
    sc = re.compile(rf'(<c r="{addr}"[^>]*?)\s*/>')
    xml2, n = sc.subn(lambda m: f"{head_of(m.group(1))}><v>{value}</v></c>", xml)
    if n:
        return xml2
    # existing <c r=".." ..> ... </c>
    ex = re.compile(rf'(<c r="{addr}"[^>]*?)>.*?</c>', re.S)
    xml3, n3 = ex.subn(lambda m: f"{head_of(m.group(1))}><v>{value}</v></c>", xml)
    return xml3 if n3 else xml


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ticker", required=True)
    ap.add_argument("--path")
    a = ap.parse_args()
    dcf = Path(a.path) if a.path else Path(
        f"/mnt/c/Users/yzsun/Desktop/DD/{a.ticker}/DCF {a.ticker}.xlsx")

    md = _market_data(a.ticker.upper())

    # label (col F) -> live value
    updates = {
        "Last Price": md["price"], "Adjusted Beta": md["beta"],
        "Shares Outstanding": md["shares_m"] or None,
        "Market Cap": md["mktcap_m"] or None,
        "Consensus PT - Median": md["target_median"], "Consensus PT - High": md["target_high"],
        "Consensus PT - Low": md["target_low"], "Lifetime High": md["lifetime_high"],
        "Lifetime Low": md["lifetime_low"], "Risk Free Rate": md["risk_free"],
        # Ticker-neutral WACC inputs. The template ships the reference company's
        # numbers here (BCYC ERP 6.0863 / CDS 115.13) and no free yfinance field
        # supplies them, so write generic defaults — never a reference company's
        # value — otherwise the stale BCYC constants flow into VALUATION WACC.
        "Equity Risk Premium": 5.5,   # generic equity-risk-premium (%)
        "Implied CDS Spread": 0,      # most pre-commercial biotechs carry no debt
    }
    # Warn loudly for any label yfinance returned None for (dropped → the template
    # default, which is the reference company's own figure, is left in place)
    # instead of silently keeping a stale value.
    for label in [k for k, v in updates.items() if v is None]:
        print(f"  !! WARNING: no live value for {label!r} (yfinance returned None); "
              f"template default kept — this may be the reference company's figure")
    updates = {k: v for k, v in updates.items() if v is not None}

    bak = dcf.with_name(f"{dcf.stem}_pre_mktdata_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(dcf, bak)

    with zipfile.ZipFile(dcf) as zf:
        spath = sheet_zip_path(zf, "BBG DAPI")
        rows = label_row_map(zf, spath)
        xml = zf.read(spath).decode("utf8", "ignore")
        names = zf.namelist()
        blobs = {n: zf.read(n) for n in names}

    applied = []
    for label, val in updates.items():
        r = rows.get(label)
        if not r:
            print(f"  ! label not found: {label}")
            continue
        for col in ("H", "I"):                 # OVERRIDE + FINAL
            xml = patch_numeric(xml, f"{col}{r}", val)
        applied.append((label, r, val))

    blobs[spath] = xml.encode("utf8")
    with zipfile.ZipFile(dcf, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in blobs:
            zo.writestr(n, blobs[n])

    print(f"BBG DAPI market data filled ({a.ticker}, listing ccy {md.get('currency')}, "
          f"completed session {md.get('price_as_of')}):")
    for label, r, val in applied:
        print(f"  {label:24} row {r:>3}  = {val}")
    print(f"backup: {bak.name}")
    print("NOTE: prices/shares are in listing currency (e.g. USD); model financials "
          "are in reporting currency (e.g. CHF) — not FX-converted this pass.")


if __name__ == "__main__":
    main()
