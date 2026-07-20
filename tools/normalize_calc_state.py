#!/usr/bin/env python3
"""Force an Excel-native, non-stale calculation state.

Root cause of the "opened the file → prompt about old data → some numbers show a
line through them" report: the LibreOffice recalc roundtrip (the harden/recalc
step) saves the workbook with `calcMode="manual"` + `calcCompleted="0"`
(+ `calcOnSave="0"`) and a leftover MS `calcFeatures` extLst. Excel reads that as
"calculation is out of date", prompts on open, and renders the stale cached
values with a strikethrough — none of which is stored in the file (which is why
no <strike> exists anywhere in it).

This makes the shipped workbook clean so Excel recalculates fresh on open and
never shows the stale prompt/strikethrough:
  * xl/workbook.xml <calcPr>  → auto calc + fullCalcOnLoad="1"; iterate settings
    kept; calcMode/calcCompleted/calcOnSave/forceFullCalc dropped
  * workbook-level <extLst> (calcFeatures) removed
  * xl/calcChain.xml dropped (+ its content-type/rels) so Excel rebuilds it

Only these package-level bits change; every cell value/formula/style is untouched.

    python tools/normalize_calc_state.py --ticker MOLN
    python tools/normalize_calc_state.py --path "/path/DCF X.xlsx"
"""
import argparse
import re
import shutil
import time
import zipfile
from pathlib import Path

CALC_KEEP = ("calcId", "iterate", "iterateCount", "iterateDelta", "refMode")


def clean_calcpr(wbx: str) -> str:
    m = re.search(r"<calcPr\b([^>]*)/>", wbx)
    attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))) if m else {}
    kept = {k: attrs[k] for k in CALC_KEEP if k in attrs}
    kept.setdefault("calcId", "191029")
    kept.setdefault("iterate", "1")
    kept.setdefault("iterateDelta", "0.0001")
    kept["fullCalcOnLoad"] = "1"                 # Excel recalcs cleanly on open
    new = "<calcPr " + " ".join(f'{k}="{v}"' for k, v in kept.items()) + "/>"
    if m:
        return wbx[:m.start()] + new + wbx[m.end():]
    # no calcPr → insert before </workbook>-ish (after sheets); safe fallback
    return wbx.replace("</workbook>", new + "</workbook>")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ticker")
    ap.add_argument("--path")
    ap.add_argument("--no-backup", action="store_true",
                    help="Normalize in place without creating an additional backup")
    a = ap.parse_args()
    dcf = Path(a.path) if a.path else Path(
        f"/mnt/c/Users/yzsun/Desktop/DD/{a.ticker}/DCF {a.ticker}.xlsx")

    bak = None
    if not a.no_backup:
        bak = dcf.with_name(
            f"{dcf.stem}_pre_calcstate_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        shutil.copy2(dcf, bak)

    with zipfile.ZipFile(dcf) as zf:
        names = zf.namelist()
        blobs = {n: zf.read(n) for n in names}

    wbx = blobs["xl/workbook.xml"].decode("utf8", "ignore")
    before = wbx
    wbx = clean_calcpr(wbx)
    wbx = re.sub(r"<extLst>.*?</extLst>", "", wbx, flags=re.S)   # drop calcFeatures
    blobs["xl/workbook.xml"] = wbx.encode("utf8")

    dropped_chain = "xl/calcChain.xml" in blobs
    blobs.pop("xl/calcChain.xml", None)
    if dropped_chain:
        ct = blobs["[Content_Types].xml"].decode("utf8", "ignore")
        ct = re.sub(r'<Override[^>]*/xl/calcChain\.xml[^>]*/>', "", ct)
        blobs["[Content_Types].xml"] = ct.encode("utf8")
        rels = blobs["xl/_rels/workbook.xml.rels"].decode("utf8", "ignore")
        rels = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", rels)
        blobs["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")

    with zipfile.ZipFile(dcf, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            if n in blobs:
                zo.writestr(n, blobs[n])

    calc = re.search(r"<calcPr[^>]*/>", wbx)
    print(f"calc state normalized: {calc.group(0) if calc else '?'}")
    print(f"  extLst removed: {before.count('<extLst>')} | calcChain dropped: {dropped_chain}")
    print(f"  backup: {bak.name if bak else 'disabled'}")


if __name__ == "__main__":
    main()
