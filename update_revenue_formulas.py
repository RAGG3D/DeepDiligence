#!/usr/bin/env python3
"""update_revenue_formulas.py

Rewrite Revenue & COGS formulas in the Pipeline sheet so they reference
TAM Solid's growth-rate rows (551-553) and COGS/Price row (562) instead
of deleted maturity-curve rows (R444-446) and COGS rate row (R455).

Also fills missing MS row C-column ratings from Peer Views.
"""

import zipfile
import re
import io
import shutil
import sys
from pathlib import Path
from datetime import datetime


_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)
_XLSX_PATH = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
_PIPE_SHEET = "xl/worksheets/sheet9.xml"

# ── column helpers ──────────────────────────────────────────────────────────
def _col_letter(n: int) -> str:
    """1-indexed column number → letter(s).  6→'F', 21→'U'."""
    if n <= 26:
        return chr(64 + n)
    return chr(64 + (n - 1) // 26) + chr(64 + (n - 1) % 26 + 1)

_COL_F = 6   # first data col (2023)
_COL_U = 21  # last data col  (2038)

_REV_STYLE = "336"   # s= attribute on Revenue/COGS formula cells

# ── Drug block specs ────────────────────────────────────────────────────────
# Each "terms" entry is (tam_row, ms_row) for one indication.
DRUGS = [
    {"rev": 15, "cogs": 16, "drug": 9,
     "terms": [(10, 11), (12, 13)]},                               # CTX-009: BTC, CRC
    {"rev": 28, "cogs": 29, "drug": 18,
     "terms": [(19, 20), (21, 22), (23, 24), (25, 26)]},           # CTX-10726: RCC,HCC,GC,EC
    {"rev": 39, "cogs": 40, "drug": 31,
     "terms": [(32, 33), (34, 35), (36, 37)]},                     # CTX-471: ES-SCLC,Mel,MPM
    {"rev": 54, "cogs": 55, "drug": 42,
     "terms": [(43, 44), (45, 46), (47, 48), (49, 50), (51, 52)]}, # CTX-8371: NSCLC,TNBC,HL,MM,HNSCC
]

# MS rows → indication abbreviations (for C-column rating fill)
_MS_INDICATION = {
    11: "BTC", 13: "CRC",
    20: "RCC", 22: "HCC", 24: "GC", 26: "EC",
    33: "ES-SCLC", 35: "Melanoma NCAM+", 37: "MPM",
    44: "NSCLC", 46: "TNBC", 48: "HL", 50: "MM", 52: "HNSCC",
}

# Peer Views rating codes → C-column label text
_RATING_TEXT = {
    "BIC": "Best-In-Class Growth",
    "T1":  "Tier One Growth",
    "AVG": "Average Growth",
}


# ── formula builders ────────────────────────────────────────────────────────
def _growth_factor(ms_row: int, ysa_expr: str) -> str:
    """Nested IF selecting growth-rate row based on C-column rating text."""
    bic = f"INDEX('TAM Solid'!$F$552:$AH$552,MIN({ysa_expr},29))"
    t1  = f"INDEX('TAM Solid'!$F$553:$AH$553,MIN({ysa_expr},29))"
    avg = f"INDEX('TAM Solid'!$F$551:$AH$551,MIN({ysa_expr},29))"
    return (
        f'IF($C${ms_row}="Best-In-Class Growth",{bic},'
        f'IF($C${ms_row}="Tier One Growth",{t1},{avg}))'
    )


def _build_revenue_formula(col: str, drug_row: int, terms: list) -> str:
    """Revenue = SUM(TAM × MS × GrowthFactor) for each indication, guarded by phase-5 check."""
    ysa = f"COLUMN({col}1)-MATCH(5,$F${drug_row}:$U${drug_row},0)-COLUMN($F1)+2"
    parts = []
    for tam_row, ms_row in terms:
        gf = _growth_factor(ms_row, ysa)
        parts.append(f"{col}{tam_row}*{col}{ms_row}*{gf}")
    inner = "+".join(parts)
    guard = f"COUNTIF($F{drug_row}:{col}{drug_row},5)&gt;0"
    return f"IF({guard},{inner},0)"


def _build_cogs_formula(col: str, drug_row: int, rev_row: int) -> str:
    """COGS = COGS_rate(from TAM Solid R562) × Revenue, guarded by phase-5 check."""
    ysa = f"COLUMN({col}1)-MATCH(5,$F${drug_row}:$U${drug_row},0)-COLUMN($F1)+2"
    cogs_lookup = f"INDEX('TAM Solid'!$K$562:$AH$562,MIN({ysa},24))"
    guard = f"COUNTIF($F{drug_row}:{col}{drug_row},5)&gt;0"
    return f"IF({guard},{cogs_lookup}*{col}{rev_row},0)"


def _build_row_cells(row: int, drug_row: int, terms_or_rev, is_cogs: bool = False) -> str:
    """Build XML for columns F-U of a Revenue or COGS row."""
    parts = []
    for n in range(_COL_F, _COL_U + 1):
        col = _col_letter(n)
        if is_cogs:
            formula = _build_cogs_formula(col, drug_row, terms_or_rev)
        else:
            formula = _build_revenue_formula(col, drug_row, terms_or_rev)
        parts.append(
            f'<c r="{col}{row}" s="{_REV_STYLE}">'
            f"<f>{formula}</f>"
            f"<v>0</v>"
            f"</c>"
        )
    return "".join(parts)


# ── row replacement ─────────────────────────────────────────────────────────
def _replace_row_cells(xml: str, row: int, new_cells: str) -> str:
    """Replace F-U cells in a row; keep A-E cells intact."""
    pat = re.compile(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', re.DOTALL)
    m = pat.search(xml)
    if not m:
        print(f"  WARNING: row {row} not found")
        return xml

    row_tag  = m.group(1)
    row_body = m.group(2)

    # Keep everything before the first F-column cell
    f_start = re.search(rf'<c r="F{row}"', row_body)
    if f_start:
        pre_f = row_body[: f_start.start()]
    else:
        # No F cell exists yet — keep all existing cells and append
        pre_f = row_body

    new_row = row_tag + pre_f + new_cells + "</row>"
    return xml[: m.start()] + new_row + xml[m.end():]


# ── C-column rating fill ────────────────────────────────────────────────────
def _fill_missing_c_ratings(xml: str, ratings: dict) -> str:
    """Insert C-column rating text for MS rows that lack it."""
    # Find style from an existing C cell in an MS row
    c_style = None
    for ms_row in _MS_INDICATION:
        m = re.search(rf'<c r="C{ms_row}"[^>]*s="(\d+)"', xml)
        if m:
            c_style = m.group(1)
            break
    if c_style is None:
        c_style = "90"  # fallback (ms_d style)

    for ms_row, ind in sorted(_MS_INDICATION.items()):
        # Skip if C cell already exists
        if re.search(rf'<c r="C{ms_row}"', xml):
            continue

        rating_code = ratings.get(ind, "AVG")
        rating_text = _RATING_TEXT.get(rating_code, "Average Growth")

        # Insert C cell before D cell in the same row
        d_m = re.search(rf'(<c r="D{ms_row}")', xml)
        if d_m:
            c_cell = (
                f'<c r="C{ms_row}" s="{c_style}" t="inlineStr">'
                f"<is><t>{rating_text}</t></is></c>"
            )
            xml = xml[: d_m.start()] + c_cell + xml[d_m.start():]
            print(f"  Filled C{ms_row} = {rating_text} ({ind})")
        else:
            print(f"  WARNING: D{ms_row} not found, cannot insert C cell")

    return xml


# ── read ratings from Peer Views ────────────────────────────────────────────
def _read_ratings(xlsx_path: Path) -> dict:
    """Read indication → rating mapping from Peer Views sheet."""
    try:
        from generate_pipeline import _read_peer_views_ratings
        return _read_peer_views_ratings(xlsx_path)
    except Exception as e:
        print(f"  WARNING: Could not read Peer Views ratings: {e}")
        print(f"  Defaulting all missing to Average Growth")
        return {}


# ── zip patching ─────────────────────────────────────────────────────────────
def _apply_patch(xlsx_path: Path, pipeline_xml: str) -> None:
    """Patch Pipeline sheet in-place, dropping stale calcChain."""
    tmp = xlsx_path.with_suffix(".~patch.xlsx")
    with zipfile.ZipFile(xlsx_path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "xl/calcChain.xml":
                continue  # removed — references stripped below
            if item.filename == _PIPE_SHEET:
                zout.writestr(item, pipeline_xml.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    tmp.replace(xlsx_path)
    sz = xlsx_path.stat().st_size
    print(f"  Written {sz:,} bytes → {xlsx_path}")


# ── main ─────────────────────────────────────────────────────────────────────
def main() -> None:
    if not _XLSX_PATH.exists():
        print(f"ERROR: {_XLSX_PATH} not found")
        sys.exit(1)

    # Backup
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = _XLSX_PATH.with_name(f"DCF_CMPX_pre_revformula_{ts}.xlsx")
    shutil.copy2(_XLSX_PATH, bak)
    print(f"Backup → {bak.name}")

    # Read Peer Views ratings
    ratings = _read_ratings(_XLSX_PATH)
    print(f"Ratings: {ratings}")

    # Read Pipeline sheet XML
    with zipfile.ZipFile(_XLSX_PATH, "r") as z:
        pipeline_xml = z.read(_PIPE_SHEET).decode("utf-8")

    original_rows = len(re.findall(r'<row r="\d+"', pipeline_xml))
    print(f"Pipeline rows: {original_rows}")

    # Step 1: Fill missing C-column ratings
    pipeline_xml = _fill_missing_c_ratings(pipeline_xml, ratings)

    # Steps 2 & 3: Patch Revenue and COGS rows
    for spec in DRUGS:
        rev_row  = spec["rev"]
        cogs_row = spec["cogs"]
        drug_row = spec["drug"]
        terms    = spec["terms"]

        print(f"  Patching Revenue R{rev_row}, COGS R{cogs_row} (drug R{drug_row})")

        # Revenue row: replace F-U cells with new formulas
        rev_cells = _build_row_cells(rev_row, drug_row, terms, is_cogs=False)
        pipeline_xml = _replace_row_cells(pipeline_xml, rev_row, rev_cells)

        # COGS row: replace F-U cells with new formulas
        cogs_cells = _build_row_cells(cogs_row, drug_row, rev_row, is_cogs=True)
        pipeline_xml = _replace_row_cells(pipeline_xml, cogs_row, cogs_cells)

    patched_rows = len(re.findall(r'<row r="\d+"', pipeline_xml))
    print(f"Pipeline rows after patch: {patched_rows}")
    assert original_rows == patched_rows, f"Row count changed: {original_rows} → {patched_rows}"

    # Write patched file
    _apply_patch(_XLSX_PATH, pipeline_xml)
    print("Done.")


if __name__ == "__main__":
    main()
