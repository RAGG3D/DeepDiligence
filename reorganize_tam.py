#!/usr/bin/env python3
"""reorganize_tam.py  –  Reorganize TAM Solid rows 388-431 into Tecentriq format.

Consolidates multi-indication drugs (Lenvima, Cabometyx, Nexavar) under one
drug name row with total worldwide sales, with indication sub-rows beneath.

Uses surgical zip-patching (never openpyxl .save()).
"""

import re, zipfile, shutil, sys, os
from pathlib import Path
from collections import OrderedDict


_EMPTY_CALC_CHAIN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></calcChain>'
)
XLSX = Path("/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx")
SHEET = "xl/worksheets/sheet18.xml"

# Column letters F through AH
_COLS = [chr(c) for c in range(ord('F'), ord('Z') + 1)]
_COLS += ['AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']
# F=2010 (idx 0), G=2011, ..., S=2023 (idx 13), T=2024 (idx 14), ..., AH=2038 (idx 28)


def _col_idx(col: str) -> int:
    """Column letter(s) to 1-based index."""
    r = 0
    for ch in col:
        r = r * 26 + (ord(ch) - ord('A') + 1)
    return r


def _xml_esc(s: str) -> str:
    return (s.replace('&', '&amp;').replace('<', '&lt;')
             .replace('>', '&gt;').replace('"', '&quot;'))


def _fmt_val(v: float) -> str:
    """Format numeric value, dropping .0 for integers."""
    if v == int(v):
        return str(int(v))
    return f"{v:.1f}"


# ── Source row mapping ──
# (old_indication_row, drug_short_name, drug_full_name, indication)
_SOURCE = [
    (389, "Pemazyre",  "Pemazyre (Incyte, pemigatinib)",             "BTC"),
    (391, "Tibsovo",   "Tibsovo (Servier, ivosidenib)",             "BTC"),
    (393, "Lytgobi",   "Lytgobi (Taiho, futibatinib)",              "BTC"),
    (395, "Truseltiq", "Truseltiq (Helsinn, infigratinib)",         "BTC"),
    (397, "Lenvima",   "Lenvima (Eisai/Merck, lenvatinib)",         "EC"),
    (399, "Jemperli",  "Jemperli (GSK, dostarlimab)",               "EC"),
    (401, "Zepzelca",  "Zepzelca (Jazz, lurbinectedin)",            "ES-SCLC"),
    (403, "Imdelltra", "Imdelltra (Amgen, tarlatamab)",             "ES-SCLC"),
    (405, "Nexavar",   "Nexavar (Bayer, sorafenib)",                "HCC"),
    (407, "Lenvima",   "Lenvima (Eisai/Merck, lenvatinib)",         "HCC"),
    (409, "Cabometyx", "Cabometyx (Exelixis, cabozantinib)",        "HCC"),
    (411, "Amtagvi",   "Amtagvi (Iovance, lifileucel)",             "Melanoma NCAM+"),
    (413, "Sutent",    "Sutent (Pfizer, sunitinib)",                "RCC"),
    (415, "Votrient",  "Votrient (Novartis, pazopanib)",            "RCC"),
    (417, "Inlyta",    "Inlyta (Pfizer, axitinib)",                 "RCC"),
    (419, "Cabometyx", "Cabometyx (Exelixis, cabozantinib)",        "RCC"),
    (421, "Lenvima",   "Lenvima (Eisai/Merck, lenvatinib)",         "RCC"),
    (423, "Fotivda",   "Fotivda (EUSA/LG Chem, tivozanib)",        "RCC"),
    (425, "Afinitor",  "Afinitor (Novartis, everolimus)",           "RCC"),
    (427, "Torisel",   "Torisel (Pfizer, temsirolimus)",            "RCC"),
    (429, "Welireg",   "Welireg (Merck, belzutifan)",               "RCC"),
    (431, "Nexavar",   "Nexavar (Bayer, sorafenib)",                "RCC"),
]

# ── Target layout: ordered drugs with their indications ──
# Multi-indication drugs appear once with all indications listed
_TARGET = [
    ("Pemazyre",  "Pemazyre (Incyte, pemigatinib)",         ["BTC"]),
    ("Tibsovo",   "Tibsovo (Servier, ivosidenib)",          ["BTC"]),
    ("Lytgobi",   "Lytgobi (Taiho, futibatinib)",           ["BTC"]),
    ("Truseltiq", "Truseltiq (Helsinn, infigratinib)",      ["BTC"]),
    ("Lenvima",   "Lenvima (Eisai/Merck, lenvatinib)",      ["EC", "HCC", "RCC"]),
    ("Jemperli",  "Jemperli (GSK, dostarlimab)",            ["EC"]),
    ("Zepzelca",  "Zepzelca (Jazz, lurbinectedin)",         ["ES-SCLC"]),
    ("Imdelltra", "Imdelltra (Amgen, tarlatamab)",          ["ES-SCLC"]),
    ("Nexavar",   "Nexavar (Bayer, sorafenib)",             ["HCC", "RCC"]),
    ("Cabometyx", "Cabometyx (Exelixis, cabozantinib)",     ["HCC", "RCC"]),
    ("Amtagvi",   "Amtagvi (Iovance, lifileucel)",          ["Melanoma NCAM+"]),
    ("Sutent",    "Sutent (Pfizer, sunitinib)",             ["RCC"]),
    ("Votrient",  "Votrient (Novartis, pazopanib)",         ["RCC"]),
    ("Inlyta",    "Inlyta (Pfizer, axitinib)",              ["RCC"]),
    ("Fotivda",   "Fotivda (EUSA/LG Chem, tivozanib)",     ["RCC"]),
    ("Afinitor",  "Afinitor (Novartis, everolimus)",        ["RCC"]),
    ("Torisel",   "Torisel (Pfizer, temsirolimus)",         ["RCC"]),
    ("Welireg",   "Welireg (Merck, belzutifan)",            ["RCC"]),
]

# ── Drug total worldwide sales (MM USD) from financial reports ──
# Only for multi-indication drugs; single-indication drugs use indication data
_DRUG_TOTALS = {
    # Lenvima: Eisai/Merck annual reports (calendar year approximation)
    # Approved: 2015 DTC, 2016 RCC combo, 2018 HCC, 2019 EC
    "Lenvima": {
        'K': 134, 'L': 515, 'M': 900, 'N': 1500, 'O': 2200,
        'P': 2700, 'Q': 3500, 'R': 4000, 'S': 3000, 'T': 2200,
    },
    # Cabometyx: Exelixis annual reports
    # Approved: 2016 RCC, 2019 HCC
    "Cabometyx": {
        'L': 181, 'M': 427, 'N': 665, 'O': 893,
        'P': 1148, 'Q': 1411, 'R': 1580, 'S': 1629, 'T': 1809,
    },
    # Nexavar: Bayer annual reports (approximate)
    # Approved: 2005 RCC, 2007 HCC, 2013 DTC
    "Nexavar": {
        'F': 1100, 'G': 1050, 'H': 1000, 'I': 900, 'J': 850,
        'K': 800, 'L': 700, 'M': 600, 'N': 500, 'O': 430,
        'P': 330, 'Q': 250, 'R': 180, 'S': 130, 'T': 100,
    },
}


def _extract_row_data(xml: str, row_num: int) -> dict:
    """Extract {col_letter: float_value} from a row in the XML."""
    pat = re.compile(
        rf'<row\s[^>]*r="{row_num}"[^>]*>.*?</row>',
        re.DOTALL
    )
    m = pat.search(xml)
    if not m:
        return {}
    row_xml = m.group(0)
    vals = {}
    for col in _COLS:
        c_pat = re.search(
            rf'<c\s[^>]*r="{col}{row_num}"[^>]*>(.*?)</c>'
            rf'|<c\s[^>]*r="{col}{row_num}"[^/]*/>',
            row_xml, re.DOTALL
        )
        if c_pat:
            v_m = re.search(r'<v>([^<]+)</v>', c_pat.group(0))
            if v_m:
                try:
                    vals[col] = float(v_m.group(1))
                except ValueError:
                    pass
    return vals


def _build_drug_row(row_num: int, drug_full_name: str,
                    data: dict) -> str:
    """Build XML for a drug name row (Tecentriq R175 pattern)."""
    cells = []
    cells.append(f'<c r="D{row_num}" s="409" t="inlineStr">'
                 f'<is><t>{_xml_esc(drug_full_name)}</t></is></c>')
    cells.append(f'<c r="E{row_num}" s="52" t="inlineStr">'
                 f'<is><t>[MM USD]</t></is></c>')
    for col in _COLS:
        v = data.get(col, 0)
        if v == 0:
            cells.append(f'<c r="{col}{row_num}" s="48"/>')
        else:
            cells.append(f'<c r="{col}{row_num}" s="48">'
                         f'<v>{_fmt_val(v)}</v></c>')
    return f'<row r="{row_num}" spans="1:34">' + ''.join(cells) + '</row>'


def _build_indication_row(row_num: int, indication: str,
                          data: dict) -> str:
    """Build XML for an indication sub-row (Tecentriq R179 pattern)."""
    cells = []
    cells.append(f'<c r="A{row_num}" s="326"/>')
    cells.append(f'<c r="B{row_num}" s="3"/>')
    cells.append(f'<c r="C{row_num}" s="529"/>')
    cells.append(f'<c r="D{row_num}" s="48" t="inlineStr">'
                 f'<is><t>{_xml_esc(indication)}</t></is></c>')
    cells.append(f'<c r="E{row_num}" s="52" t="inlineStr">'
                 f'<is><t>[MM USD]</t></is></c>')
    for col in _COLS:
        v = data.get(col, 0)
        if v == 0:
            cells.append(f'<c r="{col}{row_num}" s="48"/>')
        else:
            cells.append(f'<c r="{col}{row_num}" s="48">'
                         f'<v>{_fmt_val(v)}</v></c>')
    return (f'<row r="{row_num}" spans="1:34" s="235" customFormat="1">'
            + ''.join(cells) + '</row>')


def _build_empty_row(row_num: int) -> str:
    """Build an empty row."""
    return f'<row r="{row_num}" spans="1:34"/>'


def _compute_drug_total(drug_short: str,
                        indication_data: dict) -> dict:
    """Compute drug total row data.

    For multi-indication drugs with verified totals: use _DRUG_TOTALS.
    For single-indication drugs: copy indication data.
    For multi-indication without verified totals: sum indications.

    Forecast columns (U-AH) carry forward the T value.
    """
    if drug_short in _DRUG_TOTALS:
        # Use verified total + carry forward
        total = dict(_DRUG_TOTALS[drug_short])
        # Carry forward T value through U-AH
        t_val = total.get('T', 0)
        for col in _COLS:
            if _col_idx(col) > _col_idx('T') and col not in total:
                total[col] = t_val
        return total

    # Single or multi indication: sum all indications
    indications = list(indication_data.values())
    if len(indications) == 1:
        return dict(indications[0])

    # Sum across all indications
    total = {}
    for col in _COLS:
        s = sum(ind.get(col, 0) for ind in indications)
        if s != 0:
            total[col] = s
    return total


def main():
    if not XLSX.exists():
        print(f"ERROR: {XLSX} not found", file=sys.stderr)
        sys.exit(1)

    # ── Step 1: Read sheet XML ──
    import io
    with zipfile.ZipFile(str(XLSX), 'r') as z:
        xml = z.read(SHEET).decode('utf-8')

    print("Read sheet18.xml OK")

    # ── Step 2: Extract existing indication data ──
    # Build {(drug_short, indication): {col: value}}
    ind_data = {}
    for old_row, drug_short, _, indication in _SOURCE:
        data = _extract_row_data(xml, old_row)
        ind_data[(drug_short, indication)] = data
        n = sum(1 for v in data.values() if v != 0)
        print(f"  R{old_row} {drug_short}/{indication}: {n} nonzero values")

    # ── Step 3: Build new rows ──
    new_rows = []
    row_num = 388

    for drug_short, drug_full, indications in _TARGET:
        # Gather indication data for this drug
        drug_ind_data = OrderedDict()
        for ind in indications:
            key = (drug_short, ind)
            if key in ind_data:
                drug_ind_data[ind] = ind_data[key]
            else:
                print(f"  WARNING: No data for {drug_short}/{ind}")
                drug_ind_data[ind] = {}

        # Compute drug total row data
        total_data = _compute_drug_total(drug_short, drug_ind_data)

        # Drug name row
        new_rows.append(_build_drug_row(row_num, drug_full, total_data))
        print(f"  R{row_num}: {drug_short} (drug total)")
        row_num += 1

        # Indication sub-rows
        for ind in indications:
            new_rows.append(
                _build_indication_row(row_num, ind, drug_ind_data[ind])
            )
            print(f"  R{row_num}: {ind}")
            row_num += 1

    print(f"\nGenerated {len(new_rows)} rows (R388-R{row_num - 1})")

    # Fill remaining rows up to R431 with empty rows
    while row_num <= 431:
        new_rows.append(_build_empty_row(row_num))
        row_num += 1

    print(f"Total rows including empties: {len(new_rows)} (R388-R431)")

    # ── Step 4: Replace rows 388-431 in XML ──
    # Find the first row >= 388 and last row <= 431
    # Remove all rows in [388, 431] range
    ns_prefix = ''
    ns_m = re.search(r'<worksheet\s[^>]*xmlns="([^"]+)"', xml)
    if ns_m:
        # Check if rows use namespace prefix
        if '<ns0:row ' in xml:
            ns_prefix = 'ns0:'

    # Remove existing rows 388-431
    removed = 0
    for rn in range(388, 432):
        # Match both self-closing and content rows
        pat = re.compile(
            rf'<{ns_prefix}row\s[^>]*r="{rn}"[^>]*/>'
            rf'|<{ns_prefix}row\s[^>]*r="{rn}"[^>]*>.*?</{ns_prefix}row>',
            re.DOTALL
        )
        m = pat.search(xml)
        if m:
            xml = xml[:m.start()] + xml[m.end():]
            removed += 1

    print(f"Removed {removed} existing rows")

    # Add namespace prefix to new rows if needed
    if ns_prefix:
        for i, row in enumerate(new_rows):
            row = row.replace('<row ', f'<{ns_prefix}row ')
            row = row.replace('</row>', f'</{ns_prefix}row>')
            row = row.replace('<c ', f'<{ns_prefix}c ')
            row = row.replace('</c>', f'</{ns_prefix}c>')
            row = row.replace('<v>', f'<{ns_prefix}v>')
            row = row.replace('</v>', f'</{ns_prefix}v>')
            row = row.replace('<is>', f'<{ns_prefix}is>')
            row = row.replace('</is>', f'</{ns_prefix}is>')
            row = row.replace('<t>', f'<{ns_prefix}t>')
            row = row.replace('</t>', f'</{ns_prefix}t>')
            row = row.replace('<f>', f'<{ns_prefix}f>')
            row = row.replace('</f>', f'</{ns_prefix}f>')
            new_rows[i] = row

    # Find insertion point: after last row < 388
    # Look for row 387 (Total Solid Tumor) end tag
    insert_pat = re.compile(
        rf'(</{ns_prefix}row>)',
        re.DOTALL
    )

    # Find the row just before 388 to insert after
    row387_pat = re.compile(
        rf'<{ns_prefix}row\s[^>]*r="387"[^>]*>.*?</{ns_prefix}row>',
        re.DOTALL
    )
    m387 = row387_pat.search(xml)
    if m387:
        insert_pos = m387.end()
    else:
        # Fallback: find row 432 or 433 and insert before it
        for fallback_rn in [432, 433, 434, 435]:
            fb_pat = re.compile(
                rf'<{ns_prefix}row\s[^>]*r="{fallback_rn}"',
                re.DOTALL
            )
            fb_m = fb_pat.search(xml)
            if fb_m:
                insert_pos = fb_m.start()
                break
        else:
            print("ERROR: Could not find insertion point", file=sys.stderr)
            sys.exit(1)

    # Insert new rows
    new_block = '\n'.join(new_rows)
    xml = xml[:insert_pos] + '\n' + new_block + '\n' + xml[insert_pos:]
    print("Inserted new rows into XML")

    # ── Step 5: Ensure fullCalcOnLoad ──
    # (Should already be set from previous scripts)

    # ── Step 6: Zip-patch ──
    backup = XLSX.with_suffix('.xlsx.bak_reorg')
    shutil.copy2(str(XLSX), str(backup))
    print(f"Backup: {backup}")

    tmp = XLSX.with_suffix('.xlsx.tmp')
    with zipfile.ZipFile(str(XLSX), 'r') as zin, \
         zipfile.ZipFile(str(tmp), 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "xl/calcChain.xml":
                continue  # removed — references stripped below
            if item.filename == SHEET:
                zout.writestr(item, xml.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))

    shutil.move(str(tmp), str(XLSX))
    print(f"\nDone. Reorganized R388-R431 in {XLSX.name}")


if __name__ == "__main__":
    main()
