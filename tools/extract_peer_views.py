import zipfile
import xml.etree.ElementTree as ET

XLSX = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"

with zipfile.ZipFile(XLSX, 'r') as z:
    # List all files to find sharedStrings variant
    all_files = z.namelist()
    ss_candidates = [f for f in all_files if "sharedstring" in f.lower() or "shared" in f.lower()]
    print(f"Shared string candidates: {ss_candidates}")
    
    # Parse workbook.xml
    wb_xml = z.read("xl/workbook.xml")
    wb_root = ET.fromstring(wb_xml)
    ns_wb = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    
    sheet_rid = None
    for sheet in wb_root.findall(".//s:sheet", ns_wb):
        name = sheet.get("name")
        rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if name == "Peer Views":
            sheet_rid = rid
    
    # Parse rels
    rels_xml = z.read("xl/_rels/workbook.xml.rels")
    rels_root = ET.fromstring(rels_xml)
    
    sheet_file = None
    for rel in rels_root:
        if rel.get("Id") == sheet_rid:
            target = rel.get("Target")
            sheet_file = "xl/" + target if not target.startswith("xl/") and not target.startswith("/") else target.lstrip("/")
            break
    
    print(f"Sheet file: {sheet_file}")
    
    # Read shared strings if they exist
    shared_strings = []
    for candidate in ["xl/sharedStrings.xml", "xl/SharedStrings.xml", "xl/sharedstrings.xml"]:
        if candidate in all_files:
            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            ss_xml = z.read(candidate)
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.findall("s:si", ns):
                texts = []
                t_elem = si.find("s:t", ns)
                if t_elem is not None and t_elem.text:
                    texts.append(t_elem.text)
                else:
                    for r in si.findall("s:r", ns):
                        t2 = r.find("s:t", ns)
                        if t2 is not None and t2.text:
                            texts.append(t2.text)
                shared_strings.append("".join(texts))
            print(f"Loaded {len(shared_strings)} shared strings from {candidate}")
            break
    else:
        print("No sharedStrings.xml found - will show raw indices for t='s' cells")
    
    # Parse the sheet
    sheet_xml = z.read(sheet_file)
    root = ET.fromstring(sheet_xml)
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    
    def get_cell_value(cell):
        t = cell.get("t", "")
        v_elem = cell.find("s:v", ns)
        f_elem = cell.find("s:f", ns)
        is_elem = cell.find("s:is", ns)
        
        val = ""
        if t == "s" and v_elem is not None and v_elem.text:
            idx = int(v_elem.text)
            if shared_strings and idx < len(shared_strings):
                val = shared_strings[idx]
            else:
                val = f"[ss#{idx}]"
        elif t == "inlineStr" and is_elem is not None:
            t_elem = is_elem.find("s:t", ns)
            if t_elem is not None and t_elem.text:
                val = t_elem.text
        elif t == "str" and v_elem is not None and v_elem.text:
            val = v_elem.text
        elif v_elem is not None and v_elem.text:
            val = v_elem.text
        
        formula = ""
        if f_elem is not None:
            formula = f_elem.text or ""
            # Check for shared formula
            sf_t = f_elem.get("t", "")
            sf_si = f_elem.get("si", "")
            sf_ref = f_elem.get("ref", "")
            if sf_t == "shared":
                if sf_ref:
                    formula = f"[shared si={sf_si} ref={sf_ref}] {formula}"
                else:
                    formula = f"[shared si={sf_si}]"
        
        return val, formula
    
    def print_row(row_elem):
        r = row_elem.get("r")
        hidden = row_elem.get("hidden", "")
        spans = row_elem.get("spans", "")
        ht = row_elem.get("ht", "")
        custom_fmt = row_elem.get("customFormat", "")
        s_attr = row_elem.get("s", "")
        
        attrs = []
        if hidden: attrs.append(f"hidden={hidden}")
        if spans: attrs.append(f"spans={spans}")
        if ht: attrs.append(f"ht={ht}")
        if s_attr: attrs.append(f"s={s_attr}")
        if custom_fmt: attrs.append(f"customFormat={custom_fmt}")
        attr_str = " ".join(attrs)
        
        cells = row_elem.findall("s:c", ns)
        if not cells:
            print(f"  Row {r}: (empty) [{attr_str}]")
            return
        
        print(f"  Row {r}: [{attr_str}]")
        for cell in cells:
            ref = cell.get("r", "")
            s = cell.get("s", "")
            t = cell.get("t", "")
            val, formula = get_cell_value(cell)
            
            parts = [f"    {ref}"]
            if s: parts.append(f"s={s}")
            if t: parts.append(f"t={t}")
            if val: 
                # Truncate long values
                display = val if len(val) < 80 else val[:77] + "..."
                parts.append(f"val={display!r}")
            if formula: 
                display = formula if len(formula) < 100 else formula[:97] + "..."
                parts.append(f"f={display!r}")
            print("  ".join(parts))
    
    # Collect all rows
    rows = root.findall(".//s:sheetData/s:row", ns)
    print(f"\n=== Total rows in sheet: {len(rows)} ===")
    
    max_row = 0
    for row in rows:
        r = int(row.get("r"))
        if r > max_row:
            max_row = r
    print(f"Max row number: {max_row}")
    
    # Print requested ranges
    ranges = [
        (1, 10, "Rows 1-10 (Header area)"),
        (195, 215, "Rows 195-215 (BTC section area)"),
        (230, 250, "Rows 230-250 (CRC section area)"),
        (265, 300, "Rows 265-300 (potential v4 area)"),
    ]
    
    for start, end, label in ranges:
        print(f"\n{'='*60}")
        print(f"=== {label} ===")
        print(f"{'='*60}")
        found = False
        for row in rows:
            r = int(row.get("r"))
            if start <= r <= end:
                print_row(row)
                found = True
        if not found:
            print(f"  (no rows found in range {start}-{end})")
    
    # Show last 20 rows
    print(f"\n{'='*60}")
    print(f"=== Last 20 rows (rows {max_row-19} to {max_row}) ===")
    print(f"{'='*60}")
    for row in rows:
        r = int(row.get("r"))
        if r >= max_row - 19:
            print_row(row)
    
    # Also check merge cells and dimension
    dim = root.find("s:dimension", ns)
    if dim is not None:
        print(f"\n=== Dimension: {dim.get('ref')} ===")
    
    merges = root.findall(".//s:mergeCells/s:mergeCell", ns)
    if merges:
        print(f"\n=== Merge cells ({len(merges)} total) ===")
        for mc in merges[:50]:  # First 50
            print(f"  {mc.get('ref')}")
        if len(merges) > 50:
            print(f"  ... and {len(merges)-50} more")

