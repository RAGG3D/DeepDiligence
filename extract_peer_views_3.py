import zipfile
import xml.etree.ElementTree as ET

XLSX = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"

with zipfile.ZipFile(XLSX, 'r') as z:
    sheet_xml = z.read("xl/worksheets/sheet20.xml")
    root = ET.fromstring(sheet_xml)
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    
    rows = root.findall(".//s:sheetData/s:row", ns)
    
    def get_val(cell):
        t = cell.get("t", "")
        v_elem = cell.find("s:v", ns)
        is_elem = cell.find("s:is", ns)
        if t == "inlineStr" and is_elem is not None:
            t_elem = is_elem.find("s:t", ns)
            if t_elem is not None and t_elem.text:
                return t_elem.text
        elif v_elem is not None and v_elem.text:
            return v_elem.text
        return ""

    # Print FULL detail of the RCC 2L section (rows 273-301) 
    # This is a "v4" style section
    print("=" * 80)
    print("=== FULL RCC 2L Post-IO Section (rows 273-300) ===")
    print("=" * 80)
    
    for row in rows:
        r = int(row.get("r"))
        if 273 <= r <= 300:
            cells = row.findall("s:c", ns)
            if not cells:
                print(f"  Row {r}: (empty row tag)")
                continue
            
            # Collect all cells with full detail
            parts = []
            for cell in cells:
                ref = cell.get("r", "")
                s = cell.get("s", "")
                t = cell.get("t", "")
                val = get_val(cell)
                f_elem = cell.find("s:f", ns)
                ftext = ""
                if f_elem is not None and f_elem.text:
                    ftext = f_elem.text[:80]
                
                detail = f"{ref}[s={s}"
                if t: detail += f",t={t}"
                if val: detail += f",v={val!r}"
                if ftext: detail += f",f={ftext!r}"
                detail += "]"
                parts.append(detail)
            
            row_attrs = []
            for attr in ["ht", "hidden", "s", "customFormat", "spans"]:
                v = row.get(attr, "")
                if v: row_attrs.append(f"{attr}={v}")
            attr_str = " " + " ".join(row_attrs) if row_attrs else ""
            
            print(f"  Row {r}:{attr_str}")
            for p in parts:
                print(f"      {p}")
    
    # Also print the BTC section (203-237) - a "v3" style section for comparison
    print("\n" + "=" * 80)
    print("=== FULL BTC Section (rows 203-237) ===")
    print("=" * 80)
    
    for row in rows:
        r = int(row.get("r"))
        if 203 <= r <= 237:
            cells = row.findall("s:c", ns)
            if not cells:
                print(f"  Row {r}: (empty row tag)")
                continue
            
            parts = []
            for cell in cells:
                ref = cell.get("r", "")
                s = cell.get("s", "")
                t = cell.get("t", "")
                val = get_val(cell)
                f_elem = cell.find("s:f", ns)
                ftext = ""
                if f_elem is not None and f_elem.text:
                    ftext = f_elem.text[:80]
                
                detail = f"{ref}[s={s}"
                if t: detail += f",t={t}"
                if val: detail += f",v={val!r}"
                if ftext: detail += f",f={ftext!r}"
                detail += "]"
                parts.append(detail)
            
            row_attrs = []
            for attr in ["ht", "hidden", "s", "customFormat", "spans"]:
                v = row.get(attr, "")
                if v: row_attrs.append(f"{attr}={v}")
            attr_str = " " + " ".join(row_attrs) if row_attrs else ""
            
            print(f"  Row {r}:{attr_str}")
            for p in parts:
                print(f"      {p}")
    
    # 3. Print ALL unique D-column labels across all sections to see the full row template
    print("\n" + "=" * 80)
    print("=== All unique D-column row labels (ordered by first appearance) ===")
    print("=" * 80)
    seen = []
    for row in rows:
        r = int(row.get("r"))
        for cell in row.findall("s:c", ns):
            ref = cell.get("r", "")
            if ref == f"D{r}":
                val = get_val(cell)
                if val and val not in seen:
                    seen.append(val)
                    print(f"  Row {r}: {val!r}")

