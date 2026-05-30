import zipfile
import xml.etree.ElementTree as ET

XLSX = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"

with zipfile.ZipFile(XLSX, 'r') as z:
    sheet_xml = z.read("xl/worksheets/sheet20.xml")
    root = ET.fromstring(sheet_xml)
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    
    rows = root.findall(".//s:sheetData/s:row", ns)
    
    # 1. Find all section headers (rows with indication names in D column)
    print("=== Section Headers (indication title rows) ===")
    print(f"{'Row':>5} | {'Col A':>5} | {'Style':>5} | {'D value'}")
    print("-" * 60)
    
    section_rows = []
    for row in rows:
        r = int(row.get("r"))
        cells = row.findall("s:c", ns)
        d_cell = None
        a_cell = None
        for cell in cells:
            ref = cell.get("r", "")
            if ref == f"D{r}":
                d_cell = cell
            if ref == f"A{r}":
                a_cell = cell
        
        if d_cell is not None:
            s = d_cell.get("s", "")
            # Check for section headers: s=794 or s=474
            if s in ("794", "474"):
                is_elem = d_cell.find("s:is", ns)
                val = ""
                if is_elem is not None:
                    t_elem = is_elem.find("s:t", ns)
                    if t_elem is not None and t_elem.text:
                        val = t_elem.text
                
                a_val = ""
                if a_cell is not None:
                    is_elem_a = a_cell.find("s:is", ns)
                    if is_elem_a is not None:
                        t_elem_a = is_elem_a.find("s:t", ns)
                        if t_elem_a is not None and t_elem_a.text:
                            a_val = t_elem_a.text
                
                section_rows.append((r, a_val, s, val))
                print(f"{r:>5} | {a_val:>5} | s={s:>3} | {val}")
    
    # 2. Now print the structure of each section - first few rows + last row
    print("\n\n=== Section Structure Summary ===")
    for i, (sr, a_val, s, title) in enumerate(section_rows):
        next_sr = section_rows[i+1][0] if i+1 < len(section_rows) else 999
        
        # Count data columns in this section
        # Find the ticker row (usually section_row + 3)
        ticker_row = None
        last_data_row = sr
        for row in rows:
            r = int(row.get("r"))
            if r > sr and r < next_sr:
                cells = row.findall("s:c", ns)
                if cells:
                    last_data_row = r
                # Check for ticker row
                for cell in cells:
                    ref = cell.get("r", "")
                    if ref.startswith("E") and r == sr + 3:
                        ticker_row = r
        
        # Get tickers
        tickers = []
        if ticker_row:
            for row in rows:
                if int(row.get("r")) == ticker_row:
                    for cell in row.findall("s:c", ns):
                        ref = cell.get("r", "")
                        is_elem = cell.find("s:is", ns)
                        if is_elem is not None:
                            t_elem = is_elem.find("s:t", ns)
                            if t_elem is not None and t_elem.text:
                                col = ref.rstrip("0123456789")
                                tickers.append(f"{col}={t_elem.text}")
        
        print(f"\nSection: {title}")
        print(f"  Title row: {sr} (A={a_val!r})")
        print(f"  Last data row: {last_data_row}")
        print(f"  Span: {last_data_row - sr} rows")
        if tickers:
            print(f"  Tickers: {', '.join(tickers)}")
    
    # 3. Print the v4 area in detail (rows 300-340)
    print("\n\n" + "=" * 60)
    print("=== Rows 300-340 (post-CRC area, looking for next sections) ===")
    print("=" * 60)
    for row in rows:
        r = int(row.get("r"))
        if 298 <= r <= 340:
            hidden = row.get("hidden", "")
            cells = row.findall("s:c", ns)
            if not cells:
                print(f"  Row {r}: (empty)")
                continue
            cell_info = []
            for cell in cells:
                ref = cell.get("r", "")
                s = cell.get("s", "")
                t = cell.get("t", "")
                is_elem = cell.find("s:is", ns)
                v_elem = cell.find("s:v", ns)
                f_elem = cell.find("s:f", ns)
                
                val = ""
                if t == "inlineStr" and is_elem is not None:
                    t_elem = is_elem.find("s:t", ns)
                    if t_elem is not None and t_elem.text:
                        val = t_elem.text
                elif v_elem is not None and v_elem.text:
                    val = v_elem.text
                
                fstr = ""
                if f_elem is not None and f_elem.text:
                    fstr = f" f={f_elem.text[:60]}"
                
                cell_info.append(f"{ref}(s={s},t={t},v={val!r}{fstr})")
            
            hstr = " HIDDEN" if hidden else ""
            print(f"  Row {r}:{hstr} {' | '.join(cell_info)}")

