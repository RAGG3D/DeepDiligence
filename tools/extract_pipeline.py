import zipfile
import xml.etree.ElementTree as ET
import re
import os

XLSX_PATH = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"

def col_letter(col_num):
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result

def col_num_from_letter(letters):
    num = 0
    for ch in letters.upper():
        num = num * 26 + (ord(ch) - 64)
    return num

def get_cell_value(cell_elem, shared_strings, ns):
    t = cell_elem.get('t', '')
    v_elem = cell_elem.find(f'{ns}v')
    is_elem = cell_elem.find(f'{ns}is')
    f_elem = cell_elem.find(f'{ns}f')
    
    value = None
    if t == 's' and v_elem is not None:
        idx = int(v_elem.text)
        if idx < len(shared_strings):
            value = shared_strings[idx]
    elif t == 'inlineStr' and is_elem is not None:
        t_elem = is_elem.find(f'{ns}t')
        if t_elem is not None and t_elem.text:
            value = t_elem.text
        else:
            parts = []
            for r in is_elem.findall(f'{ns}r'):
                rt = r.find(f'{ns}t')
                if rt is not None and rt.text:
                    parts.append(rt.text)
            if parts:
                value = ''.join(parts)
    elif v_elem is not None:
        value = v_elem.text
    
    formula = None
    if f_elem is not None:
        formula = f_elem.text
    
    return value, formula, t

def parse_cell_ref(ref):
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    if m:
        return m.group(1), int(m.group(2))
    return None, None

def main():
    if not os.path.exists(XLSX_PATH):
        print(f"File not found: {XLSX_PATH}")
        return
    
    with zipfile.ZipFile(XLSX_PATH, 'r') as zf:
        # Step 1: Find Pipeline sheet file
        print("=" * 80)
        print("STEP 1: Finding Pipeline sheet")
        print("=" * 80)
        
        wb_xml = zf.read('xl/workbook.xml')
        wb_tree = ET.fromstring(wb_xml)
        wb_ns_match = re.match(r'\{([^}]+)\}', wb_tree.tag)
        wb_ns = f'{{{wb_ns_match.group(1)}}}' if wb_ns_match else ''
        
        sheets_info = []
        for sheet in wb_tree.findall(f'.//{wb_ns}sheet'):
            name = sheet.get('name')
            sheet_id = sheet.get('sheetId')
            r_id = None
            for attr_name, attr_val in sheet.attrib.items():
                if 'id' in attr_name.lower() and attr_val.startswith('rId'):
                    r_id = attr_val
            sheets_info.append((name, sheet_id, r_id))
            print(f"  Sheet: '{name}', sheetId={sheet_id}, rId={r_id}")
        
        # Parse rels
        rels_xml = zf.read('xl/_rels/workbook.xml.rels')
        rels_tree = ET.fromstring(rels_xml)
        rid_to_file = {}
        for rel in rels_tree:
            rid = rel.get('Id')
            target = rel.get('Target')
            rid_to_file[rid] = target
        
        print(f"\n  rId-to-file mappings (sheets only):")
        for rid, target in sorted(rid_to_file.items()):
            if 'sheet' in target.lower():
                print(f"    {rid} -> {target}")
        
        # Find Pipeline
        pipeline_rid = None
        for name, sid, rid in sheets_info:
            if name == 'Pipeline':
                pipeline_rid = rid
                break
        
        pipeline_target = rid_to_file.get(pipeline_rid, '')
        # Normalize path
        if pipeline_target.startswith('/xl/'):
            full_path = pipeline_target[1:]  # remove leading /
        elif pipeline_target.startswith('xl/'):
            full_path = pipeline_target
        else:
            full_path = f'xl/{pipeline_target}'
        
        print(f"\n  Pipeline rId={pipeline_rid}, target='{pipeline_target}', full_path='{full_path}'")
        
        if full_path not in zf.namelist():
            print(f"  '{full_path}' not found. Listing sheet files:")
            for name in sorted(zf.namelist()):
                if 'sheet' in name.lower() and name.endswith('.xml'):
                    print(f"    {name}")
            # Fallback: use sheet9.xml directly
            full_path = 'xl/worksheets/sheet9.xml'
            print(f"  Falling back to: {full_path}")
        
        print(f"  Using: {full_path}")
        
        # Load shared strings
        print("\n" + "=" * 80)
        print("STEP 2: Loading shared strings")
        print("=" * 80)
        
        shared_strings = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            ss_xml = zf.read('xl/sharedStrings.xml')
            ss_tree = ET.fromstring(ss_xml)
            ss_ns_match = re.match(r'\{([^}]+)\}', ss_tree.tag)
            ss_ns = f'{{{ss_ns_match.group(1)}}}' if ss_ns_match else ''
            
            for si in ss_tree.findall(f'{ss_ns}si'):
                t_elem = si.find(f'{ss_ns}t')
                if t_elem is not None and t_elem.text:
                    shared_strings.append(t_elem.text)
                else:
                    parts = []
                    for r in si.findall(f'{ss_ns}r'):
                        rt = r.find(f'{ss_ns}t')
                        if rt is not None and rt.text:
                            parts.append(rt.text)
                    shared_strings.append(''.join(parts) if parts else '')
            print(f"  Loaded {len(shared_strings)} shared strings")
        else:
            print("  No sharedStrings.xml found")
        
        # Parse Pipeline sheet
        print("\n" + "=" * 80)
        print("STEP 3: Parsing Pipeline sheet XML")
        print("=" * 80)
        
        sheet_xml = zf.read(full_path)
        sheet_tree = ET.fromstring(sheet_xml)
        ns_match = re.match(r'\{([^}]+)\}', sheet_tree.tag)
        ns = f'{{{ns_match.group(1)}}}' if ns_match else ''
        
        dim = sheet_tree.find(f'{ns}dimension')
        if dim is not None:
            print(f"  Dimension: {dim.get('ref')}")
        
        all_rows = sheet_tree.findall(f'.//{ns}sheetData/{ns}row')
        print(f"  Total rows with data: {len(all_rows)}")
        if all_rows:
            last_row = max(int(r.get('r', 0)) for r in all_rows)
            print(f"  Max row number: {last_row}")
        else:
            last_row = 0
        
        # Build cells dict
        cells = {}
        for row in all_rows:
            row_num = int(row.get('r', 0))
            for cell in row.findall(f'{ns}c'):
                ref = cell.get('r', '')
                col_let, r_num = parse_cell_ref(ref)
                if col_let is None:
                    continue
                value, formula, t = get_cell_value(cell, shared_strings, ns)
                style = cell.get('s', '')
                cells[(row_num, col_let)] = {
                    'ref': ref,
                    'value': value,
                    'formula': formula,
                    'type': t,
                    'style': style
                }
        
        # Step 4: Rows 378-385
        print("\n" + "=" * 80)
        print("STEP 4: Rows 378-385 (expected Revenue Forecasting area)")
        print("=" * 80)
        
        for row_num in range(378, 386):
            row_cells = {k: v for k, v in cells.items() if k[0] == row_num}
            if row_cells:
                sorted_cells = sorted(row_cells.items(), key=lambda x: col_num_from_letter(x[0][1]))
                print(f"\n  Row {row_num} ({len(sorted_cells)} cells):")
                for (rn, cl), info in sorted_cells:
                    val_display = repr(info['value']) if info['value'] is not None else '(empty)'
                    formula_display = f"  FORMULA: {info['formula']}" if info['formula'] else ''
                    print(f"    {info['ref']:6s} s={info['style']:4s} t={info['type']:12s} val={val_display}{formula_display}")
            else:
                print(f"\n  Row {row_num}: (no cells)")
        
        # Step 5: Find year header rows near 378-385
        print("\n" + "=" * 80)
        print("STEP 5: Year header row search (rows 375-395)")
        print("=" * 80)
        
        for test_row in range(375, 396):
            row_data = {k: v for k, v in cells.items() if k[0] == test_row}
            if row_data:
                vals = [(k[1], v['value']) for k, v in row_data.items() if v['value'] is not None]
                year_like = []
                for col, val in vals:
                    try:
                        yr = int(float(str(val)))
                        if 2005 <= yr <= 2045:
                            year_like.append((col, yr))
                    except (ValueError, TypeError):
                        pass
                if year_like:
                    print(f"\n  Row {test_row} has YEAR values:")
                    sorted_r = sorted(row_data.items(), key=lambda x: col_num_from_letter(x[0][1]))
                    for (rn, cl), info in sorted_r:
                        val = repr(info['value']) if info['value'] is not None else '(empty)'
                        print(f"    Col {cl:3s} ({info['ref']:6s}): {val}")
        
        # Step 6: First 15 rows after 378 with data
        print("\n" + "=" * 80)
        print("STEP 6: First 15 data rows after row 378")
        print("=" * 80)
        
        count = 0
        for row_num in range(379, min(last_row + 1, 500)):
            if count >= 15:
                break
            row_cells = {k: v for k, v in cells.items() if k[0] == row_num}
            if row_cells:
                count += 1
                sorted_cells = sorted(row_cells.items(), key=lambda x: col_num_from_letter(x[0][1]))
                print(f"\n  Row {row_num} ({len(sorted_cells)} cells):")
                for (rn, cl), info in sorted_cells:
                    val_display = repr(info['value']) if info['value'] is not None else '(empty)'
                    formula_display = f"  FORMULA: {info['formula']}" if info['formula'] else ''
                    print(f"    {info['ref']:6s} s={info['style']:4s} t={info['type']:12s} val={val_display}{formula_display}")
        
        # Step 7: Search for Revenue/Forecast text
        print("\n" + "=" * 80)
        print("STEP 7: Search for 'Revenue' or 'Forecast' in cell text")
        print("=" * 80)
        
        found = []
        for (row_num, col_let), info in sorted(cells.items()):
            val = info['value']
            if val and isinstance(val, str):
                vl = val.lower()
                if 'revenue' in vl or 'forecast' in vl:
                    found.append((row_num, col_let, info))
        
        if found:
            for row_num, col_let, info in found:
                print(f"  {info['ref']:8s} (row {row_num:4d}, col {col_let:3s}): '{info['value']}'")
        else:
            print("  No cells found containing 'Revenue' or 'Forecast'")
            print("  Searching for any text cells to understand sheet content...")
            text_cells = [(k, v) for k, v in cells.items() if v['value'] and isinstance(v['value'], str) and len(v['value']) > 2]
            text_cells.sort()
            print(f"  Found {len(text_cells)} text cells. First 30:")
            for (rn, cl), info in text_cells[:30]:
                print(f"    {info['ref']:8s}: '{info['value'][:60]}'")
        
        # Step 8: Structure rows 370-400
        print("\n" + "=" * 80)
        print("STEP 8: Structure rows 370-400")
        print("=" * 80)
        
        for row_num in range(370, 401):
            row_cells = {k: v for k, v in cells.items() if k[0] == row_num}
            if row_cells:
                sorted_cells = sorted(row_cells.items(), key=lambda x: col_num_from_letter(x[0][1]))
                summary_parts = []
                for (rn, cl), info in sorted_cells[:10]:
                    val = info['value']
                    if val is not None:
                        val_str = repr(str(val)[:35])
                    elif info['formula']:
                        val_str = f"={info['formula'][:30]}"
                    else:
                        val_str = "(styled)"
                    summary_parts.append(f"{cl}={val_str}")
                extra = f" +{len(sorted_cells)-10} more" if len(sorted_cells) > 10 else ""
                print(f"  R{row_num:3d} ({len(sorted_cells):2d}c): {', '.join(summary_parts)}{extra}")

if __name__ == '__main__':
    main()
