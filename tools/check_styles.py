import zipfile
import xml.etree.ElementTree as ET

XLSX = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"

with zipfile.ZipFile(XLSX, 'r') as zf:
    styles_xml = zf.read('xl/styles.xml')

root = ET.fromstring(styles_xml)
ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# ── 1. Extract fills ──
fills_el = root.find('x:fills', ns)
all_fills = fills_el.findall('x:fill', ns)
print("=" * 70)
print("FILLS (IDs 16, 17, 18)")
print("=" * 70)
for fid in [16, 17, 18]:
    fill = all_fills[fid]
    pf = fill.find('x:patternFill', ns)
    if pf is not None:
        pat_type = pf.get('patternType', '(none)')
        fg = pf.find('x:fgColor', ns)
        bg = pf.find('x:bgColor', ns)
        fg_str = dict(fg.attrib) if fg is not None else None
        bg_str = dict(bg.attrib) if bg is not None else None
        print(f"  Fill {fid}: patternType={pat_type}, fgColor={fg_str}, bgColor={bg_str}")
    else:
        print(f"  Fill {fid}: (no patternFill)")

# ── 2. Extract cellXfs ──
cellxfs_el = root.find('x:cellXfs', ns)
all_xfs = cellxfs_el.findall('x:xf', ns)
print(f"\nTotal cellXfs entries: {len(all_xfs)}")

target_ids = [
    847, 655, 654, 862, 648, 647, 850, 683, 661,
    852, 864, 669, 668, 854, 709, 708, 855, 703, 702,
    479, 675, 674, 858, 856, 843, 470, 471, 472, 473
]

print("\n" + "=" * 70)
print("CELLXFS ENTRIES")
print("=" * 70)
print(f"{'StyleID':>8} {'numFmtId':>9} {'fillId':>7} {'fontId':>7} {'borderId':>9}")
print("-" * 45)

xf_data = {}
for sid in sorted(target_ids):
    if sid < len(all_xfs):
        xf = all_xfs[sid]
        nf = xf.get('numFmtId', '?')
        fi = xf.get('fillId', '?')
        fo = xf.get('fontId', '?')
        bi = xf.get('borderId', '?')
        xf_data[sid] = {'numFmtId': nf, 'fillId': fi, 'fontId': fo, 'borderId': bi}
        print(f"{sid:>8} {nf:>9} {fi:>7} {fo:>7} {bi:>9}")
    else:
        print(f"{sid:>8}  *** OUT OF RANGE (max={len(all_xfs)-1}) ***")

# ── 3. Verify fill assignments ──
print("\n" + "=" * 70)
print("VERIFICATION")
print("=" * 70)

bic_ids = [847, 862, 850, 852, 864, 854, 855, 479]
t1_ids  = [655, 648, 683, 669, 709, 703, 675]
avg_ids = [654, 647, 661, 668, 708, 702, 674]
shared_ids = [858, 856, 843, 470, 471, 472, 473]

def check_group(name, ids, expected_fill):
    ok = []
    bad = []
    for sid in ids:
        if sid in xf_data:
            actual = xf_data[sid]['fillId']
            if str(actual) == str(expected_fill):
                ok.append(sid)
            else:
                bad.append((sid, actual))
        else:
            bad.append((sid, 'MISSING'))
    status = "PASS" if not bad else "FAIL"
    print(f"\n  {name} (expected fillId={expected_fill}): {status}")
    if ok:
        print(f"    OK: {ok}")
    if bad:
        for sid, actual in bad:
            print(f"    MISMATCH: style {sid} has fillId={actual}")

check_group("BIC styles", bic_ids, 16)
check_group("T1 styles", t1_ids, 18)
check_group("AVG styles", avg_ids, 17)

print(f"\n  Shared styles (858, 856, 843, 470-473) - actual fillIds:")
for sid in shared_ids:
    if sid in xf_data:
        print(f"    style {sid}: fillId={xf_data[sid]['fillId']}, numFmtId={xf_data[sid]['numFmtId']}, fontId={xf_data[sid]['fontId']}, borderId={xf_data[sid]['borderId']}")

# ── 4. Group by numFmtId within each category ──
print("\n" + "=" * 70)
print("GROUPED BY numFmtId (to see format patterns)")
print("=" * 70)
for name, ids in [("BIC", bic_ids), ("T1", t1_ids), ("AVG", avg_ids)]:
    print(f"\n  {name}:")
    for sid in ids:
        if sid in xf_data:
            d = xf_data[sid]
            print(f"    s={sid}: numFmt={d['numFmtId']}, font={d['fontId']}, border={d['borderId']}")
