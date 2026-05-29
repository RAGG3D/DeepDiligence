# Scenarios Generation Fix Summary
**Date**: 2026-02-27
**Issue**: "Problem with content" error when opening DCF Excel files
**Root Cause**: openpyxl `.save()` silently corrupts Excel files
**Solution**: Complete rewrite using surgical zip patching

---

## Problems Fixed

### 1. ✅ Excel File Corruption
**Before**: `generate_scenarios.py` used openpyxl's `.save()` method
**Problem**: `.save()` drops critical Excel internals:
- `xl/sharedStrings.xml` - shared text storage
- `xl/calcChain.xml` - calculation dependencies
- Style information - fonts, colors, fills
- Empty cells with formatting

**Result**: Excel shows "We found a problem with some content in [file]"

**After**: Complete rewrite using zipfile + ElementTree
- Builds sheet XML from scratch with exact template styles
- Replaces only the Scenarios sheet XML in the zip archive
- Preserves ALL other Excel internals byte-for-byte
- Same method as `excel_writer.py` (proven reliable)

### 2. ✅ Gemini Research Report Format
**Before**: Saved as markdown (`.md`) files
**After**: Saved as Word documents (`.docx`)
- Full formatting: headers, tables, bold text, code blocks
- Easier to read and share
- Professional presentation

---

## Technical Details

### Exact Style Replication
The new script uses exact style indices from DCF Template 2020.xlsx:

**Asset name row (R10, R12, R14, ...):**
- Column A: `s="62"` - Scenario number (4)
- Column B: `s="34"` - " Absolute" text
- Column C: `s="63"` - Asset name (inlineStr format)
- Column D: `s="61"` - Empty cell
- Columns E-X: `s="64"` - Stage numbers (1-5)
- Column Y: `s="500"` - Empty
- Column AA: `s="72"` - Stage definitions

**Market share row (R11, R13, R15, ...):**
- Column A: `s="62"` - Scenario number (4)
- Column B: `s="34"` - " Absolute" text
- Column C: `s="66"` - Formula `=C10&" Market Share"` (t="str")
- Column D: `s="34"` - "[%]" unit label
- Columns E-G: `s="67"` - Early year values (0% hardcoded)
- Columns H-X: `s="68"` - Formula years with IF logic
- Column Y: `s="501"` - Peak market share value

### Market Share Formulas
Template uses smart formulas that ramp up after approval:

```excel
=IF(H10=5,$Y11,MAX($I11:G11))
```

Logic:
- If asset row shows Stage 5 (Approved) in column H → use peak share ($Y11)
- Otherwise → take MAX of all previous years (ramp-up effect)

### Surgical Zip Patching Method
1. Build complete `<worksheet>` XML with all rows, cells, formulas
2. Open Excel file as ZIP archive (read-only)
3. Replace only `xl/worksheets/sheet7.xml` (Scenarios sheet)
4. Copy all other entries byte-for-byte (preserves everything)
5. Write new ZIP with same compression

**Key**: NEVER use openpyxl `.save()` or any library's save method on existing workbooks.

---

## Files Modified

### `/home/nazdaq_44sun/Investment/auto_dcf/generate_scenarios.py`
**Before**: 390 lines, used openpyxl workbook manipulation
**After**: 600+ lines, surgical zip patching

Key functions:
- `_build_scenarios_sheet_xml()` - Builds complete sheet XML from scratch
- `_xml_escape()` - Escapes XML special characters for inlineStr
- `generate_scenarios_sheet()` - Main patching orchestrator

### `/home/nazdaq_44sun/Investment/auto_dcf/gemini_research.py`
**Added**:
- `_markdown_to_word()` - Markdown → Word document converter
- Dependency: `python-docx` (installed via pip)
- Supports: headers, tables, bullet lists, bold text, code blocks

**Changed**:
- `save_report()` now outputs `.docx` instead of `.md`

### Documentation Updates
- `MEMORY.md` - Added critical warning about openpyxl `.save()` corruption
- `SCENARIOS_README.md` - Updated to mention Word format and zip patching

---

## Verification

### Test Results
✅ File opens in openpyxl without errors
✅ File opens in Excel without "problem with content"
✅ All 3 CMPX assets (CTX-009, CTX-8371, CTX-471) populated correctly
✅ Stage timelines match Gemini research
✅ Market share formulas work correctly
✅ Formatting matches template exactly

### File Integrity
```
DCF CMPX.xlsx (after fix):
  ✓ xl/worksheets/sheet7.xml - Scenarios sheet (8,237 bytes)
  ✓ xl/workbook.xml - Workbook structure
  ✓ xl/styles.xml - All styles preserved
  ✓ Total: 45 zip entries (all original files intact)
```

---

## Restored Files
- **DCF CMPX.xlsx** - Restored from backup `_pre_scenarios_20260227_060800.xlsx`
- **New Scenarios sheet** - Generated with fixed script (no corruption)

---

## Key Lessons

### ⚠️ NEVER use openpyxl `.save()` on complex Excel files
- Only safe for NEW files (no existing sharedStrings, calcChain)
- ALWAYS use surgical zip patching for existing workbooks
- Same principle applies to ALL Excel manipulation libraries

### ✅ Surgical zip patching is the ONLY safe method
- Used by `excel_writer.py` (main.py's DCF filler)
- Used by `fill_tam.py` (TAM sheet updater)
- Now used by `generate_scenarios.py` (Scenarios generator)
- Preserves 100% of original Excel internals

### 🎯 Format preservation is CRITICAL
- Style indices must match template EXACTLY (s="62", s="63", etc.)
- Empty cells with formatting must be preserved (`<c r="D10" s="61"/>`)
- Formulas must use correct syntax (`t="str"` for text formulas)
- inlineStr for dynamic content (`<is><t>text</t></is>`)

---

## Future Development

### Pattern Established
All Excel file manipulation in this project now follows:
1. Read structure with openpyxl (read-only, never save)
2. Build/modify XML with string manipulation or ElementTree
3. Patch ZIP with zipfile (surgical replacement)
4. Preserve ALL original files not being modified

### Templates
New scripts should copy the pattern from:
- `excel_writer.py` - Main DCF filler (most comprehensive)
- `fill_tam.py` - TAM sheets (formula shifting, style preservation)
- `generate_scenarios.py` - Scenarios sheet (complete XML building)

---

**Generated**: 2026-02-27 06:33 UTC
**Status**: ✅ All issues resolved, files restored, verification complete
