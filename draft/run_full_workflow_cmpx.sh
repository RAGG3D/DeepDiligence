#!/bin/bash
#
# Complete CMPX Research Workflow
# Run this script after setting GEMINI_API_KEY
#

set -e  # Exit on error

TICKER="CMPX"
COMPANY_NAME="Compass Therapeutics"
BASE_DIR="/mnt/c/Users/yzsun/Desktop/DD/${TICKER}"

echo "========================================================================"
echo "CMPX Complete Research Workflow"
echo "========================================================================"
echo ""

# Check API key
if [ -z "$GEMINI_API_KEY" ]; then
    echo "❌ ERROR: GEMINI_API_KEY not set"
    echo ""
    echo "Please set it first:"
    echo "  export GEMINI_API_KEY='your-api-key-here'"
    echo ""
    echo "Get API key from: https://aistudio.google.com/app/apikey"
    exit 1
fi

echo "✅ GEMINI_API_KEY is set"
echo ""

# Step 1: Clinical Trials (already done, but verify file exists)
echo "Step 1: Clinical Trials Data"
echo "----------------------------"
TRIALS_FILE=$(ls -t ${BASE_DIR}/CMPX_clinical_trials_*.json 2>/dev/null | head -1)
if [ -n "$TRIALS_FILE" ]; then
    echo "✅ Using existing: $TRIALS_FILE"
else
    echo "Fetching fresh clinical trials data..."
    python clinical_trials_fetcher.py --ticker ${TICKER} --company-name "${COMPANY_NAME}"
    TRIALS_FILE=$(ls -t ${BASE_DIR}/CMPX_clinical_trials_*.json | head -1)
fi
echo ""

# Step 2: Gemini Research
echo "Step 2: Gemini Deep Research"
echo "----------------------------"
echo "Calling Gemini API (this may take 5-10 minutes)..."
python gemini_research.py --ticker ${TICKER} --company-name "${COMPANY_NAME}"

# Find the generated markdown file
MD_FILE=$(ls -t ${BASE_DIR}/CMPX_gemini_research_*.md | head -1)
echo "✅ Research report generated: $MD_FILE"
echo ""

# Step 3: Convert to Word
echo "Step 3: Convert to Word Format"
echo "----------------------------"
python3 << EOF
from gemini_research import _markdown_to_word
from pathlib import Path

md_path = Path("$MD_FILE")
with open(md_path, 'r', encoding='utf-8') as f:
    content = f.read()

output_path = md_path.with_suffix('.docx')
doc = _markdown_to_word(content, "${TICKER}")
doc.save(str(output_path))
print(f"✅ Word document saved: {output_path}")
EOF
echo ""

# Step 4: Generate Scenarios Sheet
echo "Step 4: Generate Scenarios Sheet"
echo "----------------------------"
python generate_scenarios.py --ticker ${TICKER} --research-file "$MD_FILE"
echo ""

# Step 5: Verification
echo "========================================================================"
echo "Workflow Complete - Verification"
echo "========================================================================"
echo ""

# Extract and display generated rows
python3 << 'VERIFY'
import zipfile
import re

zip_path = "/mnt/c/Users/yzsun/Desktop/DD/CMPX/DCF CMPX.xlsx"
with zipfile.ZipFile(zip_path) as zf:
    sheet_xml = zf.read("xl/worksheets/sheet7.xml").decode("utf-8")

print("Generated Scenarios Sheet Rows:")
print("-" * 70)

row_count = 0
for row_num in range(10, 30):  # Check up to row 30
    # Check for asset name
    inline_match = re.search(rf'<c r="C{row_num}"[^>]*><is><t>([^<]+)</t>', sheet_xml)
    if inline_match:
        text = inline_match.group(1)
        print(f"Row {row_num}: {text}")
        row_count += 1
        continue

    # Check for formula (MS row)
    formula_match = re.search(rf'<c r="C{row_num}"[^>]*><f>([^<]+)</f>', sheet_xml)
    if formula_match:
        formula = formula_match.group(1).replace("&amp;", "&")
        # Extract indication name from formula
        if "BTC" in formula:
            print(f"Row {row_num}:   → BTC Market Share")
        elif "CRC" in formula:
            print(f"Row {row_num}:   → CRC Market Share")
        elif "SCLC" in formula:
            print(f"Row {row_num}:   → SCLC Market Share")
        elif "NSCLC" in formula:
            print(f"Row {row_num}:   → NSCLC Market Share")
        elif "TNBC" in formula:
            print(f"Row {row_num}:   → TNBC Market Share")
        elif "HL" in formula or "Hodgkin" in formula:
            print(f"Row {row_num}:   → Hodgkin Lymphoma Market Share")
        elif "HNSCC" in formula:
            print(f"Row {row_num}:   → HNSCC Market Share")
        elif "Melanoma" in formula:
            print(f"Row {row_num}:   → Melanoma Market Share")
        elif "Market Share" in formula:
            print(f"Row {row_num}:   → Market Share (combined)")
        row_count += 1
        continue

    # If we've seen some rows and now hit 3 empty ones, stop
    if row_count > 0:
        if row_num > 10:
            break

print("")
print(f"Total rows generated: {row_count}")
print("")

VERIFY

echo "========================================================================"
echo "Files Generated:"
echo "========================================================================"
ls -lh ${BASE_DIR}/CMPX_gemini_research_*.md 2>/dev/null | tail -1
ls -lh ${BASE_DIR}/CMPX_gemini_research_*.docx 2>/dev/null | tail -1
ls -lh ${BASE_DIR}/"DCF CMPX.xlsx" 2>/dev/null
echo ""
echo "✅ Complete workflow finished successfully!"
echo "========================================================================"
