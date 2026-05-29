# Scenarios Sheet Generation Workflow

This workflow uses Gemini Deep Research to analyze a biotech company's pipeline and automatically generates a Scenarios sheet in the DCF Excel file.

## Prerequisites

1. **Gemini API Key**: Obtain from [Google AI Studio](https://makersuite.google.com/app/apikey)
   - **Free tier**: Available but has very limited quotas (gemini-2.5-flash recommended)
   - **Paid tier**: Required for gemini-2.5-pro or high-volume usage
   - Set environment variable: `export GEMINI_API_KEY="your-key-here"`
   - Or add to `~/.bashrc`: `echo 'export GEMINI_API_KEY="your-key-here"' >> ~/.bashrc && source ~/.bashrc`

2. **Python packages**:
   ```bash
   pip install google-generativeai openpyxl
   ```

### ⚠️ Important: API Quota Limits

The free tier Gemini API has strict rate limits:
- **gemini-2.5-pro**: Quota often at 0 for free tier → requires paid billing
- **gemini-2.5-flash**: Higher free tier limits, suitable for testing
- **Current script**: Uses gemini-2.5-flash by default

If you hit quota errors, either:
1. Wait 24-48 hours for quota reset
2. Upgrade to paid billing in Google AI Studio
3. Use the demo workflow (manual research + generate_scenarios.py)

## Workflow Steps

### Step 1: Run Gemini Deep Research

```bash
python gemini_research.py --ticker CMPX --company-name "Compass Therapeutics"
```

**What it does:**
- Uses Gemini 2.5 Pro with Google Search Grounding
- Analyzes company's clinical pipeline (all priced-in assets)
- Generates 3-part report:
  1. Current stock price breakdown (cash + pipeline valuations)
  2. Market share projections for each asset (2024-2038)
  3. Stage timeline predictions (Phase I → II → III → BLA → Approval)

**Output:**
- **Word document**: `C:\Users\yzsun\Desktop\DD\CMPX\CMPX_gemini_research_TIMESTAMP.docx`
  - Formatted with headers, tables, bold text, code blocks
  - Easier to read and share than markdown
- Time: ~5-15 minutes depending on pipeline complexity

### Step 2: Review Research Report

Open the generated markdown file and verify:
- ✓ All pipeline assets identified
- ✓ Stage timelines have reasonable years
- ✓ Market share projections are plausible
- ✓ Sources cited for all major claims

**Edit if needed** — Gemini may miss some assets or use outdated data.

### Step 3: Generate Scenarios Sheet

```bash
python generate_scenarios.py \
    --ticker CMPX \
    --research-file "C:\Users\yzsun\Desktop\DD\CMPX\CMPX_gemini_research_20260227_*.md" \
    --company-name "Compass Therapeutics"
```

**What it does:**
- Parses Gemini research report (Word or markdown format)
- Extracts pipeline assets, stages, market shares
- Creates new "Scenarios" sheet in DCF CMPX.xlsx
- Matches exact format of DCF Template 2020.xlsx (Scenario 4 structure)
- ⚠️ **Uses surgical zip patching (NEVER openpyxl `.save()`)** to preserve all Excel internals
- Preserves ALL formatting: fonts, colors, fills, borders, formulas, empty cells
- No more "problem with content" errors!

**Output:**
- Overwrites `Scenarios` sheet in `C:\Users\yzsun\Desktop\DD\CMPX\DCF CMPX.xlsx`
- Creates backup: `DCF CMPX_pre_scenarios_TIMESTAMP.xlsx`

### Step 4: Verify in Excel

Open `DCF CMPX.xlsx` and check:
- Scenarios sheet is first tab
- Row 9: "Scenario 4 | Absolute Value: All Current Programs"
- Rows 10+: Each pipeline asset with:
  - Stage numbers (1-5) in appropriate year columns
  - Market share % rows below each asset
  - Formulas in column C referencing asset names
- Column AA: Stage definitions (Stage 1 = Phase I Start, etc.)

## Example: CMPX (Compass Therapeutics)

### Expected Pipeline Assets

Based on CMPX's public disclosures:
1. **CTX-009** (DLL3, SCLC) — Phase II
2. **CTX-8371** (PD-1/PD-L1, solid tumors) — Phase I
3. **CTX-471** (CD137, solid tumors) — Phase I

### Example Research Report Extract

```markdown
## Part 2: Pipeline Asset Market Share Projections

**Asset: CTX-009 (DLL3, SCLC)**

All Indications:
- TAM: ~15,000 patients (3rd-line+ SCLC in US)
- Competitive Landscape:
  - Current: Lurbinectedin ($~80M), topo-I (generic)
  - Pipeline: Dato-DXd (Daiichi Sankyo, exp. 2026)
- Differentiation:
  - ORR: 34% (SCLC-DLL3+) vs Lurbinectedin 30%
  - Assessment: Above-average
- Market Share Projection:
  2024: 0% (Phase II ongoing)
  2025: 0%
  2026: 5% (potential Phase III initiation)
  2027: 8%
  2028: 15% (launch year)
  2029: 22%
  2030: 25%
  ...
  2038: 18%
```

### Example Generated Scenarios Sheet

```
Row 9:  | Scenario 4 | Absolute Value: All Current Programs
Row 10: 4|  Absolute | CTX-009 (DLL3, SCLC) | | | | | 1 | 2 | 3 | 4 | 5 |...
Row 11: 4|  Absolute | =C10&" Market Share" | [%] | 0% | 0% | 0% | 5% | 8% | 15% |...
Row 12: 4|  Absolute | CTX-8371 (PD-1/PD-L1, Solid) | | 1 | 2 | 3 | 4 | 5 |...
Row 13: 4|  Absolute | =C12&" Market Share" | [%] | 0% | 0% | 0% | 2% | 5% |...
...
```

## Troubleshooting

### Gemini API Errors

**Error**: `GEMINI_API_KEY not set`
- **Fix**: Set environment variable or use `--api-key` flag

**Error**: `Model not found`
- **Fix**: Script auto-falls back to gemini-1.5-pro if gemini-2.0 unavailable

**Error**: `Empty response from Gemini`
- **Fix**: Try again (transient API issue) or reduce prompt complexity

### Parsing Errors

**Issue**: "No pipeline assets found in research report"
- **Cause**: Gemini report doesn't match expected markdown structure
- **Fix**: Manually edit report to follow Part 2/Part 3 format, or regenerate

**Issue**: Stage years look wrong
- **Cause**: Gemini made prediction errors
- **Fix**: Edit markdown file Stage years before running generate_scenarios.py

### Excel Formatting Issues

**Issue**: Formulas show as text
- **Cause**: OpenPyXL limitation
- **Fix**: Open Excel, press Ctrl+F3 (Name Manager), delete any broken names, recalculate (F9)

**Issue**: Colors/fonts don't match template
- **Fix**: Adjust style definitions in generate_scenarios.py (Font/PatternFill objects)

## Advanced Usage

### Custom Company Research

If company name can't be auto-detected from ticker:

```bash
python gemini_research.py \
    --ticker XYZ \
    --company-name "Example Biotech Inc." \
    --output-dir "/custom/path/to/output"
```

### Multi-Indication Assets

Gemini automatically splits market share if it detects drug targets multiple indications separately.

Example: BT5528 in BCYC template has:
- "BT5528 mUC Market Share"
- "BT5528 Other Cancer Market Share"

This happens when Gemini report has:
```
Indication 1: mUC
- Market Share Projection: 2024: 0%, 2025: 0%, ...

Indication 2: Other Cancers
- Market Share Projection: 2024: 0%, 2025: 0%, ...
```

### Re-running with Updated Research

To regenerate Scenarios sheet with new Gemini research:

```bash
# Step 1: Delete old research file (optional)
rm "C:\Users\yzsun\Desktop\DD\CMPX\CMPX_gemini_research_*.md"

# Step 2: Run new research
python gemini_research.py --ticker CMPX

# Step 3: Regenerate Scenarios (overwrites existing)
python generate_scenarios.py --ticker CMPX --research-file "C:\Users\yzsun\Desktop\DD\CMPX\CMPX_gemini_research_*.md"
```

Backup files are automatically created, so original Scenarios sheet is never lost.

## Files Created

```
C:\Users\yzsun\Desktop\DD\CMPX\
├── DCF CMPX.xlsx                              # Main DCF file (Scenarios sheet added/updated)
├── DCF CMPX_pre_scenarios_20260227_*.xlsx     # Backup before Scenarios generation
├── CMPX_gemini_research_20260227_*.md         # Gemini research report
└── (existing files: 10-K, events, etc.)
```

## Integration with Other Scripts

This workflow complements:
- **fill_tam.py**: Populates TAM sheets with drug revenue data (inputs to Gemini)
- **fill_events.py**: Historical Events sheet (catalyst analysis for Gemini context)
- **main.py**: DCF data fill from SEC filings

Typical full workflow:
1. `fill_tam.py` → Get latest drug revenue data
2. `fill_events.py` → Populate historical events
3. `gemini_research.py` → Analyze pipeline (uses TAM + events as context)
4. `generate_scenarios.py` → Create Scenarios sheet
5. `main.py` → Fill fundamental data from 10-K
6. Manual: Complete DCF valuation model

## Notes

- **Gemini Search Grounding**: Ensures up-to-date data (earnings, trial results, FDA actions)
- **Cost**: Gemini API charges per token. Typical research report costs ~$0.10-0.50
- **Accuracy**: Always verify Gemini's market share projections against analyst reports
- **Format Preservation**: generate_scenarios.py exactly replicates DCF Template formatting (fonts, colors, formulas)

## Support

For issues or questions, check:
1. This README
2. Script docstrings (`python gemini_research.py --help`)
3. MEMORY.md in `~/.claude/projects/.../memory/`
