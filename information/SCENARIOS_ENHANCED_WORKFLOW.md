# Enhanced Scenarios Generation Workflow

**Date**: 2026-02-28
**Purpose**: Integrate ClinicalTrials.gov data + competitive analysis + social sentiment for accurate priced-in indication identification

---

## Overview

The enhanced workflow adds three critical data sources to Gemini Deep Research:

1. **ClinicalTrials.gov** — Active trials with specific indications, phases, NCT numbers
2. **Competitive Drug Data** — Marketed drugs' revenue, ORR, PFS, OS, safety profiles
3. **Social/Analyst Sentiment** — X/Twitter discussions, analyst reports mentioning specific indications

This enables Gemini to accurately identify which indications are priced into the stock, including **data transfer effects** (e.g., drug for A/B/C cancers with only A trial → market may price in B/C if pathology/pharmacology similar).

---

## Workflow Steps

### Step 1: Fetch Clinical Trials Data

```bash
python clinical_trials_fetcher.py \
    --ticker CMPX \
    --company-name "Compass Therapeutics" \
    --output-dir "/mnt/c/Users/yzsun/Desktop/DD/CMPX"
```

**What it does:**
- Extracts NCT numbers from latest 10-K and recent 8-K filings
- Searches ClinicalTrials.gov by sponsor name
- Fetches detailed trial data for each NCT (title, phase, status, conditions, interventions, dates)
- Saves JSON output: `CMPX_clinical_trials_TIMESTAMP.json`

**Example Output:**
```json
{
  "nct_ids": ["NCT05506943", "NCT06150664", "NCT03881488"],
  "trials": {
    "NCT05506943": {
      "title": "Phase 2/3 CTX-009 in Biliary Tract Cancers",
      "phase": "PHASE2",
      "conditions": ["Biliary Tract Cancer", "Cholangiocarcinoma"],
      "interventions": ["CTX-009", "Paclitaxel"],
      "status": "ACTIVE_NOT_RECRUITING"
    }
  }
}
```

### Step 2: Run Enhanced Gemini Deep Research

```bash
python gemini_research.py \
    --ticker CMPX \
    --company-name "Compass Therapeutics" \
    --output-dir "/mnt/c/Users/yzsun/Desktop/DD/CMPX"
```

**What it does:**
1. **Loads clinical trials data** from Step 1's JSON file
2. **Formats trial data** into readable prompt for Gemini
3. **Runs Gemini 2.5 Flash** with enhanced prompt that requires:
   - Priced-in indication analysis (which indications have trials vs which are priced in via data transfer)
   - Competitive drug data (revenue, ORR, PFS, OS, safety)
   - Horizontal comparisons (2L vs 2L, not 2L vs 1L)
   - X/Twitter and analyst sentiment analysis
   - **Explicit indication source citations** (NCT numbers, 10-K page numbers, analyst report names)
4. **Saves Word document**: `CMPX_gemini_research_TIMESTAMP.docx`

**Enhanced Prompt Sections:**

#### Priced-In Indications Analysis (NEW)
```
**Asset: CTX-009 (DLL3, Biliary Tract/Colorectal)**

**Priced-In Indications:**
- Biliary Tract Cancer, 2nd-line
  - Source: NCT05506943 (Phase 2/3, active), 10-K p.45
  - Market Evidence: 70% of analyst reports focus on BTC (Jefferies 2024, SVB 2025)
  - Social: X/Twitter ~60% BTC-focused

- Colorectal Cancer, 3rd-line
  - Source: NCT05513742 (Phase 2, completed), 10-K p.48
  - Market Evidence: 30% analyst mention (Jefferies 2024)
  - Social: X/Twitter ~40% CRC discussions

**Data Transfer Analysis:**
- Small Cell Lung Cancer potential (DLL3+ expression ~80% SCLC):
  - Pharmacology: DLL3 target highly expressed in SCLC
  - Pathology: Neuroendocrine features similar to BTC
  - Market Pricing: Partially priced in (~15% analyst mentions, no active trial but referenced in 10-K pipeline)
  - Conclusion: Include SCLC with reduced market share (10% peak vs BTC 25%)
```

#### Competitive Drug Data (Enhanced)
```
**Biliary Tract Cancer, 2nd-line competitors:**

| Drug | Revenue (2024) | ORR | PFS (mo) | OS (mo) | Grade 3+ AEs | Route |
|------|----------------|-----|----------|---------|--------------|-------|
| Pemigatinib | $120M | 35% | 6.9 | 17.5 | 64% | Oral |
| Ivosidenib | $95M | 23% | 2.7 | 10.8 | 46% | Oral |
| Gemcitabine/Cisplatin (generic) | N/A | 26% | 8.0 | 11.7 | 71% | IV |

**CTX-009 + Paclitaxel (Phase 2 data):**
- ORR: 28% (DLL3+ subset: 34%)
- PFS: 5.5 months
- OS: Data pending
- Grade 3+ AEs: 58%
- Route: IV

**Differentiation Assessment:** Above-average in DLL3+ subset (34% ORR vs competitors 23-35%), similar safety profile, IV route disadvantage vs oral competitors.
```

### Step 3: Generate Scenarios Sheet with Indication Labels

```bash
python generate_scenarios.py \
    --ticker CMPX \
    --research-file "/mnt/c/Users/yzsun/Desktop/DD/CMPX/CMPX_gemini_research_*.docx" \
    --company-name "Compass Therapeutics"
```

**What it does:**
- Parses enhanced Gemini report to extract:
  - Asset name with **all indications**: `CTX-009 (DLL3, BTC/CRC/SCLC)`
  - Market share rows with **specific indication labels**:
    - `CTX-009 (DLL3, BTC/CRC/SCLC) BTC Market Share`
    - `CTX-009 (DLL3, BTC/CRC/SCLC) CRC Market Share`
    - `CTX-009 (DLL3, BTC/CRC/SCLC) SCLC Market Share`
- Generates Scenarios sheet matching DCF Template 2020.xlsx format
- Uses surgical zip patching (NEVER openpyxl .save())

**Output Format:**
```
Row 10: 4 |  Absolute | CTX-009 (DLL3, BTC/CRC/SCLC) | | 1 | 2 | 2 | 3 | 4 | 5 | ...
Row 11: 4 |  Absolute | =C10&" BTC Market Share" | [%] | 0% | 0% | 5% | 15% | 25% | ...
Row 12: 4 |  Absolute | =C10&" CRC Market Share" | [%] | 0% | 2% | 8% | 12% | 15% | ...
Row 13: 4 |  Absolute | =C10&" SCLC Market Share" | [%] | 0% | 0% | 0% | 3% | 10% | ...

Row 14: 4 |  Absolute | CTX-8371 (PD-1xPD-L1, NSCLC/TNBC) | | | 1 | 2 | 3 | 4 | ...
Row 15: 4 |  Absolute | =C14&" Market Share" | [%] | 0% | 0% | 0% | 2% | 8% | ...
```

---

## Key Changes to Scripts

### clinical_trials_fetcher.py (NEW)

**Functions:**
- `extract_nct_numbers(text)` — Extract NCT\d{8} from 10-K/press releases
- `fetch_trial_details(nct_id)` — Call ClinicalTrials.gov API v2
- `search_trials_by_sponsor(company_name)` — Search by sponsor
- `get_company_trials(ticker, company_name)` — Aggregate all trial data

**API Used:** https://clinicaltrials.gov/api/v2/studies/{nct_id}

### gemini_research.py (Modified)

**New Functions:**
- `format_clinical_trials_data(trial_data)` — Format JSON → readable prompt text
- `run_deep_research()` — Now accepts `clinical_trials_data` parameter

**Enhanced Prompt Sections:**
- **Priced-In Indications Analysis** — Which indications have trials vs data transfer
- **Data Transfer Effect** — Pathology/pharmacology similarity analysis
- **Competitive Marketed Drugs** — Revenue, ORR, PFS, OS, safety, route
- **Horizontal Comparison** — Line-matched (2L vs 2L, 3L vs 3L)
- **Social Media & Analyst Sentiment** — X/Twitter, analyst reports
- **Multi-Indication Output Format** — Asset (Target, IND1/IND2/IND3) + specific MS rows

### generate_scenarios.py (Modified - TODO)

**Parsing Enhancements Needed:**
- Parse asset names: `Drug (Target, IND1/IND2/IND3)`
- Parse indication-specific market share rows: `Drug (...) IND1 Market Share`
- Generate multiple market share rows per asset
- Use inlineStr for all text cells to avoid shared string corruption

---

## Example: CMPX Complete Workflow

```bash
# Step 1: Fetch clinical trials
python clinical_trials_fetcher.py --ticker CMPX --company-name "Compass Therapeutics"
# Output: CMPX_clinical_trials_20260228_013033.json
# Found: 5 trials (CTX-009 BTC, CTX-009 CRC, CTX-8371 multi-tumor, CTX-471 multi-tumor)

# Step 2: Run enhanced Gemini research
python gemini_research.py --ticker CMPX --company-name "Compass Therapeutics"
# Output: CMPX_gemini_research_20260228_HHMMSS.docx
# Analysis includes:
#   - CTX-009: BTC (primary, 70% priced in), CRC (30% priced in), SCLC (15% data transfer)
#   - CTX-8371: NSCLC (40%), TNBC (30%), HL (20%), other (10%)
#   - CTX-471: Completed Phase 1, limited pricing

# Step 3: Generate Scenarios sheet
python generate_scenarios.py --ticker CMPX --research-file "C:\Users\yzsun\Desktop\DD\CMPX\CMPX_gemini_research_*.docx"
# Output: DCF CMPX.xlsx Scenarios sheet updated
# Format: CTX-009 (DLL3, BTC/CRC/SCLC)
#         - BTC Market Share row
#         - CRC Market Share row
#         - SCLC Market Share row
```

---

## Critical Requirements

### ✅ Indication Source Citations (Gemini MUST provide)
- NCT number: `NCT05506943`
- 10-K page: `10-K FY2024 p.45`
- Analyst report: `Jefferies equity research 2024-11-15`
- Social media: `X/Twitter biotech community discussions`

### ✅ Competitive Drug Data (Gemini MUST include)
- **Revenue**: Latest annual revenue from TAM sheets or earnings
- **ORR**: Objective Response Rate (%)
- **PFS**: Progression-Free Survival (months)
- **OS**: Overall Survival (months)
- **Safety**: Grade 3+ Adverse Events (%)
- **Route**: Oral vs IV vs SC

### ✅ Horizontal Comparison Requirement
- **Line-matched**: Compare 2L vs 2L, NOT 2L vs 1L
- **Biomarker-matched**: PD-L1+ vs PD-L1+ if relevant
- **Population-matched**: Similar prior treatment history

### ✅ Data Transfer Analysis
- **Pathology**: Disease biology similarity (e.g., neuroendocrine features)
- **Pharmacology**: Target expression levels across indications
- **Market Evidence**: Do analysts/X mention non-trial indications?
- **Conclusion**: Include or exclude from priced-in indications

---

## File Outputs

### Clinical Trials JSON
```
C:\Users\yzsun\Desktop\DD\CMPX\
  └── CMPX_clinical_trials_20260228_013033.json
      ├── nct_ids: [...]
      ├── trials: {NCT_ID: {title, phase, conditions, ...}}
      ├── indications_summary: {indication: [nct_ids]}
      └── phases_summary: {phase: [nct_ids]}
```

### Gemini Research Word Document
```
C:\Users\yzsun\Desktop\DD\CMPX\
  └── CMPX_gemini_research_20260228_HHMMSS.docx
      ├── Part 1: Market Cap Breakdown
      ├── Part 2: Pipeline Asset Market Share Projections
      │   ├── Priced-In Indications Analysis
      │   ├── Data Transfer Effect
      │   ├── Competitive Marketed Drugs
      │   └── Market Share Tables (by indication)
      └── Part 3: Stage Timeline Predictions
```

### Scenarios Sheet
```
DCF CMPX.xlsx → Scenarios sheet
  ├── Row 9: Scenario 4 header
  ├── Rows 10-15: Pipeline assets with indication labels
  │   ├── Asset name: Drug (Target, IND1/IND2/IND3)
  │   └── Market share rows: Drug (...) IND1 Market Share
  └── Column AA: Stage definitions
```

---

## Benefits of Enhanced Workflow

1. **Accurate Priced-In Analysis** — ClinicalTrials.gov data prevents Gemini from hallucinating trial status
2. **Data Transfer Identification** — Catches cases where market prices in indications without active trials
3. **Competitive Benchmarking** — Quantified differentiation vs marketed drugs
4. **Social Sentiment** — X/Twitter biotech community often ahead of analyst reports
5. **Indication Transparency** — DCF Scenarios sheet clearly shows which cancers drive valuation

---

**Generated**: 2026-02-28
**Status**: ✅ Clinical trials fetcher complete, Gemini prompt enhanced, generate_scenarios.py update pending
