#!/usr/bin/env python3
"""
research_tam_indications.py -- Research missing TAM indication drug data via Gemini.

For each missing indication (BTC, RCC, HCC, EC, ES-SCLC, Melanoma NCAM+, MPM, HL, MM),
queries Gemini for all marketed drugs with annual revenue data and global incidence.

Outputs structured JSON per indication for use by expand_tam.py.

Usage:
    python research_tam_indications.py [--indications BTC RCC HCC ...]
        [--output-dir path] [--dry-run]
"""

import argparse
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

try:
    from google import genai
except ImportError:
    print("ERROR: google-genai not installed. Run: pip install google-genai")
    sys.exit(1)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
#  INDICATION DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════

INDICATIONS = {
    # Solid tumors → TAM Solid
    "BTC": {
        "full_name": "Biliary Tract Cancer",
        "description": "cholangiocarcinoma, gallbladder cancer, ampullary cancer",
        "sheet": "TAM Solid",
    },
    "RCC": {
        "full_name": "Renal Cell Carcinoma",
        "description": "kidney cancer (clear cell and non-clear cell subtypes)",
        "sheet": "TAM Solid",
    },
    "HCC": {
        "full_name": "Hepatocellular Carcinoma",
        "description": "primary liver cancer",
        "sheet": "TAM Solid",
    },
    "EC": {
        "full_name": "Endometrial Cancer",
        "description": "uterine/endometrial cancer (pMMR and dMMR subtypes)",
        "sheet": "TAM Solid",
    },
    "ES-SCLC": {
        "full_name": "Extensive-Stage Small Cell Lung Cancer",
        "description": "extensive-stage small cell lung cancer (~15% of all lung cancers)",
        "sheet": "TAM Solid",
    },
    "Melanoma NCAM+": {
        "full_name": "NCAM-positive Melanoma",
        "description": "melanoma subset expressing NCAM/CD56 (~15% of melanoma)",
        "sheet": "TAM Solid",
    },
    "MPM": {
        "full_name": "Malignant Pleural Mesothelioma",
        "description": "rare asbestos-related cancer of the pleura",
        "sheet": "TAM Solid",
    },
    # Blood cancers → TAM Blood
    "HL": {
        "full_name": "Hodgkin Lymphoma",
        "description": "classical Hodgkin lymphoma",
        "sheet": "TAM Blood",
    },
    "MM": {
        "full_name": "Multiple Myeloma",
        "description": "plasma cell neoplasm",
        "sheet": "TAM Blood",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
#  RESEARCH PROMPT
# ══════════════════════════════════════════════════════════════════════════════

RESEARCH_PROMPT = """You are a senior pharmaceutical market analyst. I need comprehensive drug revenue data for **{indication_full} ({indication_abbrev})**.

## Task

Compile a complete list of ALL currently marketed drugs approved for {indication_full} ({indication_desc}) with their annual global revenue data.

## Requirements

1. **Drug List**: Include ALL approved drugs (checkpoint inhibitors, targeted therapies, chemotherapy, immunotherapy, ADCs, etc.)
2. **Revenue Data**: Annual GLOBAL revenue for EACH drug from 2015-2024, in MM USD (millions of US dollars)
   - Convert non-USD revenues at historical average exchange rates
   - For multi-indication drugs (e.g., Keytruda, Opdivo), estimate the {indication_abbrev}-specific revenue share based on:
     - Published indication-level revenue splits (if available)
     - Global incidence ratios between indications
     - Clinical trial enrollment as a proxy for prescribing volume
3. **Incidence**: Global new cases per year for {indication_full}
4. **Sources**: Note the primary source for each drug's revenue (SEC filing, annual report, GlobalData, etc.)

## Output Format

Return a JSON object with this exact structure:
```json
{{
    "indication": "{indication_abbrev}",
    "indication_full": "{indication_full}",
    "incidence_global_annual": 123000,
    "incidence_note": "Source and year of incidence data",
    "drugs": [
        {{
            "name": "Drug Brand Name",
            "generic": "generic-name",
            "manufacturer": "Company Name",
            "approval_year": 2020,
            "mechanism": "e.g., PD-1 inhibitor",
            "revenues_mm_usd": {{
                "2015": 0,
                "2016": 50,
                "2017": 120,
                "2018": 200,
                "2019": 350,
                "2020": 500,
                "2021": 650,
                "2022": 780,
                "2023": 850,
                "2024": 900
            }},
            "revenue_note": "Method used to estimate indication-specific revenue",
            "source": "SEC filing, annual report, etc."
        }}
    ]
}}
```

## Important Notes
- Revenue should be the TOTAL global revenue attributable to {indication_abbrev} specifically
- For drugs like Keytruda/Opdivo that treat many cancers, clearly state your estimation method
- Include drugs even if their {indication_abbrev} revenue is small relative to total
- Use 0 for years before approval or when data is unavailable
- Be thorough: missing a significant drug is worse than including a minor one
"""


# ══════════════════════════════════════════════════════════════════════════════
#  GEMINI RESEARCH
# ══════════════════════════════════════════════════════════════════════════════

def research_indication(
    indication: str,
    client: "genai.Client",
    model: str = "gemini-2.0-flash",
) -> Optional[Dict]:
    """Research a single indication via Gemini."""
    info = INDICATIONS[indication]
    prompt = RESEARCH_PROMPT.format(
        indication_abbrev=indication,
        indication_full=info["full_name"],
        indication_desc=info["description"],
    )

    log.info(f"  Querying Gemini for {indication} ({info['full_name']})...")

    try:
        response = client.models.generate_content(
            model=model,
            contents=prompt,
        )
        text = response.text
    except Exception as e:
        log.error(f"  Gemini error for {indication}: {e}")
        return None

    # Extract JSON from response
    json_m = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
    if json_m:
        json_str = json_m.group(1)
    else:
        # Try to find raw JSON
        json_m = re.search(r'\{.*\}', text, re.DOTALL)
        if json_m:
            json_str = json_m.group()
        else:
            log.error(f"  No JSON found in Gemini response for {indication}")
            log.info(f"  Response preview: {text[:500]}")
            return None

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        log.error(f"  JSON parse error for {indication}: {e}")
        log.info(f"  JSON preview: {json_str[:500]}")
        return None

    n_drugs = len(data.get("drugs", []))
    incidence = data.get("incidence_global_annual", 0)
    log.info(f"  {indication}: {n_drugs} drugs, incidence={incidence:,}")

    return data


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Research missing TAM indication drug data via Gemini"
    )
    parser.add_argument(
        "--indications", nargs="+",
        default=list(INDICATIONS.keys()),
        choices=list(INDICATIONS.keys()),
        help="Indications to research (default: all)")
    parser.add_argument(
        "--output-dir", default=".",
        help="Output directory for JSON files (default: current dir)")
    parser.add_argument(
        "--model", default="gemini-2.0-flash",
        help="Gemini model to use")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview prompts without calling Gemini")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key and not args.dry_run:
        log.error("Set GEMINI_API_KEY or GOOGLE_API_KEY environment variable")
        return

    log.info(f"Indications: {args.indications}")
    log.info(f"Output: {output_dir}")
    log.info(f"Model: {args.model}")
    log.info(f"Mode: {'DRY RUN' if args.dry_run else 'LIVE'}")

    if args.dry_run:
        for ind in args.indications:
            info = INDICATIONS[ind]
            prompt = RESEARCH_PROMPT.format(
                indication_abbrev=ind,
                indication_full=info["full_name"],
                indication_desc=info["description"],
            )
            print(f"\n{'='*60}")
            print(f"{ind} ({info['full_name']}) → {info['sheet']}")
            print(f"{'='*60}")
            print(prompt[:500] + "...")
        return

    client = genai.Client(api_key=api_key)

    results = {}
    for ind in args.indications:
        out_file = output_dir / f"tam_{ind.lower().replace(' ', '_').replace('+', '')}_research.json"

        # Skip if already researched
        if out_file.exists():
            log.info(f"  {ind}: already exists at {out_file}, skipping")
            with open(out_file) as f:
                results[ind] = json.load(f)
            continue

        data = research_indication(ind, client, args.model)
        if data:
            data["sheet"] = INDICATIONS[ind]["sheet"]
            with open(out_file, 'w') as f:
                json.dump(data, f, indent=2)
            log.info(f"  Saved → {out_file}")
            results[ind] = data
        else:
            log.warning(f"  {ind}: research failed, skipping")

        # Rate limit
        time.sleep(2)

    # Summary
    print(f"\n{'='*60}")
    print("TAM Indication Research Summary")
    print(f"{'='*60}")
    for ind in args.indications:
        if ind in results:
            d = results[ind]
            n_drugs = len(d.get("drugs", []))
            incidence = d.get("incidence_global_annual", 0)
            total_rev_2024 = sum(
                drug.get("revenues_mm_usd", {}).get("2024", 0)
                for drug in d.get("drugs", []))
            print(f"  {ind:15s} ({INDICATIONS[ind]['sheet']:10s}): "
                  f"{n_drugs:2d} drugs, incidence={incidence:>8,}, "
                  f"2024 rev=${total_rev_2024:>8,.0f}MM")
        else:
            print(f"  {ind:15s}: FAILED")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
