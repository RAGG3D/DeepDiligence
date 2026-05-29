#!/usr/bin/env python3
"""
clinical_trials_fetcher.py -- Fetch clinical trial data from ClinicalTrials.gov

Extracts NCT numbers from company 10-K/press releases and fetches detailed trial data
including indications, phases, and status.

Usage:
    python clinical_trials_fetcher.py --ticker CMPX
"""

import argparse
import json
import logging
import re
import requests
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  CLINICALTRIALS.GOV API V2
# ══════════════════════════════════════════════════════════════════════════════

CLINICALTRIALS_API_BASE = "https://clinicaltrials.gov/api/v2"


def fetch_trial_details(nct_id: str) -> Optional[Dict]:
    """
    Fetch detailed trial data from ClinicalTrials.gov API v2.

    Returns dict with:
        - nct_id: NCT number
        - title: Official trial title
        - status: Recruiting, Active, Completed, etc.
        - phase: Phase 1, Phase 2, Phase 3, etc.
        - conditions: List of diseases/conditions
        - interventions: List of drugs/interventions
        - sponsor: Lead sponsor name
        - start_date: Study start date
        - completion_date: Primary completion date
    """
    url = f"{CLINICALTRIALS_API_BASE}/studies/{nct_id}"

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()

        # Extract relevant fields from API response
        protocol_section = data.get("protocolSection", {})
        identification = protocol_section.get("identificationModule", {})
        status_module = protocol_section.get("statusModule", {})
        design_module = protocol_section.get("designModule", {})
        conditions_module = protocol_section.get("conditionsModule", {})
        interventions_module = protocol_section.get("armsInterventionsModule", {})
        sponsor_module = protocol_section.get("sponsorCollaboratorsModule", {})

        trial_data = {
            "nct_id": nct_id,
            "title": identification.get("officialTitle", identification.get("briefTitle", "")),
            "status": status_module.get("overallStatus", "Unknown"),
            "phase": design_module.get("phases", ["Unknown"])[0] if design_module.get("phases") else "Unknown",
            "conditions": conditions_module.get("conditions", []),
            "interventions": [
                interv.get("name", "")
                for interv in interventions_module.get("interventions", [])
            ],
            "sponsor": sponsor_module.get("leadSponsor", {}).get("name", "Unknown"),
            "start_date": status_module.get("startDateStruct", {}).get("date", ""),
            "completion_date": status_module.get("primaryCompletionDateStruct", {}).get("date", ""),
        }

        return trial_data

    except requests.RequestException as e:
        logger.error(f"Failed to fetch {nct_id}: {e}")
        return None


def search_trials_by_sponsor(company_name: str, max_results: int = 50) -> List[str]:
    """
    Search for trials by sponsor/collaborator name.

    Returns list of NCT IDs.
    """
    url = f"{CLINICALTRIALS_API_BASE}/studies"
    params = {
        "query.lead": company_name,
        "pageSize": max_results,
        "format": "json"
    }

    try:
        response = requests.get(url, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()

        studies = data.get("studies", [])
        nct_ids = [
            study.get("protocolSection", {}).get("identificationModule", {}).get("nctId")
            for study in studies
        ]

        return [nct for nct in nct_ids if nct]

    except requests.RequestException as e:
        logger.error(f"Failed to search trials for {company_name}: {e}")
        return []


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACT NCT NUMBERS FROM 10-K / PRESS RELEASES
# ══════════════════════════════════════════════════════════════════════════════

def extract_nct_numbers(text: str) -> List[str]:
    """
    Extract NCT numbers (e.g., NCT12345678) from text.

    NCT format: NCT followed by 8 digits.
    """
    pattern = r'\b(NCT\d{8})\b'
    matches = re.findall(pattern, text, re.IGNORECASE)

    # Uppercase and deduplicate
    nct_ids = list(set([nct.upper() for nct in matches]))

    return sorted(nct_ids)


def fetch_latest_10k_text(ticker: str) -> Optional[str]:
    """
    Fetch latest 10-K text from SEC EDGAR.

    Returns full text content for NCT extraction.
    """
    from sec_fetcher import SECFetcher

    try:
        fetcher = SECFetcher()
        cik = fetcher.get_cik(ticker)

        # Get submissions to find latest 10-K
        url = f"https://data.sec.gov/submissions/CIK{cik.zfill(10)}.json"
        headers = {"User-Agent": "Research research@example.com"}

        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()

        # Find latest 10-K
        recent_filings = data.get("filings", {}).get("recent", {})
        forms = recent_filings.get("form", [])
        accession_numbers = recent_filings.get("accessionNumber", [])

        for form, accession in zip(forms, accession_numbers):
            if form == "10-K":
                # Construct document URL
                accession_no_dash = accession.replace("-", "")
                doc_url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession_no_dash}/{accession}.txt"

                doc_response = requests.get(doc_url, headers=headers, timeout=30)
                doc_response.raise_for_status()

                return doc_response.text

        logger.warning(f"No 10-K found for {ticker}")
        return None

    except Exception as e:
        logger.error(f"Failed to fetch 10-K for {ticker}: {e}")
        return None


def fetch_press_releases(ticker: str) -> List[str]:
    """
    Fetch recent press releases from SEC EDGAR 8-K filings.

    Returns list of text contents.
    """
    from sec_fetcher import SECFetcher

    fetcher = SECFetcher()
    cik = fetcher.get_cik(ticker)

    # Query SEC submissions API for 8-K filings
    url = f"https://data.sec.gov/submissions/CIK{cik.zfill(10)}.json"
    headers = {"User-Agent": "YourCompany research@yourcompany.com"}

    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()

        # Get recent 8-K filings (last 2 years)
        recent_filings = data.get("filings", {}).get("recent", {})
        forms = recent_filings.get("form", [])
        accession_numbers = recent_filings.get("accessionNumber", [])
        filing_dates = recent_filings.get("filingDate", [])

        press_release_texts = []

        for form, accession, filing_date in zip(forms, accession_numbers, filing_dates):
            if form != "8-K":
                continue

            # Only process recent filings (last 2 years)
            if filing_date < "2024-01-01":
                continue

            # Construct 8-K document URL
            accession_no_dash = accession.replace("-", "")
            doc_url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession_no_dash}/{accession}.txt"

            try:
                doc_response = requests.get(doc_url, headers=headers, timeout=10)
                doc_response.raise_for_status()
                press_release_texts.append(doc_response.text)

                # Limit to 20 press releases to avoid excessive processing
                if len(press_release_texts) >= 20:
                    break

                time.sleep(0.1)  # Rate limiting

            except Exception as e:
                logger.debug(f"Failed to fetch 8-K {accession}: {e}")
                continue

        return press_release_texts

    except Exception as e:
        logger.error(f"Failed to fetch press releases for {ticker}: {e}")
        return []


# ══════════════════════════════════════════════════════════════════════════════
#  AGGREGATE TRIAL DATA
# ══════════════════════════════════════════════════════════════════════════════

def get_company_trials(ticker: str, company_name: Optional[str] = None) -> Dict:
    """
    Aggregate all clinical trial data for a company.

    Returns dict:
        {
            "nct_ids": [list of NCT IDs found],
            "trials": {nct_id: trial_details_dict},
            "indications_summary": {indication: [list of nct_ids]},
            "phases_summary": {phase: [list of nct_ids]}
        }
    """
    logger.info(f"Fetching clinical trial data for {ticker}...")

    # Step 1: Extract NCT numbers from 10-K
    logger.info("Step 1: Extracting NCT numbers from latest 10-K...")
    nct_from_10k = []

    text_10k = fetch_latest_10k_text(ticker)
    if text_10k:
        nct_from_10k = extract_nct_numbers(text_10k)
        logger.info(f"  Found {len(nct_from_10k)} NCT IDs in 10-K: {nct_from_10k}")

    # Step 2: Extract NCT numbers from press releases
    logger.info("Step 2: Extracting NCT numbers from recent 8-K filings...")
    nct_from_press = []

    press_texts = fetch_press_releases(ticker)
    for text in press_texts:
        nct_from_press.extend(extract_nct_numbers(text))

    nct_from_press = list(set(nct_from_press))
    logger.info(f"  Found {len(nct_from_press)} NCT IDs in press releases: {nct_from_press}")

    # Step 3: Search by company name (if provided)
    nct_from_search = []
    if company_name:
        logger.info(f"Step 3: Searching ClinicalTrials.gov for '{company_name}'...")
        nct_from_search = search_trials_by_sponsor(company_name)
        logger.info(f"  Found {len(nct_from_search)} NCT IDs via sponsor search")

    # Combine and deduplicate
    all_nct_ids = list(set(nct_from_10k + nct_from_press + nct_from_search))
    logger.info(f"\nTotal unique NCT IDs: {len(all_nct_ids)}")

    # Step 4: Fetch detailed data for each trial
    logger.info("Step 4: Fetching detailed trial data from ClinicalTrials.gov...")
    trials = {}

    for nct_id in all_nct_ids:
        logger.info(f"  Fetching {nct_id}...")
        trial_data = fetch_trial_details(nct_id)

        if trial_data:
            trials[nct_id] = trial_data

        time.sleep(0.3)  # Rate limiting (ClinicalTrials.gov allows ~3 requests/sec)

    # Step 5: Summarize by indication and phase
    indications_summary = {}
    phases_summary = {}

    for nct_id, trial in trials.items():
        # Group by indication
        for condition in trial["conditions"]:
            if condition not in indications_summary:
                indications_summary[condition] = []
            indications_summary[condition].append(nct_id)

        # Group by phase
        phase = trial["phase"]
        if phase not in phases_summary:
            phases_summary[phase] = []
        phases_summary[phase].append(nct_id)

    return {
        "nct_ids": all_nct_ids,
        "trials": trials,
        "indications_summary": indications_summary,
        "phases_summary": phases_summary,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  SAVE RESULTS
# ══════════════════════════════════════════════════════════════════════════════

def save_trial_data(ticker: str, trial_data: Dict, output_dir: Path):
    """Save trial data as JSON."""
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{ticker}_clinical_trials_{timestamp}.json"
    output_path = output_dir / filename

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(trial_data, f, indent=2, ensure_ascii=False)

    logger.info(f"\nTrial data saved: {output_path}")
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Fetch clinical trial data from ClinicalTrials.gov"
    )
    parser.add_argument("--ticker", required=True, help="Stock ticker (e.g., CMPX)")
    parser.add_argument("--company-name", help="Company name for sponsor search")
    parser.add_argument("--output-dir", help="Output directory (default: C:\\Users\\yzsun\\Desktop\\DD\\{TICKER})")

    args = parser.parse_args()

    # Set output directory
    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        output_dir = Path(f"C:/Users/yzsun/Desktop/DD/{args.ticker}")

    # Fetch trial data
    trial_data = get_company_trials(args.ticker, args.company_name)

    # Save results
    output_path = save_trial_data(args.ticker, trial_data, output_dir)

    # Print summary
    print(f"\n{'='*80}")
    print(f"Clinical Trials Summary: {args.ticker}")
    print(f"{'='*80}")
    print(f"Total NCT IDs: {len(trial_data['nct_ids'])}")
    print(f"\nBy Phase:")
    for phase, ncts in sorted(trial_data['phases_summary'].items()):
        print(f"  {phase}: {len(ncts)} trials")
    print(f"\nTop Indications:")
    sorted_indications = sorted(
        trial_data['indications_summary'].items(),
        key=lambda x: len(x[1]),
        reverse=True
    )
    for indication, ncts in sorted_indications[:10]:
        print(f"  {indication}: {len(ncts)} trials")
    print(f"\nOutput: {output_path}")
    print(f"{'='*80}\n")


if __name__ == "__main__":
    main()
