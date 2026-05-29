"""
sec_fetcher_enhanced.py – Enhanced SEC EDGAR XBRL data retrieval with Notes details.

Key enhancements:
1. Dynamic Notes extraction (R&D, G&A, PP&E, Accrued Liabilities)
2. Auto-balancing with "Other Adjustments" rows
3. Comprehensive XBRL concept discovery
"""

import logging
import re
import time
from collections import defaultdict
from typing import Dict, List, Optional, Tuple, Set

import requests

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# Core XBRL concept mapping (Top-level statements)
# ─────────────────────────────────────────────────────────────────────────────

XBRL_MAP: List[Tuple] = [
    # ── Income Statement ──────────────────────────────────────────────────────
    ("IS", None, "Research And Development",
     ["ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost",
      "ResearchAndDevelopmentExpense"], 1 / 1000, True),

    ("IS", None, "General And Administrative",
     ["GeneralAndAdministrativeExpense"], 1 / 1000, True),

    ("IS", None, "Interest Income",
     ["InvestmentIncomeInterest",
      "OtherNonoperatingIncome",
      "InterestAndOtherIncome",
      "InterestIncomeOther"], 1 / 1000, True),

    ("IS", None, "Interest Expense",
     ["InterestExpense",
      "InterestExpenseDebt",
      "InterestExpenseNonoperating",
      "InterestAndDebtExpense"], 1 / 1000, True),

    ("IS", None, "Income Tax Provision",
     ["IncomeTaxExpenseBenefit",
      "CurrentIncomeTaxExpenseBenefit"], 1 / 1000, True),

    ("IS", None, "Weighted-Average Number Of Common Shares, Basic And Diluted",
     ["WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
      "WeightedAverageNumberOfSharesOutstandingBasic",
      "WeightedAverageNumberOfDilutedSharesOutstanding"], 1, True),

    ("IS", None, "Foreign Currency Translation Adjustment",
     ["OtherComprehensiveIncomeLossForeignCurrencyTransactionAndTranslationAdjustmentNetOfTax",
      "OtherComprehensiveIncomeForeignCurrencyTransactionAndTranslationAdjustmentNetOfTaxPortionAttributableToParent",
      "OtherComprehensiveIncomeForeignCurrencyTransactionAndTranslationGainLossArisingDuringPeriodNetOfTax",
      "OtherComprehensiveIncomeUnrealizedHoldingGainLossOnSecuritiesArisingDuringPeriodNetOfTax",
      "OtherComprehensiveIncomeLossAvailableForSaleSecuritiesAdjustmentNetOfTax"],
     1 / 1000, True),

    # ── Balance Sheet ─────────────────────────────────────────────────────────
    ("BS", None, "Cash And Cash Equivalents",
     ["CashAndCashEquivalentsAtCarryingValue",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
     1 / 1000, False),

    ("BS", None, "Prepaid Expenses And Other Current Assets",
     ["PrepaidExpenseAndOtherAssetsCurrent",
      "OtherAssetsCurrent",
      "PrepaidExpenseCurrent"], 1 / 1000, False),

    ("BS", None, "Research And Development Incentives Receivable",
     ["ValueAddedTaxReceivableCurrent",
      "IncomeTaxReceivable",
      "ResearchAndDevelopmentIncentiveReceivable",
      "AccruedIncomeTaxesCurrent"], 1 / 1000, False),

    ("BS", None, "Property And Equipment, Net",
     ["PropertyPlantAndEquipmentNet"], 1 / 1000, False),

    ("BS", None, "Operating Lease Right-Of-Use Assets",
     ["OperatingLeaseRightOfUseAsset"], 1 / 1000, False),

    ("BS", None, "Other Assets",
     ["OtherAssetsNoncurrent",
      "OtherAssetsMiscellaneousNoncurrent",
      "DeferredIncomeTaxAssetsNet"], 1 / 1000, False),

    ("BS", None, "Accounts Payable",
     ["AccountsPayableCurrent"], 1 / 1000, False),

    ("BS", None, "Accrued Expenses And Other Current Liabilities",
     ["AccountsPayableAndOtherAccruedLiabilitiesCurrent",
      "AccruedLiabilitiesCurrent",
      "EmployeeRelatedLiabilitiesCurrent"], 1 / 1000, False),

    ("BS", None, "Debt, Current Portion",
     ["DebtCurrent",
      "LongTermDebtCurrent",
      "ConvertibleNotesPayableCurrent",
      "NotesPayableRelatedPartiesClassifiedCurrent"], 1 / 1000, False),

    ("BS", None, "Long-Term Debt, Net Of Discount",
     ["LongTermDebtNoncurrent",
      "LongTermDebt"], 1 / 1000, False),

    ("BS", None, "Operating Lease Liabilities, Net Of Current Portion",
     ["OperatingLeaseLiabilityNoncurrent"], 1 / 1000, False),

    ("BS", None, "Ordinary Shares, £0.01 Nominal Value",
     ["CommonStockValue"], 1 / 1000, False),

    ("BS", None, "Additional Paid-In Capital",
     ["AdditionalPaidInCapital",
      "AdditionalPaidInCapitalCommonStock"], 1 / 1000, False),

    ("BS", None, "Accumulated Other Comprehensive (Loss) Income",
     ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"], 1 / 1000, False),

    ("BS", None, "Accumulated Deficit",
     ["RetainedEarningsAccumulatedDeficit"], 1 / 1000, False),

    # ── Cash Flow Statement ───────────────────────────────────────────────────
    ("CFS", None, "Net Cash Used In Operating Activities",
     ["NetCashProvidedByUsedInOperatingActivities"], 1 / 1000, True),

    ("CFS", None, "Net Cash Provided by (Used In) Investing Activities",
     ["NetCashProvidedByUsedInInvestingActivities"], 1 / 1000, True),

    ("CFS", None, "Net Cash Provided By Financing Activities",
     ["NetCashProvidedByUsedInFinancingActivities"], 1 / 1000, True),

    ("CFS", None, "Effect Of Exchange Rate Changes On Cash And Cash Equivalents",
     ["EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
      "EffectOfExchangeRateOnCashAndCashEquivalentsContinuingOperations"], 1 / 1000, True),
]

# ─────────────────────────────────────────────────────────────────────────────
# Notes-level extraction patterns
# ─────────────────────────────────────────────────────────────────────────────

NOTES_PATTERNS = {
    "R&D": {
        "parent_key": ("IS", None, "Research And Development"),
        "keywords": [
            "ResearchAndDevelopmentExpense",
            "ClinicalTrialExpense",
            "ResearchAndDevelopmentArrangement",
        ],
        "exclude": ["ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"],  # already top-level
    },
    "G&A": {
        "parent_key": ("IS", None, "General And Administrative"),
        "keywords": [
            "GeneralAndAdministrativeExpense",
            "ProfessionalFees",
            "LegalFees",
            "AdministrativeExpense",
        ],
        "exclude": ["GeneralAndAdministrativeExpense"],  # already top-level
    },
    "PP&E": {
        "parent_key": ("BS", None, "Property And Equipment, Net"),
        "keywords": [
            "PropertyPlantAndEquipment",
            "LeaseholdImprovements",
            "ComputerEquipment",
            "Furniture",
            "OfficeEquipment",
            "AccumulatedDepreciation",
        ],
        "exclude": ["PropertyPlantAndEquipmentNet"],  # already top-level
    },
    "Accrued": {
        "parent_key": ("BS", None, "Accrued Expenses And Other Current Liabilities"),
        "keywords": [
            "AccruedLiabilities",
            "EmployeeRelatedLiabilities",
            "AccruedProfessionalFees",
            "AccruedResearchAndDevelopmentCosts",
        ],
        "exclude": [
            "AccruedLiabilitiesCurrent",
            "AccountsPayableAndOtherAccruedLiabilitiesCurrent"
        ],
    },
}


# ─────────────────────────────────────────────────────────────────────────────
class SECFetcher:
    BASE_URL = "https://data.sec.gov/api/xbrl/companyfacts"
    SEARCH_URL = "https://efts.sec.gov/LATEST/search-index"
    COMPANY_SEARCH_URL = "https://www.sec.gov/cgi-bin/browse-edgar"

    def __init__(self, user_agent: str = "financial-research contact@example.com"):
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": user_agent})
        self._facts_cache: Dict[str, dict] = {}

    # ── CIK resolution ────────────────────────────────────────────────────────

    def get_cik(self, ticker: str) -> str:
        """Resolve ticker → zero-padded 10-digit CIK string."""
        r2 = self.session.get(
            self.COMPANY_SEARCH_URL,
            params={"CIK": ticker, "type": "10-K", "action": "getcompany"},
            timeout=20,
        )
        ciks = re.findall(r"CIK=(\d+)", r2.text)
        if ciks:
            cik = str(int(ciks[0])).zfill(10)
            logger.info(f"Resolved {ticker} → CIK {cik}")
            return cik

        tickers_url = "https://www.sec.gov/files/company_tickers.json"
        r3 = self.session.get(tickers_url, timeout=20)
        for entry in r3.json().values():
            if entry.get("ticker", "").upper() == ticker.upper():
                cik = str(entry["cik_str"]).zfill(10)
                logger.info(f"Resolved {ticker} → CIK {cik} (via tickers JSON)")
                return cik

        raise ValueError(f"Cannot resolve CIK for ticker '{ticker}'")

    # ── Company facts ─────────────────────────────────────────────────────────

    def get_company_facts(self, cik: str) -> dict:
        """Fetch full XBRL company facts JSON from SEC."""
        if cik in self._facts_cache:
            return self._facts_cache[cik]

        url = f"{self.BASE_URL}/CIK{cik}.json"
        logger.info(f"Fetching company facts: {url}")
        time.sleep(0.3)
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        self._facts_cache[cik] = data
        return data

    # ── Per-concept value extraction ──────────────────────────────────────────

    @staticmethod
    def _best_value(
        entries: list,
        year: int,
        is_flow: bool,
        fiscal_year_end_month: int = 12,
    ) -> Optional[float]:
        """Extract best value for given fiscal year from XBRL entries."""
        fye_month_str = f"{fiscal_year_end_month:02d}"

        if is_flow:
            candidates = [
                e for e in entries
                if e.get("form") in ("10-K", "10-K/A")
                and e.get("fp") == "FY"
                and e.get("end", "")[:4] == str(year)
            ]
        else:
            candidates = [
                e for e in entries
                if e.get("form") in ("10-K", "10-K/A")
                and e.get("end", "")[:4] == str(year)
                and e.get("end", "")[5:7] == fye_month_str
            ]

        if not candidates:
            return None

        candidates.sort(key=lambda x: x.get("filed", ""), reverse=True)
        return candidates[0]["val"]

    def extract_concept_by_year(
        self,
        facts: dict,
        concepts: List[str],
        years: List[int],
        is_flow: bool,
        fiscal_year_end_month: int = 12,
    ) -> Dict[int, Optional[float]]:
        """Try each concept in order; return first non-None result per year."""
        us_gaap = facts.get("facts", {}).get("us-gaap", {})
        result: Dict[int, Optional[float]] = {y: None for y in years}

        for concept in concepts:
            if concept not in us_gaap:
                continue
            units = us_gaap[concept].get("units", {})
            entries = units.get("USD") or units.get("shares") or []

            for year in years:
                if result[year] is not None:
                    continue
                val = self._best_value(entries, year, is_flow, fiscal_year_end_month)
                if val is not None:
                    result[year] = val
                    logger.debug(f"  {concept}[{year}] = {val:,.0f}")

        return result

    # ── Zero propagation ──────────────────────────────────────────────────────

    @staticmethod
    def _propagate_zero(
        year_vals: Dict[int, Optional[float]],
        years: List[int],
    ) -> Dict[int, Optional[float]]:
        """Carry forward explicit 0 values to subsequent None years."""
        result = dict(year_vals)
        for year in sorted(years):
            if result.get(year) == 0:
                for next_year in sorted(years):
                    if next_year <= year:
                        continue
                    if result.get(next_year) is None:
                        result[next_year] = 0
                    else:
                        break
        return result

    # ── Fiscal year-end month detection ───────────────────────────────────────

    @staticmethod
    def detect_fye_month(facts: dict) -> int:
        """Detect fiscal year-end month from 10-K end dates."""
        us_gaap = facts.get("facts", {}).get("us-gaap", {})
        from collections import Counter
        months: Counter = Counter()
        for concept_data in us_gaap.values():
            for unit_entries in concept_data.get("units", {}).values():
                for e in unit_entries:
                    if e.get("form") in ("10-K", "10-K/A") and e.get("fp") == "FY":
                        end = e.get("end", "")
                        if len(end) >= 7:
                            months[int(end[5:7])] += 1
        if months:
            return months.most_common(1)[0][0]
        return 12

    # ── NEW: Dynamic Notes extraction ─────────────────────────────────────────

    def discover_note_concepts(
        self,
        facts: dict,
        keywords: List[str],
        exclude: List[str],
        is_flow: bool,
    ) -> List[str]:
        """
        Discover all XBRL concepts matching keywords but not in exclude list.
        Returns list of concept names that exist in the company's filings.
        """
        us_gaap = facts.get("facts", {}).get("us-gaap", {})
        discovered: Set[str] = set()

        for concept_name in us_gaap.keys():
            # Skip if in exclude list
            if concept_name in exclude:
                continue

            # Check if concept matches any keyword
            for keyword in keywords:
                if keyword.lower() in concept_name.lower():
                    # Verify it has data for our target form
                    units = us_gaap[concept_name].get("units", {})
                    entries = units.get("USD") or units.get("shares") or []

                    # Check if it has 10-K data
                    has_10k = any(
                        e.get("form") in ("10-K", "10-K/A")
                        for e in entries
                    )
                    if has_10k:
                        discovered.add(concept_name)
                        break

        return sorted(discovered)

    def extract_notes_details(
        self,
        facts: dict,
        pattern_config: dict,
        years: List[int],
        fiscal_year_end_month: int,
    ) -> List[Tuple[str, Dict[int, Optional[float]]]]:
        """
        Extract detailed Notes-level items for a category.

        Returns:
            List of (concept_display_name, {year: value}) tuples
        """
        keywords = pattern_config["keywords"]
        exclude = pattern_config["exclude"]

        # Determine is_flow from parent key
        parent_key = pattern_config["parent_key"]
        is_flow = parent_key[0] in ("IS", "CFS")

        # Discover concepts
        concepts = self.discover_note_concepts(facts, keywords, exclude, is_flow)

        if not concepts:
            logger.debug(f"  No note concepts discovered for pattern {pattern_config}")
            return []

        logger.info(f"  Discovered {len(concepts)} note concepts")

        results: List[Tuple[str, Dict[int, Optional[float]]]] = []
        for concept in concepts:
            year_vals = self.extract_concept_by_year(
                facts, [concept], years, is_flow, fiscal_year_end_month
            )

            # Only include if has at least one non-None value
            if any(v is not None for v in year_vals.values()):
                # Convert concept name to display name
                display_name = self._concept_to_display_name(concept)
                results.append((display_name, year_vals))
                logger.debug(f"    {display_name}: {year_vals}")

        return results

    @staticmethod
    def _concept_to_display_name(concept: str) -> str:
        """
        Convert XBRL concept name to human-readable display name.

        Example:
          ResearchAndDevelopmentExpenseClinicalTrial
          → Research And Development Expense Clinical Trial
        """
        # Insert spaces before capital letters
        spaced = re.sub(r'([A-Z])', r' \1', concept).strip()
        # Remove multiple spaces
        spaced = re.sub(r'\s+', ' ', spaced)
        return spaced

    # ── NEW: Auto-balancing logic ─────────────────────────────────────────────

    def compute_balance_adjustment(
        self,
        parent_vals: Dict[int, Optional[float]],
        note_items: List[Tuple[str, Dict[int, Optional[float]]]],
        years: List[int],
    ) -> Dict[int, Optional[float]]:
        """
        Calculate the adjustment needed to make Notes sum equal to parent total.

        Returns {year: adjustment_value} where adjustment = parent - sum(notes)
        """
        adjustments: Dict[int, Optional[float]] = {}

        for year in years:
            parent = parent_vals.get(year)
            if parent is None:
                adjustments[year] = None
                continue

            notes_sum = sum(
                item_vals.get(year, 0) or 0
                for _, item_vals in note_items
            )

            adjustment = parent - notes_sum

            # Only create adjustment if non-trivial (> $1K USD)
            if abs(adjustment) > 1:
                adjustments[year] = adjustment
            else:
                adjustments[year] = None

        return adjustments

    # ── High-level builder ────────────────────────────────────────────────────

    def build_financial_data(
        self,
        ticker: str,
        years: List[int],
    ) -> Dict[Tuple, Dict[int, Optional[float]]]:
        """
        Fetch XBRL data and return comprehensive mapping including Notes details.

        Returns:
          {(col_c, col_b, col_d): {year: k_usd_value_or_None}}
        """
        cik = self.get_cik(ticker)
        facts = self.get_company_facts(cik)
        fye_month = self.detect_fye_month(facts)
        logger.info(f"{ticker} fiscal year-end month: {fye_month}")

        result: Dict[Tuple, Dict[int, Optional[float]]] = {}

        # ── Step 1: Extract top-level statements ──────────────────────────────
        logger.info("Step 1: Extracting top-level statement items")

        for col_c, col_b, col_d, concepts, factor, is_flow in XBRL_MAP:
            raw = self.extract_concept_by_year(
                facts, concepts, years, is_flow, fye_month
            )
            converted = {}
            for year, val in raw.items():
                if val is not None:
                    converted[year] = round(val * factor)
                else:
                    converted[year] = None

            converted = self._propagate_zero(converted, years)
            key = (col_c, col_b, col_d)
            result[key] = converted

            logger.info(
                f"  [{col_c}] {col_d[:40]}: "
                + " | ".join(
                    f"{y}:{v:,}" if v is not None else f"{y}:–"
                    for y, v in sorted(converted.items())
                )
            )

        # ── Step 2: Extract Notes details ─────────────────────────────────────
        logger.info("Step 2: Extracting Notes-level details")

        for category_name, pattern_config in NOTES_PATTERNS.items():
            logger.info(f"  Processing {category_name} Notes...")

            # Get parent statement values
            parent_key = pattern_config["parent_key"]
            parent_vals = result.get(parent_key, {})

            # Extract note details
            note_items = self.extract_notes_details(
                facts, pattern_config, years, fye_month
            )

            if not note_items:
                logger.info(f"    No {category_name} notes found")
                continue

            # Compute balance adjustment
            adjustment_vals = self.compute_balance_adjustment(
                parent_vals, note_items, years
            )

            # Add note items to result with col_b = 1, 2, 3...
            parent_col_c = parent_key[0]
            for idx, (display_name, item_vals) in enumerate(note_items, start=1):
                # Convert to K USD
                converted = {
                    year: round(val / 1000) if val is not None else None
                    for year, val in item_vals.items()
                }

                note_key = (parent_col_c, idx, display_name)
                result[note_key] = converted

                logger.info(
                    f"    [{parent_col_c}|{idx}] {display_name[:40]}: "
                    + " | ".join(
                        f"{y}:{v:,}" if v is not None else f"{y}:–"
                        for y, v in sorted(converted.items())
                    )
                )

            # Add adjustment row if needed
            if any(v is not None and abs(v) > 0 for v in adjustment_vals.values()):
                adj_idx = len(note_items) + 1
                adj_key = (parent_col_c, adj_idx, f"Other {category_name} (Balancing)")
                result[adj_key] = adjustment_vals

                logger.info(
                    f"    [{parent_col_c}|{adj_idx}] Other {category_name} (Balancing): "
                    + " | ".join(
                        f"{y}:{v:,}" if v is not None else f"{y}:–"
                        for y, v in sorted(adjustment_vals.items())
                    )
                )

        # ── Step 3: Validate accounting equation ──────────────────────────────
        logger.info("Step 3: Validating accounting equation")
        self._validate_accounting_equation(result, years)

        return result

    def _validate_accounting_equation(
        self,
        data: Dict[Tuple, Dict[int, Optional[float]]],
        years: List[int],
    ) -> None:
        """
        Validate: Total Assets = Total Liabilities + Total Equity
        Log warnings for any imbalances.
        """
        for year in years:
            # Sum assets (BS items with no negative convention)
            assets = sum(
                vals.get(year, 0) or 0
                for (col_c, col_b, col_d), vals in data.items()
                if col_c == "BS" and any(kw in col_d for kw in [
                    "Cash", "Receivable", "Prepaid", "Assets", "Equipment", "Lease"
                ])
            )

            # Sum liabilities
            liabilities = sum(
                vals.get(year, 0) or 0
                for (col_c, col_b, col_d), vals in data.items()
                if col_c == "BS" and any(kw in col_d for kw in [
                    "Payable", "Accrued", "Debt", "Liabilities"
                ])
            )

            # Sum equity
            equity = sum(
                vals.get(year, 0) or 0
                for (col_c, col_b, col_d), vals in data.items()
                if col_c == "BS" and any(kw in col_d for kw in [
                    "Shares", "Capital", "Comprehensive", "Deficit"
                ])
            )

            balance = assets - (liabilities + equity)

            if abs(balance) > 10:  # Allow $10K tolerance
                logger.warning(
                    f"  Year {year}: Accounting equation imbalance = {balance:,.0f} K USD"
                    f" (Assets={assets:,.0f}, Liabilities={liabilities:,.0f}, Equity={equity:,.0f})"
                )
            else:
                logger.info(
                    f"  Year {year}: Balanced ✓ "
                    f"(Assets={assets:,.0f}, L+E={liabilities + equity:,.0f})"
                )
