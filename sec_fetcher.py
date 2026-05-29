"""
sec_fetcher.py – SEC EDGAR XBRL data retrieval and financial concept mapping.

Fetches company facts from:
  https://data.sec.gov/api/xbrl/companyfacts/CIK{padded}.json

Returns a structured dict keyed by (col_c, col_b, col_d) → {year: k_usd_value}
ready to be written into the FY DATA K USD sheet.
"""

import logging
import re
import time
from typing import Dict, List, Optional, Tuple

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# XBRL concept mapping
# Each entry: (col_c, col_b, col_d, [xbrl_concepts], factor, is_flow)
#   col_c     – statement type string in col C of the Excel sheet
#   col_b     – sub-key (None, 1, 2 …) in col B
#   col_d     – exact line-item name in col D
#   concepts  – ordered list of XBRL US-GAAP concept names to try
#   factor    – multiply raw USD value by this to get the K USD storage value
#               e.g. 1/1000 → USD → K USD
#                    1      → store raw value (shares count)
#   is_flow   – True  = income/cash-flow (duration, filter fp='FY')
#               False = balance-sheet    (instant,  filter by end date)
# ─────────────────────────────────────────────────────────────────────────────

XBRL_MAP: List[Tuple] = [
    # ── Income Statement ──────────────────────────────────────────────────────
    ("IS", None, "Revenue",
     ["RevenueFromContractWithCustomerIncludingAssessedTax",
      "RevenueFromContractWithCustomerExcludingAssessedTax",
      "Revenues", "Revenue",
      "CollaborationRevenue"], 1 / 1000, True),

    ("IS", None, "Research And Development",
     # Some companies (e.g. Compass Therapeutics) use the more specific
     # "ExcludingAcquiredInProcessCost" tag for ongoing R&D.  Always try
     # the more specific tag first, then fall back to the generic one.
     ["ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost",
      "ResearchAndDevelopmentExpense"], 1 / 1000, True),

    ("IS", None, "General And Administrative",
     ["GeneralAndAdministrativeExpense"], 1 / 1000, True),

    # Interest Income – try several tags; CMPX reports from 2023+
    ("IS", None, "Interest Income",
     ["InvestmentIncomeInterest",
      "OtherNonoperatingIncome",
      "InterestAndOtherIncome",
      "InterestIncomeOther"], 1 / 1000, True),

    # Interest Expense
    ("IS", None, "Interest Expense",
     ["InterestExpense",
      "InterestExpenseDebt",
      "InterestExpenseNonoperating",
      "InterestAndDebtExpense"], 1 / 1000, True),

    # Income Tax
    ("IS", None, "Income Tax Provision",
     ["IncomeTaxExpenseBenefit",
      "CurrentIncomeTaxExpenseBenefit"], 1 / 1000, True),

    # Weighted-average shares – store ACTUAL count so the formula
    # =F18*1000/F20 in the K USD sheet computes correct EPS
    ("IS", None, "Weighted-Average Number Of Common Shares, Basic And Diluted",
     ["WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
      "WeightedAverageNumberOfSharesOutstandingBasic",
      "WeightedAverageNumberOfDilutedSharesOutstanding"], 1, True),

    # FX / OCI adjustment – works for both UK-listed (FX translation) and
    # US-listed (unrealized gains/losses on securities) companies
    ("IS", None, "Foreign Currency Translation Adjustment",
     ["OtherComprehensiveIncomeLossForeignCurrencyTransactionAndTranslationAdjustmentNetOfTax",
      "OtherComprehensiveIncomeForeignCurrencyTransactionAndTranslationAdjustmentNetOfTaxPortionAttributableToParent",
      "OtherComprehensiveIncomeForeignCurrencyTransactionAndTranslationGainLossArisingDuringPeriodNetOfTax",
      # US biotechs often report only unrealized investment gains/losses in OCI
      "OtherComprehensiveIncomeUnrealizedHoldingGainLossOnSecuritiesArisingDuringPeriodNetOfTax",
      "OtherComprehensiveIncomeLossAvailableForSaleSecuritiesAdjustmentNetOfTax"],
     1 / 1000, True),

    # ── Balance Sheet (instant items) ─────────────────────────────────────────
    ("BS", None, "Cash And Cash Equivalents",
     ["CashAndCashEquivalentsAtCarryingValue",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
     1 / 1000, False),

    # Marketable Securities - short-term investments (US Treasuries, money market, etc.)
    ("BS", None, "Marketable Securities",
     ["MarketableSecuritiesCurrent",
      "ShortTermInvestments",
      "AvailableForSaleSecuritiesCurrent",
      "InvestmentsCurrent"], 1 / 1000, False),

    # Explicitly map AR (often 0 for pre-revenue biotechs)
    ("BS", None, "Accounts Receivable",
     ["AccountsReceivableNetCurrent",
      "AccountsReceivableNet"], 1 / 1000, False),

    ("BS", None, "Prepaid Expenses And Other Current Assets",
     ["PrepaidExpenseAndOtherAssetsCurrent",
      "OtherAssetsCurrent",
      "PrepaidExpenseCurrent"], 1 / 1000, False),

    # R&D incentives / VAT receivable
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

    # Accrued Exp – combined AP+accrued tag used; AP is shown separately
    # so downstream users should be aware of potential overlap
    ("BS", None, "Accrued Expenses And Other Current Liabilities",
     ["AccountsPayableAndOtherAccruedLiabilitiesCurrent",
      "AccruedLiabilitiesCurrent",
      "EmployeeRelatedLiabilitiesCurrent"], 1 / 1000, False),

    # Deferred Revenue (often 0 for pre-revenue biotechs)
    ("BS", None, "Deferred revenue, Current Portion",
     ["DeferredRevenueCurrent",
      "ContractWithCustomerLiabilityCurrent"], 1 / 1000, False),

    ("BS", None, "Debt, Current Portion",
     ["DebtCurrent",
      "LongTermDebtCurrent",
      "ConvertibleNotesPayableCurrent",
      "NotesPayableRelatedPartiesClassifiedCurrent"], 1 / 1000, False),

    # Operating Lease Liability - Current Portion
    ("BS", None, "Operating Lease Liabilities, Current Portion",
     ["OperatingLeaseLiabilityCurrent"], 1 / 1000, False),

    ("BS", None, "Long-Term Debt, Net Of Discount",
     ["LongTermDebtNoncurrent",
      "LongTermDebt"], 1 / 1000, False),

    ("BS", None, "Operating Lease Liabilities, Net Of Current Portion",
     ["OperatingLeaseLiabilityNoncurrent"], 1 / 1000, False),

    # Deferred Revenue Non-Current
    ("BS", None, "Deferred Revenue, Net Of Current Portion",
     ["DeferredRevenueNoncurrent",
      "ContractWithCustomerLiabilityNoncurrent"], 1 / 1000, False),

    # Other Long-Term Liabilities
    ("BS", None, "Other Long-Term Liabilities",
     ["OtherLiabilitiesNoncurrent",
      "OtherAccruedLiabilitiesNoncurrent"], 1 / 1000, False),

    # Equity section
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

    # ── ISN: R&D Notes (single row = total R&D) ─────────────────────────────
    ("ISN", 1, "Research And Development Expenses",
     ["ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost",
      "ResearchAndDevelopmentExpense"], 1 / 1000, True),

    # ── ISN: G&A Notes (single row = total G&A) ─────────────────────────────
    ("ISN", 1, "General And Administrative Expenses",
     ["GeneralAndAdministrativeExpense"], 1 / 1000, True),

    # ── BSN: PP&E Notes ──────────────────────────────────────────────────────
    ("BSN", 1, "Property, Plant And Equipment, Gross",
     ["PropertyPlantAndEquipmentGross"], 1 / 1000, False),

    ("BSN", 2, "Accumulated Depreciation",
     ["AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment"],
     -1 / 1000, False),  # NEGATIVE: gross - accum dep = net

    # ── BSN: Accrued Expense Notes ───────────────────────────────────────────
    ("BSN", 1, "Accrued Employee Benefits",
     ["AccruedEmployeeBenefitsCurrent",
      "EmployeeRelatedLiabilitiesCurrent"], 1 / 1000, False),

    ("BSN", 2, "Other Accrued Liabilities",
     ["OtherAccruedLiabilitiesCurrent"], 1 / 1000, False),
]


# ── IFRS concept alternatives ───────────────────────────────────────────────
# Key: (col_c, col_b, col_d) — same as XBRL_MAP.  Value: list of ifrs-full
# concept names, tried in order.  Only used when the taxonomy is "ifrs-full".
IFRS_CONCEPTS: Dict[Tuple, List[str]] = {
    # Income Statement
    ("IS", None, "Revenue"):
        ["Revenue", "RevenueFromContractsWithCustomers"],
    ("IS", None, "Research And Development"):
        ["ResearchAndDevelopmentExpense"],
    ("IS", None, "General And Administrative"):
        ["SellingGeneralAndAdministrativeExpense"],
    ("IS", None, "Interest Income"):
        ["FinanceIncome",
         "InterestIncomeForFinancialAssetsMeasuredAtAmortisedCost"],
    ("IS", None, "Interest Expense"):
        ["FinanceCosts"],
    ("IS", None, "Income Tax Provision"):
        ["IncomeTaxExpenseContinuingOperations"],
    ("IS", None, "Weighted-Average Number Of Common Shares, Basic And Diluted"):
        ["WeightedAverageShares", "AdjustedWeightedAverageShares"],
    ("IS", None, "Foreign Currency Translation Adjustment"):
        ["OtherComprehensiveIncomeNetOfTaxExchangeDifferencesOnTranslation",
         "OtherComprehensiveIncomeNetOfTaxGainsLossesOnRemeasurementsOfDefinedBenefitPlans",
         "OtherComprehensiveIncome"],
    # Balance Sheet
    ("BS", None, "Cash And Cash Equivalents"):
        ["CashAndCashEquivalents"],
    ("BS", None, "Marketable Securities"):
        ["ShorttermDepositsNotClassifiedAsCashEquivalents"],
    ("BS", None, "Accounts Receivable"):
        ["CurrentTradeReceivables", "TradeAndOtherCurrentReceivables"],
    ("BS", None, "Prepaid Expenses And Other Current Assets"):
        ["OtherCurrentAssets",
         "CurrentPrepaymentsAndCurrentAccruedIncomeOtherThanCurrentContractAssets",
         "CurrentPrepaidExpenses"],
    ("BS", None, "Research And Development Incentives Receivable"):
        ["CurrentReceivablesFromTaxesOtherThanIncomeTax",
         "CurrentValueAddedTaxReceivables"],
    ("BS", None, "Property And Equipment, Net"):
        ["PropertyPlantAndEquipment",
         "PropertyPlantAndEquipmentIncludingRightofuseAssets"],
    ("BS", None, "Operating Lease Right-Of-Use Assets"):
        [],  # IFRS includes ROU in PPE; no separate concept typically
    ("BS", None, "Other Assets"):
        ["IntangibleAssetsOtherThanGoodwill"],
    ("BS", None, "Accounts Payable"):
        ["TradeAndOtherCurrentPayablesToTradeSuppliers",
         "TradeAndOtherCurrentPayables"],
    ("BS", None, "Accrued Expenses And Other Current Liabilities"):
        ["CurrentAccruedExpensesAndOtherCurrentLiabilities",
         "ShorttermEmployeeBenefitsAccruals"],
    ("BS", None, "Deferred revenue, Current Portion"):
        ["CurrentContractLiabilities"],
    ("BS", None, "Debt, Current Portion"):
        [],  # varies widely in IFRS
    ("BS", None, "Operating Lease Liabilities, Current Portion"):
        ["CurrentLeaseLiabilities"],
    ("BS", None, "Long-Term Debt, Net Of Discount"):
        ["NoncurrentPayables"],
    ("BS", None, "Operating Lease Liabilities, Net Of Current Portion"):
        ["NoncurrentLeaseLiabilities"],
    ("BS", None, "Deferred Revenue, Net Of Current Portion"):
        ["NoncurrentContractLiabilities"],
    ("BS", None, "Other Long-Term Liabilities"):
        ["NoncurrentProvisionsForEmployeeBenefits"],
    ("BS", None, "Ordinary Shares, \u00a30.01 Nominal Value"):
        ["IssuedCapital"],
    ("BS", None, "Additional Paid-In Capital"):
        ["AdditionalPaidinCapital"],
    ("BS", None, "Accumulated Other Comprehensive (Loss) Income"):
        [],  # IFRS: often embedded in equity movement
    ("BS", None, "Accumulated Deficit"):
        ["RetainedEarnings"],
    # Cash Flow Statement
    ("CFS", None, "Net Cash Used In Operating Activities"):
        ["CashFlowsFromUsedInOperatingActivities"],
    ("CFS", None, "Net Cash Provided by (Used In) Investing Activities"):
        ["CashFlowsFromUsedInInvestingActivities"],
    ("CFS", None, "Net Cash Provided By Financing Activities"):
        ["CashFlowsFromUsedInFinancingActivities"],
    ("CFS", None, "Effect Of Exchange Rate Changes On Cash And Cash Equivalents"):
        ["EffectOfExchangeRateChangesOnCashAndCashEquivalents"],
    # ISN/BSN notes
    ("ISN", 1, "Research And Development Expenses"):
        ["ResearchAndDevelopmentExpense"],
    ("ISN", 1, "General And Administrative Expenses"):
        ["SellingGeneralAndAdministrativeExpense"],
    ("BSN", 1, "Property, Plant And Equipment, Gross"):
        [],  # IFRS doesn't always split gross/accumulated
    ("BSN", 2, "Accumulated Depreciation"):
        [],  # IFRS doesn't always have separate concept
    ("BSN", 1, "Accrued Employee Benefits"):
        ["ShorttermEmployeeBenefitsAccruals"],
    ("BSN", 2, "Other Accrued Liabilities"):
        ["OtherCurrentLiabilities"],
    # BS totals (for residual balancing)
    ("_total", None, "Assets"):
        ["Assets"],
    ("_total", None, "Liabilities"):
        ["Liabilities"],
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
        """
        Resolve ticker → zero-padded 10-digit CIK string.
        Uses the SEC EDGAR full-text search API.
        """
        # Try the ticker lookup endpoint first
        url = "https://efts.sec.gov/LATEST/search-index"
        params = {
            "q": f'"{ticker.upper()}"',
            "dateRange": "custom",
            "startdt": "2018-01-01",
            "forms": "10-K,20-F",
        }
        resp = self.session.get(url, params=params, timeout=20)
        resp.raise_for_status()

        # Try the company search page as a fallback / primary
        r2 = self.session.get(
            self.COMPANY_SEARCH_URL,
            params={"CIK": ticker, "type": "10-K,20-F", "action": "getcompany"},
            timeout=20,
        )
        ciks = re.findall(r"CIK=(\d+)", r2.text)
        if ciks:
            cik = str(int(ciks[0])).zfill(10)
            logger.info(f"Resolved {ticker} → CIK {cik}")
            return cik

        # Also try the dedicated tickers JSON
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
        time.sleep(0.3)           # be polite to the SEC rate limiter
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        self._facts_cache[cik] = data
        return data

    # ── Taxonomy / form-type detection ──────────────────────────────────────────

    @staticmethod
    def _detect_taxonomy(facts: dict) -> Tuple[str, tuple, str]:
        """Detect which XBRL taxonomy the company uses.

        Returns (taxonomy_key, form_types, primary_currency).
          - taxonomy_key: 'us-gaap' or 'ifrs-full'
          - form_types:   ('10-K', '10-K/A') or ('20-F', '20-F/A')
          - primary_currency: 'USD', 'CHF', 'EUR', etc.
        """
        us_gaap = facts.get("facts", {}).get("us-gaap", {})
        ifrs = facts.get("facts", {}).get("ifrs-full", {})

        if len(us_gaap) >= len(ifrs):
            return "us-gaap", ("10-K", "10-K/A"), "USD"

        # IFRS company — detect primary monetary currency
        currency_counts: Dict[str, int] = {}
        for concept_data in ifrs.values():
            for unit_key in concept_data.get("units", {}):
                if unit_key not in ("shares", "pure", "employee") and "/" not in unit_key:
                    currency_counts[unit_key] = currency_counts.get(unit_key, 0) + 1
        currency = max(currency_counts, key=currency_counts.get) if currency_counts else "USD"
        return "ifrs-full", ("20-F", "20-F/A"), currency

    # ── Per-concept value extraction ──────────────────────────────────────────

    @staticmethod
    def _best_value(
        entries: list,
        year: int,
        is_flow: bool,
        fiscal_year_end_month: int = 12,
        form_types: tuple = ("10-K", "10-K/A"),
    ) -> Optional[float]:
        """
        From a list of XBRL entries for one concept, return the best value
        for the given fiscal year, or None if not found.

        Deduplication rule: if multiple filings report the same period,
        take the entry with the latest 'filed' date.
        """
        fye_month_str = f"{fiscal_year_end_month:02d}"

        if is_flow:
            # Duration items: fp='FY', end year == target year
            candidates = [
                e for e in entries
                if e.get("form") in form_types
                and e.get("fp") == "FY"
                and e.get("end", "")[:4] == str(year)
            ]
        else:
            # Instant (balance sheet): end date = YYYY-MM-DD where MM=fye month
            candidates = [
                e for e in entries
                if e.get("form") in form_types
                and e.get("end", "")[:4] == str(year)
                and e.get("end", "")[5:7] == fye_month_str
            ]

        if not candidates:
            return None

        # Take the latest-filed entry
        candidates.sort(key=lambda x: x.get("filed", ""), reverse=True)
        return candidates[0]["val"]

    def extract_concept_by_year(
        self,
        facts: dict,
        concepts: List[str],
        years: List[int],
        is_flow: bool,
        fiscal_year_end_month: int = 12,
        taxonomy: str = "us-gaap",
        form_types: tuple = ("10-K", "10-K/A"),
        currency: str = "USD",
    ) -> Dict[int, Optional[float]]:
        """
        Try each concept in order; return first non-None result per year.
        Returns {year: raw_value_or_None}.
        """
        tax_data = facts.get("facts", {}).get(taxonomy, {})
        result: Dict[int, Optional[float]] = {y: None for y in years}

        for concept in concepts:
            if concept not in tax_data:
                continue
            units = tax_data[concept].get("units", {})
            # For financial $$: use primary currency; for shares: use 'shares'
            entries = units.get(currency) or units.get("shares") or []

            for year in years:
                if result[year] is not None:
                    continue  # already found
                val = self._best_value(entries, year, is_flow,
                                        fiscal_year_end_month,
                                        form_types)
                if val is not None:
                    result[year] = val
                    logger.debug(
                        f"  {concept}[{year}] = {val:,.0f}"
                        f"  (is_flow={is_flow})"
                    )

        return result

    # ── Zero propagation ──────────────────────────────────────────────────────

    @staticmethod
    def _propagate_zero(
        year_vals: Dict[int, Optional[float]],
        years: List[int],
    ) -> Dict[int, Optional[float]]:
        """
        If value[Y] == 0 (explicitly from XBRL) and value[Y+1] is None,
        carry the 0 forward for all subsequent None years.

        Example: DebtCurrent = {2020: 7467, 2021: 0, 2022: None, 2023: None}
                 → {2020: 7467, 2021: 0, 2022: 0, 2023: 0}
        """
        result = dict(year_vals)
        for year in sorted(years):
            if result.get(year) == 0:
                for next_year in sorted(years):
                    if next_year <= year:
                        continue
                    if result.get(next_year) is None:
                        result[next_year] = 0
                    else:
                        break  # non-None value found; stop propagation
        return result

    # ── Detect fiscal year-end month ──────────────────────────────────────────

    @staticmethod
    def detect_fye_month(facts: dict) -> int:
        """
        Detect fiscal year-end month by scanning annual filing end dates.
        Supports both us-gaap (10-K) and ifrs-full (20-F) taxonomies.
        Returns the most common month (1-12).
        """
        from collections import Counter
        months: Counter = Counter()
        annual_forms = ("10-K", "10-K/A", "20-F", "20-F/A")
        for taxonomy_key in ("us-gaap", "ifrs-full"):
            tax_data = facts.get("facts", {}).get(taxonomy_key, {})
            for concept_data in tax_data.values():
                for unit_entries in concept_data.get("units", {}).values():
                    for e in unit_entries:
                        if e.get("form") in annual_forms and e.get("fp") == "FY":
                            end = e.get("end", "")
                            if len(end) >= 7:
                                months[int(end[5:7])] += 1
        if months:
            return months.most_common(1)[0][0]
        return 12  # default to December

    # ── Reporting unit detection ───────────────────────────────────────────────

    def detect_reporting_unit(
        self, cik: str, years: List[int],
        form_types: tuple = ("10-K", "10-K/A"),
    ) -> str:
        """Detect whether the company reports in thousands or millions.

        Scans the most recent annual filing HTML for phrases like
        'in thousands' or 'in millions' near financial statements.

        Returns 'K' (thousands) or 'MM' (millions). Default: 'K'.
        """
        try:
            filing_urls = self._get_10k_filing_urls(cik, years,
                                                     form_types=form_types)
            if not filing_urls:
                logger.info("No annual filings found; defaulting to K")
                return "K"

            latest_year = max(filing_urls.keys())
            url = filing_urls[latest_year]
            form_label = "20-F" if "20-F" in form_types else "10-K"
            logger.info(f"Detecting reporting unit from FY{latest_year} {form_label} ...")
            time.sleep(0.3)
            resp = self.session.get(url, timeout=120)
            resp.raise_for_status()
            text = resp.text.lower()

            # Count occurrences of each unit declaration
            k_count = len(re.findall(r'\bin\s+thousands\b', text))
            m_count = len(re.findall(r'\bin\s+millions\b', text))

            if m_count > k_count:
                logger.info(f"  Detected: MM (millions={m_count}, thousands={k_count})")
                return "MM"
            else:
                logger.info(f"  Detected: K (thousands={k_count}, millions={m_count})")
                return "K"
        except Exception as e:
            logger.warning(f"Unit detection failed ({e}); defaulting to K")
            return "K"

    # ── 10-K Filing HTML Parsing (for ISN/BSN note details) ─────────────────

    def _get_10k_filing_urls(
        self, cik: str, years: List[int],
        form_types: tuple = ("10-K", "10-K/A"),
    ) -> Dict[int, str]:
        """Get annual filing URLs mapped by fiscal year.

        Supports both 10-K (US-GAAP) and 20-F (IFRS) form types.
        Returns {fiscal_year: filing_url} for each available filing.
        """
        url = f"https://data.sec.gov/submissions/CIK{cik}.json"
        time.sleep(0.3)
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        sub = resp.json()

        recent = sub.get("filings", {}).get("recent", {})
        forms = recent.get("form", [])
        accessions = recent.get("accessionNumber", [])
        primary_docs = recent.get("primaryDocument", [])
        report_dates = recent.get("reportDate", [])

        cik_num = str(int(cik))
        filing_urls: Dict[int, str] = {}

        min_year = min(years) - 1
        max_year = max(years) + 1

        for i, form in enumerate(forms):
            if form not in form_types:
                continue
            report_year = int(report_dates[i][:4])
            if report_year < min_year or report_year > max_year:
                continue
            acc = accessions[i].replace("-", "")
            furl = (
                f"https://www.sec.gov/Archives/edgar/data/"
                f"{cik_num}/{acc}/{primary_docs[i]}"
            )
            # Keep the first (most recent) filing per fiscal year
            if report_year not in filing_urls:
                filing_urls[report_year] = furl

        return filing_urls

    @staticmethod
    def _parse_note_table(
        table,
    ) -> Tuple[List[int], List[Tuple[str, List[float]]]]:
        """Parse an HTML note table into years and line-items.

        Returns (years, items) where:
          years : fiscal years from the header (e.g. [2024, 2023])
          items : [(item_name, [value_per_year_column])]
                  values are in the table's unit (usually 000's = K USD)
        """
        rows = []
        for tr in table.find_all("tr"):
            cells = [
                td.get_text(strip=True)
                .replace("\xa0", " ")
                .replace("\u00a0", " ")
                for td in tr.find_all(["td", "th"])
            ]
            rows.append(cells)

        # ── Detect year columns from the header ──────────────────────────
        years: List[int] = []
        has_change = False
        for row in rows[:6]:
            row_text = " ".join(row)
            found = re.findall(r"\b(20\d{2})\b", row_text)
            if found:
                years = [int(y) for y in found]
                break
        for row in rows[:6]:
            if any("hange" in c for c in row):
                has_change = True
                break

        if not years:
            return [], []

        # The R&D table has a "Change" column → 3 numeric columns
        # PP&E / Accrued tables have only year columns → 2 numeric columns
        num_year_cols = len(years)

        # ── Parse data rows ──────────────────────────────────────────────
        items: List[Tuple[str, List[float]]] = []
        for row in rows:
            name = None
            values: List[float] = []
            in_negative = False

            for cell in row:
                cell = cell.strip()
                if not cell or cell == "$":
                    continue
                if cell == "(":
                    in_negative = True
                    continue
                if cell == ")":
                    in_negative = False
                    continue

                # Try to parse as a number
                cleaned = cell.replace(",", "").replace("$", "").strip()
                if cleaned in ("\u2014", "\u2013", "\u2012", "—", "-", "–", ""):
                    values.append(0.0)
                    continue

                # Handle parenthesised negatives in the same cell: "(6,361)"
                is_neg_paren = False
                if cleaned.startswith("(") and cleaned.endswith(")"):
                    cleaned = cleaned[1:-1]
                    is_neg_paren = True
                elif cleaned.startswith("("):
                    cleaned = cleaned[1:]
                    is_neg_paren = True

                try:
                    val = float(cleaned)
                    if is_neg_paren or in_negative:
                        val = -val
                        in_negative = False
                    values.append(val)
                except ValueError:
                    # Non-numeric → item name (take the first one)
                    if name is None and len(cell) > 1:
                        name = cell

            if not name or not values:
                continue

            # Skip total / net summary rows and header artifacts
            name_lower = name.lower().strip()
            if name_lower.startswith("total") or any(
                kw in name_lower
                for kw in [
                    "property and equipment, net",
                    "property and equipment\u2013at cost",
                    "property and equipment–at cost",
                    "property and equipment -at cost",
                    "(000",
                ]
            ):
                continue
            # Skip year-header artifacts like "Change"
            if name_lower in ("change", "year ended", "year", "month"):
                continue
            # Skip table header artifacts ("In thousands", "In millions", etc.)
            if name_lower.startswith("in ") and any(
                kw in name_lower for kw in ["thousand", "million", "billion"]
            ):
                continue
            # Skip share-based / stock-based compensation (handled separately)
            if "share-based" in name_lower or "stock-based" in name_lower:
                continue

            # Keep only the year-column values (drop Change column if present)
            year_values = values[:num_year_cols]
            if len(year_values) == num_year_cols:
                items.append((name, year_values))

        return years, items

    def _find_note_tables(
        self, html: str
    ) -> Dict[str, Tuple[List[int], List[Tuple[str, List[float]]]]]:
        """Find and parse R&D, PP&E, Accrued, G&A tables from filing HTML.

        Returns {section_name: (years, items)} for each found section.
        """
        soup = BeautifulSoup(html, "html.parser")
        result: Dict[str, Tuple[List[int], List[Tuple[str, List[float]]]]] = {}

        # ── R&D table: contains "Total research and development" ─────────
        for table in soup.find_all("table"):
            if "otal research and development" in table.get_text().lower():
                years, items = self._parse_note_table(table)
                if items:
                    result["rd"] = (years, items)
                break

        # ── PP&E table: after "Property and equipment consist" ───────────
        for text_node in soup.find_all(
            string=re.compile(r"[Pp]roperty and equipment consist")
        ):
            next_table = text_node.find_next("table")
            if next_table:
                years, items = self._parse_note_table(next_table)
                if items:
                    result["ppe"] = (years, items)
                break

        # ── Accrued table: after "Accrued expenses consist" ──────────────
        for text_node in soup.find_all(
            string=re.compile(r"[Aa]ccrued expenses consist")
        ):
            next_table = text_node.find_next("table")
            if next_table:
                years, items = self._parse_note_table(next_table)
                if items:
                    result["accrued"] = (years, items)
                break

        # ── G&A table: contains "Total general and administrative" ───────
        for table in soup.find_all("table"):
            if "otal general and administrative" in table.get_text().lower():
                years, items = self._parse_note_table(table)
                if items:
                    result["ga"] = (years, items)
                break

        return result

    def fetch_note_details(
        self, cik: str, years: List[int],
        form_types: tuple = ("10-K", "10-K/A"),
    ) -> Dict[str, Optional[List[Tuple[str, Dict[int, float]]]]]:
        """Fetch detailed note breakdowns from annual filing HTML.

        Downloads a minimal set of annual filings (10-K or 20-F),
        parses note tables (R&D, PP&E, Accrued, G&A), and consolidates.

        Returns {
            'rd':      [(item_name, {year: k_usd_value}), ...] or None,
            'ga':      [(item_name, {year: k_usd_value}), ...] or None,
            'ppe':     [(item_name, {year: k_usd_value}), ...] or None,
            'accrued': [(item_name, {year: k_usd_value}), ...] or None,
        }
        """
        filing_urls = self._get_10k_filing_urls(cik, years,
                                                 form_types=form_types)
        form_label = "20-F" if "20-F" in form_types else "10-K"
        logger.info(
            f"Available {form_label} filings: "
            + ", ".join(f"FY{y}" for y in sorted(filing_urls))
        )

        # ── Select minimal filings to cover all target years ─────────────
        # Each filing has 2 year columns (current + prior).
        # Prefer filing Y+1 for year Y data (newer naming, comparatives).
        needed = sorted(years, reverse=True)
        filings_to_download: Dict[int, str] = {}
        covered: set = set()

        for y in needed:
            if y in covered:
                continue
            # Prefer filing for year y (its primary column is y)
            if y in filing_urls:
                filings_to_download[y] = filing_urls[y]
                covered.add(y)
                covered.add(y - 1)
            elif y + 1 in filing_urls:
                filings_to_download[y + 1] = filing_urls[y + 1]
                covered.add(y + 1)
                covered.add(y)

        logger.info(
            f"Downloading {len(filings_to_download)} filings: "
            + ", ".join(f"FY{y}" for y in sorted(filings_to_download))
        )

        # ── Download and parse each filing ───────────────────────────────
        # sections_by_year[section][year] = {item_name: value}
        sections_by_year: Dict[str, Dict[int, Dict[str, float]]] = {
            "rd": {},
            "ga": {},
            "ppe": {},
            "accrued": {},
        }

        for fy in sorted(filings_to_download, reverse=True):
            url = filings_to_download[fy]
            logger.info(f"  Downloading FY{fy} 10-K …")
            time.sleep(0.3)
            try:
                resp = self.session.get(url, timeout=120)
                resp.raise_for_status()
            except Exception as e:
                logger.warning(f"  Failed to download FY{fy} 10-K: {e}")
                continue

            tables = self._find_note_tables(resp.text)
            for section, (table_years, items) in tables.items():
                for col_idx, ty in enumerate(table_years):
                    if ty not in years:
                        continue
                    if ty in sections_by_year[section]:
                        continue  # already have from a more recent filing

                    year_data: Dict[str, float] = {}
                    for item_name, vals in items:
                        if col_idx < len(vals):
                            year_data[item_name] = vals[col_idx]

                    if year_data:
                        sections_by_year[section][ty] = year_data
                        logger.info(
                            f"    {section.upper()} FY{ty}: "
                            f"{len(year_data)} items"
                        )

        # ── Consolidate across years ─────────────────────────────────────
        result: Dict[str, Optional[List[Tuple[str, Dict[int, float]]]]] = {}

        for section, year_data in sections_by_year.items():
            if not year_data:
                result[section] = None
                continue

            # Build ordered union of item names (most recent year first)
            all_names: List[str] = []
            seen: set = set()
            for y in sorted(year_data, reverse=True):
                for name in year_data[y]:
                    if name not in seen:
                        all_names.append(name)
                        seen.add(name)

            # Build per-item {year: value} dicts; skip all-zero items
            consolidated: List[Tuple[str, Dict[int, float]]] = []
            for name in all_names:
                yr_vals: Dict[int, float] = {}
                for y in years:
                    if y in year_data and name in year_data[y]:
                        yr_vals[y] = year_data[y][name]
                    else:
                        yr_vals[y] = 0
                # Skip items that are all zeros across all years
                if all(v == 0 for v in yr_vals.values()):
                    continue
                consolidated.append((name, yr_vals))

            logger.info(
                f"  {section.upper()} consolidated: {len(consolidated)} items "
                f"across {len(year_data)} years"
            )
            for name, yr_vals in consolidated:
                logger.info(
                    f"    {name:<45} "
                    + " | ".join(
                        f"{y}:{v:>8,.0f}" for y, v in sorted(yr_vals.items())
                    )
                )

            result[section] = consolidated

        # ── Post-processing: merge PP&E depreciation variants ────────────
        # Different filings may call it "Accumulated depreciation" vs
        # "Accumulated depreciation and amortization".  Merge into one row.
        if result.get("ppe"):
            dep_items = [
                (i, n, v)
                for i, (n, v) in enumerate(result["ppe"])
                if "depreci" in n.lower()
            ]
            if len(dep_items) > 1:
                merged_vals: Dict[int, float] = {}
                for _, _, yr_v in dep_items:
                    for y, v in yr_v.items():
                        if v != 0:
                            merged_vals[y] = v
                        elif y not in merged_vals:
                            merged_vals[y] = 0
                # Use the most common name
                merged_name = dep_items[0][1]
                # Remove old dep items (reverse order to preserve indices)
                for idx, _, _ in sorted(dep_items, reverse=True):
                    result["ppe"].pop(idx)
                result["ppe"].append((merged_name, merged_vals))
                logger.info(
                    f"  PPE: merged {len(dep_items)} depreciation items → "
                    f"'{merged_name}'"
                )

        return result

    # ── High-level builder ────────────────────────────────────────────────────

    def build_financial_data(
        self,
        ticker: str,
        years: List[int],
        cik_override: Optional[str] = None,
        unit_override: Optional[str] = None,
    ) -> Tuple[Dict[Tuple, Dict[int, Optional[float]]], Dict[Tuple, str], Dict, str, str]:
        """
        Fetch XBRL data and return a 5-tuple:
          (financial_data, rename_map, note_details, reporting_unit, currency)

        financial_data: {(col_c, col_b, col_d): {year: storage_value_or_None}}
        rename_map:     {(col_c, col_b, original_col_d): new_col_d_text}
        note_details:   {section: [(item_name, {year: value})] or None}
        reporting_unit: 'K' (thousands) or 'MM' (millions)
        currency:       'USD', 'CHF', 'EUR', etc.

        Values are converted using the factor from XBRL_MAP, scaled by
        reporting_unit: K → factor as-is, MM → factor/1000 for USD items.
        None means no data found for that cell.
        """
        if cik_override:
            cik = cik_override.zfill(10)
            logger.info(f"Using CIK override: {cik}")
        else:
            cik = self.get_cik(ticker)
        facts = self.get_company_facts(cik)
        fye_month = self.detect_fye_month(facts)
        logger.info(f"{ticker} fiscal year-end month: {fye_month}")

        # Detect taxonomy (us-gaap vs ifrs-full) and form types (10-K vs 20-F)
        taxonomy, form_types, xbrl_currency = self._detect_taxonomy(facts)
        if taxonomy == "ifrs-full":
            logger.info(
                f"IFRS company detected (taxonomy={taxonomy}, "
                f"currency={xbrl_currency}, forms={form_types})"
            )

        # Detect reporting unit
        if unit_override:
            reporting_unit = unit_override
            logger.info(f"Using unit override: {reporting_unit}")
        else:
            reporting_unit = self.detect_reporting_unit(
                cik, years, form_types=form_types,
            )

        # MM mode: scale USD factors by additional 1/1000
        # factor=1/1000 → 1/1000000, factor=1 (shares) → 1 (unchanged)
        mm_scale = 1.0 if reporting_unit == "K" else 1 / 1000

        result: Dict[Tuple, Dict[int, Optional[float]]] = {}

        for col_c, col_b, col_d, concepts, factor, is_flow in XBRL_MAP:
            effective_factor = factor if factor == 1 else factor * mm_scale

            # For IFRS companies: use IFRS concepts instead of US-GAAP
            if taxonomy == "ifrs-full":
                key = (col_c, col_b, col_d)
                ifrs_concepts = IFRS_CONCEPTS.get(key, [])
                raw = self.extract_concept_by_year(
                    facts, ifrs_concepts, years, is_flow, fye_month,
                    taxonomy=taxonomy, form_types=form_types,
                    currency=xbrl_currency,
                )
            else:
                raw = self.extract_concept_by_year(
                    facts, concepts, years, is_flow, fye_month,
                    taxonomy=taxonomy, form_types=form_types,
                    currency=xbrl_currency,
                )
            converted = {}
            has_any_data = False
            for year, val in raw.items():
                if val is not None:
                    converted[year] = round(val * effective_factor)
                    has_any_data = True
                else:
                    converted[year] = None

            # ── Zero-propagation: if a concept was explicitly 0 in year Y
            # and None in year Y+1, carry the 0 forward.  Handles "debt
            # paid off" situations where zero values are omitted from XBRL.
            converted = self._propagate_zero(converted, years)

            # ── For Balance Sheet items with no data, write 0 to override placeholders
            if not has_any_data and col_c == "BS":
                for year in years:
                    if converted[year] is None:
                        converted[year] = 0
                        logger.info(
                            f"  Setting to 0 (no XBRL data): ({col_c}, {col_b!r}, '{col_d}') year={year}"
                        )

            # ── For CFS FX effect: US companies have no FX → set all None to 0
            if col_c == "CFS" and "Exchange Rate" in col_d and not has_any_data:
                for year in years:
                    if converted[year] is None:
                        converted[year] = 0
                logger.info(
                    f"  Setting FX effect to 0 (no XBRL data): ({col_c}, {col_b!r}, '{col_d}')"
                )

            # ── For AOCI/OCI and similar items: replace first-year None with 0
            # (companies often don't report AOCI/OCI in early years if it's zero)
            if col_c in ("BS", "IS") and any(kw in col_d for kw in ["Comprehensive", "Translation"]):
                for year in sorted(years):
                    if converted[year] is None:
                        converted[year] = 0
                        logger.info(
                            f"  Setting first-year AOCI/OCI to 0: ({col_c}, {col_b!r}, '{col_d}') year={year}"
                        )
                        break  # Only fix the first None

            key = (col_c, col_b, col_d)
            result[key] = converted
            logger.info(
                f"  Mapped ({col_c}, {col_b!r}, '{col_d[:40]}') → "
                + " | ".join(
                    f"{y}:{v:,}" if v is not None else f"{y}:–"
                    for y, v in sorted(converted.items())
                )
            )

        # ── Post-processing: Add Restricted Cash to Other Assets ──────────────
        # RestrictedCashNoncurrent is often reported separately but should be
        # combined with OtherAssetsNoncurrent in the Excel template
        other_assets_key = ("BS", None, "Other Assets")
        if other_assets_key in result:
            restricted_cash_raw = self.extract_concept_by_year(
                facts,
                ["RestrictedCashNoncurrent",
                 "RestrictedCashAndCashEquivalentsNoncurrent"],
                years,
                is_flow=False,
                fiscal_year_end_month=fye_month,
                taxonomy=taxonomy, form_types=form_types,
                currency=xbrl_currency,
            )

            has_restricted = any(v is not None for v in restricted_cash_raw.values())
            if has_restricted:
                logger.info("  Adding Restricted Cash (Non-Current) to Other Assets")
                for year in years:
                    rc_val = restricted_cash_raw.get(year)
                    if rc_val is not None:
                        rc_divisor = 1000 if reporting_unit == "K" else 1000000
                        rc_rounded = round(rc_val / rc_divisor)
                        current_val = result[other_assets_key].get(year, 0) or 0
                        result[other_assets_key][year] = current_val + rc_rounded
                        logger.info(
                            f"    {year}: Other Assets {current_val:,} + "
                            f"Restricted Cash {rc_rounded:,} = "
                            f"{result[other_assets_key][year]:,}"
                        )

        # ── Post-processing: Compute Accrued Balancing Item ──────────────────
        # balancing = Total Accrued (R68) - Employee Benefits - Other Accrued
        accrued_total_key = ("BS", None, "Accrued Expenses And Other Current Liabilities")
        emp_benefits_key = ("BSN", 1, "Accrued Employee Benefits")
        other_accrued_key = ("BSN", 2, "Other Accrued Liabilities")
        balancing_key = ("BSN", 3, "Other Accrued Liabilities (Balancing)")

        if all(k in result for k in [accrued_total_key, emp_benefits_key, other_accrued_key]):
            balancing_vals: Dict[int, Optional[float]] = {}
            for year in years:
                total = result[accrued_total_key].get(year)
                emp = result[emp_benefits_key].get(year)
                other = result[other_accrued_key].get(year)
                if total is not None and emp is not None and other is not None:
                    balancing_vals[year] = total - emp - other
                elif total is not None:
                    # If sub-items are None, the balancing item is the whole total
                    balancing_vals[year] = total - (emp or 0) - (other or 0)
                else:
                    balancing_vals[year] = 0
            result[balancing_key] = balancing_vals
            logger.info(
                f"  Computed accrued balancing item: "
                + " | ".join(
                    f"{y}:{v:,}" if v is not None else f"{y}:–"
                    for y, v in sorted(balancing_vals.items())
                )
            )

        # ── Post-processing: Balance Sheet residual balancing ──────────────
        # The template computes Total Assets and Total L+SE from individual
        # items.  If our mapped items don't cover all BS items (e.g. goodwill,
        # intangibles, contingent consideration), the check (R88) won't be 0.
        # Fix: fetch XBRL totals and add residuals to catch-all rows.
        _ASSET_NAMES = [
            "Cash And Cash Equivalents",
            "Marketable Securities",
            "Prepaid Expenses And Other Current Assets",
            "Research And Development Incentives Receivable",
            "Property And Equipment, Net",
            "Operating Lease Right-Of-Use Assets",
            "Other Assets",
        ]
        _LIAB_NAMES = [
            "Accounts Payable",
            "Accrued Expenses And Other Current Liabilities",
            "Operating Lease Liabilities, Current Portion",
            "Debt, Current Portion",
            "Long-Term Debt, Net Of Discount",
            "Operating Lease Liabilities, Net Of Current Portion",
            "Deferred Revenue, Net Of Current Portion",
            "Other Long-Term Liabilities",
        ]
        _EQUITY_NAMES = [
            "Ordinary Shares, £0.01 Nominal Value",
            "Additional Paid-In Capital",
            "Accumulated Other Comprehensive (Loss) Income",
            "Accumulated Deficit",
        ]

        # Use IFRS concepts for totals if IFRS company
        asset_concepts = IFRS_CONCEPTS.get(("_total", None, "Assets"), ["Assets"]) \
            if taxonomy == "ifrs-full" else ["Assets"]
        liab_concepts = IFRS_CONCEPTS.get(("_total", None, "Liabilities"), ["Liabilities"]) \
            if taxonomy == "ifrs-full" else ["Liabilities"]

        xbrl_assets_raw = self.extract_concept_by_year(
            facts, asset_concepts, years, is_flow=False,
            fiscal_year_end_month=fye_month,
            taxonomy=taxonomy, form_types=form_types,
            currency=xbrl_currency,
        )
        xbrl_liab_raw = self.extract_concept_by_year(
            facts, liab_concepts, years, is_flow=False,
            fiscal_year_end_month=fye_month,
            taxonomy=taxonomy, form_types=form_types,
            currency=xbrl_currency,
        )

        # Divisor matches the effective factor for USD items
        bs_divisor = 1000 if reporting_unit == "K" else 1000000

        logger.info("Balance sheet residual balancing:")
        for year in years:
            xa = xbrl_assets_raw.get(year)
            xl = xbrl_liab_raw.get(year)
            if xa is None and xl is None:
                logger.info(f"  {year}: No XBRL BS totals; skipping")
                continue

            xa_k = round(xa / bs_divisor) if xa is not None else None
            xl_k = round(xl / bs_divisor) if xl is not None else None

            # ── Asset side residual → Other Assets ───────────────────
            if xa_k is not None:
                mapped_a = sum(
                    (result.get(("BS", None, n), {}).get(year) or 0)
                    for n in _ASSET_NAMES
                )
                a_res = xa_k - mapped_a
                if a_res != 0:
                    oa_key = ("BS", None, "Other Assets")
                    result[oa_key][year] = (result[oa_key].get(year) or 0) + a_res
                    logger.info(
                        f"  {year}: Other Assets += {a_res:,} "
                        f"(XBRL total={xa_k:,}, mapped={mapped_a:,})"
                    )

            # ── Liability side residual → Other LT Liabilities ───────
            if xl_k is not None:
                mapped_l = sum(
                    (result.get(("BS", None, n), {}).get(year) or 0)
                    for n in _LIAB_NAMES
                )
                l_res = xl_k - mapped_l
                if l_res != 0:
                    ol_key = ("BS", None, "Other Long-Term Liabilities")
                    result[ol_key][year] = (result[ol_key].get(year) or 0) + l_res
                    logger.info(
                        f"  {year}: Other LT Liab += {l_res:,} "
                        f"(XBRL total={xl_k:,}, mapped={mapped_l:,})"
                    )

            # ── Equity residual → APIC ──────────────────────────────
            # GAAP identity: Assets = Liabilities + Equity
            # Required equity = Assets - Liabilities; gap goes to APIC
            if xa_k is not None and xl_k is not None:
                required_e = xa_k - xl_k
                mapped_e = sum(
                    (result.get(("BS", None, n), {}).get(year) or 0)
                    for n in _EQUITY_NAMES
                )
                e_res = required_e - mapped_e
                if e_res != 0:
                    apic_key = ("BS", None, "Additional Paid-In Capital")
                    result[apic_key][year] = (result[apic_key].get(year) or 0) + e_res
                    logger.info(
                        f"  {year}: APIC += {e_res:,} "
                        f"(required equity={required_e:,}, mapped={mapped_e:,})"
                    )

        # ── Zero out BS items for pre-existence years ─────────────────────
        # If a year has no XBRL totals and the BS doesn't self-balance,
        # the company didn't exist as a separate entity → zero all BS items.
        for year in years:
            if xbrl_assets_raw.get(year) is not None:
                continue  # Already balanced above
            mapped_a = sum(
                (result.get(("BS", None, n), {}).get(year) or 0)
                for n in _ASSET_NAMES
            )
            mapped_l = sum(
                (result.get(("BS", None, n), {}).get(year) or 0)
                for n in _LIAB_NAMES
            )
            mapped_e = sum(
                (result.get(("BS", None, n), {}).get(year) or 0)
                for n in _EQUITY_NAMES
            )
            if mapped_a != mapped_l + mapped_e:
                logger.info(
                    f"  {year}: No XBRL totals and BS imbalanced "
                    f"(A={mapped_a:,} ≠ L+E={mapped_l + mapped_e:,}); "
                    f"zeroing BS items (pre-existence year)"
                )
                for n in _ASSET_NAMES + _LIAB_NAMES + _EQUITY_NAMES:
                    key = ("BS", None, n)
                    if key in result:
                        result[key][year] = 0

        # ── Build rename_map ─────────────────────────────────────────────────
        # Maps (col_c, col_b, ORIGINAL_excel_col_d) → new_col_d text
        rename_map: Dict[Tuple, str] = {}

        # IS: "Collaboration Revenues" → "Revenue"
        rename_map[("IS", None, "Collaboration Revenues")] = "Revenue"

        # BS: Repurpose unused rows for items that have no matching Excel row
        rename_map[("BS", None, "Accounts Receivable")] = "Marketable Securities"
        rename_map[("BS", None, "Deferred revenue, Current Portion")] = \
            "Operating Lease Liabilities, Current Portion"

        # ISN R&D Notes (rows R28-R36): map notes sub-keys to row positions
        # Original Excel has 9 rows (sub 1-9) with Bicycle Therapeutics names
        rename_map[("ISN", 1, None)] = "Research And Development Expenses"
        for sub in range(2, 10):
            rename_map[("ISN", sub, None)] = ""  # zero-fill rows

        # ISN G&A Notes (rows R42-R46): 5 rows
        rename_map[("ISN", 1, None, "GA")] = "General And Administrative Expenses"
        for sub in range(2, 6):
            rename_map[("ISN", sub, None, "GA")] = ""  # zero-fill rows

        # BSN PP&E Notes (rows R93-R97): 5 rows
        rename_map[("BSN", 1, None, "PPE")] = "Property, Plant And Equipment, Gross"
        rename_map[("BSN", 2, None, "PPE")] = ""  # zero
        rename_map[("BSN", 3, None, "PPE")] = ""  # zero
        rename_map[("BSN", 4, None, "PPE")] = ""  # zero
        rename_map[("BSN", 5, None, "PPE")] = "Accumulated Depreciation"

        # BSN Accrued Notes (rows R103-R107): 5 rows
        rename_map[("BSN", 1, None, "ACC")] = "Accrued Employee Benefits"
        rename_map[("BSN", 2, None, "ACC")] = "Other Accrued Liabilities (Balancing)"
        rename_map[("BSN", 3, None, "ACC")] = ""  # zero
        rename_map[("BSN", 4, None, "ACC")] = ""  # zero
        rename_map[("BSN", 5, None, "ACC")] = "Other Accrued Liabilities"

        logger.info(f"  Rename map: {len(rename_map)} entries")

        # ── Fetch note details from 10-K filing HTML ─────────────────────
        form_label = "20-F" if "20-F" in form_types else "10-K"
        logger.info(f"Fetching note details from {form_label} filings …")
        try:
            note_details = self.fetch_note_details(cik, years,
                                                    form_types=form_types)
        except Exception as e:
            logger.warning(f"Failed to fetch note details: {e}")
            note_details = {"rd": None, "ga": None, "ppe": None, "accrued": None}

        return result, rename_map, note_details, reporting_unit, xbrl_currency
