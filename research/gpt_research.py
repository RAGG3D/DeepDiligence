#!/usr/bin/env python3
"""
gpt_research.py — research/judgment layer via OpenAI GPT-5.6 Sol.

Division of labour (user directive 2026-07-07): GPT-5.6 Sol does the research and
judgement (what to fill, how to rate); Claude Code does the Excel mechanics. This
module is the connector — it calls the OpenAI API with a precise, format-locked
brief per task and saves the structured output for the Excel-writer steps to parse.

Requires an OpenAI **API key** (a ChatGPT subscription cannot be called by a
script). Set OPENAI_API_KEY in the environment or .env. Model defaults to
`gpt-5.6-sol` (override with --model; run --list-models to see what the key can use).

    python research/gpt_research.py --ticker MOLN --company-name "Molecular Partners AG" --brief all
    python research/gpt_research.py --list-models

Briefs:
  notes      FY DATA note breakdowns (R&D/G&A/PP&E/accrued), by year, CHF k
  pipeline   every clinical-stage drug + collaboration economics (MOLN net %)
  ratings    Best-In-Class / Tier One / Average per drug×indication (DB-grounded)
  events     Historical Events (Date | EVT | Category), rolling last 3 years
  movecauses cause per >=8% move date (reads {T}_moves_needing_research.json) →
             writes artifacts/{T}/{T}_moves_researched.json
"""
import argparse
import csv
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import openai

REPO = Path(__file__).resolve().parent.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from datastore.research_fact_store import (
    database_sync_prompt,
    extract_database_updates,
    scan_research_context,
    upsert_research_facts,
)

DEFAULT_MODEL = "gpt-5.6-sol"


def load_key():
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        env = REPO / ".env"
        if env.exists():
            for line in env.read_text().splitlines():
                if line.strip().startswith("OPENAI_API_KEY") and "=" in line:
                    key = line.split("=", 1)[1].strip().strip('"').strip("'")
    if not key:
        raise SystemExit("No OPENAI_API_KEY (set in env or .env). A ChatGPT "
                         "subscription will NOT work — needs a platform API key.")
    return key


# ── briefs (format-locked; the Excel writers parse these) ───────────────────────
def brief_company(ticker, company):
    return f"""Research the core company and security facts for {company} ({ticker}) before any
financial, price-history, pipeline or valuation work. This is a DATA-LINEAGE GATE, not a narrative
company profile.

SOURCE ORDER (mandatory): SEC/issuer filing and official company IR first; exchange/regulator
second; ClinicalTrials.gov/FDA for clinical/regulatory facts; reputable market-data vendor only
for prices, beta and analyst consensus. Never use a search-result snippet as the final source.

Required checks:
1. Establish legal name, ticker, exchange, trading/reporting currencies, fiscal year-end and GAAP/IFRS.
2. Trace every ticker/name/issuer discontinuity, reverse merger, shell transaction, spin-off,
   ADR ratio change, stock split and reverse split. Determine the first trading date that belongs
   to the CURRENT operating business. Predecessor/shell price history must not enter current-company
   lifetime high/low or percentiles, even when a vendor carries it under the same security history.
3. Record the latest completed trading session date. Never mix an intraday price with a partial
   daily OHLC bar or stamp it as a completed session.
4. Extract shares outstanding from the latest authoritative filing cover page, with exact as-of
   date. Cross-check vendor shares and explain any difference; do not silently use a diluted,
   weighted-average or fully diluted count as basic shares outstanding.
5. Extract both the latest interim AND latest fiscal-year-end cash, restricted cash, current and
   non-current marketable securities, other non-current assets and debt separately. Only call the sum "unrestricted liquidity" when the issuer defines/reconciles
   it that way; never relabel cash + investments as cash.
6. For every numeric fact, preserve value, unit, as-of/data-cut date, source URL and whether it is
   Reported Fact, Company Estimate/Claim, Market Data, or Analyst Assumption. Recompute every stated
   percentage from numerator/denominator and flag any source inconsistency rather than silently
   selecting one version.

Output one fenced JSON object, no comments, using exactly these top-level keys (use null/[] when
not applicable):
legal_name, ticker, exchange, country, trading_currency, reporting_currency, fiscal_year_end,
accounting_standard, current_company_identity_start_date, price_history_start_date, identity_basis,
ir_base_url, ir_news_url,
corporate_actions, shares_outstanding, latest_reported_liquidity, prior_year_reported_liquidity, conflicts, sources.
corporate_actions rows: date, type, ratio, predecessor, accounting_acquirer, first_current_business_trading_date, source.
shares_outstanding: value, as_of, source, vendor_cross_check_value, vendor_cross_check_as_of, reconciliation.
latest_reported_liquidity: as_of, cash_and_cash_equivalents, restricted_cash, current_marketable_securities,
noncurrent_marketable_securities, debt, unrestricted_liquidity, currency, unit, source.
prior_year_reported_liquidity uses the same fields and must be the latest completed fiscal year.
"""


def brief_notes(ticker, company, years):
    cols = " ".join(map(str, years))
    return f"""Read {company} ({ticker})'s 20-F annual reports and latest interim/quarterly reports. \
Current FY DATA annual columns are the latest five fully published annual periods: {cols} \
(values in reported currency thousands unless the filing says otherwise). Extract the DETAILED \
breakdown of each note below, by year, using the company's OWN category names (not generic ones).

Output four markdown tables (rows = actual reported category names; columns = {cols}; \
values in thousands; each note's categories MUST sum to the reported total — show a Total row; \
cite note #/page/filing section for each):
1. Research & Development expenses breakdown.
2. General & Administrative expenses breakdown.
3. Property, plant & equipment — gross by asset class plus accumulated depreciation / carrying amount if disclosed.
4. Accrued expenses & other current liabilities breakdown.

Dynamic interim rule for the workflow: if a newer quarterly/half-year period exists after the latest annual FY, \
state which financial statement rows can be annualized as reported YTD * 4/quarters_covered (Q1*4, H1 average*4, \
9M average*4) and which detailed note categories cannot be decomposed from interim disclosures. Do not invent \
missing note categories or backfill by generic labels."""


def brief_pipeline(ticker, company):
    return f"""For {company} ({ticker}), list EVERY drug currently in clinical development \
(Phase 1+, active/recruiting). For each: drug code, target/mechanism, all indications, phase, \
and wholly-owned vs partnered. For any PARTNERED program, read the collaboration agreement \
(20-F / press releases) and derive {ticker}'s NET economic share of that drug's product sales \
— a single blended % of sales the company keeps (royalty tier / profit-split / who books \
revenue) plus material milestones. Output a markdown table: \
Drug | Target | Indications | Phase | Owner/Partner | {ticker} economic share (%) | key terms + source. \
State explicitly which are 100% owned. Use web search on filings. Apply the company-facts data-lineage \
rules first: primary sources beat aggregators; tag each item as Reported Fact, Company Claim/Estimate, \
Market Data, or Analyst Assumption; preserve data-cut/as-of dates; recompute every percentage from its \
numerator and denominator; and surface source conflicts instead of silently choosing a convenient value."""


def brief_ratings(ticker, company, db_context=""):
    return f"""For each {company} ({ticker}) drug × indication, assign the market-uptake TIER that \
sets its share-ramp: Best-In-Class (superior head-to-head efficacy/safety, first/best in class), \
Tier One (strong, not dominant), or Average (me-too / late / crowded). Base each on a LINE-MATCHED, \
CITED head-to-head vs the specific competitors in that indication (ORR/PFS/OS/safety/route). \
Do NOT default to Average — justify every call. Prefer our internal data center's competitor \
readouts where provided below.
{db_context}
Output a markdown table: Drug | Indication (abbrev) | Rating [Best-In-Class / Tier One / Average] | \
head-to-head justification + source | database actions. Database actions must say which competitors belong in \
Peer View only (precommercial) and which belong in both Peer View and TAM (marketed with sales)."""


def brief_events(ticker, company, start_date, end_date):
    return f"""Build the COMPLETE Historical Events source file for {company} ({ticker}) for
{start_date} through {end_date}. This is an exhaustive official-release audit, not a selection of
only material events.

MANDATORY RESEARCH METHOD:
1. Find the issuer's official IR press-release/news archive and enumerate EVERY company press
   release in the date window, including all pagination/RSS/archive pages. Reconcile the enumerated
   count to the archive and do not silently drop routine releases or multiple releases on one date.
2. Open every release page. Preserve its canonical company URL, exact publication date and title.
3. When a release discloses clinical data at ASCO/AACR/ASH/ESMO/SITC/ISSVA or any academic/major
   meeting, read BOTH the full company release and the meeting abstract/poster/report. Preserve the
   abstract URL. If the abstract is not public, state abstract_unavailable; never fabricate it.
4. A clinical-data EVT MUST begin with the disclosure venue in square brackets, e.g. [ASCO GU Poster]
   or [CALL]. Then use a very compact fixed order and exact reported numbers only:
   Phase / data type (primary, initial, interim, updated, final) / N / ORR / CR / survival / safety.
   Omit unavailable metrics; do not convert denominators into percentages unless you show the math.
5. For non-data releases, summarize compactly but keep every release. Classify only as Clinical Data,
   Partnership, Regulatory, Financing, Corporate, or Other. Primary sources control conflicts.

Output ONE fenced valid JSON object and no other JSON objects, exactly:
{{"archive_url":"...","archive_count":0,"events":[{{"date":"YYYY-MM-DD","title":"exact title",
"evt":"[VENUE] Phase ...; updated data; N=...; ORR=...; CR=...; survival=...; safety=...",
"category":"Clinical Data","url":"canonical company PR URL","source_kind":"company_press_release",
"venue":"ASCO GU Poster","abstract_url":"conference URL or empty","abstract_status":"read|unavailable"}}]}}
Chronological. Every official release in the window must appear exactly once."""


def brief_catalyst(ticker, company):
    today = datetime.now().date().isoformat()
    return f"""As of {today}, identify the single nearest FUTURE clinical catalyst for {company}
({ticker}). Search the issuer's latest press releases/presentations and recent official biomedical
conference schedules/notices. Confirm exactly what is expected to be disclosed: drug, indication,
trial/phase, analysis type, population/cohort and named endpoints/data metrics. Do not infer a
conference or exact date the company has not announced. List every Catalyst framework target that
is directly affected; unrelated targets must not be included.

Output one fenced valid JSON object only:
{{"as_of":"YYYY-MM-DD","event_date":"YYYY-MM-DD or announced window","event_name":"...",
"drug":"...","indication":"...","trial":"...","phase":"...",
"expected_disclosure":"very compact exact description","relevant_targets":["exact framework target"],
"sources":["official company PR URL","official conference notice URL if any"],
"uncertainties":["what has not been announced"]}}
At least one source must be a current official company source. Never use a search snippet as source."""


def brief_postcatalyst(ticker, company, context):
    return f"""Interpret the completed clinical catalyst below for {company} ({ticker}). Read the
full official company data release AND the academic meeting abstract/poster/report when the event
was presented at a meeting. Search the broader web for analyst/medical interpretation, but primary
data control every numeric claim. Explain the result relative to the pre-event expectation, what
was clinically strong/weak, safety, durability, subgroup/censoring limitations, and the likely
reason for the actual stock reaction. Do not confuse correlation with causation.

CATALYST CONTEXT:
{context}

Output one fenced valid JSON object only:
{{"summary":"concise but substantive interpretation","data_result":"Phase/type/N/ORR/CR/survival/safety exact figures",
"expectation_comparison":"...","price_movement_interpretation":"...","limitations":["..."],
"sources":["company PR URL","conference abstract/poster URL","other primary/reputable URL"]}}
The source list must include the company release and, if applicable, the meeting source."""


# ── move-cause research (>=8% one-day move dates) ───────────────────────────────
# Two-pass flow: fill_historical_events flags unexplained >=threshold move dates in
# artifacts/{T}/{T}_moves_needing_research.json ({"moves":[{date,dod_pct}]}); this
# brief researches each date's cause and the producer writes
# artifacts/{T}/{T}_moves_researched.json ({"moves":[{date,evt,category}]}), which
# _load_move_research (fill_historical_events) consumes, preferring a real cause over
# the generic filler. Categories MUST be one of the six the events pass understands.
MOVE_CATEGORIES = ("Clinical Data", "Partnership", "Regulatory",
                   "Financing", "Corporate", "Other")
_MOVE_CATEGORY_BY_LOWER = {c.lower(): c for c in MOVE_CATEGORIES}
_MOVE_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def brief_move_causes(ticker, company, dates):
    date_lines = "\n".join(f"- {d}" for d in dates)
    cats = " | ".join(MOVE_CATEGORIES)
    return f"""For {company} ({ticker}), find the single most likely company- or \
market-specific CAUSE of the stock's large (>=8%) one-day price move on EACH date below. \
For every date, use web search for news dated on OR the trading day before it: clinical data \
readouts, partnership/collaboration deals, regulatory actions (FDA/EMA), financings/offerings, \
earnings/guidance or other corporate events, or — only if no company news exists — a sector/market \
move. Prefer the primary source (press release, 8-K/6-K, filing, reputable news) and cite it.

Dates to explain (YYYY-MM-DD):
{date_lines}

Output BOTH of the following, covering EVERY date above exactly once:

1. A markdown table, one row per date, chronological:
| Date | EVT | Category |
| :--- | :--- | :--- |
| YYYY-MM-DD | <cause, <=15 words, with source> | <Category> |

2. A fenced ```json code block (MUST be valid JSON, no comments) with EXACTLY this shape:
```json
{{"moves": [{{"date": "YYYY-MM-DD", "evt": "<cause, <=15 words>", "category": "<Category>"}}]}}
```

Category MUST be exactly one of: {cats}. If no company-specific cause is findable, use \
"Other" and describe the sector/market move. Do not invent events — if truly nothing is found, \
say so in EVT and use "Other"."""


def _norm_move_category(value):
    return _MOVE_CATEGORY_BY_LOWER.get((value or "").strip().lower(), "Other")


def _try_load_moves(blob):
    """Return a list of move dicts from a JSON string, or None."""
    try:
        data = json.loads(blob)
    except Exception:
        return None
    if isinstance(data, dict):
        items = data.get("moves")
        if items is None:
            return [data] if "date" in data else None
        return items
    return data if isinstance(data, list) else None


def parse_move_causes(text):
    """Extract [{date, evt, category}] from a movecauses response.

    Prefers a fenced ```json {"moves":[...]} block; falls back to the
    ``| Date | EVT | Category |`` markdown table. Category is normalized to the
    six allowed labels (unknown -> "Other"); rows without a date+evt are dropped;
    de-duplicated by date (first occurrence wins). Pure function — unit-tested.
    """
    moves: list = []
    seen: set = set()

    def _add(date, evt, category):
        m = _MOVE_DATE_RE.search(str(date or ""))
        evt = (evt or "").strip()
        if not m or not evt:
            return
        d = m.group(0)
        if d in seen:
            return
        seen.add(d)
        moves.append({"date": d, "evt": evt, "category": _norm_move_category(category)})

    # 1) JSON — fenced blocks first, then any {"moves": [...]} object in the text.
    candidates: list = []
    for block in re.findall(r"```(?:json)?\s*(.*?)```", text, re.DOTALL):
        block = block.strip()
        candidates.append(block)
        m = re.search(r"(\{.*\}|\[.*\])", block, re.DOTALL)
        if m:
            candidates.append(m.group(1))
    m = re.search(r"\{[^`]*?\"moves\"\s*:\s*\[.*?\]\s*\}", text, re.DOTALL)
    if m:
        candidates.append(m.group(0))
    for blob in candidates:
        items = _try_load_moves(blob)
        if not items:
            continue
        for it in items:
            if isinstance(it, dict):
                _add(it.get("date") or it.get("Date"),
                     it.get("evt") or it.get("EVT"),
                     it.get("category") or it.get("Category"))
        if moves:
            return moves

    # 2) Markdown table fallback: | Date | EVT | Category |
    for line in text.splitlines():
        s = line.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if len(cells) < 3:
            continue
        if cells[0].lower() == "date" or set(cells[0]) <= set("-: "):
            continue
        _add(cells[0], cells[1], cells[2])
    return moves


def load_move_dates(ticker, explicit_dates=None):
    """Dates (YYYY-MM-DD) whose >=8% move needs a cause.

    Uses ``explicit_dates`` when given, else reads
    artifacts/{T}/{T}_moves_needing_research.json ({"moves":[{date,dod_pct}]}).
    Returns [] when neither is available. Pure/deterministic — unit-tested.
    """
    if explicit_dates:
        return [_MOVE_DATE_RE.search(str(d)).group(0)
                for d in explicit_dates if _MOVE_DATE_RE.search(str(d))]
    path = REPO / "artifacts" / ticker / f"{ticker}_moves_needing_research.json"
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    items = data.get("moves", data) if isinstance(data, dict) else data
    dates: list = []
    for it in items or []:
        d = it.get("date") if isinstance(it, dict) else it
        m = _MOVE_DATE_RE.search(str(d or ""))
        if m and m.group(0) not in dates:
            dates.append(m.group(0))
    return dates


def write_moves_researched(ticker, moves):
    """Write artifacts/{T}/{T}_moves_researched.json (consumed by the events pass)."""
    out_dir = REPO / "artifacts" / ticker
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{ticker}_moves_researched.json"
    out_path.write_text(json.dumps({"moves": moves}, indent=2), encoding="utf-8")
    return out_path


BRIEFS = {"company": brief_company, "notes": brief_notes, "pipeline": brief_pipeline,
          "ratings": brief_ratings, "events": brief_events,
          "catalyst": brief_catalyst, "postcatalyst": brief_postcatalyst,
          "movecauses": brief_move_causes}
BUILD_BRIEFS = ["company", "notes", "pipeline", "ratings", "events", "movecauses"]


def parse_company_facts(text):
    """Parse the strict company-facts JSON block returned by the research brief."""
    candidates = re.findall(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)
    candidates.append(text)
    for blob in candidates:
        m = re.search(r"\{.*\}", blob, re.DOTALL)
        if not m:
            continue
        try:
            data = json.loads(m.group(0))
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("ticker"):
            return data
    return None


def parse_json_object(text):
    candidates = re.findall(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)
    candidates.append(text)
    for blob in candidates:
        match = re.search(r"\{.*\}", blob, re.DOTALL)
        if not match:
            continue
        try:
            data = json.loads(match.group(0))
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict):
            return data
    return None


def write_company_facts(ticker, data):
    out_dir = REPO / "artifacts" / ticker.upper()
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{ticker.upper()}_company_facts.json"
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def parse_official_events(text):
    """Parse the strict exhaustive-events JSON block returned by brief_events."""
    candidates = re.findall(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)
    candidates.append(text)
    for blob in candidates:
        match = re.search(r"\{.*\}", blob, re.DOTALL)
        if not match:
            continue
        try:
            data = json.loads(match.group(0))
        except json.JSONDecodeError:
            continue
        if not isinstance(data, dict) or not isinstance(data.get("events"), list):
            continue
        rows = []
        for event in data["events"]:
            if not isinstance(event, dict):
                continue
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(event.get("date", ""))):
                continue
            if not event.get("evt") or not event.get("url"):
                continue
            rows.append(event)
        data["events"] = rows
        return data
    return None


def write_official_events(ticker, data, output_dir=None):
    artifact_dir = REPO / "artifacts" / ticker.upper()
    artifact_dir.mkdir(parents=True, exist_ok=True)
    artifact_path = artifact_dir / f"{ticker.upper()}_official_events.json"
    payload = json.dumps(data, indent=2, ensure_ascii=False)
    artifact_path.write_text(payload, encoding="utf-8")
    output_path = None
    if output_dir:
        output_path = Path(output_dir) / f"{ticker.upper()}_events.json"
        output_path.write_text(payload, encoding="utf-8")
    return artifact_path, output_path


def latest_completed_years(count: int = 5):
    last_full_year = datetime.now().year - 1
    return list(range(last_full_year - count + 1, last_full_year + 1))


def _infer_rating_keywords(report_dir: Path, ticker: str):
    keywords = {ticker}
    if not report_dir.exists():
        return keywords
    # Build the indication vocabulary per-ticker from each report's own section
    # headers (e.g. "3.1 <indication> (ABBR)"). The parenthetical abbreviation is
    # what the datastore's section / indication_code columns key on, so capturing
    # it generalizes to any ticker's diseases — no hardcoded reference-indication list.
    for path in report_dir.glob(f"{ticker}_*_research_*.md"):
        text = path.read_text(encoding="utf-8", errors="ignore")
        for header in re.findall(r"##+\s*3\.\d+\s+([^\n]+)", text, re.I):
            for token in re.findall(r"\(([A-Za-z][A-Za-z0-9/+ -]+)\)", header):
                keywords.add(token.strip())
    return keywords


def load_db_context(ticker: str, report_dir: Path) -> str:
    """Load a compact Peer View/TAM data-center excerpt for GPT ratings."""
    keywords = _infer_rating_keywords(report_dir, ticker)
    kw_lower = {k.lower() for k in keywords}
    export = REPO / "datastore" / "export"

    lines = ["Internal data center excerpts to consult before web sources:"]
    peer_rating = export / "peer_rating.csv"
    if peer_rating.exists():
        rows = []
        with peer_rating.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                hay = " ".join(row.get(k, "") for k in ("section", "drug", "ticker")).lower()
                if any(k in hay for k in kw_lower):
                    rows.append(row)
        if rows:
            lines.append("peer_rating.csv matching rows:")
            for row in rows[:120]:
                lines.append(
                    f"- {row.get('section')} | {row.get('drug')} | "
                    f"{row.get('ticker')} | {row.get('rating')}"
                )

    tam = export / "tam_by_indication_year.csv"
    if tam.exists():
        tam_rows = []
        with tam.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                code = (row.get("indication_code") or "").lower()
                if any(k in code for k in kw_lower):
                    tam_rows.append(row)
        if tam_rows:
            lines.append("tam_by_indication_year.csv matching rows:")
            for row in tam_rows[:80]:
                lines.append(
                    f"- {row.get('indication_code')} {row.get('year')}: "
                    f"{row.get('tam_usd_m')} USDm"
                )

    if len(lines) == 1:
        lines.append("No matching rows found; explicitly state needed database additions.")
    return "\n".join(lines)


def call_gpt(client, model, prompt, label=""):
    t0 = time.time()
    # Responses API with web search (works for gpt-5.x reasoning/pro models).
    resp = client.responses.create(
        model=model, input=prompt, tools=[{"type": "web_search"}],
    )
    text = getattr(resp, "output_text", None) or ""
    print(f"  {label}: {len(text):,} chars, {int(time.time()-t0)}s")
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ticker")
    ap.add_argument("--company-name")
    ap.add_argument("--brief", default="all",
                    help="company | notes | pipeline | ratings | events | movecauses | all")
    ap.add_argument("--years", nargs="+", type=int)
    ap.add_argument("--dates", nargs="+",
                    help="movecauses: >=8%% move dates (YYYY-MM-DD) to explain; "
                         "default reads artifacts/{T}/{T}_moves_needing_research.json")
    ap.add_argument("--model", default=DEFAULT_MODEL)
    ap.add_argument("--output-dir")
    ap.add_argument("--report-dir",
                    help="Pipeline report directory used to infer rating DB keywords")
    ap.add_argument("--context-file",
                    help="postcatalyst: active-state/research JSON or text context")
    ap.add_argument("--list-models", action="store_true")
    a = ap.parse_args()

    os.environ["OPENAI_API_KEY"] = load_key()
    client = openai.OpenAI()

    if a.list_models:
        for m in client.models.list().data:
            print(" ", m.id)
        return

    years = a.years or latest_completed_years(5)
    out = Path(a.output_dir) if a.output_dir else \
        Path(f"/mnt/c/Users/yzsun/Desktop/DD/{a.ticker}/gpt_research")
    out.mkdir(parents=True, exist_ok=True)
    report_dir = Path(a.report_dir) if a.report_dir else \
        Path(f"/mnt/c/Users/yzsun/Desktop/DD/{a.ticker}/pipeline_base4")
    end_date = datetime.now().date()
    # Match the four calendar-year Historical Events blocks exactly; a rolling
    # 1,095-day cut would silently omit January-to-current-date of the first year.
    start_date = end_date.replace(year=end_date.year - 3, month=1, day=1)

    todo = BUILD_BRIEFS if a.brief == "all" else [a.brief]
    for name in todo:
        fn = BRIEFS[name]
        if name == "notes":
            prompt = fn(a.ticker, a.company_name, years)
        elif name == "events":
            prompt = fn(a.ticker, a.company_name, start_date.isoformat(), end_date.isoformat())
        elif name == "ratings":
            prompt = fn(a.ticker, a.company_name, load_db_context(a.ticker, report_dir))
        elif name == "movecauses":
            dates = load_move_dates(a.ticker, a.dates)
            if not dates:
                print(f"▶ {name}\n  skipped: no dates (pass --dates or create "
                      f"artifacts/{a.ticker}/{a.ticker}_moves_needing_research.json)")
                continue
            prompt = fn(a.ticker, a.company_name, dates)
        elif name == "postcatalyst":
            if not a.context_file:
                raise SystemExit("postcatalyst requires --context-file")
            context = Path(a.context_file).read_text(encoding="utf-8", errors="ignore")
            prompt = fn(a.ticker, a.company_name, context)
        else:
            prompt = fn(a.ticker, a.company_name)
        # Mandatory parallel database branch for every research brief.  The
        # clinical/commercial fact ledger is scanned before web research and
        # any newly sourced rows are persisted immediately after the response.
        db_scan = scan_research_context(a.ticker)
        prompt += database_sync_prompt(db_scan)
        print(f"▶ {name}")
        try:
            text = call_gpt(client, a.model, prompt, name)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            p = out / f"{a.ticker}_{name}_{ts}.md"
            p.write_text(text, encoding="utf-8")
            print(f"  saved {p}")
            database_updates = extract_database_updates(text)
            if database_updates:
                stored = upsert_research_facts(database_updates, a.ticker)
                refreshed = scan_research_context(a.ticker)
                incomplete = [
                    competitor for competitor, status in
                    refreshed["competitor_completeness"].items()
                    if not status["complete"]
                ]
                print(
                    f"  database scan: stored {len(stored)} new/corrected facts; "
                    f"incomplete competitors={len(incomplete)}"
                )
            else:
                print("  database scan: active; no new structured facts returned")
            if name == "company":
                facts = parse_company_facts(text)
                if facts:
                    jp = write_company_facts(a.ticker, facts)
                    print(f"  wrote researched company facts → {jp}")
                else:
                    print("  WARNING: no parseable company-facts JSON; market-data lifetime "
                          "statistics will require manual identity review")
            elif name == "events":
                event_data = parse_official_events(text)
                if event_data:
                    apath, opath = write_official_events(a.ticker, event_data, out)
                    print(f"  wrote {len(event_data['events'])} official events → {apath}")
                    if opath:
                        print(f"  build-compatible events file → {opath}")
                else:
                    print("  WARNING: no parseable exhaustive-events JSON; official archive "
                          "crawler must supply the Historical Events input")
            elif name in {"catalyst", "postcatalyst"}:
                payload = parse_json_object(text)
                if payload:
                    suffix = "catalyst_run_research" if name == "catalyst" else "post_catalyst_interpretation"
                    artifact = REPO / "artifacts" / a.ticker.upper() / f"{a.ticker.upper()}_{suffix}.json"
                    artifact.parent.mkdir(parents=True, exist_ok=True)
                    artifact.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
                    print(f"  wrote workflow JSON → {artifact}")
                else:
                    print(f"  WARNING: no parseable {name} JSON")
            elif name == "movecauses":
                moves = parse_move_causes(text)
                if moves:
                    jp = write_moves_researched(a.ticker, moves)
                    print(f"  wrote {len(moves)} researched moves → {jp}")
                else:
                    print("  WARNING: no parseable moves in movecauses output; "
                          "moves_researched.json not written")
        except Exception as e:
            print(f"  FAILED {name}: {e}")


if __name__ == "__main__":
    main()
