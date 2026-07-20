#!/usr/bin/env python3
"""
opus_research.py — per-drug pipeline research via Claude Opus 4.8 + web search.

Drop-in replacement for research/gemini_research.py: it produces the SAME
markdown report format the downstream parsers consume
(generate_scenarios.py / generate_pipeline.py / fill_peer_views.py), but the
research is done by Claude Opus 4.8 with the Anthropic server-side `web_search`
tool instead of Google Gemini Deep Research.

Two artefacts per drug, written to DD/{TICKER}/pipeline_base4/:
    {TICKER}_{Drug}_research_{ts}.md   — Ch1 overview, Ch3 per-indication
                                         (TAM + market-share table), Ch4 stage
                                         timeline, Ch5 PEER_VIEW blocks
    {TICKER}_{Drug}_pricing_{ts}.md    — per-indication treatment pricing

Market share: by default (`--market-share real`) Opus researches per-indication
TAM ($MM) and BASE/BULL/BEAR peak share, grounded in cited comparables. Pass
`--market-share flat` to instead write a FLAT peak share (`--flat-ms`, default
10%) into every table (0% pre-launch → ramp to peak → held to 2038).

Usage:
    python research/opus_research.py --ticker MOLN \
        --company-name "Molecular Partners AG" \
        --drugs MP0533 MP0712 \
        --trials-json ".../MOLN_clinical_trials_*.json"

Key: reads ANTHROPIC_API_KEY, else MY_PYTHON_SCRIPT_KEY (from env or repo .env).
"""
import argparse
import glob
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import anthropic

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

MODEL = "claude-opus-4-8"
WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search", "max_uses": 20}
REPO = Path(__file__).resolve().parent.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from datastore.research_fact_store import (
    database_sync_prompt,
    extract_database_updates,
    scan_research_context,
    upsert_research_facts,
)


# ── key loading ────────────────────────────────────────────────────────────────
def load_api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("MY_PYTHON_SCRIPT_KEY")
    if not key:
        env = REPO / ".env"
        if env.exists():
            for line in env.read_text().splitlines():
                line = line.strip()
                if line.startswith(("ANTHROPIC_API_KEY", "MY_PYTHON_SCRIPT_KEY")) and "=" in line:
                    key = line.split("=", 1)[1].strip().strip('"').strip("'")
                    if line.startswith("ANTHROPIC_API_KEY"):
                        break
    if not key:
        raise SystemExit("No ANTHROPIC_API_KEY / MY_PYTHON_SCRIPT_KEY found (env or .env).")
    return key


# ── clinical-trial context for a single drug ────────────────────────────────────
def format_drug_trials(drug_name: str, trial_data: Optional[Dict]) -> str:
    if not trial_data or not trial_data.get("trials"):
        return f"No pre-fetched trial data for {drug_name}."
    key = drug_name.lower().replace("[", "").replace("]", "")
    lines, found = [f"### ClinicalTrials.gov records matching {drug_name}\n"], False
    for nct, t in trial_data["trials"].items():
        ivs = " | ".join(t.get("interventions", []))
        if key.split()[0] in ivs.lower() or drug_name.lower() in ivs.lower():
            found = True
            lines += [
                f"**{nct}** — {t.get('title', 'N/A')}",
                f"- Phase: {t.get('phase', '?')}  Status: {t.get('status', '?')}",
                f"- Conditions: {', '.join(t.get('conditions', []))}",
                f"- Interventions: {ivs}",
                f"- Start: {t.get('start_date', 'N/A')}  Completion: {t.get('completion_date', 'N/A')}\n",
            ]
    if not found:
        lines.append("No exact trial match; search ClinicalTrials.gov + company filings.")
    return "\n".join(lines)


def load_tam_datastore_context() -> str:
    """Compact TAM export excerpt for the research prompt."""
    path = REPO / "datastore" / "export" / "tam_by_indication_year.csv"
    if not path.exists():
        return "Internal TAM datastore export not found; every TAM estimate must state database actions."
    import csv

    by_code: Dict[str, Dict[int, float]] = {}
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            try:
                year = int(row.get("year", ""))
                # Anchor years are template-locked to the fixed TAM Solid/Blood grid
                # (cols F=2010 … AH=2038; datastore/tam_overrides MODEL_END_YEAR=2038),
                # not build-date-derived — 2024/2030/2038 = near/mid/terminal samples.
                if year not in {2024, 2030, 2038}:
                    continue
                by_code.setdefault(row["indication_code"], {})[year] = float(row["tam_usd_m"])
            except (KeyError, TypeError, ValueError):
                continue
    lines = ["Internal TAM datastore current anchors (USD MM; exact-code match only):"]
    for code in sorted(by_code):
        vals = by_code[code]
        anchors = ", ".join(f"{y}={vals[y]:.0f}" for y in (2024, 2030, 2038) if y in vals)
        if anchors:
            lines.append(f"- {code}: {anchors}")
    return "\n".join(lines[:180])


# ── prompt builders (STRICT parser-aligned format) ──────────────────────────────
def _ms_section_flat(peak):
    return f"""`#### 3.N.1 TAM`
- Addressable patients (US/global) for the specific line/biomarker, 2024 → 2030 → 2038,
  with cited sources.

`#### 3.N.7 Market Share Projection — BASE Case (ABBR)`
Then a table with these exact first two columns (whole-number percent + literal `%`):
| Year | Market Share | Reasoning |
| :--- | :--- | :--- |
| 2024 | 0% | Pre-approval |
| <launch year> | 0% | Launch |
| <launch year + 4> | {peak}% | Ramp to peak |
| 2038 | {peak}% | Held |
IMPORTANT: use a FLAT peak market share of {peak}% for every indication this pass — do
NOT research or vary market share. Choose <launch year> from your Chapter 4 stage-5 year."""


def _ms_section_real():
    return """`#### 3.N.1 TAM` — market SIZE, researched and cited. Estimate the addressable
market in **US$ millions** for this drug's specific line/biomarker (incidence × treated
fraction × annual treatment cost). FIRST consult the internal TAM datastore excerpt below.
If the exact indication code exists and matches this drug's line/biomarker/serviceable
market, use it as the starting point; if it is absent, stale, or mismatched, search all
marketed drugs and epidemiology sources, derive new 2024/2030/2038 anchors, and state a
clear `Database TAM action:` sentence saying whether to add or override that exact code.
Emit BOTH the reasoning AND this exact table
(header text must contain "Addressable Market ($MM)" — software reads it):
| Year | Addressable Market ($MM) |
| :--- | :--- |
| 2024 | <$MM> |
| 2030 | <$MM> |
| 2038 | <$MM> |

`#### 3.N.3 Competitive Landscape` — EXHAUSTIVE, with ONE dated row per readout/dataset —
NEVER a single summary row per drug. If a competitor reported multiple datasets (e.g.
Tarlatamab DeLLphi-301 AND DeLLphi-304), give EACH its own dated row. Marketed drugs (name,
company, data date, dataset/study or NCT#, latest ANNUAL SALES $MM, ORR/PFS/OS, patent-expiry
year) AND clinical-stage drugs (name, company, data date, dataset/study or NCT#, phase, est.
approval year, efficacy). Cite each.

`#### 3.N.5 Differentiation` — quantified, line-matched comparison of {drug} vs each key
competitor (ORR/PFS/OS/safety/route). Conclude with a line `Assessment: <label>` using EXACTLY
one accepted tier label — Best-in-Class, Tier One (a.k.a. Above-average), Average, or
Below-average (Best-in-Class→BIC, Tier One/Above-average→T1, Average/Below-average→AVG) —
plus a one-sentence justification.

`#### 3.N.7 Market Share Projection — BASE Case (ABBR)` — a peak-share and uptake curve
JUSTIFIED by comparable drugs (name specific analog drugs that reached X% with a similar
profile) and adjusted for competitive entry / patent cliffs. Table (whole-number % + literal
`%`; only the first two columns are read):
| Year | Market Share | Reasoning |
| :--- | :--- | :--- |
| 2024 | 0% | Pre-approval |
| <launch year> | 0% | Launch |
| ... intermediate years ramping ... |
| <peak year> | <peak>% | Peak — [comparable analog achieved Y%] |
| 2038 | <%> | [erosion / durability] |

`#### 3.N.8 Market Share Projection — BULL Case (ABBR)` — higher peak / faster uptake
(best-in-class data, faster approval, weaker competition). Same table format.

`#### 3.N.9 Market Share Projection — BEAR Case (ABBR)` — lower peak / slower uptake
(competitive pressure, safety, reimbursement). Same table format.

Market share is the fraction of the indication TAM this drug captures. Ground every number
in cited comparables — no round-number guesses."""


def research_prompt(ticker, company, drug, trials_text, flat_ms=None, market_share="flat"):
    ms_section = _ms_section_real() if market_share == "real" \
        else _ms_section_flat(round((flat_ms or 0.10) * 100))
    tam_context = load_tam_datastore_context() if market_share == "real" else ""
    prompt = f"""You are a senior biotech equity-research analyst. Write a data-driven research \
report on **{drug}** from **{company}** (ticker {ticker}). Use the `web_search` tool \
extensively to ground every indication, TAM figure, competitor readout, launch-timing \
comparable, and price in real, cited sources (ClinicalTrials.gov, company 10-K/20-F/press \
releases, ASCO/ASH/ESMO, SEER/epidemiology, IQVIA, analyst notes).

Before analysis, apply a strict data-lineage gate: primary filings/issuer releases and original
FDA/ClinicalTrials.gov/conference sources outrank aggregators. Give every numeric fact an as-of or
data-cut date and distinguish Reported Fact, Company Estimate/Claim, Market Data, and Analyst
Assumption. Recompute all rates from numerator/denominator. If sources conflict, display the
conflict instead of silently choosing one. For safety, preserve Grade 3+, serious events,
discontinuations and denominators; never turn "generally mild/moderate" into "all mild/moderate."

## Clinical-trial context
{trials_text}

## Internal data-center context
{tam_context}

## OUTPUT — follow this markdown structure EXACTLY (downstream software parses it; \
any deviation from the headers/markers breaks it)

### Chapter 1: Drug Overview
- 3-6 sentences: modality, mechanism, current highest phase.
- A line exactly: `**Target:** <targets/mechanism>`
- A line exactly: `Economic share: NN%` — integer percent of net product sales {ticker} retains
  (`100%` if wholly owned; for a partnered program derive the blended net share from the
  collaboration terms in the 10-K/20-F or deal press releases).

### Chapter 2: Forecasting Strategy
- 2-4 sentences on which indications get a separate forecast.

### Chapter 3: Per-Indication Analysis
For EACH clinically-relevant indication, emit a section headed exactly:
`### 3.N <Indication Name> (ABBR)`   (N = 1,2,3…; ABBR is the abbreviation in parentheses,
e.g. `### 3.1 Acute Myeloid Leukemia (AML)`). Under each:

{ms_section}

### Chapter 4: Stage Timeline
A table (year may be a range; earliest is used; stage 5 = Approval turns revenue on):
| Stage | Year | Source / Reasoning |
| :--- | :--- | :--- |
| 1 (Phase I) | <year> | ... |
| 2 (Phase II) | <year> | ... |
| 3 (Phase III) | <year> | ... |
| 4 (BLA/MAA Filing) | <year> | ... |
| 5 (Approval) | <year> | comparable-based estimate |

### Chapter 5: Peer View Data Collection
For EACH indication above, emit a block delimited EXACTLY (END label byte-identical to START):
`#### PEER_VIEW_START: <Indication Name> (ABBR)`
… readouts …
`#### PEER_VIEW_END: <Indication Name> (ABBR)`

Inside, one readout per DATED data-presentation event — a drug with multiple datasets (e.g.
Tarlatamab DeLLphi-301 AND DeLLphi-304) gets a SEPARATE readout block for EACH; never collapse
them into one summary block. Head each exactly:
`##### Drug: <name> — Readout <n>`
followed by `- Key: Value` lines using these exact keys (percentages as DECIMALS e.g.
`0.37`; use `/` for anything unavailable; dates as YYYY-MM-DD):
- Drug Name: / Company: / Ticker: / Innovation: / Rating: / Target: / Result: / NCT#: /
  Treatment Line: / Phase: / Stage: / Data Date: / Conference: / N: / ORR: / CR: / PR: /
  DCR: / Median PFS: / Median OS: / GEQ G3 SAE Pct: / Route: / Latest Annual Sale: / Source:
`Rating:` is the drug's competitiveness tier in THIS indication — EXACTLY one of `BIC`
(best-in-class), `T1` (tier one / above-average), or `AVG` (average / below-average); include it
for EACH competitor and for {drug} itself.
List **{drug} FIRST** in each block (use `/` where it has no data yet), then the top
marketed competitors (by sales), then key clinical-stage competitors.

Begin now. Be exhaustive with web_search; cite sources inline."""
    return prompt + database_sync_prompt(scan_research_context(ticker))


def pricing_prompt(ticker, company, drug):
    prompt = f"""You are a pharmaceutical pricing analyst. For **{drug}** ({company}, {ticker}), \
research treatment pricing for EACH of its indications using `web_search` (WAC/ASP of \
comparable marketed drugs in the same class + indication).

Output one section per indication, headed exactly `### <Indication Name> (ABBR)` (same
abbreviation used elsewhere, e.g. `### Acute Myeloid Leukemia (AML)`), each containing a
table with this EXACT row (value in millions USD per patient; a single figure or `$lo - $hi`):
| Parameter | Value | Source |
| :--- | :--- | :--- |
| Estimated Price Per Dose ($) | ... | ... |
| Dosing Schedule | ... | ... |
| Treatment Duration (months) | ... | ... |
| Total Treatment Cost Per Patient (MM USD) | $X.XX | ... |
| Route | IV/SC/Oral | ... |

Then a `### Comparable Pricing Table` listing the drugs used. Cite all sources."""
    return prompt + database_sync_prompt(scan_research_context(ticker))


# ── Opus call ───────────────────────────────────────────────────────────────────
def call_opus(client, prompt, max_tokens=12000, label=""):
    t0 = time.time()
    parts, searches = [], 0
    resp = client.messages.create(
        model=MODEL, max_tokens=max_tokens, tools=[WEB_SEARCH_TOOL],
        messages=[{"role": "user", "content": prompt}],
    )
    for b in resp.content:
        if getattr(b, "type", "") == "text":
            parts.append(b.text)
        elif getattr(b, "type", "") == "server_tool_use":
            searches += 1
    text = "".join(parts)
    dt = int(time.time() - t0)
    trunc = " [TRUNCATED: max_tokens]" if resp.stop_reason == "max_tokens" else ""
    logger.info(f"  {label}: {len(text):,} chars, {searches} searches, {dt}s{trunc}")
    return text


# ── save ────────────────────────────────────────────────────────────────────────
def save_md(text, out_dir: Path, ticker, drug, kind):
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^\w\-]", "_", drug)
    path = out_dir / f"{ticker}_{safe}_{kind}_{ts}.md"
    path.write_text(text, encoding="utf-8")
    logger.info(f"  saved {path}")
    return path


def load_trials(trials_json: Optional[str], ticker: str) -> Optional[Dict]:
    if trials_json and "*" in trials_json:
        hits = sorted(glob.glob(trials_json))
        trials_json = hits[-1] if hits else None
    if not trials_json:
        for base in (f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}",
                     str(REPO / f"C:/Users/yzsun/Desktop/DD/{ticker}")):
            hits = sorted(glob.glob(f"{base}/*clinical_trials*.json"))
            if hits:
                trials_json = hits[-1]
                break
    if trials_json and os.path.exists(trials_json):
        logger.info(f"Loaded trials: {trials_json}")
        return json.load(open(trials_json))
    logger.warning("No trials JSON found; proceeding without trial context.")
    return None


def main():
    ap = argparse.ArgumentParser(description="Per-drug pipeline research via Claude Opus 4.8 + web search")
    ap.add_argument("--ticker", required=True)
    ap.add_argument("--company-name", required=True)
    ap.add_argument("--drugs", nargs="+", required=True,
                    help="Drug names to research (active clinical-stage assets)")
    ap.add_argument("--trials-json", help="Path (glob ok) to clinical-trials JSON")
    ap.add_argument("--output-dir")
    ap.add_argument("--flat-ms", type=float, default=0.10,
                    help="Flat peak market share for every table (default 0.10)")
    ap.add_argument("--market-share", choices=["flat", "real"], default="real",
                    help="'real' (default) = Opus researches TAM ($MM) + BASE/BULL/BEAR share "
                         "per indication; 'flat' = every table uses --flat-ms")
    ap.add_argument("--no-pricing", action="store_true")
    args = ap.parse_args()

    os.environ["ANTHROPIC_API_KEY"] = load_api_key()
    client = anthropic.Anthropic()
    out_dir = Path(args.output_dir) if args.output_dir else \
        Path(f"/mnt/c/Users/yzsun/Desktop/DD/{args.ticker}/pipeline_base4")
    trials = load_trials(args.trials_json, args.ticker)

    results = []
    for i, drug in enumerate(args.drugs, 1):
        logger.info(f"\n{'─'*60}\n[{i}/{len(args.drugs)}] {drug}\n{'─'*60}")
        tctx = format_drug_trials(drug, trials)
        try:
            rp = research_prompt(args.ticker, args.company_name, drug, tctx,
                                 args.flat_ms, args.market_share)
            rtext = call_opus(client, rp, max_tokens=24000, label=f"{drug} research")
            stored = upsert_research_facts(
                extract_database_updates(rtext), args.ticker
            )
            logger.info("  Database scan active: stored/refreshed %d facts", len(stored))
            save_md(rtext, out_dir, args.ticker, drug, "research")
            if not args.no_pricing:
                pt = pricing_prompt(args.ticker, args.company_name, drug)
                ptext = call_opus(client, pt, max_tokens=6000, label=f"{drug} pricing")
                stored = upsert_research_facts(
                    extract_database_updates(ptext), args.ticker
                )
                logger.info("  Database scan active: stored/refreshed %d facts", len(stored))
                save_md(ptext, out_dir, args.ticker, drug, "pricing")
            results.append((drug, "OK", len(rtext)))
        except Exception as e:
            logger.error(f"  FAILED {drug}: {e}")
            import traceback; traceback.print_exc()
            results.append((drug, "FAIL", str(e)[:80]))

    print(f"\n{'='*60}\nOpus research complete — {args.ticker}\n{'='*60}")
    for d, st, info in results:
        print(f"  {st:4} {d:28} {info}")


if __name__ == "__main__":
    main()
