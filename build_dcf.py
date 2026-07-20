#!/usr/bin/env python3
"""
build_dcf.py — one-command DCF builder: ticker in → complete valuation model out.

Chains the whole pipeline, seeding the workbook by COPYING the master template
(so all 20 tabs + every pre-wired valuation formula are present and match the
template cell-for-cell) and then filling each input tab. GPT-5.5 Pro owns the
research/judgment layer; this script owns workbook assembly and verification.

    python build_dcf.py --ticker MOLN \
        --company-name "Molecular Partners AG" \
        --drugs MP0533 MP0712 MP0317

Steps (skip any with --skip, resume with --from):
    1 seed        cp template → DD/{TICKER}/DCF {TICKER}.xlsx  (backup first)
    2 financials  main.py            → FY DATA / FY DATA K USD (dynamic latest years)
    3 trials      clinical_trials_fetcher.py → NCT/indication JSON
    4 gptresearch gpt_research.py    → company facts + notes/pipeline/economics/ratings/events briefs
    5 research    gemini_research.py/opus_research.py → parser-compatible per-drug reports
    5a pricing    gemini_research.py --chapter pricing → per-patient list price
    5b peerdatabase update_peer_database_from_reports.py → new peer readouts into datastore
    5c judgement  model_judgement.py → reviewed assumptions draft/approved JSON
    6 scenarios   generate_scenarios.py → Scenarios incl. drug×indication breakdown
    7 pipeline    generate_pipeline.py  → Pipeline revenue incl. ratings/economics
    7 peerviews   fill_peer_views.py    → raw Peer Views readouts
    7a datastore  fill_peer_views_from_datastore.py → restore data-center peer facts
    7b peerviewsummary build_peer_view_summary.py → current ticker drug-vs-peer page
   8 events      fill_historical_events.py → official news + prices/DoD/EVT/category
   8a movecauses gpt_research.py --brief movecauses → causes for unexplained >=8%
                  moves, then re-runs events so the workbook consumes them
   9 financials2 generate_financials.py→ RBS / RCFS (feeds the DCF)
  10 templateclean clean_workbook_template.py → remove stale template labels
   11 formats     normalize_workbook_formats.py → generated Pipeline/Scenarios styles
   12 fsalinks    fix_fsa_links.py → repair FSA/RIS/Schedules formula links
   12a statementlogic fix_statement_logic.py → RIS opex/RBS equity roll-forward
   13 excelrepair repair_excel_package.py → normalize Excel-hostile OOXML extras
   14a finalpolish final_excel_polish.py → Excel-native final formats/comments
   14b displayformats sanitize_excel_display_formats.py → zero values show as 0, not dash lines
   14c trimtabs    trim_model_tabs.py → delete build-time TAM/Peer Views/Welcome tabs
   14d namespaces  repair_excel_namespaces.py → restore Excel-compatible OOXML prefixes
   14e referencestyles apply approved styles (Pipeline uses locked 20260709 source)
   14f valuationcharts build_valuation_charts.py → Waterfall + Football charts
   14g excelnormalize Excel COM repair-save → final Excel-native package normalization
   14h calcstate   normalize_calc_state.py → automatic/non-stale Excel calculation state
   15 audit       workbook_audit.py + recalc_check.py

Bloomberg add-in artifacts are frozen to static public-data values. Market inputs use one
completed-session yfinance snapshot, while researched company facts define issuer-history
boundaries and authoritative shares/liquidity inputs.
"""
import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

REPO = Path(__file__).resolve().parent
PY = sys.executable
TEMPLATE = Path("/mnt/c/Users/yzsun/Desktop/DD/base/DCF Template 2020.xlsx")

# Per-ticker investor-relations news pages for req5's official-news scrape.
# Keyed by upper-case ticker. An explicit entry always wins over auto-discovery
# (below); leave it empty for tickers whose IR news page resolve_news_url can
# find on its own. The generalized scraper (tools/fill_historical_events.py) no
# longer needs a Q4/GlobeNewswire-shaped listing — any same-site article listing
# whose detail pages carry parseable dates works.
NEWS_URL_BY_TICKER: dict[str, str] = {}

# Common investor-relations news / press-release paths probed during
# auto-discovery when no explicit --news-url or NEWS_URL_BY_TICKER entry exists.
IR_NEWS_PATHS = [
    "/news",
    "/news-releases",
    "/press-releases",
    "/investors/news",
    "/investors/news-events/press-releases",
    "/investors/press-releases",
    "/investor-relations/news",
    "/investors/news-events/news-releases",
    "/media/press-releases",
    "/news-media",
]

_ARTICLE_HINTS = ("news", "press", "release", "announce", "media")


def _looks_article_like(href: str, base_host: str) -> bool:
    """Whether a resolved href plausibly points at a single news/press article.

    Same host as the listing, an article-ish path segment, and a slug-like tail
    below it (so nav/index links like ``/news`` are ignored). The Q4/GlobeNewswire
    detail path is accepted explicitly. Pure/deterministic — no I/O."""
    from urllib.parse import urlparse
    try:
        parsed = urlparse(href)
    except Exception:
        return False
    host = (parsed.netloc or "").lower()
    if host and base_host and host != base_host:
        return False
    path = (parsed.path or "").rstrip("/").lower()
    if not path:
        return False
    if "/news-releases/news-release-details/" in path:
        return True
    segments = [s for s in path.split("/") if s]
    if len(segments) < 2:
        return False
    return any(hint in path for hint in _ARTICLE_HINTS)


def _autodiscover_news_url(ticker: str) -> str | None:
    """Best-effort discovery of a company's IR news-listing URL.

    Resolves the corporate website via yfinance, then probes common IR/news
    paths on that host (and an ``ir.<domain>`` host), returning the first page
    that fetches and lists multiple article-like links. Returns None — never
    raises — when yfinance/requests are unavailable, no website is known, or
    nothing probes clean, so the caller cleanly falls back to no-scrape."""
    try:
        import yfinance as yf
    except Exception:
        return None
    try:
        info = yf.Ticker(ticker).info or {}
        website = info.get("website") or ""
    except Exception:
        return None
    if not website:
        return None
    try:
        import requests
        from bs4 import BeautifulSoup
        from urllib.parse import urljoin, urlparse
    except Exception:
        return None
    parsed = urlparse(website if "://" in website else "https://" + website)
    host = (parsed.netloc or parsed.path or "").lower().strip("/")
    if not host:
        return None
    bare = host[4:] if host.startswith("www.") else host
    hosts = [host]
    ir_host = f"ir.{bare}"
    if ir_host not in hosts:
        hosts.append(ir_host)
    session = requests.Session()
    headers = {"User-Agent": "DeepDiligence news discovery/1.0"}
    for h in hosts:
        for path in IR_NEWS_PATHS:
            url = f"https://{h}{path}"
            try:
                resp = session.get(url, headers=headers, timeout=15, allow_redirects=True)
            except Exception:
                continue
            if resp.status_code >= 400:
                continue
            try:
                soup = BeautifulSoup(resp.text, "html.parser")
            except Exception:
                continue
            final_url = str(getattr(resp, "url", "") or url)
            final_host = (urlparse(final_url).netloc or h).lower()
            seen: set[str] = set()
            for a_tag in soup.find_all("a", href=True):
                resolved = urljoin(final_url, a_tag["href"]).split("#", 1)[0].rstrip("/")
                if resolved in seen or not _looks_article_like(resolved, final_host):
                    continue
                seen.add(resolved)
                if len(seen) >= 3:
                    print(f"  auto-discovered IR news URL for {ticker}: {final_url}")
                    return final_url
    return None


def resolve_news_url(ticker: str, cli_news_url: str | None) -> str | None:
    """Resolve the IR news-listing URL for req5's official-news scrape.

    Priority: explicit --news-url, then the per-ticker config map, then a
    best-effort yfinance-website auto-discovery probe. Returns None when nothing
    resolves (never fabricates a URL). Signature is unchanged for callers."""
    if cli_news_url:
        return cli_news_url
    facts_path = REPO / "artifacts" / (ticker or "").upper() / f"{(ticker or '').upper()}_company_facts.json"
    if facts_path.exists():
        try:
            facts = json.loads(facts_path.read_text(encoding="utf-8"))
            researched = facts.get("ir_news_url")
            if researched:
                return str(researched)
        except Exception:
            pass
    mapped = NEWS_URL_BY_TICKER.get((ticker or "").upper())
    if mapped:
        return mapped
    return _autodiscover_news_url(ticker)


def dd_dir(ticker):
    return Path(f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}")


def dcf_path(ticker):
    return dd_dir(ticker) / f"DCF {ticker}.xlsx"


def windows_path(path: Path | str) -> str:
    value = str(path)
    if os.name == "nt":
        return value
    try:
        return subprocess.check_output(["wslpath", "-w", value], text=True).strip()
    except Exception:
        return value


def default_years():
    last_full_year = datetime.now().year - 1
    return list(range(last_full_year - 4, last_full_year + 1))


def default_event_years():
    current_year = datetime.now().year
    return list(range(current_year - 3, current_year + 1))


def run(desc, cmd, cont_on_error=False):
    print(f"\n{'='*70}\n▶ {desc}\n  $ {' '.join(str(c) for c in cmd)}\n{'='*70}")
    t0 = time.time()
    r = subprocess.run(cmd, cwd=str(REPO))
    dt = int(time.time() - t0)
    if r.returncode != 0:
        print(f"✗ {desc} FAILED (exit {r.returncode}, {dt}s)")
        if not cont_on_error:
            sys.exit(r.returncode)
        return False
    print(f"✓ {desc} ({dt}s)")
    return True


def step_seed(a):
    dst = dcf_path(a.ticker)
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        bak = dst.with_name(f"DCF {a.ticker}_pre_rebuild_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(dst, bak)
        print(f"  backed up existing → {bak.name}")
    if not TEMPLATE.exists():
        sys.exit(f"Template not found: {TEMPLATE}")
    shutil.copy2(TEMPLATE, dst)
    print(f"✓ seed: copied template → {dst}")


def latest_trials(ticker):
    for base in (dd_dir(ticker), REPO / f"C:/Users/yzsun/Desktop/DD/{ticker}"):
        hits = sorted(glob.glob(f"{base}/*clinical_trials*.json"))
        if hits:
            return hits[-1]
    return None


def latest_file(patterns):
    hits = []
    for pat in patterns:
        hits.extend(glob.glob(str(pat)))
    return sorted(hits)[-1] if hits else None


SOC_DRUGS = {
    "placebo", "standard of care", "best supportive care", "comparator",
    "cemiplimab", "pembrolizumab", "nivolumab", "ipilimumab",
}

INACTIVE_TRIAL_STATUSES = {
    "WITHDRAWN", "TERMINATED", "SUSPENDED", "NO_LONGER_AVAILABLE",
    "TEMPORARILY_NOT_AVAILABLE",
}


def _clean_drug_name(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip())


# Dose amounts ("100 mg", "10mg", "5 mcg", "250 mg/mL"), formulation words, and
# trial-arm/part designators are stripped so that one asset written many ways on
# ClinicalTrials.gov ("BHV-7000", "BHV-7000 Tablets", "Troriluzole 100 mg") is
# counted once. Conservative: this MERGES only obvious duplicates of the same
# asset; it never fuses two distinct agents.
_DOSE_RE = re.compile(
    r"\b\d+(?:\.\d+)?\s*(?:mg|mcg|ug|µg|g|ml|iu|units?|%)\b(?:\s*/\s*\w+)?", re.I)
_FORM_RE = re.compile(
    r"\b(?:tablets?|capsules?|injections?|injectable|solution|suspension|oral|"
    r"intravenous|iv|subcutaneous|sc|infusion|film[- ]?coated|extended[- ]?release|"
    r"immediate[- ]?release|er|xr|sr|sachet|powder|gel|cream|patch|inhaled|"
    r"prefilled|syringe|vial)\b", re.I)
_ARM_RE = re.compile(
    r"\b(?:part|arm|cohort|group|dose[- ]?level|dl)\s*[0-9ivxab]+\b", re.I)


def _display_drug_name(name: str) -> str:
    """Strip combo/arm/dose/formulation qualifiers while preserving the asset's
    original casing and hyphenation, yielding a clean research-friendly name
    (e.g. "MP0533 monotherapy, Part 1" -> "MP0533", "BHV-7000 Tablets" ->
    "BHV-7000"). Falls back to the cleaned original if stripping empties it."""
    s = _clean_drug_name(name)
    # Keep only the primary agent from a combination / co-administration string.
    s = re.split(r"\s*\+\s*", s)[0]
    s = re.split(r"\s+(?:with|plus|and)\s+", s, flags=re.I)[0]
    # Drop trailing qualifier tails introduced by a comma or parenthesis.
    s = re.split(r"\s*[,(]", s)[0]
    s = _ARM_RE.sub(" ", s)
    s = re.sub(r"\bmonotherapy\b", " ", s, flags=re.I)
    s = _DOSE_RE.sub(" ", s)
    s = _FORM_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or _clean_drug_name(name)


def _canonical_drug_key(name: str) -> str:
    """Collapse dose/formulation/arm/combination variants of one asset to a
    single identity so each real clinical asset is counted once.

    Conservative by design: it merges only obvious duplicates (case, dosage,
    formulation, "Part N"/"Arm A" arms, and "+ combo" partners) and keeps
    distinct assets on distinct keys. INN<->development-code synonyms that are
    not textually related (e.g. Troriluzole <-> BHV-4157) are intentionally NOT
    merged here, since fusing them reliably needs an alias table and guessing
    risks over-merging genuinely different assets.
    """
    s = _display_drug_name(name).lower()
    # Join development-code hyphen/space variants: "bhv-7000"/"bhv 7000" -> "bhv7000".
    s = re.sub(r"([a-z]{1,6})[\s-]+(\d{2,})", r"\1\2", s)
    s = re.sub(r"[^\w]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def auto_detect_drugs(ticker: str) -> list[str]:
    """ClinicalTrials.gov-derived drug list for one-command builds.

    A drug is kept if it has at least one Phase 1+ company trial that is not
    withdrawn/terminated/suspended. Completed trials are retained because an
    asset can remain in the current pipeline after its prior trial completes.
    """
    trial_path = latest_trials(ticker)
    if not trial_path:
        return []
    try:
        data = json.loads(Path(trial_path).read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"  could not read trials JSON for drug auto-detect: {exc}")
        return []

    found: dict[str, dict] = {}
    for nct, trial in (data.get("trials") or {}).items():
        status = str(trial.get("status") or "").upper()
        if status in INACTIVE_TRIAL_STATUSES:
            continue
        phase = str(trial.get("phase") or "")
        if "PHASE" not in phase.upper() and "Unknown" not in phase:
            continue
        for raw in trial.get("interventions") or []:
            drug = _clean_drug_name(raw)
            low = drug.lower()
            if not drug or low in SOC_DRUGS or any(low.startswith(x) for x in SOC_DRUGS):
                continue
            if any(x in low for x in ("placebo", "standard", "comparator", "investigator")):
                continue
            key = _canonical_drug_key(drug)
            if not key:
                continue
            display = _display_drug_name(drug)
            slot = found.setdefault(key, {"n": 0, "ncts": set(), "display": display})
            # Keep the cleanest (shortest) spelling as the research name, so
            # "MP0533" wins over "MP0533 monotherapy, Part 1".
            if len(display) < len(slot["display"]):
                slot["display"] = display
            if nct not in slot["ncts"]:
                slot["n"] += 1
                slot["ncts"].add(nct)

    # Highest trial count first, then stable alpha order.
    drugs = [slot["display"] for _, slot in
             sorted(found.items(), key=lambda kv: (-kv[1]["n"], kv[1]["display"].lower()))]
    if drugs:
        print("  auto-detected clinical assets from trials JSON: " + ", ".join(drugs))
    return drugs


def choose_research_engine(requested: str) -> str:
    req = (requested or "auto").lower()
    if req in {"gemini", "opus", "none"}:
        return req
    if os.environ.get("GEMINI_API_KEY"):
        return "gemini"
    if os.environ.get("ANTHROPIC_API_KEY"):
        return "opus"
    return "none"


def main():
    ap = argparse.ArgumentParser(description="One-command DCF builder (ticker in → model out)")
    ap.add_argument("--ticker", required=True)
    ap.add_argument("--company-name", required=True)
    ap.add_argument("--drugs", nargs="+", help="Active clinical-stage assets to research")
    ap.add_argument("--cik", help="Override CIK (else auto-resolved)")
    ap.add_argument("--years", nargs="+", type=int,
                    help="Fiscal years to fetch (default: latest five completed years)")
    ap.add_argument("--flat-ms", type=float, default=0.10)
    ap.add_argument("--research-engine", choices=["auto", "gemini", "opus", "none"],
                    default="auto",
                    help="Report generator for parser-compatible per-drug reports")
    ap.add_argument("--judgement-engine", choices=["auto", "openai", "opus", "none"],
                    default="auto",
                    help="High-capability LLM for the model-assumptions calibration pass")
    ap.add_argument("--judgement-model",
                    help="Optional provider-specific model override for the judgement pass")
    ap.add_argument("--require-assumptions-approval", action="store_true",
                    help="Write a validated assumptions draft but do not auto-promote it")
    ap.add_argument("--assumptions-file",
                    help="Approved model assumptions JSON for ratings/economic shares")
    ap.add_argument("--events-file",
                    help="Approved events JSON/CSV/markdown for Historical Events")
    ap.add_argument("--news-url",
                    help="IR news-listing URL for req5's official-news scrape "
                         "(else resolved from NEWS_URL_BY_TICKER; if neither, the "
                         "build stops at the official-news gate)")
    ap.add_argument("--skip", nargs="+", default=[],
                    help=("Step names to skip (in run order): seed financials trials gptresearch "
                          "research pricing peerdatabase judgement scenarios pipeline wiretam peerviews "
                          "datastoretabs peerviewsummary events movecauses financials2 adaptris catalyst "
                          "catalystframework catalystclean "
                          "marketdata harden templateclean formats fsalinks statementlogic "
                          "liquidity "
                          "excelrepair finalpolish displayformats trimtabs referencestyles valuationcharts "
                          "namespaces corerecalc excelnormalize calcstate audit"))
    ap.add_argument("--from", dest="from_step", help="Start from this step name")
    ap.add_argument("--continue-on-error", action="store_true")
    a = ap.parse_args()

    reports = dd_dir(a.ticker) / "pipeline_base4"
    years = a.years or default_years()
    # Step execution is driven by the textual sequence of the `if active(name):`
    # blocks below; there is no separate `order` list to drift out of sync with it.
    started = a.from_step is None

    def active(name):
        nonlocal started
        if a.from_step and name == a.from_step:
            started = True
        return started and name not in a.skip

    coe = a.continue_on_error

    if active("seed"):
        step_seed(a)

    if active("financials"):
        cmd = [PY, "main.py", "--ticker", a.ticker, "--years", *map(str, years)]
        if a.cik:
            cmd += ["--cik", a.cik]
        run("2. Financials (SEC XBRL → FY DATA, check cells → 0)", cmd, coe)

    if active("trials"):
        run("3. Clinical trials (ClinicalTrials.gov)",
            [PY, "research/clinical_trials_fetcher.py", "--ticker", a.ticker,
             "--company-name", a.company_name], cont_on_error=True)

    if active("gptresearch"):
        run("4. GPT-5.5 Pro company-facts gate + research/judgment briefs",
            [PY, "research/gpt_research.py", "--ticker", a.ticker,
             "--company-name", a.company_name, "--brief", "all",
             "--years", *map(str, years), "--report-dir", str(reports)],
            cont_on_error=True)

    if active("research"):
        drugs = a.drugs or auto_detect_drugs(a.ticker)
        engine = choose_research_engine(a.research_engine)
        if engine == "none":
            print("  no research API key available; skipping parser-compatible research step")
            drugs = drugs or []
        if not drugs and engine != "none":
            sys.exit("--drugs required for the research step; auto-detect found none")
        tj = latest_trials(a.ticker)
        if engine == "gemini" and drugs:
            cmd = [PY, "research/gemini_research.py", "--ticker", a.ticker,
                   "--company-name", a.company_name, "--drugs", *drugs,
                   "--output-dir", str(reports)]
            if tj:
                cmd += ["--trials-json", tj]
            run("5. Gemini parser-compatible per-drug reports", cmd, coe)
        elif engine == "opus" and drugs:
            cmd = [PY, "research/opus_research.py", "--ticker", a.ticker,
                   "--company-name", a.company_name, "--drugs", *drugs,
                   "--flat-ms", str(a.flat_ms), "--market-share", "real",
                   "--output-dir", str(reports)]
            if tj:
                cmd += ["--trials-json", tj]
            run("5. Opus parser-compatible per-drug reports", cmd, coe)

    if active("pricing"):
        drugs = a.drugs or auto_detect_drugs(a.ticker)
        if os.environ.get("GEMINI_API_KEY") and drugs:
            run("5a. Gemini pricing chapters (list price per completed treatment)",
                [PY, "research/gemini_research.py", "--ticker", a.ticker,
                 "--company-name", a.company_name, "--drugs", *drugs,
                 "--output-dir", str(reports), "--chapter", "pricing"],
                cont_on_error=True)
        else:
            print("  pricing chapter skipped (requires GEMINI_API_KEY and detected drugs)")

    if active("peerdatabase"):
        run("5b. Upsert researched peer readouts into DD data center",
            [PY, "tools/update_peer_database_from_reports.py", "--ticker", a.ticker,
             "--report-dir", str(reports)],
            cont_on_error=True)

    if active("judgement"):
        if a.assumptions_file:
            print("  judgement skipped: explicit --assumptions-file is already approved")
        else:
            cmd = [PY, "research/model_judgement.py", "--ticker", a.ticker,
                   "--company-name", a.company_name, "--report-dir", str(reports),
                   "--engine", a.judgement_engine,
                   "--approval-mode",
                   "draft" if a.require_assumptions_approval else "auto"]
            if a.judgement_model:
                cmd += ["--model", a.judgement_model]
            # Best-effort by design. Missing keys/SDK, incomplete inputs, an
            # invalid LLM payload, or a human-review gate must not break the
            # mechanical build; Scenarios/Pipeline then consume any existing
            # approved JSON and otherwise retain their report-derived fallbacks.
            run("5c. Critical LLM calibration (include/exclude + model assumptions)",
                cmd, cont_on_error=True)

    if active("scenarios"):
        cmd = [PY, "generate/generate_scenarios.py", "--ticker", a.ticker,
               "--report-dir", str(reports), "--dcf-file", str(dcf_path(a.ticker))]
        if a.assumptions_file:
            cmd += ["--assumptions-file", a.assumptions_file]
        run("5. Scenarios (market shares from approved assumptions/datastore)",
            cmd, coe)

    if active("pipeline"):
        # generate_pipeline now upserts this ticker's researched TAM into the data
        # center AND inlines the DB TAM ($MM) straight into each Revenue formula,
        # so the Pipeline shows only the revenue-forecast rows (no TAM rows) and
        # the model can drop the TAM Solid/Blood tabs. The old 'wiretam' step is
        # therefore obsolete (kept as a skippable no-op alias below for --from).
        cmd = [PY, "generate/generate_pipeline.py", "--ticker", a.ticker,
               "--company-name", a.company_name, "--report-dir", str(reports)]
        if a.assumptions_file:
            cmd += ["--assumptions-file", a.assumptions_file]
        run("6. Pipeline (revenue = TAM[DB, inlined] × MS × maturity × economics)",
            cmd, coe)

    if active("wiretam"):
        # Obsolete: TAM is now inlined by the pipeline step. Retained only so an
        # explicit `--from wiretam` still resolves; it intentionally does nothing.
        print("• 6b. wiretam is obsolete (TAM inlined during pipeline) — skipping")

    if active("peerviews"):
        run("7. Peer Views (drug-vs-drug readouts)",
            [PY, "fill/fill_peer_views.py", "--ticker", a.ticker,
             "--report-dir", str(reports)], cont_on_error=True)

    if active("datastoretabs"):
        # Output here is wholesale-regenerated by step 7b (build_peer_view_summary)
        # and 'Peer Views' is later deleted by trimtabs, so a failure must not abort
        # the build (Finding 29).
        run("7a. Peer Views data-center restore",
            [PY, "tools/fill_peer_views_from_datastore.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))], cont_on_error=True)

    if active("peerviewsummary"):
        run("7b. Ticker-specific Peer View summary",
            [PY, "tools/build_peer_view_summary.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))], coe)

    # Captured so the movecauses step can re-run the identical events command as a
    # 2nd pass (consuming researched move causes) without re-probing the network
    # for news-URL discovery. None when the events step did not run this build.
    event_cmd = None
    if active("events"):
        events_file = a.events_file or latest_file([
            reports / f"{a.ticker}_events.json",
            REPO / "artifacts" / a.ticker / f"{a.ticker}_official_events.json",
            dd_dir(a.ticker) / f"{a.ticker}_events.json",
            dd_dir(a.ticker) / "gpt_research" / f"{a.ticker}_events_*.md",
        ])
        event_cmd = [PY, "tools/fill_historical_events.py", "--ticker", a.ticker,
                     "--years", *map(str, default_event_years()),
                     "--company-name", a.company_name, "--require-official-news"]
        if events_file:
            event_cmd += ["--events-file", str(events_file)]
        news_url = resolve_news_url(a.ticker, a.news_url)
        if news_url:
            event_cmd += ["--news-url", news_url]
            news_bit = "official news + "
        else:
            sys.exit(
                f"Historical Events blocked for {a.ticker}: no authoritative IR news archive "
                "URL. Add ir_news_url to the company-facts artifact or pass --news-url."
            )
        if events_file:
            label = f"7b. Historical Events (approved events + {news_bit}prices/DoD)"
        else:
            label = f"7b. Historical Events ({news_bit}prices/DoD)"
        run(label, event_cmd, coe)

    if active("movecauses"):
        # The events step above writes the dates of unexplained >=8% moves to
        # artifacts/{T}/{T}_moves_needing_research.json. Research their causes and,
        # if a researched-causes file is produced, re-run events so
        # fill_historical_events (_load_move_research) folds them into the workbook.
        needing = REPO / "artifacts" / a.ticker / f"{a.ticker}_moves_needing_research.json"
        pending = 0
        if needing.exists():
            try:
                data = json.loads(needing.read_text(encoding="utf-8"))
                moves = data.get("moves", data) if isinstance(data, dict) else data
                pending = len(moves or [])
            except Exception:
                pending = 0
        if pending:
            # Research is best-effort: a missing movecauses brief or API key must
            # not abort the build (cont_on_error), it just leaves generic labels.
            run("8a. Research causes for unexplained large share moves",
                [PY, "research/gpt_research.py", "--ticker", a.ticker,
                 "--company-name", a.company_name, "--brief", "movecauses",
                 "--report-dir", str(reports)],
                cont_on_error=True)
            researched = REPO / "artifacts" / a.ticker / f"{a.ticker}_moves_researched.json"
            if researched.exists() and event_cmd is not None:
                run("8a2. Historical Events (2nd pass: consume researched move causes)",
                    event_cmd, cont_on_error=True)
            elif researched.exists():
                print("  movecauses: researched causes ready but the events step was "
                      "skipped this run; re-run 'events' to fold them into the workbook.")
            else:
                print("  movecauses: no researched-causes file produced (research is "
                      "best-effort); Historical Events keeps the generic move labels.")
        else:
            print(f"  movecauses: no unexplained large moves flagged for {a.ticker}; "
                  "nothing to research.")

    if active("financials2"):
        run("8. RBS / RCFS (restated statements → DCF)",
            [PY, "generate/generate_financials.py", "--ticker", a.ticker], coe)

    if active("adaptris"):
        run("9. Adapt RIS/RBS to this ticker's pipeline (revenue + opex flow)",
            [PY, "generate/adapt_ris.py", "--ticker", a.ticker], coe)

    if active("catalyst"):
        run("10. Catalyst targets → every drug × indication (no aggregation/drop)",
            [PY, "generate/adapt_catalyst.py", "--ticker", a.ticker], coe)

    if active("catalystframework"):
        run("10a. Catalyst Tables 2/3 + Scenarios + Post-Catalyst Price Table",
            [PY, "generate/build_catalyst_framework.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))], coe)

    if active("catalystclean"):
        run("10b. Catalyst neutral-input / zero-residue gate",
            [PY, "tools/catalyst_workflow.py", "clean", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker)), "--force"], coe)

    if active("marketdata"):
        run("11. Market data → BBG DAPI (yfinance: price/beta/shares/consensus)",
            [PY, "fill/fill_market_data.py", "--ticker", a.ticker], cont_on_error=True)
        run("11a. Freeze Bloomberg add-in formulas + clear workbook comments",
            [PY, "tools/neutralize_bloomberg_artifacts.py", "--ticker", a.ticker],
            cont_on_error=True)

    if active("harden"):
        run("11. Harden residual pre-revenue ratio cells (IFERROR) + verify",
            [PY, "tools/harden_formulas.py", "--ticker", a.ticker], cont_on_error=True)

    if active("templateclean"):
        run("11b. Clean stale template labels across support tabs",
            [PY, "tools/clean_workbook_template.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker)), "--years", *map(str, years)],
            coe)

    if active("formats"):
        run("11c. Normalize generated workbook formats",
            [PY, "tools/normalize_workbook_formats.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("fsalinks"):
        run("11d. Repair FSA/RIS/Schedules formula links",
            [PY, "tools/fix_fsa_links.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("statementlogic"):
        run("11d2. Repair RIS opex breakdown and RBS equity roll-forward",
            [PY, "tools/fix_statement_logic.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("liquidity"):
        run("11d3. Reconcile cash + current/non-current marketable securities",
            [PY, "tools/fix_liquidity_rollforward.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("excelrepair"):
        run("11f. Repair Excel package compatibility",
            [PY, "tools/repair_excel_package.py", "--ticker", a.ticker], coe)

    if active("finalpolish"):
        run("11g. Excel-native final polish (Pipeline/FY DATA/Peer View/comments)",
            [PY, "tools/final_excel_polish.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("displayformats"):
        run("11h. Sanitize display formats (no dash/strike zero display)",
            [PY, "tools/sanitize_excel_display_formats.py", "--path", str(dcf_path(a.ticker))],
            coe)

    if active("trimtabs"):
        run("11h2. Trim build-time data-center tabs from delivered model",
            [PY, "tools/trim_model_tabs.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("referencestyles"):
        # This must be the final cell-format mutation.  Display-format cleanup
        # and tab trimming run first so no later build step can reintroduce a
        # style drift after the MOLN reference system has been applied.
        run("11h3. Apply approved MOLN reference styles to every delivered tab",
            [PY, "tools/apply_reference_model_styles.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("valuationcharts"):
        # Drawing creation is intentionally after the last cell-format pass.
        # The chart worker runs in manual calculation mode, so large Catalyst
        # What-If tables cannot prevent the drawing objects from being saved.
        run("11h4. Rebuild Valuation Waterfall + Football Field charts",
            [PY, "tools/build_valuation_charts.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))],
            coe)

    if active("namespaces"):
        run("11j. Repair Excel-sensitive OOXML namespace prefixes",
            [PY, "tools/repair_excel_namespaces.py", "--path", str(dcf_path(a.ticker))],
            coe)

    if active("corerecalc"):
        run("11j2. Excel-native core recalc through Catalyst data table",
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", windows_path(REPO / "tools" / "excel_recalc_model_core.ps1"),
             "-Path", windows_path(dcf_path(a.ticker))],
            coe)

    if active("excelnormalize"):
        # Final workbook mutation.  Every XML-writing step above can invalidate
        # optional Excel caches (especially calcChain/drawing/chart caches).
        # Put Excel's own repair-save path last so the delivered file opens
        # without an Open-and-Repair prompt.
        win_path = windows_path(dcf_path(a.ticker))
        run("11k. Final Excel-native repair-save normalization",
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", windows_path(REPO / "tools" / "excel_repair_saveas.ps1"),
             "-Path", win_path],
            cont_on_error=True)

    if active("calcstate"):
        # This must follow every Excel/OOXML mutation. Manual-mode builders and
        # Excel repair saves can otherwise reintroduce the dirty calculation
        # metadata that current Excel renders as stale-value strikethrough.
        run("11l. Final calc-state normalization (no stale-value strikethrough)",
            [PY, "tools/normalize_calc_state.py", "--path", str(dcf_path(a.ticker))],
            coe)

    if active("audit"):
        run("12. Workbook stale-ticker audit",
            [PY, "tools/workbook_audit.py", "--ticker", a.ticker,
             "--strict-research"], coe)
        run("12a. Strict MOLN-reference style and Catalyst-layout audit",
            [PY, "tools/workbook_style_audit.py", "--ticker", a.ticker,
             "--path", str(dcf_path(a.ticker))], coe)
        # Hard gate (finding 13): recalc_check now exits non-zero when a FY DATA /
        # RBS check cell is non-zero or an #REF!/#VALUE! survives recalc, so a
        # structurally broken model can no longer ship as "build complete".
        # Overridable with --continue-on-error for intentional partial runs.
        run("13. Recalc/check cells",
            [PY, "tools/recalc_check.py", "--ticker", a.ticker], coe)

    print(f"\n{'='*70}\n✔ DCF build complete → {dcf_path(a.ticker)}")
    print("  Verify anytime:  python tools/recalc_check.py --ticker " + a.ticker)
    print("  Audit anytime:   python tools/workbook_audit.py --ticker " + a.ticker)
    print("="*70)


if __name__ == "__main__":
    main()
