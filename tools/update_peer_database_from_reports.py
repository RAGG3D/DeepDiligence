#!/usr/bin/env python3
"""Upsert parser-compatible PEER_VIEW blocks into the DD data center.

Research reports are the first place a new ticker's drug/peer readouts exist.
This script promotes those readouts into datastore/seed/peer_views.json and
rebuilds datastore/export/*.csv so future workbooks use the database as the
system of record instead of re-parsing old reports or carrying raw Peer Views
tabs inside each model.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from fill.fill_peer_views import DrugReadout, parse_peer_view_blocks  # noqa: E402

SEED_PATH = REPO / "datastore" / "seed" / "peer_views.json"


FIELD_TO_METRIC = {
    "innovation": "Innovation",
    "target": "Target",
    "result": "Result",
    "nct": "NCT#",
    "treatment_line": "Treatment Line",
    "phase": "Readout Phase",
    "stage": "Stage",
    "data_date": "Date",
    "conference": "Conference",
    "n_patients": "Evaluable Patients",
    "orr": "ORR",
    "bicr_orr": "BICR ORR",
    "cr": "CR",
    "pr": "PR",
    "dcr": "DCR",
    "median_pfs": "Median PFS",
    "median_rpfs": "Median rPFS",
    "median_os": "Median OS",
    "pfs_6mo": "6 Mo PFS Rate",
    "pfs_12mo": "12 Mo PFS Rate",
    "pfs_24mo": "24 Mo PFS Rate",
    "os_18mo": "18 Mo OS Rate",
    "os_24mo": "24 Mo OS Rate",
    "median_dfs": "Median DFS",
    "median_followup": "Median Follow-Up",
    "geq_g3_sae_pct": "GEQ G3 SAE/Patients",
    "geq_g3_clinical_ae": "GEQ G3 Clinical AE",
    "route": "Route",
    "dosing_schedule": "Dosing Schedule",
    "latest_annual_sale": "Latest Sale (MM USD)",
    "first_yr_sale": "1st Yr Sale (MM USD)",
    "stock_price_before": "Stock Price Day Before",
    "stock_price_after": "Stock Price Day After",
    "stock_change_1d": "Stock Change 1d",
    "stock_change_3d": "Stock Change 3d",
    "source": "Source",
}


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text or "").lower())


# ── Differentiation-Assessment → DB rating -----------------------------------
# Per-drug research reports carry a "### 3.N.5 Differentiation Assessment"
# section whose "Assessment:" verdict grades the ticker's own asset vs peers.
# We promote that verdict into the datastore rating (req4) so Pipeline ratings
# come from the DB instead of defaulting to "Average".  MOLN reports have no
# such section, so this is a no-op there (rating stays None).
_SECTION_RE = re.compile(r"^##[ \t]+3\.(\d+)[ \t]+(.+?)[ \t]*$", re.MULTILINE)
_ASSESSMENT_RE = re.compile(
    r"3\.(\d+)\.5\s+Differentiation Assessment\s+Assessment:\s*([^\n]+)",
    re.IGNORECASE,
)
_PEER_VIEW_RE = re.compile(r"####\s*PEER_VIEW_START:\s*(.+)")


def _assessment_to_rating(text: str) -> str | None:
    """Map a Differentiation-Assessment verdict OR a per-competitor "- Rating:"
    value to a DB rating code (BIC/T1/AVG).

    Accepts both the verbose tier strings the prompts emit ("Best-in-Class",
    "Tier One"/"Above-average", "Average"/"Below-average") and the short tier
    codes (BIC/T1/AVG) a "- Rating:" line may carry directly.  Absent/unknown
    -> None (unchanged behavior).
    """
    t = (text or "").strip().lower()
    # Exact short-code forms first (a bare "- Rating: BIC|T1|AVG" line).
    if t in ("bic", "best-in-class", "best in class"):
        return "BIC"
    if t in ("t1", "tier one", "tier 1", "tier-one", "above-average", "above average"):
        return "T1"
    if t in ("avg", "average", "below-average", "below average"):
        return "AVG"
    # Verbose substring forms (Differentiation-Assessment prose).
    if "best-in-class" in t or "best in class" in t:
        return "BIC"
    if "above-average" in t or "above average" in t or "tier one" in t or "tier 1" in t:
        return "T1"
    if "average" in t or "below" in t:
        return "AVG"
    return None


def _indication_keys(name: str) -> set[str]:
    """Normalized lookup keys for a "## 3.N <name>" indication heading."""
    keys: set[str] = set()
    n = _norm(name)
    if n:
        keys.add(n)
    base = re.sub(r"\s*\([^)]*\)\s*$", "", name).strip()
    if base:
        keys.add(_norm(base))
    for paren in re.findall(r"\(([^)]*)\)", name):
        pk = _norm(paren)
        if pk:
            keys.add(pk)
    return keys


def _assessment_ratings(text: str) -> dict[str, str]:
    """Map normalized indication -> rating code from a research report.

    Reports without any "3.N.5 Differentiation Assessment" section (e.g. MOLN)
    yield {}, so no ratings attach and the prior behavior (rating=None) holds.
    """
    sec_names: dict[int, str] = {}
    for m in _SECTION_RE.finditer(text):
        sec_names[int(m.group(1))] = m.group(2).strip()
    num_ratings: dict[int, str] = {}
    for m in _ASSESSMENT_RE.finditer(text):
        rating = _assessment_to_rating(m.group(2))
        if rating:
            num_ratings.setdefault(int(m.group(1)), rating)
    if not num_ratings:
        return {}
    out: dict[str, str] = {}
    # (a) key by name / parenthetical code from the "## 3.N <name>" headings.
    for num, rating in num_ratings.items():
        for key in _indication_keys(sec_names.get(num, "")):
            out.setdefault(key, rating)
    # (b) positional fallback: the Nth PEER_VIEW indication (first-appearance
    #     order) corresponds to the Nth assessment.  This covers reports whose
    #     "## 3.1 Full Indication Name (IgA Nephropathy)" heading uses a long
    #     name while the PEER_VIEW block uses a short code ("IgAN").
    seen: list[str] = []
    seen_norm: set[str] = set()
    for m in _PEER_VIEW_RE.finditer(text):
        nn = _norm(m.group(1).strip())
        if nn and nn not in seen_norm:
            seen_norm.add(nn)
            seen.append(nn)
    ordered_ratings = [num_ratings[n] for n in sorted(num_ratings)]
    for key, rating in zip(seen, ordered_ratings):
        out.setdefault(key, rating)
    return out


def _is_company_drug(readout: DrugReadout, ticker: str) -> bool:
    """True when a readout is the current ticker's own asset (not a peer).

    Requires exact normalized-ticker equality: own-drug readouts carry a clean
    "Ticker: BHVN", whereas peers whose prose mentions the ticker (e.g.
    "UCB (not NYSE:BHVN)", "PFE ..., BHVN (for Biohaven ...)") must NOT inherit
    the company's differentiation rating.
    """
    return bool(_norm(readout.ticker)) and _norm(readout.ticker) == _norm(ticker)


def _col_letter(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def _col_num(col: str) -> int:
    out = 0
    for ch in col:
        out = out * 26 + ord(ch.upper()) - 64
    return out


def _load_seed() -> list[dict[str, Any]]:
    if not SEED_PATH.exists():
        return []
    data = json.loads(SEED_PATH.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        data = data.get("sections", [])
    return data if isinstance(data, list) else []


def _save_seed(sections: list[dict[str, Any]]) -> None:
    SEED_PATH.parent.mkdir(parents=True, exist_ok=True)
    SEED_PATH.write_text(
        json.dumps(sections, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def _read_reports(report_dir: Path, ticker: str) -> dict[str, list[DrugReadout]]:
    blocks: dict[str, list[DrugReadout]] = {}
    paths = sorted(report_dir.glob(f"{ticker.upper()}_*_research_*.md"))
    latest: dict[str, Path] = {}
    prefix = f"{ticker.upper()}_"
    marker = "_research_"
    for path in paths:
        # A later modelling supplement may intentionally contain only TAM and
        # scenario curves.  It must not hide the latest full report carrying
        # PEER_VIEW blocks for the database refresh.
        if "PEER_VIEW_START:" not in path.read_text(encoding="utf-8", errors="ignore"):
            continue
        name = path.stem
        if not name.startswith(prefix) or marker not in name:
            continue
        drug, version = name[len(prefix):].split(marker, 1)
        key = drug.upper()
        current = latest.get(key)
        current_version = current.stem.split(marker, 1)[1] if current else ""
        if current is None or (version, path.stat().st_mtime) > (
            current_version, current.stat().st_mtime
        ):
            latest[key] = path
    for path in sorted(latest.values()):
        text = path.read_text(encoding="utf-8", errors="ignore")
        parsed = parse_peer_view_blocks(text)
        ratings = _assessment_ratings(text)
        for indication, readouts in parsed.items():
            rating = ratings.get(_norm(indication))
            if rating:
                for readout in readouts:
                    if _is_company_drug(readout, ticker):
                        # Stash the company's Differentiation-Assessment rating on
                        # a private ._rating (distinct from the readout's own
                        # per-competitor `rating` field) so upsert prefers it for
                        # the ticker's own asset.
                        readout._rating = rating  # type: ignore[attr-defined]
            blocks.setdefault(indication, []).extend(readouts)
    return blocks


def _readout_key(readout: DrugReadout) -> tuple[str, str, str, str]:
    return (
        _norm(readout.drug_name),
        _norm(readout.nct),
        str(readout.data_date or "").strip(),
        _norm(readout.source),
    )


def _metrics(readout: DrugReadout) -> dict[str, str]:
    out: dict[str, str] = {}
    for field, metric in FIELD_TO_METRIC.items():
        value = getattr(readout, field, "")
        if value not in (None, "", "/"):
            out[metric] = str(value)
    return out


def _display_ticker(ticker: str) -> str:
    value = (ticker or "").strip()
    if value in ("", "/"):
        return ""
    upper = value.upper()
    if "EQUITY" in upper or upper in {"PRIVATE", "UNLISTED"}:
        return value
    if re.fullmatch(r"[A-Z.]{1,5}", upper):
        return upper
    return value


def _next_col(drugs: list[dict[str, Any]]) -> str:
    used = [_col_num(str(d.get("col", ""))) for d in drugs if d.get("col")]
    return _col_letter(max([4] + used) + 1)


def upsert_peer_views(report_dir: Path, ticker: str) -> int:
    blocks = _read_reports(report_dir, ticker)
    if not blocks:
        print(f"No PEER_VIEW blocks found in {report_dir}")
        return 0

    sections = _load_seed()
    approved_ratings: dict[str, dict[str, str]] = {}
    assumptions_path = REPO / "artifacts" / ticker.upper() / f"{ticker.upper()}_model_assumptions.json"
    if assumptions_path.exists():
        try:
            assumptions = json.loads(assumptions_path.read_text(encoding="utf-8"))
            approved_ratings = {
                _norm(drug): {_norm(ind): str(rating) for ind, rating in ind_map.items()}
                for drug, ind_map in (assumptions.get("ratings") or {}).items()
                if isinstance(ind_map, dict)
            }
        except Exception as exc:
            print(f"Warning: could not load approved ratings from {assumptions_path}: {exc}")
    # A ticker refresh is a replacement, not an append-only event.  Remove its
    # own historical columns everywhere first so superseded readouts (for
    # example an endpoint that was previously "due") cannot coexist with the
    # reported result under a differently named legacy section.
    removed = 0
    for section in sections:
        drugs = section.get("drugs") or []
        kept = [d for d in drugs if _norm(d.get("ticker", "")) != _norm(ticker)]
        removed += len(drugs) - len(kept)
        section["drugs"] = kept
    by_norm = {_norm(s.get("section", "")): s for s in sections}
    written = 0

    for indication, readouts in sorted(blocks.items(), key=lambda kv: _norm(kv[0])):
        sec = by_norm.get(_norm(indication))
        if not sec:
            sec = {"section": indication, "drugs": []}
            sections.append(sec)
            by_norm[_norm(indication)] = sec

        drugs = sec.setdefault("drugs", [])
        existing_keys = set()
        for drug in drugs:
            metrics = drug.get("metrics") or {}
            existing_keys.add((
                _norm(drug.get("drug", "")),
                _norm(metrics.get("NCT#", "")),
                str(metrics.get("Date", "")).strip(),
                _norm(metrics.get("Source", "")),
            ))

        for readout in readouts:
            if not readout.drug_name:
                continue
            key = _readout_key(readout)
            if key in existing_keys:
                continue
            # Company-own-drug rating (derived from the report's Differentiation
            # Assessment and stashed on ._rating) takes precedence.  Otherwise map
            # the per-competitor "- Rating:" line so peer rows get a BIC/T1/AVG
            # code instead of a blank Peer View rating cell.
            rating = None
            if _is_company_drug(readout, ticker):
                rating = approved_ratings.get(_norm(readout.drug_name), {}).get(
                    _norm(indication)
                )
            if not rating:
                rating = getattr(readout, "_rating", None)
            if not rating:
                rating = _assessment_to_rating(getattr(readout, "rating", "") or "")
            drugs.append({
                "col": _next_col(drugs),
                "drug": readout.drug_name,
                "ticker": _display_ticker(readout.ticker),
                "rating": rating,
                "metrics": _metrics(readout),
            })
            existing_keys.add(key)
            written += 1

    if written or removed:
        _save_seed(sections)
    print(
        f"Peer database refresh: removed {removed} stale {ticker} columns; "
        f"inserted {written} current readout columns"
    )
    return written


def main() -> None:
    parser = argparse.ArgumentParser(description="Upsert PEER_VIEW report blocks into datastore")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--report-dir", required=True)
    parser.add_argument("--no-rebuild", action="store_true")
    args = parser.parse_args()

    written = upsert_peer_views(Path(args.report_dir), args.ticker.upper())
    if written and not args.no_rebuild:
        subprocess.run(
            [sys.executable, "datastore/build_datastore.py"],
            cwd=REPO,
            check=True,
        )


if __name__ == "__main__":
    main()
