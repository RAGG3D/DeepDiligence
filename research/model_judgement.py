#!/usr/bin/env python3
"""Critically calibrate report-derived DCF model assumptions with an LLM.

The pass runs after research/peer-database collection and before Scenarios and
Pipeline.  It writes a schema-validated draft first.  By default the draft is
promoted only when no approved assumptions file exists; ``--approval-mode
draft`` leaves it for explicit human review.  Existing approved/manual JSON is
never overwritten.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any


REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from core.model_assumptions import (  # noqa: E402
    AssumptionsValidationError,
    derive_inclusion_decisions,
    extract_json_object,
    normalize_model_assumptions,
)
from datastore.research_fact_store import (  # noqa: E402
    database_sync_prompt,
    scan_research_context,
    upsert_research_facts,
)


OPENAI_MODEL = "gpt-5.6-sol"
OPUS_MODEL = "claude-opus-4-8"
DEFAULT_MAX_INPUT_CHARS = 750_000


def _read_dotenv() -> dict[str, str]:
    values: dict[str, str] = {}
    path = REPO / ".env"
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def _secret(*names: str) -> str | None:
    for name in names:
        value = os.environ.get(name)
        if value:
            return value
    dotenv = _read_dotenv()
    for name in names:
        if dotenv.get(name):
            return dotenv[name]
    return None


def choose_engine(requested: str) -> str | None:
    """Choose only a high-capability judgement provider; never a weak fallback."""
    requested = (requested or "auto").lower()
    if requested == "none":
        return None
    if requested == "openai":
        return "openai" if _secret("OPENAI_API_KEY") else None
    if requested == "opus":
        return "opus" if _secret("ANTHROPIC_API_KEY", "MY_PYTHON_SCRIPT_KEY") else None
    if _secret("OPENAI_API_KEY"):
        return "openai"
    if _secret("ANTHROPIC_API_KEY", "MY_PYTHON_SCRIPT_KEY"):
        return "opus"
    return None


def _drug_from_report_name(path: Path, ticker: str) -> str:
    match = re.match(
        rf"{re.escape(ticker)}_(.+?)_research_", path.stem, flags=re.I
    )
    return (match.group(1).replace("_", "-") if match else path.stem).strip()


def latest_drug_reports(report_dir: Path, ticker: str) -> list[Path]:
    """Return only the newest report per drug, avoiding stale-version conflicts."""
    paths = set(report_dir.glob(f"{ticker}_*_research_*.md"))
    paths.update(report_dir.glob(f"{ticker.upper()}_*_research_*.md"))
    newest: dict[str, Path] = {}
    for path in paths:
        key = re.sub(r"[^a-z0-9]+", "", _drug_from_report_name(path, ticker).lower())
        previous = newest.get(key)
        if previous is None or (path.stat().st_mtime, path.name) > (
            previous.stat().st_mtime, previous.name
        ):
            newest[key] = path
    return sorted(newest.values(), key=lambda p: p.name.lower())


def load_program_manifest(report_dir: Path, ticker: str) -> tuple[dict[str, list[str]], list[Path]]:
    """Parse the exact drug/indication keys downstream generators will consume."""
    from generate.generate_scenarios import _parse_single_drug_report

    reports = latest_drug_reports(report_dir, ticker)
    manifest: dict[str, list[str]] = {}
    for path in reports:
        asset = _parse_single_drug_report(path, ticker)
        if asset is None:
            continue
        indications = list(asset.market_shares) or list(asset.indications)
        if not indications:
            continue
        slot = manifest.setdefault(asset.name, [])
        for indication in indications:
            if indication not in slot:
                slot.append(indication)
    return manifest, reports


def _latest_gpt_briefs(report_dir: Path, ticker: str) -> list[Path]:
    roots = [report_dir, report_dir.parent / "gpt_research"]
    newest: dict[str, Path] = {}
    for root in roots:
        if not root.exists():
            continue
        paths = set(root.glob(f"{ticker}_[a-zA-Z]*_*.md"))
        paths.update(root.glob(f"{ticker.upper()}_[a-zA-Z]*_*.md"))
        for path in paths:
            match = re.match(rf"{re.escape(ticker)}_([a-zA-Z]+)_", path.name, re.I)
            if not match or match.group(1).lower() not in {
                "pipeline", "ratings", "events", "notes"
            }:
                continue
            key = match.group(1).lower()
            prior = newest.get(key)
            if prior is None or (path.stat().st_mtime, path.name) > (
                prior.stat().st_mtime, prior.name
            ):
                newest[key] = path
    return [newest[key] for key in sorted(newest)]


def _tam_context(manifest: dict[str, list[str]], ticker: str) -> str:
    expected = {
        re.sub(r"[^a-z0-9]+", "", indication.lower()): indication
        for indications in manifest.values() for indication in indications
    }
    path = REPO / "datastore" / "export" / "tam_by_indication_year.csv"
    anchors: dict[str, dict[int, float]] = {ind: {} for ind in expected.values()}
    if path.exists():
        with path.open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                normalized = re.sub(
                    r"[^a-z0-9]+", "", (row.get("indication_code") or "").lower()
                )
                indication = expected.get(normalized)
                if not indication:
                    continue
                try:
                    year = int(row.get("year", ""))
                    if year in {2024, 2030, 2038}:
                        anchors[indication][year] = float(row["tam_usd_m"])
                except (KeyError, TypeError, ValueError):
                    continue
    lines = ["TAM datastore anchors (USD MM; exact indication-code matches):"]
    for indication in sorted(anchors):
        values = anchors[indication]
        if values:
            rendered = ", ".join(f"{year}={values[year]:.3f}" for year in sorted(values))
            lines.append(f"- {indication}: {rendered}")
        else:
            lines.append(f"- {indication}: NO EXACT DATASTORE MATCH; use the report-derived TAM critically")

    overrides = REPO / "datastore" / "export" / "tam_override.csv"
    scoped: dict[tuple[str, str], dict[int, float]] = {}
    if overrides.exists():
        with overrides.open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                if (row.get("source_ticker") or "").upper() != ticker.upper():
                    continue
                try:
                    year = int(row.get("year", ""))
                    if year not in {2024, 2030, 2038}:
                        continue
                    key = (row.get("source_drug") or "", row.get("indication_code") or "")
                    scoped.setdefault(key, {})[year] = float(row["tam_usd_m"])
                except (KeyError, TypeError, ValueError):
                    continue
    if scoped:
        lines.append(f"Ticker-scoped TAM overrides for {ticker.upper()}:")
        for (drug, indication), values in sorted(scoped.items()):
            rendered = ", ".join(f"{year}={values[year]:.3f}" for year in sorted(values))
            lines.append(f"- {drug}/{indication}: {rendered}")
    return "\n".join(lines)


def _maturity_context() -> str:
    path = REPO / "datastore" / "export" / "param_maturity.csv"
    if not path.exists():
        return "Maturity datastore export is missing."
    rows: dict[str, list[tuple[int, float]]] = {}
    with path.open(newline="", encoding="utf-8-sig") as handle:
        for row in csv.DictReader(handle):
            try:
                rows.setdefault(row["tier"], []).append(
                    (int(row["year_offset"]), float(row["factor"]))
                )
            except (KeyError, TypeError, ValueError):
                continue
    lines = ["Full datastore maturity curves (year_offset=factor):"]
    for tier in sorted(rows):
        curve = ", ".join(f"{year}={factor:.6g}" for year, factor in sorted(rows[tier]))
        lines.append(f"- {tier}: {curve}")
    return "\n".join(lines)


def _peer_company_path(ticker: str) -> Path:
    return REPO / "artifacts" / ticker.upper() / f"{ticker.upper()}_peer_company_data.json"


def build_context(
    report_dir: Path,
    ticker: str,
    manifest: dict[str, list[str]],
    reports: list[Path],
) -> str:
    """Assemble the locked judgement inputs, with file provenance labels."""
    from tools.build_peer_view_summary import ensure_company_peer_data_from_reports

    ensure_company_peer_data_from_reports(report_dir, ticker, set(manifest))
    sections = [
        "## PROGRAM MANIFEST (keys must be copied exactly)",
        json.dumps(manifest, indent=2, ensure_ascii=False),
        "\n## DATASTORE TAM",
        _tam_context(manifest, ticker),
        "\n## DATASTORE MATURITY CURVES",
        _maturity_context(),
    ]
    peer_path = _peer_company_path(ticker)
    sections.append("\n## COMPANY CLINICAL READOUTS (peer_company_data)")
    if peer_path.exists():
        sections.extend([
            f"Source file: {peer_path}",
            peer_path.read_text(encoding="utf-8", errors="ignore"),
        ])
    else:
        sections.append("No materializable company readout JSON was available; do not invent data.")

    sections.append("\n## PER-DRUG RESEARCH REPORTS (newest version per drug)")
    for path in reports:
        sections.extend([
            f"\n### SOURCE FILE: {path.name}",
            path.read_text(encoding="utf-8", errors="ignore"),
        ])
    briefs = _latest_gpt_briefs(report_dir, ticker)
    if briefs:
        sections.append("\n## COMPANY-WIDE RESEARCH BRIEFS (newest per brief type)")
        for path in briefs:
            sections.extend([
                f"\n### SOURCE FILE: {path.name}",
                path.read_text(encoding="utf-8", errors="ignore"),
            ])
    return "\n".join(sections)


def build_prompt(
    ticker: str,
    company_name: str,
    manifest: dict[str, list[str]],
    context: str,
) -> str:
    ratings = {drug: {ind: "AVG" for ind in inds} for drug, inds in manifest.items()}
    shares = {
        drug: {
            ind: {"base_peak": 0.0, "bull_peak": 0.0, "bear_peak": 0.0}
            for ind in inds
        }
        for drug, inds in manifest.items()
    }
    skeleton = {
        "source": "AUTO_DCF critical judgement from named input files, YYYY-MM-DD",
        "economic_share": {drug: 1.0 for drug in manifest},
        "ratings": ratings,
        "market_share": shares,
        "market_share_notes": {
            drug: "; ".join(f"INCLUDE {ind}: evidence-based reason" for ind in inds)
            for drug, inds in manifest.items()
        },
    }
    return f"""You are the final senior-biotech-investor judgement pass for {company_name} ({ticker}).
Critically audit the supplied research rather than copying its forecasts. Use ONLY the supplied
reports, company clinical readouts, and datastore context. Reconcile conflicting readouts by date,
trial population, treatment line, sample size, efficacy, durability, safety, route, competition,
ownership/economics, TAM fit, development status, and probability of a credible approval path.

For EVERY exact drug x indication pair in PROGRAM MANIFEST:
1. Decide commercial-model inclusion. INCLUDE only an active Phase 1+ program with a credible,
   differentiated and economically relevant path. EXCLUDE dormant/discontinued/preclinical-only,
   duplicate/non-addressable, unsupported, or commercially immaterial programs.
2. Encode EXCLUDE using base_peak=bull_peak=bear_peak=0.0. Encode INCLUDE with at least one
   positive peak. Never omit a pair. The guard rejects a draft that excludes every pair.
3. Assign rating using exactly BIC, T1, or AVG. BIC means credible best-in-class evidence; T1 means
   strong/above-average but not dominant; AVG includes average/below-average. Do not use prose labels.
4. Set economic_share to the company's net fraction of product sales, 0 < value <= 1.
5. Calibrate peaks as 0-1 fractions with bear_peak <= base_peak <= bull_peak. Use serviceable TAM,
   line-matched competition and maturity curves. Avoid false precision and unjustified dominant share.
6. In market_share_notes, write an explicit `INCLUDE <indication>:` or `EXCLUDE <indication>:`
   clause for EACH indication, followed by the decisive evidence and source filename(s).

Return ONLY valid JSON. The top-level keys and every nested key must match this skeleton exactly;
do not add decision/confidence/metadata fields:
{json.dumps(skeleton, indent=2, ensure_ascii=False)}

LOCKED INPUT PACKET
{context}
"""


def call_llm(engine: str, prompt: str, model: str | None = None) -> tuple[str, str]:
    if engine == "openai":
        try:
            import openai
        except ImportError as exc:
            raise RuntimeError("OpenAI SDK is not installed") from exc
        selected = model or OPENAI_MODEL
        client = openai.OpenAI(api_key=_secret("OPENAI_API_KEY"))
        response = client.responses.create(
            model=selected,
            input=prompt,
            max_output_tokens=12_000,
        )
        return (getattr(response, "output_text", None) or ""), selected
    if engine == "opus":
        try:
            import anthropic
        except ImportError as exc:
            raise RuntimeError("Anthropic SDK is not installed") from exc
        selected = model or OPUS_MODEL
        client = anthropic.Anthropic(
            api_key=_secret("ANTHROPIC_API_KEY", "MY_PYTHON_SCRIPT_KEY")
        )
        response = client.messages.create(
            model=selected,
            max_tokens=12_000,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "\n".join(
            block.text for block in response.content
            if getattr(block, "type", "") == "text" and getattr(block, "text", "")
        )
        return text, selected
    raise RuntimeError(f"Unsupported judgement engine: {engine}")


def _assumptions_candidates(report_dir: Path, ticker: str) -> list[Path]:
    names = [f"{ticker}_model_assumptions.json"]
    upper_name = f"{ticker.upper()}_model_assumptions.json"
    if upper_name not in names:
        names.append(upper_name)
    return [root / name for root in (report_dir, report_dir.parent) for name in names]


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def _decision_summary(payload: dict[str, Any]) -> str:
    decisions = derive_inclusion_decisions(payload)
    included = sum(value for drug in decisions.values() for value in drug.values())
    excluded = sum(not value for drug in decisions.values() for value in drug.values())
    return f"{included} included, {excluded} excluded drug/indication pairs"


def main() -> int:
    parser = argparse.ArgumentParser(description="LLM calibration pass for DCF model assumptions")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--company-name", required=True)
    parser.add_argument("--report-dir", required=True)
    parser.add_argument("--engine", choices=["auto", "openai", "opus", "none"], default="auto")
    parser.add_argument("--model", help="Provider-specific model override")
    parser.add_argument("--approval-mode", choices=["auto", "draft"], default="auto")
    parser.add_argument("--approve-draft", action="store_true",
                        help="Validate and promote an existing draft; does not call an LLM")
    parser.add_argument("--max-input-chars", type=int, default=DEFAULT_MAX_INPUT_CHARS)
    args = parser.parse_args()

    ticker = args.ticker.strip()
    report_dir = Path(args.report_dir)
    manifest, reports = load_program_manifest(report_dir, ticker)
    if not manifest:
        print("  judgement skipped: no parser-compatible drug/indication manifest")
        return 0

    draft_path = report_dir / f"{ticker}_model_assumptions_draft.json"
    approved_target = report_dir / f"{ticker}_model_assumptions.json"
    existing_approved = next(
        (path for path in _assumptions_candidates(report_dir, ticker) if path.exists()), None
    )

    if args.approve_draft:
        if existing_approved:
            print(f"  approval refused: approved assumptions already exist at {existing_approved}")
            return 2
        if not draft_path.exists():
            print(f"  approval refused: draft not found at {draft_path}")
            return 2
        try:
            payload = normalize_model_assumptions(
                json.loads(draft_path.read_text(encoding="utf-8")), manifest
            )
        except (OSError, json.JSONDecodeError, AssumptionsValidationError) as exc:
            print(f"  approval refused: invalid draft: {exc}")
            return 2
        _write_json(approved_target, payload)
        print(f"  approved assumptions -> {approved_target} ({_decision_summary(payload)})")
        return 0

    if existing_approved:
        print(f"  judgement skipped: preserving approved/manual assumptions at {existing_approved}")
        return 0

    engine = choose_engine(args.engine)
    if engine is None:
        print("  judgement skipped: no requested high-capability LLM API key; existing assumptions fallback remains active")
        return 0

    try:
        context = build_context(report_dir, ticker, manifest, reports)
        prompt = build_prompt(ticker, args.company_name, manifest, context)
        prompt += database_sync_prompt(scan_research_context(ticker))
        if len(prompt) > args.max_input_chars:
            raise RuntimeError(
                f"locked judgement input is {len(prompt):,} chars, above guard "
                f"{args.max_input_chars:,}; no reports were silently truncated"
            )
        raw, selected_model = call_llm(engine, prompt, args.model)
        raw_payload = extract_json_object(raw)
        database_updates = raw_payload.pop("database_updates", [])
        stored = upsert_research_facts(database_updates, ticker)
        print(f"  database scan active: stored/refreshed {len(stored)} facts")
        payload = normalize_model_assumptions(raw_payload, manifest)
        payload["source"] = (
            f"AUTO_DCF critical LLM judgement draft ({engine}/{selected_model}), {date.today()}"
        )
    except Exception as exc:
        print(f"  judgement failed safely; existing assumptions/report fallbacks remain active: {exc}")
        return 2

    _write_json(draft_path, payload)
    print(f"  validated judgement draft -> {draft_path} ({_decision_summary(payload)})")
    if args.approval_mode == "auto":
        appeared_approved = next(
            (path for path in _assumptions_candidates(report_dir, ticker) if path.exists()), None
        )
        if appeared_approved:
            print(f"  auto-promotion skipped: approved assumptions appeared at {appeared_approved}")
        else:
            _write_json(approved_target, payload)
            print(f"  auto-promoted new assumptions -> {approved_target}")
    else:
        print("  human approval required; review the draft, then rerun with --approve-draft")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
