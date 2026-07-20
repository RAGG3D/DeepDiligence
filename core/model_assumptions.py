"""Shared validation and inclusion rules for model assumptions.

The public JSON schema intentionally stays compatible with the hand-authored
``{TICKER}_model_assumptions.json`` files already consumed by Pipeline and
Scenarios.  Inclusion is encoded without a new field: a drug/indication whose
base, bull, and bear peaks are all exactly zero is excluded from the model.
Anything else is included.  Older/partial assumptions remain backward
compatible because only an explicit three-zero specification excludes a row.
"""

from __future__ import annotations

import copy
import json
import math
import re
from typing import Any, Mapping, Sequence


RATINGS = {"BIC", "T1", "AVG"}
SCHEMA_KEYS = {
    "source",
    "economic_share",
    "ratings",
    "market_share",
    "market_share_notes",
}
PEAK_KEYS = ("base_peak", "bull_peak", "bear_peak")


class AssumptionsValidationError(ValueError):
    """The LLM assumptions payload does not satisfy the locked schema."""


def _norm_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _canonical_index(values: Sequence[str], path: str) -> dict[str, str]:
    index: dict[str, str] = {}
    for value in values:
        text = str(value).strip()
        key = _norm_key(text)
        if not key:
            raise AssumptionsValidationError(f"{path}: blank key")
        if key in index and index[key] != text:
            raise AssumptionsValidationError(
                f"{path}: ambiguous normalized keys {index[key]!r} and {text!r}"
            )
        index[key] = text
    return index


def _require_dict(value: Any, path: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        raise AssumptionsValidationError(f"{path}: expected object")
    return value


def _fraction(value: Any, path: str, *, allow_zero: bool) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise AssumptionsValidationError(
            f"{path}: expected a JSON number expressed as a 0-1 fraction"
        )
    number = float(value)
    lower_ok = number >= 0.0 if allow_zero else number > 0.0
    if not math.isfinite(number) or not lower_ok or number > 1.0:
        bound = "0-1" if allow_zero else "greater than 0 and at most 1"
        raise AssumptionsValidationError(f"{path}: must be {bound}, got {value!r}")
    return number


def _canonical_object(
    value: Any,
    expected: Mapping[str, str],
    path: str,
) -> dict[str, Any]:
    obj = _require_dict(value, path)
    out: dict[str, Any] = {}
    seen: set[str] = set()
    for raw_key, item in obj.items():
        normalized = _norm_key(raw_key)
        canonical = expected.get(normalized)
        if canonical is None:
            raise AssumptionsValidationError(f"{path}: unexpected key {raw_key!r}")
        if canonical in seen:
            raise AssumptionsValidationError(f"{path}: duplicate key {raw_key!r}")
        seen.add(canonical)
        out[canonical] = item
    missing = [key for key in expected.values() if key not in seen]
    if missing:
        raise AssumptionsValidationError(f"{path}: missing keys {missing}")
    return out


def extract_json_object(text: str) -> dict[str, Any]:
    """Extract one assumptions object from a raw/fenced LLM response."""
    candidates = [m.strip() for m in re.findall(
        r"```(?:json)?\s*(.*?)```", text or "", flags=re.I | re.S
    )]
    candidates.append((text or "").strip())
    decoder = json.JSONDecoder()
    for candidate in candidates:
        try:
            value = json.loads(candidate)
            if isinstance(value, dict):
                return value
        except Exception:
            pass
        for match in re.finditer(r"\{", candidate):
            try:
                value, _end = decoder.raw_decode(candidate[match.start():])
            except Exception:
                continue
            if isinstance(value, dict) and "market_share" in value:
                return value
    raise AssumptionsValidationError("response contains no parseable assumptions JSON object")


def normalize_model_assumptions(
    payload: Any,
    expected_programs: Mapping[str, Sequence[str]],
) -> dict[str, Any]:
    """Strictly validate and canonicalize an LLM assumptions payload.

    ``expected_programs`` is the report-derived ``{drug: [indication, ...]}``
    manifest.  The LLM must cover every pair exactly once and may not introduce
    an extra drug or indication.  This is deliberately stricter than the
    backwards-compatible loaders used for older hand-authored files.
    """
    data = _require_dict(payload, "assumptions")
    keys = set(data)
    if keys != SCHEMA_KEYS:
        missing = sorted(SCHEMA_KEYS - keys)
        extra = sorted(keys - SCHEMA_KEYS)
        raise AssumptionsValidationError(
            f"assumptions: schema keys mismatch; missing={missing}, extra={extra}"
        )

    source = data.get("source")
    if not isinstance(source, str) or not source.strip():
        raise AssumptionsValidationError("source: expected a non-empty string")

    drug_index = _canonical_index(list(expected_programs), "program manifest")
    canonical_programs: dict[str, list[str]] = {}
    for raw_drug, raw_indications in expected_programs.items():
        drug = drug_index[_norm_key(raw_drug)]
        indications = [str(ind).strip() for ind in raw_indications]
        if not indications:
            raise AssumptionsValidationError(f"program manifest.{drug}: no indications")
        _canonical_index(indications, f"program manifest.{drug}")
        canonical_programs[drug] = indications

    economic = _canonical_object(data["economic_share"], drug_index, "economic_share")
    ratings = _canonical_object(data["ratings"], drug_index, "ratings")
    market_share = _canonical_object(data["market_share"], drug_index, "market_share")
    notes = _canonical_object(data["market_share_notes"], drug_index, "market_share_notes")

    normalized: dict[str, Any] = {
        "source": source.strip(),
        "economic_share": {},
        "ratings": {},
        "market_share": {},
        "market_share_notes": {},
    }
    included_count = 0
    for drug, indications in canonical_programs.items():
        normalized["economic_share"][drug] = _fraction(
            economic[drug], f"economic_share.{drug}", allow_zero=False
        )
        ind_index = _canonical_index(indications, f"program manifest.{drug}")
        drug_ratings = _canonical_object(ratings[drug], ind_index, f"ratings.{drug}")
        drug_shares = _canonical_object(
            market_share[drug], ind_index, f"market_share.{drug}"
        )
        normalized["ratings"][drug] = {}
        normalized["market_share"][drug] = {}

        note = notes[drug]
        if not isinstance(note, str) or not note.strip():
            raise AssumptionsValidationError(
                f"market_share_notes.{drug}: expected a non-empty string"
            )
        note_upper = note.upper()
        if "INCLUDE" not in note_upper and "EXCLUDE" not in note_upper:
            raise AssumptionsValidationError(
                f"market_share_notes.{drug}: must state INCLUDE and/or EXCLUDE decisions"
            )
        for indication in indications:
            decision_pattern = re.compile(
                rf"\b(?:INCLUDE|EXCLUDE)\s+{re.escape(indication)}(?![A-Za-z0-9])",
                flags=re.I,
            )
            if not decision_pattern.search(note):
                raise AssumptionsValidationError(
                    f"market_share_notes.{drug}: missing explicit decision for {indication}"
                )
        normalized["market_share_notes"][drug] = note.strip()

        for indication in indications:
            rating = drug_ratings[indication]
            if rating not in RATINGS:
                raise AssumptionsValidationError(
                    f"ratings.{drug}.{indication}: expected one of {sorted(RATINGS)}, got {rating!r}"
                )
            normalized["ratings"][drug][indication] = rating

            spec = _require_dict(
                drug_shares[indication], f"market_share.{drug}.{indication}"
            )
            if set(spec) != set(PEAK_KEYS):
                missing = sorted(set(PEAK_KEYS) - set(spec))
                extra = sorted(set(spec) - set(PEAK_KEYS))
                raise AssumptionsValidationError(
                    f"market_share.{drug}.{indication}: peak keys mismatch; "
                    f"missing={missing}, extra={extra}"
                )
            peaks = {
                key: _fraction(
                    spec[key], f"market_share.{drug}.{indication}.{key}", allow_zero=True
                )
                for key in PEAK_KEYS
            }
            bear = peaks["bear_peak"]
            base = peaks["base_peak"]
            bull = peaks["bull_peak"]
            if not (bear <= base <= bull):
                raise AssumptionsValidationError(
                    f"market_share.{drug}.{indication}: require bear <= base <= bull"
                )
            if any(peaks.values()):
                included_count += 1
            normalized["market_share"][drug][indication] = peaks

    if included_count == 0:
        raise AssumptionsValidationError(
            "market_share: guard rejected a draft that excludes every program"
        )
    return normalized


def derive_inclusion_decisions(assumptions: Mapping[str, Any]) -> dict[str, dict[str, bool]]:
    """Return ``True`` unless a pair has an explicit valid three-zero spec."""
    decisions: dict[str, dict[str, bool]] = {}
    market_share = assumptions.get("market_share") if isinstance(assumptions, dict) else None
    if not isinstance(market_share, dict):
        return decisions
    for drug, indications in market_share.items():
        if not isinstance(indications, dict):
            continue
        decisions[str(drug)] = {}
        for indication, spec in indications.items():
            excluded = False
            if isinstance(spec, dict) and all(key in spec for key in PEAK_KEYS):
                values = [spec[key] for key in PEAK_KEYS]
                excluded = all(
                    not isinstance(value, bool)
                    and isinstance(value, (int, float))
                    and float(value) == 0.0
                    for value in values
                )
            decisions[str(drug)][str(indication)] = not excluded
    return decisions


def filter_assets_by_assumptions(
    assets: Sequence[Any],
    assumptions: Mapping[str, Any],
) -> tuple[list[Any], list[tuple[str, str]]]:
    """Remove only explicitly excluded drug/indication pairs from parsed assets.

    Returned assets are shallow copies, so callers can continue scaling curves
    without changing the original parse result.  Missing/partial legacy
    assumptions never exclude anything.
    """
    decisions = derive_inclusion_decisions(assumptions)
    decision_index = {_norm_key(drug): ind_map for drug, ind_map in decisions.items()}
    out: list[Any] = []
    excluded: list[tuple[str, str]] = []
    for asset in assets:
        drug = str(getattr(asset, "name", ""))
        raw_decisions = decision_index.get(_norm_key(drug), {})
        ind_decisions = {_norm_key(ind): include for ind, include in raw_decisions.items()}
        shares = getattr(asset, "market_shares", {}) or {}
        kept_indications = [
            ind for ind in shares if ind_decisions.get(_norm_key(ind), True)
        ]
        for indication in shares:
            if indication not in kept_indications:
                excluded.append((drug, indication))
        if shares and not kept_indications:
            continue

        clone = copy.copy(asset)
        clone.market_shares = {
            ind: curve for ind, curve in shares.items() if ind in kept_indications
        }
        clone.bull_shares = {
            ind: curve for ind, curve in (getattr(asset, "bull_shares", {}) or {}).items()
            if ind in kept_indications
        }
        clone.bear_shares = {
            ind: curve for ind, curve in (getattr(asset, "bear_shares", {}) or {}).items()
            if ind in kept_indications
        }
        clone.indications = list(kept_indications) if shares else list(
            getattr(asset, "indications", []) or []
        )

        def _catalyst_is_included(catalyst: Any) -> bool:
            catalyst_key = _norm_key(getattr(catalyst, "indication", ""))
            if catalyst_key in ind_decisions:
                return ind_decisions[catalyst_key]
            # Research reports sometimes spell a catalyst indication as a long
            # name while Chapter 3 uses its abbreviation. Apply containment only
            # when it resolves to exactly one known program; ambiguous text is
            # kept rather than accidentally suppressing a real catalyst.
            matches = [
                include for indication_key, include in ind_decisions.items()
                if min(len(catalyst_key), len(indication_key)) >= 3
                and (catalyst_key in indication_key or indication_key in catalyst_key)
            ]
            return matches[0] if len(matches) == 1 else True

        clone.catalysts = [
            catalyst for catalyst in (getattr(asset, "catalysts", []) or [])
            if _catalyst_is_included(catalyst)
        ]
        out.append(clone)
    return out, excluded
