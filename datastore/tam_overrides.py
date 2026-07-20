#!/usr/bin/env python3
"""TAM override helpers for ticker-specific serviceable markets.

The main TAM database is built from marketed drug sales.  Early pipeline assets
often need a serviceable indication market that is not yet represented in those
rows, so curated report anchors are stored here as keyed data-center inputs and
published into the same CSVs the workbooks consume.
"""

from __future__ import annotations

import csv
import json
import re
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List


HERE = Path(__file__).resolve().parent
REPO = HERE.parent
SEED_PATH = HERE / "seed" / "tam_overrides.json"
EXPORT_DIR = HERE / "export"
EXCEL_EXPORT_DIR = Path("/mnt/c/Users/yzsun/Desktop/DD/_datastore")

MODEL_START_YEAR = 2010
MODEL_END_YEAR = 2038

DEFAULT_GROUP = {
    "AML": "blood",
    "MDS": "blood",
    "HL": "blood",
    "MM": "blood",
}

DEFAULT_INCIDENCE_RATE = {
    "AML": 0.000066,
    "BTC": 0.000028185349839776108,
    "EP-NEC": 0.00001,
    "LCNEC": 0.000014,
    "MDS": 0.000049,
    "SCLC": 0.000046,
}


def default_group(indication_code: str) -> str:
    return DEFAULT_GROUP.get(indication_code.upper(), "solid")


def interp_anchors(anchors: Dict[int, float], year: int) -> float:
    """CAGR interpolation/backcast across anchor years; hold after last anchor."""
    ys = sorted(int(y) for y in anchors)
    if not ys:
        return 0.0
    if year <= ys[0]:
        if len(ys) >= 2 and anchors[ys[0]] > 0 and anchors[ys[1]] > 0:
            g = (anchors[ys[1]] / anchors[ys[0]]) ** (1.0 / (ys[1] - ys[0]))
            return anchors[ys[0]] * g ** (year - ys[0])
        return anchors[ys[0]]
    if year >= ys[-1]:
        return anchors[ys[-1]]
    for a, b in zip(ys, ys[1:]):
        if a <= year <= b:
            if anchors[a] > 0 and anchors[b] > 0:
                g = (anchors[b] / anchors[a]) ** (1.0 / (b - a))
                return anchors[a] * g ** (year - a)
            return anchors[a] + (anchors[b] - anchors[a]) * (year - a) / (b - a)
    return anchors[ys[-1]]


def normalize_override(item: Dict[str, Any]) -> Dict[str, Any]:
    code = str(item["indication_code"]).strip()
    anchors = {
        int(k): float(v)
        for k, v in (item.get("anchors") or {}).items()
        if str(k).isdigit() and v is not None
    }
    out = dict(item)
    out["indication_code"] = code
    # source_ticker is now part of the override key (and the tam_override PK), so
    # it must always be a non-null, canonical string.
    out["source_ticker"] = str(item.get("source_ticker") or "").strip().upper()
    out["tam_group"] = (item.get("tam_group") or default_group(code)).lower()
    out["anchors"] = {str(k): anchors[k] for k in sorted(anchors)}
    if out.get("incidence_rate") is None and code in DEFAULT_INCIDENCE_RATE:
        out["incidence_rate"] = DEFAULT_INCIDENCE_RATE[code]
    return out


def load_overrides(path: Path = SEED_PATH) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        data = data.get("overrides", [])
    return [normalize_override(item) for item in data]


def save_overrides(overrides: Iterable[Dict[str, Any]], path: Path = SEED_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    items = sorted(
        (normalize_override(item) for item in overrides),
        key=lambda x: (x["indication_code"], x.get("source_ticker", ""), x.get("source_drug", "")),
    )
    path.write_text(json.dumps({"overrides": items}, indent=2, sort_keys=False) + "\n", encoding="utf-8")


def parse_report_tam(report_dir: Path, ticker: str) -> List[Dict[str, Any]]:
    """Parse per-drug research reports into override seed entries."""
    entries: List[Dict[str, Any]] = []
    paths = sorted(report_dir.glob(f"{ticker}_*_research_*.md"))
    latest: Dict[str, Path] = {}
    prefix = f"{ticker}_"
    marker = "_research_"
    for path in paths:
        name = path.stem
        if not name.startswith(prefix) or marker not in name:
            continue
        drug, version = name[len(prefix):].split(marker, 1)
        current = latest.get(drug.upper())
        current_version = current.stem.split(marker, 1)[1] if current else ""
        if current is None or (version, path.stat().st_mtime) > (
            current_version, current.stat().st_mtime
        ):
            latest[drug.upper()] = path
    for path in sorted(latest.values()):
        parts = path.stem.split("_")
        drug_parts: List[str] = []
        for part in parts[1:]:
            if part.lower() == "research":
                break
            drug_parts.append(part)
        drug = "-".join(drug_parts) if drug_parts else parts[1].upper()
        text = path.read_text(encoding="utf-8", errors="ignore")

        for m in re.finditer(
            r"(?m)^##+\s*3\.\d+\s+.*?\(([A-Za-z0-9/+\-]+)\)(.*?)(?=^##+\s*3\.\d+\s|\Z)",
            text,
            re.S,
        ):
            code = m.group(1).strip().upper()
            body = m.group(2)
            tam_section = re.search(r"Addressable Market \(\$MM\)(.*?)(?=\n#|\Z)", body, re.S)
            # Never fall back to the whole indication section: market-share
            # tables also contain Year + numeric columns and were previously
            # misread as $MM TAM anchors when a modelling supplement omitted TAM.
            if not tam_section:
                continue
            scope = tam_section.group(1)
            anchors: Dict[int, float] = {}
            for row in re.finditer(r"\|\s*(\d{4})\s*\|\s*\$?\s*([\d,]+(?:\.\d+)?)", scope):
                anchors[int(row.group(1))] = float(row.group(2).replace(",", ""))
            if not anchors:
                continue

            reasoning = ""
            rmatch = re.search(r"-\s*Reasoning:\s*(.+?)(?=\n\s*\| Year \|)", body, re.S)
            if rmatch:
                reasoning = re.sub(r"\s+", " ", rmatch.group(1)).strip()
            entries.append(
                normalize_override(
                    {
                        "indication_code": code,
                        "tam_group": default_group(code),
                        "source_ticker": ticker.upper(),
                        "source_drug": drug.upper(),
                        "source_file": str(path),
                        "source_note": reasoning[:700],
                        "anchors": {str(k): v for k, v in sorted(anchors.items())},
                        "updated_at": time.strftime("%Y-%m-%d"),
                    }
                )
            )
    return entries


def _override_key(item: Dict[str, Any]) -> tuple:
    """Identity of an override: an (indication_code, source_ticker) pair so that
    two tickers' serviceable markets for the same indication never collide."""
    return (item["indication_code"], item.get("source_ticker", ""))


def upsert_overrides(new_entries: Iterable[Dict[str, Any]], path: Path = SEED_PATH) -> List[Dict[str, Any]]:
    """Merge entries by (indication_code, source_ticker); latest curated entry wins.

    Keying on the source ticker as well as the indication lets a new ticker's
    curated TAM coexist with a prior ticker's, rather than clobbering it (which
    the old indication-only key did — silent cross-ticker contamination).
    """
    existing = {_override_key(item): item for item in load_overrides(path)}
    for item in new_entries:
        norm = normalize_override(item)
        old = existing.get(_override_key(norm))
        if old:
            for key in ("incidence_rate", "incidence_global_annual"):
                if norm.get(key) is None and old.get(key) is not None:
                    norm[key] = old[key]
        existing[_override_key(norm)] = norm
    merged = list(existing.values())
    save_overrides(merged, path)
    return merged


def annualized_rows(
    overrides: Iterable[Dict[str, Any]],
    start_year: int = MODEL_START_YEAR,
    end_year: int = MODEL_END_YEAR,
) -> List[tuple]:
    rows = []
    for item in overrides:
        anchors = {int(k): float(v) for k, v in (item.get("anchors") or {}).items()}
        if not anchors:
            continue
        for year in range(start_year, end_year + 1):
            rows.append(
                (
                    item["indication_code"],
                    item["tam_group"],
                    year,
                    interp_anchors(anchors, year),
                    item.get("source_ticker"),
                    item.get("source_drug"),
                    item.get("source_note"),
                )
            )
    return rows


def incidence_inputs(overrides: Iterable[Dict[str, Any]]) -> Dict[str, tuple[Any, Any]]:
    out: Dict[str, tuple[Any, Any]] = {}
    for item in overrides:
        code = item["indication_code"]
        if "incidence_rate" in item or "incidence_global_annual" in item:
            out[code] = (item.get("incidence_rate"), item.get("incidence_global_annual"))
    return out


def _overlay_csv(path: Path, fieldnames: List[str], key_fields: List[str],
                 override_rows: List[Dict[str, Any]]) -> None:
    rows: List[Dict[str, Any]] = []
    if path.exists():
        with path.open(newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    override_keys = {tuple(str(row[k]) for k in key_fields) for row in override_rows}
    rows = [
        row for row in rows
        if tuple(str(row.get(k, "")) for k in key_fields) not in override_keys
    ]
    rows.extend(override_rows)
    rows.sort(key=lambda r: tuple(r.get(k, "") for k in key_fields))
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def publish_override_csvs(
    overrides: Iterable[Dict[str, Any]],
    export_dirs: Iterable[Path] = (EXPORT_DIR, EXCEL_EXPORT_DIR),
) -> None:
    """Overlay override rows onto published CSVs without requiring DuckDB."""
    items = list(overrides)
    rows = annualized_rows(items)
    tam_rows = [
        {
            "indication_code": code,
            "year": str(year),
            "tam_usd_m": f"{tam:.10g}",
        }
        for code, _group, year, tam, _ticker, _drug, _note in rows
    ]
    group_rows = [
        {
            "tam_group": group,
            "indication_code": code,
            "year": str(year),
            "tam_usd_m": f"{tam:.10g}",
        }
        for code, group, year, tam, _ticker, _drug, _note in rows
    ]
    meta_rows = [
        {
            "indication_code": code,
            "tam_group": group,
            "year": str(year),
            "tam_usd_m": f"{tam:.10g}",
            "source_ticker": ticker or "",
            "source_drug": drug or "",
            "source_note": note or "",
        }
        for code, group, year, tam, ticker, drug, note in rows
    ]
    incidence = incidence_inputs(items)
    incidence_rows = [
        {
            "indication_code": code,
            "incidence_rate": "" if vals[0] is None else vals[0],
            "incidence_global_annual": "" if vals[1] is None else vals[1],
        }
        for code, vals in incidence.items()
    ]

    for out_dir in export_dirs:
        out_dir = Path(out_dir)
        _overlay_csv(
            out_dir / "tam_by_indication_year.csv",
            ["indication_code", "year", "tam_usd_m"],
            ["indication_code", "year"],
            tam_rows,
        )
        _overlay_csv(
            out_dir / "tam_by_group_year.csv",
            ["tam_group", "indication_code", "year", "tam_usd_m"],
            ["tam_group", "indication_code", "year"],
            group_rows,
        )
        _overlay_csv(
            out_dir / "tam_override.csv",
            ["indication_code", "tam_group", "year", "tam_usd_m", "source_ticker", "source_drug", "source_note"],
            ["indication_code", "source_ticker", "year"],
            meta_rows,
        )
        if incidence_rows:
            _overlay_csv(
                out_dir / "param_incidence.csv",
                ["indication_code", "incidence_rate", "incidence_global_annual"],
                ["indication_code"],
                incidence_rows,
            )
