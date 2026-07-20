#!/usr/bin/env python3
"""Durable incremental fact store shared by every research workflow.

The JSON seed is the system-of-record so a full DuckDB rebuild preserves new
facts.  DuckDB is updated immediately when available.  Facts are append-only by
data cut/source; a newer disclosure does not erase historical evidence.
"""
from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import duckdb


HERE = Path(__file__).resolve().parent
DEFAULT_DB = HERE / "dd.duckdb"
DEFAULT_SEED = HERE / "seed" / "research_facts.json"

FACT_COLUMNS = (
    "fact_id", "context_ticker", "subject", "indication", "comparator",
    "metric_group", "metric", "value", "unit", "population", "dose",
    "as_of_date", "source_url", "source_kind", "classification", "status",
    "retrieved_at",
)

REQUIRED_COMPETITOR_FIELDS = {
    "orr": ("efficacy", {"orr"}),
    "cr": ("efficacy", {"cr"}),
    "pr": ("efficacy", {"pr"}),
    "mpfs": ("survival", {"median_pfs", "mpfs", "pfs", "pfs_rate"}),
    "mos": ("survival", {"median_os", "mos", "os", "os_rate"}),
    "safety": ("safety", {"safety_overview"}),
    "side_effects": ("safety", {"adverse_event", "side_effect"}),
    "side_effect_rates": ("safety", {"adverse_event_rate", "side_effect_rate"}),
    "dose": ("dose", {"dose", "regimen"}),
    "price_or_forecast": ("price", {"actual_price", "forecast_price"}),
}


def _now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _fact_id(item: dict[str, Any]) -> str:
    key = "|".join(str(item.get(field) or "") for field in (
        "context_ticker", "subject", "indication", "comparator",
        "metric_group", "metric", "as_of_date", "source_url",
    ))
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:24]


def normalize_fact(item: dict[str, Any], context_ticker: str) -> dict[str, Any]:
    fact = {key: item.get(key) for key in FACT_COLUMNS}
    fact["context_ticker"] = str(item.get("context_ticker") or context_ticker).upper()
    fact["subject"] = str(item.get("subject") or "").strip()
    fact["metric_group"] = str(item.get("metric_group") or "").strip().lower()
    fact["metric"] = str(item.get("metric") or "").strip().lower()
    fact["as_of_date"] = str(item.get("as_of_date") or "").strip()
    fact["source_url"] = str(item.get("source_url") or "").strip()
    fact["classification"] = str(
        item.get("classification") or "Reported Fact"
    ).strip()
    fact["status"] = str(item.get("status") or "reported").strip().lower()
    fact["retrieved_at"] = str(item.get("retrieved_at") or _now())
    for optional in ("indication", "comparator", "value", "unit", "population", "dose", "source_kind"):
        value = item.get(optional)
        fact[optional] = None if value in (None, "") else str(value).strip()
    if not fact["subject"] or not fact["metric_group"] or not fact["metric"]:
        raise ValueError(f"research fact missing subject/metric fields: {item}")
    if not fact["as_of_date"] or not fact["source_url"].startswith(("http://", "https://")):
        raise ValueError(f"research fact missing as-of/source URL: {item}")
    if fact["status"] not in {"reported", "estimated", "unavailable", "conflict"}:
        raise ValueError(f"invalid research fact status: {fact['status']}")
    fact["fact_id"] = str(item.get("fact_id") or _fact_id(fact))
    return fact


def _load_seed(seed_path: Path) -> list[dict[str, Any]]:
    if not seed_path.exists():
        return []
    payload = json.loads(seed_path.read_text(encoding="utf-8"))
    return list(payload.get("facts") or []) if isinstance(payload, dict) else []


def _ensure_table(connection: duckdb.DuckDBPyConnection) -> None:
    connection.execute("""
        CREATE TABLE IF NOT EXISTS research_fact (
            fact_id VARCHAR PRIMARY KEY, context_ticker VARCHAR NOT NULL,
            subject VARCHAR NOT NULL, indication VARCHAR, comparator VARCHAR,
            metric_group VARCHAR NOT NULL, metric VARCHAR NOT NULL, value VARCHAR,
            unit VARCHAR, population VARCHAR, dose VARCHAR, as_of_date VARCHAR NOT NULL,
            source_url VARCHAR NOT NULL, source_kind VARCHAR,
            classification VARCHAR NOT NULL, status VARCHAR NOT NULL,
            retrieved_at VARCHAR NOT NULL
        )
    """)


def upsert_research_facts(
    items: Iterable[dict[str, Any]],
    context_ticker: str,
    db_path: Path = DEFAULT_DB,
    seed_path: Path = DEFAULT_SEED,
) -> list[dict[str, Any]]:
    normalized = [normalize_fact(item, context_ticker) for item in items]
    if not normalized:
        return []
    existing = {item["fact_id"]: item for item in _load_seed(seed_path)}
    for fact in normalized:
        existing[fact["fact_id"]] = fact
    seed_path.parent.mkdir(parents=True, exist_ok=True)
    seed_path.write_text(
        json.dumps({"facts": sorted(existing.values(), key=lambda x: (
            x["context_ticker"], x["subject"], x["metric"], x["as_of_date"], x["fact_id"]
        ))}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    if db_path.exists():
        con = duckdb.connect(str(db_path))
        try:
            _ensure_table(con)
            con.executemany(
                "INSERT OR REPLACE INTO research_fact VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                [tuple(fact[column] for column in FACT_COLUMNS) for fact in normalized],
            )
        finally:
            con.close()
    return normalized


def scan_research_context(
    context_ticker: str,
    db_path: Path = DEFAULT_DB,
) -> dict[str, Any]:
    ticker = context_ticker.upper()
    facts: list[dict[str, Any]] = []
    peer_rows: list[dict[str, Any]] = []
    if db_path.exists():
        con = duckdb.connect(str(db_path), read_only=True)
        try:
            tables = {row[0] for row in con.execute("SHOW TABLES").fetchall()}
            if "research_fact" in tables:
                rows = con.execute(
                    "SELECT * FROM research_fact WHERE context_ticker=? ORDER BY subject, as_of_date DESC",
                    [ticker],
                ).fetchall()
                columns = [item[0] for item in con.description]
                facts = [dict(zip(columns, row)) for row in rows]
            if {"peer_drug", "peer_metric"}.issubset(tables):
                rows = con.execute("""
                    SELECT d.section, d.drug, d.ticker, m.metric, m.value
                    FROM peer_drug d JOIN peer_metric m USING(section_id, col)
                    WHERE upper(coalesce(d.ticker,''))=?
                       OR upper(d.section) LIKE '%' || ? || '%'
                    ORDER BY d.section, d.drug, m.metric
                    LIMIT 500
                """, [ticker, ticker]).fetchall()
                peer_rows = [
                    dict(zip(("section", "drug", "ticker", "metric", "value"), row))
                    for row in rows
                ]
        finally:
            con.close()
    completeness = competitor_completeness(facts)
    return {
        "context_ticker": ticker,
        "fact_count": len(facts),
        "facts": facts[-300:],
        "peer_rows": peer_rows,
        "competitor_completeness": completeness,
    }


def competitor_completeness(facts: Iterable[dict[str, Any]]) -> dict[str, Any]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for fact in facts:
        competitor = str(fact.get("comparator") or "").strip()
        if competitor:
            grouped.setdefault(competitor, []).append(fact)
    result: dict[str, Any] = {}
    for competitor, rows in grouped.items():
        covered = set()
        for label, (group, metrics) in REQUIRED_COMPETITOR_FIELDS.items():
            if any(
                str(row.get("metric_group") or "").lower() == group
                and str(row.get("metric") or "").lower() in metrics
                and str(row.get("status") or "").lower() in {"reported", "estimated", "unavailable", "conflict"}
                for row in rows
            ):
                covered.add(label)
        result[competitor] = {
            "covered": sorted(covered),
            "missing": sorted(set(REQUIRED_COMPETITOR_FIELDS) - covered),
            "complete": len(covered) == len(REQUIRED_COMPETITOR_FIELDS),
        }
    return result


def database_sync_prompt(scan: dict[str, Any]) -> str:
    compact = {
        "context_ticker": scan["context_ticker"],
        "competitor_completeness": scan["competitor_completeness"],
        "recent_facts": scan["facts"][-120:],
        "peer_rows": scan["peer_rows"][:160],
    }
    return f"""

DATABASE SCAN IS ACTIVE (mandatory parallel branch):
{json.dumps(compact, ensure_ascii=False)}

For every new or corrected fact found, add a top-level `database_updates` array
to the requested JSON output; for a markdown output, append one fenced JSON
object containing that array. Each row requires: subject, indication,
comparator (when applicable), metric_group, metric, value, unit, population,
dose, as_of_date, source_url, source_kind, classification, status. Use status
`unavailable` with a source and explanation when an important field is not
reported; never invent it. For every in-scope competitor, maximize coverage of
ORR, CR, PR, mPFS/PFS, mOS/OS, safety overview, named side effects and their
rates, dose/regimen, and actual price or forecast price. Also capture DoR,
EFS, DFS, RFS and other time-to-event endpoints whenever reported. A different
time-to-event endpoint does not replace an explicit reported/unavailable record
for mPFS/PFS and mOS/OS. Keep actual and forecast price separate.
Primary-source conflicts must be stored as status `conflict`, not silently
resolved. Do not repeat unchanged rows merely to populate the array.
"""


def extract_database_updates(text: str) -> list[dict[str, Any]]:
    candidates = re.findall(r"```(?:json)?\s*(.*?)```", text, flags=re.S | re.I) + [text]
    updates: list[dict[str, Any]] = []
    for candidate in candidates:
        # Decode complete JSON values.  A non-greedy brace regex truncates as
        # soon as it meets the first nested object in database_updates.
        decoder = json.JSONDecoder()
        offsets = [match.start() for match in re.finditer(r"\{", candidate)]
        for offset in offsets:
            try:
                payload, _ = decoder.raw_decode(candidate[offset:])
            except json.JSONDecodeError:
                continue
            if isinstance(payload, dict) and isinstance(payload.get("database_updates"), list):
                updates.extend(item for item in payload["database_updates"] if isinstance(item, dict))
    unique = {}
    for item in updates:
        unique[json.dumps(item, sort_keys=True, ensure_ascii=False)] = item
    return list(unique.values())
