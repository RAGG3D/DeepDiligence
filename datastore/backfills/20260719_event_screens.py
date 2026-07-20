#!/usr/bin/env python3
"""Store price-blind clinical-disclosure facts from validated event screens."""
from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from datastore.research_fact_store import upsert_research_facts


ARTIFACTS = (
    REPO / "artifacts" / "event_screens" / "esmo2025.json",
    REPO / "artifacts" / "event_screens" / "aacr2025.json",
)


def build_fact(payload: dict, row: dict) -> dict:
    metrics = ", ".join(row["important_metrics"])
    return {
        "subject": f"{row['ticker']} clinical disclosure at {payload['event']}",
        "indication": "Event-wide clinical-data screen",
        "metric_group": "event",
        "metric": "quantitative_clinical_data_disclosure",
        "value": f"Qualified price-blind screen; reported metrics: {metrics}",
        "unit": "event qualification",
        "population": "Human clinical study participants",
        "as_of_date": payload["event_dates"]["end"],
        "source_url": row["official_clinical_sources"][0],
        "source_kind": "conference_abstract",
        "classification": "Reported Fact",
        "status": "reported",
    }


def main() -> None:
    totals: dict[str, int] = {}
    for artifact in ARTIFACTS:
        payload = json.loads(artifact.read_text(encoding="utf-8"))
        for row in payload["candidates"]:
            if not row["eligible"]:
                continue
            ticker = row["ticker"]
            stored = upsert_research_facts([build_fact(payload, row)], ticker)
            totals[ticker] = totals.get(ticker, 0) + len(stored)
    print("event-screen clinical facts:", totals)


if __name__ == "__main__":
    main()
