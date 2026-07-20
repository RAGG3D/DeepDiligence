#!/usr/bin/env python3
"""Durable Catalyst run/snapshot store.

This database is intentionally separate from ``dd.duckdb`` because the main
datastore builder recreates that file. Analyst-entered Catalyst snapshots must
survive every model/datastore rebuild.
"""
from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

import duckdb

DB_PATH = Path(__file__).resolve().parent / "catalyst_history.duckdb"


DDL = """
CREATE TABLE IF NOT EXISTS catalyst_run (
    run_id VARCHAR PRIMARY KEY,
    ticker VARCHAR NOT NULL,
    status VARCHAR NOT NULL,
    event_date VARCHAR,
    event_name VARCHAR,
    drug VARCHAR,
    indication VARCHAR,
    relevant_targets_json VARCHAR NOT NULL,
    research_json VARCHAR NOT NULL,
    workbook_path VARCHAR NOT NULL,
    created_at TIMESTAMP NOT NULL,
    closed_at TIMESTAMP
);
CREATE TABLE IF NOT EXISTS catalyst_snapshot (
    snapshot_id VARCHAR PRIMARY KEY,
    run_id VARCHAR NOT NULL,
    ticker VARCHAR NOT NULL,
    captured_at TIMESTAMP NOT NULL,
    workbook_sha256 VARCHAR NOT NULL,
    workbook_blob BLOB NOT NULL,
    catalyst_sheet_xml BLOB NOT NULL,
    catalyst_cells_json VARCHAR NOT NULL,
    price_reaction_json VARCHAR NOT NULL,
    interpretation_json VARCHAR NOT NULL,
    FOREIGN KEY (run_id) REFERENCES catalyst_run(run_id)
);
"""


def connect(path: Path = DB_PATH):
    path.parent.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(str(path))
    con.execute(DDL)
    return con


def start_run(
    ticker: str,
    research: Dict[str, Any],
    workbook_path: str,
    path: Path = DB_PATH,
) -> str:
    run_id = str(uuid.uuid4())
    relevant = research.get("relevant_targets") or []
    con = connect(path)
    try:
        con.execute(
            """INSERT INTO catalyst_run VALUES (?, ?, 'open', ?, ?, ?, ?, ?, ?, ?, ?, NULL)""",
            [
                run_id, ticker.upper(), research.get("event_date"), research.get("event_name"),
                research.get("drug"), research.get("indication"), json.dumps(relevant),
                json.dumps(research, ensure_ascii=False), workbook_path,
                datetime.now(timezone.utc).replace(tzinfo=None),
            ],
        )
    finally:
        con.close()
    return run_id


def latest_open_run(ticker: str, path: Path = DB_PATH) -> Optional[str]:
    con = connect(path)
    try:
        row = con.execute(
            """SELECT run_id FROM catalyst_run WHERE ticker=? AND status='open'
               ORDER BY created_at DESC LIMIT 1""",
            [ticker.upper()],
        ).fetchone()
        return row[0] if row else None
    finally:
        con.close()


def save_snapshot(
    run_id: str,
    ticker: str,
    workbook_sha256: str,
    workbook_blob: bytes,
    sheet_xml: bytes,
    cells: Dict[str, Any],
    price_reaction: Dict[str, Any],
    interpretation: Dict[str, Any],
    path: Path = DB_PATH,
) -> str:
    snapshot_id = str(uuid.uuid4())
    con = connect(path)
    try:
        con.begin()
        con.execute(
            """INSERT INTO catalyst_snapshot VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            [
                snapshot_id, run_id, ticker.upper(),
                datetime.now(timezone.utc).replace(tzinfo=None), workbook_sha256,
                workbook_blob, sheet_xml, json.dumps(cells, ensure_ascii=False),
                json.dumps(price_reaction, ensure_ascii=False),
                json.dumps(interpretation, ensure_ascii=False),
            ],
        )
        con.execute(
            "UPDATE catalyst_run SET status='closed', closed_at=? WHERE run_id=?",
            [datetime.now(timezone.utc).replace(tzinfo=None), run_id],
        )
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
    return snapshot_id
