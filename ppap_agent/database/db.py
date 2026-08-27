"""Database access layer for PPAP agent tools."""

from __future__ import annotations

import json
import sqlite3
from datetime import date
from pathlib import Path

import os

DEFAULT_DB = Path(__file__).resolve().parent.parent / "data" / "ppap_synthetic.db"


def resolve_db_path(db_path: Path | None = None) -> Path:
    if db_path:
        return db_path
    env = os.getenv("PPAP_DB_PATH", "")
    return Path(env) if env else DEFAULT_DB


def get_connection(db_path: Path | None = None) -> sqlite3.Connection:
    path = resolve_db_path(db_path)
    if not path.exists():
        from ppap_agent.database.seed import seed_database
        seed_database(path)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    return conn


def list_pending_ppaps(db_path: Path | None = None) -> list[dict]:
    conn = get_connection(db_path)
    rows = conn.execute(
        """SELECT p.id, p.part_number, pt.description, s.name as supplier_name,
                  p.submission_level, p.priority, p.received_date, p.due_date, p.status
           FROM ppap_submissions p
           JOIN parts pt ON p.part_number = pt.part_number
           JOIN suppliers s ON p.supplier_code = s.code
           WHERE p.status = 'pending'
           ORDER BY
             CASE p.priority WHEN 'critical' THEN 0 WHEN 'high' THEN 1 WHEN 'normal' THEN 2 ELSE 3 END,
             p.due_date ASC"""
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_ppap_package(ppap_id: str, db_path: Path | None = None) -> dict:
    conn = get_connection(db_path)
    sub = conn.execute(
        """SELECT p.*, pt.description, pt.customer, pt.commodity, s.name as supplier_name, s.quality_rating
           FROM ppap_submissions p
           JOIN parts pt ON p.part_number = pt.part_number
           JOIN suppliers s ON p.supplier_code = s.code
           WHERE p.id = ?""",
        (ppap_id,),
    ).fetchone()
    if not sub:
        conn.close()
        raise ValueError(f"PPAP {ppap_id} not found")

    elements = conn.execute(
        "SELECT * FROM ppap_elements WHERE ppap_id = ? ORDER BY element_number",
        (ppap_id,),
    ).fetchall()

    dimensions = conn.execute(
        """SELECT dr.*, ds.nominal, ds.tolerance_plus, ds.tolerance_minus, ds.critical
           FROM dimensional_results dr
           LEFT JOIN drawing_specs ds ON dr.characteristic = ds.characteristic
               AND ds.part_number = (SELECT part_number FROM ppap_submissions WHERE id = ?)
           WHERE dr.ppap_id = ?""",
        (ppap_id, ppap_id),
    ).fetchall()

    inbox = conn.execute(
        "SELECT * FROM inbox_messages WHERE ppap_id = ? ORDER BY received_at DESC",
        (ppap_id,),
    ).fetchall()

    conn.close()
    return {
        "submission": dict(sub),
        "elements": [dict(e) for e in elements],
        "dimensions": [dict(d) for d in dimensions],
        "inbox": [dict(m) for m in inbox],
    }


def get_aiag_rules(db_path: Path | None = None) -> list[dict]:
    conn = get_connection(db_path)
    rows = conn.execute("SELECT * FROM aiag_rules ORDER BY rule_code").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_decision(
    ppap_id: str,
    decision: str,
    risk_band: str,
    risk_score: float,
    reasons: list[str],
    actions: list[str],
    agent_version: str = "0.1.0",
    db_path: Path | None = None,
) -> None:
    conn = get_connection(db_path)
    conn.execute(
        """INSERT INTO review_decisions
           (ppap_id, decision, risk_band, risk_score, reasons, actions, created_at, agent_version)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            ppap_id, decision, risk_band, risk_score,
            json.dumps(reasons), json.dumps(actions),
            date.today().isoformat(), agent_version,
        ),
    )
    conn.execute(
        "UPDATE ppap_submissions SET status = ? WHERE id = ?",
        (decision, ppap_id),
    )
    conn.commit()
    conn.close()


def get_review_history(ppap_id: str | None = None, db_path: Path | None = None) -> list[dict]:
    conn = get_connection(db_path)
    if ppap_id:
        rows = conn.execute(
            "SELECT * FROM review_decisions WHERE ppap_id = ? ORDER BY created_at DESC",
            (ppap_id,),
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM review_decisions ORDER BY created_at DESC"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]
