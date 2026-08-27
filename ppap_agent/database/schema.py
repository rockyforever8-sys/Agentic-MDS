"""SQLite schema for synthetic PPAP quality data."""

from __future__ import annotations

import sqlite3
from pathlib import Path

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS suppliers (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    code TEXT NOT NULL UNIQUE,
    tier INTEGER NOT NULL,
    quality_rating REAL
);

CREATE TABLE IF NOT EXISTS parts (
    id INTEGER PRIMARY KEY,
    part_number TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL,
    customer TEXT NOT NULL,
    commodity TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS ppap_submissions (
    id TEXT PRIMARY KEY,
    part_number TEXT NOT NULL,
    supplier_code TEXT NOT NULL,
    submission_level INTEGER NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    received_date TEXT NOT NULL,
    due_date TEXT NOT NULL,
    priority TEXT NOT NULL DEFAULT 'normal',
    engineer_assigned TEXT,
    FOREIGN KEY (part_number) REFERENCES parts(part_number),
    FOREIGN KEY (supplier_code) REFERENCES suppliers(code)
);

CREATE TABLE IF NOT EXISTS ppap_elements (
    id INTEGER PRIMARY KEY,
    ppap_id TEXT NOT NULL,
    element_number INTEGER NOT NULL,
    element_name TEXT NOT NULL,
    required INTEGER NOT NULL,
    submitted INTEGER NOT NULL DEFAULT 0,
    compliant INTEGER NOT NULL DEFAULT 0,
    file_ref TEXT,
    notes TEXT,
    FOREIGN KEY (ppap_id) REFERENCES ppap_submissions(id),
    UNIQUE(ppap_id, element_number)
);

CREATE TABLE IF NOT EXISTS drawing_specs (
    id INTEGER PRIMARY KEY,
    part_number TEXT NOT NULL,
    revision TEXT NOT NULL,
    characteristic TEXT NOT NULL,
    nominal REAL NOT NULL,
    tolerance_plus REAL NOT NULL,
    tolerance_minus REAL NOT NULL,
    unit TEXT NOT NULL,
    critical INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (part_number) REFERENCES parts(part_number)
);

CREATE TABLE IF NOT EXISTS dimensional_results (
    id INTEGER PRIMARY KEY,
    ppap_id TEXT NOT NULL,
    characteristic TEXT NOT NULL,
    measured REAL NOT NULL,
    unit TEXT NOT NULL,
    within_spec INTEGER,
    FOREIGN KEY (ppap_id) REFERENCES ppap_submissions(id)
);

CREATE TABLE IF NOT EXISTS aiag_rules (
    id INTEGER PRIMARY KEY,
    rule_code TEXT NOT NULL UNIQUE,
    element_number INTEGER,
    category TEXT NOT NULL,
    description TEXT NOT NULL,
    severity TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS review_decisions (
    id INTEGER PRIMARY KEY,
    ppap_id TEXT NOT NULL,
    decision TEXT NOT NULL,
    risk_band TEXT NOT NULL,
    risk_score REAL NOT NULL,
    reasons TEXT NOT NULL,
    actions TEXT NOT NULL,
    created_at TEXT NOT NULL,
    agent_version TEXT NOT NULL,
    FOREIGN KEY (ppap_id) REFERENCES ppap_submissions(id)
);

CREATE TABLE IF NOT EXISTS inbox_messages (
    id INTEGER PRIMARY KEY,
    ppap_id TEXT NOT NULL,
    subject TEXT NOT NULL,
    sender TEXT NOT NULL,
    received_at TEXT NOT NULL,
    read INTEGER NOT NULL DEFAULT 0,
    body TEXT,
    FOREIGN KEY (ppap_id) REFERENCES ppap_submissions(id)
);
"""


def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA_SQL)
    conn.commit()
    return conn
