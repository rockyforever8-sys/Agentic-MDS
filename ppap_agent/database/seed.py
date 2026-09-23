"""Synthetic PPAP dataset for prototype demonstrations."""

from __future__ import annotations

import json
import sqlite3
from datetime import date, timedelta
from pathlib import Path

from ppap_agent.database.schema import init_db

AIAG_RULES = [
    ("AIAG-PSW-001", 18, "PSW", "Part Submission Warrant must be signed by authorized supplier representative", "critical"),
    ("AIAG-DS-002", 2, "Design Records", "Engineering drawing revision must match submission revision", "critical"),
    ("AIAG-EC-003", 3, "Engineering Change", "All open ECNs must be closed or documented with customer approval", "major"),
    ("AIAG-CP-004", 4, "Customer Approval", "Customer-specific approval required for appearance items", "major"),
    ("AIAG-DFMEA-005", 5, "DFMEA", "Design FMEA must cover all special characteristics", "major"),
    ("AIAG-PFMEA-006", 6, "PFMEA", "Process FMEA RPN > 100 must have documented mitigation", "major"),
    ("AIAG-CP-007", 7, "Control Plan", "Control plan must reference all special characteristics from drawing", "critical"),
    ("AIAG-MSA-008", 4, "MSA", "Gage R&R must be < 30% for critical characteristics", "major"),
    ("AIAG-SPC-009", 4, "SPC", "Cpk >= 1.33 required for production characteristics", "major"),
    ("AIAG-MAT-010", 9, "Material", "Material certification must match drawing callout", "critical"),
    ("AIAG-PPAP-011", 10, "Performance", "Performance test results must meet specification limits", "critical"),
    ("AIAG-DIM-012", 9, "Dimensional", "All critical dimensions must be within drawing tolerance", "critical"),
    ("AIAG-APPR-013", 14, "Appearance", "Appearance approval report required for Class A surfaces", "major"),
    ("AIAG-SAMPLE-014", 9, "Sample", "Sample production parts from significant production run", "major"),
    ("AIAG-MASTER-015", 9, "Master Sample", "Master sample retention per customer agreement", "minor"),
]

PPAP_ELEMENTS = [
    (1, "Design Records", True),
    (2, "Authorized Engineering Change Documents", True),
    (3, "Customer Engineering Approval", False),
    (4, "Design FMEA", True),
    (5, "Process Flow Diagram", True),
    (6, "Process FMEA", True),
    (7, "Control Plan", True),
    (8, "Measurement System Analysis Studies", True),
    (9, "Dimensional Results", True),
    (10, "Material / Performance Test Results", True),
    (11, "Initial Process Studies", True),
    (12, "Qualified Laboratory Documentation", True),
    (13, "Appearance Approval Report", False),
    (14, "Sample Production Parts", True),
    (15, "Master Sample", True),
    (16, "Checking Aids", False),
    (17, "Customer-Specific Requirements", True),
    (18, "Part Submission Warrant (PSW)", True),
]

SUPPLIERS = [
    ("Bosch Automotive Systems", "BOS-001", 1, 4.2),
    ("Denso Thermal Systems", "DEN-002", 1, 4.5),
    ("Magna Seating GmbH", "MAG-003", 2, 3.8),
    ("Lear Corporation", "LEA-004", 1, 4.0),
    ("Faurecia Interiors", "FAU-005", 2, 3.5),
    ("Continental AG", "CON-006", 1, 4.3),
    ("ZF Friedrichshafen", "ZF-007", 1, 4.1),
    ("Aptiv Technologies", "APT-008", 1, 3.9),
]

PARTS = [
    ("BRK-CAL-4421", "Brake Caliper Assembly - Front LH", "GM", "Chassis"),
    ("HV-BAT-MOD-88", "HV Battery Module Housing", "Ford", "EV Powertrain"),
    ("STG-COL-2210", "Steering Column Intermediate Shaft", "VW", "Steering"),
    ("DOOR-TRIM-550", "Door Trim Panel - Class A", "Stellantis", "Interior"),
    ("INJ-RAIL-3300", "Fuel Injector Rail Assembly", "Toyota", "Powertrain"),
    ("SUS-ARM-7782", "Front Lower Control Arm", "BMW", "Chassis"),
    ("SEAT-TRACK-901", "Power Seat Track Mechanism", "Rivian", "Interior"),
    ("RAD-END-445", "Radiator End Tank", "Hyundai", "Thermal"),
]

# Scenario definitions: (ppap_suffix, part_idx, supplier_idx, level, priority, scenario_type)
SCENARIOS = [
    ("001", 0, 1, 3, "high", "clean_accept"),       # All good - fast accept
    ("002", 1, 0, 5, "critical", "missing_docs"),   # Missing elements - hold
    ("003", 2, 6, 3, "normal", "dim_out_of_spec"),  # Critical dim OOS - reject
    ("004", 3, 4, 3, "high", "minor_issues"),       # Minor issues - hold with actions
    ("005", 4, 2, 3, "normal", "clean_accept"),     # Clean accept
    ("006", 5, 7, 5, "critical", "cpk_fail"),       # SPC/Cpk failure - reject
    ("007", 6, 3, 3, "normal", "revision_mismatch"),# Drawing revision mismatch - hold
    ("008", 7, 5, 3, "low", "clean_accept"),        # Clean accept
]


def _element_status(scenario: str, element_num: int) -> tuple[bool, bool, str]:
    """Return (submitted, compliant, notes) for an element given scenario."""
    notes = ""

    if scenario == "clean_accept":
        return True, True, "Complete and compliant"

    if scenario == "missing_docs":
        missing = {8, 11, 12}
        if element_num in missing:
            return False, False, "Not submitted"
        return True, True, "Complete"

    if scenario == "dim_out_of_spec":
        if element_num == 9:
            return True, False, "Critical dimension out of tolerance"
        return True, True, "Complete"

    if scenario == "minor_issues":
        if element_num == 13:
            return False, False, "AAR not submitted for Class A surface"
        if element_num == 6:
            return True, False, "PFMEA RPN 112 on OP-40 without mitigation"
        return True, True, "Complete"

    if scenario == "cpk_fail":
        if element_num in {8, 11}:
            return True, False, "Cpk 1.12 below 1.33 requirement"
        return True, True, "Complete"

    if scenario == "revision_mismatch":
        if element_num in {1, 2}:
            return True, False, "Drawing Rev C submitted, customer requires Rev D"
        return True, True, "Complete"

    return True, True, "Complete"


def _dimension_data(scenario: str, part_number: str) -> list[dict]:
    """Generate dimensional results based on scenario."""
    base_dims = {
        "BRK-CAL-4421": [
            {"characteristic": "Bore Diameter", "nominal": 42.0, "tol_p": 0.05, "tol_m": 0.05, "critical": True},
            {"characteristic": "Mounting Hole PCD", "nominal": 98.5, "tol_p": 0.1, "tol_m": 0.1, "critical": False},
            {"characteristic": "Overall Length", "nominal": 156.2, "tol_p": 0.3, "tol_m": 0.3, "critical": False},
        ],
        "HV-BAT-MOD-88": [
            {"characteristic": "Housing Width", "nominal": 320.0, "tol_p": 0.5, "tol_m": 0.5, "critical": True},
            {"characteristic": "Seal Groove Depth", "nominal": 3.2, "tol_p": 0.1, "tol_m": 0.1, "critical": True},
            {"characteristic": "Mounting Boss Height", "nominal": 12.5, "tol_p": 0.2, "tol_m": 0.2, "critical": False},
        ],
        "STG-COL-2210": [
            {"characteristic": "Spline Major Diameter", "nominal": 24.0, "tol_p": 0.03, "tol_m": 0.03, "critical": True},
            {"characteristic": "Overall Length", "nominal": 445.0, "tol_p": 0.5, "tol_m": 0.5, "critical": False},
            {"characteristic": "Runout at Bearing Journal", "nominal": 0.0, "tol_p": 0.05, "tol_m": 0.0, "critical": True},
        ],
        "DOOR-TRIM-550": [
            {"characteristic": "Gap Profile Point A", "nominal": 4.0, "tol_p": 0.5, "tol_m": 0.5, "critical": True},
            {"characteristic": "Flush Profile Point B", "nominal": 0.0, "tol_p": 0.3, "tol_m": 0.3, "critical": True},
        ],
        "INJ-RAIL-3300": [
            {"characteristic": "Rail Bore ID", "nominal": 8.0, "tol_p": 0.02, "tol_m": 0.02, "critical": True},
            {"characteristic": "Injector Port Spacing", "nominal": 35.0, "tol_p": 0.05, "tol_m": 0.05, "critical": True},
        ],
        "SUS-ARM-7782": [
            {"characteristic": "Ball Joint Bore", "nominal": 28.0, "tol_p": 0.05, "tol_m": 0.05, "critical": True},
            {"characteristic": "Bushing Press Fit OD", "nominal": 42.0, "tol_p": 0.03, "tol_m": 0.03, "critical": True},
        ],
        "SEAT-TRACK-901": [
            {"characteristic": "Track Width", "nominal": 18.5, "tol_p": 0.1, "tol_m": 0.1, "critical": False},
            {"characteristic": "Latch Engagement Depth", "nominal": 6.0, "tol_p": 0.15, "tol_m": 0.15, "critical": True},
        ],
        "RAD-END-445": [
            {"characteristic": "Tank Wall Thickness", "nominal": 2.5, "tol_p": 0.15, "tol_m": 0.15, "critical": True},
            {"characteristic": "Inlet Port Diameter", "nominal": 32.0, "tol_p": 0.2, "tol_m": 0.2, "critical": False},
        ],
    }

    dims = base_dims.get(part_number, [])
    results = []

    for d in dims:
        measured = d["nominal"]
        within = True

        if scenario == "dim_out_of_spec" and d["critical"]:
            measured = d["nominal"] + d["tol_p"] + 0.02  # Out of spec high side
            within = False
        elif scenario == "minor_issues" and d["characteristic"].startswith("Gap"):
            measured = d["nominal"] + d["tol_p"] * 0.8  # Near limit but OK
        elif scenario == "cpk_fail":
            measured = d["nominal"] + d["tol_p"] * 0.7  # Near upper limit

        results.append({
            "characteristic": d["characteristic"],
            "nominal": d["nominal"],
            "tolerance_plus": d["tol_p"],
            "tolerance_minus": d["tol_m"],
            "measured": round(measured, 4),
            "unit": "mm",
            "within_spec": within,
            "critical": d["critical"],
        })

    return results


def seed_database(db_path: Path) -> dict:
    """Populate synthetic PPAP data. Returns summary stats."""
    conn = init_db(db_path)
    today = date.today()

    for name, code, tier, rating in SUPPLIERS:
        conn.execute(
            "INSERT OR IGNORE INTO suppliers (name, code, tier, quality_rating) VALUES (?, ?, ?, ?)",
            (name, code, tier, rating),
        )

    for pn, desc, customer, commodity in PARTS:
        conn.execute(
            "INSERT OR IGNORE INTO parts (part_number, description, customer, commodity) VALUES (?, ?, ?, ?)",
            (pn, desc, customer, commodity),
        )

    for rule_code, elem, cat, desc, sev in AIAG_RULES:
        conn.execute(
            "INSERT OR IGNORE INTO aiag_rules (rule_code, element_number, category, description, severity) VALUES (?, ?, ?, ?, ?)",
            (rule_code, elem, cat, desc, sev),
        )

    scenario_map: dict[str, str] = {}

    for suffix, part_idx, sup_idx, level, priority, scenario in SCENARIOS:
        part = PARTS[part_idx]
        supplier = SUPPLIERS[sup_idx]
        ppap_id = f"PPAP-2026-{suffix}"
        scenario_map[ppap_id] = scenario

        received = today - timedelta(days=3 + int(suffix))
        due = today + timedelta(days=max(1, 14 - int(suffix)))

        conn.execute(
            """INSERT OR REPLACE INTO ppap_submissions
               (id, part_number, supplier_code, submission_level, status, received_date, due_date, priority, engineer_assigned)
               VALUES (?, ?, ?, ?, 'pending', ?, ?, ?, 'SQE-Auto-Agent')""",
            (ppap_id, part[0], supplier[1], level, received.isoformat(), due.isoformat(), priority),
        )

        for elem_num, elem_name, required in PPAP_ELEMENTS:
            submitted, compliant, notes = _element_status(scenario, elem_num)
            conn.execute(
                """INSERT OR REPLACE INTO ppap_elements
                   (ppap_id, element_number, element_name, required, submitted, compliant, file_ref, notes)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    ppap_id, elem_num, elem_name, int(required), int(submitted),
                    int(compliant), f"{ppap_id}_elem{elem_num:02d}.pdf" if submitted else None, notes,
                ),
            )

        dims = _dimension_data(scenario, part[0])
        for d in dims:
            conn.execute(
                """INSERT OR REPLACE INTO drawing_specs
                   (part_number, revision, characteristic, nominal, tolerance_plus, tolerance_minus, unit, critical)
                   VALUES (?, 'Rev-C', ?, ?, ?, ?, ?, ?)""",
                (part[0], d["characteristic"], d["nominal"], d["tolerance_plus"],
                 d["tolerance_minus"], d["unit"], int(d.get("critical", False))),
            )
            conn.execute(
                """INSERT OR REPLACE INTO dimensional_results
                   (ppap_id, characteristic, measured, unit, within_spec)
                   VALUES (?, ?, ?, ?, ?)""",
                (ppap_id, d["characteristic"], d["measured"], d["unit"], int(d["within_spec"])),
            )

        conn.execute(
            """INSERT OR REPLACE INTO inbox_messages
               (ppap_id, subject, sender, received_at, read, body)
               VALUES (?, ?, ?, ?, 0, ?)""",
            (
                ppap_id,
                f"PPAP Level {level} Submission - {part[0]}",
                f"{supplier[0]} <ppap@{supplier[1].lower()}.com>",
                received.isoformat(),
                f"Please review attached PPAP package for {part[1]}. Submission level {level}.",
            ),
        )

    conn.commit()
    conn.close()

    summary = {
        "suppliers": len(SUPPLIERS),
        "parts": len(PARTS),
        "ppap_submissions": len(SCENARIOS),
        "aiag_rules": len(AIAG_RULES),
        "scenarios": scenario_map,
    }

    meta_path = db_path.parent / "seed_meta.json"
    meta_path.write_text(json.dumps(summary, indent=2))
    return summary
