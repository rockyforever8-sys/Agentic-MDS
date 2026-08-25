#!/usr/bin/env python3
"""Pure IMDS decision logic — no Playwright, no secrets.

Used by imds_agent_v2.py and by unit tests / --self-test.
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from typing import Optional

RULE_PACK_VERSION = "rec001-v1"

# Matches both IMDS UI text and the comma-normalized form.
# Character class [(]s[)] is used so Colab/IPython does not treat this as LaTeX.
ERROR_WARNING_RE = re.compile(
    r"(\d+)\s*Error[(]s[)]\s*(?:/|,)\s*(\d+)\s*Warning[(]s[)]",
    re.IGNORECASE,
)
PASSED_RE = re.compile(r"passed all included checks", re.IGNORECASE)

TOOL_FAILURE_MESSAGES = {
    "check failed",
    "check result not found",
    "unknown",
    "extraction_failed",
}

PASS_TOKENS = {"pass", "passed", "ok"}
FAIL_TOKENS = {"fail", "failed"}


@dataclass
class CheckCounts:
    errors: Optional[int]
    warnings: Optional[int]
    passed_all: bool
    parse_ok: bool
    raw: str


@dataclass
class Decision:
    band: str  # GREEN | AMBER | RED
    overall: str  # PASS | HOLD | FAIL (Excel-facing)
    action: str  # accept | hold | reject
    reasons: list[str] = field(default_factory=list)
    reject_text: str = ""
    rule_pack_version: str = RULE_PACK_VERSION
    check_errors: Optional[int] = None
    check_warnings: Optional[int] = None

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=True)


def parse_check_counts(result_msg: str) -> CheckCounts:
    raw = (result_msg or "").strip()
    if not raw:
        return CheckCounts(None, None, False, False, raw)

    passed_all = bool(PASSED_RE.search(raw))
    match = ERROR_WARNING_RE.search(raw)
    if match:
        errors = int(match.group(1))
        warnings = int(match.group(2))
        return CheckCounts(
            errors=errors,
            warnings=warnings,
            passed_all=passed_all or (errors == 0 and warnings == 0),
            parse_ok=True,
            raw=raw,
        )

    if passed_all:
        return CheckCounts(0, 0, True, True, raw)

    return CheckCounts(None, None, False, False, raw)


def is_check_clean(result_msg: str) -> bool:
    """True only when IMDS Check reports zero errors and zero warnings."""
    counts = parse_check_counts(result_msg)
    if not counts.parse_ok:
        return False
    return (counts.errors or 0) == 0 and (counts.warnings or 0) == 0


def _norm_rule(value: str) -> str:
    return (value or "").strip().lower()


def _is_pass(value: str) -> bool:
    return _norm_rule(value) in PASS_TOKENS


def _is_fail(value: str) -> bool:
    return _norm_rule(value) in FAIL_TOKENS


def _is_tool_failure(value: str) -> bool:
    return _norm_rule(value) in TOOL_FAILURE_MESSAGES or not (value or "").strip()


def normalize_mds_id(value: str) -> str:
    return re.sub(r"\s+", "", value or "").lower()


def mds_id_number(value: str) -> str:
    return normalize_mds_id(value).split("/", 1)[0]


def mds_ids_match(opened: str, expected: str, *, require_version: bool = True) -> bool:
    """Confirm the open MDS is the one we scored before Accept/Reject."""
    if not opened or not expected:
        return False
    if opened == "EXTRACTION_FAILED" or expected == "EXTRACTION_FAILED":
        return False
    if normalize_mds_id(opened) == normalize_mds_id(expected):
        return True
    if require_version:
        return False
    return mds_id_number(opened) == mds_id_number(expected) and bool(mds_id_number(opened))


def build_reject_text(
    *,
    mds_id: str,
    reasons: list[str],
    check_result: str,
    rule_pack_version: str = RULE_PACK_VERSION,
) -> str:
    lines = [
        f"IMDS agent auto-reject [{rule_pack_version}]",
        f"MDS: {mds_id}" if mds_id else "MDS: (id not extracted)",
        f"IMDS Check: {check_result}" if check_result else "IMDS Check: (not available)",
        "Reasons:",
    ]
    if reasons:
        lines.extend(f"- {reason}" for reason in reasons)
    else:
        lines.append("- Rec 001 / IMDS Check failed (see Excel trail).")
    lines.append("Please correct and resubmit the MDS.")
    return "\n".join(lines)


def decide_overall(
    *,
    check_result: str,
    recyclate_check: str,
    biocidal_check: str,
    parts_marking_check: str,
    mds_id: str = "",
    require_parts_marking: bool = False,
    rule_pack_version: str = RULE_PACK_VERSION,
) -> Decision:
    """Map Rec 001 + IMDS Check into green / amber / red.

    GREEN: auto-accept candidate (0 errors, 0 warnings, Rec 001 pass)
    AMBER: human only (warnings, missing data, tool failure, parts marking)
    RED: auto-reject candidate (IMDS Check errors or Rec 001 recyclate/biocidal fail)

    Tool/UI failure never auto-rejects — that was a v0 hazard.
    """
    reasons: list[str] = []
    band = "GREEN"
    counts = parse_check_counts(check_result)

    if _is_tool_failure(check_result) or not counts.parse_ok:
        band = "AMBER"
        reasons.append(f"IMDS Check not conclusive: {check_result or '(empty)'}")
    elif (counts.errors or 0) > 0:
        band = "RED"
        reasons.append(f"IMDS Check errors: {counts.errors} Error(s), {counts.warnings} Warning(s)")
    elif (counts.warnings or 0) > 0:
        band = "AMBER"
        reasons.append(f"IMDS Check warnings: {counts.warnings} Warning(s) (human review)")
    elif not counts.passed_all and not is_check_clean(check_result):
        band = "AMBER"
        reasons.append(f"IMDS Check not clean: {check_result}")

    if _is_fail(recyclate_check):
        band = "RED"
        reasons.append("Recyclate question unanswered or failed on a Rec 001 material class")
    elif not _is_pass(recyclate_check):
        if band != "RED":
            band = "AMBER"
        reasons.append(f"Recyclate check not conclusive: {recyclate_check}")

    if _is_fail(biocidal_check):
        band = "RED"
        reasons.append("Biocidal 'Still in production?' unanswered or failed on a Rec 001 material class")
    elif not _is_pass(biocidal_check):
        if band != "RED":
            band = "AMBER"
        reasons.append(f"Biocidal check not conclusive: {biocidal_check}")

    marking = _norm_rule(parts_marking_check)
    if marking in {"no components", "no materials", "n/a", "na"}:
        pass
    elif _is_fail(parts_marking_check):
        if require_parts_marking:
            band = "RED"
            reasons.append("Parts marking empty on a polymer/required component (strict mode)")
        else:
            if band != "RED":
                band = "AMBER"
            reasons.append("Parts marking empty on a required component (held for human)")
    elif not _is_pass(parts_marking_check):
        if band == "GREEN":
            band = "AMBER"
        reasons.append(f"Parts marking check not conclusive: {parts_marking_check}")

    if band == "GREEN":
        overall, action = "PASS", "accept"
        reject_text = ""
        if not reasons:
            reasons.append("IMDS Check 0/0 and Rec 001 recyclate/biocidal passed")
    elif band == "RED":
        overall, action = "FAIL", "reject"
        reject_text = build_reject_text(
            mds_id=mds_id,
            reasons=reasons,
            check_result=check_result,
            rule_pack_version=rule_pack_version,
        )
    else:
        overall, action = "HOLD", "hold"
        reject_text = ""

    return Decision(
        band=band,
        overall=overall,
        action=action,
        reasons=reasons,
        reject_text=reject_text,
        rule_pack_version=rule_pack_version,
        check_errors=counts.errors,
        check_warnings=counts.warnings,
    )


def apply_autonomous_policy(
    decision: Decision,
    *,
    hold_amber: bool = False,
    mds_id: str = "",
    check_result: str = "",
) -> Decision:
    """Live inbox policy: PASS → accept; FAIL → reject; AMBER → reject unless held.

    Tool/UI failure (Check did not run) always stays HOLD so we never reject a
    sheet we could not actually score.
    """
    if decision.action != "hold":
        return decision
    if hold_amber:
        return decision
    if _is_tool_failure(check_result):
        return decision
    reasons = list(decision.reasons) + [
        "Autonomous live run: not a clean PASS, treating as FAIL"
    ]
    return Decision(
        band="RED",
        overall="FAIL",
        action="reject",
        reasons=reasons,
        reject_text=build_reject_text(
            mds_id=mds_id,
            reasons=reasons,
            check_result=check_result,
            rule_pack_version=decision.rule_pack_version,
        ),
        rule_pack_version=decision.rule_pack_version,
        check_errors=decision.check_errors,
        check_warnings=decision.check_warnings,
    )


def kill_switch_active(env_get, kill_file_exists) -> bool:
    """env_get(name) -> str|None; kill_file_exists() -> bool."""
    raw = env_get("IMDS_KILL_SWITCH")
    if raw and str(raw).strip().lower() in {"1", "true", "yes", "on"}:
        return True
    return bool(kill_file_exists())


def env_flag(raw: Optional[str], default: bool) -> bool:
    if raw is None or str(raw).strip() == "":
        return default
    return str(raw).strip().lower() in {"1", "true", "yes", "on"}


def run_self_test() -> int:
    """Return 0 on success. Prints a short fixture table."""
    cases = [
        {
            "name": "slash_zero_is_green",
            "kwargs": {
                "check_result": "0 Error(s) / 0 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
                "mds_id": "1234567890 / 1.0",
            },
            "expect_band": "GREEN",
            "expect_action": "accept",
        },
        {
            "name": "comma_zero_is_green",
            "kwargs": {
                "check_result": "0 Error(s), 0 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
            },
            "expect_band": "GREEN",
            "expect_action": "accept",
        },
        {
            "name": "passed_all_text_is_green",
            "kwargs": {
                "check_result": "The MDS has passed all included checks.",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
            },
            "expect_band": "GREEN",
            "expect_action": "accept",
        },
        {
            "name": "warnings_are_amber",
            "kwargs": {
                "check_result": "0 Error(s) / 2 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
            },
            "expect_band": "AMBER",
            "expect_action": "hold",
        },
        {
            "name": "errors_are_red",
            "kwargs": {
                "check_result": "3 Error(s), 0 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
                "mds_id": "999 / 2.0",
            },
            "expect_band": "RED",
            "expect_action": "reject",
        },
        {
            "name": "check_ui_failure_is_amber_not_reject",
            "kwargs": {
                "check_result": "Check failed",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
            },
            "expect_band": "AMBER",
            "expect_action": "hold",
        },
        {
            "name": "recyclate_fail_is_red",
            "kwargs": {
                "check_result": "The MDS has passed all included checks.",
                "recyclate_check": "FAIL",
                "biocidal_check": "PASS",
                "parts_marking_check": "PASS",
                "mds_id": "111 / 1",
            },
            "expect_band": "RED",
            "expect_action": "reject",
        },
        {
            "name": "parts_marking_fail_is_amber_by_default",
            "kwargs": {
                "check_result": "0 Error(s) / 0 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "FAIL",
            },
            "expect_band": "AMBER",
            "expect_action": "hold",
        },
        {
            "name": "parts_marking_fail_strict_is_red",
            "kwargs": {
                "check_result": "0 Error(s) / 0 Warning(s)",
                "recyclate_check": "PASS",
                "biocidal_check": "PASS",
                "parts_marking_check": "FAIL",
                "require_parts_marking": True,
                "mds_id": "222 / 1",
            },
            "expect_band": "RED",
            "expect_action": "reject",
        },
    ]

    failed = []
    print("case                         band    action")
    print("-" * 52)
    for case in cases:
        decision = decide_overall(**case["kwargs"])
        ok = decision.band == case["expect_band"] and decision.action == case["expect_action"]
        status = "OK" if ok else "FAIL"
        print(f"{case['name']:28} {decision.band:7} {decision.action:7} {status}")
        if decision.band == "RED" and "reject" not in decision.reject_text.lower():
            failed.append(f"{case['name']}: missing reject text")
            ok = False
        if not ok:
            failed.append(
                f"{case['name']}: got band={decision.band} action={decision.action} reasons={decision.reasons}"
            )

    id_ok = mds_ids_match("1234567890 / 1.0", "1234567890/1.0")
    id_bad = mds_ids_match("1234567890 / 1.0", "999 / 1.0")
    id_fail = mds_ids_match("EXTRACTION_FAILED", "123 / 1")
    if not id_ok:
        failed.append("mds_ids_match should accept spaced vs compact ID/version")
    if id_bad:
        failed.append("mds_ids_match should reject a different ID")
    if id_fail:
        failed.append("mds_ids_match should reject EXTRACTION_FAILED")

    if not is_check_clean("0 Error(s), 0 Warning(s)"):
        failed.append("is_check_clean must accept comma form")
    if is_check_clean("0 Error(s) / 1 Warning(s)"):
        failed.append("is_check_clean must reject warnings")
    if is_check_clean("Check failed"):
        failed.append("is_check_clean must reject tool failure")

    if kill_switch_active(lambda n: "1", lambda: False) is not True:
        failed.append("kill switch env")
    if kill_switch_active(lambda n: "0", lambda: True) is not True:
        failed.append("kill switch file")
    if kill_switch_active(lambda n: "0", lambda: False) is not False:
        failed.append("kill switch off")

    if failed:
        print("\nSELF-TEST FAILURES:")
        for item in failed:
            print(f"  - {item}")
        return 1

    print("\nself-test OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(run_self_test())
