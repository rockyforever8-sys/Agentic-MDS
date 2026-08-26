#!/usr/bin/env python3
"""IMDS Playwright agent — live inbox: check 10 MDS, accept+forward+propose PASS, reject FAIL.

Live session (defaults):
- Process NUM_ITERATIONS MDS (default 10)
- Overall PASS → Accept, then Forward, then Propose to recipient company IDs
- Overall FAIL → Reject with structured Rec 001 / IMDS Check text
- Tool/UI failure stays on HOLD (do not reject a sheet we could not score)
- Act on the MDS already open; verify ID/version before Accept/Reject
- Kill switch: IMDS_KILL_SWITCH=1 or imds_output/KILL
- Login secrets from environment / Colab Secrets only (no password defaults)

Does not log into IMDS during --self-test.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from imds_decisions import (
    ERROR_WARNING_RE,
    RULE_PACK_VERSION,
    apply_autonomous_policy,
    decide_overall,
    env_flag,
    kill_switch_active,
    mds_ids_match,
    run_self_test as run_decision_self_test,
)

try:
    import nest_asyncio

    nest_asyncio.apply()
except ImportError:
    pass

# nest_asyncio does not disable Playwright's own running-loop check. Jupyter and
# Google Colab already have an asyncio loop, so sync_playwright() must run in a
# child process (see orchestrate()).

# Chromium also needs OS libraries (libatk-1.0.so.0, …) that a bare Colab VM
# does not ship. `playwright install chromium` is not enough; see
# ensure_chromium_os_deps().
CHROMIUM_LAUNCH_ARGS = ["--no-sandbox", "--disable-dev-shm-usage"]
LIBATK_CANDIDATES = (
    Path("/usr/lib/x86_64-linux-gnu/libatk-1.0.so.0"),
    Path("/usr/lib/aarch64-linux-gnu/libatk-1.0.so.0"),
    Path("/lib/x86_64-linux-gnu/libatk-1.0.so.0"),
    Path("/lib/aarch64-linux-gnu/libatk-1.0.so.0"),
)
CHROMIUM_APT_PACKAGES = (
    "libatk1.0-0t64",
    "libatk1.0-0",
    "libatk-bridge2.0-0t64",
    "libatk-bridge2.0-0",
    "libcups2t64",
    "libcups2",
    "libdrm2",
    "libxkbcommon0",
    "libxcomposite1",
    "libxdamage1",
    "libxfixes3",
    "libxrandr2",
    "libgbm1",
    "libasound2t64",
    "libasound2",
    "libnss3",
    "libnspr4",
    "libpango-1.0-0",
    "libcairo2",
    "libatspi2.0-0t64",
    "libatspi2.0-0",
    "libx11-xcb1",
    "libxshmfence1",
    "libxcursor1",
    "libgtk-3-0t64",
    "libgtk-3-0",
    "fonts-liberation",
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("imds_agent")

# ---------- XPaths (Oracle ADF IMDS NT) ----------
XP_RECEIVED_MDS_MENU = "//*[@id='pt1:sdiInboxSearch::disAcr']"
XP_COMBINED_ALL = "//*[@id='pt1:dcCmds:sfIbLU:cbAll']/a"
XP_ID_FIELD = "//*[@id='pt1:dcCmds:sfIbLU:itModuleId::content']"
XP_SEARCH_BUTTON = "//*[@id='pt1:dcCmds:sfIbLU:cbSearch']/a"
XP_RESULT_ROWS = "//*[@id='pt1:dcCmds:sfIbLU:pc2:tResult::db']/table/tbody/tr"
XP_MDS_MENU = "//*[@id='pt1:pt_mFile']/div/table/tbody/tr/td[2]/a"
XP_ACCEPT = "//*[@id='pt1:pt_cmiMenuAccept']/td[2]"
XP_ACCEPT_MODAL = "//*[@id='dcPopup:ctbAcceptMds']/a/span"
XP_FORWARD_MENU = "//*[@id='pt1:pt_mMenuForward']"
XP_FORWARD_ACTION2 = "//*[@id='pt1:pt_cmiMenuForward']/td[2]"
XP_FORWARD_OK = "//*[@id='pt1:pt_dcud:ctbOk']/a"
XP_SUPPLIER_DATA = "//*[@id='pt1:sdiDetailSupplier::disAcr']"
XP_CONTACT = "//*[@id='pt1:dcSupp:socContact::content']"
XP_RECIPIENT_DATA = "//*[@id='pt1:sdiDetailRecipients::disAcr']"
XP_ADD_RECIPIENT = "//*[@id='pt1:dcReci:ctbAddRecipient::icon']"
XP_PROPOSE = "//*[@id='pt1:dcReci:ctbRecipPropose']/a/span"
XP_PROPOSE_MODAL = "//*[@id='dcPopup:ctbMultiPurpose']/a"
XP_REJECT_MENU = "//*[@id='pt1:pt_cmiMenuReject']/td[2]"
XP_REJECT_MODAL = "//*[@id='dcPopup:subViewReject:t1::oc']/table/tbody/tr/td[2]"
XP_COMPANY_ID_EXACT = [
    "//*[@id='pt1:svSearchCompanyLookup:sfSubLU:it2::content']",
    "//*[@id='pt1:svSearchCompanyLookup:sfSubLU:it2']",
]
XP_FILTER_NONE = "//*[@id='pt1:dcCmds:sfIbLU:cbNone']/a"
XP_FILTER_BROWSED = "//*[@id='pt1:dcCmds:sfIbLU:sbcBrowsed::content']"
XP_EXPAND_ALL = "//*[@id='pt1:dcIngr:ctbExpandAll']"
XP_REJECT_REASON = (
    "//textarea[contains(@id,'Reject') or contains(@id,'reason') or contains(@id,'Reason')]"
)

REQUIRED_CLASSIFICATIONS = {
    "5.1.a",
    "5.1.b",
    "5.2",
    "5.3",
    "5.4",
    "5.4.1",
    "5.4.2",
    "5.4.3",
    "5.5.1",
    "5.5.2",
    "6.1",
    "9.7",
    "7.1",
}


@dataclass
class Config:
    username: str
    password: str
    otp_secret: str
    inbox_url: str
    output_dir: Path
    num_iterations: int
    recipient_ids: list[str]
    contact_name: str
    auto_accept: bool
    auto_reject: bool
    auto_forward: bool
    hold_amber: bool
    require_parts_marking: bool
    debug_screenshots: bool
    headless: bool
    rule_pack_version: str


def _require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(
            f"{name} is not set. Add it in Colab Secrets (🔑) or the environment; "
            "do not hardcode it."
        )
    return value


def load_config() -> Config:
    from imds_secrets import apply_stored_credentials, missing_secret_keys

    apply_stored_credentials(persist=True)
    missing = missing_secret_keys()
    if missing:
        raise RuntimeError(
            "Missing private credentials: "
            + ", ".join(missing)
            + ". Add them in Colab Secrets (🔑) together with IMDS_MASTER_KEY, "
            "or unlock the encrypted vault. Never put passwords in the notebook."
        )
    output_dir = Path(os.getenv("IMDS_OUTPUT_DIR", "./imds_output"))
    output_dir.mkdir(parents=True, exist_ok=True)
    recipients_raw = os.getenv("RECIPIENT_COMPANY_IDS", "9994,293798")
    recipient_ids = [item.strip() for item in recipients_raw.split(",") if item.strip()]
    return Config(
        username=_require_env("IMDS_USERNAME"),
        password=_require_env("IMDS_PASSWORD"),
        otp_secret=_require_env("OTP_SECRET"),
        inbox_url=os.getenv(
            "IMDS_INBOX_URL", "https://www.mdsystem.com/imdsnt/faces/sentReceivedSearch"
        ),
        output_dir=output_dir,
        num_iterations=int(os.getenv("NUM_ITERATIONS", "10")),
        recipient_ids=recipient_ids,
        contact_name=os.getenv("IMDS_CONTACT_NAME", "Qu, Theresa").strip(),
        auto_accept=env_flag(os.getenv("IMDS_AUTO_ACCEPT"), True),
        auto_reject=env_flag(os.getenv("IMDS_AUTO_REJECT"), True),
        auto_forward=env_flag(os.getenv("IMDS_AUTO_FORWARD"), True),
        hold_amber=env_flag(os.getenv("IMDS_HOLD_AMBER"), False),
        require_parts_marking=env_flag(os.getenv("IMDS_REQUIRE_PARTS_MARKING"), False),
        debug_screenshots=env_flag(os.getenv("IMDS_DEBUG_SCREENSHOTS"), False),
        headless=env_flag(os.getenv("IMDS_HEADLESS"), True),
        rule_pack_version=os.getenv("IMDS_RULE_PACK_VERSION", RULE_PACK_VERSION),
    )


def kill_is_on(cfg: Config) -> bool:
    kill_file = cfg.output_dir / "KILL"
    return kill_switch_active(os.getenv, kill_file.exists)


def get_otp(secret: str) -> str:
    try:
        import pyotp
    except ImportError:
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyotp"])
        import pyotp

    cleaned = secret.replace(" ", "").upper()
    if len(cleaned) < 16:
        raise RuntimeError(
            "OTP_SECRET looks too short. Use the authenticator TOTP seed (base32), "
            "not a Gmail app password."
        )
    return pyotp.TOTP(cleaned).now()


def wait_ui(page, extra_ms: int = 0, timeout: int = 15000) -> None:
    """Wait for ADF busy indicators instead of a fixed sleep + networkidle."""
    for selector in (".AFBusyState", ".p_AFBusy", ".AFLoadingIcon"):
        try:
            loc = page.locator(selector).first
            if loc.count() > 0:
                loc.wait_for(state="hidden", timeout=timeout)
        except Exception:
            pass
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout)
    except Exception:
        pass
    if extra_ms:
        page.wait_for_timeout(extra_ms)


def visible(locator) -> bool:
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def click_ready(locator, *, force: bool = False, timeout: int = 10000) -> bool:
    """Click when visible. force=True only for ADF menus that sit under a glass pane."""
    try:
        target = locator.first
        target.wait_for(state="visible", timeout=timeout)
        target.click(force=force, timeout=timeout)
        return True
    except Exception as exc:
        log.warning("click_ready failed: %s", exc)
        return False


def save_screenshot(page, cfg: Config, name: str, *, important: bool = False) -> None:
    if not important and not cfg.debug_screenshots:
        return
    path = cfg.output_dir / name
    try:
        page.screenshot(path=str(path), full_page=True)
        log.info("Screenshot saved: %s", path)
    except Exception as exc:
        log.warning("Screenshot failed (%s): %s", name, exc)


def dismiss_modal(page, cfg: Config, *, allow_dom_cleanup: bool = False) -> bool:
    """Dismiss informational ADF modals. Never DOM-delete on Accept/Reject/Propose."""
    glass = page.locator(".AFModalGlassPane")
    try:
        if glass.count() == 0 or not glass.first.is_visible():
            return True
    except Exception:
        return True

    for selector in (
        "#pt1\\:pt_dcud\\:ctbOk",
        "button:has-text('OK')",
        "button:has-text('Close')",
        "button:has-text('Proceed')",
        "input[value='OK']",
        "input[value='Close']",
    ):
        btn = page.locator(selector).first
        if visible(btn):
            click_ready(btn, force=True)
            wait_ui(page, extra_ms=200)
            try:
                if glass.count() == 0 or not glass.first.is_visible():
                    return True
            except Exception:
                return True

    try:
        page.keyboard.press("Escape")
        wait_ui(page, extra_ms=200)
        if glass.count() == 0 or not glass.first.is_visible():
            return True
    except Exception:
        pass

    if allow_dom_cleanup:
        log.warning("Informational modal still up; not removing glass pane from the DOM.")
    else:
        log.warning("Modal still present; leaving it (no DOM delete on a decision dialog).")
        save_screenshot(page, cfg, "modal_still_present.png", important=True)
    return False


def append_audit(cfg: Config, record: dict) -> None:
    record = dict(record)
    record.setdefault("ts", datetime.now(timezone.utc).isoformat())
    record.setdefault("rule_pack_version", cfg.rule_pack_version)
    path = cfg.output_dir / "decisions.jsonl"
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=True) + "\n")


# ---------- Login / navigation ----------
def imds_login(page, cfg: Config) -> None:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

    log.info("Logging in...")
    page.goto("https://www.mdsystem.com/imdsnt", wait_until="domcontentloaded", timeout=60000)
    wait_ui(page, extra_ms=300)

    login_link = page.locator("a:has-text('Login')").first
    if visible(login_link):
        click_ready(login_link)
        wait_ui(page, extra_ms=300)

    username = page.locator("#username, input[name='username']").first
    try:
        username.wait_for(state="visible", timeout=30000)
    except PlaywrightTimeoutError:
        sso_btn = page.locator("button:has-text('Sign in'), a:has-text('Sign in')").first
        if not visible(sso_btn) or not click_ready(sso_btn):
            save_screenshot(page, cfg, "login_username_not_found.png", important=True)
            raise RuntimeError("Cannot find username field or sign-in button.")
        username = page.locator("#username, input[name='username']").first
        username.wait_for(state="visible", timeout=15000)

    username.fill(cfg.username)
    next_btn = page.locator("button:has-text('Sign in'), button:has-text('Next')").first
    if visible(next_btn):
        click_ready(next_btn)
    else:
        username.press("Enter")
    wait_ui(page, extra_ms=300)

    password = page.locator("input[type='password']").first
    try:
        password.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeoutError:
        save_screenshot(page, cfg, "login_password_not_found.png", important=True)
        raise RuntimeError("Password field not found after username submission.")
    password.fill(cfg.password)
    login_btn = page.locator(
        "button:has-text('Sign in'), button:has-text('Login'), input[value*='Sign in'], input[value*='Login']"
    ).first
    if visible(login_btn):
        click_ready(login_btn)
    else:
        password.press("Enter")
    wait_ui(page, extra_ms=400)

    try:
        otp_field = page.locator(
            "input[type='text'][placeholder*='code' i], input[type='text'][placeholder*='OTP' i], input[name='otp'], input#otp"
        ).first
        otp_field.wait_for(state="visible", timeout=15000)
        otp_field.fill(get_otp(cfg.otp_secret))
        log.info("Filled OTP (value not logged).")
        submit_otp = page.locator("button:has-text('Verify'), button:has-text('Submit')").first
        if visible(submit_otp):
            click_ready(submit_otp)
        else:
            otp_field.press("Enter")
        wait_ui(page, extra_ms=400)
    except PlaywrightTimeoutError:
        log.info("No OTP field appeared.")

    if visible(page.locator("button:has-text('Login')").first):
        save_screenshot(page, cfg, "login_failed.png", important=True)
        raise RuntimeError("Login failed.")
    log.info("Login successful.")
    save_screenshot(page, cfg, "01_after_login.png", important=True)


def navigate_to_search_page(page, cfg: Config) -> bool:
    log.info("Navigating to Received MDSs search page...")
    dismiss_modal(page, cfg, allow_dom_cleanup=True)

    received = page.locator("a:has-text('Received MDSs')").first
    if visible(received) and click_ready(received):
        wait_ui(page, extra_ms=400)
        if visible(page.locator(f"xpath={XP_ID_FIELD}")):
            return True

    back_btn = page.locator(f"xpath={XP_RECEIVED_MDS_MENU}")
    if visible(back_btn) and click_ready(back_btn, force=True):
        wait_ui(page, extra_ms=400)
        if visible(page.locator(f"xpath={XP_ID_FIELD}")):
            return True

    inbox_btn = page.locator("#pt1\\:pt_ctbToolBarInbound\\:\\:popEl")
    if visible(inbox_btn) and click_ready(inbox_btn):
        wait_ui(page, extra_ms=300)
        mds_item = page.locator("#pt1\\:pt_cmiSearchInboxB, a:has-text('MDS')").first
        if visible(mds_item) and click_ready(mds_item):
            wait_ui(page, extra_ms=400)
            if visible(page.locator(f"xpath={XP_ID_FIELD}")):
                return True

    log.error("Failed to navigate to search page.")
    save_screenshot(page, cfg, "nav_search_failed.png", important=True)
    return False


def apply_not_yet_browsed_filter(page, cfg: Config) -> None:
    """Click the native label so ADF PPR runs. Do not set checkbox.checked in JS."""
    none_btn = page.locator(f"xpath={XP_FILTER_NONE}")
    if visible(none_btn):
        click_ready(none_btn, force=True)
        wait_ui(page, extra_ms=200)

    label = page.locator("label:has-text('not yet browsed')").first
    if visible(label):
        click_ready(label)
        wait_ui(page, extra_ms=200)
    else:
        log.warning("Could not find 'not yet browsed' label.")

    save_screenshot(page, cfg, "04_after_filter_applied.png", important=True)
    search_btn = page.locator(f"xpath={XP_SEARCH_BUTTON}")
    if not (visible(search_btn) and click_ready(search_btn, force=True)):
        fallback = page.locator("a:has-text('Search'), button:has-text('Search')").first
        click_ready(fallback, force=True)
    wait_ui(page, extra_ms=500)
    dismiss_modal(page, cfg, allow_dom_cleanup=True)
    try:
        page.wait_for_selector("table", timeout=15000)
    except Exception:
        log.warning("Result table not found after search.")
    save_screenshot(page, cfg, "05_after_search.png", important=True)


def expand_tree(page, cfg: Config) -> None:
    expand = page.locator(f"xpath={XP_EXPAND_ALL}")
    if visible(expand) and click_ready(expand, force=True):
        wait_ui(page, extra_ms=400)
        dismiss_modal(page, cfg, allow_dom_cleanup=True)
        return
    log.warning("Expand All not found; expanding collapsed nodes.")
    previous = -1
    for _ in range(12):
        nodes = page.locator("[aria-expanded='false']").all()
        if not nodes or len(nodes) == previous:
            break
        previous = len(nodes)
        for node in nodes[:40]:
            try:
                if node.is_visible():
                    node.click(timeout=5000)
                    wait_ui(page, extra_ms=150)
            except Exception:
                continue


def extract_mds_id_version(page, cfg: Config) -> str:
    try:
        page.wait_for_selector("td:has-text('ID / Version')", timeout=10000)
        cells = page.locator("td").all()
        for cell in cells:
            text = (cell.text_content() or "").strip()
            if "ID / Version" in text:
                nxt = cell.locator("xpath=following-sibling::td").first
                if nxt.count():
                    val = (nxt.text_content() or "").strip()
                    if val:
                        return val
        body = page.locator("body").text_content() or ""
        match = re.search(r"(\d{6,}\s*/\s*[\d.]+)", body)
        if match:
            return match.group(1)
    except Exception as exc:
        log.warning("ID extraction failed: %s", exc)
    save_screenshot(page, cfg, "id_extraction_failed.png", important=True)
    return "EXTRACTION_FAILED"


def _label_value(page, *label_texts: str) -> str:
    for text in label_texts:
        label = page.locator(f"label:has-text('{text}')").first
        if visible(label):
            value = label.locator("xpath=following-sibling::*").first
            if value.count():
                return (value.text_content() or "").strip()
    return ""


def extract_material_classification(page) -> str:
    value = _label_value(page, "Material class", "Classification", "Material group")
    if value:
        return value
    body = page.locator("body").text_content() or ""
    match = re.search(r"\b(5\.[0-9]+(?:\.[a-z])?|6\.1|9\.7|7\.1)\b", body, re.I)
    return match.group(1) if match else ""


def capture_material_nodes(page, cfg: Config, iteration: int) -> None:
    nodes = page.locator("[role='treeitem']:has(img[src*='material'])").all()
    if not nodes:
        for img in page.locator("img[src*='btn_tree_material']").all():
            parent = img.locator("xpath=ancestor::*[@role='treeitem']")
            if parent.count():
                nodes.append(parent)
    for idx, node in enumerate(nodes, start=1):
        try:
            node.scroll_into_view_if_needed()
            node.click()
            wait_ui(page, extra_ms=150)
            save_screenshot(page, cfg, f"mds_iter{iteration}_node{idx}.png", important=True)
            dismiss_modal(page, cfg, allow_dom_cleanup=True)
        except Exception as exc:
            log.warning("Material screenshot %s failed: %s", idx, exc)


def material_nodes(page):
    nodes = page.locator("[role='treeitem']:has(img[src*='material'])").all()
    if nodes:
        return nodes
    found = []
    for img in page.locator("img[src*='btn_tree_material']").all():
        parent = img.locator("xpath=ancestor::*[@role='treeitem']")
        if parent.count():
            found.append(parent)
    return found


def run_checks_on_mds(page, cfg: Config) -> dict:
    recyclate_fail = False
    biocidal_fail = False
    parts_marking_fail = False
    nodes = material_nodes(page)
    if not nodes:
        return {
            "parts_marking_check": "No materials",
            "recyclate_check": "No materials",
            "biocidal_check": "No materials",
        }

    comp_map: dict[str, list[str]] = {}
    for node in nodes:
        try:
            node.scroll_into_view_if_needed()
            node.click()
            wait_ui(page, extra_ms=150)
            classification = extract_material_classification(page)
            recyclate = _label_value(page, "recyclate", "Recycled", "recycl")
            biocidal = _label_value(page, "Still in production", "still in production")
            if classification in REQUIRED_CLASSIFICATIONS:
                if not recyclate or recyclate.lower() == "not yet answered":
                    recyclate_fail = True
                if not (biocidal or "").strip():
                    biocidal_fail = True
            handle = node.element_handle()
            if handle:
                parent_text = page.evaluate(
                    """(el) => {
                        let p = el.parentElement;
                        while (p) {
                          if (p.getAttribute('role') === 'treeitem') {
                            const img = p.querySelector('img');
                            if (img && (img.src.includes('component') || img.src.includes('package'))) {
                              return p.innerText || '';
                            }
                          }
                          p = p.parentElement;
                        }
                        return '';
                    }""",
                    handle,
                )
                if parent_text and classification:
                    comp_map.setdefault(parent_text.strip(), []).append(classification)
            dismiss_modal(page, cfg, allow_dom_cleanup=True)
        except Exception as exc:
            log.warning("Material rule scrape failed: %s", exc)

    components = page.locator("[role='treeitem']:has(img[src*='component'])").all()
    if not components:
        components = page.locator("[role='treeitem']:has(img[src*='package'])").all()
    if not components:
        parts_result = "No components"
    else:
        for comp in components:
            try:
                comp.scroll_into_view_if_needed()
                comp.click()
                wait_ui(page, extra_ms=150)
                marking = _label_value(page, "Part marking", "Parts marking", "marking")
                text = (comp.text_content() or "").strip()
                classes = comp_map.get(text, [])
                if any(cls in REQUIRED_CLASSIFICATIONS for cls in classes) and not marking:
                    parts_marking_fail = True
                dismiss_modal(page, cfg, allow_dom_cleanup=True)
            except Exception as exc:
                log.warning("Component scrape failed: %s", exc)
        parts_result = "FAIL" if parts_marking_fail else "PASS"

    return {
        "parts_marking_check": parts_result,
        "recyclate_check": "FAIL" if recyclate_fail else "PASS",
        "biocidal_check": "FAIL" if biocidal_fail else "PASS",
    }


def run_imds_check(page, cfg: Config) -> tuple[bool, str]:
    menu = page.locator(f"xpath={XP_MDS_MENU}")
    if not (visible(menu) and click_ready(menu, force=True)):
        fallback = page.locator("a:has-text('MDS')").first
        if not (visible(fallback) and click_ready(fallback, force=True)):
            return False, "Check failed"
    wait_ui(page, extra_ms=200)
    check_item = page.locator("a:has-text('Check'), #pt1\\:pt_cmiMenuCheck").first
    if not (visible(check_item) and click_ready(check_item, force=True)):
        return False, "Check failed"
    try:
        page.wait_for_selector("table:has-text('Message'), table:has-text('Check results')", timeout=15000)
    except Exception:
        return False, "Check failed"
    wait_ui(page, extra_ms=200)
    dismiss_modal(page, cfg, allow_dom_cleanup=True)

    table = page.locator("table:has-text('Check results')").first
    if table.count():
        for row in table.locator("tr").all():
            for cell in row.locator("td").all():
                text = (cell.text_content() or "").strip()
                if text and len(text) > 20 and "Export" not in text and "hidden column" not in text:
                    return True, text
    body = page.locator("body").text_content() or ""
    match = re.search(r"(The MDS has passed all included checks\.[^.]*\.)", body)
    if match:
        return True, match.group(1)
    match = ERROR_WARNING_RE.search(body)
    if match:
        return True, f"{match.group(1)} Error(s) / {match.group(2)} Warning(s)"
    return True, "Check result not found"


def confirm_success(page, phrases: list[str], timeout: int = 10000) -> bool:
    for phrase in phrases:
        try:
            page.get_by_text(phrase, exact=False).first.wait_for(state="visible", timeout=timeout)
            return True
        except Exception:
            continue
    return False


def open_mds_menu(page, cfg: Config) -> bool:
    menu = page.locator(f"xpath={XP_MDS_MENU}")
    for _ in range(3):
        if visible(menu) and click_ready(menu, force=True):
            wait_ui(page, extra_ms=200)
            return True
        fallback = page.locator("a:has-text('MDS')").first
        if visible(fallback) and click_ready(fallback, force=True):
            wait_ui(page, extra_ms=200)
            return True
        wait_ui(page, extra_ms=400)
    save_screenshot(page, cfg, "mds_menu_not_found.png", important=True)
    return False


def accept_mds(page, cfg: Config) -> bool:
    log.info("Accepting MDS (GREEN).")
    if not open_mds_menu(page, cfg):
        return False
    accept_btn = page.locator(f"xpath={XP_ACCEPT}")
    if not (visible(accept_btn) and click_ready(accept_btn, force=True)):
        alt = page.locator("[role='menuitem']:has-text('Accept'), a:has-text('Accept')").first
        if not (visible(alt) and click_ready(alt, force=True)):
            save_screenshot(page, cfg, "accept_menu_item_not_found.png", important=True)
            return False
    wait_ui(page, extra_ms=300)
    modal = page.locator(f"xpath={XP_ACCEPT_MODAL}")
    if visible(modal):
        click_ready(modal, force=True)
    else:
        alt = page.locator("button:has-text('Accept'), input[value='Accept']").first
        if visible(alt):
            click_ready(alt, force=True)
        else:
            log.warning("Accept confirmation not found.")
            save_screenshot(page, cfg, "accept_modal_missing.png", important=True)
            return False
    wait_ui(page, extra_ms=400)
    if not confirm_success(page, ["MDS accepted", "accepted successfully", "Accept successful"]):
        log.error("Accept not confirmed; not treating as success.")
        save_screenshot(page, cfg, "accept_uncertain.png", important=True)
        return False
    save_screenshot(page, cfg, "after_accept.png", important=True)
    return True


def fill_reject_reason(page, reason: str) -> None:
    if not reason:
        return
    field = page.locator(f"xpath={XP_REJECT_REASON}").first
    if not visible(field):
        field = page.locator("textarea:visible, input[type='text']:visible").first
    if not visible(field):
        log.warning("Reject reason field not found; continuing with menu confirm only.")
        return
    try:
        field.click()
        field.fill(reason[:2000])
        log.info("Filled structured reject reason (%s chars).", min(len(reason), 2000))
    except Exception as exc:
        log.warning("Could not fill reject reason: %s", exc)


def reject_mds(page, cfg: Config, reason: str) -> bool:
    log.info("Rejecting MDS (RED).")
    if not open_mds_menu(page, cfg):
        return False
    reject_btn = page.locator(f"xpath={XP_REJECT_MENU}")
    if not (visible(reject_btn) and click_ready(reject_btn, force=True)):
        alt = page.locator("[role='menuitem']:has-text('Reject'), a:has-text('Reject')").first
        if not (visible(alt) and click_ready(alt, force=True)):
            save_screenshot(page, cfg, "reject_menu_item_not_found.png", important=True)
            return False
    wait_ui(page, extra_ms=300)
    fill_reject_reason(page, reason)
    modal = page.locator(f"xpath={XP_REJECT_MODAL}")
    if visible(modal):
        click_ready(modal, force=True)
    else:
        alt = page.locator("button:has-text('Reject'), input[value='Reject']").first
        if visible(alt):
            click_ready(alt, force=True)
        else:
            log.warning("Reject confirmation not found.")
            save_screenshot(page, cfg, "reject_modal_missing.png", important=True)
            return False
    wait_ui(page, extra_ms=400)
    if not confirm_success(page, ["MDS rejected", "rejected successfully", "Reject successful"]):
        log.error("Reject not confirmed; not treating as success.")
        save_screenshot(page, cfg, "reject_uncertain.png", important=True)
        return False
    save_screenshot(page, cfg, "after_reject.png", important=True)
    return True


def forward_mds(page, cfg: Config) -> bool:
    if not cfg.auto_forward:
        log.info("Skipping forward (IMDS_AUTO_FORWARD is off).")
        return True
    log.info("Forwarding accepted MDS.")
    if not open_mds_menu(page, cfg):
        return False
    fwd = page.locator(f"xpath={XP_FORWARD_ACTION2}")
    if not (visible(fwd) and click_ready(fwd, force=True)):
        alt = page.locator(f"xpath={XP_FORWARD_MENU}, a:has-text('Forward')").first
        if not (visible(alt) and click_ready(alt, force=True)):
            return False
    wait_ui(page, extra_ms=300)
    yes = page.locator("button:has-text('Yes'), input[value='Yes']").first
    if visible(yes):
        click_ready(yes, force=True)
        wait_ui(page, extra_ms=300)
    ok_btn = page.locator(f"xpath={XP_FORWARD_OK}")
    if visible(ok_btn):
        click_ready(ok_btn, force=True)
        wait_ui(page, extra_ms=400)
    return True


def select_contact_person(page, cfg: Config) -> bool:
    if not cfg.contact_name:
        log.info("IMDS_CONTACT_NAME not set; skipping contact selection.")
        return True
    dropdown = page.locator(f"xpath={XP_CONTACT}")
    if not visible(dropdown):
        log.warning("Contact dropdown not found.")
        return False
    try:
        tag = dropdown.evaluate("el => el.tagName")
        if tag and tag.lower() == "select":
            dropdown.select_option(label=cfg.contact_name)
            wait_ui(page, extra_ms=200)
            return True
    except Exception as exc:
        log.warning("select_option failed: %s", exc)
    click_ready(dropdown, force=True)
    wait_ui(page, extra_ms=200)
    option = page.locator(f"[role='option']:has-text('{cfg.contact_name}'), option:has-text('{cfg.contact_name}')").first
    if visible(option) and click_ready(option, force=True):
        return True
    log.warning("Contact '%s' not found.", cfg.contact_name)
    return False


def _company_id_field(frame_locator):
    for xp in XP_COMPANY_ID_EXACT:
        field = frame_locator.locator(f"xpath={xp}").first
        if visible(field):
            return field
    for pattern in ("CompanyId", "CompanyID", "svSearchCompanyLookup", "it2"):
        field = frame_locator.locator(f"xpath=//input[contains(@id, '{pattern}')]").first
        if visible(field):
            return field
    label = frame_locator.locator("label:has-text('Company ID'), label:has-text('Company Id')").first
    if visible(label):
        sibling = label.locator("xpath=following::input[1]").first
        if visible(sibling):
            return sibling
    return None


def add_recipient(page, cfg: Config, company_id: str, supplier_code: str, part_no: str) -> bool:
    add_btn = page.locator(f"xpath={XP_ADD_RECIPIENT}")
    if not (visible(add_btn) and click_ready(add_btn, force=True)):
        return False
    wait_ui(page, extra_ms=300)
    try:
        frame = page.frame_locator("iframe[src*='lookupCompany']")
        frame.locator("body").wait_for(timeout=15000)
    except Exception:
        save_screenshot(page, cfg, "no_iframe.png", important=True)
        return False

    company_field = _company_id_field(frame)
    if company_field is None:
        log.error("Company ID field not found; refusing first-visible-input fallback.")
        save_screenshot(page, cfg, "company_id_not_found.png", important=True)
        return False
    company_field.click()
    company_field.fill(company_id)
    if (company_field.input_value() or "") != company_id:
        log.warning("Company ID input mismatch after fill.")
        return False

    search_btn = frame.locator("xpath=//*[@id='pt1:svSearchCompanyLookup:sfSubLU:cbSearch']/a/span")
    if not visible(search_btn):
        search_btn = frame.locator("button:has-text('Search'), a:has-text('Search')").first
    if not (visible(search_btn) and click_ready(search_btn, force=True)):
        return False
    wait_ui(page, extra_ms=400)

    apply_btn = frame.locator("button:has-text('Apply'), a:has-text('Apply'), input[value='Apply']").first
    if not (visible(apply_btn) and click_ready(apply_btn, force=True)):
        # Browser CSS has no :has-text; click by textContent.
        clicked = frame.locator("body").evaluate(
            """() => {
                const nodes = Array.from(document.querySelectorAll('button, a, input'));
                const btn = nodes.find(el => (el.value || el.textContent || '').trim() === 'Apply');
                if (btn) { btn.click(); return true; }
                return false;
            }"""
        )
        if not clicked:
            return False
    wait_ui(page, extra_ms=400)
    try:
        page.wait_for_selector("iframe[src*='lookupCompany']", state="detached", timeout=15000)
    except Exception:
        dismiss_modal(page, cfg)

    if supplier_code and supplier_code.strip() not in {"", "-"}:
        supp = page.locator("xpath=//*[@id='pt1:dcReci:itSuppCode::content']")
        if visible(supp):
            supp.fill(supplier_code)
    if part_no and part_no.strip():
        part = page.locator("xpath=//*[@id='pt1:dcReci:itprodCode::content']")
        if visible(part):
            part.fill(part_no)

    propose = page.locator(f"xpath={XP_PROPOSE}")
    if not (visible(propose) and click_ready(propose, force=True)):
        return False
    wait_ui(page, extra_ms=300)
    modal = page.locator(f"xpath={XP_PROPOSE_MODAL}")
    if visible(modal):
        click_ready(modal, force=True)
        wait_ui(page, extra_ms=400)
    return True


def complete_forward_recipients(page, cfg: Config, supplier_code: str, part_no: str) -> bool:
    if not cfg.auto_forward:
        return True
    if not cfg.recipient_ids:
        log.warning("RECIPIENT_COMPANY_IDS empty; skipping Propose.")
        return True
    supplier_tab = page.locator(f"xpath={XP_SUPPLIER_DATA}")
    if visible(supplier_tab):
        click_ready(supplier_tab, force=True)
        wait_ui(page, extra_ms=200)
    select_contact_person(page, cfg)
    recip_tab = page.locator(f"xpath={XP_RECIPIENT_DATA}")
    if visible(recip_tab):
        click_ready(recip_tab, force=True)
        wait_ui(page, extra_ms=200)
    ok = True
    for cid in cfg.recipient_ids:
        if not add_recipient(page, cfg, cid, supplier_code, part_no):
            log.warning("Failed to add recipient %s.", cid)
            ok = False
    return ok


def verify_open_mds(page, cfg: Config, expected_id: str) -> bool:
    opened = extract_mds_id_version(page, cfg)
    if not mds_ids_match(opened, expected_id):
        log.error("Open MDS '%s' does not match scored MDS '%s'. Skipping action.", opened, expected_id)
        save_screenshot(page, cfg, "id_mismatch.png", important=True)
        append_audit(
            cfg,
            {
                "event": "id_mismatch",
                "opened": opened,
                "expected": expected_id,
            },
        )
        return False
    return True


def go_back_to_inbox(page, cfg: Config) -> None:
    back = page.locator(f"xpath={XP_RECEIVED_MDS_MENU}")
    if visible(back) and click_ready(back, force=True):
        wait_ui(page, extra_ms=400)
        dismiss_modal(page, cfg, allow_dom_cleanup=True)
        return
    try:
        page.go_back()
        wait_ui(page, extra_ms=400)
    except Exception:
        navigate_to_search_page(page, cfg)


def extract_row_meta(row_element) -> tuple[str, str, str]:
    supplier_code = part_no = status = ""
    try:
        row_tr = row_element.locator("xpath=ancestor::tr")
        cells = row_tr.locator("td").all()
        if len(cells) >= 8:
            part_no = (cells[3].text_content() or "").strip()
            supplier_code = (cells[6].text_content() or "").strip()
            if supplier_code == "-":
                supplier_code = ""
            status = (cells[7].text_content() or "").strip()
    except Exception as exc:
        log.warning("Row meta extract failed: %s", exc)
    return supplier_code, part_no, status


def write_excel(cfg: Config, results: list[dict]) -> Path | None:
    if not results:
        return None
    try:
        from openpyxl import Workbook
    except ImportError:
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Check Summary"
    headers = [
        "MDS ID / Version",
        "Decision Band",
        "Action",
        "Overall Result",
        "Check Result",
        "Check Errors",
        "Check Warnings",
        "Parts Marking Check",
        "Recyclate Check",
        "Biocidal Check",
        "Reasons",
        "Reject Text",
        "Rule Pack",
        "Supplier Code",
        "Part/Item No.",
        "Status",
        "Acted",
    ]
    ws.append(headers)
    for row in results:
        ws.append([row.get(h, "") for h in headers])
    path = cfg.output_dir / "check_summary.xlsx"
    wb.save(str(path))
    log.info("Excel summary saved to %s", path)
    return path


def _final_status(row: dict) -> str:
    if (row.get("Action") or "").lower() == "accept":
        return "ACCEPT"
    return "REJECT"


def write_status_report(cfg: Config, results: list[dict]) -> Path:
    """One-row-per-MDS accept/reject report keyed by MDS ID."""
    csv_path = cfg.output_dir / "mds_status_report.csv"
    xlsx_path = cfg.output_dir / "mds_status_report.xlsx"
    md_path = cfg.output_dir / "mds_status_report.md"
    lines = ["mds_id,result,imds_acted,reasons"]
    md = [
        "# IMDS accept / reject report",
        "",
        "| MDS ID | Result | IMDS acted | Reasons |",
        "|---|---|---|---|",
    ]
    try:
        from openpyxl import Workbook
    except ImportError:
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "MDS Status"
    ws.append(["MDS ID", "Result", "IMDS acted", "Reasons"])
    counts = {"ACCEPT": 0, "REJECT": 0}
    for row in results:
        mds_id = row.get("MDS ID / Version") or ""
        result = _final_status(row)
        counts[result] = counts.get(result, 0) + 1
        acted = row.get("Acted") or ""
        reasons = (row.get("Reasons") or "").replace("\n", " ").replace(",", ";")
        lines.append(f"{mds_id},{result},{acted},{reasons}")
        md.append(f"| {mds_id} | {result} | {acted} | {reasons} |")
        ws.append([mds_id, result, acted, row.get("Reasons") or ""])
    md.extend(["", f"Accepted: {counts.get('ACCEPT', 0)}  Rejected: {counts.get('REJECT', 0)}"])
    csv_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    md_path.write_text("\n".join(md) + "\n", encoding="utf-8")
    wb.save(str(xlsx_path))
    log.info(
        "Status report: %s ACCEPT, %s REJECT → %s",
        counts.get("ACCEPT", 0),
        counts.get("REJECT", 0),
        csv_path,
    )
    print("\n".join(md))
    return csv_path


def process_rows_and_export(page, cfg: Config) -> list[dict]:
    results: list[dict] = []
    for i in range(cfg.num_iterations):
        if kill_is_on(cfg):
            log.warning("Kill switch on; stopping after %s MDS.", i)
            append_audit(cfg, {"event": "kill_switch", "processed": i})
            break

        row_xpath = f"//*[@id='pt1:dcCmds:sfIbLU:pc2:tResult:{i}:cName']"
        log.info("Processing MDS row %s/%s", i + 1, cfg.num_iterations)
        row_element = page.locator(row_xpath)
        try:
            row_element.wait_for(state="visible", timeout=10000)
        except Exception as exc:
            log.warning("Row %s not found: %s. Stopping.", i, exc)
            break

        supplier_code, part_no, status = extract_row_meta(row_element)
        try:
            row_element.dblclick()
        except Exception:
            row_element.click()
            page.keyboard.press("Enter")
        wait_ui(page, extra_ms=400)
        dismiss_modal(page, cfg, allow_dom_cleanup=True)

        mds_id = extract_mds_id_version(page, cfg)
        expand_tree(page, cfg)
        if mds_id == "EXTRACTION_FAILED":
            mds_id = extract_mds_id_version(page, cfg)

        capture_material_nodes(page, cfg, i + 1)
        rule_results = run_checks_on_mds(page, cfg)
        first_node = page.locator("[role='treeitem']").first
        if visible(first_node):
            click_ready(first_node)
            wait_ui(page, extra_ms=150)

        check_ok, result_msg = run_imds_check(page, cfg)
        if not check_ok:
            result_msg = result_msg or "Check failed"
        save_screenshot(page, cfg, f"mds_check_{i + 1}.png", important=True)

        decision = decide_overall(
            check_result=result_msg,
            recyclate_check=rule_results.get("recyclate_check", "Unknown"),
            biocidal_check=rule_results.get("biocidal_check", "Unknown"),
            parts_marking_check=rule_results.get("parts_marking_check", "Unknown"),
            mds_id=mds_id,
            require_parts_marking=cfg.require_parts_marking,
            rule_pack_version=cfg.rule_pack_version,
        )
        decision = apply_autonomous_policy(
            decision,
            hold_amber=cfg.hold_amber,
            mds_id=mds_id,
            check_result=result_msg,
        )
        log.info(
            "MDS %s band=%s action=%s check=%s recyclate=%s biocidal=%s marking=%s",
            mds_id,
            decision.band,
            decision.action,
            result_msg,
            rule_results.get("recyclate_check"),
            rule_results.get("biocidal_check"),
            rule_results.get("parts_marking_check"),
        )

        acted = "no"
        action_error = ""
        try:
            if decision.action == "accept" and cfg.auto_accept:
                if not verify_open_mds(page, cfg, mds_id):
                    action_error = "id_mismatch"
                elif accept_mds(page, cfg):
                    acted = "accepted"
                    if cfg.auto_forward:
                        if forward_mds(page, cfg) and complete_forward_recipients(
                            page, cfg, supplier_code, part_no
                        ):
                            acted = "accepted_forwarded"
                        else:
                            action_error = "forward_failed"
                else:
                    action_error = "accept_unconfirmed"
            elif decision.action == "reject" and cfg.auto_reject:
                if not verify_open_mds(page, cfg, mds_id):
                    action_error = "id_mismatch"
                elif reject_mds(page, cfg, decision.reject_text):
                    acted = "rejected"
                else:
                    action_error = "reject_unconfirmed"
            else:
                log.info("Holding MDS %s for human review (%s).", mds_id, decision.band)
                acted = "held"
        except Exception as exc:
            action_error = str(exc)
            log.warning("Action failed for %s: %s", mds_id, exc)
            save_screenshot(page, cfg, f"action_error_{i + 1}.png", important=True)

        row = {
            "MDS ID / Version": mds_id,
            "Decision Band": decision.band,
            "Action": decision.action,
            "Overall Result": decision.overall,
            "Check Result": result_msg,
            "Check Errors": decision.check_errors if decision.check_errors is not None else "",
            "Check Warnings": decision.check_warnings if decision.check_warnings is not None else "",
            "Parts Marking Check": rule_results.get("parts_marking_check", "Unknown"),
            "Recyclate Check": rule_results.get("recyclate_check", "Unknown"),
            "Biocidal Check": rule_results.get("biocidal_check", "Unknown"),
            "Reasons": " | ".join(decision.reasons),
            "Reject Text": decision.reject_text,
            "Rule Pack": decision.rule_pack_version,
            "Supplier Code": supplier_code,
            "Part/Item No.": part_no,
            "Status": status,
            "Acted": acted,
        }
        results.append(row)
        append_audit(
            cfg,
            {
                "event": "decision",
                "mds_id": mds_id,
                "band": decision.band,
                "action": decision.action,
                "acted": acted,
                "action_error": action_error,
                "check_result": result_msg,
                "reasons": decision.reasons,
                "reject_text": decision.reject_text,
            },
        )

        if i < cfg.num_iterations - 1:
            go_back_to_inbox(page, cfg)

    write_excel(cfg, results)
    write_status_report(cfg, results)
    return results


def _asyncio_loop_is_running() -> bool:
    """True inside Jupyter/Colab (and any other running asyncio loop)."""
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return False
    return True


def _orchestrate_in_subprocess() -> int:
    """Run this file in a child process so Playwright Sync API can start.

    Playwright raises:
      "It looks like you are using Playwright Sync API inside the asyncio loop"
    when sync_playwright() is called from a Colab/ipywidgets click handler.
    """
    script = Path(__file__).resolve()
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["IMDS_INSIDE_ORCHESTRATE_SUBPROCESS"] = "1"
    log.info(
        "Colab/Jupyter asyncio loop is running — Playwright Sync API cannot start "
        "in this process. Launching %s in a subprocess.",
        script.name,
    )
    proc = subprocess.Popen(
        [sys.executable, "-u", str(script)],
        cwd=str(script.parent),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    assert proc.stdout is not None
    for line in proc.stdout:
        print(line, end="")
        sys.stdout.flush()
    rc = proc.wait()
    return int(rc if rc is not None else 1)


def libatk_present() -> bool:
    return any(path.exists() for path in LIBATK_CANDIDATES)


def _looks_like_missing_chromium_lib(exc: BaseException) -> bool:
    text = str(exc).lower()
    return any(
        needle in text
        for needle in (
            "shared libraries",
            "libatk",
            "cannot open shared object",
            "target page, context or browser has been closed",
            "targetclosederror",
        )
    )


def ensure_chromium_os_deps() -> None:
    """Install libatk and friends. Colab's `playwright install chromium` skips these."""
    if os.environ.get("IMDS_SKIP_BROWSER_DEPS") == "1":
        return
    if libatk_present():
        return
    log.info("Chromium OS libraries missing (libatk-1.0.so.0). Installing Playwright deps...")
    env = os.environ.copy()
    env["DEBIAN_FRONTEND"] = "noninteractive"
    subprocess.run(
        [sys.executable, "-m", "playwright", "install-deps", "chromium"],
        env=env,
        check=False,
    )
    if libatk_present():
        log.info("Playwright install-deps provided libatk.")
        return
    subprocess.run(["apt-get", "update", "-qq"], env=env, check=False)
    t64 = [pkg for pkg in CHROMIUM_APT_PACKAGES if pkg.endswith("t64")]
    legacy = [pkg for pkg in CHROMIUM_APT_PACKAGES if not pkg.endswith("t64")]
    for batch in (t64, legacy):
        subprocess.run(
            ["apt-get", "install", "-y", "-qq", "--no-install-recommends", *batch],
            env=env,
            check=False,
        )
        if libatk_present():
            log.info("apt installed libatk.")
            return
    log.error(
        "Still missing libatk-1.0.so.0. In Colab run: "
        "!python -m playwright install-deps chromium"
    )


def _launch_chromium(playwright, cfg):
    ensure_chromium_os_deps()
    kwargs = {"headless": cfg.headless, "args": list(CHROMIUM_LAUNCH_ARGS)}
    try:
        return playwright.chromium.launch(**kwargs)
    except Exception as exc:
        if not _looks_like_missing_chromium_lib(exc):
            raise
        log.warning("Chromium launch failed (%s). Installing OS deps and retrying once.", exc)
        os.environ.pop("IMDS_SKIP_BROWSER_DEPS", None)
        env = os.environ.copy()
        env["DEBIAN_FRONTEND"] = "noninteractive"
        subprocess.run(
            [sys.executable, "-m", "playwright", "install-deps", "chromium"],
            env=env,
            check=False,
        )
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            env=env,
            check=False,
        )
        if not libatk_present():
            subprocess.run(["apt-get", "update", "-qq"], env=env, check=False)
            t64 = [pkg for pkg in CHROMIUM_APT_PACKAGES if pkg.endswith("t64")]
            legacy = [pkg for pkg in CHROMIUM_APT_PACKAGES if not pkg.endswith("t64")]
            for batch in (t64, legacy):
                subprocess.run(
                    ["apt-get", "install", "-y", "-qq", "--no-install-recommends", *batch],
                    env=env,
                    check=False,
                )
                if libatk_present():
                    break
        return playwright.chromium.launch(**kwargs)


def _orchestrate_live() -> int:
    cfg = load_config()
    if kill_is_on(cfg):
        log.error("Kill switch is on (IMDS_KILL_SWITCH or %s). Not starting.", cfg.output_dir / "KILL")
        return 2
    log.info(
        "Live session: %s MDS; PASS→accept+forward+propose to %s; FAIL→reject; contact=%s; hold_amber=%s",
        cfg.num_iterations,
        ",".join(cfg.recipient_ids) or "(none)",
        cfg.contact_name or "(none)",
        cfg.hold_amber,
    )

    from playwright.sync_api import sync_playwright

    try:
        with sync_playwright() as playwright:
            browser = _launch_chromium(playwright, cfg)
            page = browser.new_page()
            try:
                imds_login(page, cfg)
                if not navigate_to_search_page(page, cfg):
                    raise RuntimeError("Could not open Received MDSs search")
                apply_not_yet_browsed_filter(page, cfg)
                process_rows_and_export(page, cfg)
                log.info("All done.")
                return 0
            except Exception as exc:
                log.error("Script failed: %s", exc)
                save_screenshot(page, cfg, "error.png", important=True)
                return 1
            finally:
                browser.close()
    except Exception as exc:
        log.error("Could not start Chromium: %s", exc)
        if _looks_like_missing_chromium_lib(exc) or not libatk_present():
            log.error(
                "Colab is missing Chromium OS libraries (often libatk-1.0.so.0). "
                "Re-run Cell 1, or run: !python -m playwright install-deps chromium"
            )
        return 1


def orchestrate() -> int:
    """Live IMDS session. Safe to call from Colab's green button.

    If the current process already has a running asyncio loop (Jupyter/Colab),
    Playwright Sync API cannot start here. The agent is launched as
    `python -u imds_agent_v2.py` in a child process instead.
    """
    if os.environ.get("IMDS_INSIDE_ORCHESTRATE_SUBPROCESS") == "1":
        return _orchestrate_live()
    if _asyncio_loop_is_running():
        return _orchestrate_in_subprocess()
    return _orchestrate_live()


def _write_fixture_excel(output_dir: Path) -> Path:
    """Exercise Excel + JSONL from canned decisions (no IMDS login)."""
    os.environ.setdefault("IMDS_USERNAME", "selftest")
    os.environ.setdefault("IMDS_PASSWORD", "selftest")
    os.environ.setdefault("OTP_SECRET", "JBSWY3DPEHPK3PXP")  # RFC 4226 test secret, not a real login
    os.environ["IMDS_OUTPUT_DIR"] = str(output_dir)
    cfg = load_config()
    fixtures = [
        decide_overall(
            check_result="0 Error(s), 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
            mds_id="1111111111 / 1.0",
        ),
        decide_overall(
            check_result="0 Error(s) / 2 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
            mds_id="2222222222 / 1.0",
        ),
        decide_overall(
            check_result="2 Error(s) / 0 Warning(s)",
            recyclate_check="FAIL",
            biocidal_check="PASS",
            parts_marking_check="PASS",
            mds_id="3333333333 / 1.0",
        ),
    ]
    rows = []
    for decision, mds_id in zip(
        fixtures, ["1111111111 / 1.0", "2222222222 / 1.0", "3333333333 / 1.0"]
    ):
        rows.append(
            {
                "MDS ID / Version": mds_id,
                "Decision Band": decision.band,
                "Action": decision.action,
                "Overall Result": decision.overall,
                "Check Result": "fixture",
                "Check Errors": decision.check_errors,
                "Check Warnings": decision.check_warnings,
                "Parts Marking Check": "PASS",
                "Recyclate Check": "PASS" if decision.band != "RED" else "FAIL",
                "Biocidal Check": "PASS",
                "Reasons": " | ".join(decision.reasons),
                "Reject Text": decision.reject_text,
                "Rule Pack": decision.rule_pack_version,
                "Supplier Code": "",
                "Part/Item No.": "PN-TEST",
                "Status": "not yet browsed",
                "Acted": "held" if decision.action == "hold" else decision.action,
            }
        )
        append_audit(cfg, {"event": "fixture", "mds_id": mds_id, "band": decision.band})
    path = write_excel(cfg, rows)
    write_status_report(cfg, rows)
    assert path and path.exists()
    bands = [row["Decision Band"] for row in rows]
    if bands != ["GREEN", "AMBER", "RED"]:
        raise AssertionError(f"fixture bands {bands}")
    return path


def run_self_test(tmp_dir: str | None = None) -> int:
    rc = run_decision_self_test()
    if rc != 0:
        return rc
    output = Path(tmp_dir or os.getenv("IMDS_SELFTEST_DIR", "./imds_output_selftest"))
    if output.exists():
        for child in output.glob("*"):
            if child.is_file():
                child.unlink()
    output.mkdir(parents=True, exist_ok=True)

    os.environ["IMDS_SKIP_VAULT"] = "1"
    os.environ["IMDS_SKIP_BROWSER_DEPS"] = "1"
    os.environ["IMDS_VAULT_PATH"] = str(output / "no-vault.enc")
    # Config fail-fast without secrets
    saved = {k: os.environ.pop(k, None) for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET")}
    try:
        try:
            load_config()
            print("FAIL: load_config should require secrets")
            return 1
        except RuntimeError as exc:
            if "IMDS_USERNAME" not in str(exc):
                print(f"FAIL: unexpected config error: {exc}")
                return 1
            print("OK   config fail-fast without secrets")
    finally:
        for key, value in saved.items():
            if value is not None:
                os.environ[key] = value

    excel_path = _write_fixture_excel(output)
    jsonl = output / "decisions.jsonl"
    report = output / "mds_status_report.csv"
    print(f"OK   fixture excel {excel_path} ({excel_path.stat().st_size} bytes)")
    print(f"OK   audit jsonl   {jsonl} ({jsonl.stat().st_size} bytes)")
    if not report.exists():
        print("FAIL: missing mds_status_report.csv")
        return 1
    print(f"OK   status report {report} ({report.stat().st_size} bytes)")
    print("self-test OK (agent fixtures; no IMDS login)")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="IMDS Playwright agent")
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="Run decision + Excel fixtures without logging into IMDS",
    )
    args = parser.parse_args(argv)
    if args.self_test:
        return run_self_test()
    return orchestrate()


if __name__ == "__main__":
    raise SystemExit(main())
