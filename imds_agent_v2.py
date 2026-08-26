#!/usr/bin/env python3
"""
IMDS Agent — Auto‑check, Excel summary, screenshot of every material node,
rule‑based checks, automated acceptance, forwarding, recipient assignment,
and rejection of failed MDSs.
Enhanced with explicit XPaths and robust Company ID detection using frame_locator.
Overall Result ignores Parts Marking Check per user request.
Acceptance phase uses filter "Browsed" only.
Fixed: previous-version forward prompt must be dismissed with No (Yes
creates a new own-MDS ID and breaks later search/reject).
Fixed: "Do you want to save your changes?" must be answered with Yes so
tab switches after Forward can reach Supplier Data / Recipient Data.
Never JS-strip that dialog — that leaves the new own MDS stuck and later
opens read the leftover ID.
After Forward, IMDS mints a new own-MDS ID (version 0.01). Contact person,
Add Recipient, and Propose must finish on that new ID before the next
received MDS is searched.
Company lookup uses the newest lookupCompany iframe only. Leftover lookup
dialogs are Cancelled (never JS-stripped). Do not Search an empty lookup.
"""

import os
import sys
import time
import logging
import subprocess
import re
import asyncio
from pathlib import Path

try:
    import nest_asyncio
    nest_asyncio.apply()
except ImportError:
    pass

try:
    import pyotp
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyotp"])
    import pyotp

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
from openpyxl import Workbook

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    sync_playwright = None
    class PlaywrightTimeoutError(Exception):
        pass

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger(__name__)

# ---------- Configuration ----------
# Login secrets: Colab 🔑 (IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET) or the environment.
# Never hardcode passwords in this file.
OUTPUT_DIR = os.getenv("IMDS_OUTPUT_DIR", "./imds_output")
DEFAULT_NUM_ITERATIONS = 10


def resolve_num_iterations(raw: str | None = None) -> int:
    """Live default is 10. A leftover Colab/vault value of 3 is treated as unset."""
    if raw is None:
        raw = os.getenv("NUM_ITERATIONS", "")
    text = str(raw or "").strip()
    if not text:
        return DEFAULT_NUM_ITERATIONS
    try:
        n = int(text)
    except ValueError:
        return DEFAULT_NUM_ITERATIONS
    if n < 1:
        return DEFAULT_NUM_ITERATIONS
    # Earlier debug cells and the Drive vault stored 3. Honor any other explicit count.
    if n == 3 and os.getenv("IMDS_ALLOW_THREE") not in {"1", "true", "yes"}:
        return DEFAULT_NUM_ITERATIONS
    return n


NUM_ITERATIONS = resolve_num_iterations()
RECIPIENT_IDS = [x.strip() for x in os.getenv("RECIPIENT_COMPANY_IDS", "9994,293798").split(",") if x.strip()]
IMDS_USERNAME = os.getenv("IMDS_USERNAME", "")
IMDS_PASSWORD = os.getenv("IMDS_PASSWORD", "")
OTP_SECRET = os.getenv("OTP_SECRET", "")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_live_credentials():
    """Pull Colab Secrets / encrypted vault into the module-level login globals."""
    global IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET, NUM_ITERATIONS, RECIPIENT_IDS
    from imds_secrets import apply_stored_credentials, missing_secret_keys

    apply_stored_credentials(persist=True)
    missing = missing_secret_keys()
    if missing:
        raise RuntimeError(
            "Private secrets missing: "
            + ", ".join(missing)
            + ". Add IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET in Colab Secrets (key icon)."
        )
    IMDS_USERNAME = os.environ["IMDS_USERNAME"]
    IMDS_PASSWORD = os.environ["IMDS_PASSWORD"]
    OTP_SECRET = os.environ["OTP_SECRET"]
    NUM_ITERATIONS = resolve_num_iterations()
    log.info(f"Will process up to {NUM_ITERATIONS} MDS rows.")
    RECIPIENT_IDS = [
        x.strip()
        for x in os.getenv("RECIPIENT_COMPANY_IDS", ",".join(RECIPIENT_IDS)).split(",")
        if x.strip()
    ]


def _asyncio_loop_is_running():
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return False
    return True


def _libatk_present():
    return any(
        Path(p).exists()
        for p in (
            "/usr/lib/x86_64-linux-gnu/libatk-1.0.so.0",
            "/usr/lib/aarch64-linux-gnu/libatk-1.0.so.0",
        )
    )


def _ensure_chromium_os_deps():
    """Colab needs libatk; `playwright install chromium` alone is not enough."""
    if os.getenv("IMDS_SKIP_BROWSER_DEPS") in {"1", "true", "yes"}:
        return
    if _libatk_present():
        return
    log.info("Chromium OS libraries missing (libatk). Running playwright install-deps...")
    env = os.environ.copy()
    env["DEBIAN_FRONTEND"] = "noninteractive"
    subprocess.run([sys.executable, "-m", "playwright", "install-deps", "chromium"], env=env, check=False)
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], env=env, check=False)


def _orchestrate_in_subprocess():
    """Playwright Sync API cannot start inside Colab's running asyncio loop."""
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["IMDS_INSIDE_ORCHESTRATE_SUBPROCESS"] = "1"
    script = Path(__file__).resolve()
    log.info("Colab/Jupyter asyncio loop detected — running this script in a subprocess.")
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
    return int(proc.wait() or 0)

# ---------- XPaths for key actions ----------
XP_RECEIVED_MDS_MENU = "//*[@id='pt1:sdiInboxSearch::disAcr']"
XP_COMBINED_ALL = "//*[@id='pt1:dcCmds:sfIbLU:cbAll']/a"
XP_ID_FIELD = "//*[@id='pt1:dcCmds:sfIbLU:itModuleId::content']"
XP_SEARCH_BUTTON = "//*[@id='pt1:dcCmds:sfIbLU:cbSearch']/a"
XP_FIRST_ROW = "//*[@id='pt1:dcCmds:sfIbLU:pc2:tResult::db']/table/tbody/tr"
XP_FIRST_RESULT_NAME = "//*[@id='pt1:dcCmds:sfIbLU:pc2:tResult:0:cName']"
XP_MDS_MENU = "//*[@id='pt1:pt_mFile']/div/table/tbody/tr/td[2]/a"
XP_ACCEPT = "//*[@id='pt1:pt_cmiMenuAccept']/td[2]"
XP_ACCEPT_MODAL = "//*[@id='dcPopup:ctbAcceptMds']/a/span"
XP_FORWARD_MENU = "//*[@id='pt1:pt_mMenuForward']"
XP_FORWARD_ACTION1 = "//*[@id='pt1:pt_cmiMenuForward']/td[1]"
XP_FORWARD_ACTION2 = "//*[@id='pt1:pt_cmiMenuForward']/td[2]"
# Same IDs as above; ADF often puts the clickable text in td[1]/td[2] (Accept already uses /td[2]).
XP_FORWARD_MENU_CLICK = [
    "//*[@id='pt1:pt_mMenuForward']/td[2]",
    "//*[@id='pt1:pt_mMenuForward']/td[1]",
    "//*[@id='pt1:pt_mMenuForward']//a",
    XP_FORWARD_MENU,
    "//*[@id='pt1:pt_cmiMenuForward']/td[2]",
    "//*[@id='pt1:pt_cmiMenuForward']/td[1]",
    "//*[@id='pt1:pt_cmiMenuForward']",
]
XP_FORWARD_OK = "//*[@id='pt1:pt_dcud:ctbOk']/a"
XP_SUPPLIER_DATA = "//*[@id='pt1:sdiDetailSupplier::disAcr']"
XP_CONTACT = "//*[@id='pt1:dcSupp:socContact::content']"
XP_CONTACT_DROP = "//*[@id='pt1:dcSupp:socContact::drop']"
XP_CONTACT_SELECTORS = [
    XP_CONTACT,
    "//*[@id='pt1:dcSupp:socContact']//select",
    "//label[contains(normalize-space(.),'Contact Person')]/following::select[1]",
]
XP_CONTACT_FALLBACKS = [
    "//*[@id='pt1:dcSupp:socContact']",
    "//*[@id='pt1:dcSupp:panelLabelAndMessage26']/td[2]",
    "//*[@id='pt1:dcSupp:panelLabelAndMessage26']/td[2]/table",
    "//*[@id='pt1:dcSupp:panelLabelAndMessage26']/td[2]/table/tbody",
    "//*[@id='pt1:dcSupp:panelLabelAndMessage26']/td[2]/table/tbody/tr",
    "//*[@id='pt1:dcSupp:panelLabelAndMessage26']/td[2]/table/tbody/tr/td",
    "//*[@id='pt1:dcSupp:pglSupplierContact']",
]
XP_RECIPIENT_DATA = "//*[@id='pt1:sdiDetailRecipients::disAcr']"
XP_INGREDIENTS_TAB = [
    "//*[@id='pt1:sdiIngr::disAcr']",
    "//*[@id='pt1:sdiDetailIngredients::disAcr']",
    "//*[@id='pt1:sdiDetailIngr::disAcr']",
    "//*[@id='pt1:sdiIngredients::disAcr']",
]
XP_INGREDIENTS_EXPAND = "//*[@id='pt1:dcIngr:ctbExpandAll']"
XP_TOOLBAR_ACCEPT = [
    "//*[@id='pt1:pt_ctbAccept']",
    "//*[@id='pt1:pt_ctbAccept']//a",
    "//*[@id='pt1:pt_ctbAccept::icon']",
]
XP_TOOLBAR_FORWARD = [
    "//*[@id='pt1:pt_ctbForward']",
    "//*[@id='pt1:pt_ctbForward']//a",
    "//*[@id='pt1:pt_ctbForward::icon']",
]
XP_ADD_RECIPIENT = "//*[@id='pt1:dcReci:ctbAddRecipient::icon']"
XP_ADD_RECIPIENT_FALLBACKS = [
    XP_ADD_RECIPIENT,
    "//*[@id='pt1:dcReci:ctbAddRecipient']",
    "//*[@id='pt1:dcReci:ctbAddRecipient']/a",
    "//*[@id='pt1:dcReci:ctbAddRecipient']/a/span",
    "//*[@id='pt1:dcReci:ctbAddRecipient']/table",
]
XP_PROPOSE = "//*[@id='pt1:dcReci:ctbRecipPropose']/a/span"
XP_PROPOSE_MODAL = "//*[@id='dcPopup:ctbMultiPurpose']/a"

# ---------- XPaths for Reject ----------
XP_REJECT_MENU = "//*[@id='pt1:pt_cmiMenuReject']/td[2]"
XP_REJECT_MODAL = "//*[@id='dcPopup:subViewReject:t1::oc']/table/tbody/tr/td[2]"

# ---------- Exact XPaths for Company ID field ----------
XP_COMPANY_ID_EXACT = [
    "//*[@id='pt1:svSearchCompanyLookup:sfSubLU:it2::content']",
    "//*[@id='pt1:svSearchCompanyLookup:sfSubLU:it2']/td[2]",
    "//*[@id='pt1:svSearchCompanyLookup:sfSubLU:it2']"
]
LOOKUP_COMPANY_IFRAME = "iframe[src*='lookupCompany']"
XP_LOOKUP_CANCEL = [
    "//*[@id='pt1:svSearchCompanyLookup:ctbCancel']/a/span",
    "//*[@id='pt1:svSearchCompanyLookup:ctbCancel']/a",
    "//*[@id='pt1:svSearchCompanyLookup:ctbCancel']",
    "//*[contains(@id,'ctbCancel')]/a/span[normalize-space()='Cancel']",
]

# ---------- XPaths for filter in acceptance phase ----------
XP_FILTER_NONE = "//*[@id='pt1:dcCmds:sfIbLU:cbNone']/a"
XP_FILTER_BROWSED = "//*[@id='pt1:dcCmds:sfIbLU:sbcBrowsed::content']"
# Previous-version forward prompt: Yes auto-forwards and mints a new own-MDS ID.
XP_MODAL_NO = [
    "//*[@id='pt1:pt_dcud:ctbNo']/a",
    "//*[@id='pt1:pt_dcud:ctbNo']/a/span",
    "//*[@id='pt1:pt_dcud:ctbNo']",
    "//*[contains(@id,'ctbNo')]/a",
]
XP_MODAL_YES = [
    "//*[@id='pt1:pt_dcud:ctbYes']/a",
    "//*[@id='pt1:pt_dcud:ctbYes']/a/span",
    "//*[@id='pt1:pt_dcud:ctbYes']",
    "//*[contains(@id,'ctbYes')]/a",
]


def is_forward_previous_version_prompt(text: str) -> bool:
    """True for IMDS 'previous version has been forwarded — forward new version too?'."""
    t = " ".join((text or "").lower().split())
    if "save your changes" in t:
        return False
    if "forward the new version" in t or "do you want to forward the new version" in t:
        return True
    if "you just accepted" in t and "forward" in t:
        return True
    if "has been forwarded" in t and ("previous version" in t or "new version" in t):
        return True
    return False


def is_save_changes_prompt(text: str) -> bool:
    """True for the ADF 'Do you want to save your changes?' Yes/No/Cancel dialog."""
    t = " ".join((text or "").lower().split())
    return "save your changes" in t or "do you want to save" in t


def is_empty_search_criteria_prompt(text: str) -> bool:
    """True for IMDS 'Please enter at least one search criteria!' on an empty lookup."""
    t = " ".join((text or "").lower().split())
    return "at least one search criteria" in t or "enter at least one search" in t


def recipient_id_in_text(tree_text: str, company_id: str) -> bool:
    """True when the recipient tree already lists this company ID, e.g. [9994]."""
    cid = (company_id or "").strip()
    if not cid:
        return False
    return f"[{cid}]" in (tree_text or "")


def contact_name_matches(control_text: str, contact_name: str) -> bool:
    """True when Supplier Data already shows the contact (span or select label)."""
    hay = " ".join((control_text or "").lower().split())
    needle = (contact_name or "").strip().lower()
    if not hay or not needle:
        return False
    if needle in hay:
        return True
    last = needle.split(",")[0].strip()
    return bool(last) and last in hay


def company_id_was_filled(filled_value, company_id: str) -> bool:
    """True only when the lookup Company ID input actually holds the ID we typed."""
    if filled_value is None or company_id is None:
        return False
    return str(filled_value).strip() == str(company_id).strip() and bool(str(company_id).strip())


def should_js_strip_modal(*, lookup_iframes: int, dialog_text: str = "", yes_no: bool = False) -> bool:
    """JS-removing ADF glass panes leaves lookup iframes stacked. Never strip those."""
    if lookup_iframes > 0:
        return False
    if yes_no:
        return False
    if is_save_changes_prompt(dialog_text) or is_forward_previous_version_prompt(dialog_text):
        return False
    if is_empty_search_criteria_prompt(dialog_text):
        return False
    return True


def check_results_present(text: str) -> bool:
    """True when Check has produced a result, even if the Message header table is hidden."""
    t = text or ""
    if re.search(r"\d+\s*Error\(s\)\s*/\s*\d+\s*Warning\(s\)", t):
        return True
    if "passed all included checks" in t.lower():
        return True
    if "check results" in t.lower() and re.search(r"error\(s\)", t, re.I):
        return True
    return False


def parse_mds_id_number(visible: str | None) -> str:
    if not visible or visible == "EXTRACTION_FAILED":
        return ""
    match = re.search(r"(\d{7,})", str(visible))
    return match.group(1) if match else ""


def mds_open_status(visible: str | None, expected: str | None) -> str:
    """Compare numeric MDS IDs only. Version is ignored. None is unknown, not a mismatch."""
    exp_id = parse_mds_id_number(expected)
    vis_id = parse_mds_id_number(visible)
    if not exp_id:
        return "match" if vis_id else "unknown"
    if not vis_id:
        return "unknown"
    return "match" if vis_id == exp_id else "mismatch"


def mds_id_matches(visible: str | None, expected: str | None) -> bool:
    """True when the numeric MDS ID matches. Version (2 vs 0.02 vs 1.01) is ignored."""
    return mds_open_status(visible, expected) == "match"


SUMMARY_COLUMNS = [
    "MDS ID / Version",
    "Check Result",
    "Parts Marking Check",
    "Recyclate Check",
    "Biocidal Check",
    "Overall Result",
    "Supplier Code",
    "Part/Item No.",
    "Action Result",
]


def save_check_summary(results, dest: Path | None = None) -> Path | None:
    """Write the Excel summary Colab displays. Status is omitted; Action Result is the break-out reason."""
    if not results:
        log.warning("No results to export.")
        return None
    dest = dest or Path(OUTPUT_DIR) / "check_summary.xlsx"
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Check Summary"
    ws.append(list(SUMMARY_COLUMNS))
    for row in results:
        ws.append([row.get(col, "") for col in SUMMARY_COLUMNS])
    wb.save(str(dest))
    log.info(f"Excel summary saved to {dest}")
    return dest

def get_otp():
    return pyotp.TOTP(OTP_SECRET.replace(" ", "").upper()).now()

def save_screenshot(page, name):
    path = Path(OUTPUT_DIR) / name
    page.screenshot(path=str(path), full_page=True)
    log.info(f"Screenshot saved: {path}")

# ---------- Login ----------
def imds_login(page):
    log.info("Logging in...")
    try:
        page.goto("https://www.mdsystem.com/imdsnt", wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(2000)
    except Exception as e:
        log.error(f"Failed to load login page: {e}")
        raise

    try:
        login_link = page.locator("a:has-text('Login'):visible").first
        if login_link.count() > 0:
            login_link.click(force=True)
            page.wait_for_timeout(2000)
    except:
        pass

    try:
        username = page.locator("#username, input[name='username']").first
        username.wait_for(state="visible", timeout=30000)
    except PlaywrightTimeoutError:
        log.warning("Username field not found; trying to click 'Sign in' button first.")
        try:
            sso_btn = page.locator("button:has-text('Sign in'):visible, a:has-text('Sign in'):visible").first
            if sso_btn.count() > 0:
                sso_btn.click(force=True)
                page.wait_for_timeout(3000)
                username = page.locator("#username, input[name='username']").first
                username.wait_for(state="visible", timeout=15000)
            else:
                raise RuntimeError("Cannot find username field or sign-in button.")
        except Exception as e:
            log.error(f"Failed to locate username field: {e}")
            save_screenshot(page, "login_username_not_found.png")
            raise

    username.fill(IMDS_USERNAME)
    page.wait_for_timeout(500)

    next_btn = page.locator("button:has-text('Sign in'):visible, button:has-text('Next'):visible, input[value*='Sign in']:visible").first
    if next_btn.count() > 0:
        next_btn.click(force=True)
        page.wait_for_timeout(2000)
    else:
        username.press("Enter")
        page.wait_for_timeout(2000)

    try:
        password = page.locator("input[type='password']:visible").first
        password.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeoutError:
        log.error("Password field not found after username submission.")
        save_screenshot(page, "login_password_not_found.png")
        raise

    password.fill(IMDS_PASSWORD)
    page.wait_for_timeout(500)

    login_btn = page.locator("button:has-text('Sign in'):visible, button:has-text('Login'):visible, input[value*='Sign in']:visible, input[value*='Login']:visible").first
    if login_btn.count() > 0:
        login_btn.click(force=True)
        page.wait_for_timeout(3000)
    else:
        password.press("Enter")
        page.wait_for_timeout(3000)

    try:
        otp_field = page.locator("input[type='text']:visible[placeholder*='code' i], input[type='text']:visible[placeholder*='OTP' i], input[name='otp']:visible, input#otp:visible").first
        otp_field.wait_for(state="visible", timeout=15000)
        otp = get_otp()
        log.info("Filled OTP (value not logged).")
        otp_field.fill(otp)
        page.wait_for_timeout(500)
        submit_otp = page.locator("button:has-text('Verify'):visible, button:has-text('Submit'):visible, input[value*='Verify']:visible").first
        if submit_otp.count() > 0:
            submit_otp.click(force=True)
        else:
            otp_field.press("Enter")
        page.wait_for_load_state("networkidle", timeout=30000)
        page.wait_for_timeout(3000)
    except PlaywrightTimeoutError:
        log.info("No OTP field appeared; assuming login completed without OTP or already logged in.")

    if page.locator("button:has-text('Login'):visible").count() > 0:
        raise RuntimeError("Login failed.")
    log.info("Login successful.")
    save_screenshot(page, "01_after_login.png")


def on_public_login_page(page) -> bool:
    """True on the unauthenticated IMDS landing page (Login link, no Inbox)."""
    try:
        has_inbox = page.locator("#pt1\\:pt_ctbToolBarInbound\\:\\:popEl").count() > 0
        if has_inbox:
            return False
        login_link = page.locator("a:has-text('Login'):visible").first
        forgotten = page.locator("text=User ID forgotten")
        if login_link.count() > 0 and (forgotten.count() > 0 or not has_inbox):
            return login_link.is_visible()
    except Exception:
        return False
    return False

# ---------- Dismiss Modal ----------
def visible_dialog_text(page) -> str:
    for sel in (
        ".AFModalDialog",
        "[id*='pt_dcud']",
        "#pt1\\:pt_dcud",
        ".p_AFModal",
        ".AFBlockingGlassPane",
    ):
        try:
            loc = page.locator(sel)
            if loc.count() > 0:
                text = loc.first.inner_text(timeout=500)
                if text and text.strip():
                    return text
        except Exception:
            continue
    try:
        return page.locator("body").inner_text(timeout=500)[:8000]
    except Exception:
        return ""


def _click_first_matching(page, selectors) -> str | None:
    for selector in selectors:
        try:
            if selector.startswith("//") or selector.startswith("xpath="):
                xp = selector[6:] if selector.startswith("xpath=") else selector
                btn = page.locator(f"xpath={xp}").first
            else:
                btn = page.locator(selector).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(force=True, timeout=4000)
                return selector
        except Exception:
            continue
    return None


def click_modal_no(page, reason: str = "modal") -> bool:
    no_selectors = list(XP_MODAL_NO) + [
        "button:has-text('No'):visible",
        "input[value='No']:visible",
        "a:has-text('No'):visible",
        "span:has-text('No'):visible",
    ]
    used = _click_first_matching(page, no_selectors)
    if not used:
        return False
    log.info(f"Clicked No on {reason} via {used}")
    page.wait_for_timeout(1500)
    try:
        page.wait_for_selector(".AFModalGlassPane, .AFModalDialog", state="detached", timeout=8000)
    except Exception:
        pass
    return True


def click_forward_prompt_no(page) -> bool:
    return click_modal_no(page, reason="previous-version forward prompt")


def click_modal_yes(page, reason: str = "save-changes prompt") -> bool:
    yes_selectors = list(XP_MODAL_YES) + [
        "button:has-text('Yes'):visible",
        "input[value='Yes']:visible",
        "a:has-text('Yes'):visible",
        "span:has-text('Yes'):visible",
    ]
    used = _click_first_matching(page, yes_selectors)
    if not used:
        return False
    log.info(f"Clicked Yes on {reason} via {used}")
    page.wait_for_timeout(1500)
    try:
        page.wait_for_selector(".AFModalGlassPane, .AFModalDialog", state="detached", timeout=8000)
    except Exception:
        pass
    return True


def _locator_visible(page, selector: str) -> bool:
    try:
        loc = page.locator(selector).first
        return loc.count() > 0 and loc.is_visible()
    except Exception:
        return False


def modal_dialog_visible(page) -> bool:
    """True when an ADF dialog is up — including save-changes, which may lack .AFModalGlassPane."""
    for sel in (
        ".AFModalGlassPane",
        ".AFBlockingGlassPane",
        ".AFModalDialog",
        "#pt1\\:pt_dcud",
        "#pt1\\:pt_dcud\\:ctbYes",
        "#pt1\\:pt_dcud\\:ctbNo",
        "#pt1\\:pt_dcud\\:ctbOk",
        "#pt1\\:pt_dcud\\:ctbCancel",
    ):
        if _locator_visible(page, sel):
            return True
    return False


def yes_no_buttons_visible(page) -> bool:
    try:
        no_btn = page.locator("#pt1\\:pt_dcud\\:ctbNo").first
        yes_btn = page.locator("#pt1\\:pt_dcud\\:ctbYes").first
        return (
            no_btn.count() > 0 and no_btn.is_visible()
            and yes_btn.count() > 0 and yes_btn.is_visible()
        )
    except Exception:
        return False


def lookup_company_iframe_count(page) -> int:
    try:
        return page.locator(LOOKUP_COMPANY_IFRAME).count()
    except Exception:
        return 0


def last_lookup_company_frame(page):
    """Return the newest lookupCompany iframe frame. Stacked dialogs must not use .first."""
    loc = page.locator(LOOKUP_COMPANY_IFRAME)
    try:
        n = loc.count()
    except Exception:
        n = 0
    if n <= 0:
        return None
    log.info(f"Using newest lookupCompany iframe ({n} present).")
    try:
        return loc.last.content_frame
    except Exception:
        return None


def wait_for_last_lookup_company_frame(page, timeout_ms: int = 15000):
    deadline = time.time() + timeout_ms / 1000.0
    last_err = None
    while time.time() < deadline:
        loc = page.locator(LOOKUP_COMPANY_IFRAME)
        try:
            n = loc.count()
        except Exception as e:
            last_err = e
            n = 0
        if n > 0:
            try:
                frame = loc.last.content_frame
                if frame:
                    frame.locator("body").wait_for(timeout=2500)
                    log.info(f"Found the lookupCompany iframe (newest of {n}).")
                    return frame
            except Exception as e:
                last_err = e
        page.wait_for_timeout(250)
    log.warning(f"Could not find lookupCompany iframe: {last_err}")
    return None


def _click_cancel_in_frame(frame) -> bool:
    for xp in XP_LOOKUP_CANCEL:
        try:
            btn = frame.locator(f"xpath={xp}").first
            if btn.count() > 0:
                btn.click(force=True, timeout=3000)
                return True
        except Exception:
            continue
    for sel in (
        "a:has-text('Cancel'):visible",
        "span:has-text('Cancel'):visible",
        "button:has-text('Cancel'):visible",
        "input[value='Cancel']:visible",
    ):
        try:
            btn = frame.locator(sel).first
            if btn.count() > 0:
                btn.click(force=True, timeout=3000)
                return True
        except Exception:
            continue
    return False


def close_company_lookup_dialogs(page, max_rounds: int = 8) -> bool:
    """Cancel leftover company-lookup dialogs. Never JS-strip them."""
    for _ in range(max_rounds):
        dialog_text = visible_dialog_text(page)
        if is_empty_search_criteria_prompt(dialog_text):
            used = _click_first_matching(
                page,
                [
                    "#pt1\\:pt_dcud\\:ctbOk",
                    "#pt1\\:pt_dcud\\:ctbOk > a > span",
                    "button:has-text('OK'):visible",
                    "input[value='OK']:visible",
                    "a:has-text('OK'):visible",
                    "span:has-text('OK'):visible",
                ],
            )
            if used:
                log.info(f"Clicked OK on empty search-criteria prompt via {used}")
                page.wait_for_timeout(800)
                continue
        n = lookup_company_iframe_count(page)
        if n <= 0:
            return True
        log.info(f"Cancelling leftover company lookup ({n} iframe(s)).")
        frame = last_lookup_company_frame(page)
        clicked = False
        if frame is not None:
            clicked = _click_cancel_in_frame(frame)
        if not clicked:
            used = _click_first_matching(
                page,
                list(XP_LOOKUP_CANCEL)
                + [
                    "a:has-text('Cancel'):visible",
                    "span:has-text('Cancel'):visible",
                    "button:has-text('Cancel'):visible",
                    "input[value='Cancel']:visible",
                ],
            )
            clicked = bool(used)
            if used:
                log.info(f"Clicked Cancel on company lookup via {used}")
        page.wait_for_timeout(800)
        if lookup_company_iframe_count(page) < n:
            continue
        if not clicked:
            break
    leftover = lookup_company_iframe_count(page)
    if leftover > 0:
        log.warning(
            f"{leftover} lookupCompany iframe(s) still present; not stripping lookup dialogs via JavaScript."
        )
        return False
    return True


def wait_for_glass_pane_clear(page, timeout_ms: int = 8000, allow_escape: bool = True) -> bool:
    """Dismiss leftover IMDS dialogs so they cannot intercept later clicks."""
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        if lookup_company_iframe_count(page) > 0:
            close_company_lookup_dialogs(page)
        if not modal_dialog_visible(page) and lookup_company_iframe_count(page) == 0:
            return True
        dismiss_modal(page, allow_escape=allow_escape)
        page.wait_for_timeout(400)
    return not modal_dialog_visible(page) and lookup_company_iframe_count(page) == 0


def read_visible_mds_id(page) -> str | None:
    """Read ID / Version from the details panel. Do not require Playwright is_visible().

    Scoring already finds this field after a double-click. A missing read means
    the page is still loading, not that a different MDS is open.
    """
    id_ver = re.compile(r"(\d{7,}\s*/\s*[\d.]+)")
    try:
        labels = page.locator("td:has-text('ID / Version')")
        n = min(int(labels.count()), 12)
        for i in range(n):
            cell = labels.nth(i)
            own = ""
            try:
                own = (cell.text_content() or "").strip()
            except Exception:
                own = ""
            found = id_ver.search(own)
            if found:
                return found.group(1)
            try:
                nxt = cell.locator("xpath=following-sibling::td")
                if nxt.count():
                    val = (nxt.first.text_content() or "").strip()
                    found = id_ver.search(val)
                    if found:
                        return found.group(1)
                    if parse_mds_id_number(val):
                        return val
            except Exception:
                continue
    except Exception:
        pass
    try:
        expand = page.locator(f"xpath={XP_INGREDIENTS_EXPAND}")
        if expand.count() == 0:
            return None
        body = page.locator("body").text_content(timeout=1500) or ""
        found = id_ver.search(body)
        if found:
            return found.group(1)
    except Exception:
        pass
    return None


def ingredients_tree_ready(page) -> bool:
    """True when the Ingredients tree/details are in the DOM (visibility not required)."""
    for sel in (
        f"xpath={XP_INGREDIENTS_EXPAND}",
        f"xpath={XP_INGREDIENTS_EXPAND}/table",
        "label:has-text('MDS Supplier')",
    ):
        try:
            if page.locator(sel).count() > 0:
                return True
        except Exception:
            continue
    return False


def dismiss_modal(page, allow_escape: bool = True):
    log.info("Attempting to dismiss modal...")
    if lookup_company_iframe_count(page) > 0:
        log.info("Company lookup dialog is open; cancelling leftover lookup instead of stripping it.")
        close_company_lookup_dialogs(page)
        if lookup_company_iframe_count(page) == 0 and not modal_dialog_visible(page):
            return True
    if not modal_dialog_visible(page):
        log.info("No glass pane, no modal.")
        return True
    log.info("Modal dialog detected.")
    dialog_text = visible_dialog_text(page)
    glass = page.locator(".AFModalGlassPane, .AFBlockingGlassPane")

    if is_forward_previous_version_prompt(dialog_text):
        log.info("Forward-previous-version prompt is showing; clicking No (not Yes).")
        if click_modal_no(page, reason="previous-version forward prompt"):
            return True
        log.warning("Could not click No; leaving the prompt in place rather than clicking Yes.")
        return False

    if is_save_changes_prompt(dialog_text) or (
        yes_no_buttons_visible(page) and "save" in (dialog_text or "").lower()
    ):
        log.info("Save-changes prompt is showing; clicking Yes so tab switches keep the forwarded MDS.")
        if click_modal_yes(page, reason="save-changes prompt"):
            return True
        log.warning("Could not click Yes on save-changes; not stripping the dialog.")
        return False

    if yes_no_buttons_visible(page) and "forward" in (dialog_text or "").lower():
        log.info("Yes/No forward prompt is showing; clicking No (not Yes).")
        if click_modal_no(page, reason="previous-version forward prompt"):
            return True
        return False

    if is_empty_search_criteria_prompt(dialog_text):
        log.info("Empty search-criteria prompt is showing; clicking OK then cancelling the lookup.")
        used = _click_first_matching(
            page,
            [
                "#pt1\\:pt_dcud\\:ctbOk",
                "#pt1\\:pt_dcud\\:ctbOk > a > span",
                "button:has-text('OK'):visible",
                "input[value='OK']:visible",
            ],
        )
        if used:
            log.info(f"Clicked modal button: {used}")
            page.wait_for_timeout(800)
        close_company_lookup_dialogs(page)
        return lookup_company_iframe_count(page) == 0

    for selector in [
        "#pt1\\:pt_dcud\\:ctbOk",
        "#pt1\\:pt_dcud\\:ctbOk > a > span",
        "button:has-text('OK'):visible",
        "input[value='OK']:visible",
        "button:has-text('Close'):visible",
        "input[value='Close']:visible",
        "button:has-text('Proceed'):visible",
    ]:
        try:
            btn = page.locator(selector).first
            if btn.count() > 0 and btn.is_visible():
                btn.click(force=True)
                log.info(f"Clicked modal button: {selector}")
                page.wait_for_timeout(1000)
                if lookup_company_iframe_count(page) > 0:
                    close_company_lookup_dialogs(page)
                if not modal_dialog_visible(page) and lookup_company_iframe_count(page) == 0:
                    return True
        except Exception:
            continue

    if yes_no_buttons_visible(page):
        log.warning("Yes/No dialog still showing; not clicking the glass pane or stripping it.")
        return False

    if lookup_company_iframe_count(page) > 0:
        log.warning("Leaving company lookup in place rather than clicking the glass pane.")
        close_company_lookup_dialogs(page)
        return lookup_company_iframe_count(page) == 0

    try:
        if glass.count() > 0 and glass.first.is_visible():
            glass.first.click(force=True)
            log.info("Clicked glass pane.")
            page.wait_for_timeout(1000)
            if not modal_dialog_visible(page):
                return True
    except Exception:
        pass

    if allow_escape:
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
            if not modal_dialog_visible(page):
                log.info("Pressed Escape, modal dismissed.")
                return True
        except Exception:
            pass

    if yes_no_buttons_visible(page) or is_save_changes_prompt(visible_dialog_text(page)):
        log.warning("Leaving Yes/No save/forward dialog in place; JS strip would freeze the MDS sheet.")
        return False

    lookup_n = lookup_company_iframe_count(page)
    if not should_js_strip_modal(
        lookup_iframes=lookup_n,
        dialog_text=visible_dialog_text(page),
        yes_no=yes_no_buttons_visible(page),
    ):
        log.warning("Not stripping lookup dialogs via JavaScript.")
        return False

    try:
        page.evaluate("""
            () => {
                const panes = document.querySelectorAll('.AFModalGlassPane, .AFBlockingGlassPane');
                panes.forEach(p => p.remove());
                const modals = document.querySelectorAll('.AFModalDialog');
                modals.forEach(m => m.style.display = 'none');
            }
        """)
        log.warning("Removed glass pane via JavaScript.")
        page.wait_for_timeout(500)
        return True
    except Exception as e:
        log.warning(f"JS removal failed: {e}")

    log.warning("Modal dismissal failed.")
    return False

# ---------- Navigate to Search Page ----------
def navigate_to_search_page(page):
    log.info("Navigating to Received MDSs search page...")
    close_company_lookup_dialogs(page)
    if on_public_login_page(page):
        log.warning("Session is on the public login page; logging in again.")
        imds_login(page)
    for _ in range(4):
        dismiss_modal(page, allow_escape=False)
        page.wait_for_timeout(400)
        if not modal_dialog_visible(page):
            break

    try:
        received_mds_link = page.locator("a:has-text('Received MDSs'):visible").first
        if received_mds_link.count() > 0 and received_mds_link.is_visible():
            received_mds_link.click()
            log.info("Clicked 'Received MDSs' link.")
            page.wait_for_timeout(800)
            dismiss_modal(page, allow_escape=False)
            wait_for_glass_pane_clear(page, timeout_ms=5000, allow_escape=False)
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(1500)
            if page.locator(f"xpath={XP_ID_FIELD}").count() > 0:
                log.info("Successfully navigated to search page via Received MDSs link.")
                return True
            if modal_dialog_visible(page):
                dismiss_modal(page, allow_escape=False)
                received_mds_link = page.locator("a:has-text('Received MDSs'):visible").first
                if received_mds_link.count() > 0:
                    received_mds_link.click()
                    page.wait_for_timeout(2000)
                    if page.locator(f"xpath={XP_ID_FIELD}").count() > 0:
                        log.info("Successfully navigated to search page after dismissing save-changes.")
                        return True
    except Exception as e:
        log.warning(f"Error clicking Received MDSs link: {e}")

    try:
        back_btn = page.locator(f"xpath={XP_RECEIVED_MDS_MENU}")
        if back_btn.count() > 0 and back_btn.is_visible():
            back_btn.click(force=True)
            log.info("Clicked MDS Request tab (back).")
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            if page.locator(f"xpath={XP_ID_FIELD}").count() > 0:
                log.info("Successfully navigated to search page via back button.")
                return True
    except Exception as e:
        log.warning(f"Error using back button: {e}")

    try:
        inbox_btn = page.locator("#pt1\\:pt_ctbToolBarInbound\\:\\:popEl")
        if inbox_btn.count() == 0:
            inbox_btn = page.locator("//*[@id='pt1:pt_ctbToolBarInbound::popEl']")
        if inbox_btn.count() > 0 and inbox_btn.is_visible():
            inbox_btn.click()
            page.wait_for_timeout(1500)
            log.info("Clicked Inbox button")
        else:
            log.warning("Inbox button not found; trying to go back.")
            page.go_back()
            page.wait_for_timeout(2000)

        mds_item = page.locator("#pt1\\:pt_cmiSearchInboxB")
        if mds_item.count() == 0:
            dropdown = page.locator("#pt1\\:pt_ctbToolBarInbound_Menu\\:\\:menu")
            if dropdown.count() > 0:
                mds_item = dropdown.locator("a:has-text('MDS'):visible, li:has-text('MDS'):visible").first
        if mds_item.count() > 0 and mds_item.is_visible():
            mds_item.click()
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            log.info("Clicked 'MDS' from Inbox dropdown")
        else:
            log.warning("MDS item not found; navigating to search page failed.")
            return False

        try:
            page.wait_for_selector(f"xpath={XP_ID_FIELD}", timeout=10000)
            log.info("Successfully navigated to search page via Inbox.")
            return True
        except Exception as e:
            log.warning(f"Search page not reached (ID field missing): {e}")
            return False
    except Exception as e:
        log.warning(f"Error navigating to search page via Inbox: {e}")

    log.info("All navigation methods failed; reloading main page and retrying...")
    try:
        if on_public_login_page(page):
            imds_login(page)
        else:
            page.goto("https://www.mdsystem.com/imdsnt")
            page.wait_for_load_state("networkidle", timeout=30000)
            page.wait_for_timeout(3000)
            if on_public_login_page(page):
                imds_login(page)
        inbox_btn = page.locator("#pt1\\:pt_ctbToolBarInbound\\:\\:popEl")
        if inbox_btn.count() == 0:
            inbox_btn = page.locator("//*[@id='pt1:pt_ctbToolBarInbound::popEl']")
        if inbox_btn.count() > 0 and inbox_btn.is_visible():
            inbox_btn.click()
            page.wait_for_timeout(1500)
            mds_item = page.locator("#pt1\\:pt_cmiSearchInboxB")
            if mds_item.count() == 0:
                dropdown = page.locator("#pt1\\:pt_ctbToolBarInbound_Menu\\:\\:menu")
                if dropdown.count() > 0:
                    mds_item = dropdown.locator("a:has-text('MDS'):visible, li:has-text('MDS'):visible").first
            if mds_item.count() > 0 and mds_item.is_visible():
                mds_item.click()
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(2000)
                if page.locator(f"xpath={XP_ID_FIELD}").count() > 0:
                    log.info("Successfully navigated to search page after reload.")
                    return True
    except Exception as e:
        log.warning(f"Reload fallback failed: {e}")

    log.error("Failed to navigate to search page after all attempts.")
    return False

# ---------- Navigate and Filter ----------
def navigate_and_filter(page):
    log.info("Navigating to Received MDSs and applying filter...")
    if not navigate_to_search_page(page):
        raise RuntimeError("Could not navigate to search page")

    log.info("Applying filter...")
    page.evaluate("""
        () => {
            const labels = document.querySelectorAll('label');
            let targetContainer = null;
            for (const label of labels) {
                if (label.textContent.trim() === 'not yet browsed') {
                    let parent = label.parentNode;
                    while (parent) {
                        if (parent.tagName === 'FIELDSET' || parent.tagName === 'DIV') {
                            targetContainer = parent;
                            break;
                        }
                        parent = parent.parentNode;
                    }
                    break;
                }
            }
            if (!targetContainer) {
                const containers = document.querySelectorAll('div, fieldset');
                for (const c of containers) {
                    const labelsInC = c.querySelectorAll('label');
                    let statusCount = 0;
                    for (const lbl of labelsInC) {
                        const text = lbl.textContent.trim();
                        if (['browsed', 'accepted', 'rejected', 'modified', 'cancelled by sender', 'in process at recipient', 'not yet browsed', 'Follow Up'].includes(text)) {
                            statusCount++;
                        }
                    }
                    if (statusCount >= 3) {
                        targetContainer = c;
                        break;
                    }
                }
            }
            if (targetContainer) {
                const cbs = targetContainer.querySelectorAll('input[type="checkbox"]');
                cbs.forEach(cb => { cb.checked = false; });
                const labelsInContainer = targetContainer.querySelectorAll('label');
                for (const lbl of labelsInContainer) {
                    if (lbl.textContent.trim() === 'not yet browsed') {
                        const cb = lbl.querySelector('input[type="checkbox"]');
                        if (cb) {
                            cb.checked = true;
                            const evt = new Event('change', { bubbles: true });
                            cb.dispatchEvent(evt);
                        }
                        break;
                    }
                }
            }
        }
    """)
    page.wait_for_timeout(1000)
    filter_state = page.evaluate("""
        () => {
            const result = {};
            const labels = document.querySelectorAll('label');
            for (const label of labels) {
                const text = label.textContent.trim();
                const cb = label.querySelector('input[type="checkbox"]');
                if (cb) {
                    result[text] = cb.checked;
                }
            }
            return result;
        }
    """)
    log.info(f"Filter state: {filter_state}")
    if not filter_state.get('not yet browsed', False):
        label = page.locator("label:has-text('not yet browsed'):visible").first
        if label.count() > 0:
            label.click()
            log.info("Clicked 'not yet browsed' label as fallback.")
            page.wait_for_timeout(500)
    save_screenshot(page, "04_after_filter_applied.png")
    search_clicked = page.evaluate("""
        (function() {
            var btn = document.querySelector('input[type="submit"][value="Search"]');
            if (!btn) {
                var btns = document.querySelectorAll('button, a');
                for (var i = 0; i < btns.length; i++) {
                    if (btns[i].textContent.trim() === 'Search') {
                        btn = btns[i];
                        break;
                    }
                }
            }
            if (btn) { btn.click(); return true; }
            return false;
        })();
    """)
    if not search_clicked:
        search_text = page.locator(":has-text('Search'):visible").first
        if search_text.count() > 0:
            box = search_text.bounding_box()
            if box:
                page.mouse.click(box['x'] + box['width']/2, box['y'] + box['height']/2)
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(3000)
    log.info("Search submitted.")
    page.wait_for_timeout(1000)
    dismissed = dismiss_modal(page)
    if not dismissed:
        page.wait_for_timeout(2000)
        dismiss_modal(page)
    save_screenshot(page, "05_after_search.png")
    try:
        page.wait_for_selector("table", timeout=15000)
        log.info("Table found.")
    except:
        log.warning("Table not found.")
    page.wait_for_timeout(2000)
    save_screenshot(page, "06_final_list.png")

# ---------- Expand Tree ----------
def expand_tree(page):
    log.info("Expanding MDS tree...")
    expand_selectors = [
        "xpath=//*[@id='pt1:dcIngr:ctbExpandAll']/table/tbody/tr/td[1]",
        "xpath=//*[@id='pt1:dcIngr:ctbExpandAll']//td",
        "xpath=//*[@id='pt1:dcIngr:pt_expand_all']/td[2]",
        "xpath=//*[contains(@id,'expand_all')]",
        "xpath=//*[contains(@title,'Expand All')]",
        "xpath=//*[contains(text(),'Expand All')]",
        "button:has-text('Expand All')"
    ]
    for selector in expand_selectors:
        try:
            if selector.startswith("xpath="):
                loc = page.locator(selector[6:])
            else:
                loc = page.locator(selector)
            if loc.count() > 0 and loc.is_visible():
                loc.click(force=True)
                log.info(f"Clicked Expand All via selector: {selector}")
                page.wait_for_load_state("networkidle", timeout=5000)
                page.wait_for_timeout(1000)
                dismiss_modal(page)
                return
        except Exception as e:
            log.warning(f"Selector {selector} failed: {e}")

    log.warning("Expand All button not found; falling back to iterative expansion...")
    max_iterations = 20
    total_expanded = 0
    previous_count = -1
    for iteration in range(max_iterations):
        expandables = page.locator("[aria-expanded='false']:visible").all()
        expand_images = page.locator("img[src*='btn_tree_expand']:visible, img[src*='expand']:visible, img[src*='plus']:visible").all()
        all_expandables = expandables + expand_images
        tree_items = page.locator("[role='treeitem']:visible").all()
        for item in tree_items:
            icon = item.locator(".af_tree_node_icon:has(img[src*='expand'])")
            if icon.count() > 0:
                all_expandables.append(item)
        if not all_expandables:
            log.info(f"No more expandable nodes found after {iteration+1} iterations.")
            break
        current_count = len(all_expandables)
        if current_count == previous_count:
            log.info("No new expandable elements found; stopping expansion.")
            break
        previous_count = current_count
        log.info(f"Iteration {iteration+1}: Found {len(all_expandables)} expandable elements.")
        clicked_any = False
        for el in all_expandables:
            try:
                if el.is_visible() and el.is_enabled():
                    el.scroll_into_view_if_needed(timeout=5000)
                    el.click(force=True, timeout=5000)
                    clicked_any = True
                    total_expanded += 1
                    page.wait_for_timeout(300)
            except Exception as e:
                log.warning(f"Could not click expandable: {e}")
        page.wait_for_timeout(1000)
        try:
            page.wait_for_load_state("networkidle", timeout=3000)
        except:
            pass
        if not clicked_any:
            log.info("No elements clicked; tree fully expanded.")
            break
    log.info(f"Expanded {total_expanded} nodes in total.")

# ---------- Extract MDS ID ----------
def extract_mds_id_version_early(page):
    try:
        page.wait_for_selector("td:has-text('ID / Version')", timeout=10000)
        cells = page.locator("td").all()
        for cell in cells:
            text = cell.text_content().strip()
            if "ID / Version" in text:
                next_cell = cell.locator("xpath=following-sibling::td")
                if next_cell.count():
                    val = next_cell.text_content().strip()
                    if val:
                        return val
        body = page.locator("body").text_content()
        match = re.search(r'(\d{10,}\s*/\s*[\d.]+)', body)
        if match:
            return match.group(1)
    except Exception as e:
        log.warning(f"ID extraction failed: {e}")
    save_screenshot(page, "id_extraction_failed.png")
    return "EXTRACTION_FAILED"

# ---------- Capture Material Nodes ----------
def capture_all_material_nodes(page, iteration):
    log.info("Capturing screenshots of all material nodes...")
    material_nodes = page.locator("[role='treeitem']:has(img[src*='material'])").all()
    if not material_nodes:
        material_images = page.locator("img[src*='btn_tree_material']").all()
        material_nodes = []
        for img in material_images:
            parent = img.locator("xpath=ancestor::*[@role='treeitem']")
            if parent.count():
                material_nodes.append(parent)
    if not material_nodes:
        log.warning("No material nodes found.")
        return
    log.info(f"Found {len(material_nodes)} material nodes.")
    previous_name = ""
    for idx, node in enumerate(material_nodes, start=1):
        try:
            node.scroll_into_view_if_needed()
            node.click(force=True)
            log.info(f"Clicked material node {idx}")
            page.wait_for_timeout(500)
            try:
                page.wait_for_function(
                    """(prev) => {
                        const nameLabel = document.evaluate(
                            "//label[contains(text(),'Name')]",
                            document,
                            null,
                            XPathResult.FIRST_ORDERED_NODE_TYPE,
                            null
                        ).singleNodeValue;
                        if (!nameLabel) return false;
                        const nameField = nameLabel.closest('tr')?.querySelector('td:last-child') ||
                                           nameLabel.parentElement?.querySelector('span, input, td');
                        if (!nameField) return false;
                        const text = nameField.textContent.trim();
                        return text.length > 0 && text !== prev;
                    }""",
                    arg=previous_name,
                    timeout=8000
                )
            except Exception as e:
                log.warning(f"Name field update timeout for node {idx}: {e}")
            page.wait_for_timeout(500)
            save_screenshot(page, f"mds_iter{iteration}_node{idx}.png")
            try:
                name_field = page.locator("label:has-text('Name')").locator("xpath=..").locator("td:last-child, span, input").first
                if name_field.count():
                    previous_name = name_field.text_content().strip()
            except:
                pass
            dismiss_modal(page)
        except Exception as e:
            log.warning(f"Failed to capture node {idx}: {e}")

# ---------- Extraction helpers ----------
def extract_material_classification(page):
    try:
        label = page.locator("label:has-text('Material class'), label:has-text('Classification'), label:has-text('Material group'), label:has-text('Std. Mat.-No.')").first
        if label.count():
            value = label.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
        body = page.locator("body").text_content()
        match = re.search(r'\b([5]\.[0-9]+\.[a-z]?|6\.1|9\.7|7\.1)\b', body)
        if match:
            return match.group(1)
    except Exception as e:
        log.warning(f"Could not extract classification: {e}")
    return ""

def extract_recyclate_answer(page):
    try:
        label = page.locator("label:has-text('recyclate'), label:has-text('Recycled')").first
        if label.count():
            value = label.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
        label2 = page.locator("label:has-text('recycl')").first
        if label2.count():
            value = label2.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
    except Exception as e:
        log.warning(f"Could not extract recyclate answer: {e}")
    return ""

def extract_biocidal_still_in_production(page):
    try:
        label = page.locator("label:has-text('Still in production'), label:has-text('still in production')").first
        if label.count():
            value = label.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
        label2 = page.locator("label:has-text('production')").first
        if label2.count():
            parent = label2.locator("xpath=..")
            value = parent.locator("td:last-child, span, input").first
            if value.count():
                return value.text_content().strip()
    except Exception as e:
        log.warning(f"Could not extract biocidal 'Still in production?': {e}")
    return ""

def get_component_nodes(page):
    comps = page.locator("[role='treeitem']:has(img[src*='component'])").all()
    if not comps:
        comps = page.locator("[role='treeitem']:has(img[src*='package'])").all()
    return comps

def extract_parts_marking(page):
    try:
        label = page.locator("label:has-text('Part marking'), label:has-text('Parts marking')").first
        if label.count():
            value = label.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
        label2 = page.locator("label:has-text('marking')").first
        if label2.count():
            value = label2.locator("xpath=following-sibling::*").first
            if value.count():
                return value.text_content().strip()
    except Exception as e:
        log.warning(f"Could not extract parts marking: {e}")
    return ""

def get_parent_component(page, material_node_locator):
    parent = page.evaluate(
        """(node) => {
            let parent = node.parentElement;
            while (parent) {
                if (parent.getAttribute('role') === 'treeitem') {
                    const img = parent.querySelector('img');
                    if (img && (img.src.includes('component') || img.src.includes('package'))) {
                        return parent;
                    }
                }
                parent = parent.parentElement;
            }
            return null;
        }""",
        arg=material_node_locator.element_handle()
    )
    if parent:
        return parent.text_content().strip()
    return None

# ---------- Run Checks ----------
def run_checks_on_mds(page, iteration, mds_id):
    log.info(f"Running rules on MDS {iteration}...")
    required_classifications = [
        '5.1.a', '5.1.b', '5.2', '5.3', '5.4', '5.4.1', '5.4.2', '5.4.3',
        '5.5.1', '5.5.2', '6.1', '9.7', '7.1'
    ]

    material_nodes = page.locator("[role='treeitem']:has(img[src*='material'])").all()
    if not material_nodes:
        material_images = page.locator("img[src*='btn_tree_material']").all()
        material_nodes = []
        for img in material_images:
            parent = img.locator("xpath=ancestor::*[@role='treeitem']")
            if parent.count():
                material_nodes.append(parent)

    comp_material_map = {}
    recyclate_fail = False
    biocidal_fail = False

    if not material_nodes:
        log.warning("No material nodes found.")
        return {
            "parts_marking_check": "No materials",
            "recyclate_check": "No materials",
            "biocidal_check": "No materials"
        }

    for mat_node in material_nodes:
        try:
            mat_node.scroll_into_view_if_needed()
            mat_node.click(force=True)
            page.wait_for_timeout(500)
            classification = extract_material_classification(page)
            recyclate = extract_recyclate_answer(page)
            biocidal = extract_biocidal_still_in_production(page)

            if classification in required_classifications:
                if not recyclate or recyclate.lower() == "not yet answered":
                    recyclate_fail = True
                    log.warning(f"Material with classification {classification} has recyclate answer '{recyclate}'.")
                if not biocidal or biocidal.strip() == "":
                    biocidal_fail = True
                    log.warning(f"Material with classification {classification} has empty 'Still in production?'.")

            comp_text = get_parent_component(page, mat_node)
            if comp_text:
                if comp_text not in comp_material_map:
                    comp_material_map[comp_text] = []
                if classification:
                    comp_material_map[comp_text].append(classification)

            dismiss_modal(page)
        except Exception as e:
            log.warning(f"Error processing material for rules: {e}")

    component_nodes = get_component_nodes(page)
    parts_marking_fail = False
    if not component_nodes:
        parts_marking_check_result = "No components"
    else:
        for comp_node in component_nodes:
            try:
                comp_node.scroll_into_view_if_needed()
                comp_node.click(force=True)
                page.wait_for_timeout(500)
                parts_marking = extract_parts_marking(page)
                comp_text = comp_node.text_content().strip()
                if comp_text in comp_material_map:
                    materials_classes = comp_material_map[comp_text]
                    if any(cls in required_classifications for cls in materials_classes):
                        if not parts_marking:
                            parts_marking_fail = True
                            log.warning(f"Component '{comp_text}' has required classification but Parts Marking is empty.")
                dismiss_modal(page)
            except Exception as e:
                log.warning(f"Error processing component for rules: {e}")
        parts_marking_check_result = "PASS" if not parts_marking_fail else "FAIL"

    recyclate_check_result = "PASS" if not recyclate_fail else "FAIL"
    biocidal_check_result = "PASS" if not biocidal_fail else "FAIL"

    return {
        "parts_marking_check": parts_marking_check_result,
        "recyclate_check": recyclate_check_result,
        "biocidal_check": biocidal_check_result
    }

# ---------- Click First Tree Node ----------
def click_first_tree_node(page):
    log.info("Clicking the first tree node...")
    first_node = page.locator("[role='treeitem']:visible, .af_tree_node_text:visible").first
    if first_node.count() > 0:
        try:
            first_node.scroll_into_view_if_needed()
            page.wait_for_timeout(500)
            first_node.click(force=True)
            page.wait_for_timeout(1000)
            dismiss_modal(page)
            log.info("First tree node clicked.")
            return True
        except Exception as e:
            log.warning(f"Failed to click first tree node: {e}")
            return False
    else:
        log.warning("No tree node found.")
        return False

# ---------- Run Built-in Check ----------
def run_check(page):
    log.info("Performing check...")
    try:
        mds_menu = page.locator(f"xpath={XP_MDS_MENU}")
        if mds_menu.count() > 0 and mds_menu.is_visible():
            mds_menu.click(force=True)
            log.info("Clicked MDS menu (exact XPath).")
            page.wait_for_timeout(1000)
        else:
            log.warning("MDS menu not found; trying fallback.")
            mds_menu = page.locator("a:has-text('MDS'):visible, #pt1\\:pt_mFile .x18v:visible").first
            if mds_menu.count() > 0:
                mds_menu.click(force=True)
                log.info("Clicked MDS menu (fallback).")
                page.wait_for_timeout(1000)
            else:
                log.warning("MDS menu not found.")
                return False

        check_item = page.locator("a:has-text('Check'):visible, #pt1\\:pt_cmiMenuCheck:visible").first
        check_item.wait_for(state="visible", timeout=10000)
        if check_item.count() > 0:
            check_item.click(force=True)
            log.info("Clicked Check item.")
            page.wait_for_timeout(2000)
        else:
            log.warning("Check item not found.")
            return False

        if wait_for_check_results(page, timeout_ms=45000):
            log.info("Check results appeared.")
            dismiss_modal(page)
            return True
        log.warning("Check results did not appear (no Error(s)/Warning(s) text).")
        return False
    except Exception as e:
        log.warning(f"Check failed: {e}")
        return False

def wait_for_check_results(page, timeout_ms: int = 45000) -> bool:
    """Wait for Check output. The Message header table is often hidden (pt1:dcCheck:tChkRes::ch::t)."""
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        try:
            body = page.locator("body").text_content(timeout=1000) or ""
        except Exception:
            body = ""
        if check_results_present(body):
            return True
        try:
            panel = page.locator("#pt1\\:dcCheck, [id*='dcCheck']").first
            if panel.count() > 0:
                panel_text = panel.text_content(timeout=500) or ""
                if check_results_present(panel_text):
                    return True
        except Exception:
            pass
        try:
            attached = page.locator("table:has-text('Check results')").first
            if attached.count() > 0:
                txt = attached.text_content(timeout=500) or ""
                if check_results_present(txt) or "check results" in txt.lower():
                    return True
        except Exception:
            pass
        page.wait_for_timeout(500)
    return False

def extract_check_result(page):
    try:
        check_table = page.locator("table:has-text('Check results')").first
        if check_table.count():
            rows = check_table.locator("tr").all()
            for row in rows:
                cells = row.locator("td").all()
                for cell in cells:
                    text = cell.text_content().strip()
                    if text and len(text) > 20 and "Export" not in text and "hidden column" not in text:
                        return text
        body = page.locator("body").text_content()
        match = re.search(r'(The MDS has passed all included checks\.[^.]*\.)', body)
        if match:
            return match.group(1)
        match = re.search(r'(\d+)\s*Error\(s\)\s*/\s*(\d+)\s*Warning\(s\)', body)
        if match:
            errors, warnings = match.groups()
            return f"{errors} Error(s), {warnings} Warning(s)"
        return "Check result not found"
    except Exception as e:
        log.warning(f"Could not extract check result: {e}")
    return "Unknown"

def is_check_clean(result_msg):
    if result_msg == "Check failed" or result_msg == "Check result not found":
        return False
    match = re.search(r'(\d+)\s*Error\(s\)\s*/\s*(\d+)\s*Warning\(s\)', result_msg)
    if match:
        errors = int(match.group(1))
        warnings = int(match.group(2))
        return errors == 0 and warnings == 0
    if "passed all included checks" in result_msg:
        return True
    return False

# ---------- Handle Forward Confirmation Modal ----------
def handle_forward_confirmation_modal(page):
    """Dismiss 'forward the new version as well?' with No.

    Yes auto-forwards and opens a new own MDS (new ID). Later search/open/reject
    then reads that leftover ID instead of the MDS we asked for.
    Save-changes Yes/No/Cancel is handled separately (Yes).
    Returns 'no' if we clicked No, None if no such prompt.
    """
    log.info("Checking for forward confirmation modal...")
    try:
        empty_streak = 0
        for _ in range(10):
            dialog_up = modal_dialog_visible(page)
            dialog_text = visible_dialog_text(page) if dialog_up else ""
            if dialog_up and is_save_changes_prompt(dialog_text):
                log.info("Save-changes prompt after Accept/Forward; clicking Yes.")
                click_modal_yes(page, reason="save-changes prompt")
                empty_streak = 0
                page.wait_for_timeout(400)
                continue
            is_prompt = is_forward_previous_version_prompt(dialog_text) or (
                yes_no_buttons_visible(page) and "forward" in (dialog_text or "").lower()
            )
            if dialog_up and is_prompt:
                log.info("Previous-version forward prompt detected. Clicking No to keep the original MDS ID.")
                if click_modal_no(page, reason="previous-version forward prompt"):
                    wait_for_glass_pane_clear(page, timeout_ms=5000, allow_escape=False)
                    return "no"
                log.warning("No button not found on forward confirmation; not clicking Yes.")
                return None
            page.wait_for_timeout(400)
            if not dialog_up:
                empty_streak += 1
                if empty_streak >= 5:
                    break
            else:
                empty_streak = 0
        log.info("No forward confirmation modal detected.")
        return None
    except Exception as e:
        log.warning(f"Error handling forward confirmation modal: {e}")
    return None

# ---------- Accept MDS ----------
def accept_mds(page):
    log.info("Accepting MDS (PASS workflow) using exact XPaths...")
    try:
        mds_menu = page.locator(f"xpath={XP_MDS_MENU}")
        if mds_menu.count() > 0 and mds_menu.is_visible():
            mds_menu.click(force=True)
            log.info("Clicked MDS menu (exact XPath).")
            page.wait_for_timeout(2000)
        else:
            log.warning("MDS menu not found via exact XPath; trying fallback.")
            mds_menu = page.locator("a:has-text('MDS'):visible, #pt1\\:pt_mFile .x18v:visible").first
            if mds_menu.count() > 0:
                mds_menu.click(force=True)
                log.info("Clicked MDS menu (fallback).")
                page.wait_for_timeout(2000)
            else:
                log.warning("MDS menu not found.")
                save_screenshot(page, "mds_menu_not_found.png")
                return False
    except Exception as e:
        log.warning(f"Failed to click MDS menu: {e}")
        return False

    try:
        accept_btn = page.locator(f"xpath={XP_ACCEPT}")
        if accept_btn.count() > 0 and accept_btn.is_visible():
            accept_btn.click(force=True)
            log.info("Found Accept via exact XPath.")
        else:
            log.warning("Accept button not found via exact XPath; trying fallback.")
            accept_selectors = [
                "#pt1\\:pt_cmiAccept",
                "a[id*='pt_cmiAccept']",
                "li:has-text('Accept')",
                "a:has-text('Accept')",
                "span:has-text('Accept')",
                "td:has-text('Accept')",
                "[role='menuitem']:has-text('Accept')",
                "//*[contains(text(),'Accept')]"
            ]
            accepted = False
            for selector in accept_selectors:
                try:
                    if selector.startswith("//"):
                        locator = page.locator(f"xpath={selector}")
                    else:
                        locator = page.locator(selector)
                    if locator.count() > 0 and locator.is_visible():
                        locator.click(force=True)
                        log.info(f"Found Accept via fallback selector: {selector}")
                        accepted = True
                        break
                except Exception as e:
                    log.warning(f"Fallback selector {selector} failed: {e}")
            if not accepted:
                log.warning("Accept menu item not found; trying Ingredients toolbar Accept.")
                for xp in XP_TOOLBAR_ACCEPT:
                    if _click_xpath_if_present(page, xp):
                        log.info(f"Clicked toolbar Accept via {xp}")
                        accepted = True
                        break
            if not accepted:
                log.warning("Accept menu item not found.")
                save_screenshot(page, "accept_menu_item_not_found.png")
                return False
    except Exception as e:
        log.warning(f"Error clicking Accept: {e}")
        return False

    page.wait_for_timeout(2000)

    try:
        modal_btn = page.locator(f"xpath={XP_ACCEPT_MODAL}")
        if modal_btn.count() > 0 and modal_btn.is_visible():
            modal_btn.click(force=True)
            log.info("Clicked Accept modal button (exact XPath).")
            page.wait_for_timeout(2000)
        else:
            log.warning("Accept modal button not found via exact XPath; using fallback.")
            modal_accept_selectors = [
                "button:has-text('Accept'):visible",
                "input[value='Accept']:visible",
                "button:has-text('OK'):visible",
                "input[value='OK']:visible",
                "button:has-text('Yes'):visible",
                "input[value='Yes']:visible",
                "#pt1\\:pt_dcud\\:ctbOk",
                "//*[contains(@id,'Accept')]//span[contains(text(),'Accept')]"
            ]
            modal_clicked = False
            for selector in modal_accept_selectors:
                try:
                    if selector.startswith("//"):
                        btn = page.locator(f"xpath={selector}").first
                    else:
                        btn = page.locator(selector).first
                    if btn.count() > 0 and btn.is_visible():
                        btn.click(force=True)
                        log.info(f"Found modal Accept button via fallback selector: {selector}")
                        modal_clicked = True
                        break
                except:
                    continue
            if not modal_clicked:
                ok_btn = page.locator("button:has-text('OK'):visible, button:has-text('Confirm'):visible, input[value*='OK']:visible").first
                if ok_btn.count() > 0:
                    ok_btn.click(force=True)
                    log.info("Clicked OK/Confirm on acceptance modal.")
                else:
                    dismiss_modal(page)
    except Exception as e:
        log.warning(f"Error handling Accept modal: {e}")
        dismiss_modal(page)

    try:
        page.wait_for_selector("text='MDS accepted'", timeout=10000)
        log.info("Acceptance confirmed (success message found).")
    except:
        try:
            page.wait_for_selector("button:has-text('Accept'):visible", state="hidden", timeout=10000)
            log.info("Accept button disappeared – likely accepted.")
        except:
            log.warning("Could not confirm acceptance; check manually.")
            save_screenshot(page, "accept_uncertain.png")

    save_screenshot(page, "after_accept.png")
    return True

# ---------- Reject MDS ----------
def reject_mds(page):
    log.info("Rejecting MDS...")
    menu_clicked = False
    for attempt in range(3):
        try:
            mds_menu = page.locator(f"xpath={XP_MDS_MENU}")
            if mds_menu.count() > 0 and mds_menu.is_visible():
                mds_menu.click(force=True)
                log.info("Clicked MDS menu (exact XPath).")
                menu_clicked = True
                break
            else:
                log.warning(f"Attempt {attempt+1}: MDS menu not found, trying fallback.")
                mds_menu = page.locator("a:has-text('MDS'):visible, #pt1\\:pt_mFile .x18v:visible").first
                if mds_menu.count() > 0:
                    mds_menu.click(force=True)
                    log.info("Clicked MDS menu (fallback).")
                    menu_clicked = True
                    break
        except Exception as e:
            log.warning(f"Attempt {attempt+1} to click MDS menu failed: {e}")
            page.wait_for_timeout(1000)

    if not menu_clicked:
        log.warning("MDS menu not found after retries.")
        save_screenshot(page, "mds_menu_not_found_reject.png")
        return False

    page.wait_for_timeout(2000)

    try:
        reject_btn = page.locator(f"xpath={XP_REJECT_MENU}")
        if reject_btn.count() > 0 and reject_btn.is_visible():
            reject_btn.click(force=True)
            log.info("Found Reject via exact XPath.")
        else:
            log.warning("Reject button not found via exact XPath; trying fallback.")
            reject_selectors = [
                "#pt1\\:pt_cmiReject",
                "a[id*='pt_cmiReject']",
                "li:has-text('Reject')",
                "a:has-text('Reject')",
                "span:has-text('Reject')",
                "td:has-text('Reject')",
                "[role='menuitem']:has-text('Reject')",
                "//*[contains(text(),'Reject')]"
            ]
            rejected = False
            for selector in reject_selectors:
                try:
                    if selector.startswith("//"):
                        locator = page.locator(f"xpath={selector}")
                    else:
                        locator = page.locator(selector)
                    if locator.count() > 0 and locator.is_visible():
                        locator.click(force=True)
                        log.info(f"Found Reject via fallback selector: {selector}")
                        rejected = True
                        break
                except Exception as e:
                    log.warning(f"Fallback selector {selector} failed: {e}")
            if not rejected:
                log.warning("Reject menu item not found.")
                save_screenshot(page, "reject_menu_item_not_found.png")
                return False
    except Exception as e:
        log.warning(f"Error clicking Reject: {e}")
        return False

    page.wait_for_timeout(2000)

    try:
        modal_btn = page.locator(f"xpath={XP_REJECT_MODAL}")
        if modal_btn.count() > 0 and modal_btn.is_visible():
            modal_btn.click(force=True)
            log.info("Clicked Reject modal button (exact XPath).")
            page.wait_for_timeout(2000)
        else:
            log.warning("Reject modal button not found via exact XPath; using fallback.")
            modal_reject_selectors = [
                "button:has-text('Reject'):visible",
                "input[value='Reject']:visible",
                "button:has-text('OK'):visible",
                "input[value='OK']:visible",
                "button:has-text('Yes'):visible",
                "input[value='Yes']:visible",
                "#pt1\\:pt_dcud\\:ctbOk",
                "//*[contains(@id,'Reject')]//span[contains(text(),'Reject')]"
            ]
            modal_clicked = False
            for selector in modal_reject_selectors:
                try:
                    if selector.startswith("//"):
                        btn = page.locator(f"xpath={selector}").first
                    else:
                        btn = page.locator(selector).first
                    if btn.count() > 0 and btn.is_visible():
                        btn.click(force=True)
                        log.info(f"Found modal Reject button via fallback selector: {selector}")
                        modal_clicked = True
                        break
                except:
                    continue
            if not modal_clicked:
                ok_btn = page.locator("button:has-text('OK'):visible, button:has-text('Confirm'):visible, input[value*='OK']:visible").first
                if ok_btn.count() > 0:
                    ok_btn.click(force=True)
                    log.info("Clicked OK/Confirm on rejection modal.")
                else:
                    dismiss_modal(page)
    except Exception as e:
        log.warning(f"Error handling Reject modal: {e}")
        dismiss_modal(page)

    try:
        page.wait_for_selector("text='MDS rejected'", timeout=10000)
        log.info("Rejection confirmed (success message found).")
    except:
        try:
            page.wait_for_selector("button:has-text('Reject'):visible", state="hidden", timeout=10000)
            log.info("Reject button disappeared – likely rejected.")
        except:
            log.warning("Could not confirm rejection; check manually.")
            save_screenshot(page, "reject_uncertain.png")

    save_screenshot(page, "after_reject.png")
    return True

def _click_xpath_if_present(page, xpath, *, hover_first: bool = False) -> bool:
    """Click an ADF node if it exists. Do not require is_visible() — submenu wrappers often fail that check."""
    try:
        loc = page.locator(f"xpath={xpath}")
        if loc.count() == 0:
            return False
        target = loc.first
        if hover_first:
            try:
                target.hover(force=True, timeout=3000)
                page.wait_for_timeout(400)
            except Exception:
                pass
        target.click(force=True, timeout=5000)
        return True
    except Exception:
        return False


# ---------- Forward MDS ----------
def forward_mds(page):
    log.info("Forwarding MDS using exact XPaths...")
    dismiss_modal(page)
    for xp in XP_TOOLBAR_FORWARD:
        if _click_xpath_if_present(page, xp):
            log.info(f"Clicked Ingredients toolbar Forward via {xp}")
            page.wait_for_timeout(2000)
            try:
                ok_btn = page.locator(f"xpath={XP_FORWARD_OK}")
                if ok_btn.count() > 0:
                    ok_btn.first.click(force=True)
                    log.info("Clicked OK on Forward modal (exact XPath).")
                    page.wait_for_timeout(2000)
            except Exception:
                dismiss_modal(page)
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            save_screenshot(page, "after_forward.png")
            return True
    try:
        mds_menu = page.locator(f"xpath={XP_MDS_MENU}")
        if mds_menu.count() > 0 and mds_menu.is_visible():
            mds_menu.click(force=True)
            log.info("Clicked MDS menu (exact XPath).")
            page.wait_for_timeout(2000)
        else:
            log.warning("MDS menu not found via exact XPath; trying fallback.")
            mds_menu = page.locator("a:has-text('MDS'):visible, #pt1\\:pt_mFile .x18v:visible").first
            if mds_menu.count() > 0:
                mds_menu.click(force=True)
                log.info("Clicked MDS menu (fallback).")
                page.wait_for_timeout(2000)
            else:
                log.warning("MDS menu not found.")
                save_screenshot(page, "mds_menu_not_found_forward.png")
                return False
    except Exception as e:
        log.warning(f"Failed to click MDS menu: {e}")
        return False

    forward_clicked = False
    try:
        forward_main = page.locator(f"xpath={XP_FORWARD_MENU}")
        if forward_main.count() > 0 and forward_main.is_visible():
            try:
                forward_main.hover(force=True)
                page.wait_for_timeout(400)
            except Exception:
                pass
            forward_main.click(force=True)
            log.info("Clicked Forward main menu item (exact XPath).")
            page.wait_for_timeout(1000)
            forward_clicked = True
        else:
            log.warning("Forward main menu wrapper not visible; trying same-id td/a cells.")
            for xp in XP_FORWARD_MENU_CLICK:
                if _click_xpath_if_present(page, xp, hover_first=True):
                    log.info(f"Clicked Forward via XPath: {xp}")
                    page.wait_for_timeout(800)
                    forward_clicked = True
                    break
            if not forward_clicked:
                forward_selectors = [
                    "xpath=//*[contains(@id,'cmiMenuForward')]",
                    "xpath=//*[contains(@id,'mMenuForward')]",
                    "a:has-text('Forward')",
                    "li:has-text('Forward')",
                    "td:has-text('Forward')",
                    "[role='menuitem']:has-text('Forward')",
                ]
                for selector in forward_selectors:
                    try:
                        if selector.startswith("xpath="):
                            loc = page.locator(selector[6:])
                        else:
                            loc = page.locator(selector)
                        if loc.count() == 0:
                            continue
                        loc.first.click(force=True, timeout=5000)
                        log.info(f"Clicked Forward via fallback selector: {selector}")
                        forward_clicked = True
                        page.wait_for_timeout(800)
                        break
                    except Exception as e:
                        log.warning(f"Fallback selector {selector} failed: {e}")
            if not forward_clicked:
                for xp in XP_TOOLBAR_FORWARD:
                    if _click_xpath_if_present(page, xp):
                        log.info(f"Clicked Ingredients toolbar Forward via {xp}")
                        forward_clicked = True
                        break
            if not forward_clicked:
                log.warning("Forward menu item not found.")
                save_screenshot(page, "forward_menu_not_found.png")
                return False
    except Exception as e:
        log.warning(f"Error clicking Forward main menu: {e}")
        return False

    action_clicked = False
    for xp in [XP_FORWARD_ACTION2, XP_FORWARD_ACTION1]:
        if _click_xpath_if_present(page, xp):
            log.info(f"Clicked Forward action via XPath: {xp}")
            action_clicked = True
            break

    if not action_clicked:
        log.warning("Forward action not found; trying fallback.")
        try:
            action = page.locator("td:has-text('Forward'), a:has-text('Forward')").first
            if action.count() > 0:
                action.click(force=True)
                log.info("Clicked Forward action by text fallback.")
                action_clicked = True
        except Exception as e:
            log.warning(f"Text fallback failed: {e}")

    if not action_clicked and forward_clicked:
        log.info("Forward submenu click may have been the action; continuing.")
        action_clicked = True

    if not action_clicked:
        log.warning("Forward action not found after all attempts.")
        save_screenshot(page, "forward_action_not_found.png")
        return False

    page.wait_for_timeout(2000)

    try:
        ok_btn = page.locator(f"xpath={XP_FORWARD_OK}")
        if ok_btn.count() > 0 and ok_btn.is_visible():
            ok_btn.click(force=True)
            log.info("Clicked OK on Forward modal (exact XPath).")
            page.wait_for_timeout(2000)
            for _ in range(3):
                page.wait_for_timeout(500)
                if page.locator(".AFModalGlassPane").count() == 0:
                    break
                dismiss_modal(page)
        else:
            log.warning("OK modal button not found via exact XPath; trying fallback.")
            ok_btn = page.locator("button:has-text('OK'):visible, input[value='OK']:visible, #pt1\\:pt_dcud\\:ctbOk").first
            if ok_btn.count() > 0 and ok_btn.is_visible():
                ok_btn.click(force=True)
                log.info("Clicked OK on Forward modal (fallback).")
                page.wait_for_timeout(2000)
                for _ in range(3):
                    page.wait_for_timeout(500)
                    if page.locator(".AFModalGlassPane").count() == 0:
                        break
                    dismiss_modal(page)
            else:
                log.warning("OK button not found; trying to dismiss modal.")
                dismiss_modal(page)
                page.wait_for_timeout(2000)
    except Exception as e:
        log.warning(f"Error handling Forward modal: {e}")
        dismiss_modal(page)

    page.wait_for_load_state("networkidle", timeout=15000)
    page.wait_for_timeout(2000)
    save_screenshot(page, "after_forward.png")
    return True

# ---------- Select Contact Person ----------
def select_contact_person(page, contact_name="Qu, Theresa"):
    """Select the supplier contact on the forwarded *own* MDS (ADF select is often not Playwright-visible)."""
    log.info(f"Selecting contact person: {contact_name}")

    def _control_text(loc) -> str:
        if loc is None:
            return ""
        for reader in ("inner_text", "text_content"):
            try:
                text = getattr(loc, reader)(timeout=800)
                if text and str(text).strip():
                    return str(text)
            except Exception:
                continue
        try:
            return str(loc.evaluate(
                """el => {
                    if (!el) return '';
                    if (el.options && el.selectedIndex >= 0) {
                        const opt = el.options[el.selectedIndex];
                        return opt ? (opt.text || '') : '';
                    }
                    return el.textContent || el.innerText || '';
                }"""
            ) or "")
        except Exception:
            return ""

    def _first_select():
        for xp in XP_CONTACT_SELECTORS:
            loc = page.locator(f"xpath={xp}").first
            try:
                if loc.count() > 0:
                    return loc
            except Exception:
                continue
        try:
            loc = page.locator("select[id*='socContact']").first
            if loc.count() > 0:
                return loc
        except Exception:
            pass
        return None

    def _selected() -> bool:
        sel = _first_select()
        if sel is not None and contact_name_matches(_control_text(sel), contact_name):
            return True
        try:
            content = page.locator("xpath=//*[@id='pt1:dcSupp:socContact::content']").first
            if content.count() > 0 and contact_name_matches(_control_text(content), contact_name):
                return True
        except Exception:
            pass
        return False

    try:
        dropdown = _first_select()
        if dropdown is None:
            log.warning("Contact dropdown not found via exact XPath; trying fallback XPaths.")
            found = False
            for fxp in XP_CONTACT_FALLBACKS:
                try:
                    alt = page.locator(f"xpath={fxp}").first
                    if alt.count() > 0:
                        dropdown = alt
                        log.info(f"Found contact control via fallback XPath: {fxp}")
                        found = True
                        break
                except Exception:
                    continue
            if not found:
                log.warning("Contact dropdown not found after all fallbacks.")
                return False

        if _selected():
            log.info(f"Contact already {contact_name}; skipping re-select.")
            save_screenshot(page, "after_contact_selection.png")
            return True

        try:
            tag = ""
            try:
                tag = (dropdown.evaluate("el => (el && el.tagName) || ''") or "").lower()
            except Exception:
                tag = ""
            if tag == "select":
                dropdown.select_option(label=contact_name, timeout=4000)
                log.info(f"Selected {contact_name} via select_option")
                page.wait_for_timeout(400)
                if _selected():
                    save_screenshot(page, "after_contact_selection.png")
                    return True
            else:
                log.info("Contact control is not a <select>; skipping select_option.")
        except Exception as e:
            log.warning(f"select_option failed: {e}")

        try:
            js_ok = dropdown.evaluate(
                """(el, name) => {
                    const sel = el.tagName === 'SELECT' ? el : el.querySelector('select');
                    if (!sel) return false;
                    const opt = [...sel.options].find(o => (o.text || '').includes(name));
                    if (!opt) return false;
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('input', { bubbles: true }));
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }""",
                contact_name,
            )
            if js_ok:
                log.info(f"Selected {contact_name} via DOM change event")
                page.wait_for_timeout(400)
                if _selected():
                    save_screenshot(page, "after_contact_selection.png")
                    return True
        except Exception as e:
            log.warning(f"DOM contact select failed: {e}")

        try:
            drop = page.locator(f"xpath={XP_CONTACT_DROP}").first
            if drop.count() > 0:
                drop.click(force=True, timeout=3000)
            else:
                dropdown.click(force=True, timeout=3000)
            log.info("Clicked contact dropdown.")
            page.wait_for_timeout(800)
        except Exception as e:
            log.warning(f"Contact dropdown click failed: {e}")

        option = None
        locators = [
            f"option:has-text('{contact_name}')",
            f"li:has-text('{contact_name}')",
            f"span:has-text('{contact_name}')",
            f"a:has-text('{contact_name}')",
            f"[role='option']:has-text('{contact_name}')",
            f"text='{contact_name}'",
        ]
        for loc in locators:
            try:
                el = page.locator(loc).first
                if el.count() > 0:
                    option = el
                    if el.is_visible():
                        break
            except Exception:
                continue

        if option is not None and option.count() > 0:
            option.click(force=True)
            log.info(f"Selected {contact_name} from custom dropdown")
            page.wait_for_timeout(400)
            save_screenshot(page, "after_contact_selection.png")
            return _selected() or True

        log.warning(f"Option '{contact_name}' not found.")
        return False
    except Exception as e:
        log.warning(f"Error selecting contact person: {e}")
        return False

# ---------- Complete Forward Recipients ----------
def wait_for_forwarded_own_mds(page, received_id: str, timeout_s: float = 25) -> str:
    """After Forward, IMDS opens a new own MDS (new ID, version 0.01). Stay on it.

    Contact / Add Recipient / Propose must run on this new ID. Searching the
    original received ID while this sheet is still open is what broke the 10-row run.
    """
    received_num = parse_mds_id_number(received_id)
    deadline = time.time() + timeout_s
    last = None
    while time.time() < deadline:
        dismiss_modal(page, allow_escape=False)
        last = read_visible_mds_id(page) or extract_mds_id_version_early(page)
        last_num = parse_mds_id_number(last)
        own = False
        try:
            body = page.locator("body").text_content(timeout=1000) or ""
            own = "component (own mds)" in body.lower() or "forwarded version of a received mds" in body.lower()
        except Exception:
            own = False
        if last_num and received_num and last_num != received_num:
            log.info(
                f"Forward created own MDS {last} from received {received_id}. "
                "Completing contact, recipients, and propose on this new ID (not the received ID)."
            )
            return last_num
        if own and last_num:
            log.info(f"On forwarded own MDS {last}; completing contact/recipients/propose here.")
            return last_num
        page.wait_for_timeout(500)
    log.warning(
        f"Did not see a new own-MDS ID after Forward (still {last!r} vs received {received_id}). "
        "Completing contact/recipients/propose on the open sheet anyway."
    )
    return parse_mds_id_number(last) or received_num or ""


def _open_detail_tab(page, xpath: str, name: str, screenshot: str) -> bool:
    """Click Ingredients/Supplier/Recipient tab, answering save-changes with Yes if it appears."""
    wait_for_glass_pane_clear(page, timeout_ms=4000)
    try:
        tab = page.locator(f"xpath={xpath}")
        if tab.count() == 0:
            log.warning(f"{name} tab not found.")
            return False
        tab.first.click(force=True)
        log.info(f"Clicked {name}.")
        page.wait_for_timeout(800)
        dismiss_modal(page, allow_escape=False)
        wait_for_glass_pane_clear(page, timeout_ms=4000, allow_escape=False)
        page.wait_for_timeout(800)
        save_screenshot(page, screenshot)
        return True
    except Exception as e:
        log.warning(f"Error clicking {name}: {e}")
        return False


def _click_add_recipient(page) -> bool:
    for xp in XP_ADD_RECIPIENT_FALLBACKS:
        if _click_xpath_if_present(page, xp):
            log.info(f"Clicked Add Recipient via {xp}")
            return True
    try:
        btn = page.locator("a:has-text('Add Recipient'):visible, span:has-text('Add Recipient'):visible").first
        if btn.count() > 0:
            btn.click(force=True, timeout=5000)
            log.info("Clicked Add Recipient by visible text.")
            return True
    except Exception:
        pass
    log.warning("Add Recipient button not found.")
    return False


def _recipient_already_on_sheet(page, company_id: str) -> bool:
    cid = (company_id or "").strip()
    if not cid:
        return False
    try:
        loc = page.locator(f"text=[{cid}]")
        if loc.count() > 0:
            return True
    except Exception:
        pass
    try:
        body = page.locator("body").text_content(timeout=2000) or ""
        return recipient_id_in_text(body, cid)
    except Exception:
        return False


def _propose_button_enabled(page) -> bool:
    try:
        btn = page.locator(f"xpath={XP_PROPOSE}").first
        if btn.count() == 0:
            btn = page.locator("a:has-text('Propose'):visible, button:has-text('Propose'):visible").first
        if btn.count() == 0:
            return False
        cls = (btn.get_attribute("class") or "") + " " + (btn.get_attribute("aria-disabled") or "")
        if "p_AFDisabled" in cls or "disabled" in cls.lower():
            return False
        try:
            if not btn.is_enabled():
                return False
        except Exception:
            pass
        return True
    except Exception:
        return False


def _fill_supplier_part_and_propose(page, company_id, supplier_code, part_no):
    """Fill recipient codes and Propose after the company lookup is gone."""
    if supplier_code and supplier_code.strip() not in ['', '-']:
        try:
            supp_code_field = page.locator("xpath=//*[@id='pt1:dcReci:itSuppCode::content']")
            if supp_code_field.count() > 0:
                wait_for_glass_pane_clear(page, timeout_ms=4000)
                supp_code_field.click(force=True, timeout=8000)
                supp_code_field.fill("")
                supp_code_field.fill(supplier_code)
                log.info(f"Filled Supplier Code: {supplier_code}")
        except Exception as e:
            log.warning(f"Error filling Supplier Code: {e}")
            try:
                page.locator("xpath=//*[@id='pt1:dcReci:itSuppCode::content']").fill(supplier_code, force=True, timeout=5000)
                log.info(f"Filled Supplier Code via force fill: {supplier_code}")
            except Exception as e2:
                log.warning(f"Force fill Supplier Code also failed: {e2}")
    else:
        log.warning("No valid Supplier Code to fill.")

    if part_no and part_no.strip():
        try:
            part_field = page.locator("xpath=//*[@id='pt1:dcReci:itprodCode::content']")
            if part_field.count() > 0:
                wait_for_glass_pane_clear(page, timeout_ms=4000)
                part_field.click(force=True, timeout=8000)
                part_field.fill("")
                part_field.fill(part_no)
                log.info(f"Filled Part/Item No.: {part_no}")
                save_screenshot(page, f"after_fill_part_no_{company_id}.png")
        except Exception as e:
            log.warning(f"Error filling Part/Item No.: {e}")
            try:
                page.locator("xpath=//*[@id='pt1:dcReci:itprodCode::content']").fill(part_no, force=True, timeout=5000)
                log.info(f"Filled Part/Item No. via force fill: {part_no}")
                save_screenshot(page, f"after_fill_part_no_{company_id}.png")
            except Exception as e2:
                log.warning(f"Force fill Part/Item No. also failed: {e2}")
    else:
        log.warning("No Part/Item No. to fill.")

    wait_for_glass_pane_clear(page, timeout_ms=5000)
    try:
        propose_btn = page.locator(f"xpath={XP_PROPOSE}")
        if propose_btn.count() == 0:
            propose_btn = page.locator("button:has-text('Propose'):visible").first
        if propose_btn.count() > 0 and propose_btn.is_visible():
            propose_btn.click(force=True)
            log.info("Clicked Propose button.")
            save_screenshot(page, f"after_click_propose_{company_id}.png")
            page.wait_for_timeout(2000)
        else:
            log.warning("Propose button not found.")
            return "Propose Failed"
    except Exception as e:
        log.warning(f"Error clicking Propose: {e}")
        return "Propose Failed"

    try:
        propose_modal_btn = page.locator(f"xpath={XP_PROPOSE_MODAL}")
        if propose_modal_btn.count() > 0 and propose_modal_btn.is_visible():
            propose_modal_btn.click(force=True)
            log.info("Clicked Propose on confirmation modal (exact XPath).")
            page.wait_for_timeout(2000)
            try:
                page.wait_for_selector(".AFModalGlassPane", state="detached", timeout=10000)
                log.info("Propose modal dismissed.")
            except Exception:
                log.warning("Propose modal glass pane did not detach; dismissing without JS strip.")
                dismiss_modal(page, allow_escape=False)
            save_screenshot(page, f"after_propose_modal_{company_id}.png")
        else:
            log.warning("Propose modal button not found via exact XPath; trying fallback.")
            modal_propose = page.locator("button:has-text('Propose'):visible, input[value='Propose']:visible").first
            if modal_propose.count() > 0 and modal_propose.is_visible():
                modal_propose.click(force=True)
                log.info("Clicked Propose on modal (fallback button).")
                page.wait_for_timeout(2000)
                try:
                    page.wait_for_selector(".AFModalGlassPane", state="detached", timeout=10000)
                    log.info("Propose modal dismissed.")
                except Exception:
                    log.warning("Propose modal glass pane did not detach; dismissing without JS strip.")
                    dismiss_modal(page, allow_escape=False)
                save_screenshot(page, f"after_propose_modal_{company_id}.png")
            else:
                ok_btn = page.locator("button:has-text('OK'):visible, input[value='OK']:visible, #pt1\\:pt_dcud\\:ctbOk").first
                if ok_btn.count() > 0 and ok_btn.is_visible():
                    ok_btn.click(force=True)
                    log.info("Clicked OK on modal.")
                    page.wait_for_timeout(2000)
                    save_screenshot(page, f"after_propose_modal_ok_{company_id}.png")
                else:
                    log.warning("Modal buttons not found; dismissing modal.")
                    dismiss_modal(page)
                    save_screenshot(page, f"after_propose_modal_dismiss_{company_id}.png")
    except Exception as e:
        log.warning(f"Error handling propose modal: {e}")
        dismiss_modal(page)

    save_screenshot(page, f"after_recipient_{company_id}.png")
    return None


def complete_forward_recipients(page, supplier_code, part_no):
    log.info("Completing recipient assignment for forwarded MDS...")
    close_company_lookup_dialogs(page)
    wait_for_glass_pane_clear(page, timeout_ms=5000, allow_escape=False)

    contact_ok = False
    for attempt in range(1, 4):
        _open_detail_tab(page, XP_SUPPLIER_DATA, "Supplier Data", "supplier_data.png")
        for _ in range(6):
            dismiss_modal(page, allow_escape=False)
            if select_contact_person(page, "Qu, Theresa"):
                contact_ok = True
                break
            page.wait_for_timeout(500)
        if contact_ok:
            break
        log.warning(f"Contact person not selected on attempt {attempt}/3; retrying on this own MDS.")

    if not contact_ok:
        log.warning("Contact person selection failed on the forwarded own MDS.")

    _open_detail_tab(page, XP_RECIPIENT_DATA, "Recipient Data", "recipient_data.png")

    def add_recipient(company_id):
        log.info(f"Adding recipient with Company ID: {company_id}")
        close_company_lookup_dialogs(page)
        wait_for_glass_pane_clear(page, timeout_ms=4000, allow_escape=False)

        if _recipient_already_on_sheet(page, company_id):
            log.info(
                f"Recipient [{company_id}] already on this MDS; not opening another company lookup."
            )
            if not _propose_button_enabled(page):
                log.info(f"Propose is inactive; treating {company_id} as already proposed.")
                return None
            try:
                node = page.locator(f"text=[{company_id}]").first
                if node.count() > 0:
                    node.click(force=True, timeout=4000)
                    page.wait_for_timeout(500)
            except Exception:
                pass
            return _fill_supplier_part_and_propose(page, company_id, supplier_code, part_no)

        # 1. Click Add Recipient
        if not _click_add_recipient(page):
            return "Add Recipient Failed"

        # 2. Wait for modal glass pane
        try:
            page.wait_for_selector(".AFModalGlassPane", state="attached", timeout=10000)
            log.info("Modal glass pane detected.")
        except Exception as e:
            log.warning(f"Modal glass pane did not appear: {e}")

        # 3. Use the newest lookupCompany iframe (stacked dialogs must not use .first)
        iframes_before = lookup_company_iframe_count(page)
        frame_locator = wait_for_last_lookup_company_frame(page, timeout_ms=15000)
        if frame_locator is None:
            try:
                iframe_locator = page.locator("iframe:visible").last
                if iframe_locator.count() > 0:
                    frame_locator = iframe_locator.content_frame
                    log.info("Fallback: using last visible iframe.")
            except Exception:
                frame_locator = None
        if frame_locator is None:
            log.warning("No iframe found; cannot proceed.")
            save_screenshot(page, "no_iframe.png")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        # 4. Locate the Company ID field (do not wait on input.first — ADF's first
        # input is a hidden RICH_UPDATE token and burns 10s every lookup).
        company_field = None

        # Try exact XPaths with explicit wait
        for xp in XP_COMPANY_ID_EXACT:
            try:
                field = frame_locator.locator(f"xpath={xp}")
                # Wait for it to be visible
                field.wait_for(state="visible", timeout=3000)
                if field.count() > 0 and field.is_visible():
                    company_field = field
                    log.info(f"Found Company ID field via exact XPath: {xp}")
                    break
            except Exception as e:
                log.debug(f"Exact XPath {xp} failed: {e}")

        # Strategy A: label with "Company ID" -> 'for' attribute or sibling
        if company_field is None or company_field.count() == 0:
            try:
                label = frame_locator.locator("label:has-text('Company ID'):visible, label:has-text('Company Id'):visible, label:has-text('ID'):visible").first
                if label.count() > 0 and label.is_visible():
                    for_id = label.get_attribute("for")
                    if for_id:
                        field = frame_locator.locator(f"xpath=//*[@id='{for_id}']")
                        if field.count() > 0 and field.is_visible():
                            company_field = field
                            log.info("Found Company ID field via label's 'for' attribute.")
                    else:
                        # Try sibling or following input
                        parent = label.locator("xpath=..")
                        if parent.count() > 0:
                            field = parent.locator("input:visible").first
                            if field.count() > 0 and field.is_visible():
                                company_field = field
                                log.info("Found Company ID field as sibling input.")
                            else:
                                following = label.locator("xpath=following-sibling::input[1]")
                                if following.count() > 0 and following.is_visible():
                                    company_field = following
                                    log.info("Found Company ID field as following sibling.")
            except Exception as e:
                log.warning(f"Label-based lookup failed: {e}")

        # Strategy B: by placeholder text
        if company_field is None or company_field.count() == 0:
            placeholders = ["Company ID", "Company Id", "ID", "Search", "Enter ID", "Company"]
            for ph in placeholders:
                try:
                    field = frame_locator.locator(f"input[placeholder*='{ph}']:visible").first
                    if field.count() > 0 and field.is_visible():
                        company_field = field
                        log.info(f"Found Company ID field by placeholder: {ph}")
                        break
                except:
                    continue

        # Strategy C: by id containing specific patterns
        if company_field is None or company_field.count() == 0:
            id_patterns = ["CompanyId", "company", "it8", "it2", "svSearchCompanyLookup", "CompanyID"]
            for pattern in id_patterns:
                try:
                    field = frame_locator.locator(f"input[contains(@id, '{pattern}')]:visible").first
                    if field.count() > 0 and field.is_visible():
                        company_field = field
                        log.info(f"Found Company ID field by id pattern: {pattern}")
                        break
                except:
                    continue

        # Strategy D: additional XPaths (fallback)
        if company_field is None or company_field.count() == 0:
            xpaths = [
                "//input[contains(@name, 'CompanyId')]",
                "//input[contains(@aria-label, 'Company ID')]",
                "//input[contains(@id, 'CompanyId')]",
                "//input[contains(@id, 'company')]",
                "//input[contains(@placeholder, 'ID')]"
            ]
            for xp in xpaths:
                try:
                    field = frame_locator.locator(f"xpath={xp}")
                    if field.count() > 0 and field.is_visible():
                        company_field = field
                        log.info(f"Found Company ID field via XPath: {xp}")
                        break
                except:
                    continue

        # Strategy E: fallback to first visible input (excluding buttons)
        if company_field is None or company_field.count() == 0:
            try:
                all_inputs = frame_locator.locator("input:visible").all()
                for inp in all_inputs:
                    type_attr = inp.get_attribute("type") or ""
                    if type_attr.lower() not in ["submit", "button", "reset", "hidden"]:
                        company_field = inp
                        log.info("Fallback: using first visible input (non-button).")
                        break
            except Exception as e:
                log.warning(f"Fallback input selection failed: {e}")

        if company_field is None or company_field.count() == 0:
            log.warning("Company ID field not found after all strategies.")
            save_screenshot(page, "company_id_not_found.png")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        # 5. Fill the Company ID
        filled_value = ""
        try:
            company_field.click()
            company_field.fill("")
            company_field.fill(company_id)
            log.info(f"Filled Company ID with {company_id}")
            save_screenshot(page, f"after_fill_company_id_{company_id}.png")
            filled_value = company_field.input_value()
            if filled_value != company_id:
                log.warning(f"Value mismatch: filled {company_id}, but input has {filled_value}")
        except Exception as e:
            log.warning(f"Error filling Company ID: {e}")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        if not company_id_was_filled(filled_value, company_id):
            log.warning(f"Company ID not filled (got {filled_value!r}); not clicking Search.")
            save_screenshot(page, f"company_id_not_filled_{company_id}.png")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        # 6. Click Search button inside iframe
        try:
            search_btn = frame_locator.locator("xpath=//*[@id='pt1:svSearchCompanyLookup:sfSubLU:cbSearch']/a/span")
            if search_btn.count() == 0:
                search_btn = frame_locator.locator("button:has-text('Search'):visible, input[value='Search']:visible").first
            if search_btn.count() > 0 and search_btn.is_visible():
                search_btn.click(force=True)
                log.info("Clicked Search button inside iframe.")
                save_screenshot(page, f"after_click_search_{company_id}.png")
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(2000)
            else:
                log.warning("Search button not found.")
                close_company_lookup_dialogs(page)
                return "Recipient lookup Failed"
        except Exception as e:
            log.warning(f"Error clicking Search: {e}")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        if is_empty_search_criteria_prompt(visible_dialog_text(page)):
            log.warning("Search ran without criteria; cancelling lookup.")
            close_company_lookup_dialogs(page)
            return "Recipient lookup Failed"

        # 7. Click Apply button inside iframe
        try:
            apply_btn = frame_locator.locator("button:has-text('Apply'):visible, input[value='Apply']:visible").first
            if apply_btn.count() == 0:
                apply_btn = frame_locator.locator("xpath=//*[contains(@id,'Apply')]").first
            if apply_btn.count() > 0 and apply_btn.is_visible():
                apply_btn.click(force=True)
                log.info("Clicked Apply button inside iframe.")
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(1500)
                save_screenshot(page, f"after_click_apply_{company_id}.png")
            else:
                log.warning("Apply button not found; trying JavaScript...")
                frame_locator.locator("body").evaluate("""
                    () => {
                        const btn = document.querySelector('button:has-text("Apply"), input[value="Apply"]') ||
                                   document.querySelector('[id*="Apply"]');
                        if (btn) btn.click();
                    }
                """)
                page.wait_for_timeout(1500)
                save_screenshot(page, f"after_click_apply_{company_id}_js.png")
        except Exception as e:
            log.warning(f"Error clicking Apply: {e}")
            try:
                frame_locator.locator("body").evaluate("""
                    () => {
                        const btn = document.querySelector('button:has-text("Apply"), input[value="Apply"]') ||
                                   document.querySelector('[id*="Apply"]');
                        if (btn) btn.click();
                    }
                """)
                page.wait_for_timeout(2000)
                log.info("Applied via JavaScript fallback.")
                page.wait_for_timeout(1500)
            except Exception as e2:
                log.warning(f"JavaScript fallback also failed: {e2}")
                close_company_lookup_dialogs(page)
                return "Recipient lookup Failed"

        # 8. Wait for this lookup iframe to close (stale iframes may remain until Cancel)
        closed = False
        deadline = time.time() + 15
        while time.time() < deadline:
            n = lookup_company_iframe_count(page)
            if n == 0 or n < iframes_before:
                log.info("Iframe closed after Apply.")
                closed = True
                break
            page.wait_for_timeout(400)
        if not closed:
            log.warning("Iframe did not detach; cancelling leftover lookup.")
            close_company_lookup_dialogs(page)
        wait_for_glass_pane_clear(page, timeout_ms=8000)

        return _fill_supplier_part_and_propose(page, company_id, supplier_code, part_no)

    proposed = []
    failures = []
    for cid in RECIPIENT_IDS:
        cid = cid.strip()
        if not cid:
            continue
        err = add_recipient(cid)
        if err:
            log.warning(f"Failed to add recipient {cid}: {err}")
            failures.append(f"{err} ({cid})")
        else:
            log.info(f"Successfully added recipient {cid}.")
            proposed.append(cid)

    save_screenshot(page, "after_recipients.png")
    log.info("Recipient assignment completed.")
    if not proposed:
        if not contact_ok:
            return False, "Contact person selection Failed"
        return False, failures[0] if failures else "Recipient assignment Failed"
    if failures:
        extra = f"; proposed {', '.join(proposed)}" if proposed else ""
        return False, "; ".join(failures) + extra
    if not contact_ok:
        return True, "Accepted, forwarded, proposed (contact selection failed)"
    return True, "Accepted, forwarded, proposed"

# ---------- Set Search Filters ----------
def set_search_filters(page):
    log.info("Setting search filters...")
    try:
        all_btn = page.locator(f"xpath={XP_COMBINED_ALL}")
        if all_btn.count() > 0 and all_btn.is_visible():
            all_btn.click(force=True)
            log.info("Clicked 'All' in Combined filter (exact XPath).")
            page.wait_for_timeout(500)
        else:
            log.warning("Could not find Combined 'All' button; trying fallback.")
            all_radio = page.locator("label:has-text('all'):visible, input[value='all']:visible, label:has-text('All'):visible").first
            if all_radio.count() > 0 and all_radio.is_visible():
                all_radio.click(force=True)
                log.info("Clicked 'all' option via fallback.")
                page.wait_for_timeout(500)
    except Exception as e:
        log.warning(f"Error setting Combined filter: {e}")

# ---------- Set Filter to Browsed only ----------
def _click_status_none(page) -> None:
    try:
        none_btn = page.locator(f"xpath={XP_FILTER_NONE}")
        if none_btn.count() > 0 and none_btn.is_visible():
            none_btn.click(force=True)
            log.info("Clicked NONE button.")
            page.wait_for_timeout(500)
        else:
            log.warning("NONE button not found; trying fallback.")
            page.evaluate("""
                () => {
                    const statusFieldset = document.querySelector('fieldset:has(label)');
                    if (statusFieldset) {
                        const cbs = statusFieldset.querySelectorAll('input[type="checkbox"]');
                        cbs.forEach(cb => { cb.checked = false; });
                    }
                }
            """)
            page.wait_for_timeout(500)
    except Exception as e:
        log.warning(f"Error clicking NONE button: {e}")


def set_all_status_filter(page):
    """Clear status checkboxes so search-by-ID is not hidden by Browsed-only."""
    log.info("Clearing status filter (all statuses)...")
    _click_status_none(page)


def set_browsed_filter(page):
    log.info("Setting filter to Browsed status only...")
    _click_status_none(page)

    try:
        browsed_cb = page.locator(f"xpath={XP_FILTER_BROWSED}")
        if browsed_cb.count() > 0 and browsed_cb.is_visible():
            is_checked = browsed_cb.is_checked()
            if not is_checked:
                browsed_cb.click(force=True)
                log.info("Clicked Browsed checkbox.")
            else:
                log.info("Browsed checkbox already checked.")
            page.wait_for_timeout(500)
        else:
            log.warning("Browsed checkbox not found; trying fallback label click.")
            label = page.locator("label:has-text('Browsed'):visible").first
            if label.count() > 0 and label.is_visible():
                label.click(force=True)
                log.info("Clicked Browsed label as fallback.")
                page.wait_for_timeout(500)
            else:
                log.warning("Browsed label not found either.")
    except Exception as e:
        log.warning(f"Error clicking Browsed checkbox: {e}")

def click_ingredients_tab(page) -> bool:
    """Ingredients tab is where Accept/Forward toolbar icons become active."""
    for xp in XP_INGREDIENTS_TAB:
        if _click_xpath_if_present(page, xp):
            log.info(f"Clicked Ingredients tab via {xp}")
            page.wait_for_timeout(1500)
            return True
    try:
        tab = page.get_by_text("Ingredients", exact=True)
        if tab.count() > 0:
            tab.first.click(force=True, timeout=5000)
            log.info("Clicked Ingredients tab by visible text.")
            page.wait_for_timeout(1500)
            return True
    except Exception as e:
        log.warning(f"Ingredients tab text click failed: {e}")
    return False


def wait_for_mds_content_page(page, expected_id: str | None = None) -> bool:
    """Stay on the MDS Ingredients page (ID / Version + tree).

    Match the numeric MDS ID only (ignore version). A missing ID read is
    'still loading', not a leftover sheet. Only a *different* numeric ID is a
    mismatch. Do not click Ingredients just because the ID is not parsed yet —
    that interrupted the open in Colab and produced 'Opened MDS None'.
    Do not click Ingredients on a leftover own MDS whose ID already differs.
    """
    close_company_lookup_dialogs(page)
    dismiss_modal(page, allow_escape=False)
    last_id = None
    last_status = "unknown"
    wrong_streak = 0
    tree_ready = False

    for i in range(30):
        wait_for_glass_pane_clear(page, timeout_ms=800, allow_escape=False)
        last_id = read_visible_mds_id(page)
        last_status = mds_open_status(last_id, expected_id)
        tree_ready = ingredients_tree_ready(page)

        if last_status == "match":
            log.info(f"On-screen MDS ID matches {parse_mds_id_number(expected_id) or expected_id}: {last_id}")
            break
        if last_status == "mismatch":
            wrong_streak += 1
            if wrong_streak >= 8:
                log.warning(
                    f"On-screen MDS {last_id} stayed different from expected {expected_id}; "
                    "not clicking Ingredients on the leftover sheet."
                )
                save_screenshot(page, "mds_id_mismatch.png")
                return False
        else:
            wrong_streak = 0

        if last_status == "unknown" and tree_ready and i >= 4:
            extracted = extract_mds_id_version_early(page)
            if extracted and extracted != "EXTRACTION_FAILED":
                last_id = extracted
                last_status = mds_open_status(last_id, expected_id)
            if last_status == "match":
                log.info(f"MDS ID/Version after scoring-style extract: {last_id}")
                break
            if last_status == "mismatch":
                continue
            if not expected_id:
                log.info(
                    f"Ingredients tree is open; ID/Version not parsed yet "
                    f"(expected {expected_id}). Continuing without requiring a full ID/Version string."
                )
                break
        page.wait_for_timeout(400)
    else:
        if last_status == "mismatch":
            log.warning(f"Opened MDS {last_id} does not match expected ID {expected_id}.")
            save_screenshot(page, "mds_id_mismatch.png")
            return False

    if last_status == "mismatch":
        log.warning(f"Opened MDS {last_id} does not match expected ID {expected_id}.")
        save_screenshot(page, "mds_id_mismatch.png")
        return False

    click_ingredients_tab(page)
    landmarks = [
        "td:has-text('ID / Version')",
        f"xpath={XP_INGREDIENTS_EXPAND}",
        f"xpath={XP_INGREDIENTS_EXPAND}/table",
        "label:has-text('MDS Supplier')",
        "label:has-text('Name / Trade name')",
    ]
    ready = False
    for _ in range(20):
        dismiss_modal(page, allow_escape=False)
        if ingredients_tree_ready(page):
            ready = True
            break
        for sel in landmarks:
            try:
                loc = page.locator(sel)
                if loc.count() > 0:
                    ready = True
                    break
            except Exception:
                continue
        if ready:
            break
        page.wait_for_timeout(500)
    if not ready:
        log.warning("MDS Ingredients page did not load. Accept/Forward stay inactive on the search list.")
        save_screenshot(page, "mds_content_page_not_loaded.png")
        return False
    mds_id = read_visible_mds_id(page) or extract_mds_id_version_early(page)
    log.info(f"MDS Ingredients page is open. ID/Version={mds_id}")
    status = mds_open_status(mds_id, expected_id)
    if status == "mismatch":
        log.warning(f"Opened MDS {mds_id} does not match expected ID {expected_id}.")
        save_screenshot(page, "mds_id_mismatch.png")
        return False
    if expected_id and status == "unknown":
        log.warning(
            f"Could not parse ID/Version after open (got {mds_id!r}); "
            f"Ingredients page is open, continuing with expected ID {expected_id}."
        )
    return True


def _first_result_present(page) -> bool:
    try:
        return page.locator(f"xpath={XP_FIRST_RESULT_NAME}").count() > 0
    except Exception:
        return False


def _fill_id_and_search(page, mds_id_num: str) -> bool:
    try:
        id_input = page.locator(f"xpath={XP_ID_FIELD}")
        id_input.click(force=True)
        id_input.fill("")
        id_input.fill(mds_id_num)
        log.info(f"Filled ID field with {mds_id_num}")
    except Exception as e:
        log.warning(f"Failed to fill ID field: {e}.")
        return False
    try:
        search_btn = page.locator(f"xpath={XP_SEARCH_BUTTON}")
        search_btn.click(force=True)
        log.info("Clicked Search button.")
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(3000)
        wait_for_glass_pane_clear(page, timeout_ms=3000)
        return True
    except Exception as e:
        log.warning(f"Failed to click Search button: {e}.")
        return False


def search_mds_by_id(page, mds_id_num: str) -> bool:
    """Return to Received MDSs search, filter Browsed, and search one ID.

    If Browsed-only returns no rows, retry with all statuses. Opening the MDS
    still marks it browsed so Accept stays active.
    """
    wait_for_glass_pane_clear(page, timeout_ms=5000)
    if not navigate_to_search_page(page):
        return False
    wait_for_glass_pane_clear(page, timeout_ms=3000)
    set_browsed_filter(page)
    if not _fill_id_and_search(page, mds_id_num):
        return False
    if _first_result_present(page):
        return True
    log.warning(f"No browsed rows for MDS ID {mds_id_num}; retrying with all statuses.")
    save_screenshot(page, f"no_rows_browsed_{mds_id_num}.png")
    set_all_status_filter(page)
    return _fill_id_and_search(page, mds_id_num)


def _double_click_first_result(page, mds_id_num: str) -> bool:
    wait_for_glass_pane_clear(page, timeout_ms=5000, allow_escape=False)
    name_cell = page.locator(f"xpath={XP_FIRST_RESULT_NAME}")
    row = None
    try:
        if name_cell.count() > 0:
            row = name_cell.first
            log.info("Using first result name cell tResult:0:cName")
        else:
            rows = page.locator(f"xpath={XP_FIRST_ROW}").all()
            if not rows:
                log.warning(f"No rows found for MDS ID {mds_id_num}.")
                save_screenshot(page, f"no_rows_{mds_id_num}.png")
                return False
            row = rows[0]
            log.info("Using first tbody row fallback")
    except Exception as e:
        log.warning(f"Error finding result row: {e}")
        return False

    try:
        row.dblclick(force=True)
        log.info(f"Double-clicked first result for {mds_id_num}")
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(2000)
        wait_for_glass_pane_clear(page, timeout_ms=4000, allow_escape=False)
        return True
    except Exception as e:
        log.warning(f"Double-click failed: {e}")
        try:
            row.click(force=True)
            page.keyboard.press("Enter")
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            wait_for_glass_pane_clear(page, timeout_ms=4000, allow_escape=False)
            log.info("Opened row via click + Enter fallback.")
            return True
        except Exception as e2:
            log.warning(f"Could not open result row: {e2}")
            return False


def open_first_result_on_content_page(page, mds_id_num: str) -> bool:
    """Double-click the first result and wait for the Ingredients page for that MDS ID."""
    for attempt in range(1, 4):
        if not _double_click_first_result(page, mds_id_num):
            return False
        if wait_for_mds_content_page(page, expected_id=mds_id_num):
            save_screenshot(page, f"after_open_content_{mds_id_num}.png")
            return True
        log.warning(
            f"Open attempt {attempt}/3 for {mds_id_num} did not confirm the Ingredients page."
        )
        save_screenshot(page, f"after_doubleclick_{mds_id_num}.png")
        if attempt < 3:
            if not search_mds_by_id(page, mds_id_num):
                return False
    return False


def accept_passed_mds(page, results):
    log.info("Starting acceptance of PASS MDSs using search-by-ID with exact XPaths...")

    for idx, res in enumerate(results):
        if res.get("Overall Result") != "PASS":
            log.info(f"Skipping row {idx+1} (Overall Result: {res['Overall Result']})")
            continue

        mds_id_num = res["MDS ID / Version"].split('/')[0].strip()
        supplier_code = res.get("Supplier Code", "")
        part_no = res.get("Part/Item No.", "")
        log.info(f"Searching for MDS ID: {mds_id_num}")

        if not search_mds_by_id(page, mds_id_num):
            log.warning(f"Could not search for MDS {mds_id_num}. Skipping this MDS.")
            res["Action Result"] = "Search Failed"
            save_check_summary(results)
            continue

        if not open_first_result_on_content_page(page, mds_id_num):
            log.warning(f"Could not open Ingredients page for {mds_id_num}; Accept/Forward would be inactive.")
            res["Action Result"] = "Open Failed"
            save_check_summary(results)
            continue

        if not accept_mds(page):
            log.warning("Acceptance failed; skipping forwarding.")
            res["Action Result"] = "Accept Failed"
            save_check_summary(results)
            continue

        handle_forward_confirmation_modal(page)
        page.wait_for_timeout(1500)
        wait_for_glass_pane_clear(page, timeout_ms=5000)
        current_id = read_visible_mds_id(page) or extract_mds_id_version_early(page)
        auto_forwarded = mds_open_status(current_id, mds_id_num) == "mismatch"
        forward_note = ""
        if auto_forwarded:
            new_id = parse_mds_id_number(current_id) or current_id
            log.warning(
                f"IMDS auto-forwarded after Accept; on-screen ID is {current_id} "
                f"(searched {mds_id_num}). Skipping a second Forward and proposing on {new_id}."
            )
            click_ingredients_tab(page)
            forward_note = "Auto-forwarded"
        else:
            if not wait_for_mds_content_page(page, expected_id=mds_id_num):
                log.warning("Left Ingredients page after Accept; clicking Ingredients before Forward.")
                click_ingredients_tab(page)
                wait_for_mds_content_page(page, expected_id=mds_id_num)
            if not forward_mds(page):
                log.warning("Forwarding failed; still attempting recipient/propose on the open MDS.")
                forward_note = "Forward Failed"
            else:
                log.info("Forward successful.")
            forwarded_id = wait_for_forwarded_own_mds(page, mds_id_num)
            if forwarded_id and forwarded_id != mds_id_num:
                log.info(
                    f"Received MDS {mds_id_num} is now own MDS {forwarded_id}. "
                    "Will not search the received ID again until contact/recipients/propose finish on this sheet."
                )

        ok, recipient_msg = False, "Recipient assignment Failed"
        for cycle in range(1, 4):
            ok, recipient_msg = complete_forward_recipients(page, supplier_code, part_no)
            if ok:
                break
            log.warning(
                f"Propose cycle {cycle}/3 incomplete on the forwarded own MDS "
                f"({recipient_msg}); retrying here, not searching received ID {mds_id_num}."
            )
            close_company_lookup_dialogs(page)
            dismiss_modal(page, allow_escape=False)
            wait_for_glass_pane_clear(page, timeout_ms=4000, allow_escape=False)
        if ok:
            if forward_note == "Forward Failed":
                res["Action Result"] = f"Forward Failed; {recipient_msg}"
            elif forward_note == "Auto-forwarded":
                res["Action Result"] = "Auto-forwarded, proposed"
            else:
                res["Action Result"] = recipient_msg
            log.info("Recipient assignment successful.")
        else:
            prefix = f"{forward_note}; " if forward_note == "Forward Failed" else ""
            res["Action Result"] = prefix + recipient_msg
            log.warning("Recipient assignment incomplete.")
        save_check_summary(results)

        if not navigate_to_search_page(page):
            log.warning("Could not return to Received MDSs search after propose; leftover MDS may remain open.")
        wait_for_glass_pane_clear(page, timeout_ms=4000)

    log.info("Acceptance, forwarding, and recipient assignment completed.")

# ---------- Process FAIL MDSs ----------
def reject_failed_mds(page, results):
    log.info("Starting rejection of FAIL MDSs using search-by-ID...")

    for idx, res in enumerate(results):
        if res.get("Overall Result") != "FAIL":
            log.info(f"Skipping row {idx+1} (Overall Result is not FAIL)")
            continue

        mds_id_num = res["MDS ID / Version"].split('/')[0].strip()
        log.info(f"Rejecting MDS ID: {mds_id_num}")

        if not search_mds_by_id(page, mds_id_num):
            log.warning(f"Could not search for MDS {mds_id_num}. Skipping this MDS.")
            res["Action Result"] = "Search Failed"
            save_check_summary(results)
            continue

        if not open_first_result_on_content_page(page, mds_id_num):
            log.warning(f"Could not open Ingredients page for {mds_id_num}; Reject would be inactive.")
            res["Action Result"] = "Open Failed"
            save_check_summary(results)
            continue

        if not reject_mds(page):
            log.warning(f"Rejection failed for MDS {mds_id_num}.")
            res["Action Result"] = "Reject Failed"
        else:
            log.info(f"Successfully rejected MDS {mds_id_num}.")
            res["Action Result"] = "Rejected"
        save_check_summary(results)

        if not navigate_to_search_page(page):
            log.warning("Could not return to Received MDSs search after reject.")
        wait_for_glass_pane_clear(page, timeout_ms=4000)

    log.info("Rejection of FAIL MDSs completed.")

# ---------- Main Loop ----------
def process_rows_and_export(page):
    back_xpath = "//*[@id='pt1:sdiInboxSearch::disAcr']"
    results = []

    for i in range(NUM_ITERATIONS):
        row_xpath = f"//*[@id='pt1:dcCmds:sfIbLU:pc2:tResult:{i}:cName']"
        log.info(f"Processing MDS row {i+1}/{NUM_ITERATIONS} using XPath: {row_xpath}")

        try:
            row_element = page.locator(row_xpath)
            row_element.wait_for(state="visible", timeout=10000)
            log.info(f"Row {i} found and visible.")
        except Exception as e:
            log.warning(f"Row {i} not found: {e}. Stopping.")
            break

        supplier_code = ""
        part_no = ""
        status = ""
        try:
            row_tr = row_element.locator("xpath=ancestor::tr")
            cells = row_tr.locator("td").all()
            if len(cells) >= 8:
                part_no = cells[3].text_content().strip()
                supplier_code = cells[6].text_content().strip()
                if supplier_code == "-":
                    supplier_code = ""
                status = cells[7].text_content().strip()
                log.info(f"Extracted from list: Supplier Code='{supplier_code}', Part/Item No.='{part_no}', Status='{status}'")
            else:
                log.warning(f"Row has only {len(cells)} columns; cannot extract all data.")
        except Exception as e:
            log.warning(f"Error extracting row data: {e}")

        try:
            row_element.dblclick(force=True)
            log.info("Double-clicked row.")
        except Exception as e:
            log.warning(f"Double-click failed: {e}. Trying fallback.")
            try:
                page.locator(row_xpath).dblclick(force=True)
            except:
                log.warning("Fallback double-click also failed.")
                page.locator(row_xpath).click(force=True)
                page.keyboard.press("Enter")
                log.info("Used click + Enter to open.")

        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(2000)
        dismiss_modal(page)

        mds_id = extract_mds_id_version_early(page)
        log.info(f"Extracted ID/Version: {mds_id}")

        expand_tree(page)
        dismiss_modal(page)

        if mds_id == "EXTRACTION_FAILED":
            mds_id = extract_mds_id_version_early(page)
            log.info(f"Retry ID extraction: {mds_id}")

        capture_all_material_nodes(page, i+1)

        rule_results = run_checks_on_mds(page, i+1, mds_id)

        if click_first_tree_node(page):
            page.wait_for_timeout(1000)
        else:
            log.warning("Could not click first tree node; continuing anyway.")

        check_success = run_check(page)
        if check_success:
            page.wait_for_timeout(1000)
            save_screenshot(page, f"mds_check_{i+1}.png")
            result_msg = extract_check_result(page)
        else:
            result_msg = "Check failed"

        check_clean = is_check_clean(result_msg)
        recyclate_ok = rule_results.get("recyclate_check") == "PASS"
        biocidal_ok = rule_results.get("biocidal_check") == "PASS"
        overall = "PASS" if (check_clean and recyclate_ok and biocidal_ok) else "FAIL"

        results.append({
            "MDS ID / Version": mds_id,
            "Check Result": result_msg,
            "Parts Marking Check": rule_results.get("parts_marking_check", "Unknown"),
            "Recyclate Check": rule_results.get("recyclate_check", "Unknown"),
            "Biocidal Check": rule_results.get("biocidal_check", "Unknown"),
            "Overall Result": overall,
            "Supplier Code": supplier_code,
            "Part/Item No.": part_no,
            "Action Result": "Pending action",
        })
        log.info(f"Row {i+1}: {mds_id} -> Check: {result_msg} | Parts: {rule_results['parts_marking_check']} | Recyclate: {rule_results['recyclate_check']} | Biocidal: {rule_results['biocidal_check']} | Overall: {overall} | Supplier: {supplier_code} | Part: {part_no}")

        if i < NUM_ITERATIONS - 1:
            log.info("Going back to results page...")
            try:
                back_btn = page.locator(back_xpath)
                if back_btn.count() > 0:
                    back_btn.click(force=True)
                else:
                    page.go_back()
                page.wait_for_selector("table", timeout=15000)
            except:
                page.go_back()
            dismiss_modal(page)
            page.wait_for_timeout(2000)

    save_check_summary(results)

    # Process PASS MDSs
    accept_passed_mds(page, results)

    # Process FAIL MDSs
    reject_failed_mds(page, results)

    save_check_summary(results)

# ---------- Orchestration ----------
def orchestrate():
    if os.environ.get("IMDS_INSIDE_ORCHESTRATE_SUBPROCESS") != "1" and _asyncio_loop_is_running():
        return _orchestrate_in_subprocess()

    load_live_credentials()
    _ensure_chromium_os_deps()
    if sync_playwright is None:
        raise RuntimeError("playwright is not installed. In Colab, re-run Cell 1.")
    launch_args = ["--no-sandbox", "--disable-dev-shm-usage"] if Path("/content").exists() else []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=launch_args)
        page = browser.new_page()
        try:
            imds_login(page)
            navigate_and_filter(page)
            process_rows_and_export(page)
            log.info("All done.")
            return 0
        except Exception as e:
            log.error(f"Script failed: {e}")
            try:
                page.screenshot(path=str(Path(OUTPUT_DIR) / "error.png"))
            except Exception:
                pass
            return 1
        finally:
            browser.close()

if __name__ == "__main__":
    raise SystemExit(orchestrate())

