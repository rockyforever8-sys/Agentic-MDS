#!/usr/bin/env python3
"""Original agent loads without secrets and keeps the uploaded XPath set."""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import imds_agent_v2


class OriginalAgentTests(unittest.TestCase):
    def test_source_has_no_hardcoded_imds_password(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        self.assertNotIn("jowk0001", text)
        self.assertIn("load_live_credentials", text)
        self.assertIn("XP_FORWARD_ACTION1", text)
        self.assertIn("XP_FIRST_RESULT_NAME", text)
        self.assertIn("wait_for_mds_content_page", text)
        self.assertIn("XP_INGREDIENTS_EXPAND", text)
        self.assertIn("XP_CONTACT_FALLBACKS", text)
        self.assertIn("def accept_passed_mds", text)
        self.assertIn("def reject_failed_mds", text)
        self.assertIn("a:has-text('Login'):visible", text)
        self.assertIn("previous-version forward prompt", text)
        self.assertIn("def is_save_changes_prompt", text)
        self.assertIn("Save-changes prompt is showing", text)
        self.assertIn("def wait_for_forwarded_own_mds", text)
        self.assertIn("Completing contact, recipients, and propose on this new ID", text)
        self.assertIn("Not completing contact/recipients/propose on the received sheet", text)
        self.assertNotIn("on the open sheet anyway", text)
        self.assertIn("def looks_like_inbox_table_chrome", text)
        self.assertIn("def own_mds_ready_for_recipients", text)
        self.assertIn("def ensure_inbox_row_opened_as_mds", text)
        self.assertIn("pt1:pt_cmiMenuAccept", text)
        self.assertNotIn("Accept button disappeared", text)
        self.assertNotIn("dropdown.count() == 0 or not dropdown.is_visible()", text)
        self.assertIn("retrying with all statuses", text)
        self.assertIn("not clicking Ingredients on the leftover sheet", text)
        self.assertNotIn("Looking for Yes button.", text)
        self.assertNotIn('locator("input").first.wait_for(state="visible"', text)
        self.assertIn("Action Result", text)
        self.assertIn("DEFAULT_NUM_ITERATIONS = 20", text)
        self.assertIn("def resolve_num_iterations", text)
        self.assertIn("def last_lookup_company_frame", text)
        self.assertIn("def close_company_lookup_dialogs", text)
        self.assertIn("Using newest lookupCompany iframe", text)
        self.assertIn("not clicking Search", text)
        self.assertIn("Contact already", text)
        self.assertIn("not stripping lookup dialogs", text)
        self.assertIn("def wait_for_check_results", text)
        self.assertIn("def read_check_results_text", text)
        self.assertIn("def return_to_inbox_results", text)
        self.assertIn("Check waiter missed the panel", text)
        self.assertIn("Continuing with remaining rows", text)
        self.assertIn("def wait_for_connectivity", text)
        self.assertIn("def ensure_imds_session", text)
        self.assertIn("Waiting to reconnect, then retrying this row", text)
        self.assertIn("Recipient [", text)
        self.assertIn("Contact display value:", text)
        self.assertIn("Propose Failed (Contact must be specified)", text)
        self.assertIn("Contact person selection Failed", text)
        self.assertIn("close_check_results_dialog", text)
        self.assertIn("not clicking Ingredients on a leftover or search page", text)
        self.assertIn("not clicking the disabled Propose confirm", text)
        self.assertIn("clicking No so the leftover own MDS is discarded", text)
        self.assertIn("Preferred contact", text)
        self.assertIn("def contact_option_is_usable", text)
        self.assertIn("def parse_contact_option_names", text)
        self.assertIn("allow_fallback", text)
        self.assertIn("save_changes=\"no\"", text)
        self.assertIn("save-changes leave-sheet", text)
        self.assertNotIn('page.frame_locator("iframe[src*=\'lookupCompany\']")', text)
        self.assertNotIn("Fallback: using first visible iframe.", text)
        self.assertNotIn('wait_for_selector("table:has-text(\'Message\')"', text)
        self.assertIn("def imds_chrome_present", text)
        self.assertIn("def wait_for_imds_chrome", text)
        self.assertIn("def session_logged_in_after_reconnect", text)
        self.assertIn("def should_wait_for_network_recovery", text)
        self.assertIn("waiting for Received MDSs / MDS menu", text)
        self.assertIn("could not navigate to search page", text)
        self.assertIn("def is_gadsdl_svhc_update_prompt", text)
        self.assertIn("def acknowledge_gadsdl_svhc_update", text)
        self.assertIn("GADSDL / SVHC", text)
        login_fn = text.split("def imds_login", 1)[1].split("\ndef _sel_count", 1)[0]
        self.assertIn("dismiss_modal", login_fn)
        blockers_fn = text.split("def _dismiss_nav_blockers", 1)[1].split(
            "\ndef _try_search_nav_fast", 1
        )[0]
        self.assertIn("dismiss_modal", blockers_fn)
        fast_fn = text.split("def _try_search_nav_fast", 1)[1].split(
            "\ndef recover_imds_chrome_once", 1
        )[0]
        self.assertIn("Inbox missing; retrying Received MDSs", fast_fn)
        recover_fn = text.split("def recover_imds_chrome_once", 1)[1].split(
            "\ndef leave_own_mds_for_inbox", 1
        )[0]
        self.assertIn("wait_for_imds_chrome", recover_fn)
        self.assertIn("needs_imds_relogin", recover_fn)
        self.assertIn("ensure_imds_session", recover_fn)
        session_fn = text.split("def ensure_imds_session", 1)[1].split(
            "\ndef recover_after_network_error", 1
        )[0]
        self.assertIn("reset_search_nav_chrome_recovery", session_fn)
        check_wait_fn = text.split("def wait_for_check_results", 1)[1].split(
            "\ndef extract_check_result", 1
        )[0]
        self.assertIn("needs_imds_relogin", check_wait_fn)
        leave_fn = text.split("def leave_own_mds_for_inbox", 1)[1].split(
            "\ndef navigate_to_search_page", 1
        )[0]
        self.assertIn('save_changes="no"', leave_fn)
        self.assertIn("_click_received_mds_link", leave_fn)
        self.assertIn("_click_received_mds_menu", leave_fn)
        inbox_fn = text.split("def _click_inbox_mds", 1)[1].split("\ndef _dismiss_nav_blockers", 1)[0]
        self.assertNotIn("go_back()", inbox_fn)
        self.assertNotIn("trying to go back", inbox_fn)
        nav_fn = text.split("def navigate_to_search_page", 1)[1].split(
            "\ndef apply_not_yet_browsed_filter_and_search", 1
        )[0]
        self.assertNotIn("go_back()", nav_fn)
        self.assertIn("logged_in_to_imds", nav_fn)
        self.assertIn("_dismiss_nav_blockers", nav_fn)
        self.assertIn("not looping Inbox", nav_fn)
        self.assertLess(nav_fn.find("_dismiss_nav_blockers"), nav_fn.find("_try_search_nav_fast"))
        self.assertIn("recover_imds_chrome_once", nav_fn)
        self.assertIn("ACCEPT_CONFIRM_WAIT_MS", text)
        self.assertIn("def _wait_and_click_accept_confirm", text)
        self.assertIn("def _cancel_leftover_accept_dialog", text)
        self.assertIn("Not opening MDS menu while a dialog", text)
        self.assertIn("not retrying MDS menu", text)
        self.assertIn("dcPopup:ctbAcceptMds", text)
        ret_fn = text.split("def return_to_inbox_results", 1)[1].split("\ndef process_rows_and_export", 1)[0]
        self.assertNotIn("go_back()", ret_fn)
        self.assertIn("_click_received_mds_link", ret_fn)
        self.assertIn("def recover_inbox_list", text)
        self.assertIn("def apply_not_yet_browsed_filter_and_search", text)
        self.assertIn("def ensure_ingredients_ready_for_check", text)
        self.assertIn("def is_searchable_mds_id", text)
        self.assertIn("Not searching invalid MDS ID", text)
        accept_fn = text.split("def accept_mds", 1)[1].split("\ndef reject_mds", 1)[0]
        self.assertIn("pt1:pt_cmiMenuAccept", accept_fn)
        self.assertNotIn("td:has-text('Accept')", accept_fn)
        self.assertNotIn("[role='menuitem']:has-text('Accept')", accept_fn)
        fwd_fn = text.split("def _click_exact_forward_main", 1)[1].split(
            "\ndef list_contact_option_names", 1
        )[0]
        self.assertNotIn("td:has-text('Forward')", fwd_fn)
        self.assertNotIn("a:has-text('Forward')", fwd_fn)
        self.assertNotIn("pt_dlgUserDialog", fwd_fn)
        self.assertIn("pt1:pt_cmiMenuForward", fwd_fn)
        process_fn = text.split("def process_rows_and_export", 1)[1].split(
            "\ndef orchestrate", 1
        )[0]
        self.assertIn("not checking from the inbox table", process_fn)
        pass_fn = text.split("def accept_passed_mds", 1)[1].split(
            "\ndef reject_failed_mds", 1
        )[0]
        self.assertIn("leaving the inbox without Add Recipient", pass_fn)
        self.assertIn("Forward Failed", pass_fn)
        self.assertIn("leave_own_mds_for_inbox", pass_fn)

    def test_load_live_credentials_requires_secrets(self):
        saved = {k: os.environ.pop(k, None) for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET", "IMDS_MASTER_KEY")}
        os.environ["IMDS_SKIP_VAULT"] = "1"
        try:
            with self.assertRaises(RuntimeError) as ctx:
                imds_agent_v2.load_live_credentials()
            self.assertIn("IMDS_USERNAME", str(ctx.exception))
        finally:
            os.environ.pop("IMDS_SKIP_VAULT", None)
            for key, value in saved.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value


class NumIterationsTests(unittest.TestCase):
    def test_empty_and_invalid_default_to_twenty(self):
        self.assertEqual(imds_agent_v2.resolve_num_iterations(""), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("  "), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("abc"), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("0"), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("-1"), 20)

    def test_leftover_three_and_ten_become_twenty(self):
        saved_three = os.environ.pop("IMDS_ALLOW_THREE", None)
        saved_ten = os.environ.pop("IMDS_ALLOW_TEN", None)
        try:
            self.assertEqual(imds_agent_v2.resolve_num_iterations("3"), 20)
            self.assertEqual(imds_agent_v2.resolve_num_iterations("10"), 20)
            os.environ["IMDS_ALLOW_THREE"] = "1"
            self.assertEqual(imds_agent_v2.resolve_num_iterations("3"), 3)
            os.environ["IMDS_ALLOW_TEN"] = "1"
            self.assertEqual(imds_agent_v2.resolve_num_iterations("10"), 10)
        finally:
            if saved_three is None:
                os.environ.pop("IMDS_ALLOW_THREE", None)
            else:
                os.environ["IMDS_ALLOW_THREE"] = saved_three
            if saved_ten is None:
                os.environ.pop("IMDS_ALLOW_TEN", None)
            else:
                os.environ["IMDS_ALLOW_TEN"] = saved_ten

    def test_explicit_counts_are_honored(self):
        self.assertEqual(imds_agent_v2.resolve_num_iterations("20"), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("30"), 30)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("5"), 5)


class AdfDisabledHelpers(unittest.TestCase):
    def test_element_adf_disabled_true_when_eval_says_so(self):
        class _First:
            def evaluate(self, _js):
                return True

        class _Loc:
            def count(self):
                return 1

            @property
            def first(self):
                return _First()

        self.assertTrue(imds_agent_v2._element_adf_disabled(_Loc()))

    def test_element_adf_disabled_false_when_empty(self):
        class _Loc:
            def count(self):
                return 0

        self.assertFalse(imds_agent_v2._element_adf_disabled(_Loc()))


class ForwardPromptHelpers(unittest.TestCase):
    def test_detects_previous_version_forward_prompt(self):
        text = (
            "You just accepted an MDS where the previous version 1521938290 / 2 "
            "has been forwarded. Do you want to forward the new version as well?"
        )
        self.assertTrue(imds_agent_v2.is_forward_previous_version_prompt(text))
        self.assertFalse(imds_agent_v2.is_forward_previous_version_prompt("Clicked Inbox button"))
        self.assertFalse(imds_agent_v2.is_forward_previous_version_prompt("Forward menu"))
        self.assertFalse(
            imds_agent_v2.is_forward_previous_version_prompt("Do you want to save your changes?")
        )

    def test_detects_save_changes_prompt(self):
        self.assertTrue(imds_agent_v2.is_save_changes_prompt("Do you want to save your changes?"))
        self.assertTrue(imds_agent_v2.is_save_changes_prompt("MDS - MATERIAL DATA SYSTEM\nDo you want to save your changes?\nYes\nNo\nCancel"))
        self.assertFalse(imds_agent_v2.is_save_changes_prompt("Do you want to forward the new version as well?"))
        self.assertFalse(imds_agent_v2.is_save_changes_prompt("Clicked Inbox button"))

    def test_detects_gadsdl_svhc_update_prompt(self):
        wrapped = (
            "GADSDL / SVHC Update\n"
            "GADSL / SVHC has been updated.\n"
            "I understand that a new version of all materials using jokers to hide one of the\n"
            "updated substances has to be created and released to the supply chain according\n"
            "to Rec001 Rule 3.2.1.D.\n"
            "OK\nCancel"
        )
        self.assertTrue(imds_agent_v2.is_gadsdl_svhc_update_prompt(wrapped))
        self.assertTrue(
            imds_agent_v2.is_gadsdl_svhc_update_prompt(
                "GADSL / SVHC Update  I understand  Rec001 Rule 3.2.1.D."
            )
        )
        self.assertFalse(imds_agent_v2.is_gadsdl_svhc_update_prompt("Do you want to save your changes?"))
        self.assertFalse(imds_agent_v2.is_gadsdl_svhc_update_prompt("Clicked Inbox button"))
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(lookup_iframes=0, dialog_text=wrapped, yes_no=False)
        )

    def test_mds_id_matches_numeric_id_only(self):
        self.assertTrue(imds_agent_v2.mds_id_matches("1522070544 / 2", "1522070544"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1522070544 / 2.00", "1522070544"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1503991331 / 0.02", "1503991331"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1521938290 / 1", "1521938290"))
        self.assertFalse(imds_agent_v2.mds_id_matches("1522107776 / 1.01", "1521938290"))
        self.assertFalse(imds_agent_v2.mds_id_matches("1522107776 / 1.01", "1430442417"))
        self.assertFalse(imds_agent_v2.mds_id_matches(None, "1522070544"))
        self.assertEqual(imds_agent_v2.mds_open_status(None, "1522070544"), "unknown")
        self.assertEqual(imds_agent_v2.mds_open_status("EXTRACTION_FAILED", "1522070544"), "unknown")
        self.assertEqual(imds_agent_v2.mds_open_status("1522070544 / 2", "1522070544"), "match")
        self.assertEqual(imds_agent_v2.mds_open_status("1522107776 / 1.01", "1522070544"), "mismatch")
        self.assertEqual(imds_agent_v2.parse_mds_id_number("1522107776 / 1.01"), "1522107776")
        self.assertNotEqual(
            imds_agent_v2.parse_mds_id_number("1522275960 / 0.01"),
            imds_agent_v2.parse_mds_id_number("1522267651 / 1"),
        )


class SummaryExportTests(unittest.TestCase):
    def test_summary_columns_drop_status_and_add_action_result(self):
        self.assertIn("Action Result", imds_agent_v2.SUMMARY_COLUMNS)
        self.assertNotIn("Status", imds_agent_v2.SUMMARY_COLUMNS)

    def test_save_check_summary_writes_action_result(self):
        import tempfile
        from openpyxl import load_workbook

        rows = [{
            "MDS ID / Version": "1522247238 / 1",
            "Check Result": "Check results - 0 Error(s) / 0 Warning(s)",
            "Parts Marking Check": "PASS",
            "Recyclate Check": "PASS",
            "Biocidal Check": "PASS",
            "Overall Result": "PASS",
            "Supplier Code": "606165",
            "Part/Item No.": "1428-1130007",
            "Action Result": "Accepted, forwarded, proposed",
        }]
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "check_summary.xlsx"
            written = imds_agent_v2.save_check_summary(rows, dest)
            self.assertEqual(written, dest)
            wb = load_workbook(dest)
            ws = wb.active
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            self.assertEqual(headers, list(imds_agent_v2.SUMMARY_COLUMNS))
            values = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
            self.assertEqual(values[-1], "Accepted, forwarded, proposed")
            self.assertNotIn("No", values)


class CompanyLookupHelpers(unittest.TestCase):
    def test_empty_search_criteria_prompt(self):
        self.assertTrue(
            imds_agent_v2.is_empty_search_criteria_prompt("Please enter at least one search criteria!")
        )
        self.assertTrue(
            imds_agent_v2.is_empty_search_criteria_prompt(
                "MDS - MATERIAL DATA SYSTEM\nInformation\nPlease enter at least one search criteria!"
            )
        )
        self.assertFalse(imds_agent_v2.is_empty_search_criteria_prompt("Do you want to save your changes?"))
        self.assertFalse(imds_agent_v2.is_empty_search_criteria_prompt("Clicked Search button inside iframe."))

    def test_recipient_id_in_text(self):
        tree = "Johnson Electric Industrial Manufactory Limited [9994] not yet browsed (08/26/2026)"
        self.assertTrue(imds_agent_v2.recipient_id_in_text(tree, "9994"))
        self.assertFalse(imds_agent_v2.recipient_id_in_text(tree, "293798"))
        self.assertFalse(imds_agent_v2.recipient_id_in_text("", "9994"))
        self.assertFalse(imds_agent_v2.recipient_id_in_text(tree, ""))

    def test_contact_name_matches(self):
        self.assertTrue(imds_agent_v2.contact_name_matches("Qu, Theresa", "Qu, Theresa"))
        self.assertTrue(imds_agent_v2.contact_name_matches("Qu, Theresa", "Qu"))
        self.assertFalse(imds_agent_v2.contact_name_matches("-", "Qu, Theresa"))
        self.assertFalse(imds_agent_v2.contact_name_matches("", "Qu, Theresa"))
        self.assertFalse(imds_agent_v2.contact_name_matches("Please select", "Qu, Theresa"))
        option_list = "Please select\nQu, Other\nQu, Theresa\nWong, Kam Yuen"
        self.assertFalse(imds_agent_v2.contact_name_matches(option_list, "Qu, Theresa"))
        self.assertEqual(imds_agent_v2.contact_display_value(option_list), "Please select")
        self.assertTrue(imds_agent_v2.contact_is_blank("-"))
        self.assertTrue(imds_agent_v2.contact_is_blank(""))
        self.assertTrue(imds_agent_v2.contact_is_blank("Please select"))
        self.assertFalse(imds_agent_v2.contact_is_blank("Qu, Theresa"))
        self.assertFalse(imds_agent_v2.contact_option_is_usable(""))
        self.assertFalse(imds_agent_v2.contact_option_is_usable("-"))
        self.assertFalse(imds_agent_v2.contact_option_is_usable("Please select"))
        self.assertTrue(imds_agent_v2.contact_option_is_usable("Qu, Theresa"))
        self.assertTrue(imds_agent_v2.contact_option_is_usable("Liu, Minghui"))
        dump = (
            "option-list: [\n"
            "  '',\n"
            "  'Beenah, Tan',\n"
            "  'Joe, Qiao',\n"
            "  'Liu, Minghui',\n"
            "  'Qu, Theresa',\n"
            "]"
        )
        self.assertEqual(
            imds_agent_v2.parse_contact_option_names(dump),
            ["Beenah, Tan", "Joe, Qiao", "Liu, Minghui", "Qu, Theresa"],
        )
        self.assertEqual(imds_agent_v2.preferred_contact_name(""), "Qu, Theresa")
        self.assertEqual(imds_agent_v2.preferred_contact_name("Liu, Minghui"), "Liu, Minghui")
        self.assertEqual(imds_agent_v2._save_changes_choice("no"), "no")
        self.assertEqual(imds_agent_v2._save_changes_choice("yes"), "yes")
        self.assertEqual(imds_agent_v2._save_changes_choice(None), "yes")

    def test_check_errors_blocking_prompt_and_no_js_strip(self):
        check_text = (
            "Check results - 1 Error(s) / 2 Warning(s)\n"
            "Contact must be specified\n"
            "All existing errors need to be eliminated before any further processing may take place."
        )
        self.assertTrue(imds_agent_v2.is_check_errors_blocking_prompt(check_text))
        self.assertFalse(imds_agent_v2.is_check_errors_blocking_prompt("Do you want to save your changes?"))
        self.assertEqual(
            imds_agent_v2.propose_blocked_message(check_text),
            "Propose Failed (Contact must be specified)",
        )
        self.assertEqual(
            imds_agent_v2.propose_blocked_message(
                "All existing errors need to be eliminated before any further processing may take place."
            ),
            "Propose Failed (Check errors)",
        )
        self.assertIsNone(imds_agent_v2.propose_blocked_message("Clicked Propose button."))
        self.assertFalse(imds_agent_v2.should_js_strip_modal(lookup_iframes=0, dialog_text=check_text, yes_no=False))
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(
                lookup_iframes=0,
                dialog_text="Check results - 1 Error(s) / 2 Warning(s)",
                yes_no=False,
            )
        )
        pass_text = (
            "Check results - 0 Error(s) / 0 Warning(s)\n"
            "The MDS has passed all included checks. These checks do not cover all aspects "
            "of IMDS data requirements. Further manual review may be required."
        )
        self.assertFalse(imds_agent_v2.is_check_errors_blocking_prompt(pass_text))
        self.assertTrue(imds_agent_v2.is_passing_check_results_text(pass_text))
        self.assertTrue(imds_agent_v2.is_check_results_overlay_text(pass_text))
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(lookup_iframes=0, dialog_text=pass_text, yes_no=False)
        )
        self.assertEqual(
            imds_agent_v2.preferred_check_result_message(pass_text),
            "Check results - 0 Error(s) / 0 Warning(s)",
        )
        self.assertTrue(imds_agent_v2.is_check_clean(pass_text))
        self.assertTrue(imds_agent_v2.is_check_clean("0 Error(s), 0 Warning(s)"))
        self.assertFalse(imds_agent_v2.is_check_clean("Check failed"))

    def test_company_id_was_filled(self):
        self.assertTrue(imds_agent_v2.company_id_was_filled("9994", "9994"))
        self.assertTrue(imds_agent_v2.company_id_was_filled("293798", "293798"))
        self.assertFalse(imds_agent_v2.company_id_was_filled("", "9994"))
        self.assertFalse(imds_agent_v2.company_id_was_filled(None, "293798"))
        self.assertFalse(imds_agent_v2.company_id_was_filled("9994", "293798"))

    def test_should_js_strip_modal_never_when_lookup_iframes(self):
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(lookup_iframes=2, dialog_text="", yes_no=False)
        )
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(
                lookup_iframes=1,
                dialog_text="Please enter at least one search criteria!",
                yes_no=False,
            )
        )
        self.assertFalse(
            imds_agent_v2.should_js_strip_modal(
                lookup_iframes=0, dialog_text="Do you want to save your changes?", yes_no=True
            )
        )
        self.assertTrue(imds_agent_v2.should_js_strip_modal(lookup_iframes=0, dialog_text="", yes_no=False))

    def test_check_results_present_without_visible_message_table(self):
        self.assertTrue(imds_agent_v2.check_results_present("Check results - 0 Error(s) / 0 Warning(s)"))
        self.assertTrue(imds_agent_v2.check_results_present("The MDS has passed all included checks."))
        self.assertFalse(imds_agent_v2.check_results_present("Clicked Check item."))
        self.assertFalse(imds_agent_v2.check_results_present(""))
        self.assertEqual(imds_agent_v2.parse_ui_check_counts("Check results - 0 Error(s) / 0 Warning(s)"), (0, 0))
        self.assertTrue(imds_agent_v2.is_passing_check_results_text("Check results - 0 Error(s) / 0 Warning(s)"))


class BrowsedAfterNoneFilterTests(unittest.TestCase):
    def _fn_body(self, name: str) -> str:
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        marker = f"def {name}"
        self.assertIn(marker, text)
        return text.split(marker, 1)[1].split("\ndef ", 1)[0]

    def test_original_filter_xpaths_unchanged(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        self.assertIn("XP_FILTER_NONE = \"//*[@id='pt1:dcCmds:sfIbLU:cbNone']/a\"", text)
        self.assertIn(
            "XP_FILTER_BROWSED = \"//*[@id='pt1:dcCmds:sfIbLU:sbcBrowsed::content']\"",
            text,
        )
        self.assertIn("XP_COMBINED_ALL = \"//*[@id='pt1:dcCmds:sfIbLU:cbAll']/a\"", text)

    def test_browsed_filter_clicks_none_then_forces_browsed(self):
        browsed = self._fn_body("set_browsed_filter")
        none = self._fn_body("_click_status_none")
        self.assertIn("_click_status_none", browsed)
        self.assertIn("XP_FILTER_BROWSED", browsed)
        self.assertIn("Clicked NONE button.", none)
        self.assertIn("Clicked Browsed checkbox.", browsed + self._fn_body("_click_browsed_control"))
        self.assertLess(browsed.find("_click_status_none"), browsed.find("XP_FILTER_BROWSED"))
        # Stale Playwright is_checked() after Combined None must not skip Browsed.
        self.assertNotIn("Browsed checkbox already checked.", browsed)
        self.assertNotIn("already checked", browsed.lower())
        self.assertNotIn("is_checked()", browsed)
        self.assertIn("_wait_native_checkbox", browsed)
        self.assertIn("clicking again", browsed.lower())

    def test_already_checked_does_not_skip_browsed_after_none(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        self.assertNotIn("Browsed checkbox already checked.", text)
        browsed = self._fn_body("set_browsed_filter")
        # After None, Browsed is always clicked; native input.checked is re-read.
        self.assertIn("_click_browsed_control", browsed)
        self.assertIn("_native_checkbox_checked", self._fn_body("_wait_native_checkbox"))
        native = self._fn_body("_native_checkbox_checked")
        self.assertIn("input.checked", native)

    def test_all_status_retry_uses_combined_all(self):
        all_fn = self._fn_body("set_all_status_filter")
        search_fn = self._fn_body("set_search_filters")
        self.assertIn("set_search_filters", all_fn)
        self.assertNotIn("_click_status_none", all_fn)
        self.assertIn("XP_COMBINED_ALL", search_fn)
        self.assertIn("Clicked 'All' in Combined filter", search_fn)
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        self.assertIn("retrying with all statuses", text)
        search_by_id = self._fn_body("search_mds_by_id")
        self.assertIn("set_browsed_filter", search_by_id)
        self.assertIn("set_all_status_filter", search_by_id)
        self.assertLess(search_by_id.find("set_browsed_filter"), search_by_id.find("set_all_status_filter"))


class NetworkResumeTests(unittest.TestCase):
    def test_network_wait_default_is_fifteen_minutes(self):
        self.assertEqual(imds_agent_v2.network_wait_seconds(""), 15 * 60)
        self.assertEqual(imds_agent_v2.network_wait_seconds("20"), 20 * 60)
        self.assertEqual(imds_agent_v2.network_wait_seconds("abc"), 15 * 60)

    def test_transient_network_error_not_xpath_timeout(self):
        self.assertTrue(imds_agent_v2.is_transient_network_error("net::ERR_INTERNET_DISCONNECTED"))
        self.assertTrue(imds_agent_v2.is_transient_network_error("net::ERR_NAME_NOT_RESOLVED"))
        self.assertTrue(imds_agent_v2.is_transient_network_error("Failed to load login page: Timeout"))
        self.assertFalse(imds_agent_v2.is_transient_network_error("Timeout 15000ms exceeded."))
        self.assertFalse(imds_agent_v2.is_transient_network_error("Row 6 not found"))
        self.assertFalse(imds_agent_v2.is_transient_network_error("Could not navigate to search page"))
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery("Could not navigate to search page")
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery(
                RuntimeError("Could not navigate to search page")
            )
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery(
                "Failed to navigate to search page after all attempts."
            )
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery("Could not return to the inbox list")
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery(
                "Inbox button not found; waiting for Received MDSs / MDS menu."
            )
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery("Search page not ready yet (attempt 1/4)")
        )
        self.assertFalse(
            imds_agent_v2.should_wait_for_network_recovery("inbox row index shifted after refresh")
        )
        self.assertTrue(imds_agent_v2.is_inbox_or_search_nav_failure("Could not navigate to search page"))
        self.assertTrue(
            imds_agent_v2.should_wait_for_network_recovery("net::ERR_INTERNET_DISCONNECTED")
        )

    def test_action_result_is_complete(self):
        self.assertTrue(imds_agent_v2.action_result_is_complete("Accepted, forwarded, proposed"))
        self.assertTrue(imds_agent_v2.action_result_is_complete("Rejected"))
        self.assertFalse(imds_agent_v2.action_result_is_complete("Pending action"))
        self.assertFalse(imds_agent_v2.action_result_is_complete("Open Failed"))
        self.assertFalse(imds_agent_v2.action_result_is_complete("Propose Failed (Contact must be specified)"))


class _FakeLocator:
    def __init__(self, n: int, visible: bool):
        self._n = n
        self._visible = visible
        self.first = self

    def count(self) -> int:
        return self._n

    def is_visible(self) -> bool:
        return bool(self._visible and self._n > 0)

    def locator(self, selector: str):
        return _FakeLocator(0, False)


class FakeImdsPage:
    """Minimal Playwright page stand-in for login/chrome detectors."""

    def __init__(self, present=(), url="https://www.mdsystem.com/imdsnt"):
        self.present = set(present)
        self.url = url

    def wait_for_timeout(self, _ms: int):
        return None

    def locator(self, selector: str):
        sel = selector or ""
        sel_l = sel.lower()
        key = None
        if "pt_ctbtoolbarinbound" in sel_l and "menu" not in sel_l:
            key = "inbox"
        elif "pt_cmisearchinboxb" in sel_l:
            key = "inbox"
        elif "sdiinboxsearch" in sel_l:
            key = "received_menu"
        elif "received mdss" in sel_l:
            key = "received_link"
        elif "pt_mfile" in sel_l:
            key = "mds_menu"
        elif "johnson electric" in sel_l:
            key = "org"
        elif "user id forgotten" in sel_l:
            key = "forgotten"
        elif "username" in sel_l or "userid" in sel_l.replace(" ", ""):
            key = "username"
        elif "user id" in sel_l:
            key = "user_id_label"
        elif "button:has-text('login')" in sel_l or (
            "input[value" in sel_l and "login" in sel_l
        ):
            key = "login_button"
        elif "a:has-text('login')" in sel_l:
            key = "login_link"
        hit = key in self.present if key else False
        return _FakeLocator(1 if hit else 0, hit)


class PostLoginSessionTests(unittest.TestCase):
    def test_post_login_page_is_not_public_login(self):
        leftover_login = FakeImdsPage(
            present=("inbox", "login_link", "username"),
            url="https://www.mdsystem.com/imdsnt/faces/login",
        )
        self.assertFalse(imds_agent_v2.on_public_login_page(leftover_login))
        self.assertTrue(imds_agent_v2.logged_in_to_imds(leftover_login))

        splash = FakeImdsPage(
            present=("org", "received_link", "login_link", "username", "forgotten"),
            url="https://www.mdsystem.com/imdsnt?login=1",
        )
        self.assertFalse(imds_agent_v2.on_public_login_page(splash))
        self.assertTrue(imds_agent_v2.logged_in_to_imds(splash))

        login_link_only = FakeImdsPage(present=("login_link",))
        self.assertFalse(imds_agent_v2.on_public_login_page(login_link_only))

    def test_logged_in_when_mds_or_inbox_chrome_present(self):
        inbox = FakeImdsPage(present=("inbox", "login_button"))
        self.assertTrue(imds_agent_v2.imds_chrome_present(inbox))
        self.assertTrue(imds_agent_v2.logged_in_to_imds(inbox))
        self.assertFalse(imds_agent_v2.on_public_login_page(inbox))

        mds_menu = FakeImdsPage(present=("mds_menu",))
        self.assertTrue(imds_agent_v2.logged_in_to_imds(mds_menu))

        received = FakeImdsPage(present=("received_menu",))
        self.assertTrue(imds_agent_v2.logged_in_to_imds(received))

        public = FakeImdsPage(present=("username", "login_button", "forgotten", "login_link"))
        self.assertTrue(imds_agent_v2.on_public_login_page(public))
        self.assertFalse(imds_agent_v2.logged_in_to_imds(public))

    def test_reconnect_login_success_is_not_still_logged_out(self):
        settling = FakeImdsPage(present=("login_link",))
        self.assertFalse(imds_agent_v2.logged_in_to_imds(settling))
        self.assertTrue(
            imds_agent_v2.session_logged_in_after_reconnect(settling, login_succeeded=True)
        )
        self.assertFalse(
            imds_agent_v2.session_logged_in_after_reconnect(settling, login_succeeded=False)
        )

    def test_own_mds_same_id_draft_version(self):
        received = "8827908 / 8"
        own = "8827908 / 0.01"
        self.assertTrue(imds_agent_v2.versions_indicate_own_draft(received, own))
        self.assertTrue(imds_agent_v2.own_mds_ready_for_recipients(received, own))
        self.assertFalse(
            imds_agent_v2.own_mds_ready_for_recipients(received, "8827908 / 8")
        )
        self.assertFalse(
            imds_agent_v2.own_mds_ready_for_recipients(received, "8827908")
        )
        self.assertTrue(
            imds_agent_v2.own_mds_ready_for_recipients("1467141604 / 1", "1467999999 / 0.01")
        )

    def test_accept_pass_uses_full_mds_label_for_forward(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        accept_fn = text.split("def accept_passed_mds", 1)[1].split(
            "\ndef reject_failed_mds", 1
        )[0]
        self.assertIn("received_label", accept_fn)
        self.assertIn("wait_for_forwarded_own_mds(page, received_label)", accept_fn)
        self.assertNotIn("wait_for_forwarded_own_mds(page, mds_id_num)", accept_fn)

    def test_page_text_indicates_public_login_after_session_timeout(self):
        colab_login_body = (
            "Login\nUser ID forgotten\nRequest new password\n"
            "Registration\nRegister your company\nMulti-factor Authentication"
        )
        self.assertTrue(imds_agent_v2.page_text_indicates_public_login(colab_login_body))
        check_panel = "Check results - 0 Error(s) / 0 Warning(s)"
        self.assertFalse(imds_agent_v2.page_text_indicates_public_login(check_panel))
        inbox = "Received MDSs\nnot yet browsed\nSearch"
        self.assertFalse(imds_agent_v2.page_text_indicates_public_login(inbox))

    def test_leftover_three_and_ten_still_twenty_rows(self):
        saved_three = os.environ.pop("IMDS_ALLOW_THREE", None)
        saved_ten = os.environ.pop("IMDS_ALLOW_TEN", None)
        try:
            self.assertEqual(imds_agent_v2.resolve_num_iterations("3"), 20)
            self.assertEqual(imds_agent_v2.resolve_num_iterations("10"), 20)
        finally:
            if saved_three is None:
                os.environ.pop("IMDS_ALLOW_THREE", None)
            else:
                os.environ["IMDS_ALLOW_THREE"] = saved_three
            if saved_ten is None:
                os.environ.pop("IMDS_ALLOW_TEN", None)
            else:
                os.environ["IMDS_ALLOW_TEN"] = saved_ten


GADSDL_DIALOG_TEXT = (
    "GADSDL / SVHC Update\n"
    "GADSL / SVHC has been updated on 19/11/2026. All affected confidential substances "
    "have already been revealed.\n"
    "I understand that a new version of all materials using jokers to hide one of the "
    "updated substances has to be created and released to the supply chain according to "
    "Rec001 Rule 3.2.1.D.\n"
    "OK\nCancel\nExport"
)


class _FakeGadsdlLocator:
    def __init__(self, page: "FakeGadsdlDialogPage", kind: str):
        self.page = page
        self.kind = kind
        self.first = self

    def count(self) -> int:
        if self.kind == "lookup":
            return 0
        if self.kind in {"inbox", "received_menu", "received_link", "mds_menu", "org"}:
            return 1 if self.kind in self.page.chrome else 0
        if self.kind in {"username", "forgotten", "login_link", "login_button", "user_id_label"}:
            return 1
        if self.kind in {
            "dialog",
            "gadsdl",
            "checkbox",
            "ok",
            "cancel",
            "export",
            "glass",
        }:
            return 1 if self.page.modal_up else 0
        if self.kind == "body":
            return 1
        return 0

    def is_visible(self) -> bool:
        return self.count() > 0

    def is_checked(self) -> bool:
        if self.kind != "checkbox":
            raise AttributeError("not a checkbox")
        return bool(self.page.checkbox_checked)

    def inner_text(self, timeout: int = 500) -> str:
        if self.kind == "body":
            public = "Login\nLanguage\nUser ID forgotten\n"
            return public + (GADSDL_DIALOG_TEXT if self.page.modal_up else "")
        if self.kind in {"dialog", "gadsdl", "glass"} and self.page.modal_up:
            return GADSDL_DIALOG_TEXT
        return ""

    def click(self, force: bool = False, timeout: int = 4000):
        self.page.clicks.append(self.kind)
        if self.kind == "checkbox":
            self.page.checkbox_checked = True
            return
        if self.kind == "ok":
            if not self.page.checkbox_checked:
                raise RuntimeError("OK is disabled until the acknowledgment checkbox is checked")
            self.page.modal_up = False
            return
        if self.kind == "cancel":
            self.page.modal_up = False
            return
        if self.kind == "export":
            return

    def locator(self, _selector: str):
        return _FakeGadsdlLocator(self.page, "none")


class FakeGadsdlDialogPage:
    """Playwright stand-in: public Login chrome behind a GADSDL / SVHC Update dialog."""

    def __init__(self, *, modal_up: bool = True, checkbox_checked: bool = False, chrome=()):
        self.modal_up = modal_up
        self.checkbox_checked = checkbox_checked
        self.chrome = set(chrome)
        self.clicks: list[str] = []
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms: int):
        return None

    def wait_for_selector(self, _selector: str, state: str | None = None, timeout: int = 8000):
        if state == "detached" and self.modal_up:
            raise TimeoutError("GADSDL dialog still attached")
        return None

    def evaluate(self, _script):
        return None

    class _Keyboard:
        def press(self, _key: str):
            return None

    @property
    def keyboard(self):
        return FakeGadsdlDialogPage._Keyboard()

    def locator(self, selector: str):
        return _FakeGadsdlLocator(self, self._kind(selector or ""))

    @staticmethod
    def _kind(sel: str) -> str:
        s = sel.lower()
        if "lookupcompany" in s:
            return "lookup"
        if "pt_ctbtoolbarinbound" in s and "menu" not in s:
            return "inbox"
        if "pt_cmisearchinboxb" in s:
            return "inbox"
        if "sdiinboxsearch" in s:
            return "received_menu"
        if "received mdss" in s:
            return "received_link"
        if "pt_mfile" in s:
            return "mds_menu"
        if "johnson electric" in s:
            return "org"
        if "user id forgotten" in s:
            return "forgotten"
        if "username" in s or "userid" in s.replace(" ", ""):
            return "username"
        if "user id" in s:
            return "user_id_label"
        if "button:has-text('login')" in s or ("input[value" in s and "login" in s):
            return "login_button"
        if "a:has-text('login')" in s:
            return "login_link"
        if "ctbcancel" in s or "has-text('cancel')" in s or "value='cancel'" in s:
            return "cancel"
        if "export" in s:
            return "export"
        if "ctbyes" in s or "has-text('yes')" in s:
            return "none"
        if "ctbno" in s or "has-text('no')" in s:
            return "none"
        if "checkbox" in s or "i understand" in s:
            return "checkbox"
        if "gadsdl" in s or "gadsl" in s or "svhc" in s or "rec001" in s:
            return "gadsdl"
        if "ctbok" in s or "has-text('ok')" in s or "value='ok'" in s:
            return "ok"
        if "contains(normalize-space(.), 'ok')" in s and "cancel" not in s:
            return "ok"
        if "afmodalglasspane" in s or "afblockingglasspane" in s:
            return "glass"
        if "afmodaldialog" in s or "pt_dcud" in s:
            return "dialog"
        if s == "body":
            return "body"
        return "none"


class GadsdlSvhcModalTests(unittest.TestCase):
    def test_dismiss_checks_box_then_ok_not_cancel(self):
        page = FakeGadsdlDialogPage()
        self.assertTrue(imds_agent_v2.gadsdl_svhc_update_visible(page))
        self.assertTrue(imds_agent_v2.modal_dialog_visible(page))
        self.assertTrue(imds_agent_v2.dismiss_modal(page, allow_escape=False))
        self.assertIn("checkbox", page.clicks)
        self.assertIn("ok", page.clicks)
        self.assertNotIn("cancel", page.clicks)
        self.assertNotIn("export", page.clicks)
        self.assertLess(page.clicks.index("checkbox"), page.clicks.index("ok"))
        self.assertFalse(page.modal_up)
        self.assertFalse(imds_agent_v2.gadsdl_svhc_update_visible(page))

    def test_already_checked_only_clicks_ok(self):
        page = FakeGadsdlDialogPage(checkbox_checked=True)
        self.assertTrue(imds_agent_v2.acknowledge_gadsdl_svhc_update(page))
        self.assertEqual(page.clicks, ["ok"])
        self.assertNotIn("cancel", page.clicks)

    def test_login_page_detection_with_gadsdl_overlay(self):
        overlay = FakeGadsdlDialogPage()
        self.assertFalse(imds_agent_v2.on_public_login_page(overlay))
        self.assertFalse(imds_agent_v2.logged_in_to_imds(overlay))
        self.assertTrue(imds_agent_v2.dismiss_modal(overlay, allow_escape=False))
        self.assertTrue(imds_agent_v2.on_public_login_page(overlay))
        self.assertFalse(imds_agent_v2.logged_in_to_imds(overlay))

        public = FakeGadsdlDialogPage(modal_up=False)
        self.assertTrue(imds_agent_v2.on_public_login_page(public))

        leftover_login = FakeImdsPage(present=("inbox", "login_link", "username", "forgotten"))
        self.assertFalse(imds_agent_v2.on_public_login_page(leftover_login))


class SearchableMdsIdTests(unittest.TestCase):
    def test_extraction_failed_is_not_searchable(self):
        self.assertFalse(imds_agent_v2.is_searchable_mds_id("EXTRACTION_FAILED"))
        self.assertFalse(imds_agent_v2.is_searchable_mds_id(""))
        self.assertFalse(imds_agent_v2.is_searchable_mds_id(None))
        self.assertFalse(imds_agent_v2.is_searchable_mds_id("Open Failed"))
        self.assertTrue(imds_agent_v2.is_searchable_mds_id("1527311944 / 2"))
        self.assertTrue(imds_agent_v2.is_searchable_mds_id("1527311944"))

    def test_search_mds_by_id_skips_invalid_without_page(self):
        self.assertFalse(imds_agent_v2.search_mds_by_id(None, "EXTRACTION_FAILED"))
        self.assertFalse(imds_agent_v2.search_mds_by_id(None, ""))
        self.assertFalse(imds_agent_v2.search_mds_by_id(None, "Open Failed"))

    def test_accept_reject_skip_invalid_ids(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        accept_fn = text.split("def accept_passed_mds", 1)[1].split("\ndef reject_failed_mds", 1)[0]
        reject_fn = text.split("def reject_failed_mds", 1)[1].split("\nINBOX_SEARCH_TAB", 1)[0]
        self.assertIn("is_searchable_mds_id", accept_fn)
        self.assertIn("is_searchable_mds_id", reject_fn)
        self.assertIn("Skipped (invalid MDS ID)", accept_fn)
        self.assertIn("Skipped (invalid MDS ID)", reject_fn)
        self.assertNotIn('split("/")[0]', accept_fn.replace(" ", ""))
        self.assertIn("Not searching invalid MDS ID", text)


class InboxRecoverTests(unittest.TestCase):
    def test_ensure_session_inbox_missing_does_not_wait_or_relogin(self):
        page = FakeImdsPage(present=(), url="https://www.mdsystem.com/imdsnt")
        self.assertFalse(imds_agent_v2.page_looks_offline(page))
        self.assertFalse(imds_agent_v2.on_public_login_page(page))
        with (
            mock.patch.object(
                imds_agent_v2, "recover_inbox_list", return_value=True
            ) as recover,
            mock.patch.object(
                imds_agent_v2, "wait_for_imds_chrome", return_value=True
            ),
            mock.patch.object(
                imds_agent_v2, "wait_for_connectivity"
            ) as wait,
            mock.patch.object(imds_agent_v2, "imds_login") as login,
        ):
            self.assertTrue(imds_agent_v2.ensure_imds_session(page))
            recover.assert_called()
            wait.assert_not_called()
            login.assert_not_called()

    def test_reconnect_and_lost_table_reapply_filter(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        ensure_fn = text.split("def ensure_imds_session", 1)[1].split(
            "\ndef recover_after_network_error", 1
        )[0]
        recover_fn = text.split("def recover_inbox_list", 1)[1].split(
            "\ndef return_to_inbox_results", 1
        )[0]
        process_fn = text.split("def process_rows_and_export", 1)[1].split(
            "\ndef orchestrate", 1
        )[0]
        self.assertIn("apply_not_yet_browsed_filter_and_search", ensure_fn)
        self.assertIn("recover_inbox_list", ensure_fn)
        self.assertIn("apply_not_yet_browsed_filter_and_search", recover_fn)
        self.assertIn("recover_inbox_list", process_fn)
        self.assertIn("continuing remaining visible rows", process_fn)

    def test_process_rows_does_not_relogin_on_lost_list(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        process_fn = text.split("def process_rows_and_export", 1)[1].split(
            "\ndef orchestrate", 1
        )[0]
        self.assertIn("not a network drop", process_fn)
        self.assertIn("without re-login", process_fn)
        self.assertIn("page_looks_offline", process_fn)
        self.assertIn("needs_imds_relogin", process_fn)


class _IngredientsLocator(_FakeLocator):
    def __init__(self, page: "FakeIngredientsPage", n: int, visible: bool, kind: str):
        super().__init__(n, visible)
        self.page = page
        self.kind = kind

    def click(self, force: bool = False, timeout: int = 5000):
        self.page.clicks.append(self.kind)
        if self.kind == "ingredients":
            self.page.ingredients = True

    def hover(self, force: bool = False, timeout: int = 3000):
        return None


class FakeIngredientsPage:
    """Stand-in: MDS sheet that is not on Ingredients until the tab is clicked."""

    def __init__(self):
        self.ingredients = False
        self.clicks: list[str] = []
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms: int):
        return None

    def locator(self, selector: str):
        sel = (selector or "").lower()
        ready = self.ingredients
        if "sdidetailingredients" in sel or "sdiingr" in sel or "sdiingredients" in sel:
            return _IngredientsLocator(self, 1, True, "ingredients")
        if "ctbexpandall" in sel or "mds supplier" in sel or "treeitem" in sel:
            return _IngredientsLocator(self, 1 if ready else 0, ready, "tree")
        if "pt_mfile" in sel or "has-text('mds')" in sel:
            return _IngredientsLocator(self, 1 if ready else 0, ready, "mds_menu")
        if "lookupcompany" in sel:
            return _IngredientsLocator(self, 0, False, "none")
        if "pt_dcud" in sel or "afmodal" in sel:
            return _IngredientsLocator(self, 0, False, "none")
        return _IngredientsLocator(self, 0, False, "none")

    def get_by_text(self, text: str, exact: bool = False):
        if text == "Ingredients":
            return _IngredientsLocator(self, 1, True, "ingredients")
        return _IngredientsLocator(self, 0, False, "none")


class IngredientsReadyTests(unittest.TestCase):
    def test_ingredients_tab_retry_before_no_materials(self):
        page = FakeIngredientsPage()
        self.assertFalse(imds_agent_v2.ingredients_tree_ready(page))
        self.assertFalse(imds_agent_v2.mds_menu_available(page))
        with mock.patch.object(imds_agent_v2, "expand_tree"):
            with mock.patch.object(imds_agent_v2, "dismiss_modal"):
                self.assertTrue(imds_agent_v2.ensure_ingredients_ready_for_check(page))
        self.assertIn("ingredients", page.clicks)
        self.assertTrue(imds_agent_v2.ingredients_tree_ready(page))
        self.assertTrue(imds_agent_v2.mds_menu_available(page))

    def test_process_and_check_call_ingredients_ready(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        process_fn = text.split("def process_rows_and_export", 1)[1].split(
            "\ndef orchestrate", 1
        )[0]
        check_fn = text.split("def run_check", 1)[1].split("\ndef wait_for_check_results", 1)[0]
        collect_fn = text.split("def collect_material_nodes", 1)[1].split(
            "\ndef capture_all_material_nodes", 1
        )[0]
        self.assertIn("ensure_ingredients_ready_for_check", process_fn)
        self.assertIn("ensure_ingredients_ready_for_check", check_fn)
        self.assertIn("ensure_ingredients_ready_for_check", collect_fn)
        self.assertIn("clicking Ingredients tab before Check", text)


class OrdinaryOkDialogTests(unittest.TestCase):
    def test_ordinary_notice_clicks_ok_once(self):
        class _OkLoc:
            def __init__(self, page, kind):
                self.page = page
                self.kind = kind
                self.first = self

            def count(self):
                if self.kind in {"ok", "ok_span", "dialog"}:
                    return 1 if self.page.modal_up else 0
                return 0

            def is_visible(self):
                return self.count() > 0

            def inner_text(self, timeout=500):
                return "Information\nSearch complete.\nOK" if self.page.modal_up else ""

            def click(self, force=False, timeout=4000):
                self.page.clicks.append(self.kind)
                self.page.modal_up = False

            def locator(self, _sel):
                return _OkLoc(self.page, "none")

        class _OkPage:
            def __init__(self):
                self.modal_up = True
                self.clicks = []
                self.url = "https://www.mdsystem.com/imdsnt"

            def wait_for_timeout(self, _ms):
                return None

            def locator(self, selector: str):
                s = (selector or "").lower()
                if "lookupcompany" in s:
                    return _OkLoc(self, "none")
                if "ctbyes" in s or "has-text('yes')" in s:
                    return _OkLoc(self, "none")
                if "ctbno" in s or "has-text('no')" in s:
                    return _OkLoc(self, "none")
                if "ctbcancel" in s or "has-text('cancel')" in s:
                    return _OkLoc(self, "none")
                if "ctbok > a > span" in s or "ctbok > a > span" in s.replace("\\", ""):
                    return _OkLoc(self, "ok_span")
                if "ctbok" in s or "has-text('ok')" in s or "value='ok'" in s:
                    return _OkLoc(self, "ok")
                if "pt_dcud" in s or "afmodal" in s:
                    return _OkLoc(self, "dialog")
                if s == "body":
                    return _OkLoc(self, "dialog")
                return _OkLoc(self, "none")

        page = _OkPage()
        self.assertTrue(imds_agent_v2.dismiss_modal(page, allow_escape=False))
        self.assertEqual(page.clicks.count("ok") + page.clicks.count("ok_span"), 1)
        self.assertFalse(page.modal_up)


INBOX_CHROME = (
    'Menu Export There are hidden column(s). You can show them using "View" button.'
)


class _KindLoc:
    def __init__(self, page, kind, n=0, visible=False, text="", element_id=""):
        self.page = page
        self.kind = kind
        self._n = n
        self._visible = visible
        self._text = text
        self.element_id = element_id
        self.first = self

    def count(self):
        return self._n

    def is_visible(self):
        return bool(self._visible and self._n > 0)

    def click(self, force=False, timeout=5000):
        self.page.clicks.append(self.kind or self.element_id or "unknown")
        if self.kind == "accept_menu":
            self.page.accept_menu_up = False
            if hasattr(self.page, "accept_clicked_at"):
                self.page.accept_clicked_at = getattr(self.page, "waited_ms", 0)
            if hasattr(self.page, "dialog_up"):
                self.page.dialog_up = True
        if self.kind == "confirm":
            self.page.confirmed = True
            if hasattr(self.page, "dialog_up"):
                self.page.dialog_up = False
        if self.kind == "cancel":
            self.page.cancelled = True
            if hasattr(self.page, "dialog_up"):
                self.page.dialog_up = False
        if self.kind == "dialog_forward":
            self.page.clicked_user_dialog = True
        if self.kind == "inbox_row":
            self.page.dblclicks += 1

    def dblclick(self, force=False, timeout=5000):
        self.page.dblclicks += 1

    def hover(self, force=False, timeout=3000):
        return None

    def text_content(self, timeout=1500):
        return self._text

    def inner_text(self, timeout=500):
        return self._text

    def all(self):
        return [self] if self._n else []

    def locator(self, selector):
        if "following-sibling" in (selector or ""):
            return _KindLoc(
                self.page, "id_neighbor", n=1, visible=True, text=self.page.id_neighbor
            )
        return _KindLoc(self.page, "none")

    def nth(self, _i):
        return self


class FakeAcceptNoConfirmPage:
    """MDS menu + pt_cmiMenuAccept; confirm control never appears (index 17 false success)."""

    def __init__(self):
        self.clicks = []
        self.accept_menu_up = True
        self.confirmed = False
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def wait_for_load_state(self, *_a, **_k):
        return None

    def wait_for_selector(self, selector, **_k):
        raise TimeoutError(selector)

    def locator(self, selector: str):
        s = (selector or "").lower().replace("\\", "")
        if "lookupcompany" in s:
            return _KindLoc(self, "none")
        if "pt_mfile" in s or "has-text('mds')" in s:
            return _KindLoc(self, "mds_menu", n=1, visible=True)
        if "pt_cmimenuaccept" in s or "id='pt1:pt_cmimenuaccept'" in s:
            n = 1 if self.accept_menu_up else 0
            return _KindLoc(self, "accept_menu", n=n, visible=bool(n))
        if "ctbacceptmds" in s or "ctbaccept" in s or "ctbok" in s:
            return _KindLoc(self, "confirm", n=0)
        if "pt_dcud" in s or "afmodal" in s:
            return _KindLoc(self, "dialog", n=0)
        if s.startswith("td:has-text('accept')") or s.startswith(
            "[role='menuitem']:has-text('accept')"
        ):
            raise AssertionError(f"broad Accept locator must not be used: {selector}")
        return _KindLoc(self, "none")


class FakeForwardDialogPage:
    """Forward action missing; leftover user-dialog td contains the word Forward."""

    def __init__(self):
        self.clicks = []
        self.clicked_user_dialog = False
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def wait_for_load_state(self, *_a, **_k):
        return None

    def locator(self, selector: str):
        s = (selector or "").lower().replace("\\", "")
        if "lookupcompany" in s:
            return _KindLoc(self, "none")
        if "pt_mfile" in s or "has-text('mds')" in s:
            return _KindLoc(self, "mds_menu", n=1, visible=True)
        if "pt_mmenuforward" in s:
            return _KindLoc(self, "fwd_main", n=1, visible=True)
        if "pt_cmimenuforward" in s or "pt_ctbforward" in s:
            return _KindLoc(self, "fwd_action", n=0)
        if "has-text('forward')" in s or "pt_dlguserdialog" in s:
            return _KindLoc(
                self,
                "dialog_forward",
                n=1,
                visible=False,
                element_id="pt1:pt_dcud:pt_dlgUserDialog::contentContainer",
            )
        if "pt_dcud" in s or "afmodal" in s:
            return _KindLoc(self, "dialog", n=0)
        if "ctbok" in s:
            return _KindLoc(self, "ok", n=0)
        return _KindLoc(self, "none")


class FakeInboxChromeExtractPage:
    def __init__(self, body=INBOX_CHROME, neighbor=INBOX_CHROME):
        self.body = body
        self.id_neighbor = neighbor
        self.clicks = []
        self.dblclicks = 0
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def wait_for_selector(self, selector, timeout=10000):
        return True

    def locator(self, selector: str):
        s = (selector or "")
        sl = s.lower()
        if sl == "td" or sl == "td:has-text('id / version')":
            cell = _KindLoc(self, "id_label", n=1, visible=True, text="ID / Version")
            return cell
        if "tresult" in sl:
            return _KindLoc(self, "tresult", n=1, visible=True)
        if "ctbexpandall" in sl or "mds supplier" in sl or "treeitem" in sl:
            return _KindLoc(self, "tree", n=0)
        if sl == "body":
            return _KindLoc(self, "body", n=1, visible=True, text=self.body)
        return _KindLoc(self, "none")


class FakeOwnMdsWaitPage:
    def __init__(self, on_screen: str):
        self.on_screen = on_screen
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def locator(self, selector: str):
        sl = (selector or "").lower()
        if sl == "body":
            return _KindLoc(self, "body", n=1, visible=True, text=self.on_screen)
        if "id / version" in sl:
            return _KindLoc(self, "id_label", n=1, visible=True, text=f"ID / Version {self.on_screen}")
        if "lookupcompany" in sl or "pt_dcud" in sl or "afmodal" in sl:
            return _KindLoc(self, "none")
        return _KindLoc(self, "none")


class AcceptForwardInboxTests(unittest.TestCase):
    def test_inbox_table_chrome_is_not_mds_id(self):
        self.assertTrue(imds_agent_v2.looks_like_inbox_table_chrome(INBOX_CHROME))
        self.assertFalse(imds_agent_v2.looks_like_mds_id_value(INBOX_CHROME))
        self.assertEqual(imds_agent_v2.parse_mds_id_number(INBOX_CHROME), "")
        self.assertFalse(imds_agent_v2.is_searchable_mds_id(INBOX_CHROME))
        self.assertTrue(imds_agent_v2.looks_like_mds_id_value("1470791560 / 1"))
        self.assertFalse(imds_agent_v2.looks_like_mds_id_value("EXTRACTION_FAILED"))

    def test_extract_rejects_inbox_table_chrome(self):
        page = FakeInboxChromeExtractPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            self.assertEqual(
                imds_agent_v2.extract_mds_id_version_early(page), "EXTRACTION_FAILED"
            )

    def test_own_mds_gate_blocks_add_recipient_on_received_id(self):
        self.assertFalse(
            imds_agent_v2.own_mds_ready_for_recipients("1470791560", "1470791560 / 1")
        )
        self.assertTrue(
            imds_agent_v2.own_mds_ready_for_recipients("1470791560", "1616000123 / 0.01")
        )
        page = FakeOwnMdsWaitPage("1470791560 / 1")
        with (
            mock.patch.object(imds_agent_v2, "dismiss_modal"),
            mock.patch.object(imds_agent_v2, "read_visible_mds_id", return_value="1470791560 / 1"),
            mock.patch.object(
                imds_agent_v2, "extract_mds_id_version_early", return_value="1470791560 / 1"
            ),
        ):
            self.assertEqual(
                imds_agent_v2.wait_for_forwarded_own_mds(page, "1470791560", timeout_s=0.2),
                "",
            )

    def test_accept_passed_skips_recipients_when_forward_did_not_mint(self):
        results = [
            {
                "Overall Result": "PASS",
                "MDS ID / Version": "1470791560 / 1",
                "Action Result": "Pending action",
                "Supplier Code": "",
                "Part/Item No.": "1616-1YY0331-MT(A)",
            }
        ]
        page = FakeImdsPage(present=("mds_menu", "inbox"))
        with (
            mock.patch.object(imds_agent_v2, "page_looks_offline", return_value=False),
            mock.patch.object(imds_agent_v2, "on_public_login_page", return_value=False),
            mock.patch.object(imds_agent_v2, "search_mds_by_id", return_value=True),
            mock.patch.object(
                imds_agent_v2, "open_first_result_on_content_page", return_value=True
            ),
            mock.patch.object(imds_agent_v2, "accept_mds", return_value=True),
            mock.patch.object(imds_agent_v2, "handle_forward_confirmation_modal"),
            mock.patch.object(imds_agent_v2, "wait_for_glass_pane_clear"),
            mock.patch.object(imds_agent_v2, "_recover_forward_ready_after_accept"),
            mock.patch.object(
                imds_agent_v2, "read_visible_mds_id", return_value="1470791560 / 1"
            ),
            mock.patch.object(
                imds_agent_v2, "extract_mds_id_version_early", return_value="1470791560 / 1"
            ),
            mock.patch.object(imds_agent_v2, "wait_for_mds_content_page", return_value=True),
            mock.patch.object(imds_agent_v2, "forward_mds", return_value=True) as fwd,
            mock.patch.object(
                imds_agent_v2, "wait_for_forwarded_own_mds", return_value=""
            ),
            mock.patch.object(imds_agent_v2, "complete_forward_recipients") as recip,
            mock.patch.object(imds_agent_v2, "leave_own_mds_for_inbox", return_value=True),
            mock.patch.object(imds_agent_v2, "navigate_to_search_page", return_value=True),
            mock.patch.object(imds_agent_v2, "save_check_summary"),
            mock.patch.object(imds_agent_v2, "close_check_results_dialog"),
            mock.patch.object(imds_agent_v2, "_dismiss_leftover_user_dialogs_before_mds_menu"),
        ):
            imds_agent_v2.accept_passed_mds(page, results)
        recip.assert_not_called()
        self.assertEqual(results[0]["Action Result"], "Forward Failed")
        self.assertGreaterEqual(fwd.call_count, 2)

    def test_accept_uses_exact_id_and_missing_confirm_is_not_success(self):
        page = FakeAcceptNoConfirmPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            with mock.patch.object(
                imds_agent_v2, "_dismiss_leftover_user_dialogs_before_mds_menu"
            ):
                self.assertFalse(imds_agent_v2.accept_mds(page))
        self.assertFalse(page.confirmed)
        self.assertIn("accept_menu", page.clicks)
        self.assertNotIn("disappeared", " ".join(page.clicks))

    def test_forward_does_not_click_user_dialog(self):
        page = FakeForwardDialogPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            with mock.patch.object(
                imds_agent_v2, "_dismiss_leftover_user_dialogs_before_mds_menu"
            ):
                self.assertFalse(imds_agent_v2.forward_mds(page))
        self.assertFalse(page.clicked_user_dialog)
        self.assertNotIn("dialog_forward", page.clicks)


class FakeAcceptSlowConfirmPage:
    """Confirm control appears only after a longer wait (draft 0.01 leftover dcPopup)."""

    def __init__(self):
        self.clicks = []
        self.accept_menu_up = True
        self.confirmed = False
        self.waited_ms = 0
        self.accept_clicked_at = None
        self.url = "https://www.mdsystem.com/imdsnt"

    def _confirm_ready(self):
        if self.accept_clicked_at is None:
            return False
        return (self.waited_ms - self.accept_clicked_at) >= 4000

    def wait_for_timeout(self, ms):
        self.waited_ms += int(ms or 0)

    def wait_for_load_state(self, *_a, **_k):
        return None

    def wait_for_selector(self, selector, **_k):
        s = (selector or "").lower()
        if "mds accepted" in s and self.confirmed:
            return True
        raise TimeoutError(selector)

    def locator(self, selector: str):
        s = (selector or "").lower().replace("\\", "")
        if "lookupcompany" in s:
            return _KindLoc(self, "none")
        if "pt_mfile" in s or "has-text('mds')" in s:
            return _KindLoc(self, "mds_menu", n=1, visible=True)
        if "pt_cmimenuaccept" in s or "id='pt1:pt_cmimenuaccept'" in s:
            n = 1 if self.accept_menu_up else 0
            return _KindLoc(self, "accept_menu", n=n, visible=bool(n))
        if "ctbacceptmds" in s or "ctbaccept" in s or (
            "dcpopup" in s and "accept" in s
        ):
            n = 1 if self._confirm_ready() and not self.confirmed else 0
            return _KindLoc(self, "confirm", n=n, visible=bool(n))
        if "pt_dcud" in s or "afmodal" in s or s in {"#dcpopup", "[id='dcpopup']"}:
            return _KindLoc(self, "dialog", n=0)
        if s.startswith("td:has-text('accept')") or s.startswith(
            "[role='menuitem']:has-text('accept')"
        ):
            raise AssertionError(f"broad Accept locator must not be used: {selector}")
        return _KindLoc(self, "none")


class FakeAcceptLeftoverDialogPage:
    """Accept menu opens a leftover dcPopup; confirm never appears; Cancel restores chrome."""

    def __init__(self):
        self.clicks = []
        self.accept_menu_up = True
        self.dialog_up = False
        self.confirmed = False
        self.cancelled = False
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def wait_for_load_state(self, *_a, **_k):
        return None

    def wait_for_selector(self, selector, **_k):
        raise TimeoutError(selector)

    def locator(self, selector: str):
        s = (selector or "").lower().replace("\\", "")
        if "lookupcompany" in s:
            return _KindLoc(self, "none")
        if "ctbok" in s:
            return _KindLoc(self, "ok", n=0)
        if "cmi" not in s and "cancel" not in s and (
            "ctbaccept" in s
            or "has-text('accept" in s
            or "value='accept'" in s
            or "button:has-text('accept" in s
        ):
            return _KindLoc(self, "confirm", n=0)
        if "ctbcancel" in s or "has-text('cancel')" in s or "value='cancel'" in s:
            n = 1 if self.dialog_up else 0
            return _KindLoc(self, "cancel", n=n, visible=bool(n))
        if "pt_cmimenuaccept" in s or "id='pt1:pt_cmimenuaccept'" in s:
            n = 1 if self.accept_menu_up and not self.dialog_up else 0
            return _KindLoc(self, "accept_menu", n=n, visible=bool(n))
        if "pt_mfile" in s or "has-text('mds')" in s:
            return _KindLoc(self, "mds_menu", n=1, visible=True)
        if (
            "afmodal" in s
            or "glass" in s
            or "pt_dcud" in s
            or "dcpopup" in s
        ):
            n = 1 if self.dialog_up else 0
            return _KindLoc(self, "dialog", n=n, visible=bool(n), text="Accept MDS")
        if s.startswith("td:has-text('accept')") or s.startswith(
            "[role='menuitem']:has-text('accept')"
        ):
            raise AssertionError(f"broad Accept locator must not be used: {selector}")
        return _KindLoc(self, "none")


class _GoneLoc:
    def __init__(self, page):
        self.page = page
        self.first = self

    def count(self):
        return 0

    def is_visible(self):
        return False

    def click(self, **_k):
        raise TimeoutError("missing")

    def locator(self, _sel):
        return self

    def inner_text(self, **_k):
        return ""

    def text_content(self, **_k):
        return ""


class FakeChromeGonePage:
    """Inbox chrome is gone; leftover Login DOM must not trigger re-login."""

    def __init__(self):
        self.gotos = []
        self.clicks = []
        self.url = "https://www.mdsystem.com/imdsnt"

    def wait_for_timeout(self, _ms):
        return None

    def wait_for_load_state(self, *_a, **_k):
        raise TimeoutError("networkidle")

    def wait_for_selector(self, selector, **_k):
        raise TimeoutError(selector)

    def goto(self, url, **_k):
        self.gotos.append(url)

    def locator(self, _selector: str):
        return _GoneLoc(self)


class AcceptChromeRecoveryTests(unittest.TestCase):
    def test_confirm_wait_is_longer_than_two_seconds(self):
        self.assertGreaterEqual(imds_agent_v2.ACCEPT_CONFIRM_WAIT_MS, 8000)
        self.assertIn("#dcPopup\\:ctbAcceptMds", imds_agent_v2.CSS_ACCEPT_CONFIRM)
        self.assertIn("//*[@id='dcPopup:ctbAcceptMds']/a/span", imds_agent_v2.XP_ACCEPT_CONFIRM)

    def test_slow_accept_confirm_is_waited_for(self):
        page = FakeAcceptSlowConfirmPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            self.assertTrue(imds_agent_v2.accept_mds(page))
        self.assertTrue(page.confirmed)
        self.assertGreaterEqual(page.waited_ms - (page.accept_clicked_at or 0), 4000)

    def test_no_mds_menu_fallback_while_dialog_up(self):
        page = FakeAcceptLeftoverDialogPage()
        page.dialog_up = True
        self.assertFalse(imds_agent_v2._open_mds_file_menu(page))
        self.assertNotIn("mds_menu", page.clicks)

    def test_leftover_accept_dialog_is_cancelled_not_mds_menu_retried(self):
        page = FakeAcceptLeftoverDialogPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            self.assertFalse(imds_agent_v2.accept_mds(page))
        self.assertFalse(page.confirmed)
        self.assertTrue(page.cancelled)
        self.assertIn("cancel", page.clicks)
        self.assertEqual(page.clicks.count("mds_menu"), 1)

    def test_missing_confirm_is_not_success(self):
        page = FakeAcceptNoConfirmPage()
        with mock.patch.object(imds_agent_v2, "save_screenshot"):
            self.assertFalse(imds_agent_v2.accept_mds(page))
        self.assertFalse(page.confirmed)

    def test_leave_own_mds_uses_save_changes_no(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        leave_fn = text.split("def leave_own_mds_for_inbox", 1)[1].split(
            "\ndef navigate_to_search_page", 1
        )[0]
        self.assertIn('save_changes="no"', leave_fn)
        self.assertIn("Leaving own MDS with save-changes No", leave_fn)
        pass_fn = text.split("def accept_passed_mds", 1)[1].split(
            "\ndef reject_failed_mds", 1
        )[0]
        self.assertIn("leave_own_mds_for_inbox", pass_fn)
        reject_fn = text.split("def reject_failed_mds", 1)[1].split(
            "\nINBOX_SEARCH_TAB", 1
        )[0]
        self.assertIn("leave_own_mds_for_inbox", reject_fn)

    def test_dismiss_modal_before_inbox_wait(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        nav_fn = text.split("def navigate_to_search_page", 1)[1].split(
            "\ndef apply_not_yet_browsed_filter_and_search", 1
        )[0]
        self.assertLess(
            nav_fn.find("_dismiss_nav_blockers"),
            nav_fn.find("_try_search_nav_fast"),
        )
        blockers_fn = text.split("def _dismiss_nav_blockers", 1)[1].split(
            "\ndef _try_search_nav_fast", 1
        )[0]
        self.assertIn("dismiss_modal", blockers_fn)
        self.assertIn("BEFORE waiting for Inbox chrome", blockers_fn)

    def test_search_nav_recovers_chrome_once_not_four_times_per_id(self):
        imds_agent_v2.reset_search_nav_chrome_recovery()
        page = FakeChromeGonePage()
        order = []

        def rec_dismiss(*_a, **_k):
            order.append("dismiss")
            return True

        def rec_chrome(*_a, **_k):
            order.append("chrome")
            return False

        with (
            mock.patch.object(imds_agent_v2, "dismiss_modal", side_effect=rec_dismiss),
            mock.patch.object(imds_agent_v2, "close_check_results_dialog"),
            mock.patch.object(imds_agent_v2, "close_company_lookup_dialogs"),
            mock.patch.object(imds_agent_v2, "wait_for_glass_pane_clear", return_value=True),
            mock.patch.object(imds_agent_v2, "wait_for_imds_chrome", side_effect=rec_chrome),
            mock.patch.object(imds_agent_v2, "imds_login") as login,
            mock.patch.object(imds_agent_v2, "on_public_login_page", return_value=False),
            mock.patch.object(imds_agent_v2, "logged_in_to_imds", return_value=True),
            mock.patch.object(imds_agent_v2, "page_looks_offline", return_value=False),
        ):
            self.assertFalse(imds_agent_v2.navigate_to_search_page(page))
            first_chrome = order.count("chrome")
            first_goto = len(page.gotos)
            self.assertGreaterEqual(order.count("dismiss"), 1)
            self.assertLess(order.index("dismiss"), order.index("chrome"))
            self.assertFalse(imds_agent_v2.navigate_to_search_page(page))
            self.assertLessEqual(order.count("chrome") - first_chrome, 1)
            self.assertEqual(len(page.gotos), first_goto)
            login.assert_not_called()
            self.assertGreaterEqual(first_goto, 1)
            self.assertLessEqual(first_goto, 1)


PASSING_CHECK_DIALOG = (
    "Check results\n"
    "The MDS has passed all included checks. These checks do not cover all aspects "
    "of IMDS data requirements. Further manual review may be required.\n"
    "Accept\nReject\nCancel"
)


class _CheckAcceptLocator:
    def __init__(self, page: "FakePassingCheckDialog", kind: str):
        self.page = page
        self.kind = kind
        self.first = self

    def count(self) -> int:
        if not self.page.modal_up:
            return 0
        if self.kind in {"accept", "dialog"}:
            return 1
        return 0

    def is_visible(self) -> bool:
        return self.count() > 0

    def click(self, force=False, timeout=0):
        if self.kind != "accept":
            raise RuntimeError(f"unexpected click {self.kind}")
        self.page.clicks.append("accept")
        self.page.modal_up = False

    def inner_text(self, timeout=0):
        if self.kind == "dialog" and self.page.modal_up:
            return self.page.text
        return ""


class FakePassingCheckDialog:
    def __init__(self, text: str = PASSING_CHECK_DIALOG, modal_up: bool = True):
        self.text = text
        self.modal_up = modal_up
        self.clicks: list[str] = []

    def wait_for_timeout(self, _ms: int):
        return None

    def locator(self, selector: str):
        sel = (selector or "").lower()
        if "lookupcompany" in sel:
            return _CheckAcceptLocator(self, "none")
        if "accept" in sel and ("afmodaldialog" in sel or "dcpopup" in sel or "ctbaccept" in sel or "pt_dcud" in sel):
            return _CheckAcceptLocator(self, "accept")
        if "afmodaldialog" in sel or "afmodalglasspane" in sel or "afblockingglasspane" in sel or "pt_dcud" in sel or "dcpopup" in sel:
            return _CheckAcceptLocator(self, "dialog")
        return _CheckAcceptLocator(self, "none")


class PassingCheckAcceptTests(unittest.TestCase):
    def test_clicks_visible_check_dialog_accept_not_cancel(self):
        page = FakePassingCheckDialog()
        self.assertTrue(imds_agent_v2.passing_check_dialog_visible(page))
        self.assertTrue(imds_agent_v2.click_passing_check_accept(page))
        self.assertEqual(page.clicks, ["accept"])
        self.assertFalse(page.modal_up)
        self.assertFalse(imds_agent_v2.passing_check_dialog_visible(page))
        self.assertFalse(imds_agent_v2.click_passing_check_accept(page))

    def test_accept_confirm_does_not_skip_visible_check_dialog(self):
        page = FakePassingCheckDialog()
        self.assertTrue(imds_agent_v2._click_accept_confirm_control(page))
        self.assertEqual(page.clicks, ["accept"])
        self.assertFalse(imds_agent_v2.passing_check_dialog_visible(page))

    def test_blocking_check_dialog_is_not_an_accept_confirm(self):
        page = FakePassingCheckDialog(
            text=(
                "Check results - 1 Error(s) / 0 Warning(s)\n"
                "Contact must be specified\n"
                "All existing errors need to be eliminated before any further processing may take place."
            )
        )
        self.assertFalse(imds_agent_v2.passing_check_dialog_visible(page))
        self.assertFalse(imds_agent_v2.click_passing_check_accept(page))
        self.assertEqual(page.clicks, [])
        self.assertTrue(page.modal_up)


if __name__ == "__main__":
    unittest.main()
