#!/usr/bin/env python3
"""Colab/Jupyter has a running asyncio loop; Playwright Sync API cannot start there."""
from __future__ import annotations

import asyncio
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import imds_agent_v2


class ColabLoopTests(unittest.TestCase):
    def tearDown(self) -> None:
        os.environ.pop("IMDS_INSIDE_ORCHESTRATE_SUBPROCESS", None)

    def test_loop_not_running_in_sync_unittest(self):
        self.assertFalse(imds_agent_v2._asyncio_loop_is_running())

    def test_loop_running_inside_asyncio(self):
        async def _check() -> bool:
            return imds_agent_v2._asyncio_loop_is_running()

        self.assertTrue(asyncio.run(_check()))

    def test_orchestrate_uses_subprocess_inside_running_loop(self):
        with patch.object(imds_agent_v2, "_orchestrate_in_subprocess", return_value=17) as sub:
            with patch.object(imds_agent_v2, "_orchestrate_live", return_value=99):
                async def _go() -> int:
                    return imds_agent_v2.orchestrate()

                self.assertEqual(asyncio.run(_go()), 17)
                sub.assert_called_once()

    def test_orchestrate_live_when_child_flag_set_even_inside_loop(self):
        os.environ["IMDS_INSIDE_ORCHESTRATE_SUBPROCESS"] = "1"
        with patch.object(imds_agent_v2, "_orchestrate_live", return_value=3) as live:
            with patch.object(imds_agent_v2, "_orchestrate_in_subprocess", return_value=17) as sub:
                async def _go() -> int:
                    return imds_agent_v2.orchestrate()

                self.assertEqual(asyncio.run(_go()), 3)
                live.assert_called_once()
                sub.assert_not_called()

    def test_orchestrate_live_without_loop(self):
        with patch.object(imds_agent_v2, "_orchestrate_live", return_value=5) as live:
            with patch.object(imds_agent_v2, "_orchestrate_in_subprocess", return_value=17) as sub:
                self.assertEqual(imds_agent_v2.orchestrate(), 5)
                live.assert_called_once()
                sub.assert_not_called()

    def test_subprocess_command_is_unbuffered_script(self):
        fake = MagicMock()
        fake.stdout = iter(["hello from child\n"])
        fake.wait.return_value = 0
        with patch("imds_agent_v2.subprocess.Popen", return_value=fake) as popen:
            rc = imds_agent_v2._orchestrate_in_subprocess()
        self.assertEqual(rc, 0)
        cmd = popen.call_args.args[0]
        self.assertEqual(cmd[1], "-u")
        self.assertTrue(str(cmd[2]).endswith("imds_agent_v2.py"))
        env = popen.call_args.kwargs["env"]
        self.assertEqual(env["IMDS_INSIDE_ORCHESTRATE_SUBPROCESS"], "1")
        self.assertEqual(env["PYTHONUNBUFFERED"], "1")


    def test_ensure_skip_flag_does_not_apt(self):
        os.environ["IMDS_SKIP_BROWSER_DEPS"] = "1"
        try:
            with patch.object(imds_agent_v2, "libatk_present", return_value=False):
                with patch("imds_agent_v2.subprocess.run") as run:
                    imds_agent_v2.ensure_chromium_os_deps()
                    run.assert_not_called()
        finally:
            os.environ.pop("IMDS_SKIP_BROWSER_DEPS", None)

    def test_ensure_skips_when_libatk_present(self):
        os.environ.pop("IMDS_SKIP_BROWSER_DEPS", None)
        with patch.object(imds_agent_v2, "libatk_present", return_value=True):
            with patch("imds_agent_v2.subprocess.run") as run:
                imds_agent_v2.ensure_chromium_os_deps()
                run.assert_not_called()

    def test_looks_like_missing_libatk_message(self):
        self.assertTrue(
            imds_agent_v2._looks_like_missing_chromium_lib(
                RuntimeError("error while loading shared libraries: libatk-1.0.so.0")
            )
        )
        self.assertTrue(
            imds_agent_v2._looks_like_missing_chromium_lib(
                RuntimeError("BrowserType.launch: Target page, context or browser has been closed")
            )
        )
        self.assertFalse(imds_agent_v2._looks_like_missing_chromium_lib(RuntimeError("timeout")))

    def test_first_visible_skips_hidden_adf_clones(self):
        class _Item:
            def __init__(self, shown: bool) -> None:
                self._shown = shown

            def is_visible(self) -> bool:
                return self._shown

        class _Loc:
            def __init__(self, flags: list[bool]) -> None:
                self._flags = flags

            def count(self) -> int:
                return len(self._flags)

            def nth(self, i: int) -> _Item:
                return _Item(self._flags[i])

        self.assertIsNone(imds_agent_v2.first_visible(_Loc([False, False])))
        hit = imds_agent_v2.first_visible(_Loc([False, True, True]))
        self.assertIsNotNone(hit)
        self.assertTrue(hit.is_visible())

    def test_username_selectors_include_adf_user_id(self):
        joined = " ".join(imds_agent_v2.USERNAME_SELECTORS)
        self.assertIn("::content", joined)
        self.assertIn("UserId", joined)
        self.assertIn("#username", joined)


if __name__ == "__main__":
    unittest.main()
