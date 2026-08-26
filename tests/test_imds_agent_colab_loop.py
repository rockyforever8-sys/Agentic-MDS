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


if __name__ == "__main__":
    unittest.main()
