#!/usr/bin/env python3
"""Compatibility entry point for the Colab-exported filename.

The live IMDS agent is imds_agent_v2.py (original XPaths and actions).
Login secrets come from Colab Secrets / the environment — never from this file.
"""
from imds_agent_v2 import orchestrate

if __name__ == "__main__":
    raise SystemExit(orchestrate())
