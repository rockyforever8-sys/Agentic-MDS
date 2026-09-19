# Paste this into a NEW Colab notebook (File → New notebook).
# Do not paste Agentic_MDS.ipynb — that file is JSON and causes:
#   NameError: name 'true' is not defined

import os, pathlib, subprocess

ROOT = pathlib.Path("/content/Agentic-MDS")
REPO = "https://github.com/rockyforever8-sys/Agentic-MDS.git"
BRANCH = "cursor/row5-inbox-recover-07ca"
PIN = "70ac29307a6c70ffefa9c0c4c75ca44b7fe7c791"
REF = os.environ.get("IMDS_GIT_REF", PIN)
if not (ROOT / ".git").exists():
    subprocess.check_call(["git", "clone", "--depth", "50", "--branch", BRANCH, REPO, str(ROOT)])
else:
    subprocess.check_call(["git", "-C", str(ROOT), "fetch", "--depth", "50", "origin", BRANCH])
os.chdir(ROOT)
subprocess.check_call(["git", "checkout", "--detach", REF])
print("Working directory:", os.getcwd())
print("git HEAD:", subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip())
print("Next: run Cell 1 in Colab_Start_Here.ipynb, or:")
print("  !python imds_agent_v2.py")
