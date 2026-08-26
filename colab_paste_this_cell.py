# Paste this into a NEW Colab notebook (File → New notebook).
# Do not paste Agentic_MDS.ipynb — that file is JSON and causes:
#   NameError: name 'true' is not defined

import os, pathlib, subprocess

ROOT = pathlib.Path("/content/Agentic-MDS")
REPO = "https://github.com/rockyforever8-sys/Agentic-MDS.git"
if not (ROOT / ".git").exists():
    subprocess.check_call(["git", "clone", "--depth", "1", REPO, str(ROOT)])
else:
    subprocess.check_call(["git", "-C", str(ROOT), "fetch", "--depth", "1", "origin", "main"])
    subprocess.check_call(["git", "-C", str(ROOT), "checkout", "-B", "main", "origin/main"])
os.chdir(ROOT)
print("Working directory:", os.getcwd())
print("Next: Runtime → Run all in Colab_Start_Here.ipynb, or run:")
print("  !python imds_agent_v2.py --self-test")
