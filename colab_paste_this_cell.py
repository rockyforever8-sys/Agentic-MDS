# Paste this into a NEW Colab notebook (File → New notebook).
# Do not paste Agentic_MDS.ipynb — that file is JSON and causes:
#   NameError: name 'true' is not defined

import os, pathlib, subprocess

ROOT = pathlib.Path("/content/Agentic-MDS")
REPO = "https://github.com/rockyforever8-sys/Agentic-MDS.git"
REF = os.environ.get("IMDS_GIT_REF", "cursor/colab-playwright-asyncio-07ca")
if not (ROOT / ".git").exists():
    try:
        subprocess.check_call(["git", "clone", "--depth", "1", "--branch", REF, REPO, str(ROOT)])
    except subprocess.CalledProcessError:
        subprocess.check_call(["git", "clone", "--depth", "1", REPO, str(ROOT)])
else:
    fetched = False
    for _ref in (REF, "main"):
        try:
            subprocess.check_call(["git", "-C", str(ROOT), "fetch", "--depth", "1", "origin", _ref])
            subprocess.check_call(["git", "-C", str(ROOT), "checkout", "-B", _ref, f"origin/{_ref}"])
            fetched = True
            break
        except subprocess.CalledProcessError:
            print("Could not fetch origin/" + _ref)
    if not fetched:
        raise RuntimeError("git fetch failed for " + REF + " and main")
os.chdir(ROOT)
print("Working directory:", os.getcwd())
print("Next: Runtime → Run all in Colab_Start_Here.ipynb, or run:")
print("  !python imds_agent_v2.py --self-test")
