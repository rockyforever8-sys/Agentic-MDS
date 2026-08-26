# Paste this into a NEW Colab notebook (File → New notebook).
# Do not paste Agentic_MDS.ipynb — that file is JSON and causes:
#   NameError: name 'true' is not defined

import os, pathlib, subprocess

ROOT = pathlib.Path("/content/Agentic-MDS")
if not (ROOT / "imds_agent_v2.py").exists():
    subprocess.check_call(
        [
            "git",
            "clone",
            "--depth",
            "1",
            "https://github.com/rockyforever8-sys/Agentic-MDS.git",
            str(ROOT),
        ]
    )
os.chdir(ROOT)
print("Working directory:", os.getcwd())
print("Next: Runtime → Run all in Colab_Start_Here.ipynb, or run:")
print("  !python imds_agent_v2.py --self-test")
