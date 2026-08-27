# Paste this entire cell into Google Colab and run it.
# Do not paste a .ipynb JSON file.

import os, pathlib, shutil, subprocess, sys

ROOT = pathlib.Path("/content/ppap_agent_repo")
CANDIDATES = [
    ("https://github.com/rockyforever8-sys/Agentic-MDS.git", "cursor/ppap-quality-agent-17d5"),
    ("https://github.com/rockyforever8-sys/Agentic-MDS.git", "cursor/ppap-langgraph-prototype-17d5"),
]


def _ok(path: pathlib.Path) -> bool:
    return (path / "ppap_agent" / "__init__.py").exists()


if not _ok(ROOT):
    last = None
    for repo, ref in CANDIDATES:
        try:
            if ROOT.exists():
                shutil.rmtree(ROOT)
            print(f"Cloning {repo} @{ref} ...")
            subprocess.check_call(["git", "clone", "--depth", "1", "--branch", ref, repo, str(ROOT)])
            if _ok(ROOT):
                print(f"Ready: {ref}")
                break
        except subprocess.CalledProcessError as exc:
            last = exc
            if ROOT.exists():
                shutil.rmtree(ROOT, ignore_errors=True)
    else:
        raise RuntimeError(f"Clone failed: {last}")

os.chdir(ROOT)
sys.path.insert(0, str(ROOT))

try:
    from IPython import get_ipython
    get_ipython().run_line_magic("pip", "install -q langgraph langchain-core rich")
except Exception:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "langgraph", "langchain-core", "rich"])

from colab_ppap_demo import main
main()
