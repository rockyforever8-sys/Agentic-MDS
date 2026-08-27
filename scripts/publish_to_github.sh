#!/usr/bin/env bash
# Publish Agentic-PPAP to its own GitHub repository.
#
# Prerequisites:
#   1. Create an EMPTY repo at https://github.com/rockyforever8-sys/Agentic-PPAP
#      (no README, no .gitignore, no license — empty repo)
#   2. Run from this project root:
#        bash scripts/publish_to_github.sh
#
set -euo pipefail

REPO_URL="${PPAP_REPO_URL:-https://github.com/rockyforever8-sys/Agentic-PPAP.git}"
BRANCH="${PPAP_BRANCH:-main}"

echo "Publishing to ${REPO_URL} (branch: ${BRANCH})..."

git remote remove ppap-origin 2>/dev/null || true
if [[ -n "${PPAP_GITHUB_TOKEN:-}" ]]; then
  AUTH_URL="https://${PPAP_GITHUB_TOKEN}@github.com/rockyforever8-sys/Agentic-PPAP.git"
  git remote add ppap-origin "${AUTH_URL}"
else
  git remote add ppap-origin "${REPO_URL}"
fi

# Push current standalone project as main on the new repo
if [[ -n "${PPAP_GITHUB_TOKEN:-}" ]]; then
  GIT_CONFIG_GLOBAL=/dev/null GIT_CONFIG_SYSTEM=/dev/null \
    git -c credential.helper= push -u ppap-origin HEAD:"${BRANCH}"
else
  git push -u ppap-origin HEAD:"${BRANCH}"
fi

echo ""
echo "Done! Repository: https://github.com/rockyforever8-sys/Agentic-PPAP"
echo "Colab: https://colab.research.google.com/github/rockyforever8-sys/Agentic-PPAP/blob/main/PPAP_Colab_Start_Here.ipynb"
