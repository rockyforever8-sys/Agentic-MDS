#!/usr/bin/env bash
# Publish Agentic-PPAP from your local machine (uses your GitHub login).
#
# Run this on your computer after creating the empty Agentic-PPAP repo:
#   bash scripts/publish_local.sh
#
set -euo pipefail

SOURCE_REPO="${SOURCE_REPO:-https://github.com/rockyforever8-sys/Agentic-MDS.git}"
SOURCE_BRANCH="${SOURCE_BRANCH:-cursor/ppap-quality-agent-17d5}"
TARGET_REPO="${TARGET_REPO:-https://github.com/rockyforever8-sys/Agentic-PPAP.git}"
TARGET_BRANCH="${TARGET_BRANCH:-main}"
WORKDIR="${WORKDIR:-/tmp/agentic-ppap-publish}"

rm -rf "${WORKDIR}"
git clone --depth 1 --branch "${SOURCE_BRANCH}" "${SOURCE_REPO}" "${WORKDIR}"
cd "${WORKDIR}"
git remote set-url origin "${TARGET_REPO}"
git push -u origin "HEAD:${TARGET_BRANCH}"

echo ""
echo "Published to ${TARGET_REPO} (${TARGET_BRANCH})"
echo "https://github.com/rockyforever8-sys/Agentic-PPAP"
