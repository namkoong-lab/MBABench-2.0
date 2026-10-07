#!/usr/bin/env bash
# One-time interactive sign-in for the gui-agents pipeline: launches the
# automation Chrome for one provider and leaves it open for you to sign in.
#
#   scripts/setup_chrome.sh claude     # claude.ai
#   scripts/setup_chrome.sh chatgpt    # chatgpt.com
#
# Port and profile come from infra/configs (configs.default.yaml, overridden by
# the gitignored infra/configs/configs.yaml); the engine reads the same values.
set -euo pipefail
cd "$(dirname "$0")/.."
exec uv run python -m claude_web_agent.chrome_setup "$@"
