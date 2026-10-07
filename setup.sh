#!/usr/bin/env bash
# setup.sh — create the Python environment for MBABench with uv.
#
#   ./setup.sh            # creates .venv (or venv_path from config/config.yaml) and installs everything
#
# Run it again after pulling; it updates the environment in place.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_ROOT"

if ! command -v uv >/dev/null 2>&1; then
    echo "uv is not installed. Install it with:  curl -LsSf https://astral.sh/uv/install.sh | sh" >&2
    exit 1
fi

# Create config/config.yaml from the defaults on first run.
if [[ ! -f config/config.yaml ]]; then
    cp config/config_default.yaml config/config.yaml
    echo "Created config/config.yaml from config/config_default.yaml — put your API keys there (or export them)."
fi

# venv_path from config/config.yaml (null = <repo>/.venv).
VENV_PATH="$(sed -nE 's/^venv_path:[[:space:]]*"?([^"#]*)"?.*$/\1/p' config/config.yaml | head -1 | xargs || true)"
case "${VENV_PATH:-}" in
    ""|null|Null|NULL|"~"|none|None|NONE) VENV_PATH="$REPO_ROOT/.venv" ;;
    "~/"*) VENV_PATH="$HOME/${VENV_PATH#\~/}" ;;
    /*) ;;
    *) VENV_PATH="$REPO_ROOT/$VENV_PATH" ;;
esac
export UV_PROJECT_ENVIRONMENT="$VENV_PATH"

echo "Syncing the environment at $VENV_PATH ..."
uv sync

# The GUI pipeline's fallback browser path needs Playwright's Chromium build.
uv run playwright install chromium >/dev/null 2>&1 || echo "note: playwright install chromium failed; only the GUI pipeline's fallback path needs it."

echo
echo "Done. Next:  uv run python scripts/download_dataset.py"
