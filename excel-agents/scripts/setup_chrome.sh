#!/usr/bin/env bash
# One-time interactive Microsoft 365 sign-in for the excel-agents pipeline.
# Thin wrapper: all port/profile/binary settings come from infra/configs,
# which the runtime reads too — nothing to keep in sync.
set -euo pipefail
cd "$(dirname "$0")/.."
exec uv run python -m excel_agent.chrome_browser "$@"
