#!/usr/bin/env python3
"""One-time interactive Chrome setup for the excel-agents pipeline.

Launches the automation Chrome (same binary/port/profile the engine will
attach to — both read infra/configs) and holds it open so you can complete
the Microsoft 365 sign-in, including 2FA. The session persists in the
profile directory, so subsequent automated runs reuse the login until the
cookies expire.

Run via scripts/setup_chrome.sh, or directly:

    uv run python -m excel_agent.chrome_browser

There is nothing to keep in sync between this script and the runtime: both
resolve the port, binary, and profile from the same merged config.
"""

import asyncio
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from infra.configs import ConfigError, load_configs

from excel_agent.core.browser_manager import (
    check_cdp_health,
    is_cdp_port_open,
    launch_chrome_cdp,
    resolve_browser_settings,
    wait_for_chrome_ready,
)


def _load_settings() -> dict:
    try:
        cfg = load_configs()
    except ConfigError as e:
        print(f"❌ Config load failed:\n{e}")
        sys.exit(2)
    browser = getattr(cfg, "browser", None)
    block = {
        k: getattr(browser, k, None)
        for k in ("cdp_port", "profile_dir", "chrome_binary", "headless", "timeout")
    } if browser else {}
    return resolve_browser_settings({"browser": block})


async def main() -> int:
    settings = _load_settings()
    port = settings["cdp_port"]

    print("=" * 70)
    print("Excel-agents Chrome setup")
    print(f"  CDP port : {port}")
    print(f"  Profile  : {settings['profile_dir']}")
    print("=" * 70)

    if is_cdp_port_open(port):
        if check_cdp_health(port):
            print(f"✅ Chrome already running with CDP on port {port}.")
            print("   Sign in to OneDrive in that window if you haven't yet:")
            print("   https://onedrive.live.com")
            return 0
        print(
            f"❌ Port {port} is open but not answering CDP. Close whatever "
            f"holds it (or change browser.cdp_port in configs.yaml) and re-run."
        )
        return 1

    process = launch_chrome_cdp(settings)
    if not process:
        return 1
    if not await wait_for_chrome_ready(port):
        return 1

    print()
    print("👤 In the Chrome window that just opened:")
    print("   1. Go to https://onedrive.live.com")
    print("   2. Sign in with the Microsoft 365 account (complete any 2FA)")
    print("   3. Confirm you can browse your files without a sign-in prompt")
    print()
    print("The session persists in the profile — leave Chrome running for")
    print("automated runs, or close it; the engine relaunches it as needed.")
    print("This script exits now; Chrome stays open (it is detached).")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
