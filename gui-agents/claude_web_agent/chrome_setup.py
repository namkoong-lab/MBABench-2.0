#!/usr/bin/env python3
"""One-time interactive Chrome setup for the gui-agents pipeline.

Launches the automation Chrome for one provider (claude or chatgpt) with the
same binary, CDP port and profile directory the engine attaches to — both read
infra/configs — and leaves it open so you can sign in to claude.ai or
chatgpt.com, including 2FA. The session persists in the profile directory, so
automated runs reuse the login until the cookies expire.

Run via scripts/setup_chrome.sh, or directly:

    uv run python -m claude_web_agent.chrome_setup claude
    uv run python -m claude_web_agent.chrome_setup chatgpt
"""
import argparse
import asyncio
import socket
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from infra.configs import ConfigError, load_configs  # noqa: E402
from claude_web_agent.browser_manager import (  # noqa: E402
    CDP_PORT,
    DEFAULT_PROFILE_DIRS,
    launch_chrome_cdp,
    resolve_profile_dir,
)

SIGN_IN_URL = {"claude": "https://claude.ai", "chatgpt": "https://chatgpt.com"}


def _settings(provider: str) -> tuple[int, str]:
    try:
        cfg = load_configs()
    except ConfigError as e:
        print(f"❌ Config load failed:\n{e}")
        sys.exit(2)
    block = getattr(getattr(cfg, f"{provider}_web", None), "browser", None)
    port = int(getattr(block, "cdp_port", None) or CDP_PORT)
    profile = getattr(block, "profile_dir", None) or DEFAULT_PROFILE_DIRS[provider]
    return port, resolve_profile_dir(profile)


def _port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", port)) == 0


async def _wait_ready(port: int, timeout: int = 30) -> bool:
    for _ in range(timeout * 2):
        if _port_open(port):
            return True
        await asyncio.sleep(0.5)
    return False


async def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("provider", choices=("claude", "chatgpt"))
    args = ap.parse_args()
    port, profile = _settings(args.provider)

    print("=" * 70)
    print(f"gui-agents Chrome setup — {args.provider}")
    print(f"  CDP port : {port}")
    print(f"  Profile  : {profile}")
    print("=" * 70)

    if _port_open(port):
        print(f"✅ Chrome is already running with CDP on port {port}.")
        print(f"   Sign in at {SIGN_IN_URL[args.provider]} in that window if you haven't yet.")
        return 0

    if launch_chrome_cdp(headless=False, profile_dir=profile, cdp_port=port) is None:
        return 1
    if not await _wait_ready(port):
        print(f"❌ Chrome did not open port {port} within 30 s.")
        return 1

    print()
    print("👤 In the Chrome window that just opened:")
    print(f"   1. Go to {SIGN_IN_URL[args.provider]}")
    print("   2. Sign in with the paid account the runs will use (complete any 2FA)")
    print("   3. Confirm the composer loads without a sign-in prompt")
    print()
    print("The session persists in the profile. Leave Chrome running for automated")
    print("runs, or close it: the engine relaunches it on the same port and profile.")
    print("This script exits now; Chrome stays open (it is detached).")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
