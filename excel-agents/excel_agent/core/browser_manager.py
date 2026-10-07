"""Browser management for the Excel Agent Engine — Chrome-over-CDP only.

One real Chrome instance (launched by scripts/setup_chrome.sh, or lazily by
this module) holds the persisted Microsoft 365 session in a dedicated
automation profile; the engine attaches to it over the Chrome DevTools
Protocol. Using the real browser's TLS fingerprint is what keeps the
Microsoft/Cloudflare login flows working.

Every knob (cdp_port, profile_dir, chrome_binary, headless, timeout) comes
from the merged config's `browser` block — the setup script reads the same
config, so the two can never disagree about port or profile.

Cleanup policy: this module NEVER kills browsers wholesale. The only kill
it can perform is scoped to the automation profile directory (a token no
personal Chrome ever has on its command line), and even that runs only when
a Chrome is confirmed listening-but-unresponsive on OUR port. A responsive
foreign process on the port is an error to report, not something to kill.
"""

import asyncio
import logging
import os
import json
import re
import platform
import socket
import subprocess
import time
import urllib.request
from pathlib import Path

logger = logging.getLogger(__name__)

_REPO_ROOT = Path(__file__).resolve().parents[2]

# Chrome binaries, tried in order when browser.chrome_binary is null.
# Canary first — less likely to auto-update mid-session.
CHROME_CDP_PATHS = [
    "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
    os.path.expanduser(
        "~/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary"
    ),
    "/usr/bin/google-chrome-canary",
    "/usr/bin/google-chrome-unstable",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    os.path.expanduser("~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
]


def get_modifier_key():
    """Meta on macOS, Control elsewhere."""
    return "Meta" if platform.system().lower() == "darwin" else "Control"


def get_select_all_key():
    return f"{get_modifier_key()}+A"


def get_save_as_key():
    return f"{get_modifier_key()}+Shift+S"


def resolve_browser_settings(config: dict) -> dict:
    """The engine-config `browser` block with paths resolved.

    profile_dir: relative paths resolve against the repo root (the
    gitignored browser_profiles/ tree), same rule as gui-agents.
    """
    block = dict(config.get("browser") or {})
    port = int(block.get("cdp_port") or 9222)
    profile = block.get("profile_dir") or "browser_profiles/chrome-excel"
    profile_path = Path(profile)
    if not profile_path.is_absolute():
        profile_path = _REPO_ROOT / profile_path
    return {
        "cdp_port": port,
        "profile_dir": str(profile_path),
        "chrome_binary": block.get("chrome_binary") or None,
        "headless": bool(block.get("headless", False)),
        "timeout": int(block.get("timeout") or 30000),
    }


def find_chrome(explicit: str | None = None) -> str | None:
    """The Chrome binary to use: explicit config value, else auto-detect."""
    if explicit:
        return explicit if os.path.exists(explicit) else None
    for path in CHROME_CDP_PATHS:
        if os.path.exists(path):
            return path
    return None


def is_cdp_port_open(port: int) -> bool:
    """TCP-level check: is something listening on the CDP port?"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.settimeout(2)
        return sock.connect_ex(("127.0.0.1", port)) == 0
    finally:
        sock.close()

def ensure_cdp_page_target(port: int) -> bool:
    """Make sure the automation Chrome has at least one page target.

    A Chrome whose last tab was closed keeps listening on the CDP port but
    Playwright's connect_over_cdp then fails with "Protocol error
    (Browser.setDownloadBehavior): Browser context management is not
    supported" (seen 2026-09-10 after a probe closed the final tab). Opening
    a blank page through the HTTP endpoint restores the default context.
    Returns True when a page target exists afterwards.
    """
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=5) as r:
            targets = json.loads(r.read().decode("utf-8") or "[]")
        if any(t.get("type") == "page" for t in targets):
            return True
        logger.warning(
            f"⚠️ Chrome on port {port} has no page target — opening a blank "
            f"page so CDP attach can succeed"
        )
        req = urllib.request.Request(
            f"http://127.0.0.1:{port}/json/new?about:blank", method="PUT"
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            r.read()
        return True
    except Exception as e:
        logger.warning(f"⚠️ Could not verify/open a page target on port {port}: {e}")
        return False


def check_cdp_health(port: int, timeout: int = 10) -> bool:
    """Does the process on the port answer CDP's /json/version?"""
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/json/version")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status == 200
    except Exception:
        return False


def kill_automation_chrome(profile_dir: str) -> None:
    """Kill ONLY the Chrome running our automation profile.

    The profile directory appears verbatim in the automation Chrome's
    command line (--user-data-dir=<profile_dir>) and in no personal
    browser's, so a -f match on it cannot touch the user's own Chrome,
    Firefox, or any sibling pipeline using a different profile.
    """
    # No leading dashes in the pattern: macOS pkill parsed the old
    # "--user-data-dir=..." token as an option ("illegal option -- -") and
    # killed nothing, so every CDP-recovery retry found Chrome "already
    # running" (2026-09-10). The trailing boundary keeps a sibling lane's
    # profile that merely extends this one (chrome-excel vs
    # chrome-excel-9227) out of reach.
    pattern = f"user-data-dir={re.escape(profile_dir)}( |$)"
    logger.warning(f"🧹 Killing automation Chrome (profile {profile_dir})")
    if os.name == "nt":
        # No pkill on Windows; scoped kill needs WMI. Refuse rather than
        # taskkill every chrome.exe.
        logger.error(
            "Scoped Chrome cleanup is not implemented on Windows — close "
            "the automation Chrome window manually."
        )
        return
    result = subprocess.run(
        ["pkill", "-f", pattern], check=False, capture_output=True, text=True
    )
    if result.returncode not in (0, 1):  # 1 = no process matched
        logger.error(f"pkill failed (rc={result.returncode}): {result.stderr.strip()}")


def launch_chrome_cdp(settings: dict) -> subprocess.Popen | None:
    """Launch Chrome with remote debugging on the configured port/profile.

    The process is detached (its own session) so it outlives the engine —
    the persisted login is shared by every attempt in a batch. stdout and
    stderr both go to DEVNULL: a PIPE nobody drains fills its 64KB buffer
    during long tasks and stalls Chrome mid-run.
    """
    chrome_path = find_chrome(settings.get("chrome_binary"))
    if not chrome_path:
        logger.error("❌ Chrome not found!")
        logger.error("   Install Chrome or Chrome Canary, or set")
        logger.error("   browser.chrome_binary in infra/configs/configs.yaml")
        return None

    profile_dir = settings["profile_dir"]
    port = settings["cdp_port"]
    args = [
        chrome_path,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-background-timer-throttling",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-features=ProfilePicker,ChromeWhatsNewUI",
        "--disable-popup-blocking",
        "--disable-session-crashed-bubble",
    ]
    if settings.get("headless"):
        args.append("--headless=new")

    logger.info(f"🚀 Launching Chrome with CDP on port {port}: {chrome_path}")
    logger.info(f"   Profile: {profile_dir}")
    os.makedirs(profile_dir, exist_ok=True)

    popen_kwargs: dict = {
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
    }
    if os.name == "nt":
        popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        popen_kwargs["start_new_session"] = True
    return subprocess.Popen(args, **popen_kwargs)


async def wait_for_chrome_ready(port: int, timeout: int = 30) -> bool:
    """Wait for the CONFIGURED port to accept CDP connections."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if is_cdp_port_open(port):
            logger.info(f"✅ Chrome is ready for CDP on port {port}")
            return True
        await asyncio.sleep(0.5)
    logger.error(f"❌ Chrome didn't start on port {port} within {timeout}s")
    return False


class BrowserManager:
    """Attach to (or lazily launch) the automation Chrome over CDP."""

    def __init__(self, config: dict):
        self.settings = resolve_browser_settings(config)
        self.cdp_port = self.settings["cdp_port"]
        self.profile_dir = self.settings["profile_dir"]
        self.cdp_url = f"http://127.0.0.1:{self.cdp_port}"
        self.timeout = self.settings["timeout"]

    def is_cdp_mode(self):
        return True  # CDP is the only mode in this pipeline

    async def _ensure_chrome(self) -> None:
        """Ensure a healthy Chrome is listening on our port, or raise."""
        if is_cdp_port_open(self.cdp_port):
            if check_cdp_health(self.cdp_port):
                logger.info(
                    f"✅ Chrome already running with CDP on port {self.cdp_port}"
                )
                return
            # Listening but not answering CDP. Restart it — scoped to our
            # profile only. If the squatter is some other process, the
            # scoped kill won't free the port and we fail with a clear
            # message instead of killing anything else.
            logger.warning(
                f"⚠️ Port {self.cdp_port} is open but not answering CDP — "
                f"restarting the automation Chrome"
            )
            kill_automation_chrome(self.profile_dir)
            await asyncio.sleep(3)
            if is_cdp_port_open(self.cdp_port):
                raise RuntimeError(
                    f"Port {self.cdp_port} is held by a process that is not "
                    f"the automation Chrome (profile {self.profile_dir}). "
                    f"Free the port or set a different browser.cdp_port in "
                    f"infra/configs/configs.yaml."
                )

        logger.info(f"🚀 Chrome not running on port {self.cdp_port}, launching...")
        process = launch_chrome_cdp(self.settings)
        if not process:
            raise RuntimeError(
                "Chrome not found. Install Chrome or Chrome Canary, or set "
                "browser.chrome_binary in infra/configs/configs.yaml."
            )
        if not await wait_for_chrome_ready(self.cdp_port):
            raise RuntimeError(
                f"Chrome failed to start with CDP on port {self.cdp_port}"
            )

    async def launch_browser(self, playwright):
        """Attach over CDP. Returns (browser, context)."""
        last_error: Exception | None = None
        for attempt in range(3):
            await self._ensure_chrome()
            ensure_cdp_page_target(self.cdp_port)
            try:
                browser = await playwright.chromium.connect_over_cdp(
                    self.cdp_url, timeout=30000
                )
                logger.info(f"✅ Connected to Chrome via CDP (port {self.cdp_port})")
            except Exception as e:
                last_error = e
                logger.warning(
                    f"⚠️ CDP connection attempt {attempt + 1}/3 failed: {e}"
                )
                if attempt < 2:
                    # Restart our own Chrome only, then retry.
                    kill_automation_chrome(self.profile_dir)
                    await asyncio.sleep(3)
                    continue
                raise RuntimeError(
                    f"Failed to connect to Chrome on port {self.cdp_port} "
                    f"after 3 attempts: {e}"
                ) from e

            contexts = browser.contexts
            if contexts:
                context = contexts[0]
                logger.info(
                    f"✅ Using existing context ({len(context.pages)} page(s))"
                )
            else:
                context = await browser.new_context(ignore_https_errors=True)
                logger.info("✅ Created new browser context")
            context.set_default_timeout(self.timeout)
            return browser, context

        raise RuntimeError(
            f"Failed to connect to Chrome after 3 attempts: {last_error}"
        )

    async def close_browser(self, context, browser=None):
        """CDP mode: the shared Chrome (and its login) stays alive for the
        next attempt. Task pages are closed by the engine itself."""
        logger.debug("CDP mode: keeping shared Chrome/context alive")
