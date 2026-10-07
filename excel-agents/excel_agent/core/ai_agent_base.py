"""
AI Agent base class for Excel add-in automation.

Provides shared logic for TabAI, Claude, and future AI agents.
"""

import asyncio
import logging
import re
import time
from abc import ABC, abstractmethod
from enum import Enum

from .browser_manager import get_select_all_key

logger = logging.getLogger(__name__)


class AIAgentState(Enum):
    """AI agent interface states."""

    RUNNING = "running"
    READY_TO_SEND = "ready_to_send"
    IDLE = "idle"
    UNKNOWN = "unknown"


# Frames that are Excel Online / OneDrive chrome, never the add-in. The
# add-in-frame brute-force passes must skip these — matching Excel's own
# contenteditable surfaces would type the benchmark prompt into the
# workbook being graded.
_HOST_FRAME_NAME_PREFIXES = ("wacframe",)
_HOST_FRAME_URL_MARKERS = ("officeapps", "onedrive.live.com", "sharepoint")


def is_host_frame(frame) -> bool:
    name = (frame.name or "").lower()
    if name.startswith(_HOST_FRAME_NAME_PREFIXES):
        return True
    url = (frame.url or "").lower()
    return any(m in url for m in _HOST_FRAME_URL_MARKERS)


async def _first_visible(page, selector: str):
    """Find first visible element matching selector across page and frames."""
    try:
        for ctx in [page] + page.frames:
            try:
                loc = ctx.locator(selector).first
                if await loc.count() > 0 and await loc.is_visible():
                    return loc
            except Exception:
                # Context may be closed, skip it
                continue
    except Exception:
        # Page may be closed entirely
        pass
    return None


class AIAgentCore(ABC):
    """Base class for AI agent automation (TabAI, Claude, etc.)"""

    def __init__(self, page, config, shutdown_event=None, completion_logger=None):
        """
        Initialize AI agent core.

        Args:
            page: Playwright page instance
            config: Configuration dictionary
            shutdown_event: Optional shutdown event
            completion_logger: Optional completion logger for timing
        """
        self.page = page
        self.config = config
        self.shutdown_event = shutdown_event
        self.completion_logger = completion_logger
        # Set by a core when the vendor answered a prompt with a capacity /
        # outage notice instead of doing the work. The engine classifies such
        # an attempt as infra (unrecorded, retried), never as success or as
        # the model's own failure.
        self.provider_unavailable = False
        # Epoch seconds when a vendor usage limit resets (parsed from the
        # panel banner), or None. The engine holds until then before the
        # infra retry so the lane pauses itself instead of churning.
        self.provider_retry_after = None

    # Panel-level vendor notices that are NOT chat replies (no <article>):
    # 2026-09-11 21:51 the Claude add-in showed "Rate limit exceeded. Limits
    # will reset at Fri Sep 11 2026 22:40:00 GMT-0400 ... Get extra usage"
    # plus "Something went wrong — let us know" after one tool step; no
    # response article ever rendered and wait_for_completion sat in
    # "Waiting for Claude to start" for over an hour (would have recorded a
    # TIMEOUT row at the 5h cap). Phrases here are ones a modelling reply
    # never contains.
    _PANEL_NOTICE_RE = re.compile(
        r"rate limit exceeded"
        r"|limits will reset"
        r"|get extra usage"
        r"|something went wrong"
        r"|(high|unusually high|heavy|peak) demand"
        r"|unable to (respond|process|complete)"
        r"|temporarily unavailable",
        re.IGNORECASE,
    )
    # Whitespace-tolerant: panel innerText can carry non-breaking spaces and
    # line breaks inside the date (2026-09-12 03:01: the banner matched but
    # the space-literal form of this pattern parsed no reset time).
    _LIMIT_RESET_RE = re.compile(
        r"reset\s+at\s+(\w{3}\s+\w{3}\s+\d{1,2}\s+\d{4}\s+\d{1,2}:\d{2}:\d{2})\s+GMT\s*([+-]\d{4})",
        re.IGNORECASE,
    )
    # Locale short form seen 2026-09-12 03:06: "Limits will reset at
    # 9/12/2026, 3:40:00 AM." — local time, no zone.
    _LIMIT_RESET_SHORT_RE = re.compile(
        r"reset\s+at\s+(\d{1,2}/\d{1,2}/\d{4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?)\s*([AP]M)",
        re.IGNORECASE,
    )

    @classmethod
    def panel_notice(cls, panel_text: str | None) -> str | None:
        """The matched vendor-notice phrase in `panel_text`, else None."""
        m = cls._PANEL_NOTICE_RE.search(panel_text or "")
        return m.group(0) if m else None

    @classmethod
    def parse_limit_reset_epoch(cls, panel_text: str | None) -> float | None:
        """Epoch seconds of 'Limits will reset at <date> GMT±hhmm', else None."""
        text = (panel_text or "").replace(" ", " ")
        from datetime import datetime
        m = cls._LIMIT_RESET_RE.search(text)
        if m:
            try:
                stamp = re.sub(r"\s+", " ", m.group(1)).strip()
                return datetime.strptime(
                    f"{stamp} {m.group(2)}", "%a %b %d %Y %H:%M:%S %z"
                ).timestamp()
            except Exception:
                return None
        m = cls._LIMIT_RESET_SHORT_RE.search(text)
        if m:
            try:
                clock = m.group(2) if m.group(2).count(":") == 2 else m.group(2) + ":00"
                local = datetime.strptime(
                    f"{m.group(1)} {clock} {m.group(3).upper()}", "%m/%d/%Y %I:%M:%S %p"
                )
                return local.timestamp()  # naive = local wall clock, as the panel shows it
            except Exception:
                return None
        return None

    # Vendor capacity / outage notices. 2026-09-11 ~10:00: Anthropic "high
    # demand" — the Claude add-in replied with one short notice in 22s, the
    # Stop button came and went, and the engine recorded a 1-formula workbook
    # as SUCCESS (attempt 1404); an earlier task (1402) was cut off the same
    # way after building most of the model. Only a SHORT last reply is tested
    # so a genuine model reply that mentions "capacity" (a case about plant
    # capacity, say) can never trip it.
    _CAPACITY_NOTICE_MAX_CHARS = 600
    _CAPACITY_NOTICE_RE = re.compile(
        r"(high|unusually high|heavy|peak) demand"
        r"|capacity (constraints?|issues?|limits?)"
        r"|(is|are) (currently )?(at|over) capacity"
        r"|unable to (respond|process|complete)"
        r"|try again (later|in a (few|little))"
        r"|temporarily unavailable"
        r"|something went wrong"
        r"|an error occurred",
        re.IGNORECASE,
    )

    @classmethod
    def looks_like_capacity_notice(cls, text: str | None) -> bool:
        """True when `text` is a short vendor notice about demand/capacity."""
        t = (text or "").strip()
        if not t or len(t) > cls._CAPACITY_NOTICE_MAX_CHARS:
            return False
        return bool(cls._CAPACITY_NOTICE_RE.search(t))

    @abstractmethod
    def get_agent_type(self) -> str:
        """Return agent type for logging: 'tabai', 'claude_excel_agent', etc."""
        pass

    @abstractmethod
    def get_addon_name(self) -> str:
        """Return addon name for installation: 'TabAI', 'Claude by Anthropic', etc."""
        pass

    @abstractmethod
    def get_open_button_text(self) -> str:
        """Return open button text: 'Open TabAI', 'Open Claude', etc."""
        pass

    def requires_addins_menu(self) -> bool:
        """Return True if agent requires Add-ins menu (default: False)."""
        return False

    def get_ribbon_launcher_name(self) -> str | None:
        """Label of the add-in's OWN ribbon launcher, or None.

        An installed add-in gets a RibbonButton of its own in the Home
        ribbon ("Claude", "ChatGPT" — probed 2026-09-10); at narrow window
        widths Excel Online collapses it into the "More Options" overflow
        flyout as a menuitem with the same name. Clicking either opens the
        panel directly, skipping the Add-ins -> My Add-ins popup.
        """
        return None

    def get_config_section(self) -> dict:
        """Get agent-specific config section (tabai, claude_excel_agent, etc.)."""
        return self.config.get(self.get_agent_type(), {})

    async def get_state(self) -> AIAgentState:
        """
        Get current AI agent state.

        Returns:
            AIAgentState enum value
        """
        # Priority 1: running
        stop = await _first_visible(self.page, '[data-testid="stop-button"]')
        if stop:
            return AIAgentState.RUNNING

        # Check mic presence (mic exists both idle and ready)
        mic = await _first_visible(self.page, '[data-testid="mic-button"]')

        # Check send
        send = await _first_visible(self.page, '[data-testid="send-button"]')
        if send:
            enabled = False
            try:
                # Prefer is_enabled; fall back to absence of disabled attribute
                enabled = await send.is_enabled()
                if not enabled:
                    disabled_attr = await send.get_attribute("disabled")
                    enabled = disabled_attr is None
            except Exception:
                pass

            if enabled:
                return AIAgentState.READY_TO_SEND
            if mic:
                return AIAgentState.IDLE

        return AIAgentState.UNKNOWN

    # ------------------------------------------------------------------
    # JS snippet: find and click an add-in tile inside the Fluent UI
    # ms-Layer popup (the "My Add-ins" dropdown).
    #
    # The add-in tiles are generic <DIV>s with no aria-label, title, or
    # role.  The display name may be longer than get_addon_name() — e.g.
    # "Claude by Anthropic for Excel" vs "Claude by Anthropic".
    #
    # Algorithm:
    #   1. Collect every element inside a .ms-Layer-content container
    #      whose textContent includes the addon name (substring match).
    #   2. Pick the element with the shortest textContent (= narrowest /
    #      most specific match — the tile, not a parent container).
    #   3. Click it.
    # ------------------------------------------------------------------
    _JS_CLICK_ADDON_IN_LAYER = """
    (addonName) => {
        const lower = addonName.toLowerCase();
        let layers = Array.from(document.querySelectorAll(
            '.ms-Layer-content, [class*="Layer-content"]'
        ));
        // Prefer the My Add-ins popup: the ribbon overflow flyout is a
        // layer too and lists the add-in's own launcher under the same
        // name, so an unscoped scan can click the wrong layer.
        const mine = layers.filter(l =>
            (l.textContent || '').toLowerCase().includes('my add-ins'));
        if (mine.length) layers = mine;
        let bestEl = null;
        let bestLen = Infinity;
        for (const layer of layers) {
            const walker = document.createTreeWalker(
                layer, NodeFilter.SHOW_ELEMENT
            );
            while (walker.nextNode()) {
                const el = walker.currentNode;
                const text = (el.textContent || '').trim().toLowerCase();
                if (text.includes(lower) && text.length < bestLen) {
                    bestEl = el;
                    bestLen = text.length;
                }
            }
        }
        if (bestEl) {
            bestEl.click();
            return bestLen;
        }
        return -1;
    }
    """

    # Ribbon overflow ("More Options") anchor, most to least stable.
    _OVERFLOW_ANCHOR_SELECTORS = (
        "#RibbonOverflowMenu-overflow",
        '[data-unique-id="Overflow-RibbonOverflowMenu-overflow"]',
        'button[data-automation-type="RibbonFlyoutAnchor"][aria-label="More Options"]',
    )

    def _frames_excel_first(self):
        """Page frames with the Excel host frame first (where the ribbon lives)."""
        frames = list(self.page.frames)
        frames.sort(key=lambda f: (f.name or "") != "WacFrame_Excel_0")
        return frames

    async def _click_overflow_menuitem(self, name: str, wait_seconds: float = 4.0) -> bool:
        """Open the ribbon's "More Options" overflow flyout and click the
        menuitem whose name/label is exactly `name` (case-insensitive).

        Excel Online collapses ribbon groups into this flyout at narrow
        window widths (~960px); the collapsed items keep their names
        (`[role=menuitem][name="Add-ins"]`, `[name="ChatGPT"]`, ...) but
        their `data-unique-id` is `Overflow-AddinControlN` with N in
        install order, so match on the name, never the id. Escape-closes
        the flyout and returns False when the item is not there.
        """
        target = name.strip().lower()
        for frame in self._frames_excel_first():
            anchor = None
            for sel in self._OVERFLOW_ANCHOR_SELECTORS:
                try:
                    el = await asyncio.wait_for(frame.query_selector(sel), timeout=3)
                    if el and await el.is_visible() and await el.is_enabled():
                        anchor = el
                        break
                except (Exception, asyncio.TimeoutError):
                    continue
            if not anchor:
                continue
            try:
                await anchor.click(timeout=3000)
            except Exception as e:
                logger.debug(f"🔍 Overflow anchor click failed: {e}")
                continue
            deadline = asyncio.get_event_loop().time() + wait_seconds
            while asyncio.get_event_loop().time() < deadline:
                await asyncio.sleep(0.3)
                try:
                    items = await frame.query_selector_all('[role="menuitem"]')
                except Exception:
                    items = []
                for item in items:
                    try:
                        if not await item.is_visible():
                            continue
                        label = (
                            await item.get_attribute("name")
                            or await item.get_attribute("aria-label")
                            or await item.text_content()
                            or ""
                        ).strip().lower()
                        if label == target:
                            await item.click(timeout=3000)
                            return True
                    except Exception:
                        continue
            try:
                await frame.press("body", "Escape")
            except Exception:
                pass
        return False

    async def _click_ribbon_launcher(self, name: str, max_seconds: int) -> bool:
        """Click the add-in's own launcher: the Home-ribbon RibbonButton
        labelled `name` when visible, else the same entry inside the
        "More Options" overflow flyout. Retries until `max_seconds` — the
        ribbon can take tens of seconds to materialize on a fresh copy.
        """
        deadline = asyncio.get_event_loop().time() + max_seconds
        selector = (
            f'button[data-automation-type="RibbonButton"][aria-label="{name}"]'
        )
        while asyncio.get_event_loop().time() < deadline:
            for frame in self._frames_excel_first():
                try:
                    el = await asyncio.wait_for(frame.query_selector(selector), timeout=3)
                    if el and await el.is_visible() and await el.is_enabled():
                        await el.click(timeout=3000)
                        logger.info(f"✅ Clicked '{name}' ribbon button")
                        return True
                except (Exception, asyncio.TimeoutError):
                    continue
            if await self._click_overflow_menuitem(name):
                logger.info(f"✅ Clicked '{name}' in the ribbon overflow menu")
                return True
            await asyncio.sleep(1.5)
        logger.info(f"🔍 No '{name}' ribbon launcher found within {max_seconds}s")
        return False

    async def _click_addon_in_layer(self, addon_name: str) -> bool:
        """Click an add-in tile in the My Add-ins Fluent UI popup.

        Uses JS evaluate to find the narrowest element inside a
        ms-Layer-content container whose text contains ``addon_name``,
        then clicks it.  Retries up to 5 times with 2 s between attempts
        to allow the popup to render.

        Returns True if clicked, False otherwise.
        """
        excel_frame = next(
            (f for f in self.page.frames if f.name == "WacFrame_Excel_0"),
            None,
        )
        targets = [excel_frame] if excel_frame else self.page.frames

        for attempt in range(5):
            for frame in targets:
                try:
                    result = await asyncio.wait_for(
                        frame.evaluate(self._JS_CLICK_ADDON_IN_LAYER, addon_name),
                        timeout=3,
                    )
                    if isinstance(result, (int, float)) and result >= 0:
                        logger.info(
                            f"✅ Clicked '{addon_name}' in ms-Layer"
                            f" (textLen={result}, frame={frame.name or '?'})"
                        )
                        return True
                except (Exception, asyncio.TimeoutError):
                    continue
            if attempt < 4:
                logger.debug(f"⏳ Add-in not found in layer, retry {attempt + 1}/5...")
                await asyncio.sleep(2.0)

        logger.error(f"❌ Could not find '{addon_name}' in ms-Layer popup")
        return False

    async def find_and_click(self, max_seconds: int = 15) -> bool:
        """
        Find and click AI agent, then open button, then verify panel opened.

        Flow depends on agent:
        - TabAI: Click ribbon tab → Click "Open TabAI" → Panel opens
        - Claude: Click "Add-ins" → Click "Claude by Anthropic" → Panel opens

        Args:
            max_seconds: Maximum time to complete the process

        Returns:
            True if panel opened successfully
        """
        addon_name = self.get_addon_name()
        open_button_text = self.get_open_button_text()

        # Step 0: the add-in's own ribbon launcher (direct button when the
        # ribbon is wide, "More Options" overflow menuitem when it is
        # collapsed). Both open the panel directly — no My Add-ins popup.
        launcher = self.get_ribbon_launcher_name()
        if launcher:
            logger.info(f"🔍 Step 0: Looking for the '{launcher}' ribbon launcher...")
            if await self._click_ribbon_launcher(launcher, max_seconds):
                await asyncio.sleep(3.0)
                if await self._verify_panel_opened():
                    return True
                logger.warning(
                    f"⚠️ '{launcher}' launcher clicked but no panel — "
                    f"falling back to the Add-ins menu path"
                )

        # Step 1: Click either Add-ins (Claude) or addon ribbon tab (TabAI)
        # Each step gets its own full budget — Step 0 can't starve Step 1,
        # Step 1 can't starve Step 2. The budget MUST start here, after Step 0:
        # 2026-09-11 (full run, 9222 lane, tasks 16-17) the Claude launcher
        # dropped out of the ribbon overflow mid-run, Step 0 spent the whole
        # budget looking for it, and a budget computed before Step 0 left this
        # fallback with zero seconds — "Could not find Add-ins after retries"
        # fired instantly on every attempt and a task went infra with the
        # Add-ins path never tried.
        step1_end = asyncio.get_event_loop().time() + max_seconds
        if self.requires_addins_menu():
            logger.info("🔍 Step 1: Looking for Add-ins button...")
            target_name = "Add-ins"
        else:
            logger.info(f"🔍 Step 1: Looking for {addon_name} menu item...")
            target_name = addon_name

        # Use data-automation-type="RibbonTab" to ONLY match ribbon tabs, not filenames
        # This attribute is Microsoft-specific and only exists on ribbon tabs (Home, Insert, TabAI, etc.)
        # The filename element uses data-unique-id="CommitNewDocumentTitle-input" instead
        menu_selectors = [
            # Primary: Microsoft's ribbon tab automation attribute (most reliable)
            f'[data-automation-type="RibbonTab"]:has-text("{target_name}")',
            # Fallback: role="tab" combined with aria attributes (exact match)
            f'button[role="tab"][aria-label="{target_name}"]',
            f'[role="tab"] .ms-Button-label:text-is("{target_name}")',
            # Last resort: aria-label/title on any button (these are exact matches)
            f'button[aria-label="{target_name}"]',
            f'button[title="{target_name}"]',
        ]

        menu_clicked = False
        PER_CALL_TIMEOUT = 3  # Hard cap per Playwright call to prevent hangs

        while asyncio.get_event_loop().time() < step1_end and not menu_clicked:
            for selector in menu_selectors:
                if asyncio.get_event_loop().time() >= step1_end:
                    break
                try:
                    # Try in main page (with per-call timeout)
                    element = await asyncio.wait_for(
                        self.page.query_selector(selector),
                        timeout=PER_CALL_TIMEOUT,
                    )
                    if element and await asyncio.wait_for(
                        element.is_visible(), timeout=PER_CALL_TIMEOUT
                    ):
                        try:
                            await asyncio.wait_for(
                                element.scroll_into_view_if_needed(),
                                timeout=PER_CALL_TIMEOUT,
                            )
                            await asyncio.sleep(0.3)
                        except Exception:
                            pass
                        await asyncio.wait_for(
                            element.click(), timeout=PER_CALL_TIMEOUT
                        )
                        logger.info(f"✅ Clicked {target_name} menu: {selector}")
                        menu_clicked = True
                        break

                    # Try in frames
                    for frame in self.page.frames:
                        if asyncio.get_event_loop().time() >= step1_end:
                            break
                        try:
                            element = await asyncio.wait_for(
                                frame.query_selector(selector),
                                timeout=PER_CALL_TIMEOUT,
                            )
                            if element and await asyncio.wait_for(
                                element.is_visible(), timeout=PER_CALL_TIMEOUT
                            ):
                                try:
                                    await asyncio.wait_for(
                                        element.scroll_into_view_if_needed(),
                                        timeout=PER_CALL_TIMEOUT,
                                    )
                                    await asyncio.sleep(0.3)
                                except Exception:
                                    pass
                                await asyncio.wait_for(
                                    element.click(), timeout=PER_CALL_TIMEOUT
                                )
                                logger.info(
                                    f"✅ Clicked {target_name} menu in frame: {selector}"
                                )
                                menu_clicked = True
                                break
                        except (Exception, asyncio.TimeoutError):
                            continue

                    if menu_clicked:
                        break
                except asyncio.TimeoutError:
                    logger.debug(f"⏱️ query_selector timed out for: {selector}")
                    continue
                except Exception:
                    continue

            if not menu_clicked and self.requires_addins_menu():
                # Collapsed ribbon: "Add-ins" lives inside the "More
                # Options" overflow flyout as a menuitem.
                if await self._click_overflow_menuitem(target_name):
                    logger.info(f"✅ Clicked {target_name} via the ribbon overflow menu")
                    menu_clicked = True
                    break
            if not menu_clicked:
                await asyncio.sleep(1.5)

        if not menu_clicked:
            logger.error(f"❌ Could not find {target_name} after retries")
            return False

        # Step 2: Click addon name (Claude) or "Open" button (TabAI)
        # Fresh budget — Step 1's time usage doesn't affect Step 2
        step2_end = asyncio.get_event_loop().time() + max_seconds

        if self.requires_addins_menu():
            logger.info(f"🔍 Step 2: Looking for '{addon_name}' in add-ins list...")
            await asyncio.sleep(3.0)

            # The My Add-ins popup is a Fluent UI Layer inside WacFrame_Excel_0.
            # Add-in tiles are generic DIVs with no aria-label/title/role, and
            # the display name may differ from get_addon_name() (e.g. the DOM
            # text is "Claude by Anthropic for Excel", not "Claude by Anthropic").
            # Playwright exact-match selectors fail in this scenario.
            #
            # Strategy: use JS evaluate in the Excel frame to find the narrowest
            # element inside the ms-Layer popup whose text contains the addon
            # name, then click it. This is safe from filename collisions because
            # the ms-Layer popup is separate from the spreadsheet area.
            open_clicked = await self._click_addon_in_layer(addon_name)
        else:
            # Non-addins-menu path (e.g. TabAI "Open" button)
            logger.info(f"🔍 Step 2: Looking for '{open_button_text}' button...")
            await asyncio.sleep(3.0)
            step2_selectors = [
                f'text="{open_button_text}"',
                f'button:text-is("{open_button_text}")',
                f'[aria-label="{open_button_text}"]',
                f'[title="{open_button_text}"]',
            ]

            open_clicked = False
            remaining_time = step2_end - asyncio.get_event_loop().time()

            for _ in range(int(max(remaining_time, 1) * 1.5)):
                if asyncio.get_event_loop().time() >= step2_end:
                    break
                for selector in step2_selectors:
                    if asyncio.get_event_loop().time() >= step2_end:
                        break
                    try:
                        element = await asyncio.wait_for(
                            self.page.query_selector(selector),
                            timeout=PER_CALL_TIMEOUT,
                        )
                        if element and await asyncio.wait_for(
                            element.is_visible(), timeout=PER_CALL_TIMEOUT
                        ):
                            try:
                                await asyncio.wait_for(
                                    element.scroll_into_view_if_needed(),
                                    timeout=PER_CALL_TIMEOUT,
                                )
                                await asyncio.sleep(0.3)
                            except Exception:
                                pass
                            await asyncio.wait_for(
                                element.click(), timeout=PER_CALL_TIMEOUT
                            )
                            logger.info(f"✅ Clicked: {selector}")
                            open_clicked = True
                            break

                        for frame in self.page.frames:
                            if asyncio.get_event_loop().time() >= step2_end:
                                break
                            try:
                                element = await asyncio.wait_for(
                                    frame.query_selector(selector),
                                    timeout=PER_CALL_TIMEOUT,
                                )
                                if element and await asyncio.wait_for(
                                    element.is_visible(), timeout=PER_CALL_TIMEOUT
                                ):
                                    try:
                                        await asyncio.wait_for(
                                            element.scroll_into_view_if_needed(),
                                            timeout=PER_CALL_TIMEOUT,
                                        )
                                        await asyncio.sleep(0.3)
                                    except Exception:
                                        pass
                                    await asyncio.wait_for(
                                        element.click(), timeout=PER_CALL_TIMEOUT
                                    )
                                    logger.info(f"✅ Clicked in frame: {selector}")
                                    open_clicked = True
                                    break
                            except (Exception, asyncio.TimeoutError):
                                continue

                        if open_clicked:
                            break
                    except asyncio.TimeoutError:
                        logger.debug(f"⏱️ query_selector timed out for: {selector}")
                        continue
                    except Exception:
                        continue

                if open_clicked:
                    break

                await asyncio.sleep(1.5)

        if not open_clicked:
            logger.error(
                f"❌ Could not find {'addon name' if self.requires_addins_menu() else 'open button'} after retries"
            )
            return False

        # Step 3: Verify panel opened with retries
        logger.info(f"⏳ Step 3: Waiting for {addon_name} panel to open...")
        await asyncio.sleep(5)

        panel_opened = False
        for verify_attempt in range(3):
            if self.shutdown_event and self.shutdown_event.is_set():
                logger.info("Shutdown requested during panel open")
                return False
            if await self._verify_panel_opened():
                logger.info(f"✅ {addon_name} panel opened successfully")
                panel_opened = True
                break
            elif verify_attempt < 2:
                logger.info(
                    f"⚠️ Panel not yet open, waiting longer (attempt {verify_attempt + 1}/3)..."
                )
                await asyncio.sleep(3.0)
            else:
                logger.error(f"❌ {addon_name} panel did not open after retries")

        if not panel_opened:
            return False

        return True

    async def handle_initial_setup(self) -> bool:
        """
        Handle agent-specific initial setup after panel opens.

        Override this in subclasses if the agent needs first-time configuration.
        Default implementation does nothing.

        Returns:
            True if setup succeeded or not needed
        """
        return True

    async def verify_session_health(self) -> bool:
        """Verify AI agent session is healthy. Override for agent-specific checks."""
        return True

    async def upload_files(self, file_paths: str | list) -> bool:
        """Upload files into the add-in panel. Provider-specific — every
        concrete agent must override (ClaudeCore and ChatGPTCore do)."""
        raise NotImplementedError(
            f"{type(self).__name__} does not implement upload_files()"
        )

    async def _dump_upload_dom(self, label: str) -> None:
        """Best-effort diagnostic dump when a file upload path fails.

        Writes ``upload_dumps/upload_dump_<label>_<ts>.png`` (screenshot)
        and ``.json`` (visible role/aria/testid/text of elements in the
        main page, every iframe, and any Fluent ``ms-Layer`` overlay).
        Never raises — purely diagnostic so we can see what the add-in
        UI actually looks like when a selector misses.
        """
        try:
            import json
            from datetime import datetime
            from pathlib import Path

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = Path("upload_dumps")
            try:
                out_dir.mkdir(exist_ok=True)
            except Exception:
                pass
            stem = out_dir / f"upload_dump_{label}_{ts}"

            try:
                await self.page.screenshot(path=str(stem) + ".png", full_page=False)
            except Exception as e:
                logger.debug(f"dump screenshot failed: {e}")

            js = r"""
            () => {
                const roots = [document.body, ...document.querySelectorAll(
                    '.ms-Layer, .ms-Layer-content, [class*="Layer-content"]'
                )];
                const seen = new Set();
                const nodes = [];
                for (const root of roots) {
                    if (!root) continue;
                    try {
                        root.querySelectorAll(
                            '[role], [aria-label], [data-testid], button, [class*="menu-item" i]'
                        ).forEach(el => {
                            if (seen.has(el)) return;
                            seen.add(el);
                            const r = el.getBoundingClientRect();
                            if (r.width === 0 || r.height === 0) return;
                            const st = window.getComputedStyle(el);
                            if (st.visibility === 'hidden' || st.display === 'none') return;
                            const text = (el.textContent || '').trim().slice(0, 120);
                            nodes.push({
                                tag: el.tagName.toLowerCase(),
                                role: el.getAttribute('role') || '',
                                aria: el.getAttribute('aria-label') || '',
                                testid: el.getAttribute('data-testid') || '',
                                text: text,
                            });
                        });
                    } catch (e) {}
                }
                return nodes.slice(0, 400);
            }
            """

            dump = {"frames": []}
            try:
                main = await self.page.evaluate(js)
                dump["frames"].append(
                    {"frame": "__main__", "url": self.page.url, "nodes": main}
                )
            except Exception as e:
                logger.debug(f"dump main evaluate failed: {e}")

            for frame in self.page.frames:
                try:
                    nodes = await frame.evaluate(js)
                    dump["frames"].append(
                        {
                            "frame": frame.name or "",
                            "url": (frame.url or "")[:200],
                            "nodes": nodes,
                        }
                    )
                except Exception:
                    continue

            try:
                Path(str(stem) + ".json").write_text(json.dumps(dump, indent=2))
                logger.error(f"📸 Upload DOM dump written: {stem}.png + {stem}.json")
            except Exception as e:
                logger.debug(f"dump write failed: {e}")
        except Exception as e:
            logger.debug(f"_dump_upload_dom outer failure: {e}")

    async def _verify_panel_opened(self) -> bool:
        """
        Verify that the AI agent panel has opened by checking for input field.

        Returns:
            True if panel is open with input field visible
        """
        input_selectors = [
            'textarea[placeholder*="Ask"]',
            'textarea[data-testid="chat-input"]',
            'textarea[role="textbox"]',
            "textarea",
        ]

        # Check main page
        for selector in input_selectors:
            try:
                element = await self.page.query_selector(selector)
                if element and await element.is_visible():
                    logger.info(f"   ✅ Found input field: {selector}")
                    return True
            except Exception:
                pass

        # Check frames
        for frame in self.page.frames:
            for selector in input_selectors:
                try:
                    element = await frame.query_selector(selector)
                    if element and await element.is_visible():
                        logger.info(f"   ✅ Found input field in frame: {selector}")
                        return True
                except Exception:
                    pass

        logger.warning("   ⚠️ Could not find input field after clicking")
        return False

    async def _verify_text_in_field(self, element, expected_text: str) -> bool:
        """
        Verify that text is actually present in the input field.

        Args:
            element: The input element
            expected_text: The text we expect to find

        Returns:
            True if text matches (approximately), False otherwise
        """
        try:
            # Get the actual text content
            tag_name = await element.evaluate("el => el.tagName.toLowerCase()")

            if tag_name == "div":
                # For contenteditable divs, check textContent
                actual_text = await element.evaluate("el => el.textContent")
            else:
                # For textarea/input, check value
                actual_text = await element.input_value()

            # Normalize whitespace for comparison
            actual_normalized = " ".join(actual_text.split())
            expected_normalized = " ".join(expected_text.split())

            # Check if they match (allowing for minor differences)
            matches = actual_normalized.strip() == expected_normalized.strip()

            if matches:
                logger.debug(f"✅ Text verified in field: '{actual_text[:50]}...'")
            else:
                logger.warning(
                    f"❌ Text mismatch - Expected: '{expected_text[:50]}...', Got: '{actual_text[:50]}...'"
                )

            return matches

        except Exception as e:
            logger.debug(f"Could not verify text: {e}")
            return False

    async def _type_text_with_keyboard(
        self, element, text: str, clear_first: bool = True
    ) -> bool:
        """
        Type text normally - use fill() for textarea, insertText for contenteditable.
        NOTE: Element should already be focused/clicked before calling this.

        Args:
            element: The input element (should already be focused)
            text: Text to type
            clear_first: Whether to clear existing text first

        Returns:
            True if typing succeeded
        """
        try:
            # Check element type
            tag_name = await element.evaluate("el => el.tagName.toLowerCase()")
            is_contenteditable = await element.get_attribute("contenteditable")

            # Ensure element is focused (but don't click again - already clicked in submit_prompt)
            await element.focus()
            await asyncio.sleep(0.2)

            if tag_name == "textarea" or tag_name == "input":
                # For textarea/input: ALWAYS explicitly clear first, then fill
                if clear_first:
                    # Explicitly clear first to ensure clean state
                    await element.fill("")
                    await asyncio.sleep(0.1)
                    await element.fill(text)
                else:
                    # Append to existing value
                    current_value = await element.input_value()
                    await element.fill(current_value + text)
                await asyncio.sleep(0.2)
            elif tag_name == "div" and is_contenteditable:
                # For contenteditable div: clear and set text directly
                if clear_first:
                    # Explicitly clear first
                    await element.evaluate("el => el.textContent = ''")
                    await asyncio.sleep(0.1)
                    # Then set the new text
                    await element.evaluate(f"el => el.textContent = {repr(text)}")
                    await element.dispatch_event("input")
                else:
                    # Use insertText to append at cursor position
                    await element.evaluate(
                        f"""el => {{
                        const selection = window.getSelection();
                        if (selection.rangeCount > 0) {{
                            const range = selection.getRangeAt(0);
                            range.collapse(false); // Move to end
                            selection.removeAllRanges();
                            selection.addRange(range);
                        }}
                        document.execCommand('insertText', false, {repr(text)});
                    }}"""
                    )
                    await element.dispatch_event("input")
                await asyncio.sleep(0.2)
            else:
                # Fallback: use fill() if possible
                try:
                    if clear_first:
                        await element.fill("")
                        await asyncio.sleep(0.1)
                    await element.fill(text)
                except Exception:
                    # Last resort: keyboard (but avoid if possible)
                    logger.warning("⚠️ Using keyboard fallback - may cause issues")
                    if clear_first:
                        await self.page.keyboard.press(get_select_all_key())
                        await asyncio.sleep(0.1)
                        await self.page.keyboard.press("Delete")
                        await asyncio.sleep(0.1)
                    await element.fill(text)  # Still try fill() after clearing

            logger.debug(f"✅ Entered text: '{text[:50]}...'")
            return True

        except Exception as e:
            logger.debug(f"Text entry failed: {e}")
            return False

    async def submit_prompt(
        self, prompt: str, prompt_number: int, has_attachments: bool = False
    ) -> bool:
        """
        Submit a prompt to the AI agent with verification.

        Args:
            prompt: The prompt text to submit (empty string if only submitting files)
            prompt_number: The prompt number (for logging, 0 for file-only submission)
            has_attachments: Whether files are attached

        Returns:
            True if prompt was submitted successfully
        """
        try:
            if prompt:
                logger.info(f"📝 Submitting prompt #{prompt_number}: {prompt[:100]}...")
            else:
                logger.info("📝 Submitting files only (no prompt text)...")

            # Find the input field
            input_selectors = [
                'div[contenteditable="true"][role="textbox"][data-chat-input="true"]',
                'div[contenteditable="true"][role="textbox"]',
                'textarea[placeholder*="Ask"]',
                'textarea[data-testid="chat-input"]',
                'textarea[role="textbox"]',
            ]

            element = None

            # Find the input element
            for selector in input_selectors:
                # Try main page first
                try:
                    elem = await self.page.query_selector(selector)
                    if elem:
                        element = elem
                        logger.debug(f"Found input in main page: {selector}")
                        break
                except Exception:
                    pass

                # Try frames
                if not element:
                    for frame in self.page.frames:
                        try:
                            elem = await frame.query_selector(selector)
                            if elem:
                                element = elem
                                logger.debug(f"Found input in frame: {selector}")
                                break
                        except Exception:
                            continue

                if element:
                    break

            if not element:
                logger.error("❌ Could not find input field to submit prompt")
                return False

            # Click/focus the element
            try:
                await element.click()
                await asyncio.sleep(0.6)
            except Exception:
                try:
                    await element.focus()
                    await asyncio.sleep(0.3)
                except Exception as e:
                    logger.debug(f"Could not click/focus element: {e}")

            # Special case: Empty prompt with attachments (file-only submission)
            # Just submit without entering any text
            if has_attachments and not prompt:
                logger.info("📎 Files already attached, submitting without text entry")
                # Skip text entry, go straight to submission
                text_entered = True
            else:
                # Normal prompt submission with text entry
                # Strategy: Use ONE method per attempt to avoid double-typing
                # For normal prompts, try fill() first attempt, then _type_text_with_keyboard on retries
                max_retries = 3
                text_entered = False

                for attempt in range(max_retries):
                    if attempt > 0:
                        logger.info(
                            f"🔄 Retry {attempt}/{max_retries} - entering text again..."
                        )

                    # Choose method: ONE method per attempt (no fallback in same attempt)
                    # - First attempt: try fill()
                    # - Retries: use _type_text_with_keyboard
                    use_fill_method = attempt == 0

                    if use_fill_method:
                        # Try fill() only on first attempt
                        try:
                            # Explicitly clear first, then fill
                            await element.fill("")
                            await asyncio.sleep(0.1)
                            await element.fill(prompt)
                            logger.debug("Used fill() method")
                        except Exception as e:
                            logger.debug(
                                f"fill() failed: {e}, will retry next attempt with _type_text_with_keyboard"
                            )
                            # Skip verification and continue to next attempt
                            if attempt < max_retries - 1:
                                await asyncio.sleep(1)
                                continue
                            # Last attempt failed, fall through to use _type_text_with_keyboard
                            use_fill_method = False

                    if not use_fill_method:
                        # Use _type_text_with_keyboard (for retries or if fill() failed on last attempt)
                        logger.info("⌨️  Using _type_text_with_keyboard method")
                        await self._type_text_with_keyboard(
                            element, prompt, clear_first=True
                        )

                    # Wait a bit for the UI to update
                    await asyncio.sleep(0.5)

                    # Verify the text is actually there
                    try:
                        tag_name = await element.evaluate(
                            "el => el.tagName.toLowerCase()"
                        )
                        if tag_name == "textarea" or tag_name == "input":
                            actual_text = await element.input_value()
                        else:
                            actual_text = await element.evaluate("el => el.textContent")

                        # Check if prompt text is present in the field
                        actual_lower = actual_text.lower().strip()
                        prompt_lower = prompt.lower().strip()

                        # Check if prompt text appears in the actual text
                        # Use first 30 chars of prompt to avoid issues with quotes/special chars
                        prompt_key = (
                            prompt_lower[:30]
                            if len(prompt_lower) > 30
                            else prompt_lower
                        )

                        if prompt_key in actual_lower:
                            text_entered = True
                            logger.info(
                                "✅ Text verified in input field (found prompt text)"
                            )
                            break
                        else:
                            logger.debug(
                                f"Text check: found '{actual_text[:80]}...', looking for '{prompt[:80]}...'"
                            )
                    except Exception as e:
                        logger.debug(
                            f"Verification check failed: {e}, assuming text entered"
                        )
                        text_entered = True  # Assume success if we can't verify
                        break

                    if attempt < max_retries - 1:
                        await asyncio.sleep(1)

                if not text_entered:
                    logger.error(
                        "❌ Prompt text never verified in the input field after "
                        "3 attempts — refusing to send (an empty send would "
                        "record this prompt as delivered when it was not)"
                    )
                    return False

            # Proceed to submit even if verification is uncertain
            await asyncio.sleep(0.5)

            # Click send button - try multiple methods
            send_clicked = False

            # Method 1: Find send button by common selectors
            send_selectors = [
                'button[data-testid="send-button"]',
                'button[aria-label*="Send" i]',
                'button:has-text("Send")',
                'button[type="submit"]',
                'button[aria-label="Send"]',
            ]

            # Try main page first
            for selector in send_selectors:
                try:
                    send_btn = await self.page.query_selector(selector)
                    if send_btn and await send_btn.is_visible():
                        await send_btn.click()
                        send_clicked = True
                        logger.info(f"✅ Clicked send button: {selector}")
                        break
                except Exception:
                    continue

            # Try frames if not found on main page
            if not send_clicked:
                for frame in self.page.frames:
                    for selector in send_selectors:
                        try:
                            send_btn = await frame.query_selector(selector)
                            if send_btn and await send_btn.is_visible():
                                await send_btn.click()
                                send_clicked = True
                                logger.info(
                                    f"✅ Clicked send button in frame: {selector}"
                                )
                                break
                        except Exception:
                            continue
                    if send_clicked:
                        break

            # Method 2: Press Enter if button not found
            if not send_clicked:
                logger.info("⌨️  Using Enter key to submit")
                await element.focus()
                await asyncio.sleep(0.2)
                await self.page.keyboard.press("Enter")
                send_clicked = True

            logger.info(f"✅ Prompt #{prompt_number} submitted")
            return True

        except Exception as e:
            logger.error(f"❌ Error submitting prompt: {e}")
            return False

    async def get_button_count(self) -> dict:
        """
        Get count of upvote/downvote buttons.

        Returns:
            Dict with 'upvote' and 'downvote' counts
        """
        try:
            upvote_count = 0
            downvote_count = 0

            # Check main page
            upvote_buttons = await self.page.query_selector_all(
                '[data-testid="upvote-button"]'
            )
            downvote_buttons = await self.page.query_selector_all(
                '[data-testid="downvote-button"]'
            )

            upvote_count += len(upvote_buttons)
            downvote_count += len(downvote_buttons)

            # Check frames
            for frame in self.page.frames:
                try:
                    frame_upvotes = await frame.query_selector_all(
                        '[data-testid="upvote-button"]'
                    )
                    frame_downvotes = await frame.query_selector_all(
                        '[data-testid="downvote-button"]'
                    )
                    upvote_count += len(frame_upvotes)
                    downvote_count += len(frame_downvotes)
                except Exception:
                    continue

            return {"upvote": upvote_count, "downvote": downvote_count}

        except Exception as e:
            logger.debug(f"Error counting buttons: {e}")
            return {"upvote": 0, "downvote": 0}

    async def wait_for_completion(
        self, prompt_number: int, initial_counts: dict = None
    ) -> bool:
        """
        Wait for prompt to complete processing.

        Args:
            prompt_number: Prompt number (for logging)
            initial_counts: Initial button counts before submission

        Returns:
            True if completion detected
        """
        try:
            if initial_counts is None:
                initial_counts = {"upvote": 0, "downvote": 0}

            agent_config = self.get_config_section()
            max_wait = agent_config.get("max_wait_per_prompt_seconds", 900)
            log_interval = agent_config.get("check_interval_seconds", 60)

            logger.info(
                f"⏳ Waiting for prompt #{prompt_number} to complete (max {max_wait}s)..."
            )

            start_time = time.monotonic()
            last_log = 0.0
            check_frequency = 2  # Check every 2 seconds
            saw_running = False  # gate for the idle fallback below

            while (time.monotonic() - start_time) < max_wait:
                # Check for shutdown signal
                if self.shutdown_event and self.shutdown_event.is_set():
                    logger.warning(
                        f"⏱️  Shutdown signal received during prompt #{prompt_number} - exiting gracefully"
                    )
                    return False

                await asyncio.sleep(check_frequency)
                elapsed = time.monotonic() - start_time

                # Log status at configured intervals
                if elapsed - last_log >= log_interval:
                    state = await self.get_state()
                    logger.info(f"   [{int(elapsed)}s] State: {state.value}")
                    last_log = elapsed

                # Check state
                state = await self.get_state()

                if state == AIAgentState.RUNNING:
                    saw_running = True
                    continue

                # Check for new buttons (primary completion indicator)
                current_counts = await self.get_button_count()
                if (
                    current_counts["upvote"] > initial_counts["upvote"]
                    or current_counts["downvote"] > initial_counts["downvote"]
                ):
                    logger.info(
                        f"✅ Prompt #{prompt_number} completed! (New buttons detected after {int(elapsed)}s)"
                    )
                    return True

                # If idle/ready but no new buttons, might be done — but only
                # once the agent was actually observed RUNNING. Right after
                # submit the cleared input also reads as idle, and accepting
                # that would declare completion before processing started.
                if not saw_running:
                    continue
                if state in [AIAgentState.IDLE, AIAgentState.READY_TO_SEND]:
                    await asyncio.sleep(2)
                    state_recheck = await self.get_state()
                    if state_recheck in [AIAgentState.IDLE, AIAgentState.READY_TO_SEND]:
                        logger.info(
                            f"✅ Prompt #{prompt_number} completed! (Ready state detected after {int(elapsed)}s)"
                        )
                        return True

            logger.error(
                f"❌ Timeout waiting for prompt #{prompt_number} after {max_wait}s"
            )
            return False

        except Exception as e:
            logger.error(f"❌ Error waiting for completion: {e}")
            return False

    async def process_all_prompts(self, files_to_upload: list = None) -> bool:
        """
        Process all prompts sequentially.

        Args:
            files_to_upload: Optional list of files to attach to the FIRST
                prompt. They are never sent as a turn of their own — an
                attachment-only send lets an agentic add-in start working the
                open workbook before it has been told what the task is.

        Returns:
            True if all prompts processed successfully
        """
        try:
            prompts = self.config.get("prompts", [])
            if not prompts:
                logger.error("❌ No prompts found in config")
                return False

            if isinstance(prompts, str):
                prompts = [prompts]

            # Step 1: If files exist, attach them to the composer now. They are
            # sent WITH the first prompt, never as a text-less turn of their own.
            # 2026-09-10 (rung 3, tasks 36/51): an empty send with only
            # House_Standards_v1.md attached made both add-ins start building
            # the model from the workbook's Instructions sheet before the task
            # prompt had been delivered — Claude was already editing cells and
            # dismissing permission dialogs three minutes in. The attachments
            # must ride on prompt #1 so the model's first turn IS the prompt.
            attach_to_first_prompt = False
            if files_to_upload:
                logger.info(f"\n{'='*80}")
                logger.info("📤 FILE ATTACH (sent with prompt #1)")
                logger.info(f"{'='*80}")
                logger.info(
                    f"📤 Attaching {len(files_to_upload)} file(s) to the first prompt..."
                )

                upload_success = await self.upload_files(files_to_upload)
                if not upload_success:
                    logger.error("❌ File upload failed")
                    return False
                logger.info("✅ Files attached; they go out with prompt #1")
                await asyncio.sleep(3)  # Give time for files to be processed
                attach_to_first_prompt = True

            # Step 2: Send the prompts in order; #1 carries the attachments.
            logger.info(f"🚀 Processing {len(prompts)} prompt(s)")

            for i, prompt in enumerate(prompts, 1):
                # Check for shutdown signal
                if self.shutdown_event and self.shutdown_event.is_set():
                    logger.warning(
                        f"⏱️  Shutdown signal received before prompt #{i} - exiting gracefully"
                    )
                    return False

                logger.info(f"\n{'='*80}")
                logger.info(f"📝 PROMPT {i}/{len(prompts)}")
                logger.info(f"{'='*80}")

                # Start prompt logging
                if self.completion_logger:
                    self.completion_logger.start_prompt(prompt)

                # Get initial button counts
                initial_counts = await self.get_button_count()

                # Prompt #1 carries the attachments (if any); later prompts none.
                has_attachments = attach_to_first_prompt and i == 1
                if not await self.submit_prompt(
                    prompt, i, has_attachments=has_attachments
                ):
                    logger.error(f"❌ Failed to submit prompt {i}")
                    if self.completion_logger:
                        self.completion_logger.end_prompt(success=False)
                    return False

                # Wait for completion
                if not await self.wait_for_completion(i, initial_counts):
                    logger.error(f"❌ Prompt {i} did not complete")
                    if self.completion_logger:
                        self.completion_logger.end_prompt(success=False)
                    return False

                logger.info(f"✅ Prompt {i} completed successfully")

                # End prompt logging
                if self.completion_logger:
                    self.completion_logger.end_prompt(success=True)

                await asyncio.sleep(3)

            logger.info(f"\n🎉 All {len(prompts)} prompts completed!")
            return True

        except Exception as e:
            logger.error(f"❌ Error processing prompts: {e}")
            if self.completion_logger and self.completion_logger.current_prompt:
                self.completion_logger.end_prompt(success=False)
            return False
