"""
Claude by Anthropic core interaction logic.
"""

import asyncio
import logging
import re
import time
from pathlib import Path

from .ai_agent_base import AIAgentCore, is_host_frame as _is_host_frame

logger = logging.getLogger(__name__)


def _label_matches(target_lower: str, shown_lower: str) -> bool:
    """Whole-token model-label match.

    Exact match, or `target` appearing in `shown` bounded by characters
    that cannot continue a model name (letters, digits, '.'). A plain
    substring test let "Fable 5" pass on a "Fable 5.1" button (and would
    let "Opus 4" pass on "Opus 4.8"); the gui pipeline fixed the same
    bug in 7ac76d6.
    """
    target_lower = target_lower.strip().lower()
    shown_lower = shown_lower.strip().lower()
    if not target_lower:
        return False
    if target_lower == shown_lower:
        return True
    return re.search(
        rf"(?<![\w.]){re.escape(target_lower)}(?![\w.])", shown_lower
    ) is not None


class ClaudeCore(AIAgentCore):
    """Claude by Anthropic-specific implementation."""

    def __init__(self, page, config, shutdown_event, completion_logger):
        super().__init__(page, config, shutdown_event, completion_logger)
        self._claude_frame = None
        self._accepted_all_edits = False
        # Response count sampled in submit_prompt BEFORE the send click, so
        # wait_for_completion's "new response appeared" baseline predates
        # the response it is looking for.
        self._pre_send_response_count: int | None = None

    def get_agent_type(self) -> str:
        return "claude_excel_agent"

    def get_addon_name(self) -> str:
        return "Claude by Anthropic"

    def get_open_button_text(self) -> str:
        return "Open Claude"

    def requires_addins_menu(self) -> bool:
        return True

    def get_ribbon_launcher_name(self) -> str:
        # The installed add-in's own Home-ribbon button / overflow menuitem
        # (probed 2026-09-10: "Claude", data-unique-id Ribbon-AddinControlN).
        return "Claude"

    # Selectors for the Claude chat input, ordered by likelihood.
    _INPUT_SELECTORS = [
        'div[contenteditable="true"][data-testid="chat-input"]',
        'textarea[data-testid="chat-input"]',
        'div[contenteditable="true"][role="textbox"]',
    ]

    async def _verify_panel_opened(self) -> bool:
        """Claude-specific: poll for chat-input with configurable boot timeout."""
        boot_timeout = self.config.get("panel_boot_timeout_seconds", 20)
        poll_interval = 2
        elapsed = 0

        while elapsed < boot_timeout:
            if self.shutdown_event and self.shutdown_event.is_set():
                logger.info("   Shutdown requested during panel verification")
                return False
            self._claude_frame = None  # fresh lookup each time
            frame = await self._get_claude_frame()
            if frame:
                logger.info(
                    f"   ✅ Panel verified: Claude chat-input found after {elapsed}s"
                )
                return True
            if elapsed > 0 and elapsed % 6 == 0:
                logger.info(
                    f"   ⏳ [{elapsed}s] Waiting for Claude app to initialize..."
                )
            await asyncio.sleep(poll_interval)
            elapsed += poll_interval

        # Last-ditch: dump what frames and elements exist for debugging
        logger.warning(f"   ⚠️ Claude chat-input not found after {boot_timeout}s")
        try:
            frame_names = []
            for f in self.page.frames:
                try:
                    url = f.url[:80] if f.url else "(no url)"
                    textareas = await f.query_selector_all("textarea")
                    editables = await f.query_selector_all('[contenteditable="true"]')
                    frame_names.append(
                        f"{f.name or '(anon)'}({url}) "
                        f"textareas={len(textareas)} editables={len(editables)}"
                    )
                except Exception:
                    frame_names.append(f"{f.name or '(anon)'} (error)")
            logger.warning(f"   Frames: {frame_names}")
        except Exception:
            pass
        return False

    # URL fragments that identify the Claude add-in iframe.
    _FRAME_URL_HINTS = ["pivot.claude.ai", "claude.ai"]

    async def _get_claude_frame(self):
        """Find and cache the Claude add-in iframe.

        Strategy: try frames matching known URLs first, then fall back to
        any frame that contains a chat-input element.
        """
        if self._claude_frame:
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await self._claude_frame.query_selector(selector)
                    if el:
                        return self._claude_frame
            except Exception:
                pass
            self._claude_frame = None

        # Pass 1: frames whose URL matches known hints (fast path)
        for f in self.page.frames:
            url = f.url or ""
            if not any(hint in url for hint in self._FRAME_URL_HINTS):
                continue
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await f.query_selector(selector)
                    if el:
                        self._claude_frame = f
                        return f
            except Exception:
                continue

        # Pass 2: brute-force all frames (URL changed or new domain).
        # Host (Excel/OneDrive) frames are excluded and the input must be
        # visible — binding to Excel's own contenteditable would type the
        # prompt into the workbook under grading.
        for f in self.page.frames:
            if _is_host_frame(f):
                continue
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await f.query_selector(selector)
                    if el and await el.is_visible():
                        self._claude_frame = f
                        return f
            except Exception:
                continue
        return None

    async def _find_input_element(self, frame=None):
        """Find the chat input element in the given frame (or cached frame).

        Returns the element handle, or None.
        """
        frame = frame or await self._get_claude_frame()
        if not frame:
            return None
        for selector in self._INPUT_SELECTORS:
            try:
                el = await frame.query_selector(selector)
                if el:
                    return el
            except Exception:
                continue
        return None

    async def _find_web_search_button(self):
        """Find the 'Search the web' <button> in the Claude add-in frame.

        Scoped to the Claude iframe only (not every frame in the page).
        The Plus popover is plain-styled — no ``role="menu"`` or
        ``data-state`` — so the stablest anchor is the visible label.
        Uses ``has-text`` (case-insensitive substring) which is fine here
        because the popover only contains a handful of short-labeled
        items and "Search the web" doesn't collide with any of them.
        """
        frame = await self._get_claude_frame()
        if not frame:
            return None
        try:
            btn = await frame.query_selector('button:has-text("Search the web")')
            if btn and await btn.is_visible():
                return btn
        except Exception:
            pass
        return None

    async def _disable_web_search(self):
        """Disable 'Search the web' in the Claude add-in if it's enabled.

        Opens the Plus (+) menu and reads the toggle state from the
        "Search the web" button using two independent signals. ON when
        either holds:
          (a) className contains ``accent-secondary``. The add-in
              currently ships ``text-accent-secondary-100`` when active,
              and the substring match is forward-compatible with any
              ``*accent-secondary*`` successor.
          (b) the button contains 2+ ``<svg>`` descendants. OFF renders
              just the leading icon; ON adds a trailing checkmark SVG.
        Either signal alone is fragile (class names are Tailwind-ish
        tokens; icon structure can change), but both breaking at once
        would require a simultaneous visual + semantic redesign — much
        less likely than either one in isolation.
        """
        try:
            logger.info("🔧 Checking 'Search the web' toggle...")

            # Open the Plus menu.
            _frame, plus_btn = await self._find_plus_button()
            if not plus_btn:
                logger.debug("Could not find Plus button for web search check")
                return
            await plus_btn.click()

            # Wait for the "Search the web" button to appear. The popover
            # isn't Radix so there's no data-state=open to wait on; the
            # button's visibility is itself the open signal. ~2s budget.
            search_button = None
            for _ in range(20):
                await asyncio.sleep(0.1)
                search_button = await self._find_web_search_button()
                if search_button:
                    break

            if not search_button:
                logger.debug(
                    "'Search the web' button did not appear after opening "
                    "Plus menu — skipping"
                )
                await self.page.keyboard.press("Escape")
                await asyncio.sleep(0.3)
                return

            # Read ON/OFF via the layered signal described above.
            is_enabled = await search_button.evaluate(
                """el => {
                    const cls = el.className || '';
                    if (cls.includes('accent-secondary')) return true;
                    // ON state renders a checkmark SVG alongside the icon.
                    if (el.querySelectorAll('svg').length >= 2) return true;
                    return false;
                }"""
            )

            if is_enabled:
                logger.info("🌐 'Search the web' is ENABLED — clicking to disable...")
                await search_button.click()
                await asyncio.sleep(0.5)
                logger.info("✅ Disabled 'Search the web'")
            else:
                logger.info("✅ 'Search the web' is already disabled")

            await self.page.keyboard.press("Escape")
            await asyncio.sleep(0.3)

        except Exception as e:
            logger.debug(f"Could not check web search toggle: {e}")
            try:
                await self.page.keyboard.press("Escape")
                await asyncio.sleep(0.3)
            except Exception:
                pass

    async def _close_panel(self) -> bool:
        """Close the Claude add-in panel to allow fresh reopen."""
        try:
            # Try Office task pane close button
            close_selectors = [
                'button[aria-label="Close task pane"]',
                'button[aria-label="Close"]',
                'button[title="Close"]',
                '[data-automation-id="PanelCloseButton"]',
            ]
            for selector in close_selectors:
                for ctx in [self.page] + self.page.frames:
                    try:
                        el = await ctx.query_selector(selector)
                        if el and await el.is_visible():
                            await el.click()
                            logger.info(f"✅ Closed panel via: {selector}")
                            self._claude_frame = None
                            self._accepted_all_edits = False
                            await asyncio.sleep(3)
                            return True
                    except Exception:
                        continue

            # Fallback: toggle via Add-ins ribbon
            logger.info("Close button not found, toggling via Add-ins...")
            for ctx in [self.page] + self.page.frames:
                # Skip detached frames — clicking against one leaks an
                # uncaught Playwright Future that asyncio warns about.
                if hasattr(ctx, "is_detached") and ctx.is_detached():
                    continue
                try:
                    await ctx.click('text="Add-ins"', timeout=3000)
                    await asyncio.sleep(2)
                    break
                except Exception:
                    continue
            # Click add-in tile in the My Add-ins popup (uses JS evaluate
            # to handle name mismatches like "Claude by Anthropic for Excel")
            if await self._click_addon_in_layer(self.get_addon_name()):
                self._claude_frame = None
                self._accepted_all_edits = False
                await asyncio.sleep(3)
                return True

            return False
        except Exception as e:
            logger.debug(f"Error closing panel: {e}")
            return False

    async def verify_session_health(self) -> bool:
        """Verify Claude iframe is responsive (not just present)."""
        try:
            self._claude_frame = None
            frame = await self._get_claude_frame()
            if not frame:
                logger.warning("❌ Health check: Claude frame not found")
                return False

            textarea = await self._find_input_element(frame)
            if not textarea or not await textarea.is_visible():
                logger.warning("❌ Health check: input not visible")
                return False

            # Test interactivity
            await textarea.focus()
            await asyncio.sleep(0.3)

            logger.info("✅ Session health check passed")
            return True
        except Exception as e:
            logger.warning(f"❌ Health check failed: {e}")
            return False

    async def _select_model(self):
        """Select the configured Claude model via the add-in model dropdown.

        Reads ``claude_excel_agent.ui_model_label`` from config (injected
        by the runner from the agent-identity registry) and matches it
        (case-insensitive) against the selector button and dropdown menu
        items. The value is whatever string the Claude add-in's UI shows
        — e.g. "Opus 4.6", "Sonnet 4.6". No whitelist, so new models are
        a registry entry away, never a code change.

        Selector strategy (most to least stable):
          * button:    ``button[aria-label="Model selector"]``
          * menu:      ``[role="menu"][data-state="open"]`` linked back
                       to the button via ``aria-labelledby="<button id>"``
                       — catches silent click failures and avoids picking
                       up an unrelated open menu.
          * menu item: ``[role="menuitem"]`` inside that menu.

        Returns True when no model is pinned, or when the selector button
        verifiably shows the pinned label afterwards; False otherwise —
        the caller treats False as a setup failure (infra, unrecorded),
        never as "run on whatever model happens to be active".
        """
        target = self.config.get("claude_excel_agent", {}).get("ui_model_label")
        if not target:
            return True

        target_lower = str(target).strip().lower()
        if not target_lower:
            return True

        frame = None
        try:
            logger.info("Selecting Claude Excel model: %s", target)
            frame = await self._get_claude_frame()
            if not frame:
                logger.error("Claude frame not found — cannot pin the model")
                return False

            # Anchor the button by its stable aria-label.
            model_btn = await frame.query_selector(
                'button[aria-label="Model selector"]'
            )
            if not model_btn or not await model_btn.is_visible():
                logger.error(
                    "Model selector button not found (aria-label='Model "
                    "selector') — cannot pin the model; aborting setup"
                )
                return False

            # Already on target? Button's title attribute is the exact
            # model string ("Opus 4.7"); text_content as a fallback.
            current = (await model_btn.get_attribute("title") or "").strip().lower()
            if not current:
                current = (await model_btn.text_content() or "").strip().lower()
            if _label_matches(target_lower, current):
                logger.info("Model '%s' already selected", target)
                return True

            # Remember the button's id so we can lock onto *its* popover
            # (Radix links each menu's aria-labelledby to its trigger).
            btn_id = await model_btn.get_attribute("id")

            await model_btn.click()

            # Wait for the popover to open (~2s budget). Prefer the menu
            # whose aria-labelledby points back to our trigger; fall back
            # to any open menu.
            menu = None
            for _ in range(20):
                await asyncio.sleep(0.1)
                if btn_id:
                    menu = await frame.query_selector(
                        '[role="menu"][data-state="open"]'
                        f'[aria-labelledby="{btn_id}"]'
                    )
                    if menu:
                        break
                menu = await frame.query_selector('[role="menu"][data-state="open"]')
                if menu:
                    break
            if not menu:
                logger.error(
                    "Model dropdown did not open (no [role=menu]"
                    "[data-state=open] found) — cannot verify the pinned model"
                )
                await self._dismiss_dropdown(frame)
                return False

            # Scan items inside the opened menu only. Exact match on
            # text_content first, then substring.
            items = await menu.query_selector_all('[role="menuitem"]')
            try:
                dump = []
                for it in items:
                    dump.append(
                        f"{(await it.text_content() or '').strip()!r}"
                        f"[disabled={await it.get_attribute('aria-disabled')}"
                        f"/{await it.get_attribute('data-disabled')}"
                        f" state={await it.get_attribute('data-state')}]"
                    )
                logger.info("Model menu items: %s", "; ".join(dump))
            except Exception as e:
                logger.info("Model menu dump failed: %s", e)
            for item in items:
                try:
                    if not await item.is_visible():
                        continue
                    text = (await item.text_content() or "").strip().lower()
                    if text == target_lower:
                        await item.click()
                        logger.info("Clicked model item '%s'", target)
                        await asyncio.sleep(0.5)
                        return await self._verify_or_retry_model(frame, target, target_lower)
                except Exception:
                    continue
            for item in items:
                try:
                    if not await item.is_visible():
                        continue
                    text = (await item.text_content() or "").strip().lower()
                    if _label_matches(target_lower, text):
                        await item.click()
                        logger.info("Clicked model item '%s' (whole-token)", target)
                        await asyncio.sleep(0.5)
                        return await self._verify_or_retry_model(frame, target, target_lower)
                except Exception:
                    continue

            logger.error(
                "Model '%s' not found in the add-in dropdown — aborting "
                "setup rather than running on an unverified model",
                target,
            )
            await self._dismiss_dropdown(frame)
            return False

        except Exception as e:
            logger.error("Could not select model '%s': %s", target, e)
            if frame:
                await self._dismiss_dropdown(frame)
            return False

    async def _verify_or_retry_model(self, frame, target: str, target_lower: str) -> bool:
        """Verify the pin; on failure log what the panel shows and retry the
        selection once via keyboard (hover the item, press Enter), which is
        how a Radix menu item selects when a synthetic click is ignored."""
        if await self._verify_selected_model(frame, target_lower):
            return True
        # First-time selection of some models (Fable 5.1) opens an
        # "About this model" modal with Cancel / Got it; the switch only
        # commits after "Got it". Acknowledge it and re-verify.
        try:
            ack = await frame.query_selector('button:has-text("Got it")')
            if ack and await ack.is_visible():
                await ack.click()
                logger.info("Acknowledged the 'About this model' modal (Got it)")
                await asyncio.sleep(0.5)
                if await self._verify_selected_model(frame, target_lower):
                    return True
        except Exception as e:
            logger.warning("Model-notice acknowledgement errored: %s", e)
        try:
            body = await frame.evaluate("document.body.innerText")
            logger.warning("Panel text after failed pin (600 chars): %r", (body or "")[:600])
            still_open = await frame.query_selector('[role="menu"][data-state="open"]')
            logger.warning("Model menu still open after click: %s", bool(still_open))
            if still_open:
                await self._dismiss_dropdown(frame)
                await asyncio.sleep(0.3)
            model_btn = await frame.query_selector('button[aria-label="Model selector"]')
            if not model_btn:
                return False
            await model_btn.click()
            menu = None
            for _ in range(20):
                await asyncio.sleep(0.1)
                menu = await frame.query_selector('[role="menu"][data-state="open"]')
                if menu:
                    break
            if not menu:
                logger.error("Model dropdown did not reopen for the keyboard retry")
                return False
            for item in await menu.query_selector_all('[role="menuitem"]'):
                text = (await item.text_content() or "").strip().lower()
                if text == target_lower or _label_matches(target_lower, text):
                    await item.hover()
                    await asyncio.sleep(0.2)
                    await item.focus()
                    await frame.page.keyboard.press("Enter")
                    logger.info("Keyboard-selected model item '%s'", target)
                    await asyncio.sleep(0.5)
                    return await self._verify_selected_model(frame, target_lower)
            logger.error("Model '%s' not found on the keyboard retry", target)
            await self._dismiss_dropdown(frame)
            return False
        except Exception as e:
            logger.error("Keyboard model retry errored: %s", e)
            return False

    async def _verify_selected_model(self, frame, target_lower: str) -> bool:
        """Re-read the model selector and confirm it shows the pinned label.

        The selector's title re-renders asynchronously after a menu click, so
        poll for up to ~4s before declaring the pin unverified.
        """
        try:
            shown = ""
            for _ in range(8):
                btn = await frame.query_selector('button[aria-label="Model selector"]')
                if not btn:
                    logger.error("Model selector button vanished during verification")
                    return False
                shown = (await btn.get_attribute("title") or "").strip().lower()
                if not shown:
                    shown = (await btn.text_content() or "").strip().lower()
                if _label_matches(target_lower, shown):
                    logger.info("✅ Model verified: selector shows %r", shown)
                    return True
                await asyncio.sleep(0.5)
            logger.error(
                "Model verification FAILED: selector shows %r, pinned %r — "
                "aborting setup",
                shown,
                target_lower,
            )
            return False
        except Exception as e:
            logger.error("Model verification errored: %s", e)
            return False

    async def _dismiss_dropdown(self, frame):
        """Press Escape to close any open popover; swallow any failure."""
        try:
            await frame.press("body", "Escape")
            await asyncio.sleep(0.3)
        except Exception:
            try:
                await self.page.keyboard.press("Escape")
            except Exception:
                pass

    async def _ensure_extended_thinking(self, frame) -> bool:
        """Pin the add-in's "Toggle extended thinking" button ON.

        The Claude add-in exposes no effort tiers (probed 2026-09-10);
        extended thinking is its only reasoning control, so the cohort's
        "highest effort" means this toggle pressed. Verified by re-reading
        aria-pressed after any click; a toggle that cannot be read or
        turned on is a setup failure (infra, unrecorded), never a run at
        an unknown thinking setting. Returns True when pressed.
        """
        selector = (
            '[data-testid="extended-thinking-toggle"], '
            'button[aria-label="Toggle extended thinking"]'
        )
        try:
            toggle = await frame.query_selector(selector)
            if not toggle or not await toggle.is_visible():
                logger.error(
                    "Extended-thinking toggle not found — cannot pin "
                    "thinking; aborting setup"
                )
                return False
            pressed = (await toggle.get_attribute("aria-pressed") or "").lower()
            if pressed == "true":
                logger.info("✅ Extended thinking already ON")
                return True
            await toggle.click()
            await asyncio.sleep(0.6)
            toggle = await frame.query_selector(selector)
            pressed = (
                (await toggle.get_attribute("aria-pressed") or "").lower()
                if toggle else ""
            )
            if pressed == "true":
                logger.info("✅ Extended thinking turned ON and verified")
                return True
            logger.error(
                "Extended thinking verification FAILED (aria-pressed=%r) — "
                "aborting setup",
                pressed,
            )
            return False
        except Exception as e:
            logger.error("Could not pin extended thinking: %s", e)
            return False

    async def handle_initial_setup(self) -> bool:
        """Select model, pin extended thinking, click 'Accept all edits',
        disable web search."""
        if self._accepted_all_edits:
            return True

        try:
            logger.info("🔍 Setting up 'Accept all edits' mode...")
            frame = await self._get_claude_frame()
            if not frame:
                logger.warning("⚠️ Could not find Claude frame for setup")
                return False

            # Select + verify the pinned model. A miss is a setup failure
            # (engine → PANEL_FAILED → infra retry, nothing recorded).
            if not await self._select_model():
                return False
            if not await self._ensure_extended_thinking(frame):
                return False

            # Disable web search before anything else
            await self._disable_web_search()

            # Dismiss any stale menus/dropdowns before looking for buttons
            await frame.press("body", "Escape")
            await asyncio.sleep(1.5)

            # Find the permission-mode trigger button.
            #
            # Selector priority (most to least stable):
            #   1. aria-label="Toggle permission mode" — stable a11y label
            #   2. button:has-text("Ask before edits")  — current-mode text
            #   3. button:has-text("Accept all edits")  — current-mode text
            # We used to primarily hunt for `span.font-small` + text contains
            # "edits"; that relied on both a hashed Tailwind class and loose
            # substring matching and is now a last-resort fallback only.
            mode_btn = await frame.query_selector(
                'button[aria-label="Toggle permission mode"]'
            )
            if not mode_btn:
                mode_btn = await frame.query_selector(
                    'button:has-text("Ask before edits")'
                )
            if not mode_btn:
                mode_btn = await frame.query_selector(
                    'button:has-text("Accept all edits")'
                )

            if not mode_btn or not await mode_btn.is_visible():
                logger.info(
                    "ℹ️ Permission mode button not found — assuming it's" " already set"
                )
                self._accepted_all_edits = True
                return True

            # Quick early-exit: if the trigger button's own text already says
            # "Accept all edits", we're there. Saves opening the dropdown.
            try:
                btn_text = (await mode_btn.text_content() or "").strip().lower()
                if "accept all" in btn_text:
                    logger.info("✅ Already in 'Accept all edits' mode")
                    self._accepted_all_edits = True
                    return True
            except Exception:
                pass

            # Open the permission-mode dropdown.
            await mode_btn.click()
            logger.info("✅ Clicked permission mode button (opening dropdown)")

            # The permission-mode popover is plain-styled (no Radix
            # data-state), so we poll for the "Accept all edits" option
            # to appear as our open-signal. ~2s budget.
            accept_btn = None
            for _ in range(20):
                await asyncio.sleep(0.1)
                accept_btn = await frame.query_selector(
                    'button:has-text("Accept all edits")'
                )
                if accept_btn and await accept_btn.is_visible():
                    break

            if not accept_btn:
                logger.warning("⚠️ 'Accept all edits' option did not appear in dropdown")
                await frame.press("body", "Escape")
                await asyncio.sleep(0.3)
                return True  # non-fatal, let task continue

            # Is 'Accept all edits' already selected? Selected state adds
            # a trailing checkmark SVG wrapped in a div with class
            # `text-accent-secondary-100` (Claude's "active" accent token).
            # Layered signal — any of:
            #   (a) button's own class contains 'accent-secondary'
            #   (b) a descendant has class containing 'accent-secondary'
            #   (c) button has 2+ SVG descendants (icon + checkmark).
            # Any one alone is fragile; all three breaking at once would
            # require a coordinated visual + semantic redesign.
            already_selected = await accept_btn.evaluate(
                """el => {
                    if ((el.className || '').includes('accent-secondary')) return true;
                    if (el.querySelector('[class*="accent-secondary"]')) return true;
                    if (el.querySelectorAll('svg').length >= 2) return true;
                    return false;
                }"""
            )

            if already_selected:
                logger.info("✅ 'Accept all edits' already selected")
                self._accepted_all_edits = True
            else:
                await accept_btn.click()
                logger.info("✅ Clicked 'Accept all edits'")
                self._accepted_all_edits = True
                await asyncio.sleep(0.5)

            # Always dismiss the dropdown after we're done
            await frame.press("body", "Escape")
            await asyncio.sleep(0.5)

            return True

        except Exception as e:
            logger.warning(f"⚠️ Error in initial setup: {e}")
            # Try to close any open menus
            try:
                frame = await self._get_claude_frame()
                if frame:
                    await frame.press("body", "Escape")
                    await asyncio.sleep(0.3)
            except Exception:
                pass
            return False

    async def submit_prompt(
        self, prompt: str, prompt_number: int, has_attachments: bool = False
    ) -> bool:
        """
        Submit a prompt to Claude using the chat-input textarea.

        Args:
            prompt: The prompt text to submit (empty string if only submitting files)
            prompt_number: The prompt number (for logging, 0 for file-only submission)
            has_attachments: Whether files are attached (for Claude, files are handled separately)
        """
        try:
            # On first prompt (or file-only submission), ensure "Accept all edits" is set
            if prompt_number == 1 or (prompt_number == 0 and has_attachments):
                await self.handle_initial_setup()

            frame = await self._get_claude_frame()
            if not frame:
                logger.error("❌ Could not find Claude frame")
                return False

            # Special case: Empty prompt with attachments (file-only submission)
            # For Claude, files are uploaded separately, so just submit without text
            if has_attachments and not prompt:
                logger.info("📎 Files already attached, submitting without text entry")
                # Find textarea to focus it
                textarea = await self._find_input_element(frame)
                if textarea:
                    await textarea.click()
                    await asyncio.sleep(0.3)
            else:
                # Normal prompt submission with text entry
                # Find and fill the textarea
                textarea = await self._find_input_element(frame)
                if not textarea:
                    logger.error("❌ Could not find chat-input element")
                    return False

                await textarea.click(force=True)
                await asyncio.sleep(0.3)
                # Clear first, then fill
                await textarea.fill("")
                await asyncio.sleep(0.1)
                await textarea.fill(prompt)
                logger.info("✅ Filled prompt in Claude textarea")

            await asyncio.sleep(0.5)

            # Dismiss blocking dialogs if they appeared between prompts
            await self._dismiss_clarification_dialog()
            await self._dismiss_permission_dialog()

            # Baseline the response count BEFORE sending (see __init__).
            try:
                self._pre_send_response_count = await self._get_response_count(frame)
            except Exception:
                self._pre_send_response_count = None

            # Click send button
            send_btn = await frame.query_selector('[data-testid="send-button"]')
            if not send_btn:
                send_btn = await frame.query_selector(
                    'button[aria-label="Send message"]'
                )
                if not send_btn:
                    send_btn = await frame.query_selector('button[aria-label="Send"]')

            if send_btn:
                # force=True bypasses Playwright's actionability check so the
                # click succeeds even when div.flex-1 (textarea wrapper)
                # overlaps the button in narrow panel layouts.
                await send_btn.click(force=True)
                logger.info("✅ Clicked send button")
            else:
                # Fallback: press Enter
                if textarea:
                    await textarea.press("Enter")
                    logger.info("✅ Pressed Enter to send")
                else:
                    logger.error("❌ Could not find textarea or send button")
                    return False

            return True

        except Exception as e:
            logger.error(f"❌ Failed to submit prompt: {e}")
            return False

    async def get_button_count(self) -> dict:
        """Count Claude's feedback buttons (thumbs up/down). Requires mouse hover."""
        try:
            frame = await self._get_claude_frame()
            if not frame:
                return {"upvote": 0, "downvote": 0}

            # Move mouse to chat area to reveal feedback buttons
            chat_area = await frame.query_selector("div.font-claude-response-small")
            if chat_area:
                try:
                    await chat_area.hover()
                    await asyncio.sleep(0.3)
                except Exception:
                    pass

            # Count thumbs up/down buttons - try various selectors
            upvote_count = 0
            downvote_count = 0

            # Claude uses SVG icons with titles like "Thumbs Up" / "Thumbs Down"
            # or buttons with aria-labels
            up_selectors = [
                'button[aria-label*="thumbs up" i]',
                'button[aria-label*="like" i]',
                'svg title:text("Thumbs Up")',
                '[data-testid*="upvote"]',
                '[data-testid*="like"]',
            ]
            down_selectors = [
                'button[aria-label*="thumbs down" i]',
                'button[aria-label*="dislike" i]',
                'svg title:text("Thumbs Down")',
                '[data-testid*="downvote"]',
                '[data-testid*="dislike"]',
            ]

            for sel in up_selectors:
                try:
                    btns = await frame.query_selector_all(sel)
                    upvote_count += len(btns)
                except Exception:
                    continue

            for sel in down_selectors:
                try:
                    btns = await frame.query_selector_all(sel)
                    downvote_count += len(btns)
                except Exception:
                    continue

            return {"upvote": upvote_count, "downvote": downvote_count}

        except Exception as e:
            logger.debug(f"Error counting Claude buttons: {e}")
            return {"upvote": 0, "downvote": 0}

    async def _get_response_count(self, frame) -> int:
        """Get current number of Claude responses.

        Claude responses are wrapped in <article> tags:
        <article class="flex justify-start mb-3 group relative">
          <div>...<div class="text-text-100 ... font-claude-response-small">...</div></div>
        </article>

        User messages are just <div class="flex justify-end mb-3">
        """
        try:
            # Count <article> elements - each Claude response is an article
            responses = await frame.query_selector_all("article")
            return len(responses) if responses else 0
        except Exception:
            return 0

    async def _has_stop_button(self, frame) -> bool:
        """Check if the Stop button is visible (Claude is processing)."""
        try:
            # Stop button has aria-label="Stop" - most reliable
            stop_btn = await frame.query_selector('button[aria-label="Stop"]')
            if stop_btn and await stop_btn.is_visible():
                return True
        except Exception:
            pass
        return False

    async def _dismiss_clarification_dialog(self) -> bool:
        """Dismiss the clarification question dialog if it's blocking the UI.

        The Claude add-in shows a fixed overlay (div[role="dialog"] with
        z-[1000]) at the bottom of the screen with numbered options and a
        "Skip" button. This overlay captures all pointer events, blocking
        the send button and stalling the agent.

        Detection (two approaches tried per frame):
        1. Visible "Skip" button confirmed by a "Minimize" aria-label button.
        2. div[role="dialog"] with class containing "fixed", then look for
           Skip / Minimize / last button inside it as fallback.

        Returns True if the dialog was found and dismissed, False otherwise.
        """
        for f in self.page.frames:
            try:
                # --- Approach 1: visible "Skip" button + Minimize confirm ---
                buttons = await f.query_selector_all("button")
                skip_btn = None
                for btn in buttons:
                    try:
                        text = await btn.text_content()
                        if text and text.strip() == "Skip":
                            if await btn.is_visible():
                                skip_btn = btn
                                break
                    except Exception:
                        continue

                if skip_btn:
                    minimize_btn = await f.query_selector(
                        'button[aria-label="Minimize"]'
                    )
                    if minimize_btn and await minimize_btn.is_visible():
                        logger.info(
                            "Clarification dialog detected (Skip btn), clicking Skip..."
                        )
                        await skip_btn.click(force=True)
                        await asyncio.sleep(1)
                        logger.info("Clarification dialog dismissed")
                        return True

                # --- Approach 2: div[role="dialog"] with fixed positioning ---
                dialog = await f.query_selector('div[role="dialog"]')
                if not dialog or not await dialog.is_visible():
                    continue

                cls = await dialog.get_attribute("class") or ""
                if "fixed" not in cls:
                    continue

                logger.info(
                    "Clarification dialog detected (div[role='dialog']), dismissing..."
                )

                # Try clicking Skip inside the dialog first
                dialog_buttons = await dialog.query_selector_all("button")
                dismissed = False
                for btn in dialog_buttons:
                    try:
                        text = (await btn.text_content() or "").strip()
                        if text == "Skip":
                            await btn.click(force=True)
                            dismissed = True
                            break
                    except Exception:
                        continue

                # Fallback: Minimize button
                if not dismissed:
                    minimize = await dialog.query_selector(
                        'button[aria-label="Minimize"]'
                    )
                    if minimize:
                        await minimize.click(force=True)
                        dismissed = True

                # Last resort: click the last button in the dialog
                if not dismissed and dialog_buttons:
                    try:
                        await dialog_buttons[-1].click(force=True)
                        dismissed = True
                    except Exception:
                        pass

                if dismissed:
                    await asyncio.sleep(1)
                    logger.info("Clarification dialog dismissed")
                    return True
                else:
                    logger.info("No dismiss button found in dialog, pressing Escape...")
                    await self.page.keyboard.press("Escape")
                    await asyncio.sleep(1)
                    return True

            except Exception:
                continue

        return False

    async def _dismiss_permission_dialog(self) -> bool:
        """Auto-approve the 'Claude wants to build...' permission dialog.

        The Claude add-in shows a modal when it wants to modify the
        spreadsheet.  Three buttons are offered:

        - **Deny** (Esc)
        - **Dangerously always allow** (Ctrl+Enter)  ← preferred
        - **Allow once** (Enter)  ← fallback

        We click "Dangerously always allow" so the dialog never reappears
        for the rest of the session.  If that button isn't found we fall
        back to "Allow once".

        Returns True if a dialog was found and dismissed, False otherwise.
        """
        for f in self.page.frames:
            try:
                # Prefer to scope to the permission modal — the add-in
                # ships role="dialog" on its container. Fall back to the
                # whole frame if no dialog element is present (older UI).
                dialogs = await f.query_selector_all('[role="dialog"]')
                scopes = dialogs if dialogs else [f]

                allow_always_btn = None
                allow_once_btn = None
                for scope in scopes:
                    buttons = await scope.query_selector_all("button")
                    for btn in buttons:
                        try:
                            # text_content() concatenates all descendant
                            # text, including the kbd shortcut suffix
                            # (e.g. "Allow once" + "Enter" → "Allow onceEnter").
                            # Use substring matches, never equality.
                            text = (await btn.text_content() or "").strip()
                            if not text:
                                continue
                            text_lower = text.lower()
                            if "dangerously" in text_lower and "allow" in text_lower:
                                if await btn.is_visible():
                                    allow_always_btn = btn
                            elif "allow once" in text_lower:
                                if await btn.is_visible():
                                    allow_once_btn = btn
                        except Exception:
                            continue
                    if allow_always_btn or allow_once_btn:
                        break  # already found in this dialog

                target = allow_always_btn or allow_once_btn
                if target:
                    label = (
                        "Dangerously always allow"
                        if target is allow_always_btn
                        else "Allow once"
                    )
                    logger.info(f"🔓 Permission dialog detected, clicking '{label}'...")
                    await target.click(force=True)
                    await asyncio.sleep(1)
                    logger.info("✅ Permission dialog dismissed")
                    return True

            except Exception:
                continue

        return False

    async def wait_for_completion(
        self, prompt_number: int, initial_counts: dict = None
    ) -> bool:
        """Wait for Claude to complete by checking Stop button presence.

        Logic:
        1. Wait for Stop button to APPEAR (Claude started processing)
        2. Wait for Stop button to DISAPPEAR (Claude finished)
        3. Brief stabilization to ensure Claude is fully done
        """
        agent_config = self.get_config_section()
        max_wait = agent_config.get("max_wait_per_prompt_seconds", 900)

        # Baseline from submit_prompt (sampled before the send click);
        # falls back to sampling now only if submit never set it.
        self._claude_frame = None
        frame = await self._get_claude_frame()
        if self._pre_send_response_count is not None:
            initial_response_count = self._pre_send_response_count
            self._pre_send_response_count = None
        else:
            initial_response_count = (
                await self._get_response_count(frame) if frame else 0
            )

        logger.info(
            f"⏳ Waiting for prompt #{prompt_number} (starting with {initial_response_count} responses)..."
        )

        start_time = time.monotonic()
        check_interval = 1  # Check every second for responsiveness
        saw_stop_button = False  # Must see Stop button before accepting completion

        while (time.monotonic() - start_time) < max_wait:
            if self.shutdown_event and self.shutdown_event.is_set():
                return False

            await asyncio.sleep(check_interval)
            elapsed = int(time.monotonic() - start_time)

            # Refresh frame reference
            self._claude_frame = None
            frame = await self._get_claude_frame()
            if not frame:
                continue

            # Vendor notice banners are not chat replies (no <article>), so
            # the response-count logic never sees them. Scan the panel text
            # every 15s: a usage-limit / outage banner means the prompt will
            # never complete — bail as provider_unavailable (infra, unrecorded)
            # instead of sitting here until the 5h cap records a TIMEOUT row.
            if elapsed % 15 == 0:
                try:
                    panel_text = await frame.evaluate(
                        "() => (document.body && document.body.innerText) "
                        "? document.body.innerText.slice(-2000) : ''"
                    )
                except Exception:
                    panel_text = ""
                notice = self.panel_notice(panel_text)
                if notice:
                    self.provider_unavailable = True
                    self.provider_retry_after = self.parse_limit_reset_epoch(panel_text)
                    tail = " ".join((panel_text or "")[-500:].split())
                    logger.error(
                        f"❌ Claude panel shows a vendor notice ({notice!r}); "
                        f"limit reset epoch={self.provider_retry_after} — "
                        "aborting the prompt as provider_unavailable. "
                        f"Panel tail: {tail!r}"
                    )
                    return False

            # Scroll to bottom periodically
            if elapsed % 5 == 0:
                try:
                    await frame.evaluate(
                        "window.scrollTo(0, document.body.scrollHeight)"
                    )
                except Exception:
                    pass

            try:
                stop_visible = await self._has_stop_button(frame)
                current_response_count = await self._get_response_count(frame)

                if stop_visible:
                    # Claude is processing
                    saw_stop_button = True
                    # Check for blocking dialogs every ~5 seconds
                    if elapsed % 5 == 0:
                        await self._dismiss_clarification_dialog()
                        await self._dismiss_permission_dialog()
                    if elapsed % 10 == 0:
                        logger.info(
                            f"   [{elapsed}s] Claude processing... (responses: {current_response_count})"
                        )
                    continue

                # Stop button not visible
                if not saw_stop_button:
                    # Haven't seen Claude start yet - keep waiting
                    # Also check for blocking dialogs (edge case: Claude
                    # asks a question before the Stop button appears)
                    if elapsed % 5 == 0:
                        await self._dismiss_clarification_dialog()
                        await self._dismiss_permission_dialog()
                    if elapsed % 10 == 0:
                        logger.info(
                            f"   [{elapsed}s] Waiting for Claude to start... (responses: {current_response_count})"
                        )
                    continue

                # Saw stop button, now it's gone - Claude finished!
                # Brief stabilization to ensure fully done
                logger.info(
                    f"   [{elapsed}s] Stop button gone, verifying completion..."
                )
                await asyncio.sleep(3)

                # Re-check that stop button is still gone
                self._claude_frame = None
                frame = await self._get_claude_frame()
                if frame:
                    still_stopped = not await self._has_stop_button(frame)
                    final_count = await self._get_response_count(frame)

                    if still_stopped and final_count > initial_response_count:
                        # A finished turn whose last reply is a vendor
                        # capacity notice is not a completed prompt — the
                        # work was never done. Flag it for the engine to
                        # treat as infra (unrecorded, retried).
                        last_text = ""
                        try:
                            articles = await frame.query_selector_all("article")
                            if articles:
                                last_text = await articles[-1].text_content() or ""
                        except Exception:
                            last_text = ""
                        if self.looks_like_capacity_notice(last_text):
                            self.provider_unavailable = True
                            logger.error(
                                "❌ Claude answered with a capacity/outage notice "
                                f"instead of working: {last_text.strip()[:200]!r}"
                            )
                            return False
                        logger.info(
                            f"✅ Prompt #{prompt_number} completed! (responses: {final_count}, was {initial_response_count})"
                        )
                        return True
                    elif not still_stopped:
                        # Stop button reappeared - Claude started again
                        logger.info(f"   [{elapsed}s] Claude resumed processing...")
                        continue
                    else:
                        # No new responses - might be an issue, but keep waiting
                        logger.info(
                            f"   [{elapsed}s] Stop gone but no new responses (count: {final_count})"
                        )
                        saw_stop_button = False  # Reset to catch next cycle

            except Exception as e:
                logger.debug(f"   [{elapsed}s] Check error: {e}")

        logger.error(f"❌ Timeout waiting for prompt #{prompt_number}")
        return False

    async def _find_plus_button(self):
        """Find the composer "+" trigger that opens the attach popover.

        Returns ``(frame, button)`` or ``(None, None)``.

        The Claude panel ships **two** buttons with ``aria-label="More options"``:
        the attach ``+`` next to the composer, and the header ``⋮`` three-dot
        menu. They are distinguished by their inner SVG icon — the attach
        button's SVG contains ``<title>Plus</title>`` (semantic a11y label
        for the glyph), the three-dot menu does not. ``query_selector`` with
        only the aria-label returns the first document-order match (which is
        the header ``⋮``) and silently clicks the wrong control, so we
        require the Plus SVG title as part of the selector.
        """
        frame = await self._get_claude_frame()
        if not frame:
            return None, None

        # Primary: SVG <title>Plus</title> uniquely marks the attach button.
        try:
            btn = await frame.query_selector(
                'button[aria-label="More options"]:has(svg > title:text-is("Plus"))'
            )
            if btn and await btn.is_visible():
                return frame, btn
        except Exception:
            pass

        # Fallback: match by the Plus glyph's SVG path data (d attribute
        # starts with "M224,128..."), in case the <title> element is removed.
        try:
            btn = await frame.query_selector(
                'button[aria-label="More options"]:has(svg path[d^="M224,128"])'
            )
            if btn and await btn.is_visible():
                return frame, btn
        except Exception:
            pass

        return None, None

    async def _find_menu_item(self, text: str):
        """Search all frames + main page for a menu item with the given text.

        Primary strategy is ``get_by_role("menuitem", name=...)`` — Playwright's
        accessible-name computation walks nested icon/label structure and
        normalizes whitespace, so it matches the real DOM shape that Anthropic
        ships (``<div role="menuitem"><svg/><span>Add files or photos</span></div>``).
        CSS ``:text-is`` / ``:has-text`` selectors stay as fallbacks.

        Returns an element handle or ``None``.
        """
        name_re = re.compile(re.escape(text), re.I)

        contexts = list(self.page.frames) + [self.page]
        for ctx in contexts:
            # Primary: ARIA role + accessible name.
            try:
                loc = ctx.get_by_role("menuitem", name=name_re).first
                if await loc.count() > 0 and await loc.is_visible():
                    return await loc.element_handle()
            except Exception:
                pass

            # Fallback 1: role=menuitem with substring text match (in case
            # accessible name is empty but visible text is present).
            try:
                el = await ctx.query_selector(f'[role="menuitem"]:has-text("{text}")')
                if el and await el.is_visible():
                    return el
            except Exception:
                pass

            # Fallback 2: original CSS-text strategies.
            try:
                el = await ctx.query_selector(f'span:text-is("{text}")')
                if el and await el.is_visible():
                    return el

                el = await ctx.query_selector(f'button:has-text("{text}")')
                if el and await el.is_visible():
                    return el

                el = await ctx.query_selector(f':text-is("{text}")')
                if el and await el.is_visible():
                    return el
            except Exception:
                continue

        return None

    async def _plus_menu_is_open(self) -> bool:
        """Return True if the Plus popover is currently showing any menuitem.

        Used to avoid the toggle trap: if an earlier setup step
        (``_disable_web_search``) left the menu open, clicking Plus again
        would close it instead of opening it.
        """
        for ctx in list(self.page.frames) + [self.page]:
            try:
                el = await ctx.query_selector('[role="menuitem"]')
                if el and await el.is_visible():
                    return True
            except Exception:
                continue
        return False

    async def _open_plus_menu_and_find(self, target_text: str, max_retries: int = 3):
        """Open the Plus popover (if not already open) and return a menu item.

        Handles the toggle trap: if the popover is already on screen from an
        earlier setup step (e.g. ``_disable_web_search``), clicking Plus again
        would close it. We search first, and only click Plus when the menu is
        confirmed closed.

        Returns the menu item element handle, or ``None`` after ``max_retries``.
        """
        for attempt in range(max_retries):
            # If a prior step left the popover open, search directly — no click.
            if await self._plus_menu_is_open():
                menu_item = await self._find_menu_item(target_text)
                if menu_item:
                    return menu_item
                # Menu is open but item isn't there. Dismiss and fall through
                # to click Plus fresh on the same attempt.
                try:
                    await self.page.keyboard.press("Escape")
                except Exception:
                    pass
                await asyncio.sleep(0.5)

            frame, plus_btn = await self._find_plus_button()
            if not plus_btn:
                logger.error("❌ Could not find Plus button")
                return None

            await plus_btn.click()
            await asyncio.sleep(1.0)

            menu_item = await self._find_menu_item(target_text)
            if menu_item:
                return menu_item

            # Give the popover a beat longer in case of animation.
            await asyncio.sleep(1.0)
            menu_item = await self._find_menu_item(target_text)
            if menu_item:
                return menu_item

            logger.warning(
                f"⚠️ '{target_text}' not found on attempt {attempt + 1}/{max_retries}"
            )

            # On the final attempt, capture a full screenshot + DOM dump so we
            # can see what the add-in is actually rendering.
            if attempt == max_retries - 1:
                await self._dump_upload_dom("claude_plus_menu")
                for f in self.page.frames:
                    try:
                        spans = await f.query_selector_all("span")
                        texts = []
                        for s in spans[:20]:
                            t = await s.text_content()
                            if t and t.strip() and len(t.strip()) < 50:
                                texts.append(repr(t.strip()))
                        if texts:
                            logger.error(f"   Frame {f.name or f.url[:50]}: {texts}")
                    except Exception:
                        continue

            try:
                if frame:
                    await frame.press("body", "Escape")
                else:
                    await self.page.keyboard.press("Escape")
            except Exception:
                pass
            await asyncio.sleep(1.0)

        return None

    async def upload_files(self, file_paths: str | list) -> bool:
        """Upload files to Claude (one at a time)."""
        if isinstance(file_paths, str):
            file_paths = [file_paths]

        if not file_paths:
            return True

        MAX_SIZE = 25 * 1024 * 1024  # 25 MB

        logger.info(f"📎 Checking {len(file_paths)} file(s) for upload...")
        for fp in file_paths:
            p = Path(fp)
            if not p.exists():
                logger.error(f"❌ File not found: {fp}")
                return False
            if p.stat().st_size > MAX_SIZE:
                logger.error(f"❌ File too large: {p.name}")
                return False
            logger.info(f"   📄 {p.name} ({p.stat().st_size / 1024 / 1024:.2f} MB)")

        try:
            for i, fp in enumerate(file_paths):
                p = Path(fp)
                logger.info(f"📤 Uploading {i + 1}/{len(file_paths)}: {p.name}")

                uploaded = False
                menu = await self._open_plus_menu_and_find("Add files or photos")
                if menu:
                    logger.info("✅ Found 'Add files or photos' — clicking...")
                    try:
                        async with self.page.expect_file_chooser(timeout=8000) as fc:
                            await menu.click()
                        chooser = await fc.value
                        await chooser.set_files([fp])
                        logger.info(f"   ✅ Selected: {p.name}")
                        uploaded = True
                    except Exception as e:
                        logger.warning(
                            f"⚠️ File chooser did not fire after menu click: {e}"
                        )

                # Fallback: some React chat UIs keep a hidden <input type="file">
                # that the visible "+" proxies to. Driving it directly bypasses
                # the native chooser entirely.
                if not uploaded:
                    if await self._set_files_on_hidden_input(fp):
                        logger.info(
                            f"   ✅ Selected via hidden input fallback: {p.name}"
                        )
                        uploaded = True

                if not uploaded:
                    logger.error("❌ 'Add files or photos' not found after retries")
                    await self._dump_upload_dom("claude_upload_failed")
                    return False

                await asyncio.sleep(min(5 + p.stat().st_size / 1024 / 100, 30))

            logger.info(f"✅ All {len(file_paths)} file(s) uploaded!")
            return True

        except Exception as e:
            logger.error(f"❌ Upload failed: {e}")
            import traceback

            logger.error(traceback.format_exc())
            await self._dump_upload_dom("claude_upload_exception")
            return False

    async def _set_files_on_hidden_input(self, file_path) -> bool:
        """Set files on a hidden ``<input type="file">`` INSIDE the Claude
        frame (or its child frames) only.

        Never searches the whole page: Excel Online ships its own hidden
        file inputs earlier in frame order, and setting files on one
        "succeeds" silently while nothing is attached in the Claude panel
        — the task would then run without its input files.
        """
        scope = await self._get_claude_frame()
        if scope is None:
            return False
        ctxs = [scope] + [
            f for f in self.page.frames if f is not scope and f.parent_frame is scope
        ]
        for ctx in ctxs:
            try:
                inputs = await ctx.query_selector_all('input[type="file"]')
            except Exception:
                continue
            for inp in inputs:
                try:
                    await inp.set_input_files([str(file_path)])
                    return True
                except Exception:
                    continue
        return False
