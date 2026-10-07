"""
ChatGPT Excel add-in core interaction logic.

Handles ChatGPT-specific implementation extending AIAgentCore.
Mirrors the structure of claude_core.py with ChatGPT selectors and UI flow.
"""

import asyncio
import logging
import re
import time
from pathlib import Path

from .ai_agent_base import AIAgentCore, is_host_frame as _is_host_frame

logger = logging.getLogger(__name__)


class ChatGPTCore(AIAgentCore):
    """ChatGPT Excel add-in specific implementation."""

    def __init__(self, page, config, shutdown_event, completion_logger):
        super().__init__(page, config, shutdown_event, completion_logger)
        self._chatgpt_frame = None
        self._setup_completed = False

    def get_agent_type(self) -> str:
        return "chatgpt_excel_agent"

    def get_addon_name(self) -> str:
        return "ChatGPT"

    def get_open_button_text(self) -> str:
        return "Open ChatGPT"

    def requires_addins_menu(self) -> bool:
        return True

    def get_ribbon_launcher_name(self) -> str:
        # The installed add-in's own Home-ribbon button / overflow menuitem
        # (probed 2026-09-10: "ChatGPT", data-unique-id Ribbon-AddinControlN).
        return "ChatGPT"

    # ------------------------------------------------------------------
    # Selectors for the ChatGPT chat input, ordered by likelihood.
    # ------------------------------------------------------------------
    _INPUT_SELECTORS = [
        'div[contenteditable="true"][role="textbox"]',
        'div[contenteditable="true"]',
        'textarea[placeholder*="Ask anything"]',
        'textarea[role="textbox"]',
    ]

    # URL fragments that identify the ChatGPT add-in iframe.
    _FRAME_URL_HINTS = ["chatgpt.com", "openai.com"]

    # ------------------------------------------------------------------
    # Frame detection
    # ------------------------------------------------------------------

    async def _get_chatgpt_frame(self):
        """Find and cache the ChatGPT add-in iframe.

        Strategy: try frames matching known URLs first, then fall back to
        any frame that contains a chat-input element.
        """
        if self._chatgpt_frame:
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await self._chatgpt_frame.query_selector(selector)
                    if el:
                        return self._chatgpt_frame
            except Exception:
                pass
            self._chatgpt_frame = None

        # Pass 1: frames whose URL matches known hints (fast path)
        for f in self.page.frames:
            url = f.url or ""
            if not any(hint in url for hint in self._FRAME_URL_HINTS):
                continue
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await f.query_selector(selector)
                    if el:
                        self._chatgpt_frame = f
                        return f
            except Exception:
                continue

        # Pass 2: brute-force all frames (URL changed or new domain).
        # Host (Excel/OneDrive) frames are excluded and the input must be
        # visible — the unqualified contenteditable selector would
        # otherwise match Excel's own formula bar during panel boot.
        for f in self.page.frames:
            if _is_host_frame(f):
                continue
            try:
                for selector in self._INPUT_SELECTORS:
                    el = await f.query_selector(selector)
                    if el and await el.is_visible():
                        self._chatgpt_frame = f
                        return f
            except Exception:
                continue
        return None

    async def _find_input_element(self, frame=None):
        """Find the chat input element in the given frame (or cached frame).

        Returns the element handle, or None.
        """
        frame = frame or await self._get_chatgpt_frame()
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

    # ------------------------------------------------------------------
    # Panel verification
    # ------------------------------------------------------------------

    async def _verify_panel_opened(self) -> bool:
        """ChatGPT-specific: poll for chat-input with configurable boot timeout."""
        boot_timeout = self.config.get("panel_boot_timeout_seconds", 20)
        poll_interval = 2
        elapsed = 0

        while elapsed < boot_timeout:
            if self.shutdown_event and self.shutdown_event.is_set():
                logger.info("   Shutdown requested during panel verification")
                return False
            self._chatgpt_frame = None  # fresh lookup each time
            frame = await self._get_chatgpt_frame()
            if frame:
                logger.info(
                    f"   ✅ Panel verified: ChatGPT chat-input found after {elapsed}s"
                )
                return True
            if elapsed > 0 and elapsed % 6 == 0:
                logger.info(
                    f"   ⏳ [{elapsed}s] Waiting for ChatGPT app to initialize..."
                )
            await asyncio.sleep(poll_interval)
            elapsed += poll_interval

        # Last-ditch: dump what frames and elements exist for debugging
        logger.warning(f"   ⚠️ ChatGPT chat-input not found after {boot_timeout}s")
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

    # ------------------------------------------------------------------
    # Initial setup: Thinking effort + Apply edits automatically
    #
    # Both are config-driven. If the corresponding key under
    # `chatgpt_excel_agent:` in the template is unset, the UI is left
    # alone. No hardcoded model / mode / toggle names.
    # ------------------------------------------------------------------

    async def handle_initial_setup(self) -> bool:
        """Run configured post-panel-open setup steps."""
        if self._setup_completed:
            return True

        try:
            logger.info("🔧 Running ChatGPT initial setup...")
            frame = await self._get_chatgpt_frame()
            if not frame:
                logger.warning("⚠️ Could not find ChatGPT frame for setup")
                return False

            # Pin + verify the model, then the thinking effort — both from
            # the identity registry via the combined "Model and thinking
            # effort" menu (2026-08-27 UI). A miss on either is a setup
            # failure (engine → PANEL_FAILED → infra retry, nothing
            # recorded) — never "run on whatever the panel happened to
            # show": a fresh chat was observed defaulting to GPT-5.5
            # Thinking, so an unpinned model is a silently wrong cohort.
            if not await self._select_model(frame):
                return False
            if not await self._select_thinking_effort(frame):
                return False
            # The edits toggle is convenience, not identity — best-effort.
            await self._apply_edits_toggle(frame)

            self._setup_completed = True
            logger.info("✅ ChatGPT initial setup complete")
            return True

        except Exception as e:
            logger.warning(f"⚠️ Error in initial setup: {e}")
            try:
                frame = await self._get_chatgpt_frame()
                if frame:
                    await frame.press("body", "Escape")
                    await asyncio.sleep(0.3)
            except Exception:
                pass
            return False

    # ------------------------------------------------------------------
    # The combined "Model and thinking effort" trigger (2026-08-27 UI).
    #
    # One composer button (id `reasoning-effort-select` — the pre-drift id
    # survived) opens one Radix menu holding a "Model" group ("GPT-5.5
    # Thinking" / "GPT-5.6 Sol") and a "Thinking effort" group ("Instant" /
    # "Medium" / "High" / "Extra High"). The button's aria-label carries
    # both axes: "Model and thinking effort: <model> · <effort>", so a
    # single read verifies everything; its text renders the model only.
    # Menu items expose no aria-checked, so verification always reads the
    # trigger, never the items.
    # ------------------------------------------------------------------

    async def _get_effort_pill(self, frame):
        """The combined model/effort trigger button, by stable id first."""
        for sel in (
            "button#reasoning-effort-select",
            'button[aria-label^="Model and thinking effort"]',
            'button[aria-label^="Thinking effort"]',  # pre-drift pill
        ):
            el = await frame.query_selector(sel)
            if el:
                return el
        return None

    @staticmethod
    def _parse_trigger_state(aria_label, text):
        """(model, effort) as the combined trigger currently shows them.

        Handles both generations: the combined form
        "Model and thinking effort: <model> · <effort>" and the pre-drift
        effort-only pill "Thinking effort: <effort>". Either slot is None
        when the trigger does not state it — callers must treat a None
        slot as unverified, never as a match.
        """
        label = (aria_label or "").strip()
        tail = label.split(":", 1)[1].strip() if ":" in label else ""
        if "·" in tail:
            model, effort = (part.strip() for part in tail.split("·", 1))
            return (model or None), (effort or None)
        if label.lower().startswith("thinking effort"):
            return None, (tail or (text or "").strip() or None)
        if tail:
            return tail, None
        return ((text or "").strip() or None), None

    async def _read_trigger(self, frame):
        """(model, effort) from the current trigger, (None, None) if absent."""
        pill = await self._get_effort_pill(frame)
        if not pill:
            return None, None
        return self._parse_trigger_state(
            await pill.get_attribute("aria-label"),
            await pill.text_content(),
        )

    async def _open_trigger_menu(self, frame):
        """Click the trigger and wait for its Radix menu; None on failure."""
        pill = await self._get_effort_pill(frame)
        if not pill or not await pill.is_visible():
            return None
        await pill.click()
        for _ in range(20):
            await asyncio.sleep(0.1)
            menu = await frame.query_selector('[role="menu"][data-state="open"]')
            if menu:
                return menu
        return None

    @staticmethod
    async def _click_menu_item(menu, target_lower):
        """Click the menu item whose primary label matches target.

        Equality on the `span.truncate` label first (an item's text also
        contains its description, and effort labels nest — "High" is a
        substring of "Extra High" — so equality must win); substring on the
        truncate label, then on the full text, only as drift fallbacks.
        """
        items = await menu.query_selector_all('[role="menuitem"]')

        async def _label(item):
            label_el = await item.query_selector("span.truncate")
            return (
                ((await label_el.text_content()) if label_el else "")
                .strip()
                .lower()
            )

        for match in ("equal", "truncate_sub", "full_sub"):
            for item in items:
                try:
                    if not await item.is_visible():
                        continue
                    if match == "equal":
                        hit = (await _label(item)) == target_lower
                    elif match == "truncate_sub":
                        label = await _label(item)
                        hit = bool(label) and target_lower in label
                    else:
                        full = (await item.text_content() or "").strip().lower()
                        hit = target_lower in full
                    if hit:
                        await item.click()
                        return True
                except Exception:
                    continue
        return False

    async def _select_model(self, frame) -> bool:
        """Pin the add-in model via the combined model/effort menu.

        Reads ``chatgpt_excel_agent.ui_model_label`` from config (injected
        by the runner from the agent-identity registry) and matches it
        case-insensitively against the trigger's model slot and the menu's
        Model items. No whitelist — new models are a registry entry away,
        never a code change. Leave unset to skip (legacy identities are
        already refused at resolve time).

        Returns True when no model is pinned, or when the trigger verifiably
        shows the pinned model afterwards; False otherwise — the caller
        treats False as a setup failure (infra, unrecorded), never as "run
        on whatever model happens to be active".
        """
        target = self.config.get("chatgpt_excel_agent", {}).get("ui_model_label")
        if not target:
            return True
        target_lower = str(target).strip().lower()
        if not target_lower:
            return True

        async def _attempt():
            pill = await self._get_effort_pill(frame)
            if not pill or not await pill.is_visible():
                return "pill_missing"
            model, _ = self._parse_trigger_state(
                await pill.get_attribute("aria-label"),
                await pill.text_content(),
            )
            if model and model.strip().lower() == target_lower:
                return "already"

            menu = await self._open_trigger_menu(frame)
            if not menu:
                await self._dismiss_dropdown(frame)
                return "menu_missing"
            if not await self._click_menu_item(menu, target_lower):
                await self._dismiss_dropdown(frame)
                return "not_found"

            await asyncio.sleep(0.5)
            model_after, _ = await self._read_trigger(frame)
            if model_after and model_after.strip().lower() == target_lower:
                return "selected"
            logger.warning(
                "Model click did not change trigger (shows '%s')", model_after
            )
            return "click_no_effect"

        try:
            for attempt in (1, 2):
                logger.info("Selecting ChatGPT model: %s (attempt %d)", target, attempt)
                result = await _attempt()
                if result == "selected":
                    logger.info("✅ Model verified: '%s'", target)
                    return True
                if result == "already":
                    logger.info("✅ Model '%s' already selected", target)
                    return True
                if result == "pill_missing":
                    logger.error(
                        "Model/effort trigger not found — the pinned model "
                        "cannot be applied or verified; aborting setup"
                    )
                    return False
                if result in ("menu_missing", "not_found", "click_no_effect"):
                    if attempt == 1:
                        logger.info(
                            "Model attempt 1 failed (%s); retrying", result
                        )
                        await asyncio.sleep(0.5)
                        continue
                    logger.error(
                        "Model '%s' not applied after 2 attempts (%s) — "
                        "aborting setup rather than running on an unverified "
                        "model",
                        target,
                        result,
                    )
                    return False
            return False
        except Exception as e:
            logger.error("Could not set model: %s", e)
            await self._dismiss_dropdown(frame)
            return False

    async def _select_thinking_effort(self, frame) -> bool:
        """Pin the 'Thinking effort' pill to the configured value.

        Returns True when no effort is pinned, or when the pill verifiably
        shows the pinned value afterwards; False otherwise — the caller
        treats False as a setup failure.

        Reads ``chatgpt_excel_agent.thinking_effort`` from config and matches
        it case-insensitively against the combined trigger's effort slot and
        the menu's Thinking-effort items. No whitelist — any current or
        future label works ("Instant" / "Medium" / "High" / "Extra High" on
        the 2026-08-27 UI). Leave unset to skip.

        Shares the combined "Model and thinking effort" trigger and menu
        with `_select_model` (see the section comment above). Verification
        reads the trigger's parsed effort slot with EQUALITY — effort
        labels nest ("High" ⊂ "Extra High"), so substring verification
        would report false matches.
        """
        target = self.config.get("chatgpt_excel_agent", {}).get("thinking_effort")
        if not target:
            return True

        target_lower = str(target).strip().lower()
        if not target_lower:
            return True

        async def _attempt():
            """Return 'selected', 'already', 'pill_missing', 'menu_missing',
            'not_found', or 'click_no_effect'."""
            pill = await self._get_effort_pill(frame)
            if not pill or not await pill.is_visible():
                return "pill_missing"

            _, effort = self._parse_trigger_state(
                await pill.get_attribute("aria-label"),
                await pill.text_content(),
            )
            if effort and effort.strip().lower() == target_lower:
                return "already"

            menu = await self._open_trigger_menu(frame)
            if not menu:
                await self._dismiss_dropdown(frame)
                return "menu_missing"
            if not await self._click_menu_item(menu, target_lower):
                await self._dismiss_dropdown(frame)
                return "not_found"

            # Verify the trigger actually changed. Radix popovers
            # occasionally accept the click visually but the parent React
            # state misses the update — verify so we can retry instead of
            # running with the wrong effort. A vanished or unreadable
            # trigger is NOT success: verification needs positive evidence.
            await asyncio.sleep(0.5)
            _, after = await self._read_trigger(frame)
            if after and after.strip().lower() == target_lower:
                return "selected"
            logger.warning(
                "Thinking-effort click did not change trigger (shows '%s')", after
            )
            return "click_no_effect"

        try:
            for attempt in (1, 2):
                logger.info(
                    "Selecting Thinking effort: %s (attempt %d)", target, attempt
                )
                result = await _attempt()
                if result == "selected":
                    logger.info("✅ Thinking effort verified: '%s'", target)
                    return True
                if result == "already":
                    logger.info("✅ Thinking effort '%s' already selected", target)
                    return True
                if result == "pill_missing":
                    logger.error(
                        "Thinking-effort pill not found — the pinned effort "
                        "cannot be applied or verified; aborting setup"
                    )
                    return False
                if result in ("menu_missing", "not_found", "click_no_effect"):
                    if attempt == 1:
                        logger.info(
                            "Thinking-effort attempt 1 failed (%s); retrying",
                            result,
                        )
                        await asyncio.sleep(0.5)
                        continue
                    logger.error(
                        "Thinking effort '%s' not applied after 2 attempts "
                        "(%s) — aborting setup rather than running on an "
                        "unverified effort",
                        target,
                        result,
                    )
                    return False
            return False
        except Exception as e:
            logger.error("Could not set thinking effort: %s", e)
            await self._dismiss_dropdown(frame)
            return False

    async def _dismiss_dropdown(self, frame):
        """Press Escape to close any open popover; swallow any failure."""
        try:
            await frame.press("body", "Escape")
            await asyncio.sleep(0.3)
        except Exception:
            pass

    async def _find_settings_button(self, frame):
        """Find the "…" settings / kebab menu button in the ChatGPT frame.

        Selectors, ordered from most to least robust:

          1. Exact aria-label "Open settings menu" — the label ChatGPT
             currently ships (stable, human-facing, designed for a11y).
          2. aria-haspopup="menu" combined with aria-label containing
             settings/options/menu/more — catches future label variants
             while still constraining to popup-triggers.
          3. aria-label substring fallback (no haspopup constraint).

        The old SVG-path heuristic was retired once we had the stable
        aria-label; keeping only attribute-based matches.
        """
        selectors = [
            # Exact label we've confirmed in the current UI.
            'button[aria-label="Open settings menu"]',
            # Popup-opener + descriptive label (resilient to label tweaks).
            'button[aria-haspopup="menu"][aria-label*="settings" i]',
            'button[aria-haspopup="menu"][aria-label*="option" i]',
            'button[aria-haspopup="menu"][aria-label*="menu" i]',
            'button[aria-haspopup="menu"][aria-label*="more" i]',
            # Bare aria-label substring (last resort — may match unrelated
            # buttons, but still far more principled than SVG sniffing).
            'button[aria-label*="settings" i]',
            'button[aria-label*="option" i]',
            'button[aria-label*="more" i]',
        ]
        for sel in selectors:
            try:
                btn = await frame.query_selector(sel)
                if btn and await btn.is_visible():
                    return btn
            except Exception:
                continue

        return None

    async def _apply_edits_toggle(self, frame):
        """Set the 'Apply edits automatically' switch to the configured value.

        Reads ``chatgpt_excel_agent.apply_edits_automatically`` from config:
          * ``True``  → ensure toggle is ON
          * ``False`` → ensure toggle is OFF
          * unset / non-bool → skip, leave the UI alone

        Selector chain (most to least robust):
          * settings trigger: ``_find_settings_button`` (aria-label anchored)
          * popover:          ``[role="menu"][data-state="open"]`` whose
                              ``aria-labelledby`` matches the trigger's id
                              (Radix links each popover back to its opener)
          * switch:           ``<label>`` with text "Apply edits automatically"
                              → ``for`` attr → ``[id=<for>]`` inside popover
                              (fallback: the only ``button[role="switch"]``
                              inside the popover)
        """
        desired = self.config.get("chatgpt_excel_agent", {}).get(
            "apply_edits_automatically"
        )
        if not isinstance(desired, bool):
            return

        try:
            logger.info("Setting 'Apply edits automatically': desired=%s", desired)

            # Locate the "…" settings trigger, across all frames if needed.
            settings_btn = await self._find_settings_button(frame)
            trigger_frame = frame
            if not settings_btn:
                for f in self.page.frames:
                    settings_btn = await self._find_settings_button(f)
                    if settings_btn:
                        trigger_frame = f
                        break
            if not settings_btn:
                logger.warning(
                    "Settings (…) button not found — cannot set apply-edits toggle"
                )
                return

            # Remember the trigger's id so we can match the specific popover
            # this button opens (Radix wires aria-labelledby = trigger id).
            trigger_id = await settings_btn.get_attribute("id")

            await settings_btn.click()

            # Wait for the popover. Prefer the one aria-labelledby the
            # trigger; fall back to any open menu. ~2s budget.
            popover = None
            popover_frame = trigger_frame
            for _ in range(20):
                await asyncio.sleep(0.1)
                candidates = []
                for f in [trigger_frame] + list(self.page.frames):
                    try:
                        if trigger_id:
                            el = await f.query_selector(
                                '[role="menu"][data-state="open"]'
                                f'[aria-labelledby="{trigger_id}"]'
                            )
                            if el:
                                candidates.append((el, f))
                        if not candidates:
                            el = await f.query_selector(
                                '[role="menu"][data-state="open"]'
                            )
                            if el:
                                candidates.append((el, f))
                    except Exception:
                        continue
                if candidates:
                    popover, popover_frame = candidates[0]
                    break

            if not popover:
                logger.warning(
                    "Settings popover did not open (no [role=menu]"
                    "[data-state=open] found) — skipping"
                )
                await self._dismiss_dropdown(frame)
                return

            # Find the switch, scoped to the popover.
            switch = None
            labels = await popover.query_selector_all("label")
            for label in labels:
                try:
                    text = (await label.text_content() or "").strip()
                    if "Apply edits automatically" not in text:
                        continue
                    label_for = await label.get_attribute("for")
                    if label_for:
                        # Attribute selectors with colons in the value (Radix
                        # ids like ":r2f:") work fine when double-quoted.
                        switch = await popover.query_selector(f'[id="{label_for}"]')
                    break
                except Exception:
                    continue

            if not switch:
                # Fallback: the sole role=switch inside the popover.
                switch = await popover.query_selector('button[role="switch"]')

            if not switch:
                logger.warning(
                    "'Apply edits automatically' toggle not found in popover"
                )
                await self._dismiss_dropdown(frame)
                return

            # Read current state. data-state and aria-checked track the
            # same fact; check both so we're insensitive to which one
            # Radix happens to update first.
            state = (await switch.get_attribute("data-state") or "").lower()
            aria_checked = (await switch.get_attribute("aria-checked") or "").lower()
            is_on = state == "checked" or aria_checked == "true"

            if is_on == desired:
                logger.info(
                    "'Apply edits automatically' already %s",
                    "on" if desired else "off",
                )
            else:
                await switch.click()
                await asyncio.sleep(0.5)
                state2 = (await switch.get_attribute("data-state") or "").lower()
                aria2 = (await switch.get_attribute("aria-checked") or "").lower()
                now_on = state2 == "checked" or aria2 == "true"
                if now_on == desired:
                    logger.info(
                        "'Apply edits automatically' set to %s",
                        "on" if desired else "off",
                    )
                else:
                    logger.warning(
                        "'Apply edits automatically' click did not change"
                        " state (still %s; wanted %s)",
                        "on" if now_on else "off",
                        "on" if desired else "off",
                    )

            await self._dismiss_dropdown(popover_frame)

        except Exception as e:
            logger.debug("Could not set apply-edits toggle: %s", e)
            await self._dismiss_dropdown(frame)

    # ------------------------------------------------------------------
    # Panel close
    # ------------------------------------------------------------------

    async def _close_panel(self) -> bool:
        """Close the ChatGPT add-in panel to allow fresh reopen."""
        try:
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
                            self._chatgpt_frame = None
                            self._setup_completed = False
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
                self._chatgpt_frame = None
                self._setup_completed = False
                await asyncio.sleep(3)
                return True

            return False
        except Exception as e:
            logger.debug(f"Error closing panel: {e}")
            return False

    # ------------------------------------------------------------------
    # Session health
    # ------------------------------------------------------------------

    async def verify_session_health(self) -> bool:
        """Verify ChatGPT iframe is responsive (not just present)."""
        try:
            self._chatgpt_frame = None
            frame = await self._get_chatgpt_frame()
            if not frame:
                logger.warning("❌ Health check: ChatGPT frame not found")
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

    # ------------------------------------------------------------------
    # Prompt submission
    # ------------------------------------------------------------------

    async def submit_prompt(
        self, prompt: str, prompt_number: int, has_attachments: bool = False
    ) -> bool:
        """Submit a prompt to ChatGPT using the chat input."""
        try:
            # On first prompt (or file-only submission), ensure setup is done
            if prompt_number == 1 or (prompt_number == 0 and has_attachments):
                await self.handle_initial_setup()

            frame = await self._get_chatgpt_frame()
            if not frame:
                logger.error("❌ Could not find ChatGPT frame")
                return False

            # Special case: Empty prompt with attachments (file-only submission)
            if has_attachments and not prompt:
                logger.info("📎 Files already attached, submitting without text entry")
                textarea = await self._find_input_element(frame)
                if textarea:
                    await textarea.click()
                    await asyncio.sleep(0.3)
            else:
                # Normal prompt submission with text entry
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
                logger.info("✅ Filled prompt in ChatGPT input")

            await asyncio.sleep(0.5)

            # Click send button
            send_btn = await self._find_send_button(frame)
            if send_btn:
                await send_btn.click(force=True)
                logger.info("✅ Clicked send button")
            else:
                # Fallback: Ctrl+Enter on the focused input. Plain Enter
                # on a contenteditable inserts a newline (does not
                # submit), and on the in-Excel add-in often loses focus
                # to the formula bar.
                if textarea:
                    try:
                        await textarea.focus()
                    except Exception:
                        pass
                    await textarea.press("Control+Enter")
                    logger.info("✅ Pressed Ctrl+Enter to send (fallback)")
                else:
                    logger.error("❌ Could not find textarea or send button")
                    return False

            return True

        except Exception as e:
            logger.error(f"❌ Failed to submit prompt: {e}")
            return False

    async def _find_send_button(self, frame):
        """Find the send/submit button in the ChatGPT frame.

        The send button contains an SVG with an upward arrow path:
        M11.293 5.293a1 1 0 0 1 1.414 0l5 5a1...

        Searches the cached ChatGPT frame first, then every other page
        frame. The in-Excel ChatGPT add-in occasionally splits its input
        and send button across sibling iframes, so single-frame lookup
        misses the button even when the selector is exact.
        """
        selectors = [
            '[data-testid="send-button"]',
            'button[aria-label="Send message"]',
            'button[aria-label="Send"]',
            'button[aria-label*="Send" i]',
        ]

        # Build the frame search order: cached frame first, then every
        # other page frame.
        frames = [frame] + [f for f in self.page.frames if f is not frame]

        for f in frames:
            for sel in selectors:
                try:
                    btn = await f.query_selector(sel)
                    if btn and await btn.is_visible():
                        if f is not frame:
                            logger.info(
                                "Send button found in sibling frame (url=%s)",
                                (f.url or "")[:80],
                            )
                        return btn
                except Exception:
                    continue

            # SVG-path fallback per frame.
            try:
                buttons = await f.query_selector_all("button")
                for btn in buttons:
                    try:
                        if not await btn.is_visible():
                            continue
                        svg = await btn.query_selector("svg")
                        if not svg:
                            continue
                        path_el = await svg.query_selector("path")
                        if not path_el:
                            continue
                        d_attr = await path_el.get_attribute("d") or ""
                        if "11.293" in d_attr and "5.293" in d_attr:
                            if f is not frame:
                                logger.info(
                                    "Send button found via SVG-path in sibling "
                                    "frame (url=%s)",
                                    (f.url or "")[:80],
                                )
                            return btn
                    except Exception:
                        continue
            except Exception:
                continue

        # Diagnostic: log per-frame match counts so the next failure
        # tells us which frame holds (or hides) the button.
        try:
            diag = []
            for f in frames:
                try:
                    n = len(await f.query_selector_all('button[aria-label*="Send" i]'))
                    if n:
                        diag.append(f"url={(f.url or '')[:60]!r} send-like={n}")
                except Exception:
                    continue
            if diag:
                logger.warning("Send-button miss; per-frame matches: %s", diag)
        except Exception:
            pass

        return None

    # ------------------------------------------------------------------
    # File upload
    # ------------------------------------------------------------------

    async def upload_files(self, file_paths: str | list) -> bool:
        """Upload files to ChatGPT via the Plus (+) button."""
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
            frame = await self._get_chatgpt_frame()
            if not frame:
                logger.error("❌ Could not find ChatGPT frame for upload")
                return False

            for i, fp in enumerate(file_paths):
                p = Path(fp)
                logger.info(f"📤 Uploading {i + 1}/{len(file_paths)}: {p.name}")

                # Multi-file uploads (3rd+) sometimes fail because a stale
                # "+" popover from the previous file is still open and the
                # next click toggles it shut instead of opening a fresh one.
                # Press Escape to ensure clean state before each upload.
                if i > 0:
                    try:
                        await self.page.keyboard.press("Escape")
                        await asyncio.sleep(0.2)
                    except Exception:
                        pass

                uploaded = False

                # Try the menuitem path up to 3 times — the popover can fail
                # to render on first click after prior uploads.
                for menuitem_attempt in range(3):
                    plus_btn = await self._find_plus_button(frame)
                    if not plus_btn:
                        if menuitem_attempt == 0:
                            logger.error("❌ Plus (+) button not found for file upload")
                            await self._dump_upload_dom("chatgpt_plus_missing")
                            return False
                        break

                    if menuitem_attempt == 0:
                        logger.info("✅ Found Plus (+) button — clicking...")
                    else:
                        logger.info(
                            f"🔁 Retrying menuitem path (attempt {menuitem_attempt + 1}/3)..."
                        )
                    await plus_btn.click()
                    # Let the popover render. ChatGPT's `+` opens a menu
                    # (Upload files / Skills / Apps) instead of firing the
                    # file chooser directly.
                    await asyncio.sleep(0.4 + 0.3 * menuitem_attempt)

                    upload_item = await self._find_chatgpt_upload_menuitem(frame)
                    if upload_item:
                        if menuitem_attempt == 0:
                            logger.info(
                                "✅ Found 'Upload files' menu item — clicking..."
                            )
                        try:
                            async with self.page.expect_file_chooser(
                                timeout=8000
                            ) as fc:
                                await upload_item.click()
                            chooser = await fc.value
                            await chooser.set_files([fp])
                            logger.info(f"   ✅ Selected: {p.name}")
                            uploaded = True
                            break
                        except Exception as e:
                            logger.warning(
                                f"⚠️ File chooser did not fire after menuitem click: {e}"
                            )
                    # Close any stale popover before retrying.
                    try:
                        await self.page.keyboard.press("Escape")
                        await asyncio.sleep(0.3)
                    except Exception:
                        pass

                if not uploaded:
                    # Fallback 1: older `+`-is-the-chooser behaviour, in case
                    # a future build reverts. Re-click `+` and race the event.
                    try:
                        plus_btn2 = await self._find_plus_button(frame)
                        if plus_btn2:
                            async with self.page.expect_file_chooser(
                                timeout=3000
                            ) as fc:
                                await plus_btn2.click()
                            chooser = await fc.value
                            await chooser.set_files([fp])
                            logger.info(
                                f"   ✅ Selected via legacy +-direct path: {p.name}"
                            )
                            uploaded = True
                    except Exception:
                        pass

                if not uploaded:
                    # Fallback 2: hidden <input type="file">. Scope to the
                    # ChatGPT frame so we don't accidentally upload to one
                    # of Excel's own file inputs (which would succeed
                    # silently but leave the visible UI without a file
                    # chip, blocking the send button later).
                    if await self._set_files_on_hidden_input(fp, scope=frame):
                        logger.info(
                            f"   ✅ Selected via hidden input fallback: {p.name}"
                        )
                        uploaded = True

                if not uploaded:
                    logger.error("❌ File upload failed for this file")
                    await self._dump_upload_dom("chatgpt_upload_failed")
                    return False

                await asyncio.sleep(min(5 + p.stat().st_size / 1024 / 100, 30))

            logger.info(f"✅ All {len(file_paths)} file(s) uploaded!")
            return True

        except Exception as e:
            logger.error(f"❌ Upload failed: {e}")
            import traceback

            logger.error(traceback.format_exc())
            await self._dump_upload_dom("chatgpt_upload_exception")
            return False

    async def _find_chatgpt_upload_menuitem(self, chatgpt_frame):
        """Find the "Upload files" menuitem in the popover that opens after
        clicking `+`.

        The popover may render inside the ChatGPT iframe or in Excel's top-level
        ``ms-Layer`` overlay, so we search both. ARIA role+name is the primary
        anchor; the visible text string is a fallback.
        """
        name_re = re.compile(r"upload\s+file", re.I)
        contexts = [chatgpt_frame, self.page] + list(self.page.frames)

        for ctx in contexts:
            if ctx is None:
                continue
            try:
                loc = ctx.get_by_role("menuitem", name=name_re).first
                if await loc.count() > 0 and await loc.is_visible():
                    return await loc.element_handle()
            except Exception:
                pass

            for label in ("Upload files", "Upload file"):
                try:
                    el = await ctx.query_selector(
                        f'[role="menuitem"]:has-text("{label}")'
                    )
                    if el and await el.is_visible():
                        return el
                except Exception:
                    pass
                try:
                    el = await ctx.query_selector(f'button:has-text("{label}")')
                    if el and await el.is_visible():
                        return el
                except Exception:
                    pass

        return None

    async def _set_files_on_hidden_input(self, file_path, scope=None) -> bool:
        """Find any ``<input type="file">`` (including hidden) and set files.

        Bypasses the native file-chooser dialog entirely — some React chat UIs
        render a hidden input that the visible `+` button proxies to.

        ``scope`` (optional) restricts the search to the given frame and its
        child frames. When omitted, every page frame is searched. Scoping is
        important on multi-frame pages (e.g. the ChatGPT add-in inside Excel
        Online) — uploading to the wrong frame's input succeeds silently but
        leaves the visible UI unaware of the file, which then blocks the
        send button.
        """
        if scope is not None:
            scoped = [scope]
            scoped.extend(
                f
                for f in self.page.frames
                if f is not scope and f.parent_frame is scope
            )
            ctxs = scoped
        else:
            ctxs = list(self.page.frames) + [self.page]
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

    async def _find_plus_button(self, frame):
        """Find the Plus (+) file attachment button in the ChatGPT frame.

        The button contains an SVG with a plus/cross path:
        M12 3.59998...C12.5891 3.59998 13.0666 4.07754 13.0666 4.66664V10.9333H19.3333...
        """
        # Strategy 1: aria-label
        for sel in [
            'button[aria-label*="attach" i]',
            'button[aria-label*="upload" i]',
            'button[aria-label*="file" i]',
            'button[aria-label*="add" i]',
        ]:
            try:
                btn = await frame.query_selector(sel)
                if btn and await btn.is_visible():
                    return btn
            except Exception:
                continue

        # Strategy 2: find button with plus SVG path signature
        try:
            buttons = await frame.query_selector_all("button")
            for btn in buttons:
                try:
                    if not await btn.is_visible():
                        continue
                    svg = await btn.query_selector("svg")
                    if not svg:
                        continue
                    path_el = await svg.query_selector("path")
                    if not path_el:
                        continue
                    d_attr = await path_el.get_attribute("d") or ""
                    # The plus icon has "M12 3.59998" in its path
                    if "M12 3.59998" in d_attr or "12 3.6" in d_attr:
                        return btn
                except Exception:
                    continue
        except Exception:
            pass

        return None

    # ------------------------------------------------------------------
    # Completion detection
    # ------------------------------------------------------------------

    async def _has_stop_button(self, frame) -> bool:
        """Check if the Stop button is visible (ChatGPT is processing).

        The stop button contains a square SVG:
        M6 8C6 6.89543 6.89543 6 8 6H16C17.1046 6 18 6.89543 18 8V16...
        wrapped in <span class="_ButtonInner_1jdeq_4">
        """
        try:
            # Strategy 1: aria-label
            for sel in [
                'button[aria-label="Stop"]',
                'button[aria-label*="stop" i]',
                'button[data-testid="stop-button"]',
            ]:
                btn = await frame.query_selector(sel)
                if btn and await btn.is_visible():
                    return True

            # Strategy 2: find button with _ButtonInner span containing square SVG
            buttons = await frame.query_selector_all("button")
            for btn in buttons:
                try:
                    if not await btn.is_visible():
                        continue
                    # Check for the ButtonInner span pattern
                    inner = await btn.query_selector("span")
                    if not inner:
                        continue
                    svg = await inner.query_selector("svg")
                    if not svg:
                        continue
                    path_el = await svg.query_selector("path")
                    if not path_el:
                        continue
                    d_attr = await path_el.get_attribute("d") or ""
                    # The stop-square path starts with "M6 8C6 6.89543"
                    if "M6 8C6 6.89543" in d_attr:
                        return True
                except Exception:
                    continue

        except Exception:
            pass
        return False

    async def _get_response_count(self, frame) -> int:
        """Get current number of ChatGPT responses.

        ChatGPT responses may use data-message-author-role="assistant"
        or similar DOM structure. Falls back to generic detection.
        """
        try:
            # Try several selectors for ChatGPT response containers
            for sel in [
                '[data-message-author-role="assistant"]',
                "article",
                'div[data-testid*="conversation-turn"]',
            ]:
                responses = await frame.query_selector_all(sel)
                if responses:
                    return len(responses)
            return 0
        except Exception:
            return 0

    async def wait_for_completion(
        self, prompt_number: int, initial_counts: dict = None
    ) -> bool:
        """Wait for ChatGPT to complete by checking Stop button presence.

        Logic (mirrors Claude):
        1. Wait for Stop button to APPEAR (ChatGPT started processing)
        2. Wait for Stop button to DISAPPEAR (ChatGPT finished)
        3. Brief stabilization to ensure ChatGPT is fully done
        """
        agent_config = self.get_config_section()
        max_wait = agent_config.get("max_wait_per_prompt_seconds", 900)

        # Get initial response count for logging
        self._chatgpt_frame = None
        frame = await self._get_chatgpt_frame()
        initial_response_count = await self._get_response_count(frame) if frame else 0

        logger.info(
            f"⏳ Waiting for prompt #{prompt_number} "
            f"(starting with {initial_response_count} responses)..."
        )

        start_time = time.monotonic()
        check_interval = 1  # Check every second for responsiveness
        saw_stop_button = False

        while (time.monotonic() - start_time) < max_wait:
            if self.shutdown_event and self.shutdown_event.is_set():
                return False

            await asyncio.sleep(check_interval)
            elapsed = int(time.monotonic() - start_time)

            # Refresh frame reference
            self._chatgpt_frame = None
            frame = await self._get_chatgpt_frame()
            if not frame:
                continue

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
                    # ChatGPT is processing
                    saw_stop_button = True
                    if elapsed % 10 == 0:
                        logger.info(
                            f"   [{elapsed}s] ChatGPT processing... "
                            f"(responses: {current_response_count})"
                        )
                    continue

                # Stop button not visible
                if not saw_stop_button:
                    # Haven't seen ChatGPT start yet — keep waiting
                    if elapsed % 10 == 0:
                        logger.info(
                            f"   [{elapsed}s] Waiting for ChatGPT to start... "
                            f"(responses: {current_response_count})"
                        )
                    continue

                # Saw stop button, now it's gone — ChatGPT finished!
                logger.info(
                    f"   [{elapsed}s] Stop button gone, verifying completion..."
                )
                await asyncio.sleep(3)

                # Re-check that stop button is still gone
                self._chatgpt_frame = None
                frame = await self._get_chatgpt_frame()
                if frame:
                    still_stopped = not await self._has_stop_button(frame)
                    final_count = await self._get_response_count(frame)

                    if not still_stopped:
                        # Stop button reappeared — ChatGPT started again
                        logger.info(f"   [{elapsed}s] ChatGPT resumed processing...")
                        continue

                    # Stop button gone and stayed gone — ChatGPT finished.
                    # Trust the stop button signal. Response count selectors
                    # may not match the add-in iframe DOM, so don't gate on it.
                    logger.info(
                        f"✅ Prompt #{prompt_number} completed! "
                        f"(responses: {final_count}, was {initial_response_count})"
                    )
                    return True

            except Exception as e:
                logger.debug(f"   [{elapsed}s] Check error: {e}")

        logger.error(f"❌ Timeout waiting for prompt #{prompt_number}")
        return False

    # ------------------------------------------------------------------
    # Feedback buttons (for base class completion detection fallback)
    # ------------------------------------------------------------------

    async def get_button_count(self) -> dict:
        """Count ChatGPT's feedback buttons (thumbs up/down)."""
        try:
            frame = await self._get_chatgpt_frame()
            if not frame:
                return {"upvote": 0, "downvote": 0}

            upvote_count = 0
            downvote_count = 0

            up_selectors = [
                'button[aria-label*="thumbs up" i]',
                'button[aria-label*="like" i]',
                '[data-testid*="thumbs-up"]',
                '[data-testid*="upvote"]',
            ]
            down_selectors = [
                'button[aria-label*="thumbs down" i]',
                'button[aria-label*="dislike" i]',
                '[data-testid*="thumbs-down"]',
                '[data-testid*="downvote"]',
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
            logger.debug(f"Error counting ChatGPT buttons: {e}")
            return {"upvote": 0, "downvote": 0}
