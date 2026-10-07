"""
ChatGPT Web Agent - Automate interactions with chatgpt.com web interface.

Uses Playwright to:
1. Navigate to a ChatGPT project
2. Enable agent mode + extended thinking
3. Upload files and submit prompts
4. Wait for response completion
5. Download Excel artifacts
"""

import asyncio
import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import unquote

from claude_web_agent.web_agent import WebAgent, WebAgentState, ConversationMessage
from claude_web_agent.dom_diagnostics import dump_final_message_dom

logger = logging.getLogger(__name__)


class ChatGPTWebAgent(WebAgent):
    """Agent for automating ChatGPT web interface."""

    CHATGPT_BASE_URL = "https://chatgpt.com"

    SELECTORS = {
        # Chat input — visible ProseMirror contenteditable div
        "chat_input": 'div.ProseMirror[contenteditable="true"]',
        # Hidden textarea (for page-type detection; display:none so not usable for visibility checks)
        "textarea_project": 'textarea[placeholder*="New chat in"]',
        "textarea_conversation": 'textarea[placeholder="Ask anything"]',
        # Buttons
        # 2026-09-22 composer: the label is a bare "Send" and the testid is
        # gone; the older wordings stay first so the other generation is
        # untouched. Never match "Send" as a substring of a longer label.
        "send_button": 'button:has-text("Send prompt"), button[aria-label="Send prompt"], [data-testid="send-button"], button[aria-label="Send"]',
        "plus_menu_button": '[data-testid="composer-plus-btn"]',
        "stop_button": 'button:has-text("Stop")',
        "answer_now_button": 'button:has-text("Answer now")',
        # File upload
        "add_files_menuitem": '[role="menuitem"]:has-text("Add photos & files")',
        # State detection
        "login_button": 'button:has-text("Log in")',
        "thinking_active": 'button:has-text("Pro thinking")',
        "thinking_complete": 'button:text-matches("Thought for \\d")',
        # Model — pill in the composer toolbar that opens the dropdown
        "model_selector": 'button.__composer-pill[aria-haspopup="menu"]',
    }

    # Phrases that mean the session hit a usage/plan limit. Without this
    # the agent sits in its wait loop for the full max_wait_per_prompt
    # (hours) after a limit screen appears. Phrases must be specific — the
    # financial documents being generated legitimately contain "limit".
    USAGE_LIMIT_PHRASES = (
        "You've hit your limit",
        "You’ve hit your limit",
        "You've reached your limit",
        "You’ve reached your limit",
        "out of usage credits",
        "usage cap reached",
        "limit resets",
        # Live 2026-09-10 (Pro account, Ultra effort, full run): "You've
        # reached your usage limit. Add credits to continue, or try again at
        # Sep 14, 2026, 9:22 PM." — "usage" sits between "your" and "limit",
        # so the two phrases above did not match and the lane waited out
        # the accept window instead of failing fast.
        "reached your usage limit",
        "hit your usage limit",
        "Add credits to continue",
    )

    # Server-side stream failures. Generation stops, no artifact is
    # produced, and the turn is left with a Retry button that resumes it.
    STREAM_ERROR_PHRASES = (
        "Error in message stream",
        "Something went wrong while generating",
        "A network error occurred",
    )

    # A turn that died with NO way to resume it. Live 2026-09-21, chat mode at
    # Pro, ~80 min into task 14 (DailyCash): the assistant turn collapsed to a
    # bare "Thinking failed" chip — no text, no file, no Retry button — so
    # neither stream-error recovery nor "Continue" applies. The wait loop
    # used to sit on it until the dead-page watchdog fired 30 minutes later.
    # Matched against the WHOLE response text, never as a substring: a model
    # may well write the words in prose. Chat mode only — work-mode turns go
    # quiet for minutes mid-build and have their own end-of-turn handling.
    TURN_FAILED_TEXTS = ("Thinking failed",)
    TURN_FAILED_GRACE_SEC = 90

    def __init__(self, page, config: dict, shutdown_event=None, completion_logger=None):
        super().__init__(page, config, shutdown_event, completion_logger)
        self.agent_config = config.get("chatgpt_web", {})
        self.project_id = self.agent_config.get("project_id", "")
        self.project_slug = self.agent_config.get("project_slug", "")
        self.max_wait_per_prompt = self.agent_config.get(
            "max_wait_per_prompt_seconds", 1800
        )
        self.check_interval = self.agent_config.get("check_interval_seconds", 3)
        # Tracks how many response articles existed BEFORE the first prompt,
        # so download_all_artifacts searches all articles from the conversation.
        # Set once before the first prompt; not overwritten by subsequent prompts.
        self._baseline_article_count = 0
        self._baseline_set = False

    @property
    def project_url(self) -> str:
        # null/empty project_id = no project scope: the homepage is a new
        # chat (mirrors claude_web's project_id fallback). Conversations
        # then land in the account's main history, not in a project.
        if not self.project_id:
            return self.CHATGPT_BASE_URL
        slug_part = f"-{self.project_slug}" if self.project_slug else ""
        return f"{self.CHATGPT_BASE_URL}/g/g-p-{self.project_id}{slug_part}/project"

    async def navigate_to_new_chat(self) -> bool:
        """Navigate to the ChatGPT project page (which is a new chat).

        Returns True if the page loaded (even if login is required).
        The engine's auth-wait loop handles the login case.
        """
        try:
            # Reset baseline for the new conversation so download_all_artifacts
            # will search all articles from this chat (not carry over from prior task).
            self._baseline_article_count = 0
            self._baseline_set = False

            if self.project_id:
                logger.info(
                    f"Navigating to ChatGPT project: {self.project_url}"
                )
            else:
                logger.warning(
                    "chatgpt_web.project_id is empty — starting at the "
                    "chatgpt.com homepage (no project scope; conversations "
                    "land in the account's main chat history)"
                )

            # NOTE: Do NOT call set_viewport_size on CDP pages — it uses
            # Emulation.setDeviceMetricsOverride which crashes ChatGPT tabs.
            await self.page.goto(
                self.project_url, wait_until="domcontentloaded", timeout=60000
            )

            # Debug: log where we ended up
            current_url = self.page.url
            page_title = await self.page.title()
            logger.info(f"Page loaded — URL: {current_url}")
            logger.info(f"Page loaded — Title: {page_title}")

            # Check if we were redirected to an auth page (not on chatgpt.com)
            if "chatgpt.com" not in current_url:
                logger.warning(f"Redirected to auth page: {current_url}")
                return True  # Let the engine's auth-wait loop handle it

            # Check if we need to authenticate (login button on chatgpt.com)
            state = await self.get_state()
            if state == WebAgentState.AUTH_REQUIRED:
                logger.warning("Authentication required - please log in manually")
                return True  # Let the engine's auth-wait loop handle it

            logger.info(f"Page state: {state.value}")

            # Wait for chat input — the React SPA may need time to hydrate.
            # Try ProseMirror editor first, then fall back to paragraph placeholder.
            # Use wait_for(state="attached") since is_visible() is unreliable on CDP.
            chat_input = self.page.locator(
                'div.ProseMirror[contenteditable="true"], '
                'p[data-placeholder*="New chat"]'
            )

            for attempt_label in ("initial load", "after reload"):
                try:
                    await chat_input.first.wait_for(state="attached", timeout=15000)
                    logger.info(f"ChatGPT chat input visible ({attempt_label})")
                    # Assert Chat/Work mode BEFORE any uploads — toggling
                    # later can clear attached files, and the toggle's
                    # selection persists across sessions.
                    if not await self.ensure_mode():
                        logger.error("Failed to set Chat/Work mode")
                        return False
                    return True
                except Exception:
                    if attempt_label == "initial load":
                        logger.info("Chat input not visible yet — reloading page")
                        await self.page.reload(
                            wait_until="domcontentloaded", timeout=30000
                        )
                        await self.page.wait_for_timeout(3000)

            # Last resort: dump page content for debugging
            try:
                body_text = await self.page.locator("body").inner_text()
                logger.error(
                    f"Could not find chat input. Page text (first 500 chars): {body_text[:500]}"
                )
            except Exception:
                logger.error("Could not find chat input on ChatGPT page")
            return False
        except Exception as e:
            logger.error(f"Navigation failed: {e}")
            return False

    async def get_state(self) -> WebAgentState:
        """Detect ChatGPT page state.

        Uses JS evaluate instead of Playwright ``is_visible()`` because the
        latter is unreliable on Chrome CDP connections.
        """
        try:
            current_url = self.page.url
            if "chatgpt.com" not in current_url:
                return WebAgentState.AUTH_REQUIRED

            state_info = await self.page.evaluate(
                """() => {
                const btns = Array.from(document.querySelectorAll('button'));
                const hasLogin = btns.some(b => b.textContent.trim() === 'Log in');
                const hasStop = btns.some(b => b.textContent.trim() === 'Stop');
                const hasThinking = btns.some(b => b.textContent.includes('Pro thinking'));
                const hasInput = !!document.querySelector(
                    'div.ProseMirror[contenteditable="true"], p[data-placeholder]'
                );
                return { hasLogin, hasStop, hasThinking, hasInput };
            }"""
            )

            if state_info["hasLogin"]:
                return WebAgentState.AUTH_REQUIRED

            if state_info["hasStop"] or state_info["hasThinking"]:
                return WebAgentState.RUNNING
            if state_info["hasInput"]:
                return WebAgentState.READY

            return WebAgentState.UNKNOWN
        except Exception as e:
            logger.error(f"State detection error: {e}")
            return WebAgentState.ERROR

    async def _check_button_text(self, text: str) -> bool:
        """CDP-safe check: does any button on the page contain *text*?

        Checks both ``textContent`` and ``aria-label`` because ChatGPT uses
        aria-labels like "Agent, click to remove" while the visible text is
        just "Agent".  Playwright's ``is_visible()`` is unreliable on CDP
        connections, so we query the DOM directly via JavaScript.
        """
        try:
            return await self.page.evaluate(
                "(t) => Array.from(document.querySelectorAll('button'))"
                ".some(b => b.textContent.includes(t)"
                " || (b.getAttribute('aria-label') || '').includes(t))",
                text,
            )
        except Exception:
            return False

    # chatgpt_web.model config values → labels in the model submenu of the
    # composer pill (verified live 2026-07-21). Unknown values fall back to
    # underscores→spaces/dashes best-effort matching.
    MODEL_LABELS = {
        "gpt_6_astra": "GPT-6 Astra",       # verified live 2026-09-10
        "gpt_5_6_sol": "GPT-5.6 Sol",
        "gpt_5_5": "GPT-5.5",
        "gpt_5_4": "GPT-5.4",
        "gpt_5_3": "GPT-5.3",
        "o3": "o3",
    }

    # chatgpt_web.intelligence config values → top-level radio labels in the
    # pill menu. NOTE the pill/menu axis is "Intelligence" (reasoning
    # effort); the model itself lives in a nested submenu (the old flat
    # model-switcher-* testids no longer exist).
    INTELLIGENCE_LABELS = {
        "instant": "Instant",
        "medium": "Medium",
        "high": "High",
        "xhigh": "Extra High",
        "pro": "Pro",
    }

    # One-axis chatgpt_web.model values. The UI has since split model and
    # intelligence into separate pickers, but these three name whole cohorts
    # in task_attempts (`chatgpt_instant`, `chatgpt_thinking`,
    # `chatgpt_web_pro` — see infra/configs/agent_identity.py, whose tables
    # are append-only), so a config must still be able to reproduce one.
    # The value carries no model, only an intelligence level; route it there.
    ONE_AXIS_MODEL_TO_INTELLIGENCE = {
        "instant": "instant",
        "thinking": "high",
        "pro": "pro",
    }

    # chatgpt_web.mode — the Chat/Work toggle (verified live 2026-07-21):
    # Radix button[role=radio] pair with aria-checked/data-state. The
    # selection persists across sessions, so it is asserted every task in
    # both directions.
    MODE_VALUES = ("chat", "work")

    # How long ensure_mode waits for that toggle to hydrate. The composer
    # toolbar renders AFTER the ProseMirror input that navigate_to_new_chat
    # waits on — measured ~1s behind it on a warm chatgpt.com homepage
    # (2026-08-21), and the gap is invisible on a project URL only because
    # its redirect burns that second first. A single read races it.
    MODE_TOGGLE_WAIT_SEC = 8.0
    MODE_TOGGLE_POLL_SEC = 0.5

    # Work-mode picker. TWO generations coexist (ensure_work_settings
    # detects per-run, so an A/B flip back keeps working):
    #  - Advanced-rows (pre-2026-08-28): pill menu → "Advanced" → Model /
    #    Effort / Speed rows with flyouts. The rows carry no roles/testids —
    #    they're div.__menu-item, text-anchored. After picking an off-slider
    #    effort (Ultra), the slider DECOUPLES (shows "Reset to default");
    #    state must be verified from the rows + pill text, never the slider.
    #  - Slider generation (verified live 2026-08-28): see the _SLIDER_*
    #    block below — the Advanced rows are gone entirely.
    WORK_MODEL_LABELS = {
        "gpt_6_astra": "GPT-6 Astra",       # verified live 2026-09-10 (slider picker)
        "gpt_5_6_sol": "GPT-5.6 Sol",
        "gpt_5_6_terra": "GPT-5.6 Terra",
        "gpt_5_6_luna": "GPT-5.6 Luna",
        "gpt_5_5": "GPT-5.5",
    }
    WORK_EFFORT_LABELS = {
        "light": "Light",
        "medium": "Medium",
        "high": "High",
        "xhigh": "Extra High",
        "max": "Max",
        "ultra": "Ultra",
    }
    WORK_SPEED_LABELS = {
        "standard": "Standard",
        "fast": "Fast",
    }

    # Slider-generation work picker (verified live 2026-08-28). The pill
    # menu is a two-view widget (data-view simple|advanced):
    #  - model:  menuitemradio list behind a [role=menuitem]
    #            aria-label="Select model" view toggle. Radios read
    #            "5.6 Sol" (no "GPT-" prefix); "Default" is a
    #            recommended-mix pseudo-model, NOT an explicit selection.
    #  - effort: a "Power" slider row (aria-label="Power"), keyboard
    #            Left/Right; live state is announced through
    #            aria-describedby as "<Label>, <n> of <total>.".
    #  - speed:  menuitemcheckbox aria-label="Enable fast mode"
    #            (unchecked = Standard, checked = Fast).
    # Ultra exists ONLY under an explicit model: with "Default" the slider
    # tops out at Extra High ("5 of 5"); selecting 5.6 Sol re-ranges it to
    # 6 stops ("Extra High, 4 of 6" → Max → Ultra). So the model radio
    # must be clicked even when the toggle already displays the right
    # name — the toggle shows what "Default" RESOLVES to, not that an
    # explicit model is pinned. Anchored on aria-labels, not the hashed
    # utility classes.
    _SLIDER_TOGGLE_SEL = (
        '[role="menu"][data-state="open"] '
        '[role="menuitem"][aria-label="Select model"]'
    )
    _SLIDER_POWER_SEL = (
        '[role="menu"][data-state="open"] [role="menuitem"][aria-label="Power"]'
    )
    _SLIDER_FAST_SEL = (
        '[role="menu"][data-state="open"] '
        '[role="menuitemcheckbox"][aria-label="Enable fast mode"]'
    )
    # UI order, low → high; drives the arrow-walk direction. Positions are
    # NOT stable across models ("Extra High" is 5 of 5 under Default but
    # 4 of 6 under 5.6 Sol), so all comparisons go by label.
    # CAUTION: REOPENING the pill menu snaps a committed Ultra back to Max
    # (observed live 2026-08-28: close at Ultra → pill "5.6 Sol Ultra";
    # open + Escape → pill "5.6 Sol Max"). Two consequences: (a) every
    # ensure_work_settings call must re-walk the slider — "already Ultra"
    # will essentially never be read back after an open; (b) nothing may
    # reopen this menu between setting the effort and sending the prompt,
    # or the send goes out at Max. The engine's flow (set → close → pill
    # check → send) respects this.
    _SLIDER_EFFORT_LADDER = ("Light", "Medium", "High", "Extra High", "Max", "Ultra")

    # Slider-generation CHAT picker (probed live 2026-09-21, two Pro
    # accounts). Chat mode now renders the same widget family as work mode
    # — the "Select model" toggle, the "Power" slider row and a
    # menuitemradio model list — and the model/Effort submenus that
    # ensure_model_and_intelligence drove are gone. What differs from work:
    #  - models:  "Latest" / "GPT-5.6 Sol" / "GPT-5.5". The GPT-6 generation
    #             is offered ONLY as "Latest"; work mode's "GPT-6 Astra"
    #             radio does not exist here and the word "Astra" appears
    #             nowhere in chat. So the chat cohort is keyed `gpt_6`, not
    #             `gpt_6_astra`, and its radio is "Latest".
    #  - ladder:  Instant / Medium / High / Extra High / Pro (5 stops).
    #  - no fast-mode checkbox.
    #  - the top stop HOLDS across a menu reopen (work's Ultra snaps back).
    #  - pill:    model token + tier with no separator — "6Pro" under Latest
    #             at Pro. Under Latest the token is what "Latest" currently
    #             resolves to, so the pill is the only on-screen evidence of
    #             the model; it is matched EXACTLY, and a re-pointed "Latest"
    #             ("6.1Pro") stops the lane rather than mixing models under
    #             one label. (At Instant the pill drops the token and reads
    #             just "Instant" — only the pinned tier's form is relied on.)
    # These tables are read by the chat slider branch only; work mode never
    # touches them.
    CHAT_SLIDER_MODEL_RADIOS = {
        "gpt_6": "Latest",
    }
    CHAT_SLIDER_PILL_TOKENS = {
        "gpt_6": "6",
    }
    _CHAT_SLIDER_LADDER = ("Instant", "Medium", "High", "Extra High", "Pro")

    # JS-dispatched hover/click — skip pointer-events actionability checks
    # (submenu flyouts overlay their sibling menu items).
    _JS_HOVER = (
        "el => ['pointerover','pointerenter','mouseover','mouseenter','mousemove']"
        ".forEach(t => el.dispatchEvent("
        "new MouseEvent(t, {bubbles: true, cancelable: true, view: window})))"
    )
    _JS_CLICK = "el => el.click()"

    # Full synthetic pointer sequence. React's work-picker rows ignore a bare
    # el.click(); they need the pointer/mouse pair. Verified live 2026-08-02
    # selecting Effort -> Ultra (pill confirmed "5.6 Sol Ultra").
    _JS_POINTER_CLICK = """el => {
        const r = el.getBoundingClientRect();
        const opts = {bubbles: true, cancelable: true, view: window,
                      clientX: r.left + r.width / 2,
                      clientY: r.top + r.height / 2};
        for (const t of ['pointerdown','mousedown','pointerup','mouseup','click']) {
            el.dispatchEvent(new MouseEvent(t, opts));
        }
    }"""

    def _resolve_targets(self) -> tuple:
        """Resolve (model_label, intelligence_label) from config.

        Handles the one-axis ``model: pro|thinking|instant`` form by routing
        it to the intelligence axis (with a warning) when no explicit
        ``intelligence`` is set.
        """
        model = self.agent_config.get("model")
        intelligence = self.agent_config.get("intelligence")

        # Only the chat path calls this resolver, so effort/speed in the
        # config can only be a misconfiguration (they are work-mode knobs).
        # Warn instead of silently ignoring — a chat run configured with
        # effort: ultra ran at the session-default intelligence, unnoticed
        # (observed 2026-08-12).
        if self.agent_config.get("effort") or self.agent_config.get("speed"):
            logger.warning(
                "chatgpt_web.effort/speed apply only to WORK mode and are "
                "ignored in chat mode. Chat's Effort submenu (the UI "
                "renamed intelligence to 'Effort' 2026-08-12) is driven by "
                "chatgpt_web.intelligence (instant|medium|high|xhigh|pro)."
            )

        if model and model.lower() in self.ONE_AXIS_MODEL_TO_INTELLIGENCE:
            routed = self.ONE_AXIS_MODEL_TO_INTELLIGENCE[model.lower()]
            logger.warning(
                "chatgpt_web.model=%r names an intelligence level, not a "
                "model — the UI splits the two. Routing it to "
                "intelligence=%r; unless you are reproducing that cohort, "
                "set chatgpt_web.model to a model (e.g. gpt_5_6_sol) and "
                "chatgpt_web.intelligence explicitly.",
                model,
                routed,
            )
            if not intelligence:
                intelligence = routed
            model = None

        model_label = None
        if model:
            model_label = self.MODEL_LABELS.get(
                model.lower(), model.replace("_", " ").strip()
            )

        intel_label = None
        if intelligence:
            intel_label = self.INTELLIGENCE_LABELS.get(intelligence.lower())
            if intel_label is None:
                logger.error(
                    "Unknown chatgpt_web.intelligence %r. Valid: %s",
                    intelligence,
                    ", ".join(self.INTELLIGENCE_LABELS),
                )
                raise ValueError(f"unknown intelligence {intelligence!r}")

        return model_label, intel_label

    async def _get_pill(self):
        """The composer pill that opens the Intelligence/model menu.

        Its visible text is the CURRENT state, but WHAT it shows differs by
        UI generation: the older pill reads model+tier glued together
        ("6Pro"), the 2026-09-22 composer reads "Thinking effortPro" — the
        tier only. Callers that need the model must read it inside the menu
        (see _chat_menu_model_text), never from this button's text.
        """
        for sel in (
            self.SELECTORS["model_selector"],
            'button[aria-label="Select ChatGPT model"]',
            'form button[aria-haspopup="menu"]',
        ):
            pill = await self.page.query_selector(sel)
            if pill:
                return pill
        return None

    async def _open_pill_menu(self) -> bool:
        # Already open? (Clicking the pill again would toggle it closed.)
        menu = await self.page.query_selector('[role="menu"][data-state="open"]')
        if menu:
            return True
        pill = await self._get_pill()
        if pill is None:
            logger.error("ChatGPT composer pill not found")
            return False
        for _ in range(2):
            try:
                await pill.click()
            except Exception:
                await pill.evaluate(self._JS_CLICK)
            await asyncio.sleep(1.2)
            menu = await self.page.query_selector('[role="menu"][data-state="open"]')
            if menu:
                return True
        logger.error("ChatGPT pill menu did not open")
        return False

    async def _ensure_work_rows_visible(self) -> bool:
        """Make sure the Advanced rows (Model/Effort/Speed) are on screen,
        opening the pill menu and expanding the Advanced section as needed."""
        if await self._find_work_row("Model") is not None:
            return True
        if not await self._open_pill_menu():
            return False
        if await self._find_work_row("Model") is not None:
            return True
        adv = await self._find_work_row("Advanced")
        if adv is None:
            h = await self.page.evaluate_handle(
                """() => Array.from(document.querySelectorAll('*'))
                    .filter(el => el.getClientRects().length > 0 &&
                            el.children.length === 0 &&
                            (el.textContent || '').trim() === 'Advanced')
                    .map(el => el.closest('button, .__menu-item, [role]') || el)[0]
                    || null"""
            )
            adv = h.as_element()
        if adv is not None:
            await self._click_el(adv)
            await asyncio.sleep(1.2)
        return await self._find_work_row("Model") is not None

    async def _close_pill_menu(self) -> None:
        try:
            await self.page.keyboard.press("Escape")
            await asyncio.sleep(0.3)
            await self.page.keyboard.press("Escape")
            await asyncio.sleep(0.3)
        except Exception:
            pass

    async def _click_visible_radio(self, label: str) -> bool:
        """Click the visible menuitemradio whose text starts with ``label``.

        startswith, not equality: rows append hints (e.g. "Instant5.5",
        "GPT-5.4Leaving on July 23")."""
        radios = await self.page.query_selector_all('[role="menuitemradio"]')
        for r in radios:
            try:
                if not await r.is_visible():
                    continue
                text = ((await r.text_content()) or "").strip()
                if text.lower().startswith(label.lower()):
                    await r.evaluate(self._JS_CLICK)
                    await asyncio.sleep(1.0)
                    return True
            except Exception:
                continue
        return False

    async def _model_submenu_parent(self):
        """The menuitem that carries the CURRENT model name and opens the
        model submenu (role=menuitem + aria-haspopup=menu).

        2026-08-12 (observed live): the chat pill menu gained a second
        submenu trigger — "Effort", holding the former top-level
        intelligence radios — so prefer the row captioned "Model" and
        fall back to first-visible for older single-trigger builds."""
        items = await self.page.query_selector_all(
            '[role="menuitem"][aria-haspopup="menu"]'
        )
        first_visible = None
        for it in items:
            if await it.is_visible():
                if first_visible is None:
                    first_visible = it
                if "model" in (((await it.text_content()) or "").lower()):
                    return it
        return first_visible

    async def _effort_submenu_parent(self):
        """The chat-mode "Effort" submenu trigger (new 2026-08-12): the
        reasoning tiers (Instant..Pro) moved from top-level radios into
        this submenu. Returns None on older builds — callers then fall
        back to the top-level radio path."""
        items = await self.page.query_selector_all(
            '[role="menuitem"][aria-haspopup="menu"]'
        )
        for it in items:
            if await it.is_visible():
                if "effort" in (((await it.text_content()) or "").lower()):
                    return it
        return None

    async def _open_model_submenu(self, parent) -> bool:
        """Expand the model flyout. Radix submenus need TRUSTED pointer
        events — synthetic JS hover does NOT open them (verified live
        2026-07-21) — so try a real hover first and fall back to JS
        hover/click for the occluded case. Open state is read from the
        parent's data-state/aria-expanded."""
        for attempt in ("hover", "js_hover", "js_click"):
            try:
                if attempt == "hover":
                    await parent.hover(timeout=3000)
                elif attempt == "js_hover":
                    await parent.evaluate(self._JS_HOVER)
                else:
                    await parent.evaluate(self._JS_CLICK)
            except Exception:
                pass
            await asyncio.sleep(1.3)
            try:
                if (await parent.get_attribute("data-state")) == "open":
                    return True
                if (await parent.get_attribute("aria-expanded")) == "true":
                    return True
            except Exception:
                continue
        return False

    @staticmethod
    def _pill_intel_matches(pill_text: str, intel_label: str) -> bool:
        """True if the chat-mode pill reflects the wanted intelligence.

        The chat pill concatenates the model and intelligence labels with no
        separator (e.g. GPT-5.5 at Pro reads "5.5Pro"), so an exact match
        against the bare intelligence label never succeeds. Intelligence is
        the trailing component, so match on suffix — with the one guard that
        "High" must not spuriously match a pill ending in "Extra High".
        """
        p = (pill_text or "").strip().lower()
        lab = (intel_label or "").strip().lower()
        if not p or not lab:
            return False
        if not p.endswith(lab):
            return False
        if lab == "high" and p.endswith("extra high"):
            return False
        return True

    @staticmethod
    def _model_row_shows(row_text: str, model_label: str) -> bool:
        """True iff the submenu-parent row is showing model_label.

        Two things have to be normalized away before comparing.

        The row's text_content() concatenates its caption and value, so it
        reads 'ModelGPT-5.6 Sol' rather than the bare value (observed live
        2026-08-12) — strip the caption.

        The value is then written in EITHER form: the menu lists the option
        as 'GPT-5.5', but once that option is selected the row and the pill
        both shorten it to '5.5' ('Model5.5', pill '5.5Instant') — verified
        live 2026-08-21. Selecting the model therefore made verification
        fail against the very label that had just been clicked. Compare with
        the 'GPT-' prefix dropped from both sides so either rendering
        matches.
        """
        def norm(s: str) -> str:
            s = s.strip().lower()
            if s.startswith("gpt-"):
                s = s[len("gpt-"):]
            return s.lstrip()

        t = row_text.strip()
        if t.lower().startswith("model"):
            t = t[len("model"):].lstrip()
        return norm(t).startswith(norm(model_label))

    async def ensure_model_and_intelligence(self) -> bool:
        """Select the configured model (submenu) and intelligence (radios).

        Fails loudly on unknown labels or verification mismatch — silently
        benchmarking the wrong model is far more expensive than a crash.
        Both settings verified: intelligence against the pill text, model
        against the submenu-parent row text.
        """
        try:
            model_label, intel_label = self._resolve_targets()
        except ValueError:
            return False

        if not model_label and not intel_label:
            logger.info("No ChatGPT model/intelligence configured — using defaults")
            return True

        # Slider-generation chat picker (2026-09-21) — see the
        # CHAT_SLIDER_* block. Detected per run, like work mode's two
        # generations; on the submenu generation the menu is closed again
        # and the code below runs exactly as it always has.
        try:
            if not await self._open_pill_menu_patient():
                return False
            if await self._slider_picker_present():
                model = self.agent_config.get("model")
                model = model.lower() if model else None
                if model in self.ONE_AXIS_MODEL_TO_INTELLIGENCE:
                    model = None  # routed to the tier by _resolve_targets
                return await self._ensure_chat_slider(
                    model, model_label, intel_label
                )
            await self._close_pill_menu()
        except Exception as e:
            logger.error(f"Error selecting model/intelligence (slider): {e}")
            await self._close_pill_menu()
            return False

        try:
            # ---- Model (nested submenu) ----
            if model_label:
                if not await self._open_pill_menu():
                    return False
                parent = await self._model_submenu_parent()
                if parent is None:
                    logger.error("Model submenu parent not found — UI drift?")
                    await self._close_pill_menu()
                    return False
                current = ((await parent.text_content()) or "").strip()
                if self._model_row_shows(current, model_label):
                    logger.info(f"Model already {current!r}")
                    await self._close_pill_menu()
                else:
                    if not await self._open_model_submenu(parent):
                        logger.error("Model submenu did not expand")
                        await self._close_pill_menu()
                        return False
                    if not await self._click_visible_radio(model_label):
                        logger.error(
                            f"Model {model_label!r} not found in submenu"
                        )
                        await self._close_pill_menu()
                        return False
                    await self._close_pill_menu()
                    # Verify: reopen, submenu parent row shows current model.
                    if not await self._open_pill_menu():
                        return False
                    parent = await self._model_submenu_parent()
                    now = (
                        ((await parent.text_content()) or "").strip()
                        if parent
                        else ""
                    )
                    await self._close_pill_menu()
                    if not self._model_row_shows(now, model_label):
                        logger.error(
                            f"Model verification failed: wanted {model_label!r}, "
                            f"menu shows {now!r}"
                        )
                        return False
                    logger.info(f"Model verified: {now}")

            # ---- Intelligence (top-level radios) ----
            if intel_label:
                pill = await self._get_pill()
                pill_text = (
                    ((await pill.text_content()) or "").strip() if pill else ""
                )
                if self._pill_intel_matches(pill_text, intel_label):
                    logger.info(f"Intelligence already {pill_text!r}")
                else:
                    if not await self._open_pill_menu():
                        return False
                    # 2026-08-12 UI: the tiers live in an "Effort" submenu
                    # (formerly top-level radios). Open it when present;
                    # older builds fall through to the top-level radios.
                    effort_parent = await self._effort_submenu_parent()
                    if effort_parent is not None and not (
                        await self._open_model_submenu(effort_parent)
                    ):
                        logger.error("Effort submenu did not expand")
                        await self._close_pill_menu()
                        return False
                    if not await self._click_visible_radio(intel_label):
                        logger.error(
                            f"Intelligence {intel_label!r} not found in menu"
                        )
                        await self._close_pill_menu()
                        return False
                    await self._close_pill_menu()
                    pill = await self._get_pill()
                    pill_text = (
                        ((await pill.text_content()) or "").strip() if pill else ""
                    )
                    if not self._pill_intel_matches(pill_text, intel_label):
                        logger.error(
                            f"Intelligence verification failed: wanted "
                            f"{intel_label!r}, pill reads {pill_text!r}"
                        )
                        return False
                    logger.info(f"Intelligence verified: {pill_text}")

            return True
        except Exception as e:
            logger.error(f"Error selecting model/intelligence: {e}")
            await self._close_pill_menu()
            return False

    async def ensure_mode(self) -> bool:
        """Assert the Chat/Work toggle matches ``chatgpt_web.mode``.

        Defaults to "chat". Asserted in both directions because the
        selection persists across sessions. Backward compatible: a surface
        without the toggle passes for chat mode and fails loudly for work.
        """
        mode = (self.agent_config.get("mode") or "chat").lower()
        if mode not in self.MODE_VALUES:
            logger.error(
                f"Unknown chatgpt_web.mode {mode!r}. Valid: "
                f"{', '.join(self.MODE_VALUES)}"
            )
            return False
        label = "Work" if mode == "work" else "Chat"

        async def _find():
            # Two generations. Old: button[role=radio] with aria-checked.
            # New (2026-09-22): a [role=group][aria-label="Composer mode"] of
            # plain buttons with aria-pressed — no radio role anywhere, so the
            # old lookup returned None and a chat run passed as "chat-only
            # surface" while a work run failed outright.
            h = await self.page.evaluate_handle(
                """(label) => {
                    const vis = el => el.getClientRects().length > 0;
                    const hit = el => (el.textContent || '').trim().endsWith(label);
                    const radios = Array.from(document.querySelectorAll(
                        'button[role="radio"]')).filter(vis);
                    const found = radios.find(hit);
                    if (found) return found;
                    const group = Array.from(document.querySelectorAll(
                        '[role="group"]')).filter(vis).find(
                            g => (g.getAttribute('aria-label') || '')
                                .toLowerCase().includes('composer mode'));
                    if (!group) return null;
                    return Array.from(group.querySelectorAll('button'))
                        .filter(vis).find(hit) || null;
                }""",
                label,
            )
            return h.as_element()

        async def _checked(el) -> bool:
            for attr in ("aria-checked", "aria-pressed"):
                if (await el.get_attribute(attr)) == "true":
                    return True
            return False

        # Poll rather than read once: the toggle hydrates after the chat
        # input. This matters in BOTH directions — a work run fails outright
        # on a missed toggle, and a chat run silently keeps whatever the
        # previous task left selected, since the selection persists.
        radio = await _find()
        waited = 0.0
        while radio is None and waited < self.MODE_TOGGLE_WAIT_SEC:
            await asyncio.sleep(self.MODE_TOGGLE_POLL_SEC)
            waited += self.MODE_TOGGLE_POLL_SEC
            radio = await _find()
        if radio is not None and waited:
            logger.info(f"Chat/Work toggle appeared after {waited:.1f}s")

        if radio is None:
            if mode == "chat":
                logger.info(
                    f"Chat/Work toggle not present after "
                    f"{self.MODE_TOGGLE_WAIT_SEC:.0f}s — chat-only surface"
                )
                return True
            logger.error(
                f"chatgpt_web.mode=work but the Chat/Work toggle was not "
                f"found after {self.MODE_TOGGLE_WAIT_SEC:.0f}s — UI drift, "
                f"or not a home/project page."
            )
            return False

        if await _checked(radio):
            logger.info(f"Mode already '{mode}'")
            return True

        # Radix toggles want trusted pointer events; JS click as fallback.
        try:
            await radio.click(timeout=4000)
        except Exception:
            await radio.evaluate(self._JS_CLICK)
        await asyncio.sleep(1.5)
        radio = await _find()
        if radio is not None and await _checked(radio):
            logger.info(f"Mode set to '{mode}' (verified)")
            return True
        logger.error(f"Mode toggle to '{mode}' did not verify")
        return False

    async def _click_el(self, el) -> bool:
        """Real hover+click with JS-dispatch fallback."""
        try:
            await el.hover(timeout=2500)
        except Exception:
            pass
        try:
            await el.click(timeout=4000)
            return True
        except Exception:
            try:
                await el.evaluate(self._JS_CLICK)
                return True
            except Exception:
                return False

    async def _find_work_row(self, prefix: str):
        """A work-mode picker row ("Model…", "Effort…", "Speed…").

        Scoped to an OPEN [role="menu"]: `.__menu-item` also matches the
        left sidebar, and a conversation title starting with the prefix
        (live 2026-08-10: "Model Durumunu Kontrol" in the project sidebar)
        made _ensure_work_rows_visible short-circuit without ever opening
        the pill menu — the role-scoped click then found nothing and the
        task failed with "Work picker row 'Model' not clickable"."""
        h = await self.page.evaluate_handle(
            """(prefix) => Array.from(document.querySelectorAll(
                '.__menu-item, [role^="menuitem"]'
            )).filter(el => el.getClientRects().length > 0)
              .filter(el => el.closest('[role="menu"]'))
              .find(el => (el.textContent || '').trim().startsWith(prefix)) || null""",
            prefix,
        )
        return h.as_element()

    async def _select_work_option(self, row_prefix: str, option_label: str) -> bool:
        """Open a work-picker row's submenu and click the option.

        Option matching: exact text first, else the SHORTEST visible item
        starting with the label (options concatenate descriptions, e.g.
        "UltraConsumes usage limits faster", "Fast1.5x speed, more usage").
        Verified by re-reading the row text afterwards.
        """
        if not await self._ensure_work_rows_visible():
            logger.error(f"Work picker rows not visible for {row_prefix!r}")
            return False
        row = await self._find_work_row(row_prefix)
        if row is None:
            logger.error(f"Work picker row {row_prefix!r} not found")
            return False

        current = ((await row.text_content()) or "").strip()
        if current[len(row_prefix):].strip().startswith(option_label):
            logger.info(f"{row_prefix} already {option_label!r}")
            return True

        # KEYBOARD navigation. Every pointer path fails on this picker
        # (verified live 2026-07-21): React recreates the flyout nodes
        # continuously, so ElementHandle clicks go stale ("not attached"),
        # locator clicks never see a stable element, and raw coordinate
        # clicks dismiss the menu without selecting. The menu supports
        # keyboard nav with DOM focus on the .__menu-item itself
        # (document.activeElement — NOT [data-highlighted]): ArrowDown
        # walks rows, ArrowRight opens a row's flyout, Enter selects.
        # NOTE: never press Left/Right except on a confirmed row — the
        # power slider at the top of this menu binds Left/Right to adjust
        # effort.
        async def _focused_text() -> str:
            return await self.page.evaluate(
                """() => {
                    const ae = document.activeElement;
                    return ae ? (ae.textContent || '').trim().slice(0, 60) : '';
                }"""
            )

        # ROLE-BASED selection. Focus-walking the picker is not viable: a
        # "power" slider takes initial focus, and `.__menu-item` also
        # matches the LEFT SIDEBAR (New chat / Search / Library /
        # conversation names), so a walk drifts into the sidebar. The picker
        # rows carry proper roles — [role="menuitem"] with text
        # "<Row>\n<Value>" — and a synthetic pointer-event click on them is
        # stable (verified live 2026-08-02: Effort -> Ultra, pill confirmed).
        async def _click_by_text(prefix: str, within_flyout: bool):
            handle = await self.page.evaluate_handle(
                """(args) => {
                    const [prefix] = args;
                    const rx = new RegExp('^' + prefix.replace(
                        /[.*+?^${}()|[\\]\\\\]/g, '\\\\$&'));
                    const cands = [...document.querySelectorAll(
                        '[role="menuitem"],[role="menuitemradio"]')]
                        .filter(e => e.offsetParent !== null)
                        .filter(e => rx.test((e.innerText || '').trim()));
                    // shortest match wins: options concatenate descriptions
                    // ("UltraConsumes usage limits faster")
                    cands.sort((a, b) =>
                        (a.innerText || '').length - (b.innerText || '').length);
                    return cands[0] || null;
                }""",
                [prefix],
            )
            el = handle.as_element()
            if el is None:
                return False
            await el.evaluate(self._JS_POINTER_CLICK)
            return True

        # 1. Open the target row's flyout.
        if not await _click_by_text(row_prefix, False):
            logger.error(f"Work picker row {row_prefix!r} not clickable")
            return False
        await asyncio.sleep(1.2)

        # 2. Select the option inside it.
        if not await _click_by_text(option_label, True):
            logger.error(
                f"Option {option_label!r} not found in {row_prefix!r} flyout"
            )
            await self.page.keyboard.press("Escape")
            return False
        await asyncio.sleep(1.5)

        # Selecting an option usually CLOSES the menu — reopen to verify.
        if not await self._ensure_work_rows_visible():
            logger.error(f"Could not reopen picker to verify {row_prefix!r}")
            return False
        row = await self._find_work_row(row_prefix)
        now = ((await row.text_content()) or "").strip() if row else ""
        if now[len(row_prefix):].strip().startswith(option_label):
            logger.info(f"{row_prefix} set to {option_label!r} (verified)")
            return True
        logger.error(
            f"{row_prefix} verification failed: wanted {option_label!r}, "
            f"row reads {now!r}"
        )
        return False

    async def _slider_picker_present(self) -> bool:
        """True when the open pill menu is the slider generation."""
        return await self.page.query_selector(self._SLIDER_TOGGLE_SEL) is not None

    async def _slider_effort_state(self):
        """(label, position, total) from the Power row's live announcement,
        e.g. ("Extra High", 4, 6). None when unreadable — callers fail
        loudly rather than guess."""
        try:
            state = await self.page.evaluate(
                """(sel) => {
                    const sc = document.querySelector(sel);
                    if (!sc) return null;
                    const ids = (sc.getAttribute('aria-describedby') || '')
                        .split(/\\s+/);
                    for (const id of ids) {
                        const el = document.getElementById(id);
                        const t = el ? (el.textContent || '').trim() : '';
                        // No $ anchor: at the top stop the announcement
                        // appends a sentence — "Ultra, 6 of 6. Consumes
                        // usage limits faster" (live 2026-08-28).
                        const m = t.match(/^(.+?), (\\d+) of (\\d+)\\.?/);
                        if (m) return [m[1], parseInt(m[2], 10), parseInt(m[3], 10)];
                    }
                    return null;
                }""",
                self._SLIDER_POWER_SEL,
            )
            return tuple(state) if state else None
        except Exception:
            return None

    async def _slider_set_model(self, model_label: str) -> bool:
        """Pin the model radio in the slider picker's advanced view.

        Clicks the radio even when it already appears selected: selection
        flips the widget back to the simple view (observed behavior), which
        the effort step needs, and a same-value select is idempotent.
        """
        short = model_label.replace("GPT-", "")
        toggle = await self.page.query_selector(self._SLIDER_TOGGLE_SEL)
        if toggle is None:
            logger.error("Slider picker: 'Select model' toggle vanished")
            return False
        await toggle.evaluate(self._JS_POINTER_CLICK)
        await asyncio.sleep(1.0)
        # Match the radio's FIRST text line: "Default" carries a
        # description line, plain models don't.
        clicked = await self.page.evaluate(
            """(args) => {
                const [short, full] = args;
                const radios = [...document.querySelectorAll(
                    '[role="menu"][data-state="open"] [role="menuitemradio"]')];
                const t = radios.find(r => {
                    const first = ((r.innerText || '').trim()
                        .split('\\n')[0] || '').trim();
                    return first === short || first === full;
                });
                if (!t) return null;
                const r = t.getBoundingClientRect();
                const opts = {bubbles: true, cancelable: true, view: window,
                              clientX: r.left + r.width / 2,
                              clientY: r.top + r.height / 2};
                for (const ev of ['pointerdown','mousedown','pointerup',
                                  'mouseup','click'])
                    t.dispatchEvent(new MouseEvent(ev, opts));
                return ((t.innerText || '').trim().split('\\n')[0] || '').trim();
            }""",
            [short, model_label],
        )
        if clicked is None:
            logger.error(f"Slider picker: model radio {short!r} not found")
            return False
        await asyncio.sleep(1.2)
        # The radios stay in the DOM after the view flips back — verify
        # aria-checked without reopening the advanced view. Accept the
        # radio's first line as either the short ("5.6 Sol", 2026-08-28 UI)
        # or the full ("GPT-5.6 Sol", 2026-09-06 UI) form — the click above
        # already did; verifying on the short form alone failed every
        # work-mode run once the UI restored the "GPT-" prefix.
        checked = await self.page.evaluate(
            """(args) => {
                const [short, full] = args;
                const radios = [...document.querySelectorAll(
                    '[role="menu"][data-state="open"] [role="menuitemradio"]')];
                const t = radios.find(r => {
                    const first = ((r.innerText || '').trim()
                        .split('\\n')[0] || '').trim();
                    return first === short || first === full;
                });
                return t ? t.getAttribute('aria-checked') === 'true' : false;
            }""",
            [short, model_label],
        )
        if not checked:
            logger.error(
                f"Slider picker: model {short!r} did not verify as checked"
            )
            return False
        logger.info(f"Slider picker: model {short!r} selected (verified)")
        return True

    async def _slider_set_effort(
        self, effort_label: str, ladder: Optional[tuple] = None
    ) -> bool:
        """Walk the Power slider to the target effort with arrow keys.

        Label-driven: the announcement text is re-read after every press,
        and positions are never trusted (they shift with the model's
        available range).

        ``ladder`` is the UI order of the stops, low → high. Omitted, it is
        the work-mode ladder, so work-mode callers behave exactly as before;
        the chat slider branch passes ``_CHAT_SLIDER_LADDER``."""
        work_ladder = ladder is None
        ladder = self._SLIDER_EFFORT_LADDER if work_ladder else ladder
        if effort_label not in ladder:
            logger.error(
                f"Slider picker: effort {effort_label!r} not in ladder "
                f"{ladder}"
            )
            return False
        target_idx = ladder.index(effort_label)
        state = await self._slider_effort_state()
        if state is None:
            logger.error("Slider picker: cannot read Power slider state")
            return False
        label, pos, total = state
        if label == effort_label:
            logger.info(
                f"Slider picker: effort already {effort_label!r} "
                f"({pos} of {total})"
            )
            return True
        if target_idx >= total:
            hint = (
                "Ultra requires an explicit model, not 'Default'"
                if work_ladder
                else "this model's chat ladder stops short of it"
            )
            logger.error(
                f"Slider picker: {effort_label!r} needs stop "
                f"{target_idx + 1} but the slider offers {total} — {hint}"
            )
            return False
        sc = await self.page.query_selector(self._SLIDER_POWER_SEL)
        if sc is None:
            logger.error("Slider picker: Power row not found")
            return False
        await sc.evaluate("el => el.focus()")
        await asyncio.sleep(0.3)
        focused = await self.page.evaluate(
            "(sel) => document.activeElement === document.querySelector(sel)",
            self._SLIDER_POWER_SEL,
        )
        if not focused:
            logger.error("Slider picker: Power row did not take focus")
            return False
        # Bounded walk: one press per iteration, re-read, stop on match.
        for _ in range(len(ladder) + 2):
            state = await self._slider_effort_state()
            if state is None:
                logger.error("Slider picker: Power state unreadable mid-walk")
                return False
            label, pos, total = state
            if label == effort_label:
                logger.info(
                    f"Slider picker: effort set to {effort_label!r} "
                    f"({pos} of {total}, verified)"
                )
                return True
            try:
                cur_idx = ladder.index(label)
            except ValueError:
                logger.error(
                    f"Slider picker: unknown effort label {label!r} — "
                    f"UI drifted again"
                )
                return False
            key = "ArrowRight" if target_idx > cur_idx else "ArrowLeft"
            await self.page.keyboard.press(key)
            await asyncio.sleep(0.6)
        logger.error(
            f"Slider picker: arrow walk never reached {effort_label!r} "
            f"(stuck at {label!r})"
        )
        return False

    async def _slider_set_speed(self, speed_label: str) -> bool:
        """Fast-mode checkbox: unchecked = Standard, checked = Fast."""
        want = speed_label == "Fast"
        cb = await self.page.query_selector(self._SLIDER_FAST_SEL)
        if cb is None:
            logger.error("Slider picker: fast-mode toggle not found")
            return False
        cur = (await cb.get_attribute("aria-checked")) == "true"
        if cur == want:
            logger.info(f"Slider picker: speed already {speed_label!r}")
            return True
        await cb.evaluate(self._JS_POINTER_CLICK)
        await asyncio.sleep(0.8)
        cb = await self.page.query_selector(self._SLIDER_FAST_SEL)
        cur = cb is not None and (await cb.get_attribute("aria-checked")) == "true"
        if cur != want:
            logger.error(
                f"Slider picker: fast-mode did not verify for {speed_label!r}"
            )
            return False
        logger.info(f"Slider picker: speed set to {speed_label!r} (verified)")
        return True

    @staticmethod
    def _chat_pill_matches(pill_text: str, token: str, intel_label: str) -> bool:
        """True iff the chat pill reads exactly ``<token><tier>``.

        Whitespace- and case-insensitive, otherwise exact: "6Pro" and
        "6 Pro" both satisfy ("6", "Pro"); "6.1Pro", "6 AstraPro", "5.5Pro"
        and a bare "Pro" do not. Exactness is the point — under "Latest"
        the token is the only on-screen evidence of which model answers."""
        def squash(s: str) -> str:
            return "".join((s or "").split()).lower()

        want = squash(token) + squash(intel_label)
        return bool(want) and squash(pill_text) == want

    async def _open_pill_menu_patient(self) -> bool:
        """_open_pill_menu, then a slower second try. Chat path only.

        On a loaded machine the menu can render later than the 1.2 s
        _open_pill_menu allows; its second click then lands on the
        late-opening menu and closes it again (live 2026-09-21, load ~20,
        right after a Work→Chat flip: "pill menu did not open"). So: wait,
        look before clicking, click once, and poll for the open state."""
        if await self._open_pill_menu():
            return True
        open_sel = '[role="menu"][data-state="open"]'
        for _ in range(2):
            await asyncio.sleep(3.0)
            if await self.page.query_selector(open_sel):
                logger.info("ChatGPT pill menu opened late")
                return True
            pill = await self._get_pill()
            if pill is None:
                continue
            try:
                await pill.click()
            except Exception:
                await pill.evaluate(self._JS_CLICK)
            for _ in range(10):
                await asyncio.sleep(0.5)
                if await self.page.query_selector(open_sel):
                    logger.info("ChatGPT pill menu opened on the patient retry")
                    return True
        return False

    async def _chat_menu_model_text(self) -> Optional[str]:
        """The open picker's "Select model" row text, e.g. "6\\nPro".

        This row carries the model token in every generation seen so far,
        which the 2026-09-22 composer button no longer does. None when the
        menu is closed or the row is absent."""
        try:
            return await self.page.evaluate(
                """(sel) => { const el = document.querySelector(sel);
                    return el ? (el.innerText || '').trim() : null; }""",
                self._SLIDER_TOGGLE_SEL,
            )
        except Exception:
            return None

    async def _chat_slider_radio_state(self, radio_label: str) -> Optional[bool]:
        """aria-checked of the chat model radio whose FIRST text line is
        ``radio_label`` (rows append hints: "GPT-5.5 / Leaving on …").
        None when no such radio is in the open menu."""
        return await self.page.evaluate(
            """(label) => {
                const radios = [...document.querySelectorAll(
                    '[role="menu"][data-state="open"] [role="menuitemradio"]')]
                    .filter(r => r.getClientRects().length > 0);
                const t = radios.find(r => ((r.innerText || '').trim()
                    .split('\\n')[0] || '').trim() === label);
                return t ? t.getAttribute('aria-checked') === 'true' : null;
            }""",
            radio_label,
        )

    async def _chat_slider_set_model(self, radio_label: str) -> bool:
        """Select ``radio_label`` in the chat slider picker's model list.

        Unlike work mode, an already-checked radio is left alone: work's
        toggle shows what "Default" RESOLVES to, so a displayed name proved
        nothing there, whereas this reads the radio's own aria-checked — an
        explicit selection. Verified on aria-checked after a click, reopening
        the menu first if the click closed it."""
        state = await self._chat_slider_radio_state(radio_label)
        if state is None:
            # List collapsed behind the view toggle — expand it once.
            toggle = await self.page.query_selector(self._SLIDER_TOGGLE_SEL)
            if toggle is not None:
                await toggle.evaluate(self._JS_POINTER_CLICK)
                await asyncio.sleep(1.0)
                state = await self._chat_slider_radio_state(radio_label)
        if state is None:
            logger.error(
                f"Chat slider picker: model radio {radio_label!r} not found"
            )
            return False
        if state:
            logger.info(f"Chat slider picker: model already {radio_label!r}")
            return True

        clicked = await self.page.evaluate(
            """(label) => {
                const radios = [...document.querySelectorAll(
                    '[role="menu"][data-state="open"] [role="menuitemradio"]')]
                    .filter(r => r.getClientRects().length > 0);
                const t = radios.find(r => ((r.innerText || '').trim()
                    .split('\\n')[0] || '').trim() === label);
                if (!t) return false;
                const r = t.getBoundingClientRect();
                const opts = {bubbles: true, cancelable: true, view: window,
                              clientX: r.left + r.width / 2,
                              clientY: r.top + r.height / 2};
                for (const ev of ['pointerdown','mousedown','pointerup',
                                  'mouseup','click'])
                    t.dispatchEvent(new MouseEvent(ev, opts));
                return true;
            }""",
            radio_label,
        )
        if not clicked:
            logger.error(
                f"Chat slider picker: model radio {radio_label!r} vanished"
            )
            return False
        await asyncio.sleep(1.2)
        if not await self._open_pill_menu_patient():
            return False
        if await self._chat_slider_radio_state(radio_label) is not True:
            logger.error(
                f"Chat slider picker: model {radio_label!r} did not verify "
                f"as checked"
            )
            return False
        logger.info(f"Chat slider picker: model {radio_label!r} selected (verified)")
        return True

    async def _ensure_chat_slider(
        self, model: Optional[str], model_label: Optional[str],
        intel_label: Optional[str],
    ) -> bool:
        """Chat mode on the slider-generation picker (pill menu is open).

        Model radio, then the Power slider on the chat ladder, then the pill.
        Nothing reopens the menu after the slider is set."""
        radio_label = None
        token = None
        if model:
            radio_label = self.CHAT_SLIDER_MODEL_RADIOS.get(model, model_label)
            token = self.CHAT_SLIDER_PILL_TOKENS.get(
                model, (model_label or "").replace("GPT-", "")
            )
        if radio_label and not await self._chat_slider_set_model(radio_label):
            await self._close_pill_menu()
            return False

        if intel_label:
            # The model-list view disables the Power row; a reopen lands on
            # the simple view. Harmless before the slider is set.
            disabled = await self.page.evaluate(
                """(sel) => { const el = document.querySelector(sel);
                    return !!el && (el.getAttribute('aria-disabled') === 'true'
                                    || el.hasAttribute('data-disabled')); }""",
                self._SLIDER_POWER_SEL,
            )
            if disabled:
                await self._close_pill_menu()
                if not await self._open_pill_menu_patient():
                    return False
            if not await self._slider_set_effort(
                intel_label, ladder=self._CHAT_SLIDER_LADDER
            ):
                await self._close_pill_menu()
                return False
        # Read the model+tier the menu itself shows BEFORE closing it: the
        # 2026-09-22 composer's button reads "Thinking effortPro" — the tier
        # only — so the pill alone can no longer prove which model answers.
        menu_text = await self._chat_menu_model_text()
        await self._close_pill_menu()

        pill = await self._get_pill()
        pill_text = ((await pill.text_content()) or "").strip() if pill else ""
        source, shown = "pill", pill_text
        if intel_label and token:
            ok = self._chat_pill_matches(pill_text, token, intel_label)
            if not ok and menu_text is not None:
                # Older pill wording is gone — fall back to the menu row,
                # which still carries the model token ("6" + "Pro").
                ok = self._chat_pill_matches(menu_text, token, intel_label)
                source, shown = "menu 'Select model' row", menu_text
            want = f"{token}{intel_label}"
        elif intel_label:
            ok = self._pill_intel_matches(pill_text, intel_label)
            want = f"…{intel_label}"
        else:
            ok, want = True, ""
        if not ok:
            logger.error(
                f"Chat model/tier verification failed: wanted {want!r}, pill "
                f"reads {pill_text!r}, menu row reads {menu_text!r}"
            )
            return False
        logger.info(f"Chat settings verified ({source}: {shown!r})")
        return True

    async def ensure_work_settings(self) -> bool:
        """Work-mode picker: model / effort / speed, either generation.

        Detects which picker the open pill menu renders (slider generation
        vs Advanced rows — see the class comments) and drives that one.
        Fails loudly on unknown labels or verification mismatch. Final
        cross-check reads the pill text (``<model-short> <effort>``, e.g.
        "5.6 Sol Ultra") — NOT the slider, which in the rows generation
        decouples after an off-slider effort like Ultra is chosen.
        """
        model = self.agent_config.get("model")
        effort = self.agent_config.get("effort")
        speed = self.agent_config.get("speed") or "standard"

        model_label = None
        if model:
            model_label = self.WORK_MODEL_LABELS.get(
                model.lower(), self.MODEL_LABELS.get(model.lower())
            )
            if model_label is None:
                logger.error(
                    f"Unknown chatgpt_web.model {model!r} for work mode. "
                    f"Valid: {', '.join(self.WORK_MODEL_LABELS)}"
                )
                return False
        effort_label = None
        if effort:
            effort_label = self.WORK_EFFORT_LABELS.get(effort.lower())
            if effort_label is None:
                logger.error(
                    f"Unknown chatgpt_web.effort {effort!r}. Valid: "
                    f"{', '.join(self.WORK_EFFORT_LABELS)}"
                )
                return False
        speed_label = self.WORK_SPEED_LABELS.get(speed.lower())
        if speed_label is None:
            logger.error(
                f"Unknown chatgpt_web.speed {speed!r}. Valid: "
                f"{', '.join(self.WORK_SPEED_LABELS)}"
            )
            return False

        try:
            if not await self._open_pill_menu():
                logger.error("Work pill menu did not open")
                return False

            if await self._slider_picker_present():
                # Slider generation (2026-08-28). Model first — it is what
                # unlocks the Ultra stop on the effort slider.
                if model_label and not await self._slider_set_model(model_label):
                    await self._close_pill_menu()
                    return False
                if effort_label and not await self._slider_set_effort(effort_label):
                    await self._close_pill_menu()
                    return False
                if not await self._slider_set_speed(speed_label):
                    await self._close_pill_menu()
                    return False
                await self._close_pill_menu()
            else:
                # Advanced-rows generation (pre-2026-08-28).
                if not await self._ensure_work_rows_visible():
                    logger.error("Advanced rows not reachable in work pill menu")
                    await self._close_pill_menu()
                    return False

                for prefix, label in (
                    ("Model", model_label),
                    ("Effort", effort_label),
                    ("Speed", speed_label),
                ):
                    if label is None:
                        continue
                    if not await self._select_work_option(prefix, label):
                        await self._close_pill_menu()
                        return False

                await self._close_pill_menu()

            # Cross-check the pill: shows "<model-short> <effort>" (model
            # without the "GPT-" prefix).
            pill = await self._get_pill()
            pill_text = ((await pill.text_content()) or "").strip() if pill else ""
            if effort_label and effort_label.lower() not in pill_text.lower():
                logger.error(
                    f"Pill verification failed: effort {effort_label!r} not "
                    f"in pill text {pill_text!r}"
                )
                return False
            if model_label:
                short = model_label.replace("GPT-", "")
                if short.lower() not in pill_text.lower():
                    logger.error(
                        f"Pill verification failed: model {short!r} not in "
                        f"pill text {pill_text!r}"
                    )
                    return False
            logger.info(f"Work settings verified (pill: {pill_text!r})")
            return True
        except Exception as e:
            logger.error(f"Error configuring work settings: {e}")
            await self._close_pill_menu()
            return False

    async def ensure_features_enabled(self) -> bool:
        """Assert mode, then select the mode's picker settings.

        chat mode → model + intelligence, on whichever picker generation the
                    pill menu renders (slider, or the older submenus).
        work mode → model + effort + speed, likewise (slider or Advanced rows).
        """
        await asyncio.sleep(2)
        # Re-assert mode (cheap when already correct — navigation set it
        # before files were uploaded; flipping HERE would risk dropping
        # attachments, so a flip at this point is logged by ensure_mode).
        if not await self.ensure_mode():
            return False
        mode = (self.agent_config.get("mode") or "chat").lower()
        if mode == "work":
            return await self.ensure_work_settings()
        return await self.ensure_model_and_intelligence()

    # The composer's permanent hidden <input type=file>, newest UI first.
    # 2026-09-22: the id disappeared — the composer rebuild ships React-
    # generated ids ("_r_hp_") and identifies the inputs by aria-label
    # instead, with photo/video inputs alongside the general one. A run that
    # still looked only for #upload-files lost both chooser paths too and
    # failed every task on that account with upload_failed (lane B, task 70).
    # Never match an input with an `accept` filter: those are the photo and
    # camera paths and they reject .xlsx.
    UPLOAD_INPUT_SELECTORS = (
        'input#upload-files[type="file"]',
        'input[type="file"][aria-label="Attach files"]',
        'input[type="file"][multiple]:not([accept])',
        'input[type="file"]:not([accept])',
    )

    @staticmethod
    def _chip_selector(stem: str) -> str:
        """Selector for the composer's attachment tile of a file whose name
        contains ``stem``, across both composer generations.

        Old: a div[role="group"] carrying the displayed name as aria-label.
        New (2026-09-22): a plain <button aria-label="Name.xlsx"> beside a
        <button aria-label="Remove Name.xlsx">. The Remove twin has to be
        excluded or every file counts twice. Missing the tile does not just
        mis-report: upload_files re-SETS the input up to four times when no
        tile is seen, so a stale selector attaches the same file repeatedly
        before giving up.
        """
        return (
            f'[role="group"][aria-label*="{stem}"], '
            f'button[aria-label*="{stem}"]:not([aria-label^="Remove"])'
        )

    async def _direct_upload_input(self):
        """The hidden file input to set files on, or None if the page has
        none. Returns a Locator already narrowed to one element."""
        for sel in self.UPLOAD_INPUT_SELECTORS:
            try:
                loc = self.page.locator(sel)
                if await loc.count() > 0:
                    logger.info(f"Composer file input: {sel}")
                    return loc.first
            except Exception:
                continue
        return None

    async def upload_files(self, file_paths: list[str]) -> bool:
        """Upload files via the + menu > Add photos & files flow.

        Falls back to the "Add files and more" button (bottom-left of composer)
        if the + menu approach fails — the button text changes depending on
        whether agent mode is active.
        """
        logger.info("upload_files v4 (direct-input + plus-wait + row-retry) active")
        try:
            for file_path in file_paths:
                logger.info(f"Uploading file: {file_path}")

                uploaded = False

                # Approach 0 (2026-08-12, observed live on task 1
                # ApfelInc, chat surface): the + menu row renders with the
                # expected text but clicking it no longer fires a
                # filechooser event, so both chooser-based approaches
                # timed out. The composer form carries a permanent hidden
                # <input id="upload-files" multiple> with no accept filter
                # (the two accept="image/*" inputs are the photo/camera
                # paths); setting files on it directly fires the React
                # change handler and attaches the file. Success requires
                # the attachment tile to appear — otherwise fall through
                # to the menu approaches for builds without this input.
                try:
                    direct_input = await self._direct_upload_input()
                    if direct_input is not None:
                        # Match the tile on the file's STEM, not its full
                        # name: chatgpt.com de-duplicates an upload whose
                        # name already exists in the account by appending a
                        # counter, so ApfelInc.xlsx renders as
                        # "ApfelInc(2).xlsx" on any run after the first. The
                        # rename also lands ASYNCHRONOUSLY — the tile appears
                        # under the local name and is relabelled when the
                        # server answers — so an exact-name match was a race
                        # that passed or failed on timing (observed live
                        # 2026-08-21, task 1). The tile is a div[role=group]
                        # carrying the displayed name as its aria-label;
                        # it is NOT a <button>, so the old
                        # button[aria-label=...] half never matched at all.
                        stem = Path(file_path).stem.replace('"', '')
                        chip = self.page.locator(self._chip_selector(stem))
                        before = await chip.count()

                        # Retry the set: for the first ~3s after
                        # navigate_to_new_chat returns, the composer accepts
                        # the change event and DROPS it — no tile ever
                        # appears, not even after 30s (measured live
                        # 2026-08-21: at 0s delay the file is lost, at 3s it
                        # attaches). navigate returns as soon as the chat
                        # input is ATTACHED, which is earlier than whatever
                        # the uploader is waiting on; React's own props are
                        # on the input from the first millisecond, so they
                        # are not a usable readiness signal. Because a
                        # dropped set never materializes later, re-setting
                        # cannot double-attach.
                        for set_try in range(4):
                            await self._set_input_files_cdp_safe(
                                direct_input.first, file_path)
                            try:
                                await chip.nth(before).wait_for(
                                    state="attached", timeout=6000)
                                uploaded = True
                                break
                            except Exception:
                                logger.info(
                                    f"composer dropped the file — re-setting "
                                    f"(try {set_try + 1}/4)"
                                )
                        if not uploaded:
                            raise RuntimeError(
                                "the hidden file input accepted the file but "
                                "no attachment tile appeared"
                            )
                        shown = await chip.nth(before).get_attribute(
                            "aria-label")
                        logger.info(
                            f"Uploaded via the hidden file input "
                            f"(direct set_input_files); tile shows {shown!r}")
                except Exception as e_direct:
                    logger.info(
                        f"Direct hidden-input upload failed — "
                        f"falling back to + menu ({e_direct!r:.200})"
                    )

                # Approach 1: + panel > "Add photos & files" row.
                # The + panel is no longer an ARIA menu (2026-07 UI): rows
                # are div.__menu-item[tabindex=0] with no role/testid, so
                # get_by_role("menuitem", ...) matches nothing. Anchor on
                # the row class + text; keep the old role-based lookups as
                # fallbacks for older builds.
                add_files = None
                try:
                    # Dismiss any stale popup first
                    await self.page.keyboard.press("Escape")
                    await self.page.wait_for_timeout(300)

                    plus_btn = self.page.locator(self.SELECTORS["plus_menu_button"])
                    # The + button renders AFTER the chat input — an
                    # instant count() check ~1s post-load sees 0 and
                    # skipped this whole approach (observed live
                    # 2026-07-21 on the task-7 gate). Wait for it.
                    try:
                        await plus_btn.wait_for(
                            state="visible", timeout=10000)
                    except Exception:
                        pass
                    if not uploaded and await plus_btn.count() > 0:
                        await plus_btn.click(timeout=5000)
                        await self.page.wait_for_timeout(1200)
                        # Auto-waiting row lookup with re-click retries: a
                        # single fixed 1.2s wait raced the menu render and
                        # failed the whole upload (observed live 2026-07-21
                        # on the task-7 gate). Rows carry a subtitle now
                        # ("Upload from computer"), so match loosely.
                        row = self.page.locator(
                            'div.__menu-item:has-text("Add photos"), '
                            'div.__menu-item:has-text("Upload from computer")'
                        )
                        for menu_try in range(3):
                            try:
                                await row.first.wait_for(
                                    state="visible", timeout=4000)
                                add_files = row.first
                                break
                            except Exception:
                                logger.info(
                                    f"+ menu row not visible (try "
                                    f"{menu_try + 1}/3) — re-clicking +"
                                )
                                try:
                                    await plus_btn.click(timeout=5000)
                                except Exception:
                                    pass
                                await self.page.wait_for_timeout(1200)
                        if add_files is None:
                            for name in (
                                "Add photos & files",
                                "Add files and more",
                                "Add files",
                            ):
                                cand = self.page.get_by_role("menuitem", name=name)
                                if await cand.count() > 0:
                                    add_files = cand.first
                                    break

                        try:
                            if add_files is None:
                                raise RuntimeError(
                                    "'Add photos & files' row not found"
                                )
                            async with self.page.expect_file_chooser(
                                timeout=10000
                            ) as fc_info:
                                await add_files.click(timeout=5000)

                            file_chooser = await fc_info.value
                            await self._set_files_cdp_safe(file_chooser, file_path)
                            uploaded = True
                        except Exception as e1:
                            import traceback as _tb
                            _frames = _tb.format_exc().splitlines()
                            logger.info(
                                f"+ panel 'Add photos & files' not found, trying fallback ({e1!r:.200})\n"
                                + "\n".join(_frames[:6] + ["..."] + _frames[-12:])
                            )
                            await self.page.keyboard.press("Escape")
                            await self.page.wait_for_timeout(500)
                except Exception as e0:
                    logger.info(f"+ menu approach failed, trying fallback ({e0!r:.200})")

                # Approach 2: the composer's own "Add files" button, for the
                # composer states that expose it instead of routing uploads
                # through the + menu. Anchor via aria-label where possible;
                # text fallbacks are last-resort.
                if not uploaded:
                    try:
                        # NOTE: no fuzzy aria-label*="file" match here — it
                        # grabbed a sidebar chat titled "…Excel File
                        # Creation" and burned 30s failing to click it
                        # (observed live 2026-07-21).
                        add_files_btn = self.page.locator(
                            'button[aria-label="Add files and more"], '
                            'button[aria-label="Add photos & files"], '
                            'button:has-text("Add files and more"), '
                            'button:has-text("Add files")'
                        )
                        if await add_files_btn.count() > 0:
                            async with self.page.expect_file_chooser(
                                timeout=10000
                            ) as fc_info:
                                await add_files_btn.first.click()

                            file_chooser = await fc_info.value
                            await self._set_files_cdp_safe(file_chooser, file_path)
                            uploaded = True
                            logger.info("Used 'Add files' button fallback")
                    except Exception as e:
                        logger.warning(f"Fallback file upload also failed: {e}")

                if not uploaded:
                    raise RuntimeError(f"Could not upload file: {file_path}")

                await self.page.wait_for_timeout(2000)

                # Verify the file appears as an attachment tile. Same stem
                # match as approach 0 above (chatgpt.com may rename the
                # upload to "<stem>(n).<ext>"), and .first because the
                # composer holds one tile per uploaded file — a bare
                # locator over several of them is a strict-mode violation,
                # which surfaced as a spurious "not confirmed" warning on
                # an upload that had in fact succeeded.
                stem = Path(file_path).stem.replace('"', '')
                attachment = self.page.locator(self._chip_selector(stem))
                try:
                    await attachment.first.wait_for(
                        state="attached", timeout=5000)
                    shown = await attachment.first.get_attribute("aria-label")
                    logger.info(f"File attached: {shown!r}")
                except Exception:
                    logger.warning(
                        f"File attachment not confirmed for: "
                        f"{Path(file_path).name}")

            # The chip appearing only means the file was ACCEPTED, not that
            # ChatGPT finished processing the upload. When their file
            # pipeline is slow (verified live 2026-07-23), send stays
            # gated ("File upload pending", send button disabled) for
            # tens of seconds. If we race ahead to type the prompt, open
            # the model menu, and click send, the send silently no-ops and
            # the whole task fails. Block here until uploads truly settle.
            if not await self._wait_for_uploads_complete(len(file_paths)):
                logger.error(
                    "Uploads never finished processing — not proceeding "
                    "(would fail to submit)")
                return False
            # Remember what the composer is supposed to be carrying: an
            # attachment can still vanish between here and the send (see
            # _missing_attachments).
            self._expected_attachments = [str(p) for p in file_paths]
            return True
        except Exception as e:
            logger.error(f"File upload failed: {e}")
            return False

    async def _missing_attachments(self) -> list[str]:
        """Files uploaded earlier whose tile is no longer in the composer.

        Live 2026-09-22/23 (lane B, several tasks): uploads settled, the
        picker verified, and then ONE of the two tiles quietly disappeared —
        the case workbook on one task, House_Standards on another. With an
        attachment missing the composer refuses to submit: the Send click
        does nothing, Enter does nothing, and the attempt burns both wait
        windows before failing. Checking costs one DOM query and turns a
        silent 75-second stall into a re-upload.
        """
        missing = []
        for path in getattr(self, "_expected_attachments", []) or []:
            stem = Path(path).stem.replace('"', "")
            try:
                if await self.page.locator(self._chip_selector(stem)).count() == 0:
                    missing.append(path)
            except Exception:
                continue
        return missing

    async def _restore_attachments(self) -> bool:
        """Re-upload any attachment that vanished. True if the composer ends
        up carrying everything it should."""
        missing = await self._missing_attachments()
        if not missing:
            return True
        logger.warning(
            "Attachment(s) vanished from the composer before send: "
            + ", ".join(Path(p).name for p in missing)
            + " — re-uploading"
        )
        direct = await self._direct_upload_input()
        if direct is None:
            return False
        for path in missing:
            stem = Path(path).stem.replace('"', "")
            chip = self.page.locator(self._chip_selector(stem))
            await self._set_input_files_cdp_safe(direct, path)
            try:
                await chip.first.wait_for(state="attached", timeout=15000)
            except Exception:
                logger.error(f"Re-upload of {Path(path).name} did not attach")
                return False
        if not await self._wait_for_uploads_complete(len(missing)):
            return False
        still = await self._missing_attachments()
        if still:
            logger.error(
                "Still missing after re-upload: "
                + ", ".join(Path(p).name for p in still)
            )
            return False
        logger.info("Attachments restored — proceeding to send")
        return True

    async def _wait_for_uploads_complete(
        self, n_files: int, timeout_sec: int = 240
    ) -> bool:
        """Block until attachment processing is done: the send button is
        enabled AND no 'upload pending' state remains, held stable across a
        few polls. Returns False if it never settles within timeout."""
        deadline = asyncio.get_event_loop().time() + timeout_sec
        stable = 0
        last_log = 0.0
        while asyncio.get_event_loop().time() < deadline:
            try:
                st = await self.page.evaluate(
                    """() => {
                    const s = document.querySelector(
                        'button[data-testid="send-button"], '
                        + 'button[aria-label="Send prompt"], '
                        + 'button[aria-label="Send"]');
                    const t = document.body.innerText || '';
                    return {
                        enabled: s ? (!s.disabled
                            && s.getClientRects().length > 0) : false,
                        pending: t.includes('File upload pending')
                            || t.includes('upload pending'),
                    };
                }""")
            except Exception:
                st = {"enabled": False, "pending": True}
            if st["enabled"] and not st["pending"]:
                stable += 1
                if stable >= 3:  # ~4.5s of steady "ready"
                    logger.info(
                        f"Uploads processed and ready ({n_files} file(s))")
                    return True
            else:
                stable = 0
            now = asyncio.get_event_loop().time()
            if now - last_log > 20:
                logger.info(
                    f"Waiting for uploads to finish processing "
                    f"(enabled={st['enabled']}, pending={st['pending']})")
                last_log = now
            await self.page.wait_for_timeout(1500)
        return False

    async def submit_prompt(self, prompt: str, prompt_number: int = 1) -> bool:
        """Type prompt text and click send.

        Retries once through a page RELOAD when the composer refuses to
        submit: live 2026-09-22/23 on one account, a composer holding the
        prompt and both attachments, with Send enabled and focus in the
        editor, silently swallowed every click AND the Enter key — the same
        click succeeded on the same page seconds after a reload. The engine's
        own attempt-retry did not clear it, so the recovery belongs here.
        """
        if await self._submit_prompt_once(prompt, prompt_number):
            return True
        if getattr(self, "_send_reload_used", False):
            return False
        self._send_reload_used = True
        logger.warning("Send was swallowed — reloading the page and retrying once")
        try:
            await self.page.reload(wait_until="domcontentloaded", timeout=90000)
            await self.page.wait_for_timeout(8000)
            files = list(getattr(self, "_expected_attachments", []) or [])
            if files:
                self._expected_attachments = []
                if not await self.upload_files(files):
                    logger.error("Re-upload after reload failed")
                    return False
            # The picker is re-asserted by _submit_prompt_once itself on the
            # first prompt, so it is not repeated here.
        except Exception as e:
            logger.error(f"Reload recovery failed: {e}")
            return False
        return await self._submit_prompt_once(prompt, prompt_number)

    async def _submit_prompt_once(self, prompt: str, prompt_number: int = 1) -> bool:
        try:
            logger.info(f"Submitting prompt {prompt_number} ({len(prompt)} chars)")

            # Focus the chat input (ProseMirror div OR paragraph placeholder)
            editor = self.page.locator(
                'div.ProseMirror[contenteditable="true"], '
                'p[data-placeholder*="New chat"]'
            )
            try:
                await editor.first.wait_for(state="attached", timeout=10000)
            except Exception:
                logger.error("Chat input not found")
                return False
            await editor.first.click()
            await self.page.wait_for_timeout(300)

            # Clear any leftover text
            await self.page.keyboard.press("Meta+a")
            await self.page.keyboard.press("Backspace")
            await self.page.wait_for_timeout(200)

            # Fill prompt. Playwright fill() first; on failure VERIFY what
            # the composer holds before doing anything else. 2026-09-06
            # (sol_B, task 86): fill() of the 77k-char prompt raised on its
            # 30s timeout AFTER the text had landed in ProseMirror, and the
            # unconditional clipboard-paste fallback then delivered a SECOND
            # copy, which ChatGPT converts into a "Pasted text.txt"
            # attachment — the turn went out as inline prompt + pasted-text
            # document + workbook. No Aug-29 cohort row ever took the paste
            # path (0/68), so a duplicated prompt is a cohort-integrity
            # break, not a cosmetic one. Order of preference now: fill;
            # else keyboard.insertText (inline, no paste-to-file
            # conversion); else clipboard paste. Every step is checked by
            # reading the composer's own text length.
            target_len = len(prompt.split())
            async def _composer_words() -> int:
                try:
                    return await editor.first.evaluate(
                        "el => (el.innerText || '').trim().split(/\\s+/).filter(Boolean).length"
                    )
                except Exception:
                    return -1
            def _landed(n: int) -> bool:
                return n >= 0 and abs(n - target_len) <= max(3, target_len // 100)
            try:
                # 30s (the default) is not enough for a 77k-char prompt on the
                # no-project composer (2026-09-06: raised at 30.6s with the
                # text already landed; Aug-29 in-project fills took ~8s).
                await editor.first.fill(prompt, timeout=120_000)
            except Exception:
                n = await _composer_words()
                if _landed(n):
                    logger.info(
                        f"fill() raised but the composer already holds the prompt "
                        f"({n} words) — not pasting a second copy"
                    )
                else:
                    logger.info(
                        f"fill() failed (composer holds {n} words of {target_len}); "
                        "clearing and retrying via insertText"
                    )
                    await editor.first.click()
                    await self.page.keyboard.press("Meta+a")
                    await self.page.keyboard.press("Backspace")
                    await self.page.wait_for_timeout(200)
                    try:
                        await self.page.keyboard.insert_text(prompt)
                    except Exception as e:
                        logger.info(f"insertText failed ({e}); falling back to clipboard paste")
                    await self.page.wait_for_timeout(1000)
                    n = await _composer_words()
                    if not _landed(n):
                        logger.warning(
                            f"insertText left {n} words of {target_len}; falling back "
                            "to clipboard paste (ChatGPT may attach it as Pasted text)"
                        )
                        await self.page.keyboard.press("Meta+a")
                        await self.page.keyboard.press("Backspace")
                        await self.page.evaluate(
                            "(text) => navigator.clipboard.writeText(text)", prompt
                        )
                        await self.page.keyboard.press("Meta+v")
                        await self.page.wait_for_timeout(500)
            await self.page.wait_for_timeout(1000)
            n = await _composer_words()
            logger.info(f"Composer holds {n} words (prompt has {target_len})")

            # Select model + intelligence after files are attached and the
            # prompt is typed, but before sending (only on first prompt).
            # Hard-fail: sending on the wrong model silently benchmarks the
            # wrong thing, which is worse than a failed attempt.
            if prompt_number == 1:
                if not await self.ensure_features_enabled():
                    logger.error(
                        "Model/intelligence selection failed — not sending"
                    )
                    return False

            # Click send button. It is briefly DISABLED after the prompt is
            # typed while ChatGPT finishes processing the attachments
            # server-side — a plain wait_for(visible) returns immediately
            # (a disabled button is still visible), we click into the void,
            # and the send silently no-ops. Seen live 2026-07-23: two tasks
            # in a row failed to submit with send-button disabled=true.
            # Poll for the button to become ENABLED, then click; retry.
            # Last check before the send: everything that was uploaded must
            # still be attached, or the composer silently refuses to submit.
            if not await self._restore_attachments():
                logger.error(
                    "Composer is missing an attachment and it could not be "
                    "restored — failing this attempt rather than sending an "
                    "incomplete prompt"
                )
                return False

            url_before = self.page.url
            send_btn = self.page.locator(self.SELECTORS["send_button"])
            sent_click = False
            for attempt in range(45):  # up to ~90s for attachment processing
                try:
                    enabled = await self.page.evaluate(
                        """() => {
                        const b = document.querySelector(
                            'button[data-testid="send-button"], '
                            + 'button[aria-label="Send prompt"], '
                            + 'button[aria-label="Send"]');
                        return b ? !b.disabled
                            && b.getClientRects().length > 0 : false;
                    }""")
                except Exception:
                    enabled = False
                if enabled:
                    try:
                        await send_btn.first.click(timeout=3000)
                        sent_click = True
                        break
                    except Exception:
                        pass  # re-poll and retry
                await self.page.wait_for_timeout(2000)
            if not sent_click:
                logger.info(
                    "Send button never became enabled — trying Enter key")
                await self.page.keyboard.press("Enter")

            # Wait for confirmation that the prompt was sent.
            # For prompt 1: URL changes from project page to /c/{id}
            # For prompts 2+: URL already has /c/ — check generation indicators
            already_in_conversation = "/c/" in url_before
            enter_tried = False
            for tick in range(45):  # 45s max wait for send confirmation
                await self.page.wait_for_timeout(1000)
                # A click that silently does nothing leaves the prompt sitting
                # in the composer (live 2026-09-22, lane B CFForecast: uploads
                # and picker fine, then "URL unchanged" three attempts in a
                # row). Text still in the composer after ~10s means the turn
                # was never submitted, so press Enter once — the same key a
                # person would use. Harmless if the click did land: the
                # composer is empty by then and Enter does nothing.
                if tick == 10 and not enter_tried:
                    enter_tried = True
                    try:
                        left = await self.page.evaluate(
                            """() => { const c = document.querySelector(
                                'div.ProseMirror[contenteditable="true"]');
                                return c ? (c.innerText || '').trim().length : 0; }"""
                        )
                    except Exception:
                        left = 0
                    if left > 0:
                        logger.warning(
                            f"No send confirmation after 10s and {left} chars "
                            f"still in the composer — pressing Enter"
                        )
                        try:
                            await self.page.keyboard.press("Enter")
                        except Exception as e:
                            logger.warning(f"Enter fallback failed: {e}")
                current_url = self.page.url
                if current_url != url_before and "/c/" in current_url:
                    logger.info(f"Prompt sent — conversation: {current_url}")
                    return True
                # Check if generation started (Stop button visible) — CDP-safe
                if await self._check_button_text("Stop"):
                    logger.info("Prompt sent (Stop button appeared)")
                    return True
                # For follow-up prompts, also check if a new article appeared
                if already_in_conversation and await self._is_generating():
                    logger.info("Prompt sent (generation indicators detected)")
                    return True

            logger.error(
                f"Prompt may not have been sent (URL unchanged: {self.page.url})"
            )
            return False

        except Exception as e:
            logger.error(f"Submit prompt failed: {e}")
            return False

    async def _is_generating(self) -> bool:
        """Check if ChatGPT is still generating using JS DOM queries.

        CDP tabs may not render overlay elements as "visible" to Playwright,
        so we query the DOM directly via JavaScript. Checks for:
        - Stop button via data-testid (reliable during Code Interpreter execution)
        - Stop button via text (present during Pro thinking)
        - Answer now button (present during extended thinking)
        - Pro thinking indicator (extended thinking in progress)
        - result-streaming class (text streaming)
        - "Writing code" / "Analyzing" text (agent mode code execution)
        - "ChatGPT is generating" status text
        - Send button absence (no send button while generating)
        """
        try:
            return await self.page.evaluate(
                """() => {
                // Most reliable: data-testid="stop-button" (present during ALL generation phases)
                const hasStopBtn = !!document.querySelector('[data-testid="stop-button"]');
                // Only check buttons INSIDE the main chat area, not sidebar history items
                // (sidebar items like "Early Stopping in Experiments" have aria-labels containing "Stop")
                const mainArea = document.querySelector('main') || document.body;
                const btns = Array.from(mainArea.querySelectorAll('button'));
                const hasStop = btns.some(b => b.textContent.trim() === 'Stop');
                // 2026-09-22 composer: the Stop control carries only an
                // aria-label — its text is empty — so the test above misses
                // it and a healthy turn reads as "generation never started"
                // (lane B, every task). Exact label, and mainArea scope, as
                // sidebar chat titles have matched a loose "Stop" before.
                const hasStopAria = btns.some(
                    b => (b.getAttribute('aria-label') || '').trim() === 'Stop');
                const hasAnswerNow = btns.some(b => b.textContent.trim() === 'Answer now');
                const hasThinking = btns.some(b => b.textContent.includes('Pro thinking'));
                const hasGenerating = !!document.querySelector('[class*="result-streaming"]');
                // Tool-use indicators (the model running code mid-answer):
                // scoped to the conversation area so sidebar chat titles
                // cannot match.
                const mainText = mainArea.innerText || '';
                const hasWritingCode = mainText.includes('Writing code');
                const hasAnalyzing = /Analyz(ing|ed)/.test(mainText) && (hasStop || hasStopBtn);
                return hasStopBtn || hasStop || hasStopAria || hasAnswerNow || hasThinking || hasGenerating || hasWritingCode || hasAnalyzing;
            }"""
            )
        except Exception:
            return False

    async def _check_usage_limit(self):
        """CDP-safe scan for usage-limit phrases. Returns the matched
        phrase (for logging) or None."""
        try:
            return await self.page.evaluate(
                "(phrases) => { const t = document.body.innerText || '';"
                " return phrases.find(p => t.includes(p)) || null; }",
                list(self.USAGE_LIMIT_PHRASES),
            )
        except Exception:
            return None

    async def _set_input_files_cdp_safe(self, input_locator, file_path) -> None:
        """set_input_files directly on a hidden <input type=file> (no
        filechooser event involved), with the same raw-CDP fallback as
        _set_files_cdp_safe for files over Playwright's 50MB CDP cap."""
        try:
            await input_locator.set_input_files(file_path)
            return
        except Exception as e:
            if "larger than 50Mb" not in str(e):
                raise
        logger.info(
            "File exceeds Playwright's 50MB CDP cap — using raw CDP "
            "DOM.setFileInputFiles on the hidden input"
        )
        el = await input_locator.element_handle()
        await el.evaluate("el => el.setAttribute('data-cdp-upload', '1')")
        cdp = await self.page.context.new_cdp_session(self.page)
        try:
            doc = await cdp.send("DOM.getDocument")
            node = await cdp.send("DOM.querySelector", {
                "nodeId": doc["root"]["nodeId"],
                "selector": 'input[data-cdp-upload="1"]',
            })
            if not node.get("nodeId"):
                raise RuntimeError("tagged file input not found via CDP")
            files = file_path if isinstance(file_path, list) else [file_path]
            await cdp.send("DOM.setFileInputFiles", {
                "files": [str(p) for p in files],
                "nodeId": node["nodeId"],
            })
        finally:
            await cdp.detach()
            try:
                await el.evaluate(
                    "el => el.removeAttribute('data-cdp-upload')")
            except Exception:
                pass

    async def _set_files_cdp_safe(self, file_chooser, file_path) -> None:
        """set_files with a raw-CDP fallback for files over Playwright's
        50MB CDP-connection cap (same guard hit on the Claude agent
        2026-07-21; DOM.setFileInputFiles has no size limit)."""
        try:
            await file_chooser.set_files(file_path)
            return
        except Exception as e:
            if "larger than 50Mb" not in str(e):
                raise
        logger.info(
            "File exceeds Playwright's 50MB CDP cap — using raw CDP "
            "DOM.setFileInputFiles"
        )
        el = file_chooser.element
        await el.evaluate("el => el.setAttribute('data-cdp-upload', '1')")
        cdp = await self.page.context.new_cdp_session(self.page)
        try:
            doc = await cdp.send("DOM.getDocument")
            node = await cdp.send("DOM.querySelector", {
                "nodeId": doc["root"]["nodeId"],
                "selector": 'input[data-cdp-upload="1"]',
            })
            if not node.get("nodeId"):
                raise RuntimeError("tagged file input not found via CDP")
            files = file_path if isinstance(file_path, list) else [file_path]
            await cdp.send("DOM.setFileInputFiles", {
                "files": [str(p) for p in files],
                "nodeId": node["nodeId"],
            })
        finally:
            await cdp.detach()
            try:
                await el.evaluate(
                    "el => el.removeAttribute('data-cdp-upload')")
            except Exception:
                pass

    async def _recover_content_load_error(self) -> bool:
        """Recover from ChatGPT's transient "Content failed to load" error.

        That error replaces the response with a placeholder — real DOM:
            <h2>Content failed to load</h2>
            <button class="btn … btn-secondary"><div>Try again</div></button>
        and leaves the response empty until "Try again" is clicked. Requires
        BOTH the error text AND a button whose text is exactly "Try again"
        (specific to this state — the healthy-response button says "Regenerate",
        not "Try again"), then clicks it. Returns True iff it clicked.
        """
        try:
            clicked = await self.page.evaluate(
                r"""() => {
                const main = document.querySelector('main') || document.body;
                const errRe = /content failed to load|something went wrong|couldn.?t load/i;
                const hasErr = Array.from(main.querySelectorAll('h1,h2,h3,p,div,span'))
                    .some(e => e.children.length === 0 && errRe.test(e.textContent || ''));
                if (!hasErr) return false;
                const btn = Array.from(main.querySelectorAll('button'))
                    .find(b => (b.textContent || '').trim().toLowerCase() === 'try again');
                if (!btn) return false;
                btn.click();
                return true;
            }"""
            )
            return bool(clicked)
        except Exception as e:
            logger.warning(f"content-load-error recovery check failed: {e}")
            return False

    async def _recover_stream_error(self) -> bool:
        """If the turn died with a stream error, click its Retry button.

        Returns True only when a Retry was actually clicked. Retry
        regenerates the turn from scratch — it does not resume — so the
        caller must treat this as a full-cost restart of the prompt.
        """
        try:
            return await self.page.evaluate(
                """(phrases) => {
                const body = document.body.innerText || '';
                if (!phrases.some(p => body.includes(p))) return false;
                // Only retry errors on a COMMITTED conversation (/c/ URL).
                // If the error hit at submission time, Retry restores the
                // draft and bounces to the site root (observed live
                // 2026-07-21), stranding the page — let the engine's
                // attempt-retry rebuild the session instead.
                if (!location.pathname.includes('/c/')) return false;
                const btn = Array.from(document.querySelectorAll('button'))
                    .filter(b => b.getClientRects().length > 0)
                    .find(b => (b.innerText || '').trim() === 'Retry');
                if (!btn) return false;
                btn.click();
                return true;
            }""",
                list(self.STREAM_ERROR_PHRASES),
            )
        except Exception:
            return False

    async def _ensure_conversation_view(self, timeout_sec: int = 90) -> bool:
        """Verify we are about to poll the CONVERSATION, not the project page.

        `_submit_prompt` returns True on ANY send signal — a changed URL, the
        Stop button, or generation indicators. When the send is made from a
        project composer the SPA can keep the page on ``/g/g-p-.../project``
        while the turn runs elsewhere. The project page never renders
        assistant turns, so the wait loop reads an empty response forever and
        the dead-page watchdog eventually kills a perfectly healthy run. That
        burned three tasks x ~45 min on 2026-08-09 (289/352/306) and left the
        artifact scan matching the UPLOADED input file listed on the project
        page — a false "the model produced nothing" every time.

        Deliberately does NOT fall back to opening "the newest conversation in
        the project": that list has been observed serving stale entries, and
        attaching to the wrong conversation would record ANOTHER task's
        workbook as this task's solution. Failing fast is recoverable;
        silently grading the wrong artifact is not.
        """
        if "/c/" in self.page.url:
            return True
        deadline = asyncio.get_event_loop().time() + timeout_sec
        while asyncio.get_event_loop().time() < deadline:
            await self.page.wait_for_timeout(2000)
            if "/c/" in self.page.url:
                logger.info(f"Conversation view reached: {self.page.url}")
                return True
        logger.error(
            f"Never navigated into a conversation after send (still at "
            f"{self.page.url}) — failing this attempt now instead of polling a "
            f"page that cannot show a response. Check whether the composer "
            f"posted from the project surface."
        )
        return False

    async def _activity_fingerprint(self) -> str:
        """Cheap liveness probe for the dead-page watchdog.

        Work-mode Sol runs spreadsheet/PDF tools for many minutes while
        emitting assistant turns whose visible content is EMPTY. During
        those stretches `generating` can read False and the extracted
        response text stays "", which the empty-idle watchdog alone reads as
        a dead run (observed 2026-08-09 on tasks 289/352/306: ~20 messages,
        every one with end_turn=False, no output file yet).

        Turn/article COUNTS grow whenever the model emits anything at all,
        including empty progress turns. Deliberately excludes body text —
        sidebar timestamps ("2 minutes ago") churn on their own and would
        make this always-alive, disarming the watchdog completely.
        """
        try:
            return await self.page.evaluate(
                """() => {
                  const turns = document.querySelectorAll('[data-turn]').length;
                  const arts = document.querySelectorAll(
                      "[data-message-author-role='assistant']").length;
                  return turns + ':' + arts;
                }"""
            )
        except Exception:
            return ""

    async def _count_xlsx_tiles(self) -> int:
        """Excel file tiles currently in the conversation (inputs render as
        tiles too, so callers compare against a submission-time baseline)."""
        try:
            return await self.page.evaluate(
                """() => [...document.querySelectorAll(
                    'button[aria-label$=".xlsx"], button[aria-label$=".xls"]'
                )].length"""
            )
        except Exception:
            return -1

    async def _count_response_articles(self) -> int:
        """Count the assistant turns currently on the page.

        Must match the tiers _extract_last_response reads, because only some
        responses carry data-message-author-role: verified live 2026-08-20
        across 14 conversations, short plain-text answers carried
        role='assistant' (5/14) while longer agentic ones carried only
        [data-turn]/.agent-turn (9/14). Counting the role attribute alone
        returned 0 on those 9, which silently disabled both callers — the
        "a new turn appeared" test for generation-started, and the
        articles_changed liveness signal during generation.
        """
        try:
            return await self.page.evaluate(
                """() => {
                const assistants = document.querySelectorAll("[data-message-author-role='assistant']");
                if (assistants.length > 0) return assistants.length;
                const turns = Array.from(document.querySelectorAll('[data-turn]'))
                    .filter(el => el.getAttribute('data-turn') !== 'user');
                if (turns.length > 0) return turns.length;
                const agentTurns = document.querySelectorAll('.agent-turn').length;
                if (agentTurns > 0) return agentTurns;
                // 2026-09-22 composer: an exchange is a [data-turn-key]
                // wrapper holding a [data-user-message-bubble] and a hidden
                // [data-chatgpt-agent-turn-start] anchor. None of the older
                // markers exist there, so this returned 0 and the wait loop
                // lost both its "a turn appeared" and liveness signals.
                return document.querySelectorAll(
                    '[data-chatgpt-agent-turn-start]').length;
            }"""
            )
        except Exception:
            return 0

    async def wait_for_response(self, prompt_number: int = 1) -> Optional[str]:
        """Wait for ChatGPT to finish responding.

        Uses a three-phase approach:
        1. Snapshot existing articles so we only track NEW responses.
        2. Wait for generation to actually start (new article OR generation
           indicators like Stop/Pro thinking).
        3. Wait for completion using content stabilization — the response text
           must stop changing for a sustained period AND generation indicators
           must be gone.

        This prevents premature completion when:
        - Stale articles from prior conversations are on the page
        - There's a gap between extended thinking and code execution
        - The model pauses between code execution steps
        """
        logger.info(f"Waiting for response to prompt {prompt_number}...")

        if not await self._ensure_conversation_view():
            return None

        start_time = asyncio.get_event_loop().time()

        # Phase 0: Snapshot how many response articles exist BEFORE this response
        baseline_article_count = await self._count_response_articles()
        if not self._baseline_set:
            # Only store baseline before the FIRST prompt so that
            # download_all_artifacts searches ALL conversation articles,
            # not just those from the last prompt.
            self._baseline_article_count = baseline_article_count
            self._baseline_set = True
            logger.info(
                f"Baseline response articles on page (first prompt): {baseline_article_count}"
            )
        else:
            logger.info(
                f"Current response articles on page: {baseline_article_count} (baseline kept at {self._baseline_article_count})"
            )

        # Minimum time before accepting completion (Pro tasks need
        # substantial time for Excel building). If a file card is detected
        # (Spreadsheet/.xlsx), accept sooner — the model may have finished
        # quickly with just a file output.
        min_elapsed_sec = 120
        min_elapsed_sec_with_file = 30  # Accept quickly if file card present

        # 2026-08-01 UI rollout: work mode emits separate short assistant
        # turns with quiet (no-indicator) sandbox gaps between them, so
        # "text stable + not generating" no longer implies DONE. Real
        # completion signal = a NEW file tile beyond the ones present at
        # submission (inputs render as tiles too, so compare against a
        # baseline count). Without a new tile, only accept completion
        # after the text has been continuously stable for a long window.
        WORK_STABLE_ACCEPT_SEC = 600
        work_mode = (self.agent_config.get("mode") or "").lower() == "work"
        baseline_tiles = 0
        if work_mode:
            baseline_tiles = max(0, await self._count_xlsx_tiles())
        stable_since = None  # wall time when current stability run began
        cur_tiles = baseline_tiles

        # Give ChatGPT time to start generating
        await self.page.wait_for_timeout(5000)

        # Phase 1: Wait for generation to start (up to 120s)
        # Require EITHER generation indicators OR a NEW article (beyond baseline)
        generation_started = False
        for _ in range(60):
            if self.shutdown_event and self.shutdown_event.is_set():
                return None

            if await self._is_generating():
                generation_started = True
                logger.info("Response generation started (indicators detected)")
                break

            # Check for a NEW ChatGPT article (beyond baseline count)
            current_count = await self._count_response_articles()
            if current_count > baseline_article_count:
                generation_started = True
                logger.info(
                    f"New ChatGPT response article appeared ({current_count} > {baseline_article_count})"
                )
                break

            await self.page.wait_for_timeout(2000)

        if not generation_started:
            logger.error("Response generation did not start within 120s")
            return None

        # Phase 2: Wait for completion using content stabilization
        # The response text must stop changing AND generation indicators must be
        # gone for required_stable consecutive checks.
        stable_count = 0
        required_stable = 5  # 5 consecutive checks × check_interval (15s at 3s)
        last_response_text = ""
        last_article_count = baseline_article_count

        # Usage-limit fast-fail requires PERSISTENCE: a transient toast (or
        # a false-positive text match) must not abandon a healthy in-flight
        # response — verified live 2026-07-21, where a single-hit check
        # killed an Ultra response that was still generating 10+ minutes
        # later. Only fail after several consecutive sightings.
        limit_streak = 0
        LIMIT_STREAK_REQUIRED = 3

        # Server-side stream failures ("Error in message stream") end
        # generation with a Retry button and no result. Observed live
        # 2026-07-21 at ~65 min into an Ultra run that had nearly finished.
        # Retry does NOT resume — verified live: the errored turn is
        # discarded and regenerated from the prompt, losing the work. It is
        # still the cheapest recovery (a full task restart re-uploads and
        # costs the same generation anyway), but because each retry costs a
        # fresh full-length generation, keep the cap low.
        stream_retries = 0
        MAX_STREAM_RETRIES = 2

        # Dead-page bailout (see the empty-idle branch below).
        # 2026-08-09: raised 300 -> 1800. Five minutes of silence is NORMAL in
        # work mode (Sol runs tools for long stretches emitting empty turns);
        # at 300s this watchdog killed three healthy mid-build runs in a row.
        # The branch now also resets on turn-count growth, so this limit only
        # bites when the page is genuinely producing nothing at all.
        empty_idle_since = None
        last_fingerprint = ""
        EMPTY_IDLE_LIMIT_SEC = 1800

        while (asyncio.get_event_loop().time() - start_time) < self.max_wait_per_prompt:
            if self.shutdown_event and self.shutdown_event.is_set():
                return None

            if stream_retries < MAX_STREAM_RETRIES:
                clicked = await self._recover_stream_error()
                if clicked:
                    stream_retries += 1
                    logger.warning(
                        f"Stream error detected — clicked Retry "
                        f"({stream_retries}/{MAX_STREAM_RETRIES}); resuming"
                    )
                    stable_count = 0
                    await asyncio.sleep(5)
                    continue

            # Sibling recovery: ChatGPT's transient "Content failed to load"
            # placeholder (h2 + a "Try again" button — distinct from the
            # stream-error "Retry" state above). It wipes the response until
            # Try again is clicked, so without this the response sits empty
            # until the per-prompt timeout. Reuses the stream-retry budget.
            if stream_retries < MAX_STREAM_RETRIES:
                if await self._recover_content_load_error():
                    stream_retries += 1
                    logger.warning(
                        f"'Content failed to load' detected — clicked Try "
                        f"again ({stream_retries}/{MAX_STREAM_RETRIES})"
                    )
                    stable_count = 0
                    # Wait long enough for the reload to actually complete
                    # before re-checking, so we don't hammer the button while
                    # the page is still reloading from the previous click.
                    await asyncio.sleep(20)
                    continue

            limit_phrase = await self._check_usage_limit()
            if limit_phrase:
                limit_streak += 1
                logger.warning(
                    f"Usage-limit phrase {limit_phrase!r} on page "
                    f"({limit_streak}/{LIMIT_STREAK_REQUIRED} consecutive)"
                )
                if limit_streak >= LIMIT_STREAK_REQUIRED:
                    logger.error(
                        "Usage/plan limit persisted — failing fast instead "
                        "of waiting out max_wait_per_prompt"
                    )
                    return None
            else:
                limit_streak = 0

            generating = await self._is_generating()

            # Also track article count — agent mode creates multiple response articles
            current_article_count = await self._count_response_articles()
            articles_changed = current_article_count != last_article_count
            if articles_changed:
                logger.info(
                    f"Article count changed: {last_article_count} -> {current_article_count}"
                )
                last_article_count = current_article_count

            # Always sample response text (even during generation) for monitoring
            current_response = await self._extract_last_response() or ""

            if generating or articles_changed:
                stable_count = 0
                stable_since = None
                empty_idle_since = None
                # Track text even during generation so we know progress
                last_response_text = current_response
            else:
                # Check if response text is still changing (content stabilization)
                text_changed = current_response != last_response_text
                last_response_text = current_response

                if text_changed:
                    stable_count = 0  # Content still growing
                    stable_since = None
                else:
                    stable_count += 1
                    if stable_since is None:
                        stable_since = asyncio.get_event_loop().time()

                elapsed = asyncio.get_event_loop().time() - start_time

                # Dead turn with nothing to click (see TURN_FAILED_TEXTS):
                # fail the attempt now so the engine starts a fresh chat,
                # instead of idling out the 30-minute dead-page watchdog.
                if (
                    not work_mode
                    and stable_since is not None
                    and current_response.strip() in self.TURN_FAILED_TEXTS
                    and (asyncio.get_event_loop().time() - stable_since)
                    >= self.TURN_FAILED_GRACE_SEC
                ):
                    logger.error(
                        f"Turn ended with {current_response.strip()!r} and no "
                        f"Retry control ({int(elapsed)}s elapsed) — failing "
                        f"this attempt so the engine retries in a fresh chat"
                    )
                    return None

                if stable_count >= required_stable:
                    # Content is stable and generation stopped.
                    # Accept if we have content and enough time elapsed.
                    if work_mode:
                        # Text keywords are useless here — progress notes
                        # mention .xlsx constantly. Only a NEW tile counts.
                        cur_tiles = await self._count_xlsx_tiles()
                        if cur_tiles < 0:
                            cur_tiles = baseline_tiles
                        has_file_indicator = cur_tiles > baseline_tiles
                        if not has_file_indicator:
                            stable_for = (
                                asyncio.get_event_loop().time() - stable_since
                                if stable_since is not None else 0
                            )
                            if stable_for < WORK_STABLE_ACCEPT_SEC:
                                # Quiet sandbox gap, not completion — keep
                                # waiting. Must not skip the end-of-loop
                                # sleep, so sleep here before continue.
                                await self.page.wait_for_timeout(
                                    self.check_interval * 1000
                                )
                                continue
                            # 2026-09-10: the turn ENDED (stable for the
                            # whole accept window, not generating) and no
                            # file was handed over. Return whatever text is
                            # on the page — even a bare "Worked for 45m"
                            # note — so the engine sends "Continue" in THIS
                            # conversation, where the sandbox and its files
                            # still exist. Before this, a note under 50
                            # chars fell into the dead-page branch below,
                            # sat 30 min, then restarted the task in a NEW
                            # chat and threw the work away (the 2026-09-08
                            # sol restarts). The dead-page guard now only
                            # covers a page with NO assistant turn at all.
                            turn_seen = current_article_count > baseline_article_count
                            is_error_surface = any(
                                p in current_response
                                for p in self.STREAM_ERROR_PHRASES
                            )
                            if turn_seen and not is_error_surface:
                                logger.warning(
                                    f"Work-mode turn ended with no new file "
                                    f"({int(stable_for)}s stable, "
                                    f"{len(current_response)} chars, "
                                    f"articles {baseline_article_count}->"
                                    f"{current_article_count}) — returning "
                                    f"so the engine can send 'Continue' in "
                                    f"this conversation"
                                )
                                return (
                                    current_response
                                    or "[work-mode turn ended with no file]"
                                )
                    else:
                        # File card responses can be very short (e.g. "Here is your file:\nSpreadsheet" = ~37 chars)
                        has_file_indicator = any(
                            kw in current_response
                            for kw in [
                                "Spreadsheet",
                                ".xlsx",
                                ".xls",
                                "Excel file",
                                "Download",
                            ]
                        )
                    # An error banner is itself ~70 chars of text, so it
                    # sails past the length check and gets accepted as a
                    # real answer (observed live 2026-07-21: repeated
                    # "Response complete … 76 chars" on failed turns).
                    # Never treat an error surface as content.
                    is_error_surface = any(
                        p in current_response
                        for p in self.STREAM_ERROR_PHRASES
                    )
                    content_ok = (
                        not is_error_surface
                        and (len(current_response) > 50 or has_file_indicator)
                    )
                    if content_ok:
                        effective_min = (
                            min_elapsed_sec_with_file
                            if has_file_indicator
                            else min_elapsed_sec
                        )
                        if elapsed >= effective_min:
                            logger.info(
                                f"Response complete ({int(elapsed)}s elapsed, "
                                f"{len(current_response)} chars, file_indicator={has_file_indicator})"
                            )
                            return current_response
                        else:
                            logger.info(
                                f"Content stable but too early ({int(elapsed)}s < {min_elapsed_sec}s min), "
                                f"continuing to wait..."
                            )
                            stable_count = 0
                    else:
                        # Content too short and no file indicators. Keep
                        # waiting briefly — but a page that stays idle AND
                        # empty is dead (e.g. a submission-time network
                        # error bounced ChatGPT to the root page with the
                        # draft restored; observed live 2026-07-21, where
                        # this branch looped for the full 2h task cap).
                        # Fail the attempt so the engine can rebuild the
                        # session instead of burning the whole budget.
                        # Turn/article growth means the model IS working, even
                        # with no visible text — do not count that as idle.
                        fingerprint = await self._activity_fingerprint()
                        if fingerprint and fingerprint != last_fingerprint:
                            last_fingerprint = fingerprint
                            empty_idle_since = None
                            logger.info(
                                f"Empty response but page activity "
                                f"({fingerprint}) — resetting idle watchdog"
                            )
                        elif empty_idle_since is None:
                            empty_idle_since = asyncio.get_event_loop().time()
                        elif (asyncio.get_event_loop().time()
                                - empty_idle_since) > EMPTY_IDLE_LIMIT_SEC:
                            logger.error(
                                f"Page idle with no response text for "
                                f">{EMPTY_IDLE_LIMIT_SEC}s — treating as "
                                f"dead page, failing this attempt"
                            )
                            return None
                        stable_count = 0

            elapsed = int(asyncio.get_event_loop().time() - start_time)
            # Log every ~30s (use range check since loop interval may skip exact multiples)
            if elapsed > 0 and elapsed % 30 < (self.check_interval + 1):
                resp_len = len(last_response_text)
                # 2026-09-10: enough state per line to classify a restart
                # from the log alone (detector misread vs ChatGPT-side
                # death): turn count vs baseline, file tiles vs baseline,
                # how long the page has been quiet, and the response tail.
                stable_for = (
                    int(asyncio.get_event_loop().time() - stable_since)
                    if stable_since is not None else 0
                )
                if work_mode:
                    t = await self._count_xlsx_tiles()
                    if t >= 0:
                        cur_tiles = t
                tail = last_response_text[-80:].replace("\n", " / ")
                logger.info(
                    f"Waiting... {elapsed}s elapsed, generating={generating}, "
                    f"response_len={resp_len}, stable={stable_count}/{required_stable} "
                    f"({stable_for}s), articles={last_article_count}"
                    f"(base {baseline_article_count}), tiles={cur_tiles}"
                    f"(base {baseline_tiles}), tail={tail!r}"
                )

            await self.page.wait_for_timeout(self.check_interval * 1000)

        # Timeout — return whatever we have if it looks reasonable
        logger.warning(f"Response timeout after {self.max_wait_per_prompt}s")
        final_response = await self._extract_last_response()
        if final_response and len(final_response) > 50:
            logger.info(
                f"Returning partial response ({len(final_response)} chars) after timeout"
            )
            return final_response
        return None

    async def _extract_last_response(self) -> Optional[str]:
        """Extract text from the last ChatGPT response.

        Uses JS evaluate instead of Playwright locators for CDP reliability.
        Two strategies, the second firing only when the first matched
        nothing on the page:
        1. data-message-author-role='assistant' — carried by short
           plain-text answers
        2. [data-turn] / .agent-turn — carried by longer agentic answers,
           which have no author-role attribute at all

        Both are needed: verified live 2026-08-20 over 14 conversations,
        5 exposed only strategy 1 and 9 only strategy 2.
        """
        try:
            text = await self.page.evaluate(
                """() => {
                // Strategy 1: data-message-author-role (chat-mode DOM)
                const assistants = document.querySelectorAll("[data-message-author-role='assistant']");
                if (assistants.length > 0) {
                    return assistants[assistants.length - 1].innerText;
                }
                // Strategy 2: WORK MODE — assistant turns carry no
                // author-role attribute. Since the 2026-08-01 UI rollout,
                // work mode emits MULTIPLE assistant [data-turn] elements
                // (one short progress note per phase) instead of a single
                // growing turn, and the generating indicator goes dark
                // between them. Reading only the LAST turn returned a
                // ~76-char stub -> false "dead page" while Sol kept working
                // server-side (tasks 147/392, 2026-08-01 05:26-05:59).
                // AGGREGATE every non-user turn instead: on the old
                // single-turn DOM this concatenates one element (identical
                // behavior); on the new DOM the total keeps growing with
                // each progress note, so stability detection still works.
                const turns = Array.from(document.querySelectorAll('[data-turn]'))
                    .filter(el => el.getAttribute('data-turn') !== 'user');
                if (turns.length > 0) {
                    return turns.map(el => el.innerText || '').join('\\n\\n');
                }
                const agentTurns = document.querySelectorAll('.agent-turn');
                if (agentTurns.length > 0) {
                    return Array.from(agentTurns).map(el => el.innerText || '').join('\\n\\n');
                }
                // Strategy 3 (2026-09-22 composer): no author-role, no
                // [data-turn], no .agent-turn. An exchange is one
                // [data-turn-key] wrapper whose text starts with the user's
                // own message (inside [data-user-message-bubble]); the
                // assistant's words are whatever follows it. Subtracting the
                // bubble matters — otherwise every poll reads the prompt
                // back, the text never changes, and a live turn looks
                // finished the moment it starts.
                const keyed = document.querySelectorAll('[data-turn-key]');
                if (keyed.length > 0) {
                    const turn = keyed[keyed.length - 1];
                    let text = turn.innerText || '';
                    const bubble = turn.querySelector('[data-user-message-bubble]');
                    if (bubble) {
                        const said = bubble.innerText || '';
                        const at = text.indexOf(said);
                        if (at >= 0) text = text.slice(at + said.length);
                    }
                    return text.replace(/^\\s*(You said:)?\\s*/, '');
                }
                return null;
            }"""
            )
            if text:
                text = text.strip()
                return text or None
            return None
        except Exception as e:
            logger.error(f"Response extraction failed: {e}")
            return None

    async def _download_via_backend_api(self, download_path: Path) -> list[str]:
        """Strategy 0: pull .xlsx bytes from ChatGPT's own backend API.

        The DOM path can require opening the file-preview panel, which
        renders the whole spreadsheet grid and OOM-crashes the renderer on
        multi-MB workbooks (UI_DRIFT_PLAYBOOK §1c). Sandbox outputs are
        referenced as ``sandbox:/mnt/data/<name>.xlsx`` in assistant
        messages; the authenticated chain (session cookies ride along on
        ``page.context.request``):

            1. GET /api/auth/session                        -> accessToken
            2. GET /backend-api/conversation/{cid}          -> sandbox refs
            3. GET .../interpreter/download?message_id&sandbox_path
                                                            -> download_url
            4. GET download_url                             -> xlsx bytes

        Gotchas encoded below: the ``sandbox:`` prefix must be stripped
        (raw /mnt/data path), the download id is MINTED by step 3 (there is
        no stable id in the DOM), and messages are walked oldest→newest so
        the last write of a filename wins (QA passes rewrite files).

        Safe by construction: bytes must validate (HTTP 200, len > 200, zip
        magic ``PK``); on ANY error or ambiguity this returns [] and the
        caller falls through to the DOM download flow.
        """
        saved: list[str] = []
        try:
            m = re.search(r"/c/([0-9a-f-]{8,})", self.page.url)
            if not m:
                logger.info("No conversation id in URL — skipping API download")
                return []
            cid = m.group(1)
            req = self.page.context.request

            sess = await req.get("https://chatgpt.com/api/auth/session")
            if not sess.ok:
                logger.info(f"auth/session returned {sess.status} — skipping")
                return []
            token = (await sess.json()).get("accessToken")
            if not token:
                logger.info("No accessToken in session — skipping API download")
                return []
            headers = {"Authorization": f"Bearer {token}"}

            conv = await req.get(
                f"https://chatgpt.com/backend-api/conversation/{cid}",
                headers=headers,
            )
            if not conv.ok:
                logger.info(f"conversation fetch returned {conv.status} — skipping")
                return []
            data = await conv.json()

            # Oldest → newest so the LAST write of each filename wins.
            nodes = [
                n
                for n in (data.get("mapping") or {}).values()
                if isinstance(n, dict) and n.get("message")
            ]
            nodes.sort(key=lambda n: (n["message"].get("create_time") or 0))
            # TWO path shapes reach the same interpreter/download endpoint:
            #   chat / code-interpreter: "sandbox:/mnt/data/<name>.xlsx"
            #   work mode (Sol):         "/workspace/scratch/<id>/outputs/<id>/<name>.xlsx"
            # Work-mode names routinely contain SPACES and also appear
            # percent-encoded, so the character class may not exclude \s —
            # it stops at a quote or backslash instead. Verified live
            # 2026-07-30 recovering a 24.6 MB workbook from task 296.
            # A space is allowed ONLY when not followed by "/", so a match can
            # never run from one path into the next one mentioned in prose.
            path_re = re.compile(
                r"(?:sandbox:)?(/(?:mnt/data|workspace)/"
                r"(?:[^\"'\\\s]|\s(?!/))*?\.xlsx)",
                re.IGNORECASE,
            )
            # filename -> [(message_id, raw_path), ...]; the same file shows up
            # both encoded and decoded, and only one variant may mint.
            refs: dict[str, list] = {}
            for n in nodes:
                msg = n["message"]
                blob = json.dumps(msg.get("content") or {})
                for sm in path_re.finditer(blob):
                    raw_path = sm.group(1)
                    fname = unquote(raw_path.rsplit("/", 1)[-1])
                    cand = (msg.get("id"), raw_path)
                    refs.setdefault(fname, [])
                    if cand not in refs[fname]:
                        refs[fname].append(cand)

            if not refs:
                logger.info("No sandbox .xlsx refs in conversation — skipping")
                return []
            logger.info(
                f"Backend-API: {len(refs)} candidate artifact(s): {list(refs)}"
            )

            # The scan above sees every workbook path mentioned anywhere in the
            # conversation, including scratch files the model wrote and moved on
            # from an hour earlier. It cannot tell those from the deliverable —
            # but the finished page can, because it renders exactly the
            # artifact(s) the model handed over. Narrow to those when the page
            # says so, and keep everything when it doesn't.
            surfaced = await self._surfaced_artifact_names()
            if surfaced:
                scoped = {f: c for f, c in refs.items() if f in surfaced}
                if scoped and len(scoped) != len(refs):
                    logger.info(
                        f"Backend-API: final turn surfaces {sorted(surfaced)} — "
                        f"dropping {sorted(set(refs) - set(scoped))}"
                    )
                    refs = scoped
                elif not scoped:
                    logger.info(
                        f"Backend-API: no sandbox ref matches the surfaced "
                        f"artifact(s) {sorted(surfaced)} — keeping all candidates"
                    )

            for fname, candidates in refs.items():
                # sandbox_path must be the RAW /mnt/data/... or /workspace/...
                # path — leaving the sandbox: prefix on returns file_not_found.
                # Try newest candidate first, then the other encoding variant.
                for mid, raw_path in reversed(candidates):
                    mint = await req.get(
                        f"https://chatgpt.com/backend-api/conversation/{cid}"
                        f"/interpreter/download",
                        params={"message_id": mid, "sandbox_path": raw_path},
                        headers=headers,
                    )
                    if not mint.ok:
                        logger.info(f"download mint {mint.status} for {fname}")
                        continue
                    info = await mint.json()
                    dl_url = info.get("download_url")
                    if info.get("status") != "success" or not dl_url:
                        logger.info(f"download mint rejected for {fname}: {info}")
                        continue
                    blob_resp = await req.get(dl_url)
                    if not blob_resp.ok:
                        continue
                    body = await blob_resp.body()
                    if len(body) <= 200 or not body.startswith(b"PK"):
                        logger.info(
                            f"Rejected {fname}: {len(body)} bytes, not a zip/xlsx"
                        )
                        continue
                    target = download_path / fname
                    target.write_bytes(body)
                    saved.append(str(target))
                    logger.info(
                        f"Backend-API download: {target} ({len(body)} bytes)"
                    )
                    break
        except Exception as e:
            logger.info(f"Backend-API download unavailable ({e}) — using DOM flow")
        return saved

    async def _surfaced_artifact_names(self) -> set[str]:
        """Workbook filenames the FINAL assistant turn presents for download.

        This is the only signal that distinguishes the deliverable from a
        scratch file: the conversation history contains both, the finished
        turn contains only what the model handed over. Verified live
        2026-08-28 — the turn carries ``button[aria-label="FlightPlan.xlsx"]``
        (inline link, file card) beside ``button[aria-label="Download file"]``.

        Read-only: queries attributes, never clicks and never opens the
        preview, so it carries none of the §1c OOM risk.

        Returns an empty set on drift or error — the caller then keeps every
        candidate, so a selector change degrades to today's behavior rather
        than dropping a real deliverable.
        """
        try:
            names = await self.page.evaluate(
                """() => {
                  const turns = [...document.querySelectorAll(
                    "[data-message-author-role='assistant']")];
                  if (!turns.length) return [];
                  const last = turns[turns.length - 1];
                  const root = last.closest('article') || last.parentElement || last;
                  const out = new Set();
                  for (const el of root.querySelectorAll('[aria-label]')) {
                    const a = (el.getAttribute('aria-label') || '').trim();
                    if (/\\.xlsx?$/i.test(a)) out.add(a);
                  }
                  return [...out];
                }"""
            )
            surfaced = {n for n in (names or []) if n}
            if surfaced:
                logger.info(f"Final turn surfaces artifact(s): {sorted(surfaced)}")
            return surfaced
        except Exception as e:
            logger.info(
                f"Could not read surfaced artifacts ({e}) — keeping all candidates"
            )
            return set()

    async def _download_via_work_tiles(self, download_path: Path) -> list[str]:
        """Work-mode file downloads (verified live 2026-07-21).

        Work-mode outputs render as file TILES — ``button[aria-label=
        "<name>.xlsx"]`` covered by a ``[data-default-action]`` overlay div
        (which intercepts real pointer events; a JS click on the overlay
        works). Clicking opens a Library preview dialog whose toolbar has
        ``button[aria-label="Download"]``. There are no sandbox refs (this
        is not code interpreter) and no icon-button download in the card.

        NOTE: the preview renders the sheet grid, so very large workbooks
        carry the §1c OOM risk — no lighter path exists on this surface.
        """
        saved: list[str] = []
        try:
            tiles = [
                t
                for t in await self.page.query_selector_all(
                    'button[aria-label$=".xlsx"], button[aria-label$=".xls"]'
                )
                if await t.is_visible()
            ]
        except Exception:
            tiles = []
        if not tiles:
            return []
        self.last_download_saw_buttons = True
        logger.info(f"Work-mode surface: {len(tiles)} file tile(s) found")

        use_cdp = False
        try:
            cdp = await self.page.context.new_cdp_session(self.page)
            await cdp.send(
                "Browser.setDownloadBehavior",
                {
                    "behavior": "allowAndName",
                    "downloadPath": str(download_path.resolve()),
                    "eventsEnabled": True,
                },
            )
            use_cdp = True
        except Exception as e:
            logger.info(f"CDP download setup failed ({e}); using expect_download")

        seen: set = set()
        for tile in reversed(tiles):  # newest tiles last in DOM
            fname = (await tile.get_attribute("aria-label")) or "work_output.xlsx"
            if fname in seen:
                continue
            seen.add(fname)
            try:
                overlay_handle = await tile.evaluate_handle(
                    """el => {
                        let card = el;
                        for (let i = 0; i < 5 && card; i++) {
                            card = card.parentElement;
                            const ov = card ? card.querySelector('[data-default-action]') : null;
                            if (ov) return ov;
                        }
                        return el;
                    }"""
                )
                clicker = overlay_handle.as_element() or tile
                await clicker.evaluate(self._JS_CLICK)

                # Wait for a Download control. Two UI generations coexist:
                # the older preview dialog with aria-label "Download", and
                # (verified live 2026-08-28) a "Download file" button rendered
                # straight onto the card, with no dialog at all — so the
                # non-dialog selectors must be tried too or this path waits out
                # its deadline on a dialog that never opens. Ordered
                # dialog-first so the scoped match still wins when one exists.
                #
                # "Download apps" in the sidebar also contains "Download":
                # match the labels exactly, never by substring.
                dl_selectors = (
                    '[role="dialog"] button[aria-label="Download"]',
                    '[role="dialog"] button[aria-label="Download file"]',
                    'button[aria-label="Download file"]',
                )
                dl = None
                deadline = asyncio.get_event_loop().time() + 10
                while asyncio.get_event_loop().time() < deadline:
                    for sel in dl_selectors:
                        cand = await self.page.query_selector(sel)
                        if cand and await cand.is_visible():
                            dl = cand
                            break
                    if dl is not None:
                        break
                    await asyncio.sleep(0.5)
                if dl is None:
                    logger.warning(
                        f"Preview dialog / Download button never appeared "
                        f"for {fname!r}"
                    )
                    await self.page.keyboard.press("Escape")
                    continue

                if use_cdp:
                    files_before = set(download_path.iterdir())
                    try:
                        await dl.click(timeout=4000)
                    except Exception:
                        await dl.evaluate(self._JS_CLICK)
                    new_file = None
                    deadline = asyncio.get_event_loop().time() + 20
                    while asyncio.get_event_loop().time() < deadline:
                        fresh = [
                            f
                            for f in set(download_path.iterdir()) - files_before
                            if not f.name.endswith(".crdownload")
                            and f.is_file()
                            and f.stat().st_size > 0
                        ]
                        if fresh:
                            new_file = fresh[0]
                            break
                        await asyncio.sleep(0.4)
                    if new_file is None:
                        logger.warning(f"No file appeared for tile {fname!r}")
                    else:
                        target = download_path / fname
                        counter = 1
                        while target.exists():
                            target = download_path / (
                                f"{Path(fname).stem}_{counter}{Path(fname).suffix}"
                            )
                            counter += 1
                        new_file.rename(target)
                        saved.append(str(target))
                        logger.info(f"Downloaded work-mode file: {target}")
                else:
                    async with self.page.expect_download(
                        timeout=20000
                    ) as dl_info:
                        try:
                            await dl.click(timeout=4000)
                        except Exception:
                            await dl.evaluate(self._JS_CLICK)
                    download = await dl_info.value
                    target = download_path / fname
                    await download.save_as(str(target))
                    if target.stat().st_size == 0:
                        target.unlink(missing_ok=True)
                        continue
                    saved.append(str(target))
                    logger.info(f"Downloaded work-mode file: {target}")

                await self.page.keyboard.press("Escape")
                await asyncio.sleep(0.6)
            except Exception as e:
                logger.warning(f"Work tile download failed for {fname!r}: {e}")
                try:
                    await self.page.keyboard.press("Escape")
                except Exception:
                    pass
                continue
        return saved

    async def download_all_artifacts(
        self, download_dir: Optional[str] = None, timeout: int = 30000
    ) -> list[str]:
        """Download all Excel artifacts from the ChatGPT conversation.

        Strategy 0 pulls bytes straight from ChatGPT's backend API (see
        _download_via_backend_api); strategy 0.5 handles Work-mode file
        tiles (see _download_via_work_tiles); both silently defer to the
        DOM flow below on any failure.

        The DOM flow clicks the inline preview card ChatGPT renders for a
        produced file. Its structure is:

            paragraph
              └── generic (outer container)
                    └── generic (header row)
                          ├── generic (left: img icon + filename text)
                          └── generic (right: icon-only buttons)
                                ├── button (expand/preview)
                                └── button (download)   ← TARGET
                    └── generic (sheet tabs: "model", "answers", etc.)

        The download buttons have NO text — just an <img> icon. We identify
        artifact cards using JS to find the correct structure, then click the
        last button in the action row (which is the download button).

        Falls back to sandbox download links (/mnt/data/ paths) if no
        preview cards are found.
        """
        downloaded = []
        self.last_download_saw_buttons = False
        download_path = Path(download_dir) if download_dir else Path(".")
        download_path.mkdir(parents=True, exist_ok=True)

        # Strategy 0: backend API — no DOM scraping, no preview rendering.
        api_files = await self._download_via_backend_api(download_path)
        if api_files:
            self.last_download_saw_buttons = True
            return api_files

        # Strategy 0.5: Work-mode file tiles (preview dialog → Download).
        work_files = await self._download_via_work_tiles(download_path)
        if work_files:
            return work_files

        try:
            # Scroll to the bottom of the conversation to ensure all file cards
            # are rendered in the DOM before searching.
            await self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            await self.page.wait_for_timeout(2000)

            # Always-on diagnostic: record how the finished file is presented
            # (control form + final-message HTML) so a UI format drift is
            # self-documenting in the log rather than a silent zero-file
            # result. Only reached when the no-render strategies (backend
            # API, work tiles) found nothing.
            await dump_final_message_dom(
                self.page, logger, "chatgpt",
                ["[data-message-author-role='assistant']", "article"],
            )

            # Strategy 1: Find artifact preview cards via JS DOM inspection.
            # Only search articles AFTER the baseline count (i.e., new articles
            # from the current response), to avoid picking up file cards from
            # previous conversations in the project.
            baseline = self._baseline_article_count
            logger.info(
                f"Searching for artifacts in articles after baseline={baseline}"
            )

            scan_result = await self.page.evaluate(
                """(baseline) => {
                // File card keywords: ChatGPT may show "Spreadsheet", the actual
                // filename with .xlsx extension, or other file type labels.
                const FILE_KEYWORDS = ['.xlsx', '.xls', 'Spreadsheet', 'Excel'];

                // Roots to skip in document-wide fallback (composer, nav, etc.)
                const SKIP_TAGS = new Set(['CODE', 'PRE', 'NAV', 'HEADER', 'FORM',
                                           'TEXTAREA', 'INPUT', 'SCRIPT', 'STYLE']);

                function ancestorChain(el) {
                    const parts = [];
                    let cur = el;
                    while (cur && cur !== document.body && parts.length < 8) {
                        let label = cur.tagName;
                        const role = cur.getAttribute && cur.getAttribute('data-message-author-role');
                        if (role) label += '[role=' + role + ']';
                        else if (cur.className && typeof cur.className === 'string') {
                            const firstCls = cur.className.trim().split(/\\s+/)[0];
                            if (firstCls) label += '.' + firstCls.substring(0, 24);
                        }
                        parts.push(label);
                        cur = cur.parentElement;
                    }
                    return parts.join(' < ');
                }

                function scanRoot(root, artifactsOut, seenCards) {
                    const walker = document.createTreeWalker(
                        root,
                        NodeFilter.SHOW_TEXT,
                        { acceptNode: (node) => {
                            const text = node.textContent.trim();
                            if (!text) return NodeFilter.FILTER_REJECT;
                            const isFile = FILE_KEYWORDS.some(kw => text.includes(kw));
                            if (!isFile) return NodeFilter.FILTER_REJECT;
                            // Reject if inside a skipped ancestor, OR inside a
                            // user message (user-uploaded file cards have the
                            // same shape as assistant output cards, so only the
                            // author-role ancestor distinguishes them).
                            let el = node.parentElement;
                            while (el && el !== root) {
                                if (SKIP_TAGS.has(el.tagName)) return NodeFilter.FILTER_REJECT;
                                const role = el.getAttribute &&
                                    el.getAttribute('data-message-author-role');
                                if (role === 'user') return NodeFilter.FILTER_REJECT;
                                // Belt-and-suspenders: user-message wrapper has
                                // items-end rtl:items-start on the outer flex div.
                                if (el.className && typeof el.className === 'string' &&
                                    el.className.includes('items-end') &&
                                    el.className.includes('rtl:items-start')) {
                                    return NodeFilter.FILTER_REJECT;
                                }
                                el = el.parentElement;
                            }
                            return NodeFilter.FILTER_ACCEPT;
                        }}
                    );

                    let node;
                    while (node = walker.nextNode()) {
                        const filename = node.textContent.trim();
                        let container = node.parentElement;
                        // Aria-labels of the per-message action toolbar that sits
                        // under every assistant reply. These look like icon-only
                        // buttons and otherwise satisfy the artifact-card heuristic,
                        // so we exclude them explicitly — clicking them switches
                        // model, opens Share, downvotes the response, etc.
                        const MSG_ACTION_ARIAS = new Set([
                            'Copy response', 'Copy', 'Good response', 'Bad response',
                            'Share', 'Switch model', 'More actions', 'Edit',
                            'Read aloud', 'Try again', 'Regenerate'
                        ]);
                        for (let depth = 0; depth < 8 && container; depth++) {
                            const buttons = container.querySelectorAll('button');
                            // Artifact cards have icon-only buttons (SVG icons).
                            const iconButtons = Array.from(buttons).filter(b => {
                                const hasIcon = b.querySelector('img') || b.querySelector('svg');
                                const isSmall = b.textContent.trim().length === 0 ||
                                                b.textContent.trim().length < 5;
                                const aria = (b.getAttribute('aria-label') || '').trim();
                                if (MSG_ACTION_ARIAS.has(aria)) return false;
                                return hasIcon && isSmall;
                            });
                            const isFileCard = container.className &&
                                container.className.includes('rounded-2xl') &&
                                container.className.includes('my-4');
                            if (iconButtons.length >= 1 || isFileCard) {
                                // Dedup by button-set: if any of these icon
                                // buttons was already tagged by a previous
                                // text-node match, skip — it's the same card
                                // reached via a different walker path.
                                if (iconButtons.some(b => b.hasAttribute('data-artifact-btn'))) {
                                    break;
                                }
                                const cardIdx = artifactsOut.length;
                                // Tag EVERY icon button — we don't know which
                                // one is download (icons are sprite-hashed and
                                // classes don't disambiguate in pro mode), so
                                // Python will try each in turn.
                                const buttonMetas = iconButtons.map((b, i) => {
                                    const id = 'art-' + cardIdx + '-btn-' + i;
                                    b.setAttribute('data-artifact-btn', id);
                                    const useEl = b.querySelector('use');
                                    const rect = b.getBoundingClientRect();
                                    return {
                                        id: id,
                                        ariaLabel: b.getAttribute('aria-label') || '',
                                        title: b.getAttribute('title') || '',
                                        spriteHref: useEl ? (useEl.getAttribute('href') ||
                                                             useEl.getAttribute('xlink:href') || '') : '',
                                        classes: b.className || '',
                                        rectX: Math.round(rect.x),
                                        rectY: Math.round(rect.y),
                                        rectW: Math.round(rect.width),
                                        rectH: Math.round(rect.height),
                                        isRoundedFull: b.className.includes('rounded-full'),
                                        innerHtml: b.innerHTML.substring(0, 300),
                                    };
                                });
                                const displayName = filename.includes('.xls')
                                    ? filename : 'ai_attempt.xlsx';
                                artifactsOut.push({
                                    filename: displayName,
                                    buttons: buttonMetas,
                                    containerHtml: container.outerHTML.substring(0, 1500),
                                    foundVia: root === document.body ? 'document' : 'article',
                                });
                                break;
                            }
                            container = container.parentElement;
                        }
                    }
                }

                // Assistant turns, in the two shapes chatgpt.com serves:
                // short plain-text answers carry the author-role attribute,
                // longer agentic ones only [data-turn]. Same tiers as
                // _extract_last_response / _count_response_articles.
                let responseElements = Array.from(
                    document.querySelectorAll("[data-message-author-role='assistant']")
                );
                if (responseElements.length === 0) {
                    responseElements = Array.from(document.querySelectorAll('[data-turn]'))
                        .filter(el => el.getAttribute('data-turn') !== 'user');
                }
                const totalAssistant = responseElements.length;
                const newArticles = responseElements.slice(baseline);
                const artifacts = [];
                const seenCards = new Set();

                // Pass 1: scan assistant articles after the baseline.
                for (const article of newArticles) {
                    scanRoot(article, artifacts, seenCards);
                }

                // Pass 2 (fallback): if nothing found, scan entire document body.
                // Pro mode / canvas sometimes renders file cards outside the
                // assistant article (e.g. side panel, attachment tray).
                let fallbackUsed = false;
                if (artifacts.length === 0) {
                    fallbackUsed = true;
                    scanRoot(document.body, artifacts, seenCards);
                }

                // Diagnostics: where do file-keyword text nodes live on the page?
                const sightings = [];
                const diagWalker = document.createTreeWalker(
                    document.body,
                    NodeFilter.SHOW_TEXT,
                    { acceptNode: (node) => {
                        const text = node.textContent.trim();
                        if (!text) return NodeFilter.FILTER_REJECT;
                        return FILE_KEYWORDS.some(kw => text.includes(kw))
                            ? NodeFilter.FILTER_ACCEPT
                            : NodeFilter.FILTER_REJECT;
                    }}
                );
                let dn;
                while ((dn = diagWalker.nextNode()) && sightings.length < 10) {
                    sightings.push({
                        text: dn.textContent.trim().substring(0, 120),
                        ancestors: ancestorChain(dn.parentElement),
                    });
                }

                return {
                    artifacts,
                    diagnostics: {
                        totalAssistantArticles: totalAssistant,
                        baselineSkipped: baseline,
                        newArticlesScanned: newArticles.length,
                        fallbackUsed,
                        fileKeywordSightings: sightings,
                    },
                };
            }""",
                baseline,
            )

            artifact_info = scan_result["artifacts"]
            self.last_download_saw_buttons = bool(artifact_info)
            diag = scan_result["diagnostics"]
            logger.info(
                f"Found {len(artifact_info)} artifact preview card(s) "
                f"(fallback={diag['fallbackUsed']}, "
                f"assistant_articles={diag['totalAssistantArticles']}, "
                f"new_scanned={diag['newArticlesScanned']})"
            )
            if not artifact_info:
                sightings = diag["fileKeywordSightings"]
                if sightings:
                    logger.warning(
                        f"No cards matched, but {len(sightings)} file-keyword "
                        f"text node(s) exist on page. Ancestor chains:"
                    )
                    for s in sightings:
                        logger.warning(f"  '{s['text']}' in {s['ancestors']}")
                else:
                    logger.warning(
                        "No file-keyword text nodes found anywhere on page — "
                        "response may not have produced a file yet."
                    )

            # Set up CDP download behavior so Chrome saves files to our
            # directory.  page.expect_download() does NOT work on CDP
            # connections — Chrome handles downloads natively.
            download_path.mkdir(parents=True, exist_ok=True)
            use_cdp_download = False
            try:
                cdp = await self.page.context.new_cdp_session(self.page)
                await cdp.send(
                    "Browser.setDownloadBehavior",
                    {
                        "behavior": "allowAndName",
                        "downloadPath": str(download_path.resolve()),
                        "eventsEnabled": True,
                    },
                )
                use_cdp_download = True
                logger.info(f"CDP download path: {download_path.resolve()}")
            except Exception as e:
                logger.info(
                    f"CDP download setup failed ({e}), using Playwright fallback"
                )

            # Per-button click budget. Each card may have N icon buttons and
            # we try them in order; the REAL download button is the first one
            # that causes a new file to appear within this window.
            per_button_wait_sec = 6

            for info in artifact_info:
                filename = info["filename"]
                card_html = info.get("containerHtml", "")
                buttons = info.get("buttons", [])
                found_via = info.get("foundVia", "?")
                logger.info(
                    f"Artifact card for {filename} (via {found_via}) has "
                    f"{len(buttons)} icon button(s):"
                )
                for b in buttons:
                    logger.info(
                        f"  - {b['id']} rounded_full={b['isRoundedFull']} "
                        f"aria={b['ariaLabel']!r} title={b['title']!r} "
                        f"sprite={b['spriteHref']!r} "
                        f"rect=({b['rectX']},{b['rectY']},{b['rectW']}x{b['rectH']})"
                    )

                if not buttons:
                    logger.warning(
                        f"No icon buttons for {filename}. Card DOM: {card_html}"
                    )
                    continue

                # Ordering priority:
                #   0. Button whose SVG sprite href matches a known download
                #      icon fragment. Sprite IDs are stable *within* a ChatGPT
                #      build but rotate on redeploys, so we maintain a list of
                #      known hints. Update DOWNLOAD_SPRITE_HINTS when a new
                #      build rolls out and the log shows a different sprite.
                #   1. rounded-full icon buttons (broad fallback — download is
                #      almost always styled as a round icon button).
                #   2. anything else.
                DOWNLOAD_SPRITE_HINTS = ["#1a3695"]

                def button_rank(b):
                    sprite = b.get("spriteHref", "") or ""
                    if any(hint in sprite for hint in DOWNLOAD_SPRITE_HINTS):
                        return 0
                    if b.get("isRoundedFull"):
                        return 1
                    return 2

                ordered = sorted(buttons, key=button_rank)

                saved_path: Optional[Path] = None
                tried_button_ids = []
                for btn_meta in ordered:
                    btn_id = btn_meta["id"]
                    tried_button_ids.append(btn_id)
                    btn = self.page.locator(f'[data-artifact-btn="{btn_id}"]')
                    if await btn.count() == 0:
                        logger.info(f"  {btn_id}: locator missing, skipping")
                        continue

                    try:
                        await btn.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass

                    if use_cdp_download:
                        files_before = set(download_path.iterdir())
                        try:
                            await btn.click(force=True, timeout=5000)
                        except Exception as click_err:
                            logger.info(
                                f"  {btn_id}: click(force=True) failed "
                                f"({click_err}); dispatching DOM click"
                            )
                            try:
                                await btn.dispatch_event("click")
                            except Exception as de:
                                logger.info(f"  {btn_id}: dispatch_event failed ({de})")
                                continue

                        # Wait per-button for a new completed file.
                        deadline = (
                            asyncio.get_event_loop().time() + per_button_wait_sec
                        )
                        new_file = None
                        while asyncio.get_event_loop().time() < deadline:
                            new_files = set(download_path.iterdir()) - files_before
                            complete = [
                                f for f in new_files
                                if not f.name.endswith(".crdownload")
                            ]
                            if complete:
                                new_file = complete[0]
                                break
                            await asyncio.sleep(0.3)

                        if new_file:
                            target = download_path / filename
                            if new_file.name != filename:
                                new_file.rename(target)
                                new_file = target
                            logger.info(
                                f"  {btn_id}: triggered download -> {new_file}"
                            )
                            saved_path = new_file
                            break
                        else:
                            logger.info(
                                f"  {btn_id}: no file within {per_button_wait_sec}s"
                            )
                            # Clicking a non-download button (preview, expand)
                            # may open a canvas/modal that occludes the
                            # remaining buttons. Dismiss before the next try.
                            try:
                                await self.page.keyboard.press("Escape")
                                await asyncio.sleep(0.3)
                            except Exception:
                                pass
                    else:
                        # Playwright download event path (non-CDP connections)
                        try:
                            async with self.page.expect_download(
                                timeout=per_button_wait_sec * 1000
                            ) as dl_info:
                                try:
                                    await btn.click(force=True, timeout=5000)
                                except Exception:
                                    await btn.dispatch_event("click")
                            download = await dl_info.value
                            target = download_path / filename
                            await download.save_as(str(target))
                            logger.info(
                                f"  {btn_id}: triggered download -> {target}"
                            )
                            saved_path = target
                            break
                        except Exception as e:
                            logger.info(f"  {btn_id}: no download event ({e})")
                            continue

                if saved_path:
                    downloaded.append(str(saved_path))
                else:
                    logger.warning(
                        f"None of {tried_button_ids} produced a download for "
                        f"{filename}. Full card DOM: {card_html}"
                    )

            if downloaded:
                return downloaded

            # Strategy 2: Fallback — sandbox download links in NEW turns only.
            # The model sometimes links a file saved to /mnt/data/ instead of
            # rendering an inline preview card.
            logger.info("No preview cards found, trying sandbox download links...")
            download_links = await self.page.evaluate(
                """(baseline) => {
                let assistantTurns = Array.from(
                    document.querySelectorAll("[data-message-author-role='assistant']")
                );
                if (assistantTurns.length === 0) {
                    assistantTurns = Array.from(document.querySelectorAll('[data-turn]'))
                        .filter(el => el.getAttribute('data-turn') !== 'user');
                }
                const newArticles = assistantTurns.slice(baseline);
                const links = [];
                for (const article of newArticles) {
                    const articleLinks = Array.from(article.querySelectorAll('a[href*="sandbox"]'));
                    for (const a of articleLinks) {
                        if (a.href && (a.href.includes('.xlsx') || a.href.includes('.xls'))) {
                            links.push({ href: a.href, text: a.textContent.trim() });
                        }
                    }
                }
                return links;
            }""",
                baseline,
            )

            for link_info in download_links:
                try:
                    logger.info(f"Trying sandbox link: {link_info['text']}")
                    link = self.page.locator(
                        f'a[href*="sandbox"]:has-text("{link_info["text"]}")'
                    )

                    if use_cdp_download:
                        files_before = set(download_path.iterdir())
                        await link.first.click()

                        deadline = asyncio.get_event_loop().time() + timeout / 1000
                        save_path = None
                        while asyncio.get_event_loop().time() < deadline:
                            current_files = set(download_path.iterdir())
                            new_files = current_files - files_before
                            complete = [
                                f
                                for f in new_files
                                if not f.name.endswith(".crdownload")
                            ]
                            if complete:
                                save_path = complete[0]
                                break
                            await asyncio.sleep(0.5)

                        if save_path:
                            logger.info(f"Downloaded via sandbox link: {save_path}")
                            downloaded.append(str(save_path))
                        else:
                            logger.warning(
                                f"Sandbox download timeout for {link_info['text']}"
                            )
                    else:
                        async with self.page.expect_download(
                            timeout=timeout
                        ) as dl_info:
                            await link.first.click()
                        download = await dl_info.value
                        save_path = download_path / download.suggested_filename
                        await download.save_as(str(save_path))
                        logger.info(f"Downloaded via sandbox link: {save_path}")
                        downloaded.append(str(save_path))

                except Exception as e:
                    logger.warning(
                        f"Sandbox download failed for {link_info['text']}: {e}"
                    )
                    continue

            if not downloaded:
                logger.warning(
                    "No artifacts downloaded (no preview cards or sandbox links found)"
                )

        except Exception as e:
            logger.error(f"Artifact download failed: {e}")

        return downloaded

    async def get_conversation_history(self) -> list[dict]:
        """Extract conversation as list of message dicts.

        Reads [data-turn], whose attribute value IS the role, so both sides
        of the conversation come from one query in document order.
        [data-message-author-role] is the fallback for a page that renders
        the attribute but not the turn wrapper.

        Verified live 2026-08-20: chatgpt.com serves no <article> and no <h6>
        elements at all (0 of each across 14 conversations), so the heading
        walk this used to do returned an empty transcript on every run.
        """
        messages = []
        try:
            messages = await self.page.evaluate(
                """() => {
                const read = (el, role) => ({
                    role: role === 'user' ? 'user' : 'assistant',
                    content: (el.innerText || '').trim(),
                });
                let turns = Array.from(document.querySelectorAll('[data-turn]'));
                if (turns.length > 0) {
                    return turns.map(el => read(el, el.getAttribute('data-turn')));
                }
                turns = Array.from(
                    document.querySelectorAll('[data-message-author-role]')
                );
                return turns.map(
                    el => read(el, el.getAttribute('data-message-author-role'))
                );
            }"""
            )
            messages = [m for m in messages if m.get("content")]
        except Exception as e:
            logger.error(f"Conversation history extraction failed: {e}")

        return messages

    async def process_all_prompts(self, files_to_upload: list = None) -> bool:
        """Process all prompts: upload files, enable features, send prompts, wait."""
        prompts = self.config.get("prompts", [])
        if not prompts:
            logger.error("No prompts configured")
            return False

        # Upload files if provided (engine may have already uploaded in Phase 1)
        if files_to_upload:
            if not await self.upload_files(files_to_upload):
                logger.error("File upload failed")
                return False

        # Process each prompt
        for i, prompt in enumerate(prompts, 1):
            logger.info(f"Processing prompt {i}/{len(prompts)}")

            if not await self.submit_prompt(prompt, i):
                logger.error(f"Failed to submit prompt {i}")
                return False

            response = await self.wait_for_response(i)
            if response is None:
                logger.error(f"No response for prompt {i}")
                return False

            self.messages.append(
                ConversationMessage(
                    role="user", content=prompt, timestamp=datetime.now()
                )
            )
            self.messages.append(
                ConversationMessage(
                    role="assistant", content=response, timestamp=datetime.now()
                )
            )
            self.current_response_count += 1

        logger.info(f"All {len(prompts)} prompt(s) processed successfully")
        return True
