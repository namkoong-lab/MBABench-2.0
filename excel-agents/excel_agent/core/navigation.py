"""
Navigation utilities for OneDrive/Excel Online.

Handles folder navigation and file opening in OneDrive interface.
"""

import asyncio
import logging
import re

logger = logging.getLogger(__name__)


def _normalize_name(value: str) -> str:
    """
    Normalize filename/folder name for exact matching (ignoring special chars).

    Removes spaces, hyphens, underscores, punctuation and converts to lowercase.
    Keeps only alphanumeric characters for comparison.
    """
    value = (value or "").lower()
    # Keep alphanumerics only (removes spaces, hyphens, dots, etc)
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


class Navigation:
    """Handles navigation in OneDrive/Excel Online."""

    @staticmethod
    async def _click_folder_fuzzy(
        page, desired_name: str, max_attempts: int = 30, step_px: int = 300
    ) -> bool:
        """
        Find and click folder/file by normalized name (ignoring special characters).

        Normalizes both the target and each row's NAME-bearing elements
        (anchor/button/titled-span texts and title attributes) and clicks
        only on EXACT normalized equality. Substring matching against the
        whole row text is deliberately not done — row text includes dates
        and sizes, and "round1" must never match a "round10" row.

        Scrolls incrementally (small steps) so we don't overshoot the target.

        Args:
            page: Playwright page instance
            desired_name: Name of folder/file to find
            max_attempts: Maximum scroll attempts (default 30)
            step_px: Pixels to scroll per attempt (default 300)

        Returns:
            True if folder/file was clicked, False otherwise
        """
        target_norm = _normalize_name(desired_name)

        async def row_name_matches(row) -> bool:
            """True iff any name-bearing element in the row equals the
            target after normalization."""
            candidates: list[str] = []
            try:
                cells = row.locator("a, button, [title], [data-automationid*='name' i]")
                n = await cells.count()
            except Exception:
                return False
            for j in range(min(n, 12)):
                cell = cells.nth(j)
                try:
                    title = await cell.get_attribute("title")
                    if title:
                        candidates.append(title)
                    text = await cell.text_content()
                    if text:
                        candidates.append(text)
                except Exception:
                    continue
            return any(_normalize_name(c) == target_norm for c in candidates)

        async def scan_visible_rows() -> bool:
            try:
                rows = page.locator('[role="row"]')
                count = await rows.count()
            except Exception:
                return False
            for i in range(min(count, 50)):
                row = rows.nth(i)
                if not await row_name_matches(row):
                    continue
                try:
                    await row.scroll_into_view_if_needed()
                    await asyncio.sleep(0.05)
                    # OneDrive requires double-click to open folders
                    await row.dblclick(force=True)
                    return True
                except Exception:
                    continue
            return False

        # Try without scrolling first
        if await scan_visible_rows():
            return True

        # Scroll incrementally and scan after each step
        for _ in range(max_attempts):
            try:
                await page.mouse.wheel(0, step_px)
            except Exception:
                pass
            await asyncio.sleep(0.3)
            if await scan_visible_rows():
                return True
        return False

    @staticmethod
    async def _try_find_and_click_folder(page, folder_name: str) -> bool:
        """
        Helper to try finding and clicking a folder (used for retry logic).

        Returns:
            True if folder was clicked, False otherwise
        """
        # Wait for rows to render (quick check)
        try:
            await page.wait_for_selector('[role="row"]', timeout=800)
        except Exception:
            await asyncio.sleep(0.6)  # Slowed down for slow browser

        # Exact-match selectors only. Substring forms (has-text,
        # [title*=]) collide on numbered task names — "round1" would
        # match "round10" — so anything looser goes through the
        # normalized-equality fuzzy scan below instead.
        folder_selectors = [
            f'text="{folder_name}"',
            f'[aria-label="{folder_name}"]',
            f'[title="{folder_name}"]',
        ]

        clicked = False
        for selector in folder_selectors:
            try:
                element = await page.query_selector(selector)
                if element and await element.is_visible():  # Check visibility
                    # Scroll element into view if needed
                    await element.scroll_into_view_if_needed()
                    await asyncio.sleep(0.3)  # Slowed down for slow browser

                    # OneDrive requires double-click to open folders
                    await element.dblclick(force=True)
                    logger.info(f"✅ Double-clicked folder: {folder_name} ({selector})")
                    clicked = True
                    break
            except Exception as e:
                logger.debug(f"🔍 Selector {selector} failed: {e}")
                continue

        if not clicked:
            logger.debug(
                f"🔍 Direct selectors failed, trying fuzzy scan: {folder_name}"
            )
            try:
                clicked = await Navigation._click_folder_fuzzy(
                    page, folder_name, max_attempts=30, step_px=300
                )
                if clicked:
                    logger.info(f"✅ Clicked folder via fuzzy scan: {folder_name}")
            except Exception as e:
                logger.debug(f"🔍 Fuzzy scan failed: {e}")

        return clicked

    @staticmethod
    async def navigate_to_direct_url(page, url: str):
        """
        Navigate directly to a OneDrive/SharePoint URL, bypassing folder navigation.

        Use this when you have a direct link to a task folder instead of navigating
        through the OneDrive folder hierarchy.

        Args:
            page: Playwright page instance
            url: Full OneDrive or SharePoint URL to the task folder

        Returns:
            Page object or None if navigation failed
        """
        try:
            logger.info("Navigating to direct URL: %s", url)
            await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            await asyncio.sleep(3)

            # Wait for file list to appear
            for retry in range(5):
                try:
                    await page.wait_for_selector(
                        '[role="row"], [data-automationid*="item"]', timeout=5000
                    )
                    logger.info(
                        "Direct URL navigation successful — folder contents loaded"
                    )
                    return page
                except Exception:
                    if retry < 4:
                        logger.info(
                            "Waiting for folder contents (attempt %d/5)...", retry + 1
                        )
                        await asyncio.sleep(2.0)
                    else:
                        logger.warning(
                            "Folder contents not detected, but page loaded — continuing"
                        )
                        return page

        except Exception as e:
            logger.error("Failed to navigate to direct URL: %s", e)
            return None

    @staticmethod
    def _build_shorthand_path(
        task_name: str,
        task_source: str,
        base_path: list = None,
    ) -> list:
        """
        Build a OneDrive folder path from the task-source shorthand fields.

        Path layout:
            base_path + source_folder + task_name + "Task"
        where source_folder is the task_source itself, except for the
        historical alias `wallstreetprep` -> `wsp`.
        """
        if not task_source:
            raise ValueError(
                "task_source is required for shorthand navigation. "
                "Set `task_source` in your task config, or provide an "
                "explicit `onedrive_path` / `direct_url` instead."
            )

        if base_path is None:
            base_path = ["My files", "main_tasks"]

        # Historical alias kept for backward compatibility with the bundled
        # example sources. Any other source name is used as-is.
        folder = "wsp" if task_source == "wallstreetprep" else task_source
        return base_path + [folder, task_name, "Task"]

    @staticmethod
    async def navigate_to_task_folder(
        page,
        task_name: str,
        task_source: str = "",
        base_path: list = None,
        direct_url: str = None,
        onedrive_path: list = None,
    ):
        """
        Navigate to the task folder on OneDrive.

        Resolution order:
            1. direct_url       — goto() the URL, skip folder navigation
            2. onedrive_path    — click through these exact folder segments
            3. Task-source shorthand — build path from base_path + task_source + task_name

        Args:
            page: Playwright page instance.
            task_name: Name of the task.
            task_source: Optional source name used by the shorthand path
                builder (any string is accepted; doubles as the folder name).
            base_path: Optional OneDrive path segments used as the parent of
                the source folder (defaults to ["My files", "main_tasks"]).
            direct_url: Direct OneDrive/SharePoint URL to the task folder.
            onedrive_path: Explicit list of OneDrive folder names to click through.

        Returns:
            Page object or None if navigation failed.
        """
        # Priority 1: Direct URL
        if direct_url:
            return await Navigation.navigate_to_direct_url(page, direct_url)

        try:
            # Priority 2: Explicit onedrive_path
            if onedrive_path:
                file_path = list(onedrive_path)
                logger.info(f"📁 Using explicit onedrive_path: {' > '.join(file_path)}")
            else:
                # Priority 3: Task-source shorthand path construction
                file_path = Navigation._build_shorthand_path(
                    task_name, task_source, base_path
                )

            logger.info(
                f"📁 Navigating to Task folder ({task_source or 'custom'}): "
                f"{' > '.join(file_path)}"
            )

            # Navigate through the folder structure
            for i, folder_name in enumerate(file_path):
                logger.info(
                    f"📂 Step {i+1}/{len(file_path)}: Looking for folder: {folder_name}"
                )

                # Try to find and click folder (fast attempt first)
                clicked = await Navigation._try_find_and_click_folder(page, folder_name)

                # If fast attempt failed, wait longer and retry
                if not clicked:
                    logger.info(
                        f"⚠️ First attempt failed for '{folder_name}', waiting longer and retrying..."
                    )
                    await asyncio.sleep(5)  # Longer wait for slow-loading pages
                    clicked = await Navigation._try_find_and_click_folder(
                        page, folder_name
                    )

                if not clicked:
                    logger.error(f"❌ Could not find folder: {folder_name}")
                    logger.error(
                        f"   Failed at step {i+1}/{len(file_path)} of navigation"
                    )
                    return None

                # Wait for navigation and content loading
                await asyncio.sleep(1.0)

                # Try to wait for file/folder list to appear with retries
                for retry in range(3):
                    try:
                        await page.wait_for_selector(
                            '[role="row"], [data-automationid*="item"]', timeout=3000
                        )
                        logger.info("✅ Folder contents loaded")
                        break
                    except Exception:
                        if retry < 2:
                            logger.info(
                                f"⚠️ Waiting for folder contents (attempt {retry + 1}/3)..."
                            )
                            await asyncio.sleep(2.0)  # Slowed down for slow browser
                        else:
                            # Last attempt failed, give a brief moment for content to load
                            await asyncio.sleep(1.0)  # Slowed down for slow browser

            logger.info("✅ Successfully navigated to Task folder")
            return page

        except Exception as e:
            logger.error(f"❌ Error navigating to task folder: {e}")
            return None

    @staticmethod
    async def open_specific_file(page, file_name: str):
        """
        Open a specific file by name.

        Args:
            page: Playwright page instance
            file_name: Name of file to open

        Returns:
            Page object of opened file, or None if failed
        """
        try:
            logger.info(f"📄 Opening file: {file_name}")

            # Wait for folder contents to load
            await asyncio.sleep(1.0)

            # Try to wait for file list to be ready with retries
            for retry in range(3):
                try:
                    await page.wait_for_selector(
                        '[data-automationid*="file"], [data-automationid*="item"], .ms-List-cell, .od-ItemTile',
                        timeout=4000,  # Increased timeout for slow browser
                    )
                    logger.info("✅ File list is ready")
                    break
                except Exception:
                    if retry < 2:
                        logger.info(
                            f"⚠️ Waiting for file list (attempt {retry + 1}/3)..."
                        )
                        await asyncio.sleep(2.0)  # Slowed down for slow browser
                    else:
                        logger.info(
                            "⚠️ File list not detected, but continuing with search..."
                        )
                        await asyncio.sleep(1.0)  # Slowed down for slow browser

            # Exact-match selectors only. The old substring forms —
            # especially div:has-text — matched the outermost container
            # of the whole file list and force-clicked its center,
            # opening an arbitrary file. Looser matching goes through the
            # normalized-equality fuzzy scan below.
            file_selectors = [
                f'text="{file_name}"',
                f'[aria-label="{file_name}"]',
                f'[title="{file_name}"]',
            ]

            clicked = False
            for selector in file_selectors:
                try:
                    element = await page.query_selector(selector)
                    if element and await element.is_visible():  # Check visibility
                        # Scroll element into view if needed
                        await element.scroll_into_view_if_needed()
                        await asyncio.sleep(0.3)  # Slowed down for slow browser

                        # Try to click with force=True to bypass overlay issues
                        await element.click(force=True)
                        logger.info(f"✅ Clicked file: {file_name} ({selector})")
                        clicked = True
                        break
                except Exception as e:
                    logger.debug(f"🔍 Selector {selector} failed: {e}")
                    continue

            if not clicked:
                logger.info(
                    f"🔍 Direct selectors failed, trying fuzzy row scan for: {file_name}"
                )

                # Use fuzzy matching (same as folder navigation)
                try:
                    clicked = await Navigation._click_folder_fuzzy(
                        page, file_name, max_attempts=25, step_px=300
                    )
                    if clicked:
                        logger.info(f"✅ Found file via fuzzy scan: {file_name}")
                except Exception as e:
                    logger.error(f"❌ Fuzzy scan failed: {e}")

            if not clicked:
                logger.error(f"❌ Could not find file: {file_name}")
                return None

            # Wait for file to open (Excel files typically open in new tab)
            await asyncio.sleep(2.0)

            # Check if a new tab opened - wait and retry for slow browser
            context = page.context
            new_tab_appeared = False
            for tab_check in range(3):
                if len(context.pages) > 1:
                    new_tab_appeared = True
                    break
                if tab_check < 2:
                    logger.info(
                        f"⚠️ Waiting for new tab to appear (check {tab_check + 1}/3)..."
                    )
                    await asyncio.sleep(2.0)  # Wait longer for slow browser

            if new_tab_appeared or len(context.pages) > 1:
                # Switch to the new tab (last one opened)
                new_page = context.pages[-1]
                logger.info("✅ File opened in new tab, switching to it...")

                try:
                    await new_page.wait_for_load_state(
                        "load", timeout=15000
                    )  # Increased timeout for slow browser
                    logger.info("✅ File loaded successfully")
                    return new_page
                except Exception as e:
                    logger.warning(f"⚠️ File load timeout: {e}")
                    # Give it one more moment before returning
                    await asyncio.sleep(1.0)
                    return new_page
            else:
                logger.info("✅ File opened in same tab")
                return page

        except Exception as e:
            logger.error(f"❌ Error opening file: {e}")
            return None
