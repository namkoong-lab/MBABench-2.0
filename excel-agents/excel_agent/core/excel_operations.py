"""
Excel operations for the Excel Agent Engine.

Handles Excel file creation, renaming, and workbook operations.
"""

import asyncio
import logging

from .browser_manager import get_save_as_key

logger = logging.getLogger(__name__)


class ExcelOperations:
    """Handles Excel file operations in Excel Online."""

    @staticmethod
    async def _try_click_element(page, text: str, timeout: int = 3000) -> bool:
        """
        Helper to try clicking an element in frames or main page.
        Uses multiple selector strategies for better reliability.
        Optimized to check quickly first, then wait if needed.

        Returns:
            True if element was clicked, False otherwise
        """
        # Build multiple selector strategies
        selectors = [text]  # Start with original selector

        # Also try case variations if it's a text selector
        if text.startswith("text="):
            text_part = text.split("=", 1)[1]
            selectors.extend(
                [
                    f'text="{text_part}"',  # Quoted version
                    f"text={text_part.lower()}",  # Lowercase
                ]
            )

        # Quick check first: try query_selector (non-blocking) before wait_for_selector
        # This avoids long waits if element is already visible
        for frame in page.frames:
            for selector in selectors:
                try:
                    element = await frame.query_selector(selector)
                    if element and await element.is_visible():
                        await element.click(timeout=1000)
                        logger.debug(
                            f"✅ Clicked '{text}' in frame using selector: {selector} (quick check)"
                        )
                        return True
                except Exception:
                    continue

        # Quick check on main page
        for selector in selectors:
            try:
                element = await page.query_selector(selector)
                if element and await element.is_visible():
                    await element.click(timeout=1000)
                    logger.debug(
                        f"✅ Clicked '{text}' in main page using selector: {selector} (quick check)"
                    )
                    return True
            except Exception:
                continue

        # If quick check failed, wait for element to appear (with shorter per-attempt timeout)
        # Use a fraction of timeout per selector to avoid long waits
        per_selector_timeout = max(
            500, timeout // (len(selectors) * 2)
        )  # At least 500ms, but divide timeout

        # Try frames first (where Excel UI usually is)
        for frame in page.frames:
            for selector in selectors:
                try:
                    element = await frame.wait_for_selector(
                        selector, timeout=per_selector_timeout, state="visible"
                    )
                    await element.click(timeout=1000)
                    logger.debug(
                        f"✅ Clicked '{text}' in frame using selector: {selector}"
                    )
                    return True
                except Exception:
                    continue

        # Try main page
        for selector in selectors:
            try:
                element = await page.wait_for_selector(
                    selector, timeout=per_selector_timeout, state="visible"
                )
                await element.click(timeout=1000)
                logger.debug(
                    f"✅ Clicked '{text}' in main page using selector: {selector}"
                )
                return True
            except Exception:
                continue

        return False

    @staticmethod
    async def _find_input(page, selectors: list) -> tuple:
        """
        Helper to find input field in frames or main page.

        Returns:
            Tuple of (input_element, location_description) or (None, None)
        """
        # Try frames first
        for frame in page.frames:
            for selector in selectors:
                try:
                    element = await frame.query_selector(selector)
                    if element and await element.is_visible():
                        return element, f"frame ({selector})"
                except Exception:
                    continue

        # Try main page
        for selector in selectors:
            try:
                element = await page.query_selector(selector)
                if element and await element.is_visible():
                    return element, f"main page ({selector})"
            except Exception:
                continue

        return None, None

    @staticmethod
    async def rename_workbook(
        page, new_name: str, retry: bool = False, start_number: int = 0
    ) -> tuple[bool, int]:
        """
        Rename the workbook using File tab -> Create a Copy method with incremental naming.

        Args:
            page: Playwright page instance
            new_name: New name for the workbook
            retry: If True, uses longer waits (for retry attempts)
            start_number: Starting number for filename attempts (for retries)

        Returns:
            Tuple of (success: bool, next_number: int) - next_number is where to start if retrying
        """
        try:
            logger.info(f"✏️ Renaming Excel file to: {new_name}")
            await asyncio.sleep(4 if retry else 2)  # longer settle on retries

            # Retry all 3 steps together as a sequence
            max_retries = 5
            sequence_complete = False

            for attempt in range(max_retries):
                if attempt > 0:
                    logger.info(
                        f"🔄 Retrying entire sequence (attempt {attempt + 1}/{max_retries})..."
                    )
                    # Close any open menus/dialogs before retry
                    try:
                        await page.keyboard.press("Escape")
                        await asyncio.sleep(1.0)
                        await page.keyboard.press("Escape")
                        await asyncio.sleep(1.0)
                    except Exception:
                        pass
                    # Exponential backoff: 2s, 4s, 6s, 8s
                    wait_time = 2.0 * attempt
                    if wait_time > 0:
                        await asyncio.sleep(wait_time)

                # Step 1: Click File tab
                logger.info("🔍 Step 1: Looking for File tab...")
                if not await ExcelOperations._try_click_element(
                    page, "text=File", timeout=6000
                ):
                    logger.warning(
                        f"⚠️ File tab not found (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                logger.info("✅ Clicked File tab")
                await asyncio.sleep(4.0)  # Wait for File menu to expand

                # Step 2: Click "Create a Copy"
                logger.info("🔍 Step 2: Looking for 'Create a Copy' button...")
                if not await ExcelOperations._try_click_element(
                    page, "text=Create a Copy", timeout=8000
                ):
                    logger.warning(
                        f"⚠️ Create a Copy button not found (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                logger.info("✅ Clicked 'Create a Copy'")

                # Step 3: Click "Create a copy online" (the submenu item)
                # Instead of verifying the submenu appeared separately, just
                # try clicking the item directly with a generous timeout.
                # The old verification loop iterated all frames × selectors
                # with 5s timeouts each, taking ~2 min on slow VMs while the
                # submenu disappeared.
                logger.info("🔍 Step 3: Looking for 'Create a copy online'...")
                await asyncio.sleep(2.0)  # Brief wait for submenu to render

                create_copy_online_clicked = False
                submenu_selectors = [
                    'text="Create a copy online"',
                    "text=Create a copy online",
                    "text=create a copy online",
                ]

                # Try up to 10s total with quick polling
                for poll_round in range(5):
                    for selector_variant in submenu_selectors:
                        if await ExcelOperations._try_click_element(
                            page, selector_variant, timeout=2000
                        ):
                            create_copy_online_clicked = True
                            break
                    if create_copy_online_clicked:
                        break
                    await asyncio.sleep(1.0)

                if not create_copy_online_clicked:
                    logger.warning(
                        f"⚠️ 'Create a copy online' not found (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                logger.info("✅ Clicked 'Create a copy online'")
                await asyncio.sleep(2.0)

                # Verify Step 3 worked by checking if filename input appears
                verification_passed = False
                for verify_frame in page.frames:
                    try:
                        element = await verify_frame.wait_for_selector(
                            'input[type="text"]', timeout=3000, state="visible"
                        )
                        if element:
                            verification_passed = True
                            break
                    except Exception:
                        continue
                if not verification_passed:
                    logger.warning(
                        f"⚠️ Create a copy online click verification failed (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                logger.debug("✅ Verified 'Create a copy online' click succeeded")

                # All steps completed successfully
                sequence_complete = True
                break

            if not sequence_complete:
                logger.warning(
                    "⚠️ Could not complete File > Create a Copy > Create a copy online sequence after retries"
                )
                return (False, start_number)

            await asyncio.sleep(3.0)  # Final wait for dialog to be fully ready

            # Step 4: Handle the dialog with incremental naming
            logger.info("🔍 Step 4: Looking for filename input...")
            input_selectors = [
                'input[type="text"]',
                'input[aria-label*="name" i]',
                'input[placeholder*="name" i]',
                "input",
            ]

            name_input, location = await ExcelOperations._find_input(
                page, input_selectors
            )
            if not name_input:
                logger.warning("⚠️ Could not find filename input")
                return (False, start_number)
            logger.info(f"✅ Found filename input in {location}")

            # Try incremental naming (don't add .xlsx as it's automatically included)
            base_name = (
                new_name.replace(" ", "_").replace(".xlsx", "").replace(".XLSX", "")
            )

            for attempt in range(start_number, start_number + 30):
                file_name = f"{attempt}_{base_name}"
                logger.info(
                    f"🔄 Trying filename: {file_name} (attempt {attempt - start_number + 1}/30)"
                )

                # Fill the input field (handle stale reference by refinding)
                try:
                    await name_input.click()
                    await name_input.select_text()
                    await name_input.fill(file_name)
                except Exception:
                    logger.debug("⚠️ Input field stale, refinding...")
                    name_input, location = await ExcelOperations._find_input(
                        page, input_selectors
                    )
                    if not name_input:
                        logger.warning("⚠️ Could not re-find filename input")
                        return (False, start_number)
                    await name_input.click()
                    await name_input.select_text()
                    await name_input.fill(file_name)

                await asyncio.sleep(1)  # Give UI time to react to input

                # Click Create a Copy button in dialog
                if not await ExcelOperations._try_click_element(
                    page, 'button:has-text("Create a Copy")', timeout=5000
                ):
                    logger.warning(
                        f"⚠️ Could not find 'Create a Copy' button in dialog for {file_name}"
                    )
                    # UI bug - return failure so outer retry can try again with same number
                    return (False, attempt)  # Return SAME attempt number to retry
                logger.info("✅ Clicked 'Create a Copy' button")
                await asyncio.sleep(2)

                # Check if "Rename" button exists (file name collision)
                if await ExcelOperations._try_click_element(
                    page, 'button:has-text("Rename")', timeout=2000
                ):
                    logger.info(
                        f"⚠️ File '{file_name}.xlsx' exists (confirmed), trying next number"
                    )
                    await asyncio.sleep(0.5)
                    continue  # Move to next filename

                logger.info(f"✅ Successfully created copy as '{file_name}.xlsx'")
                return (True, attempt + 1)  # Success - return next number

            logger.warning(
                f"⚠️ Could not find available filename after 30 attempts (tried {start_number}-{start_number + 29})"
            )
            return (False, start_number + 30)
        except Exception as e:
            logger.error(f"❌ Rename error: {e}")
            return (False, start_number)

    @staticmethod
    async def create_solution_model(page, base_name: str = None):
        """
        Create a new Excel file (does NOT rename it).

        Args:
            page: Playwright page instance
            base_name: Unused, kept for compatibility

        Returns:
            Page object of new Excel file, or False on failure
        """
        try:
            logger.info("📊 Creating new Excel workbook for solution model...")

            # Click "Create or upload" button
            new_selectors = [
                'span.text_89a70e57:has-text("Create or upload")',
                'span:has-text("Create or upload")',
                'text="Create or upload"',
                'button:has-text("New")',
                '[data-automationid="newCommand"]',
                '.ms-Button:has-text("New")',
                '[aria-label="New"]',
                'button[title="New"]',
                '[data-testid="new-button"]',
                'button[class*="new"]',
                'button[class*="create"]',
            ]

            new_clicked = False
            for selector in new_selectors:
                try:
                    new_button = await page.query_selector(selector)
                    if new_button and await new_button.is_visible():
                        await new_button.click()
                        logger.info(f"✅ Clicked New button: {selector}")
                        new_clicked = True
                        break
                except Exception as e:
                    logger.debug(f"🔍 New button selector '{selector}' failed: {e}")
                    continue

            if not new_clicked:
                logger.error("❌ Could not find New/Create button")
                return False

            await asyncio.sleep(2)

            # Click "Excel workbook"
            excel_selectors = [
                'span.ms-ContextualMenu-itemText.label-151:has-text("Excel workbook")',
                'span.ms-ContextualMenu-itemText:has-text("Excel workbook")',
                'span:has-text("Excel workbook")',
                'text="Excel workbook"',
                '[title="Excel workbook"]',
                '.ms-Button:has-text("Excel")',
                '[aria-label="Excel workbook"]',
                'button[title*="Excel"]',
                'button[class*="excel"]',
                'button:has-text("Excel")',
                '[data-testid="excel-workbook"]',
            ]

            excel_clicked = False
            for selector in excel_selectors:
                try:
                    excel_button = await page.query_selector(selector)
                    if excel_button and await excel_button.is_visible():
                        await excel_button.click()
                        logger.info(f"✅ Clicked Excel workbook: {selector}")
                        excel_clicked = True
                        break
                except Exception as e:
                    logger.debug(f"🔍 Excel button selector '{selector}' failed: {e}")
                    continue

            if not excel_clicked:
                logger.error("❌ Could not find Excel workbook option")
                return False

            # Wait for Excel to load
            await asyncio.sleep(5)

            # Check if a new tab opened (Excel files open in new tabs)
            context = page.context
            current_pages = len(context.pages)

            if current_pages > 1:  # New tab opened
                # Switch to the new tab (last one opened)
                new_page = context.pages[-1]
                logger.info("🔄 Excel opened in new tab, switching to it...")

                try:
                    # Wait for the new tab to load
                    await new_page.wait_for_load_state("load", timeout=15000)
                    logger.info("✅ Switched to Excel tab")
                except Exception as e:
                    logger.info(f"⚠️ New tab load timeout, but continuing: {e}")

                return new_page
            else:
                logger.info("✅ File created successfully (same tab)")
                return page

        except Exception as e:
            logger.error(f"❌ Error creating Solution Model: {e}")
            return False

    @staticmethod
    async def rename_excel_file(page, new_name: str) -> bool:
        """
        Robustly rename the workbook using Excel Online UI.

        Args:
            page: Playwright page instance
            new_name: New name for the file

        Returns:
            True if rename succeeded, False otherwise
        """
        try:
            logger.info(f"✏️ Renaming Excel file to: {new_name}")
            await asyncio.sleep(2)

            # Try to open rename UI
            opened = False
            filename_button_selectors = [
                'button[title*="rename" i]',
                'button[aria-label*="name" i]',
                '[data-automationid="DocumentTitleField"]',
                '[aria-label*="File name" i]',
                ".od-DocumentTitle",
                'button[role="button"]:has(span[title$=".xlsx"])',
            ]
            for s in filename_button_selectors:
                try:
                    el = await page.query_selector(s)
                    if el:
                        await el.click()
                        await asyncio.sleep(1)
                        opened = True
                        logger.info(f"✅ Opened rename popover via: {s}")
                        break
                except Exception as e:
                    logger.debug(f"Rename trigger selector failed {s}: {e}")

            if not opened:
                # Keyboard fallback
                try:
                    await page.keyboard.press(get_save_as_key())
                    await asyncio.sleep(1)
                    opened = True
                    logger.info("✅ Opened rename via keyboard fallback")
                except Exception as e:
                    logger.debug(f"Keyboard fallback failed: {e}")

            # Handle dialog/input across page and frames
            if opened:
                contexts = [page] + list(page.frames)
                for ctx in contexts:
                    try:
                        # Prefer elements within a dialog/callout
                        dialog = await ctx.query_selector(
                            '[role="dialog"], .ms-Dialog, .ms-Callout, .ms-Callout-main, .ms-Dialog-main'
                        )
                        search_root = dialog or ctx

                        # Find filename input
                        input_sel_candidates = [
                            'input[type="text"]',
                            'input[aria-label*="File name" i]',
                            'input[role="textbox"]',
                        ]
                        name_input = None
                        for inp_sel in input_sel_candidates:
                            try:
                                elem = await search_root.query_selector(inp_sel)
                                if elem and await elem.is_visible():
                                    name_input = elem
                                    break
                            except Exception:
                                continue

                        if name_input:
                            # Compose desired filename including extension
                            file_name = (
                                new_name
                                if new_name.lower().endswith(".xlsx")
                                else f"{new_name}.xlsx"
                            )
                            await name_input.click()
                            await name_input.fill(file_name)
                            await asyncio.sleep(0.2)

                            # Tick "Replace existing file" if present
                            try:
                                checkbox = await search_root.query_selector(
                                    'input[type="checkbox"]'
                                )
                                if checkbox and not await checkbox.is_checked():
                                    await checkbox.click()
                                    await asyncio.sleep(0.1)
                            except Exception as e:
                                logger.debug(f"Checkbox handling skipped: {e}")

                            # Click Save or press Enter
                            saved = False
                            for save_sel in [
                                'button:text("Save")',
                                'button[aria-label*="Save" i]',
                                "text=Save",
                            ]:
                                try:
                                    if save_sel.startswith("text="):
                                        await search_root.click(save_sel)
                                    else:
                                        btn = await search_root.query_selector(save_sel)
                                        if btn:
                                            await btn.click()
                                    saved = True
                                    break
                                except Exception:
                                    continue
                            if not saved:
                                await ctx.keyboard.press("Enter")

                            logger.info("✅ Rename submitted via Save As dialog")
                            return True
                    except Exception as e:
                        logger.debug(f"Dialog handling error in context: {e}")

            logger.warning("⚠️ Rename UI not found; skipping")
            return False
        except Exception as e:
            logger.error(f"❌ Error renaming file: {e}")
            return False
