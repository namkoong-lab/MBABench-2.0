"""
File organizer for the Excel Agent Engine.

Handles downloading Excel files and organizing log files into date-based folder structure.
"""

import asyncio
import logging
import re
import shutil
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional

from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from .excel_operations import ExcelOperations

logger = logging.getLogger(__name__)


class TaskStatus(str, Enum):
    """Fine-grained task completion status."""

    # --- Agent statuses (ran inside the AI panel) ---
    SUCCESS = "success"
    TIMEOUT = "timeout"
    PROMPT_FAILED = "prompt_failed"
    # --- Pipeline statuses (infrastructure / pre-agent / post-agent) ---
    DOWNLOAD_FAILED = "download_failed"
    FILE_CORRUPTED = "file_corrupted"
    NAV_FAILED = "nav_failed"
    EXCEL_FAILED = "excel_failed"
    PANEL_FAILED = "panel_failed"
    # The vendor answered with a capacity / outage notice instead of working
    # (Anthropic "high demand", 2026-09-11). Infra: unrecorded, retried.
    PROVIDER_UNAVAILABLE = "provider_unavailable"
    UNKNOWN = "unknown"


AGENT_STATUSES = {
    TaskStatus.SUCCESS,
    TaskStatus.TIMEOUT,
    TaskStatus.PROMPT_FAILED,
}
PIPELINE_STATUSES = {
    TaskStatus.DOWNLOAD_FAILED,
    TaskStatus.FILE_CORRUPTED,
    TaskStatus.NAV_FAILED,
    TaskStatus.EXCEL_FAILED,
    TaskStatus.PANEL_FAILED,
    TaskStatus.PROVIDER_UNAVAILABLE,
    TaskStatus.UNKNOWN,
}


@dataclass
class ValidationResult:
    """Result of validating a downloaded Excel file."""

    is_valid: bool
    status: TaskStatus
    message: str
    file_path: Optional[Path] = None
    sheet_names: list = field(default_factory=list)


_AGENT_FOLDER_LABELS = {
    "chatgpt_excel_agent": "chatgptExcel",
    "claude_excel_agent": "claudeExcel",
    "tabai": "tabaiExcel",
}


def get_run_folder_label(agent_type: str) -> str:
    """Return the folder label for a given agent type."""
    return _AGENT_FOLDER_LABELS.get(agent_type, agent_type)


class FileOrganizer:
    """Organizes task files into date-based folder structure."""

    @staticmethod
    def create_run_folders(run_dir: Path) -> dict:
        """Flat folder set inside one attempt's staging directory.

        Used by the infra runner, which gives every attempt its own
        directory — no date segment, so a run crossing midnight still
        writes every artifact to the same place the runner reads.
        """
        run_dir = Path(run_dir)
        folders = {
            "root": run_dir,
            "solutions": run_dir / "solutions",
            "general_logs": run_dir / "general_logs",
            "json_logs": run_dir / "json_logs",
        }
        for folder in folders.values():
            folder.mkdir(parents=True, exist_ok=True)
        return folders

    @staticmethod
    def create_date_folders(
        base_dir: Path, date: str, agent_type: str = "claude_excel_agent"
    ) -> dict:
        """Create date-based folder structure.

        Folder name: {YYYYMMDD}_{agentLabel}/
        """
        folder_label = get_run_folder_label(agent_type)
        date_dir = base_dir / f"{date}_{folder_label}"
        folders = {
            "root": date_dir,
            "solutions": date_dir / "solutions",
            "general_logs": date_dir / "general_logs",
            "json_logs": date_dir / "json_logs",
        }
        for folder in folders.values():
            folder.mkdir(parents=True, exist_ok=True)
        return folders

    @staticmethod
    async def download_excel_file(
        page,
        output_dir: Path,
        task_name: str,
        agent_name: str = "tabai",
        task_source: str = "",
    ) -> Optional[Path]:
        """Download Excel file using File > Create a Copy > Download a Copy."""
        dl_start_time = datetime.now()
        # Reload page to get clean Excel state (Claude panel may have occluded ribbon)
        try:
            logger.info("🔄 Reloading page for clean download state...")
            await page.reload(wait_until="load", timeout=30000)
            await asyncio.sleep(5)  # Let Excel Online fully render
            # Dismiss any "Restore" dialogs or overlays
            try:
                await page.keyboard.press("Escape")
                await asyncio.sleep(1)
            except Exception:
                pass
            logger.info("✅ Page reloaded for download")
        except Exception as e:
            logger.warning(f"⚠️ Page reload failed: {e} — proceeding with current state")

        # Wait for the ribbon to be back before spending the retry budget.
        # 2026-09-11 (Telecom, ChatGPT lane): Page.reload timed out at 30s on
        # a heavy workbook, the five 6-second File-tab looks that followed all
        # missed, and 12 minutes of finished model work was discarded as
        # download_failed. Excel Online can take well over a minute to render
        # after a reload; a finished workbook is worth waiting for.
        ready_deadline = asyncio.get_event_loop().time() + 120
        ribbon_ready = False
        while asyncio.get_event_loop().time() < ready_deadline:
            for frame in page.frames:
                try:
                    el = await frame.query_selector('text="File"')
                    if el and await el.is_visible():
                        ribbon_ready = True
                        break
                except Exception:
                    continue
            if ribbon_ready:
                break
            await asyncio.sleep(3)
        if not ribbon_ready:
            logger.warning("⚠️ Ribbon (File tab) not visible after 120s — trying anyway")

        max_retries = 5
        download_complete = False
        download = None

        for attempt in range(max_retries):
            if attempt > 0:
                logger.info(
                    f"🔄 Retrying download sequence (attempt {attempt + 1}/{max_retries})..."
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
            else:
                await asyncio.sleep(2)

            try:
                # Step 1: Click File tab
                logger.info("🔍 Step 1: Looking for File tab...")
                if not await ExcelOperations._try_click_element(
                    page, "text=File", timeout=12000
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
                # Wait longer for submenu to appear (submenus can be slow)
                await asyncio.sleep(4.0)

                # Verify Step 2 worked by checking if "Download a Copy" appears
                verification_passed = False
                verification_selectors = [
                    "text=Download a Copy",
                    'text="Download a Copy"',
                    "text=download a copy",  # Lowercase
                ]
                for verify_frame in page.frames:
                    for selector in verification_selectors:
                        try:
                            element = await verify_frame.wait_for_selector(
                                selector, timeout=3000, state="visible"
                            )
                            verification_passed = True
                            logger.debug(
                                f"✅ Verified 'Create a Copy' - found submenu with selector: {selector}"
                            )
                            break
                        except Exception:
                            continue
                    if verification_passed:
                        break
                if not verification_passed:
                    logger.warning(
                        f"⚠️ Create a Copy click verification failed (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                await asyncio.sleep(2.0)  # Additional wait before clicking download

                # Step 3: Click "Download a Copy" and wait for download
                logger.info("🔍 Step 3: Looking for 'Download a Copy' button...")

                # Find element FIRST (quick check, outside download listener to avoid timeout)
                download_element = None
                download_selectors = [
                    "text=Download a Copy",
                    'text="Download a Copy"',
                    "text=download a copy",  # Lowercase
                ]

                # Quick check for element (non-blocking)
                for frame in page.frames:
                    for selector in download_selectors:
                        try:
                            element = await frame.query_selector(selector)
                            if element and await element.is_visible():
                                download_element = element
                                break
                        except Exception:
                            continue
                    if download_element:
                        break

                if not download_element:
                    # Fallback: wait for element with reasonable timeout
                    for frame in page.frames:
                        for selector in download_selectors:
                            try:
                                download_element = await frame.wait_for_selector(
                                    selector, timeout=5000, state="visible"
                                )
                                break
                            except Exception:
                                continue
                        if download_element:
                            break

                if not download_element:
                    logger.warning(
                        f"⚠️ Download a Copy button not found (attempt {attempt + 1}/{max_retries})"
                    )
                    continue

                # NOW set up download listener and click (element is ready, click should be fast)
                try:
                    async with page.expect_download(timeout=15000) as download_info:
                        # Click with longer timeout to ensure stability
                        await download_element.click(timeout=5000)
                        logger.info("✅ Clicked 'Download a Copy'")

                    download = await download_info.value
                    download_complete = True
                    break
                except (PlaywrightTimeoutError, asyncio.TimeoutError):
                    logger.warning(
                        f"⚠️ Download timeout (attempt {attempt + 1}/{max_retries})"
                    )
                    continue
                except Exception as e:
                    logger.warning(
                        f"⚠️ Download error: {e} (attempt {attempt + 1}/{max_retries})"
                    )
                    continue

            except Exception as e:
                logger.warning(
                    f"⚠️ Download sequence error: {e} (attempt {attempt + 1}/{max_retries})"
                )
                continue

        if not download_complete or not download:
            dl_elapsed = (datetime.now() - dl_start_time).total_seconds()
            logger.error(
                f"❌ Could not complete download sequence after retries "
                f"({dl_elapsed:.0f}s elapsed)"
            )
            return None

        # Save the downloaded file
        try:
            safe_task_name = task_name.replace("/", "-").replace(" ", "_")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if download.suggested_filename and download.suggested_filename.endswith(
                (".xlsx", ".xls")
            ):
                base_name = Path(download.suggested_filename).stem
                extension = Path(download.suggested_filename).suffix
                # Use just timestamp + suggested filename (which already contains
                # the task name). Adding safe_task_name again would double it and
                # exceed Windows MAX_PATH (260 chars) for long task names.
                unique_filename = f"{timestamp}_{base_name}{extension}"
            else:
                # Build the filename. Drop the task_source segment if empty so
                # custom-source projects don't get an awkward double underscore.
                source_segment = f"_{task_source}" if task_source else ""
                unique_filename = f"{timestamp}_{safe_task_name}_Solution{source_segment}_{agent_name}_Model.xlsx"

            output_path = output_dir / unique_filename
            base_stem = output_path.stem
            counter = 1
            while output_path.exists():
                output_path = output_dir / f"{base_stem}_{counter}{output_path.suffix}"
                counter += 1

            output_path.parent.mkdir(parents=True, exist_ok=True)
            await download.save_as(str(output_path))
            # save_as over a CDP-attached browser can write 0 bytes while
            # Chrome saves the same stream natively to its own download
            # directory (2026-08-27: a second Playwright client attaching
            # mid-run reset the browser-wide download behavior; both capture
            # attempts produced empty files while ~/Downloads held the intact
            # workbook, and the 28-min attempt was discarded as infra). Never
            # trust save_as blindly: verify the byte count and recover the
            # native artifact when it exists.
            if output_path.stat().st_size == 0:
                rescued = FileOrganizer._rescue_native_download(
                    download.suggested_filename, dl_start_time
                )
                if rescued is not None:
                    shutil.copy2(rescued, output_path)
                    logger.warning(
                        f"⚠️ save_as wrote 0 bytes; recovered the workbook "
                        f"from the browser's native download dir: {rescued}"
                    )
            dl_elapsed = (datetime.now() - dl_start_time).total_seconds()
            file_size_mb = output_path.stat().st_size / (1024 * 1024)
            logger.info(
                f"✅ Downloaded: {output_path.name} "
                f"({file_size_mb:.1f} MB, {dl_elapsed:.0f}s)"
            )
            return output_path

        except Exception as e:
            logger.error(f"❌ Error saving downloaded file: {e}")
            logger.debug(traceback.format_exc())
            return None

    @staticmethod
    def _rescue_native_download(
        suggested_filename: Optional[str], started: datetime
    ) -> Optional[Path]:
        """Newest matching file in the browser's native download directory.

        When Playwright's managed download artifact is empty (see caller),
        Chrome has usually saved the real stream itself under the suggested
        filename — deduplicated as "name (N).ext" — in its default download
        directory (the engine launches Chrome without overriding it, so
        ~/Downloads). Return the newest non-empty candidate written since
        `started`, or None; on None the caller's existing validation-failure
        path stands.
        """
        if not suggested_filename:
            return None
        native_dir = Path.home() / "Downloads"
        if not native_dir.is_dir():
            return None
        stem = Path(suggested_filename).stem
        suffix = Path(suggested_filename).suffix
        pattern = re.compile(rf"^{re.escape(stem)}( \(\d+\))?{re.escape(suffix)}$")
        threshold = started.timestamp() - 5
        try:
            candidates = [
                p
                for p in native_dir.iterdir()
                if p.is_file()
                and pattern.match(p.name)
                and p.stat().st_mtime >= threshold
                and p.stat().st_size > 0
            ]
        except OSError:
            return None
        if not candidates:
            return None
        return max(candidates, key=lambda p: p.stat().st_mtime)

    @staticmethod
    def copy_log_file(source_path: Path, dest_dir: Path) -> Optional[Path]:
        """Copy log file to destination directory."""
        try:
            if not source_path.exists():
                return None
            dest_path = dest_dir / source_path.name
            shutil.copy2(source_path, dest_path)
            return dest_path
        except Exception as e:
            logger.error(f"❌ Copy error: {e}")
            return None

    @staticmethod
    def validate_excel_file(file_path: Optional[Path]) -> ValidationResult:
        """
        Validate a downloaded Excel file for correctness.

        Checks:
        1. File exists and path is not None
        2. File size > 0
        3. openpyxl can open it (not corrupted)

        Returns:
            ValidationResult with is_valid, status, message, file_path, sheet_names
        """
        # Check 1: File exists
        if file_path is None or not Path(file_path).exists():
            return ValidationResult(
                is_valid=False,
                status=TaskStatus.DOWNLOAD_FAILED,
                message=f"File not found: {file_path}",
                file_path=file_path,
            )

        file_path = Path(file_path)

        # Check 2: File size > 0
        if file_path.stat().st_size == 0:
            return ValidationResult(
                is_valid=False,
                status=TaskStatus.FILE_CORRUPTED,
                message=f"File is empty (0 bytes): {file_path.name}",
                file_path=file_path,
            )

        # Check 3: openpyxl can open it
        try:
            import openpyxl
        except ImportError:
            logger.warning(
                "openpyxl not installed — skipping Excel validation. "
                "Install with: pip install openpyxl"
            )
            return ValidationResult(
                is_valid=True,
                status=TaskStatus.SUCCESS,
                message="Validation skipped (openpyxl not available)",
                file_path=file_path,
            )

        try:
            wb = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        except Exception as e:
            return ValidationResult(
                is_valid=False,
                status=TaskStatus.FILE_CORRUPTED,
                message=f"Cannot open Excel file: {e}",
                file_path=file_path,
            )

        logger.info(
            f"✅ Excel validation passed: {file_path.name} (sheets: {sheet_names})"
        )
        return ValidationResult(
            is_valid=True,
            status=TaskStatus.SUCCESS,
            message="Validation passed",
            file_path=file_path,
            sheet_names=sheet_names,
        )

    @staticmethod
    async def organize_task_files(
        page,
        task_name: str,
        log_file_path: Optional[str],
        json_file_path: Optional[str],
        base_dir: Optional[Path] = None,
        agent_name: str = "claude_excel_agent",
        task_source: str = "",
        run_dir: Optional[Path] = None,
    ) -> dict:
        """Organize all task files.

        With `run_dir` (the infra runner's per-attempt staging dir), files
        land flat inside it; otherwise a date-based folder is created under
        `base_dir` (standalone runs).
        """
        try:
            if run_dir is not None:
                folders = FileOrganizer.create_run_folders(run_dir)
            else:
                base_dir = base_dir or Path.cwd()
                today = datetime.now().strftime("%Y%m%d")
                folders = FileOrganizer.create_date_folders(
                    base_dir, today, agent_name
                )

            excel_path = await FileOrganizer.download_excel_file(
                page, folders["solutions"], task_name, agent_name, task_source
            )

            # Validate the downloaded Excel file
            validation = FileOrganizer.validate_excel_file(excel_path)
            if not validation.is_valid:
                logger.warning(
                    f"⚠️ Excel validation failed: {validation.status.value} — {validation.message}"
                )
            else:
                logger.info(f"✅ Excel validation: {validation.message}")

            results = {
                "excel": excel_path,
                "validation": validation,
                "log": (
                    FileOrganizer.copy_log_file(
                        Path(log_file_path), folders["general_logs"]
                    )
                    if log_file_path
                    else None
                ),
                # JSON completion logs are written directly to json_logs/
                # by CompletionLogger — no copy needed
                "json": json_file_path,
            }
            return results
        except Exception as e:
            logger.error(f"❌ File organization error: {type(e).__name__}: {e!r}")
            logger.debug(traceback.format_exc())
            return {
                "excel": None,
                "validation": ValidationResult(
                    is_valid=False,
                    status=TaskStatus.UNKNOWN,
                    message=f"File organization error: {type(e).__name__}: {e!r}",
                ),
                "log": None,
                "json": None,
            }
