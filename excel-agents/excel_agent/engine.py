#!/usr/bin/env python3
"""Excel Agent Engine — one attempt of one task in Excel Online.

Drives a real Chrome (over CDP) through OneDrive: navigate to the task
folder, open the template workbook (or create a blank one for template-less
tasks), "Create a Copy" under a standardized name, open the AI add-in panel
(Claude or ChatGPT), run the configured prompts, then download and validate
the resulting workbook.

Normally spawned by `python -m infra.run` with a generated temp config and
`--no-hold`; can also be run standalone against a hand-written config.

Exit-code contract (consumed by infra/run.py — keep in sync):
    0   success — prompts completed, workbook downloaded and validated
    1   agent failure — the AI ran but failed (prompt_failed, timeout).
        Recorded as an attempt with agent_failed=true.
    2   config error — bad/missing config; nothing was attempted.
    3   infra (pipeline) failure — nav/Excel/panel/download problem before
        or after the agent's own work. NOT recorded; the runner retries
        in place.

The engine ALWAYS opens the task's template workbook when one is
configured — including on retries. (The predecessor gated this on
attempt_number == 0, so any retried attempt silently ran on a blank
workbook and could still be recorded as a success.)

The solution workbook's path is written into the completion JSON
(solution_file) and echoed as an `ENGINE_SOLUTION_FILE=` line, so the
runner reads it back exactly instead of guessing from directory globs.
"""

import argparse
import asyncio
import logging
import signal
import sys
import time
from datetime import datetime
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import yaml
from playwright.async_api import async_playwright

from excel_agent.core import (
    AGENT_STATUSES,
    BrowserManager,
    ChatGPTCore,
    ClaudeCore,
    CompletionLogger,
    ExcelOperations,
    FileManager,
    FileOrganizer,
    Navigation,
    TaskStatus,
    ValidationResult,
    setup_logging,
)
from excel_agent.core.logging_setup import (
    configure_safe_stdout,
)

logger = logging.getLogger(__name__)

EXIT_SUCCESS = 0
EXIT_AGENT_FAILURE = 1
EXIT_CONFIG_ERROR = 2
EXIT_INFRA_FAILURE = 3

AGENT_CORES = {
    "claude_excel_agent": ClaudeCore,
    "chatgpt_excel_agent": ChatGPTCore,
}

# Global shutdown event for graceful termination
shutdown_event = asyncio.Event()


def _handle_signal(signum, frame):
    try:
        shutdown_event.set()
    except Exception:
        pass
    # Restore default handler so a second Ctrl+C force quits
    signal.signal(signal.SIGINT, signal.SIG_DFL)


async def run_automation(config: dict) -> str:
    """Single attempt. Returns "success" | "agent_failure" | "pipeline_failure"."""
    agent_type = config["agent_type"]  # validated in main()

    max_sec_per_task = config.get(agent_type, {}).get("max_sec_per_task", 0) or 0
    if max_sec_per_task > 0:
        logger.info(
            f"⏱️  Task timeout: {max_sec_per_task}s ({max_sec_per_task // 60}m)"
        )
    else:
        logger.info("⏱️  No task timeout configured")

    async def _timeout_guard():
        if max_sec_per_task > 0:
            await asyncio.sleep(max_sec_per_task)
            logger.error(f"⏱️  TIMEOUT: task exceeded {max_sec_per_task}s")
            shutdown_event.set()

    guard_task = None

    try:
        browser_mgr = BrowserManager(config)
        file_mgr = FileManager()

        task_name = config["task_name"]
        task_source = config.get("task_source", "")
        direct_url = config.get("direct_url")
        onedrive_path = config.get("onedrive_path")
        template_file = config.get("template_file")
        upload_files = config.get("upload_files") or []
        solution_name = config.get("solution_name")
        local_files_base = config.get("local_files_base")
        attempt_number = int(config.get("attempt_number", 0) or 0)
        current_log_file_path = config.get("_log_file_path")

        agent_name = agent_type

        files_to_upload = file_mgr.resolve_upload_files(
            [str(f) for f in upload_files], local_files_base
        )
        if len(files_to_upload) != len(upload_files):
            # A benchmark attempt without its inputs is not a real attempt.
            logger.error(
                f"❌ Only {len(files_to_upload)}/{len(upload_files)} upload "
                f"files resolved — aborting as infra failure"
            )
            return "pipeline_failure"

        if template_file and str(template_file).lower() != "blank":
            workbook_filename = Path(str(template_file)).name
            logger.info(f"📊 Will open template workbook: {workbook_filename}")
        else:
            workbook_filename = None
            logger.info("📝 No template workbook — will create a blank one")

        # Per-attempt output directory. The runner passes run_dir (its
        # staging dir); standalone runs get a dated folder in CWD.
        run_dir_cfg = config.get("run_dir")
        if run_dir_cfg:
            folders = FileOrganizer.create_run_folders(Path(run_dir_cfg))
        else:
            today = datetime.now().strftime("%Y%m%d")
            folders = FileOrganizer.create_date_folders(
                Path.cwd(), today, agent_type
            )
        run_dir = folders["root"]

        async with async_playwright() as playwright:
            browser, context = await browser_mgr.launch_browser(playwright)

            # CDP mode: always drive a fresh page, never touch existing tabs.
            page = await context.new_page()
            logger.info(
                f"✅ Created fresh page (CDP, {len(context.pages)} total tabs)"
            )

            final_task_status = None  # None = no failure yet
            task_success = False
            solution_path = None
            excel_page = None
            task_page = None

            async def _close_task_pages(skip_save=False):
                nonlocal excel_page, task_page
                CLOSE_TIMEOUT = 30

                async def _do_close():
                    if not skip_save:
                        try:
                            if excel_page and not excel_page.is_closed():
                                from excel_agent.core.browser_manager import (
                                    get_modifier_key,
                                )

                                logger.info("💾 Saving Excel file before closing...")
                                await excel_page.keyboard.press(
                                    f"{get_modifier_key()}+S"
                                )
                                await asyncio.sleep(7)
                                logger.info("✅ Save complete, closing pages...")
                        except Exception as e:
                            logger.debug(f"Save attempt: {e}")
                    else:
                        logger.info("⏩ Skipping save (task failed), closing pages...")
                    for pg in (excel_page, task_page):
                        try:
                            if pg and pg is not page and not pg.is_closed():
                                await pg.close()
                        except Exception:
                            pass

                try:
                    await asyncio.wait_for(_do_close(), timeout=CLOSE_TIMEOUT)
                except TimeoutError:
                    logger.warning(
                        f"⚠️ _close_task_pages timed out after {CLOSE_TIMEOUT}s"
                    )

            retry_suffix = f"_retry{attempt_number}" if attempt_number > 0 else ""
            attempt_task_id = f"{task_name}{retry_suffix}"

            completion_logger = CompletionLogger(
                log_dir=str(folders["json_logs"]),
                task_identifier=attempt_task_id,
                agent_name=agent_name,
                prompt_version=config.get("prompt_version"),
                task_source=task_source,
                max_sec_per_task=max_sec_per_task,
            )

            guard_task = asyncio.create_task(_timeout_guard())

            try:
                if shutdown_event.is_set():
                    logger.warning("⏱️ Shutdown event set — skipping")
                    final_task_status = TaskStatus.TIMEOUT

                # ======================================================
                # PIPELINE PHASE: navigate → workbook → panel
                # ======================================================

                if final_task_status is None:
                    logger.info("🌐 Navigating to OneDrive...")
                    try:
                        await page.goto(
                            "https://onedrive.live.com",
                            wait_until="load",
                            timeout=60000,
                        )
                        await asyncio.sleep(2)
                        logger.info(f"✅ OneDrive loaded: {page.url}")
                    except Exception as e:
                        logger.error(f"❌ Failed to load OneDrive: {e}")
                        final_task_status = TaskStatus.NAV_FAILED

                if final_task_status is None:
                    logger.info(
                        f"📁 Navigating to task folder (source: {task_source or 'custom'})"
                    )
                    nav = Navigation()
                    task_page = None
                    NAV_TIMEOUT = 180

                    async def _nav_with_retries():
                        nonlocal task_page
                        for nav_attempt in range(3):
                            if nav_attempt > 0:
                                logger.warning(
                                    f"⚠️ Navigation attempt {nav_attempt + 1}/3..."
                                )
                                try:
                                    await page.goto(
                                        "https://onedrive.live.com",
                                        wait_until="load",
                                        timeout=60000,
                                    )
                                    await asyncio.sleep(2)
                                except Exception as e:
                                    logger.error(f"❌ Could not reload OneDrive: {e}")
                                    continue
                            task_page = await nav.navigate_to_task_folder(
                                page,
                                task_name,
                                task_source,
                                base_path=config.get("file_path"),
                                direct_url=direct_url,
                                onedrive_path=onedrive_path,
                            )
                            if task_page:
                                break

                    try:
                        await asyncio.wait_for(_nav_with_retries(), timeout=NAV_TIMEOUT)
                    except TimeoutError:
                        logger.error(f"❌ Navigation timed out after {NAV_TIMEOUT}s")
                        task_page = None

                    if not task_page:
                        logger.error("❌ Failed to reach the task folder")
                        final_task_status = TaskStatus.NAV_FAILED

                # --- Open the template workbook (EVERY attempt) or create one
                if final_task_status is None:
                    excel_ops = ExcelOperations()
                    if solution_name:
                        base_name = f"{solution_name}_{agent_name}"
                    else:
                        source_segment = f"_{task_source}" if task_source else ""
                        base_name = (
                            f"{task_name}_Solution{source_segment}_{agent_name}_Model"
                        )

                    if workbook_filename:
                        logger.info("=" * 60)
                        logger.info("📊 OPENING TEMPLATE WORKBOOK AND CREATING COPY")
                        logger.info("=" * 60)
                        logger.info(f"   Opening file: {workbook_filename}")
                        excel_page = await nav.open_specific_file(
                            task_page, workbook_filename
                        )
                    else:
                        logger.info("=" * 60)
                        logger.info("📊 CREATING NEW WORKBOOK")
                        logger.info("=" * 60)
                        excel_page = await excel_ops.create_solution_model(task_page)

                    if not excel_page:
                        logger.error("❌ Failed to open/create Excel file")
                        final_task_status = TaskStatus.EXCEL_FAILED

                # --- Standardized copy with incremental naming
                if final_task_status is None:
                    RENAME_TIMEOUT_S = 240
                    logger.info(
                        f"✏️ Creating copy: N_{base_name}.xlsx "
                        f"(hard timeout: {RENAME_TIMEOUT_S}s)"
                    )
                    copy_success = False
                    next_number = 0

                    async def _rename_phase():
                        nonlocal copy_success, next_number
                        for attempt in range(4):
                            if attempt > 0:
                                try:
                                    await excel_page.keyboard.press("Escape")
                                    await asyncio.sleep(1.0)
                                    await excel_page.keyboard.press("Escape")
                                    await asyncio.sleep(1.0)
                                except Exception:
                                    pass
                                wait_time = (2**attempt) * 2
                                logger.info(f"   Waiting {wait_time}s before retry...")
                                await asyncio.sleep(wait_time)
                            else:
                                await asyncio.sleep(3)
                            copy_success, next_number = await excel_ops.rename_workbook(
                                excel_page,
                                base_name,
                                retry=(attempt > 0),
                                start_number=next_number,
                            )
                            if copy_success:
                                break

                    try:
                        await asyncio.wait_for(_rename_phase(), timeout=RENAME_TIMEOUT_S)
                    except TimeoutError:
                        logger.error(
                            f"⏱️ RENAME TIMEOUT: exceeded {RENAME_TIMEOUT_S}s"
                        )
                        copy_success = False

                    if not copy_success:
                        logger.error("❌ Could not create copy after 4 attempts")
                        logger.error("   Aborting to avoid modifying the template")
                        final_task_status = TaskStatus.EXCEL_FAILED

                # --- Open the add-in panel
                if final_task_status is None:
                    await asyncio.sleep(3)
                    await excel_page.keyboard.press("Escape")
                    await asyncio.sleep(1.0)

                    ai_agent = AGENT_CORES[agent_type](
                        excel_page, config, shutdown_event, completion_logger
                    )
                    logger.info(f"🤖 Using {ai_agent.get_addon_name()}")

                    # 480 so three open attempts fit at the 60s find_and_click
                    # budget below (2026-08-27: the Add-ins ribbon took ~40s+
                    # to materialize on a fresh copy under two-lane load; the
                    # old 15s budget lost that race seven straight times).
                    PANEL_OPEN_TIMEOUT_S = 480
                    CLOSE_PANEL_TIMEOUT = 15
                    agent_opened = False

                    async def _panel_open_phase():
                        nonlocal agent_opened
                        for agent_attempt in range(3):
                            if agent_attempt > 0:
                                logger.info("🔄 Closing stale panel before retry...")
                                try:
                                    await asyncio.wait_for(
                                        ai_agent._close_panel(),
                                        timeout=CLOSE_PANEL_TIMEOUT,
                                    )
                                except TimeoutError:
                                    logger.warning("⚠️ _close_panel() timed out")
                                logger.warning(
                                    f"⚠️ {ai_agent.get_addon_name()} open "
                                    f"attempt {agent_attempt + 1}/3..."
                                )
                                await asyncio.sleep(3)
                                if agent_attempt >= 1:
                                    try:
                                        logger.info("🔄 Refreshing Excel page...")
                                        await excel_page.reload(
                                            wait_until="load", timeout=30000
                                        )
                                        await asyncio.sleep(8)
                                    except Exception as e:
                                        logger.error(f"❌ Could not reload page: {e}")

                            if await ai_agent.find_and_click(max_seconds=60):
                                if await ai_agent.verify_session_health():
                                    agent_opened = True
                                    break
                                logger.warning(
                                    "⚠️ Panel opened but health check failed"
                                )
                                try:
                                    await asyncio.wait_for(
                                        ai_agent._close_panel(),
                                        timeout=CLOSE_PANEL_TIMEOUT,
                                    )
                                except TimeoutError:
                                    logger.warning("⚠️ _close_panel() timed out")

                    try:
                        await asyncio.wait_for(
                            _panel_open_phase(), timeout=PANEL_OPEN_TIMEOUT_S
                        )
                    except TimeoutError:
                        logger.error(
                            f"⏱️ PANEL TIMEOUT: exceeded {PANEL_OPEN_TIMEOUT_S}s"
                        )
                        agent_opened = False

                    # Setup (model/effort pin + verification) runs outside
                    # the panel timeout. A pin that cannot be verified is a
                    # setup failure — infra, unrecorded — never a run on an
                    # unverified model.
                    if agent_opened:
                        if not await ai_agent.handle_initial_setup():
                            logger.error(
                                f"❌ {ai_agent.get_addon_name()} setup failed "
                                f"(model/effort pin unverified)"
                            )
                            agent_opened = False

                    if not agent_opened:
                        logger.error(
                            f"❌ Could not open {ai_agent.get_addon_name()} panel"
                        )
                        final_task_status = TaskStatus.PANEL_FAILED

                # ======================================================
                # AGENT PHASE
                # ======================================================
                if final_task_status is None:
                    completion_logger.start_task(
                        attempt_task_id, attempt_number=attempt_number + 1
                    )

                    logger.info("🚀 Starting prompt processing...")
                    prompt_success = False
                    try:
                        if not await ai_agent.process_all_prompts(
                            files_to_upload=files_to_upload
                        ):
                            logger.error("❌ Prompt processing failed")
                            if shutdown_event.is_set():
                                completion_logger.end_task(
                                    task_status=TaskStatus.TIMEOUT,
                                    error_msg="Task timed out",
                                )
                                final_task_status = TaskStatus.TIMEOUT
                            elif getattr(ai_agent, "provider_unavailable", False):
                                # The vendor answered with a capacity notice
                                # (Anthropic "high demand", 2026-09-11) — the
                                # model never got to work. Infra, not the
                                # model's failure: unrecorded and retried.
                                completion_logger.end_task(
                                    task_status=TaskStatus.PROVIDER_UNAVAILABLE,
                                    error_msg="Vendor capacity/outage notice",
                                )
                                final_task_status = TaskStatus.PROVIDER_UNAVAILABLE
                                # A usage limit with a known reset time: hold
                                # here (bounded) so the runner's infra retry
                                # lands after the reset instead of churning
                                # through the queue as infra-skips.
                                retry_after = getattr(
                                    ai_agent, "provider_retry_after", None
                                )
                                if retry_after:
                                    hold = min(max(0.0, retry_after - time.time()), 3600.0)
                                else:
                                    # No parseable reset time: still hold, or
                                    # the runner's 20s retry re-hits the same
                                    # limit and burns all three tries in
                                    # minutes (2026-09-12 03:02).
                                    hold = 1800.0
                                if hold > 0:
                                    logger.warning(
                                        f"⏸️ Vendor notice — holding {hold/60:.0f} min "
                                        "before the infra retry"
                                    )
                                    await asyncio.sleep(hold)
                            else:
                                completion_logger.end_task(
                                    task_status=TaskStatus.PROMPT_FAILED,
                                    error_msg="Prompt processing failed",
                                )
                                final_task_status = TaskStatus.PROMPT_FAILED
                        else:
                            logger.info("✅ All prompts completed successfully!")
                            prompt_success = True

                        # --- Download + validate (also archival on failure)
                        DOWNLOAD_TIMEOUT = 600

                        async def _download_once():
                            return await asyncio.wait_for(
                                FileOrganizer.organize_task_files(
                                    excel_page,
                                    task_name,
                                    current_log_file_path,
                                    str(completion_logger.session_file),
                                    agent_name=agent_name,
                                    task_source=task_source,
                                    run_dir=run_dir,
                                ),
                                timeout=DOWNLOAD_TIMEOUT,
                            )

                        try:
                            logger.info("📁 Downloading + validating workbook...")
                            try:
                                file_results = await _download_once()
                            except TimeoutError:
                                file_results = {
                                    "excel": None,
                                    "validation": ValidationResult(
                                        is_valid=False,
                                        status=TaskStatus.DOWNLOAD_FAILED,
                                        message=(
                                            f"Download timed out after "
                                            f"{DOWNLOAD_TIMEOUT}s"
                                        ),
                                    ),
                                }

                            validation = file_results.get("validation")

                            if not prompt_success:
                                if validation and validation.is_valid:
                                    logger.info(
                                        "📁 Archival download OK (prompt still failed)"
                                    )
                                    solution_path = file_results.get("excel")
                            else:
                                if validation and validation.is_valid:
                                    final_task_status = TaskStatus.SUCCESS
                                    task_success = True
                                    solution_path = file_results.get("excel")
                                    completion_logger.end_task(
                                        task_status=TaskStatus.SUCCESS
                                    )
                                    logger.info("✅ Post-task validation PASSED")
                                else:
                                    if validation:
                                        logger.warning(
                                            f"⚠️ Validation FAILED: "
                                            f"{validation.status.value} — "
                                            f"{validation.message}"
                                        )
                                    # One re-download for corruption/timeouts
                                    redownload_success = False
                                    if validation and validation.status in (
                                        TaskStatus.DOWNLOAD_FAILED,
                                        TaskStatus.FILE_CORRUPTED,
                                    ):
                                        logger.info("🔄 Attempting re-download...")
                                        if (
                                            validation.file_path
                                            and Path(validation.file_path).exists()
                                        ):
                                            try:
                                                Path(validation.file_path).unlink()
                                            except Exception:
                                                pass
                                        try:
                                            file_results2 = await _download_once()
                                        except TimeoutError:
                                            file_results2 = {}
                                        validation2 = file_results2.get("validation")
                                        if validation2 and validation2.is_valid:
                                            final_task_status = TaskStatus.SUCCESS
                                            task_success = True
                                            solution_path = file_results2.get("excel")
                                            completion_logger.end_task(
                                                task_status=TaskStatus.SUCCESS
                                            )
                                            logger.info("✅ Re-download PASSED")
                                            redownload_success = True

                                    if not redownload_success:
                                        val_status = (
                                            validation.status
                                            if validation
                                            else TaskStatus.UNKNOWN
                                        )
                                        completion_logger.end_task(
                                            task_status=val_status,
                                            error_msg=(
                                                validation.message
                                                if validation
                                                else None
                                            ),
                                        )
                                        final_task_status = val_status

                        except Exception as e:
                            logger.warning(
                                f"⚠️ File organization failed: "
                                f"{type(e).__name__}: {e!r}"
                            )
                            if prompt_success:
                                final_task_status = TaskStatus.DOWNLOAD_FAILED
                                completion_logger.end_task(
                                    task_status=TaskStatus.DOWNLOAD_FAILED,
                                    error_msg=f"{type(e).__name__}: {e!r}",
                                )

                    finally:
                        if completion_logger.current_task:
                            completion_logger.end_task(
                                task_status=(final_task_status or TaskStatus.UNKNOWN),
                                error_msg="Task not completed (unexpected exit)",
                            )

            finally:
                # Record the solution path (even for archival downloads) so
                # the runner reads it back exactly — never a glob guess.
                completion_logger.set_solution_file(solution_path)
                if solution_path:
                    print(f"ENGINE_SOLUTION_FILE={solution_path}", flush=True)
                await _close_task_pages(skip_save=not task_success)
                if guard_task is not None:
                    guard_task.cancel()

            if shutdown_event.is_set() and not task_success:
                if final_task_status is None:
                    final_task_status = TaskStatus.TIMEOUT

            if task_success:
                result = "success"
                logger.info("✅ Task completed successfully!")
            elif final_task_status is not None and final_task_status in AGENT_STATUSES:
                result = "agent_failure"
                logger.error(f"❌ Agent failure: {final_task_status.value}")
            else:
                result = "pipeline_failure"
                status_val = final_task_status.value if final_task_status else "unknown"
                logger.error(f"❌ Pipeline failure: {status_val}")

            # Close our navigation page; keep the shared Chrome alive.
            try:
                if page and not page.is_closed() and len(context.pages) > 1:
                    await page.close()
            except Exception as e:
                logger.debug(f"Could not close main page: {e}")
            await browser_mgr.close_browser(context, browser)

            return result

    except Exception as e:
        logger.error(f"❌ Automation error: {type(e).__name__}: {e}")
        return "pipeline_failure"
    finally:
        if guard_task is not None:
            guard_task.cancel()


def _validate_config(config: dict) -> list[str]:
    """Config problems that make the attempt meaningless. Non-empty → exit 2."""
    errors = []
    agent_type = config.get("agent_type")
    if agent_type not in AGENT_CORES:
        errors.append(
            f"agent_type={agent_type!r} is not one of {sorted(AGENT_CORES)}"
        )
    if not config.get("task_name"):
        errors.append("task_name is missing")
    prompts = config.get("prompts")
    if not prompts or not isinstance(prompts, list):
        errors.append("prompts must be a non-empty list of prompt strings")
    if not (
        config.get("direct_url")
        or config.get("onedrive_path")
        or (config.get("file_path") and config.get("task_source") is not None)
    ):
        errors.append(
            "no navigation target: set direct_url, onedrive_path, or "
            "file_path (OneDrive base segments) + task_source"
        )
    return errors


def main():
    configure_safe_stdout()

    parser = argparse.ArgumentParser(
        description="Excel Agent Engine — one attempt of one task in Excel Online"
    )
    parser.add_argument("--config", required=True, help="Path to engine config YAML")
    parser.add_argument(
        "--no-hold",
        action="store_true",
        help="Exit immediately after completion (the runner always passes this)",
    )
    args = parser.parse_args()

    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    config_path = Path(args.config)
    if not config_path.exists():
        print(f"❌ Config file not found: {config_path}")
        sys.exit(EXIT_CONFIG_ERROR)

    with open(config_path) as f:
        config = yaml.safe_load(f) or {}

    errors = _validate_config(config)
    if errors:
        print("❌ Config validation failed:")
        for e in errors:
            print(f"   - {e}")
        sys.exit(EXIT_CONFIG_ERROR)

    global logger
    task_name = config.get("task_name", "unknown_task")
    agent_name = config.get("agent_type")
    logger, log_file_path = setup_logging(
        config, __name__, task_name=task_name, agent_name=agent_name
    )
    config["_log_file_path"] = log_file_path

    logger.info("=" * 80)
    logger.info("🚀 Excel Agent Engine Starting")
    logger.info("=" * 80)
    logger.info(f"📋 Task: {task_name}")
    logger.info(f"📋 Agent: {agent_name}")
    logger.info(f"📋 Config: {config_path}")

    try:
        result = asyncio.run(run_automation(config))
    except KeyboardInterrupt:
        logger.info("👋 Interrupted — recording nothing (infra exit)")
        sys.exit(EXIT_INFRA_FAILURE)
    except Exception as e:
        logger.error(f"❌ Engine crashed: {type(e).__name__}: {e}")
        import traceback

        logger.debug(traceback.format_exc())
        sys.exit(EXIT_INFRA_FAILURE)

    if not args.no_hold and not shutdown_event.is_set():
        logger.info("⏸️  Browser staying open for inspection (Ctrl+C to exit)...")
        try:
            while not shutdown_event.is_set():
                import time

                time.sleep(1)
        except KeyboardInterrupt:
            pass

    if result == "success":
        print("\n🎉 SUCCESS")
        sys.exit(EXIT_SUCCESS)
    elif result == "agent_failure":
        print("\n❌ AGENT FAILURE")
        sys.exit(EXIT_AGENT_FAILURE)
    else:
        print("\n❌ INFRA FAILURE")
        sys.exit(EXIT_INFRA_FAILURE)


if __name__ == "__main__":
    main()
