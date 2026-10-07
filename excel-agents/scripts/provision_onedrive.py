#!/usr/bin/env python3
"""Provision the OneDrive folder tree the excel-agents engine navigates.

For every task of the downloaded benchmark (<data_root>/tasks/task_id=<N>/,
see scripts/download_dataset.py at the repo root), ensures

    <onedrive_base_path> / <task_name> / Task /

exists on OneDrive and that the task's template workbook (found among its
starting files by the same rule the runner uses) is uploaded into the Task
folder under its original filename — which is exactly the name the engine
will open. Tasks with no workbook among their starting files still get
their folder chain (the engine creates a blank workbook there).

One-time / occasional operator tool, driven through the same automation
Chrome + Microsoft 365 session as the engine. Watch it run: OneDrive's UI
drifts, and a provisioning mistake is cheap to catch here and expensive to
catch as NAV_FAILED retries later.

Usage (from excel-agents/):
    uv run python scripts/provision_onedrive.py --dry-run     # list the plan
    uv run python scripts/provision_onedrive.py --stage       # manual path:
        # build the tree locally under onedrive_staging/, then drag it into
        # OneDrive web yourself (one folder drag; no browser automation)
    uv run python scripts/provision_onedrive.py               # automated upload
    uv run python scripts/provision_onedrive.py --verify      # check + manifest
    uv run python scripts/provision_onedrive.py --task-ids 1 2 3

--verify walks the tree read-only and writes onedrive_manifest.json
(task_id -> path, workbook, verified_at) next to this repo's root.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from excel_agent.core.browser_manager import BrowserManager
from excel_agent.core.excel_operations import ExcelOperations
from excel_agent.core.file_manager import FileManager
from excel_agent.core.navigation import Navigation, _normalize_name
from infra.configs import (
    ConfigError,
    data_root,
    describe_local_roots,
    load_configs,
    output_root,
)
from infra.run import _ns_to_dict, onedrive_task_path
from playwright.async_api import async_playwright
from task_io.sources.bundle_source import BundleTaskSource

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("provision_onedrive")

MANIFEST_PATH = _REPO_ROOT / "onedrive_manifest.json"


def _build_source(task_ids):
    """A task source for provisioning: every matching task of the downloaded
    benchmark, regardless of attempts (the attempt filter is off, so the
    agent name is a placeholder)."""
    return BundleTaskSource(
        data_root=data_root(),
        output_root=output_root(),
        agent_model_name="__provisioning__",
        prompt_version=None,
        task_ids=list(task_ids or []),
        skip_already_attempted=False,
    )


async def _visible_row_names(page) -> set[str]:
    """Normalized names of the rows currently listed in the OneDrive view."""
    names: set[str] = set()
    try:
        rows = page.locator('[role="row"]')
        count = await rows.count()
    except Exception:
        return names
    for i in range(min(count, 250)):
        row = rows.nth(i)
        try:
            cells = row.locator("a, button, [title]")
            n = await cells.count()
        except Exception:
            continue
        for j in range(min(n, 8)):
            cell = cells.nth(j)
            try:
                for value in (
                    await cell.get_attribute("title"),
                    await cell.text_content(),
                ):
                    if value and value.strip():
                        names.add(_normalize_name(value))
            except Exception:
                continue
    return names


_SCROLL_LIST_JS = """(dy) => {
  let moved = false;
  for (const el of document.querySelectorAll('*')) {
    const cs = getComputedStyle(el);
    if ((cs.overflowY === 'auto' || cs.overflowY === 'scroll')
        && el.scrollHeight > el.clientHeight + 50
        && el.querySelector('[role="row"]')) {
      const before = el.scrollTop;
      el.scrollTop = dy < 0 ? 0 : before + dy;
      moved = moved || el.scrollTop !== before;
    }
  }
  return moved;
}"""


async def _folder_visible(
    page, name: str, max_scrolls: int = 40, step_px: int = 500
) -> bool:
    """True iff a row named `name` exists in the current OneDrive listing.

    OneDrive virtualizes the list: only ~60 rows are in the DOM and the
    rest render as its list container scrolls. Reading the rendered rows
    once therefore reported every alphabetically-late task as missing once
    the task folder held more than ~60 entries (verify run 2026-09-10: 60 ok
    / 41 "missing", the 8 late-alphabet originals among them). Mouse-wheel
    events do not reliably reach the list container either, so scroll the
    container itself (any scrollable element that holds [role=row]) in
    steps and rescan; stop when the row set stops growing (end of list).
    Scroll back to the top afterwards so the caller's own click scan starts
    from a known position.
    """
    target = _normalize_name(name)
    # Let the new listing render: wait until [role=row] is non-empty and
    # its count holds steady across two samples (up to ~10 s). Scanning
    # earlier reads the PREVIOUS folder's rows and finds no list to scroll.
    last = -1
    for _ in range(10):
        await asyncio.sleep(1.0)
        try:
            count = await page.locator('[role="row"]').count()
        except Exception:
            count = 0
        if count > 0 and count == last:
            break
        last = count
    seen = await _visible_row_names(page)
    if target in seen:
        return True
    found = False
    stale = 0
    for _ in range(max_scrolls):
        try:
            await page.evaluate(_SCROLL_LIST_JS, step_px)
        except Exception:
            break
        await asyncio.sleep(0.6)
        names = await _visible_row_names(page)
        if target in names:
            found = True
            break
        if names <= seen:
            stale += 1
            if stale >= 5:
                break
        else:
            stale = 0
        seen |= names
    logger.debug(f"   scroll-scan for {name}: found={found} rows_seen={len(seen)}")
    try:
        await page.evaluate(_SCROLL_LIST_JS, -1)
        await asyncio.sleep(0.5)
    except Exception:
        pass
    return found


async def _create_folder(page, name: str) -> bool:
    """Create a folder named `name` in the current OneDrive view."""
    logger.info(f"   📁 Creating folder: {name}")
    opened = False
    for selector in [
        'span:has-text("Create or upload")',
        'text="Create or upload"',
        'button:has-text("New")',
        '[aria-label="New"]',
    ]:
        try:
            el = await page.query_selector(selector)
            if el and await el.is_visible():
                await el.click()
                opened = True
                break
        except Exception:
            continue
    if not opened:
        logger.error("   ❌ Could not open the Create/New menu")
        return False
    await asyncio.sleep(1.5)

    clicked_folder = False
    for selector in [
        'span:has-text("Folder")',
        'text="Folder"',
        '[aria-label="Folder"]',
    ]:
        try:
            el = await page.query_selector(selector)
            if el and await el.is_visible():
                await el.click()
                clicked_folder = True
                break
        except Exception:
            continue
    if not clicked_folder:
        logger.error("   ❌ Could not find the 'Folder' menu item")
        await page.keyboard.press("Escape")
        return False
    await asyncio.sleep(1.5)

    name_input, _loc = await ExcelOperations._find_input(
        page, ['input[type="text"]', 'input[aria-label*="name" i]', "input"]
    )
    if not name_input:
        logger.error("   ❌ Folder-name input did not appear")
        await page.keyboard.press("Escape")
        return False
    await name_input.click()
    await name_input.fill(name)
    await asyncio.sleep(0.5)

    for selector in ['button:has-text("Create")', 'text="Create"']:
        try:
            el = await page.query_selector(selector)
            if el and await el.is_visible():
                await el.click()
                break
        except Exception:
            continue
    else:
        await page.keyboard.press("Enter")
    await asyncio.sleep(2.5)
    return await _folder_visible(page, name)


async def _enter_or_create_folder(page, name: str, create: bool) -> bool:
    """Double-click into `name`, creating it first when allowed."""
    if not await _folder_visible(page, name):
        if not create:
            logger.error(f"   ❌ Folder missing: {name}")
            return False
        if not await _create_folder(page, name):
            return False
    clicked = await Navigation._try_find_and_click_folder(page, name)
    if not clicked:
        logger.error(f"   ❌ Could not open folder: {name}")
        return False
    await asyncio.sleep(1.5)
    return True


async def _upload_file(page, local_path: Path) -> bool:
    """Upload one local file into the current OneDrive folder."""
    logger.info(f"   📤 Uploading: {local_path.name}")
    opened = False
    for selector in [
        'span:has-text("Create or upload")',
        'text="Create or upload"',
        'button:has-text("Upload")',
        '[aria-label="Upload"]',
    ]:
        try:
            el = await page.query_selector(selector)
            if el and await el.is_visible():
                await el.click()
                opened = True
                break
        except Exception:
            continue
    if not opened:
        logger.error("   ❌ Could not open the Create/Upload menu")
        return False
    await asyncio.sleep(1.5)

    item = None
    for selector in [
        'span:has-text("Files upload")',
        'text="Files upload"',
        'span:has-text("File upload")',
        'text="Files"',
    ]:
        try:
            el = await page.query_selector(selector)
            if el and await el.is_visible():
                item = el
                break
        except Exception:
            continue
    if not item:
        logger.error("   ❌ Could not find the file-upload menu item")
        await page.keyboard.press("Escape")
        return False

    try:
        async with page.expect_file_chooser(timeout=10000) as fc:
            await item.click()
        chooser = await fc.value
        await chooser.set_files([str(local_path)])
    except Exception as e:
        logger.error(f"   ❌ File chooser failed: {e}")
        return False

    # Give the upload time to land, then confirm the row exists.
    for _ in range(30):
        await asyncio.sleep(2)
        if await _folder_visible(page, local_path.name):
            logger.info(f"   ✅ Uploaded: {local_path.name}")
            return True
    logger.error(f"   ❌ Upload not visible after waiting: {local_path.name}")
    return False


async def _navigate_home(page) -> bool:
    try:
        await page.goto("https://onedrive.live.com", wait_until="load", timeout=60000)
        await asyncio.sleep(3)
        return True
    except Exception as e:
        logger.error(f"❌ Could not load OneDrive: {e}")
        return False


async def provision(args) -> int:
    try:
        cfg = load_configs()
    except ConfigError as e:
        logger.error(f"Config load failed:\n{e}")
        return 2

    base_path = list(_ns_to_dict(cfg).get("onedrive_base_path") or [])
    if not base_path:
        logger.error("onedrive_base_path is empty in the config")
        return 2

    logger.info(f"Local layout: {describe_local_roots()}")
    logger.info(f"OneDrive base path: {' > '.join(base_path)}")

    try:
        source = _build_source(args.task_ids)
    except ValueError as e:
        logger.error(str(e))
        return 2
    try:
        specs = list(source.iter_tasks())
    finally:
        source.close()
    if not specs:
        logger.warning("No tasks matched.")
        return 0

    plan = []
    for spec in specs:
        local_paths = [str(p) for p in spec.upload_files]
        workbook = FileManager.find_workbook_file(local_paths)
        plan.append((spec, Path(workbook) if workbook else None))

    logger.info(f"\nPlan: {len(plan)} task folder(s) under {' > '.join(base_path)}")
    # OneDrive rejects these characters in folder/file names. A task name
    # carrying one can't be provisioned under its verbatim name — flag it
    # loudly rather than let it surface as NAV_FAILED retries later.
    _ONEDRIVE_ILLEGAL = set('"*:<>?/\\|')
    for spec, workbook in plan:
        wb = workbook.name if workbook else "(no workbook — folder only)"
        bad = _ONEDRIVE_ILLEGAL & set(spec.task_name)
        flag = f"   ⚠️ task name has OneDrive-illegal chars: {sorted(bad)}" if bad else ""
        logger.info(f"  {spec.task_name}/Task/  <- {wb}{flag}")
    if args.dry_run:
        logger.info("[DRY RUN] Nothing was created or uploaded.")
        return 0

    if args.stage is not None:
        # MANUAL-UPLOAD PATH: build the exact folder tree locally, then drag
        # it into OneDrive in one go (OneDrive web uploads folders
        # recursively). No browser involved.
        import shutil

        stage_root = Path(args.stage)
        if not stage_root.is_absolute():
            stage_root = _REPO_ROOT / stage_root
        staged = skipped = 0
        total_bytes = 0
        for spec, workbook in plan:
            task_dir = stage_root / spec.task_name / "Task"
            task_dir.mkdir(parents=True, exist_ok=True)
            if workbook is None:
                skipped += 1
                continue
            dest = task_dir / workbook.name
            shutil.copy2(workbook, dest)
            total_bytes += dest.stat().st_size
            staged += 1
        logger.info(
            f"\nStaged {staged} workbook(s) "
            f"({total_bytes / 1024 / 1024:.1f} MB; {skipped} folder-only) "
            f"under:\n  {stage_root}"
        )
        logger.info(
            f"\nNext: in OneDrive web, open "
            f"'{' > '.join(base_path)}' (create those folders if new), then "
            f"drag the task folder(s) inside {stage_root.name}/ "
            f"(e.g. '{plan[0][0].task_name}') into it. When uploads finish, "
            f"run --verify to confirm the layout and write the manifest."
        )
        return 0

    if not args.yes:
        try:
            answer = input(
                f"\n{'Verify' if args.verify else 'Provision'} "
                f"{len(plan)} task folder(s) on OneDrive? [y/N]: "
            ).strip().lower()
        except EOFError:
            answer = ""
        if answer not in {"y", "yes"}:
            logger.info("Aborted by user.")
            return 0

    browser_cfg = {"browser": _ns_to_dict(getattr(cfg, "browser", None)) or {}}
    browser_mgr = BrowserManager(browser_cfg)
    manifest: dict = {}
    if MANIFEST_PATH.exists():
        try:
            manifest = json.loads(MANIFEST_PATH.read_text())
        except Exception:
            manifest = {}

    ok = failed = 0
    async with async_playwright() as playwright:
        _browser, context = await browser_mgr.launch_browser(playwright)
        page = await context.new_page()
        try:
            for spec, workbook in plan:
                logger.info(f"\n▶ {spec.task_name}")
                if not await _navigate_home(page):
                    failed += 1
                    continue
                segments = onedrive_task_path(base_path, spec.task_name)
                reached = True
                for segment in segments:
                    if not await _enter_or_create_folder(
                        page, segment, create=not args.verify
                    ):
                        reached = False
                        break
                if not reached:
                    failed += 1
                    continue

                workbook_ok = True
                if workbook is not None:
                    present = await _folder_visible(page, workbook.name)
                    if present:
                        logger.info(f"   ✅ Workbook already present: {workbook.name}")
                    elif args.verify:
                        logger.error(f"   ❌ Workbook missing: {workbook.name}")
                        workbook_ok = False
                    else:
                        workbook_ok = await _upload_file(page, workbook)

                if workbook_ok:
                    ok += 1
                    manifest[str(spec.task_id)] = {
                        "path": segments,
                        "workbook": workbook.name if workbook else None,
                        "verified_at": datetime.now().isoformat(),
                    }
                else:
                    failed += 1
        finally:
            try:
                if not page.is_closed() and len(context.pages) > 1:
                    await page.close()
            except Exception:
                pass
            await browser_mgr.close_browser(context)

    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2, sort_keys=True))
    logger.info(f"\nDone. ok={ok} failed={failed}. Manifest: {MANIFEST_PATH}")
    return 0 if failed == 0 else 1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Provision (or verify) the OneDrive task tree from the "
        "downloaded benchmark (<data_root>/tasks)"
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the plan; touch nothing.")
    parser.add_argument("--stage", nargs="?", const="onedrive_staging",
                        default=None, metavar="DIR",
                        help="Manual-upload path: build the exact folder tree "
                        "locally under DIR (default onedrive_staging/) for a "
                        "single drag into OneDrive web. No browser.")
    parser.add_argument("--verify", action="store_true",
                        help="Read-only: check folders + workbooks, write the manifest.")
    parser.add_argument("--task-ids", nargs="*", type=int, default=None,
                        help="Restrict to these task ids (default: every task).")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="Skip the confirmation prompt.")
    args = parser.parse_args()
    return asyncio.run(provision(args))


if __name__ == "__main__":
    sys.exit(main())
