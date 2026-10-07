"""
File management utilities for the Excel Agent Engine.

Handles file discovery, path resolution, and task file management.

Supports two modes:
  - **Explicit (recommended for new projects):** caller provides `upload_files`
    and `template_file` directly per task. See `tasks_configs/examples/sample_tasks.yaml`.
  - **Task-source shorthand:** auto-discovers files from a sibling
    `main_tasks/{task_source}/{task_name}/Task/` directory. Useful when many
    tasks share a common parent. See `tasks_configs/examples/task_source_format.yaml`.
"""

import logging
import os
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)


class FileManager:
    """Manages file discovery and path operations."""

    # ------------------------------------------------------------------
    # Explicit file resolution
    # ------------------------------------------------------------------

    @staticmethod
    def resolve_upload_files(
        upload_files: List[str],
        local_files_base: Optional[str] = None,
    ) -> List[str]:
        """
        Resolve upload_files paths to absolute paths.

        Args:
            upload_files: Relative (or absolute) file paths from the task YAML.
            local_files_base: Base directory for resolving relative paths.
                              Defaults to CWD if not set.

        Returns:
            List of absolute file paths (only those that exist on disk).
        """
        base = Path(local_files_base) if local_files_base else Path.cwd()
        resolved = []
        for fp in upload_files:
            p = Path(fp)
            if not p.is_absolute():
                p = base / p
            p = p.resolve()
            if p.exists():
                resolved.append(str(p))
                logger.info(f"  📤 Upload file: {p.name}")
            else:
                logger.warning(f"  ⚠️ Upload file not found, skipping: {p}")
        logger.info(f"📤 Resolved {len(resolved)}/{len(upload_files)} upload files")
        return resolved

    # ------------------------------------------------------------------
    # Task-source shorthand helpers
    # ------------------------------------------------------------------

    @staticmethod
    def get_local_task_files(task_name: str, task_source: str) -> List[str]:
        """
        Get all files from local main_tasks/{task_source}/{task_name}/Task/ directory.

        Args:
            task_name: Name of the task folder.
            task_source: Identifies the parent dataset folder under `main_tasks/`.
                Any string is accepted; it doubles as the folder name. The
                example sources bundled with this project are `fmwc`,
                `modeloff`, and `wallstreetprep` (the last is aliased to
                `wsp/` on disk for historical reasons).

        Returns:
            List of absolute file paths
        """
        try:
            if not task_source:
                logger.error(
                    "❌ get_local_task_files called without a task_source. "
                    "Set `task_source` in your task config, or use the "
                    "explicit `upload_files` format instead."
                )
                return []

            logger.info(
                f"🔍 Searching for local files for task: '{task_name}' (source: {task_source})"
            )

            # Map task_source to folder name (wallstreetprep -> wsp historical alias)
            folder_name = "wsp" if task_source == "wallstreetprep" else task_source

            # Construct path relative to current working directory:
            #   ../main_tasks/{folder_name}/{task_name}/Task
            task_path = Path("..") / "main_tasks" / folder_name / task_name / "Task"
            logger.info(f"   Trying path 1 (relative to cwd): {task_path.absolute()}")

            # Fallback: resolve relative to this file's location, assuming the
            # repo sits alongside a sibling main_tasks/ directory.
            if not task_path.exists():
                script_dir = Path(__file__).parent.parent  # core -> excel_agent
                task_path = (
                    script_dir.parent.parent
                    / "main_tasks"
                    / folder_name
                    / task_name
                    / "Task"
                )
                logger.info(
                    f"   Trying path 2 (relative to script): {task_path.absolute()}"
                )

            if not task_path.exists():
                logger.error("❌ Local task directory not found at either location!")
                logger.error(f"   Task name: '{task_name}'")
                logger.error(f"   Task source: '{task_source}'")
                logger.error(f"   Last tried: {task_path.absolute()}")
                return []

            logger.info(f"✅ Found local task directory: {task_path.absolute()}")

            # Get all files in the directory, ignoring macOS system files
            all_files = []
            # macOS system files to ignore
            ignored_patterns = [
                ".DS_Store",
                "._*",  # macOS resource fork files
                ".AppleDouble",
                ".LSOverride",
                "Thumbs.db",  # Windows thumbnail cache
                "desktop.ini",  # Windows folder settings
            ]

            # Use os.scandir instead of pathlib.iterdir: DirEntry caches
            # is_file()/is_dir() from the directory enumeration, which works
            # for long paths on Windows. pathlib's iterdir + Path.is_file
            # does a separate stat() call that silently returns False near
            # MAX_PATH (260 chars), causing files to be invisibly skipped.
            with os.scandir(task_path) as entries:
                for entry in entries:
                    if not entry.is_file():
                        continue
                    file_name = entry.name
                    # Skip macOS system files and hidden files starting with .
                    if file_name.startswith(".") or file_name in ignored_patterns:
                        continue
                    # Skip macOS resource fork files (._filename)
                    if file_name.startswith("._"):
                        continue
                    all_files.append(os.path.abspath(entry.path))

            logger.info(f"📁 Found {len(all_files)} local files:")
            for file_path in all_files:
                logger.info(f"     📄 {Path(file_path).name}")

            return all_files

        except Exception as e:
            logger.error(f"❌ Error getting local task files: {e}")
            import traceback

            logger.error(traceback.format_exc())
            return []

    @staticmethod
    def find_workbook_file(
        file_list: List[str], task_source: str = ""
    ) -> Optional[str]:
        """
        Find a pre-existing workbook in the local task folder.

        For the example sources bundled with this project, source-specific
        filename conventions are used:
            fmwc           -> picks `*model.xlsx`
            modeloff       -> picks the first `.xlsx` / `.xlsm`
            wallstreetprep -> picks `*-before.xlsx`

        For any other source (or no source), a generic detector picks the
        first `.xlsx` / `.xlsm` that does not contain "solution" in its name.

        Args:
            file_list: List of file paths
            task_source: Optional source name; controls the matching pattern.

        Returns:
            Path to workbook file, or None if not found.
        """
        logger.info(
            f"🔍 Searching for workbook file in {len(file_list)} files (source: {task_source or 'generic'})..."
        )

        for f in file_list:
            filename = Path(f).name
            filename_lower = filename.lower()

            # Skip files with "solution" in the name (those are outputs, not inputs)
            if "solution" in filename_lower:
                continue

            if task_source == "fmwc":
                if filename_lower.endswith("model.xlsx"):
                    logger.info(f"✅ Found fmwc model file: {filename}")
                    return f

            elif task_source == "wallstreetprep":
                if filename_lower.endswith("-before.xlsx"):
                    logger.info(f"✅ Found wallstreetprep before file: {filename}")
                    return f

            else:
                # Generic detection (also covers `modeloff`): first .xlsx/.xlsm
                if filename_lower.endswith((".xlsx", ".xlsm")):
                    logger.info(f"✅ Found workbook file: {filename}")
                    return f

        logger.warning(
            f"❌ No workbook file found (source: {task_source or 'generic'})"
        )
        return None

    @staticmethod
    def get_files_to_upload(
        local_file_paths: List[str], workbook_file: Optional[str]
    ) -> List[str]:
        """
        Determine which local file paths to upload based on workbook presence.

        Args:
            local_file_paths: All local file paths
            workbook_file: Path to workbook file if found (.xlsx or .xlsm)

        Returns:
            List of files to upload
        """
        if workbook_file:
            # Upload all files EXCEPT the workbook file
            workbook_name = Path(workbook_file).name
            files_to_upload = [
                f for f in local_file_paths if Path(f).name != workbook_name
            ]
            logger.info(
                f"📤 Workbook exists locally ({workbook_name}) - will upload {len(files_to_upload)} other files"
            )
        else:
            # Upload ALL files
            files_to_upload = local_file_paths
            logger.info(
                f"📤 No workbook found locally - will upload all {len(files_to_upload)} files"
            )

        for file_path in files_to_upload:
            logger.info(f"  📤 Will upload: {Path(file_path).name}")

        return files_to_upload
