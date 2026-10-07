"""
Excel file validation for downloaded artifacts.

Checks that downloaded Excel files exist, are non-empty, and can be
opened by openpyxl (not corrupted).
"""

import logging
from pathlib import Path

from claude_web_agent.task_status import TaskStatus

logger = logging.getLogger(__name__)


def validate_excel_file(file_path) -> tuple[bool, TaskStatus, str]:
    """
    Validate a downloaded Excel file.

    Checks:
    1. File exists and has non-zero size
    2. openpyxl can open it (not corrupted)

    Args:
        file_path: Path to the Excel file

    Returns:
        (is_valid, status, message) tuple
    """
    file_path = Path(file_path)

    # Check existence and size
    if not file_path.exists():
        return False, TaskStatus.DOWNLOAD_FAILED, f"File does not exist: {file_path}"

    if file_path.stat().st_size == 0:
        return False, TaskStatus.DOWNLOAD_FAILED, f"File is empty: {file_path}"

    # Check openpyxl can open it
    try:
        import openpyxl

        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        wb.close()
    except ImportError:
        logger.warning("openpyxl not installed — skipping corruption check")
        return True, TaskStatus.SUCCESS, "openpyxl not available, skipping validation"
    except Exception as e:
        return False, TaskStatus.FILE_CORRUPTED, f"Cannot open Excel file: {e}"

    return True, TaskStatus.SUCCESS, "Valid"
