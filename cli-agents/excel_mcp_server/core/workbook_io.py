"""Workbook I/O: loading, saving, file path resolution, auto-fit."""
import os
import re
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from . import shared_state


def _get_file_path(filename: str) -> Path:
    """Get full path to Excel file."""
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    return shared_state.STORAGE_PATH / filename


def _auto_fit_columns(wb: Workbook) -> None:
    """Auto-fit column widths based on cell content length."""
    for ws in wb.worksheets:
        for col_cells in ws.columns:
            max_length = 0
            # A column whose first cell lies inside a merged range starts with a
            # MergedCell, which has .column but no .column_letter: that crashed
            # every save of such a workbook ("'MergedCell' object has no
            # attribute 'column_letter'"), so copy_file could never create
            # solution.xlsx for it (2026-09-19, CoastalAggregates).
            col_letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                if isinstance(cell, MergedCell):
                    continue
                if cell.value is not None:
                    cell_len = len(str(cell.value))
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        cell_len = min(cell_len, 15)
                    max_length = max(max_length, cell_len)
            adjusted_width = min(max(max_length + 2, 10), 50)
            ws.column_dimensions[col_letter].width = adjusted_width


def _save_workbook_sync(wb: Workbook, file_path: Path) -> None:
    """Save the workbook atomically, with fsync, so an interrupted save cannot destroy it.

    2026-09-22: a format_cells call over 1.4M cells overran its 120 s tool timeout and the
    client killed the server process while openpyxl was rewriting solution.xlsx in place.
    The file was left a truncated zip - 15.6 MB, no central directory, so no styles, no
    shared strings, no workbook index - and the attempt's whole deliverable was
    ungradeable (gpt-6-astra task 89, attempt 2005; 1 of 208 attempts in that run).
    Writing to a temp file beside it and renaming over the target makes the swap atomic:
    a save killed at any moment leaves the previous good workbook untouched. The temp name
    is dot-prefixed so list_files (globs *.xlsx) never shows it and the attempt packager
    never uploads it if a SIGKILL leaves one behind.
    """
    _auto_fit_columns(wb)
    tmp_path = file_path.with_name(f".{file_path.name}.tmp-{os.getpid()}")
    try:
        wb.save(tmp_path)
        with open(tmp_path, 'r+b') as f:
            os.fsync(f.fileno())
        os.replace(tmp_path, file_path)          # atomic on the same filesystem
    except BaseException:
        try:
            tmp_path.unlink()
        except OSError:
            pass
        raise
    dir_fd = os.open(str(file_path.parent), os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(dir_fd)
    finally:
        os.close(dir_fd)


def _slugify(text: str) -> str:
    """Create a filesystem-safe slug from a title."""
    slug = re.sub(r"[^A-Za-z0-9\-_. ]+", "", text).strip().lower().replace(" ", "-")
    return slug[:80] if slug else "issue"


def _load_workbook(filename: str) -> Workbook:
    """Load Excel workbook (raw, formulas visible)."""
    file_path = _get_file_path(filename)
    if not file_path.exists():
        raise FileNotFoundError(f"Excel file '{filename}' not found")
    return openpyxl.load_workbook(file_path, data_only=False)


def _load_workbook_view(filename: str) -> Workbook:
    """Load Excel workbook returning last calculated values (view mode)."""
    file_path = _get_file_path(filename)
    if not file_path.exists():
        raise FileNotFoundError(f"Excel file '{filename}' not found")
    return openpyxl.load_workbook(file_path, data_only=True)
