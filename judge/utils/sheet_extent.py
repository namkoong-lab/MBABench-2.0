"""Iteration bounds for worksheets whose declared extent is inflated.

openpyxl's ``max_column`` is the widest column holding *any* cell record,
including styled-but-empty ones. A sheet with 53,549 data rows and a few
formatted blanks out at column XFD (task 84's golden solution, "Solution
Model - Cashflows Sum", 2026-09-07) therefore declares 53,549 x 16,384 =
877M cells, and ``iter_rows()`` materialises every one of them -- the judge
process swapped the machine to a halt inside a GC pass and never logged
again.

``iteration_bounds`` leaves every ordinary sheet untouched: the declared
extent is returned verbatim while ``max_row * max_column`` stays within
``JUDGE_MAX_SHEET_CELLS`` (default 20M -- the largest sheet the judge has
processed to date is 297 x 16,384 = 4.9M). Only beyond that does the width
shrink to the last column holding a value, so the CSVs the judge reads
still carry every populated cell.
"""

import os

ENV_MAX_SHEET_CELLS = "JUDGE_MAX_SHEET_CELLS"
DEFAULT_MAX_SHEET_CELLS = 20_000_000


def max_sheet_cells() -> int:
    raw = os.environ.get(ENV_MAX_SHEET_CELLS)
    return int(raw) if raw else DEFAULT_MAX_SHEET_CELLS


def iteration_bounds(ws, max_cells=None):
    """Return ``(max_row, max_col, narrowed)`` for iterating ``ws``.

    ``narrowed`` is False (and the bounds equal the declared extent) unless
    the declared cell count exceeds ``max_cells``; then ``max_col`` is the
    last column holding a non-None value. Needs a regular (not read-only)
    worksheet for the narrowing step; a sheet without ``_cells`` is
    returned unchanged.
    """
    limit = max_sheet_cells() if max_cells is None else max_cells
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1
    if max_row * max_col <= limit:
        return max_row, max_col, False
    cells = getattr(ws, "_cells", None)
    if not cells:
        return max_row, max_col, False
    value_col = max(
        (c for (_r, c), cell in cells.items() if cell.value is not None),
        default=1,
    )
    return max_row, value_col, True


def iter_rows_kwargs(ws, max_cells=None):
    """kwargs for ``iter_rows`` on both the formula and data-only views.

    Empty (byte-identical behaviour) unless the sheet was narrowed.
    """
    max_row, max_col, narrowed = iteration_bounds(ws, max_cells)
    if not narrowed:
        return {}, False
    return {"min_row": 1, "max_row": max_row, "min_col": 1, "max_col": max_col}, True
