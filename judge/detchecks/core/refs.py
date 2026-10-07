"""A1-reference helpers shared by the reader and the checks (stdlib only).

Rows and columns are 1-based ints everywhere in detchecks (row 1 = "1", col 1 = "A").
"""
from __future__ import annotations

import re
from typing import Iterable, Optional

MAX_ROW = 1048576
MAX_COL = 16384

_COL_CACHE: dict[str, int] = {}
_LETTER_CACHE: dict[int, str] = {}
_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")
_PART_RE = re.compile(r"^\$?([A-Za-z]{0,3})\$?([0-9]{0,7})$")
_SAFE_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def col_to_index(letters: str) -> int:
    """'A' -> 1, 'XFD' -> 16384 (case-insensitive)."""
    n = _COL_CACHE.get(letters)
    if n is None:
        n = 0
        for ch in letters.upper():
            n = n * 26 + (ord(ch) - 64)
        _COL_CACHE[letters] = n
    return n


def index_to_col(n: int) -> str:
    """1 -> 'A', 16384 -> 'XFD'."""
    s = _LETTER_CACHE.get(n)
    if s is None:
        k, s = n, ""
        while k > 0:
            k, r = divmod(k - 1, 26)
            s = chr(65 + r) + s
        _LETTER_CACHE[n] = s
    return s


def split_ref(ref: str) -> Optional[tuple[int, int]]:
    """'B7' / '$B$7' -> (7, 2); None when it is not a single-cell A1 reference."""
    m = _REF_RE.match(ref)
    if not m:
        return None
    return int(m.group(2)), col_to_index(m.group(1))


def make_ref(row: int, col: int, absolute: bool = False) -> str:
    if absolute:
        return f"${index_to_col(col)}${row}"
    return f"{index_to_col(col)}{row}"


def parse_range(ref: str) -> Optional[tuple[int, int, int, int]]:
    """'A1:C5' / 'A1' / 'A:C' / '3:7' / '$A$1:$C$5' -> (r1, c1, r2, c2), normalised so
    r1<=r2 and c1<=c2.  Whole columns span rows 1..1048576, whole rows cols 1..16384.
    None when unparseable (e.g. a defined name or an error)."""
    ref = ref.strip()
    if not ref:
        return None
    if ":" in ref:
        a, b = ref.split(":", 1)
    else:
        a = b = ref
    ma, mb = _PART_RE.match(a), _PART_RE.match(b)
    if not ma or not mb:
        return None
    ca, ra, cb, rb = ma.group(1), ma.group(2), mb.group(1), mb.group(2)
    if (not ca and not ra) or (not cb and not rb):
        return None
    if bool(ca) != bool(cb) or bool(ra) != bool(rb):
        return None
    c1 = col_to_index(ca) if ca else 1
    c2 = col_to_index(cb) if cb else MAX_COL
    r1 = int(ra) if ra else 1
    r2 = int(rb) if rb else MAX_ROW
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def parse_sqref(sqref: str | None) -> list[tuple[int, int, int, int]]:
    """Space-separated list of ranges (sqref attribute / <xm:sqref>) -> list of bounds.
    Unparseable parts are skipped."""
    out = []
    for part in (sqref or "").split():
        b = parse_range(part)
        if b is not None:
            out.append(b)
    return out


def range_to_str(r1: int, c1: int, r2: int, c2: int) -> str:
    a = make_ref(r1, c1)
    return a if (r1, c1) == (r2, c2) else f"{a}:{make_ref(r2, c2)}"


def in_bounds(row: int, col: int, b: tuple[int, int, int, int]) -> bool:
    return b[0] <= row <= b[2] and b[1] <= col <= b[3]


def quote_sheet(name: str) -> str:
    """Excel-style sheet qualifier: Sheet1 -> Sheet1, Solution Model -> 'Solution Model',
    O'Brien -> 'O''Brien'."""
    if _SAFE_SHEET.match(name) and not _REF_RE.match(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def location(sheet: str, ref: str | None = None) -> str:
    """Mistake location string: "'Solution Model'!A1:N1", or the bare sheet name when ref is None."""
    if ref is None:
        return sheet
    return f"{quote_sheet(sheet)}!{ref}"


def group_cells(cells: Iterable[tuple[int, int]]) -> list[tuple[int, int, int, int]]:
    """Group (row, col) cells into rectangles: first horizontal runs per row, then runs
    with the same column span on consecutive rows are stacked.  Output sorted by
    (top row, left col).  Every input cell is covered exactly once."""
    rows: dict[int, list[int]] = {}
    for r, c in set(cells):
        rows.setdefault(r, []).append(c)
    runs: list[tuple[int, int, int]] = []  # (row, c1, c2)
    for r in sorted(rows):
        cols = sorted(rows[r])
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
                continue
            runs.append((r, start, prev))
            start = prev = c
        runs.append((r, start, prev))
    open_rects: dict[tuple[int, int], list[int]] = {}  # (c1, c2) -> [r1, r2]
    done: list[tuple[int, int, int, int]] = []
    for r, c1, c2 in runs:          # runs are sorted by row
        key = (c1, c2)
        rect = open_rects.get(key)
        if rect is not None and rect[1] == r - 1:
            rect[1] = r
        else:
            if rect is not None:
                done.append((rect[0], c1, rect[1], c2))
            open_rects[key] = [r, r]
    for (c1, c2), (r1, r2) in open_rects.items():
        done.append((r1, c1, r2, c2))
    done.sort()
    return done


def group_cells_str(cells: Iterable[tuple[int, int]]) -> list[str]:
    return [range_to_str(*b) for b in group_cells(cells)]
