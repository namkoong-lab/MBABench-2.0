"""Values of individual cells, read through the grading's value source (Patrick 2026-10-05).

The conditional-format evaluator (checks/_cfeval.py) needs the cells a rule's formula refers to -
on the rule's own sheet or on another one ('LEFT($D$4,4)="FAIL"', "ABS(C7)>'Assumptions'!$B$9").
The engine streams each sheet once for all checks, and a sheet's rules are only known at its tail,
so those cells are read on demand here:

    vals = cell_values(wb)                     # one per workbook, shared by every check
    vals.want("Assumptions", (9, 2, 9, 2))     # register rectangles (r1, c1, r2, c2) per sheet
    vals.get("Assumptions", 9, 2)              # value, BLANK (None) or an Unavailable

The first get() on a sheet with registered-but-unread rectangles streams that sheet ONCE (SheetStream
with the same value source and trust policy the checks use: an Excel-saved file's caches, the
recalc pipeline's LibreOffice copy, constants) and keeps the values inside every rectangle
registered so far.  A value is exactly what a check's cell.value gives; an empty position is BLANK.
Unavailable (with the reason): a formula result whose value is untrusted, a member of an array /
data-table range that is not written in the file, a sheet that does not exist or cannot be read,
more than MAX_VALUES values kept.  Nothing here raises for a cell: the caller decides.
"""
from __future__ import annotations

from typing import Optional

from ..errors import GradingError
from .refs import location, make_ref, parse_range
from .sheet import SheetStream
from .values import make_context

BLANK = None
MAX_VALUES = 2_000_000          # values kept per workbook (memory); beyond: Unavailable


class Unavailable:
    """A cell value that cannot be used (why)."""
    __slots__ = ("why",)

    def __init__(self, why: str):
        self.why = why

    def __repr__(self):
        return f"Unavailable({self.why!r})"


def _in(box, r, c) -> bool:
    return box[0] <= r <= box[2] and box[1] <= c <= box[3]


class _SheetValues:
    __slots__ = ("info", "loaded", "pending", "values", "arrays", "error", "overflow", "n_streams")

    def __init__(self, info):
        self.info = info
        self.loaded: list = []       # rectangles whose present cells are in `values`
        self.pending: list = []      # registered, not streamed yet
        self.values: dict = {}       # (r, c) -> value | Unavailable
        self.arrays: list = []       # (r1, c1, r2, c2, anchor) of multi-cell array / data-table ranges
        self.error: Optional[str] = None
        self.overflow = False
        self.n_streams = 0


class CellValues:
    """On-demand cell values of one workbook (module doc)."""

    def __init__(self, pkg, value_source=None):
        self.pkg = pkg
        self.value_source = value_source
        self._sheets: dict = {}      # sheet name as in the workbook -> _SheetValues
        self.n_values = 0

    def _state(self, sheet: str):
        info = self.pkg.sheet(sheet)
        if info is None:
            return None
        st = self._sheets.get(info.name)
        if st is None:
            st = self._sheets[info.name] = _SheetValues(info)
        return st

    def want(self, sheet: str, box: tuple) -> None:
        """Register a rectangle (r1, c1, r2, c2) of `sheet` to be read with the next stream of that sheet."""
        st = self._state(sheet)
        if st is None or st.error is not None:
            return
        box = tuple(int(x) for x in box)
        if any(b[0] <= box[0] and b[1] <= box[1] and b[2] >= box[2] and b[3] >= box[3] for b in st.loaded + st.pending):
            return
        st.pending.append(box)

    def get(self, sheet: str, row: int, col: int):
        """The value of sheet!(row, col): a value, BLANK, or Unavailable."""
        st = self._state(sheet)
        if st is None:
            return Unavailable(f"there is no sheet named {sheet!r} in the workbook")
        if st.error is None and not any(_in(b, row, col) for b in st.loaded):
            if not any(_in(b, row, col) for b in st.pending):
                st.pending.append((row, col, row, col))
            self._load(st)
        if st.error is not None:
            return Unavailable(st.error)
        v = st.values.get((row, col), _MISSING)
        if v is not _MISSING:
            return v
        if st.overflow:
            return Unavailable(f"more than {MAX_VALUES:,} referenced cells: {location(st.info.name, make_ref(row, col))} "
                               f"was not kept")
        for r1, c1, r2, c2, anchor in st.arrays:
            if r1 <= row <= r2 and c1 <= col <= c2:
                return Unavailable(f"{location(st.info.name, make_ref(row, col))} is a member of the array range "
                                   f"{make_ref(r1, c1)}:{make_ref(r2, c2)} (anchor {anchor}) that is not written in the file")
        return BLANK

    def _load(self, st: _SheetValues) -> None:
        boxes = st.pending
        if not boxes:
            return
        st.pending = []
        info = st.info
        if info.kind not in ("worksheet", "macrosheet", "dialogsheet") or not info.part:
            st.error = f"sheet {info.name!r} holds no cells ({info.kind})"
            return
        r1 = min(b[0] for b in boxes)
        c1 = min(b[1] for b in boxes)
        r2 = max(b[2] for b in boxes)
        c2 = max(b[3] for b in boxes)
        one = boxes[0] if len(boxes) == 1 else None
        values = st.values
        arrays = st.arrays
        name = info.name

        def on_cell(cell):
            f = cell.formula
            if f is not None and f.ref and f.kind in ("array", "dataTable"):
                b = parse_range(f.ref)
                if b is not None and (b[0], b[1]) != (b[2], b[3]) and \
                        any(b[0] <= x[2] and x[0] <= b[2] and b[1] <= x[3] and x[1] <= b[3] for x in boxes):
                    arrays.append((b[0], b[1], b[2], b[3], cell.ref))
            r, c = cell.row, cell.col
            if r < r1 or r > r2 or c < c1 or c > c2:
                return
            if one is None and not any(_in(b, r, c) for b in boxes):
                return
            if (r, c) in values:
                return
            if self.n_values >= MAX_VALUES:
                st.overflow = True
                return
            values[(r, c)] = _value_of(cell, name)
            self.n_values += 1

        try:
            ss = SheetStream(self.pkg, info)
            try:
                ss.read_head()
                ctx = make_context(self.pkg, info, self.value_source, needs_values=True)
                ss.read_body(None, on_cell, ctx)
            finally:
                ss.close()
        except GradingError as e:
            st.error = f"sheet {name!r} could not be read: {e}"
            return
        except Exception as e:  # noqa: BLE001 - a reader problem is the caller's 'unavailable', never a crash here
            st.error = f"sheet {name!r} could not be read: {type(e).__name__}: {e}"
            return
        st.n_streams += 1
        st.loaded.extend(boxes)


_MISSING = object()


def _value_of(cell, sheet: str):
    """What cell.value gives for a check, or Unavailable."""
    if cell.is_blank:
        return BLANK
    if cell.is_formula_result:
        if not cell.value_trusted:
            return Unavailable(f"{location(sheet, cell.ref)} is a formula whose value is untrusted "
                               f"(source {cell.value_source})")
        try:
            return cell.value
        except GradingError as e:
            return Unavailable(f"{location(sheet, cell.ref)}: {e}")
    if cell.t == "d":
        return Unavailable(f"{location(sheet, cell.ref)} holds an ISO date (t=\"d\")")
    try:
        return cell.value
    except GradingError as e:
        return Unavailable(f"{location(sheet, cell.ref)}: {e}")


def cell_values(wb) -> CellValues:
    """The workbook's shared CellValues (created on first use, with the value source the engine set on
    wb.value_source: the recalc pipeline's copy, or None for the file's own caches)."""
    cv = getattr(wb, "_cell_values", None)
    if cv is None:
        cv = CellValues(wb, getattr(wb, "value_source", None))
        try:
            wb._cell_values = cv
        except AttributeError:
            pass
    return cv
