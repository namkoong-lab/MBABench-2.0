"""93 Potential Dangers/No hidden rows/columns.

Rule (handoff bucket 1; rubric "free of hidden rows or columns. Use grouping instead of hiding"):
every worksheet, dialog sheet and macro sheet of the delivered workbook (hidden sheets and case
tabs included, no starting-file diff) fails on any row or column that is hidden:

  rows     hidden="1"; or zero height: ht <= 0 (with or without customHeight); or, on a sheet
           whose rows are zero-height by default (sheetFormatPr zeroHeight="1", or
           defaultRowHeight <= 0), any row that has no height of its own (a <row> without ht,
           or a row not written at all - the unused grid hidden).
  columns  <col hidden="1">; or zero width: <col width <= 0>; or a column without a width of
           its own (no <col>, or a <col> without width) on a sheet whose defaultColWidth <= 0.
           Hidden unused columns out to XFD count like any other ("hiding the unused empty
           grid counts as hidden"); empty hidden rows/columns are not exempt.

The one exemption is a COLLAPSED OUTLINE GROUP (rubric: group instead of hiding).  A row is
exempt when all of these hold:
  * it is hidden by the hidden flag (whatever its stored height: Excel itself writes
    width="0" hidden="1" for hidden columns, e.g. Excel-saved toys 22 T4/T6;
    FLAGGED_ZERO_SIZE_EXEMPT); a zero height WITHOUT the flag is not what collapsing writes
    and is never exempt;
  * its outlineLevel L >= 1;
  * the maximal run of consecutive rows with outlineLevel >= L that contains it is hidden
    ENTIRELY (every row of the run hidden by the flag or by zero height).  That is what
    collapsing a group produces; a row hidden inside an expanded group is hiding, not grouping.
    (Runs at a lower level k < L contain the level-L run, so testing the level-L run is enough.)
  * only when REQUIRE_COLLAPSED_FLAG is on (default off): the run's summary row - the row
    after the run when sheetPr/outlinePr summaryBelow is on (default), else the row before -
    carries collapsed="1".
Columns: the same with <col outlineLevel>, <col collapsed> and summaryRight.

The decision is made from the rows' own outlineLevel, never from a collapsed flag alone: after
Excel's Ungroup the summary row keeps collapsed="1" although no group exists (93 toy T3 Fail).
Outline symbols switched off and sheet protection do not cancel the exemption (see the doc's
questions; GROUP_NEEDS_OUTLINE_SYMBOLS switches the first).

Mistakes: one per run of consecutive offending rows (or columns) that share a reason and the
same explanation of why they are not exempt, e.g.
'Solution Model'!BT:XFD; one sheet-level mistake when the sheet hides all rows without a height
of their own (zeroHeight).  Totals per sheet are in stats.  Rows/columns that are near zero but
not zero (0 < ht < 0.75 pt, 0 < width < 0.5 char) are not failed; they are listed in stats.

Reads only <row> attributes and the sheet head (sheetFormatPr, cols, sheetPr): no cell value.
"""
from __future__ import annotations

from ..core.refs import MAX_COL, MAX_ROW, index_to_col, location, parse_range
from .base import Check

# ---------------------------------------------------------------- rule options
REQUIRE_COLLAPSED_FLAG = False       # True: an exempt group also needs collapsed="1" on its summary row/col
GROUP_NEEDS_OUTLINE_SYMBOLS = False  # True: no exemption on a sheet whose outline symbols are switched off
FLAGGED_ZERO_SIZE_EXEMPT = True      # a hidden="1" member of a collapsed group is exempt even when its stored
                                     # height/width is 0 (Excel writes width="0" hidden="1" for hidden columns);
                                     # False: such a member fails (it might stay invisible when expanded)
ZERO_ROW_HEIGHT_PT = 0.0             # a row height <= this is a zero height (hidden)
ZERO_COL_WIDTH_CH = 0.0              # a column width <= this is a zero width (hidden)
NEAR_ZERO_ROW_PT = 0.75              # 0 < ht < this: reported in stats only (not failed)
NEAR_ZERO_COL_CH = 0.5               # 0 < width < this: reported in stats only (not failed)

# reason codes -> description fragments
R_FLAG = "flag"          # hidden="1"
R_ZERO = "zero"          # zero height / width of its own
R_DEFAULT = "default"    # no size of its own and a zero sheet default


def _runs(idx):
    """Sorted ints -> [(first, last)] runs of consecutive values."""
    out = []
    start = prev = None
    for i in idx:
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            out.append((start, prev))
            start = prev = i
    if start is not None:
        out.append((start, prev))
    return out


def _coded_runs(code: dict):
    """{index: code} -> [(first, last, code)] runs of consecutive indexes sharing a code."""
    out = []
    for i in sorted(code):
        cd = code[i]
        if out and i == out[-1][1] + 1 and cd == out[-1][2]:
            out[-1][1] = i
        else:
            out.append([i, i, cd])
    return [tuple(x) for x in out]


def collapsed_members(levels: dict, invisible, flag_ok, collapsed_marks=None, summary_after=True,
                      require_mark=False) -> set:
    """Indexes exempt as members of a collapsed outline group.

    levels: {index: outlineLevel >= 1}; invisible(i): index i is hidden (flag or zero size);
    flag_ok(i): i is hidden by the flag only (would show again when the group is expanded).
    An index i with level L is exempt when flag_ok(i) and the maximal run of consecutive
    indexes with level >= L containing i is entirely invisible (and, with require_mark, the
    run's summary index carries the collapsed mark)."""
    exempt = set()
    if not levels:
        return exempt
    ordered = sorted(levels)
    top = max(levels.values())
    for k in range(1, top + 1):
        members = [i for i in ordered if levels[i] >= k]
        for a, b in _runs(members):
            run = range(a, b + 1)
            if not all(invisible(i) for i in run):
                continue
            if require_mark:
                summary = b + 1 if summary_after else a - 1
                if collapsed_marks is None or summary not in collapsed_marks:
                    continue
            for i in run:
                if levels[i] == k and flag_ok(i):
                    exempt.add(i)
    return exempt


# why-not-exempt codes (one explanation per offending row/column)
W_NO_GROUP = "no_group"        # no outline level
W_ZERO = "zero"                # zero size without the hidden flag, inside a group
W_SYMBOLS = "symbols"          # group on a sheet whose outline symbols are off (option)
W_ZERO_IN_COLLAPSED = "zero_collapsed"   # flagged zero-size member of a collapsed group (option off)
W_PARTIAL = "partial"          # inside a group that is not collapsed as a whole (or lacks the mark)


def _why_not_exempt(level: int, why: str, group_ok: bool, in_collapsed: bool) -> str:
    """Why one hidden index is not exempt (a W_* code)."""
    if level <= 0:
        return W_NO_GROUP
    if why != R_FLAG:
        return W_ZERO
    if not group_ok:
        return W_SYMBOLS
    if in_collapsed:
        return W_ZERO_IN_COLLAPSED
    return W_PARTIAL


def _group_context(code: str, axis: str) -> str:
    """Why a hidden run is not exempt, as a description fragment."""
    size = "height" if axis == "row" else "width"
    if code == W_NO_GROUP:
        return " outside any outline group"
    if code == W_ZERO:
        return f" inside an outline group, but a zero {size} is not what collapsing a group writes and stays hidden when the group is expanded"
    if code == W_SYMBOLS:
        return " inside an outline group whose outline symbols are switched off"
    if code == W_ZERO_IN_COLLAPSED:
        return (f" inside a collapsed outline group, but with a zero {size} of its own it stays hidden when the "
                f"group is expanded")
    if REQUIRE_COLLAPSED_FLAG:
        return (" inside an outline group that is not a collapsed group (not hidden as a whole, or no "
                "collapsed mark on its summary " + axis + ")")
    return " inside an outline group that is not collapsed as a whole (hiding, not grouping)"


def _row_ref(a, b):
    return f"{a}:{b}"


def _col_ref(a, b):
    return f"{index_to_col(a)}:{index_to_col(b)}"


class C93(Check):
    number = 93
    key = "Potential Dangers/No hidden rows/columns"
    needs_rows = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        self._per_sheet = []
        self._n_sheets = 0

    # ------------------------------------------------------------------ per sheet
    def sheet_start(self, head):
        f = head.format
        self._default_zero_rows = bool(f.zero_height) or (f.default_row_height is not None
                                                           and f.default_row_height <= ZERO_ROW_HEIGHT_PT)
        self._row_reason = {}        # r -> R_FLAG | R_ZERO | R_DEFAULT (hidden rows only)
        self._row_flagged = set()    # rows with hidden="1"
        self._row_flag_ok = set()    # flag-hidden rows that may be exempt (non-zero height, or the option)
        self._row_level = {}         # r -> outlineLevel (>= 1)
        self._row_marks = set()      # rows with collapsed="1"
        self._n_written = 0
        self._n_own_or_flag = 0      # written rows with their own ht or the hidden flag
        self._near_zero_rows = []

    def row(self, row):
        r = row.r
        self._n_written += 1
        ht = row.ht
        own = ht is not None
        if own or row.hidden:
            self._n_own_or_flag += 1
        if own:
            zero = ht <= ZERO_ROW_HEIGHT_PT
            if not zero and ht < NEAR_ZERO_ROW_PT and len(self._near_zero_rows) < 50:
                self._near_zero_rows.append((r, ht))
        else:
            zero = self._default_zero_rows
        if row.hidden:
            self._row_reason[r] = R_FLAG
            self._row_flagged.add(r)
            if not zero or FLAGGED_ZERO_SIZE_EXEMPT:
                self._row_flag_ok.add(r)
        elif zero:
            self._row_reason[r] = R_ZERO if own else R_DEFAULT
        if row.outline_level > 0:
            self._row_level[r] = row.outline_level
        if row.collapsed:
            self._row_marks.add(r)

    def sheet_end(self, head, tail):
        self._n_sheets += 1
        name = head.name
        sp = head.sheet_pr
        symbols = sp.show_outline_symbols and all(v.show_outline_symbols for v in head.views)
        group_ok = symbols or not GROUP_NEEDS_OUTLINE_SYMBOLS
        hid_note = " (hidden sheet)" if head.state != "visible" else ""

        # ---------------- rows
        row_reason = self._row_reason
        collapsed_r = set()          # flagged members of collapsed groups (exempt unless zero size + option off)
        if group_ok:
            collapsed_r = collapsed_members(self._row_level, lambda r: r in row_reason,
                                            lambda r: r in self._row_flagged, self._row_marks, sp.summary_below,
                                            require_mark=REQUIRE_COLLAPSED_FLAG)
        exempt_r = {r for r in collapsed_r if r in self._row_flag_ok}
        bad_rows = {r: why for r, why in row_reason.items() if r not in exempt_r}
        unsized_hidden = 0
        if self._default_zero_rows:
            unsized_hidden = MAX_ROW - self._n_own_or_flag   # not written, or written without ht
        af = tail.auto_filter or {}
        af_rows = None
        if af.get("filter_columns"):
            af_rows = parse_range(af.get("ref") or "")

        # ---------------- columns
        f = head.format
        default_zero_cols = f.default_col_width is not None and f.default_col_width <= ZERO_COL_WIDTH_CH
        col_reason = [None] * (MAX_COL + 1)
        col_level = {}
        col_marks = set()
        col_flag_ok = set()
        col_flagged = set()
        near_zero_cols = []
        covered = [False] * (MAX_COL + 1)
        # the shared reading of <col> entries (core.sheet.paint_cols): where entries overlap the later
        # one - by min, then file order - wins on the columns it covers, as a whole (a later visible
        # entry clears an earlier hidden one); before 2026-10-04 this loop only ever added reasons
        for lo, hi, ci in head.col_segments():
            own = ci.width is not None
            zero = (ci.width <= ZERO_COL_WIDTH_CH) if own else default_zero_cols
            if own and not zero and ci.width < NEAR_ZERO_COL_CH and len(near_zero_cols) < 50:
                near_zero_cols.append((_col_ref(lo, hi), ci.width))
            why = R_FLAG if ci.hidden else (R_ZERO if own and zero else (R_DEFAULT if zero else None))
            for c in range(lo, hi + 1):
                covered[c] = True
                if why is not None:
                    col_reason[c] = why
                    if ci.hidden:
                        col_flagged.add(c)
                        if not zero or FLAGGED_ZERO_SIZE_EXEMPT:
                            col_flag_ok.add(c)
                if ci.outline_level > 0:
                    col_level[c] = ci.outline_level
                if ci.collapsed:
                    col_marks.add(c)
        if default_zero_cols:
            for c in range(1, MAX_COL + 1):
                if not covered[c]:
                    col_reason[c] = R_DEFAULT
        collapsed_c = set()
        if group_ok:
            collapsed_c = collapsed_members(col_level, lambda c: col_reason[c] is not None,
                                            lambda c: c in col_flagged, col_marks, sp.summary_right,
                                            require_mark=REQUIRE_COLLAPSED_FLAG)
        exempt_c = {c for c in collapsed_c if c in col_flag_ok}
        bad_cols = {c: col_reason[c] for c in range(1, MAX_COL + 1)
                    if col_reason[c] is not None and c not in exempt_c}

        # ---------------- mistakes
        n_row_mistakes = self._report_rows(name, bad_rows, hid_note, af_rows, group_ok, collapsed_r)
        if unsized_hidden > 0:
            src = ('zeroHeight="1"' if f.zero_height else f"defaultRowHeight={f.default_row_height:g}")
            self.add_mistake(location(name), f"Sheet '{name}'{hid_note} hides every row that has no height of "
                                             f"its own (sheetFormatPr {src}): {unsized_hidden:,} rows, typically "
                                             f"the unused grid, are hidden. Hidden rows are not allowed.")
        n_col_mistakes = self._report_cols(name, bad_cols, hid_note, f, col_level, group_ok, collapsed_c)

        self._per_sheet.append({
            "sheet": name, "state": head.state,
            "hidden_rows": sum(1 for w in bad_rows.values() if w == R_FLAG),
            "zero_height_rows": sum(1 for w in bad_rows.values() if w == R_ZERO),
            "default_zero_rows_written": sum(1 for w in bad_rows.values() if w == R_DEFAULT),
            "rows_hidden_by_zero_default": unsized_hidden,
            "hidden_cols": sum(1 for w in bad_cols.values() if w == R_FLAG),
            "zero_width_cols": sum(1 for w in bad_cols.values() if w == R_ZERO),
            "default_zero_width_cols": sum(1 for w in bad_cols.values() if w == R_DEFAULT),
            "grouped_collapsed_rows_exempt": len(exempt_r),
            "grouped_collapsed_cols_exempt": len(exempt_c),
            "outline_symbols_shown": symbols,
            "protected": tail.is_protected,
            "near_zero_rows": self._near_zero_rows[:20],
            "near_zero_cols": near_zero_cols[:20],
            "n_row_mistakes": n_row_mistakes, "n_col_mistakes": n_col_mistakes,
        })

    def _report_rows(self, name, bad_rows, hid_note, af_rows, group_ok, collapsed) -> int:
        # rows hidden only by the sheet's zero default height (R_DEFAULT) are reported once, at
        # sheet level, together with the rows that are not written at all
        n = 0
        for why in (R_FLAG, R_ZERO):
            code = {r: _why_not_exempt(self._row_level.get(r, 0), why, group_ok, r in collapsed)
                    for r, w in bad_rows.items() if w == why}
            for a, b, cd in _coded_runs(code):
                cnt = b - a + 1
                what = f"Row {a}" if cnt == 1 else f"Rows {a}:{b} ({cnt:,} rows)"
                verb = "is" if cnt == 1 else "are"
                if why == R_FLAG:
                    how = f"{verb} hidden (hidden=\"1\")"
                else:
                    how = f"{verb} hidden by a zero row height (ht=0, no hidden flag)"
                ctx = _group_context(cd, "row")
                if cd == W_NO_GROUP:
                    marks = [m for m in (a - 1, b + 1) if m in self._row_marks]
                    if marks and why == R_FLAG:
                        ctx += (f"; the collapsed mark on row {marks[0]} does not make a group without "
                                f"outline levels on the hidden rows")
                if af_rows is not None and why == R_FLAG and af_rows[0] < a and b <= af_rows[2]:
                    ctx += "; the rows lie in an AutoFilter range (filtered-out rows count as hidden)"
                self.add_mistake(location(name, _row_ref(a, b)),
                                 f"{what} on sheet '{name}'{hid_note} {how}{ctx}. Use grouping instead of hiding.")
                n += 1
        return n

    def _report_cols(self, name, bad_cols, hid_note, fmt, col_level, group_ok, collapsed) -> int:
        n = 0
        for why in (R_FLAG, R_ZERO, R_DEFAULT):
            code = {c: _why_not_exempt(col_level.get(c, 0), why, group_ok, c in collapsed)
                    for c, w in bad_cols.items() if w == why}
            for a, b, cd in _coded_runs(code):
                cnt = b - a + 1
                ref = _col_ref(a, b)
                what = f"Column {index_to_col(a)}" if cnt == 1 else f"Columns {ref} ({cnt:,} columns)"
                verb = "is" if cnt == 1 else "are"
                if why == R_FLAG:
                    how = f"{verb} hidden (hidden=\"1\")"
                elif why == R_ZERO:
                    how = f"{verb} hidden by a zero column width (width=0, no hidden flag)"
                else:
                    how = (f"{verb} hidden: no width of their own and the sheet's default column width is "
                           f"{fmt.default_col_width:g}")
                if b == MAX_COL and cnt > 1:
                    how += ", out to the last column XFD (the unused grid hidden),"
                ctx = _group_context(cd, "column")
                self.add_mistake(location(name, ref),
                                 f"{what} on sheet '{name}'{hid_note} {how}{ctx}. Use grouping instead of hiding.")
                n += 1
        return n

    # ------------------------------------------------------------------ verdict
    def finish(self) -> dict:
        flagged = [s for s in self._per_sheet if s["n_row_mistakes"] or s["n_col_mistakes"]
                   or s["rows_hidden_by_zero_default"]]
        tot = lambda k: sum(s[k] for s in self._per_sheet)  # noqa: E731
        stats = {
            "n_sheets_checked": self._n_sheets,
            "hidden_rows": tot("hidden_rows"), "zero_height_rows": tot("zero_height_rows"),
            "rows_hidden_by_zero_default": tot("rows_hidden_by_zero_default"),
            "hidden_cols": tot("hidden_cols"), "zero_width_cols": tot("zero_width_cols"),
            "default_zero_width_cols": tot("default_zero_width_cols"),
            "grouped_collapsed_rows_exempt": tot("grouped_collapsed_rows_exempt"),
            "grouped_collapsed_cols_exempt": tot("grouped_collapsed_cols_exempt"),
            "per_sheet": [s for s in self._per_sheet if any(
                s[k] for k in ("hidden_rows", "zero_height_rows", "default_zero_rows_written",
                               "rows_hidden_by_zero_default", "hidden_cols", "zero_width_cols",
                               "default_zero_width_cols", "grouped_collapsed_rows_exempt",
                               "grouped_collapsed_cols_exempt", "near_zero_rows", "near_zero_cols"))],
            "options": {"require_collapsed_flag": REQUIRE_COLLAPSED_FLAG,
                        "group_needs_outline_symbols": GROUP_NEEDS_OUTLINE_SYMBOLS,
                        "flagged_zero_size_exempt": FLAGGED_ZERO_SIZE_EXEMPT},
        }
        sheets = ", ".join(f"'{s['sheet']}'" for s in flagged)
        exempt = stats["grouped_collapsed_rows_exempt"] + stats["grouped_collapsed_cols_exempt"]
        ex = f" {exempt} row(s)/column(s) are hidden inside collapsed outline groups (allowed)." if exempt else ""
        return self.verdict(f"No hidden rows or columns on {self._n_sheets} sheet(s).{ex}",
                            f"{{n}} block(s) of hidden rows/columns on: {sheets}.{ex}", stats)
