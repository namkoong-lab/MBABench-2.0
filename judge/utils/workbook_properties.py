"""Workbook/sheet properties block (judge v6, item D — "serve the blind evidence").

The 2026-09-01 evidence sweep found ~34 rubric checks graded on properties
the judge is never shown: true tab order (the file listing was sorted
alphabetically), hidden sheets/rows/columns (served as ordinary content),
data validation, column widths and row heights, cell comments, conditional
formats, hyperlinks, defined names, calc mode, print setup. This module
extracts all of that ONCE at CSV-extraction time, saves it beside the CSVs
(`_workbook_properties.json`, so the per-workbook cache carries it and repeat
gradings reuse identical facts), and renders it as a compact, deterministic
text block for the judge's seed prompt.

Where a property cannot be read the block says so explicitly ("unknown"),
so the model can tell "absent" from "not provided".

Cache generation: files written here ride in `*_csv_cache_v9` — a v2 cache
has no properties file and the loaders degrade to the old behaviour
(alphabetical listing, no block), which is why the generation was bumped.
`_v9` (2026-09-20, judge v12): schema 6. The content-fit scan measures a
number in its own font's glyph widths against the column-width unit of the
workbook's Normal font (`numeric_display_width`); before, every face was
measured as Calibri scaled by size/11, so Arial 10 came out 9% narrow and
grading 1092's WACC!C31:C39 (### in Excel) was never flagged. The implicit-
intersection scan also records which formulas reference each flagged cell.
A v8 cache carries the old fit counts: hence the bump.
`_v8` (2026-09-18, judge v11): schema 5 adds the IMPLICIT INTERSECTION scan
per sheet (utils/implicit_intersection.py) — plain formulas that Excel
evaluates to #VALUE! while the cached value (LibreOffice/openpyxl) looks
fine; grading 1075 Sens_Engine!D448. Data validations now carry their
error-alert state (`alert`, `error_style`) and render the full rule
(operator, both bounds) with a per-sheet tally: 8 of the 12 jv9 GUI
attempts had every validation alert-off and the judge passed 7 of them.
Rounding statements per sheet beside the count of formulas using a
rounding function (check 105). The freeze position carries the rows and
columns it locks with their size, tagged EXCESSIVE at render time (check
122). A v7 cache would silently lack all of it: hence the bump.
`_v7` (2026-09-16, judge v9): schema 4 adds the evidence flags from the
toy-reliability walkthrough — PERIOD SERIES scan per sheet (orientation,
OUT OF ORDER, unlabeled gaps, VERTICAL PERIOD SERIES on a horizontal tab,
content past an End-style marker), content-vs-column-width fit (NUMERIC
exceeds width / TEXT cut off), and formulas carrying a typed date-like
string literal. WIDE OUTLIER (column widths) is render-time from the
stored widths. Extraction now also receives the data-only workbook.
`_v6` (2026-09-10 pm): light yellows at the 60-degree hue boundary are named
`light_yellow`, not `olive` (canary on check 47); schema unchanged.
`_v5` (2026-09-10, judge v7 tier 2): schema 3 adds the active cell per
sheet, a count of styled-but-empty cells inside the used range, the number
of spill/array ranges per sheet, and the resolved hex beside every theme
colour (tab colours, CF styles); hidden defined names leave the rendered
list for the footnote. The cell extractor changed alongside (theme colours
tokenised, `[Red]`/locale number formats rendered, spill anchors and
children tagged, blue-family colour names, `[TEXT]` on numeric-looking
text constants).
`_v4` (2026-09-09, judge v7): schema 2 adds cell hyperlinks (openpyxl fills
`ws._hyperlinks` only when writing), page breaks, row/column grouping,
conditional-format styling, hidden defined names, a zip-based VBA test and
the attempt's original filename; the cell extractor changed alongside
(dates rendered like Excel, accounting padding, blank-format hiding, array /
data-table tagging). A v3 cache lacks all of that, so the generation moved.
"""

from __future__ import annotations

import datetime as _dt
import json
import re
import weakref
from pathlib import Path
from typing import Any, Optional

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

try:
    from .logger import logger
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger
try:
    from . import implicit_intersection as _ii
except ImportError:  # bare-module import path
    import implicit_intersection as _ii

FILENAME = "_workbook_properties.json"
SCHEMA_VERSION = 6   # 6 (2026-09-20, judge v12): typeface-aware numeric fit; implicit-intersection dependents
                     # 5 (2026-09-18, judge v11): implicit-intersection scan (Excel-only #VALUE!)
                     # 4 (2026-09-16, judge v9): period series scan, content fit, date literals in formulas
                     # 3 (2026-09-10): active cell, styled empty cells, spill counts, theme hex on colours
                     # 2 (2026-09-09): hyperlinks, page breaks, grouping, CF styles, hidden names, vba, origin
_MAX_LIST = 25          # per-list cap in the rendered text (JSON keeps everything)
_MAX_COMMENT_CHARS = 160
DEFAULT_COL_WIDTH = 8.43

# --- judge v9 evidence-flag thresholds (module constants; see README v9) ----
WIDE_OUTLIER_RATIO = 2.5       # a run of >=2 equal-width columns this many times wider than its neighbours
WIDE_OUTLIER_MIN_NEIGHBOUR = 4.0   # spacer columns (width < 4) are not neighbours
PERIOD_MIN_RUN = 3             # labels in a row/column before it counts as a period series
PERIOD_YEAR_MIN, PERIOD_YEAR_MAX = 1990, 2100
FIT_MARGIN_CHARS = 1.0         # text must exceed the column by more than this to be counted
NUMERIC_FIT_MARGIN_CHARS = 1.5 # a number must exceed the column by more than this to be called ### (estimator error band)
# judge v11 (check 122): "not more than 2/3 of a regular screen view should be blocked by freeze
# panes". A laptop grid shows roughly 450 pt of rows and 200 characters of columns at 100% zoom.
# Rows: two thirds of that. Columns: the golden sweep (2026-09-19) found label blocks up to 156
# characters wide frozen in three goldens (CapitalinMotion, FinancialRelations, DailyCash), so
# columns are tagged only when they alone fill a regular screen; goldens' rows top out at 118 pt.
FREEZE_MAX_ROWS_PT = 300.0     # ~20 standard 15-pt rows
FREEZE_MAX_COLS_CHARS = 200.0
# judge v12 (check 76): a sheet is called multi-page only when the estimate clears one page by
# this much; the estimate reads heights, widths, margins and the fit settings, never a renderer.
PRINT_MULTI_PAGE_MIN = 1.3
_PAPER_PT = {1: (612.0, 792.0), 5: (612.0, 1008.0), 8: (841.9, 1190.6), 9: (595.3, 841.9)}   # Letter, Legal, A3, A4
_COL_UNIT_PT = 5.25            # one column-width unit: 7 px at 96 dpi
# Evidence served for ONE rubric_9 check only. While that check is retired
# (project_configs judge.retired_checks) its line is left out of the rendered
# block, so the judge cannot cite it under another check; taking the check off
# the retired list renders it again. Render-time only: the JSON always keeps
# the data, so neither direction needs a schema or cache bump.
STYLED_EMPTY_CHECK = 28        # "No unused formatting" (retired 2026-09-19, judge v11)


def _safe(fn, default="unknown"):
    try:
        v = fn()
        return default if v is None else v
    except Exception:  # noqa: BLE001 — property reads must never break extraction
        return default


def _runs(pairs: list[tuple[int, Any]]) -> list[dict]:
    """[(index, value)] sorted by index -> [{first, last, value}] runs of
    consecutive indexes with equal values."""
    runs: list[dict] = []
    for idx, val in sorted(pairs, key=lambda p: p[0]):
        if runs and runs[-1]["last"] == idx - 1 and runs[-1]["value"] == val:
            runs[-1]["last"] = idx
        else:
            runs.append({"first": idx, "last": idx, "value": val})
    return runs


def _col_runs_text(runs: list[dict], fmt=lambda v: v) -> str:
    parts = []
    for r in runs:
        a, b = get_column_letter(r["first"]), get_column_letter(r["last"])
        parts.append(f"{a}{'' if a == b else ':' + b}={fmt(r['value'])}")
    return ", ".join(parts)


def _row_runs_text(runs: list[dict], fmt=lambda v: v) -> str:
    parts = []
    for r in runs:
        a, b = r["first"], r["last"]
        parts.append(f"{a}{'' if a == b else '-' + str(b)}={fmt(r['value'])}")
    return ", ".join(parts)


def _ranges_text(indexes: list[int], col: bool) -> str:
    if not indexes:
        return "none"
    runs = _runs([(i, True) for i in indexes])
    parts = []
    for r in runs:
        if col:
            a, b = get_column_letter(r["first"]), get_column_letter(r["last"])
        else:
            a, b = str(r["first"]), str(r["last"])
        parts.append(a if a == b else f"{a}:{b}" if col else f"{a}-{b}")
    return ", ".join(parts)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------


def _color_text(color, palette=None) -> Optional[str]:
    """openpyxl Color -> 'rgb:FFC000' / 'theme:4 tint +0.40 (8FAADC)' /
    'indexed:12' / None.

    `.rgb` on a theme colour is a descriptor ERROR STRING in openpyxl
    ("Values must be of type <class 'str'>"), so read the type first. With a
    `palette` (theme_palette.load_palette) the resolved hex is shown beside
    the theme slot.
    """
    if color is None:
        return None
    ctype = getattr(color, "type", None)
    if ctype == "rgb":
        rgb = getattr(color, "rgb", None)
        return f"rgb:{rgb}" if isinstance(rgb, str) and len(rgb) in (6, 8) else None
    if ctype == "theme":
        tint = getattr(color, "tint", 0) or 0
        text = f"theme:{getattr(color, 'theme', '?')}" + (f" tint {tint:+.2f}" if tint else "")
        if palette is not None:
            try:
                from . import theme_palette as _theme
            except ImportError:  # bare-module import path
                import theme_palette as _theme
            hx = _safe(lambda: _theme.resolve_color(palette, color), None)
            if hx:
                text += f" ({hx})"
        return text
    if ctype == "indexed":
        return f"indexed:{getattr(color, 'indexed', '?')}"
    return None


# Defined names that add-ins/system tooling plant in workbooks (Capital IQ,
# @RISK, Palisade, print areas, solver) — counted, not listed.
_SYSTEM_NAME_PREFIXES = ("IQ_", "IQB_", "Risk", "Pal_", "_xlnm", "solver_", "_xlfn", "Slicer_")


# ---------------------------------------------------------------------------
# judge v9 evidence flags (extraction-time; toy-reliability walkthrough 2026-09-16)
# ---------------------------------------------------------------------------

_MONTHS = {m: i for i, m in enumerate(
    ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}
_RE_YEAR = re.compile(r"^(?:FY|CY|FYE|YE)?\s?'?(\d{4})\s?[AEFPB]?$", re.I)          # 2026, FY2026, 2026E
_RE_FY2 = re.compile(r"^(?:FY|CY)\s?'?(\d{2})\s?[AEFPB]?$", re.I)                    # FY26
_RE_Q_Y = re.compile(r"^Q([1-4])\s*[-/' ]?\s*(?:FY|CY)?\s?'?(\d{4}|\d{2})\s?[AEFPB]?$", re.I)   # Q1 2028, Q1'28
_RE_Y_Q = re.compile(r"^(?:FY|CY)?\s?'?(\d{4})\s*[-/ ]?\s*Q([1-4])$", re.I)          # 2028 Q1, 2028-Q1
_RE_MON_Y = re.compile(r"^([A-Za-z]{3})[a-z]*\.?\s*[-/' ]?\s*'?(\d{4}|\d{2})$")      # Jan-26, Jan 2026, January 2026
_RE_Y_MON = re.compile(r"^(\d{4})\s*[-/ ]\s*([A-Za-z]{3})[a-z]*\.?$")                # 2026-Jan
_RE_DATE_LITERAL = re.compile(
    r"\"(?:\d{1,2}[./-]\d{1,2}[./-](?:\d{4}|\d{2})|\d{4}[./-]\d{1,2}[./-]\d{1,2})\""
)
_RE_END_MARKER = re.compile(
    r"^[\s<>\-=*_#]*end(?:\s+of)?(?:\s+(?:sheet|model|calculations?|calcs?|tab|worksheet|section))?[\s<>\-=*_#.!]*$",
    re.I,
)


def _two_digit_year(yy: int) -> int:
    return 2000 + yy if yy < 70 else 1900 + yy


def _period_key(v) -> Optional[tuple]:
    """Period-style label -> (year, month, day) sort key, or None.

    Plain numbers are NOT classified here: a row of values that happen to lie
    in 1990-2100 is data, not a timeline. `_numeric_year_run` admits numeric
    years only when a run is consecutive (2026, 2027, 2028 ...)."""
    if isinstance(v, _dt.datetime):
        return (v.year, v.month, v.day) if PERIOD_YEAR_MIN <= v.year <= PERIOD_YEAR_MAX else None
    if isinstance(v, _dt.date):
        return (v.year, v.month, v.day) if PERIOD_YEAR_MIN <= v.year <= PERIOD_YEAR_MAX else None
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s or len(s) > 24:
        return None
    # a year-only label sorts at the END of its year (12, 31): a fiscal-year
    # total column after Q4 (Q4'22, FY'22, Q1'23) is in order
    m = _RE_YEAR.match(s)
    if m:
        y = int(m.group(1))
        return (y, 12, 31) if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None
    m = _RE_FY2.match(s)
    if m:
        return (_two_digit_year(int(m.group(1))), 12, 31)
    m = _RE_Q_Y.match(s)
    if m:
        q, y = int(m.group(1)), m.group(2)
        y = int(y) if len(y) == 4 else _two_digit_year(int(y))
        return (y, q * 3, 0) if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None
    m = _RE_Y_Q.match(s)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        return (y, q * 3, 0) if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None
    m = _RE_MON_Y.match(s)
    if m and m.group(1).lower() in _MONTHS:
        y = m.group(2)
        y = int(y) if len(y) == 4 else _two_digit_year(int(y))
        return (y, _MONTHS[m.group(1).lower()], 0) if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None
    m = _RE_Y_MON.match(s)
    if m and m.group(2).lower() in _MONTHS:
        y = int(m.group(1))
        return (y, _MONTHS[m.group(2).lower()], 0) if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None
    return None


def _numeric_year(v) -> Optional[int]:
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    if float(v) != int(v):
        return None
    y = int(v)
    return y if PERIOD_YEAR_MIN <= y <= PERIOD_YEAR_MAX else None


def _label_text(v) -> str:
    if isinstance(v, _dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour or v.minute) == 0 else v.isoformat(" ")
    if isinstance(v, _dt.date):
        return v.isoformat()
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _is_formula(v) -> bool:
    return (isinstance(v, str) and v.startswith("=") and v != "=") or type(v).__name__ in ("ArrayFormula", "DataTableFormula")


def _has_content(v) -> bool:
    return v is not None and not (isinstance(v, str) and not v.strip())


def _line_runs(cells: list, positions: list[int]):
    """Period runs along one row or column.

    `cells` are the values in order, `positions` their 1-based indexes. Yields
    dicts {start, end, keys, labels, gaps} where a gap is a blank cell inside
    the run (recorded for the "unlabeled gap" evidence). String/date labels
    qualify on their own; plain numeric years qualify only as a monotonic
    run stepping by 0 or 1 with at least two distinct years (2026, 2027,
    2028 ... or a monthly model's 2027, 2028, 2028, 2028 ...).
    """
    n = len(cells)
    i = 0
    while i < n:
        key = _period_key(cells[i])
        ny = _numeric_year(cells[i]) if key is None else None
        if key is None and ny is None:
            i += 1
            continue
        start = i
        keys: list = []
        labels: list = []
        gaps: list = []
        numeric_only = key is None
        j = i
        last_num = None
        direction = None
        while j < n:
            v = cells[j]
            k = _period_key(v)
            y = _numeric_year(v) if k is None else None
            if k is not None:
                if numeric_only and keys:
                    break   # a numeric run does not absorb string labels (different header row style)
                numeric_only = False
                keys.append(k); labels.append(_label_text(v)); j += 1
                continue
            if y is not None and numeric_only:
                # numeric years: steps of 0 or 1 in one direction (a monthly
                # model repeats the year across its months; 2027 2028 2028 ...)
                if last_num is not None:
                    step = y - last_num
                    if abs(step) > 1:
                        break
                    if step != 0:
                        if direction is None:
                            direction = step
                        elif step != direction:
                            break
                keys.append((y, 12, 31)); labels.append(str(y)); last_num = y; j += 1
                continue
            if y is not None and not numeric_only:
                break
            # blank cell: a gap only when the run continues right after it
            if not _has_content(v) and keys and j + 1 < n and (
                _period_key(cells[j + 1]) is not None or (numeric_only and _numeric_year(cells[j + 1]) is not None)
            ) and (not gaps or gaps[-1] != j - 1):
                gaps.append(j); j += 1
                continue
            break
        if len(keys) >= PERIOD_MIN_RUN and not (numeric_only and len({k[0] for k in keys}) < 2):
            end = j - 1
            while end > start and not _has_content(cells[end]):
                end -= 1
            yield {"start": positions[start], "end": positions[end], "keys": keys, "labels": labels,
                   "gaps": [positions[g] for g in gaps], "numeric": numeric_only}
        i = max(j, i + 1)


def _out_of_order(keys: list) -> bool:
    """True only for a NON-monotonic run. A strictly descending series (a
    newest-first statement, DEC '25 ... DEC '20) is ordered, just reversed;
    goldens carry those on source-data sheets."""
    asc = all(b >= a for a, b in zip(keys, keys[1:]))
    desc = all(b <= a for a, b in zip(keys, keys[1:]))
    return not (asc or desc)


def _scan_period_series(ws, ws_values, bounds: dict) -> dict:
    """PERIOD SERIES evidence for rubric_9 checks 108/123/124/125 (judge v9).

    Walks the data-only sheet once. Horizontal runs (a row of >=3 period
    labels) define the sheet's main timeline. A vertical run is reported as
    VERTICAL PERIOD SERIES only on a sheet that has a horizontal run and only
    when the cells beside the run are formulas (an Assumptions-style year
    list with typed inputs beside it is a register and stays silent). Also
    reports the last End-style marker and any content below it.
    """
    out: dict[str, Any] = {"horizontal": [], "vertical": [], "out_of_order": [], "gaps": [],
                           "end_marker": None}
    rows_vals: dict[int, list] = {}
    col_cells: dict[int, list] = {}
    value_rows: dict[int, set] = {}
    end_marker = None
    min_c = bounds.get("min_col", 1)
    for r, row in enumerate(ws_values.iter_rows(**bounds), bounds.get("min_row", 1)):
        if not row:
            continue
        # read-only sheets yield EmptyCell objects without .row/.column, so
        # the row number comes from the enumeration, never from the cell
        vals = [c.value for c in row]
        rows_vals[r] = vals
        vr = set()
        for off, v in enumerate(vals):
            if _has_content(v):
                c = min_c + off
                vr.add(c)
                if isinstance(v, str) and len(v) <= 40 and _RE_END_MARKER.match(v):
                    end_marker = (r, c, v.strip())
        if vr:
            value_rows[r] = vr
    if not rows_vals:
        return out
    row_numbers = sorted(rows_vals)
    width = max(len(v) for v in rows_vals.values())
    positions_h = [min_c + k for k in range(width)]

    # horizontal
    for r in row_numbers:
        vals = rows_vals[r]
        for run in _line_runs(vals, positions_h[:len(vals)]):
            a, b = get_column_letter(run["start"]), get_column_letter(run["end"])
            rec = {"range": f"{a}{r}:{b}{r}", "orientation": "horizontal", "n": len(run["keys"]),
                   "first": run["labels"][0], "last": run["labels"][-1], "numeric": run["numeric"],
                   "descending": bool(run["keys"] and run["keys"][-1] < run["keys"][0])}
            if _out_of_order(run["keys"]):
                rec["out_of_order"] = True
                out["out_of_order"].append({"range": rec["range"], "labels": run["labels"][:12]})
            # unlabeled gaps: blank header over a column that carries values just below
            real_gaps = []
            for g in run["gaps"]:
                below = any(g in value_rows.get(rr, ()) for rr in range(r + 1, r + 16))
                if below:
                    real_gaps.append(f"{get_column_letter(g)}{r}")
            if real_gaps:
                rec["gaps"] = real_gaps
                out["gaps"].append({"range": rec["range"], "cells": real_gaps})
            out["horizontal"].append(rec)
    has_main_timeline = bool(out["horizontal"])

    # vertical (only meaningful on a sheet with a horizontal timeline)
    if has_main_timeline:
        h_cells = set()
        for rec in out["horizontal"]:
            (c1, r1, c2, _r2) = range_boundaries(rec["range"])
            for c in range(c1, c2 + 1):
                h_cells.add((r1, c))
        for k in range(width):
            c = min_c + k
            col_vals = [rows_vals[r][k] if k < len(rows_vals[r]) else None for r in row_numbers]
            for run in _line_runs(col_vals, row_numbers):
                r1, r2 = run["start"], run["end"]
                if any((r, c) in h_cells for r in range(r1, r2 + 1)):
                    continue
                rows_in = [r for r in range(r1, r2 + 1) if _has_content(rows_vals.get(r, [None] * width)[k] if k < len(rows_vals.get(r, [])) else None)]
                beside_formula = 0
                beside_any = 0
                for r in rows_in:
                    v = _safe(lambda: ws.cell(row=r, column=c + 1).value, None)
                    if _has_content(v) or _is_formula(v):
                        beside_any += 1
                        if _is_formula(v):
                            beside_formula += 1
                if not rows_in or beside_any == 0 or beside_formula * 2 < beside_any:
                    continue
                col = get_column_letter(c)
                out["vertical"].append({
                    "range": f"{col}{r1}:{col}{r2}", "orientation": "vertical", "n": len(run["keys"]),
                    "first": run["labels"][0], "last": run["labels"][-1],
                    "formulas_beside": f"{beside_formula}/{beside_any}",
                    "out_of_order": _out_of_order(run["keys"]),
                })

    if end_marker is not None and has_main_timeline:
        r, c, text = end_marker
        below = sorted(rr for rr in value_rows if rr > r)
        out["end_marker"] = {"cell": f"{get_column_letter(c)}{r}", "text": text,
                             "rows_below": len(below),
                             "first_row_below": below[0] if below else None,
                             "last_row_below": below[-1] if below else None}
    return out


def _date_literal_formulas(ws, bounds: dict) -> list[str]:
    """Cells whose formula text carries a typed date-like string literal
    ("12.12.2028", "2028-12-12", "12/12/2028") — checks 2/10/81 (judge v9).
    Plain date cell values are not touched."""
    hits: list[str] = []
    for row in ws.iter_rows(**bounds):
        for cell in row:
            v = cell.value
            text = None
            if isinstance(v, str) and v.startswith("="):
                text = v
            elif type(v).__name__ == "ArrayFormula":
                text = str(getattr(v, "text", "") or "")
            if text and '"' in text and _RE_DATE_LITERAL.search(text):
                hits.append(cell.coordinate)
    return hits


# A precision CLAIM ("rounded to $0.01", "results round to two decimals"), not any use of the
# word: the golden sweep met product names ("Round moulder"), case instructions ("please round
# your answers ...", "rounded up") and notes about "rounding noise", none of which is a label.
_RE_ROUND_WORD = re.compile(r"\bround(?:ed|s|ing)?\s+to\b", re.I)
_RE_ROUND_FUNC = re.compile(r"\b(?:ROUND|ROUNDUP|ROUNDDOWN|MROUND)\s*\(", re.I)
_MAX_ROUNDING_STATEMENTS = 6


def _rounding_statements(ws, bounds: dict) -> dict:
    """Judge v11 (check 105): the sheet's typed rounding statements ("USD,
    rounded to $0.01") beside the number of formulas on the sheet that use a
    rounding function. A label that says figures are "rounded" while no
    formula rounds them describes the display, not the model; the judge
    decides, this only serves the two facts."""
    statements, n_formulas, n_round = [], 0, 0
    for row in ws.iter_rows(**bounds):
        for cell in row:
            v = cell.value
            if isinstance(v, str):
                if v.startswith("="):
                    n_formulas += 1
                    if _RE_ROUND_FUNC.search(v):
                        n_round += 1
                elif _RE_ROUND_WORD.search(v):
                    text = " ".join(v.split())
                    statements.append({"cell": cell.coordinate,
                                       "text": text if len(text) <= 120 else text[:120] + "…"})
            elif type(v).__name__ == "ArrayFormula":
                n_formulas += 1
                if _RE_ROUND_FUNC.search(str(getattr(v, "text", "") or "")):
                    n_round += 1
    return {"statements": statements, "n_formulas": n_formulas, "n_round_formulas": n_round}


def _rounding_lines(rs) -> list[str]:
    # A sheet with no formulas (the case's Instructions text, a change log) has no figures
    # of its own to round: nothing to compare the statement with, so nothing is rendered.
    if not isinstance(rs, dict) or not rs.get("statements") or not rs.get("n_formulas"):
        return []
    st = rs["statements"]
    shown = "; ".join(f'{d["cell"]} "{d["text"]}"' for d in st[:_MAX_ROUNDING_STATEMENTS])
    more = f"; (+{len(st) - _MAX_ROUNDING_STATEMENTS} more)" if len(st) > _MAX_ROUNDING_STATEMENTS else ""
    return [
        f"     rounding statements: {shown}{more} — formulas on this sheet using a rounding function "
        f"(ROUND/ROUNDUP/ROUNDDOWN/MROUND): {rs.get('n_round_formulas', 0):,} of {rs.get('n_formulas', 0):,}"
    ]


def _char_weight(ch: str) -> float:
    """Width of one character in units of the default font's '0' (Calibri 11:
    digits 1.0, lowercase ~0.9, capitals ~1.1, narrow glyphs ~0.5)."""
    if ch in " .,:;'|!il`":
        return 0.5
    if ch in "$%()-+/\\[]{}\"*^tfrjI":
        return 0.7
    if ch in "@#&mwMW":
        return 1.3
    if ch.isupper():
        return 1.1
    if ch.isdigit():
        return 1.0
    return 0.9


def display_width(text: str, font_size: Optional[float] = None, bold: bool = False) -> float:
    """Approximate width of `text` in Excel column-width units (characters of
    the default 11pt font). Proportional-font correction is coarse on purpose:
    the fit flags carry a margin and the golden sweep calibrates them."""
    w = sum(_char_weight(ch) for ch in text)
    if font_size and font_size > 0:
        w *= float(font_size) / 11.0
    if bold:
        w *= 1.08
    return w


# Glyph advance widths in em, read from the font files with fontTools (2026-09-20): the faces the
# corpus uses on numeric cells (Arial, Aptos Narrow, Calibri, Arial Narrow; Roboto on the case-set
# Questions tabs, from its published metrics, the font is not installed here) and Excel's other
# common ones. Order: digit, punctuation (. , '), space, parenthesis, minus/hyphen, plus, percent,
# slash/colon, capital (mean A-Z), lowercase (mean a-z), digit of the bold cut.
_FACE_EM = {
    "calibri":         (0.507, 0.251, 0.226, 0.303, 0.306, 0.498, 0.715, 0.386, 0.554, 0.456, 0.507),
    "aptos narrow":    (0.507, 0.260, 0.187, 0.304, 0.306, 0.507, 0.761, 0.313, 0.551, 0.443, 0.507),
    "aptos":           (0.534, 0.286, 0.203, 0.293, 0.340, 0.534, 0.826, 0.339, 0.603, 0.483, 0.534),
    "arial":           (0.556, 0.278, 0.278, 0.333, 0.333, 0.584, 0.889, 0.278, 0.677, 0.490, 0.556),
    "arial narrow":    (0.456, 0.228, 0.228, 0.273, 0.273, 0.479, 0.729, 0.228, 0.555, 0.401, 0.456),
    "times new roman": (0.500, 0.250, 0.250, 0.333, 0.333, 0.564, 0.833, 0.278, 0.667, 0.459, 0.500),
    "verdana":         (0.636, 0.364, 0.352, 0.454, 0.454, 0.818, 1.076, 0.454, 0.687, 0.562, 0.711),
    "tahoma":          (0.546, 0.303, 0.312, 0.383, 0.363, 0.728, 0.977, 0.382, 0.608, 0.489, 0.637),
    "cambria":         (0.554, 0.205, 0.220, 0.382, 0.332, 0.554, 0.890, 0.490, 0.600, 0.488, 0.592),
    "georgia":         (0.614, 0.270, 0.241, 0.375, 0.374, 0.643, 0.817, 0.469, 0.681, 0.499, 0.701),
    "trebuchet ms":    (0.524, 0.367, 0.301, 0.367, 0.367, 0.524, 0.600, 0.524, 0.587, 0.501, 0.586),
    "courier new":     (0.600, 0.600, 0.600, 0.600, 0.600, 0.600, 0.600, 0.600, 0.600, 0.600, 0.600),
    "roboto":          (0.562, 0.230, 0.248, 0.342, 0.276, 0.567, 0.732, 0.412, 0.655, 0.535, 0.562),
}
_FACE_ALIASES = {"helvetica": "arial", "helvetica neue": "arial", "liberation sans": "arial", "arimo": "arial",
                 "carlito": "calibri", "liberation serif": "times new roman"}
_DEFAULT_FACE = "calibri"   # unknown faces are measured as the narrowest common one: a miss, never a false ###
_NORMAL_FONT_CACHE: "weakref.WeakKeyDictionary" = weakref.WeakKeyDictionary()


def _face_em(face: Optional[str]) -> tuple:
    key = (face or "").strip().lower()
    return _FACE_EM.get(_FACE_ALIASES.get(key, key), _FACE_EM[_DEFAULT_FACE])


def _px_per_em(size: Optional[float]) -> float:
    return (float(size) if size and size > 0 else 11.0) * 96.0 / 72.0


def normal_font(wb) -> tuple[str, float]:
    """(face, size) of the workbook's Normal style. Column widths are counted in digits of THIS
    font (ECMA-376 18.3.1.13), whatever font a cell uses."""
    try:
        return _NORMAL_FONT_CACHE[wb]
    except (KeyError, TypeError):
        pass
    face, size = "Calibri", 11.0
    font = _safe(lambda: wb._named_styles["Normal"].font, None) or _safe(lambda: wb._fonts[0], None)
    if font is not None:
        face = getattr(font, "name", None) or face
        size = float(getattr(font, "sz", None) or size)
    try:
        _NORMAL_FONT_CACHE[wb] = (face, size)
    except TypeError:
        pass
    return face, size


def column_unit_px(face: Optional[str] = None, size: Optional[float] = None) -> float:
    """Pixels in one column-width unit: the widest digit of the Normal font, in whole pixels
    (7 for Calibri 11, Aptos Narrow 11 and Arial 10 alike)."""
    return max(1.0, float(round(_face_em(face)[0] * _px_per_em(size))))


def numeric_display_width(text: str, face: Optional[str] = None, font_size: Optional[float] = None,
                          bold: bool = False, unit_px: float = 7.0) -> float:
    """Width of a rendered number or date in column-width units, from the cell font's own glyph
    widths (judge v12). `display_width` measures every face as Calibri scaled by size/11, which
    made Arial 10 (digits 7.4 px, the same as Calibri 11) come out 9% narrow: grading 1092,
    WACC!C31:C39, "3,276,619.94" in a 9.0 column shows ### in Excel and was never flagged.
    Bold leaves the digits of most faces unchanged; where it widens them the table says so."""
    m = _face_em(face)
    cut = (m[10] / m[0]) if bold and m[0] else 1.0
    letters = 1.06 if bold else 1.0
    em = 0.0
    for ch in text:
        if ch.isdigit() or ch in "$€£¥":
            em += m[0] * cut
        elif ch in ".,'":
            em += m[1] * cut
        elif ch in "  ":
            em += m[2]
        elif ch in "()[]":
            em += m[3] * cut
        elif ch in "-−–":
            em += m[4] * cut
        elif ch == "+":
            em += m[5] * cut
        elif ch == "%":
            em += m[6] * cut
        elif ch in "/:":
            em += m[7] * cut
        elif ch.isupper():
            em += m[8] * cut * letters
        else:
            em += m[9] * cut * letters
    return em * _px_per_em(font_size) / (unit_px or 7.0)


class _WidthMap:
    """Column -> width lookup over <col> runs (a run may span to 16384)."""

    def __init__(self, runs: list[tuple[int, int, float]], default: float):
        self.runs = sorted(runs)
        self.default = default

    def get(self, c: int, default=None) -> float:
        for lo, hi, w in self.runs:
            if lo <= c <= hi:
                return w
            if lo > c:
                break
        return self.default if default is None else default


def _column_width_map(ws) -> tuple[_WidthMap, float]:
    """Width lookup for the sheet plus its default width. A sheet-level
    defaultColWidth of 0 (Excel writes it) means the application default."""
    default = _safe(lambda: float(ws.sheet_format.defaultColWidth), None) or DEFAULT_COL_WIDTH
    runs = []
    for lo, hi, dim in _col_dims(ws):
        w = getattr(dim, "width", None)
        if w is not None and getattr(dim, "customWidth", True):
            runs.append((lo, hi, round(float(w), 2)))
    return _WidthMap(runs, float(default)), float(default)


def column_fit_summary(ws, fit_cells: list, value_cells: dict) -> dict:
    """Content-vs-column-width evidence for rubric_9 check 69 (judge v9).

    `fit_cells`: [(row, col, kind, width_chars)] collected by the cell
    extractor from the DISPLAY strings it already rendered (kind is
    "num" for numbers/dates rendered as Excel shows them, measured by
    `numeric_display_width` since judge v12, "text" for strings, measured
    by `display_width`); `value_cells`: {row: set(cols)} of populated cells. Wrapped,
    shrink-to-fit and merged cells are excluded by the collector.

    Three classes. NUMERIC exceeds width (Excel would render ###) is the one
    labelled a problem and the only one RENDERED (judge v10). Text wider than
    its column beside a non-empty neighbour is cut off on screen; 64 of the
    101 goldens carry such header labels and no width threshold separates
    them from a defect, so the count is kept in the stored JSON only (it was
    read as a defect every time when rendered, toy 69 Pass 0/3). Text
    overflowing into an empty neighbour is normal and only counted.
    """
    widths, default = _column_width_map(ws)
    hidden = set()
    for lo, hi, dim in _col_dims(ws):
        if getattr(dim, "hidden", False) and hi - lo <= 2000:
            hidden.update(range(lo, hi + 1))
    numeric: list[dict] = []
    cut: list[dict] = []
    overflow = 0
    for (r, c, kind, need) in fit_cells:
        if c in hidden:
            continue
        cap = widths.get(c)
        if cap <= 0.5:
            continue   # effectively hidden by width
        ref = f"{get_column_letter(c)}{r}"
        if kind == "num":
            if need > cap + NUMERIC_FIT_MARGIN_CHARS:
                numeric.append({"ref": ref, "need": round(need, 1), "width": cap})
        elif need > cap + FIT_MARGIN_CHARS:
            if (c + 1) in value_cells.get(r, ()):
                cut.append({"ref": ref, "need": round(need, 1), "width": cap})
            else:
                overflow += 1
    return {
        "numeric_overflow": {"count": len(numeric), "examples": numeric[:_MAX_LIST]},
        "text_cut_off": {"count": len(cut), "examples": cut[:_MAX_LIST]},
        "text_overflow_into_empty": overflow,
    }


def _print_estimate(ws, props: dict) -> Optional[dict]:
    """How many pages the sheet prints on (judge v12, check 76): the print area, else the used
    range, measured from row heights and column widths against the paper, orientation, margins
    and scaling the file stores. Excel's automatic breaks are not stored anywhere, so without
    this the judge had to guess which sheets run past a page: the check flipped on 5 of the 12
    jv9/jv11 GUI attempts, all of them workbooks with breaks on some long sheets and none on
    others (grading 1092: Summary, 133 rows, fit to one page wide, no breaks). Sparse arithmetic
    over the stored dimensions, so a sheet declared at full height costs nothing."""
    area = ws.print_area
    area = (area[0] if isinstance(area, (list, tuple)) and area else area) or ""
    ref = str(area).split(",")[0].rsplit("!", 1)[-1].replace("$", "")
    basis = "print area"
    if not ref:
        ref, basis = str(props.get("used_range") or ""), "used range"
    if not ref or ref == "unknown":
        return None
    c1, r1, c2, r2 = range_boundaries(ref)
    if None in (c1, r1, c2, r2):
        return None
    default_h = float(ws.sheet_format.defaultRowHeight or 15.0)
    height = default_h * (r2 - r1 + 1)
    for idx, dim in ws.row_dimensions.items():
        if r1 <= int(idx) <= r2:
            h = float(dim.height) if dim.height is not None else default_h
            height += (0.0 if getattr(dim, "hidden", False) else h) - default_h
    widths, default_w = _column_width_map(ws)
    hidden_cols = set(props.get("hidden_cols") or []) if isinstance(props.get("hidden_cols"), list) else set()
    width = sum(0.0 if c in hidden_cols else widths.get(c) for c in range(c1, min(c2, c1 + 2000) + 1)) * _COL_UNIT_PT
    if c2 - c1 > 2000:
        width += default_w * (c2 - c1 - 2000) * _COL_UNIT_PT
    ps, pm = ws.page_setup, ws.page_margins
    pw, ph = _PAPER_PT.get(int(ps.paperSize or 1), _PAPER_PT[1])
    if ps.orientation == "landscape":
        pw, ph = ph, pw
    avail_w = pw - 72.0 * (float(pm.left or 0) + float(pm.right or 0))
    avail_h = ph - 72.0 * (float(pm.top or 0) + float(pm.bottom or 0))
    if avail_w <= 0 or avail_h <= 0 or height <= 0 or width <= 0:
        return None
    fit = bool(props.get("fit_to_page"))
    fit_w = (1 if ps.fitToWidth is None else int(ps.fitToWidth)) if fit else None
    fit_h = (1 if ps.fitToHeight is None else int(ps.fitToHeight)) if fit else None
    if fit:
        scale = min([1.0] + ([avail_w * fit_w / width] if fit_w else []) + ([avail_h * fit_h / height] if fit_h else []))
        scale = max(scale, 0.10)
    else:
        scale = float(ps.scale or 100) / 100.0
    title_h = 0.0
    titles = str(ws.print_title_rows or "").replace("$", "")
    if ":" in titles:
        t1, t2 = (int(x) for x in titles.split(":"))
        title_h = sum(float(ws.row_dimensions[r].height or default_h) if r in ws.row_dimensions else default_h
                      for r in range(t1, t2 + 1)) * scale
    tall = height * scale / avail_h
    if tall > 1 and avail_h > title_h * 2:
        tall = 1 + (height * scale - avail_h) / (avail_h - title_h)
    return {"basis": basis, "range": ref, "pages_tall": round(tall, 2), "pages_wide": round(width * scale / avail_w, 2),
            "scale": round(scale, 2), "fit_w": fit_w, "fit_h": fit_h,
            "row_breaks_in_range": sum(1 for b in (props.get("row_breaks") or []) if isinstance(b, int) and r1 <= b < r2),
            "col_breaks_in_range": sum(1 for b in (props.get("col_breaks") or []) if isinstance(b, int) and c1 <= b < c2)}


def _print_estimate_lines(s: dict) -> list[str]:
    pe = s.get("print_estimate")
    if not isinstance(pe, dict) or _WIDE_OUTLIER_SKIP_SHEETS.search(str(s.get("name") or "")):
        return []   # the case's own brief / questions sheet is not the agent's print setup
    tall, wide = float(pe.get("pages_tall") or 0), float(pe.get("pages_wide") or 0)
    how = (f"fit to {pe['fit_w'] or 'any'} wide x {pe['fit_h'] or 'any'} tall" if pe.get("fit_w") is not None
           else "no fit-to-page")
    size = f"about {tall:.1f} pages tall x {wide:.1f} wide" if max(tall, wide) > 1.0 else "fits one page"
    line = (f"     print estimate ({pe.get('basis')} {pe.get('range')}): {size} at {round(100 * float(pe.get('scale') or 1))}% "
            f"({how}); manual breaks inside it: {pe.get('row_breaks_in_range', 0)} row, {pe.get('col_breaks_in_range', 0)} col")
    over = []
    if tall > PRINT_MULTI_PAGE_MIN and not pe.get("row_breaks_in_range"):
        over.append("down")
    if wide > PRINT_MULTI_PAGE_MIN and not pe.get("col_breaks_in_range"):
        over.append("across")
    if over:
        line += f" — MULTI-PAGE, NO MANUAL BREAKS ({' and '.join(over)}): Excel cuts the pages wherever they run out"
    return [line]


def _sheet_properties(ws, index: int, output_name: Optional[str], palette=None,
                      ws_values=None, column_fit: Optional[dict] = None) -> dict:
    kind = "chartsheet" if type(ws).__name__ == "Chartsheet" else "worksheet"
    props: dict[str, Any] = {
        "name": ws.title,
        "output_name": output_name,
        "index": index,
        "kind": kind,
        "state": _safe(lambda: ws.sheet_state, "visible"),
        "tab_color": _safe(lambda: _color_text(ws.sheet_properties.tabColor, palette), None),
    }
    if kind == "chartsheet":
        return props

    props.update({
        "max_row": _safe(lambda: ws.max_row, 0),
        "max_column": _safe(lambda: ws.max_column, 0),
        "zoom": _safe(lambda: ws.sheet_view.zoomScale, None),
        "gridlines": _safe(lambda: ws.sheet_view.showGridLines, None),
        "freeze_panes": _safe(lambda: ws.freeze_panes, None),
        "freeze_extent": _safe(lambda: _freeze_extent(ws), None),
        "merged_ranges": _safe(lambda: sorted(str(r) for r in ws.merged_cells.ranges), []),
        "protected": _safe(lambda: bool(ws.protection.sheet), None),
        "print_area": _safe(lambda: ws.print_area, None),
        "print_title_rows": _safe(lambda: ws.print_title_rows, None),
        "print_title_cols": _safe(lambda: ws.print_title_cols, None),
        "page_orientation": _safe(lambda: ws.page_setup.orientation, None),
        "fit_to_page": _safe(
            lambda: bool(ws.sheet_properties.pageSetUpPr.fitToPage)
            if ws.sheet_properties.pageSetUpPr else False,
            None,
        ),
    })

    # used range + comments + hyperlinks + counts in one pass over the cells
    min_r = min_c = None
    max_r = max_c = 0
    n_values = n_formulas = n_spills = 0
    comments = []
    links = []
    styled_empty: list[tuple[int, int, str]] = []  # (row, col, coordinate), row-major
    try:
        try:
            from .sheet_extent import iter_rows_kwargs
        except ImportError:  # bare-module import path
            from sheet_extent import iter_rows_kwargs
        _bounds, _ = iter_rows_kwargs(ws)
        for row in ws.iter_rows(**_bounds):
            for cell in row:
                v = cell.value
                if v is None:
                    if _is_styled_empty(cell):
                        styled_empty.append((cell.row, cell.column, cell.coordinate))
                    continue
                r, c = cell.row, cell.column
                min_r = r if min_r is None or r < min_r else min_r
                min_c = c if min_c is None or c < min_c else min_c
                max_r, max_c = max(max_r, r), max(max_c, c)
                n_values += 1
                if isinstance(v, str) and v.startswith("=") and v != "=":
                    # a bare "=" is a spilled child (<f ca="1"/>), not a formula
                    n_formulas += 1
                elif type(v).__name__ == "ArrayFormula":
                    n_formulas += 1
                    ref = str(getattr(v, "ref", "") or "")
                    if ":" in ref and ref.split(":")[0] != ref.split(":")[1]:
                        n_spills += 1
                elif type(v).__name__ == "DataTableFormula":
                    n_formulas += 1
            for cell in row:
                cm = getattr(cell, "comment", None)
                if cm is not None:
                    comments.append({
                        "ref": cell.coordinate,
                        "author": getattr(cm, "author", None),
                        "text": (getattr(cm, "text", "") or "")[:_MAX_COMMENT_CHARS],
                    })
                # Hyperlink objects live on the cell after a load; the
                # sheet-level ws._hyperlinks list is only populated on write.
                hl = getattr(cell, "hyperlink", None)
                if hl is not None:
                    links.append({
                        "ref": cell.coordinate,
                        "target": getattr(hl, "target", None) or getattr(hl, "location", None),
                        "display": getattr(hl, "display", None),
                    })
        props["used_range"] = (
            f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
            if min_r is not None else None
        )
        props["n_values"] = n_values
        props["n_formulas"] = n_formulas
        props["n_spill_anchors"] = n_spills
        props["comments"] = comments
        props["hyperlinks"] = sorted(links, key=lambda d: d["ref"])
        # Styled cells with no value INSIDE the used range (rubric_9 check 28,
        # "no unused formatting"): the extractor skips them, so this count is
        # the only place the judge can see leftover formatting sprawl.
        if min_r is not None:
            inside = [
                coord for (r, c, coord) in styled_empty
                if min_r <= r <= max_r and min_c <= c <= max_c
            ]
        else:
            inside = []
        props["styled_empty_cells"] = {"count": len(inside), "examples": inside[:10]}
    except Exception:  # noqa: BLE001
        props["used_range"] = "unknown"
        props["comments"] = "unknown"
        props["hyperlinks"] = "unknown"
        props["styled_empty_cells"] = "unknown"

    # hidden rows / cols, widths / heights, outline grouping
    try:
        hidden_rows, heights, row_groups = [], [], []
        for idx, dim in ws.row_dimensions.items():
            if getattr(dim, "hidden", False):
                hidden_rows.append(int(idx))
            h = getattr(dim, "height", None)
            if h is not None and getattr(dim, "customHeight", None) is not False:
                heights.append((int(idx), round(float(h), 1)))
            lvl = getattr(dim, "outlineLevel", 0) or 0
            if lvl:
                row_groups.append((int(idx), int(lvl)))
        props["hidden_rows"] = sorted(hidden_rows)
        props["row_heights"] = _runs(heights)
        props["row_groups"] = _runs(row_groups)
        props["default_row_height"] = _safe(lambda: ws.sheet_format.defaultRowHeight, None)
    except Exception:  # noqa: BLE001
        props["hidden_rows"] = "unknown"
        props["row_heights"] = "unknown"
        props["row_groups"] = "unknown"
    try:
        hidden_cols, width_runs, col_groups = [], [], []
        for lo, hi, dim in _col_dims(ws):
            span = hi - lo + 1
            if getattr(dim, "hidden", False) and span <= 2000:
                hidden_cols.extend(range(lo, hi + 1))
            w = getattr(dim, "width", None)
            if w is not None and getattr(dim, "customWidth", True):
                # a <col min=9 max=308> run is one entry, however wide (judge
                # v9: the old 200-column cap collapsed real model timelines
                # to their first column and misread every other width)
                width_runs.append({"first": lo, "last": hi, "value": round(float(w), 2)})
            lvl = getattr(dim, "outlineLevel", 0) or 0
            if lvl and span <= 2000:
                for c in range(lo, hi + 1):
                    col_groups.append((c, int(lvl)))
        props["hidden_cols"] = sorted(set(hidden_cols))
        props["column_widths"] = _merge_runs(width_runs)
        props["col_groups"] = _runs(col_groups)
        props["default_col_width"] = _safe(lambda: ws.sheet_format.defaultColWidth, None)
    except Exception:  # noqa: BLE001
        props["hidden_cols"] = "unknown"
        props["column_widths"] = "unknown"
        props["col_groups"] = "unknown"

    # active cell (rubric_9 check 62). openpyxl stores one <selection> per
    # pane; on a frozen-pane sheet the first one belongs to a corner pane, so
    # the cursor is the selection whose pane is the view's activePane.
    try:
        sv = ws.sheet_view
        sels = list(getattr(sv, "selection", None) or [])
        pane = getattr(sv, "pane", None)
        active_pane = getattr(pane, "activePane", None) if pane is not None else None
        chosen = next((sel for sel in sels if getattr(sel, "pane", None) == active_pane), None)
        if chosen is None and sels:
            chosen = sels[-1]
        ac = None
        if chosen is not None:
            ac = getattr(chosen, "activeCell", None)
            if not ac:
                sq = str(getattr(chosen, "sqref", "") or "")
                ac = sq.split()[0].split(":")[0] if sq else None
        # No <selection> element at all is Excel's default: cursor on A1.
        props["active_cell"] = ac or ("A1" if not sels else "unknown")
        tlc = getattr(sv, "topLeftCell", None)
        props["top_left_cell"] = str(tlc) if tlc else None
    except Exception:  # noqa: BLE001
        props["active_cell"] = "unknown"
        props["top_left_cell"] = None

    # page breaks (manual breaks only; automatic ones are not stored)
    try:
        props["row_breaks"] = sorted(int(b.id) for b in ws.row_breaks.brk)
        props["col_breaks"] = sorted(int(b.id) for b in ws.col_breaks.brk)
    except Exception:  # noqa: BLE001
        props["row_breaks"] = "unknown"
        props["col_breaks"] = "unknown"
    props["print_estimate"] = _safe(lambda: _print_estimate(ws, props), "unknown")   # judge v12, check 76

    # data validation / conditional formatting / hyperlinks
    try:
        dvs = []
        for dv in (ws.data_validations.dataValidation if ws.data_validations else []):
            dvs.append({
                "sqref": str(dv.sqref),
                "type": dv.type,
                "operator": dv.operator,
                "formula1": dv.formula1,
                "formula2": dv.formula2,
                "allow_blank": dv.allowBlank,
                # judge v11: a validation whose error alert is off never rejects
                # or flags an entry (Excel accepts anything typed); the judge
                # passed 7 of 8 attempts whose validations were all alert-off.
                "alert": bool(dv.showErrorMessage),
                "error_style": (dv.errorStyle or "stop") if dv.showErrorMessage else None,
            })
        props["data_validations"] = sorted(dvs, key=lambda d: d["sqref"])
    except Exception:  # noqa: BLE001
        props["data_validations"] = "unknown"
    try:
        cfs = []
        for cf in ws.conditional_formatting:
            cfs.append({
                "sqref": str(cf.sqref),
                "rules": [
                    {"type": r.type, "operator": getattr(r, "operator", None),
                     "formula": list(getattr(r, "formula", []) or []),
                     "style": _dxf_style(getattr(r, "dxf", None), palette)}
                    for r in cf.rules
                ],
            })
        props["conditional_formats"] = sorted(cfs, key=lambda d: d["sqref"])
    except Exception:  # noqa: BLE001
        props["conditional_formats"] = "unknown"

    # judge v9 evidence flags. Each is independent and never fatal.
    try:
        try:
            from .sheet_extent import iter_rows_kwargs as _irk
        except ImportError:  # bare-module import path
            from sheet_extent import iter_rows_kwargs as _irk
        _b, _ = _irk(ws)
    except Exception:  # noqa: BLE001
        _b = {}
    try:
        props["date_literal_formulas"] = _date_literal_formulas(ws, _b)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"  [properties] date-literal scan failed on '{ws.title}': {e}")
        props["date_literal_formulas"] = "unknown"
    try:
        props["rounding_statements"] = _rounding_statements(ws, _b)   # judge v11, check 105
    except Exception as e:  # noqa: BLE001
        logger.warning(f"  [properties] rounding-statement scan failed on '{ws.title}': {e}")
        props["rounding_statements"] = "unknown"
    try:
        # judge v11: plain formulas Excel evaluates to #VALUE! by implicit intersection
        props["implicit_intersection"] = _ii.scan_sheet(ws, _b)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"  [properties] implicit-intersection scan failed on '{ws.title}': {e}")
        props["implicit_intersection"] = "unknown"
    if ws_values is not None:
        try:
            props["period_series"] = _scan_period_series(ws, ws_values, _b)
        except Exception as e:  # noqa: BLE001
            logger.warning(f"  [properties] period-series scan failed on '{ws.title}': {e}")
            props["period_series"] = "unknown"
    else:
        props["period_series"] = "unknown"
    props["column_fit"] = column_fit if column_fit is not None else "unknown"
    return props


def _freeze_extent(ws) -> Optional[dict]:
    """Rows/columns locked by the sheet's freeze position and their size:
    `A35` locks 34 rows; the judge was only ever shown the cell and never
    cited an excessive freeze (12 jv9 GUI gradings: 660/480/475/306 pt
    locked on four attempts, goldens at most 8 rows / 118 pt). Hidden rows
    take no screen space and are not counted."""
    fp = ws.freeze_panes
    if not fp:
        return None
    m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", str(fp))
    if not m:
        return None
    from openpyxl.utils.cell import column_index_from_string
    n_rows, n_cols = int(m.group(2)) - 1, column_index_from_string(m.group(1).upper()) - 1
    default_h = float(getattr(ws.sheet_format, "defaultRowHeight", None) or 15.0)
    rows_pt = 0.0
    for r in range(1, n_rows + 1):
        d = ws.row_dimensions[r] if r in ws.row_dimensions else None
        if d is not None and d.hidden:
            continue
        rows_pt += float(d.height) if d is not None and d.height else default_h
    widths, default_w = _column_width_map(ws)
    cols_chars = sum(float(widths.get(c, default_w) or default_w) for c in range(1, n_cols + 1))
    return {"rows": n_rows, "cols": n_cols, "rows_pt": round(rows_pt, 1), "cols_chars": round(cols_chars, 1)}


def freeze_text(sheet_props: dict) -> str:
    """'A35 (34 rows ≈ 660 pt frozen, 0 columns) EXCESSIVE: …' — the tag is
    render-time from the stored extent, so the thresholds can move without
    re-extraction. Caches without the extent render the cell alone."""
    fp = sheet_props.get("freeze_panes")
    if not fp:
        return "none"
    ext = sheet_props.get("freeze_extent")
    if not isinstance(ext, dict):
        return str(fp)
    rows, cols = int(ext.get("rows") or 0), int(ext.get("cols") or 0)
    rows_pt, cols_chars = float(ext.get("rows_pt") or 0), float(ext.get("cols_chars") or 0)
    parts = [f"{rows} row{'s' if rows != 1 else ''}" + (f" ≈ {rows_pt:.0f} pt" if rows else "") + " frozen",
             f"{cols} column{'s' if cols != 1 else ''}" + (f" ≈ {cols_chars:.0f} characters wide" if cols else "")]
    text = f"{fp} ({', '.join(parts)})"
    over = []
    if rows_pt > FREEZE_MAX_ROWS_PT:
        over.append(f"{rows_pt:.0f} pt of rows")
    if cols_chars > FREEZE_MAX_COLS_CHARS:
        over.append(f"{cols_chars:.0f} characters of columns")
    if over:
        text += (" EXCESSIVE: " + " and ".join(over)
                 + " locked, more than two thirds of a regular screen; what lies beyond cannot be scrolled into a useful view")
    return text


def _col_dims(ws):
    """(lo, hi, dim) for every column_dimensions entry, 1-based inclusive."""
    from openpyxl.utils import column_index_from_string
    out = []
    try:
        for key, dim in ws.column_dimensions.items():
            lo = int(getattr(dim, "min", None) or 0) or None
            hi = int(getattr(dim, "max", None) or 0) or None
            if lo is None:
                lo = hi = column_index_from_string(key)
            hi = hi or lo
            if hi < lo:
                lo, hi = hi, lo
            out.append((lo, hi, dim))
    except Exception:  # noqa: BLE001
        pass
    return sorted(out, key=lambda t: t[0])


def _merge_runs(runs: list[dict]) -> list[dict]:
    """Sort width runs and merge adjacent equal-value runs."""
    merged: list[dict] = []
    for r in sorted(runs, key=lambda r: r["first"]):
        if merged and merged[-1]["last"] == r["first"] - 1 and merged[-1]["value"] == r["value"]:
            merged[-1]["last"] = r["last"]
        else:
            merged.append(dict(r))
    return merged


def _is_styled_empty(cell) -> bool:
    """A value-less cell carrying visible formatting: any border side, bold,
    a fill with a fill type, or a number format other than General.
    `cell.has_style` alone is too broad (every cell touched by a style row)."""
    try:
        if not getattr(cell, "has_style", False):
            return False
        font = cell.font
        if font is not None and getattr(font, "bold", None):
            return True
        fill = cell.fill
        if fill is not None and getattr(fill, "fill_type", None):
            return True
        nf = cell.number_format
        if nf and nf != "General":
            return True
        border = cell.border
        if border is not None and any(
            getattr(getattr(border, side, None), "style", None)
            for side in ("top", "bottom", "left", "right")
        ):
            return True
    except Exception:  # noqa: BLE001
        return False
    return False


def _dxf_style(dxf, palette=None) -> Optional[dict]:
    """The style a conditional-format rule applies (openpyxl resolves dxfId
    on load): font colour, fill and bold. None when the rule carries none.
    Solid dxf fills usually store the colour in bgColor, so both are read."""
    if dxf is None:
        return None
    out: dict[str, Any] = {}
    try:
        font = getattr(dxf, "font", None)
        if font is not None:
            fc = _color_text(getattr(font, "color", None), palette)
            if fc and not fc.endswith(":00000000"):
                out["font"] = fc
            if getattr(font, "bold", None):
                out["bold"] = True
        fill = getattr(dxf, "fill", None)
        if fill is not None:
            bg = _color_text(getattr(fill, "bgColor", None), palette)
            fg = _color_text(getattr(fill, "fgColor", None), palette)
            chosen = None
            for cand in (bg, fg):
                if cand and not cand.endswith(":00000000"):
                    chosen = cand
                    break
            if chosen:
                out["fill"] = chosen
    except Exception:  # noqa: BLE001
        return out or None
    return out or None


def extract_workbook_properties(workbook, excel_file_path, name_map: dict | None = None,
                                workbook_values=None, column_fit: dict | None = None) -> dict:
    """Everything the rubric grades that is not a cell value.

    `name_map` maps original sheet names to the output (filtered/safe) names
    the CSVs were saved under; sheets skipped by the filter map to None.
    `workbook_values` (judge v9) is the data-only load of the same file — the
    period-series scan reads cached values; without it that block is
    "unknown". `column_fit` maps sheet name -> column_fit_summary() output
    collected by the cell extractor.
    """
    path = Path(excel_file_path)
    wb: dict[str, Any] = {
        "filename": path.name,
        "bytes": _safe(lambda: path.stat().st_size, None),
        # load_workbooks never sets keep_vba, so vba_archive is always None;
        # the zip listing is the reliable test and needs no load-flag change.
        "has_vba": _safe(lambda: _zip_has_vba(path), None),
        "calc_mode": _safe(lambda: workbook.calculation.calcMode, None),
        "full_calc_on_load": _safe(lambda: workbook.calculation.fullCalcOnLoad, None),
        "iterative_calc": _safe(lambda: workbook.calculation.iterate, None),
        "defined_names": [],
        "external_links": [],
        "active_sheet": _safe(lambda: workbook.active.title if workbook.active else None, None),
    }
    try:
        names = []
        dn = workbook.defined_names
        items = dn.items() if hasattr(dn, "items") else [(d.name, d) for d in dn.definedName]
        for name, d in items:
            names.append({"name": name, "refers_to": getattr(d, "attr_text", None), "scope": None,
                          "hidden": bool(getattr(d, "hidden", False))})
        for ws in workbook.worksheets:
            local = getattr(ws, "defined_names", None)
            if local and hasattr(local, "items"):
                for name, d in local.items():
                    names.append({"name": name, "refers_to": getattr(d, "attr_text", None),
                                  "scope": ws.title,
                                  "hidden": bool(getattr(d, "hidden", False))})
        wb["defined_names"] = sorted(names, key=lambda d: (d["scope"] or "", d["name"]))
    except Exception:  # noqa: BLE001
        wb["defined_names"] = "unknown"
    try:
        wb["external_links"] = sorted(
            str(getattr(getattr(l, "file_link", None), "Target", None) or "?")
            for l in workbook._external_links
        )
    except Exception:  # noqa: BLE001
        wb["external_links"] = "unknown"

    try:
        from . import theme_palette as _theme
    except ImportError:  # bare-module import path
        import theme_palette as _theme
    palette = _safe(lambda: _theme.load_palette(workbook), None)

    sheets = []
    name_map = name_map or {}
    column_fit = column_fit or {}
    for i, ws in enumerate(workbook._sheets if hasattr(workbook, "_sheets") else workbook.worksheets, 1):
        ws_values = None
        if workbook_values is not None and type(ws).__name__ != "Chartsheet":
            ws_values = _safe(lambda: workbook_values[ws.title], None)
        sheets.append(_sheet_properties(ws, i, name_map.get(ws.title, ws.title), palette,
                                        ws_values=ws_values, column_fit=column_fit.get(ws.title)))
    _attach_implicit_intersection_dependents(workbook, sheets)
    return {"schema": SCHEMA_VERSION, "workbook": wb, "sheets": sheets}


def _attach_implicit_intersection_dependents(workbook, sheets: list) -> None:
    """judge v12 (check 32): list, beside every flagged IMPLICIT INTERSECTION cell, the formulas
    that reference it. One workbook-wide pass, and only when something was flagged."""
    flagged = {s["name"]: [e["cell"] for e in s["implicit_intersection"].get("examples") or []]
               for s in sheets if isinstance(s.get("implicit_intersection"), dict)
               and s["implicit_intersection"].get("count")}
    if not flagged:
        return
    try:
        try:
            from .sheet_extent import iter_rows_kwargs as _irk
        except ImportError:  # bare-module import path
            from sheet_extent import iter_rows_kwargs as _irk
        deps = _ii.find_dependents(workbook, flagged, bounds_for=lambda ws: _irk(ws)[0])
    except Exception as e:  # noqa: BLE001 - evidence must never break extraction
        logger.warning(f"implicit-intersection dependents scan failed: {e}")
        return
    for s in sheets:
        if s["name"] in flagged:
            for ex in s["implicit_intersection"]["examples"]:
                ex["referenced_by"] = deps.get((s["name"], ex["cell"].replace("$", "").upper()), [])


def _zip_has_vba(path: Path) -> Optional[bool]:
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            return "xl/vbaProject.bin" in z.namelist()
    except (zipfile.BadZipFile, OSError):
        return None


ORIGIN_FILENAME = "_attempt_origin.json"


def load_origin(directory) -> Optional[dict]:
    """The attempt's provenance sidecar written by grade.setup_task_folder
    (original filename + source path); None when absent (hand-assembled folders)."""
    if not directory:
        return None
    p = Path(directory) / ORIGIN_FILENAME
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def save_properties(output_dir: Path, props: dict) -> Path:
    p = Path(output_dir) / FILENAME
    p.write_text(json.dumps(props, indent=2, default=str), encoding="utf-8")
    return p


def load_properties(directory) -> Optional[dict]:
    if not directory:
        return None
    p = Path(directory) / FILENAME
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        logger.warning(f"  [properties] unreadable {p}: {e}")
        return None


# ---------------------------------------------------------------------------
# Ordering + rendering
# ---------------------------------------------------------------------------


def order_file_list(file_list: list[str], props: Optional[dict]) -> list[str]:
    """Order `*_full.csv` names by the workbook's TRUE tab order.

    Names not present in the properties (older caches, filtered sheets) keep
    an alphabetical order after the known ones, so a v2 cache degrades to
    exactly the old listing.
    """
    if not props:
        return sorted(file_list)
    rank = {}
    for s in props.get("sheets", []):
        out_name = s.get("output_name")
        if out_name:
            from .excel_utils import create_safe_filename  # local import: avoid a cycle at import time

            rank[f"{create_safe_filename(out_name)}_full.csv"] = s["index"]
    known = sorted((f for f in file_list if f in rank), key=lambda f: rank[f])
    unknown = sorted(f for f in file_list if f not in rank)
    return known + unknown


def _group_text(runs, col: bool) -> str:
    """Outline runs -> `12-20=L1, 25-30=L2` / `C:F=L1`; 'none' / 'unknown'."""
    if runs == "unknown":
        return "unknown"
    if not runs:
        return "none"
    return (_col_runs_text if col else _row_runs_text)(
        runs[:_MAX_LIST], fmt=lambda lvl: f"L{lvl}"
    ) + (f", (+{len(runs) - _MAX_LIST} more runs)" if len(runs) > _MAX_LIST else "")


def _hidden_text(indexes, groups, col: bool) -> str:
    """Hidden ranges, each marked `(grouped)` when it sits inside an outline
    group — the rubric prefers collapsed groups to plain hiding (check 93)."""
    if indexes == "unknown":
        return "unknown"
    if not indexes:
        return "none"
    grouped: set = set()
    if isinstance(groups, list):
        for g in groups:
            grouped.update(range(int(g["first"]), int(g["last"]) + 1))
    parts = []
    for r in _runs([(i, True) for i in indexes]):
        a, b = r["first"], r["last"]
        if col:
            a_t, b_t = get_column_letter(a), get_column_letter(b)
            label = a_t if a == b else f"{a_t}:{b_t}"
        else:
            label = str(a) if a == b else f"{a}-{b}"
        if all(i in grouped for i in range(a, b + 1)):
            label += " (grouped)"
        parts.append(label)
    return ", ".join(parts)


def _cf_rule_text(rule: dict) -> str:
    """`cellIs equal "MODEL OK" -> fill rgb:C6EFCE font rgb:9C0006 bold`."""
    bits = [rule.get("type") or "?"]
    if rule.get("operator"):
        bits.append(str(rule["operator"]))
    formulas = rule.get("formula") or []
    if formulas:
        bits.append(", ".join(str(f)[:40] for f in formulas[:2]))
    style = rule.get("style")
    if isinstance(style, dict) and style:
        st = []
        if style.get("fill"):
            st.append(f"fill {style['fill']}")
        if style.get("font"):
            st.append(f"font {style['font']}")
        if style.get("bold"):
            st.append("bold")
        bits.append("-> " + " ".join(st))
    return " ".join(bits)


_DV_OPERATORS = {
    "between": "between {a} and {b}", "notBetween": "not between {a} and {b}",
    "equal": "= {a}", "notEqual": "<> {a}", "greaterThan": "> {a}", "lessThan": "< {a}",
    "greaterThanOrEqual": ">= {a}", "lessThanOrEqual": "<= {a}",
}


def _data_validation_rule(d: dict) -> str:
    """'D7 whole between 1 and 50 — alert OFF (any entry accepted)'."""
    t = d.get("type") or "?"
    a, b = d.get("formula1"), d.get("formula2")
    op = d.get("operator")
    if t in ("list", "custom") or op is None and b is None:
        rule = f"{d['sqref']} {t}{' ' + str(a) if a else ''}"
    else:
        tmpl = _DV_OPERATORS.get(op or "between", "{a} {b}")
        rule = f"{d['sqref']} {t} " + tmpl.format(a=a, b=b if b is not None else "?")
    alert = d.get("alert")
    if alert is None:
        return rule + " — alert unknown"
    if not alert:
        return rule + " — alert OFF (any entry accepted)"
    style = d.get("error_style") or "stop"
    what = {"stop": "stop, entry rejected", "warning": "warning, entry allowed after a prompt",
            "information": "information, entry allowed"}.get(str(style), str(style))
    return rule + f" — alert ON ({what})"


def _data_validation_text(dvs) -> str:
    """Judge v11: per-sheet tally of the error-alert state, then every rule."""
    if dvs == "unknown":
        return "unknown"
    if not dvs:
        return "none"
    known = [d for d in dvs if d.get("alert") is not None]
    n, n_on = len(dvs), sum(1 for d in known if d.get("alert"))
    if len(known) < n:
        head = f"{n} (alert state unknown for {n - len(known)})"
    elif n_on == n:
        head = f"{n} (error alert ON for all)"
    elif n_on == 0:
        head = f"{n} (error alert OFF for ALL — none of them rejects or flags an entry)"
    else:
        head = f"{n} ({n_on} with the error alert ON, {n - n_on} OFF)"
    return head + ": " + _fmt_list(dvs, fn=_data_validation_rule)


def _fmt_list(items, limit=_MAX_LIST, fn=str) -> str:
    if items == "unknown":
        return "unknown"
    if not items:
        return "none"
    shown = [fn(x) for x in items[:limit]]
    more = len(items) - limit
    return "; ".join(shown) + (f"; (+{more} more)" if more > 0 else "")


EXCEL_MAX_COLUMNS = 16384
EXCEL_MAX_ROWS = 1048576


def oversized_range_tag(sheet_props: dict) -> str:
    """Evidence tag for rubric_9 check 25 (Reasonable file size): a sheet whose
    declared extent reaches Excel's full width or height while its content
    ends far earlier — the "oversized used range" the rubric names as bloat.

    Deliberately extreme-only (2026-09-14): a ratio rule (declared >= 10x
    content) flagged 59 of 471 cached real workbooks, golden solutions
    included, on harmless 300-column extents. Full-sheet extents flag 3, all
    one golden's sheet. Computed at render time from the stored properties,
    so no cache generation bump. Empty/unknown used range -> no tag.
    """
    ur = sheet_props.get("used_range")
    if not ur or ur == "unknown":
        return ""
    try:
        _, _, last_col, last_row = range_boundaries(ur)
    except Exception:  # noqa: BLE001
        return ""
    parts = []
    mc = sheet_props.get("max_column") or 0
    mr = sheet_props.get("max_row") or 0
    if mc >= EXCEL_MAX_COLUMNS and last_col:
        parts.append(
            f"declared {mc:,} columns, content ends at column {last_col} "
            f"({mc // last_col}x): OVERSIZED RANGE"
        )
    if mr >= EXCEL_MAX_ROWS and last_row:
        parts.append(
            f"declared {mr:,} rows, content ends at row {last_row} "
            f"({mr // last_row}x): OVERSIZED RANGE"
        )
    return (" — " + "; ".join(parts)) if parts else ""


_WIDE_OUTLIER_SKIP_SHEETS = re.compile(r"instruction|question|brief|readme", re.I)


def wide_outlier_tags(sheet_props: dict) -> list[str]:
    """Evidence for rubric_9 check 70 (Reasonable column widths), judge v9:
    a run of >= 2 consecutive equal-width columns at least WIDE_OUTLIER_RATIO
    times wider than the nearest equal-width run (length >= 2, not a spacer)
    on BOTH sides, and no longer than the longer neighbour run (a group
    inside a field, not the field and not an edge). Lone columns (a label or question column) are
    never compared, and the case's own Instructions/Questions sheets are
    skipped (guidance excludes them from check 70). Computed at render time
    from the stored widths. Returns
    lines like `BA:BD width 34.9 vs neighbours 12.9 (2.7x): WIDE OUTLIER`.
    """
    cw = sheet_props.get("column_widths")
    if not isinstance(cw, list) or not cw:
        return []
    if _WIDE_OUTLIER_SKIP_SHEETS.search(str(sheet_props.get("name") or "")):
        return []   # the case's own brief / questions sheet: excluded from check 70 by guidance
    ur = sheet_props.get("used_range")
    try:
        first_c, _, last_c, _ = range_boundaries(ur) if ur and ur != "unknown" else (None, None, None, None)
    except Exception:  # noqa: BLE001
        first_c = last_c = None
    default = sheet_props.get("default_col_width") or DEFAULT_COL_WIDTH
    try:
        default = float(default)
    except (TypeError, ValueError):
        default = DEFAULT_COL_WIDTH
    if not (first_c and last_c):
        return []   # no content, nothing to compare
    lo, hi = first_c, last_c
    widths: dict[int, float] = {}
    for run in cw:
        try:
            a, b, w = int(run["first"]), int(run["last"]), float(run["value"])
        except (KeyError, TypeError, ValueError):
            continue
        for c in range(max(a, lo), min(b, hi) + 1):
            widths[c] = w
    if not widths:
        return []
    hidden = set(sheet_props.get("hidden_cols") or []) if isinstance(sheet_props.get("hidden_cols"), list) else set()
    cols = [c for c in range(lo, hi + 1) if c not in hidden]
    if not cols:
        return []
    runs: list[dict] = []
    for c in cols:
        w = widths.get(c, default)
        if runs and runs[-1]["last"] == c - 1 and abs(runs[-1]["w"] - w) < 0.01:
            runs[-1]["last"] = c
        else:
            runs.append({"first": c, "last": c, "w": w})
    def _neighbour(idx: int, step: int):
        j = idx + step
        while 0 <= j < len(runs):
            r = runs[j]
            if (r["last"] - r["first"] + 1) >= 2 and r["w"] >= WIDE_OUTLIER_MIN_NEIGHBOUR:
                return r
            j += step
        return None
    tags = []
    for i, r in enumerate(runs):
        if (r["last"] - r["first"] + 1) < 2 or r["w"] < WIDE_OUTLIER_MIN_NEIGHBOUR * WIDE_OUTLIER_RATIO:
            continue
        left, right = _neighbour(i, -1), _neighbour(i, +1)
        nbrs = [n for n in (left, right) if n is not None]
        if len(nbrs) < 2:
            # the group must sit INSIDE the field: a wide pair at the right
            # edge of the used range (long-text note columns) or beside the
            # sheet's margin columns is not an outlier (Telecom golden, toys
            # 69/70 Pass)
            continue
        if any(r["w"] < WIDE_OUTLIER_RATIO * n["w"] for n in nbrs):
            continue
        # an outlier is a GROUP inside a wider field of same-width columns:
        # the run must be no longer than its longest neighbour run. A model's
        # whole timeline block (G:DW at 16 beside A:F label columns at 5.8)
        # is the field itself, not an outlier — 10 goldens have that shape.
        span = r["last"] - r["first"] + 1
        if span > max(n["last"] - n["first"] + 1 for n in nbrs):
            continue
        nw = sorted({round(n["w"], 1) for n in nbrs})
        ratio = r["w"] / max(n["w"] for n in nbrs)
        a, b = get_column_letter(r["first"]), get_column_letter(r["last"])
        tags.append(
            f"{a}:{b} width {r['w']:.1f} vs neighbours {'/'.join(f'{x:.1f}' for x in nw)} "
            f"({ratio:.1f}x): WIDE OUTLIER"
        )
    return tags


def _period_series_lines(ps) -> list[str]:
    if ps == "unknown" or not isinstance(ps, dict):
        return ["     period series: unknown"]
    lines = []
    h = ps.get("horizontal") or []
    if not h:
        lines.append("     period series: none detected")
    else:
        shown = h[:4]
        desc = "; ".join(f"{r['range']} {r['first']} → {r['last']}" for r in shown)
        more = f"; (+{len(h) - len(shown)} more)" if len(h) > len(shown) else ""
        lines.append(f"     period series: {len(h)} horizontal ({desc}{more})")
    for rec in (ps.get("out_of_order") or [])[:10]:
        lines.append(f"     PERIOD SERIES OUT OF ORDER: {rec['range']} {', '.join(rec.get('labels', []))}")
    for rec in (ps.get("gaps") or [])[:10]:
        lines.append(f"     unlabeled gap at {', '.join(rec['cells'])} (period header {rec['range']} continues past it over value-bearing columns)")
    for rec in (ps.get("vertical") or [])[:10]:
        lines.append(
            f"     VERTICAL PERIOD SERIES: {rec['range']} {rec['first']} → {rec['last']} down rows "
            f"({rec['formulas_beside']} cells beside are formulas) on a sheet whose main timeline runs across columns"
            + (" — OUT OF ORDER" if rec.get("out_of_order") else "")
        )
    em = ps.get("end_marker")
    if isinstance(em, dict) and em.get("rows_below"):
        lines.append(
            f"     content continues {em['rows_below']} rows past the \"{em['text']}\" marker at {em['cell']} "
            f"(rows {em['first_row_below']}-{em['last_row_below']})"
        )
    return lines


def _column_fit_lines(cf) -> list[str]:
    """Judge v10: only the NUMERIC ### class is rendered (see column_fit_summary)."""
    if cf == "unknown" or not isinstance(cf, dict):
        return ["     content fit: unknown"]
    num = cf.get("numeric_overflow") or {}
    n_num = int(num.get("count", 0) or 0)
    if not n_num:
        return ["     content fit: no numeric value exceeds its column width (nothing would render ###)"]
    ex = num.get("examples") or []
    return ["     content fit: "
            f"NUMERIC exceeds width (would render ###): {', '.join(e['ref'] for e in ex[:8])}"
            f"{', ...' if n_num > 8 else ''} ({n_num} cell{'s' if n_num != 1 else ''}; e.g. {ex[0]['ref']} needs ~{ex[0]['need']} chars in width {ex[0]['width']})"]


def render_properties_text(
    props: Optional[dict],
    listed_files: Optional[set] = None,
    origin: Optional[dict] = None,
    retired_checks=None,
) -> str:
    """Compact deterministic text for the seed prompt.

    `listed_files` (the `*_full.csv` names actually served) marks sheets
    whose CSV was dropped (ignored/filtered) so the judge is not sent
    looking for a file that is not there. `origin` (attempt only) is the
    provenance sidecar from setup_task_folder: the attempt is staged as
    ai_attempt.xlsx, so the delivered filename/extension (check 77) is
    only known from it. `retired_checks` (the grading's retired check
    numbers) leaves out evidence that exists for a retired check alone
    (STYLED_EMPTY_CHECK); omitted, everything renders.
    """
    if not props:
        return "  (workbook properties not available — older extraction cache)"
    retired = set(retired_checks or ())
    wb = props.get("workbook", {})
    lines = []
    size = wb.get("bytes")
    size_txt = f"{size / 1024:.0f} KB" if isinstance(size, (int, float)) else "unknown size"
    calc_mode = wb.get("calc_mode")
    original = (origin or {}).get("original_filename")
    lines.append(
        f"Workbook {wb.get('filename', '?')} ({size_txt})"
        f"{f'; original filename: {original}' if original else ''}; calc mode: "
        f"{calc_mode if calc_mode and calc_mode != 'unknown' else 'auto (Excel default, none set)'}"
        f"{' (full calc on load)' if wb.get('full_calc_on_load') else ''}; iterative calc: "
        f"{'on' if wb.get('iterative_calc') else ('off' if wb.get('iterative_calc') is not None else 'unknown')}; "
        f"VBA: {'yes' if wb.get('has_vba') else ('no' if wb.get('has_vba') is not None else 'unknown')}; "
        f"active sheet: {wb.get('active_sheet') or 'unknown'}"
    )
    dn = wb.get("defined_names")
    if isinstance(dn, list):
        # Hidden names (add-ins such as @RISK plant dozens per file) are
        # counted, not listed, so the model's own names keep the slots
        # (rubric_9 checks 9/29: hidden add-in names do not count against
        # the Name Manager). The JSON keeps every name.
        user_names = [d for d in dn if not str(d.get("name", "")).startswith(_SYSTEM_NAME_PREFIXES)]
        n_sys = len(dn) - len(user_names)
        visible = [d for d in user_names if not d.get("hidden")]
        n_hidden = len(user_names) - len(visible)
        footnote = [f"+{n_sys} add-in/system"] * bool(n_sys) + [f"+{n_hidden} hidden"] * bool(n_hidden)
        lines.append(
            "Defined names: " + _fmt_list(
                visible,
                fn=lambda d: (
                    f"{d['name']}{' (' + d['scope'] + ')' if d.get('scope') else ''}"
                    f" -> {d.get('refers_to')}"
                ),
            )
            + (f" [{', '.join(footnote)} names not listed]" if footnote else "")
        )
    else:
        lines.append("Defined names: unknown")
    lines.append("External links: " + _fmt_list(wb.get("external_links")))
    lines.append("Sheets in TRUE TAB ORDER (index. name [state]):")
    from .excel_utils import create_safe_filename

    for s in props.get("sheets", []):
        state = s.get("state", "visible")
        tag = f" [{state.upper()}]" if state and state != "visible" else ""
        if s.get("kind") == "chartsheet":
            lines.append(f"  {s['index']}. {s['name']}{tag} — chart sheet (no cell data)")
            continue
        out = s.get("output_name")
        served = ""
        if out is None:
            served = " — not served (filtered out)"
        elif listed_files is not None and f"{create_safe_filename(out)}_full.csv" not in listed_files:
            served = " — not served (ignored sheet)"
        elif out != s["name"]:
            served = f" — served as {out}_full.csv"
        head = (
            f"  {s['index']}. {s['name']}{tag}{served}: "
            f"{s.get('max_row')}x{s.get('max_column')} (used {s.get('used_range') or 'empty'}; "
            f"{s.get('n_values', '?')} values, {s.get('n_formulas', '?')} formulas"
            f"{', ' + str(s['n_spill_anchors']) + ' spill/array ranges' if s.get('n_spill_anchors') else ''})"
            f"{oversized_range_tag(s)}"
        )
        lines.append(head)
        fp = s.get("freeze_panes")
        detail = [
            f"freeze panes: {freeze_text(s)}",
            f"gridlines: {'on' if s.get('gridlines') in (None, True) else 'off'}",
            f"zoom: {s.get('zoom') or 100}",
            f"tab color: {s.get('tab_color') or 'none'}",
            f"protected: {'yes' if s.get('protected') else 'no'}",
            f"merged ranges: {len(s.get('merged_ranges') or []) if s.get('merged_ranges') != 'unknown' else 'unknown'}",
            "active cell: " + str(s.get("active_cell") or "unknown")
            + (f" (opens scrolled to {s['top_left_cell']})"
               if s.get("top_left_cell") and s.get("top_left_cell") != "A1" else ""),
        ]
        lines.append("     " + "; ".join(detail))
        if STYLED_EMPTY_CHECK not in retired:
            se = s.get("styled_empty_cells")
            if isinstance(se, dict):
                n_se = int(se.get("count", 0) or 0)
                ex = list(se.get("examples") or [])
                se_txt = "none" if n_se == 0 else (
                    f"{n_se} (e.g. {', '.join(ex)}{', ...' if n_se > len(ex) else ''})"
                )
            else:
                se_txt = "unknown" if se == "unknown" else "none"
            lines.append("     styled empty cells in used range: " + se_txt)
        hr, hc = s.get("hidden_rows", []), s.get("hidden_cols", [])
        rg, cg = s.get("row_groups", []), s.get("col_groups", [])
        lines.append(
            "     hidden rows: " + _hidden_text(hr, rg, col=False)
            + "; hidden cols: " + _hidden_text(hc, cg, col=True)
        )
        lines.append(
            "     grouped rows: " + _group_text(rg, col=False)
            + "; grouped cols: " + _group_text(cg, col=True)
        )
        rb, cb = s.get("row_breaks", []), s.get("col_breaks", [])
        lines.append(
            "     page breaks: rows " + ("unknown" if rb == "unknown" else (", ".join(str(b) for b in rb) if rb else "none"))
            + "; cols " + ("unknown" if cb == "unknown" else (", ".join(get_column_letter(b) for b in cb) if cb else "none"))
        )
        lines.extend(_print_estimate_lines(s))
        cw = s.get("column_widths")
        dcw = s.get("default_col_width")
        lines.append(
            "     column widths: "
            + (_col_runs_text(cw[:_MAX_LIST]) + (f", (+{len(cw) - _MAX_LIST} more runs)" if len(cw) > _MAX_LIST else "")
               if isinstance(cw, list) and cw else ("unknown" if cw == "unknown" else "all default"))
            + f" (default {dcw if dcw else 8.43})"
        )
        for tag in wide_outlier_tags(s):
            lines.append("     column width outlier: " + tag)
        lines.extend(_column_fit_lines(s.get("column_fit", "unknown")))
        lines.extend(_period_series_lines(s.get("period_series", "unknown")))
        dl = s.get("date_literal_formulas")
        if isinstance(dl, list) and dl:
            lines.append(
                "     formulas with a typed date-like string literal: "
                + ", ".join(dl[:_MAX_LIST]) + (f", (+{len(dl) - _MAX_LIST} more)" if len(dl) > _MAX_LIST else "")
            )
        lines.extend(_ii.render_lines(s.get("implicit_intersection")))
        lines.extend(_rounding_lines(s.get("rounding_statements")))
        rh = s.get("row_heights")
        lines.append(
            "     custom row heights: "
            + (_row_runs_text(rh[:_MAX_LIST]) + (f", (+{len(rh) - _MAX_LIST} more runs)" if len(rh) > _MAX_LIST else "")
               if isinstance(rh, list) and rh else ("unknown" if rh == "unknown" else "none"))
        )
        lines.append("     data validation: " + _data_validation_text(s.get("data_validations")))
        lines.append(
            "     conditional formats: " + _fmt_list(
                s.get("conditional_formats"),
                fn=lambda d: f"{d['sqref']} ({'; '.join(_cf_rule_text(r) for r in d.get('rules', []))})",
            )
        )
        lines.append(
            "     comments/notes: " + _fmt_list(
                s.get("comments"),
                fn=lambda d: f"{d['ref']}: \"{(d.get('text') or '').strip()[:80]}\"",
            )
        )
        lines.append("     hyperlinks: " + _fmt_list(s.get("hyperlinks"), fn=lambda d: f"{d['ref']} -> {d.get('target')}"))
        lines.append(
            f"     print: area {s.get('print_area') or 'none'}; title rows {s.get('print_title_rows') or 'none'}; "
            f"orientation {s.get('page_orientation') or 'default'}; fit to page: "
            f"{'yes' if s.get('fit_to_page') else 'no'}"
        )
    return "\n".join(lines)
