"""61 Formatting/Consistent zoom level.

Rule as implemented (see docs/checks/61.md):
fail when the sheets of the delivered workbook do not all open at the same zoom level
(House Standards: "One zoom level throughout"; rubric: "Every sheet should use the same
zoom level across the entire workbook").  Whole workbook, no starting-file diff, no
tolerance: 79% next to 80% is two levels.

* A sheet's zoom is the zoom it OPENS at: <sheetView zoomScale> of the sheet's view for
  workbook window 0 (detchecks SheetHead.view).  zoomScale is the zoom of the CURRENT view
  (ECMA-376 18.3.1.87), whatever the view mode (normal, page layout, page break preview).
  A missing zoomScale is 100 (schema default; toy 61 T3), and so is zoomScale="0" (core
  SheetView.zoom).  A sheet with no <sheetView> opens at 100.
* zoomScaleNormal / zoomScalePageLayoutView / zoomScaleSheetLayoutView are per-mode
  memories Excel uses only when the user switches view mode; Excel ignores them for the
  opening zoom (toy Overview 61 T1/T3: a sheet with only zoomScaleNormal=80 shows 100%).
  They are recorded in stats ("per-view" reading) but never decide (ZOOM_READING).
* Counted sheets: worksheets, dialog sheets and macro sheets (sheets with a cell grid),
  hidden and very hidden ones included (whole workbook).  Chart sheets have no grid
  (their zoom sizes a chart page, often zoomToFit): excluded, listed in stats
  (COUNTED_KINDS).
* Mistakes: the reference level is the level most counted sheets open at (ties: the level
  of the earliest such sheet in tab order); one mistake per sheet at another level,
  location = the bare sheet name.

Unknown zoom: a counted sheet whose zoomScale is not an integer, or an integer outside
10..400 other than 0, has an unknown opening zoom.  So has a sheet whose <sheetViews>
comes AFTER <sheetData> (schema-invalid element order: Excel does not open such a file
as-is, it offers a repair; review 2026-10-04).  Such a sheet never gets a mistake; it is
listed in stats.undecidable_sheets.  When the readable sheets already differ (spread > 2 x
ZOOM_TOLERANCE, so no value of the unknown zoom can make them one level) the verdict is
fail; when it is the only counted sheet the verdict is pass (one sheet cannot differ from
itself); otherwise the unknown zoom decides and the check raises GradingError (no guess).
A sheet whose part cannot be read raises in the engine.
"""
from __future__ import annotations

import re
from collections import Counter

from ..errors import GradingError
from .base import Check

# Policy switches (each a question for Patrick in docs/checks/61.md)
ZOOM_READING = "zoomScale"            # "zoomScale" (Excel's opening zoom) | "per_view" (zoomScaleNormal etc. first)
ZOOM_TOLERANCE = 0                    # points; 0 = levels must be identical ("one zoom level")
COUNTED_KINDS = ("worksheet", "dialogsheet", "macrosheet")   # chart sheets excluded (no grid)
INCLUDE_HIDDEN_SHEETS = True          # whole workbook: hidden / veryHidden sheets count
DEFAULT_ZOOM = 100
VALID_RANGE = (10, 400)               # ECMA-376: zoomScale is restricted to 10..400

PER_VIEW_ATTR = {"normal": "zoomScaleNormal", "pageLayout": "zoomScalePageLayoutView",
                 "pageBreakPreview": "zoomScaleSheetLayoutView"}
_INT = re.compile(r"^\s*\+?\d+\s*$")


def parse_zoom(raw, attr: str, sheet: str):
    """Stored zoom attribute -> int, None when absent.  0 -> None (treated as absent, as
    the core's SheetView.zoom does).  Non-integers and values outside 10..400 raise."""
    if raw is None:
        return None
    if not _INT.match(str(raw)):
        raise GradingError(f"sheet {sheet!r}: {attr}={raw!r} is not an integer zoom; "
                           f"the zoom Excel opens it at is unknown")
    z = int(raw)
    if z == 0:
        return None
    if not VALID_RANGE[0] <= z <= VALID_RANGE[1]:
        raise GradingError(f"sheet {sheet!r}: {attr}={z} is outside the valid zoom range "
                           f"{VALID_RANGE[0]}..{VALID_RANGE[1]}; the zoom Excel opens it at is unknown")
    return z


def _read(raw, attr: str, sheet: str, strict: bool, problems: list):
    """parse_zoom; when not strict (a stats-only reading) a bad value is recorded, not raised."""
    if strict:
        return parse_zoom(raw, attr, sheet)
    try:
        return parse_zoom(raw, attr, sheet)
    except GradingError as e:
        problems.append(str(e))
        return None


def view_zooms(view, sheet: str, decide: str | None = ZOOM_READING) -> dict:
    """{'zoomScale': opening zoom, 'per_view': per-mode reading, 'explicit': bool, 'mode': view mode,
        'attrs': stored zoom attributes, 'problems': [..]} for a SheetView (None = no sheetView).
    Only the reading named by `decide` raises on an unreadable value (it decides); the other
    reading is stats only and records the problem instead.  decide=None never raises."""
    if view is None:
        return {"zoomScale": DEFAULT_ZOOM, "per_view": DEFAULT_ZOOM, "explicit": False, "mode": "normal",
                "attrs": {}, "problems": []}
    a = view.attrs
    problems: list[str] = []
    zs = _read(a.get("zoomScale"), "zoomScale", sheet, decide == "zoomScale", problems)
    mode = a.get("view") or "normal"
    pv_attr = PER_VIEW_ATTR.get(mode)
    if pv_attr is None:
        if decide == "per_view":
            raise GradingError(f"sheet {sheet!r}: unknown sheetView view={mode!r}")
        problems.append(f"unknown view mode {mode!r}")
        pv = None
    else:
        pv = _read(a.get(pv_attr), pv_attr, sheet, decide == "per_view", problems)
    opening = zs if zs is not None else DEFAULT_ZOOM
    stored = {k: a[k] for k in ("zoomScale", "zoomScaleNormal", "zoomScalePageLayoutView",
                                "zoomScaleSheetLayoutView", "zoomToFit") if a.get(k) is not None}
    return {"zoomScale": opening, "per_view": pv if pv is not None else opening,
            "explicit": zs is not None, "mode": mode, "attrs": stored, "problems": problems}


def reference_level(rows: list[dict], key: str) -> int:
    """Most common level among rows; ties -> the level of the earliest row (tab order)."""
    cnt = Counter(r[key] for r in rows)
    first = {}
    for r in rows:
        first.setdefault(r[key], r["index"])
    return max(cnt, key=lambda z: (cnt[z], -first[z]))


def _plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


class C61(Check):
    number = 61
    key = "Formatting/Consistent zoom level"
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet", "chartsheet")

    def start(self, wb):
        super().start(wb)
        self.rows: list[dict] = []
        self.excluded: list[dict] = []
        self.extra_windows: list[str] = []
        self.problems: dict[str, list] = {}
        self.undecidable: dict[str, str] = {}
        self.late_views: list[str] = []                       # sheets whose <sheetViews> follows <sheetData>
        self._views_at_start: dict[int, tuple] = {}           # sheet index -> (head.views object, copy)

    def sheet_start(self, head):
        self._views_at_start[head.index] = (head.views, list(head.views))
        counted = head.kind in COUNTED_KINDS and (head.state == "visible" or INCLUDE_HIDDEN_SHEETS)
        if len(head.views) > 1:
            self.extra_windows.append(head.name)
        try:
            z = view_zooms(head.view, head.name, ZOOM_READING if counted else None)
        except GradingError as e:        # this sheet's opening zoom is unknown: decided in finish()
            self.undecidable[head.name] = str(e)
            return False
        if z["problems"]:
            self.problems[head.name] = z["problems"]
        row = {"index": head.index, "name": head.name, "kind": head.kind, "state": head.state,
               "zoomScale": z["zoomScale"], "per_view": z["per_view"], "explicit": z["explicit"],
               "mode": z["mode"], "attrs": z["attrs"], "has_view": head.view is not None}
        if counted:
            self.rows.append(row)
        else:
            self.excluded.append({"sheet": head.name, "kind": head.kind, "state": head.state,
                                  "zoom": row["zoomScale"], "attrs": row["attrs"]})
        return False                     # the head is all this check needs: sheet data is skipped

    def sheet_end(self, head, tail):
        """Between sheet_start and sheet_end the engine parses the tail (everything after
        </sheetData>) and applies every top-level element it knows to the head, a <sheetViews>
        included.  When head.views is no longer what sheet_start saw, the sheet carried a
        <sheetViews> AFTER <sheetData>: schema-invalid element order, which Excel does not open
        as-is (it offers a repair), so the zoom the sheet opens at is unknown.  Before the review
        of 2026-10-04 such a sheet was silently read as having no view (100%)."""
        snap = self._views_at_start.pop(head.index, None)
        if snap is None:
            return
        obj, copy = snap
        if head.views is obj and head.views == copy:
            return
        self.late_views.append(head.name)
        row = next((r for r in self.rows if r["index"] == head.index), None)
        if row is None:                  # excluded (chart sheet) or already unknown for another reason: stats only
            return
        self.rows.remove(row)
        self.undecidable[head.name] = (f"sheet {head.name!r}: its <sheetViews> follows <sheetData> (schema-invalid "
                                       f"element order; Excel does not open such a file as-is); the zoom Excel "
                                       f"opens it at is unknown")

    # ------------------------------------------------------------------ decision
    def _deviants(self, key: str):
        if not self.rows:
            return None, []
        ref = reference_level(self.rows, key)
        return ref, [r for r in self.rows if abs(r[key] - ref) > ZOOM_TOLERANCE]

    def _certain_fail(self, key: str) -> bool:
        """The readable sheets differ by more than 2 x ZOOM_TOLERANCE: no reference level can
        cover them all, whatever the unknown zooms are."""
        levels = [r[key] for r in self.rows]
        return bool(levels) and max(levels) - min(levels) > 2 * ZOOM_TOLERANCE

    def _why(self, r: dict, use: str = "zoomScale") -> str:
        a = r["attrs"]
        if use == "per_view":            # non-default reading: the per-mode attribute decided
            how = ", ".join(f"{k}={v}" for k, v in a.items()) or "no zoom attributes"
            how += f" (per-view reading, {PER_VIEW_ATTR.get(r['mode'], 'zoomScale')} first)"
        elif not r["has_view"]:
            how = "it has no sheet view, so Excel opens it at the default 100%"
        elif not r["explicit"]:
            how = "it stores no zoomScale, so Excel opens it at the default 100%"
            if a.get("zoomScaleNormal") is not None and str(a["zoomScaleNormal"]).strip() != str(r["zoomScale"]):
                how += f" (zoomScaleNormal={a['zoomScaleNormal']} is not the opening zoom)"
        else:
            how = f"zoomScale={a.get('zoomScale')}"
            zn = a.get("zoomScaleNormal")
            if zn is not None and str(zn).strip() != str(r["zoomScale"]):
                how += f"; zoomScaleNormal={zn} is not the opening zoom"
        if r["mode"] != "normal":
            how += f"; {r['mode']} view"
        return how

    def finish(self) -> dict:
        if self.wb is None:
            raise GradingError(f"{self.key}: no workbook")
        if not self.rows and not self.excluded and not self.undecidable:
            raise GradingError(f"{self.key}: the workbook has no sheets")
        use = "zoomScale" if ZOOM_READING == "zoomScale" else "per_view"
        if ZOOM_READING not in ("zoomScale", "per_view"):
            raise GradingError(f"{self.key}: unknown ZOOM_READING {ZOOM_READING!r}")
        single = len(self.rows) + len(self.undecidable) <= 1     # one counted sheet cannot differ from itself
        if self.undecidable and not single and not self._certain_fail(use):
            # the readable sheets agree (or there are none), so the unknown zoom(s) decide: no guess
            raise GradingError(f"{self.key}: the opening zoom of {len(self.undecidable)} sheet(s) is unknown and "
                               f"the readable sheets do not already differ: " + " | ".join(self.undecidable.values()))
        ref, bad = self._deviants(use)
        other = "per_view" if use == "zoomScale" else "zoomScale"
        ref_o, bad_o = self._deviants(other)
        n = len(self.rows)
        at_ref = [r for r in self.rows if r not in bad]
        for r in bad:
            hidden = "" if r["state"] == "visible" else f" ({r['state']} sheet)"
            self.add_mistake(r["name"],
                             f"Sheet '{r['name']}'{hidden} opens at {r[use]}% zoom ({self._why(r, use)}) while "
                             f"{len(at_ref)} of {n} sheets open at {ref}%; use one zoom level throughout.",
                             severity="minor")
        levels: dict[int, list[str]] = {}
        for r in self.rows:
            levels.setdefault(r[use], []).append(r["name"])
        stats = {
            "zoom_reading": ZOOM_READING, "tolerance": ZOOM_TOLERANCE,
            "n_sheets_counted": n,
            "zooms": {r["name"]: r[use] for r in self.rows},
            "levels": {str(z): names for z, names in sorted(levels.items())},
            "reference_zoom": ref,
            "sheets_without_zoomScale": [r["name"] for r in self.rows if not r["explicit"]],
            "zoomScaleNormal_differs": {r["name"]: r["attrs"]["zoomScaleNormal"] for r in self.rows
                                        if r["attrs"].get("zoomScaleNormal") is not None
                                        and str(r["attrs"]["zoomScaleNormal"]).strip() != str(r["zoomScale"])},
            "non_normal_views": {r["name"]: r["mode"] for r in self.rows if r["mode"] != "normal"},
            "hidden_sheets_counted": [r["name"] for r in self.rows if r["state"] != "visible"],
            "excluded_sheets": self.excluded,
            "sheets_with_extra_window_views": self.extra_windows,
            "unreadable_stats_only_attributes": self.problems,
            "undecidable_sheets": self.undecidable,
            "sheet_views_after_sheet_data": self.late_views,
            "other_reading": {"reading": other, "reference_zoom": ref_o,
                              "deviating_sheets": [r["name"] for r in bad_o],
                              "decision": "pass" if not bad_o else "fail"},
            "reading_sensitive": bool(bad) != bool(bad_o),
        }
        if n == 0 and self.undecidable:          # the only counted sheet has an unknown zoom
            only = next(iter(self.undecidable))
            return self.verdict(f"Only one sheet with a cell grid ('{only}', zoom unknown: stats.undecidable_sheets); "
                                f"one sheet cannot differ from itself.", "unreachable", stats)
        if n == 0:
            names = ", ".join(f"'{e['sheet']}' ({e['kind']})" for e in self.excluded)
            return self.verdict(f"No sheet with a cell grid to compare (only {names}); nothing can differ.",
                                "unreachable", stats)
        desc = "; ".join(f"{z}%: " + (", ".join(f"'{x}'" for x in names) if len(names) <= 3
                                      else _plural(len(names), "sheet"))
                         for z, names in sorted(levels.items(), key=lambda kv: (-len(kv[1]), kv[0])))
        undec = ""
        if self.undecidable:
            names = ", ".join(f"'{x}'" for x in list(self.undecidable)[:5]) \
                + (f" and {len(self.undecidable) - 5} more" if len(self.undecidable) > 5 else "")
            undec = (f" The opening zoom of {len(self.undecidable)} further sheet(s) ({names}) is unknown "
                     f"(stats.undecidable_sheets); the verdict does not depend on it.")
        same = (f"All {_plural(n, 'sheet')} open at {ref}% zoom." if len(levels) == 1 else
                f"All {_plural(n, 'sheet')} open within {ZOOM_TOLERANCE} points of {ref}% zoom "
                f"({min(levels)}-{max(levels)}%; ZOOM_TOLERANCE={ZOOM_TOLERANCE}).")
        return self.verdict(same + undec,
                            f"Zoom varies across the workbook ({desc}); {{n}} sheet(s) differ from the "
                            f"common {ref}%." + undec, stats)
