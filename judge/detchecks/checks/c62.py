"""62 Formatting/Active cell reset to A1 on all sheets.

Rule as implemented (Patrick's ruling 2026-10-02: STRICT A1; see docs/checks/62.md):
fail on every sheet whose cursor (active cell) is not A1 when the workbook opens.  The
frozen-pane-corner / "designated home cell" exception is dropped: a frozen sheet whose
cursor sits on the first unfrozen cell (A4, F4, ...) fails too.

* The cursor is read from the sheet's view for workbook window 0 (SheetHead.view): the
  <selection> of the view's ACTIVE pane (<pane activePane>, topLeft without a pane), NOT
  simply the first <selection> (toys 62 T2 'Utilities' and T3 'Solution Model' carry a
  stale first selection of another pane).
* No sheetView, or no selection at all on a view without panes: the cursor is A1 (Excel's
  default; Excel omits an A1 selection, even on a scrolled view: toy 94 T4 'NXTI IS').
* A view with panes whose ACTIVE pane (not topLeft) has no selection (another pane may have
  one): the cursor is that pane's top-left cell (pane topLeftCell, else the frozen corner),
  as Excel does (measured by Patrick 2026-10-03, attempt 1408: frozen sheets whose only
  selection is A1 in the topLeft pane open on D7 / F10, the first scrolling cell;
  ACTIVE_PANE_WITHOUT_SELECTION = "pane_top_left").  Such sheets are listed in
  stats.active_pane_without_selection.  The measurement covered panes parked AT the frozen
  corner only: when the scrolling pane is parked beyond it (GPT-6 1353: pane topLeftCell R6,
  corner F6) the pane's topLeftCell stays the primary reading and the corner is kept as a
  second candidate (both fail strict A1; re-review 2026-10-04, `62-scrolled-pane-location`).
* A view with panes whose active pane is topLeft and has no topLeft selection: A1 (Excel's
  default) when the view's topLeftCell is A1; when the view is scrolled (topLeftCell C1), A1
  and C1 (the pane's top-left cell, by the 1408 reading) are both candidates: not measured
  (re-review 2026-10-04, `62-topleft-active-ignores-view-scroll`; 0 corpus sheets).
* A selection without activeCell: A1 when its sqref is absent or starts at A1.
* A missing sqref is the schema default "A1" (ECMA-376 CT_Selection), so
  <selection activeCell="D10"/> is the same shape as <selection activeCell="D10" sqref="A1"/>.
* A view WITHOUT <pane> whose selections carry a pane label other than topLeft (stray labels:
  openpyxl keeps them after `ws.freeze_panes = None`): every such selection's cell is a
  candidate next to the topLeft reading (default A1), because Excel's resolution is not
  measured (re-review 2026-10-04, `62-stray-pane-label-silent-a1`; 0 corpus sheets).  Listed
  in stats.selections_labelled_without_pane.
* Several <selection> elements for the active pane: the LAST one decides (Patrick, in Excel,
  2026-10-06: attempt 2379 'Income Statement' stores bottomRight A1 then B9 and opens on B9;
  attempt 2402 'Balance Sheet' stores A1 then A4 and opens on A4).  The earlier ones are
  noted (stats.notes "N selections for the active pane"), not candidates.  The shape comes
  from openpyxl: `ws.freeze_panes` appends the pane's selection after one the agent set.
* Shapes whose cursor Excel resolves in an unverified way (a selection without activeCell
  whose sqref starts elsewhere; an activeCell outside its own sqref, explicit or default;
  the three shapes above) keep every plausible cursor as a candidate.  All candidates A1 ->
  the sheet passes; none A1 -> it fails (location = the primary reading); mixed -> the
  sheet is undecidable.
* Counted sheets: worksheets, dialog sheets and macro sheets, hidden and very hidden ones
  included (whole workbook; INCLUDE_HIDDEN_SHEETS).  Chart sheets have no cell cursor:
  listed in stats only.
* Scroll position (sheetView/pane topLeftCell) is not judged (CHECK_SCROLL = False, a
  question for Patrick); views that open scrolled away from A1 are listed in stats.

Decision: fail when at least one counted sheet fails for certain (undecidable sheets are
then listed in stats.undecidable_sheets, never in mistakes); otherwise GradingError when a
sheet is undecidable (an unreadable activeCell/sqref/activePane on the selection that
decides, or mixed candidates); otherwise pass.  The verdict therefore raises only when an
unknown cursor could change it (no guess).
Mistakes: one per failing sheet, location "'Sheet'!<cursor cell>".
"""
from __future__ import annotations

from ..core.refs import in_bounds, location, make_ref, parse_range, split_ref
from ..errors import GradingError
from .base import Check

# Policy switches (each a question for Patrick in docs/checks/62.md)
STRICT_A1 = True                      # Patrick 2026-10-02: no frozen-corner / home-cell exception
INCLUDE_HIDDEN_SHEETS = True          # whole workbook: hidden / veryHidden sheets count
COUNTED_KINDS = ("worksheet", "dialogsheet", "macrosheet")   # chart sheets have no cursor
CHECK_SCROLL = False                  # a view scrolled away from A1 is stats only
ACTIVE_PANE_WITHOUT_SELECTION = "pane_top_left"   # Excel 2026-10-03: the active pane's top-left cell.
                                      # Old readings kept for comparison: "A1" | "unverified" (both candidates)

PANES = ("topLeft", "topRight", "bottomLeft", "bottomRight")
A1 = (1, 1)
DEFAULT_SQREF = "A1"                  # ECMA-376 CT_Selection/@sqref default


def _cell(text: str, what: str, sheet: str) -> tuple[int, int]:
    rc = split_ref(text.strip()) if text else None
    if rc is None or not (1 <= rc[0] <= 1048576 and 1 <= rc[1] <= 16384):
        raise GradingError(f"sheet {sheet!r}: {what} {text!r} is not a cell reference; "
                           f"the cursor position cannot be read")
    return rc


def _sqref(text, sheet: str) -> list[tuple[int, int, int, int]]:
    out = []
    for part in (text or "").split():
        b = parse_range(part)
        if b is None or part.count(":") > 1 or b[2] > 1048576 or b[3] > 16384 or b[0] < 1 or b[1] < 1:
            raise GradingError(f"sheet {sheet!r}: selection sqref {text!r} is unreadable; "
                               f"the cursor position cannot be read")
        out.append(b)
    return out


def _view_top_left(view) -> tuple[int, int]:
    return split_ref((view.top_left_cell or "A1").strip()) or A1


def _selection_cells(s, sheet: str) -> tuple[list[tuple[int, int]], list[str]]:
    """Plausible cursor cells of one <selection> (first = the primary reading) and notes on
    its odd shapes.  Raises GradingError when its activeCell or sqref is unreadable."""
    sq = _sqref(s.sqref, sheet)
    if s.active_cell:
        ac = _cell(s.active_cell, "activeCell", sheet)
        eff = sq or _sqref(DEFAULT_SQREF, sheet)              # absent sqref = schema default A1
        if any(in_bounds(ac[0], ac[1], b) for b in eff):
            return [ac], []
        shown = s.sqref if sq else f"{DEFAULT_SQREF} (sqref absent: schema default)"
        return [ac, eff[0][:2]], [f"activeCell {s.active_cell} outside its selection {shown}"]
    if sq and sq[0][:2] != A1:
        return [A1, sq[0][:2]], [f"selection {s.sqref} without activeCell"]
    return [A1], []


def frozen_corner(view):
    """First unfrozen cell of a frozen view (where Ctrl+Home lands), else None."""
    p = view.pane if view is not None else None
    if p is None or p.state not in ("frozen", "frozenSplit") or (not p.x_split and not p.y_split):
        return None
    tl = _view_top_left(view)
    return (tl[0] + int(p.y_split), tl[1] + int(p.x_split))


def pane_corner(view):
    """Top-left cell of the view's ACTIVE pane when the scrolling panes sit at the frozen
    corner (the pane's top-left as derived from the splits alone), or None when there is no
    frozen corner (split panes, no splits) or the active pane is topLeft."""
    corner = frozen_corner(view)
    if corner is None:
        return None
    vtl = _view_top_left(view)
    return {"bottomRight": corner, "bottomLeft": (corner[0], vtl[1]), "topRight": (vtl[0], corner[1])}.get(view.active_pane)


def pane_top_left(view):
    """Top-left cell of the view's ACTIVE pane as the file stores it (the cell Excel shows
    first in that pane), or None when it cannot be derived."""
    if view is None:
        return A1
    vtl = _view_top_left(view)
    ap = view.active_pane
    if ap == "topLeft" or view.pane is None:
        return vtl
    ptl = split_ref(view.pane.top_left_cell.strip()) if view.pane.top_left_cell else None
    if ptl is None:
        return pane_corner(view)
    return {"bottomRight": ptl, "bottomLeft": (ptl[0], vtl[1]), "topRight": (vtl[0], ptl[1])}.get(ap)


def cursor(view, sheet: str) -> dict:
    """Where the sheet's cursor is when the workbook opens.

    Returns {'cands': [(r, c), ...] (deduplicated, first = the primary reading),
             'pane': active pane, 'selection': sqref of the deciding selection or None,
             'how': short explanation, 'notes': [unverified-shape notes],
             'first_selection': (pane, activeCell) of the FIRST <selection> or None,
             'no_active_selection': the active pane's top-left cell when the active pane of a
                                    view with panes has no selection and that cell is not A1,
             'stray_labels': "pane cell, ..." for selections labelled for another pane on a
                             view without <pane>, else None}
    Raises GradingError when the deciding selection data is unreadable."""
    if view is None:
        return {"cands": [A1], "pane": "topLeft", "selection": None, "how": "no sheet view (default A1)",
                "notes": [], "first_selection": None, "no_active_selection": None, "stray_labels": None}
    ap = view.active_pane
    if ap not in PANES:
        raise GradingError(f"sheet {sheet!r}: <pane activePane={ap!r}> is not a pane name; "
                           f"the cursor position cannot be read")
    if ACTIVE_PANE_WITHOUT_SELECTION not in ("A1", "unverified", "pane_top_left"):
        raise GradingError(f"unknown ACTIVE_PANE_WITHOUT_SELECTION {ACTIVE_PANE_WITHOUT_SELECTION!r}")
    first = (view.selections[0].pane, view.selections[0].active_cell) if view.selections else None
    matching = [s for s in view.selections if s.pane == ap]
    notes: list[str] = []
    cands: list[tuple[int, int]] = []
    alt = None
    if not matching:
        how = (f"no selection for the active pane {ap} (default A1)" if view.selections
               else "no selection (default A1)")
        cands = [A1]
        if view.pane is not None and ap != "topLeft":
            ptl = pane_top_left(view)
            if ptl is None or ptl != A1:
                alt = make_ref(*ptl) if ptl else "unknown"
                if ACTIVE_PANE_WITHOUT_SELECTION == "A1":
                    pass
                elif ptl is None:
                    raise GradingError(f"sheet {sheet!r}: the active pane {ap} has no selection and its "
                                       f"top-left cell cannot be derived; the cursor position cannot be read")
                elif ACTIVE_PANE_WITHOUT_SELECTION == "pane_top_left":
                    cands = [ptl]
                    how = (f"no selection for the active pane {ap}: Excel puts the cursor on the pane's "
                           f"top-left cell")
                else:                                             # "unverified"
                    cands.append(ptl)
                    notes.append(f"no selection for the active pane {ap} (A1 or the pane's top-left cell {alt})")
                # re-review 2026-10-04 (62-scrolled-pane-location): the Excel measurement (1408) had the
                # scrolling pane AT the frozen corner; a pane parked beyond it keeps the corner as a candidate
                corner = pane_corner(view)
                if ACTIVE_PANE_WITHOUT_SELECTION != "A1" and corner is not None and corner != ptl:
                    cands.append(corner)
                    notes.append(f"the scrolling pane is parked at {alt}, past the frozen corner "
                                 f"{make_ref(*corner)}; Excel was measured only with the pane at the corner, "
                                 f"so the cursor is {alt} or {make_ref(*corner)}")
        elif view.pane is not None:
            # re-review 2026-10-04 (62-topleft-active-ignores-view-scroll): the active topLeft pane of a
            # scrolled view has no selection: A1 (Excel's default) or the pane's top-left cell (1408 reading)
            vtl = _view_top_left(view)
            if vtl != A1 and ACTIVE_PANE_WITHOUT_SELECTION != "A1":
                alt = make_ref(*vtl)
                cands.append(vtl)
                notes.append(f"no selection for the active pane topLeft of a view scrolled to {alt} "
                             f"(A1 by Excel's default, or the pane's top-left cell {alt}; not measured)")
    else:
        if len(matching) > 1:
            # Patrick in Excel, 2026-10-06 (attempts 2379, 2402): the LAST stored selection of the active
            # pane is the one Excel opens on; the earlier ones are noted, not candidates
            notes.append(f"{len(matching)} selections for the active pane {ap}; Excel opens on the last one "
                         f"(measured 2026-10-06), earlier: "
                         + ", ".join(s.active_cell or s.sqref or DEFAULT_SQREF for s in matching[:-1]))
            matching = matching[-1:]
        cands, nn = _selection_cells(matching[0], sheet)
        notes += nn
        how = f"selection of the active pane {ap}"
    # re-review 2026-10-04 (62-stray-pane-label-silent-a1): selections labelled for a pane that does not
    # exist (no <pane>; openpyxl keeps the labels after ws.freeze_panes = None) are candidates too
    stray_labels = None
    if view.pane is None:
        stray = [s for s in view.selections if s.pane != "topLeft"]
        if stray:
            extra: list[tuple[int, int]] = []
            for s in stray:
                cc, nn = _selection_cells(s, sheet)
                extra += cc
                notes += nn
            stray_labels = ", ".join(f"{s.pane} {s.active_cell or s.sqref or DEFAULT_SQREF}" for s in stray)
            non_a1 = [make_ref(*rc) for rc in dict.fromkeys(extra) if rc != A1]
            if non_a1:
                cands += extra
                notes.append(f"selection(s) labelled {stray_labels} on a view without <pane> "
                             f"(stray pane labels; Excel's reading is not measured)")
    uniq = list(dict.fromkeys(cands))
    return {"cands": uniq, "pane": ap, "selection": matching[0].sqref if matching else None, "how": how,
            "notes": notes, "first_selection": first, "no_active_selection": alt, "stray_labels": stray_labels}


def scroll_state(view) -> dict | None:
    """Stats only: where the view opens scrolled to when that is not the top-left of the sheet."""
    if view is None:
        return None
    out = {}
    if view.top_left_cell and (split_ref(view.top_left_cell.strip()) or A1) != A1:
        out["view_top_left"] = view.top_left_cell
    corner = frozen_corner(view)
    if corner and view.pane.top_left_cell:
        ptl = split_ref(view.pane.top_left_cell.strip())
        if ptl is not None and ptl != corner:
            out["scrolling_pane_top_left"] = view.pane.top_left_cell
    return out or None


class C62(Check):
    number = 62
    key = "Formatting/Active cell reset to A1 on all sheets"
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet", "chartsheet")

    def start(self, wb):
        super().start(wb)
        if not STRICT_A1:
            raise GradingError(f"{self.key}: only the strict-A1 rule is implemented (Patrick 2026-10-02)")
        if ACTIVE_PANE_WITHOUT_SELECTION not in ("A1", "unverified", "pane_top_left"):
            raise GradingError(f"{self.key}: unknown ACTIVE_PANE_WITHOUT_SELECTION {ACTIVE_PANE_WITHOUT_SELECTION!r}")
        self.per_sheet: dict[str, dict] = {}
        self.failing: list[str] = []
        self.undecidable: dict[str, str] = {}
        self.excluded: list[dict] = []
        self.scrolled: dict[str, dict] = {}
        self.stale_first: dict[str, str] = {}
        self.no_active_sel: dict[str, str] = {}
        self.stray_labels: dict[str, str] = {}
        self.extra_windows: list[str] = []
        self.n_counted = 0

    def sheet_start(self, head):
        if head.kind not in COUNTED_KINDS:
            self.excluded.append({"sheet": head.name, "kind": head.kind, "state": head.state})
            return False
        if head.state != "visible" and not INCLUDE_HIDDEN_SHEETS:
            self.excluded.append({"sheet": head.name, "kind": head.kind, "state": head.state})
            return False
        self.n_counted += 1
        v = head.view
        if len(head.views) > 1:
            self.extra_windows.append(head.name)
        sc = scroll_state(v)
        if sc:
            self.scrolled[head.name] = sc
        try:
            c = cursor(v, head.name)
        except GradingError as e:            # this sheet's cursor is unknown: decided in finish()
            self.undecidable[head.name] = str(e)
            self.per_sheet[head.name] = {"cursor": "?", "active_pane": getattr(v, "active_pane", None),
                                         "how": "unreadable", "notes": [str(e)]}
            return False
        at_a1 = [rc == A1 for rc in c["cands"]]
        cell = make_ref(*c["cands"][0])
        self.per_sheet[head.name] = {"cursor": "/".join(make_ref(*rc) for rc in c["cands"]),
                                     "active_pane": c["pane"], "how": c["how"],
                                     **({"notes": c["notes"]} if c["notes"] else {})}
        if c["first_selection"] and c["first_selection"][0] != c["pane"]:
            self.stale_first[head.name] = f"{c['first_selection'][0]} {c['first_selection'][1]}"
        if c["no_active_selection"]:
            self.no_active_sel[head.name] = c["no_active_selection"]
        if c["stray_labels"]:
            self.stray_labels[head.name] = c["stray_labels"]
        if all(at_a1):
            if CHECK_SCROLL and sc:
                where = sc.get("view_top_left") or sc.get("scrolling_pane_top_left")
                self.failing.append(head.name)
                self.add_mistake(location(head.name, "A1"),
                                 f"Sheet '{head.name}' has its cursor on A1 but opens scrolled to {where}, "
                                 f"so A1 is not in view; scroll back to A1 before saving.", severity="minor")
            return False
        if any(at_a1):
            self.undecidable[head.name] = (
                f"sheet {head.name!r}: the cursor is A1 or "
                f"{', '.join(make_ref(*rc) for rc in c['cands'] if rc != A1)} depending on how Excel resolves "
                f"{'; '.join(c['notes'])} (unverified), so strict A1 cannot be decided")
            return False
        self.failing.append(head.name)
        hidden = "" if head.state == "visible" else f" ({head.state} sheet)"
        extra = []
        corner = frozen_corner(v)
        if corner and c["cands"][0] == corner:
            extra.append("this is the frozen-pane corner, where Ctrl+Home lands; the rule is strict A1, so "
                         "select A1 itself")
        if c["selection"] and c["selection"].strip() != cell:
            extra.append(f"selected range {c['selection']}")
        if c["notes"]:
            extra.append("; ".join(c["notes"]))
        others = [make_ref(*rc) for rc in c["cands"][1:]]
        self.add_mistake(location(head.name, cell),
                         f"Sheet '{head.name}'{hidden} opens with the cursor on {cell}"
                         + (f" (or {', '.join(others)})" if others else "")
                         + f" ({c['how']}), not A1"
                         + (f" - {'; '.join(extra)}" if extra else "")
                         + ". Park the cursor at A1 on every sheet before saving.",
                         severity="minor")
        return False                     # the head is all this check needs: sheet data is skipped

    def finish(self) -> dict:
        if self.wb is None:
            raise GradingError(f"{self.key}: no workbook")
        if self.n_counted == 0 and not self.excluded:
            raise GradingError(f"{self.key}: the workbook has no sheets")
        if self.undecidable and not self.failing:
            # no sheet fails for certain, so the unknown cursor(s) decide the verdict: no guess
            raise GradingError(f"{self.key}: {len(self.undecidable)} sheet(s) with an undecidable cursor and no "
                               f"sheet failing for certain: " + " | ".join(self.undecidable.values()))
        stats = {"strict_a1": STRICT_A1, "check_scroll": CHECK_SCROLL,
                 "active_pane_without_selection_reading": ACTIVE_PANE_WITHOUT_SELECTION,
                 "n_sheets_counted": self.n_counted,
                 "failing_sheets": self.failing,
                 "undecidable_sheets": self.undecidable,
                 "cursors": self.per_sheet,
                 "first_selection_not_active_pane": self.stale_first,
                 "active_pane_without_selection": self.no_active_sel,
                 "selections_labelled_without_pane": self.stray_labels,
                 "scrolled_views": self.scrolled,
                 "hidden_sheets_counted": [s.name for s in self.wb.sheets
                                           if s.state != "visible" and s.name in self.per_sheet],
                 "excluded_sheets": self.excluded,
                 "sheets_with_extra_window_views": self.extra_windows}
        if self.n_counted == 0:
            return self.verdict("No sheet with a cell cursor (only "
                                + ", ".join(f"'{e['sheet']}' ({e['kind']})" for e in self.excluded) + ").",
                                "unreachable", stats)
        listing = ", ".join(f"'{n}' ({self.per_sheet[n]['cursor']})" for n in self.failing[:8]) \
            + (f" and {len(self.failing) - 8} more" if len(self.failing) > 8 else "")
        undec = ""
        if self.undecidable:
            names = ", ".join(f"'{n}'" for n in list(self.undecidable)[:5]) \
                + (f" and {len(self.undecidable) - 5} more" if len(self.undecidable) > 5 else "")
            undec = (f" The cursor of {len(self.undecidable)} further sheet(s) ({names}) could not be decided "
                     f"(stats.undecidable_sheets); the verdict does not depend on it.")
        return self.verdict(f"All {self.n_counted} sheet(s) open with the cursor at A1.",
                            f"{{n}} of {self.n_counted} sheet(s) do not open with the cursor at A1: {listing}."
                            + undec, stats)
