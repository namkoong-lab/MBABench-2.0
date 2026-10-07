"""80 Formulas/Avoid volatile functions.

Rule as implemented (see docs/checks/80.md):
fail on every formula site that calls INDIRECT, OFFSET, TODAY, NOW, RAND, RANDBETWEEN, INFO
or CELL (the rubric's list; RANDARRAY is counted in stats only).  Formula sites:
  * cell formulas (normal, shared - every child of a flagged master -, array), every
    worksheet and macro sheet, hidden sheets included;
  * defined names that a formula site uses, directly or through other names (visible or
    hidden, built-in or not): a volatile call that lives only in the Name Manager fails when
    the name feeds a cell, CF, DV, chart, sparkline, text link or pivot; a name nothing uses
    (e.g. a dynamic Print_Area) is reported in stats only (COUNT_UNUSED_NAMES);
  * conditional-format formulas (main and x14, colour-scale / data-bar / icon-set thresholds
    included; a 'timePeriod' rule without a stored formula counts as TODAY) and
    data-validation formulas (main and x14).
A call is a real call parsed by detchecks.core.formula: names in text cells or inside
string literals ("OFFSET(", ="TODAY()") never count, nor does a column called NOW (NOW$2),
a LET/LAMBDA parameter called today, or _xludf.TODAY (a user function).  Calls of the
built-in through an eta value (BYROW(r,_xleta.RAND)) count.  A cell that uses a volatile
NAME is not a separate mistake; the name is, with its users listed.

Date-stamp exception (rubric "except where clearly justified (e.g., date-stamping)"): a CELL
whose whole formula is TODAY() or NOW() ("alone": optional spaces and '=', plus the no-op
wrappers Excel keeps as typed - leading unary '+' / '@' and redundant parentheses, so
=+TODAY(), =(NOW()), =@TODAY() are alone too) passes when
nothing references it - no cell formula (any reference, range or whole column/row whose area
contains the cell, shared copies included), no defined name, CF or DV formula, chart series,
sparkline, shape text link, form control, data-table input or pivot-cache source.  A
built-in name (Print_Area) that covers the cell references it only when a formula uses the
name.  A range ending in a defined name (SUM(Start:End)) is placed through the names.
Anything else that contains TODAY/NOW (="Prepared "&TEXT(TODAY(),..), TODAY()-D6,
INT(NOW()), TODAY()+0, -TODAY()) is not alone and fails.  Finding the references needs every
formula of the workbook once more, so when a stamp exists a second streaming pass re-reads
the sheets that hold formulas.

A volatile function name written with whitespace before its '(' (OFFSET (A1,1,1)) is not a
call to the parser (a NAME beside a parenthesised group); whether Excel reads such stored
text as a call is not measured (question 11 in docs/checks/80.md).  Such a site is
undecidable, never guessed either way.

Undecidable sites never change a certain verdict: a stamp reference the check cannot place
(dynamic or relative name endpoint, unknown table), a volatile name whose use cannot be
traced (broken localSheetId), and a 'NAME (' spelling (in a cell, CF or DV formula, a used
defined name, or a referenced TODAY ()/NOW () stamp) are deferred to finish().  When some
other mistake already fails the file they are recorded in stats (date_stamps_undecided,
stamp_refs_unplaced, volatile_names_use_undecidable, space_calls_undecided) and counted in
the summary; when the verdict would depend on them the check raises GradingError (no guess,
no fallback).
"""
from __future__ import annotations

import re

from ..core import formula as F
from ..core.refs import (MAX_COL, MAX_ROW, group_cells, index_to_col, location, parse_range, parse_sqref,
                         quote_sheet, range_to_str)
from ..errors import GradingError
from ._fscan import (EMPTY, FormulaScanCheck, cf_rule_texts, ctrl_formulas, parse_text, name_label,
                     name_location, plural, short)

# The rubric's list (handoff).  formula.VOLATILE_FUNCTIONS also lists RANDARRAY: not counted.
VOLATILE = frozenset({"INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN", "INFO", "CELL"})
STATS_ONLY = frozenset({"RANDARRAY"})
_PREFILTER = tuple(sorted(VOLATILE | STATS_ONLY))
# "alone": the whole cell formula is TODAY() or NOW(), at most wrapped in no-op operators -
# leading unary '+' / '@' and redundant parentheses, which Excel keeps exactly as typed
# (=+TODAY() for a Lotus-style '+TODAY()').  Not '-', not '+0'.
# Groups: 1 = the leading '+' '@' '(' run, 2 = the function, 3 = whitespace between the
# function name and its '(' (see spaced_volatile_names), 4 = the closing parentheses (the
# parentheses of 1 and 4 must balance).
_RE_STAMP = re.compile(r"\s*=?\s*((?:[+@(]\s*)*)(TODAY|NOW)(\s*)\(\s*\)((?:\s*\))*)\s*", re.I)
_FN_PREFIXES = ("_XLFN.", "_XLWS.")


def stamp_function(text: str):
    """(func, spaced) when `text` is a lone date-stamp formula (see _RE_STAMP), else None.
    func is 'TODAY' or 'NOW'; spaced: written 'TODAY ()' (whitespace before '(')."""
    m = _RE_STAMP.fullmatch(text)
    if m is None or m.group(1).count("(") != m.group(4).count(")"):
        return None
    return m.group(2).upper(), bool(m.group(3))


def spaced_volatile_names(f, text: str) -> frozenset:
    """Volatile function names that `text` writes with whitespace before their '('
    ('OFFSET (A1,1,1)', 'RANDBETWEEN\\n(1,9)').  core.formula reads such a name as a NAME
    operand beside a parenthesised group, not as a call; how Excel reads the stored text is
    not measured (question 11 in docs/checks/80.md).  Only bare names count: a sheet-qualified
    'S!OFFSET (', a table column '[Today (x)]' or a sheet called 'Now (2)' are not names here."""
    body = text[1:] if text.startswith("=") else text          # parse()'s offsets skip one '='
    out = set()
    for o in f.operands:
        if o.kind != "name" or o.sheet is not None or o.external is not None:
            continue
        nm = (o.name or "").upper()
        for p in _FN_PREFIXES:
            if nm.startswith(p):
                nm = nm[len(p):]
        if nm not in VOLATILE:
            continue
        rest = body[o.end:]
        after = rest.lstrip()
        if after.startswith("(") and len(after) < len(rest):
            out.add(nm)
    return frozenset(out)


# Policy switches (each a question for Patrick in docs/checks/80.md)
COUNT_UNUSED_NAMES = False             # a volatile name no formula site uses (e.g. a dynamic Print_Area) is stats only
CF_TIME_PERIOD_IS_TODAY = True         # a timePeriod CF rule with no stored formula counts as TODAY
STAMP_IGNORES_BUILTIN_NAMES = True     # Print_Area/Print_Titles/_FilterDatabase covering a stamp do not "reference" it
                                       # (unless a formula uses the name)
MAX_LISTED_EXAMPLES = 5
STAMP = "STAMP"

_REF_KINDS = ("cell", "range", "whole_column", "whole_row", "trimmed_range", "spill", "structured")
# a 3-D qualifier (Jan:Dec!A1) or any qualifier with ':' before '!' (conservative)
_RE_3D = re.compile(r":'?[^()+\-*/&=<>^,;{}\"!:]*'?!")


class _Cand:
    """A TODAY()/NOW()-alone cell (or array range) that may be a date stamp.  spaced: written
    'TODAY ()' - unreferenced it passes whatever Excel reads; referenced it is undecidable."""
    __slots__ = ("sheet", "b", "func", "spaced", "n_refs", "examples", "regex")

    def __init__(self, sheet, b, func, spaced=False):
        self.sheet, self.b, self.func, self.spaced = sheet, b, func, spaced
        self.n_refs = 0
        self.examples: list = []
        self.regex = None
        if b[0] == b[2] and b[1] == b[3]:
            self.regex = re.compile(r"(?<![A-Za-z0-9_.])\$?" + index_to_col(b[1]) + r"\$?" + str(b[0])
                                    + r"(?![0-9])", re.I)

    def hit(self, where: str):
        self.n_refs += 1
        if len(self.examples) < 3 and where not in self.examples:
            self.examples.append(where)


class _Box:
    """A range with defined-name endpoints (Start:End, A1:End) placed through the names: the
    endpoints' RefParts (a name's parts are absolute) and their bounding box."""
    __slots__ = ("parts", "bounds", "raw")

    def __init__(self, parts, raw):
        self.parts, self.raw = tuple(parts), raw
        rs = [x for p in parts for x in ((1, MAX_ROW) if p.kind == "col" else (p.row, p.row))]
        cs = [x for p in parts for x in ((1, MAX_COL) if p.kind == "row" else (p.col, p.col))]
        self.bounds = (min(rs), min(cs), max(rs), max(cs))


def _intersects(a, b) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _shift_bounds(op, dr: int, dc: int, drmax: int = None, dcmax: int = None):
    """Bounds of a reference operand whose relative parts move by dr..drmax rows and
    dc..dcmax columns (drmax/dcmax default to dr/dc: one copy).  Parts pushed off the grid
    are clipped when a span is given, and make the reference #REF! (None) for one copy."""
    if drmax is None and dr == 0 and dc == 0:
        return op.bounds
    span = drmax is not None
    drmax = dr if drmax is None else drmax
    dcmax = dc if dcmax is None else dcmax
    r1 = c1 = 10 ** 9
    r2 = c2 = 0
    for p in op.parts:
        if p.kind == "cell":
            ra = (p.row, p.row) if p.row_abs else (p.row + dr, p.row + drmax)
            ca = (p.col, p.col) if p.col_abs else (p.col + dc, p.col + dcmax)
        elif p.kind == "col":
            ra = (1, MAX_ROW)
            ca = (p.col, p.col) if p.col_abs else (p.col + dc, p.col + dcmax)
        elif p.kind == "row":
            ra = (p.row, p.row) if p.row_abs else (p.row + dr, p.row + drmax)
            ca = (1, MAX_COL)
        else:
            return None
        if not span and (ra[0] < 1 or ra[1] > MAX_ROW or ca[0] < 1 or ca[1] > MAX_COL):
            return None                                   # this copy is #REF!
        ra = (max(1, ra[0]), min(MAX_ROW, ra[1]))
        ca = (max(1, ca[0]), min(MAX_COL, ca[1]))
        if ra[0] > ra[1] or ca[0] > ca[1]:
            return None
        r1, r2, c1, c2 = min(r1, ra[0]), max(r2, ra[1]), min(c1, ca[0]), max(c2, ca[1])
    return (r1, c1, r2, c2) if r2 else None


class C80(FormulaScanCheck):
    number = 80
    key = "Formulas/Avoid volatile functions"

    # ------------------------------------------------------------------ setup
    def start(self, wb):
        super().start(wb)
        self.n_randarray = 0
        self.randarray_examples: list = []
        self._space_sites: list = []      # deferred "cannot tell whether ... calls" messages (cells, CF, DV)
        self._unplaced_refs: list = []    # deferred "cannot place reference" messages (stamp references)
        self._cands: list = []
        self._cands_by_sheet: dict = {}
        self._formula_sheets: list = []
        self._requested: set = set()
        self._aux: list = []              # (sheet, what, sqref, ranges, [texts]) CF / DV / sparklines
        self._data_tables: list = []      # (sheet, row, col, [input refs])
        self._tables = None
        self._bi_idents = None            # built-in names used by formulas (stamp references), lazily
        # (index, name, (functions, text, spaced)): spaced = the name only writes 'OFFSET (' -
        # whether it calls the function is undecidable (decided in finish() if the name is used)
        self._name_items = []
        flagged = []
        for i, dn in enumerate(wb.defined_names):
            self.site = name_location(dn)
            try:
                vol, spaced = self._scan(dn.text)
            except F.FormulaError as e:
                # Patrick 2026-10-05 (every attempt graded): an unparsable definition is skipped
                self.note_default("unparsable_formula", f"{self.site}: {e}")
                continue
            if vol or spaced:
                flagged.append(i)
                self._name_items.append((i, dn, (vol or spaced, short(dn.text, 100), not vol)))
        self.set_flagged_names(flagged, use_decides=not COUNT_UNUSED_NAMES)

    # ------------------------------------------------------------------ rule
    def _scan(self, text):
        """(volatile functions `text` calls, volatile names it writes as 'NAME (' when it calls
        none).  RANDARRAY calls are counted in stats."""
        if not text or not F.quick_may_call(text, _PREFILTER):
            return EMPTY, EMPTY
        f = parse_text(text)
        called = {c.name for c in f.functions if c.builtin and not c.local}
        if called & STATS_ONLY:
            self.n_randarray += 1
            if len(self.randarray_examples) < MAX_LISTED_EXAMPLES:
                self.randarray_examples.append(f"{self.site}: ={short(text, 80)}")
        vol = frozenset(called & VOLATILE)
        return vol, (EMPTY if vol else spaced_volatile_names(f, text))

    def classify(self, text, sheet):
        vol, spaced = self._scan(text)
        if spaced:                          # cell, CF or DV formula: undecidable, decided in finish()
            self._space_sites.append(_space_message(self.key, self.site, spaced, text))
        return (vol, short(text, 100)) if vol else None

    def _classify_cell_text(self, text, cell):
        st = stamp_function(text)
        if st:
            self.n_formula_texts += 1
            return (STAMP, st), self.safe_uses(text, cell.sheet)
        return super()._classify_cell_text(text, cell)

    def record(self, cell, result):
        key, detail = result
        if key != STAMP:
            return super().record(cell, result)
        b = (cell.row, cell.col, cell.row, cell.col)
        f = cell.formula
        if f.kind == "array" and f.ref and ":" in f.ref:
            rb = parse_sqref(f.ref)
            if rb:
                b = rb[0]
        c = _Cand(cell.sheet, b, *detail)
        self._cands.append(c)
        self._cands_by_sheet.setdefault(cell.sheet.lower(), []).append(c)

    def describe_cells(self, sheet, rng, n, funcs, text):
        cells = "Formula" if n == 1 else f"{n} formulas"
        return (f"{cells} in {rng} call{'s' if n == 1 else ''} {', '.join(sorted(funcs))} (volatile)"
                f"{'' if n == 1 else ', e.g.'}: ={text}")

    def describe_rule(self, what, sheet, sqref, keys, texts):
        funcs = sorted({f for k, _ in keys for f in k})
        shown = f": ={short(texts[0], 100)}" if texts else f" ({keys[0][1]})"
        return f"{what} on {sqref} calls {', '.join(funcs)} (volatile){shown}"

    def cf_rule_findings(self, cf):
        out = []
        if not CF_TIME_PERIOD_IS_TODAY:
            return out
        for r in cf.rules:
            if (r.type or "") == "timePeriod" and not any(F.quick_may_call(t, ("TODAY",)) for t in r.formulas):
                out.append((frozenset({"TODAY"}), f"timePeriod '{r.time_period}' rule"))
        return out

    def other_formula(self, cell, f):
        if f.kind == "dataTable":
            ins = [f.attrs.get(k) for k in ("r1", "r2") if f.attrs.get(k)]
            if ins:
                self._data_tables.append((cell.sheet, cell.ref, ins))
            self._sheet_formula_cells += 1

    # ------------------------------------------------------------------ per sheet
    def sheet_end(self, head, tail):
        super().sheet_end(head, tail)
        name = head.name
        for cf in tail.conditional_formats:
            texts = [t for r in cf.rules for t in cf_rule_texts(r)]
            if texts:
                self._aux.append((name, "conditional format", cf.sqref, cf.ranges, texts, True))
        for dv in tail.data_validations:
            texts = [t for t in (dv.formula1, dv.formula2) if t]
            if texts:
                self._aux.append((name, "data validation", dv.sqref, dv.ranges, texts, True))
        for g in tail.sparkline_groups:
            if g.formulas:
                self._aux.append((name, "sparkline", "", [], list(g.formulas), False))
        if self._sheet_formula_cells:
            self._formula_sheets.append(name)
        if self._cands:
            for s in self._formula_sheets:
                if s not in self._requested:
                    self._requested.add(s)
                    self.request_second_pass(s, cells=None)

    # ------------------------------------------------------------------ dependents of stamps
    def _table_bounds(self):
        if self._tables is None:
            self._tables = []
            for si in self.wb.sheets:
                if si.kind not in ("worksheet", "macrosheet") or not si.part:
                    continue
                for t in self.wb.tables(si):
                    b = parse_range(t.ref or "")
                    if b is None:
                        raise GradingError(f"{self.key}: table {t.display_name!r} on '{si.name}' has an "
                                           f"unreadable ref {t.ref!r}")
                    self._tables.append(((t.display_name or t.name or "").lower(), (t.name or "").lower(),
                                         si.name.lower(), b))
        return self._tables

    def _sheets_of(self, op, own: str):
        """Lower-case sheet names of this workbook an operand points at (None: external)."""
        if op.external is not None:
            return None
        if op.ambiguous_book and op.sheet.lower() not in {s.lower() for s in self.sheet_names}:
            return None
        if op.sheet is None:
            return (own.lower(),) if own is not None else tuple(self._cands_by_sheet)
        if op.sheet_end is None:
            return (op.sheet.lower(),)
        low = [s.lower() for s in self.sheet_names]
        a, b = op.sheet.lower(), op.sheet_end.lower()
        if a not in low or b not in low:
            return ()                                    # 3-D span over a missing sheet: #REF!
        i, j = sorted((low.index(a), low.index(b)))
        return tuple(low[i:j + 1])

    def _unplaced(self, what: str, own, text: str, why: str = ""):
        """Defer a stamp reference that cannot be placed: the message is raised in finish()
        only when the verdict depends on it (no other mistake), else recorded in stats."""
        self._unplaced_refs.append(
            f"{self.key}: cannot place {what} in ={short(text, 80)} on '{own}'{f' ({why})' if why else ''} "
            f"while deciding whether a TODAY()/NOW() stamp is referenced")

    def _targets(self, text: str, own: str, cell_rc=None):
        """_targets_of(), or no target when the formula text cannot be parsed: skipped (Patrick 2026-10-05: every
        attempt graded; stats.defaults.unparsable_formula)."""
        try:
            return self._targets_of(text, own, cell_rc)
        except F.FormulaError as e:
            self.note_default("unparsable_formula", f"{self.site or own}: {e}")
            return []

    def _targets_of(self, text: str, own: str, cell_rc=None):
        """[(sheets, op, fixed_bounds)] for the reference operands of `text` that point at a
        sheet holding a stamp candidate.  fixed_bounds is set for table references.  A
        reference that cannot be placed is deferred (_unplaced), never guessed."""
        f = parse_text(text)
        out = []
        for o in f.operands:
            if o.kind not in _REF_KINDS:
                continue
            if o.kind == "range" and o.bounds is None and any(p.kind == "name" for p in o.parts):
                placed = self._name_range(o, own, text)
                if placed is not None and placed[0] in self._cands_by_sheet:
                    out.append(((placed[0],), placed[1], None))
                continue
            if o.kind != "structured":       # a table reference carries no sheet: placed via its table
                sheets = self._sheets_of(o, own)
                if not sheets or not any(s in self._cands_by_sheet for s in sheets):
                    continue
            elif o.external is not None:
                continue
            if o.kind == "structured":
                hits = []
                for dn, tn, ts, b in self._table_bounds():
                    if o.table is not None and o.table.lower() in (dn, tn):
                        hits.append((ts, b))
                    elif o.table is None and cell_rc is not None and ts == (own or "").lower() \
                            and b[0] <= cell_rc[0] <= b[2] and b[1] <= cell_rc[1] <= b[3]:
                        hits.append((ts, b))
                if not hits:
                    self._unplaced(f"table reference {o.raw!r}", own, text, "no such table")
                    continue
                for ts, b in hits:
                    if ts in self._cands_by_sheet:
                        out.append(((ts,), None, b))
                continue
            if o.bounds is None:
                self._unplaced(f"reference {o.raw!r}", own, text)
                continue
            out.append((sheets, o, None))
        if out and any(k == "op" and t == ":" for k, t, _, _ in f.tokens):
            # a range operator outside a reference (A1:INDEX(C1:C10,5), (A1:B2):C3) spans the box
            # between its ends: add the bounding box of the references of each sheet (conservative)
            by_sheets: dict = {}
            for sheets, o, fixed in out:
                if fixed is None:
                    by_sheets.setdefault(sheets, []).append(o)
            for sheets, ops in by_sheets.items():
                if len(ops) > 1:
                    out.append((sheets, tuple(ops), None))
        return out

    def _name_range(self, o, own, text):
        """(lower-case sheet, _Box) of a range whose endpoints include defined names (Start:End,
        A1:End), or None when it references nothing (another workbook, an undefined name:
        #NAME?) or cannot be placed (a name endpoint that is not one fixed reference: deferred
        through _unplaced, never guessed)."""
        if o.external is not None:
            return None
        tab, _ = self.name_table()

        def fail(why):
            self._unplaced(f"reference {o.raw!r}", own, text, why)
            return None
        parts, sheets = [], set()
        for p in o.parts:
            if p.kind in ("cell", "col", "row"):
                sh = self._sheets_of(o, own)
                if sh is None:
                    return None
                if len(sh) != 1:
                    return fail("3-D or unplaced qualifier")
                parts.append(p)
                sheets.add(sh[0])
                continue
            if p.kind != "name":
                return None                                  # #REF! endpoint: references nothing
            dn = tab.resolve(p.text, own, o.sheet, o.workbook_scoped)
            if dn is None:
                if any(b.name.lower() in (p.text.lower(), "_xlnm." + p.text.lower()) for _, b in self.bad_scope_names):
                    return fail(f"name {p.text!r} has a localSheetId that names no sheet")
                return None                                  # undefined name: #NAME?
            nf = parse_text(dn.text)
            ops = nf.operands
            r = ops[0] if len(ops) == 1 else None
            if nf.functions or r is None or r.kind not in ("cell", "range", "whole_column", "whole_row") \
                    or r.external is not None or r.sheet_end is not None or r.bounds is None \
                    or any((q.kind in ("cell", "row") and not q.row_abs) or (q.kind in ("cell", "col") and not q.col_abs)
                           for q in r.parts):
                return fail(f"name {p.text!r} = {short(dn.text, 60)} is not one fixed reference")
            sh = r.sheet if r.sheet is not None else (dn.scope or own)
            if sh is None:
                return fail(f"name {p.text!r} = {short(dn.text, 60)} has no sheet")
            parts.extend(r.parts)
            sheets.add(sh.lower())
        if len(sheets) != 1:
            return fail("its endpoints lie on different sheets")
        return sheets.pop(), _Box(parts, o.raw)

    def _builtin_cands(self, text, sheet):
        try:
            return self._builtin_cands_of(text, sheet)
        except F.FormulaError:
            return []                       # unparsable: skipped (recorded by _targets / classify)

    def _builtin_cands_of(self, text, sheet):
        """Stamp candidates inside a built-in name (Print_Area, Print_Titles, _FilterDatabase)
        that `text` uses: a formula that reads Print_Area reads the stamp."""
        if not STAMP_IGNORES_BUILTIN_NAMES or not text or not self._cands:
            return ()
        if self._bi_idents is None:
            tab, widx = self.name_table()
            self._bi_entries = {id(tn): tn for k, tn in enumerate(tab.names)
                                if self.wb.defined_names[widx[k]].builtin and tn.text}
            self._bi_idents = tuple({tn.name.lower() for tn in self._bi_entries.values()})
            self._bi_cover: dict = {}
        if not self._bi_idents:
            return ()
        low = text.lower()
        if not any(x in low for x in self._bi_idents):
            return ()
        out = []
        for u in F.names_used_by(parse_text(text), sheet, self._name_tab, include_indirect_literals=True,
                                 sheet_names=self.sheet_names):
            tn = self._bi_entries.get(id(u))
            if tn is None:
                continue
            if id(tn) not in self._bi_cover:
                hits: list = []
                self._test(self._targets(tn.text, None), "", any_offset=True, collect=hits)
                self._bi_cover[id(tn)] = list(dict.fromkeys(hits))
            out.extend(self._bi_cover[id(tn)])
        return out

    def _test(self, targets, where: str, dr=0, dc=0, drmax=None, dcmax=None, any_offset=False, collect=None):
        for sheets, op, fixed in targets:
            if fixed is not None:
                b = fixed
            elif isinstance(op, tuple):                   # bounding box of a range-operator span
                bs = [_shift_bounds(o, -MAX_ROW, -MAX_COL, MAX_ROW, MAX_COL) if any_offset
                      else _shift_bounds(o, dr, dc, drmax, dcmax) for o in op]
                bs = [x for x in bs if x is not None]
                b = (min(x[0] for x in bs), min(x[1] for x in bs), max(x[2] for x in bs),
                     max(x[3] for x in bs)) if bs else None
            elif any_offset:
                b = _shift_bounds(op, -MAX_ROW, -MAX_COL, MAX_ROW, MAX_COL)
            else:
                b = _shift_bounds(op, dr, dc, drmax, dcmax)
            if b is None:
                continue
            for s in sheets:
                for c in self._cands_by_sheet.get(s, ()):
                    if _intersects(b, c.b):
                        if collect is not None:
                            collect.append(c)
                        else:
                            c.hit(where)

    def _cand_tables(self):
        """Lower-case names of the tables whose range holds a stamp candidate."""
        if not hasattr(self, "_ctabs"):
            self._ctabs = tuple({n for dn, tn, ts, b in self._table_bounds()
                                 for c in self._cands_by_sheet.get(ts, ()) if _intersects(b, c.b)
                                 for n in (dn, tn) if n})
        return self._ctabs

    def _endpoint_names(self):
        """Lower-case identifiers of the defined names whose definition may point at a sheet that
        holds a stamp candidate (its name in the text, or no '!' at all): a range ending in such a
        name (Start:End) can cover the stamp although neither name does."""
        if not hasattr(self, "_epn"):
            tab, _ = self.name_table()
            out = set()
            for tn in tab.names:
                low = (tn.text or "").lower()
                if "!" not in low or any(sh in low or sh.replace("'", "''") in low for sh in self._cands_by_sheet):
                    out.add(tn.name.lower())
            self._epn = tuple(out)
        return self._epn

    def _may_touch(self, text: str, own: str, per_cell: bool) -> bool:
        m = F.mask_strings(text)
        low = m.lower()
        own_l = own.lower()
        if "[" in m and any(t in low for t in self._cand_tables()):
            return True
        if ":" in m and any(x in low for x in self._endpoint_names()):
            return True                                   # Start:End may span a stamp's sheet
        for s, cands in self._cands_by_sheet.items():
            if s == own_l:
                if not per_cell or ":" in m or "[" in m:
                    return True
                if any(c.regex is None or c.regex.search(m) for c in cands):
                    return True
            elif "!" in m:
                if s in low or s.replace("'", "''") in low or _RE_3D.search(m):
                    return True
        return False

    def second_pass_start(self, head):
        self._sp = {}

    def second_pass_cell(self, cell):
        f = cell.formula
        if f is None:
            return
        text = f.text
        where = location(cell.sheet, cell.ref)
        if text:
            shared = f.kind == "shared"
            tg = self._targets(text, cell.sheet, (cell.row, cell.col)) \
                if self._may_touch(text, cell.sheet, per_cell=not shared) else []
            bi = self._builtin_cands(text, cell.sheet)
            if shared:
                self._sp[f.si] = (tg, cell.row, cell.col, bi)
            if tg:
                self._test(tg, where)
            for c in bi:
                c.hit(where)
        elif f.kind == "shared":
            m = self._sp.get(f.si)
            if m is None:
                # no master before it: skipped (Patrick 2026-10-05: every attempt graded; recorded in pass 1)
                return
            tg, mr, mc, bi = m
            if tg:
                self._test(tg, where, cell.row - mr, cell.col - mc)
            for c in bi:
                c.hit(where)

    def _dependents_outside_cells(self):
        # defined names: unqualified references inside a name mean "the sheet that uses it";
        # relative parts move with the using cell, so they may reach any row/column
        for dn in self.wb.defined_names:
            if STAMP_IGNORES_BUILTIN_NAMES and dn.builtin:
                continue
            if not dn.text:
                continue
            tg = self._targets(dn.text, None)
            self._test(tg, name_location(dn), any_offset=True)
            for c in self._builtin_cands(dn.text, name_scope_or_none(dn)):
                c.hit(name_location(dn))
        # CF / DV (relative to the top-left cell of the sqref) and sparklines
        for sheet, what, sqref, ranges, texts, relative in self._aux:
            where = f"{what} on {location(sheet, ','.join(sqref.split()) or '?')}"
            if relative and ranges:
                ar, ac = ranges[0][0], ranges[0][1]
                drmin, drmax = min(r[0] for r in ranges) - ar, max(r[2] for r in ranges) - ar
                dcmin, dcmax = min(r[1] for r in ranges) - ac, max(r[3] for r in ranges) - ac
            else:
                drmin = drmax = dcmin = dcmax = 0
            for t in texts:
                self._test(self._targets(t, sheet), where, drmin, dcmin, drmax, dcmax)
                for c in self._builtin_cands(t, sheet):
                    c.hit(where)
        # data-table input cells
        for sheet, ref, ins in self._data_tables:
            for t in ins:
                self._test(self._targets(t, sheet), f"data table at {location(sheet, ref)}")
        # charts and shape text links (absolute references)
        side = [(t, ch.sheet, f"chart {ch.part}") for ch in self.wb.charts for t in ch.formulas]
        side += [(t, dr.sheet, f"shape text link on '{dr.sheet}'") for dr in self.wb.drawings for t in dr.textlinks]
        side += [(t, sh, f"form control on '{sh}' ({a})") for sh, _, a, t in ctrl_formulas(self.wb)]
        for t, sh, where in side:
            self._test(self._targets(t, sh), where)
            for c in self._builtin_cands(t, sh):
                c.hit(where)
        # pivot-cache sources
        for pc in self.wb.pivot_caches:
            if pc.source_target:
                continue                                  # another workbook
            if pc.source_sheet and pc.source_ref:
                t = f"{quote_sheet(pc.source_sheet)}!{pc.source_ref}"
                self._test(self._targets(t, None), f"pivot cache {pc.part}")
            elif pc.source_name:
                for dn, tn, ts, b in self._table_bounds():
                    if pc.source_name.lower() in (dn, tn):
                        for c in self._cands_by_sheet.get(ts, ()):
                            if _intersects(b, c.b):
                                c.hit(f"pivot cache {pc.part}")

    # ------------------------------------------------------------------ verdict
    def finish(self):
        self.tally_side_uses()
        unused, undecided_names, spaced_unused = [], [], []
        space_undecided = list(self._space_sites)          # cells, CF, DV: 'OFFSET (' spellings
        n_names_counted = 0
        for i, dn, (funcs, _, spaced) in self._name_items:
            label = f"{dn.name}{' (built-in)' if dn.builtin else ''}={short(dn.text, 80)}"
            if not COUNT_UNUSED_NAMES:
                try:
                    used = self.name_used(i)
                except GradingError as e:          # its use cannot be traced: decided below
                    undecided_names.append((dn, str(e)))
                    continue
                if not used:
                    (spaced_unused if spaced else unused).append(label)
                    continue
            if spaced:                             # used, but does it call the function?  decided below
                space_undecided.append(_space_message(self.key, name_location(dn), funcs, dn.text)
                                       + f"; {self.usage_text(i)}")
                continue
            n_names_counted += 1
            self._wb_items.append((name_location(dn),
                                   f"The {name_label(dn)} calls {', '.join(sorted(funcs))} (volatile): "
                                   f"={short(dn.text, 100)}; {self.usage_text(i)}."))
        stamps_ok, stamps_bad, stamps_undecided = [], [], []
        n_unplaced_cands = 0
        if self._cands:
            self._dependents_outside_cells()
            groups: dict = {}
            for c in self._cands:
                where = location(c.sheet, range_to_str(*c.b))
                if c.n_refs and c.spaced:
                    # referenced 'TODAY ()': a referenced stamp if Excel reads a call, else no call
                    stamps_undecided.append(where)
                    space_undecided.append(
                        f"{self.key}: cannot tell whether {where} ={c.func} () calls {c.func}: the function "
                        f"name is followed by whitespace before '(' (Excel's reading of a stored 'NAME (' is "
                        f"not measured; docs/checks/80.md question 11) and the cell is referenced by "
                        f"{plural(c.n_refs, 'formula site')} ({', '.join(c.examples)})")
                elif c.n_refs:
                    groups.setdefault((c.sheet, c.func), []).append(c)
                elif self._unplaced_refs:
                    # an unplaceable reference may cover any candidate nothing else references
                    stamps_undecided.append(where)
                    n_unplaced_cands += 1
                else:
                    stamps_ok.append(where)
            for (sheet, func), cs in groups.items():
                cells = {}
                for c in cs:
                    for r in range(c.b[0], c.b[2] + 1):
                        for col in range(c.b[1], c.b[3] + 1):
                            cells[(r, col)] = c
                for b in group_cells(cells):
                    inside = [c for c in cs if _intersects(c.b, b)]
                    n = sum(c.n_refs for c in inside)
                    ex = list(dict.fromkeys(e for c in inside for e in c.examples))[:3]
                    rng = range_to_str(*b)
                    stamps_bad.append(location(sheet, rng))
                    self._sheet_items.append((
                        location(sheet, rng),
                        f"={func}() in {rng} would be a date stamp, but it is referenced by "
                        f"{plural(n, 'formula site')} ({', '.join(ex)}{' ...' if n > len(ex) else ''}); "
                        f"TODAY()/NOW() is exempt only as a stamp nothing references."))
        # Undecidable sites (a volatile name whose use cannot be traced, a stamp reference that
        # cannot be placed, an 'OFFSET (' spelling) decide nothing when another mistake already
        # fails the file; when the verdict would depend on them there is no guess: raise.
        undecidable = ([why for _, why in undecided_names] + space_undecided
                       + (self._unplaced_refs if n_unplaced_cands else []))
        if undecidable and not self._wb_items and not self._sheet_items:
            more = f" (+{len(undecidable) - 1} more undecidable site(s))" if len(undecidable) > 1 else ""
            raise GradingError(undecidable[0] + more)
        self.emit()
        n_undecided = len(undecided_names) + len(space_undecided) + n_unplaced_cands
        stats = {"n_formula_cells": self.n_formula_cells, "n_formula_texts": self.n_formula_texts,
                 "n_flagged_cells": self.n_flagged_cells, "n_names": len(self.wb.defined_names),
                 "n_names_flagged": n_names_counted,
                 "volatile_names_unused_not_counted": unused,
                 "volatile_names_use_undecidable": [f"{dn.name}={short(dn.text, 80)}: {why}" for dn, why in undecided_names],
                 "n_cf_rules": self.n_cf_rules, "n_data_validations": self.n_dv,
                 "date_stamps_exempt": stamps_ok, "date_stamps_referenced": stamps_bad,
                 "date_stamps_undecided": stamps_undecided,
                 "n_stamp_refs_unplaced": len(self._unplaced_refs),
                 "stamp_refs_unplaced": self._unplaced_refs[:MAX_LISTED_EXAMPLES],
                 "stamp_second_pass_sheets": sorted(self._requested),
                 "randarray_sites_not_counted": self.n_randarray, "randarray_examples": self.randarray_examples,
                 "n_space_calls_undecided": len(space_undecided),
                 "space_calls_undecided": space_undecided[:MAX_LISTED_EXAMPLES],
                 "space_call_names_unused": spaced_unused}
        exempt = f" {plural(len(stamps_ok), 'unreferenced TODAY()/NOW() date stamp')} exempt." if stamps_ok else ""
        undec = (f" {plural(n_undecided, 'undecidable site')} recorded in stats "
                 f"(the verdict does not depend on them)." if n_undecided else "")
        return self.verdict(
            f"No volatile function in {plural(self.n_formula_cells, 'formula cell')}, "
            f"{plural(len(self.wb.defined_names), 'defined name')} and the CF/DV rules.{exempt}",
            "{n} place(s) call volatile functions." + exempt + undec, stats)


def name_scope_or_none(dn):
    try:
        return dn.scope
    except GradingError:
        return None


def _space_message(key: str, site: str, names, text: str) -> str:
    return (f"{key}: cannot tell whether {site} calls {', '.join(sorted(names))}: ={short(text, 80)} writes "
            f"the function name with whitespace before '(' (Excel's reading of a stored 'NAME (' is not "
            f"measured; docs/checks/80.md question 11)")
