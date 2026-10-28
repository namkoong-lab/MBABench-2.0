"""100 Potential Dangers/Automatic calculation mode.

Rule (rubric_9 #100; Patrick's hierarchy of 2026-10-28, judge v15; details in docs/checks/100.md).
Rubric: "The workbook should be delivered with calculation set to automatic - and, where intentional
circularity requires it, with iterative calculation enabled and that requirement stated. Manual
calculation mode means displayed values may not reflect the inputs currently shown." / bad: "saved
in manual calculation mode without disclosure ... or iterative calculation is required but not
enabled".  House rules v3: "Circularity only where the economics require it, each instance noted and
fitted with a breaker switch, with iterative calculation enabled and stated."

Evaluated in this order; the first rule that fires decides:
  1. calculation mode MANUAL (xl/workbook.xml <calcPr calcMode="manual">)        -> FAIL by code
     (Patrick: manual always fails, the rubric's "without disclosure" is ignored; "autoNoTable" is
     automatic; no <calcPr> = automatic, iteration off);
  2. the formula dependency graph is built and searched for cycles (below);
  3. no cycle                                                                    -> PASS by code
     (iteration on or off; "iteration on without a cycle" is recorded in stats, not a mistake);
  4. a cycle, iterative calculation OFF (<calcPr iterate="1"> absent)            -> FAIL by code,
     one mistake per cycle: "circular reference; iterative calculation is off";
  5. a cycle, iteration ON                                                       -> UNDECIDED
     (base.py): the LLM judge is asked the one question in LLM_QUESTION - does the workbook state
     that the model is intentionally circular, where, and that iterative calculation is required
     (cover sheet, notes, legend or comments)? - and its verdict counts for this attempt: pass only
     if it does (rule 6: anything else fails, the LLM's call).  A breaker switch excuses nothing by
     itself: switch, iteration and statement are all required (Patrick 2026-10-28).

The graph (static: a cycle through an IF branch that is never both active still counts - consistent
with the house rules' iteration requirement even with a breaker switch OFF, Patrick 2026-10-28):
  * nodes: every formula cell of every worksheet / macro sheet (hidden ones too): normal, shared
    master and every child (the master's references shifted by the child's offset, as
    core.formula.shift does), array / spill anchors; a written non-anchor member of an array range
    is a node that depends on its anchor, and a reference into the range (members some writers leave
    unwritten included) reaches the anchor; every defined name with a readable scope (a sheet-local
    name's unqualified references are on its sheet; a workbook name's unqualified references are
    relative to the active sheet in Excel and are skipped, counted in stats);
  * edges: from a formula to each cell, range, spill reference (A1# -> its anchor), structured
    reference (resolved to the table's rows and columns, [#This Row] to the formula's row) and
    defined name it reads, directly or as MyLambda(...); a 3-D reference reaches every sheet of the
    span; a reference to a constant cell, an external workbook, #REF!, R1C1 text or a sheet that does
    not exist leads nowhere;
  * ranges are range nodes (a whole column is never expanded into a million cells): a range node
    leads to the formula cells inside it through a per-sheet column index, and a tall range is cut
    along the dyadic row blocks of a segment tree first, so the cumulative family =SUM($B$1:B5),
    =SUM($B$1:B6), ... shares its blocks instead of enumerating n^2 cells;
  * what-if data tables (<f t="dataTable">) are constants: no edges (not a cycle);
  * a reference that a function uses for its GEOMETRY only is never dereferenced, so it is no
    edge (GEOMETRY_FUNCS: ROWS, COLUMNS, ROW, COLUMN, AREAS, SHEET, SHEETS, ISREF, FORMULATEXT,
    ISFORMULA): the counter idioms =ROWS($B$14:B14) in B14 and =COLUMNS($E$21:E$21)-1 in E21 are
    not circular in Excel - attempt 2539 (Excel Online, iteration off) caches 1, 2, 3 for them,
    the v15 sanity run's first cut flagged 340 of them as self-loops;
  * the base / source reference of INDEX (first argument) and OFFSET (first argument) is
    dereferenced only in the part computed at run time (=SUM(E8:INDEX(E:E,n)) in E4 reads E8:En,
    =INDEX($A$1:$A$10,ROW()-1) in A5 reads A4): no edge, the miss is accepted and recorded with the
    INDIRECT / OFFSET calls (stats.dynamic_refs; such files already fail Avoid volatile functions
    (80) where OFFSET / INDIRECT are involved);
  * cycles: Tarjan's strongly connected components (iterative) over cells, ranges, blocks and names;
    a cycle is a component of more than one node or a self-loop (=A1+1 in A1; =SUM(A1:A10) in A5).
    Reported by its cells, grouped per sheet (capped), smallest cycle first.
Values are never read (needs_values = False): no LibreOffice run for this check.  A workbook whose
graph walk exceeds WORK_LIMIT successor steps raises GradingError (no fallback): the limit is set far
above the corpus (attempt 3237, 5.9 million formula cells, ~90 million steps) so that only a
pathological file reaches it.
"""
from __future__ import annotations

import time
from bisect import bisect_left, bisect_right
from typing import Iterator, Optional

import re

from ..core import formula as F
from ..core.recalc import EXCEL_FUNCTIONS
from ..core.refs import MAX_COL, MAX_ROW, col_to_index, group_cells_str, location, make_ref, parse_range, split_ref
from ..errors import GradingError
from ._fscan import name_location
from .base import Check

parse_text = getattr(F.parse, "__wrapped__", F.parse)      # no lru cache: one parse per distinct text (see _fscan)

LEAF_ROWS = 64                  # a dyadic row block this tall or shorter enumerates its cells directly
SMALL_BLOCK_CELLS = 48          # ... or any block holding at most this many formula cells
UNIVERSE_LEVEL = 20             # 2 ** 20 = 1048576 = MAX_ROW: the segment tree's root block
MAX_CYCLES_LISTED = 10
MAX_CELLS_PER_CYCLE = 12        # cells quoted per cycle (per sheet, after grouping into rectangles)
MAX_DYNAMIC_EXAMPLES = 8
TEMPLATE_CACHE_MAX = 500_000    # distinct formula texts remembered
WORK_LIMIT = 1_000_000_000      # successor steps; beyond it the walk is pathological and the check raises (the
                                # largest corpus attempt, 5.9 million formula cells, needs ~90 million; the walk is
                                # linear in edges, ~1 million steps per second)
DYNAMIC_FUNCS = frozenset({"INDIRECT", "OFFSET"})
# functions whose reference arguments are used for their geometry only (never dereferenced): no edge
GEOMETRY_FUNCS = frozenset({"ROWS", "COLUMNS", "ROW", "COLUMN", "AREAS", "SHEET", "SHEETS", "ISREF", "FORMULATEXT",
                            "ISFORMULA"})
# functions whose given reference argument is a base / source of which only a computed part is read
REFERENCE_ARG_FUNCS = {"INDEX": (0,), "OFFSET": (0,)}
FAST_PATH = True                # read plain formulas with _fast_entries instead of the parser (same entries)

LLM_QUESTION = ("The deterministic grader found a circular reference through these cells: {cycles}. "
                "Iterative calculation is enabled ({settings}). Does the workbook state that the model is "
                "intentionally circular, where the circularity is, and that iterative calculation is required "
                "(cover sheet, notes, legend or comments)? Pass only if it does; a breaker switch alone is not "
                "a statement.")


# ------------------------------------------------------------------------------------------ fast path
# Plain formulas - direct A1 references (sheet-qualified or not), numbers, strings, operators and calls of
# known Excel functions - are the bulk of every large attempt (1.5 million distinct texts in a 67 MB
# claudecode file), and the full parser costs ~50 us each.  _fast_entries reads such a text with one
# regex pass and yields exactly the entries _entries would (INDEX / OFFSET base arguments skipped and
# counted, INDIRECT / OFFSET calls counted, geometry-only arguments skipped); it returns None (-> the
# parser) at the first thing it does not handle: a defined name or table, a spill / implicit-intersection
# marker, an error literal, an array constant, R1C1 text, a 3-D span, a trimmed range, two references
# side by side (the intersection operator), a call of a name or of an unknown function, an external book,
# a separator outside a call.  test_checks_100 checks it against the parser on every form it accepts and
# rejects; the v15 sanity run compared the two on 3.6 million distinct real formula texts (0 mismatches).
_FAST_TOKEN = re.compile(r"""
    (?P<str>"(?:[^"]|"")*")
  | (?P<ref>(?:(?P<q>'(?:[^']|'')+'|[\w.]+)!)?
            (?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+))
    (?![\w(\[#@!.:])
  | (?P<fn>[A-Za-z_][\w.]*)\(
  | (?P<lit>TRUE|FALSE)(?![\w(])
  | (?P<num>\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|\.\d+(?:[Ee][+-]?\d+)?)(?![\w(.])
  | (?P<op><>|<=|>=|[-+*/^&=%,;<>])
  | (?P<open>\()
  | (?P<close>\))
  | (?P<ws>\s+)
""", re.X)
_FAST_PREFIXES = ("_XLFN.", "_XLWS.", "_XLL.", "_XLUDF.", "_XLETA.")
_RE_PART = re.compile(r"^(\$?)([A-Za-z]{0,3})(\$?)(\d*)$")


def _fast_part(p: str) -> Optional[tuple]:
    m = _RE_PART.match(p)
    if not m:
        return None
    cabs, col, rabs, row = m.groups()
    if col and row:
        r, c = int(row), col_to_index(col)
        if not (1 <= r <= MAX_ROW and c <= MAX_COL):
            return None
        return (r, bool(rabs), c, bool(cabs))
    if col:
        c = col_to_index(col)
        return None if c > MAX_COL else (None, True, c, bool(cabs))
    if row:
        r = int(row)
        return None if not 1 <= r <= MAX_ROW else (r, bool(rabs), None, True)
    return None


def _fast_entries(text: str, sheet_idx: dict, info: dict) -> Optional[tuple]:
    """Entries of a plain formula without the parser (module comment), or None.  `info` receives
    what the parser path would have counted: "geometry" (references skipped as geometry-only
    arguments), "index_args" (INDEX / OFFSET base references skipped), "dynamic" (INDIRECT / OFFSET
    calls)."""
    out = []
    stack: list = []                            # [function name or None, argument index] per open parenthesis
    pos = 0
    n = len(text)
    prev = None                                 # kind of the previous non-blank token
    blank = False
    geo = index_args = dynamic = 0
    while pos < n:
        m = _FAST_TOKEN.match(text, pos)
        if m is None:
            return None
        kind = m.lastgroup
        pos = m.end()
        if kind == "ws":
            blank = True
            continue
        if kind == "ref":
            if prev == "ref" and blank:
                return None                     # the intersection operator (A1:A5 B2:B9)
            q = m.group("q")
            sheet = None
            if q is not None:
                sheet = q[1:-1].replace("''", "'") if q.startswith("'") else q
                if sheet.lower() not in sheet_idx:
                    return None                 # a missing sheet, a 3-D span or a book: the parser decides
            body = m.group("ref")[len(q) + 1 if q is not None else 0:]
            parts = body.split(":")
            e1 = _fast_part(parts[0])
            e2 = _fast_part(parts[-1]) if len(parts) > 1 else e1
            if e1 is None or e2 is None:
                return None
            inner = stack[-1][0] if stack else None
            if inner in GEOMETRY_FUNCS:
                geo += 1
            elif inner in REFERENCE_ARG_FUNCS and stack[-1][1] in REFERENCE_ARG_FUNCS[inner]:
                index_args += 1
            else:
                out.append(("rect", sheet, None, e1, e2))
        elif kind == "fn":
            name = m.group("fn").upper()
            changed = True
            while changed:
                changed = False
                for pf in _FAST_PREFIXES:
                    if name.startswith(pf):
                        name = name[len(pf):]
                        changed = True
            if name not in EXCEL_FUNCTIONS:
                return None                     # a defined-name call or an unknown function: the parser decides
            if name in DYNAMIC_FUNCS:
                dynamic += 1
            stack.append([name, 0])
        elif kind == "open":
            stack.append([None, 0])
        elif kind == "close":
            if not stack:
                return None
            stack.pop()
        elif kind == "op" and m.group("op") in (",", ";"):
            if not stack or stack[-1][0] is None:
                return None                     # a separator outside a call: not a formula this path reads
            stack[-1][1] += 1
        prev = kind
        blank = False
    if stack:
        return None
    info["geometry"] = geo
    info["index_args"] = index_args
    info["dynamic"] = dynamic
    return tuple(out)


# ------------------------------------------------------------------------------------------ helpers
def _endpoint(p) -> tuple:
    """(row, row_abs, col, col_abs) of a RefPart; None = the whole span (a column / row part)."""
    if p.kind == "cell":
        return (p.row, p.row_abs, p.col, p.col_abs)
    if p.kind == "col":
        return (None, True, p.col, p.col_abs)
    return (p.row, p.row_abs, None, True)


def _shift(ep: tuple, dr: int, dc: int) -> Optional[tuple]:
    """The endpoint moved with a shared-formula copy; None when it is pushed off the grid (#REF!)."""
    r, ra, c, ca = ep
    if r is not None and not ra:
        r += dr
    if c is not None and not ca:
        c += dc
    if (r is not None and not 1 <= r <= MAX_ROW) or (c is not None and not 1 <= c <= MAX_COL):
        return None
    return r, c


def _rect(a: tuple, b: tuple) -> tuple:
    r1 = min(a[0] or 1, b[0] or 1)
    r2 = max(a[0] or MAX_ROW, b[0] or MAX_ROW)
    c1 = min(a[1] or 1, b[1] or 1)
    c2 = max(a[1] or MAX_COL, b[1] or MAX_COL)
    return r1, c1, r2, c2


def dyadic_blocks(r1: int, r2: int) -> list:
    """Canonical segment-tree blocks over rows 1..2**UNIVERSE_LEVEL covering [r1, r2]: [(level, k)],
    block (level, k) = rows k * 2**level + 1 .. (k + 1) * 2**level."""
    out = []
    a, b = r1 - 1, r2                     # half-open, 0-based
    lvl = 0
    while a < b:
        if a & 1:
            out.append((lvl, a))
            a += 1
        if b & 1:
            b -= 1
            out.append((lvl, b))
        a >>= 1
        b >>= 1
        lvl += 1
    return out


def _split_top(spec: str) -> list:
    """Top-level comma split of a structured-reference spec, brackets respected."""
    out, depth, cur = [], 0, []
    for ch in spec:
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
        if ch == "," and depth == 0:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    out.append("".join(cur))
    return out


def _unescape_col(name: str) -> str:
    return name.replace("'[", "[").replace("']", "]").replace("'#", "#").replace("'@", "@").replace("''", "'").strip()


def parse_structured_spec(spec: str) -> Optional[tuple]:
    """(rows, col_lo, col_hi) of the text after the table name: rows is "this" or a frozenset of
    {"headers", "data", "totals"}; columns None = every column.  None when the spec is not understood.
    Forms: [Col], [[#This Row],[Col]], [@Col], [@[Col]], [[#Totals],[Col]], [[#Headers],[#Data]],
    [#All], [[Col1]:[Col2]], [] (= data)."""
    if not (spec.startswith("[") and spec.endswith("]")):
        return None
    inner = spec[1:-1]
    if inner.strip() == "":
        return frozenset({"data"}), None, None
    parts: set = set()
    this_row = False
    cols = None
    for item in _split_top(inner):
        it = item.strip()
        if it.startswith("@"):
            this_row = True
            it = it[1:].strip()
            if not it:
                continue
        if it.startswith("[") and it.endswith("]") and "]:[" not in it:
            it = it[1:-1].strip()
        if it.startswith("#"):
            key = it.lower()
            if key == "#this row":
                this_row = True
            elif key == "#all":
                parts |= {"headers", "data", "totals"}
            elif key in ("#data", "#headers", "#totals"):
                parts.add(key[1:])
            else:
                return None
            continue
        if "]:[" in it:
            a, b = it.split("]:[", 1)
            cols = (_unescape_col(a.lstrip("[")), _unescape_col(b.rstrip("]")))
        else:
            cols = (_unescape_col(it), _unescape_col(it))
    if this_row:
        return "this", cols[0] if cols else None, cols[1] if cols else None
    return frozenset(parts or {"data"}), cols[0] if cols else None, cols[1] if cols else None


# ------------------------------------------------------------------------------------------ the check
class C100(Check):
    number = 100
    key = "Potential Dangers/Automatic calculation mode"
    needs_cells = True
    needs_values = False
    sheet_kinds = ("worksheet", "macrosheet")

    # ---------------------------------------------------------------- hooks
    def start(self, wb):
        super().start(wb)
        cp = dict(wb.calc_pr or {})
        self.calc_pr = cp
        self.calc_mode = (cp.get("calcMode") or "auto").strip() or "auto"
        self.manual = self.calc_mode.lower() == "manual"
        self.iterate = (cp.get("iterate") or "").strip().lower() in ("1", "true")
        self.iterate_count = cp.get("iterateCount")
        self.iterate_delta = cp.get("iterateDelta")
        self.full_calc_on_load = (cp.get("fullCalcOnLoad") or "").strip().lower() in ("1", "true")
        self._sheet_idx = {s.name.lower(): s.index for s in wb.sheets}
        self._sheet_name = {s.index: s.name for s in wb.sheets}
        self._order = [s.index for s in sorted(wb.sheets, key=lambda s: s.index)]
        self._cells: dict = {}          # si -> {(r, c): (template id, dr, dc)}
        self._members: dict = {}        # si -> {(r, c): (anchor r, anchor c)} written array members
        self._arrays: dict = {}         # si -> [(r1, c1, r2, c2, (ar, ac))]
        self._cols: dict = {}           # si -> {c: sorted rows}  (built in finish)
        self._colkeys: dict = {}        # si -> sorted columns
        self._tables: dict = {}         # table name (lower) -> geometry
        self._sheet_tables: dict = {}   # si -> [geometry]
        self._tpls: list = []
        self._tpl_by_text: dict = {}
        self._masters: dict = {}
        self.site = ""
        good = []
        self.n_names_bad_scope = 0
        for dn in wb.defined_names:
            if dn.scope_error is not None:
                self.n_names_bad_scope += 1
            elif not dn.builtin and (dn.text or "").strip():
                good.append(dn)
        self._name_tab = F.NameTable(good, sheet_names=[s.name for s in wb.sheets])
        self._name_defs: dict = {}      # key -> formula.DefinedName (resolved)
        self._name_tpl: dict = {}       # key -> template id
        self.n_formula_cells = 0
        self.n_member_cells = 0
        self.n_data_table_cells = 0
        self.n_external_refs = 0
        self.n_r1c1_refs = 0
        self.n_unresolved_sheet_refs = 0
        self.n_global_name_unqualified = 0
        self.n_struct_unresolved = 0
        self.n_struct_approximated = 0
        self.n_mixed_ranges = 0
        self.dynamic = {"count": 0, "examples": [], "index_offset_reference_args": 0}
        self.n_geometry_refs = 0
        self.n_fast_texts = 0
        self.n_edges = 0
        self.n_work = 0

    def sheet_start(self, head):
        self._si = head.info.index
        self._cells.setdefault(self._si, {})
        self._members.setdefault(self._si, {})
        self._arrays.setdefault(self._si, [])
        self._masters = {}

    def cell(self, cell):
        f = cell.formula
        arr = cell.array
        if f is None and arr is None:
            return
        si = self._si
        if arr is not None and (f is None or f.is_empty_marker):          # non-anchor member of a range
            af = arr.formula
            if af is None or af.kind == "dataTable":
                self.n_data_table_cells += af is not None
                return
            rc = split_ref(arr.anchor_ref)
            if rc is not None:
                self._members[si][(cell.row, cell.col)] = rc
                self.n_member_cells += 1
            return
        if f.is_empty_marker:
            return
        if f.kind == "dataTable":
            self.n_data_table_cells += 1
            return
        self.site = f"'{cell.sheet}'!{cell.ref}"
        if f.kind == "shared" and not f.text:                              # shared child: the master shifted
            m = self._masters.get(f.si)
            if m is None:
                self.note_default("unparsable_formula", f"{self.site}: shared formula si={f.si} has no master "
                                                        f"formula before it")
                return
            tid, mr, mc = m
            self._cells[si][(cell.row, cell.col)] = (tid, cell.row - mr, cell.col - mc)
            self.n_formula_cells += 1
            return
        tid = self._template(f.text)
        if f.kind == "shared":
            self._masters[f.si] = (tid, cell.row, cell.col)
        elif f.kind == "array":
            b = parse_range(f.ref or "")
            if b is not None and b != (cell.row, cell.col, cell.row, cell.col):
                self._arrays[si].append((b[0], b[1], b[2], b[3], (cell.row, cell.col)))
        self._cells[si][(cell.row, cell.col)] = (tid, 0, 0)
        self.n_formula_cells += 1

    def sheet_end(self, head, tail):
        si = head.info.index
        geoms = []
        try:
            tables = self.wb.tables(head.info)
        except GradingError as e:
            self.note_default("unparsable_formula", f"'{head.name}': table parts cannot be read ({e})")
            tables = []
        for t in tables:
            b = parse_range(t.ref or "")
            if b is None:
                continue
            cols = [str(cn.get("name") or "").lower() for cn in (t.columns or [])]
            g = (si, b[0], b[1], b[2], b[3], int(t.header_row_count or 0), int(t.totals_row_count or 0), cols)
            geoms.append(g)
            for nm in (t.name, t.display_name):
                if nm:
                    self._tables[nm.lower()] = g
        self._sheet_tables[si] = geoms

    # ---------------------------------------------------------------- templates
    def _template(self, text: str) -> int:
        tid = self._tpl_by_text.get(text)
        if tid is not None:
            return tid
        info: dict = {}
        fast = _fast_entries(text, self._sheet_idx, info) if FAST_PATH else None
        if fast is not None:
            entries: tuple = fast
            self.n_fast_texts += 1
            self.n_geometry_refs += info["geometry"]
            self.dynamic["index_offset_reference_args"] += info["index_args"]
            if info["dynamic"]:
                self.dynamic["count"] += info["dynamic"]
                if len(self.dynamic["examples"]) < MAX_DYNAMIC_EXAMPLES:
                    self.dynamic["examples"].append(f"{self.site}: INDIRECT / OFFSET in ={text[:100]}")
        else:
            try:
                f = parse_text(text)
            except F.FormulaError as e:
                self.note_default("unparsable_formula", f"{self.site}: {e}")
                entries = ()
            else:
                entries = self._entries(f)
        tid = len(self._tpls)
        self._tpls.append(entries)
        if len(self._tpl_by_text) < TEMPLATE_CACHE_MAX:
            self._tpl_by_text[text] = tid
        return tid

    def _entries(self, f) -> tuple:
        out = []
        for o in f.operands:
            k = o.kind
            if k not in ("cell", "range", "whole_column", "whole_row", "trimmed_range", "spill", "structured"):
                continue                                   # names: below; literals / params / errors: nothing
            if o.external is not None or (o.ambiguous_book and o.sheet and o.sheet.lower() not in self._sheet_idx):
                self.n_external_refs += 1
                continue
            if k == "structured":
                spec = parse_structured_spec((o.body or "")[len(o.table or ""):])
                if spec is None:
                    self.n_struct_approximated += 1
                    spec = (frozenset({"data"}), None, None)
                out.append(("table", (o.table or "").lower() or None) + spec)
                continue
            if o.r1c1:
                self.n_r1c1_refs += 1
                continue
            inner = o.func_path[-1] if o.func_path else None
            if inner in GEOMETRY_FUNCS:
                self.n_geometry_refs += 1                  # ROWS($B$14:B14) in B14: never dereferenced
                continue
            if inner in REFERENCE_ARG_FUNCS and o.arg_index in REFERENCE_ARG_FUNCS[inner]:
                self.dynamic["index_offset_reference_args"] += 1   # INDEX / OFFSET base: the read part is computed
                continue
            parts = o.parts
            if not parts or any(p.kind not in ("cell", "col", "row") for p in parts):
                self.n_mixed_ranges += any(p.kind == "name" for p in parts)
                continue                                   # #REF!, A1:Name (the name is reached below)
            e1 = _endpoint(parts[0])
            e2 = _endpoint(parts[-1]) if len(parts) > 1 else e1
            if o.spill or k == "spill":
                out.append(("spill", o.sheet, e1))
            else:
                out.append(("rect", o.sheet, o.sheet_end, e1, e2))
        for r in f.names_referenced():
            if r.external is not None or (r.ambiguous_book and r.sheet and r.sheet.lower() not in self._sheet_idx):
                self.n_external_refs += 1
                continue
            out.append(("name", r.name, r.sheet, r.workbook_scoped))
        for c in f.functions:
            if c.builtin and not c.local and c.name in DYNAMIC_FUNCS:
                self.dynamic["count"] += 1
                if len(self.dynamic["examples"]) < MAX_DYNAMIC_EXAMPLES:
                    self.dynamic["examples"].append(f"{self.site}: {c.name} in ={f.text[:100]}")
        return tuple(out)

    # ---------------------------------------------------------------- resolution
    def _sheets_of(self, sheet: Optional[str], sheet_end: Optional[str], own: int) -> list:
        if sheet is None:
            return [own]
        a = self._sheet_idx.get(sheet.lower())
        if a is None:
            self.n_unresolved_sheet_refs += 1
            return []
        if sheet_end is None:
            return [a]
        b = self._sheet_idx.get(sheet_end.lower())
        if b is None:
            self.n_unresolved_sheet_refs += 1
            return []
        lo, hi = min(a, b), max(a, b)
        return [s for s in self._order if lo <= s <= hi]

    def _cell_node(self, si: int, r: int, c: int) -> Optional[tuple]:
        if (r, c) in self._cells.get(si, ()):
            return ("c", si, r, c)
        m = self._members.get(si, {}).get((r, c))
        if m is not None:
            return ("c", si, m[0], m[1])
        for r1, c1, r2, c2, anchor in self._arrays.get(si, ()):
            if r1 <= r <= r2 and c1 <= c <= c2:
                return ("c", si, anchor[0], anchor[1])
        return None

    def _own_table(self, si: int, r: int, c: int):
        for g in self._sheet_tables.get(si, ()):
            if g[1] <= r <= g[3] and g[2] <= c <= g[4]:
                return g
        return None

    def _expand(self, e: tuple, si: int, r: int, c: int, dr: int, dc: int) -> Iterator[tuple]:
        kind = e[0]
        if kind == "rect":
            _, sheet, sheet_end, e1, e2 = e
            a, b = _shift(e1, dr, dc), _shift(e2, dr, dc)
            if a is None or b is None:
                return
            r1, c1, r2, c2 = _rect(a, b)
            for tsi in self._sheets_of(sheet, sheet_end, si):
                if (r1, c1) == (r2, c2):
                    n = self._cell_node(tsi, r1, c1)
                    if n is not None:
                        yield n
                else:
                    yield ("r", tsi, r1, c1, r2, c2)
        elif kind == "spill":
            _, sheet, e1 = e
            a = _shift(e1, dr, dc)
            if a is None or a[0] is None or a[1] is None:
                return
            for tsi in self._sheets_of(sheet, None, si):
                n = self._cell_node(tsi, a[0], a[1])
                if n is not None:
                    yield n
        elif kind == "name":
            _, ident, qual, wb_scoped = e
            hit = self._name_tab.resolve(ident, self._sheet_name.get(si), qual, wb_scoped)
            if hit is None:
                return
            key = (hit.name.lower(), (hit.scope or "").lower() or None)
            self._name_defs.setdefault(key, hit)
            yield ("n", key)
        elif kind == "table":
            _, tname, rows, col_lo, col_hi = e
            g = self._tables.get(tname) if tname else self._own_table(si, r, c)
            if g is None:
                self.n_struct_unresolved += 1
                return
            tsi, tr1, tc1, tr2, tc2, h, tt, cols = g
            if col_lo is None:
                c1, c2 = tc1, tc2
            else:
                try:
                    i, j = cols.index(col_lo.lower()), cols.index(col_hi.lower())
                except ValueError:
                    self.n_struct_approximated += 1
                    i, j = 0, max(len(cols) - 1, 0)
                c1, c2 = tc1 + min(i, j), tc1 + max(i, j)
                c2 = min(c2, tc2)
            if rows == "this":
                if not (tr1 <= r <= tr2):
                    return
                r1 = r2 = r
            else:
                lo, hi = [], []
                if "headers" in rows and h:
                    lo.append(tr1)
                    hi.append(tr1 + h - 1)
                if "data" in rows:
                    lo.append(tr1 + h)
                    hi.append(tr2 - tt)
                if "totals" in rows and tt:
                    lo.append(tr2 - tt + 1)
                    hi.append(tr2)
                if not lo:
                    return
                r1, r2 = min(lo), max(hi)
                if r1 > r2:
                    return
            if (r1, c1) == (r2, c2):
                n = self._cell_node(tsi, r1, c1)
                if n is not None:
                    yield n
            else:
                yield ("r", tsi, r1, c1, r2, c2)

    # ---------------------------------------------------------------- successors
    def _succ(self, node: tuple) -> Iterator[tuple]:
        kind = node[0]
        if kind == "c":
            _, si, r, c = node
            info = self._cells[si].get((r, c))
            if info is None:                                   # a written member: depends on its anchor
                m = self._members[si].get((r, c))
                if m is not None:
                    yield ("c", si, m[0], m[1])
                return
            tid, dr, dc = info
            for e in self._tpls[tid]:
                yield from self._expand(e, si, r, c, dr, dc)
        elif kind == "r":
            _, si, r1, c1, r2, c2 = node
            for ar1, ac1, ar2, ac2, anchor in self._arrays.get(si, ()):
                if ar1 <= r2 and r1 <= ar2 and ac1 <= c2 and c1 <= ac2:
                    yield ("c", si, anchor[0], anchor[1])
            for lvl, k in dyadic_blocks(r1, r2):
                if (1 << lvl) <= LEAF_ROWS:
                    yield from self._cells_in(si, (k << lvl) + 1, (k + 1) << lvl, c1, c2)
                else:
                    yield ("b", si, lvl, k, c1, c2)
        elif kind == "b":
            _, si, lvl, k, c1, c2 = node
            lo, hi = (k << lvl) + 1, (k + 1) << lvl
            n = self._count_in(si, lo, hi, c1, c2, SMALL_BLOCK_CELLS + 1)
            if n == 0:
                return
            if n <= SMALL_BLOCK_CELLS or (1 << lvl) <= LEAF_ROWS:
                yield from self._cells_in(si, lo, hi, c1, c2)
            else:
                yield ("b", si, lvl - 1, 2 * k, c1, c2)
                yield ("b", si, lvl - 1, 2 * k + 1, c1, c2)
        else:                                                  # ("n", key)
            key = node[1]
            tid = self._name_tpl.get(key)
            dn = self._name_defs.get(key)
            if dn is None:
                return
            if tid is None:
                self.site = name_location(dn)
                tid = self._name_tpl[key] = self._template(dn.text)
            scope_si = self._sheet_idx.get(dn.scope.lower()) if dn.scope else None
            for e in self._tpls[tid]:
                if e[0] in ("rect", "spill") and e[1] is None and scope_si is None:
                    self.n_global_name_unqualified += 1     # relative to the active sheet: unknowable
                    continue
                if e[0] == "table" and e[2] == "this":
                    continue
                yield from self._expand(e, scope_si if scope_si is not None else self._order[0], 0, 0, 0, 0)

    def _columns_in(self, si: int, c1: int, c2: int) -> list:
        keys = self._colkeys.get(si, [])
        return keys[bisect_left(keys, c1):bisect_right(keys, c2)]

    def _count_in(self, si: int, r1: int, r2: int, c1: int, c2: int, stop: int) -> int:
        n = 0
        cols = self._cols.get(si, {})
        for c in self._columns_in(si, c1, c2):
            rows = cols[c]
            n += bisect_right(rows, r2) - bisect_left(rows, r1)
            if n >= stop:
                break
        return n

    def _cells_in(self, si: int, r1: int, r2: int, c1: int, c2: int) -> Iterator[tuple]:
        cols = self._cols.get(si, {})
        for c in self._columns_in(si, c1, c2):
            rows = cols[c]
            for i in range(bisect_left(rows, r1), bisect_right(rows, r2)):
                yield ("c", si, rows[i], c)

    # ---------------------------------------------------------------- cycles
    def _build_index(self):
        for si, cells in self._cells.items():
            cols: dict = {}
            for (r, c) in cells:
                cols.setdefault(c, []).append(r)
            for c in self._members.get(si, {}):
                cols.setdefault(c[1], []).append(c[0])
            for rows in cols.values():
                rows.sort()
            self._cols[si] = cols
            self._colkeys[si] = sorted(cols)

    def _find_cycles(self) -> list:
        """Strongly connected components with more than one node or a self-loop (Tarjan, iterative,
        successors enumerated lazily)."""
        index: dict = {}
        low: dict = {}
        onstack: set = set()
        stack: list = []
        self_loop: set = set()
        counter = 0
        sccs: list = []
        roots: list = [("c", si, r, c) for si in self._order for (r, c) in self._cells.get(si, {})]
        roots += [("n", key) for key in list(self._name_defs)]
        for root in roots:
            if root in index:
                continue
            index[root] = low[root] = counter
            counter += 1
            stack.append(root)
            onstack.add(root)
            work = [(root, self._succ(root))]
            while work:
                v, it = work[-1]
                advanced = False
                for w in it:
                    self.n_edges += 1
                    self.n_work += 1
                    if self.n_work > WORK_LIMIT:
                        raise GradingError(f"{self.key}: the dependency graph walk exceeded {WORK_LIMIT:,} steps "
                                           f"({self.n_formula_cells} formula cells): cannot search for cycles")
                    if w == v:
                        self_loop.add(v)
                        continue
                    if w not in index:
                        index[w] = low[w] = counter
                        counter += 1
                        stack.append(w)
                        onstack.add(w)
                        work.append((w, self._succ(w)))
                        advanced = True
                        break
                    if w in onstack and index[w] < low[v]:
                        low[v] = index[w]
                if advanced:
                    continue
                work.pop()
                if work:
                    u = work[-1][0]
                    if low[v] < low[u]:
                        low[u] = low[v]
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = stack.pop()
                        onstack.discard(w)
                        comp.append(w)
                        if w == v:
                            break
                    if len(comp) > 1 or v in self_loop:
                        sccs.append(comp)
        return sccs

    def _describe(self, comp: list) -> dict:
        per_sheet: dict = {}
        names = []
        for n in comp:
            if n[0] == "c":
                per_sheet.setdefault(n[1], set()).add((n[2], n[3]))
            elif n[0] == "n":
                names.append(n[1])
        sheets = []
        n_cells = 0
        for si in sorted(per_sheet):
            cells = per_sheet[si]
            n_cells += len(cells)
            sheets.append((self._sheet_name[si], group_cells_str(cells), len(cells)))
        parts = []
        for name, ranges, n in sheets:
            shown = ", ".join(location(name, rng) for rng in ranges[:MAX_CELLS_PER_CYCLE])
            more = f" (+{len(ranges) - MAX_CELLS_PER_CYCLE} more block(s))" if len(ranges) > MAX_CELLS_PER_CYCLE else ""
            parts.append(f"{shown}{more}")
        for key in names:
            dn = self._name_defs.get(key)
            parts.append(name_location(dn) if dn is not None else f"Name Manager: {key[0]}")
        return {"n_cells": n_cells, "n_names": len(names), "sheets": [s[0] for s in sheets],
                "text": "; ".join(parts), "first": (sheets[0][0], sheets[0][1][0]) if sheets else None}

    # ---------------------------------------------------------------- verdict
    def _settings_text(self) -> str:
        bits = [f"calcMode={self.calc_mode}", f"iterate={'on' if self.iterate else 'off'}"]
        if self.iterate:
            bits.append(f"iterateCount={self.iterate_count or '100 (default)'}")
            bits.append(f"iterateDelta={self.iterate_delta or '0.001 (default)'}")
        return ", ".join(bits)

    def finish(self) -> dict:
        t0 = time.perf_counter()
        stats = {"calc_mode": self.calc_mode, "manual": self.manual, "iterate": self.iterate,
                 "iterate_count": self.iterate_count, "iterate_delta": self.iterate_delta,
                 "full_calc_on_load": self.full_calc_on_load, "calc_pr": self.calc_pr,
                 "n_formula_cells": self.n_formula_cells, "array_member_cells": self.n_member_cells,
                 "data_table_cells_ignored": self.n_data_table_cells, "distinct_formula_texts": len(self._tpls),
                 "names_with_formulas": len(self._name_tab), "names_with_unknown_scope": self.n_names_bad_scope,
                 "external_refs_ignored": self.n_external_refs, "r1c1_refs_ignored": self.n_r1c1_refs,
                 "unresolved_sheet_refs": self.n_unresolved_sheet_refs, "mixed_range_endpoints": self.n_mixed_ranges,
                 "structured_refs_unresolved": self.n_struct_unresolved,
                 "structured_refs_approximated": self.n_struct_approximated,
                 "geometry_refs_skipped": self.n_geometry_refs, "fast_path_texts": self.n_fast_texts,
                 "dynamic_refs": self.dynamic}
        if self.manual:
            self.add_mistake("workbook", f"calculation mode is manual (xl/workbook.xml calcPr calcMode=\"manual\"): "
                                         f"displayed values may not reflect the inputs currently shown; the workbook "
                                         f"must be delivered in automatic calculation mode.")
            stats.update({"n_cycles": None, "decided_by": "code", "rule": 1,
                          "graph_s": round(time.perf_counter() - t0, 3)})
            return self.verdict("", "Calculation mode is manual (rule 1): {n} mistake.", stats)
        self._build_index()
        sccs = self._find_cycles()
        sccs.sort(key=len)
        cycles = [self._describe(comp) for comp in sccs]
        stats.update({"n_edges": self.n_edges, "n_cycles": len(cycles),
                      "largest_cycle": max((c["n_cells"] + c["n_names"] for c in cycles), default=0),
                      "cycles": [{"cells": c["n_cells"], "names": c["n_names"], "where": c["text"][:400]}
                                 for c in cycles[:MAX_CYCLES_LISTED]],
                      "global_name_unqualified_refs_skipped": self.n_global_name_unqualified,
                      "iteration_on_without_cycle": bool(self.iterate and not cycles),
                      "graph_s": round(time.perf_counter() - t0, 3)})
        if not cycles:
            stats.update({"decided_by": "code", "rule": 3})
            return self.verdict(
                f"Calculation mode is automatic ({self._settings_text()}) and the {self.n_formula_cells} formula "
                f"cells form no circular reference"
                + (" (iterative calculation is on without a cycle: recorded, not a mistake)" if self.iterate else "")
                + ".", "", stats)
        if not self.iterate:
            stats.update({"decided_by": "code", "rule": 4})
            for cyc in cycles:
                where = cyc["first"]
                loc = location(where[0], where[1]) if where else "Name Manager"
                self.add_mistake(loc, f"circular reference through {cyc['n_cells']} cell(s)"
                                      + (f" and {cyc['n_names']} defined name(s)" if cyc["n_names"] else "")
                                      + f" ({cyc['text']}); iterative calculation is off (calcPr iterate not set), "
                                      f"so Excel cannot resolve it: enable iterative calculation and state the "
                                      f"requirement on the cover, or break the loop.")
            return self.verdict("", f"{len(cycles)} circular reference(s) with iterative calculation off "
                                    f"(rule 4): {{n}} mistake(s).", stats)
        stats.update({"rule": 5})
        listed = "; ".join(f"cycle {i + 1} ({c['n_cells']} cell(s)): {c['text']}" for i, c in enumerate(cycles[:MAX_CYCLES_LISTED]))
        if len(cycles) > MAX_CYCLES_LISTED:
            listed += f"; and {len(cycles) - MAX_CYCLES_LISTED} more cycle(s)"
        note = LLM_QUESTION.format(cycles=listed, settings=self._settings_text())
        return self.undecided(
            f"{len(cycles)} circular reference(s) with iterative calculation on ({self._settings_text()}): whether "
            f"the workbook states the circularity and the iteration requirement is the LLM judge's call (rule 5).",
            note, stats, details={"n_cycles": len(cycles), "decided_by": "llm"})
