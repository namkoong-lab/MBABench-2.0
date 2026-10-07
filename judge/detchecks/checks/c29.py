"""29 Error Checks/Clean Name Manager.

Rule as implemented (see docs/checks/29.md):
Only VISIBLE defined names are graded ("counted"): hidden names (hidden="1") and built-in
names (_xlnm.*, plus the bare sheet-local forms in BARE_BUILTIN_NAMES) never count, whatever
their prefix (toy T3: an add-in name made visible counts like any other).  Each counted name
must be

  * USED: referenced by a consumer - a cell formula (worksheets and macro sheets, every
    sheet whatever its state; shared-formula children use the same names as their master),
    a conditional format (rule formulas and colour-scale/data-bar/icon thresholds, main and
    x14), a data validation (formula1/formula2, main and x14), a chart (series, categories,
    titles, data-label ranges), a sparkline, a shape text link, a form control (ctrlProps,
    legacy VML x:Fmla*), a table's calculated/totals formula, a pivot cache source, a
    slicer/timeline/query-table cache (Excel's own Slicer_X / NativeTimeline_X /
    ExternalData_n names), an in-workbook hyperlink target (COUNT_HYPERLINK_LOCATIONS), a
    literal INDIRECT("Name") / HYPERLINK("#Name") or a run-time reference whose name part is
    fixed text, INDIRECT("'"&A1&"'!Name") (COUNT_TEXT_LITERAL_REFS) - or by a name that is
    itself used (a used name's references are uses, through hidden names too), by any other
    name (TRANSITIVE_USE=False), or by a hidden / built-in name (HIDDEN_AND_BUILTIN_NAMES_USE).
    A LAMBDA name is used when called like a function (F(..)); an Excel built-in called as a
    function (YEAR(..)) is the built-in, never a name called Year.  Names are matched as
    tokens (no prefix confusion: F_x inside F_x_v1) and resolved with Excel's scope rules (a
    sheet-local name shadows a global one of the same spelling on its own sheet; Sheet!Name,
    [0]!Name).
  * NOT BROKEN: its definition contains #REF! (outside string literals), is empty, or
    (BROKEN_MISSING_SHEET) refers to a sheet the workbook does not have.

One mistake per offending name ("Name Manager: Name", or "Name Manager: 'Sheet'!Name" for a
sheet-local name).  No starting-file diff: the whole delivered workbook is graded.

No guess (raises GradingError), only when a counted name that is NOT broken would otherwise
be flagged unused (a broken name fails whatever an unread consumer does): the workbook has a
VBA project, ActiveX controls or Power Query (consumers this check cannot read), the name is
a macro name (function / vbProcedure / xlm), a name whose scope is unknown (bad localSheetId)
or whose definition cannot be parsed (hidden / built-in name with malformed text) may use it,
or a run-time text reference (INDIRECT / EVALUATE / HYPERLINK whose text is computed) can
produce its spelling (DYNAMIC_TEXT_GATE; decided from the shape of the argument only, see
_name_part).  Malformed text of a counted name's definition, or of a consumer formula the
check needs to parse, raises FormulaError at once.
"""
from __future__ import annotations

import re
from html import unescape
from typing import Optional

from ..core import formula as F
from ..core.package import local
from ..core.refs import location
from ..errors import GradingError
from .base import Check

# ----------------------------------------------------------------------------- policies
TRANSITIVE_USE = True                # a name used only by unused counted names (or a cycle) is unused;
                                     # False = any reference from another counted name counts
HIDDEN_AND_BUILTIN_NAMES_USE = True  # names that are not counted (hidden, built-in) are consumers (roots);
                                     # False = they pass use on only when they are used themselves
COUNT_TEXT_LITERAL_REFS = True       # INDIRECT("Name"), INDIRECT("Na"&"me"), HYPERLINK("#Name") and a
                                     # run-time reference with a fixed name part count as use
COUNT_HYPERLINK_LOCATIONS = True     # a cell hyperlink into the workbook (<hyperlink location="Name">) counts
BROKEN_MISSING_SHEET = True          # a definition naming a sheet the workbook lacks is broken
BROKEN_ERROR_CONSTANT = False        # a definition that is just an error constant (=#N/A, Excel's slicer
                                     # names) is NOT broken; it must still be used
BARE_BUILTIN_NAMES = ("print_area", "print_titles", "_filterdatabase")   # sheet-local names written
                                     # without the _xlnm. prefix (ChatGPT writer) are built-ins too
DYNAMIC_TEXT_GATE = True             # raise when a computed INDIRECT/EVALUATE/HYPERLINK text can produce an
                                     # otherwise unused name (cannot evaluate it)

CTRL_FORMULA_ATTRS = ("fmlaLink", "fmlaRange", "fmlaGroup", "fmlaTxbx")
OBJECT_PART_TYPES = ("slicercache", "timelinecache", "querytable")
_VML_FMLA = re.compile(r"<(?:[\w.\-]+:)?Fmla(Link|Range|Txbx|Group|Pict)\s*>(.*?)</(?:[\w.\-]+:)?Fmla\1\s*>",
                       re.S | re.I)
_DYN_FUNCS = ("INDIRECT", "EVALUATE") + (("HYPERLINK",) if COUNT_TEXT_LITERAL_REFS else ())
_TEXT_FUNCS_LOW = ("indirect", "hyperlink", "evaluate")
# One identifier run = the characters Excel allows in a defined name (core formula._RE_NAME).
# A name token in a formula is always a whole run (it is delimited by characters outside this
# class), so "does the text spell a name" is a set lookup of the runs (cost independent of the
# number of names).
_RUN = re.compile(r"[A-Za-z0-9_.\\?\u0080-￿]+")
# Storage prefixes Excel puts in front of a function name (core formula._FUNC_PREFIXES).  A call
# of a non-built-in function stored as _xludf.Rate( / _xll.Rate( is a reference to the NAME Rate
# (docs/formula.md), but the identifier run is '_xludf.rate', so the bare spelling is added too.
_STRIP_PREFIXES = ("_xlfn.", "_xlws.", "_xll.", "_xludf.", "_xleta.")
MAX_LISTED = 60
MAX_PATTERNS = 2000                  # distinct run-time reference shapes kept (more collapse to "any name")
# formula.parse is lru-cached; the check parses each text once, so use the uncached function
# (the cache only keeps memory alive on 600k-formula workbooks; same reason as checks/_fscan.py).
_parse = getattr(F.parse, "__wrapped__", F.parse)
_ANY = ".*"                          # the pattern of a reference text that can be any name
_MAYBE = -1                          # owner sentinel: a reference that may or may not be live


def _short(text: str, n: int = 100) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 3] + "..."


def _runs(low: str) -> set:
    """Identifier runs of a lower-cased formula text, plus, for a run starting with Excel's
    function-storage prefixes, the spelling with those prefixes stripped ('_xludf.rate' ->
    also 'rate').  A name token in a formula is always one of these, so 'does the text spell
    a name' is a set lookup."""
    out = set(_RUN.findall(low))
    for r in list(out):
        if r.startswith("_xl"):
            s = r
            stripped = True
            while stripped:
                stripped = False
                for p in _STRIP_PREFIXES:
                    if s.startswith(p):
                        s = s[len(p):]
                        stripped = True
                        break
            if s and s != r:
                out.add(s)
    return out


def _split_last_bang(s: str):
    """(qualifier or None, rest) split at the last '!' outside single quotes."""
    inq = False
    k = -1
    for i, ch in enumerate(s):
        if ch == "'":
            inq = not inq
        elif ch == "!" and not inq:
            k = i
    return (s[:k], s[k + 1:]) if k >= 0 else (None, s)


def _merge(pieces: list) -> list:
    """Join adjacent literal pieces, collapse adjacent computed pieces."""
    out = []
    for p in pieces:
        if out and p[0] == "L" and out[-1][0] == "L":
            out[-1] = ("L", out[-1][1] + p[1])
        elif out and p[0] == "W" and out[-1][0] == "W":
            continue
        else:
            out.append(p)
    return out


def _regex(pieces: list, anchored: bool = True) -> str:
    """Pattern (lower case, for re.fullmatch) of the texts a list of pieces can spell."""
    rx = "" if anchored else _ANY
    for k, t in pieces:
        rx += re.escape(t.lower()) if k == "L" else _ANY
    return rx


def _display(pieces: list) -> str:
    parts = []
    for k, t in pieces:
        parts.append('"' + t.replace('"', '""') + '"' if k == "L" else "…")
    return _short("&".join(parts), 60)


class C29(Check):
    number = 29
    key = "Error Checks/Clean Name Manager"
    needs_cells = True
    sheet_kinds = ("worksheet", "macrosheet")      # where cell formulas live (charts: wb.charts)

    # ------------------------------------------------------------------ setup
    def start(self, wb):
        super().start(wb)
        self.sheet_names = [s.name for s in wb.sheets]
        self._sheets_low = {s.lower() for s in self.sheet_names}
        self._bang_sheet = any("!" in s for s in self.sheet_names)
        every = list(wb.defined_names)
        # a localSheetId that names no sheet: scope unknown (reader rule); handled in _unscoped()
        self.unscoped = [d for d in every if d.scope_error is not None]
        self.raw = [d for d in every if d.scope_error is None]
        self.tab = F.NameTable(self.raw, sheet_names=self.sheet_names)
        self.names = self.tab.names
        self.pos = {id(n): i for i, n in enumerate(self.names)}
        self.kind = [self._kind(n.name, n.scope is not None, n.builtin, n.hidden) for n in self.names]
        self.counted = [i for i, k in enumerate(self.kind) if k == "counted"]
        self.used: dict[int, str] = {}             # counted index -> how it is used (first site)
        self.used_any: set = set()                 # every name index known to be used (any kind)
        self.broken: dict[int, list] = {}
        self.cand: set = set(self.counted)         # counted names not yet known to be used
        self.edges: dict[int, set] = {}            # name index -> names (any kind) its definition uses
        self.maybe_edges: dict[int, set] = {}      # uncounted name with an unparseable definition -> names it spells
        self.malformed: dict[int, str] = {}        # that name -> the parse error
        self.rev: dict[int, set] = {}
        self.maybe: dict[int, str] = {}            # name index -> why it may be used (scope unknown)
        self.by_spell: dict[str, set] = {}
        for i, n in enumerate(self.names):
            self.by_spell.setdefault(n.name.lower(), set()).add(i)
        self._fcache: dict = {}                    # parsed definitions
        self._spell: set = set()
        self._dirty = True
        self.hash_seen = False                     # some text / literal starts with '#' or '['
        self.dyn_patterns: dict = {}               # regex -> (where, display) of computed references
        self.dyn_patterns_hash: dict = {}          # the same for HYPERLINKs whose text starts computed
        self.dyn_maybe_hash: dict = {}             # fixed-name HYPERLINK targets that need a '#' text
        self.dyn_counts = {"literal": 0, "cannot_name": 0, "fixed_name": 0, "computed_name": 0,
                           "hyperlink_computed_start": 0, "evaluate": 0}
        self.n_dynamic = 0
        self.dynamic_sites: list = []
        self.n_formula_texts = 0
        self.n_parsed = 0
        self.n_sites_parsed = {}
        self.error_constant_names: list = []
        for d in self.unscoped:
            if self._kind(d.name, True, d.builtin, d.hidden) == "counted":
                raise GradingError(f"{self.key}: visible defined name {d.name!r}: {d.scope_error}; which sheet "
                                   f"it belongs to (and so which formulas use it) is unknown")
        if not self.counted:
            return                                 # nothing to grade: no other name can matter
        for i in list(self.counted):
            try:
                self.broken[i] = self._problems(i)
            except F.FormulaError as e:
                # Patrick 2026-10-05 (every attempt graded): a definition that cannot be parsed is skipped - the
                # name is not graded and its definition uses nothing
                self.note_default("unparsable_formula", f"{self._label(i)}: {e}")
                self.kind[i] = "unparsable"
                self.counted.remove(i)
                self.cand.discard(i)
        if not self.counted:
            return
        self._name_graph()
        self._unscoped()
        for graph in (self.edges, self.maybe_edges):
            for i, targets in graph.items():
                for t in targets:
                    self.rev.setdefault(t, set()).add(i)
        if HIDDEN_AND_BUILTIN_NAMES_USE:
            for i, targets in self.edges.items():
                if self.kind[i] != "counted" and (targets or self.maybe_edges.get(i)):
                    self._mark(i, f"the {self.kind[i]} name {self.names[i].name}")
        if not TRANSITIVE_USE:
            for i, targets in list(self.edges.items()):
                if self.kind[i] == "counted":
                    for t in targets:
                        self._mark(t, f"the name {self.names[i].name}")
        if self.cand:
            self._side_parts()

    @staticmethod
    def _kind(name: str, is_local: bool, builtin: bool, hidden: bool) -> str:
        if builtin or (is_local and name.lower() in BARE_BUILTIN_NAMES):
            return "built-in"
        if hidden:
            return "hidden"
        return "counted"

    def _label(self, i: int) -> str:
        n = self.names[i]
        return f"name '{n.name}'" + (f" (local to sheet '{n.scope}')" if n.scope else "")

    def _def(self, i: int):
        """Parsed definition of name i (once).  Raises FormulaError on malformed text."""
        f = self._fcache.get(i)
        if f is None:
            f = self._fcache[i] = self._parse_at(self.names[i].text, f"the definition of {self._label(i)}")
        return f

    def _def_lenient(self, i: int):
        """Parsed definition of an UNCOUNTED name i, or None when its text cannot be parsed
        (recorded in self.malformed; the verdict may still not depend on it)."""
        if i in self.malformed:
            return None
        f = self._fcache.get(i)
        if f is None:
            self.n_parsed += 1
            try:
                f = self._fcache[i] = _parse(self.names[i].text)
            except F.FormulaError as e:
                self.malformed[i] = str(e)
                return None
        return f

    def _problems(self, i: int) -> list:
        n = self.names[i]
        out = []
        if not (n.text or "").strip():
            return ["its definition is empty"]
        f = self._def(i)
        if f.ref_errors():
            out.append("its definition contains #REF! (it points at a deleted range or sheet)")
        if BROKEN_MISSING_SHEET:
            missing = []
            for o in f.operands:
                if o.external is not None or o.ambiguous_book:
                    continue
                for s in (o.sheet, o.sheet_end):
                    if s is not None and s.lower() not in self._sheets_low and s not in missing:
                        missing.append(s)
            if missing:
                out.append("it refers to sheet " + ", ".join(f"'{s}'" for s in missing)
                           + " which the workbook does not have")
        body = f.body.strip(" \t\r\n")
        while body.startswith("(") and body.endswith(")"):
            body = body[1:-1].strip(" \t\r\n")
        if body.upper() in F.ERROR_LITERALS and body.upper() != "#REF!":
            if BROKEN_ERROR_CONSTANT:
                out.append(f"its definition is just the error value {body.upper()}")
            else:
                self.error_constant_names.append(n.name)
        return out

    # ------------------------------------------------------------------ usage graph
    def _name_graph(self) -> None:
        """Edges of every name (any kind) whose definition can lead to a counted name.
        Fixed point: a name is relevant when a counted name is reachable from it; only
        definitions spelling a relevant name (or calling INDIRECT/EVALUATE/HYPERLINK) are
        parsed.  An UNCOUNTED name whose text cannot be parsed does not raise here: it gets a
        'maybe' edge to every name it spells (and, if it mentions INDIRECT/EVALUATE/HYPERLINK,
        an any-name pattern), so the check raises at verdict time only if one of those names
        would otherwise be called unused (a counted name's text always parses or raises, in
        _problems)."""
        relevant = set(self.counted)
        spell = {self.names[i].name.lower() for i in relevant}
        runs = {}
        changed = True
        while changed:
            changed = False
            for i, n in enumerate(self.names):
                if i not in self.edges and n.text:
                    if i not in runs:
                        low = n.text.lower()
                        runs[i] = (_runs(low), any(x in low for x in _TEXT_FUNCS_LOW))
                    r, dyn = runs[i]
                    if dyn or not spell.isdisjoint(r):
                        f = self._def(i) if self.kind[i] == "counted" else self._def_lenient(i)
                        if f is None:                  # unparseable uncounted definition: skipped, it uses nothing
                            self.edges[i] = set()      # (Patrick 2026-10-05: every attempt graded)
                            if self.kind[i] != "unparsable":
                                self.note_default("unparsable_formula", f"{self._label(i)}: "
                                                                        f"{self.malformed.get(i, 'malformed text')}")
                            continue
                        self.edges[i] = {t for t in self._uses(f, n.scope) if t != i}
                        self._note_dynamic(f, n.scope, f"the definition of {self._label(i)}", owner=i)
                        if not self.hash_seen and ('"#' in n.text or '"[' in n.text):
                            self.hash_seen = True
                if i not in relevant:
                    reach = self.edges.get(i, set()) | self.maybe_edges.get(i, set())
                    if reach and not reach.isdisjoint(relevant):
                        relevant.add(i)
                        spell.add(n.name.lower())
                        changed = True

    def _unscoped(self) -> None:
        """Names whose localSheetId names no sheet.  Raise only when that can change the
        verdict: a counted one (cannot be resolved), one spelled like a counted name (could
        shadow it on its unknown sheet); otherwise its definition is resolved under every
        possible scope: targets in every reading are used (when it is a consumer), targets
        in only some readings are 'maybe' (raise if they end up unused)."""
        if not self.unscoped:
            return
        counted_spell = {self.names[i].name.lower() for i in self.counted}
        all_spell = {n.name.lower() for n in self.names}
        for d in self.unscoped:
            k = self._kind(d.name, True, d.builtin, d.hidden)
            if d.name.lower() in counted_spell:
                raise GradingError(f"{self.key}: {k} defined name {d.name!r}: {d.scope_error}; it may shadow "
                                   f"the visible name of the same spelling on its unknown sheet")
            low = (d.text or "").lower()
            if not low or (all_spell.isdisjoint(_runs(low)) and not any(x in low for x in _TEXT_FUNCS_LOW)):
                continue
            try:
                f = self._parse_at(d.text, f"the definition of {k} name '{d.name}'")
            except F.FormulaError as e:
                self.note_default("unparsable_formula", str(e))      # skipped (Patrick 2026-10-05)
                continue
            readings = []
            for sc in [None] + self.sheet_names:
                readings.append({self.pos[id(h)] for h in F.names_used_by(
                    f, sc, self.tab, include_indirect_literals=COUNT_TEXT_LITERAL_REFS)})
            if any(x in low for x in _TEXT_FUNCS_LOW):
                self._note_dynamic(f, None, f"the definition of {k} name '{d.name}'",
                                   owner=None if HIDDEN_AND_BUILTIN_NAMES_USE else _MAYBE)
            certain = set.intersection(*readings) if readings else set()
            possible = set.union(*readings) if readings else set()
            why = f"the {k} name '{d.name}' ({d.scope_error}) refers to it under some scope"
            if HIDDEN_AND_BUILTIN_NAMES_USE:
                for t in certain:
                    self._mark(t, f"the {k} name {d.name}")
                for t in possible - certain:
                    self._maybe(t, why)
            else:
                for t in possible:
                    self._maybe(t, why)

    def _maybe(self, i: int, why: str) -> None:
        """Name i may be used (by a reference we cannot place); so may everything it reaches
        (closure taken at verdict time, when the name graph is complete)."""
        self.maybe.setdefault(i, why)

    def _use_from(self, owner, t: int, how: str) -> None:
        """A definite reference to name t from a consumer (owner None), from the definition of
        name `owner` (an edge), or from a reference that may or may not be live (_MAYBE)."""
        if owner is None:
            self._mark(t, how)
        elif owner == _MAYBE:
            self._maybe(t, how)
        else:
            self.edges.setdefault(owner, set()).add(t)

    def _mark(self, i: int, how: str) -> None:
        """i is used (by `how`); everything a used name's definition references is used too."""
        stack = [(i, how)]
        while stack:
            j, h = stack.pop()
            if j in self.used_any:
                continue
            self.used_any.add(j)
            self._dirty = True
            if self.kind[j] == "counted":
                self.used[j] = h
                self.cand.discard(j)
            via = f"the {'used' if self.kind[j] == 'counted' else self.kind[j]} name {self.names[j].name}"
            for t in self.edges.get(j, ()):
                if t not in self.used_any:
                    stack.append((t, via))
            for t in self.maybe_edges.get(j, ()):      # a used name whose text cannot be parsed
                self._maybe(t, f"the {self.kind[j]} name '{self.names[j].name}' spells it in a definition "
                               f"this check cannot parse ({self.malformed.get(j, 'malformed text')})")

    def _refresh(self) -> None:
        """Spellings worth finding in a formula: the remaining candidates and every not-yet-used
        name from which a candidate is reachable (a hidden helper when hidden names are not
        roots)."""
        self._dirty = False
        reach = set(self.cand)
        stack = list(self.cand)
        while stack:
            t = stack.pop()
            for s in self.rev.get(t, ()):
                if s not in reach and s not in self.used_any:
                    reach.add(s)
                    stack.append(s)
        self._spell = {self.names[i].name.lower() for i in reach}

    def _may_use(self, text: str) -> bool:
        """Conservative pre-filter: False guarantees the text references no remaining candidate
        (or a name leading to one) and has no INDIRECT / EVALUATE / HYPERLINK (always parsed:
        INDIRECT("Na"&"me") and run-time references).  Also notes '#'/'[' string literals."""
        if self._dirty:
            self._refresh()
        if not self.hash_seen and ('"#' in text or '"[' in text):
            self.hash_seen = True
        low = text.lower()
        if "indirect" in low or "hyperlink" in low or "evaluate" in low:
            return True
        return bool(self._spell) and not self._spell.isdisjoint(_runs(low))

    def _parse_at(self, text: str, where: str):
        self.n_parsed += 1
        try:
            return _parse(text)
        except F.FormulaError as e:
            raise F.FormulaError(f"{self.key}: cannot parse the formula text of {where} "
                                 f"(={_short(text, 120)}): {e}") from e

    def _uses(self, f, scope: Optional[str]) -> list:
        """Names (any kind) a parsed formula references directly (Excel scope rules)."""
        out = []
        for hit in F.names_used_by(f, scope, self.tab, include_indirect_literals=COUNT_TEXT_LITERAL_REFS):
            i = self.pos.get(id(hit))
            if i is not None:
                out.append(i)
        return out

    def _site(self, text: Optional[str], scope: Optional[str], where: str, kind: str) -> None:
        """A consumer formula text (cell, CF, DV, chart ...) living on sheet `scope`."""
        if text and self.cand and self._may_use(text):
            self._site_parse(text, scope, where, kind)

    def _site_parse(self, text: str, scope: Optional[str], where: str, kind: str) -> None:
        try:
            f = self._parse_at(text, where)
        except F.FormulaError as e:
            # Patrick 2026-10-05 (every attempt graded): a consumer formula that cannot be parsed is skipped
            self.note_default("unparsable_formula", str(e))
            return
        self.n_sites_parsed[kind] = self.n_sites_parsed.get(kind, 0) + 1
        for i in self._uses(f, scope):
            self._mark(i, where)
        self._note_dynamic(f, scope, where, owner=None)

    def _resolve_text(self, text: Optional[str], scope: Optional[str]) -> Optional[int]:
        """A plain reference text that is not a formula (hyperlink location, pivot source
        name, object cache name): the name (any kind) it designates, else None.  Never raises:
        a text that is not 'Name' / 'Sheet!Name' / '[0]!Name' designates no name."""
        s = (text or "").strip()
        if s.startswith("#"):
            s = s[1:]
        q, ident = _split_last_bang(s)
        ident = ident.strip()
        if not ident or not F.is_valid_name(ident):
            return None
        qsheet, wbs = None, False
        if q is not None:
            try:
                qq = F.parse_qualifier(q)
            except F.FormulaError:
                return None
            if qq.external is not None:
                return None
            qsheet, wbs = qq.sheet, qq.workbook_scoped
        hit = self.tab.resolve(ident, scope, qsheet, wbs)
        return self.pos.get(id(hit)) if hit is not None else None

    # ------------------------------------------------------------------ run-time text references
    def _note_dynamic(self, f, scope: Optional[str], where: str, owner: Optional[int]) -> None:
        """INDIRECT / EVALUATE / HYPERLINK calls whose first argument is not plain literal text
        (literal ones are resolved by names_used_by).  owner: the name whose definition this
        is (a fixed-name reference becomes an edge of that name), None for a consumer."""
        for c in f.functions:
            if c.local or c.name not in _DYN_FUNCS or not c.arg_spans:
                continue
            got = self._pieces(f, c)
            if got is None:
                continue
            pieces, placeholder = got
            computed = any(p[0] == "W" for p in pieces)
            if not computed and not placeholder:
                self.dyn_counts["literal"] += 1
                if c.name == "EVALUATE":
                    self._literal_evaluate(pieces[0][1] if pieces else "", scope, where, owner)
                continue
            self.n_dynamic += 1
            if len(self.dynamic_sites) < 5:
                self.dynamic_sites.append(f"{where}: {c.name}({_display(pieces)})")
            site = f"{c.name}({_display(pieces)}) in {where}"
            if c.name == "EVALUATE":                   # formula text: a name can sit anywhere in it
                self.dyn_counts["evaluate"] += 1
                self._add_pattern(_ANY, site, False)
                continue
            deferred = False
            if c.name == "HYPERLINK":                  # only '#...' / '[book]...' texts reach a name
                if pieces[0][0] == "L":
                    s = pieces[0][1].lstrip()
                    if s.startswith("#"):
                        pieces = _merge([("L", s[1:])] + pieces[1:])
                    elif s.startswith("["):
                        k = s.find("]")
                        pieces = _merge([("W", None) if k < 0 else ("L", s[k + 1:])] + pieces[1:])
                    else:
                        self.dyn_counts["cannot_name"] += 1   # a file / web address
                        continue
                else:
                    deferred = True                    # the computed start must supply the '#'
                    self.dyn_counts["hyperlink_computed_start"] += 1
            kind, val = self._name_part(pieces)
            if kind == "fixed" and not self._ref_names(val):
                kind = "none"                          # a cell, range or no reference at all
            if kind == "none":
                self.dyn_counts["cannot_name"] += 1
            elif kind == "fixed":
                self.dyn_counts["fixed_name"] += 1
                self._fixed_name_ref(val, pieces, scope, site, owner, deferred)
            else:
                self.dyn_counts["computed_name"] += 1
                for rx in val:
                    self._add_pattern(rx, site, deferred)

    def _add_pattern(self, rx: str, site: str, deferred: bool) -> None:
        if not DYNAMIC_TEXT_GATE:
            return
        bucket = self.dyn_patterns_hash if deferred else self.dyn_patterns
        if rx in bucket:
            return
        if len(bucket) >= MAX_PATTERNS:
            rx = _ANY
            if rx in bucket:
                return
        bucket[rx] = site

    def _literal_evaluate(self, text: str, scope: Optional[str], where: str, owner: Optional[int]) -> None:
        """EVALUATE("Rate*2"): the literal is formula text; its names are used."""
        if not COUNT_TEXT_LITERAL_REFS:
            return
        try:
            g = _parse(text)
        except F.FormulaError as e:
            # EVALUATE text that cannot be parsed is skipped: it uses no name (Patrick 2026-10-05: every attempt
            # graded; until then any name it spelled could be used)
            self.note_default("unparsable_formula", f"EVALUATE(\"{_short(text, 40)}\") in {where}: {e}")
            return
        for t in self._uses(g, scope):
            self._use_from(owner, t, f"{where} (EVALUATE text)")

    def _pieces(self, f, c):
        """The first argument of call c split at top-level '&': ('L', text) for a string
        literal, ('W', None) for anything computed.  ADDRESS(...) is replaced by a cell-shaped
        placeholder when its result is sure to be '$'-absolute A1 (no name can contain '$') or
        carries a sheet ('...!' + an address).  Returns (pieces, used_placeholder) or None."""
        a, b = c.arg_spans[0]
        toks = [t for t in f.tokens if a <= t[2] and t[3] <= b and t[0] != "ws"]
        if not toks:
            return None
        groups, cur, depth = [], [], 0
        for t in toks:
            k = t[0]
            if k in ("(", "{"):
                depth += 1
            elif k in (")", "}"):
                depth -= 1
            if depth == 0 and k == "op" and t[1] == "&":
                groups.append(cur)
                cur = []
            else:
                cur.append(t)
        groups.append(cur)
        out, placeholder = [], False
        for g in groups:
            if len(g) == 1 and g[0][0] == "str":
                out.append(("L", g[0][1][1:-1].replace('""', '"')))
                continue
            if g and g[0][0] == "func" and g[-1][0] == ")":
                call = next((x for x in f.functions if x.start == g[0][2] and x.end == g[-1][3]), None)
                if call is not None and call.name == "ADDRESS" and not call.local:
                    shape = self._address_shape(f, call)
                    if shape is not None:
                        out.extend(shape)
                        placeholder = True
                        continue
            out.append(("W", None))
        return _merge(out), placeholder

    @staticmethod
    def _address_shape(f, call):
        n = call.n_args

        def arg(i):
            return f.arg_text(call, i).strip() if i < n else None
        if n >= 5 and arg(4):
            return [("W", None), ("L", "!$A$1")]         # 'sheet'!<address>: never a name
        absn, a1 = arg(2), arg(3)
        if (absn is None or absn in ("1", "2", "3")) and (a1 is None or a1.upper() in ("TRUE", "1")):
            return [("L", "$A$1")]                       # $B$1 / B$1 / $B1
        return None

    def _name_part(self, pieces: list):
        """What the reference text built from `pieces` can name.  The name part of a
        reference is the text after its last '!' (names never contain '!').  Model: a
        computed piece is unknown text that may hold a whole qualifier ('Sheet!...'), but
        does not add range / union / intersection operators (':', ',', ' ').
        Returns ('none', None) | ('fixed', text after the last '!') | ('pattern', [regex])."""
        idx_w = [k for k, p in enumerate(pieces) if p[0] == "W"]
        if not idx_w:                                     # only ADDRESS placeholders + literals
            whole = "".join(p[1] for p in pieces)
            return "fixed", whole[whole.rfind("!") + 1:]
        j = idx_w[-1]
        tail = pieces[j + 1][1] if j + 1 < len(pieces) else ""
        if "!" in tail:
            s = tail[tail.rfind("!") + 1:]
            return ("none", None) if "'" in s else ("fixed", s)
        bang = [k for k in range(j) if pieces[k][0] == "L" and "!" in pieces[k][1]]
        alts = []
        if bang:
            k = bang[-1]
            suffix = pieces[k][1][pieces[k][1].rfind("!") + 1:]
            if "'" in suffix:
                alts.append(_regex([("L", tail)], anchored=False))
            else:
                alts.append(_regex([("L", suffix)] + pieces[k + 1:]))
                if self._bang_sheet:                      # that '!' may sit inside a quoted sheet name
                    alts.append(_regex([("L", tail)], anchored=False))
        else:
            alts.append(_regex(pieces))                   # the whole text is the name
            p0 = pieces[0]
            lead = p0[1].lstrip("'").lower() if p0[0] == "L" else ""
            if (p0[0] == "W" or p0[1].startswith("[") or not lead
                    or any(s.lower().startswith(lead) for s in self.sheet_names)):
                alts.append(_regex([("L", tail)], anchored=False))   # qualifier ends inside a computed piece
        return "pattern", alts

    @staticmethod
    def _ref_names(s: str) -> list:
        """Names a fixed reference text spells ('Rate', 'Start:End'); [] for cells, ranges or
        text that is no reference (INDIRECT of it is #REF!)."""
        s = (s or "").strip()
        if not s:
            return []
        try:
            g = _parse(s)
        except F.FormulaError:
            return []                                     # run-time text, not stored formula text
        o = g.pure_reference(transparent=())
        return list(o.names) if o is not None else []

    def _fixed_name_ref(self, s: str, pieces: list, scope, site: str, owner, deferred: bool) -> None:
        """A run-time reference whose name part is fixed text: it uses that name on whatever
        sheet the computed qualifier names."""
        spells = self._ref_names(s)
        if not spells or not COUNT_TEXT_LITERAL_REFS:
            return
        qualified = any(p[0] == "L" and "!" in p[1] for p in pieces)
        for sp in spells:
            targets = set()
            if qualified or any(p[0] == "W" for p in pieces):
                for x in self.sheet_names:
                    h = self.tab.resolve(sp, None, x)
                    if h is not None:
                        targets.add(self.pos[id(h)])
                h = self.tab.resolve(sp, None, None, True)
                if h is not None:
                    targets.add(self.pos[id(h)])
            else:
                h = self.tab.resolve(sp, scope)
                if h is not None:
                    targets.add(self.pos[id(h)])
            if not targets:
                continue
            if deferred:                               # live only if some text supplies the leading '#'
                if DYNAMIC_TEXT_GATE:
                    for t in targets:
                        self.dyn_maybe_hash.setdefault(t, site)
            elif len(targets) == 1:
                self._use_from(owner, next(iter(targets)), f"{site} (its text always ends in the name)")
            elif DYNAMIC_TEXT_GATE:
                for t in targets:
                    self._maybe(t, f"{site} names '{sp}' on a sheet computed at run time, and more than "
                                   f"one name has that spelling")

    # ------------------------------------------------------------------ workbook-level consumers
    def _side_parts(self) -> None:
        wb = self.wb
        for ch in wb.charts:
            for t in ch.formulas:
                self._site(t, ch.sheet, f"chart {ch.part}" + (f" on '{ch.sheet}'" if ch.sheet else ""), "chart")
        for dr in wb.drawings:
            for t in dr.textlinks:
                self._site(t, dr.sheet, f"a shape text link on '{dr.sheet}'", "shape")
        for pc in wb.pivot_caches:
            if pc.source_name and not pc.source_target:
                i = self._resolve_text(pc.source_name, pc.source_sheet)
                if i is not None:
                    self._mark(i, f"pivot cache {pc.part}")
        for si in wb.sheets:
            if not si.part:
                continue
            if si.kind == "worksheet":
                for tb in wb.tables(si):
                    for col in tb.columns:
                        for t in (col.get("calculated"), col.get("totals_formula")):
                            self._site(t, si.name, f"table {tb.display_name or tb.name} on '{si.name}'", "table")
            for rel in wb.rels_of_type(si.part, "ctrlProp"):
                root = wb.xml(rel.part) if rel.part else None
                if root is None:
                    continue
                for k, v in root.attrib.items():
                    if local(k) in CTRL_FORMULA_ATTRS and v:
                        self._site(v, si.name, f"a form control on '{si.name}' ({rel.part})", "control")
            for rel in wb.rels_of_type(si.part, "vmlDrawing"):
                data = wb.read(rel.part) if rel.part else None
                if not data or b"Fmla" not in data:
                    continue
                for _k, v in _VML_FMLA.findall(data.decode("utf-8", "replace")):
                    v = unescape(v).strip()
                    if v:
                        self._site(v, si.name, f"a form control on '{si.name}' ({rel.part})", "control")
        if not self.cand:
            return
        for part, types in wb.reachable_parts().items():
            tl = {t.lower() for t in types}
            if not tl & set(OBJECT_PART_TYPES):
                continue
            root = wb.xml(part)
            nm = root.get("name") if root is not None else None
            if not nm:
                continue
            for i in self.counted:                       # Excel ties the object to its name by spelling
                if self.names[i].name.lower() == nm.lower():
                    self._mark(i, f"the Excel object {part} ({local(root.tag)})")

    # ------------------------------------------------------------------ streaming
    def wants_sheet(self, info) -> bool:
        return bool(self.cand) and super().wants_sheet(info)

    def sheet_start(self, head):
        if not self.cand:
            return False
        return None

    def cell(self, cell):
        if not self.cand:
            return
        f = cell.formula
        if f is None:
            if not self.hash_seen and cell.t == "inlineStr" and cell.raw and cell.raw.lstrip()[:1] in ("#", "["):
                self.hash_seen = True
            return
        t = f.text
        if not t:            # shared child (uses its master's names), data-table anchor, <f/> marker
            return
        self.n_formula_texts += 1
        if self._may_use(t):
            self._site_parse(t, cell.sheet, location(cell.sheet, cell.ref), "cell")

    def sheet_end(self, head, tail):
        if not self.cand:
            return
        sh = head.name
        for cf in tail.conditional_formats:
            where = f"a conditional format on {location(sh, _short(','.join(cf.sqref.split()), 60) or '?')}"
            for r in cf.rules:
                for t in r.formulas:
                    self._site(t, sh, where, "conditional format")
                for _typ, val in r.cfvos:
                    self._site(val, sh, where + " (threshold)", "conditional format")
        for dv in tail.data_validations:
            where = f"a data validation on {location(sh, _short(','.join(dv.sqref.split()), 60) or '?')}"
            self._site(dv.formula1, sh, where, "data validation")
            self._site(dv.formula2, sh, where, "data validation")
        for g in tail.sparkline_groups:
            for t in g.formulas:
                self._site(t, sh, f"a sparkline on '{sh}'", "sparkline")
        if COUNT_HYPERLINK_LOCATIONS:
            for h in tail.hyperlinks:
                tgt = h.location or (h.target if h.target and h.target.startswith("#") else None)
                i = self._resolve_text(tgt, sh)
                if i is not None:
                    self._mark(i, f"the hyperlink in {location(sh, h.ref)}")

    # ------------------------------------------------------------------ verdict
    def _gates(self, unused: list) -> None:
        """Consumers this check cannot read: raise instead of calling the names unused.
        Called only with names whose verdict depends on it (unused and not broken: a broken
        name fails whatever VBA / ActiveX / Power Query / run-time text might do)."""
        problem = self._gate_problem(unused)
        if problem:
            raise GradingError(f"{self.key}: {problem}")

    def _gate_problem(self, unused: list) -> Optional[str]:
        """Why the names in `unused` cannot be called unused, or None when they can."""
        wb = self.wb
        if not unused:
            return None
        labels = ", ".join(self.names[i].name for i in unused[:10]) + (" ..." if len(unused) > 10 else "")
        macro = [i for i in unused if self.raw[i].function or self.raw[i].vb_procedure or self.raw[i].xlm]
        if macro:
            return (f"visible macro name(s) {', '.join(self.names[i].name for i in macro)} (function/vbProcedure/xlm) "
                    f"are not referenced by any formula; macros are invoked from code or controls this check "
                    f"cannot read")
        if wb.has_vba:
            return (f"visible name(s) {labels} look unused, but the workbook has a VBA/XLM project "
                    f"({', '.join(wb.vba_parts[:3])}) that may use them; VBA code is not read")
        activex = [n for n in wb.names if n.lower().startswith("xl/activex/")]
        if activex:
            return (f"visible name(s) {labels} look unused, but the workbook has ActiveX controls ({activex[0]}) "
                    f"whose linked cells / list ranges are not read")
        pq = self._power_query()
        if pq:
            return (f"visible name(s) {labels} look unused, but the workbook has Power Query ({pq}); its M code "
                    f"(which can read named ranges) is not decoded")
        roots = dict(self.maybe)
        pats = dict(self.dyn_patterns)
        if self.dyn_patterns_hash or self.dyn_maybe_hash:
            if not self.hash_seen:
                self.hash_seen = any(s.lstrip()[:1] in ("#", "[") for s in wb.shared_strings if s)
            if self.hash_seen:
                pats.update(self.dyn_patterns_hash)
                for t, site in self.dyn_maybe_hash.items():
                    roots.setdefault(t, f"{site} may name it")
        maybe = {}
        for r, why in roots.items():                     # everything a maybe-used name reaches
            stack = [r]
            while stack:
                j = stack.pop()
                if j in maybe:
                    continue
                maybe[j] = why
                stack.extend(self.edges.get(j, ()))
                stack.extend(self.maybe_edges.get(j, ()))
        unsure = [f"{self.names[i].name} ({maybe[i]})" for i in unused if i in maybe]
        if unsure:
            return (f"name(s) {'; '.join(unsure[:5])} look unused, but a reference whose scope or sheet is "
                    f"unknown, or whose text cannot be parsed, may use them; this check does not guess")
        if not pats:
            return None
        compiled = [(re.compile(rx, re.S), site) for rx, site in pats.items()]
        reach = []
        for i in unused:
            nm = self.names[i].name
            if not F.is_valid_name(nm):
                continue                                  # spelled like a cell: a text reference is the cell
            low = nm.lower()
            site = next((s for rx, s in compiled if rx.fullmatch(low)), None)
            if site is not None:
                reach.append(f"{nm} ({site} can produce it)")
        if reach:
            return (f"name(s) {'; '.join(reach[:5])} look unused, but {self.n_dynamic} reference(s) are built "
                    f"from text at run time and could produce them; this check does not evaluate "
                    f"INDIRECT/EVALUATE/HYPERLINK text")
        return None

    def _power_query(self) -> Optional[str]:
        wb = self.wb
        for r in wb.rels_of_type(wb.main_part or "", "connections"):
            data = wb.read(r.part) if r.part else None
            if data and (b"Microsoft.Mashup.OleDb" in data or b"$Workbook$" in data):
                return f"connection in {r.part}"
        for n in wb.names:
            nl = n.lower()
            if nl.startswith("customxml/item") and nl.endswith(".xml") and "props" not in nl:
                if b"DataMashup" in (wb.read(n) or b"")[:4000]:
                    return f"DataMashup in {n}"
        return None

    def _location(self, i: int) -> str:
        n = self.names[i]
        return f"Name Manager: {location(n.scope, n.name) if n.scope else n.name}"

    def finish(self) -> dict:
        names = self.names
        n_hidden = sum(1 for k in self.kind if k == "hidden")
        n_builtin = sum(1 for k in self.kind if k == "built-in")
        unused = [i for i in self.counted if i not in self.used]
        # the gates decide only names whose verdict depends on them: unused AND not broken (a
        # broken name fails whatever an unread consumer might do - review finding
        # 29-rr-gate-ignores-broken)
        self._gates([i for i in unused if not self.broken.get(i)])
        bad = []
        for i in self.counted:
            probs = self.broken.get(i, [])
            is_unused = i not in self.used
            if not probs and not is_unused:
                continue
            bad.append(i)
            n = names[i]
            what = (f"Visible defined name '{n.name}'" + (f" (local to sheet '{n.scope}')" if n.scope else "")
                    + f" = {_short(n.text)}")
            parts = []
            if probs:
                parts.append(f"{what} is broken: {'; '.join(probs)}")
            if is_unused:
                referrers = [names[j].name for j, ts in self.edges.items() if i in ts and j != i]
                why = ("unused: only names that are themselves unused refer to it (" + ", ".join(referrers[:5]) + ")"
                       if referrers else
                       "unused: no cell formula, conditional format, data validation, chart, sparkline, "
                       "control, table, pivot, hyperlink or used name refers to it")
                if probs and self._gate_problem([i]) is not None:
                    # broken, and whether it is used cannot be decided (VBA, run-time text ...):
                    # say what was read without claiming it is unused
                    why = ("not referenced by any cell formula, conditional format, data validation, chart, "
                           "sparkline, control, table, pivot, hyperlink or used name this check reads (whether "
                           "code or run-time text uses it is not decided: it is broken either way)")
                parts.append(("It is also " if probs else f"{what} is ") + why)
                vp = F.name_validity_problem(n.name)
                if vp:
                    parts.append(f"Its spelling is not a usable name ({vp}), so formulas that write it "
                                 f"refer to that reference, not to the name")
            self.add_mistake(self._location(i), ". ".join(parts) + ". Delete or repair it.")
        stats = {
            "n_names": len(names) + len(self.unscoped), "n_counted": len(self.counted),
            "n_hidden_not_counted": n_hidden, "n_builtin_not_counted": n_builtin,
            "n_unknown_scope_not_counted": len(self.unscoped),
            "n_used": len(self.counted) - len(unused),
            "n_unused": len(unused), "n_broken": sum(1 for i in self.counted if self.broken.get(i)),
            "n_formula_texts_scanned": self.n_formula_texts, "n_formula_texts_parsed": self.n_parsed,
            "sites_parsed": dict(self.n_sites_parsed), "n_dynamic_text_refs": self.n_dynamic,
            "dynamic_text_refs": dict(self.dyn_counts),
            "error_constant_names_not_broken": self.error_constant_names[:MAX_LISTED],
            "counted_names": [{"name": names[i].name, "scope": names[i].scope,
                               "used_by": self.used.get(i), "problems": self.broken.get(i, [])}
                              for i in self.counted[:MAX_LISTED]],
            "policy": {"transitive_use": TRANSITIVE_USE, "hidden_and_builtin_names_use": HIDDEN_AND_BUILTIN_NAMES_USE,
                       "count_text_literal_refs": COUNT_TEXT_LITERAL_REFS,
                       "count_hyperlink_locations": COUNT_HYPERLINK_LOCATIONS,
                       "broken_missing_sheet": BROKEN_MISSING_SHEET, "broken_error_constant": BROKEN_ERROR_CONSTANT,
                       "bare_builtin_names": list(BARE_BUILTIN_NAMES), "dynamic_text_gate": DYNAMIC_TEXT_GATE},
        }
        listing = ", ".join(names[i].name for i in bad[:8]) + (" ..." if len(bad) > 8 else "")
        if not self.counted:
            ok = (f"No visible defined names to check ({n_hidden} hidden and {n_builtin} built-in "
                  f"name(s) are not counted).")
        else:
            ok = (f"All {len(self.counted)} visible defined name(s) are used and unbroken ({n_hidden} hidden and "
                  f"{n_builtin} built-in name(s) are not counted).")
        return self.verdict(ok, f"{{n}} of {len(self.counted)} visible defined name(s) are unused or broken: "
                                f"{listing}.", stats)
