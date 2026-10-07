"""Shared plumbing for the formula-scanning checks 80, 87 and 95 (group "formulascan").

Not a check itself (not registered).  It streams every cell formula once and hands each
formula TEXT to the subclass exactly once per shared group:

* normal / array / shared-master formulas: ``classify(text, sheet)`` is called with the
  stored text;
* shared-formula children carry no text: they inherit the master's result.  This is exact
  for the three checks: which functions are called (80), whether a reference is a whole
  column/row (87; a whole column moves only sideways and stays whole) and which workbook a
  reference points to (95) never change when Excel copies a formula.  (One corner: a copy
  pushed off the grid becomes #REF!; see each check's doc.  87 re-checks the copies of a
  master whose bounded range can reach the sheet edge when copied: C87.cell.)  A child
  whose master has not been seen earlier on the sheet raises GradingError, as the reader
  does.
* data-table anchors (no text) and empty <f/> markers go to ``other_formula``.

Each finding is a pair (key, detail): cells are grouped into rectangles per key on each
sheet (core.refs.group_cells) and described with the detail of the rectangle's top-left
cell.  CF formulas (rule formulas and colour-scale / data-bar / icon-set threshold
formulas, main and x14), DV formulas (main and x14) and defined names are classified with
the same function.  Mistakes are buffered and emitted in finish(): workbook-level ones
(names, links) first, then sheets in tab order.

Defined names: ``set_flagged_names`` switches on usage tallies (which formula sites use a
flagged name directly or through other names).  80/87 count a flagged name only when it
is used; the tallies also feed the mistake descriptions.  ``name_table`` is the core
NameTable over every name whose scope is readable, plus prefix-less aliases of built-in
names (a formula may write Print_Area for _xlnm.Print_Area).  A name whose localSheetId
names no sheet is left out; it matters only when a flagged name's use depends on it
(``name_used`` then raises).
"""
from __future__ import annotations

from typing import Optional

from ..core import formula as F
from ..core.package import local
from ..core.refs import group_cells, location, range_to_str
from ..errors import GradingError
from .base import Check

EMPTY = frozenset()
MAX_EXAMPLES = 3
# formula.parse is lru-cached (4,096 entries).  The checks parse each cell text once, so the cache
# only keeps memory alive: on a 35 MB attempt with 662k distinct OFFSET formulas (gpt6 1476) it
# held ~510 MB (peak 678 MB -> 167 MB without it, and 40 s -> 32 s).  Same function, no cache.
parse_text = getattr(F.parse, "__wrapped__", F.parse)
MAX_DETAILS = 20000          # per sheet: flagged cells whose own detail (example text) is kept
# Usage tallies of a flagged name stop parsing once it has this many uses (the verdict needs one;
# the description shows up to MAX_EXAMPLES): later texts that can only reach saturated names are not
# parsed, and the description then says "at least N".  (A 60k-formula workbook whose every formula
# used a flagged name took 14 s / 270 MB for the tallies alone.)
USE_TALLY_SATURATION = MAX_EXAMPLES
# form-control (xl/ctrlProps) attributes that hold formulas
CTRL_FORMULA_ATTRS = ("fmlaLink", "fmlaRange", "fmlaTxbx", "fmlaGroup")


def cf_rule_texts(rule) -> list:
    """Every formula text of a CF rule: its <formula>/<xm:f> texts, then the threshold values
    of a colour scale / data bar / icon set (<cfvo val=..> or <x14:cfvo><xm:f>): Excel stores a
    threshold of type num / percent / percentile / formula as a formula (=$C$6, =MAX(..))."""
    out = [t for t in rule.formulas if t]
    out.extend(v for _, v in rule.cfvos if v)
    return out


def ctrl_formulas(wb):
    """(sheet name, ctrlProp part, attribute, formula text) of every form control of the
    workbook (xl/ctrlProps parts related to a sheet: combo-box input ranges, cell links ...)."""
    for si in wb.sheets:
        if not si.part:
            continue
        for rel in wb.rels_of_type(si.part, "ctrlProp"):
            root = wb.xml(rel.part) if rel.part else None
            if root is None:
                continue
            for a in CTRL_FORMULA_ATTRS:
                t = next((v for k_, v in root.attrib.items() if local(k_) == a), None)
                if t:
                    yield si.name, rel.part, a, t


# ----------------------------------------------------------------------------- labels
def name_scope(dn) -> Optional[str]:
    """Sheet scope of a reader DefinedName; the literal '?' when its localSheetId names no
    sheet (the reader's scope raises then; a label must not)."""
    try:
        return dn.scope
    except GradingError:
        return "?"


def name_location(dn) -> str:
    """Location of a defined-name finding.  'Name Manager: Rate', or for a sheet-local name
    "Name Manager: 'My Sheet'!Rate".  A sheet name can never contain ':', so this is never
    mistaken for a sheet."""
    sc = name_scope(dn)
    if sc is None:
        return f"Name Manager: {dn.name}"
    return f"Name Manager: {location(sc, dn.name)}"


def name_label(dn) -> str:
    sc = name_scope(dn)
    bits = []
    if sc is not None:
        bits.append(f"local to sheet '{sc}'")
    if dn.hidden:
        bits.append("hidden")
    if dn.builtin:
        bits.append("built-in")
    return f"defined name '{dn.name}'" + (f" ({', '.join(bits)})" if bits else "")


def sqref_location(sheet: str, sqref: str) -> str:
    """"'Sheet'!A1:A9,C2" for a (space-separated) sqref."""
    return location(sheet, ",".join((sqref or "").split()) or "?")


def short(text: str, n: int = 120) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 3] + "..."


def plural(n: int, word: str, many: Optional[str] = None) -> str:
    return f"{n} {word if n == 1 else (many or word + 's')}"


# ----------------------------------------------------------------------------- base class
class FormulaScanCheck(Check):
    needs_cells = True
    sheet_kinds = ("worksheet", "macrosheet")      # where cell formulas live

    # ---------------------------------------------------------------- subclass API
    def classify(self, text: str, sheet: Optional[str]):
        """(key, detail) for a formula text that breaks the rule, else None.  sheet is the
        sheet the formula lives on (None for a workbook-scoped defined name)."""
        raise NotImplementedError

    def describe_cells(self, sheet: str, rng: str, n: int, key, detail) -> str:
        raise NotImplementedError

    def describe_rule(self, what: str, sheet: str, sqref: str, keys: list, texts: list) -> str:
        raise NotImplementedError

    def other_formula(self, cell, f) -> None:
        """Data-table anchors and empty <f/> markers (no text)."""

    def safe_classify(self, text: str, sheet: Optional[str]):
        """classify(), or None when the formula text cannot be parsed: skipped for this check (Patrick
        2026-10-05: every attempt graded; stats.defaults.unparsable_formula)."""
        try:
            return self.classify(text, sheet)
        except F.FormulaError as e:
            self.note_default("unparsable_formula", f"{self.site}: {e}")
            return None

    def safe_uses(self, text: str, sheet: Optional[str]) -> frozenset:
        """uses_of(), or nothing when the text cannot be parsed (recorded by safe_classify)."""
        try:
            return self.uses_of(text, sheet)
        except F.FormulaError:
            return EMPTY

    def record(self, cell, result) -> None:
        """Default: remember a finding of a formula cell for this sheet's grouping.  Memory:
        positions are kept compactly per key ({row: [cols]}); a cell's detail only for the
        first MAX_DETAILS flagged cells of the sheet (a rectangle whose top-left detail was
        not kept is described with the first detail of its key)."""
        key, detail = result
        rows = self._hits.get(key)
        if rows is None:
            rows = self._hits[key] = {}
            self._first_detail[key] = detail
        cols = rows.get(cell.row)
        if cols is None:
            rows[cell.row] = [cell.col]
        else:
            cols.append(cell.col)
        self._n_hits += 1
        if len(self._details) < MAX_DETAILS:
            self._details[(cell.row, cell.col)] = detail

    # ---------------------------------------------------------------- engine hooks
    def start(self, wb):
        super().start(wb)
        self.sheet_names = [s.name for s in wb.sheets]
        self._wb_items: list = []          # (location, description) workbook-level, emitted first
        self._sheet_items: list = []       # (location, description) per sheet, tab order
        self.n_formula_cells = 0           # cells with their own formula (shared children included)
        self.n_formula_texts = 0           # distinct formula texts classified (masters, normal, array)
        self.n_flagged_cells = 0
        self.n_cf_rules = 0
        self.n_dv = 0
        self._uses_on = False
        self._name_tab = None              # name_table() cache
        self.bad_scope_names: list = []    # [(index, dn)] names whose localSheetId names no sheet
        self.site = ""                     # where the text being classified lives (stats / messages)

    def sheet_start(self, head):
        self._masters: dict = {}
        self._reset_hits()
        self._sheet_formula_cells = 0

    def _reset_hits(self):
        self._hits: dict = {}              # key -> {row: [cols]}
        self._details: dict = {}           # (row, col) -> detail, first MAX_DETAILS flagged cells
        self._first_detail: dict = {}      # key -> detail of its first flagged cell
        self._n_hits = 0

    def cell(self, cell):
        f = cell.formula
        if f is None:
            return
        text = f.text
        if text:
            res = self._classify_cell_text(text, cell)
            if f.kind == "shared":
                self._masters[f.si] = res
        elif f.kind == "shared":
            res = self._masters.get(f.si)
            if res is None:
                # no master before it: its formula text cannot be read - skipped (Patrick 2026-10-05: every
                # attempt graded)
                self.note_default("unparsable_formula", f"'{cell.sheet}'!{cell.ref}: shared formula si={f.si} has "
                                                        f"no master formula before it")
                return
        else:
            self.other_formula(cell, f)
            return
        self.n_formula_cells += 1
        self._sheet_formula_cells += 1
        found, uses = res
        if uses:
            self._tally(uses, location(cell.sheet, cell.ref))
        if found is not None:
            self.record(cell, found)

    def _classify_cell_text(self, text: str, cell):
        self.n_formula_texts += 1
        self.site = f"{cell.sheet}!{cell.ref}"
        return self.safe_classify(text, cell.sheet), self.safe_uses(text, cell.sheet)

    def sheet_end(self, head, tail):
        name = head.name
        self.flush_hits(name)
        for cf in tail.conditional_formats:
            texts = [t for r in cf.rules for t in cf_rule_texts(r)]
            self.n_cf_rules += len(cf.rules)
            self.scan_rule(f"Conditional format{' (x14)' if cf.source == 'x14' else ''}", name, cf.sqref, texts,
                           extra=self.cf_rule_findings(cf))
        for dv in tail.data_validations:
            texts = [t for t in (dv.formula1, dv.formula2) if t]
            self.n_dv += 1
            self.scan_rule(f"Data validation{' (x14)' if dv.source == 'x14' else ''}", name, dv.sqref, texts)
        if self._uses_on:
            for g in tail.sparkline_groups:
                for t in g.formulas:
                    u = self.safe_uses(t, name)
                    if u:
                        self._tally(u, f"sparkline on '{name}'")

    def cf_rule_findings(self, cf) -> list:
        """Extra (key, detail) findings of a CF block that are not in its formula texts."""
        return []

    def scan_rule(self, what: str, sheet: str, sqref: str, texts: list, extra=()):
        keys, bad = [], []
        self.site = f"{what} on {sqref_location(sheet, sqref)}"
        for t in texts:
            r = self.safe_classify(t, sheet)
            u = self.safe_uses(t, sheet)
            if u:
                self._tally(u, f"{what.lower()} on {sqref_location(sheet, sqref)}")
            if r is not None:
                keys.append(r)
                bad.append(t)
        keys.extend(extra)
        if keys:
            self._sheet_items.append((sqref_location(sheet, sqref),
                                      self.describe_rule(what, sheet, ",".join((sqref or "").split()), keys, bad)))

    def flush_hits(self, sheet: str):
        """Group this sheet's formula-cell findings into rectangles per key."""
        if not self._hits:
            return
        self.n_flagged_cells += self._n_hits
        rects = []
        for k, rows in self._hits.items():
            for b in group_cells((r, c) for r, cols in rows.items() for c in cols):
                rects.append((b, k))
        rects.sort(key=lambda x: (x[0][0], x[0][1]))
        for b, k in rects:
            rng = range_to_str(*b)
            n = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
            detail = self._details.get((b[0], b[1]), self._first_detail[k])
            self._sheet_items.append((location(sheet, rng), self.describe_cells(sheet, rng, n, k, detail)))
        self._reset_hits()

    def emit(self):
        for loc, desc in self._wb_items:
            self.add_mistake(loc, desc)
        for loc, desc in self._sheet_items:
            self.add_mistake(loc, desc)

    # ---------------------------------------------------------------- name table
    def name_table(self):
        """(NameTable, wb_index_of) for this workbook, built once.  NameTable entries: every
        defined name whose scope is readable, then a prefix-less alias of each built-in name
        (_xlnm.Print_Area -> Print_Area, same scope; skipped when a real name already has that
        key).  wb_index_of[k] is the wb.defined_names index of table entry k.  Names whose
        localSheetId names no sheet are kept in self.bad_scope_names [(index, dn)]."""
        if self._name_tab is None:
            entries, idx = [], []
            self.bad_scope_names = []
            names = self.wb.defined_names
            for i, dn in enumerate(names):
                try:
                    sc = dn.scope
                except GradingError:
                    self.bad_scope_names.append((i, dn))
                    continue
                entries.append((dn.name, sc, dn.text or "", dn.hidden))
                idx.append(i)
            keys = {(e[0].lower(), (e[1] or "").lower()) for e in entries}
            for k in range(len(entries)):
                nm, sc, tx, hd = entries[k]
                if nm.lower().startswith("_xlnm.") and len(nm) > 6:
                    key = (nm[6:].lower(), (sc or "").lower())
                    if key not in keys:
                        keys.add(key)
                        entries.append((nm[6:], sc, tx, hd))
                        idx.append(idx[k])
            self._name_tab = F.NameTable(entries, sheet_names=self.sheet_names)
            self._name_tab_idx = idx
        return self._name_tab, self._name_tab_idx

    # ---------------------------------------------------------------- name usage tallies
    def set_flagged_names(self, flagged_idx, use_decides: bool = False) -> None:
        """flagged_idx: indexes into wb.defined_names of names that break the rule.  Builds
        the transitive 'reaches a flagged name' map so formula sites can be tallied.
        use_decides: whether the check counts a flagged name only when it is used (80/87);
        then a name with an unreadable scope that is flagged, or whose definition may reach a
        flagged name, makes that name's use undecidable (name_used raises if it is otherwise
        unused)."""
        self._uses = {i: [0, [], False] for i in flagged_idx}      # idx -> [count, examples, capped]
        self._uncertain: dict = {}
        if not flagged_idx:
            return
        tab, widx = self.name_table()
        tnames = tab.names
        fl = set(flagged_idx)
        reach = {id(tnames[k]): {widx[k]} for k in range(len(tnames)) if widx[k] in fl}
        changed = True
        while changed:
            changed = False
            idents = {tnames[j].name.lower() for j in range(len(tnames)) if id(tnames[j]) in reach}
            for tn in tnames:
                low = (tn.text or "").lower()
                if not any(x in low for x in idents):
                    continue
                acc = set()
                try:
                    tf = parse_text(tn.text)
                except F.FormulaError:
                    continue                      # an unparsable definition reaches nothing (Patrick 2026-10-05)
                for u in F.names_used_by(tf, tn.scope, tab, sheet_names=self.sheet_names):
                    acc |= reach.get(id(u), set())
                cur = reach.setdefault(id(tn), set())
                if acc - cur:
                    cur |= acc
                    changed = True
        self._reach = {k: frozenset(v) for k, v in reach.items() if v}
        ident_reach: dict = {}
        for tn in tnames:
            r = self._reach.get(id(tn))
            if r:
                low = tn.name.lower()
                ident_reach[low] = ident_reach.get(low, frozenset()) | r
        self._ident_reach = ident_reach
        self._idents = tuple(sorted(ident_reach))
        self._uses_on = bool(self._idents)
        self._saturated: set = set()
        # names whose localSheetId names no sheet: no formula can be resolved to them
        for i, dn in self.bad_scope_names:
            why = f"{name_label(dn)} has a localSheetId that names no sheet"
            if i in fl:
                self._uncertain.setdefault(i, []).append(why)
            low = (dn.text or "").lower()
            for x in self._idents:
                if x in low:
                    for j in ident_reach[x]:
                        self._uncertain.setdefault(j, []).append(f"{why} and its definition mentions '{x}'")
        if not use_decides:
            self._uncertain = {}

    def uses_of(self, text: str, sheet: Optional[str]) -> frozenset:
        """Flagged names (indexes) that `text` uses, directly or through other names.  A
        literal INDIRECT("Name") counts as a use (Excel evaluates the name).  Texts that can
        only reach names already used USE_TALLY_SATURATION times are not parsed (their
        names' counts become lower bounds)."""
        if not self._uses_on or not text:
            return EMPTY
        low = text.lower()
        hit = [x for x in self._idents if x in low]
        if not hit:
            return EMPTY
        possible = frozenset().union(*(self._ident_reach[x] for x in hit))
        if possible <= self._saturated:
            for i in possible:
                self._uses[i][2] = True
            return EMPTY
        acc = set()
        for u in F.names_used_by(parse_text(text), sheet, self._name_tab, include_indirect_literals=True,
                                 sheet_names=self.sheet_names):
            acc |= self._reach.get(id(u), EMPTY)
        return frozenset(acc)

    def tally_side_uses(self) -> None:
        """Uses of flagged names outside cells/CF/DV: chart series, shape text links, form
        controls (ctrlProps input ranges / cell links) and pivot caches built on a named
        range (sparklines are tallied in sheet_end)."""
        if not self._uses_on:
            return
        for ch in self.wb.charts:
            for t in ch.formulas:
                u = self.safe_uses(t, ch.sheet)
                if u:
                    self._tally(u, f"chart {ch.part}")
        for dr in self.wb.drawings:
            for t in dr.textlinks:
                u = self.safe_uses(t, dr.sheet)
                if u:
                    self._tally(u, f"shape text link on '{dr.sheet}'")
        for sheet, part, attr, t in ctrl_formulas(self.wb):
            u = self.safe_uses(t, sheet)
            if u:
                self._tally(u, f"form control on '{sheet}' ({attr})")
        for pc in self.wb.pivot_caches:
            if pc.source_name and not pc.source_target:
                u = self.safe_uses(pc.source_name, pc.source_sheet)
                if u:
                    self._tally(u, f"pivot cache {pc.part}")

    def name_used(self, idx: int) -> bool:
        """Whether some formula site uses flagged name idx.  Raises GradingError when it is
        otherwise unused and a name with an unreadable scope may decide it."""
        t = self._uses.get(idx) if hasattr(self, "_uses") else None
        if t and t[0]:
            return True
        why = getattr(self, "_uncertain", {}).get(idx)
        if why:
            dn = self.wb.defined_names[idx]
            raise GradingError(f"{self.key}: cannot tell whether {name_label(dn)} is used: "
                               f"{'; '.join(dict.fromkeys(why))}")
        return False

    def _tally(self, idxs, where: str):
        for i in idxs:
            t = self._uses.get(i)
            if t is None:
                continue
            t[0] += 1
            if len(t[1]) < MAX_EXAMPLES:
                t[1].append(where)
            if t[0] >= USE_TALLY_SATURATION:
                self._saturated.add(i)

    def usage_text(self, idx: int) -> str:
        t = self._uses.get(idx) if hasattr(self, "_uses") else None
        if not t or not t[0]:
            if idx in {i for i, _ in getattr(self, "bad_scope_names", ())}:
                return "its use cannot be traced (its localSheetId names no sheet)"
            return "no formula in the workbook uses it"
        n, ex, capped = t
        more = f" and {n - len(ex)} more" if n > len(ex) else ""
        return (f"used by {'at least ' if capped else ''}{plural(n, 'formula site')} (directly or through "
                f"other names): {', '.join(ex)}{more}")
