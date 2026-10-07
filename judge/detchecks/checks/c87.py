"""87 Formulas/Avoid whole-column references.

Rule as implemented (see docs/checks/87.md):
fail on every formula site that holds a whole-column or whole-row reference -
A:A, $A:$C, Sheet!A:A, 'My Sheet'!$B:$B, Jan:Dec!A:A, [1]Data!A:A, 1:1, $17:$17, and
A1:A1048576 / A1:XFD1 (Excel shows those as A:A / 1:1).  Formula sites:
  * cell formulas (normal, shared - every child of a flagged master -, array), every
    worksheet and macro sheet, hidden sheets included;
  * defined names that a formula site uses, directly or through other names (visible or
    hidden, built-in or not); unused names - above all Print_Titles / Print_Area print
    settings - are reported in stats only (COUNT_UNUSED_NAMES);
  * conditional-format and data-validation formulas (main and x14).
Never counted: trimmed ranges (A:.A, _xlfn._TRO_TRAILING(A:A), TRIMRANGE(A:A)), table
references (Tbl[Col]), spill references (A1#), Name1:Name2, text inside string literals
or text cells ("1:1", "A:B"), chart/sparkline/pivot sources, and bounded ranges that run to
the sheet edge (A2:A1048576: counted in stats only, question for Patrick).
A cell that uses a flagged name is not a separate mistake; the name is, with its users
listed.  Tokenising is detchecks.core.formula's; the conservative pre-filter
quick_may_have_whole_refs skips texts that cannot hold such a reference.
Shared-formula copies inherit their master's result, except that a copy of an unflagged
master whose bounded range has one end absolute at the sheet edge ($A$1:A1048575 copied
one row down = $A$1:A1048576 = column A) is re-classified from its own shifted text.
A flagged name whose use cannot be traced (its own localSheetId, or that of a name on the
chain to it, names no sheet) raises GradingError only when nothing else fails for certain;
once a cell / CF / DV site or another counted name already fails the workbook, the verdict
does not depend on it and it goes to stats (whole_ref_names_use_undecidable_not_counted).
"""
from __future__ import annotations

import re

from ..core import formula as F
from ..core.refs import MAX_COL, MAX_ROW
from ..errors import GradingError
from ._fscan import FormulaScanCheck, parse_text, name_label, name_location, plural, short

# A defined name counts only when a formula site uses it (directly or through other names):
# an unused name - e.g. a Print_Area or Print_Titles set to whole columns/rows - is stats only.
COUNT_UNUSED_NAMES = False
# Conditional-format and data-validation formulas are formulas (rubric "per formula").
COUNT_CF_DV = True
MAX_EDGE_EXAMPLES = 5

KIND_WORD = {"whole_column": "whole-column", "whole_row": "whole-row"}
# pre-filter for shared masters whose copies may turn a bounded range into a whole column/row:
# a part absolute at a sheet edge ($1, $1048576, $A, $XFD); string literals masked first
_RE_EDGE_ABS = re.compile(r"\$(?:1(?![0-9])|1048576(?![0-9])|[Aa](?=\$?[0-9])|[Xx][Ff][Dd](?=\$?[0-9]))")


def _definition(dn, n: int = 80) -> str:
    """A defined name's definition for messages, shortened.  Excel stores it without the
    leading '='; a writer that keeps one would otherwise print as '(==S!$A:$A)'."""
    t = (dn.text or "").strip()
    if t.startswith("="):
        t = t[1:].lstrip()
    return short(t, n)


def _edge_capable(parts) -> bool:
    """A bounded range (cell endpoints) that a shared copy can stretch to a whole column or row:
    one row absolute at row 1 / 1048576 and another row relative (or the same for columns)."""
    rows_abs_edge = any(p.row_abs and p.row in (1, MAX_ROW) for p in parts)
    rows_rel = any(not p.row_abs for p in parts)
    cols_abs_edge = any(p.col_abs and p.col in (1, MAX_COL) for p in parts)
    cols_rel = any(not p.col_abs for p in parts)
    return (rows_abs_edge and rows_rel) or (cols_abs_edge and cols_rel)


def _whole_after_shift(parts, dr: int, dc: int) -> bool:
    """Whether the range is a whole column/row in the copy moved by (dr, dc).  A copy pushed
    off the grid is #REF! (not whole)."""
    rows, cols = [], []
    for p in parts:
        r = p.row if p.row_abs else p.row + dr
        c = p.col if p.col_abs else p.col + dc
        if not (1 <= r <= MAX_ROW and 1 <= c <= MAX_COL):
            return False
        rows.append(r)
        cols.append(c)
    return (min(rows) == 1 and max(rows) == MAX_ROW) or (min(cols) == 1 and max(cols) == MAX_COL)


class C87(FormulaScanCheck):
    number = 87
    key = "Formulas/Avoid whole-column references"

    def start(self, wb):
        super().start(wb)
        self.n_edge = 0
        self.edge_examples: list = []
        self.n_trimmed = 0
        flagged = []
        self._name_items = []
        for i, dn in enumerate(wb.defined_names):
            self.site = name_location(dn)
            r = self.safe_classify(dn.text, None)
            if r is not None:
                flagged.append(i)
                self._name_items.append((i, dn, r))
        self.set_flagged_names(flagged, use_decides=not COUNT_UNUSED_NAMES)
        self.n_shared_copies_reclassified = 0

    # ------------------------------------------------------------------ rule
    def classify(self, text, sheet):
        if not text or not F.quick_may_have_whole_refs(text):
            return None
        f = parse_text(text)
        whole = []
        for o in f.operands:
            if o.kind in KIND_WORD:
                whole.append(o)
            elif o.kind == "range" and o.reaches_sheet_edge:
                self.n_edge += 1
                if len(self.edge_examples) < MAX_EDGE_EXAMPLES:
                    self.edge_examples.append(f"{self.site}: {o.raw}")
            elif o.kind == "trimmed_range":
                self.n_trimmed += 1
        if not whole:
            return None
        kinds = frozenset(KIND_WORD[o.kind] for o in whole)
        refs = tuple(dict.fromkeys(o.raw for o in whole))
        return kinds, refs

    # ------------------------------------------------------------------ shared copies at the edge
    def sheet_start(self, head):
        super().sheet_start(head)
        self._edge_masters: dict = {}      # si -> (row, col, text, [parts of edge-capable ranges])

    def cell(self, cell):
        super().cell(cell)
        f = cell.formula
        if f is None or f.kind != "shared":
            return
        if f.text:
            self._edge_masters.pop(f.si, None)
            m = self._masters.get(f.si)
            if m is not None and m[0] is None and ":" in f.text and _RE_EDGE_ABS.search(F.mask_strings(f.text)):
                try:
                    ops = parse_text(f.text).operands
                except F.FormulaError:
                    ops = []                       # unparsable: skipped (recorded by safe_classify)
                rngs = [o.parts for o in ops
                        if o.kind == "range" and o.bounds is not None and _edge_capable(o.parts)]
                if rngs:
                    self._edge_masters[f.si] = (cell.row, cell.col, f.text, rngs)
            return
        e = self._edge_masters.get(f.si)
        if e is None:
            return
        mr, mc, text, rngs = e
        dr, dc = cell.row - mr, cell.col - mc
        if any(_whole_after_shift(parts, dr, dc) for parts in rngs):
            self.site = f"{cell.sheet}!{cell.ref}"
            res = self.classify(F.shift(text, dr, dc), cell.sheet)
            if res is not None:
                self.n_shared_copies_reclassified += 1
                self.record(cell, res)

    def scan_rule(self, what, sheet, sqref, texts, extra=()):
        if COUNT_CF_DV:
            super().scan_rule(what, sheet, sqref, texts, extra)

    # ------------------------------------------------------------------ descriptions
    def describe_cells(self, sheet, rng, n, kinds, refs):
        k = " and ".join(sorted(kinds))
        cells = "Formula" if n == 1 else f"{n} formulas"
        ex = ", ".join(refs[:3])
        return (f"{cells} in {rng} use{'s' if n == 1 else ''} {k} reference(s) "
                f"(e.g. {ex}); "
                f"bound the range, or use a table column, TrimRef (A:.A) or a defined bounded range.")

    def describe_rule(self, what, sheet, sqref, keys, texts):
        refs = list(dict.fromkeys(r for _, rr in keys for r in rr))
        return (f"{what} formula(s) on {sqref} use whole-column/row reference(s) {', '.join(refs[:3])} "
                f"(rule: {short(texts[0], 80)}).")

    # ------------------------------------------------------------------ verdict
    def finish(self):
        self.tally_side_uses()
        unused, undecidable, errors = [], [], []
        for i, dn, (kinds, refs) in self._name_items:
            tag = f"{dn.name}{' (built-in)' if dn.builtin else ''}={_definition(dn)}"
            if not COUNT_UNUSED_NAMES:
                try:
                    used = self.name_used(i)
                except GradingError as e:
                    # its use cannot be traced (a localSheetId that names no sheet, on it or on the
                    # chain to it): decided below, once every certain site is known
                    undecidable.append(f"{tag}: {str(e).replace(self.key + ': ', '', 1)}")
                    errors.append(e)
                    continue
                if not used:
                    unused.append(tag)
                    continue
            self._wb_items.append((name_location(dn),
                                   f"The {name_label(dn)} refers to {' and '.join(sorted(kinds))} reference(s) "
                                   f"{', '.join(refs[:3])} (={_definition(dn)}); {self.usage_text(i)}. "
                                   f"A name that still spans a whole column/row is not a remedy."))
        # every sheet has been flushed (cells, CF, DV) and every decidable name is in: an undecidable
        # name matters only when nothing else fails for certain (same policy as Active cell reset
        # to A1 (62): undecidable units go to stats once the verdict is settled, else no guess)
        if errors and not (self._wb_items or self._sheet_items):
            raise GradingError(" | ".join(dict.fromkeys(str(e) for e in errors)))
        self.emit()
        stats = {"n_formula_cells": self.n_formula_cells, "n_formula_texts": self.n_formula_texts,
                 "n_flagged_cells": self.n_flagged_cells, "n_names": len(self.wb.defined_names),
                 "n_names_flagged": len(self._name_items) - len(unused) - len(undecidable),
                 "whole_ref_names_unused_not_counted": unused,
                 "whole_ref_names_use_undecidable_not_counted": undecidable, "n_cf_rules": self.n_cf_rules,
                 "n_data_validations": self.n_dv, "count_cf_dv": COUNT_CF_DV,
                 "sheet_edge_ranges_not_counted": self.n_edge, "sheet_edge_examples": self.edge_examples,
                 "trimmed_ranges_seen": self.n_trimmed,
                 "shared_copies_reclassified": self.n_shared_copies_reclassified}
        return self.verdict(
            f"No whole-column/row reference in {plural(self.n_formula_cells, 'formula cell')}, "
            f"{plural(len(self.wb.defined_names), 'defined name')} and the CF/DV rules.",
            "{n} place(s) use whole-column/row references.", stats)
