"""50 Formatting/Green font for cross-sheet links.

Rule (Patrick's confirmed colour rule 2026-10-02, widened 2026-10-26 for judge v15 by the new house
rules, v3: "green: links to another sheet/tab (also when used inside formulas)"; details in
docs/checks/50.md and colour_rules.py): every formula cell of the delivered workbook (all sheets,
hidden ones too) that reads ANOTHER sheet of this workbook must have a GREEN font - a pure link
(=Inputs!B5, class POINTER) as much as a calculation that uses another sheet's cells
(=Inputs!B5*C3, =SUM(Inputs!A1:A9), =WACC*B2 with WACC on Inputs, =Tbl[Col]*2 with Tbl on
Inputs: class CROSS).  Under v13/v14 (ruling 2026-10-03) only pointers were judged here and a
cross-sheet calculation passed in black or green (49's business); since v15 a black cross-sheet
calculation fails 50, and 50 only: 49 judges same-sheet formulas alone, so every formula is
charged once.

A pointer is a formula whose whole text is one reference, optionally wrapped in parentheses,
unary + / -, '@', a spill '#', or one of the pass-through functions ANCHORARRAY, SINGLE, TOROW,
TOCOL, TRANSPOSE (one argument each): =Inputs!B5, =+Inputs!B5, =(Inputs!B5), =-Inputs!B5,
=Inputs!B5#, =@Inputs!B5:B9, =TOROW(Inputs!B5:B9).  A defined name or table that resolves to
another sheet counts (=WACC with WACC = Inputs!$B$3).  A self-qualified reference to the
cell's own sheet ('Monthly Model'!G66 on 'Monthly Model') is not a link.  Spill / array
members take their anchor's formula; each member's own font is judged.  The two classes are
told apart in the mistake text and the stats only; both must be green.

Not judged here: own-sheet formulas (check 49), formulas reading another workbook (check 51),
HYPERLINK navigation cells (any colour, ruling 2026-10-03).  No majority / convention vote:
every non-green cross-sheet formula is a mistake (Patrick's rule supersedes the old judge
note's majority rule).  Green cells are never classified.  A non-green formula is parsed when
the conservative pre-filter (colour_rules.Classifier.may_leave_sheet: a '!' sheet qualifier, a
'[' of a table or book, or an identifier spelled like a defined name or table, outside string
literals unless INDIRECT is called) says it may reach another sheet, so =B2*C3 or =SUM(A1:A9)
in black is never parsed; the classifier's per-workbook cache is shared with 49 and 51, so a
text is parsed once for the three checks.  Malformed text or an undefined table in a formula
that may leave its sheet is skipped under the every-attempt-graded default unparsable_formula
(Patrick 2026-10-05).

INDIRECT (Patrick 2026-10-06, as 49): a BLACK formula whose INDIRECT address is built from one
own-sheet constant cell (=INDIRECT(A20)*2) reads that cell in a second pass - an address naming
another sheet makes the formula CROSS (fails), an own-sheet address OWN (fine), an address that
reads another workbook is 51's; a formula-cell address or parts naming no sheet are skipped
(default indirect_address_unknown).  A sheet-qualified address whose sheet name is computed is
taken to read another sheet (default indirect_sheet_unverified).  A blue / red / grey ... cell
of that kind is wrong whatever the address names and is charged by 49 without reading it, so
50 leaves it alone (charged once).  A black formula whose only possible other-sheet reading
goes through a defined name of unknown scope (bad localSheetId) cannot be decided and raises
GradingError (no fallback), as a black pure name does; such a cell in any other non-green colour
is 49's "wrong either way" charge.
"""
from __future__ import annotations

from typing import Optional

from ..errors import GradingError
from .colour_rules import (_RE_INDIRECT_CALL, CROSS, EXTERNAL, OWN, POINTER, UNCLASSIFIED, ColourCheck, Pending,
                           describe_colour, pointer_shaped, quick_reads_other_sheet, sheets_named_in_text)

# Read the own-sheet cell an =INDIRECT(A20) address comes from (second pass) when that decides a black cell.
INDIRECT_READ_ADDRESS_CELL = True
# A non-green formula with a DIRECT reference to another sheet of this workbook (quick_reads_other_sheet, the
# scan 49 trusts to accept green cells) and nothing that could read another workbook (may_read_other_book) is
# CROSS or POINTER for sure: charged without a full parse (pointer_shaped tells the two apart for the message).
# On a 1.6-million-formula attempt with 1.5 million distinct texts this cuts 50 from ~80 s to a few seconds.
QUICK_DIRECT_REFERENCE = True
QUICK_CACHE_MAX = 200_000

_WHAT = {
    POINTER: "link to another sheet (the formula only points at another sheet's cell or range)",
    CROSS: "calculation that references another sheet",
}


class C50(ColourCheck):
    number = 50
    key = "Formatting/Green font for cross-sheet links"
    safe_family = "green"

    def start(self, wb):
        super().start(wb)
        self.n_address_cells = 0
        self.n_left = {OWN: 0, EXTERNAL: 0}
        self.n_quick_direct = 0
        self._quick: dict = {}

    def prefilter(self, text, fam, own) -> bool:
        # may leave its sheet by a written reference / name / table, or through an INDIRECT address
        # (=INDIRECT(A20): the address lives in a cell, so the text itself names no sheet)
        return self.clf.may_leave_sheet(text) or bool(_RE_INDIRECT_CALL.search(text))

    def quick_class(self, text: str, own: str):
        """CROSS / POINTER when a direct reference to another sheet of this workbook settles the class
        without a parse (QUICK_DIRECT_REFERENCE); None when the full classifier is needed."""
        if not QUICK_DIRECT_REFERENCE:
            return None
        key = (own, text)
        hit = self._quick.get(key, 0)
        if hit == 0:
            hit = None
            if quick_reads_other_sheet(text, own) and not self.clf.may_read_other_book(text):
                hit = POINTER if pointer_shaped(text) else CROSS
            if len(self._quick) < QUICK_CACHE_MAX:
                self._quick[key] = hit
        return hit

    def judge(self, cls, fam: str, argb: str, cell) -> Optional[str]:
        if cls.kind == UNCLASSIFIED:
            return None
        if cls.kind == EXTERNAL:
            self.n_left[EXTERNAL] += 1                  # 51's business
            return None
        if cls.unsure and cls.pure_name:
            raise GradingError(f"{self.key}: '{cell.sheet}'!{cell.ref} is a {fam} formula that is only a "
                               f"{cls.unsure}; whether it is a link to another sheet cannot be decided")
        if cls.kind == POINTER:
            return POINTER
        if cls.kind == CROSS:
            if cls.indirect_assumed and not cls.other_sheets:
                # Patrick 2026-10-06: sheet-qualified INDIRECT address whose sheet name is computed
                self.note_default("indirect_sheet_unverified",
                                  f"'{cell.sheet}'!{cell.ref}: {str(cell.formula_text)[:120]}")
            return CROSS
        # OWN as far as the text says
        if fam != "black":
            # blue / red / grey ...: an own-sheet formula is 49's; one that may still read another sheet
            # (unsure name, computed INDIRECT address) is wrong either way and 49 charges it unread
            self.n_left[OWN] += 1
            return None
        if cls.unsure:
            raise GradingError(f"{self.key}: '{cell.sheet}'!{cell.ref}: black formula using a {cls.unsure}; "
                               f"whether it reads another sheet (must be green) or only its own sheet (fine in "
                               f"black) cannot be decided")
        if cls.indirect_unresolved:
            try:
                tc = self.clf.classify(cell.formula_text, cell.sheet) if self._child else cls
            except GradingError as e:
                self.note_default("unparsable_formula", f"'{cell.sheet}'!{cell.ref}: {e}")
                return None
            if INDIRECT_READ_ADDRESS_CELL and tc.indirect_cells:
                return Pending(tc.indirect_cells)
            if INDIRECT_READ_ADDRESS_CELL and tc.indirect_part_cells:
                return Pending(tc.indirect_part_cells, parts=True)
            self.note_default("indirect_address_unknown",
                              f"'{cell.sheet}'!{cell.ref}: {str(cell.formula_text)[:120]}")
            return None
        self.n_left[OWN] += 1
        return None

    def pending_is_mistake(self, sheet: str, targets: tuple, parts: bool = False) -> Optional[str]:
        """A black =INDIRECT(A20)-style cell: CROSS (a mistake) when an address cell's text names
        another sheet of this workbook; fine when it reads another workbook (51's) or stays on its
        own sheet; skipped (default indirect_address_unknown) when the address is held in a formula
        cell or its parts name no sheet."""
        unknown = None
        other = False
        for rc in targets:
            self.n_address_cells += 1
            st = self._addr.get(rc)
            if st is None:
                continue                                  # empty cell: INDIRECT("") is #REF!, reads nothing
            if st[0] == "formula":
                unknown = f"'{sheet}'!{st[1]} holds a formula; its value would be needed"
                continue
            v = st[1]
            if isinstance(v, str) and v.strip():
                if parts:
                    if any(n.lower() in self.clf._sheets_low and n.lower() != sheet.lower()
                           for n in sheets_named_in_text(v)):
                        other = True
                    continue
                ops = self.clf.indirect_operands(v, sheet)
                if any(self.clf._is_external(o) for o in ops):
                    return None                           # reads another workbook: check 51's business
                if self.clf._other_sheets(ops, sheet):
                    other = True
        if other:
            return CROSS
        if unknown or parts:
            self.note_default("indirect_address_unknown",
                              f"black INDIRECT on '{sheet}': {unknown or 'no address part names a sheet'}")
        return None

    def describe(self, why: str, argb: str) -> str:
        return f"{_WHAT[why]}; font {describe_colour(argb)}, must be green (cross-sheet links must be green)"

    def required_families(self, why) -> frozenset:
        # Patrick 2026-10-06 (attempt 2348 Cover!C5): a conditional format painting the cell green excuses it
        return frozenset({"green"})

    def finish(self) -> dict:
        st = self.base_stats()
        st["cells_charged_by_direct_reference"] = self.n_quick_direct
        st["indirect_address_cells_read"] = self.n_address_cells
        st["own_sheet_cells_left_to_check_49"] = self.n_left[OWN]
        st["external_cells_left_to_check_51"] = self.n_left[EXTERNAL]
        n_bad = st["offending_cells"]
        by = st["offending_cells_by_reason"]
        return self.verdict(
            f"Every formula that reads another sheet is green ({self.n_formula_cells} formula cells: "
            f"{self.fam_counts['green']} green; none of the other "
            f"{self.n_formula_cells - self.fam_counts['green']} links to or calculates with another sheet).",
            f"{n_bad} cell(s) reading another sheet in {{n}} block(s) are not green "
            f"({by.get(POINTER, 0)} pure link(s), {by.get(CROSS, 0)} calculation(s)).", st)
