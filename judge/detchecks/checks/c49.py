"""49 Formatting/Black font for calculations.

Rule (Patrick's confirmed colour rule, 2026-10-02, applied fully 2026-10-03; details in
docs/checks/49.md and colour_rules.py): every formula cell of the delivered workbook (all
sheets, hidden ones too) is classified by what its formula reads; 49 judges two classes:

    OWN       uses only its own sheet, or no reference at all        -> must be black
    CROSS     calculates with another sheet's cells                  -> black or green
    POINTER   only points at another sheet's cell / range            -> NOT judged here (check 50)
    EXTERNAL  reads another workbook                                 -> NOT judged here (check 51)

So 49 fails a non-black formula that is OWN, and a CROSS formula that is neither black nor
green (guidance: "Green is acceptable only when the formula itself references another sheet").
Black cells are never classified (black is fine in every class); green cells with a direct
reference to another sheet are CROSS, POINTER or EXTERNAL, all fine here, and are accepted
without a full parse.  A red / blue / grey / white ... cell needs its class (it is fine only
as a POINTER or EXTERNAL cell); when the formula cannot be classified it still fails if it
cannot possibly be a pointer or read another workbook (cheap conservative pre-filters), else
GradingError.  HYPERLINK navigation cells are ignored (ruling 2026-10-03, "not a calculation";
colour_rules.is_hyperlink_formula: a HYPERLINK call alone, or - since the 2026-10-04 review -
inside IFERROR / IFNA / IF guards whose other arguments are literals, e.g. the ChatGPT writer's
=IFERROR(HYPERLINK("#'Summary'!A1","OPEN"),"OPEN")).

A green formula that is otherwise OWN and whose INDIRECT address is built from cell contents:
when the address is one own-sheet cell (=INDIRECT(A20)*2) that cell is read in a second pass
(a constant: its text is the address; a formula: GradingError); any other computed address
cannot be decided and raises GradingError (no guess).  In any other non-black colour such a
cell fails whatever the address names (OWN or CROSS are both wrong), so the address is not
read and the mistake says so (reason INDIRECT_UNKNOWN, 2026-10-04) instead of claiming the
formula uses only its own sheet.
"""
from __future__ import annotations

from typing import Optional

from ..errors import GradingError
from .colour_rules import (CROSS, EXTERNAL, OWN, POINTER, UNCLASSIFIED, ColourCheck, Pending, describe_colour,
                           sheets_named_in_text,
                           quick_reads_other_sheet)

# Colour families allowed per JUDGED class under check 49 (POINTER -> 50, EXTERNAL -> 51).
ALLOWED = {
    OWN: frozenset({"black"}),
    CROSS: frozenset({"black", "green"}),
}
NOT_JUDGED = frozenset({POINTER, EXTERNAL})
# Read the own-sheet cell an =INDIRECT(A20) address comes from (second pass) when that decides a
# green cell.  False: such a cell raises GradingError like any computed address.
INDIRECT_READ_ADDRESS_CELL = True
QUICK_CACHE_MAX = 200_000          # memoised quick-scan results (sheet, formula text)
# Reason for a non-black, non-green cell whose only possible other-sheet reading is an INDIRECT
# address built from cell contents: it uses its own sheet or calculates with another sheet's
# cells through that address, and this colour is wrong either way (the address is not read).
INDIRECT_UNKNOWN = "indirect"

_WHAT = {
    OWN: "formula using only its own sheet",
    CROSS: "calculation reading another sheet",
    INDIRECT_UNKNOWN: "formula whose INDIRECT address is built from cell contents (it either stays on its own sheet or "
                      "calculates with another sheet's cells through that address; this colour is wrong either way)",
    UNCLASSIFIED: "formula (what it reads could not be worked out, and need not be: it cannot be a link to "
                  "another sheet or workbook, and this colour is wrong for every other formula)",
}
_RULE = {
    OWN: "formulas that use only their own sheet must be black",
    CROSS: "calculations must be black (green is accepted where the formula reads another sheet)",
    INDIRECT_UNKNOWN: "calculations must be black (green only where the formula reads another sheet)",
    UNCLASSIFIED: "calculations must be black (green only where the formula reads another sheet)",
}


class C49(ColourCheck):
    number = 49
    key = "Formatting/Black font for calculations"
    safe_family = "black"

    def start(self, wb):
        super().start(wb)
        self.n_green_direct = 0
        self.n_unclassified = 0
        self.n_address_cells = 0
        self.n_left = {POINTER: 0, EXTERNAL: 0}
        self._quick: dict = {}

    def prefilter(self, text, fam, own) -> bool:
        # a green formula with a direct reference to another sheet / workbook is CROSS, POINTER
        # or EXTERNAL, none of which can fail 49 in green: no parse needed (memoised per text:
        # shared children and copied formulas repeat it)
        if fam != "green":
            return True
        key = (own, text)
        hit = self._quick.get(key)
        if hit is None:
            hit = quick_reads_other_sheet(text, own)
            if len(self._quick) < QUICK_CACHE_MAX:
                self._quick[key] = hit
        if hit:
            self.n_green_direct += 1
            return False
        return True

    def class_needed(self, fam, text=None) -> bool:
        # red / blue / grey / white ... are wrong for OWN and CROSS; the class is needed only
        # when the formula might be a POINTER or EXTERNAL cell (not judged here)
        if fam in ("green", None) or text is None:
            return True
        return self.clf.may_be_pointer(text) or self.clf.may_read_other_book(text)

    def _raise(self, cell, what: str):
        raise GradingError(f"{self.key}: '{cell.sheet}'!{cell.ref}: {what}")

    def judge(self, cls, fam: str, argb: str, cell):
        if cls.kind == UNCLASSIFIED:
            self.n_unclassified += 1
            return UNCLASSIFIED
        if fam == "green":
            if cls.kind != OWN:
                if cls.unsure and not cls.direct_other:
                    self._raise(cell, f"green formula whose only other-sheet reading goes through a {cls.unsure}; "
                                      f"whether it reads another sheet (green allowed) cannot be decided")
                if cls.kind in NOT_JUDGED:
                    self.n_left[cls.kind] += 1
                elif cls.indirect_assumed and not cls.other_sheets:
                    # Patrick 2026-10-06: sheet-qualified INDIRECT address whose sheet name is computed
                    self.note_default("indirect_sheet_unverified",
                                      f"'{cell.sheet}'!{cell.ref}: {str(cell.formula_text)[:120]}")
                return None
            if cls.unsure:
                self._raise(cell, f"green formula using a {cls.unsure}; whether it reads another sheet "
                                  f"(green allowed) or only its own sheet (must be black) cannot be decided")
            if cls.indirect_unresolved:
                try:
                    tc = self.clf.classify(cell.formula_text, cell.sheet) if self._child else cls
                except GradingError as e:
                    # Patrick 2026-10-05 (every attempt graded): a formula text that cannot be read is skipped
                    self.note_default("unparsable_formula", f"'{cell.sheet}'!{cell.ref}: {e}")
                    return None
                if INDIRECT_READ_ADDRESS_CELL and tc.indirect_cells:
                    return Pending(tc.indirect_cells)
                if INDIRECT_READ_ADDRESS_CELL and tc.indirect_part_cells:
                    return Pending(tc.indirect_part_cells, parts=True)
                # Patrick 2026-10-06 (every attempt graded): no literal and no cell names the sheet -> skipped
                self.note_default("indirect_address_unknown",
                                  f"'{cell.sheet}'!{cell.ref}: {str(cell.formula_text)[:120]}")
                return None
            return OWN
        # red, blue, grey, white ...: wrong for OWN and CROSS; pointers -> 50, externals -> 51
        if cls.kind == EXTERNAL:
            if cls.unsure_ext and not cls.direct_ext:
                self._raise(cell, f"{fam} formula whose only other-workbook reading goes through a {cls.unsure}; "
                                  f"whether it reads another workbook (judged by check 51, not here) cannot be "
                                  f"decided")
            self.n_left[EXTERNAL] += 1
            return None
        if cls.unsure and (cls.pure_name or cls.unsure_ext):
            self._raise(cell, f"{fam} formula using a {cls.unsure}; whether it is a link to another sheet or "
                              f"workbook (judged by checks 50 / 51, not here) or a calculation (must be black) "
                              f"cannot be decided")
        if cls.kind == POINTER:
            self.n_left[POINTER] += 1
            return None
        if cls.kind == OWN and cls.indirect_unresolved:
            return INDIRECT_UNKNOWN     # own sheet, or another sheet through the address: wrong either way
        return cls.kind                 # OWN or CROSS (a computed INDIRECT never reaches another workbook)

    def pending_is_mistake(self, sheet: str, targets: tuple, parts: bool = False) -> Optional[str]:
        """A green =INDIRECT(A20)-style cell: fine when any address cell's text is a reference
        to another sheet or workbook; else an own-sheet formula in green.  `parts` (Patrick
        2026-10-06): the cells are only parts of the address (=INDIRECT(MID(I33,...))): fine
        when any one's text names another sheet as a qualifier ("... Assumptions!E6"); else the
        address is unknown and the cell is skipped (default indirect_address_unknown).  An
        address held in a formula cell is unknown too (49 never reads formula values): skipped."""
        unknown = None
        for rc in targets:
            self.n_address_cells += 1
            st = self._addr.get(rc)
            if st is None:
                continue                                  # empty cell: INDIRECT("") is #REF!
            if st[0] == "formula":
                unknown = f"'{sheet}'!{st[1]} holds a formula; its value would be needed"
                continue
            v = st[1]
            if isinstance(v, str) and v.strip():
                if parts:
                    if any(n.lower() in self.clf._sheets_low and n.lower() != sheet.lower()
                           for n in sheets_named_in_text(v)):
                        return None
                    continue
                ops = self.clf.indirect_operands(v, sheet)
                if any(self.clf._is_external(o) for o in ops) or self.clf._other_sheets(ops, sheet):
                    return None
        if unknown or parts:
            self.note_default("indirect_address_unknown",
                              f"green INDIRECT on '{sheet}': {unknown or 'no address part names a sheet'}")
            return None
        return OWN

    def describe(self, why: str, argb: str) -> str:
        return f"{_WHAT[why]} in {describe_colour(argb)} font; {_RULE[why]}"

    def required_families(self, why) -> frozenset:
        # Patrick 2026-10-06: a conditional format painting the cell black (or green for a calculation
        # that reads another sheet) excuses a wrong base font
        return frozenset({"black", "green"}) if why == CROSS else frozenset({"black"})

    def finish(self) -> dict:
        st = self.base_stats()
        st["green_cells_accepted_by_direct_reference"] = self.n_green_direct
        st["unclassified_offending_cells"] = self.n_unclassified
        st["indirect_address_cells_read"] = self.n_address_cells
        st["pointer_cells_left_to_check_50"] = self.n_left[POINTER]
        st["external_cells_left_to_check_51"] = self.n_left[EXTERNAL]
        return self.verdict(
            f"All {self.n_formula_cells} formula cells are black, or green where the formula reads "
            f"another sheet ({self.fam_counts['green']} green; links to another sheet or workbook and "
            f"HYPERLINK cells are judged elsewhere / ignored).",
            f"{st['offending_cells']} formula cell(s) in {{n}} block(s) are not black and not excused "
            f"by reading another sheet.", st)
