"""50 Formatting/Green font for cross-sheet links.

Rule (Patrick's confirmed colour rule, 2026-10-02; details in docs/checks/50.md and
colour_rules.py): every formula cell of the delivered workbook (all sheets, hidden ones too)
that ONLY points at a cell or range on ANOTHER sheet of this workbook must have a GREEN font.

A pointer is a formula whose whole text is one reference, optionally wrapped in parentheses,
unary + / -, '@', a spill '#', or one of the pass-through functions ANCHORARRAY, SINGLE, TOROW,
TOCOL, TRANSPOSE (one argument each): =Inputs!B5, =+Inputs!B5, =(Inputs!B5), =-Inputs!B5,
=Inputs!B5#, =@Inputs!B5:B9, =TOROW(Inputs!B5:B9).  A defined name or table that resolves to
another sheet counts (=WACC with WACC = Inputs!$B$3).  A self-qualified reference to the
cell's own sheet ('Monthly Model'!G66 on 'Monthly Model') is not a link.  Spill / array
members take their anchor's formula; each member's own font is judged.

Not judged here: calculations that use another sheet (=Inputs!B5*C3: black or green both
pass), formulas reading another workbook (check 51), own-sheet formulas (check 49).
No majority / convention vote: every non-green pointer is a mistake (Patrick's rule
supersedes the judge note's majority rule).  Green cells are never classified, and a non-green
formula is parsed only when the conservative pre-filter (colour_rules.Classifier.may_be_pointer:
can leave its sheet, no operator / separator / string, and the SHAPE of one reference) says it
may be a pointer; so =SUM(Inputs!A1:A5 or =Inputs!B5+Nope[Col] in black never raise here
(second review, 2026-10-04).  A pointer-shaped text whose class cannot be worked out
(=Nope[Col] with no table Nope, a pure name with an unknown scope) raises GradingError.
"""
from __future__ import annotations

from typing import Optional

from ..errors import GradingError
from .colour_rules import POINTER, ColourCheck, describe_colour


class C50(ColourCheck):
    number = 50
    key = "Formatting/Green font for cross-sheet links"
    safe_family = "green"

    def prefilter(self, text, fam, own) -> bool:
        return self.clf.may_be_pointer(text)

    def judge(self, cls, fam: str, argb: str, cell) -> Optional[str]:
        if cls.unsure and cls.pure_name:
            raise GradingError(f"{self.key}: '{cell.sheet}'!{cell.ref} is a {fam} formula that is only a "
                               f"{cls.unsure}; whether it is a link to another sheet cannot be decided")
        return POINTER if cls.kind == POINTER else None

    def describe(self, why: str, argb: str) -> str:
        return (f"link to another sheet (the formula only points at another sheet's cell or range) in "
                f"{describe_colour(argb)} font; cross-sheet links must be green")

    def required_families(self, why) -> frozenset:
        # Patrick 2026-10-06 (attempt 2348 Cover!C5): a conditional format painting the pointer green excuses it
        return frozenset({"green"})

    def finish(self) -> dict:
        st = self.base_stats()
        n_ptr_bad = st["offending_cells"]
        return self.verdict(
            f"Every formula that only points at another sheet is green ({self.n_formula_cells} formula cells: "
            f"{self.fam_counts['green']} green; none of the other "
            f"{self.n_formula_cells - self.fam_counts['green']} is a pure cross-sheet link).",
            f"{n_ptr_bad} cross-sheet link cell(s) in {{n}} block(s) are not green.", st)
