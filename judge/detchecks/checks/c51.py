"""51 Formatting/Red font for external links.

Rule (Patrick's confirmed colour rule, 2026-10-02; details in docs/checks/51.md and
colour_rules.py): every formula cell of the delivered workbook (all sheets, hidden ones too)
whose formula reads another WORKBOOK anywhere - directly ([1]Prices!$C$5,
'C:\\p\\[Book.xlsx]S'!A1, [1]!Name), inside a calculation (=[1]Rates!$B$2*C25), through a
defined name whose definition reads another workbook, or through a literal
INDIRECT("'[Book.xlsx]S'!A1") - must have a RED font.  A pure pointer to another workbook is
judged here, not by check 50.  Spill / array members take their anchor's formula.

No such cell -> pass ("satisfied by default", rubric), whatever else the workbook holds.
External references that live only in defined names, charts, data validation or conditional
formats (not used by any cell) are check 95's business.  Red cells are never classified.
"""
from __future__ import annotations

from typing import Optional

from ..errors import GradingError
from .colour_rules import EXTERNAL, ColourCheck, describe_colour


class C51(ColourCheck):
    number = 51
    key = "Formatting/Red font for external links"
    safe_family = "red"

    def start(self, wb):
        super().start(wb)
        self.n_table_unplaced = 0      # formulas whose table could not be placed; never external

    def prefilter(self, text, fam, own) -> bool:
        return self.clf.may_read_other_book(text)

    def sheets_needed(self, fam, text=None) -> bool:
        # WHICH sheet of this workbook a formula reads never matters here: a structured
        # reference / bare table name without a [n]! qualifier is this workbook's table, so a
        # table that is not defined or cannot be read leaves the verdict unchanged (R2, 2026-10-04)
        return False

    def judge(self, cls, fam: str, argb: str, cell) -> Optional[str]:
        if cls.kind == EXTERNAL:                # certain whatever an unknown-scope name turns out to be (R1)
            return EXTERNAL
        if cls.sheets_unknown:
            self.n_table_unplaced += 1
        if cls.unsure_ext:
            raise GradingError(f"{self.key}: '{cell.sheet}'!{cell.ref} is a {fam} formula using a {cls.unsure}, "
                               f"one of whose possible definitions may read another workbook; whether the cell "
                               f"reads another workbook cannot be decided")
        return None

    def describe(self, why: str, argb: str) -> str:
        return (f"formula reading another workbook in {describe_colour(argb)} font; external links "
                f"must be red")

    def required_families(self, why) -> frozenset:
        # Patrick 2026-10-06: a conditional format painting the external link red excuses a wrong base font
        return frozenset({"red"})

    def finish(self) -> dict:
        st = self.base_stats()
        st["external_link_parts"] = [{"index": e.index, "kind": e.kind, "target": e.target}
                                     for e in self.wb.external_links]
        st["non_red_cells_reading_another_workbook"] = st["offending_cells"]
        st["table_lookup_failures_not_needing_it"] = self.n_table_unplaced
        return self.verdict(
            f"No formula cell that is not red reads another workbook ({self.n_formula_cells} formula cells, "
            f"{self.fam_counts['red']} red; {len(self.wb.external_links)} externalLink part(s)): satisfied "
            f"(by default when no formula reads another workbook).",
            f"{st['offending_cells']} formula cell(s) reading another workbook, in {{n}} block(s), are "
            f"not red.", st)
