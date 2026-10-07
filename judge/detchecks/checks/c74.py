"""74 Formatting/No merged cells.

Rule: fail on any merged range (<mergeCell ref=...> with more than one cell) on any
worksheet, hidden sheets included, read from the DELIVERED file - except the Instructions
sheet, which the rubric exempts ("Merged cells on the Instructions Tab are allowed").

Identifying the Instructions sheet: the sheet whose name, trimmed and compared
case-insensitively, is exactly "Instructions".  In the toys and in every delivered corpus
attempt that has a case instructions tab (372 of 374 files checked 2026-10-02, corpus/attempts
+ corpus/gpt6) the tab is named exactly "Instructions"; renamed or copied tabs ("Instructions
(2)", "Instructions & Notes") are NOT exempt, because without the starting file a rename
cannot be told from an agent-made sheet.

1x1 merges (B2:B2) are ignored: they merge nothing.  Center Across Selection is not a merge
(it is an alignment), so it never shows up here.  Only the sheet tail is needed: sheet data
is skipped at byte level.
"""
from __future__ import annotations

from ..core.refs import location
from .base import Check

INSTRUCTIONS_NAME = "instructions"


def is_instructions_sheet(name: str) -> bool:
    return name.strip().casefold() == INSTRUCTIONS_NAME


class C74(Check):
    number = 74
    key = "Formatting/No merged cells"
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        self._n_sheets = 0
        self._sheets_with = []
        self._exempt = []
        self._single = 0

    def sheet_end(self, head, tail):
        self._n_sheets += 1
        merges = [m for m in tail.merges if not m.is_single_cell]
        self._single += len(tail.merges) - len(merges)
        if not merges:
            return
        if is_instructions_sheet(head.name):
            self._exempt.append({"sheet": head.name, "n_merges": len(merges),
                                 "refs": [m.ref for m in merges[:10]]})
            return
        self._sheets_with.append({"sheet": head.name, "state": head.state, "n_merges": len(merges)})
        for m in merges:
            n = (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1)
            self.add_mistake(location(head.name, m.ref),
                             f"Merged cells {m.ref} ({n} cells) on sheet '{head.name}'"
                             f"{' (hidden sheet)' if head.state != 'visible' else ''}; "
                             f"use Center Across Selection instead of merging.")

    def finish(self) -> dict:
        stats = {"n_sheets_checked": self._n_sheets, "sheets_with_merges": self._sheets_with,
                 "instructions_exempt": self._exempt, "n_single_cell_merges_ignored": self._single}
        sheets = ", ".join(f"'{s['sheet']}' ({s['n_merges']})" for s in self._sheets_with)
        exempt = (f" Merges on the Instructions sheet are exempt ({sum(e['n_merges'] for e in self._exempt)})."
                  if self._exempt else "")
        return self.verdict(f"No merged cells on {self._n_sheets} sheet(s).{exempt}",
                            f"{self.mistakes.total} merged range(s) found on: {sheets}.{exempt}", stats)
