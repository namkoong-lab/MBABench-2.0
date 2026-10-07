"""92 Potential Dangers/No hidden sheets.

Rule (handoff, Patrick 2026-10-02): fail on every sheet whose state is hidden or
veryHidden, whichever tab it is (worksheet, chartsheet, dialog or macro sheet; case tab or
agent tab).  No "disclosed and justified" exception, no inheritance exemption.

Reads only workbook.xml (<sheet state=...>): no sheet part is opened.
A hidden workbook WINDOW (<workbookView visibility="hidden">) is not a hidden sheet; it is
recorded in stats only.
"""
from __future__ import annotations

from ..errors import GradingError
from .base import Check

STATES = ("visible", "hidden", "veryHidden")


class C92(Check):
    number = 92
    key = "Potential Dangers/No hidden sheets"

    def wants_sheet(self, info) -> bool:
        return False                      # workbook-level only

    def finish(self) -> dict:
        hidden = []
        for s in self.wb.sheets:
            if s.state not in STATES:
                raise GradingError(f"{self.key}: sheet {s.name!r} has an unknown state {s.state!r}")
            if s.state != "visible":
                hidden.append(s)
                how = "very hidden (state=\"veryHidden\": not even listed in Excel's Unhide dialog)" \
                    if s.state == "veryHidden" else "hidden (state=\"hidden\")"
                self.add_mistake(s.name, f"Sheet '{s.name}' (tab {s.index + 1}, {s.kind}) is {how}. "
                                         f"Hidden sheets are not allowed; unhide it or remove it.")
        hidden_windows = sum(1 for v in self.wb.book_views if (v.get("visibility") or "visible") != "visible")
        stats = {"n_sheets": len(self.wb.sheets),
                 "n_hidden": sum(1 for s in hidden if s.state == "hidden"),
                 "n_very_hidden": sum(1 for s in hidden if s.state == "veryHidden"),
                 "hidden_sheets": [{"name": s.name, "state": s.state, "kind": s.kind} for s in hidden],
                 "hidden_workbook_windows": hidden_windows}
        listing = ", ".join(f"'{s.name}' ({s.state})" for s in hidden)
        return self.verdict(f"All {len(self.wb.sheets)} sheets are visible.",
                            f"{len(hidden)} of {len(self.wb.sheets)} sheets are hidden: {listing}.", stats)
