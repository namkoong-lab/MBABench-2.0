"""Base class for deterministic checks + verdict helpers.

Verdict contract (every check, every file):
    {"engine": "harness", "decision": "pass" | "fail", "summary": str,
     "mistakes": [{"location": str, "description": str, "severity": "major" | "minor"}],
     "stats": {...}}
Binary: pass <=> mistakes == [];  fail <=> at least one mistake.
Mistake lists are capped (MAX_MISTAKES, default 25); the total count is always in
stats["n_mistakes"] and in the summary.

Hook order the engine guarantees (see detchecks/docs/reader.md):
    start(wb)
    for each sheet in tab order, if wants_sheet(info):
        sheet_start(head)            -> return False to skip this sheet's rows/cells
        row(row) / cell(cell) ...    (only when needs_rows / needs_cells)
        sheet_end(head, tail)
    [second pass, only for sheets requested with request_second_pass()]
        second_pass_start(head); second_pass_cell(cell) ...; second_pass_end(head, tail)
    finish() -> verdict
"""
from __future__ import annotations

from typing import Iterable, Optional

from ..core.refs import group_cells, location, range_to_str
from ..errors import GradingError

MAX_MISTAKES = 25
SEVERITIES = ("major", "minor")

# Patrick 2026-10-05: "For production runs, every attempt must be graded."  Where a check used to raise
# because it could not decide, it now decides by one of these fixed defaults and records every use in
# stats["defaults"][kind] = {"count", "rule", "examples"} (Check.note_default; present only when used).
# Internal errors (bugs), a missing LibreOffice binary and config errors still raise.
EVERY_ATTEMPT_GRADED = "Patrick 2026-10-05: every attempt graded"
DEFAULT_RULES = {
    "unresolved_colour": "an unresolvable colour (invalid rgb, theme / indexed index outside the palette, invalid "
                         "tint) is read as Excel's default for its slot: a font colour as automatic (black), a "
                         "fill as none",
    "unverified_number_format": "a number format whose display Excel's reading of is not verified (numfmt "
                                "certain=False, e.g. unquoted '0 days') is graded by the engine's best rendering",
    "table_style_visible": "a cell whose visibility depends on a table style the reader cannot resolve is "
                           "assumed visible",
    "unwritten_array_member": "members of an array / spill range that are not written in the file are skipped",
    "untrusted_value": "a formula value that is still untrusted (no Excel cache and no usable recalculation "
                       "value) is skipped for this check",
    "delivered_name_unknown": "the delivered file name is unknown (no or a malformed _attempt_origin.json): the "
                              "format is judged from the file's content",
    "unparsable_formula": "formula text this check cannot parse (or read) is skipped for this check",
    # Patrick 2026-10-06: green INDIRECT formulas whose target sheet is only knowable from text or a value
    "indirect_sheet_unverified": "a green INDIRECT formula whose address is sheet-qualified in form "
                                 "('<name>'!...) with the sheet name computed is taken to read another sheet",
    "indirect_address_unknown": "a green INDIRECT formula whose address is built from cell contents that no "
                                "literal or constant cell resolves to a sheet is skipped for this check",
    # Patrick 2026-10-06 (attempt 2777, measured in Excel): Excel repairs the file by dropping the format
    "unreadable_number_format": "a cell number format that cannot be read (an invalid code, an undefined "
                                "numFmtId) is General, as Excel shows it after dropping the format on repair",
}
MAX_DEFAULT_EXAMPLES = 10


class Mistakes:
    """Capped mistake collector: keeps the first `cap` entries, counts all."""

    def __init__(self, cap: int = MAX_MISTAKES):
        self.cap = cap
        self.items: list[dict] = []
        self.total = 0

    def add(self, location: str, description: str, severity: str = "major"):
        if severity not in SEVERITIES:
            raise ValueError(f"severity must be one of {SEVERITIES}, got {severity!r}")
        self.total += 1
        if len(self.items) < self.cap:
            self.items.append({"location": location, "description": description, "severity": severity})

    def __len__(self):
        return self.total

    @property
    def truncated(self) -> bool:
        return self.total > len(self.items)


def cells_to_ranges(cells: Iterable[tuple[int, int]]) -> list[str]:
    """Group offending (row, col) cells into A1 range strings ('B3:D7', 'F2')."""
    return [range_to_str(*b) for b in group_cells(cells)]


def make_verdict(mistakes: Mistakes, summary: str, stats: Optional[dict] = None) -> dict:
    stats = dict(stats or {})
    stats["n_mistakes"] = mistakes.total
    stats["n_mistakes_listed"] = len(mistakes.items)
    stats["mistakes_truncated"] = mistakes.truncated
    decision = "pass" if mistakes.total == 0 else "fail"
    v = {"engine": "harness", "decision": decision, "summary": summary,
         "mistakes": list(mistakes.items), "stats": stats}
    validate_verdict(v)
    return v


def validate_verdict(v) -> None:
    """Raise GradingError if a verdict breaks the contract."""
    if not isinstance(v, dict):
        raise GradingError(f"verdict is not a dict: {type(v).__name__}")
    missing = {"engine", "decision", "summary", "mistakes", "stats"} - set(v)
    if missing:
        raise GradingError(f"verdict misses keys {sorted(missing)}")
    if v["engine"] != "harness":
        raise GradingError(f"verdict engine must be 'harness', got {v['engine']!r}")
    if v["decision"] not in ("pass", "fail"):
        raise GradingError(f"verdict decision must be pass/fail, got {v['decision']!r}")
    if not isinstance(v["summary"], str) or not v["summary"]:
        raise GradingError("verdict summary must be a non-empty string")
    if not isinstance(v["stats"], dict):
        raise GradingError("verdict stats must be a dict")
    ms = v["mistakes"]
    if not isinstance(ms, list):
        raise GradingError("verdict mistakes must be a list")
    for m in ms:
        if not isinstance(m, dict) or set(m) != {"location", "description", "severity"} \
                or m["severity"] not in SEVERITIES or not isinstance(m["location"], str) \
                or not isinstance(m["description"], str):
            raise GradingError(f"malformed mistake {m!r}")
    if (v["decision"] == "pass") != (len(ms) == 0):
        raise GradingError(f"pass/fail invariant broken: decision={v['decision']} with {len(ms)} mistakes")


class Check:
    """Subclass, set the class attributes, override the hooks you need."""

    number: int = 0                 # rubric_9 check number (1..132 in file order)
    key: str = ""                   # "<Category>/<check name>" exactly as in rubric_9.json
    needs_cells: bool = False       # cell() is called for every <c> of streamed sheets
    needs_rows: bool = False        # row() is called for every <row> of streamed sheets
    needs_values: bool = False      # reads cell.value of formula results (provenance applies);
                                    # without it, cell.value of a formula result raises
    sheet_kinds: tuple = ("worksheet",)   # sheet kinds whose hooks run (worksheet, chartsheet, dialogsheet, macrosheet)
    accepts_unparsed: bool = False  # True only for checks that grade non-SpreadsheetML files (77)
    max_mistakes: int = MAX_MISTAKES
    live: bool = True               # False: the verdict is recorded only; the LLM verdict stands at
                                    # scoring (handoff ruling 2026-10-04; today only 65 is off)

    def __init__(self):
        self.wb = None
        self.mistakes = Mistakes(self.max_mistakes)
        self.stats: dict = {}
        self._second_pass: dict[str, Optional[set]] = {}
        self._defaults: dict[str, dict] = {}

    # ------------------------------------------------------------------ hooks
    def start(self, wb) -> None:
        """Workbook-level data is ready (wb is a detchecks.core.package.Package)."""
        self.wb = wb

    def wants_sheet(self, info) -> bool:
        """Called before a sheet is opened (info: SheetInfo).  False = no hooks for it.
        A sheet whose relationship is broken (part None) or of an unrecognised type (kind
        'unknown') is wanted too, so the engine tries to open it and raises GradingError
        instead of silently skipping it.  Override only to opt out deliberately."""
        return info.kind in self.sheet_kinds or (bool(self.sheet_kinds)
                                                 and (info.part is None or info.kind == "unknown"))

    def sheet_start(self, head):
        """head: SheetHead.  Return False to skip this sheet's rows/cells (sheet_end still runs)."""
        return None

    def row(self, row) -> None:
        pass

    def cell(self, cell) -> None:
        pass

    def sheet_end(self, head, tail) -> None:
        """tail: SheetTail (merges, CF, DV, hyperlinks, ...).  Decide buffered candidates here."""
        pass

    def finish(self) -> dict:
        """Return the verdict (use self.verdict(...))."""
        raise NotImplementedError

    # ------------------------------------------------------------------ second pass
    def request_second_pass(self, sheet_name: str, cells: Optional[Iterable[tuple[int, int]]] = None):
        """Ask the engine to stream `sheet_name` once more after the first pass, calling
        second_pass_cell() for the given (row, col) cells (all cells when None).  Use only
        when unavoidable, e.g. a decision on one sheet needs cells of an EARLIER sheet that
        were not buffered.  Requests made during the first pass or in sheet_end are honoured."""
        cur = self._second_pass.get(sheet_name, set())
        if cells is None or cur is None:
            self._second_pass[sheet_name] = None
        else:
            cur = set(cur)
            cur.update(cells)
            self._second_pass[sheet_name] = cur

    def second_pass_start(self, head) -> None:
        pass

    def second_pass_cell(self, cell) -> None:
        pass

    def second_pass_end(self, head, tail) -> None:
        pass

    # ------------------------------------------------------------------ helpers
    def add_mistake(self, location_: str, description: str, severity: str = "major"):
        self.mistakes.add(location_, description, severity)

    def add_cell_mistakes(self, sheet: str, cells: Iterable[tuple[int, int]], description: str,
                          severity: str = "major") -> int:
        """One mistake per rectangle of adjacent offending cells; returns #rectangles.
        The literal texts '{range}' and '{n}' (cells in that rectangle) in description are
        replaced (plain replace, not str.format)."""
        rects = group_cells(cells)
        for b in rects:
            rng = range_to_str(*b)
            n = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
            self.mistakes.add(location(sheet, rng), description.replace("{range}", rng).replace("{n}", str(n)),
                              severity)
        return len(rects)

    def note_default(self, kind: str, example: str = "", n: int = 1) -> None:
        """Record that something was graded under one of the every-attempt-graded defaults (DEFAULT_RULES,
        Patrick 2026-10-05): stats["defaults"][kind] counts it and keeps the first examples."""
        rec = self._defaults.get(kind)
        if rec is None:
            rec = self._defaults[kind] = {"count": 0, "rule": f"{EVERY_ATTEMPT_GRADED}: {DEFAULT_RULES[kind]}",
                                          "examples": []}
        rec["count"] += n
        if example and len(rec["examples"]) < MAX_DEFAULT_EXAMPLES:
            rec["examples"].append(example[:300])

    def defaults_used(self, kind: Optional[str] = None) -> int:
        """How many times a default (one kind, or any) was used so far."""
        if kind is not None:
            return self._defaults.get(kind, {}).get("count", 0)
        return sum(r["count"] for r in self._defaults.values())

    def require_value(self, cell):
        """cell.value, or GradingError when it is an untrusted formula result (no fallback).  Checks call it
        only for trusted values: an untrusted one is skipped by default (DEFAULT_RULES['untrusted_value'])."""
        if not cell.value_trusted:
            prov = getattr(self.wb, "provenance", None)
            why = (f"writer={prov.writer}, value_path={'given' if prov and prov.value_path else 'none'}"
                   if prov else "no provenance")
            raise GradingError(f"{self.key}: needs the value of {cell.sheet}!{cell.ref} "
                               f"(source={cell.value_source}) but it is untrusted ({why})")
        return cell.value

    def verdict(self, pass_summary: str, fail_summary: str, stats: Optional[dict] = None) -> dict:
        """Build the verdict from self.mistakes.  The literal text '{n}' in fail_summary is
        replaced by the total number of mistakes (plain replace, not str.format)."""
        st = dict(self.stats)
        st.update(stats or {})
        if self._defaults:
            st["defaults"] = {k: {"count": v["count"], "rule": v["rule"], "examples": list(v["examples"])}
                              for k, v in sorted(self._defaults.items())}
        if self.mistakes.total == 0:
            summary = pass_summary
        else:
            summary = fail_summary.replace("{n}", str(self.mistakes.total))
            if self.mistakes.truncated:
                summary += f" (first {len(self.mistakes.items)} of {self.mistakes.total} listed)"
        return make_verdict(self.mistakes, summary, st)
