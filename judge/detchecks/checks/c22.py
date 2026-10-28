"""22 Error Checks/No formula errors.

Rule (rubric_9 #22; built 2026-10-04, docs/checks/22.md): FAIL if any cell on any sheet of the
delivered workbook (hidden sheets, rows and columns included) DISPLAYS an Excel error value
(#REF!, #N/A, #DIV/0!, #VALUE!, #NUM!, #NULL!, #NAME?, #SPILL!, #CALC!, #GETTING_DATA, ...) per
its trusted value:
  * the file's own cache when the file was saved by Excel, else the recalculation pipeline
    (core/recalc.py: LibreOffice copy, rerouted to a real-Excel copy when LibreOffice has gaps);
  * every formula cell's value is read (this check needs all values); an untrusted or missing
    value -> GradingError (no fallback).
What counts:
  * a formula whose value is an error (own formula, shared child, array / spill member - each
    member cell displays the error, so each is listed);
  * a typed error constant (t="e" without a formula, e.g. a pasted #N/A) - it is displayed;
  * a deliberate =NA() (the rubric has no exception), a broken reference (=E79-#REF!), a
    misspelled function (#NAME?), a blocked spill (#SPILL!, stored as #VALUE! with vm -> rich
    error value decoded from xl/metadata.xml + richData when the value is the file's own cache);
  * implicit intersection (toy T5): a legacy (non-array) formula that uses a multi-cell range as
    a single value from a cell outside that range displays #VALUE! in Excel although the cache
    / LibreOffice show a number.  Detected with judge/utils/implicit_intersection.py (read-only
    import; only what that detector flags is flagged), on every plain formula, whatever the
    value source (an Excel cache can be stale: T5 has fullCalcOnLoad).
LibreOffice gap cells (judge v15, Patrick 2026-10-28; LIVE since v15, recorded only under v13/v14
because of exactly this): with Excel recalculation OFF, a cell LibreOffice cannot compute (a
function it lacks - LAMBDA, MAP, the CUBE functions ... -, a spill reference, an array-evaluation
difference; core/recalc.py classify_lo_error, the full set in wb.provenance.gap_cells) shows #NAME?
/ #VALUE! in the copy although Excel computes it.  Per attempt:
  * any error in a NON-gap cell (a typed error constant, an implicit intersection, a genuine
    LibreOffice error such as a misspelled or unprefixed function, an Excel-cached error) -> FAIL,
    decided by code; gap cells are listed in stats beside it, never as mistakes;
  * no error cell at all -> PASS, decided by code;
  * the only cells showing errors are gap cells -> UNDECIDED (base.py): the LLM judge decides this
    attempt, told in llm_note which cells are gaps and why (sheet!ref, the formula, the function
    LibreOffice lacks) and that every other cell was checked by code and is error-free.
  A file saved by Excel, or one recalculated by Excel (RecalcPolicy.excel_allowed, kept built and
  OFF), has no gap cells, so code always decides it.
What does not count:
  * errors caught by IFERROR / IFNA / ISERROR ... (the cell displays the fallback, not an error);
  * text that merely reads "#N/A" (a string cell or a formula's string result);
  * circular references with iteration off: not 22's domain (check 100) - the value source
    decides what such a cell shows (Excel: 0 / last value);
  * concealment does NOT excuse: an error hidden by a number format (;;;), white font, a hidden
    row / column / sheet or a zero column width still counts (it is an error value; concealment
    is check 94's domain) - the rubric says "any cell on any sheet".
"""
from __future__ import annotations

import importlib.util
import os
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Optional

from ..core import formula as F
from ..core.refs import col_to_index, location, make_ref, split_ref
from ..core.sheet import ExcelError
from ..errors import GradingError
from .base import Check

II_PATH = os.environ.get("DETCHECKS_II_PATH") or os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "utils", "implicit_intersection.py")
MAX_EXAMPLES = 8
FORMULA_SNIPPET = 90
MAX_GAP_CELLS_IN_NOTE = 40      # gap cells quoted in the llm_note (the rest are counted)
GAP_NOTE_SNIPPET = 120          # characters of a gap cell's formula quoted in the note

# Excel rich-value error types (xl/richData/rdrichvalue.xml _error structure, key errorType)
RICH_ERROR_TYPES = {0: "#NULL!", 1: "#DIV/0!", 2: "#VALUE!", 3: "#REF!", 4: "#NAME?", 5: "#NUM!", 6: "#N/A",
                    7: "#GETTING_DATA", 8: "#SPILL!", 9: "#CONNECT!", 10: "#BLOCKED!", 11: "#UNKNOWN!",
                    12: "#FIELD!", 13: "#CALC!", 14: "#EXTERNAL!", 15: "#PYTHON!"}

_RANGE_RX = re.compile(r"(?<![A-Za-z0-9_.\]])\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![A-Za-z0-9_(])")
_II = None


def _ii_module():
    """The repo's implicit-intersection detector, loaded from its file (read-only; the judge
    package itself is not imported)."""
    global _II
    if _II is None:
        if not os.path.isfile(II_PATH):
            raise GradingError(f"implicit-intersection detector not found at {II_PATH} (set DETCHECKS_II_PATH)")
        import sys
        spec = importlib.util.spec_from_file_location("detchecks_implicit_intersection", II_PATH)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = mod               # dataclasses in the module need it registered
        try:
            spec.loader.exec_module(mod)
        except Exception:
            sys.modules.pop(spec.name, None)
            raise
        _II = mod
    return _II


def may_intersect_outside(text: str, row: int, col: int) -> bool:
    """Cheap pre-filter: can a plain A1:B2 range in the text be implicitly intersected from
    OUTSIDE it (the only case the detector flags)?  False guarantees no flag."""
    if ":" not in text:
        return False
    for m in _RANGE_RX.finditer(F.mask_strings(text)):
        c1, r1, c2, r2 = col_to_index(m.group(1)), int(m.group(2)), col_to_index(m.group(3)), int(m.group(4))
        if c1 > c2:
            c1, c2 = c2, c1
        if r1 > r2:
            r1, r2 = r2, r1
        if (r1, c1) == (r2, c2):
            continue
        if c1 == c2:
            if not (r1 <= row <= r2):
                return True
        elif r1 == r2:
            if not (c1 <= col <= c2):
                return True
        elif not (r1 <= row <= r2 and c1 <= col <= c2):
            return True
    return False


def implicit_intersection_hits(text: str, ref: str) -> list:
    """Ranges of a plain formula that Excel intersects from outside (detector output), [] if none."""
    return _ii_module().check_formula("=" + text, ref)


# ---------------------------------------------------------------------------- rich errors (#SPILL! & co)
def rich_error_map(pkg) -> dict:
    """{vm index (1-based, as in <c vm="n">): error name} for rich-value errors, or {} when
    the workbook has no rich-value metadata or it cannot be read (then a #VALUE! with vm is
    reported as '#VALUE! (rich error)')."""
    try:
        meta = pkg.member("xl/metadata.xml")
        if not meta:
            return {}
        root = pkg.xml(meta)
        loc = lambda t: t.rsplit("}", 1)[-1]  # noqa: E731
        types = []
        future = {}
        value_bks = []
        for el in root:
            ln = loc(el.tag)
            if ln == "metadataTypes":
                types = [t.get("name") for t in el if loc(t.tag) == "metadataType"]
            elif ln == "futureMetadata":
                bks = []
                for bk in el:
                    if loc(bk.tag) != "bk":
                        continue
                    i = None
                    for sub in bk.iter():
                        if loc(sub.tag) == "rvb":
                            i = int(sub.get("i"))
                    bks.append(i)
                future[el.get("name")] = bks
            elif ln == "valueMetadata":
                for bk in el:
                    if loc(bk.tag) == "bk":
                        rc = next((x for x in bk if loc(x.tag) == "rc"), None)
                        value_bks.append((int(rc.get("t")), int(rc.get("v"))) if rc is not None else None)
        rv_part = pkg.member("xl/richData/rdrichvalue.xml")
        st_part = pkg.member("xl/richData/rdrichvaluestructure.xml")
        if not rv_part or not st_part or not value_bks:
            return {}
        structs = []
        for s in pkg.xml(st_part):
            if loc(s.tag) == "s":
                structs.append((s.get("t"), [k.get("n") for k in s if loc(k.tag) == "k"]))
        rvs = []
        for rv in pkg.xml(rv_part):
            if loc(rv.tag) == "rv":
                rvs.append((int(rv.get("s")), [v.text for v in rv if loc(v.tag) == "v"]))
        out = {}
        for n, item in enumerate(value_bks, 1):
            if item is None:
                continue
            t, v = item
            if not (1 <= t <= len(types)) or types[t - 1] != "XLRICHVALUE":
                continue
            bks = future.get("XLRICHVALUE", [])
            if not (0 <= v < len(bks)) or bks[v] is None or not (0 <= bks[v] < len(rvs)):
                continue
            s, vals = rvs[bks[v]]
            if not (0 <= s < len(structs)):
                continue
            kind, keys = structs[s]
            if kind != "_error" or "errorType" not in keys:
                continue
            et = int(vals[keys.index("errorType")])
            out[n] = RICH_ERROR_TYPES.get(et, f"rich error type {et}")
        return out
    except Exception:  # noqa: BLE001 - metadata is decoration for the message; the error itself is still counted
        return {}


# ---------------------------------------------------------------------------- the check
class C22(Check):
    number = 22
    key = "Error Checks/No formula errors"
    live = True                                   # judge v15 (Patrick 2026-10-28: gap cells -> undecided, module doc);
                                                  # recorded only under v13/v14 (Patrick 2026-10-04)
    needs_cells = True
    needs_values = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        _ii_module()                              # loud if the detector is missing
        self._rich = None
        self.n_formula = 0
        self.n_values_read = 0
        self.n_constant_errors = 0
        self.n_error_cells = 0
        self.n_ii = 0
        self.n_ii_checked = 0
        self.by_code = defaultdict(int)
        self.hidden_sheets: list = []
        self.examples: list = []
        self.sheets_with_errors: list = []
        # judge v15: LibreOffice gap cells (module doc); None = no pipeline ran (no gap information)
        prov = getattr(wb, "provenance", None)
        self._gaps = getattr(prov, "gap_cells", None) if prov is not None else None
        self._lo_version = ((prov.recalc or {}).get("lo_version") if prov is not None and prov.recalc else None)
        self.n_gap_error_cells = 0
        self.gap_cells: list = []                 # [(sheet, ref, lo value, formula text, what LibreOffice lacks)]
        self.gap_sheets: list = []

    def _is_gap(self, cell) -> bool:
        g = self._gaps
        if not g:
            return False
        if (cell.sheet, cell.row, cell.col) in g:
            return True
        if cell.array is not None:                # a member of an array / spill range: the gap is keyed by its anchor
            rc = split_ref(cell.array.anchor_ref)
            return rc is not None and (cell.sheet, rc[0], rc[1]) in g
        return False

    def _gap_what(self, cell) -> str:
        g = self._gaps or {}
        what = g.get((cell.sheet, cell.row, cell.col))
        if what is None and cell.array is not None:
            rc = split_ref(cell.array.anchor_ref)
            what = g.get((cell.sheet, rc[0], rc[1])) if rc is not None else None
        return what or "a function LibreOffice lacks"

    def _rich_name(self, vm: int) -> str:
        if self._rich is None:
            self._rich = rich_error_map(self.wb)
        return self._rich.get(vm, "#VALUE! (rich error)")

    def sheet_start(self, head):
        self.hits = defaultdict(list)             # (code, kind) -> [(row, col)]
        self.first_formula = {}                   # (code, kind) -> (ref, formula text, detail) of the first cell

    def _record(self, cell, code: str, kind: str, text: Optional[str], detail: str = ""):
        k = (code, kind)
        self.hits[k].append((cell.row, cell.col))
        if k not in self.first_formula:
            self.first_formula[k] = (cell.ref, text, detail)
        if len(self.examples) < MAX_EXAMPLES:
            self.examples.append(f"{location(cell.sheet, cell.ref)} {code}"
                                 + (f" ({kind}{': ' + detail if detail else ''})" if kind != "formula" else "")
                                 + (f": ={text[:FORMULA_SNIPPET]}" if text else ""))

    def cell(self, cell):
        if cell.is_formula_result:
            self.n_formula += 1
            if not cell.value_trusted:
                # Patrick 2026-10-05 (every attempt graded): an untrusted value is skipped for this check
                self.note_default("untrusted_value", f"{location(cell.sheet, cell.ref)} (source={cell.value_source})")
                return
            self.n_values_read += 1
            v = self.require_value(cell)
            try:
                text = cell.formula_text if cell.has_formula else (
                    cell.array.formula.text if cell.array is not None and cell.array.formula is not None else None)
            except GradingError as e:
                # a formula text that cannot be read (a shared child without its master): its value is still
                # judged, the implicit-intersection test is skipped (Patrick 2026-10-05: every attempt graded)
                self.note_default("unparsable_formula", f"{location(cell.sheet, cell.ref)}: {e}")
                text = None
            if isinstance(v, ExcelError):
                code = str(v)
                if code == "#VALUE!" and cell.vm is not None and cell.value_source == "cached":
                    code = self._rich_name(cell.vm)
                if cell.value_source == "recalc" and self._is_gap(cell):
                    # judge v15: LibreOffice cannot compute this cell; Excel may well (module doc)
                    self.n_gap_error_cells += 1
                    if len(self.gap_cells) < MAX_GAP_CELLS_IN_NOTE:
                        self.gap_cells.append((cell.sheet, cell.ref, code, text or "", self._gap_what(cell)))
                    if not self.gap_sheets or self.gap_sheets[-1] != cell.sheet:
                        self.gap_sheets.append(cell.sheet)
                    return
                self._record(cell, code, "formula", text)
                return
            # implicit intersection: plain formulas only (array / data-table formulas and spill
            # members are array-evaluated by Excel)
            f = cell.formula
            if f is not None and f.kind in ("normal", "shared") and text and may_intersect_outside(text, cell.row, cell.col):
                self.n_ii_checked += 1
                hits = implicit_intersection_hits(text, cell.ref)
                if hits:
                    self._record(cell, "#VALUE!", "implicit intersection", text,
                                 f"{hits[0]['range']} used as a single value from outside it")
            return
        if cell.t == "e" and cell.raw:
            self._record(cell, cell.raw.strip() or "#?", "typed error constant", None)

    def sheet_end(self, head, tail):
        if not self.hits:
            return
        hidden = head.state != "visible"
        n_sheet = 0
        for (code, kind), cells in sorted(self.hits.items(), key=lambda kv: min(kv[1])):
            n_sheet += len(cells)
            self.by_code[code] += len(cells)
            if kind == "typed error constant":
                self.n_constant_errors += len(cells)
            elif kind == "implicit intersection":
                self.n_ii += len(cells)
            ref, text, detail = self.first_formula[(code, kind)]
            if kind == "typed error constant":
                what = f"display the typed error value {code}"
            elif kind == "implicit intersection":
                what = (f"display #VALUE! in Excel (implicit intersection: {detail}; the stored value is a number "
                        f"because the writer did not intersect)")
            else:
                what = f"display {code}"
            ex = f" (e.g. {ref}: ={text[:FORMULA_SNIPPET]})" if text else ""
            self.add_cell_mistakes(head.name, cells, "{n} cell(s) {range} " + what + ex
                                   + (" [hidden sheet]" if hidden else "") + ".")
        self.n_error_cells += n_sheet
        self.sheets_with_errors.append(head.name)
        if hidden:
            self.hidden_sheets.append(head.name)
        self.hits = defaultdict(list)

    def _gap_note(self) -> str:
        """What the LLM judge is told when only gap cells show errors (module doc)."""
        m = re.search(r"LibreOffice(?:Dev)?[/ ](\d+(?:\.\d+)*)", self._lo_version or "")
        engine = f"LibreOffice {m.group(1)}" if m else "LibreOffice"
        lines = [f"{location(s, r)} shows {code} = {('=' + f[:GAP_NOTE_SNIPPET]) if f else '(array member)'}"
                 f"  [{what}]" for s, r, code, f, what in self.gap_cells]
        more = self.n_gap_error_cells - len(self.gap_cells)
        if more > 0:
            lines.append(f"... and {more} more such cell(s) (sheets: {', '.join(self.gap_sheets)})")
        return (f"The deterministic grader could not decide this check by code. {self.n_gap_error_cells} cell(s) "
                f"show #NAME?/#VALUE! only because the recalculation engine ({engine}) lacks the function or "
                f"feature named in brackets; Excel computes them. Judge whether each of these formulas is valid "
                f"Excel (spelling, argument count, references), not the displayed error:\n"
                + "\n".join(lines)
                + "\nEvery other cell of the workbook has been checked by code and is error-free.")

    def finish(self) -> dict:
        stats = {"formula_cells": self.n_formula, "formula_values_read": self.n_values_read,
                 "error_cells": self.n_error_cells, "error_cells_by_code": dict(sorted(self.by_code.items())),
                 "typed_error_constants": self.n_constant_errors,
                 "implicit_intersection_cells": self.n_ii, "implicit_intersection_checked": self.n_ii_checked,
                 "sheets_with_errors": self.sheets_with_errors, "hidden_sheets_with_errors": self.hidden_sheets,
                 "examples": self.examples,
                 # judge v15: LibreOffice gap cells showing an error (never a mistake; module doc)
                 "gap_error_cells": self.n_gap_error_cells,
                 "gap_error_examples": [f"{location(s, r)} {code}" + (f": ={f[:FORMULA_SNIPPET]}" if f else "")
                                        + f" [{what}]" for s, r, code, f, what in self.gap_cells[:MAX_EXAMPLES]],
                 "gap_sheets": self.gap_sheets,
                 "gap_information": "none" if self._gaps is None else "pipeline"}
        if self.mistakes.total == 0 and self.n_gap_error_cells:
            return self.undecided(
                f"Only LibreOffice gap cells show errors ({self.n_gap_error_cells} cell(s) LibreOffice cannot "
                f"compute; every other one of the {self.n_formula} formula cells is error-free): the LLM judge "
                f"decides this attempt, told which cells are gaps.",
                self._gap_note(), stats,
                details={"gap_error_cells": self.n_gap_error_cells, "gap_sheets": self.gap_sheets,
                         "decided_by": "llm"})
        stats["decided_by"] = "code"
        gap_note = (f" {self.n_gap_error_cells} LibreOffice gap cell(s) show errors Excel would not (not counted)."
                    if self.n_gap_error_cells else "")
        return self.verdict(f"No cell displays an Excel error value ({self.n_formula} formula cells checked).",
                            "{n} block(s) of cells display Excel error values "
                            f"({self.n_error_cells} cell(s): " +
                            ", ".join(f"{k} x{v}" for k, v in sorted(self.by_code.items())) + f").{gap_note}", stats)
