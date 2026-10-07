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
from ..core.refs import col_to_index, location, make_ref
from ..core.sheet import ExcelError
from ..errors import GradingError
from .base import Check

II_PATH = os.environ.get("DETCHECKS_II_PATH") or os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "utils", "implicit_intersection.py")
MAX_EXAMPLES = 8
FORMULA_SNIPPET = 90

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
    live = False                                  # Patrick 2026-10-04: stays with the LLM judge as in judge v12 while he
                                                  # decides on recalculation; built and toy-gated, verdict recorded only
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

    def finish(self) -> dict:
        stats = {"formula_cells": self.n_formula, "formula_values_read": self.n_values_read,
                 "error_cells": self.n_error_cells, "error_cells_by_code": dict(sorted(self.by_code.items())),
                 "typed_error_constants": self.n_constant_errors,
                 "implicit_intersection_cells": self.n_ii, "implicit_intersection_checked": self.n_ii_checked,
                 "sheets_with_errors": self.sheets_with_errors, "hidden_sheets_with_errors": self.hidden_sheets,
                 "examples": self.examples}
        return self.verdict(f"No cell displays an Excel error value ({self.n_formula} formula cells checked).",
                            "{n} block(s) of cells display Excel error values "
                            f"({self.n_error_cells} cell(s): " +
                            ", ".join(f"{k} x{v}" for k, v in sorted(self.by_code.items())) + ").", stats)
