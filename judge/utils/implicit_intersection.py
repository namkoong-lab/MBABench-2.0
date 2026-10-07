"""Implicit-intersection errors: formulas that show #VALUE! in Excel but a
number everywhere else (judge v11, 2026-09-18).

Background. A formula written by a library (openpyxl, LibreOffice export,
ChatGPT's writer) is stored without the array marker (`<f t="array">`, the
Ctrl+Shift+Enter / dynamic-array flag Excel adds itself when a person types
the formula). Excel evaluates an unmarked formula under its legacy rules:
wherever a multi-cell range is used as a single value — as the operand of an
operator or the argument of a single-value function such as ABS or ROUND —
Excel picks the one cell of the range in the formula's own row or column
("implicit intersection", shown as `@` in current Excel). When the formula
sits outside the range there is no such cell and Excel returns #VALUE!.
LibreOffice and the Python engines that produce the cached values do the
block arithmetic instead, so the value the judge is served looks fine
(grading 1075, Sens_Engine!D448: cached 4.9e-08, Excel #VALUE!).

The rules below were probed against Excel for Mac 16.112 (155 formulas,
2026-09-18; see tests_offline/test_implicit_intersection.py, which carries
every probe with Excel's verdict). Precision is the design goal: a construct
that was not probed is treated as safe (no flag), never as an error.

Contexts while walking a formula's parse tree
  SCALAR    a multi-cell range here is implicitly intersected  -> flag when
            the formula cell lies outside the range
  ARRAY_SP  inside SUMPRODUCT: operators and single-value functions are
            array-evaluated
  AGG_TOP   argument of an aggregator (SUM, MAX, NPV, MATCH lookup_array,
            VLOOKUP table ...) outside SUMPRODUCT: a bare range is fine,
            operators and single-value functions inside get SCALAR
  AGG_ARR   the same aggregator inside SUMPRODUCT: operators stay
            array-evaluated, single-value functions inside still get SCALAR
            (SUMPRODUCT(MAX(ABS(r1-r2))) is #VALUE!, SUMPRODUCT(MAX(r1-r2))
            is fine)
  ARRAY_X   anything not probed, plus functions that swallow the error or
            never raise it (IFERROR, IS*, COUNT*, N, ROW, INDEX's array,
            LOOKUP's vectors, OFFSET, CHOOSE, dynamic-array functions ...):
            nothing inside is ever flagged
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Optional

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

SCALAR, ARRAY_SP, AGG_TOP, AGG_ARR, ARRAY_X = "SCALAR", "ARRAY_SP", "AGG_TOP", "AGG_ARR", "ARRAY_X"

# Every argument is array-evaluated (verified: SUMPRODUCT).
_ARRAY_SP_FUNCS = {"SUMPRODUCT"}

# Single-value ("value"-typed) functions. Verified in Excel: ABS ROUND SQRT
# EXP LN LEN TEXT INT MOD POWER SIGN NOT. The rest take the same kind of
# single-value parameters.
_VALUE_FUNCS = {
    "ABS", "ROUND", "SQRT", "EXP", "LN", "LEN", "TEXT", "INT", "MOD", "POWER", "SIGN", "NOT",
    "ROUNDUP", "ROUNDDOWN", "TRUNC", "LOG", "LOG10", "CEILING", "FLOOR", "MROUND", "QUOTIENT",
    "CEILING.MATH", "FLOOR.MATH", "CEILING.PRECISE", "FLOOR.PRECISE",
    "FACT", "SIN", "COS", "TAN", "ATAN", "ASIN", "ACOS", "DEGREES", "RADIANS",
    "VALUE", "LEFT", "RIGHT", "MID", "UPPER", "LOWER", "PROPER", "TRIM", "CLEAN", "SUBSTITUTE",
    "REPLACE", "REPT", "FIND", "SEARCH", "EXACT", "CHAR", "CODE", "CONCATENATE", "DOLLAR", "FIXED",
    "DATE", "DATEVALUE", "TIME", "TIMEVALUE", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
    "WEEKDAY", "WEEKNUM", "EOMONTH", "EDATE", "DAYS", "DAYS360", "YEARFRAC", "DATEDIF",
    "PMT", "IPMT", "PPMT", "PV", "FV", "RATE", "NPER", "EFFECT", "NOMINAL", "SLN", "DB", "DDB", "SYD",
    "ISPMT", "CUMIPMT", "CUMPRINC",
}

# Aggregators: a bare range is fine, an expression inside is not. Verified:
# SUM MAX MIN AVERAGE LARGE SMALL MEDIAN PRODUCT STDEV SUMSQ PERCENTILE AND OR
# NPV MAXA MINA MATCH VLOOKUP. The rest share the same parameter kind.
_AGG_FUNCS = {
    "SUM", "MAX", "MIN", "AVERAGE", "LARGE", "SMALL", "MEDIAN", "PRODUCT", "STDEV", "SUMSQ",
    "PERCENTILE", "AND", "OR", "NPV", "MAXA", "MINA", "MATCH", "VLOOKUP",
    "AVERAGEA", "STDEV.S", "STDEV.P", "STDEVP", "STDEVA", "STDEVPA", "VAR", "VAR.S", "VAR.P",
    "VARP", "VARA", "VARPA", "GEOMEAN", "HARMEAN", "PERCENTILE.INC", "PERCENTILE.EXC",
    "QUARTILE", "QUARTILE.INC", "QUARTILE.EXC", "MODE", "MODE.SNGL", "DEVSQ", "AVEDEV",
    "KURT", "SKEW", "SKEW.P", "XOR", "HLOOKUP",
}
# 1-based argument positions of aggregators that take a single value.
_AGG_VALUE_ARGS = {
    "NPV": {1}, "SMALL": {2}, "LARGE": {2},
    "PERCENTILE": {2}, "PERCENTILE.INC": {2}, "PERCENTILE.EXC": {2},
    "QUARTILE": {2}, "QUARTILE.INC": {2}, "QUARTILE.EXC": {2},
    "MATCH": {1, 3}, "VLOOKUP": {1, 3, 4}, "HLOOKUP": {1, 3, 4},
}
# VLOOKUP/HLOOKUP's table argument is never array-evaluated, even inside
# SUMPRODUCT (verified).
_AGG_TOP_ALWAYS = {("VLOOKUP", 2), ("HLOOKUP", 2)}

# INDEX's array and LOOKUP's vectors are array-evaluated (verified:
# MAX(INDEX(ABS(r1-r2),0)) and LOOKUP(2,1/(r>3),r2) are fine).
_ARRAY_X_ARGS = {"INDEX": {1}, "LOOKUP": {2, 3}}

_OP_CHILD = {SCALAR: SCALAR, ARRAY_SP: ARRAY_SP, ARRAY_X: ARRAY_X, AGG_TOP: SCALAR, AGG_ARR: AGG_ARR}

_RE_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_RE_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
_RE_ROW = re.compile(r"^\$?(\d+)$")

MAX_EXAMPLES = 10
_MAX_FORMULA_CHARS = 140


@dataclass
class Node:
    kind: str                      # range | other | func | op | paren | array
    text: str = ""
    name: str = ""                 # func: normalised name
    children: list = field(default_factory=list)


class _Unsupported(Exception):
    """A construct the walker does not model (intersection/union operators,
    range operators built from functions). The formula is skipped, never
    flagged."""


# ---------------------------------------------------------------------------
# parsing (openpyxl's tokenizer, a flat operator model — precedence does not
# matter because every operand of every operator gets the same context)
# ---------------------------------------------------------------------------

def _func_name(token_value: str) -> str:
    name = token_value[:-1].upper()
    while name.startswith(("_XLFN.", "_XLWS.")):
        name = name.split(".", 1)[1]
    return name


class _Parser:
    def __init__(self, formula: str):
        items = [t for t in Tokenizer(formula).items]
        self.toks = []
        prev_operandish = False
        for t in items:
            if t.type == "WHITE-SPACE":
                # a space between two operands is Excel's intersection operator
                self._pending_space = prev_operandish
                continue
            if getattr(self, "_pending_space", False) and (
                t.type == "OPERAND" or (t.type in ("FUNC", "PAREN", "ARRAY") and t.subtype == "OPEN")
            ):
                raise _Unsupported("intersection operator")
            self._pending_space = False
            self.toks.append(t)
            prev_operandish = t.type == "OPERAND" or (
                t.type in ("FUNC", "PAREN", "ARRAY") and t.subtype == "CLOSE"
            )
        self.i = 0

    def _peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def _next(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def parse(self) -> Node:
        node = self.expr()
        if self._peek() is not None:
            raise _Unsupported("trailing tokens")
        return node

    def expr(self) -> Node:
        operands = [self.unary()]
        ops = []
        while True:
            t = self._peek()
            if t is not None and t.type == "OPERATOR-INFIX":
                self._next()
                if t.value.strip() in (":", "", ","):
                    raise _Unsupported("range/union operator")
                ops.append(t.value)
                operands.append(self.unary())
            else:
                break
        if not ops:
            return operands[0]
        return Node("op", text="".join(ops), children=operands)

    def unary(self) -> Node:
        prefixes = []
        while (t := self._peek()) is not None and t.type == "OPERATOR-PREFIX":
            prefixes.append(self._next().value)
        node = self.primary()
        postfix = []
        while (t := self._peek()) is not None and t.type == "OPERATOR-POSTFIX":
            postfix.append(self._next().value)
        if prefixes or postfix:
            return Node("op", text="".join(prefixes + postfix), children=[node])
        return node

    def primary(self) -> Node:
        t = self._peek()
        if t is None:
            return Node("other", text="")          # empty argument
        if t.type == "OPERAND":
            self._next()
            if t.subtype == "RANGE":
                if t.value.startswith(":") or t.value.endswith(":"):
                    raise _Unsupported("range operator")
                return Node("range", text=t.value)
            return Node("other", text=t.value)
        if t.type == "FUNC" and t.subtype == "OPEN":
            self._next()
            args = []
            if (p := self._peek()) is not None and p.type == "FUNC" and p.subtype == "CLOSE":
                self._next()
                return Node("func", name=_func_name(t.value), children=args)
            args.append(self.expr())
            while (p := self._peek()) is not None and p.type == "SEP":
                self._next()
                args.append(self.expr())
            c = self._next() if self._peek() is not None else None
            if c is None or c.type != "FUNC" or c.subtype != "CLOSE":
                raise _Unsupported("unbalanced function")
            return Node("func", name=_func_name(t.value), children=args)
        if t.type == "PAREN" and t.subtype == "OPEN":
            self._next()
            inner = self.expr()
            c = self._next() if self._peek() is not None else None
            if c is None or c.type != "PAREN" or c.subtype != "CLOSE":
                raise _Unsupported("unbalanced parenthesis")
            return Node("paren", children=[inner])
        if t.type == "ARRAY" and t.subtype == "OPEN":
            depth = 0
            while (p := self._peek()) is not None:
                self._next()
                if p.type == "ARRAY":
                    depth += 1 if p.subtype == "OPEN" else -1
                    if depth == 0:
                        break
            return Node("array")
        if t.type == "SEP":
            return Node("other", text="")          # empty argument before a separator
        raise _Unsupported(f"token {t.type}/{t.subtype}")


# ---------------------------------------------------------------------------
# geometry
# ---------------------------------------------------------------------------

def _parse_range(text: str) -> Optional[tuple]:
    """(min_col, min_row, max_col, max_row) with None for whole rows/cols;
    None when the operand is not a plain A1-style reference (names, tables,
    external books, 3-D references are not modelled)."""
    if "[" in text or "]" in text:
        return None
    ref = text.rsplit("!", 1)[-1] if "!" in text else text
    if ":" in text.rsplit("!", 1)[0] if "!" in text else False:
        return None                                    # 3-D reference Sheet1:Sheet3!A1
    parts = ref.split(":")
    if len(parts) > 2 or not all(parts):
        return None
    if len(parts) == 1:
        m = _RE_CELL.match(parts[0])
        if not m:
            return None
        c, r = column_index_from_string(m.group(1).upper()), int(m.group(2))
        return (c, r, c, r)
    a, b = parts
    ma, mb = _RE_CELL.match(a), _RE_CELL.match(b)
    if ma and mb:
        c1, r1 = column_index_from_string(ma.group(1).upper()), int(ma.group(2))
        c2, r2 = column_index_from_string(mb.group(1).upper()), int(mb.group(2))
        return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))
    if _RE_COL.match(a) and _RE_COL.match(b):
        c1, c2 = column_index_from_string(a.strip("$").upper()), column_index_from_string(b.strip("$").upper())
        return (min(c1, c2), None, max(c1, c2), None)
    if _RE_ROW.match(a) and _RE_ROW.match(b):
        r1, r2 = int(a.strip("$")), int(b.strip("$"))
        return (None, min(r1, r2), None, max(r1, r2))
    return None


def intersection_impossible(bounds: tuple, col: int, row: int) -> bool:
    """True when Excel's implicit intersection of `bounds` from the formula
    cell (col, row) has no cell to return. Whole rows/columns always
    intersect; a single cell is not a range."""
    c1, r1, c2, r2 = bounds
    if r1 is None or c1 is None:
        return False                                    # whole column / whole row
    if c1 == c2 and r1 == r2:
        return False                                    # single cell
    if c1 == c2:                                        # one column: same row needed
        return not (r1 <= row <= r2)
    if r1 == r2:                                        # one row: same column needed
        return not (c1 <= col <= c2)
    return not (r1 <= row <= r2 and c1 <= col <= c2)    # 2-D block: both needed


# ---------------------------------------------------------------------------
# walk
# ---------------------------------------------------------------------------

def _walk(node: Node, ctx: str, col: int, row: int, via: str, hits: list) -> None:
    if node.kind == "range":
        if ctx == SCALAR:
            b = _parse_range(node.text)
            if b is not None and intersection_impossible(b, col, row):
                hits.append({"range": node.text, "via": via})
        return
    if node.kind in ("other", "array"):
        return
    if node.kind == "paren":
        _walk(node.children[0], ctx, col, row, via, hits)
        return
    if node.kind == "op":
        child = _OP_CHILD[ctx]
        for ch in node.children:
            _walk(ch, child, col, row, node.text, hits)
        return
    # function call
    name = node.name
    for i, arg in enumerate(node.children, 1):
        if name in _ARRAY_SP_FUNCS:
            child = ARRAY_SP
        elif name == "IF":
            child = ARRAY_X if ctx == ARRAY_X else SCALAR
        elif name in _ARRAY_X_ARGS:
            child = ARRAY_X if (i in _ARRAY_X_ARGS[name] or ctx == ARRAY_X) else (
                ARRAY_SP if ctx == ARRAY_SP else SCALAR)
        elif name in _AGG_FUNCS:
            if ctx == ARRAY_X:
                child = ARRAY_X
            elif i in _AGG_VALUE_ARGS.get(name, ()):
                child = ARRAY_SP if ctx == ARRAY_SP else SCALAR
            elif (name, i) in _AGG_TOP_ALWAYS:
                child = AGG_TOP
            else:
                child = AGG_ARR if ctx in (ARRAY_SP, AGG_ARR) else AGG_TOP
        elif name in _VALUE_FUNCS:
            child = ARRAY_SP if ctx == ARRAY_SP else (ARRAY_X if ctx == ARRAY_X else SCALAR)
        else:
            child = ARRAY_X                             # not modelled: never flag inside
        _walk(arg, child, col, row, name, hits)


def check_formula(formula: str, coordinate: str) -> list[dict]:
    """Ranges of a plain (non-array) formula that Excel would implicitly
    intersect from outside them. Empty list = nothing flagged (also for any
    construct the walker does not model)."""
    if not isinstance(formula, str) or not formula.startswith("=") or ":" not in formula:
        return []
    try:
        col_letters, row = coordinate_from_string(coordinate)
        col = column_index_from_string(col_letters)
        tree = _Parser(formula).parse()
    except _Unsupported:
        return []
    except Exception:  # noqa: BLE001 - tokenizer quirks: never flag what we cannot read
        return []
    hits: list[dict] = []
    try:
        _walk(tree, SCALAR, col, row, "formula", hits)
    except Exception:  # noqa: BLE001
        return []
    return hits


def scan_sheet(ws, bounds: Optional[dict] = None) -> dict:
    """{"count": n, "examples": [{"cell", "formula", "ranges", "via"}, ...]}
    over the plain formulas of a worksheet (array and data-table formulas are
    skipped: Excel array-evaluates those)."""
    examples: list[dict] = []
    count = 0
    for row in ws.iter_rows(**(bounds or {})):
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=") and ":" in v):
                continue
            hits = check_formula(v, cell.coordinate)
            if not hits:
                continue
            count += 1
            if len(examples) < MAX_EXAMPLES:
                seen, ranges = set(), []
                for h in hits:
                    if h["range"] not in seen:
                        seen.add(h["range"])
                        ranges.append(h["range"])
                examples.append({
                    "cell": cell.coordinate,
                    "formula": v if len(v) <= _MAX_FORMULA_CHARS else v[:_MAX_FORMULA_CHARS] + "…",
                    "ranges": ranges,
                    "via": hits[0]["via"],
                })
    return {"count": count, "examples": examples}


MAX_DEPENDENTS = 8


def find_dependents(workbook, flagged: dict, bounds_for=None) -> dict:
    """{(sheet, cell): ["Checks!D19", ...]}: formulas anywhere in the workbook that reference a
    flagged cell directly (judge v12). `flagged` maps sheet title -> cell coordinates;
    `bounds_for(ws)` returns the iter_rows kwargs of the sheet-extent guard.

    The cached value of a flagged cell looks fine, so everything downstream of it looks fine
    too: grading 1092 passed Master error flag while Sens_Engine!D448 (#VALUE! in Excel) fed
    Checks!F19 and the roll-up COUNTIF(F6:F42,FALSE) skipped the erroring row. Only single-cell
    operands count; a range that merely contains the cell (the COUNTIF above) is not listed,
    because whether it passes the error on depends on the function around it."""
    targets = {(s, c.replace("$", "").upper()) for s, cells in (flagged or {}).items() for c in cells}
    if not targets:
        return {}
    needles = {c for _, c in targets}
    quick = re.compile("|".join(
        r"(?<![A-Za-z$])\$?" + re.escape(m.group(1)) + r"\$?" + re.escape(m.group(2)) + r"(?![0-9])"
        for m in (_RE_CELL.match(c) for c in sorted(needles)) if m), re.I)
    found: dict = {t: [] for t in targets}
    for ws in getattr(workbook, "worksheets", []):
        kwargs = bounds_for(ws) if bounds_for else {}
        for row in ws.iter_rows(**(kwargs or {})):
            for cell in row:
                v = cell.value
                text = v if isinstance(v, str) else getattr(v, "text", None)
                if not (isinstance(text, str) and text.startswith("=") and quick.search(text)):
                    continue
                try:
                    tokens = Tokenizer(text).items
                except Exception:  # noqa: BLE001 - tokenizer quirks: a dependent we cannot read is not listed
                    continue
                for tok in tokens:
                    if tok.type != "OPERAND" or tok.subtype != "RANGE" or ":" in tok.value.rsplit("!", 1)[-1]:
                        continue
                    sheet, _, ref = tok.value.rpartition("!")
                    sheet = sheet.strip("'").replace("''", "'") if sheet else ws.title
                    key = (sheet, ref.replace("$", "").upper())
                    here = f"{ws.title}!{cell.coordinate}"
                    if key in found and here not in found[key] and key != (ws.title, cell.coordinate):
                        found[key].append(here)
    return {k: v for k, v in found.items() if v}


def render_lines(ii: Any) -> list[str]:
    """Properties-block lines for one sheet (empty when nothing was flagged)."""
    if not isinstance(ii, dict) or not ii.get("count"):
        return []
    n = int(ii.get("count") or 0)
    ex = ii.get("examples") or []
    parts = []
    for e in ex:
        via = e.get("via") or "formula"
        how = f"as a single value in {via}( )" if via.isalpha() else f"as a single value beside '{via}'"
        part = f"{e['cell']} {e['formula']} — uses {', '.join(e['ranges'])} {how}"
        deps = e.get("referenced_by") or []
        if deps:
            shown = ", ".join(deps[:MAX_DEPENDENTS])
            extra = f" (+{len(deps) - MAX_DEPENDENTS} more)" if len(deps) > MAX_DEPENDENTS else ""
            part += f" — referenced by {shown}{extra}: these cells show the error in Excel too"
        parts.append(part)
    more = f"; (+{n - len(ex)} more)" if n > len(ex) else ""
    return [
        f"     IMPLICIT INTERSECTION — {n} plain formula{'s' if n != 1 else ''} without the array marker "
        f"{'use' if n != 1 else 'uses'} a multi-cell range as a single value from a cell outside that range; "
        "Excel shows #VALUE! there (the served value comes from a non-Excel recalculation and is not what "
        "Excel displays): " + "; ".join(parts) + more
    ]
