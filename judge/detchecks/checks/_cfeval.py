"""Conditional-format rules evaluated for one cell - shared by Negatives in parentheses (65), Zeros as
dashes (66), Sufficient column widths (69) and No white-on-white hiding (94).  Patrick, 2026-10-05:
"CONDITIONAL FORMATS never stop a grading".

    env = cf_env(wb, sheet)                        # the sheet the rules belong to
    prefetch(env, tail.conditional_formats)        # register the cells the rules read (core.lookup)
    expression_fires(formula, env, where, value)   # an 'expression' rule  -> True / False / UNKNOWN / Unevaluable
    cellis_fires(rule, value, env, where)          # a 'cellIs' rule       -> the same
    cellis_operands(rule, env, where, value)       # its operands' values  -> [values] / UNKNOWN / Unevaluable
    formula_value(formula, env, where, value)      # any formula's value   -> value / UNKNOWN / Unevaluable

where = (row, col, anchor_row, anchor_col): the cell being judged and the top-left cell of the rule's
first range.  Relative references in a CF formula are relative to that anchor, so for each cell of
the range they are shifted by (row - anchor_row, col - anchor_col) (wrapping round the grid, as
Excel does); absolute parts ($) stay.  where None: the caller screens a whole range, so a
position-dependent formula (a relative reference, ROW(), COLUMN()) is UNKNOWN.
value: the judged cell's own value - a reference to the cell itself reads it - or UNKNOWN_VALUE when
the caller screens hypothetical values (then a self-reference is UNKNOWN).
Every other referenced cell - same sheet or another one - is read through core.lookup, i.e. through
the grading's value source exactly as the checks read values: constants, Excel's caches, the
recalculation copy; an untrusted formula result is Unevaluable, never guessed.

Evaluated (Excel semantics):
  literals            numbers, "text", TRUE / FALSE, error literals (#N/A ...)
  references          A1, $A$1, $A1, A$1, Sheet!A1, 'My Sheet'!$B$2; a defined name that is a
                      constant or one absolute cell (CheckTolerance -> Assumptions!$F$12)
  operators           unary + -, %, ^, * /, + -, & (text), = <> < > <= >= (numbers < text < logicals,
                      text case-insensitive, numbers to 15 significant digits; an empty cell is 0 / "" /
                      FALSE against a number / text / logical); errors propagate; text that is a plain
                      number is coerced in arithmetic ("5" + 1)
  functions           LEFT RIGHT MID LEN UPPER LOWER TRIM VALUE TEXT (through core.numfmt) EXACT SEARCH
                      (wildcards, as containsText) FIND CONCATENATE ISNUMBER ISTEXT ISBLANK ISERROR ISERR
                      ISNA ISLOGICAL ISNONTEXT AND OR NOT IF IFERROR ABS ROUND ROUNDUP ROUNDDOWN INT TRUNC
                      MOD ISEVEN ISODD ROW COLUMN
  a rule fires when its formula gives TRUE or a non-zero number (an error, FALSE, 0 or an empty cell: no).
Unevaluable(why): anything else - a range or whole-column operand, another function (COUNTIF, YEAR,
INDIRECT ...), an external or 3-D reference, an array constant, an untrusted or unavailable referenced
value, text that may be a date, a formula that returns text, a rule type other than cellIs /
containsText / notContainsText / beginsWith / endsWith / containsBlanks / notContainsBlanks /
containsErrors / notContainsErrors / expression (top10, aboveAverage, duplicateValues, timePeriod ...).
What an unevaluable rule means is the check's policy (Patrick 2026-10-05): No white-on-white hiding
(94) assumes the text is visible; the other checks grade the cell by its own formatting as if the rule
were off.  Each check records every such assumption in its stats.
"""
from __future__ import annotations

import math
import re
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Context, Decimal, InvalidOperation
from functools import lru_cache
from typing import Optional

from ..core import formula as F
from ..core import numfmt as N
from ..core.lookup import BLANK, Unavailable, cell_values
from ..core.refs import MAX_COL, MAX_ROW
from ..core.sheet import ExcelError
from ..errors import GradingError
from ._cftext import _search_rx

UNKNOWN = object()          # depends on the cell's value or position, which the caller does not know: branch
UNKNOWN_VALUE = object()    # the judged cell's own value, when the caller does not know it (a screen)
MAX_NAME_DEPTH = 5


class Unevaluable:
    """The result for a rule this evaluator cannot evaluate, with the reason."""
    __slots__ = ("why",)

    def __init__(self, why: str):
        self.why = why

    def __repr__(self):
        return f"Unevaluable({self.why!r})"

    def __bool__(self):
        raise TypeError("an Unevaluable rule result has no truth value; apply the check's policy")


def is_unevaluable(x) -> bool:
    return isinstance(x, Unevaluable)


class _Dependent(Exception):
    """Position or value dependence (-> UNKNOWN)."""


class _NoEval(Exception):
    """Not evaluable (-> Unevaluable(why))."""


# ---------------------------------------------------------------------------- environment
class CfEnv:
    """What the evaluator needs about a sheet's rules: the sheet's name, the workbook's cell values
    (core.lookup; None: no other cell can be read) and its defined names."""

    def __init__(self, sheet: str, values=None, defined_names=(), date1904: bool = False):
        self.sheet = sheet
        self.values = values
        self.date1904 = date1904
        self._names: dict = {}
        for d in defined_names or ():
            if getattr(d, "builtin", False) or not getattr(d, "name", None):
                continue
            try:
                scope = d.scope
            except GradingError:
                continue                                  # broken localSheetId: never used
            self._names.setdefault(d.name.upper(), {})[scope] = d.text or ""

    def name_text(self, name: str, sheet: Optional[str]) -> Optional[str]:
        """The definition of `name` as seen from `sheet`: a sheet-scoped name first, then the workbook's."""
        by_scope = self._names.get(name.upper())
        if not by_scope:
            return None
        for scope, text in by_scope.items():
            if scope is not None and sheet is not None and scope.casefold() == sheet.casefold():
                return text
        return by_scope.get(None)


def cf_env(wb, sheet: str) -> CfEnv:
    """The environment for the conditional formats of `sheet` in workbook `wb` (a Package)."""
    return CfEnv(sheet, cell_values(wb), getattr(wb, "defined_names", ()) or (), bool(getattr(wb, "date1904", False)))


class _Cx:
    __slots__ = ("env", "row", "col", "r0", "c0", "value", "depth")

    def __init__(self, env, where, value, depth=0):
        self.env = env
        if where is None:
            self.row = self.col = self.r0 = self.c0 = None
        else:
            self.row, self.col, self.r0, self.c0 = where
        self.value = value
        self.depth = depth


# ---------------------------------------------------------------------------- parsing
_CMP_OPS = ("=", "<>", "<", ">", "<=", ">=")


class _ParseError(Exception):
    pass


class _Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else None

    def take(self):
        if self.i >= len(self.t):
            raise _ParseError("unexpected end of formula")
        tok = self.t[self.i]
        self.i += 1
        return tok

    def _op(self, *ops):
        tok = self.peek()
        return tok is not None and tok[0] == "op" and tok[1] in ops

    def expr(self):
        a = self.concat()
        while self._op(*_CMP_OPS):
            op = self.take()[1]
            a = ("cmp", op, a, self.concat())
        return a

    def concat(self):
        a = self.additive()
        while self._op("&"):
            self.take()
            a = ("cat", a, self.additive())
        return a

    def additive(self):
        a = self.mult()
        while self._op("+", "-"):
            op = self.take()[1]
            a = ("bin", op, a, self.mult())
        return a

    def mult(self):
        a = self.power()
        while self._op("*", "/"):
            op = self.take()[1]
            a = ("bin", op, a, self.power())
        return a

    def power(self):
        a = self.unary()
        while self._op("^"):
            self.take()
            a = ("bin", "^", a, self.unary())
        return a

    def unary(self):                      # Excel: negation binds tighter than ^ (-2^2 = 4)
        if self._op("-"):
            self.take()
            return ("neg", self.unary())
        if self._op("+"):
            self.take()
            return self.unary()
        return self.postfix()

    def postfix(self):
        a = self.primary()
        while self._op("%"):
            self.take()
            a = ("pct", a)
        return a

    def primary(self):
        kind, text = self.take()[:2]
        if kind == "num":
            return ("lit", float(text))
        if kind == "str":
            return ("lit", text[1:-1].replace('""', '"'))
        if kind == "err":
            return ("lit", ExcelError(text.upper()))
        if kind == "(":
            x = self.expr()
            if self.take()[0] != ")":
                raise _ParseError("unbalanced parenthesis")
            return x
        if kind == "func":
            name, _prefix, local = F._norm_func(text)
            if self.take()[0] != "(":
                raise _ParseError(f"no '(' after {text}")
            args = []
            if self.peek() is not None and self.peek()[0] == ")":
                self.take()
            else:
                while True:
                    tok = self.peek()
                    if tok is not None and tok[0] in (",", ")"):
                        args.append(("missing",))
                    else:
                        args.append(self.expr())
                    sep = self.take()[0]
                    if sep == ")":
                        break
                    if sep != ",":
                        raise _ParseError(f"unexpected {sep!r} in the arguments of {name}")
            if local:
                return ("unsupported", f"function {text} is not evaluated")
            return ("call", name, tuple(args))
        if kind == "operand":
            up = text.upper()
            if up in ("TRUE", "FALSE"):
                return ("lit", up == "TRUE")
            try:
                info = F._ref_info(text)
            except GradingError as e:
                return ("unsupported", f"operand {text!r} cannot be read ({e})")
            q = info.qualifier
            if q is not None and (q.external or q.sheet_end or q.workbook_scoped or q.ambiguous_book):
                return ("unsupported", f"reference {text!r} to another workbook or across sheets is not evaluated")
            sheet = q.sheet if q is not None else None
            if info.kind == "cell" and not info.r1c1:
                return ("ref", sheet, info.parts[0])
            if info.kind == "name":
                return ("name", sheet, info.name)
            if info.kind == "error":
                return ("lit", ExcelError("#REF!"))
            return ("unsupported", f"{info.kind.replace('_', ' ')} operand {text!r} is not evaluated")
        raise _ParseError(f"unexpected {text!r}")


@lru_cache(maxsize=8192)
def parse(formula: str):
    """The formula's syntax tree (tuples), ('unsupported', why) when it cannot be read."""
    body = (formula or "").strip()
    if body.startswith("="):
        body = body[1:].strip()
    if not body:
        return ("unsupported", "empty formula")
    try:
        toks = [t for t in F.lex(body) if t[0] != "ws"]
    except GradingError as e:
        return ("unsupported", f"formula {formula!r} cannot be read ({e})")
    p = _Parser(toks)
    try:
        node = p.expr()
    except _ParseError as e:
        return ("unsupported", f"formula {formula!r} is not evaluated ({e})")
    except RecursionError:
        return ("unsupported", f"formula {formula!r} is nested too deeply")
    if p.i != len(toks):
        return ("unsupported", f"formula {formula!r} is not evaluated (unexpected {toks[p.i][1]!r})")
    return node


# ---------------------------------------------------------------------------- values
def _is_err(v) -> bool:
    return isinstance(v, ExcelError)


def _rank(v) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def _r15(x: float) -> float:
    return float(f"{x:.15g}") if math.isfinite(x) else x


def compare(a, b):
    """Excel's comparison of two values: -1 / 0 / 1, or the first error.  Numbers < text < logicals;
    text case-insensitive; numbers to 15 significant digits; an empty cell (BLANK) is 0, "" or FALSE
    against a number, text or logical (and equals another empty cell)."""
    if _is_err(a):
        return a
    if _is_err(b):
        return b
    if a is BLANK and b is BLANK:
        return 0
    if a is BLANK:
        a = (0.0, "", False)[_rank(b)]
    if b is BLANK:
        b = (0.0, "", False)[_rank(a)]
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 0:
        x, y = _r15(float(a)), _r15(float(b))
    elif ra == 1:
        x, y = a.casefold(), b.casefold()
    else:
        x, y = bool(a), bool(b)
    return (x > y) - (x < y)


_CMP_FN = {"=": lambda k: k == 0, "<>": lambda k: k != 0, "<": lambda k: k < 0, ">": lambda k: k > 0,
           "<=": lambda k: k <= 0, ">=": lambda k: k >= 0}
_NUM_TEXT_RX = re.compile(r"^([+-]?)\$?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?\s*(%?)$")


def _text_number(s: str):
    """A text used as a number: a plain number ('5', ' 1,000.5 ', '-2E3', '5%', '$5', '(5)') -> float;
    a text without a digit -> #VALUE!; anything else with a digit (a date '1/2/2026', a time, '5 kg')
    -> not evaluated (Excel may read it as a date / time)."""
    t = s.strip()
    neg = False
    if len(t) >= 2 and t[0] == "(" and t[-1] == ")":
        t, neg = t[1:-1].strip(), True
    m = _NUM_TEXT_RX.match(t)
    if m:
        sign, digits, exp, pct = m.groups()
        try:
            x = float(digits.replace(",", "")) * (10.0 ** int(exp) if exp else 1.0)
        except (ValueError, OverflowError):
            raise _NoEval(f"text {s!r} as a number") from None
        if sign == "-":
            x = -x
        if pct:
            x /= 100.0
        return -x if neg else x
    if not any(ch.isdigit() for ch in t):
        return ExcelError("#VALUE!")
    raise _NoEval(f"text {s!r} used as a number may be a date or time")


def _num(v):
    """A value used as a number: float, or an error."""
    if _is_err(v):
        return v
    if v is BLANK:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        return _text_number(v)
    raise _NoEval(f"value {v!r} used as a number")


def _num_text(x: float) -> str:
    """Excel's text for a number in a formula (15 significant digits, no grouping)."""
    if not math.isfinite(x):
        raise _NoEval("a non-finite number")
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    if not 1e-5 <= abs(x) < 1e15:
        raise _NoEval(f"the text of the number {x!r} (scientific notation) is not evaluated")
    d = Context(prec=15, rounding=ROUND_HALF_UP).create_decimal(repr(x))
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def _text(v):
    """A value used as text: str, or an error."""
    if _is_err(v):
        return v
    if v is BLANK:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return _num_text(float(v))
    if isinstance(v, str):
        return v
    raise _NoEval(f"value {v!r} used as text")


def _logical(v, from_ref: bool):
    """A value used as a logical: bool, None (skip: text or an empty cell read through a reference, for
    AND / OR), or an error."""
    if _is_err(v):
        return v
    if isinstance(v, bool):
        return v
    if v is BLANK:
        return None if from_ref else False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        if from_ref:
            return None
        if v.strip().upper() in ("TRUE", "FALSE"):
            raise _NoEval(f"the text {v!r} used as a logical value")
        return ExcelError("#VALUE!")
    raise _NoEval(f"value {v!r} used as a logical value")


def _decimal_round(x: float, digits: float, mode) -> float:
    try:
        n = int(math.floor(digits)) if digits >= 0 else -int(math.floor(-digits))
        d = Decimal(repr(x))
        q = d.quantize(Decimal(1).scaleb(-n), rounding=mode)
        return float(q)
    except (InvalidOperation, ValueError, OverflowError):
        raise _NoEval(f"rounding {x!r} to {digits!r} digits") from None


def _int(v) -> int:
    return int(math.trunc(v))


# ---------------------------------------------------------------------------- evaluation
def _shift(part, cx, *, row: bool) -> int:
    v = part.row if row else part.col
    if (part.row_abs if row else part.col_abs):
        return v
    pos, anchor = (cx.row, cx.r0) if row else (cx.col, cx.c0)
    if pos is None or anchor is None:
        raise _Dependent()
    size = MAX_ROW if row else MAX_COL
    return (v + (pos - anchor) - 1) % size + 1          # Excel wraps relative references round the grid


def _ref_value(sheet, part, cx):
    rr = _shift(part, cx, row=True)
    cc = _shift(part, cx, row=False)
    env = cx.env
    own = env.sheet if env is not None else None
    same_sheet = sheet is None or (own is not None and sheet.casefold() == own.casefold())
    if same_sheet and cx.row is not None and (rr, cc) == (cx.row, cx.col):
        if cx.value is UNKNOWN_VALUE:
            raise _Dependent()
        return cx.value
    if env is None or env.values is None:
        raise _NoEval("the value of another cell is needed and no value source is available")
    target = own if sheet is None else sheet
    if target is None:
        raise _NoEval("the rule's sheet is unknown")
    v = env.values.get(target, rr, cc)
    if isinstance(v, Unavailable):
        raise _NoEval(v.why)
    return v


def _name_value(sheet, name, cx):
    env = cx.env
    if env is None:
        raise _NoEval(f"defined name {name} (no workbook)")
    text = env.name_text(name, sheet or env.sheet)
    if text is None:
        return ExcelError("#NAME?")
    if cx.depth >= MAX_NAME_DEPTH:
        raise _NoEval(f"defined name {name} refers to names too deeply")
    node = parse(text)
    if node[0] not in ("lit", "ref", "name", "neg") or (node[0] == "neg" and node[1][0] != "lit"):
        raise _NoEval(f"defined name {name} = {text!r} is not a constant or a single cell")
    if node[0] == "ref":
        part = node[2]
        if not (part.row_abs and part.col_abs):
            raise _NoEval(f"defined name {name} = {text!r} has a relative reference")
        if node[1] is None:
            raise _NoEval(f"defined name {name} = {text!r} has no sheet")
    # a name is evaluated on its own: no position, the cell's own value never involved
    return _eval(node, _Cx(env, None, UNKNOWN_VALUE, cx.depth + 1))


def _first_err(*vals):
    for v in vals:
        if _is_err(v):
            return v
    return None


def _eval(node, cx):
    k = node[0]
    if k == "lit":
        return node[1]
    if k == "ref":
        return _ref_value(node[1], node[2], cx)
    if k == "name":
        return _name_value(node[1], node[2], cx)
    if k == "missing":
        return BLANK
    if k == "unsupported":
        raise _NoEval(node[1])
    if k == "neg":
        v = _num(_eval(node[1], cx))
        return v if _is_err(v) else -v
    if k == "pct":
        v = _num(_eval(node[1], cx))
        return v if _is_err(v) else v / 100.0
    if k == "bin":
        op = node[1]
        a = _num(_eval(node[2], cx))
        b = _num(_eval(node[3], cx))
        e = _first_err(a, b)
        if e is not None:
            return e
        try:
            if op == "+":
                return a + b
            if op == "-":
                return a - b
            if op == "*":
                return a * b
            if op == "/":
                return ExcelError("#DIV/0!") if b == 0 else a / b
            if a == 0 and b <= 0:
                return ExcelError("#NUM!" if b == 0 else "#DIV/0!")
            r = a ** b
            return ExcelError("#NUM!") if isinstance(r, complex) or not math.isfinite(r) else r
        except (OverflowError, ZeroDivisionError):
            return ExcelError("#NUM!")
    if k == "cat":
        a = _text(_eval(node[1], cx))
        b = _text(_eval(node[2], cx))
        e = _first_err(a, b)
        return e if e is not None else a + b
    if k == "cmp":
        c = compare(_eval(node[2], cx), _eval(node[3], cx))
        return c if _is_err(c) else _CMP_FN[node[1]](c)
    if k == "call":
        return _call(node[1], node[2], cx)
    raise _NoEval(f"node {k}")


def _args(name, args, lo, hi):
    if not lo <= len(args) <= hi:
        raise _NoEval(f"{name} with {len(args)} argument(s)")


def _call(name, args, cx):
    ev = (lambda i: _eval(args[i], cx))
    if name in ("ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR", "ISERR", "ISNA", "ISLOGICAL", "ISNONTEXT"):
        _args(name, args, 1, 1)
        v = ev(0)
        if name == "ISNUMBER":
            return isinstance(v, (int, float)) and not isinstance(v, bool)
        if name == "ISTEXT":
            return isinstance(v, str) and not _is_err(v)
        if name == "ISNONTEXT":
            return not (isinstance(v, str) and not _is_err(v))
        if name == "ISBLANK":
            return v is BLANK and args[0][0] in ("ref", "name")
        if name == "ISERROR":
            return _is_err(v)
        if name == "ISERR":
            return _is_err(v) and str(v) != "#N/A"
        if name == "ISNA":
            return _is_err(v) and str(v) == "#N/A"
        return isinstance(v, bool)                          # ISLOGICAL
    if name in ("AND", "OR"):
        if not args:
            raise _NoEval(f"{name}()")
        vals = []
        for a in args:
            v = _logical(_eval(a, cx), a[0] in ("ref", "name"))
            if _is_err(v):
                return v
            if v is not None:
                vals.append(v)
        if not vals:
            return ExcelError("#VALUE!")
        return all(vals) if name == "AND" else any(vals)
    if name == "NOT":
        _args(name, args, 1, 1)
        v = _logical(ev(0), False)
        return v if _is_err(v) else not v
    if name == "IF":
        _args(name, args, 1, 3)
        c = _logical(ev(0), False)
        if _is_err(c):
            return c
        if c:
            return ev(1) if len(args) > 1 else True
        return ev(2) if len(args) > 2 else False
    if name == "IFERROR":
        _args(name, args, 2, 2)
        v = ev(0)
        return ev(1) if _is_err(v) else v
    if name in ("LEFT", "RIGHT"):
        _args(name, args, 1, 2)
        t = _text(ev(0))
        n = _num(ev(1)) if len(args) > 1 and args[1][0] != "missing" else (0.0 if len(args) > 1 else 1.0)
        e = _first_err(t, n)
        if e is not None:
            return e
        if n < 0:
            return ExcelError("#VALUE!")
        n = _int(n)
        if name == "LEFT":
            return t[:n]
        return t[-n:] if n else ""
    if name == "MID":
        _args(name, args, 3, 3)
        t, s, n = _text(ev(0)), _num(ev(1)), _num(ev(2))
        e = _first_err(t, s, n)
        if e is not None:
            return e
        if s < 1 or n < 0:
            return ExcelError("#VALUE!")
        s, n = _int(s), _int(n)
        return t[s - 1:s - 1 + n]
    if name in ("LEN", "UPPER", "LOWER", "TRIM"):
        _args(name, args, 1, 1)
        t = _text(ev(0))
        if _is_err(t):
            return t
        if name == "LEN":
            return float(len(t))
        if name == "UPPER":
            return t.upper()
        if name == "LOWER":
            return t.lower()
        return " ".join(x for x in t.split(" ") if x)       # TRIM: ASCII spaces only, runs collapsed
    if name in ("CONCATENATE",):
        out = []
        for a in args:
            t = _text(_eval(a, cx))
            if _is_err(t):
                return t
            out.append(t)
        return "".join(out)
    if name == "EXACT":
        _args(name, args, 2, 2)
        a, b = _text(ev(0)), _text(ev(1))
        e = _first_err(a, b)
        return e if e is not None else a == b
    if name in ("SEARCH", "FIND"):
        _args(name, args, 2, 3)
        needle, hay = _text(ev(0)), _text(ev(1))
        start = _num(ev(2)) if len(args) > 2 else 1.0
        e = _first_err(needle, hay, start)
        if e is not None:
            return e
        start = _int(start)
        if start < 1 or start > len(hay) + (1 if not needle else 0):
            return ExcelError("#VALUE!")
        if name == "FIND":
            i = hay.find(needle, start - 1)
            return ExcelError("#VALUE!") if i < 0 else float(i + 1)
        if not needle:
            return float(start)
        m = _search_rx(needle.casefold()).search(hay.casefold(), start - 1)
        return ExcelError("#VALUE!") if m is None else float(m.start() + 1)
    if name == "VALUE":
        _args(name, args, 1, 1)
        v = ev(0)
        if _is_err(v) or (isinstance(v, (int, float)) and not isinstance(v, bool)):
            return v
        if isinstance(v, bool):
            return ExcelError("#VALUE!")
        return _text_number(_text(v))
    if name == "TEXT":
        _args(name, args, 2, 2)
        v, fmt = ev(0), _text(ev(1))
        e = _first_err(v, fmt)
        if e is not None:
            return e
        if "*" in fmt:
            raise _NoEval(f"TEXT with a fill character in {fmt!r}")
        if isinstance(v, str) or v is BLANK or isinstance(v, bool):
            raise _NoEval("TEXT of a value that is not a number")
        try:
            r = N.render(float(v), fmt, date1904=cx.env.date1904 if cx.env is not None else False)
        except GradingError as ex:
            raise _NoEval(f"TEXT format {fmt!r}: {ex}") from None
        if not r.certain:
            raise _NoEval(f"TEXT under {fmt!r} is not verified")
        return r.text
    if name in ("ABS", "INT"):
        _args(name, args, 1, 1)
        x = _num(ev(0))
        if _is_err(x):
            return x
        return abs(x) if name == "ABS" else float(math.floor(x))
    if name in ("ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"):
        _args(name, args, 1 if name == "TRUNC" else 2, 2)
        x = _num(ev(0))
        n = _num(ev(1)) if len(args) > 1 else 0.0
        e = _first_err(x, n)
        if e is not None:
            return e
        mode = {"ROUND": ROUND_HALF_UP, "ROUNDUP": ROUND_UP}.get(name, ROUND_DOWN)
        return _decimal_round(x, n, mode)
    if name == "MOD":
        _args(name, args, 2, 2)
        a, b = _num(ev(0)), _num(ev(1))
        e = _first_err(a, b)
        if e is not None:
            return e
        if b == 0:
            return ExcelError("#DIV/0!")
        return a - b * math.floor(a / b)
    if name in ("ISEVEN", "ISODD"):
        _args(name, args, 1, 1)
        x = _num(ev(0))
        if _is_err(x):
            return x
        even = _int(x) % 2 == 0
        return even if name == "ISEVEN" else not even
    if name in ("ROW", "COLUMN"):
        _args(name, args, 0, 1)
        is_row = name == "ROW"
        if not args:
            pos = cx.row if is_row else cx.col
            if pos is None:
                raise _Dependent()
            return float(pos)
        a = args[0]
        if a[0] != "ref":
            raise _NoEval(f"{name} of something other than one cell")
        return float(_shift(a[2], cx, row=is_row))
    raise _NoEval(f"function {name} is not evaluated")


def _run(formula: str, env, where, value):
    """The formula's value, or raises _Dependent / _NoEval."""
    node = parse(formula)
    return _eval(node, _Cx(env, where, value))


def formula_value(formula: Optional[str], env=None, where=None, value=UNKNOWN_VALUE):
    """The value of a CF formula for one cell (module doc): a value, UNKNOWN or Unevaluable."""
    if formula is None:
        return Unevaluable("no formula")
    try:
        return _run(formula, env, where, value)
    except _Dependent:
        return UNKNOWN
    except _NoEval as e:
        return Unevaluable(str(e))
    except RecursionError:
        return Unevaluable(f"formula {formula!r} is nested too deeply")
    except Exception as e:  # noqa: BLE001 - an arithmetic corner (overflow ...) or an evaluator bug: never a guess,
        #                     never a stopped grading - the rule is unevaluable, the check's policy holds (recorded)
        return Unevaluable(f"formula {formula!r} could not be evaluated ({type(e).__name__}: {e})")


def expression_fires(formula: Optional[str], env=None, where=None, value=UNKNOWN_VALUE):
    """Does an 'expression' rule fire on the cell?  True / False / UNKNOWN / Unevaluable."""
    v = formula_value(formula, env, where, value)
    if v is UNKNOWN or is_unevaluable(v):
        return v
    return truth(v, formula)


def truth(v, formula=None):
    """A CF formula's result as 'fires': TRUE or a non-zero number; an error, FALSE, 0, an empty cell do
    not; a text result is Unevaluable."""
    if _is_err(v) or v is BLANK:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    return Unevaluable(f"formula {formula!r} returns text ({v!r}), not TRUE / FALSE")


_NUM_LIT_RX = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$")


def literal(f: Optional[str]):
    """A CF operand that is a plain literal (number, "text", TRUE / FALSE) -> its value; else UNKNOWN."""
    if f is None:
        return UNKNOWN
    t = f.strip()
    if t.startswith("="):
        t = t[1:].strip()
    if _NUM_LIT_RX.match(t):
        return float(t)
    if len(t) >= 2 and t[0] == '"' and t[-1] == '"' and '"' not in t[1:-1].replace('""', ""):
        return t[1:-1].replace('""', '"')
    if t.upper() in ("TRUE", "FALSE"):
        return t.upper() == "TRUE"
    return UNKNOWN


def cellis_operands(rule, env=None, where=None, value=UNKNOWN_VALUE):
    """The values of a cellIs rule's operand formulas (literals directly, anything else through the
    evaluator, relative to the cell): a list, UNKNOWN or Unevaluable."""
    if not rule.formulas:
        return Unevaluable("cellIs rule without an operand")
    out = []
    for f in rule.formulas:
        lit = literal(f)
        if lit is UNKNOWN:
            lit = formula_value(f, env, where, value)
            if lit is UNKNOWN or is_unevaluable(lit):
                return lit
        out.append(lit)
    return out


_CELLIS_OPS = {"equal": "=", "notEqual": "<>", "greaterThan": ">", "lessThan": "<",
               "greaterThanOrEqual": ">=", "lessThanOrEqual": "<="}


def cellis_fires(rule, value, env=None, where=None):
    """A cellIs rule on a cell holding the concrete `value` (not BLANK): True / False / UNKNOWN /
    Unevaluable.  An error in the cell or an operand: the comparison is an error, so no format.
    between / notBetween take the two operands as min and max, in either order."""
    ops = cellis_operands(rule, env, where, value)
    if ops is UNKNOWN or is_unevaluable(ops):
        return ops
    if _is_err(value) or any(_is_err(o) for o in ops):
        return False
    op = rule.operator or "between"
    if op in ("between", "notBetween"):
        if len(ops) < 2:
            return Unevaluable("between needs two operands")
        a, b = ops[0], ops[1]
        lo, hi = (a, b) if compare(a, b) <= 0 else (b, a)
        inside = compare(value, lo) >= 0 and compare(value, hi) <= 0
        return inside if op == "between" else not inside
    sym = _CELLIS_OPS.get(op)
    if sym is None:
        return Unevaluable(f"cellIs operator {op!r}")
    return _CMP_FN[sym](compare(value, ops[0]))


# ---------------------------------------------------------------------------- prefetch
def _formula_cells(node, out: list):
    """(sheet or None, RefPart) of every single-cell reference in the tree, and ('name', sheet, name)."""
    k = node[0]
    if k == "ref":
        out.append((node[1], node[2]))
    elif k == "name":
        out.append(("name", node[1], node[2]))
    elif k in ("neg", "pct"):
        _formula_cells(node[1], out)
    elif k in ("bin", "cmp"):
        _formula_cells(node[2], out)
        _formula_cells(node[3], out)
    elif k == "cat":
        _formula_cells(node[1], out)
        _formula_cells(node[2], out)
    elif k == "call":
        for a in node[2]:
            _formula_cells(a, out)


def _span(v: int, absolute: bool, lo: int, hi: int, anchor: int, size: int) -> list:
    """[(a, b)] row / column spans a reference reaches over the cells lo..hi of a range."""
    if absolute:
        return [(v, v)]
    a, b = v + (lo - anchor), v + (hi - anchor)
    if b - a + 1 >= size:
        return [(1, size)]
    a, b = (a - 1) % size + 1, (b - 1) % size + 1
    return [(a, b)] if a <= b else [(a, size), (1, b)]


def prefetch(env: CfEnv, conditional_formats) -> None:
    """Register, with the workbook's cell values (core.lookup), every cell the formulas of these rules
    (expression rules and cellIs operands) can read for any cell of their ranges, so each referenced
    sheet is streamed once.  Defined names resolving to one absolute cell are followed."""
    if env is None or env.values is None:
        return
    for cf in conditional_formats or ():
        if not cf.ranges:
            continue
        r0, c0 = cf.ranges[0][0], cf.ranges[0][1]
        for rule in cf.rules:
            if rule.type not in ("expression", "cellIs"):
                continue
            for f in rule.formulas or ():
                if rule.type == "cellIs" and literal(f) is not UNKNOWN:
                    continue
                refs: list = []
                _formula_cells(parse(f), refs)
                for x in refs:
                    if x[0] == "name":
                        _prefetch_name(env, x[1], x[2], 0)
                        continue
                    sheet, part = x
                    target = sheet or env.sheet
                    for (r1, c1, r2, c2) in cf.ranges:
                        for ra, rb in _span(part.row, part.row_abs, r1, r2, r0, MAX_ROW):
                            for ca, cb in _span(part.col, part.col_abs, c1, c2, c0, MAX_COL):
                                env.values.want(target, (ra, ca, rb, cb))


def _prefetch_name(env: CfEnv, sheet, name, depth):
    if depth >= MAX_NAME_DEPTH:
        return
    text = env.name_text(name, sheet or env.sheet)
    if not text:
        return
    node = parse(text)
    if node[0] == "ref" and node[1] is not None and node[2].row_abs and node[2].col_abs:
        env.values.want(node[1], (node[2].row, node[2].col, node[2].row, node[2].col))
    elif node[0] == "name":
        _prefetch_name(env, node[1], node[2], depth + 1)


# ---------------------------------------------------------------------------- built-in presets (94)
# Excel's preset highlight styles (Home > Conditional Formatting > Highlight Cells Rules / Top-Bottom
# Rules): (font colour, fill colour) as Excel writes them in the dxf (font <color>, fill <bgColor>).
# Patrick 2026-10-05: "assume the built-in Excel conditional formats always have visible text".
PRESET_STYLES = {
    ("9C0006", "FFC7CE"): "Light Red Fill with Dark Red Text",
    ("9C6500", "FFEB9C"): "Yellow Fill with Dark Yellow Text",
    ("9C5700", "FFEB9C"): "Yellow Fill with Dark Yellow Text",      # the same preset in newer Excel builds
    ("006100", "C6EFCE"): "Green Fill with Dark Green Text",
    (None, "FFC7CE"): "Light Red Fill",
    ("9C0006", None): "Red Text",
}


def preset_style(font_rgb, fill_rgb, has_number_format: bool) -> Optional[str]:
    """The name of Excel's built-in preset style a dxf's colours match (font / painted fill as
    'RRGGBB' / 'FFRRGGBB' or None), or None.  A dxf with a number format is never a preset.  ('Red
    Border' changes no colour, so it never conceals anything anyway.)"""
    if has_number_format:
        return None
    key = (font_rgb[-6:].upper() if isinstance(font_rgb, str) else None,
           fill_rgb[-6:].upper() if isinstance(fill_rgb, str) else None)
    if key == (None, None):
        return None
    return PRESET_STYLES.get(key)
