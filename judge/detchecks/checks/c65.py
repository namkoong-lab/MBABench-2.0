"""65 Formatting/Negatives in parentheses.

Rule (rubric_9 #65: "Negative values use parentheses, e.g., (123)" / bad: "Negative values use a
minus sign, e.g., -123"; Patrick 2026-10-03: graded in code, whole workbook).  Full write-up in
docs/checks/65.md.

Every cell that DISPLAYS a negative number must display it in parentheses.  A cell displays a
negative number when it holds a numeric value <= NEG_THRESHOLD (-1e-6; constants and formula
results including shared-formula children and array / spill members; not booleans, errors, text,
ISO dates) whose rendering under its number format shows at least one non-zero digit (a value
that rounds to zero on display, '-0.00', is not judged).  A constant is numeric when its stored
type is numeric (t absent or "n").  A formula result is decided from its TRUSTED value (the
file's own cache when trusted, else the recalculation copy) whatever type the delivered file
stores for it: an agent tool's t="str" empty cache or a t="e" #NAME? placeholder says nothing
about what Excel shows, so the copy's value decides; a genuine text / boolean / error result is
not a number and is skipped.  Every worksheet counts: hidden sheets, rows and columns included;
no inherited-cell or Q&A-sheet exemption.
What the cell displays is worked out as Excel does (detchecks.core.numfmt):
  * the section its number format selects for the value: 1 section -> automatic minus (FAIL);
    2-3 sections -> the negative section; conditional sections per Excel (no automatic minus in a
    conditional or fallback section, measured 2026-10-03); built-in ids from the en-US table
    (5-8 and 37-44 parenthesise, 1-4 / 9-13 / 48 / General add a minus);
  * a conditional format that fires on the cell and carries a number format replaces it (main
    and x14 rules, priority order; rules evaluated with the shared evaluator, checks/_cfeval.py,
    including expressions that read other cells and sheets).  Patrick 2026-10-05, "CONDITIONAL
    FORMATS never stop a grading": a rule that cannot be evaluated is OFF and a conditional-format
    number format whose display is not verified gives way to the cell's own format (recorded in
    stats.cf_assumptions); a data bar / icon set with showValue="0" hides the number.
Outcome per displayed negative (constants below):
  parens    '(' before and ')' after the digits, no dash outside them: '(1,006)', '($5.00)',
            ' (5)' accounting, '(5.00%)', '(3.81E-04)'; escaped '\\(' and quoted '"("' too;
            [Red] neither helps nor hurts                                             -> pass
  minus     a dash before the digits: '-5', '-1.34E-06', '-5.00%', UK accounting '-5 '   -> FAIL
  unsigned  no sign and no parentheses: '#,##0;[Red]#,##0' -> '1,006', '[>0]0;0' -> '5' -> FAIL
  mixed     parentheses AND a dash: '-(5)', '(5.00)-'                                 -> FAIL
  text      a literal negative section with no digit ('"neg"', '\\-', '"T-1"')          -> pass (NO_DIGIT_NEGATIVE_PASSES)
Literal text of the format that contains a letter or a digit ('"FY24-25 "', '" – est."',
'"T-1: "', '\\k') is a label, not part of the number: its letters, digits and dashes are not
the number's digits or sign, so '"T-1: "0;"T-1: "(0)' -> 'T-1: (5)' and '0.0;(0.0)" – est."' ->
'(5.0) – est.' pass; its parentheses still count, so '0;"Net ("0")"' -> 'Net (5)' passes and
'0;"(est) "0' -> '(est) 5' is unsigned.  A literal made only of punctuation ('"-"', '" -"',
'"("') stays part of the number: '0;(0)"-"' -> '(5)-' is mixed.  (The number is classified from
a rendering with the labels' letters, digits and dashes removed; parentheses must enclose it.)
  blank     an empty negative section ('#,##0;;')                                      -> pass (BLANK_NEGATIVE_PASSES; check 94)
  hidden    data bar / icon set with showValue="0"                                     -> pass (HIDDEN_NEGATIVE_PASSES)
  datetime  a date / time / elapsed format (Excel shows '#####')                        -> not judged (DATE_TIME_IN_SCOPE)
  zero      rounds to zero on display ('-0', '-0.00', '(0.00)')                         -> not judged (ROUNDS_TO_ZERO_COUNTS)
Non-anchor cells of a merged range are not displayed and not judged (SKIP_MERGED_NON_ANCHOR).

Values are read only where the decision needs them: a formula result is read only when its
number format's display of a negative depends on the value or would fail (a one-section format,
a conditional format, a minus / unsigned negative section) or a conditional format on the cell
could change that (second pass).  Formats whose every negative passes (parentheses, date, text,
blank) need no value.  An untrusted value is not refused at once: the cell is buffered until the
sheet tail (merges, conditional formats) is known and GradingError (no fallback) is raised only if
a negative there could still fail (MAX_DEFERRED cells per sheet; the rest in a second pass).
Members of a multi-cell array range not written in the file (some writers store only the anchor)
display under the row / column style; when that style could show a negative as a non-pass their
values are needed and unavailable -> GradingError.
"""
from __future__ import annotations

import itertools
import re
import unicodedata
from bisect import bisect_left, bisect_right
from collections import defaultdict
from typing import Optional

from ..core import numfmt as N
from ..core.refs import group_cells, location, make_ref, parse_range, range_to_str
from ..core.sheet import ExcelError
from ..errors import GradingError
from ._cfeval import (UNKNOWN, UNKNOWN_VALUE, Unevaluable, cellis_fires, cellis_operands, cf_env, expression_fires,
                      is_unevaluable, prefetch)
from ._cfeval import literal as _literal
from ._cftext import text_rule_fires
from .base import Check

# ---------------------------------------------------------------------------- policy switches
NEG_THRESHOLD = -1e-6           # values <= this are negatives; (-1e-6, 0) are floating leftovers, not judged
ROUNDS_TO_ZERO_COUNTS = False   # True: a negative that displays as zero ('-0.00' for -0.004 under 0.00) is judged too
NO_DIGIT_NEGATIVE_PASSES = True  # a literal negative section without digits ('"neg"', '\-') shows no number
BLANK_NEGATIVE_PASSES = True    # an empty negative section shows nothing (check 94 charges format blanks)
HIDDEN_NEGATIVE_PASSES = True   # hidden by a data bar / icon set with showValue="0"
DATE_TIME_IN_SCOPE = False      # negatives under date / time formats ('#####') are not numbers on display
SKIP_MERGED_NON_ANCHOR = True   # values hidden under a merged range (not its top-left cell) are not displayed
CF_COUNTS = True                # conditional-format number formats count where they fire
MAX_UNKNOWN_RULES = 8           # more position- / value-dependent CF rules than this in a screen: the value is read
MAX_ASSUMPTION_EXAMPLES = 10
MAX_DEFERRED = 50_000           # untrusted formula cells buffered per sheet until the tail is known (rest: 2nd pass)
MAX_EXAMPLES = 8
REP_NEG = -1234567.891          # representative negative for value-independent format classification

# ---------------------------------------------------------------------------- display classes
PARENS, MINUS, UNSIGNED, MIXED, TEXT, BLANK, HIDDEN, DATETIME, ZERO, NONNEG, UNCERTAIN = (
    "parens", "minus", "unsigned", "mixed", "text", "blank", "hidden", "datetime", "zero", "nonneg", "uncertain")
DASH_CHARS = "-–—−"     # hyphen, en dash, em dash, minus sign (as core numfmt)
_TRAIL_SKIP = ")]% "


def fails(cls: str) -> bool:
    """Does a displayed negative of this class break the rule?"""
    if cls in (MINUS, UNSIGNED, MIXED):
        return True
    if cls == TEXT:
        return not NO_DIGIT_NEGATIVE_PASSES
    if cls == BLANK:
        return not BLANK_NEGATIVE_PASSES
    if cls == HIDDEN:
        return not HIDDEN_NEGATIVE_PASSES
    if cls == DATETIME:
        return DATE_TIME_IN_SCOPE
    if cls == ZERO:
        return ROUNDS_TO_ZERO_COUNTS
    return False                                  # PARENS, NONNEG


def fixed_pass(cls: str) -> bool:
    """A class every negative under a (non-conditional) format shares, and that passes: the
    format never needs the cell's value."""
    return cls in (PARENS, DATETIME, BLANK, TEXT) and not fails(cls)


def _is_label(text: str) -> bool:
    """A literal that contains a letter or a digit is a label, not part of the number."""
    return any(ch.isalnum() for ch in text)


def _label_residue(text: str) -> str:
    """What is left of a label literal once its letters, digits and dashes (none of them the
    number's sign or digits) are removed: its parentheses, spaces and other punctuation, which
    may still wrap the number ('"Net ("0")"' -> 'Net (5)', '"("0" bn)"' -> '(5 bn)')."""
    return "".join(ch for ch in text if not ch.isalnum() and ch not in DASH_CHARS)


def strip_label_literals(section: str) -> str:
    """The section code with the letters, digits and dashes of its label literals removed: a
    quoted / escaped literal that contains a letter or a digit ('"FY24-25 "', '" – est."',
    '"(USD "', '\\k') keeps only its other characters ('" "', '"  ."', '"( "', nothing).
    Punctuation-only literals ('"-"', '"("', '\\-', '\\$'), bracket tags, '_x' padding and
    '*x' fill are kept whole.  Same lexical rules as numfmt.split_sections."""
    out = []
    i, n = 0, len(section)
    while i < n:
        ch = section[i]
        if ch == '"':
            j = section.find('"', i + 1)
            if j < 0:
                out.append(section[i:])
                break
            lit = section[i + 1:j]
            if not _is_label(lit):
                out.append(section[i:j + 1])
            else:
                rest = _label_residue(lit)
                if rest:
                    out.append(f'"{rest}"')
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            if not _is_label(section[i + 1]):
                out.append(section[i:i + 2])
            i += 2
            continue
        if ch in "_*" and i + 1 < n:
            out.append(section[i:i + 2])
            i += 2
            continue
        if ch == "[":
            j = section.find("]", i + 1)
            if j < 0:
                out.append(section[i:])
                break
            out.append(section[i:j + 1])
            i = j + 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _wrapped(num: str, first: int, last: int) -> bool:
    """Do parentheses enclose the number's digits num[first..last]: a '(' still open before
    the first digit and a ')' after the last digit that closes it?  '(5)', ' $(5.00)',
    'Net (5)' yes; '() 5 ()' (two labels' parentheses) and '(5' no."""
    depth = 0
    for ch in num[:first]:
        if ch == "(":
            depth += 1
        elif ch == ")" and depth:
            depth -= 1
    if not depth:
        return False
    depth = 0
    for ch in num[last + 1:]:
        if ch == "(":
            depth += 1
        elif ch == ")":
            if not depth:
                return True
            depth -= 1
    return False


def _number_rendering(r, value: float, code: str, date1904: bool):
    """The rendering of the number alone: the section Excel used, with the letters, digits and
    dashes of its label literals removed (so 'FY24-25 ' or ' – est.' do not pass for a sign or
    for the number's digits).  r itself when the section has no label literal."""
    if r.section_index is None:
        return r
    secs = N.split_sections(code)
    if r.section_index >= len(secs):
        return r
    stripped = strip_label_literals(secs[r.section_index])
    if stripped == secs[r.section_index]:
        return r
    secs[r.section_index] = stripped
    return N.render(value, ";".join(secs), date1904=date1904)


def classify(value: float, code: str, date1904: bool = False, best: bool = False) -> tuple[str, str]:
    """(display class, rendered text) of a numeric value under a format code.  best=True: the engine's best
    rendering even where Excel's reading is not verified (certain=False; Patrick 2026-10-05: every attempt
    graded) - '#####' then shows no number (TEXT)."""
    if value > NEG_THRESHOLD:
        return NONNEG, ""
    r = N.render(value, code, date1904=date1904)
    if not r.certain and not best:
        return UNCERTAIN, r.text
    f = N.parse_format(code)
    if r.kind in ("date", "elapsed") or (r.section_index is not None
                                         and f.sections[r.section_index].kind in ("date", "elapsed")):
        return DATETIME, r.text
    if r.is_hash:
        return (TEXT if best else UNCERTAIN), r.text
    if r.is_blank:
        return BLANK, r.text
    if not any(ch.isdigit() for ch in r.text):
        return TEXT, r.text
    rn = _number_rendering(r, value, code, date1904)       # label literals removed
    num = rn.text.strip()
    digits = [i for i, ch in enumerate(num) if ch.isdigit()]
    if not digits:
        return TEXT, r.text                               # every digit shown belongs to a label ('T-1')
    if all(num[i] == "0" for i in digits):
        return ZERO, r.text
    first, last = digits[0], digits[-1]
    lead = any(ch in DASH_CHARS for ch in num[:first])
    trail = num[last + 1:]
    k = 0
    while k < len(trail) and (trail[k] in _TRAIL_SKIP or unicodedata.category(trail[k]) == "Sc"):
        k += 1
    trailing = k < len(trail) and trail[k] in DASH_CHARS
    dash = lead or trailing
    if _wrapped(num, first, last):
        return (MIXED if dash else PARENS), r.text
    return (MINUS if dash else UNSIGNED), r.text


# ---------------------------------------------------------------------------- CF evaluation
NEG = object()                                    # "some negative value", unknown which
_REF = r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})"
_SELF_CMP_RX = re.compile(r"^\s*" + _REF + r"\s*(<>|<=|>=|=|<|>)\s*(.+?)\s*$")
_CMP_SELF_RX = re.compile(r"^\s*(.+?)\s*(<>|<=|>=|=|<|>)\s*" + _REF + r"\s*$")
_FLIP = {"<": ">", ">": "<", "<=": ">=", ">=": "<=", "=": "=", "<>": "<>"}
_OPS = {"=": lambda k: k == 0, "<>": lambda k: k != 0, "<": lambda k: k < 0, ">": lambda k: k > 0,
        "<=": lambda k: k <= 0, ">=": lambda k: k >= 0}
_CELLIS = {"equal": "=", "notEqual": "<>", "greaterThan": ">", "lessThan": "<",
           "greaterThanOrEqual": ">=", "lessThanOrEqual": "<="}
literal = _literal          # a CF operand / expression as a literal (float, str, bool) or UNKNOWN


def _rank(v) -> int:
    return 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0


def _cmp(a, b) -> int:
    """Excel comparison: numbers < text < logicals; text case-insensitive."""
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.casefold(), b.casefold()
    return (a > b) - (a < b)


def _cmps(value, lit) -> frozenset:
    """Possible comparison results of the cell value with a literal.  value is a float, or NEG
    (an unknown negative number: below every text / logical and every literal >= 0)."""
    if value is NEG:
        if _rank(lit) != 0 or lit >= 0:
            return frozenset({-1})
        return frozenset({-1, 0, 1})
    return frozenset({_cmp(value, lit)})


def _apply(op: str, ks: frozenset):
    res = {_OPS[op](k) for k in ks}
    return res.pop() if len(res) == 1 else UNKNOWN


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def fires(rule, value, where=None, env=None):
    """True / False when the CF rule fires / does not fire on a cell holding `value` (a float,
    or NEG for an unknown negative); UNKNOWN when that depends on what the caller does not know
    (the negative's size, a position-dependent rule without `where`); an Unevaluable when the
    rule cannot be evaluated (checks/_cfeval.py; the caller then takes it as off - Patrick
    2026-10-05).  where = (row, col, anchor_row, anchor_col): relative references in a CF
    formula are relative to the top-left cell of the rule's first range; env = the sheet's
    _cfeval.CfEnv (other cells' values; None: none can be read)."""
    t = rule.type
    if t in ("dataBar", "iconSet", "colorScale"):
        return True                               # applies to every number in its range
    if t == "cellIs":
        if value is not NEG:
            return cellis_fires(rule, value, env, where)
        ops = cellis_operands(rule, env, where, UNKNOWN_VALUE)
        if ops is UNKNOWN or is_unevaluable(ops):
            return ops
        if any(isinstance(o, ExcelError) for o in ops):      # the comparison is an error: no format
            return False
        ops = [0.0 if o is None else o for o in ops]                         # an empty cell compares as 0
        op = rule.operator or "between"
        if op in ("between", "notBetween"):
            if len(ops) < 2:
                return Unevaluable("between needs two operands")
            lo, hi = sorted(ops[:2], key=lambda x: (_rank(x), x.casefold() if isinstance(x, str) else x))
            inside = {a >= 0 and b <= 0 for a in _cmps(value, lo) for b in _cmps(value, hi)}
            if len(inside) != 1:
                return UNKNOWN
            r = inside.pop()
            return r if op == "between" else not r
        sym = _CELLIS.get(op)
        return Unevaluable(f"cellIs operator {op!r}") if sym is None else _apply(sym, _cmps(value, ops[0]))
    if t == "containsBlanks" or t == "containsErrors":
        return False                              # a number is neither blank nor an error
    if t == "notContainsBlanks" or t == "notContainsErrors":
        return True
    if t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        if rule.text is None:
            return Unevaluable(f"{t} rule without its text")
        if value is NEG:
            return UNKNOWN
        return text_rule_fires(t, rule.text, N.general_text(float(value)))
    if t == "expression":
        if len(rule.formulas) != 1:
            return Unevaluable(f"expression rule with {len(rule.formulas)} formulas")
        if value is NEG:
            r = _neg_self_compare(rule.formulas[0], where)
            if r is not UNKNOWN:
                return r
            return expression_fires(rule.formulas[0], env, where, UNKNOWN_VALUE)
        return expression_fires(rule.formulas[0], env, where, value)
    return Unevaluable(f"rule type {t!r} is not evaluated")   # top10, aboveAverage, duplicates, timePeriod ...


def _neg_self_compare(formula: Optional[str], where):
    """For an unknown negative: '<self> <op> <literal>' in either order where <self> is the cell
    itself (B2<0 holds for every negative), else UNKNOWN."""
    if formula is None or where is None:
        return UNKNOWN
    f = formula.strip()
    if f.startswith("="):
        f = f[1:].strip()
    r, c, r0, c0 = where
    m = _SELF_CMP_RX.match(f)
    if m:
        dc, col, dr, row, op, rhs = m.groups()
    else:
        m = _CMP_SELF_RX.match(f)
        if not m:
            return UNKNOWN
        rhs, op, dc, col, dr, row = m.groups()
        op = _FLIP[op]
    lit = literal(rhs)
    if lit is UNKNOWN:
        return UNKNOWN
    rr = int(row) if dr else int(row) + (r - r0)
    cc = _col_num(col) if dc else _col_num(col) + (c - c0)
    if (rr, cc) != (r, c):
        return UNKNOWN                            # the rule reads another cell
    return _apply(op, _cmps(NEG, lit))


def rule_text(rule) -> str:
    what = rule.type or "?"
    if rule.formulas:
        what += " " + ", ".join(repr(f) for f in rule.formulas[:2])
    elif rule.text is not None:
        what += f" {rule.text!r}"
    return what


class _Rule:
    """A CF rule that can change how a negative displays: it sets a number format, hides the
    value (data bar / icon set showValue=0), or stops lower-priority rules (stopIfTrue)."""
    __slots__ = ("prio", "order", "rule", "boxes", "anchor", "fmt", "fmt_err", "hide", "stop")


# ---------------------------------------------------------------------------- the check
class _Style:
    __slots__ = ("code", "err", "cls", "text", "value_free")


class C65(Check):
    number = 65
    key = "Formatting/Negatives in parentheses"
    live = False                                  # Patrick 2026-10-04: built and toy-gated, OFF until v3 (the LLM
                                                  # keeps grading it); the verdict is recorded only
    needs_cells = True
    needs_rows = True                             # row default styles of array members missing from the file
    needs_values = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        self.st = wb.styles
        self.date1904 = bool(wb.date1904)
        self._styles: dict = {}
        self._fmt_cache: dict = {}
        self._outcome_cache: dict = {}
        self.n_values_read = 0
        self.n_deferred_total = 0
        self.n_absent_members = 0
        self.counts = defaultdict(int)            # negative cells by outcome
        self.n_failing = 0
        self.cf_rule_sheets: list = []
        self.cf_second_pass: list = []
        self.deferred_second_pass: list = []
        self.hidden_sheet_failures: list = []
        self.examples: list = []
        self.pending: dict = {}
        self._p = None
        self._env = None
        self.cf_assumed = 0                       # negative cells graded under a conditional-format assumption
        self.cf_assumed_decisive = 0              # ... where the assumption decides the cell's verdict
        self.cf_assumed_examples: list = []
        self.cf_unevaluable: dict = {}            # (sheet, rule order) -> description, rules met in judged cells

    # -------------------------------------------------------------- formats
    def _fmt(self, code: str) -> tuple[str, str, bool]:
        """(class of the representative negative, its text, value_free) for a format code.
        value_free: every negative under the code displays in a passing, value-independent way
        (no conditional sections; parentheses / date / text / blank)."""
        x = self._fmt_cache.get(code)
        if x is None:
            cls, text = classify(REP_NEG, code, self.date1904)
            free = fixed_pass(cls) and not N.parse_format(code).has_conditions
            x = (cls, text, free)
            self._fmt_cache[code] = x
        return x

    def _style(self, s: int) -> _Style:
        x = self._styles.get(s)
        if x is None:
            x = _Style()
            try:
                x.code = N.resolve_format(self.st.num_fmt_id(s), self.st.num_fmts)
                x.cls, x.text, x.value_free = self._fmt(x.code)
                x.err = None
            except GradingError as e:
                # Patrick 2026-10-06 (attempt 2777, measured in Excel): an unreadable format is General
                # (Excel drops it on repair); recorded when a negative is judged under it
                x.code, x.err = "General", str(e)
                x.cls, x.text, x.value_free = self._fmt("General")
            self._styles[s] = x
        return x

    # -------------------------------------------------------------- first pass
    def sheet_start(self, head):
        self.cand: list = []                      # (row, col, style, value): negatives that need a decision
        self.deferred: list = []                  # (row, col, style, source): untrusted values, decided at sheet_end
        self.n_deferred = 0
        self._arrays: dict = {}                   # anchor ref -> [r1, c1, r2, c2, members seen, members, anchor top-left]
        self._arr_max_r2 = 0
        self._row_styles = defaultdict(list)      # custom row style -> rows (only rows inside open array ranges)
        self._row = (0, None, False)              # (row, custom style, recorded)

    def row(self, row):
        s = row.style
        rec = False
        if s is not None and self._arrays and row.r <= self._arr_max_r2:
            self._row_styles[s].append(row.r)
            rec = True
        self._row = (row.r, s, rec)

    def _array_bookkeeping(self, cell):
        """Count the members of multi-cell array / data-table ranges that are written in the file;
        ranges with all members present are dropped (only incomplete ones are kept)."""
        a = cell.array
        if a is not None:
            rec = self._arrays.get(a.anchor_ref)
            if rec is not None:
                rec[4] += 1
                if rec[4] >= rec[5]:
                    del self._arrays[a.anchor_ref]
            return
        f = cell.formula
        if f is None or f.kind not in ("array", "dataTable") or not f.ref:
            return
        b = parse_range(f.ref)
        if b is None or (b[0] == b[2] and b[1] == b[3]):
            return
        need = (b[2] - b[0] + 1) * (b[3] - b[1] + 1) - 1
        self._arrays[cell.ref] = [b[0], b[1], b[2], b[3], 0, need, (b[0], b[1]) == (cell.row, cell.col)]
        self._arr_max_r2 = max(self._arr_max_r2, b[2])
        r, s, rec = self._row
        if s is not None and not rec and r == cell.row:
            self._row_styles[s].append(r)
            self._row = (r, s, True)

    def _numeric(self, cell, read: bool):
        """The cell's number (float), or None when it holds no number.  read=False: only when
        known without a formula-value read that could raise (stats)."""
        if cell.is_formula_result:
            if read:
                if not cell.value_trusted:
                    # Patrick 2026-10-05 (every attempt graded): an untrusted value is skipped for this check
                    self.note_default("untrusted_value", f"{location(cell.sheet, cell.ref)} (source={cell.value_source})")
                    return None
                self.n_values_read += 1
                v = self.require_value(cell)
            elif cell.value_trusted:
                v = cell.value
            else:
                return None
        else:
            v = cell.value
        if isinstance(v, bool) or not isinstance(v, float):
            return None
        return v

    def cell(self, cell):
        if cell.is_blank:
            return
        if cell.is_formula_result:
            # the delivered type says nothing reliable about a formula result (an agent tool
            # writes t="str" with an empty cache, LibreOffice a #NAME? placeholder): the
            # trusted value decides below (_numeric skips genuine text / boolean / error results)
            self._array_bookkeeping(cell)
        elif cell.t != "n":
            return                                # constants: text, booleans, errors, ISO dates
        st = self._styles.get(cell.s) or self._style(cell.s)
        if st.value_free:
            # every negative under this format passes whatever the value (a CF number format that
            # could change that is handled by the second pass); known negatives are counted for the stats
            v = self._numeric(cell, read=False)
            if v is not None and v <= NEG_THRESHOLD:
                self.counts[f"base_{st.cls}"] += 1
            return
        if cell.is_formula_result and not cell.value_trusted:
            # the value is needed unless the tail (a merge, a conditional format) makes every
            # negative here pass: decide at sheet_end; skipped (recorded) there if it is still needed
            self.n_deferred += 1
            if len(self.deferred) < MAX_DEFERRED:
                self.deferred.append((cell.row, cell.col, cell.s, cell.value_source))
            return
        v = self._numeric(cell, read=True)
        if v is None or v > NEG_THRESHOLD:
            return
        if st.err is not None:
            self.note_default("unreadable_number_format", f"{cell.sheet}!{cell.ref} = {v!r}: {st.err}")
        self.cand.append((cell.row, cell.col, cell.s, v))

    # -------------------------------------------------------------- conditional formats
    def _cf_rules(self, tail) -> list:
        out = []
        order = 0
        for cf in tail.conditional_formats:
            if not cf.ranges:
                continue
            anchor = (cf.ranges[0][0], cf.ranges[0][1])
            for rule in cf.rules:
                order += 1
                x = _Rule()
                x.prio = rule.priority if rule.priority is not None else 10 ** 9
                x.order, x.rule, x.boxes, x.anchor = order, rule, list(cf.ranges), anchor
                x.fmt = x.fmt_err = None
                x.hide = rule.type in ("dataBar", "iconSet") and rule.show_value is False
                x.stop = bool(rule.stop_if_true)
                if rule.type not in ("dataBar", "iconSet", "colorScale"):
                    d = rule.dxf if rule.dxf is not None else self.st.dxf(rule.dxf_id)
                    if d is not None:
                        if d.num_fmt_code is not None:
                            x.fmt = d.num_fmt_code
                        elif d.num_fmt_id is not None:
                            try:
                                x.fmt = N.resolve_format(d.num_fmt_id, self.st.num_fmts)
                            except GradingError as e:
                                x.fmt, x.fmt_err = "", str(e)
                if x.fmt is None and not x.hide and not x.stop:
                    continue
                out.append(x)
        out.sort(key=lambda x: (x.prio, x.order))
        return out

    @staticmethod
    def _cover(rules, r: int, c: int, index: Optional[dict] = None) -> tuple:
        """Indexes (priority order) of the rules whose ranges contain (r, c).  index: a per-sheet
        {col: [(r1, r2, i)]} cache filled lazily."""
        if index is None:
            return tuple(i for i, x in enumerate(rules)
                         if any(b[0] <= r <= b[2] and b[1] <= c <= b[3] for b in x.boxes))
        col = index.get(c)
        if col is None:
            col = [(b[0], b[2], i) for i, x in enumerate(rules) for b in x.boxes if b[1] <= c <= b[3]]
            index[c] = col
        return tuple(sorted({i for r1, r2, i in col if r1 <= r <= r2}))

    @staticmethod
    def _cover_box(rules, r1: int, c1: int, r2: int, c2: int) -> tuple:
        """Indexes of the rules whose ranges intersect a rectangle."""
        return tuple(i for i, x in enumerate(rules)
                     if any(b[0] <= r2 and r1 <= b[2] and b[1] <= c2 and c1 <= b[3] for b in x.boxes))

    def _effects(self, rules, ids, fired) -> Optional[list]:
        """All (format code or None, hidden) outcomes of the covering rules (priority order):
        the first firing rule with a number format sets it; a firing hidden-value data bar
        hides the number; stopIfTrue ends the evaluation.  Unknown rules branch."""
        unknown = [k for k, f in enumerate(fired) if f is UNKNOWN]
        if len(unknown) > MAX_UNKNOWN_RULES:
            return None
        out = set()
        for combo in itertools.product((True, False), repeat=len(unknown)):
            fl = list(fired)
            for k, i in enumerate(unknown):
                fl[i] = combo[k]
            fmt, hide = None, False
            for i, f in zip(ids, fl):
                if not f:
                    continue
                x = rules[i]
                if x.fmt is not None and fmt is None:
                    if x.fmt_err is not None:
                        raise GradingError(f"{self.key}: a conditional format's number format cannot be read: "
                                           f"{x.fmt_err}")
                    fmt = x.fmt
                hide = hide or x.hide
                if x.stop:
                    break
            out.add((fmt, hide))
        return sorted(out, key=lambda e: (e[0] is None, e[0] or "", e[1]))

    def _fire(self, x, v, where):
        """fires() for a _Rule: an Unevaluable when its own number format cannot be read either."""
        if x.fmt_err is not None:
            return Unevaluable(f"its number format cannot be read ({x.fmt_err})")
        return fires(x.rule, v, where, self._env)

    def _outcomes(self, v: float, st: _Style, rules, ids, pos: tuple, branch: bool = False) -> list:
        """[(class, text, format code, via_cf)] a negative cell can show under its CF rules.  A rule that
        cannot be evaluated is OFF (Patrick 2026-10-05); branch=True: it branches like an unknown rule
        (only to tell whether that assumption decides the verdict)."""
        fired = tuple(self._fire(rules[i], v, pos + rules[i].anchor) for i in ids)
        fired = tuple(((UNKNOWN if branch else False) if is_unevaluable(f) else f) for f in fired)
        key = (v, st.code, ids, fired)
        res = self._outcome_cache.get(key)
        if res is not None:
            return res
        effects = self._effects(rules, ids, fired) if ids else [(None, False)]
        if effects is None:
            raise GradingError(f"{self.key}: a cell is covered by more than {MAX_UNKNOWN_RULES} conditional-format "
                               f"rules whose outcome depends on its position or value")
        res = []
        for fmt, hide in effects:
            code = fmt if fmt is not None else st.code
            if hide:
                cls, text = HIDDEN, ""
            else:
                cls, text = classify(v, code, self.date1904)
            res.append((cls, text, code, fmt is not None))
        self._outcome_cache[key] = res
        return res

    def _settled(self, res: list, v: float, st: _Style, pos) -> list:
        """An outcome under a CONDITIONAL-format number format whose display is not verified gives way to
        the cell's own format (Patrick 2026-10-05).  The same list when there is none."""
        if not any(x[0] == UNCERTAIN and x[3] for x in res):
            return res
        out = [x for x in res if not (x[0] == UNCERTAIN and x[3])]
        return out + [x for x in self._outcomes(v, st, [], (), pos) if x not in out]

    def _neg_could_fail(self, st: _Style, rules, ids, pos: Optional[tuple], branch: bool = False) -> bool:
        """Could SOME negative under this base format (and these CF rules) break the rule or be
        undecidable?  If not, the cell's value is not needed.  pos = (row, col), or None for a
        whole range (rules that read the cell's position become unknown).  A rule that cannot be
        evaluated is off (branch=True: it may fire - is that assumption ever decisive here?); a CF
        number format whose display is not verified gives way to the cell's own (Patrick 2026-10-05)."""
        if st.err is not None:
            return True
        try:
            fired = tuple(self._fire(rules[i], NEG, None if pos is None else pos + rules[i].anchor) for i in ids)
            fired = tuple(((UNKNOWN if branch else False) if is_unevaluable(f) else f) for f in fired)
            effects = self._effects(rules, ids, fired) if ids else [(None, False)]
        except GradingError:
            return True
        if effects is None:
            return True
        for fmt, hide in effects:
            if hide:
                if fails(HIDDEN):
                    return True
                continue
            if fmt is not None and self._fmt(fmt)[0] == UNCERTAIN:
                if branch:
                    return True                   # the assumption could decide: read the value, record it
                fmt = None                        # the cell's own format instead
            if fmt is None:
                if not st.value_free:
                    return True
            elif not self._fmt(fmt)[2]:
                return True
        return False

    def _decide(self, sheet: str, ref: str, v: float, st: _Style, rules, ids, r: int, c: int):
        """(fails?, class, text, code, via_cf) for one negative cell; GradingError when its own format's
        display is not verified (no guessing).  Conditional formats never raise (Patrick 2026-10-05): an
        unevaluable rule is off and a CF display that is not verified gives way to the cell's own format
        - each such assumption recorded in stats.cf_assumptions."""
        res = self._outcomes(v, st, rules, ids, (r, c))
        why = []
        decisive = False
        settled = self._settled(res, v, st, (r, c))
        if settled is not res or len({fails(x[0]) for x in settled}) > 1:
            why.append("its conditional formatting's display is not verified (" +
                       ", ".join(sorted({repr(x[2]) for x in res if x[3]})) + "): graded by its own format")
            decisive = True
            settled = self._outcomes(v, st, [], (), (r, c))
        opened = [(rules[i], f) for i in ids
                  for f in (self._fire(rules[i], v, (r, c) + rules[i].anchor),) if is_unevaluable(f)]
        if opened:
            why.append("rule(s) assumed off: " + "; ".join(f"{rule_text(x.rule)} ({f.why})" for x, f in opened))
            for x, f in opened:
                self._note_unevaluable(sheet, x, f.why)
            if not decisive:
                try:
                    alt = self._settled(self._outcomes(v, st, rules, ids, (r, c), branch=True), v, st, (r, c))
                    decisive = {fails(x[0]) for x in alt} != {fails(x[0]) for x in settled} or \
                        any(x[0] == UNCERTAIN for x in alt)
                except GradingError:
                    decisive = True
        res = settled
        if why:
            self._assume(sheet, ref, v, "; ".join(why), decisive)
        if any(x[0] == UNCERTAIN for x in res):
            # Patrick 2026-10-05 (every attempt graded): the engine's best rendering of the cell's own format
            res = [(classify(v, x[2], self.date1904, best=True) + x[2:]) if x[0] == UNCERTAIN else x for x in res]
            self.note_default("unverified_number_format", f"{location(sheet, ref)} = {v!r} under "
                                                          f"{res[0][2]!r} graded as {res[0][1].strip()!r}")
        bad = [x for x in res if fails(x[0])]
        if bad:
            return (True,) + bad[0]
        return (False,) + res[0]

    def _assume(self, sheet: str, ref: str, v, why: str, decisive: bool):
        """Record a cell graded under a conditional-format assumption (stats.cf_assumptions)."""
        self.cf_assumed += 1
        self.cf_assumed_decisive += int(decisive)
        if len(self.cf_assumed_examples) < MAX_ASSUMPTION_EXAMPLES and (decisive or len(self.cf_assumed_examples) < 3):
            self.cf_assumed_examples.append(f"{location(sheet, ref)} = {v!r}{' (decisive)' if decisive else ''}: "
                                            f"{why}"[:300])

    def _note_unevaluable(self, sheet: str, x, why: str):
        k = (sheet, x.order)
        if k not in self.cf_unevaluable and len(self.cf_unevaluable) < 1000:
            where = ",".join(range_to_str(*b) for b in x.boxes[:3]) + (",..." if len(x.boxes) > 3 else "")
            self.cf_unevaluable[k] = f"{location(sheet, where)}: {rule_text(x.rule)} ({why})"[:300]

    # -------------------------------------------------------------- merges
    @staticmethod
    def _merge_index(tail) -> dict:
        """{col: (row starts, [(r1, r2, c1)] sorted by r1, disjoint?)} for the non-anchor test."""
        if not SKIP_MERGED_NON_ANCHOR:
            return {}
        by_col = defaultdict(list)
        for m in tail.merges:
            if m.is_single_cell:
                continue
            for col in range(m.c1, min(m.c2, m.c1 + 16384) + 1):
                by_col[col].append((m.r1, m.r2, m.c1))
        idx = {}
        for col, items in by_col.items():
            items.sort()
            disjoint = all(items[k][0] > items[k - 1][1] for k in range(1, len(items)))
            idx[col] = ([x[0] for x in items], items, disjoint)
        return idx

    @staticmethod
    def _hidden_by_merge(idx: dict, r: int, c: int) -> bool:
        """Is (r, c) a cell of a merged range other than its top-left one?"""
        e = idx.get(c)
        if e is None:
            return False
        starts, items, disjoint = e
        if disjoint:
            k = bisect_right(starts, r) - 1
            if k < 0:
                return False
            r1, r2, c1 = items[k]
            return r <= r2 and (r, c) != (r1, c1)
        return any(r1 <= r <= r2 and (r, c) != (r1, c1) for r1, r2, c1 in items)

    # -------------------------------------------------------------- untrusted values
    def _check_untrusted(self, sheet, r, c, s, source, rules, merges, index):
        """An untrusted formula value: skipped (Patrick 2026-10-05: every attempt graded), recorded in
        stats.defaults.untrusted_value when a negative there could have failed."""
        if merges and self._hidden_by_merge(merges, r, c):
            return
        ids = self._cover(rules, r, c, index) if rules else ()
        if self._neg_could_fail(self._style(s), rules, ids, (r, c)):
            self.note_default("untrusted_value", f"{location(sheet, make_ref(r, c))} (source={source})")

    def _check_deferred(self, sheet, rules, merges):
        index: dict = {}
        for r, c, s, src in self.deferred:
            self._check_untrusted(sheet, r, c, s, src, rules, merges, index)

    # -------------------------------------------------------------- array members missing from the file
    def _absent_styles(self, head, r1, c1, r2, c2) -> set:
        """Styles Excel gives to cells of the rectangle that have no <c>: the row's style when
        the row has customFormat, else the column's style, else 0."""
        styles = set()
        n_custom = 0
        for s, rows in self._row_styles.items():
            lo, hi = bisect_left(rows, r1), bisect_right(rows, r2)
            if hi > lo:
                styles.add(s)
                n_custom += hi - lo
        if n_custom < r2 - r1 + 1:
            covered = 0
            for ci in head.cols:
                lo, hi = max(ci.min, c1), min(ci.max, c2)
                if lo <= hi:
                    styles.add(ci.style or 0)
                    covered += hi - lo + 1
            if covered < c2 - c1 + 1:
                styles.add(0)
        return styles

    def _check_absent_members(self, head, rules):
        for s in self._row_styles.values():
            s.sort()
        for anchor, (r1, c1, r2, c2, seen, need, top_left) in self._arrays.items():
            absent = need - seen
            if absent <= 0:
                continue
            self.n_absent_members += absent
            rng = range_to_str(r1, c1, r2, c2)
            if not top_left:
                # Patrick 2026-10-05 (every attempt graded): members not written in the file are skipped
                self.note_default("unwritten_array_member", f"{head.name}!{rng} (anchor {anchor}, not its top-left "
                                                            f"cell): {absent} member(s) skipped", absent)
                continue
            ids = self._cover_box(rules, r1, c1, r2, c2) if rules else ()
            # rows / columns that can hold a missing member (the anchor itself is written)
            rr1 = r1 + 1 if c1 == c2 else r1
            cc1 = c1 + 1 if r1 == r2 else c1
            if any(self._neg_could_fail(self._style(s), rules, ids, None) for s in sorted(self._absent_styles(head, rr1, cc1, r2, c2))):
                # their values are not available (no cell in the file): skipped (Patrick 2026-10-05)
                self.note_default("unwritten_array_member", f"{head.name}!{rng} (anchor {anchor}): {absent} "
                                                            f"member(s) not written in the file skipped", absent)

    # -------------------------------------------------------------- sheet end
    def sheet_end(self, head, tail):
        self._outcome_cache = {}                  # keyed by rule indexes: valid for one sheet only
        rules = self._cf_rules(tail) if CF_COUNTS else []
        self._env = cf_env(self.wb, head.name)    # the rules' formulas read other cells through it
        if rules:
            self.cf_rule_sheets.append(head.name)
            prefetch(self._env, tail.conditional_formats)
        merges = self._merge_index(tail) if (self.cand or rules or self.deferred) else {}
        self.n_deferred_total += self.n_deferred
        self._check_deferred(head.name, rules, merges)
        if self._arrays:
            self._check_absent_members(head, rules)
        hits = self._judge_all(head.name, self.cand, rules, merges)
        cf_second = bool(rules) and self._needs_second_pass(rules)
        overflow = self.n_deferred > len(self.deferred)
        if cf_second or overflow:
            boxes = [b for x in rules for b in x.boxes]
            self.pending[head.name] = {"rules": rules, "merges": merges, "cf": cf_second, "overflow": overflow,
                                       "seen_untrusted": 0, "env": self._env,
                                       "done": {(r, c) for r, c, _s, _v in self.cand},
                                       "hits": hits, "index": {},
                                       "bbox": ((min(b[0] for b in boxes), min(b[1] for b in boxes),
                                                 max(b[2] for b in boxes), max(b[3] for b in boxes))
                                                if boxes else (0, 0, -1, -1))}
            if cf_second:
                self.cf_second_pass.append(head.name)
            if overflow:
                self.deferred_second_pass.append(head.name)
            self.request_second_pass(head.name, None)
        else:
            self._commit(head, hits)
        self.cand = []
        self.deferred = []
        self._arrays = {}
        self._row_styles = defaultdict(list)

    def _judge_all(self, sheet, cands, rules, merges) -> dict:
        hits = {}                                  # (code, via_cf, cls) -> [(r, c, text)]
        index: dict = {}
        for r, c, s, v in cands:
            if merges and self._hidden_by_merge(merges, r, c):
                self.counts["merged_non_anchor"] += 1
                continue
            st = self._style(s)
            ids = self._cover(rules, r, c, index) if rules else ()
            self._judge_one(sheet, r, c, v, st, rules, ids, hits)
        return hits

    def _judge_one(self, sheet, r, c, v, st, rules, ids, hits):
        bad, cls, text, code, via_cf = self._decide(sheet, make_ref(r, c), v, st, rules, ids, r, c)
        self.counts[f"judged_{cls}"] += 1
        if bad:
            hits.setdefault((code, via_cf, cls), []).append((r, c, text.strip()))
            if len(self.examples) < MAX_EXAMPLES:
                self.examples.append(f"{location(sheet, make_ref(r, c))} = {v!r} shows {text.strip()!r} "
                                     f"under {code!r}{' (conditional format)' if via_cf else ''}")

    def _needs_second_pass(self, rules) -> bool:
        """Could a CF number format show a negative as a non-pass on a cell whose BASE format
        passes for every negative (such cells were not read in the first pass)?"""
        for x in rules:
            if x.fmt is None:
                continue
            if x.fmt_err is not None or not self._fmt(x.fmt)[2]:
                if self._fire(x, NEG, None) is not False:      # fires, depends on the cell, or cannot be evaluated
                    return True
        return False

    _WHAT = {
        MINUS: "display with a minus sign, e.g. {ex}",
        UNSIGNED: "display without a sign or parentheses, e.g. {ex}",
        MIXED: "display a minus sign as well as parentheses, e.g. {ex}",
        TEXT: "display text instead of the number, e.g. {ex}",
        BLANK: "display nothing",
        HIDDEN: "are hidden by a data bar / icon set",
        ZERO: "round to zero on display, e.g. {ex}",
        DATETIME: "display as a date or time",
    }

    def _commit(self, head, hits: dict):
        hidden = head.state != "visible"
        order = sorted(hits.items(), key=lambda kv: min((r, c) for r, c, _t in kv[1]))
        n_sheet = 0
        for (code, via_cf, cls), cells in order:
            n_sheet += len(cells)
            src = f"conditional-format number format {code!r}" if via_cf else f"number format {code!r}"
            by_pos = {(r, c): t for r, c, t in cells}
            for b in group_cells(by_pos):
                rng = range_to_str(*b)
                n = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
                ex = repr(by_pos[min(k for k in by_pos if b[0] <= k[0] <= b[2] and b[1] <= k[1] <= b[3])])
                what = self._WHAT.get(cls, "display {ex}").replace("{ex}", ex)
                self.mistakes.add(location(head.name, rng),
                                  f"{n} negative value(s) {rng} {what} ({src}), not in parentheses"
                                  f"{' (hidden sheet)' if hidden else ''}.")
        self.n_failing += n_sheet
        if hidden and n_sheet:
            self.hidden_sheet_failures.append(head.name)

    # -------------------------------------------------------------- second pass
    def second_pass_start(self, head):
        p = self.pending[head.name]
        self._p = p
        self._env = p["env"]
        self._outcome_cache = {}

    def second_pass_cell(self, cell):
        if cell.is_blank or (cell.t != "n" and not cell.is_formula_result):
            return                                # same type gate as cell(): constants only
        p = self._p
        r, c = cell.row, cell.col
        st = self._style(cell.s)
        if not st.value_free:
            # decided in the first pass, except untrusted values beyond the first MAX_DEFERRED
            # (same predicate, same stream order as in cell())
            if p["overflow"] and cell.is_formula_result and not cell.value_trusted:
                p["seen_untrusted"] += 1
                if p["seen_untrusted"] > MAX_DEFERRED:
                    self._check_untrusted(cell.sheet, r, c, cell.s, cell.value_source, p["rules"], p["merges"],
                                          p["index"])
            return
        if not p["cf"]:
            return
        bb = p["bbox"]
        if not (bb[0] <= r <= bb[2] and bb[1] <= c <= bb[3]) or (r, c) in p["done"]:
            return
        rules = p["rules"]
        ids = self._cover(rules, r, c, p["index"])
        if not ids or not any(rules[i].fmt is not None for i in ids):
            return
        if p["merges"] and self._hidden_by_merge(p["merges"], r, c):
            return
        if not self._neg_could_fail(st, rules, ids, (r, c)):
            # value-independent screen: no value needed - unless a rule that cannot be evaluated (assumed
            # off) could make a negative here fail: then a readable value is read, so the assumption is recorded
            if (cell.is_formula_result and not cell.value_trusted) or \
                    not self._neg_could_fail(st, rules, ids, (r, c), branch=True):
                return
        v = self._numeric(cell, read=True)
        if v is None or v > NEG_THRESHOLD:
            return
        self._judge_one(cell.sheet, r, c, v, st, rules, ids, p["hits"])

    def second_pass_end(self, head, tail):
        self._commit(head, self._p["hits"])
        self._p = None

    # -------------------------------------------------------------- verdict
    def finish(self) -> dict:
        c = self.counts
        stats = {
            "negative_cells_failing": self.n_failing,
            "negative_cells_by_display": {k: v for k, v in sorted(c.items())},
            "formula_values_read": self.n_values_read,
            "untrusted_values_not_needed": self.n_deferred_total,
            "array_members_not_in_file": self.n_absent_members,
            "cf_rule_sheets": self.cf_rule_sheets,
            "cf_second_pass_sheets": self.cf_second_pass,
            "deferred_second_pass_sheets": self.deferred_second_pass,
            "hidden_sheets_with_failures": self.hidden_sheet_failures,
            "examples": self.examples,
            # Patrick 2026-10-05: conditional formats never stop a grading - every assumption is recorded
            "cf_assumptions": {
                "policy": "a conditional-format rule that cannot be evaluated is OFF, and a conditional format "
                          "whose display is not verified gives way to the cell's own number format "
                          "(Patrick 2026-10-05)",
                "cells": self.cf_assumed, "decisive_cells": self.cf_assumed_decisive,
                "examples": self.cf_assumed_examples,
                "unevaluable_rules": len(self.cf_unevaluable),
                "unevaluable_rule_examples": list(self.cf_unevaluable.values())[:MAX_ASSUMPTION_EXAMPLES]},
            "policy": {"neg_threshold": NEG_THRESHOLD, "rounds_to_zero_counts": ROUNDS_TO_ZERO_COUNTS,
                       "no_digit_negative_passes": NO_DIGIT_NEGATIVE_PASSES,
                       "blank_negative_passes": BLANK_NEGATIVE_PASSES,
                       "hidden_negative_passes": HIDDEN_NEGATIVE_PASSES, "date_time_in_scope": DATE_TIME_IN_SCOPE,
                       "skip_merged_non_anchor": SKIP_MERGED_NON_ANCHOR, "cf_counts": CF_COUNTS},
        }
        return self.verdict("Every displayed negative number is in parentheses.",
                            "{n} block(s) of negative numbers display with a minus sign or without parentheses "
                            f"({self.n_failing} cell(s)).", stats)
