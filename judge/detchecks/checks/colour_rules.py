"""Shared rules for the font-colour checks 49 / 50 / 51 (group "colours").

Patrick's confirmed colour rule (handoff, 2026-10-02).  Every formula cell is put in one class:

    EXTERNAL  reads another WORKBOOK anywhere in the formula (directly, through a defined
              name, or through a literal INDIRECT("'[Book.xlsx]S'!A1"))       -> red    (51)
    POINTER   otherwise, the formula is ONLY a reference to a cell / range on ANOTHER sheet
              of this workbook (=Inputs!B5, also =+Inputs!B5, =(Inputs!B5), =-Inputs!B5,
              =Inputs!B5#, =@Inputs!B5:B9, =ANCHORARRAY(Inputs!B5), =TOROW(Inputs!B5:B9), a
              defined name / table that resolves to another sheet)            -> green  (50)
    CROSS     otherwise, the formula calculates with another sheet's cells
              (=Inputs!B5*C3, =SUM(Inputs!A1:A9), =WACC*B2 with WACC on Inputs) -> black or green
    OWN       otherwise: uses only its own sheet (incl. self-qualified 'Own'!A1), or no
              reference at all (=5, =TODAY(), ="text")                         -> black  (49)

Each class is judged by exactly one check (confirmed rule applied fully, Patrick 2026-10-03):
49 judges OWN (black) and CROSS (black or green); 50 judges POINTER (green); 51 judges
EXTERNAL (red).  A blue pointer therefore fails 50 only, a blue external link 51 only.

The colour that counts is the cell's FONT colour (cellXfs -> fonts -> color, resolved with
the workbook's theme / indexed palette / tint by detchecks.core.styles).  Number-format colour
tags ([Blue]0) do not change it (see the constants and docs/checks/49.md).  A conditional-format
font colour EXCUSES a cell whose base font is wrong when a rule covering it paints a family the
cell's class accepts, fires or not (Patrick 2026-10-06, attempt 2348: a black pointer shown green
when "OK" / red when "FAIL" is compliant; CF_FONT_COLOUR_EXCUSES, ColourCheck._cf_excuse,
stats.cf_excused); it never makes a base-correct cell fail.  Colour families (HSV bands on the
resolved RGB) are defined here once.

HYPERLINK cells (Patrick's ruling 2026-10-03: navigation links may be any colour, "not a
calculation"): a navigation cell is ignored by 49, 50 and 51 (is_hyperlink_formula).  After
optional whitespace / one leading '=', the whole formula is either a single built-in
HYPERLINK(...) call (nothing before its name, nothing after its closing parenthesis), or - since
2026-10-04, second review - one of the guard functions HYPERLINK_GUARDS (IFERROR, IFNA, IF)
around such a call whose every other argument is a single literal or empty:
=IFERROR(HYPERLINK("#'Summary'!A1","OPEN"),"OPEN") (the ChatGPT writer's standard navigation
cell), =IF(TRUE,HYPERLINK(...),""), guards may nest.  The HYPERLINK call's own arguments are
free (they build the link).  Anything that calculates around the link is graded like any other
formula: a guard argument that reads a cell or calculates (=IF(B9>0,HYPERLINK(...),""),
=IFERROR(HYPERLINK(...),C1*2)), =HYPERLINK(...)&"", =(HYPERLINK(...)), =+HYPERLINK(...),
HYPERLINK inside any other function.  Spill members and shared-formula children take their
anchor's / master's text, so they are ignored with it.

Formula cells: every cell with its own formula (normal, shared master or child, array or
dataTable anchor) and every non-anchor member of an array / dynamic-array spill / data-table
range.  Members take their anchor's class; their own font colour is judged.  Empty <f/>
markers outside any array range carry no formula: they are not formula cells (stats only).

Classification is syntactic (formula text + defined names + table parts); cell values are
never read, except that 49 reads the one constant cell an =INDIRECT(A20) address comes from
when (and only when) that decides a green cell.  A check works out only what its decision
needs (NO FALLBACK, but raise only where the decision needs it):
  * a cell in the check's safe colour (49 black, 50 green, 51 red) is never classified;
  * a cheap conservative pre-filter skips formulas that cannot be in a failing class
    (49: green formulas with a direct reference to another sheet; 50: formulas that cannot be
    a pure reference; 51: formulas that cannot read another workbook);
  * a font colour that cannot be resolved raises only when the cell's class makes the colour
    decide the verdict (49: OWN / CROSS cells; 50: pointers; 51: external cells);
  * 49: a non-black, non-green formula that cannot be classified still fails when cheap scans
    show it can be neither a pointer nor a reader of another workbook (else it raises);
  * a defined name whose scope is unknown (bad localSheetId) raises only where the possible
    definitions could change this check's decision.
"""
from __future__ import annotations

import colorsys
import re
import weakref
from typing import Optional

from ..core import formula as F
from ..core.refs import group_cells, in_bounds, index_to_col, location, range_to_str
from ..core.styles import is_unknown
from ..errors import GradingError
from .base import Check

# ----------------------------------------------------------------------------- policy constants
# The colour that counts is the font colour.  The displayed-colour variants are NOT implemented;
# ColourCheck.start() raises if one of these is switched on (see docs/checks/49.md).
NUMFMT_COLOUR_COUNTS = False      # [Blue]/[Red]/[ColorN] number-format tags change the colour
# Patrick 2026-10-06 (spot-check case 20, attempt 2348 Cover!C5: a black pointer whose conditional
# formats paint it green when "OK" and red when "FAIL" is compliant): a conditional-format font
# colour EXCUSES a cell whose base font is wrong when some rule covering the cell paints a colour
# the cell's class accepts - whether or not the rule fires now.  It never makes a base-correct cell
# fail (check 57's topic).  An unresolvable dxf font colour excuses nothing (recorded).
CF_FONT_COLOUR_EXCUSES = True

# One-argument functions that only pass a reference through (no calculation): a formula that is
# nothing but one of these around an other-sheet reference is still a POINTER.  ANCHORARRAY /
# SINGLE are how Excel stores A1# / @A1:B5; TOROW / TOCOL / TRANSPOSE only reshape the range
# (toy 50/T4 Fail: =TOROW('1_Historical_Data'!$C$5:$C$40) is "a pure cross-sheet pull").
POINTER_WRAPPERS = ("ANCHORARRAY", "SINGLE", "TOROW", "TOCOL", "TRANSPOSE")
# Unary minus keeps a formula a pointer (=-Inputs!B5 is a sign-flipped link).  False makes it a
# CROSS calculation (black or green).  Unary plus and parentheses are always transparent.
POINTER_ALLOWS_NEGATION = True
# What-if data-table cells (<f t="dataTable">, TABLE()) evaluate the table's own-sheet formula
# with own-sheet input cells (Excel requires the input cells on the table's sheet): OWN class.
DATA_TABLE_CLASS = "own"
# A bare table name used as a reference (=SUM(Tbl), =Tbl) is the table's data body, like Tbl[]
# (Excel keeps table names and defined names in one namespace).  False = an unknown name.
BARE_TABLE_NAME_IS_REFERENCE = True
# INDIRECT addresses assembled with '&': functions whose result is always a number (or an
# error), so their text can never add a sheet qualifier ('!') to the address: "B"&ROW() is
# an own-sheet address, "Inputs!B"&ROW() one on Inputs.
NUMERIC_FUNCS = frozenset({"ROW", "COLUMN", "ROWS", "COLUMNS", "MATCH", "XMATCH", "INT", "MOD", "ABS",
                           "ROUND", "ROUNDUP", "ROUNDDOWN", "MAX", "MIN", "SUM", "COUNT", "COUNTA",
                           "COUNTIF", "COUNTIFS", "LEN"})

# Colour families on the resolved RGB (HSV: hue in degrees, s and v in 0..1).
BLACK_MAX_V = 0.20            # any hue this dark is black (000000, 1A1A1A, 0D0D26 navy-black)
BLACK_GREY_MAX_V = 0.30       # near-grey (s <= BLACK_GREY_MAX_S) up to v 0.30 is black:
BLACK_GREY_MAX_S = 0.25       #   404040 ("Text 1, lighter 25%") yes; 595959 (lighter 35%), 808080 no
GREEN_HUE = (75.0, 165.0)     # 008000, 00B050, 00FF00, 006100, 375623, 548235, 70AD47 (accent6), 92D050
GREEN_MIN_S = 0.35
GREEN_MIN_V = 0.25
RED_HUE_LOW = 15.0            # red: hue < 15 or >= 340
RED_HUE_HIGH = 340.0          # FF0000, C00000, C0504D, 9C0006, 800000; not FF6600 (orange), FF99CC (pink)
RED_MIN_S = 0.50
RED_MIN_V = 0.35

DEFAULT_FONT_ARGB = "FF000000"   # an unresolvable font colour reads as automatic = black (Patrick 2026-10-05)

OWN, CROSS, POINTER, EXTERNAL = "own", "cross", "pointer", "external"
CLASSES = (OWN, CROSS, POINTER, EXTERNAL)
UNCLASSIFIED = "unclassified"     # 49 only: colour wrong in every class, class could not be worked out


# ----------------------------------------------------------------------------- colours
def hsv(argb: str) -> tuple:
    h6 = argb[-6:]
    r, g, b = (int(h6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return h * 360.0, s, v


def family(argb: str) -> str:
    """'black' | 'green' | 'red' | 'other' for a resolved 'FFRRGGBB' / 'RRGGBB' colour."""
    h, s, v = hsv(argb)
    if v <= BLACK_MAX_V or (v <= BLACK_GREY_MAX_V and s <= BLACK_GREY_MAX_S):
        return "black"
    if GREEN_HUE[0] <= h < GREEN_HUE[1] and s >= GREEN_MIN_S and v >= GREEN_MIN_V:
        return "green"
    if (h < RED_HUE_LOW or h >= RED_HUE_HIGH) and s >= RED_MIN_S and v >= RED_MIN_V:
        return "red"
    return "other"


def colour_name(argb: str) -> str:
    """Plain-language name for messages only (never used to decide)."""
    fam = family(argb)
    if fam != "other":
        return fam
    h, s, v = hsv(argb)
    if s < 0.15:
        return "white" if v >= 0.93 else "grey"
    if 190.0 <= h < 260.0:
        return "blue"
    if 165.0 <= h < 190.0:
        return "teal"
    if 15.0 <= h < 50.0:
        return "orange"
    if 50.0 <= h < 75.0:
        return "yellow"
    if 260.0 <= h < 300.0:
        return "purple"
    if 300.0 <= h < 340.0:
        return "pink"
    if 75.0 <= h < 165.0:
        return "pale green"
    return "pale red"


def describe_colour(argb: str) -> str:
    return f"{colour_name(argb)} ({argb[-6:]})"


# ----------------------------------------------------------------------------- text scans
_SHEET_REF_KINDS = frozenset(F.REFERENCE_KINDS) | {"error"}
_LITERAL_KINDS = frozenset(F.LITERAL_KINDS)
# A quoted sheet name, a [...] group (external book index or structured-reference spec, one
# nesting level, ' escapes: the same pattern as core.formula's lexer) or a string literal,
# matched left to right as the lexer does.
_BRACKET = r"\[(?:[^\[\]']|'.|\[(?:[^\[\]']|'.)*\])*\]"
_RE_SCAN = re.compile(r"(?P<q>'(?:[^']|'')*')|(?P<b>" + _BRACKET + r')|(?P<s>"(?:[^"]|"")*")', re.S)
_NOT_IN_POINTER = frozenset('*/^&=<>%,;{}"')
# Shape of a pure reference on the masked text (strings blanked, quoted sheet names -> 'Q',
# [...] groups -> 'B'): a head of '(' / unary + - @ / one-argument wrapper calls (any _xl*.
# prefix), ONE operand holding no whitespace and none of ( ) + - @, then a tail of ')' / spill
# '#' / stray + - @ (the parser tolerates a trailing operator, so the tail keeps the test
# conservative).  Anything else - a non-wrapper call (SUM(Inputs!A1:A5), SUM(Nope[Col])), two
# operands (Inputs!B5+Nope[Col], the intersection Inputs!A1:A5 Inputs!B1:B5) - is never a pointer.
_RE_POINTER_HEAD = re.compile(r"(?:\s|[(+\-@]|(?:_xl[a-z]+\.)*(?:" + "|".join(POINTER_WRAPPERS) + r")\()+", re.I)
_RE_POINTER_TAIL = re.compile(r"(?:\s|[)#+\-@])+$")
_RE_POINTER_CORE = re.compile(r"^[^\s()+\-@]+$")
_IDENT_RX = re.compile("[A-Za-z_\\\\\u00C0-\uFFFF][\\w.?\u00C0-\uFFFF]*")   # defined-name spelling (Excel allows ? and .)
# Qualifiers that formula.parse_qualifier may read as another workbook without a '[': a path
# ('C:\d\B.xlsx'!N, 'https://h/B.xlsx'!N; Excel quotes them) or a name with a workbook extension
# right before '!' (formula._RE_BOOK_EXT: Other.xlsx!Rate).  Used only as a conservative pre-filter.
_BOOK_EXT_BANG = re.compile(r"\.(?:xl[a-z]{1,2}|csv|ods)'?!", re.I)
_QUOTED_QUAL = re.compile(r"'((?:[^']|'')*)'!")
# A sheet qualifier (quoted, or unquoted incl. a 3-D span) directly followed by an A1 cell
# reference, at the start of an operand (start of text, or after an operator / separator /
# space).  Used on text whose string literals and [...] groups are blanked.
_RE_QUALIFIED_CELL = re.compile(
    r"(?<![^\s=(,+\-*/^&<>{;@])"
    r"(?:'((?:[^']|'')+)'|([^\W\d][\w.]*(?::[^\W\d][\w.]*)?))"
    r"!\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?![\w.(\[!'])")
# INDIRECT( as a call (searched in masked text) and the string literal that starts its address,
# followed by '&' and a part that is not a string literal (searched in the original text).
_RE_INDIRECT_CALL = re.compile(r"(?<![\w.])(?:_xlfn\.)?INDIRECT\(", re.I)
_RE_LEAD_LITERAL = re.compile(r'[ \t\r\n]*"((?:[^"]|"")*)"[ \t\r\n]*&[ \t\r\n]*(?![" \t\r\n])')
# Cheap pre-test for is_hyperlink_formula (the parse confirms the call shape).
_RE_HYPERLINK_CALL = re.compile(r"HYPERLINK[ \t\r\n]*\(", re.I)
_WS = " \t\r\n"
# Guard functions that may wrap a navigation HYPERLINK call without making the cell a
# calculation, provided every other argument is a single literal (or empty): the ChatGPT writer's
# =IFERROR(HYPERLINK("#'Summary'!A1","OPEN"),"OPEN").  Second review of 49, 2026-10-04 (ruling
# 2026-10-03 "HYPERLINK cells any colour", rationale "not a calculation").  () = top-level only.
HYPERLINK_GUARDS = ("IFERROR", "IFNA", "IF")


def is_hyperlink_formula(text: Optional[str]) -> bool:
    """True for a HYPERLINK navigation cell (Patrick 2026-10-03: any colour in 49 / 50 / 51).
    After optional whitespace and one leading '=', the whole formula is one built-in
    HYPERLINK(...) call, or a HYPERLINK_GUARDS call around it whose other arguments are each one
    literal or empty (guards may nest; the HYPERLINK call's own arguments are free).  Anything
    else outside the link (an operator, another function, a guard argument that reads a cell or
    calculates, parentheses, unary +, '@') makes it an ordinary formula.  Unparseable text ->
    False (the classifier then decides, or raises)."""
    if not text or _RE_HYPERLINK_CALL.search(text) is None:
        return False
    try:
        f = F.parse(text)
    except F.FormulaError:
        return False
    body = f.body
    for i, c in enumerate(f.functions):
        if c.depth == 0:                    # the outermost call: nothing may stand around it
            return (body[:c.start].strip(_WS) == "" and body[c.end:].strip(_WS) == ""
                    and _navigation_call(f, i))
    return False


def _navigation_call(f, i: int) -> bool:
    """Call i of f is a built-in HYPERLINK call, or a HYPERLINK_GUARDS call each of whose
    arguments is empty, one literal, or (exactly, with nothing around it) such a call again -
    with at least one HYPERLINK reached."""
    c = f.functions[i]
    if not c.builtin or c.local or c.implicit or c.spilled:
        return False
    if c.name == "HYPERLINK":
        return True
    if c.name not in HYPERLINK_GUARDS:
        return False
    body = f.body
    found = False
    for k, (a, b) in enumerate(c.arg_spans):
        content = body[a:b].strip(_WS)
        if not content:
            continue                                                  # omitted argument
        inner = [j for j, g in enumerate(f.functions) if g.parent == i and g.parent_arg == k]
        if inner:
            g = f.functions[inner[0]]
            if len(inner) != 1 or body[g.start:g.end] != content or not _navigation_call(f, inner[0]):
                return False
            found = True
            continue
        ops = [o for o in f.operands if a <= o.start and o.end <= b]
        if len(ops) != 1 or ops[0].kind not in _LITERAL_KINDS or ops[0].raw != content:
            return False
    return found


def _mask(text: str, *, quoted: Optional[str], brackets: bool, bracket_fill: Optional[str] = None) -> str:
    """String literals blanked (quotes kept); [...] groups blanked when `brackets` (to spaces, or
    to `bracket_fill` when given); a quoted sheet name replaced by `quoted` (None = kept)."""
    def sub(m):
        g = m.lastgroup
        s = m.group()
        if g == "s":
            return '"' + " " * (len(s) - 2) + '"'
        if g == "b":
            if not brackets:
                return s
            return " " * len(s) if bracket_fill is None else bracket_fill
        return s if quoted is None else quoted
    return _RE_SCAN.sub(sub, text)


def pointer_shaped(text: str) -> bool:
    """False guarantees the formula text is not a pure reference (so never a POINTER): after
    masking string literals, quoted sheet names and [...] groups, and ignoring one leading '=',
    it must have the shape head + one operand + tail described at _RE_POINTER_HEAD.  It is a
    cheap, conservative shape test, not a parse: True only says 'may be a pure reference'."""
    if text.startswith("="):
        text = text[1:]
    t = _mask(text, quoted="Q", brackets=True, bracket_fill="B")
    if any(ch in t for ch in _NOT_IN_POINTER):
        return False
    m = _RE_POINTER_HEAD.match(t)
    head, rest = (t[:m.end()], t[m.end():]) if m else ("", t)
    mt = _RE_POINTER_TAIL.search(rest)
    core, tail = (rest[:mt.start()], rest[mt.start():]) if mt else (rest, "")
    # the parentheses opened in the head must all close in the tail (the parser rejects the rest)
    return bool(core) and _RE_POINTER_CORE.match(core) is not None and head.count("(") == tail.count(")")


def bookish_qualifier(text: str) -> bool:
    """False guarantees that no qualifier in text names another workbook by path or file name."""
    if "!" not in text:
        return False
    if "\\" in text or _BOOK_EXT_BANG.search(text):
        return True
    if "/" in text and "'" in text:
        return any("/" in m.group(1) for m in _QUOTED_QUAL.finditer(F.mask_strings(text)))
    return False


def quick_reads_other_sheet(text: str, own: str) -> bool:
    """True guarantees that the formula reads another sheet of this workbook or another
    workbook (so its class is CROSS, POINTER or EXTERNAL, never OWN), shown cheaply by
      * outside string literals and [...] groups, a sheet qualifier naming a sheet other
        than `own` directly followed by a valid A1 cell reference (Inputs!B5,
        'My Sheet'!$A$1:B9, Jan:Dec!B5, [1]Prices!A1); or
      * an INDIRECT call whose address starts with a string literal holding a qualifier of
        another sheet or workbook, followed by '&' and a part that is not a string literal
        (INDIRECT("'Close Prices'!B"&(r+1)): the rule of indirect_literal_qualifier).
    False means 'not shown this cheaply' (parse it)."""
    if "!" not in text:
        return False
    t = _mask(text, quoted=None, brackets=True)
    own_low = own.lower()
    for m in _RE_QUALIFIED_CELL.finditer(t):
        row = int(m.group(4))
        if not (1 <= row <= F.MAX_ROW) or F.col_index(m.group(3)) > F.MAX_COL:
            continue
        q = m.group(1)
        q = q.replace("''", "'") if q is not None else m.group(2)
        if any(part.lower() != own_low for part in q.split(":")):
            return True
    if '"' in text:
        for m in _RE_INDIRECT_CALL.finditer(t):           # INDIRECT( outside strings / brackets
            lm = _RE_LEAD_LITERAL.match(text, m.end())
            if lm is None:
                continue
            lit = lm.group(1).replace('""', '"')
            k = _bang_outside_quotes(lit)
            if k <= 0:
                continue
            try:
                qual = F.parse_qualifier(lit[:k])
            except F.FormulaError:
                continue                                    # INDIRECT returns #REF!: reads nothing
            if qual.external is not None:
                return True                                 # another workbook
            if qual.workbook_scoped or (qual.ambiguous_book and qual.sheet.lower() == own_low):
                continue
            if any(sh is not None and sh.lower() != own_low for sh in (qual.sheet, qual.sheet_end)):
                return True
    return False


def _inert(text: Optional[str]) -> bool:
    """A defined-name definition that reads no cell under any scope: it parses, every operand
    is a literal and it calls only built-in functions other than INDIRECT (=0.05, =12*0.08)."""
    try:
        f = F.parse(text or "")
    except F.FormulaError:
        return False
    return (all(o.kind in _LITERAL_KINDS for o in f.operands)
            and all(c.builtin and not c.local and c.name != "INDIRECT" for c in f.functions))


# ----------------------------------------------------------------------------- classification
class Cls:
    """Classification of one formula text on one sheet (shared by cache: never mutate).

    kind                 OWN / CROSS / POINTER / EXTERNAL (UNCLASSIFIED only for 49's messages)
    other_sheets         sheets of this workbook read, other than the own sheet
    indirect_unresolved  an INDIRECT address is built from cell contents and is not known
    indirect_cells       ((row, col), ...) own-sheet cells whose content is such an address
                         (=INDIRECT(A20)); None when some unknown address is not of that form
    unsure               why the class may differ (a defined name with an unknown scope whose
                         possible definitions differ), else None
    unsure_ext           whether such a name may also change whether another workbook is read
    direct_other         the formula's own operands (not through names) read another sheet or
                         workbook: then no name can make the class OWN
    direct_ext           the formula's own operands read another workbook
    pure_name            the formula is a pure reference whose operand is a name (=WACC)
    note                 for UNCLASSIFIED: why the class could not be worked out; for EXTERNAL:
                         a table lookup that failed (the kind is certain, other_sheets may be short)
    sheets_unknown       UNCLASSIFIED because a structured reference / bare table name could not
                         be placed on a sheet (table not defined, table parts unreadable): the
                         formula reads no other workbook (a table without [n]! is this workbook's),
                         but OWN / CROSS / POINTER cannot be told apart.  classify() raises `note`
                         for such a result unless the caller passes sheets_needed=False (51).
    """
    __slots__ = ("kind", "other_sheets", "indirect_unresolved", "indirect_cells", "indirect_part_cells",
                 "indirect_assumed", "unsure", "unsure_ext", "direct_other", "direct_ext", "pure_name", "note",
                 "sheets_unknown")

    def __init__(self, kind, other_sheets=(), indirect_unresolved=False, *, indirect_cells=None,
                 indirect_part_cells=None, indirect_assumed=False, unsure=None,
                 unsure_ext=False, direct_other=False, direct_ext=False, pure_name=False, note=None,
                 sheets_unknown=False):
        self.sheets_unknown = sheets_unknown
        self.kind = kind
        self.other_sheets = tuple(sorted(other_sheets))
        self.indirect_unresolved = indirect_unresolved
        self.indirect_cells = indirect_cells
        # Patrick 2026-10-06 (every attempt graded): own-sheet cells that are PARTS of an unknown
        # INDIRECT address (=INDIRECT(MID(I33,...)); 49 reads them for a sheet qualifier), and
        # whether a CROSS class rests on a sheet-qualified address whose sheet name is computed
        # ("'"&B3&"'!"&ADDRESS(22,D3): taken to read another sheet, recorded as a default)
        self.indirect_part_cells = indirect_part_cells
        self.indirect_assumed = indirect_assumed
        self.unsure = unsure
        self.unsure_ext = unsure_ext
        self.direct_other = direct_other
        self.direct_ext = direct_ext
        self.pure_name = pure_name
        self.note = note

    def __repr__(self):
        return (f"Cls({self.kind}, {self.other_sheets}, indirect_unresolved={self.indirect_unresolved}, "
                f"indirect_cells={self.indirect_cells}, unsure={self.unsure!r}, unsure_ext={self.unsure_ext})")


class Classifier:
    """Classifies formula texts of one workbook (cached per (sheet, text))."""

    CACHE_MAX = 200_000
    _per_workbook: "weakref.WeakKeyDictionary" = weakref.WeakKeyDictionary()

    @classmethod
    def of(cls, wb) -> "Classifier":
        """One classifier (and cache) per opened workbook, shared by 49, 50 and 51 when they
        grade the same file together (classification does not depend on the check)."""
        hit = cls._per_workbook.get(wb)
        if hit is None:
            hit = cls._per_workbook[wb] = cls(wb)
        return hit

    def __init__(self, wb):
        self._wb = weakref.ref(wb)      # no strong reference: the per-workbook registry is weak-keyed
        self.sheet_names = [s.name for s in wb.sheets]
        self._sheets_low = {s.lower() for s in self.sheet_names}
        good, self.bad_names = [], {}
        for dn in wb.defined_names:
            if dn.scope_error is not None:          # localSheetId names no sheet: scope unknown
                self.bad_names[dn.name.lower()] = dn.scope_error
            else:
                good.append(dn)
        self.names = F.NameTable(good, sheet_names=self.sheet_names)
        idents = {n.name.lower() for n in self.names} | set(self.bad_names)
        self.name_idents = idents | {n.split(".", 1)[1] for n in idents if n.startswith("_xlnm.")}
        # tables (read from the table parts via the sheet relationships; no sheet XML is read).
        # A table part that cannot be read raises only where a formula needs tables.
        self._tables: dict = {}
        self._tables_error: Optional[str] = None
        try:
            for si in wb.sheets:
                if si.kind != "worksheet" or si.part is None:
                    continue
                for t in wb.tables(si):
                    for nm in (t.display_name, t.name):
                        if nm:
                            self._tables[nm.lower()] = t.sheet
        except GradingError as e:
            self._tables, self._tables_error = {}, f"the workbook's table parts cannot be read ({e})"
        self.ref_idents = self.name_idents | set(self._tables) if BARE_TABLE_NAME_IS_REFERENCE else self.name_idents
        # defined names with an unknown scope: a reference to one may resolve to it or to a
        # same-named definition (or to nothing).  Harmless when every candidate is inert;
        # bad_ext: some candidate may read another workbook.
        self._bad_harmless: set = set()
        self._bad_ext: dict = {}
        for ident in self.bad_names:
            texts = [dn.text for dn in wb.defined_names if dn.name.lower() == ident]
            if all(_inert(t) for t in texts):
                self._bad_harmless.add(ident)
                self._bad_ext[ident] = False
            else:
                self._bad_ext[ident] = any("[" in (t or "") or bookish_qualifier(t or "") or self.mentions_name(t or "")
                                           for t in texts)
        self._cache: dict = {}
        self._ext_names: dict = {}

    # -- table name -> sheet
    def table_sheet(self, table: str, raw: str) -> str:
        if self._tables_error is not None:
            raise GradingError(f"structured reference {raw!r}: {self._tables_error}")
        hit = self._tables.get(table.lower())
        if hit is None:
            raise GradingError(f"structured reference {raw!r}: table {table!r} is not defined in this "
                               f"workbook, so the sheet it reads is unknown")
        return hit

    def bare_table_sheet(self, o) -> Optional[str]:
        """Sheet of a bare table name used as a reference (=SUM(Tbl)), else None.  A defined
        name of the same spelling wins (Excel never allows both)."""
        if not BARE_TABLE_NAME_IS_REFERENCE or o.kind != "name" or o.sheet is not None or o.external is not None \
                or o.workbook_scoped:
            return None
        low = o.name.lower()
        if low in self.name_idents:
            return None
        if self._tables_error is not None:
            raise GradingError(f"name {o.name!r} is not a defined name and may be a table, but "
                               f"{self._tables_error}")
        return self._tables.get(low)

    # -- cheap pre-filters (False guarantees the class cannot matter)
    def mentions_name(self, text: str) -> bool:
        """An identifier spelled like a defined name or (bare) table name appears in text
        (any identifier when the table parts could not be read)."""
        if self._tables_error is not None and BARE_TABLE_NAME_IS_REFERENCE:
            return _IDENT_RX.search(text) is not None
        if not self.ref_idents:
            return False
        return any(m.group(0).lower() in self.ref_idents for m in _IDENT_RX.finditer(text))

    @staticmethod
    def _scan_text(text: str) -> str:
        """Text to pre-filter: string literals blanked, unless the formula calls INDIRECT
        (whose literal address is a reference, see Classifier._classify)."""
        if '"' not in text or "INDIRECT" in text.upper():
            return text
        return F.mask_strings(text)

    def may_leave_sheet(self, text: str) -> bool:
        """False guarantees the text references neither another sheet nor another workbook:
        no '!' (sheet qualifier), no '[' (external book or table) and no identifier spelled
        like a defined name or table, outside string literals (inside them too when INDIRECT
        is called)."""
        t = self._scan_text(text)
        return "!" in t or "[" in t or self.mentions_name(t)

    def may_be_pointer(self, text: str) -> bool:
        """False guarantees the formula is not a POINTER: it cannot leave its sheet, or (string
        literals, quoted sheet names and [...] groups masked) it holds a character no pure
        reference has - an operator other than unary +/- (* / ^ & = < > %), an argument
        separator or a string - or it does not have the shape of one reference (pointer_shaped:
        a call other than the one-argument wrappers, or two operands, rule it out).  One leading
        '=' (tolerated by the parser) is ignored.  A formula rejected here is never parsed, so
        malformed text that cannot be a pointer (SUM(Inputs!A1:A5 with no ')') does not raise."""
        if text.startswith("="):
            text = text[1:]
        if not self.may_leave_sheet(text):
            return False
        return pointer_shaped(text)

    def may_read_other_book(self, text: str) -> bool:
        """False guarantees the formula cannot read another workbook: no '[' (book index, path
        or URL with [Book.xlsx]), no workbook-like qualifier before '!' (a path with \\ or /,
        or a name ending .xlsx/.xls/.csv/.ods...), and no defined name whose definition
        (transitively) may read another workbook - outside string literals, unless INDIRECT
        is called."""
        t = self._scan_text(text)
        return "[" in t or bookish_qualifier(t) or self.mentions_external_name(t)

    def mentions_external_name(self, text: str) -> bool:
        if not self.name_idents:
            return False
        for m in _IDENT_RX.finditer(text):
            ident = m.group(0).lower()
            if ident not in self.name_idents:
                continue
            hit = self._ext_names.get(ident)
            if hit is None:
                hit = True                       # unknown / unreadable: let classify() decide (or raise)
                try:
                    hit = any(F.references_external_workbook(dn.text, names=self.names, scope_sheet=dn.scope,
                                                             sheet_names=self.sheet_names)
                              or self._name_text_has_book(dn.text or "")
                              for dn in self.names if dn.name.lower() in (ident, "_xlnm." + ident))
                except GradingError:
                    hit = True
                if ident in self.bad_names:
                    hit = hit or self._bad_ext[ident]
                self._ext_names[ident] = hit
            if hit:
                return True
        return False

    @staticmethod
    def _name_text_has_book(text: str) -> bool:
        return "[" in text or bookish_qualifier(text)

    def classify(self, text: Optional[str], own: str, *, sheets_needed: bool = True) -> Cls:
        """The Cls of `text` on sheet `own`.  Raises GradingError (FormulaError) on malformed
        text, and - unless sheets_needed is False - when a structured reference or bare table
        name cannot be placed on a sheet and the class therefore stays unknown (Cls.sheets_unknown;
        51 never needs the sheets: a table of this workbook is never another workbook)."""
        if text is None:
            raise GradingError(f"a formula on sheet {own!r} has no text to classify")
        key = (own, text)
        hit = self._cache.get(key)
        if hit is None:
            hit = self._classify(text, own)
            if len(self._cache) < self.CACHE_MAX:
                self._cache[key] = hit
        if hit.sheets_unknown and sheets_needed:
            raise GradingError(hit.note)
        return hit

    def _bad_refs(self, f) -> list:
        """(ident, why) for each use in f of a defined name whose scope is unknown and whose
        possible definitions are not all inert."""
        if not self.bad_names:
            return []
        out = []
        for r in f.names_referenced():
            low = r.name.lower()
            why = self.bad_names.get(low)
            if why is not None and r.external is None and low not in self._bad_harmless:
                out.append((low, f"defined name {r.name!r} whose scope is unknown ({why})"))
        return out

    def _is_external(self, o) -> bool:
        return o.external is not None or (o.ambiguous_book and o.sheet.lower() not in self._sheets_low)

    def _other_sheets(self, operands, own: str, problems: Optional[list] = None) -> set:
        """Sheets of THIS workbook (other than own) named by reference operands; a table
        reference (also a bare table name) is placed on its table's sheet; [@Col] and
        unqualified references are on the own sheet (also inside a defined name: Excel
        resolves them on the using sheet).  A table that cannot be placed (not defined, table
        parts unreadable) raises GradingError, or - when `problems` is a list - is skipped and
        its message appended there (the caller decides whether the sheets matter)."""
        own_low = own.lower()
        out = set()
        for o in operands:
            if o.kind == "name":
                try:
                    tsh = self.bare_table_sheet(o)
                except GradingError as e:
                    if problems is None:
                        raise
                    problems.append(str(e))
                    continue
                if tsh is not None and tsh.lower() != own_low:
                    out.add(tsh)
                continue
            if o.kind not in _SHEET_REF_KINDS or self._is_external(o):
                continue
            if o.kind == "structured" and o.table is not None:
                try:
                    sheets = (self.table_sheet(o.table, o.raw),)
                except GradingError as e:
                    if problems is None:
                        raise
                    problems.append(str(e))
                    continue
            elif o.sheet is None:
                continue
            else:
                sheets = (o.sheet, o.sheet_end)
            for sh in sheets:
                if sh is not None and sh.lower() != own_low:
                    out.add(sh)
        return out

    def _classify(self, text: str, own: str) -> Cls:
        f = F.parse(text)                                   # FormulaError on malformed text
        bad = self._bad_refs(f)
        reach = F.expand(f, own, self.names, sheet_names=self.sheet_names)
        for dn in reach.names:
            bad += self._bad_refs(dn.formula())
        direct = [o for via, o in reach.operands if not via]
        named = [o for via, o in reach.operands if via]          # reached through defined names
        direct_ext = any(self._is_external(o) for o in direct)
        ext = direct_ext or any(self._is_external(o) for o in named)
        # table lookups that fail are collected, not raised: whether the formula reads another
        # workbook never depends on them (second review of 51, 2026-10-04)
        problems: list = []
        other = self._other_sheets(direct, own, problems)
        direct_other = bool(other)
        if named:
            other |= self._other_sheets(named, own, problems)
        unresolved = False
        assumed = False
        cells: Optional[list] = []
        part_cells: list = []
        if "INDIRECT" in reach.functions:
            toks: dict = {}
            for via, call in reach.calls:
                if call.name != "INDIRECT" or call.local or not call.arg_spans:
                    continue
                fx = via[-1].formula() if via else f
                tk = toks.get(id(fx))
                if tk is None:
                    tk = toks[id(fx)] = [t for t in fx.tokens if t[0] != "ws"]
                res = self._indirect(fx, call, tk, own, bad, problems)
                if res is None:
                    unresolved = True
                    cell = None if via else indirect_cell_arg(fx, call, own)
                    if cell is None:
                        cells = None
                        if not via:
                            part_cells += indirect_arg_cells(fx, call, own)
                    elif cells is not None:
                        cells.append(cell)
                    continue
                r_ext, r_other, r_assumed = res
                ext = ext or r_ext
                other |= r_other
                assumed = assumed or r_assumed
                if not via:
                    direct_ext = direct_ext or r_ext
                    direct_other = direct_other or bool(r_other) or r_assumed
        unsure = bad[0][1] if bad else None
        unsure_ext = any(self._bad_ext.get(ident, True) for ident, _ in bad)
        # pure reference: needed for POINTER (other sheets read) and for pure_name (unsure only)
        pr = f.pure_reference(transparent=POINTER_WRAPPERS) if len(f.operands) == 1 and (other or unsure) else None
        extra = dict(indirect_cells=tuple(cells) if (unresolved and cells) else None,
                     indirect_part_cells=tuple(dict.fromkeys(part_cells)) if (unresolved and part_cells) else None,
                     indirect_assumed=assumed, unsure=unsure,
                     unsure_ext=unsure_ext, direct_other=direct_other or direct_ext, direct_ext=direct_ext,
                     pure_name=pr is not None and pr.kind == "name")
        if ext:                                 # certain whatever the tables are (a table without [n]! is this workbook's)
            return Cls(EXTERNAL, other, unresolved, note=problems[0] if problems else None, **extra)
        if not problems and (other or assumed):
            try:
                tgt = self._pointer_sheets(f, own, own, 0, pr) if pr is not None else None
            except GradingError as e:           # a table reached only here (defensive: _other_sheets saw it first)
                problems.append(str(e))
                tgt = None
            if not problems:
                if tgt and any(s.lower() != own.lower() for s in tgt):
                    return Cls(POINTER, other, unresolved, **extra)
                return Cls(CROSS, other, unresolved, **extra)
        if problems:                            # not EXTERNAL; OWN / CROSS / POINTER undecidable
            return Cls(UNCLASSIFIED, other, unresolved, note=problems[0], sheets_unknown=True, **extra)
        return Cls(OWN, (), unresolved, **extra)

    def _indirect(self, fx, call, toks, own: str, bad: list, problems: Optional[list] = None) -> Optional[tuple]:
        """(reads another workbook?, other sheets read) of one INDIRECT call when its address
        is known syntactically, else None.  Known: a literal address; a computed address whose
        leading literal holds the qualifier ("'Stock Prices'!"&ADDRESS(r,c,4)); an address
        assembled from literals, ADDRESS(...) calls with a literal or no sheet, and numeric
        parts ("B"&ROW(), ADDRESS(ROW(),2,4,TRUE,"Inputs"))."""
        lit = literal_text(fx, call, 0, toks)
        if lit is not None:
            iops = self.indirect_operands(lit, own, bad, problems)
            return any(self._is_external(o) for o in iops), self._other_sheets(iops, own, problems), False
        q = indirect_literal_qualifier(fx, call, toks)
        if q is not None:
            try:
                qual = F.parse_qualifier(q)
            except F.FormulaError:
                return False, set(), False           # INDIRECT returns #REF!: reads nothing
            if qual.external is not None or (qual.ambiguous_book and qual.sheet.lower() not in self._sheets_low):
                return True, set(), False
            if qual.workbook_scoped:
                return False, set(), False
            return (False, {sh for sh in (qual.sheet, qual.sheet_end) if sh is not None and sh.lower() != own.lower()},
                    False)
        sample = indirect_sample(fx, call, toks)
        res = None
        if sample is not None:
            try:
                fs = F.parse(sample)
            except F.FormulaError:
                fs = None
            if fs is not None and not fs.functions and len(fs.operands) == 1:
                o = fs.operands[0]
                if o.kind in _LITERAL_KINDS:
                    return False, set(), False       # not a reference: INDIRECT returns #REF!
                if o.kind in _SHEET_REF_KINDS and o.kind != "structured":
                    res = self._is_external(o), self._other_sheets([o], own, problems), False
        if res is not None:
            return res
        # Patrick 2026-10-06 (every attempt graded): an address the parts cannot stand in for.  Its string
        # literals, wherever they sit in the argument (inside MID / FIND too), decide when they name a
        # sheet: "Assumptions!" names another sheet of this workbook -> reads it; a literal '!' whose
        # name is computed ("'"&B3&"'!"&ADDRESS(22,D3)) -> taken to read another sheet (assumed; 49
        # records the default).  No sheet qualifier in any literal -> unknown (None).
        kind, sheets = indirect_literal_sheets(fx, call, toks, own, self._sheets_low)
        if kind == "known":
            return False, sheets, False
        if kind == "assumed":
            return False, set(), True
        return None

    def indirect_operands(self, lit: str, own: str, bad: Optional[list] = None, problems: Optional[list] = None) -> list:
        """Reference operands of a literal INDIRECT address ("Inputs!B5", "'[B.xlsx]S'!A1",
        "Rate" -> the name's references, "Tbl" -> the table).  Text that is not a reference
        makes INDIRECT return #REF! (it reads nothing), so it yields no operand.  Uses of
        defined names with an unknown scope are appended to `bad`; a name that may be a table
        while the table parts cannot be read raises, or is dropped with its message appended to
        `problems` when that is a list (it is never another workbook)."""
        try:
            f = F.parse(lit)
        except F.FormulaError:
            return []
        if len(f.operands) != 1 or f.functions:
            return []                               # INDIRECT accepts one reference or name only
        if bad is not None:
            bad += self._bad_refs(f)
        reach = F.expand(f, own, self.names, sheet_names=self.sheet_names)
        if bad is not None:
            for dn in reach.names:
                bad += self._bad_refs(dn.formula())
        out = []
        for _, o in reach.operands:
            if o.is_reference:
                out.append(o)
                continue
            try:
                if self.bare_table_sheet(o) is not None:
                    out.append(o)
            except GradingError as e:
                if problems is None:
                    raise
                problems.append(str(e))
        return out

    def _pointer_sheets(self, f, using_sheet: str, scope: Optional[str], depth: int, op=None) -> Optional[tuple]:
        """Sheets a pure-reference formula points at, or None when f is not a pure reference
        (a name is followed to its definition; unqualified references are on the using sheet)."""
        if op is None:
            op = f.pure_reference(transparent=POINTER_WRAPPERS)
        if op is None or depth > 20:
            return None
        if not POINTER_ALLOWS_NEGATION and any(t[0] == "op" and t[1] == "-" for t in f.tokens):
            return None
        if self._is_external(op):
            return None
        if op.kind == "name":
            dn = self.names.resolve(op.name, scope, op.sheet, op.workbook_scoped)
            if dn is None:
                tsh = self.bare_table_sheet(op)
                return (tsh,) if tsh is not None else None
            return self._pointer_sheets(dn.formula(), using_sheet, dn.scope, depth + 1)
        if op.kind == "structured":
            return (using_sheet,) if op.table is None else (self.table_sheet(op.table, op.raw),)
        if op.sheet is None:
            return (using_sheet,)
        return tuple(s for s in (op.sheet, op.sheet_end) if s is not None)


def _arg_tokens(call, i: int, toks: list) -> list:
    a, b = call.arg_spans[i]
    return [t for t in toks if a <= t[2] and t[3] <= b]


def indirect_literal_qualifier(f, call, toks: Optional[list] = None) -> Optional[str]:
    """Sheet / book qualifier of an INDIRECT address whose first argument starts with string
    literals that already contain the '!' ("'Stock Prices'!"&ADDRESS(r,c,4) -> "'Stock Prices'",
    "Calc!B"&ROW() -> "Calc"), else None (the qualifier, if any, is computed)."""
    tk = _arg_tokens(call, 0, toks if toks is not None else [t for t in f.tokens if t[0] != "ws"])
    out, i = [], 0
    while i < len(tk) and tk[i][0] == "str":
        out.append(tk[i][1][1:-1].replace('""', '"'))
        if i + 1 < len(tk) and not (tk[i + 1][0] == "op" and tk[i + 1][1] == "&"):
            return None                     # not a top-level concatenation
        i += 2
    k = _bang_outside_quotes("".join(out))
    return "".join(out)[:k] if k > 0 else None


def _bang_outside_quotes(text: str) -> int:
    """Index of the first '!' outside '...' quotes and [...] brackets, or -1."""
    q, depth, i = False, 0, 0
    while i < len(text):
        c = text[i]
        if q:
            if c == "'":
                if text[i + 1:i + 2] == "'":
                    i += 2
                    continue
                q = False
        elif c == "'":
            q = True
        elif c == "[":
            depth += 1
        elif c == "]":
            depth = max(0, depth - 1)
        elif c == "!" and depth == 0:
            return i
        i += 1
    return -1


def literal_text(f, call, i: int, toks: Optional[list] = None) -> Optional[str]:
    """Value of argument i of `call` when it is string literals joined by '&', else None."""
    tk = _arg_tokens(call, i, toks if toks is not None else [t for t in f.tokens if t[0] != "ws"])
    if not tk:
        return None
    out, expect = [], True
    for kind, ttext, _, _ in tk:
        if expect and kind == "str":
            out.append(ttext[1:-1].replace('""', '"'))
            expect = False
        elif not expect and kind == "op" and ttext == "&":
            expect = True
        else:
            return None
    return None if expect else "".join(out)


def indirect_sample(f, call, toks: list) -> Optional[str]:
    """A stand-in for an INDIRECT address assembled with '&' from parts whose contribution to
    the qualifier and shape is known, else None.  Parts: string literals (kept), ADDRESS(...)
    with no / an empty sheet argument ('$A$1') or a literal sheet ("'Inputs'!$A$1"), and
    numeric parts - numbers, + - * / ^ %, parentheses and NUMERIC_FUNCS calls ('1').  The
    stand-in names the same sheets as every value the real address can take."""
    tk = _arg_tokens(call, 0, toks)
    if not tk:
        return None
    by_start = {c.start: c for c in f.functions}
    parts, cur, depth = [], [], 0
    for t in tk:
        if t[0] in ("(", "{"):
            depth += 1
        elif t[0] in (")", "}"):
            depth -= 1
        if depth == 0 and t[0] == "op" and t[1] == "&":
            parts.append(cur)
            cur = []
        else:
            cur.append(t)
    parts.append(cur)
    out = []
    for part in parts:
        piece = _sample_piece(f, part, by_start)
        if piece is None:
            return None
        out.append(piece)
    return "".join(out)


def _sample_piece(f, part: list, by_start: dict) -> Optional[str]:
    if not part:
        return None
    if len(part) == 1 and part[0][0] == "str":
        return part[0][1][1:-1].replace('""', '"')
    if part[0][0] == "func":
        c = by_start.get(part[0][2])
        if c is not None and c.name == "ADDRESS" and c.builtin and not c.local and c.end == part[-1][3]:
            if len(c.arg_spans) < 5 or not f.body[c.arg_spans[4][0]:c.arg_spans[4][1]].strip(" \t\r\n"):
                return "$A$1"
            s = literal_text(f, c, 4, part)
            if s is None:
                return None
            if s == "":
                return "$A$1"
            q = s if s.startswith("'") else "'" + s.replace("'", "''") + "'"
            return q + "!$A$1"
    i = 0
    while i < len(part):
        k, tt, ts, _ = part[i]
        if k in ("num", "(", ")") or (k == "op" and tt in "+-*/^%"):
            i += 1
            continue
        if k == "func":
            c = by_start.get(ts)
            if c is None or c.local or not c.builtin or c.name not in NUMERIC_FUNCS:
                return None
            while i < len(part) and part[i][2] < c.end:
                i += 1
            continue
        return None
    return "1"


def indirect_cell_arg(f, call, own: str) -> Optional[tuple]:
    """(row, col) when INDIRECT's address argument is exactly one cell reference on the own
    sheet (=INDIRECT(A20), =INDIRECT('Calc'!$A$20) on Calc), else None."""
    a, b = call.arg_spans[0]
    ops = [o for o in f.operands if a <= o.start and o.end <= b]
    if len(ops) != 1 or ops[0].kind != "cell" or ops[0].external is not None or ops[0].sheet_end is not None:
        return None
    o = ops[0]
    if o.sheet is not None and o.sheet.lower() != own.lower():
        return None
    if f.body[a:b].strip(" \t\r\n") != o.raw:
        return None
    return o.bounds[0], o.bounds[1]


def indirect_arg_cells(f, call, own: str) -> list:
    """(row, col) of every own-sheet single-cell operand inside INDIRECT's address argument
    (=INDIRECT(MID($I33, FIND("!", $I33), 99)) -> I33): the cells whose constant text may
    name the sheet the address reaches (Patrick 2026-10-06)."""
    a, b = call.arg_spans[0]
    out = []
    for o in f.operands:
        if not (a <= o.start and o.end <= b) or o.kind != "cell" or o.external is not None \
                or o.sheet_end is not None:
            continue
        if o.sheet is not None and o.sheet.lower() != own.lower():
            continue
        out.append((o.bounds[0], o.bounds[1]))
    return out


_RE_QUALIFIER_IN_TEXT = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.][^!'\"\[\]:;,&()\s]*))!")


def sheets_named_in_text(text: str) -> list:
    """Sheet names that appear as a qualifier ("Assumptions!E6", "'Case 03'!X22") anywhere in
    `text` (a string literal or a constant cell's text), in order."""
    out = []
    for m in _RE_QUALIFIER_IN_TEXT.finditer(text or ""):
        out.append((m.group(1) or m.group(2)).replace("''", "'"))
    return out


def indirect_literal_sheets(f, call, toks: list, own: str, sheets_low) -> tuple:
    """What the string literals inside INDIRECT's address argument say about the sheet it
    reaches (Patrick 2026-10-06, every attempt graded).  ("known", {other sheets}) when a
    literal names a sheet of this workbook as a qualifier ("Assumptions!" inside MID / FIND
    counts; only the own sheet named -> ("known", set())); ("assumed", None) when the literals
    hold a '!' qualifier whose name is computed ("'"&B3&"'!"&ADDRESS(22,D3)); (None, None)
    when no literal carries a qualifier."""
    tk = _arg_tokens(call, 0, toks)
    lits = [t[1][1:-1].replace('""', '"') for t in tk if t[0] == "str"]
    if not lits:
        return None, None
    named = []
    for lit in lits:
        for name in sheets_named_in_text(lit):
            if name.lower() in sheets_low:
                named.append(name)
    if named:
        return "known", {n for n in named if n.lower() != own.lower()}
    # "assumed" needs the '!' in a literal that is a top-level '&' part of the address itself
    # ("'"&B3&"'!"&...), not one inside a function call (FIND("!", I33) locates text)
    parts, cur, depth = [], [], 0
    for t in tk:
        if t[0] in ("(", "{"):
            depth += 1
        elif t[0] in (")", "}"):
            depth -= 1
        if depth == 0 and t[0] == "op" and t[1] == "&":
            parts.append(cur)
            cur = []
        else:
            cur.append(t)
    parts.append(cur)
    top = [p[0][1][1:-1].replace('""', '"') for p in parts if len(p) == 1 and p[0][0] == "str"]
    if _bang_outside_quotes("".join(top)) >= 0:
        return "assumed", None
    return None, None


class Pending:
    """49: a cell decided after the second pass reads its INDIRECT address cells.  `parts`:
    the targets are parts of the address (=INDIRECT(MID(I33,...))), not the whole address."""
    __slots__ = ("targets", "parts")

    def __init__(self, targets, parts: bool = False):
        self.targets = tuple(targets)
        self.parts = parts


# ----------------------------------------------------------------------------- base check
class ColourCheck(Check):
    """Streams every formula cell, resolves its font colour family and, for cells outside
    `safe_family` that pass `prefilter`, classifies the formula and asks `judge` whether the
    cell is a mistake.  Offending cells are grouped per (reason, colour) into rectangles."""

    needs_cells = True
    sheet_kinds = ("worksheet", "macrosheet", "dialogsheet")
    safe_family = "black"          # cells in this family can never fail the check

    def start(self, wb):
        super().start(wb)
        if NUMFMT_COLOUR_COUNTS:
            raise GradingError(f"{self.key}: the displayed-colour variant (number-format colour tags) is not "
                               f"implemented; see docs/checks/49.md")
        self.st = wb.styles
        self.cf_excused: dict = {"count": 0, "examples": []}          # cells excused by a CF font colour
        self.cf_unresolved: dict = {"count": 0, "examples": []}       # dxf font colours that could not be read
        self.clf = Classifier.of(wb)
        self._fam_cache: dict = {}
        self.n_formula_cells = 0
        self.n_member_cells = 0
        self.n_empty_markers = 0
        self.n_textless_arrays = 0
        self.n_classified = 0
        self.fam_counts = {"black": 0, "green": 0, "red": 0, "other": 0, "unresolved": 0}
        self.class_counts = {c: 0 for c in CLASSES + (UNCLASSIFIED,)}
        self.bad_counts: dict = {}
        self.bad_by_sheet: dict = {}
        self.sheets_seen = 0
        self._pending: dict = {}       # sheet -> [((row, col), argb, fkey, info, targets)]
        self._addr: dict = {}          # second pass: (row, col) -> ("const", value) | ("formula", ref)
        self.n_pending = 0
        self.n_hyperlink = 0           # HYPERLINK navigation cells ignored (Patrick 2026-10-03)
        self._hyper: dict = {}         # memo: formula text -> is_hyperlink_formula

    def style_colour(self, s: int, where: str) -> tuple:
        """(argb, family) of a style's font.  A colour the core cannot resolve (styles.UnknownColour) is read as
        Excel's default for a font colour, automatic = black (Patrick 2026-10-05: every attempt graded;
        recorded by cell() in stats.defaults.unresolved_colour)."""
        hit = self._fam_cache.get(s)
        if hit is None:
            font = self.st.font(s)
            argb = self.st.resolve(font.color, "font")
            if is_unknown(argb):
                hit = (DEFAULT_FONT_ARGB, "black", f"font colour {font.color.describe()} of style {s} cannot be "
                                                   f"resolved ({argb.why}): read as automatic (black)")
            else:
                hit = (argb, family(argb), None)
            self._fam_cache[s] = hit
        return hit[0], hit[1]

    def sheet_start(self, head):
        self.sheets_seen += 1
        self._sheet = head.name
        self._shared: dict = {}        # si -> (master text, row, col)
        self._cand: dict = {}          # (reason, argb) -> {(row, col): fkey}
        self._fkeys: dict = {}         # fkey -> (text, row, col) for messages
        self._child = False

    def cell(self, cell):
        f = cell.formula
        arr = cell.array
        if f is None and arr is None:
            return
        if f is not None and f.is_empty_marker and arr is None:
            self.n_empty_markers += 1
            return
        # an array anchor written without any formula text (<f t="array" ref="A26:A76"/>, seen in an
        # agent-written file) and its members hold no formula: treated like an empty <f/> marker
        af = f if (f is not None and not f.is_empty_marker) else (arr.formula if arr is not None else None)
        if af is not None and af.kind == "array" and not (af.text or "").strip():
            self.n_textless_arrays += 1
            return
        own = self._sheet
        self.n_formula_cells += 1
        if f is not None and f.kind == "shared" and f.text:
            self._shared[f.si] = (f.text, cell.row, cell.col)
        where = f"'{own}'!{cell.ref}"
        argb, fam = self.style_colour(cell.s, where)
        defaulted = self._fam_cache[cell.s][2]
        if defaulted:
            self.fam_counts["unresolved"] += 1
            self.note_default("unresolved_colour", f"{where}: {defaulted}")
        else:
            self.fam_counts[fam] += 1
        if fam == self.safe_family:
            return
        # the formula that decides the class
        child = False
        if arr is not None and (f is None or f.is_empty_marker):          # spill / array / data-table member
            self.n_member_cells += 1
            af = arr.formula
            if af.kind == "dataTable":
                text, fkey, info = None, ("dt", arr.anchor_ref), (None, arr.r1, arr.c1)
            else:
                text, fkey, info = af.text, ("anchor", arr.anchor_ref), (af.text, arr.r1, arr.c1)
        elif f.kind == "dataTable":
            text, fkey, info = None, ("dt", cell.ref), (None, cell.row, cell.col)
        elif f.kind == "shared" and not f.text:                           # shared child: the master's class
            child = True
            info = self._shared.get(f.si)
            if info is not None:
                text = info[0]
            else:
                try:
                    text = cell.formula_text                               # no master: formula_text raises
                except GradingError as e:
                    # Patrick 2026-10-05 (every attempt graded): a formula text that cannot be read is skipped
                    self.note_default("unparsable_formula", f"{where}: {e}")
                    return
            fkey = ("si", f.si)
        else:
            text, fkey, info = f.text, ("cell", cell.row, cell.col), (f.text, cell.row, cell.col)
        if text is not None and self._is_hyperlink(text):
            self.n_hyperlink += 1                  # navigation link: any colour (ruling 2026-10-03)
            return
        if text is None:
            cls = Cls(DATA_TABLE_CLASS)
        else:
            if not self.prefilter(text, fam, own):
                return
            cls = self._classify_cell(text, child, cell, fam)
            if cls is None:
                return                              # unparsable: skipped (recorded)
        self.n_classified += 1
        self.class_counts[cls.kind] += 1
        self._child = child
        why = self.judge(cls, fam, argb, cell)
        if why is None:
            return
        if isinstance(why, Pending):
            self.n_pending += 1
            self._pending.setdefault(own, []).append(((cell.row, cell.col), argb, fkey, info, why))
            self.request_second_pass(own, why.targets)
            return
        self._cand.setdefault((why, argb), {})[(cell.row, cell.col)] = fkey
        if fkey not in self._fkeys and info is not None:
            self._fkeys[fkey] = info

    def _is_hyperlink(self, text: str) -> bool:
        hit = self._hyper.get(text)
        if hit is None:
            hit = is_hyperlink_formula(text)
            if len(self._hyper) < 50_000:
                self._hyper[text] = hit
        return hit

    def _classify_cell(self, text: str, child: bool, cell, fam: str) -> Optional[Cls]:
        """The cell's class; None when it is needed but the formula cannot be read - malformed text, or a
        structured reference / bare table name whose table the workbook does not define (Excel cannot parse
        such a formula either) or whose table parts cannot be read: the cell is skipped (Patrick 2026-10-05:
        every attempt graded; stats.defaults.unparsable_formula)."""
        sheets = self.sheets_needed(fam, text)
        try:
            cls = self.clf.classify(text, self._sheet, sheets_needed=sheets)
            if child and cls.kind == POINTER:
                # translation keeps sheets and shape, except a reference pushed off the grid
                # (Sheet!#REF!), which is no pointer: confirm on this cell's own text
                cls = self.clf.classify(cell.formula_text, self._sheet, sheets_needed=sheets)
            return cls
        except GradingError as e:
            if self.class_needed(fam, text):
                self.note_default("unparsable_formula", f"'{self._sheet}'!{cell.ref}: {e}")
                return None
            return Cls(UNCLASSIFIED, note=str(e))

    # -- subclass API
    def prefilter(self, text: str, fam: Optional[str], own: str) -> bool:
        """False = the formula cannot be in a class that fails this check (no parse needed)."""
        return True

    def class_needed(self, fam: Optional[str], text: Optional[str] = None) -> bool:
        """False = a cell of this colour is decided without its class (the class only words
        the message, so a formula that cannot be classified does not raise)."""
        return True

    def sheets_needed(self, fam: Optional[str], text: Optional[str] = None) -> bool:
        """False = this check's decision never depends on WHICH sheet of this workbook a
        formula reads, only on whether it reads another workbook (51): a structured reference
        or bare table name whose table is not defined or whose table parts cannot be read then
        gives Cls.sheets_unknown (kind UNCLASSIFIED, never EXTERNAL) instead of raising."""
        return True

    def judge(self, cls: Cls, fam: str, argb: str, cell):
        """None (fine), a reason string (mistake) or a Pending (decide after a second pass)."""
        raise NotImplementedError

    def describe(self, why: str, argb: str) -> str:
        raise NotImplementedError

    def pending_is_mistake(self, sheet: str, targets: tuple, parts: bool = False) -> Optional[str]:
        raise NotImplementedError

    def sheet_end(self, head, tail):
        self._emit(head, self._cf_excuse(head, tail, self._cand), self._fkeys)
        self._cand, self._fkeys, self._shared = {}, {}, {}

    # ------------------------------------------------------------ conditional-format excuse
    def required_families(self, why) -> frozenset:
        """Font families the cell's class accepts (its base font is in none of them, else it would
        not be a candidate); a conditional-format font colour in one of them excuses the cell."""
        raise NotImplementedError

    def _cf_font_rules(self, sheet: str, tail) -> list:
        """[(ranges, family)] for every conditional-format rule on the sheet whose dxf sets a font
        colour that resolves; an unresolvable one is recorded and excuses nothing."""
        out = []
        for cf in getattr(tail, "conditional_formats", None) or []:
            for rule in cf.rules:
                if rule.type in ("colorScale", "dataBar", "iconSet"):
                    continue
                d = rule.dxf if rule.dxf is not None else self.st.dxf(rule.dxf_id)
                if d is None or d.font is None or d.font.color is None:
                    continue
                argb = self.st.resolve(d.font.color, "font")
                if argb is None or is_unknown(argb):
                    self.cf_unresolved["count"] += 1
                    if len(self.cf_unresolved["examples"]) < 10:
                        self.cf_unresolved["examples"].append(
                            f"{location(sheet, ','.join(cf.sqref.split()))}: conditional-format font colour {argb}")
                    continue
                out.append((cf.ranges, family(argb)))
        return out

    def _cf_excuse(self, head, tail, cand: dict) -> dict:
        """Drop candidate cells that a conditional format covering them paints in a family their
        class accepts (CF_FONT_COLOUR_EXCUSES, Patrick 2026-10-06); records them in cf_excused."""
        if not CF_FONT_COLOUR_EXCUSES or not cand:
            return cand
        rules = self._cf_font_rules(head.name, tail)
        if not rules:
            return cand
        out = {}
        for (why, argb), cells in cand.items():
            ok = self.required_families(why)
            keep = {}
            for rc, fkey in cells.items():
                painted = next((fam for ranges, fam in rules
                                if fam in ok and any(in_bounds(rc[0], rc[1], b) for b in ranges)), None)
                if painted is not None:
                    self.cf_excused["count"] += 1
                    if len(self.cf_excused["examples"]) < 10:
                        self.cf_excused["examples"].append(
                            f"{location(head.name, f'{index_to_col(rc[1])}{rc[0]}')}: {describe_colour(argb)} base "
                            f"font, a conditional format paints it {painted}")
                else:
                    keep[rc] = fkey
            if keep:
                out[(why, argb)] = keep
        return out

    def second_pass_start(self, head):
        self._addr = {}

    def second_pass_cell(self, cell):
        if cell.formula is not None or cell.array is not None:
            self._addr[(cell.row, cell.col)] = ("formula", cell.ref)
        else:
            self._addr[(cell.row, cell.col)] = ("const", cell.value)

    def second_pass_end(self, head, tail):
        cand, fkeys = {}, {}
        for rc, argb, fkey, info, pend in self._pending.pop(head.name, []):
            why = self.pending_is_mistake(head.name, pend.targets, pend.parts)
            if why is not None:
                cand.setdefault((why, argb), {})[rc] = fkey
                if fkey not in fkeys and info is not None:
                    fkeys[fkey] = info
        self._emit(head, self._cf_excuse(head, tail, cand), fkeys)
        self._addr = {}

    def _emit(self, head, cand: dict, fkeys: dict):
        sheet = head.name
        for (why, argb), cells in cand.items():
            for b in group_cells(cells):
                rng = range_to_str(*b)
                n = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
                r, c = b[0], b[1]            # group_cells rectangles are fully covered
                hid = " (hidden sheet)" if head.state != "visible" else ""
                self.add_mistake(location(sheet, rng),
                                 f"{rng} ({n} cell{'s' if n != 1 else ''}){hid}: {self.describe(why, argb)}"
                                 f"{self._example(fkeys, cells[(r, c)], r, c)}")
            self.bad_counts[why] = self.bad_counts.get(why, 0) + len(cells)
            self.bad_by_sheet[sheet] = self.bad_by_sheet.get(sheet, 0) + len(cells)

    @staticmethod
    def _example(fkeys: dict, fkey, row, col) -> str:
        info = fkeys.get(fkey)
        ref = f"{index_to_col(col)}{row}"
        if info is None:
            return ""
        text, r0, c0 = info
        if fkey[0] == "dt":
            return f" (e.g. {ref}: what-if data table at {fkey[1]})"
        if fkey[0] == "anchor":
            return f" (e.g. {ref}: spilled from {fkey[1]} ={_short(text)})"
        if fkey[0] == "si" and (row, col) != (r0, c0):
            text = F.shift(text, row - r0, col - c0)
        return f" (e.g. {ref}: ={_short(text)})"

    def base_stats(self) -> dict:
        return {"sheets_streamed": self.sheets_seen,
                "formula_cells": self.n_formula_cells,
                "array_member_cells_examined": self.n_member_cells,
                "empty_f_markers_ignored": self.n_empty_markers,
                "textless_array_cells_ignored": self.n_textless_arrays,
                "font_families": dict(self.fam_counts),
                "hyperlink_cells_ignored": self.n_hyperlink,
                "cells_classified": self.n_classified,
                "classified_by_class": dict(self.class_counts),
                "offending_cells": sum(self.bad_counts.values()),
                "offending_cells_by_reason": dict(self.bad_counts),
                "offending_cells_by_sheet": dict(self.bad_by_sheet),
                "colour_basis": "font colour (number-format colour tags not applied; a conditional-format font "
                                "colour the class accepts excuses a cell, Patrick 2026-10-06)",
                "cf_excused": dict(self.cf_excused),
                "cf_unresolved_font_colours": dict(self.cf_unresolved)}


def _short(text: Optional[str], n: int = 90) -> str:
    t = (text or "").replace("\n", " ").replace("\r", " ")
    return t if len(t) <= n else t[:n - 3] + "..."
