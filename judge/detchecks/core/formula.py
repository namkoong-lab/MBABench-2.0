"""Excel formula parsing for the deterministic checks (pure functions, no I/O).

Used by checks 29, 49, 50, 51, 80, 87, 95.  Input is formula text exactly as stored in
the OOXML part (cell <f>, <definedName>, CF <formula>, DV <formula1>/<xm:f>), without
the leading '=' (one leading '=' is tolerated and stripped).

    f = parse("SUM('My Sheet'!A:A)+Rate")
    f.functions        -> (FunctionCall(name='SUM', ...),)
    f.operands         -> (Operand(kind='whole_column', sheet='My Sheet', ...),
                           Operand(kind='name', name='Rate', ...))
    f.strings          -> ()                       string literals (never references)

Design notes
- Own lexer.  openpyxl's Tokenizer was probed on every construct listed in
  docs/formula.md and gets several wrong: the spill operator (A1# raises), TAB/CR
  (glued onto the next operand), a function that ends a range (A1:OFFSET( becomes one
  FUNC token), '@' glued to operands.  The lexer here is regex driven, raises
  FormulaError on malformed text, and is differential-tested against the Tokenizer on
  every formula of the toy workbooks (tests/test_formula.py).
- Operands are classified from their text alone (context-free, cached).  Sheet names,
  defined names and tables are resolved by the caller (NameTable below), never guessed.
- Errors are loud: unbalanced parentheses, unterminated strings, unknown characters or
  an operand that is neither a reference nor a valid name raise FormulaError (a
  GradingError), so a check can never silently pass a formula it did not understand.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from typing import Iterable, Optional

from openpyxl.formula.translate import Translator, TranslatorError

from detchecks.errors import GradingError

MAX_ROW = 1048576
MAX_COL = 16384


class FormulaError(GradingError):
    """The formula text could not be parsed (or a helper could not decide)."""


class AmbiguousReferenceError(FormulaError):
    """A qualifier such as Other.xlsx!Rate may be an external workbook or a local sheet
    named 'Other.xlsx'; pass sheet_names= to decide."""


# --------------------------------------------------------------------------- built-ins
# Excel worksheet functions (openpyxl's list + post-2010 functions + legacy names).
# Used only to decide whether an unprefixed call X(...) is Excel's function or a call
# of a defined (LAMBDA) name X.  Prefixed calls (_xlfn.X) are always built-ins.
BUILTIN_FUNCTIONS = frozenset("""
ABS ACCRINT ACCRINTM ACOS ACOSH ACOT ACOTH ADDRESS AGGREGATE AMORDEGRC AMORLINC ANCHORARRAY AND
ARABIC AREAS ARRAYTOTEXT ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF
AVERAGEIFS BAHTTEXT BASE BESSELI BESSELJ BESSELK BESSELY BETA.DIST BETA.INV BETADIST BETAINV
BIN2DEC BIN2HEX BIN2OCT BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST BITAND BITLSHIFT BITOR
BITRSHIFT BITXOR BYCOL BYROW CALL CEILING CEILING.MATH CEILING.PRECISE CELL CHAR CHIDIST CHIINV
CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST CHOOSE CHOOSECOLS CHOOSEROWS
CLEAN CODE COLUMN COLUMNS COMBIN COMBINA COMPLEX CONCAT CONCATENATE CONFIDENCE CONFIDENCE.NORM
CONFIDENCE.T CONVERT COPILOT CORREL COS COSH COT COTH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS
COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR COVARIANCE.P COVARIANCE.S CRITBINOM CSC
CSCH CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE
CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS DAYS360 DB DBCS DCOUNT DCOUNTA DDB
DEC2BIN DEC2HEX DEC2OCT DECIMAL DEGREES DELTA DETECTLANGUAGE DEVSQ DGET DISC DMAX DMIN DOLLAR
DOLLARDE DOLLARFR DPRODUCT DROP DSTDEV DSTDEVP DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT
ENCODEURL EOMONTH ERF ERF.PRECISE ERFC ERFC.PRECISE ERROR.TYPE EUROCONVERT EVEN EXACT EXP EXPAND
EXPON.DIST EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FACT FACTDOUBLE FALSE FDIST FIELDVALUE
FILTER FILTERXML FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST
FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR
FORMULATEXT FREQUENCY FTEST FV FVSCHEDULE GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN
GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GETPIVOTDATA GROUPBY GROWTH HARMEAN HEX2BIN HEX2DEC
HEX2OCT HLOOKUP HOUR HSTACK HYPERLINK HYPGEOM.DIST HYPGEOMDIST IF IFERROR IFNA IFS IMABS IMAGE
IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2
IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMSUM IMTAN INDEX INDIRECT INFO INT
INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER
ISO.CEILING ISODD ISOMITTED ISOWEEKNUM ISPMT ISREF ISTEXT JIS KURT LAMBDA LARGE LCM LEFT LEFTB LEN
LENB LET LINEST LN LOG LOG10 LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST LOOKUP LOWER
MAKEARRAY MAP MATCH MAX MAXA MAXIFS MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINIFS MINUTE
MINVERSE MIRR MMULT MOD MODE MODE.MULT MODE.SNGL MONTH MROUND MULTINOMIAL N NA NEGBINOM.DIST
NEGBINOMDIST NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV
NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERSTRING NUMBERVALUE OCT2BIN OCT2DEC
OCT2HEX ODD ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PDURATION PEARSON PERCENTILE
PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT
PERMUTATIONA PHI PHONETIC PI PIVOTBY PMT POISSON POISSON.DIST POWER PPMT PRICE PRICEDISC PRICEMAT
PROB PRODUCT PROPER PV PY QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RADIANS RAND RANDARRAY
RANDBETWEEN RANK RANK.AVG RANK.EQ RATE RECEIVED REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST
REGISTER.ID REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS RRI RSQ RTD
SCAN SEARCH SEARCHB SEC SECH SECOND SEQUENCE SERIESSUM SHEET SHEETS SIGN SIN SINGLE SINH SKEW
SKEW.P SLN SLOPE SMALL SORT SORTBY SQL.REQUEST SQRT SQRTPI STANDARDIZE STDEV STDEV.P STDEV.S STDEVA
STDEVP STDEVPA STEYX STOCKHISTORY SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2
SUMX2PY2 SUMXMY2 SWITCH SYD T T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TAN TANH
TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE
TINV TOCOL TODAY TOROW TRANSLATE TRANSPOSE TREND TRIM TRIMMEAN TRIMRANGE TRUE TRUNC TTEST TYPE
UNICHAR UNICODE UNIQUE UPPER USDOLLAR VALUE VALUETOTEXT VAR VAR.P VAR.S VARA VARP VARPA VDB VLOOKUP
VSTACK WEBSERVICE WEEKDAY WEEKNUM WEIBULL WEIBULL.DIST WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS XIRR
XLOOKUP XMATCH XNPV XOR YEAR YEARFRAC YEN YIELD YIELDDISC YIELDMAT Z.TEST ZTEST _TRO_ALL
_TRO_LEADING _TRO_TRAILING
""".split())

VOLATILE_FUNCTIONS = frozenset({"INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN",
                                "INFO", "CELL", "RANDARRAY"})
# Function names whose first argument is a trimmed range (A:.A is stored as _TRO_TRAILING(A:A)).
TRIM_FUNCTIONS = {"_TRO_TRAILING": "trailing", "_TRO_LEADING": "leading", "_TRO_ALL": "all",
                  "TRIMRANGE": "trimrange"}
# Storage prefixes Excel adds to function names; stripped by normalisation.
_FUNC_PREFIXES = ("_xlfn.", "_xlws.", "_xll.", "_xludf.", "_xleta.")
_PARAM_PREFIX = "_xlpm."

ERROR_LITERALS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
                  "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!",
                  "#BUSY!", "#UNKNOWN!", "#PYTHON!", "#EXTERNAL!")

# Operand kinds -----------------------------------------------------------------------
REFERENCE_KINDS = frozenset({"cell", "range", "whole_column", "whole_row", "trimmed_range",
                             "spill", "structured"})
LITERAL_KINDS = frozenset({"number", "string", "bool", "error", "array"})
_SHIFTABLE = frozenset({"cell", "range", "whole_column", "whole_row", "trimmed_range", "spill"})

# --------------------------------------------------------------------------- lexer
_WS_CHARS = " \t\r\n"
_RE_ERR = re.compile("|".join(re.escape(e) for e in sorted(ERROR_LITERALS, key=len, reverse=True)), re.I)
_BRACKET = r"\[(?:[^\[\]']|'.|\[(?:[^\[\]']|'.)*\])*\]"
_RE_OPERAND = re.compile(
    r"(?:'(?:[^']|'')*'"                # quoted qualifier segment ('My Sheet', '[1]Data')
    r"|" + _BRACKET +                  # [1] / [Book.xlsx] / structured-reference spec
    r"|(?<=!)#REF!"                    # Sheet1!#REF!
    r"|[^ \t\r\n\"'()\[\]{},;+\-*/^&=<>%@#]"
    r")+", re.S | re.I)


def _outside_split_index(text: str, ch: str, last: bool) -> int:
    """Index of the first/last `ch` outside '...' quotes and [...] brackets, or -1."""
    q = False
    depth = 0
    found = -1
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if q:
            if c == "'":
                if i + 1 < n and text[i + 1] == "'":
                    i += 2
                    continue
                q = False
        elif depth:
            if c == "'" and i + 1 < n:          # escape inside structured references
                i += 2
                continue
            if c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
        elif c == "'":
            q = True
        elif c == "[":
            depth += 1
        elif c == ch:
            found = i
            if not last:
                return i
        i += 1
    return found


_RE_MASTER = re.compile(
    r"(?P<ws>[ \t\r\n]+)"
    r'|(?P<str>"(?:[^"]|"")*")'
    r"|(?P<num>(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)(?![A-Za-z0-9_$!.:\[\\?])"
    r"|(?P<word>" + _RE_OPERAND.pattern + r")(?P<call>\()?"
    r"|(?P<hash>\#)"
    r"|(?P<op2><>|<=|>=)"
    r"|(?P<punct>[(){},;+\-*/^&=<>%@])", re.S | re.I)
_PUNCT_KIND = {"(": "(", ")": ")", "{": "{", "}": "}", ",": ",", ";": ";"}


def lex(text: str) -> list:
    """Tokens of a formula body (no leading '='), as (kind, text, start, end) tuples.
    Kinds: ws str num err operand func ( ) { } , ; op spill.  A 'func' token holds the
    name only; the '(' that follows is its own token.  Raises FormulaError on
    malformed text."""
    s = text
    toks: list = []
    append = toks.append
    match = _RE_MASTER.match
    i, n = 0, len(s)
    while i < n:
        m = match(s, i)
        if m is None:
            c = s[i]
            if c == '"':
                raise FormulaError(f"unterminated string literal at position {i} in {text!r}")
            if c == "'":
                raise FormulaError(f"unterminated quoted sheet name at position {i} in {text!r}")
            if c in "[]":
                raise FormulaError(f"unbalanced bracket at position {i} in {text!r}")
            raise FormulaError(f"unexpected character {c!r} at position {i} in {text!r}")
        g = m.lastgroup
        e = m.end()
        if g == "word" or g == "call":
            ws, we = m.span("word")
            word = m.group("word")
            if m.group("call") is not None:
                if ":" in word:
                    k = _outside_split_index(word, ":", last=True)
                else:
                    k = -1
                if k > 0:      # A1:OFFSET( -> operand A1, op ':', func OFFSET
                    append(("operand", word[:k], ws, ws + k))
                    append(("op", ":", ws + k, ws + k + 1))
                    append(("func", word[k + 1:], ws + k + 1, we))
                elif k == 0:   # ...):OFFSET(
                    append(("op", ":", ws, ws + 1))
                    append(("func", word[1:], ws + 1, we))
                else:
                    append(("func", word, ws, we))
                append(("(", "(", we, we + 1))
            elif word[0] == ":" and len(word) > 1:          # INDEX(...):B5
                append(("op", ":", ws, ws + 1))
                append(("operand", word[1:], ws + 1, we))
            elif word[-1] == ":" and len(word) > 1:         # A1:(...) / A1:INDEX handled above
                append(("operand", word[:-1], ws, we - 1))
                append(("op", ":", we - 1, we))
            elif word == ":":
                append(("op", ":", ws, we))
            else:
                append(("operand", word, ws, we))
        elif g == "ws":
            append(("ws", m.group(), i, e))
        elif g == "str":
            append(("str", m.group(), i, e))
        elif g == "num":
            append(("num", m.group(), i, e))
        elif g == "punct":
            c = m.group()
            append((_PUNCT_KIND.get(c, "op"), c, i, e))
        elif g == "op2":
            append(("op", m.group(), i, e))
        else:  # hash
            prev = toks[-1][0] if toks else None
            if prev == "operand" or prev == ")":
                append(("spill", "#", i, e))               # A1#  /  (...)#
            else:
                me = _RE_ERR.match(s, i)
                if not me:
                    raise FormulaError(f"unknown error literal at position {i} in {text!r}")
                e = me.end()
                append(("err", me.group(), i, e))
        i = e
    return toks


# --------------------------------------------------------------------------- operands
_RE_CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$")
_RE_COL = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_RE_ROW = re.compile(r"^(\$?)(\d{1,7})$")
_RE_R1C1 = re.compile(r"^(?:[Rr]\d*[Cc]\d*|[Rr]\d*|[Cc]\d*)$")
_RE_NAME = re.compile(r"^[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\?\u0080-\uffff]*$")
_RE_BOOK_EXT = re.compile(r"\.(?:xl[a-z]{1,2}|xlsx|xlsm|xlsb|xls|xltx|xltm|csv|ods)$", re.I)


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letters(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def is_cell_reference(token: str) -> bool:
    """A1-style single-cell address within the grid (A1..XFD1048576), '$' allowed."""
    m = _RE_CELL.match(token)
    return bool(m) and col_index(m.group(2)) <= MAX_COL and 1 <= int(m.group(4)) <= MAX_ROW


def is_r1c1_reference(token: str) -> bool:
    """R, C, RC, R1, C5, R2C3 ... (R1C1 notation), case-insensitive.  Note R5 / C7 are
    also A1 cells; is_cell_reference decides first."""
    return bool(_RE_R1C1.match(token))


def name_validity_problem(name: str) -> Optional[str]:
    """None when `name` is a valid Excel defined-name, else the reason it is not.
    Excel resolves a token that looks like a cell (LIM1, TAX2024, XFD1048576) or an R1C1
    reference (R, C, RC, R2C3) as that reference, never as a name."""
    if not name:
        return "empty name"
    if len(name) > 255:
        return "longer than 255 characters"
    if name.lower().startswith("_xlnm."):
        return None                      # built-in (_xlnm.Print_Area ...)
    if not _RE_NAME.match(name):
        return "contains characters Excel does not allow in names"
    if is_cell_reference(name):
        return f"'{name}' is the cell address {name.upper()}"
    if is_r1c1_reference(name):
        return f"'{name}' reads as an R1C1 reference"
    if name.upper() in ("TRUE", "FALSE"):
        return "is a boolean literal"
    return None


def is_valid_name(name: str) -> bool:
    return name_validity_problem(name) is None


def _unquote(q: str) -> str:
    if len(q) >= 2 and q[0] == "'" and q[-1] == "'":
        return q[1:-1].replace("''", "'")
    return q


@dataclass(frozen=True, slots=True)
class Qualifier:
    """Parsed 'sheet / book' prefix of a reference (text before the first '!')."""
    raw: str                       # as written, e.g. "'[1]Sheet 1'"
    sheet: Optional[str]           # unquoted sheet name (first sheet of a 3-D span)
    sheet_end: Optional[str]       # last sheet of a 3-D span (Sheet1:Sheet3!A1), else None
    external: Optional[str]        # '1' (book index), 'Book.xlsx', 'C:\\dir\\Book.xlsx'; None = this book
    workbook_scoped: bool          # [0]!Name : explicitly this workbook's global name
    ambiguous_book: bool           # Other.xlsx!Name: a local sheet 'Other.xlsx' or another workbook


@lru_cache(maxsize=8192)
def parse_qualifier(raw: str) -> Qualifier:
    q = _unquote(raw.strip())
    external = None
    rest = q
    wb_scoped = False
    ambiguous = False
    if q.startswith("["):
        j = q.find("]")
        if j < 0:
            raise FormulaError(f"unbalanced '[' in qualifier {raw!r}")
        book = q[1:j]
        rest = q[j + 1:]
        if book == "0":
            wb_scoped = True
        else:
            external = book
    elif "[" in q and "]" in q:                       # C:\dir\[Book.xlsx]Sheet / URL
        a = q.index("[")
        b = q.index("]", a)
        external = q[:a] + q[a + 1:b]
        rest = q[b + 1:]
    elif "\\" in q or "/" in q:                       # path to a workbook: book-level name
        external = q
        rest = ""
    sheet = sheet_end = None
    if rest:
        if ":" in rest:
            sheet, sheet_end = rest.split(":", 1)
        else:
            sheet = rest
    if external is None and not wb_scoped and sheet is not None and sheet_end is None \
            and _RE_BOOK_EXT.search(sheet):
        ambiguous = True
    return Qualifier(raw, sheet, sheet_end, external, wb_scoped, ambiguous)


@dataclass(frozen=True, slots=True)
class RefPart:
    """One endpoint of a reference body: 'cell' | 'col' | 'row' | 'name' | 'r1c1' | 'error'."""
    kind: str
    text: str
    row: Optional[int] = None
    col: Optional[int] = None
    row_abs: bool = False
    col_abs: bool = False


def _classify_part(p: str) -> RefPart:
    m = _RE_CELL.match(p)
    if m:
        c, r = col_index(m.group(2)), int(m.group(4))
        if c <= MAX_COL and 1 <= r <= MAX_ROW:
            return RefPart("cell", p, r, c, bool(m.group(3)), bool(m.group(1)))
    m = _RE_COL.match(p)
    if m and col_index(m.group(2)) <= MAX_COL:
        return RefPart("col", p, None, col_index(m.group(2)), False, bool(m.group(1)))
    m = _RE_ROW.match(p)
    if m and 1 <= int(m.group(2)) <= MAX_ROW:
        return RefPart("row", p, int(m.group(2)), None, bool(m.group(1)), False)
    if p.upper() == "#REF!":
        return RefPart("error", p)
    if _RE_R1C1.match(p):
        return RefPart("r1c1", p)
    if _RE_NAME.match(p):
        return RefPart("name", p)
    return RefPart("invalid", p)


@dataclass(frozen=True, slots=True)
class _RefInfo:
    kind: str                       # base kind before context (spill/trim/param) is applied
    qualifier: Optional[Qualifier]
    body: str
    parts: tuple
    shape: Optional[str]            # cell | range | column | row | sheet | None
    bounds: Optional[tuple]
    trim: Optional[str]
    name: Optional[str]
    table: Optional[str]
    r1c1: bool


@lru_cache(maxsize=65536)
def _ref_info(raw: str) -> _RefInfo:
    k = _outside_split_index(raw, "!", last=False)
    qual = None
    body = raw
    if k >= 0:
        qual = parse_qualifier(raw[:k])
        body = raw[k + 1:]
    if not body:
        raise FormulaError(f"reference {raw!r} has a qualifier but no address")
    if body.upper() == "#REF!":
        return _RefInfo("error", qual, body, (RefPart("error", body),), None, None, None, None, None, False)
    if "[" in body:                     # qualifier ([1], [Book.xlsx]) already split off
        b = body.find("[")
        table = body[:b] or None
        return _RefInfo("structured", qual, body, (), None, None, None, None, table, False)
    raw_parts = []
    j = 0
    while True:
        idx = _outside_split_index(body[j:], ":", last=False)
        if idx < 0:
            raw_parts.append(body[j:])
            break
        raw_parts.append(body[j:j + idx])
        j += idx + 1
    trim = None
    parts = []
    for pi, p in enumerate(raw_parts):
        # a later part may carry its own qualifier: Sheet1!A1:Sheet1!B5
        kk = _outside_split_index(p, "!", last=True)
        if kk >= 0:
            p = p[kk + 1:]
        lead_dot = p.startswith(".") and pi > 0
        trail_dot = p.endswith(".") and pi < len(raw_parts) - 1
        core = p[1 if lead_dot else 0: len(p) - (1 if trail_dot else 0)]
        rp = _classify_part(core)
        if (lead_dot or trail_dot) and rp.kind in ("cell", "col", "row"):
            t = "trailing" if lead_dot else "leading"
            trim = t if trim in (None, t) else "all"
            parts.append(rp)
        elif lead_dot or trail_dot:
            parts.append(_classify_part(p))            # a name such as 'A.' -- not trim syntax
        else:
            parts.append(rp)
    if len(parts) > 1:
        parts = _pair_areas(parts, raw)
    parts = tuple(parts)
    kinds = {p.kind for p in parts}
    if "invalid" in kinds:
        bad = next(p.text for p in parts if p.kind == "invalid")
        raise FormulaError(f"unrecognised reference {bad!r} in operand {raw!r}")
    if len(parts) == 1:
        p = parts[0]
        if p.kind == "cell":
            return _RefInfo("cell", qual, body, parts, "cell", (p.row, p.col, p.row, p.col), None, None, None, False)
        if p.kind == "r1c1":
            return _RefInfo("cell", qual, body, parts, None, None, None, None, None, True)
        if p.kind == "error":
            return _RefInfo("error", qual, body, parts, None, None, None, None, None, False)
        if p.kind == "name":
            return _RefInfo("name", qual, body, parts, None, None, None, p.text, None, False)
        if p.kind == "col" and "$" not in p.text:
            if _RE_R1C1.match(p.text):   # R, C, RC: R1C1 notation, never a name
                return _RefInfo("cell", qual, body, (RefPart("r1c1", p.text),), None, None, None, None, None, True)
            # a bare 1-3 letter word (Tax, A) is a valid name
            return _RefInfo("name", qual, body, (RefPart("name", p.text),), None, None, None, p.text, None, False)
        raise FormulaError(f"unrecognised reference {raw!r}")
    shape = None
    bounds = None
    if kinds <= {"cell", "col", "row"}:
        # bounding box of every area (A:A:B5 -> A1:B1048576, i.e. whole columns A:B)
        r1 = min(1 if p.kind == "col" else p.row for p in parts)
        r2 = max(MAX_ROW if p.kind == "col" else p.row for p in parts)
        c1 = min(1 if p.kind == "row" else p.col for p in parts)
        c2 = max(MAX_COL if p.kind == "row" else p.col for p in parts)
        bounds = (r1, c1, r2, c2)
        full_rows = r1 == 1 and r2 == MAX_ROW
        full_cols = c1 == 1 and c2 == MAX_COL
        shape = "sheet" if (full_rows and full_cols) else "column" if full_rows else "row" if full_cols else "range"
    if shape is None:
        kind = "range"               # Name1:Name2, A1:Name, A1:#REF! ...
    elif trim:
        kind = "trimmed_range"
    elif shape == "column":
        kind = "whole_column"
    elif shape == "row":
        kind = "whole_row"
    elif shape == "sheet":           # A:XFD -> whole_column; 1:1048576, A1:XFD1048576 -> whole_row
        kind = "whole_column" if "col" in kinds else "whole_row"
    else:
        kind = "range"
    return _RefInfo(kind, qual, body, parts, shape, bounds, trim, None, None, False)


def _pair_areas(parts: list, raw: str) -> list:
    """Excel's areas are a cell, COL:COL or ROW:ROW.  A column-shaped endpoint that is not
    paired with another column (Tax in Tax:Rate, B in A1:B) is a defined name; a row
    number that is not paired with another row (A1:1) cannot be stored.  Pairs are taken
    left to right (A:A:B5 = column area A:A, then cell B5)."""
    out = list(parts)
    i, n = 0, len(out)
    while i < n:
        p = out[i]
        nxt = out[i + 1] if i + 1 < n else None
        if p.kind in ("col", "row") and nxt is not None and nxt.kind == p.kind:
            i += 2
            continue
        if p.kind == "col":
            if _RE_R1C1.match(p.text):
                out[i] = RefPart("r1c1", p.text)
            elif _RE_NAME.match(p.text):
                out[i] = RefPart("name", p.text)
            else:                    # '$A' alone
                raise FormulaError(f"column {p.text!r} is not paired with another column in {raw!r}")
        elif p.kind == "row":
            raise FormulaError(f"row {p.text!r} is not paired with another row in {raw!r}")
        i += 1
    return out


@dataclass(slots=True, unsafe_hash=True)
class Operand:
    """One operand of a formula.

    kind: cell | range | whole_column | whole_row | trimmed_range | spill | structured |
          name | param (LET/LAMBDA local) | error | number | string | bool | array
    """
    kind: str
    raw: str                          # text as written (incl. qualifier, '#' of a spill not included)
    start: int                        # offsets into Formula.body
    end: int
    sheet: Optional[str] = None       # unquoted sheet name; None = no sheet qualifier
    sheet_end: Optional[str] = None   # 3-D span end sheet
    external: Optional[str] = None    # book index '1' / file name / path; None = this workbook
    ambiguous_book: bool = False      # qualifier like Other.xlsx: local sheet or another book
    workbook_scoped: bool = False     # [0]!Name
    qualifier: Optional[str] = None   # raw qualifier text (before '!')
    body: Optional[str] = None        # text after the qualifier
    name: Optional[str] = None        # kind name/param: the identifier (param without _xlpm.)
    table: Optional[str] = None       # kind structured: table name (None for [@Col])
    shape: Optional[str] = None       # cell | range | column | row | sheet (geometry, any kind)
    bounds: Optional[tuple] = None    # (r1, c1, r2, c2), 1-based, whole column -> rows 1..MAX
    trim: Optional[str] = None        # trailing | leading | all | trimrange
    implicit: bool = False            # '@' implicit intersection (or _xlfn.SINGLE argument)
    spill: bool = False               # A1# or _xlfn.ANCHORARRAY(A1)
    r1c1: bool = False                # R2C3-style token (not a name)
    func_path: tuple = ()             # enclosing function names, outermost first
    arg_index: Optional[int] = None   # argument position in the innermost function
    call_index: Optional[int] = None  # index into Formula.functions of the innermost call
    parts: tuple = ()                 # RefPart endpoints of the reference body
    value: object = None              # literal value (number/string/bool/error/array rows)

    @property
    def is_reference(self) -> bool:
        return self.kind in REFERENCE_KINDS

    @property
    def is_literal(self) -> bool:
        return self.kind in LITERAL_KINDS

    @property
    def names(self) -> tuple:
        """Defined-name identifiers this operand mentions (Rate, or both ends of N1:N2)."""
        if self.kind == "name":
            return (self.name,)
        return tuple(p.text for p in self.parts if p.kind == "name") if self.kind == "range" else ()

    @property
    def reaches_sheet_edge(self) -> bool:
        """A bounded range that runs to row 1048576 or column XFD (A2:A1048576)."""
        b = self.bounds
        return bool(b) and (b[2] == MAX_ROW or b[3] == MAX_COL)


@dataclass(slots=True, unsafe_hash=True)
class FunctionCall:
    name: str                 # normalised: upper case, prefixes stripped (_xlfn._xlws.SORT -> SORT)
    raw: str                  # as written, without '(' (e.g. '_xlfn._xlws.SORT')
    start: int                # offset of the name in Formula.body
    end: int                  # offset after the closing ')'
    depth: int                # 0 = not nested inside another call
    parent: Optional[int]     # index of the enclosing call in Formula.functions
    parent_arg: Optional[int] # argument position inside the parent
    arg_spans: tuple          # ((start, end), ...) offsets of each argument's text
    prefix: str = ""          # '_xlfn.' / '_xlfn._xlws.' / '_xludf.' ...
    qualifier: Optional[str] = None   # Sheet1!MyFn(...) -> 'Sheet1'
    builtin: bool = True      # an Excel function (not a defined-name / LET-local call)
    local: bool = False       # call of a LET/LAMBDA parameter (_xlpm.f(...) or f bound by LET)
    eta: bool = False         # _xleta.SUM passed as a value (BYROW(r, _xleta.SUM))
    implicit: bool = False    # @FUNC(...)
    spilled: bool = False     # FUNC(...)#

    @property
    def n_args(self) -> int:
        return len(self.arg_spans)


class Formula:
    """Parse result.  Immutable by convention (instances are cached and shared)."""

    __slots__ = ("text", "body", "functions", "operands", "strings", "params", "_tokens")

    def __init__(self, text, body, functions, operands, strings, params, tokens):
        self.text = text
        self.body = body
        self.functions = functions
        self.operands = operands
        self.strings = strings
        self.params = params
        self._tokens = tokens

    def __repr__(self):
        return f"Formula({self.body!r})"

    @property
    def tokens(self) -> tuple:
        """(kind, text, start, end) tuples of the lexer output (re-lexed on demand; not
        kept in the cache to save memory)."""
        return tuple(lex(self.body))

    # -- functions --------------------------------------------------------------------
    def function_names(self) -> set:
        return {c.name for c in self.functions if not c.local}

    def calls(self, *names: str) -> list:
        want = {n.upper() for n in names}
        return [c for c in self.functions if c.name in want and not c.local]

    def arg_text(self, call: "FunctionCall | int", i: int) -> str:
        c = self.functions[call] if isinstance(call, int) else call
        a, b = c.arg_spans[i]
        return self.body[a:b].strip(_WS_CHARS)

    def operands_in_arg(self, call: "FunctionCall | int", i: int) -> list:
        c = self.functions[call] if isinstance(call, int) else call
        a, b = c.arg_spans[i]
        return [o for o in self.operands if a <= o.start and o.end <= b]

    # -- operands ---------------------------------------------------------------------
    def references(self) -> list:
        return [o for o in self.operands if o.kind in REFERENCE_KINDS]

    def names_referenced(self) -> list:
        """Syntactic name uses: NameRef for every name operand / range endpoint and every
        call of a non-built-in function (possible LAMBDA name).  LET/LAMBDA locals,
        string literals and external names are excluded (external ones flagged)."""
        return _names_referenced(self)

    def ref_errors(self) -> list:
        """Operands carrying #REF!: a bare #REF!, Sheet!#REF!, or a range endpoint
        (A1:#REF!).  Text inside string literals never counts."""
        out = []
        for o in self.operands:
            if o.kind == "error" and (o.body or o.raw).upper().endswith("#REF!"):
                out.append(o)
            elif any(p.kind == "error" for p in o.parts):
                out.append(o)
        return out

    def whole_column_or_row_refs(self, include_sheet_edge: bool = False) -> list:
        out = [o for o in self.operands if o.kind in ("whole_column", "whole_row")]
        if include_sheet_edge:
            out += [o for o in self.operands if o.kind == "range" and o.reaches_sheet_edge]
        return out

    def references_other_sheet(self, own_sheet: Optional[str], sheet_names=None, table_sheets=None) -> bool:
        return bool(self.other_sheets(own_sheet, sheet_names, table_sheets))

    def other_sheets(self, own_sheet: Optional[str], sheet_names=None, table_sheets=None) -> set:
        """Sheet names (as written, unquoted) of this workbook referenced directly, other
        than own_sheet (case-insensitive).  3-D spans contribute both ends.  A table
        reference (Tbl[Col]) needs table_sheets={table: sheet}, else FormulaError."""
        return _other_sheets(self.operands, own_sheet, sheet_names, table_sheets)

    def references_external_workbook(self, sheet_names=None) -> bool:
        return any(_is_external(o, sheet_names) for o in self.operands)

    def external_references(self, sheet_names=None) -> list:
        return [o for o in self.operands if _is_external(o, sheet_names)]

    def pure_reference(self, transparent=("ANCHORARRAY", "SINGLE")) -> Optional[Operand]:
        """The single reference/name operand when the whole formula is just that operand,
        optionally wrapped in parentheses, unary +/-, '@', a spill '#', or calls listed in
        `transparent` (each with that operand as its only argument).  Else None."""
        sig = [t for t in self.tokens if t[0] != "ws"]
        ops = [o for o in self.operands]
        if len(ops) != 1 or ops[0].kind not in REFERENCE_KINDS | {"name"}:
            return None
        allowed_funcs = {t.upper() for t in transparent}
        for kind, ttext, tstart, _ in sig:
            if kind in ("(", ")", "spill"):
                continue
            if kind == "op" and ttext in ("+", "-", "@"):
                continue
            if kind == "func":
                if _norm_func(ttext)[0] in allowed_funcs:
                    continue
                return None
            if kind == "operand" and tstart == ops[0].start:
                continue
            return None
        for c in self.functions:
            if c.n_args != 1:
                return None
        return ops[0]


@dataclass(frozen=True, slots=True)
class NameRef:
    name: str                       # identifier as written
    sheet: Optional[str]            # qualifier sheet (Sheet1!Rate) or None
    external: Optional[str]         # external book of [1]!Rate, else None
    workbook_scoped: bool           # [0]!Rate
    as_call: bool                   # MyFn(...) call of a (possibly LAMBDA) name
    start: int
    ambiguous_book: bool = False    # Other.xlsx!Name: another workbook or a sheet 'Other.xlsx'


@lru_cache(maxsize=4096)
def _norm_func(raw: str):
    """'_xlfn._xlws.SORT' -> ('SORT', '_xlfn._xlws.', local?)"""
    name = raw
    prefix = ""
    low = name.lower()
    changed = True
    while changed:
        changed = False
        for p in _FUNC_PREFIXES:
            if low.startswith(p):
                prefix += name[:len(p)]
                name = name[len(p):]
                low = low[len(p):]
                changed = True
    local = False
    if low.startswith(_PARAM_PREFIX):
        prefix += name[:len(_PARAM_PREFIX)]
        name = name[len(_PARAM_PREFIX):]
        local = True
    return name.upper(), prefix, local


def _parse_array(toks, i, body):
    """toks[i] is '{'; returns (rows, end_index_of_close, strings)."""
    rows = [[]]
    strings = []
    sign = ""
    j = i + 1
    while j < len(toks):
        k, tt, _, _ = toks[j]
        if k == "}":
            return rows, j, strings
        if k == "ws":
            pass
        elif k == ",":
            pass
        elif k == ";":
            rows.append([])
        elif k == "op" and tt in ("+", "-"):
            sign = tt if tt == "-" else sign
        elif k == "num":
            v = float(tt)
            rows[-1].append(-v if sign == "-" else v)
            sign = ""
        elif k == "str":
            sv = tt[1:-1].replace('""', '"')
            rows[-1].append(sv)
            strings.append(sv)
        elif k == "err":
            rows[-1].append(tt.upper())
        elif k == "operand" and tt.upper() in ("TRUE", "FALSE"):
            rows[-1].append(tt.upper() == "TRUE")
        else:
            raise FormulaError(f"unexpected {tt!r} inside array constant in {body!r}")
        j += 1
    raise FormulaError(f"unterminated array constant in {body!r}")


@lru_cache(maxsize=4096)
def parse(text: str) -> Formula:
    """Parse formula text (as stored in the file, no leading '=').  Cached.
    Raises FormulaError on text Excel could not have stored."""
    if text is None:
        raise FormulaError("formula text is None")
    body = text[1:] if text.startswith("=") else text
    toks = lex(body)
    _check_adjacent(toks, body)
    functions: list = []          # mutable dicts, frozen at the end
    operands: list = []
    strings: list = []
    frames: list = []             # all open frames: [type, call_index, arg_index, arg_start]
    fstack: list = []             # the open 'func' frames only (same list objects)
    path: list = []               # names of the open calls, outermost first
    path_t: tuple = ()
    last_closed_call = None       # call index closed by the most recent ')'
    pending_implicit = False
    n = len(toks)
    i = 0
    while i < n:
        k, tt, ts, te = toks[i]
        if k == "ws":
            i += 1
            continue
        inner = fstack[-1] if fstack else None
        if k == "func":
            raw = tt
            qual = None
            if "!" in raw:
                kq = _outside_split_index(raw, "!", last=True)
                if kq >= 0:
                    qual, raw = raw[:kq], raw[kq + 1:]
            if not _RE_NAME.match(raw):
                raise FormulaError(f"invalid function name {tt!r} in {body!r}")
            name, prefix, local = _norm_func(raw)
            functions.append({
                "name": name, "raw": raw, "start": ts, "end": None,
                "depth": len(fstack), "parent": inner[1] if inner else None,
                "parent_arg": inner[2] if inner else None, "args": [], "prefix": prefix,
                "qualifier": qual, "local": local,
                "eta": False, "implicit": pending_implicit, "spilled": False,
            })
            pending_implicit = False
            fr = ["func", len(functions) - 1, 0, toks[i + 1][3]]   # toks[i+1] is '('
            frames.append(fr)
            fstack.append(fr)
            path.append(name)
            path_t = tuple(path)
            last_closed_call = None
            i += 2
            continue
        if k == "(":
            frames.append(["paren", None, 0, te])
            pending_implicit = False
            last_closed_call = None
            i += 1
            continue
        if k == ")":
            if not frames:
                raise FormulaError(f"unbalanced ')' at position {ts} in {body!r}")
            fr = frames.pop()
            last_closed_call = None
            if fr[0] == "func":
                fstack.pop()
                path.pop()
                path_t = tuple(path)
                fn = functions[fr[1]]
                fn["args"].append((fr[3], ts))
                fn["end"] = te
                if len(fn["args"]) == 1 and not body[fr[3]:ts].strip(_WS_CHARS):
                    fn["args"] = []          # f() has no arguments
                last_closed_call = fr[1]
            i += 1
            continue
        if k == "spill":
            if last_closed_call is not None:
                functions[last_closed_call]["spilled"] = True
            elif operands and toks[i - 1][0] == "operand":
                o = operands[-1]
                operands[-1] = _replace(o, spill=True, kind="spill" if o.kind == "cell" else o.kind)
            last_closed_call = None
            i += 1
            continue
        last_closed_call = None
        if k == ",":
            if frames and frames[-1][0] == "func":
                fr = frames[-1]
                functions[fr[1]]["args"].append((fr[3], ts))
                fr[2] += 1
                fr[3] = te
            # else: union operator (A1,B1) inside parentheses
            i += 1
            continue
        if k == ";":
            raise FormulaError(f"';' outside an array constant at position {ts} in {body!r}")
        if k == "}":
            raise FormulaError(f"unbalanced '}}' at position {ts} in {body!r}")
        if k == "op":
            if tt == "@":
                pending_implicit = True
            i += 1
            continue
        ctx_path = path_t
        arg_index = inner[2] if inner else None
        call_index = inner[1] if inner else None
        if k == "{":
            rows, j, strs = _parse_array(toks, i, body)
            strings.extend(strs)
            operands.append(Operand("array", body[ts:toks[j][3]], ts, toks[j][3],
                                    func_path=ctx_path, arg_index=arg_index, call_index=call_index,
                                    value=tuple(tuple(r) for r in rows)))
            pending_implicit = False
            i = j + 1
            continue
        if k == "str":
            sv = tt[1:-1].replace('""', '"')
            strings.append(sv)
            operands.append(Operand("string", tt, ts, te, func_path=ctx_path,
                                    arg_index=arg_index, call_index=call_index, value=sv))
        elif k == "num":
            operands.append(Operand("number", tt, ts, te, func_path=ctx_path,
                                    arg_index=arg_index, call_index=call_index, value=float(tt)))
        elif k == "err":
            operands.append(Operand("error", tt, ts, te, func_path=ctx_path,
                                    arg_index=arg_index, call_index=call_index, value=tt.upper()))
        elif k == "operand":
            word = tt
            up = word.upper()
            if up in ("TRUE", "FALSE"):
                operands.append(Operand("bool", word, ts, te, func_path=ctx_path,
                                        arg_index=arg_index, call_index=call_index, value=(up == "TRUE")))
            elif word[:7].lower() == "_xleta.":
                name, prefix, _ = _norm_func(word)
                functions.append({
                    "name": name, "raw": word, "start": ts, "end": te,
                    "depth": len(fstack), "parent": call_index, "parent_arg": arg_index,
                    "args": [], "prefix": prefix, "qualifier": None, "local": False,
                    "eta": True, "implicit": False, "spilled": False,
                })
            else:
                info = _ref_info(word)
                q = info.qualifier
                kind = info.kind
                nm = info.name
                if kind == "name" and q is None and nm[0] == "_":
                    low = nm.lower()
                    if low.startswith(_PARAM_PREFIX):
                        kind, nm = "param", nm[6:]
                    elif low.startswith("_xlfn." + _PARAM_PREFIX):
                        kind, nm = "param", nm[12:]
                if q is None:
                    operands.append(Operand(
                        kind, word, ts, te, body=info.body, name=nm, table=info.table,
                        shape=info.shape, bounds=info.bounds, trim=info.trim, implicit=pending_implicit,
                        r1c1=info.r1c1, func_path=ctx_path, arg_index=arg_index,
                        call_index=call_index, parts=info.parts))
                else:
                    operands.append(Operand(
                        kind, word, ts, te, sheet=q.sheet, sheet_end=q.sheet_end,
                        external=q.external, ambiguous_book=q.ambiguous_book and kind in ("name", "structured"),
                        workbook_scoped=q.workbook_scoped, qualifier=q.raw, body=info.body,
                        name=nm, table=info.table, shape=info.shape, bounds=info.bounds,
                        trim=info.trim, implicit=pending_implicit, r1c1=info.r1c1,
                        func_path=ctx_path, arg_index=arg_index, call_index=call_index,
                        parts=info.parts))
        else:  # pragma: no cover - lexer kinds are exhaustive
            raise FormulaError(f"unexpected token {tt!r} in {body!r}")
        pending_implicit = False
        i += 1
    if frames:
        raise FormulaError(f"unbalanced '(' in {body!r}")

    # ---- context pass: LET/LAMBDA parameters ----------------------------------------
    # Excel binds a LET name only AFTER its value (LET(wacc,WACC,...): the second WACC is
    # the defined name) and a LAMBDA parameter only in the LAMBDA's body.  A parameter
    # declared as _xlpm.x is referenced as _xlpm.x everywhere (that is how Excel stores
    # it), so a plain x inside its scope is the defined name / built-in, not the local.
    # Only a plain declaration (LET(x,1,x+1), as some libraries write it) binds plain x.
    scopes = []                       # (name_lower, scope_start, scope_end) of plain declarations
    declared = set()                  # lower-case names of every declaration
    decl_starts = set()               # offsets of plain declaration operands
    by_arg = None
    for fi, fn in enumerate(functions):
        if fn["name"] not in ("LET", "LAMBDA") or fn["local"] or fn["eta"] or not fn["args"]:
            continue
        args = fn["args"]
        is_let = fn["name"] == "LET"
        decl = range(0, len(args) - 1, 2) if is_let else range(len(args) - 1)
        if by_arg is None:
            by_arg = {}
            for o in operands:
                by_arg.setdefault((o.call_index, o.arg_index), []).append(o)
            for f2 in functions:            # a call inside an argument makes it not a bare name
                if f2["parent"] is not None:
                    by_arg.setdefault((f2["parent"], f2["parent_arg"]), []).append(None)
        for a in decl:
            inarg = by_arg.get((fi, a), ())
            if len(inarg) == 1 and inarg[0] is not None and inarg[0].kind in ("name", "param") \
                    and inarg[0].qualifier is None:
                d = inarg[0]
                low = d.name.lower()
                declared.add(low)
                if d.kind == "param":
                    continue                # _xlpm.x: its uses carry the prefix themselves
                decl_starts.add(d.start)
                # LET: in scope from the end of its value argument; LAMBDA: the body only
                scope_start = args[a + 1][1] if is_let else args[-1][0]
                scopes.append((low, scope_start, fn["end"]))
    if decl_starts:
        for oi, o in enumerate(operands):
            if o.kind == "name" and o.qualifier is None:
                low = o.name.lower()
                if o.start in decl_starts or any(s == low and a <= o.start < b for s, a, b in scopes):
                    operands[oi] = _replace(o, kind="param")
        for fn in functions:
            if not fn["local"] and fn["qualifier"] is None and not fn["prefix"]:
                low = fn["name"].lower()
                if any(s == low and a <= fn["start"] < b for s, a, b in scopes):
                    fn["local"] = True
    # ---- context pass: trimmed / spilled / implicit wrappers ------------------------
    for oi, o in enumerate(operands):
        if o.call_index is None or o.arg_index != 0:
            continue
        fn = functions[o.call_index]
        nm = fn["name"]
        if fn["local"]:
            continue
        if nm in TRIM_FUNCTIONS:
            if o.kind in ("cell", "range", "whole_column", "whole_row"):
                operands[oi] = _replace(o, kind="trimmed_range", trim=TRIM_FUNCTIONS[nm])
        elif nm == "ANCHORARRAY":
            if len(fn["args"]) == 1 and o.kind in ("cell", "name", "spill") and \
                    sum(1 for x in operands if x.call_index == o.call_index) == 1:
                operands[oi] = _replace(o, kind="spill" if o.kind == "cell" else o.kind, spill=True)
        elif nm == "SINGLE":
            if len(fn["args"]) == 1:
                operands[oi] = _replace(o, implicit=True)

    calls = []
    for fn in functions:
        pre = fn["prefix"].lower()
        if fn["local"]:
            builtin = False
        elif "_xludf." in pre or "_xll." in pre:
            builtin = False             # user-defined / add-in function
        elif fn["eta"] or "_xlfn." in pre or "_xlws." in pre:
            builtin = True
        else:
            builtin = fn["name"] in BUILTIN_FUNCTIONS and fn["qualifier"] is None
        calls.append(FunctionCall(fn["name"], fn["raw"], fn["start"], fn["end"], fn["depth"], fn["parent"],
                                  fn["parent_arg"], tuple(fn["args"]), fn["prefix"],
                                  _unquote(fn["qualifier"]) if fn["qualifier"] is not None else None,
                                  builtin, fn["local"], fn["eta"], fn["implicit"], fn["spilled"]))
    params = frozenset(declared | {o.name.lower() for o in operands if o.kind == "param"})
    return Formula(text, body, tuple(calls), tuple(operands), tuple(strings), params, None)


# Token kinds that end / start an operand.  Two of them side by side with nothing in
# between ('#REF!A1', '"a"B1', '1"x"', '(A1)B1') is text Excel's file parser rejects
# (toy 22/T4: '#REF!A1' is refused, '#REF!' and 'Sheet!#REF!' are accepted).  The one
# legal pair is ')(' : an immediately invoked LAMBDA(...)(...).
_ATOM_END = frozenset({"operand", "num", "str", "err", "spill", "}", ")"})
_ATOM_START = frozenset({"operand", "num", "str", "err", "func", "{", "("})


def _check_adjacent(toks, body):
    for j in range(1, len(toks)):
        pk = toks[j - 1][0]
        k = toks[j][0]
        if pk in _ATOM_END and k in _ATOM_START and not (pk == ")" and k == "("):
            raise FormulaError(f"missing operator between {toks[j - 1][1]!r} and {toks[j][1]!r} "
                               f"at position {toks[j][2]} in {body!r}")


def _replace(o: Operand, **kw) -> Operand:
    d = {f: getattr(o, f) for f in Operand.__dataclass_fields__}
    d.update(kw)
    return Operand(**d)


# --------------------------------------------------------------------------- helpers
def _is_external(o: Operand, sheet_names=None) -> bool:
    if o.external is not None:
        return True
    if o.ambiguous_book:
        if sheet_names is None:
            raise AmbiguousReferenceError(
                f"qualifier {o.qualifier!r} in {o.raw!r} may be another workbook or a sheet of this "
                f"workbook; pass sheet_names= to decide")
        return o.sheet.lower() not in {s.lower() for s in sheet_names}
    return False


_SHEET_REF_KINDS = REFERENCE_KINDS | {"error"}


def _table_map(table_sheets) -> Optional[dict]:
    if table_sheets is None:
        return None
    items = table_sheets.items() if hasattr(table_sheets, "items") else table_sheets
    out = {}
    for t, s in items:
        if not isinstance(t, str) or not isinstance(s, str) or not t or not s:
            raise FormulaError(f"table_sheets entries must be (table name, sheet name) strings, got {(t, s)!r}")
        out[t.lower()] = s
    return out


def _other_sheets(operands, own_sheet, sheet_names=None, table_sheets=None) -> set:
    """Sheets of THIS workbook named by reference operands (cells, ranges, spills, tables,
    Sheet!#REF!), other than own_sheet.  A qualified NAME (Sheet1!Rate) is not counted:
    where a name points is only known after resolving it (use expand / names=).

    Structured references: Excel never writes a sheet before a table name, so the sheet
    of Tbl[Col] comes from table_sheets ({table name: sheet name}, case-insensitive).
    A table reference without that map, or to a table missing from it, raises
    FormulaError (no guess).  [@Col] / [Col] without a table name is the formula's own
    table, i.e. its own sheet."""
    own = own_sheet.lower() if own_sheet else None
    tmap = _table_map(table_sheets)
    out = set()
    for o in operands:
        if o.external is not None or o.kind not in _SHEET_REF_KINDS:
            continue
        if o.ambiguous_book and _is_external(o, sheet_names):
            continue
        if o.kind == "structured" and o.table is not None:
            if tmap is None:
                raise FormulaError(f"structured reference {o.raw!r}: the sheet of table {o.table!r} is "
                                   f"unknown; pass table_sheets= ({{table name: sheet}}) to decide")
            hit = tmap.get(o.table.lower())
            if hit is None and o.sheet is None:
                raise FormulaError(f"structured reference {o.raw!r}: table {o.table!r} is not in table_sheets")
            sheets = (hit,) if hit is not None else (o.sheet, o.sheet_end)
        elif o.sheet is None:
            continue
        else:
            sheets = (o.sheet, o.sheet_end)
        for s in sheets:
            if s is not None and (own is None or s.lower() != own):
                out.add(s)
    return out


def _names_referenced(f: Formula) -> list:
    out = []
    for o in f.operands:
        if o.kind == "name":
            out.append(NameRef(o.name, o.sheet, o.external, o.workbook_scoped, False, o.start, o.ambiguous_book))
        elif o.kind == "range":
            for p in o.parts:
                if p.kind == "name":
                    out.append(NameRef(p.text, o.sheet, o.external, o.workbook_scoped, False, o.start))
    for c in f.functions:
        if not c.builtin and not c.local and not c.eta:
            q = parse_qualifier(c.qualifier) if c.qualifier else None
            out.append(NameRef(c.raw[len(c.prefix):],
                               q.sheet if q else None, q.external if q else None,
                               q.workbook_scoped if q else False, True, c.start,
                               bool(q and q.ambiguous_book)))
    out.sort(key=lambda r: r.start)
    return out


def _as_formula(formula) -> Formula:
    return formula if isinstance(formula, Formula) else parse(formula)


def function_calls(formula) -> list:
    """FunctionCall list (built-ins and name calls; LET-local calls flagged .local)."""
    return list(_as_formula(formula).functions)


def names_referenced(formula) -> list:
    return _as_formula(formula).names_referenced()


def whole_column_or_row_refs(formula, names=None, scope_sheet=None, include_sheet_edge=False) -> list:
    """Whole-column / whole-row reference operands (A:A, $17:$17, Sheet!A:C, A1:A1048576).
    Trimmed ranges (A:.A, _TRO_TRAILING(A:A), TRIMRANGE(A:A)), structured references and
    text inside string literals never count.  With `names` (NameTable or list), names the
    formula uses are expanded transitively and their whole references are included."""
    f = _as_formula(formula)
    if names is None:
        return f.whole_column_or_row_refs(include_sheet_edge)
    r = expand(f, scope_sheet, names)
    return [o for _, o in r.operands if o.kind in ("whole_column", "whole_row")
            or (include_sheet_edge and o.kind == "range" and o.reaches_sheet_edge)]


def references_other_sheet(formula, own_sheet, names=None, scope_sheet=None, sheet_names=None,
                           table_sheets=None) -> bool:
    """True when the formula reads a cell of another sheet of this workbook (sheet name
    compared case-insensitively with own_sheet; a self-qualified 'Own'!A1 is NOT another
    sheet).  With `names`, references inside the names it uses count too (unqualified
    references inside a name are relative to the using sheet).  table_sheets
    ({table name: sheet}) places structured references; a table reference without it
    raises FormulaError."""
    f = _as_formula(formula)
    if names is None:
        return f.references_other_sheet(own_sheet, sheet_names, table_sheets)
    r = expand(f, scope_sheet if scope_sheet is not None else own_sheet, names, sheet_names=sheet_names)
    return bool(_other_sheets([o for _, o in r.operands], own_sheet, sheet_names, table_sheets))


def references_external_workbook(formula, names=None, scope_sheet=None, sheet_names=None) -> bool:
    """True when the formula (or, with `names`, a name it uses) references another
    workbook: [1]Sheet!A1, '[1]Sheet 1'!A1, [1]!Name, [Book.xlsx]S!A1, 'C:\\p\\[B.xlsx]S'!A1,
    'C:\\p\\Book.xlsx'!Name.  [0]!Name is this workbook."""
    f = _as_formula(formula)
    if names is None:
        return f.references_external_workbook(sheet_names)
    r = expand(f, scope_sheet, names, sheet_names=sheet_names)
    return any(_is_external(o, sheet_names) for _, o in r.operands)


# --------------------------------------------------------------------------- defined names
@dataclass(frozen=True, slots=True)
class DefinedName:
    name: str
    scope: Optional[str]          # sheet name for a sheet-local name, None = workbook
    text: str                     # definition (refersTo) without leading '='
    hidden: bool = False
    index: int = 0                # position in the input list

    @property
    def builtin(self) -> bool:
        return self.name.lower().startswith("_xlnm.")

    @property
    def validity_problem(self) -> Optional[str]:
        return name_validity_problem(self.name)

    def formula(self) -> Formula:
        return parse(self.text)

    @property
    def is_lambda(self) -> bool:
        """The definition is a LAMBDA (callable as MyFn(...)): its outermost expression is
        one LAMBDA(...) call (an immediately invoked LAMBDA(...)(x) is not callable)."""
        f = parse(self.text)
        top = [c for c in f.functions if c.depth == 0]
        if len(top) != 1 or top[0].name != "LAMBDA" or top[0].local:
            return False
        rest = (f.body[:top[0].start] + f.body[top[0].end:]).strip(_WS_CHARS + "()")
        return rest == ""


def _xml_bool(v, what: str) -> bool:
    """True/False from a bool, 0/1, None (attribute absent) or an xsd:boolean string."""
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, int) and v in (0, 1):
        return bool(v)
    if isinstance(v, str) and v.strip().lower() in ("0", "false"):
        return False
    if isinstance(v, str) and v.strip().lower() in ("1", "true"):
        return True
    raise FormulaError(f"{what}: {v!r} is not a boolean")


def _coerce_defined_name(n, i: int) -> DefinedName:
    """A NameTable entry -> DefinedName, validated.  Accepted: this module's DefinedName,
    the reader's detchecks.core.package.DefinedName (or any object with name / scope /
    text / hidden attributes), or a tuple (name, scope_sheet_or_None, text[, hidden])."""
    if isinstance(n, DefinedName):
        name, scope, text, hidden, idx = n.name, n.scope, n.text, n.hidden, n.index
    elif isinstance(n, (tuple, list)):
        if len(n) not in (3, 4):
            raise FormulaError(f"NameTable entry {n!r}: expected (name, scope_sheet_or_None, text[, hidden])")
        name, scope, text = n[0], n[1], n[2]
        hidden = n[3] if len(n) == 4 else False
        idx = i
    elif all(hasattr(n, a) for a in ("name", "scope", "text", "hidden")):
        name, scope, text, hidden, idx = n.name, n.scope, n.text, n.hidden, i
    else:
        raise FormulaError(f"NameTable entry of type {type(n).__name__} is not supported: {n!r}")
    if not isinstance(name, str) or not name:
        raise FormulaError(f"NameTable entry {n!r}: the name must be a non-empty string")
    if scope is not None and (not isinstance(scope, str) or not scope):
        raise FormulaError(f"defined name {name!r}: scope must be a sheet NAME or None, got {scope!r} "
                           f"(map workbook.xml localSheetId to the sheet name first)")
    if text is None:
        text = ""
    if not isinstance(text, str):
        raise FormulaError(f"defined name {name!r}: definition must be text, got {text!r}")
    if text.startswith("="):
        text = text[1:]
    hidden = _xml_bool(hidden, f"defined name {name!r} hidden flag")
    if isinstance(n, DefinedName) and n.text == text and n.hidden is hidden:
        return n                                       # keep identity (tables rebuilt by _table)
    return DefinedName(name, scope, text, hidden, idx)


class NameTable:
    """Defined names with Excel's scope rules.

    names: iterable of entries, each a DefinedName (this module's or the reader's
    detchecks.core.package.DefinedName, e.g. pkg.defined_names) or a tuple
    (name, scope_sheet_or_None, text[, hidden]).  scope must be a sheet NAME (not a
    localSheetId); hidden may be a bool, 0/1 or an XML boolean string ("0", "true").
    Anything else raises FormulaError.
    sheet_names (optional): the workbook's sheet names; only needed to decide the rare
    qualifier 'Other.xlsx!Name' (another workbook, or a sheet named like a file).
    Lookups are case-insensitive.  A sheet-local name beats a workbook name of the same
    spelling on its own sheet; Sheet1!Rate looks up Rate local to Sheet1, then global.
    """

    def __init__(self, names: Iterable, sheet_names: Optional[Iterable[str]] = None):
        self.names: list = []
        self._by_key: dict = {}
        self.sheet_names = None if sheet_names is None else {x.lower() for x in sheet_names}
        if isinstance(names, (str, bytes)) or not hasattr(names, "__iter__"):
            raise FormulaError(f"NameTable expects an iterable of defined names, got {type(names).__name__}")
        for i, n in enumerate(names):
            n = _coerce_defined_name(n, i)
            self.names.append(n)
            key = (n.name.lower(), n.scope.lower() if n.scope else None)
            self._by_key.setdefault(key, n)

    def __iter__(self):
        return iter(self.names)

    def __len__(self):
        return len(self.names)

    def get(self, name: str, scope: Optional[str] = None) -> Optional[DefinedName]:
        return self._by_key.get((name.lower(), scope.lower() if scope else None))

    def resolve(self, ident: str, scope_sheet: Optional[str] = None,
                qualifier_sheet: Optional[str] = None, workbook_scoped: bool = False) -> Optional[DefinedName]:
        low = ident.lower()
        if low.startswith(_PARAM_PREFIX):
            return None
        if workbook_scoped:
            return self._by_key.get((low, None))
        if qualifier_sheet is not None:
            return self._by_key.get((low, qualifier_sheet.lower())) or self._by_key.get((low, None))
        if scope_sheet is not None:
            hit = self._by_key.get((low, scope_sheet.lower()))
            if hit is not None:
                return hit
        return self._by_key.get((low, None))


def _table(names, sheet_names=None) -> NameTable:
    if isinstance(names, NameTable):
        if sheet_names is not None and names.sheet_names is None:
            return NameTable(names.names, sheet_names)
        return names
    return NameTable(names, sheet_names)


def names_used_by(formula, scope_sheet: Optional[str], names, include_indirect_literals: bool = False,
                  sheet_names=None) -> list:
    """Defined names the formula references DIRECTLY (not transitively), in order of first
    use.  scope_sheet: the sheet the formula lives on (cell / CF / DV formulas), the scope
    sheet of a local name's definition, or None for a workbook-level name's definition.

    Counts: name operands (Rate, Sheet1!Rate, [0]!Rate, both ends of N1:N2, Name#) and
    calls of a defined name (MyFn(...); an Excel built-in function name called as a
    function is the built-in, unless stored as _xludf.X).  Never counts: LET/LAMBDA
    parameters, external names ([1]!Rate), string literals.  With
    include_indirect_literals=True, a literal first argument of INDIRECT("Rate") or
    HYPERLINK("#Rate") that resolves to a name also counts."""
    f = _as_formula(formula)
    tab = _table(names, sheet_names)
    out = []
    seen = set()
    for r in f.names_referenced():
        if r.external is not None:
            continue
        if r.ambiguous_book:
            if tab.sheet_names is None:
                raise AmbiguousReferenceError(
                    f"qualifier {r.sheet!r} of name {r.name!r} may be another workbook or a sheet; "
                    f"build the NameTable with sheet_names= to decide")
            if r.sheet.lower() not in tab.sheet_names:
                continue                                     # another workbook's name
        hit = tab.resolve(r.name, scope_sheet, r.sheet, r.workbook_scoped)
        if hit is not None and id(hit) not in seen:
            seen.add(id(hit))
            out.append(hit)
    if include_indirect_literals:
        for c in f.functions:
            if c.local or c.name not in ("INDIRECT", "HYPERLINK") or not c.arg_spans:
                continue
            lit = _literal_concat(f, c, 0)
            if lit is None:
                continue
            if c.name == "HYPERLINK":
                if not lit.startswith("#"):
                    continue
                lit = lit[1:]
            k = _outside_split_index(lit, "!", last=True)
            qs, ident = (lit[:k], lit[k + 1:]) if k >= 0 else (None, lit)
            ident = ident.strip()
            if not ident or not is_valid_name(ident):
                continue
            qsheet = parse_qualifier(qs).sheet if qs else None
            hit = tab.resolve(ident, scope_sheet, qsheet)
            if hit is not None and id(hit) not in seen:
                seen.add(id(hit))
                out.append(hit)
    return out


def _literal_concat(f: Formula, c: FunctionCall, i: int) -> Optional[str]:
    """Value of argument i when it is string literals joined by '&', else None."""
    a, b = c.arg_spans[i]
    toks = [t for t in f.tokens if a <= t[2] and t[3] <= b and t[0] != "ws"]
    if not toks:
        return None
    out = []
    expect = True
    for kind, ttext, _, _ in toks:
        if expect and kind == "str":
            out.append(ttext[1:-1].replace('""', '"'))
            expect = False
        elif not expect and kind == "op" and ttext == "&":
            expect = True
        else:
            return None
    return None if expect else "".join(out)


@dataclass(frozen=True)
class Reach:
    """Everything a formula reaches directly and through defined names (transitively).
    via: tuple of DefinedName objects walked to get there (() = the formula itself)."""
    functions: frozenset            # normalised names of built-in (and eta) calls
    calls: tuple                    # ((via, FunctionCall), ...)
    operands: tuple                 # ((via, Operand), ...)
    names: tuple                    # DefinedName objects reached, in BFS order
    cycle: bool                     # a name refers back to itself (directly or not)

    def calls_any(self, *fnames) -> bool:
        want = {x.upper() for x in fnames}
        return bool(self.functions & want)


def expand(formula, scope_sheet: Optional[str], names, include_indirect_literals: bool = False,
           sheet_names=None) -> Reach:
    """Fully expand the functions / operands a formula reaches through defined names.

    Names are resolved with Excel scope rules (NameTable.resolve); each name's own
    definition is parsed with its own scope (sheet-local -> that sheet, global -> None).
    Cycle-safe.  A name definition that cannot be parsed raises FormulaError (no silent
    skip): it is a grading problem the check must surface."""
    f = _as_formula(formula)
    tab = _table(names, sheet_names)
    calls = []
    operands = []
    reached = []
    seen = set()
    cycle = False
    stack = [((), f, scope_sheet)]
    while stack:
        via, fx, scope = stack.pop(0)
        for c in fx.functions:
            if not c.local:
                calls.append((via, c))
        for o in fx.operands:
            operands.append((via, o))
        for n in names_used_by(fx, scope, tab, include_indirect_literals):
            if any(n is v for v in via):
                cycle = True
                continue
            if id(n) in seen:
                continue
            seen.add(id(n))
            reached.append(n)
            try:
                nf = parse(n.text)
            except FormulaError as e:
                raise FormulaError(f"defined name {n.name!r} has an unparsable definition: {e}") from e
            stack.append((via + (n,), nf, n.scope))
    fnames = frozenset(c.name for _, c in calls if c.builtin)
    return Reach(fnames, tuple(calls), tuple(operands), tuple(reached), cycle)


def functions_reached(formula, scope_sheet, names) -> frozenset:
    """Normalised built-in function names called directly or through defined names."""
    return expand(formula, scope_sheet, names).functions


# --------------------------------------------------------------------------- translation
def _split_cell_ref(ref: str):
    ref = ref.strip()
    k = _outside_split_index(ref, "!", last=True)
    if k >= 0:
        ref = ref[k + 1:]
    m = _RE_CELL.match(ref)
    if not m:
        raise FormulaError(f"not a cell address: {ref!r}")
    return int(m.group(4)), col_index(m.group(2))


_RE_PART_SHIFT = re.compile(r"^(\.?)(\$?)([A-Za-z]{1,3})?(\$?)(\d{1,7})?(\.?)$")


def _shift_part(p: str, dr: int, dc: int) -> Optional[str]:
    """Shift one endpoint ('$A1', 'B$2', 'C', '$5', '.D' ...).  None = off the grid."""
    m = _RE_PART_SHIFT.match(p)
    if not m or (m.group(3) is None and m.group(5) is None):
        return p                                   # a name endpoint or #REF!: unchanged
    lead, cabs, col, rabs, row, trail = m.groups()
    if col is None:                                # '$5': the '$' belongs to the row
        rabs, cabs = cabs, ""
    out_col = out_row = ""
    try:
        if col is not None:
            if cabs or not dc:
                out_col = cabs + col
            else:
                out_col = Translator.translate_col(col.upper(), dc)
                if col_index(out_col) > MAX_COL:
                    return None
        if row is not None:
            if rabs or not dr:
                out_row = rabs + row
            else:
                out_row = Translator.translate_row(row, dr)
                if int(out_row) > MAX_ROW:
                    return None
    except (TranslatorError, ValueError):
        return None
    return f"{lead}{out_col}{out_row}{trail}"


def _shift_operand(raw: str, dr: int, dc: int, parts: tuple = ()) -> str:
    """parts: the operand's RefParts (one per ':' piece); only cell / col / row pieces move
    (a name endpoint such as Tax in Tax:Rate is kept byte-for-byte)."""
    k = _outside_split_index(raw, "!", last=False)
    prefix, body = (raw[:k + 1], raw[k + 1:]) if k >= 0 else ("", raw)
    pieces = []
    j = 0
    while True:
        idx = _outside_split_index(body[j:], ":", last=False)
        if idx < 0:
            pieces.append(body[j:])
            break
        pieces.append(body[j:j + idx])
        j += idx + 1
    if parts and len(parts) != len(pieces):     # pragma: no cover - same splitting as _ref_info
        raise FormulaError(f"internal: cannot align the endpoints of {raw!r} for shifting")
    out = []
    for pi, p in enumerate(pieces):
        if parts and parts[pi].kind not in ("cell", "col", "row"):
            out.append(p)
            continue
        kk = _outside_split_index(p, "!", last=True)
        pq, pb = (p[:kk + 1], p[kk + 1:]) if kk >= 0 else ("", p)
        s = _shift_part(pb, dr, dc)
        if s is None:
            return prefix + "#REF!"
        out.append(pq + s)
    return prefix + ":".join(out)


def shift(formula: str, drow: int, dcol: int) -> str:
    """Move every relative A1 reference by (drow, dcol), as Excel does when a formula is
    filled / a shared formula is expanded.  Absolute parts ($A, $1), names, structured
    references, strings and errors are unchanged.  A reference pushed off the grid
    becomes #REF! (keeping its sheet qualifier), as in Excel.  Returns text without '='."""
    f = parse(formula)
    if drow == 0 and dcol == 0:
        return f.body
    out = []
    last = 0
    for o in f.operands:
        if o.kind not in _SHIFTABLE or o.r1c1:
            continue
        out.append(f.body[last:o.start])
        out.append(_shift_operand(o.raw, drow, dcol, o.parts))
        last = o.end
    out.append(f.body[last:])
    return "".join(out)


def translate(master_formula: str, from_ref: str, to_ref: str) -> str:
    """Shared-formula expansion: the master's text (stored at from_ref) as it reads at
    to_ref.  Wraps openpyxl's Translator row/column arithmetic but uses this module's
    lexer, so spill refs (A1#), TAB/CR, A1:OFFSET(...), trimmed refs (A:.A), quoted /
    external / 3-D qualifiers and structured references are handled.  No leading '='."""
    r0, c0 = _split_cell_ref(from_ref)
    r1, c1 = _split_cell_ref(to_ref)
    return shift(master_formula, r1 - r0, c1 - c0)


# --------------------------------------------------------------------------- quick filters
# A quoted sheet name ('Q"1' may hold a double quote), a [...] group (external book or
# structured reference) or a string literal, matched left to right like the lexer does;
# only string literals are blanked.
_RE_MASK_SCAN = re.compile(r"'(?:[^']|'')*'|" + _BRACKET + r'|"(?:[^"]|"")*"', re.S)
_RE_QUICK_WHOLE = re.compile(
    r"(?<![A-Za-z0-9_.\\?])\$?[A-Za-z]{1,3}\.?:\.?\$?[A-Za-z]{1,3}(?![A-Za-z0-9_.(\\?])"
    r"|(?<![A-Za-z0-9_.\\?$])\$?\d{1,7}\.?:\.?\$?\d{1,7}(?![0-9A-Za-z_.(])"
    r"|1048576|[Xx][Ff][Dd]\$?\d")


def _mask_one(m) -> str:
    s = m.group()
    return '"' + " " * (len(s) - 2) + '"' if s[0] == '"' else s


def mask_strings(text: str) -> str:
    """Formula text with every string literal's content blanked (same length).  Quoted
    sheet names ('Q"1'!A1) and [...] groups are skipped, so a '"' inside them does not
    start a string."""
    return _RE_MASK_SCAN.sub(_mask_one, text)


def quick_may_have_whole_refs(text: str) -> bool:
    """Conservative pre-filter for whole_column_or_row_refs (direct text only, names not
    expanded): False guarantees the formula has no whole-column/row operand (incl. sheet-
    edge ranges).  Tested against the full parser on every toy formula."""
    if ":" not in text:
        return False
    return bool(_RE_QUICK_WHOLE.search(mask_strings(text)))


def quick_may_call(text: str, function_names: Iterable[str]) -> bool:
    """Conservative pre-filter: False guarantees none of the (upper-case) function names
    is called directly (as FUNC( or _xleta.FUNC).  Names are not expanded."""
    up = mask_strings(text).upper()
    return any(fn in up for fn in function_names)
