"""Excel number-format engine for the deterministic checks (pure functions, no I/O).

Used by checks 66 (zeros as dashes), 94 (formats that print nothing) and 73 (approximate
display length for row heights).  What matters most is: blank detection, dash
detection, minus versus parentheses, and the approximate display length; exact digits
are rendered as Excel would in the en-US locale, but are secondary.

    r = render(0, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)')
    r.text == ' -   '    r.zero_as_dash is True    r.is_blank is False
    render(5, ';;;').is_blank is True
    render(0, 43).zero_as_dash is True          # built-in id
    resolve_format(8, {8: '#,##0.00\\ "€"'})    # custom definition beats the built-in table

Excel rules implemented (ECMA-376 Part 1 18.8.30/31 plus observed Excel behaviour):
- Sections ';' (outside quotes/brackets/escapes): 1 = all numbers (negatives get an
  automatic '-'), 2 = (>=0 ; <0), 3 = (>0 ; <0 ; =0), 4th = text (booleans too).  A last section with
  '@' (fewer than 4 sections) is the text section; a lone '@' format shows numbers as
  General.  The second/third sections never add an automatic minus.
- Conditions [>=1000], [<=-0.005], [=0], [<>1], E-notation [>=1E11]: first matching of
  the first two conditional sections, else the remaining section; no section -> '#####'.
- Colour tags [Red] ... [White], [ColorN] (N = 1..56, default or custom palette).
- Literals "..." and \\x and the characters Excel shows unquoted ($ - + ( ) : space ...),
  _x padding (one blank the width of x), *x fill (repeats x; omitted from .text),
  @ text placeholder, % (x100 each), trailing commas (/1000 each), thousands separator,
  0 # ? placeholders (? = blank-padded), E+/E- scientific (engineering when the
  integer part is ##0), fractions (# ?/?, # ??/100), General inside a section,
  dates/times (m = minutes after h / before s), elapsed [h] [m] [s], fractional seconds,
  AM/PM, A/P, locale tags [$€-407] / [$-409], [DBNum1] and friends (ignored).

Measured in Excel by Patrick (2026-10-03, detchecks/out/excel_session_answers.md Q4/Q8),
now certain: booleans go through the TEXT section like text ("TRUE" under 0;-0;0;"txt" shows
"txt", under ;;; nothing); a numeric section's colour tag does not colour text or booleans
([Red]0 on "abc" shows the font colour); a section chosen by its own condition and the
fallback section after conditions get NO automatic minus ([<0]0 and [>0]0;0 both show -5
as 5); an empty format code is General; a spaces-only code prints nothing; elapsed [h]:mm
keeps counting past 9999-12-31 (50000 -> 1200000:00); errors are shown unformatted (colour
tags do not apply).  A conditional-format font colour beats a number-format colour tag (the
checks apply that).

Known limits (documented in docs/numfmt.md): column width is unknown here, so a value
too wide for its column ('#####') and General's width-dependent rounding are not
modelled; era/Buddhist year codes (e, g, b) are approximated.  Still certain=False: a
negative number in the first section when only the SECOND section has a condition, and
'#####' when no section applies (both unmeasured); and every number under a code with a
section that mixes placeholders (0 # ? outside fractional seconds, General, @) with unquoted
date-time letters - '0 bps', '0 days', '#,##0 d' as openpyxl writes them (finding 66-S3,
mixed_date_letters; text under such a code without a text section is shown as typed, certain).
"""
from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass, replace as _dc_replace
from decimal import ROUND_HALF_UP, Decimal, localcontext
from functools import lru_cache
from typing import Optional, Union

from detchecks.errors import GradingError


class NumFmtError(GradingError):
    """A number format or value the engine cannot render (loud by design)."""


# --------------------------------------------------------------------------- built-in table
# en-US Excel display of the built-in ids.  ECMA-376's id 44 lacks its ';' separators
# (openpyxl copies that bug); the corrected code is used here.
BUILTIN_FORMATS: dict = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);\\("$"#,##0\\)',
    6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
    8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    # 27-36, 50-58: East Asian date/time formats (zh-CN spelling; date in every locale
    # that defines them).
    27: 'yyyy"年"m"月"',
    28: 'm"月"d"日"',
    29: 'm"月"d"日"',
    30: "m-d-yy",
    31: 'yyyy"年"m"月"d"日"',
    32: 'h"时"mm"分"',
    33: 'h"时"mm"分"ss"秒"',
    34: 'AM/PMh"时"mm"分"',
    35: 'AM/PMh"时"mm"分"ss"秒"',
    36: 'yyyy"年"m"月"',
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
    50: 'yyyy"年"m"月"',
    51: 'm"月"d"日"',
    52: 'yyyy"年"m"月"',
    53: 'm"月"d"日"',
    54: 'm"月"d"日"',
    55: 'AM/PMh"时"mm"分"',
    56: 'AM/PMh"时"mm"分"ss"秒"',
    57: 'yyyy"年"m"月"',
    58: 'm"月"d"日"',
    # 59-62, 67-81: Thai-locale formats ('t' = Thai digits) -> Western equivalents
    59: "0",
    60: "0.00",
    61: "#,##0",
    62: "#,##0.00",
    67: "0%",
    68: "0.00%",
    69: "# ?/?",
    70: "# ??/??",
    71: "d/m/yyyy",
    72: "d-mmm-yy",
    73: "d-mmm",
    74: "mmm-yy",
    75: "h:mm",
    76: "h:mm:ss",
    77: "d/m/yyyy h:mm",
    78: "mm:ss",
    79: "[h]:mm:ss",
    80: "mm:ss.0",
    81: "d/m/yyyy",
}
# Ids whose display depends on the Excel UI locale (currency symbol, date order, CJK/Thai).
LOCALE_DEPENDENT_IDS = frozenset({5, 6, 7, 8, 14, 15, 16, 17, 22} | set(range(27, 37)) | set(range(50, 82)))
# 23-26 and 63-66 are not defined by ECMA-376 and have no entry: resolving them raises.


def builtin_format(num_fmt_id: int) -> Optional[str]:
    """Format code of a built-in id (en-US display), or None when the id has none."""
    return BUILTIN_FORMATS.get(int(num_fmt_id))


def resolve_format(num_fmt_id, custom_formats: Optional[dict] = None) -> str:
    """Code for a cellXfs numFmtId: the workbook's own <numFmt> definition wins (also for
    ids below 164 - toy 94/T5 redefines id 8), then the built-in table.  An id with
    neither raises NumFmtError (Excel's handling of undefined ids is not assumed)."""
    fid = int(num_fmt_id)
    if custom_formats:
        code = custom_formats.get(fid)
        if code is None:
            code = custom_formats.get(str(fid))
        if code is not None:
            return code
    code = BUILTIN_FORMATS.get(fid)
    if code is None:
        raise NumFmtError(f"numFmtId {fid} has no definition in styles.xml and is not a built-in format")
    return code


# --------------------------------------------------------------------------- colours
COLOR_NAMES = {"black": "000000", "blue": "0000FF", "cyan": "00FFFF", "green": "00FF00",
               "magenta": "FF00FF", "red": "FF0000", "white": "FFFFFF", "yellow": "FFFF00"}
DEFAULT_INDEXED_PALETTE = (
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
)
_RE_COLOR = re.compile(r"^(black|blue|cyan|green|magenta|red|white|yellow|colou?r\s*(\d{1,2}))$", re.I)


def color_tag_rgb(tag: Optional[str], indexed_palette=None) -> Optional[str]:
    """'Red' -> 'FF0000'; 'Color10' -> palette[17] (custom <indexedColors> if given)."""
    if not tag:
        return None
    m = _RE_COLOR.match(tag.strip())
    if not m:
        return None
    if m.group(2):
        k = int(m.group(2))
        if not 1 <= k <= 56:
            return None
        pal = indexed_palette or DEFAULT_INDEXED_PALETTE
        idx = k + 7
        if idx >= len(pal):
            return None
        v = pal[idx]
        return v[-6:].upper() if v else None
    return COLOR_NAMES[m.group(1).lower()]


# --------------------------------------------------------------------------- parsing
_RE_COND = re.compile(r"^\s*(<=|>=|<>|<|>|=)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*$")
_RE_ELAPSED = re.compile(r"^(h+|m+|s+)$", re.I)
_DATE_LETTERS = set("yYmMdDhHsS")
_DASHES = "-–—−"          # hyphen, en dash, em dash, minus sign


@dataclass(frozen=True)
class Section:
    raw: str
    tokens: tuple                 # ((kind, value), ...)
    color: Optional[str]          # colour tag text ('Red', 'Color10') or None
    condition: Optional[tuple]    # (op, threshold) or None
    kind: str                     # empty | general | text | date | elapsed | number | scientific | fraction | literal
    has_at: bool
    locale_tags: tuple            # raw [$...] / [DBNum1] ... contents

    @property
    def is_empty(self) -> bool:
        return self.kind == "empty"


@dataclass(frozen=True)
class NumberFormat:
    code: str
    sections: tuple               # all sections as written
    numeric: tuple                # indexes of the sections used for numbers
    text_index: Optional[int]     # index of the text section, or None (text shown as typed)
    general_numbers: bool         # lone '@' (or '@' only section): numbers shown as General
    has_conditions: bool
    verified: bool = True         # False: Excel's reading of this code is unverified - a section mixes a number /
    #                               text placeholder with unquoted date-time letters ('0 bps', '0 days', '#,##0 d':
    #                               mixed_date_letters); renders under it are certain=False (see render)
    unverified_why: Optional[str] = None   # why verified is False (for messages), else None

    @property
    def is_date(self) -> bool:
        """The first numeric section is a date/time/elapsed format."""
        if not self.numeric:
            return False
        return self.sections[self.numeric[0]].kind in ("date", "elapsed")


def split_sections(code: str) -> list:
    """Split a format code on ';' outside "quotes", [brackets] and \\escapes / _x / *x."""
    out, cur = [], []
    i, n = 0, len(code)
    while i < n:
        c = code[i]
        if c == '"':
            j = code.find('"', i + 1)
            if j < 0:
                raise NumFmtError(f"unterminated quoted literal in number format {code!r}")
            cur.append(code[i:j + 1])
            i = j + 1
            continue
        if c in "\\_*" and i + 1 < n:
            cur.append(code[i:i + 2])
            i += 2
            continue
        if c == "[":
            j = code.find("]", i + 1)
            if j < 0:
                raise NumFmtError(f"unbalanced '[' in number format {code!r}")
            cur.append(code[i:j + 1])
            i = j + 1
            continue
        if c == ";":
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(c)
        i += 1
    out.append("".join(cur))
    return out


def _tokenize_section(sec: str, code: str) -> list:
    toks = []
    i, n = 0, len(sec)
    low = sec.lower()
    while i < n:
        c = sec[i]
        if c == '"':
            j = sec.find('"', i + 1)
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
            continue
        if c == "\\":
            if i + 1 < n:
                toks.append(("lit", sec[i + 1]))
            i += 2
            continue
        if c == "_":
            if i + 1 < n:
                toks.append(("pad", sec[i + 1]))
            i += 2
            continue
        if c == "*":
            if i + 1 < n:
                toks.append(("fill", sec[i + 1]))
            i += 2
            continue
        if c == "[":
            j = sec.find("]", i + 1)
            inner = sec[i + 1:j]
            i = j + 1
            if _RE_COLOR.match(inner.strip()):
                toks.append(("color", inner.strip()))
                continue
            m = _RE_COND.match(inner)
            if m:
                toks.append(("cond", (m.group(1), float(m.group(2)))))
                continue
            if _RE_ELAPSED.match(inner):
                toks.append(("elapsed", inner.lower()))
                continue
            if inner.startswith("$"):
                sym = inner[1:].split("-", 1)[0]
                toks.append(("locale", inner))
                if sym:
                    toks.append(("lit", sym))
                continue
            toks.append(("locale", inner))       # [DBNum1], [NatNum1], [ENG] ... no output
            continue
        if low.startswith("general", i):
            toks.append(("general", sec[i:i + 7]))
            i += 7
            continue
        if c in "0#?":
            toks.append(("digit", c))
            i += 1
            continue
        if c == ".":
            toks.append(("dot", "."))
            i += 1
            continue
        if c == ",":
            toks.append(("comma", ","))
            i += 1
            continue
        if c == "%":
            toks.append(("pct", "%"))
            i += 1
            continue
        if c in "eE" and i + 1 < n and sec[i + 1] in "+-":
            toks.append(("exp", sec[i:i + 2]))
            i += 2
            continue
        if c == "@":
            toks.append(("at", "@"))
            i += 1
            continue
        if low.startswith("am/pm", i):
            toks.append(("ampm", sec[i:i + 5]))
            i += 5
            continue
        if low.startswith("a/p", i):
            toks.append(("ampm", sec[i:i + 3]))
            i += 3
            continue
        if c in _DATE_LETTERS:
            j = i
            while j < n and sec[j].lower() == c.lower():
                j += 1
            toks.append(("date", sec[i:j].lower()))
            i = j
            continue
        if c in "eEgGbB":                       # era / Buddhist year codes (date sections only)
            j = i
            while j < n and sec[j].lower() == c.lower():
                j += 1
            toks.append(("datex", sec[i:j]))
            i = j
            continue
        if c == "/":
            toks.append(("slash", "/"))
            i += 1
            continue
        toks.append(("lit", c))
        i += 1
    return toks


def _section_kind(toks) -> str:
    if not toks:
        return "empty"                          # ';;;' sections: nothing at all
    kinds = {k for k, _ in toks}
    visible = kinds - {"color", "cond", "locale"}
    if not visible:
        return "general"                        # '[Red]' / '[Blue][>100]': General in that colour
    if "elapsed" in kinds:
        return "elapsed"
    if "date" in kinds or "ampm" in kinds:
        return "date"
    if "general" in kinds:
        return "general"
    if "at" in kinds and "digit" not in kinds:
        return "text"
    if "digit" in kinds:
        if "exp" in kinds:
            return "scientific"
        if "slash" in kinds and _fraction_layout(toks) is not None:
            return "fraction"
        return "number"
    return "literal"


_DATE_TOKEN_KINDS = frozenset({"date", "datex", "ampm", "elapsed"})


def mixed_date_letters(toks) -> Optional[str]:
    """Does one section mix a number / text placeholder with unquoted date-time letters?  Returns a short
    description ('0 with d, y, s') or None.  Excel's reading of such a section is unverified (finding
    66-S3: openpyxl writes '0 bps', '0 days', '#,##0 d' unquoted; the engine would read the unit words as
    date codes, and what Excel shows is unmeasured), so parse_format marks the whole code unverified.
      placeholders: the digit placeholders 0 # ? (except the 0s of a fractional-seconds group: a '.'
                    right after a seconds code s / ss / [s] / [ss], as in 'h:mm:ss.000', 'mm:ss.0'),
                    General, and the text placeholder @;
      date-time letters (unquoted, unescaped, outside brackets): y m d h s (date / time codes, any case),
                    e g b (era / Buddhist-year codes; E followed by + or - is scientific notation, not a
                    letter), AM/PM and A/P, and the elapsed brackets [h] [m] [s].
    A genuine date / time code (no placeholder) and a number code whose letters are quoted or escaped
    ('0" bps"', '0\\x', '"FY"0"E"') are not mixed; an unquoted letter that is no date code ('0.0x') is a
    literal and is not affected."""
    letters, holders = [], []
    prev_seconds = False
    i, n = 0, len(toks)
    while i < n:
        k, v = toks[i]
        if k == "dot" and prev_seconds and i + 1 < n and toks[i + 1] == ("digit", "0"):
            i += 1
            while i < n and toks[i] == ("digit", "0"):
                i += 1                              # fractional seconds: part of the time code
            prev_seconds = False
            continue
        if k in _DATE_TOKEN_KINDS:
            letters.append(v if k != "elapsed" else f"[{v}]")
        elif k in ("digit", "general", "at"):
            holders.append(v)
        prev_seconds = (k == "date" and v[:1] == "s") or (k == "elapsed" and v[:1] == "s")
        i += 1
    if letters and holders:
        return f"{''.join(dict.fromkeys(holders))} with {', '.join(dict.fromkeys(letters))}"
    return None


@lru_cache(maxsize=4096)
def parse_format(code: str) -> NumberFormat:
    """Parse a format code (cached).  Raises NumFmtError on malformed codes.  A code with a section
    that mixes placeholders with unquoted date-time letters (mixed_date_letters) parses as before but
    is unverified (verified=False, unverified_why): numbers rendered under it are certain=False, and
    so is text when the code has a text section (render)."""
    if code is None:
        raise NumFmtError("number format code is None")
    verified = True                             # every other code's reading is measured (2026-10-03)
    why = None
    if code == "":
        code = "General"                        # empty formatCode: General (Excel 2026-10-03, no repair prompt)
    elif code.strip() == "":
        pass                                    # spaces only: a literal section, prints nothing (Excel 2026-10-03)
    elif code.strip().lower() == "general":
        code = "General"
    raw_secs = split_sections(code)
    if len(raw_secs) > 4:
        raise NumFmtError(f"number format {code!r} has {len(raw_secs)} sections (Excel allows 4)")
    sections = []
    for s in raw_secs:
        toks = _tokenize_section(s, code)
        color = next((v for k, v in toks if k == "color"), None)
        cond = next((v for k, v in toks if k == "cond"), None)
        kind = _section_kind(toks)
        has_at = any(k == "at" for k, _ in toks)
        locale = tuple(v for k, v in toks if k == "locale")
        sections.append(Section(s, tuple(toks), color, cond, kind, has_at, locale))
        mix = mixed_date_letters(toks)
        if mix is not None and verified:
            verified = False
            why = (f"section {len(sections)} of {code!r} mixes placeholders and unquoted date-time letters "
                   f"({mix}); Excel's reading of it is not verified")
    n = len(sections)
    text_index = None
    general_numbers = False
    if n == 4:
        text_index = 3
        numeric = (0, 1, 2)
    elif sections[-1].has_at:
        text_index = n - 1
        numeric = tuple(range(n - 1))
        if n == 1:
            general_numbers = True
            numeric = ()
    else:
        numeric = tuple(range(n))
    has_cond = any(sections[i].condition is not None for i in numeric[:2])
    return NumberFormat(code, tuple(sections), numeric, text_index, general_numbers, has_cond, verified, why)


def _test(cond, v) -> bool:
    op, x = cond
    if op == "<":
        return v < x
    if op == "<=":
        return v <= x
    if op == ">":
        return v > x
    if op == ">=":
        return v >= x
    if op == "=":
        return v == x
    return v != x


def select_section(fmt: Union[str, NumberFormat], value: float):
    """(section_index or None, auto_minus, certain) for a number.
    None means no section applies (Excel shows '#####')."""
    f = fmt if isinstance(fmt, NumberFormat) else parse_format(fmt)
    num = f.numeric
    if not num:
        return None, value < 0, True           # General numbers (lone '@')
    secs = f.sections
    if not f.has_conditions:
        if len(num) == 1:
            return num[0], value < 0, True
        if len(num) == 2:                       # '0;(0)' and '0;-0;@' (zero -> first)
            return (num[0], False, True) if value >= 0 else (num[1], False, True)
        if value > 0:
            return num[0], False, True
        if value < 0:
            return num[1], False, True
        return num[2], False, True
    c0 = secs[num[0]].condition
    c1 = secs[num[1]].condition if len(num) > 1 else None
    # Excel (measured 2026-10-03): no automatic minus in a section chosen by its own condition
    # ([<0]0;0 on -5 -> 5) nor in the fallback section ([>0]0;0 on -5 -> 5).
    if c0 is not None and _test(c0, value):
        return num[0], False, True
    if c1 is not None and _test(c1, value):
        return num[1], False, True
    if c0 is not None and c1 is not None:
        if len(num) >= 3:
            return num[2], False, True
        return None, False, False               # '#####': unmeasured
    if c0 is not None:                          # only the first section is conditional
        if len(num) >= 2:
            return num[1], False, True
        return None, False, False
    return num[0], value < 0, value >= 0        # only the second is conditional: unmeasured for negatives


# --------------------------------------------------------------------------- result
@dataclass(frozen=True)
class Rendered:
    text: str                  # display at natural width: '_x' -> one space, '?' -> space, '*x' fill omitted
    is_blank: bool             # nothing visible (empty, or only padding / space fill)
    zero_as_dash: bool         # numeric 0 displayed as a dash (no digit), e.g. accounting ' -   '
    shows_dash: bool           # visible text is dash-like (a dash, no digit) whatever the value
    color_tag: Optional[str]   # '[Red]' -> 'Red' of the section used
    color_rgb: Optional[str]   # tag resolved to RRGGBB (default or given palette)
    section_index: Optional[int]
    kind: str                  # general number scientific fraction date elapsed text literal empty bool error hash
    is_hash: bool              # Excel shows '#####' (negative date/time, no section for the value)
    fill_char: Optional[str]
    minus: bool                # a minus sign is visible next to the number
    parens: bool               # the number is wrapped in literal parentheses
    certain: bool              # False where Excel's behaviour is not verified (see docs)
    locale_dependent: bool     # built-in id whose display depends on Excel's locale
    format_code: str

    @property
    def visible_text(self) -> str:
        return self.text.strip()

    @property
    def display_length(self) -> int:
        """Approximate width in characters (padding and '?' blanks count, fill does not)."""
        return len(self.text)


_NUMERIC_KINDS = frozenset({"general", "number", "scientific", "fraction", "literal", "text"})


def _finish(text, value, *, color, section_index, kind, fill=None, certain=True, locale=False,
            code="", is_hash=False, palette=None, numeric=True) -> Rendered:
    vis = text.strip()
    has_digit = any(ch.isdigit() for ch in vis)
    blank = vis == "" and (fill is None or fill.isspace())
    dashy = not has_digit and (any(ch in _DASHES for ch in vis) or (fill is not None and fill in _DASHES))
    minus = parens = False
    is_num = isinstance(value, (int, float)) and not isinstance(value, bool)
    if numeric and has_digit and kind in _NUMERIC_KINDS:
        first = next(i for i, ch in enumerate(vis) if ch.isdigit())
        last = max(i for i, ch in enumerate(vis) if ch.isdigit())
        # a sign only for a negative value: the dash of '"FY"-0' or '"Year -"0' shown
        # before a positive number is a literal
        minus = is_num and value < 0 and any(ch in "-\u2212" for ch in vis[:first])
        parens = "(" in vis[:first] and ")" in vis[last + 1:]
    zero_dash = numeric and is_num and value == 0 and dashy
    return Rendered(text, blank, zero_dash, dashy, color, color_tag_rgb(color, palette), section_index,
                    kind, is_hash, fill, minus, parens, certain, locale, code)


# --------------------------------------------------------------------------- numbers
def _sig15(x: float) -> Decimal:
    """|x| as Excel holds it for display: 15 significant digits (exact decimal)."""
    d = Decimal(repr(abs(x)))
    return Decimal(format(d, ".15g")) if d != 0 else d


def _quantize(d: Decimal, places: int) -> Decimal:
    """Round a non-negative Decimal half away from zero at `places` decimals.  The
    context precision is sized to the number, so 1E+30 under '#,##0' gives all 31
    digits instead of decimal.InvalidOperation (default precision is 28)."""
    with localcontext() as ctx:
        ctx.prec = max(28, d.adjusted() + places + 3)
        return d.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def _dec_round(x: float, places: int) -> Decimal:
    """Round |x| half away from zero at `places` decimals, from Excel's 15 significant
    digits (so 0.125 -> 0.13 and 2.675 -> 2.68 as Excel shows)."""
    return _quantize(_sig15(x), places)


def _scaled15(a: float, pct: int, scale: int) -> Decimal:
    """|a| x 100**pct / 1000**scale ('%' and trailing commas) at 15 significant digits.
    Scaled in floating point first (as before), exactly in Decimal when that overflows."""
    v = a * (100 ** pct) / (1000 ** scale) if (pct or scale) else a
    if math.isinf(v):
        d = _sig15(a)
        return d.scaleb(2 * pct - 3 * scale) if d != 0 else d
    return _sig15(v)


def general_text(value: float) -> str:
    """Excel's General display of a number in a standard-width cell (<= 11 characters):
    up to 10 significant digits; scientific (6 significant digits) for |v| >= 1E11 or
    |v| < 1E-4 (Excel shows 0.0001 but 1E-05).  Width-dependent shortening is not modelled."""
    if value == 0:
        return "0"
    a = abs(value)
    sign = "-" if value < 0 else ""

    def sci() -> str:
        d = _sig15(a)
        exp = d.adjusted()
        mant = _quantize(d.scaleb(-exp), 5)
        if mant >= 10:
            exp += 1
            mant = _quantize(d.scaleb(-exp), 5)
        ms = format(mant, "f").rstrip("0").rstrip(".")
        return f"{sign}{ms}E{'+' if exp >= 0 else '-'}{abs(exp):02d}"

    if a >= 1e11 or a < 1e-4:
        return sci()
    int_digits = len(str(int(a))) if a >= 1 else 1
    dec = max(0, 10 - int_digits)
    r = _dec_round(a, dec)
    s = format(r, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s == "0":                                 # very small: fall back to scientific
        return sci()
    return sign + s


def _number_layout(toks):
    """(int_idx, dec_idx, grouping, scale_commas, pct, dot, exp) of a numeric section."""
    digit_idx = [i for i, (k, _) in enumerate(toks) if k == "digit"]
    exp = next((i for i, (k, _) in enumerate(toks) if k == "exp"), None)
    limit = exp if exp is not None else len(toks)
    dot = next((i for i, (k, _) in enumerate(toks) if k == "dot" and i < limit), None)
    int_end = dot if dot is not None else limit
    int_idx = [i for i in digit_idx if i < int_end]
    dec_idx = [i for i in digit_idx if dot is not None and dot < i < limit]
    grouping = False
    scale = 0
    for i, (k, _) in enumerate(toks):
        if k != "comma" or i >= limit or not any(d < i for d in digit_idx):
            continue
        if i < int_end:
            if any(i < d < int_end for d in digit_idx):
                grouping = True                 # #,##0
            else:
                scale += 1                      # #,##0, / 0,.00
        elif not any(i < d < limit for d in digit_idx):
            scale += 1                          # 0.0,,
    pct = sum(1 for k, _ in toks if k == "pct")
    return int_idx, dec_idx, grouping, scale, pct, dot, exp


def _group_str(digits: str) -> str:
    out = []
    for k, ch in enumerate(reversed(digits)):
        if k and k % 3 == 0:
            out.append(",")
        out.append(ch)
    return "".join(reversed(out))


def _int_slots(toks, int_idx, int_digits, grouping):
    """Map integer placeholder index -> text.  int_digits '' for an integer part of 0."""
    slot = {}
    if not int_idx:
        return slot
    ph = [toks[i][1] for i in int_idx]
    if grouping:
        zeros = sum(1 for c in ph if c == "0")
        d = int_digits.rjust(zeros, "0")
        text = _group_str(d) if d else ""
        qs = sum(1 for c in ph if c == "?")
        blanks = max(0, min(qs, len(ph) - len(d)))
        slot[int_idx[0]] = " " * blanks + text
        for i in int_idx[1:]:
            slot[i] = ""
        return slot
    n = len(int_idx)
    d = int_digits
    if len(d) > n:                              # extra digits go to the first placeholder
        slot[int_idx[0]] = d[:len(d) - n + 1]
        for i, ch in zip(int_idx[1:], d[len(d) - n + 1:]):
            slot[i] = ch
        return slot
    pad = n - len(d)
    for j, i in enumerate(int_idx):
        if j < pad:
            c = toks[i][1]
            slot[i] = "0" if c == "0" else (" " if c == "?" else "")
        else:
            slot[i] = d[j - pad]
    return slot


def _dec_slots(toks, dec_idx, frac_digits: str):
    slot = {}
    if not dec_idx:
        return slot
    ph = [toks[i][1] for i in dec_idx]
    digits = list(frac_digits.ljust(len(dec_idx), "0"))
    k = len(digits) - 1
    while k >= 0 and digits[k] == "0" and ph[k] in "#?":
        digits[k] = " " if ph[k] == "?" else ""
        k -= 1
    for i, ch in zip(dec_idx, digits):
        slot[i] = ch
    return slot


def _emit(toks, slot, a=None):
    out = []
    fill = None
    for i, (kind, val) in enumerate(toks):
        pre = slot.get(("pre", i))
        if pre:
            out.append(pre)
        if kind == "digit":
            out.append(slot.get(i, ""))
        elif kind == "dot":
            out.append(slot.get(i, "."))
        elif kind in ("lit", "slash", "datex", "date"):
            out.append(val)
        elif kind == "pct":
            out.append("%")
        elif kind == "pad":
            out.append(" ")
        elif kind == "fill":
            fill = val
        elif kind == "general":
            out.append(general_text(a) if a is not None else "")
        elif kind == "exp":
            out.append(slot.get(i, "E+"))
        elif kind == "ampm":
            out.append(val)
    return "".join(out), fill


def _render_fixed(toks, a: float):
    """Render |value| a with a fixed-point section.  Returns (text, fill)."""
    int_idx, dec_idx, grouping, scale, pct, dot, _ = _number_layout(toks)
    r = _quantize(_scaled15(a, pct, scale), len(dec_idx))
    ip, _, fp = format(r, "f").partition(".")
    int_digits = "" if ip == "0" else ip
    slot = _int_slots(toks, int_idx, int_digits, grouping)
    if not int_idx and int_digits and dot is not None:
        slot[("pre", dot)] = int_digits          # '.00' still shows the integer part
    slot.update(_dec_slots(toks, dec_idx, fp))
    return _emit(toks, slot)


def _render_general_section(toks, a: float):
    """A section holding General plus literals ('"Total "General', 'General;(General)'),
    or only tags ('[Red]'), which Excel shows as General."""
    if not any(k == "general" for k, _ in toks):
        txt, fill = _emit(toks, {}, a)
        return general_text(a) + txt, fill
    return _emit(toks, {}, a)


def _render_scientific(toks, a: float):
    int_idx, dec_idx, grouping, scale, pct, dot, exp_i = _number_layout(toks)
    exp_ph = [i for i, (k, _) in enumerate(toks) if k == "digit" and i > exp_i]
    d = _scaled15(a, pct, 0)
    int_ph = "".join(toks[i][1] for i in int_idx)
    nint = max(1, len(int_idx))
    engineering = nint > 1 and int_ph.startswith("#")
    zeros = max(1, int_ph.count("0"))
    ndec = len(dec_idx)
    if d == 0:
        e = 0
        mant = Decimal(0).quantize(Decimal(1).scaleb(-ndec))
    else:
        e0 = d.adjusted()
        step = nint if engineering else 1
        e = (e0 // nint) * nint if engineering else e0 - (zeros - 1)
        mant = _quantize(d.scaleb(-e), ndec)
        if mant >= 10 ** (nint if engineering else zeros):
            e += step
            mant = _quantize(d.scaleb(-e), ndec)
    ip, _, fp = format(mant, "f").partition(".")
    int_digits = "" if ip == "0" else ip
    slot = _int_slots(toks, int_idx, int_digits, False)
    slot.update(_dec_slots(toks, dec_idx, fp))
    nexp0 = sum(1 for i in exp_ph if toks[i][1] == "0")
    es = str(abs(e)).rjust(nexp0, "0")
    sign_mode = toks[exp_i][1][1]
    slot[exp_i] = "E" + ("-" if e < 0 else ("+" if sign_mode == "+" else ""))
    if exp_ph:
        slot[exp_ph[0]] = es
        for i in exp_ph[1:]:
            slot[i] = ""
    return _emit(toks, slot)


def _fraction_layout(toks):
    """(int_idx, num_idx, den_idx, den_token_idx, fixed_den) for '# ?/?', '?/?',
    '# ??/100', '0 ?/8'; None when the '/' is not a fraction."""
    slash = next((i for i, (k, _) in enumerate(toks) if k == "slash"), None)
    if slash is None:
        return None
    num_idx = []
    j = slash - 1
    while j >= 0 and toks[j][0] == "digit":
        num_idx.insert(0, j)
        j -= 1
    if not num_idx:
        return None
    int_idx = [i for i in range(0, j + 1) if toks[i][0] == "digit"]
    den_tok = []
    j = slash + 1
    while j < len(toks) and (toks[j][0] == "digit" or (toks[j][0] == "lit" and toks[j][1].isdigit())):
        den_tok.append(j)
        j += 1
    if not den_tok:
        return None
    den_text = "".join(toks[i][1] for i in den_tok)
    fixed = int(den_text) if den_text.isdigit() and toks[den_tok[0]][0] == "lit" else None
    den_idx = [] if fixed else den_tok
    return int_idx, num_idx, den_idx, den_tok, fixed


def _render_fraction(toks, a: float):
    int_idx, num_idx, den_idx, den_tok, fixed = _fraction_layout(toks)
    big = a >= 1e15                             # no fractional digit survives 15-digit precision
    if int_idx:
        whole = int(_sig15(a)) if big else int(math.floor(a))
        rest = 0.0 if big else a - whole
    else:
        whole, rest = 0, a
    if fixed:
        den = fixed
        num = int(_dec_round(rest * den, 0))
    elif big and not int_idx:
        num, den = int(_sig15(a)), 1
    else:
        num, den = _best_fraction(rest, 10 ** len(den_idx) - 1)
    if num == den and int_idx:
        whole, num = whole + 1, 0
    blank_frac = bool(int_idx) and num == 0
    nw = len(num_idx)
    dw = len(den_tok)
    slot = {}
    if int_idx:
        show_whole = whole != 0 or num == 0
        any_zero = any(toks[i][1] == "0" for i in int_idx)
        slot[int_idx[0]] = str(whole) if show_whole else ("0" if any_zero else "")
        for i in int_idx[1:]:
            slot[i] = ""
    slot[num_idx[0]] = " " * nw if blank_frac else str(num).rjust(nw)
    for i in num_idx[1:]:
        slot[i] = ""
    slash = next(i for i, (k, _) in enumerate(toks) if k == "slash")
    if blank_frac:
        slot[slash] = " "
    den_s = " " * dw if blank_frac else (str(den).ljust(dw) if not fixed else str(den))
    toks2 = list(toks)
    for n_, i in enumerate(den_tok):
        toks2[i] = ("digit", toks[i][1])
        slot[i] = den_s if n_ == 0 else ""
    toks2[slash] = ("dot", "/")                 # emitted through the slot (or '/')
    if not blank_frac:
        slot[slash] = "/"
    return _emit(toks2, slot)


def _best_fraction(x: float, maxden: int):
    if x <= 0:
        return 0, 1
    best = (0, 1, abs(x))
    for den in range(1, maxden + 1):
        num = int(math.floor(x * den + 0.5))
        err = abs(x - num / den)
        if err < best[2] - 1e-15:
            best = (num, den, err)
            if err == 0:
                break
    return best[0], best[1]


# --------------------------------------------------------------------------- dates
_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December")
_DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
MAX_DATE_SERIAL = 2958465.99999999     # 9999-12-31 23:59:59


def serial_to_datetime(serial: float, date1904: bool = False):
    """(year, month, day, weekday 0=Mon, hour, minute, second, fraction_of_second) for an
    Excel serial (1900 system with its fictitious 1900-02-29 = 60, serial 0 = 1900-01-00)."""
    days = int(math.floor(serial))
    secs = (serial - days) * 86400.0
    if date1904:
        d = _dt.date(1904, 1, 1) + _dt.timedelta(days=days)
        y, m, dd, wd = d.year, d.month, d.day, d.weekday()
    elif days == 0:
        y, m, dd, wd = 1900, 1, 0, 5          # Excel shows 1/0/1900 (a Saturday)
    elif days == 60:
        y, m, dd, wd = 1900, 2, 29, 2         # the 1900 leap-year bug
    else:
        d = _dt.date(1899, 12, 31) + _dt.timedelta(days=days if days < 60 else days - 1)
        y, m, dd, wd = d.year, d.month, d.day, d.weekday()
    h = int(secs // 3600)
    mi = int((secs % 3600) // 60)
    s = int(secs % 60)
    frac = secs - (h * 3600 + mi * 60 + s)
    return y, m, dd, wd, h, mi, s, frac


def _render_date(toks, value: float, date1904: bool):
    """Returns (text, fill) or None for '#####'."""
    if value < 0 or value > MAX_DATE_SERIAL:
        return None
    # precision of fractional seconds shown
    frac_digits = 0
    for i, (k, v) in enumerate(toks):
        if k == "dot" and i + 1 < len(toks) and toks[i + 1] == ("digit", "0"):
            j = i + 1
            while j < len(toks) and toks[j] == ("digit", "0"):
                j += 1
            frac_digits = max(frac_digits, j - i - 1)
    serial = float(_dec_round(value * 86400.0, frac_digits)) / 86400.0 if value else 0.0
    y, mo, d, wd, h, mi, s, frac = serial_to_datetime(serial, date1904)
    ampm = any(k == "ampm" for k, _ in toks)
    date_pos = [i for i, (k, _) in enumerate(toks) if k in ("date", "elapsed")]
    out = []
    fill = None
    i = 0
    n = len(toks)
    while i < n:
        k, v = toks[i]
        if k == "date":
            c = v[0]
            ln = len(v)
            pos = date_pos.index(i)
            prev = toks[date_pos[pos - 1]][1] if pos > 0 else ""
            nxt = toks[date_pos[pos + 1]][1] if pos + 1 < len(date_pos) else ""
            if c == "y":
                out.append(f"{y % 100:02d}" if ln <= 2 else f"{y:04d}")
            elif c == "m" and ln <= 2 and (prev[:1] == "h" or nxt[:1] == "s"):
                out.append(f"{mi:02d}" if ln == 2 else str(mi))
            elif c == "m":
                out.append(_MONTHS[mo - 1][0] if ln >= 5 else _MONTHS[mo - 1] if ln == 4
                           else _MONTHS[mo - 1][:3] if ln == 3 else f"{mo:02d}" if ln == 2 else str(mo))
            elif c == "d":
                out.append(_DAYS[wd] if ln >= 4 else _DAYS[wd][:3] if ln == 3 else f"{d:02d}" if ln == 2 else str(d))
            elif c == "h":
                hh = (h % 12 or 12) if ampm else h
                out.append(f"{hh:02d}" if ln >= 2 else str(hh))
            elif c == "s":
                out.append(f"{s:02d}" if ln >= 2 else str(s))
        elif k == "ampm":
            tag = "AM" if h < 12 else "PM"
            out.append(tag if len(v) == 5 else tag[0])
        elif k == "datex":
            c = v[0].lower()
            out.append(f"{y:04d}" if c == "e" else f"{(y + 543) % 100:02d}" if c == "b" and len(v) <= 2
                       else f"{y + 543:04d}" if c == "b" else "R")
        elif k == "dot":
            if frac_digits and i + 1 < n and toks[i + 1] == ("digit", "0"):
                fd = format(_dec_round(frac, frac_digits), "f").partition(".")[2][:frac_digits]
                out.append("." + fd.ljust(frac_digits, "0"))
                j = i + 1
                while j < n and toks[j] == ("digit", "0"):
                    j += 1
                i = j
                continue
            out.append(".")
        elif k in ("lit", "slash"):
            out.append(v)
        elif k == "digit":
            out.append(v)
        elif k == "comma":
            out.append(",")
        elif k == "pct":
            out.append("%")
        elif k == "pad":
            out.append(" ")
        elif k == "fill":
            fill = v
        i += 1
    return "".join(out), fill


def _render_elapsed(toks, value: float):
    if value < 0:
        return None
    frac_digits = 0
    for i, (k, v) in enumerate(toks):
        if k == "dot" and i + 1 < len(toks) and toks[i + 1] == ("digit", "0"):
            j = i + 1
            while j < len(toks) and toks[j] == ("digit", "0"):
                j += 1
            frac_digits = max(frac_digits, j - i - 1)
    secs = value * 86400.0
    if math.isinf(secs):                        # beyond float range: multiply exactly
        total = _quantize(_sig15(value) * 86400, frac_digits)
    else:
        total = _dec_round(secs, frac_digits)
    whole = int(total)
    frac = float(total - whole)
    unit = next(v[0] for k, v in toks if k == "elapsed")
    H, rem = divmod(whole, 3600)
    M, S = divmod(rem, 60)
    if unit == "m":
        M, S = divmod(whole, 60)
    elif unit == "s":
        S = whole
    out = []
    fill = None
    i, n = 0, len(toks)
    while i < n:
        k, v = toks[i]
        if k == "elapsed":
            val = H if v[0] == "h" else M if v[0] == "m" else S
            out.append(str(val).rjust(len(v), "0"))
        elif k == "date":
            c = v[0]
            if c == "h":
                out.append(str(H).rjust(len(v), "0"))
            elif c == "m":
                out.append(str(M).rjust(len(v), "0"))
            elif c == "s":
                out.append(str(S).rjust(len(v), "0"))
            else:
                out.append(v)
        elif k == "dot" and frac_digits and i + 1 < n and toks[i + 1] == ("digit", "0"):
            fd = format(Decimal(repr(frac)).quantize(Decimal(1).scaleb(-frac_digits), rounding=ROUND_HALF_UP), "f")
            out.append("." + fd.partition(".")[2].ljust(frac_digits, "0"))
            j = i + 1
            while j < n and toks[j] == ("digit", "0"):
                j += 1
            i = j
            continue
        elif k in ("lit", "slash", "digit"):
            out.append(v)
        elif k == "dot":
            out.append(".")
        elif k == "pad":
            out.append(" ")
        elif k == "fill":
            fill = v
        i += 1
    return "".join(out), fill


def _render_literal(toks, text_value=None):
    out = []
    fill = None
    for k, v in toks:
        if k in ("lit", "slash", "datex", "date", "digit"):
            out.append(v)
        elif k == "comma":
            out.append(",")
        elif k == "dot":
            out.append(".")
        elif k == "pct":
            out.append("%")
        elif k == "pad":
            out.append(" ")
        elif k == "fill":
            fill = v
        elif k == "at":
            out.append(text_value if text_value is not None else "")
        elif k == "general":
            out.append(text_value if text_value is not None else "")
    return "".join(out), fill


# --------------------------------------------------------------------------- render
class ErrorValue(str):
    """Marks a cached error value ('#DIV/0!') so render() does not treat it as text."""


def _iso_to_serial(s: str, date1904: bool) -> float:
    try:
        t = _dt.datetime.fromisoformat(s.replace("Z", ""))
    except ValueError as e:
        raise NumFmtError(f"cannot read ISO date value {s!r}: {e}") from e
    base = _dt.datetime(1904, 1, 1) if date1904 else _dt.datetime(1899, 12, 30)
    serial = (t - base).total_seconds() / 86400.0
    if not date1904 and serial < 61:
        serial -= 1                              # before the fictitious 1900-02-29
    return serial


def render(value, code_or_id, *, custom_formats: Optional[dict] = None, value_type: Optional[str] = None,
           date1904: bool = False, indexed_palette=None) -> Rendered:
    """Render a cell value through a number format.

    value: int/float (numbers, dates as serials), str (text), bool, None (empty cell), or
           ErrorValue('#N/A').  value_type, when given, is the OOXML cell type and wins:
           'n' number, 's'/'str'/'inlineStr' text, 'b' boolean, 'e' error, 'd' ISO date.
    code_or_id: a format code, or a numFmtId (int) resolved with custom_formats
           ({id: code} from styles.xml <numFmts>), then the built-in table.
    """
    locale = False
    if isinstance(code_or_id, bool):
        raise NumFmtError(f"invalid number format {code_or_id!r}")
    if isinstance(code_or_id, int):
        code = resolve_format(code_or_id, custom_formats)
        locale = code_or_id in LOCALE_DEPENDENT_IDS and not (
            custom_formats and (code_or_id in custom_formats or str(code_or_id) in custom_formats))
    else:
        code = code_or_id if code_or_id is not None else "General"
    f = parse_format(code)
    r = _render_value(f, value, code, locale, value_type, date1904, indexed_palette)
    if not f.verified and r.certain and not _shown_whatever_the_code(f, r):
        r = _dc_replace(r, certain=False)
    return r


def _shown_whatever_the_code(f: NumberFormat, r: Rendered) -> bool:
    """Under an unverified code, displays that do not depend on how Excel reads it: an empty cell, an
    error value (always shown unformatted, Excel 2026-10-03), and text / a logical when the code has no
    text section (shown as typed - the numeric sections never apply to text, and a code Excel refused
    would be General, which shows text as typed too).  Numbers, and text under a text section, are
    certain=False."""
    if r.kind in ("empty", "error"):
        return True
    return r.kind in ("text", "bool") and f.text_index is None


def _render_value(f: NumberFormat, value, code, locale, value_type, date1904, indexed_palette) -> Rendered:
    vt = value_type
    if vt is not None:
        if vt == "b":
            value = bool(int(value)) if isinstance(value, str) and value in ("0", "1") else bool(value)
        elif vt == "e":
            value = ErrorValue(str(value))
        elif vt in ("s", "str", "inlineStr"):
            value = "" if value is None else str(value)
        elif vt == "d":
            value = _iso_to_serial(str(value), date1904)
        elif vt == "n":
            if value is not None and not isinstance(value, (int, float)):
                try:
                    value = float(value)
                except (TypeError, ValueError) as e:
                    raise NumFmtError(f"numeric cell value {value!r} is not a number") from e
        else:
            raise NumFmtError(f"unknown cell value type {vt!r}")
    if value is None:
        return _finish("", None, color=None, section_index=None, kind="empty", code=code, locale=locale,
                       numeric=False, palette=indexed_palette)
    if isinstance(value, ErrorValue):
        return _finish(str(value), value, color=None, section_index=None, kind="error", code=code,
                       locale=locale, numeric=False, palette=indexed_palette)
    if isinstance(value, bool):
        # Excel 2026-10-03: a logical goes through the text section like the text TRUE / FALSE
        r = _render_text(f, "TRUE" if value else "FALSE", code, locale, indexed_palette)
        return _dc_replace(r, kind="bool")
    if isinstance(value, str):
        return _render_text(f, value, code, locale, indexed_palette)
    if not isinstance(value, (int, float)):
        raise NumFmtError(f"cannot render value of type {type(value).__name__}: {value!r}")
    v = float(value)
    if math.isnan(v) or math.isinf(v):
        raise NumFmtError(f"cannot render non-finite number {value!r}")
    return _render_number(f, v, value, code, locale, date1904, indexed_palette)


def _render_text(f: NumberFormat, s: str, code, locale, palette) -> Rendered:
    if f.text_index is None:
        # no text section: shown as typed, in the font colour - a numeric section's colour tag
        # does not colour text (Excel 2026-10-03: [Red]0 on "abc" is black)
        return _finish(s, s, color=None, section_index=None, kind="text", code=code, locale=locale,
                       numeric=False, palette=palette)
    sec = f.sections[f.text_index]
    if sec.kind == "general":                   # tag-only text section ('[Red]'): text as typed
        return _finish(s, s, color=sec.color, section_index=f.text_index, kind="text",
                       code=code, locale=locale, numeric=False, palette=palette)
    txt, fill = _render_literal(sec.tokens, s)
    return _finish(txt, s, color=sec.color, section_index=f.text_index, kind="text", fill=fill,
                   code=code, locale=locale, numeric=False, palette=palette)


def _render_number(f: NumberFormat, v: float, orig, code, locale, date1904, palette) -> Rendered:
    idx, auto_minus, certain = select_section(f, v)
    if f.general_numbers:
        return _finish(general_text(v), orig, color=None, section_index=None, kind="general",
                       code=code, locale=locale, palette=palette)
    if idx is None:
        return _finish("#####", orig, color=None, section_index=None, kind="hash", is_hash=True,
                       certain=certain, code=code, locale=locale, palette=palette)
    sec = f.sections[idx]
    toks = sec.tokens
    kind = sec.kind
    a = abs(v)
    fill = None
    if kind == "empty":
        return _finish("", orig, color=sec.color, section_index=idx, kind="empty", certain=certain,
                       code=code, locale=locale, palette=palette)
    if kind in ("date", "elapsed"):
        res = _render_date(toks, v, date1904) if kind == "date" else _render_elapsed(toks, v)
        if res is None:
            return _finish("#####", orig, color=sec.color, section_index=idx, kind="hash", is_hash=True,
                           certain=certain, code=code, locale=locale, palette=palette)
        txt, fill = res
        if auto_minus and v < 0:
            txt = "-" + txt
        return _finish(txt, orig, color=sec.color, section_index=idx, kind=kind, fill=fill,
                       certain=certain, code=code, locale=locale, palette=palette)
    if kind == "general":
        txt, fill = _render_general_section(toks, a)
    elif kind == "scientific":
        txt, fill = _render_scientific(toks, a)
    elif kind == "fraction":
        txt, fill = _render_fraction(toks, a)
    elif kind == "number":
        txt, fill = _render_fixed(toks, a)
    elif kind == "text":                        # '@' inside a numeric section: number as General
        txt, fill = _render_literal(toks, general_text(a))
    else:                                       # literal-only section ("-", "n/a")
        txt, fill = _render_literal(toks, None)
    if auto_minus and v < 0:
        txt = "-" + txt
    return _finish(txt, orig, color=sec.color, section_index=idx, kind=kind, fill=fill, certain=certain,
                   code=code, locale=locale, palette=palette)


# --------------------------------------------------------------------------- conveniences
def is_date_format(code_or_id, custom_formats: Optional[dict] = None) -> bool:
    code = resolve_format(code_or_id, custom_formats) if isinstance(code_or_id, int) else code_or_id
    return parse_format(code or "General").is_date


def zero_hidden_by_show_zeros_off(code_or_id, custom_formats: Optional[dict] = None) -> Optional[bool]:
    """On a sheet with showZeros="0" (Options > "Show a zero in cells that have zero value"
    off), does Excel hide an exact numeric 0 under this format?  Measured by Patrick in Excel
    2026-10-03 (excel_session_answers.md Q5): General and single-section formats go blank; a
    format with an explicit zero section still prints that section (#,##0;(#,##0);"-" -> "-",
    0.00;-0.00;0.00 -> "0.00", General;General;"zero" -> "zero").
      True   hidden: General, a lone '@', one or two numeric sections without conditions (zero
             uses the first, positive section: no zero section of its own)
      False  shown: three numeric sections without conditions (the third is the zero section)
      None   unmeasured: a format with conditional sections ([=0], [>0] ...)
    Two-section formats are not in the measured set; they are read like single-section ones
    (no explicit zero section)."""
    code = resolve_format(code_or_id, custom_formats) if isinstance(code_or_id, int) else code_or_id
    f = parse_format(code if code is not None else "General")
    if f.general_numbers or not f.numeric:
        return True
    if f.has_conditions:
        return None
    return len(f.numeric) < 3


def zero_display(code_or_id, custom_formats: Optional[dict] = None) -> str:
    """How an exact 0 shows: 'dash' | 'blank' | 'digit' | 'other' | 'hash'."""
    r = render(0, code_or_id, custom_formats=custom_formats)
    if r.is_hash:
        return "hash"
    if r.is_blank:
        return "blank"
    if r.zero_as_dash:
        return "dash"
    if any(ch.isdigit() for ch in r.text):
        return "digit"
    return "other"
