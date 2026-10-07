"""94 Potential Dangers/No white-on-white hiding.

Rule (handoff bucket 2, Patrick 2026-10-02; full write-up in docs/checks/94.md):

A POPULATED cell fails when its content cannot be seen:
  (a) its text colour is effectively the same as its OWN background: APCA lightness contrast
      |Lc(text, background)| < CONCEAL_LC (12), unless the two colours clearly differ in hue
      (HUE_EXEMPT: green text on a red fill is visible).  Text colour = a firing conditional
      format's font colour, else the number format's colour tag for the value's section
      ([White], [Color2] ...; text and logicals use only the text section's tag, errors none),
      else the font colour (theme 0 resolves to white; rich-text runs are judged one by one).
      Precedence measured in Excel 2026-10-03.  Background = the cell's own fill
      (solid fgColor, hatch patterns blended, gradients: concealed only against EVERY stop
      and their average); no fill = the sheet's white.  Never compared with white as such:
      white text on a dark fill is visible.
  (b) its number format prints nothing for its value (';;;', an empty section such as the zero
      section of '0;-0;;@', '#,###' on 0, a space-only literal section): numfmt.render(...).is_blank.
Populated = the displayed value is not empty: a number, boolean, error, or text with a
non-space character.  Empty cells (no value, or a formula returning "") never count.
Conditional formatting counts where it fires: when some combination of the rules' dxf font
colours / fills / number formats could flip a cell's verdict, the rules (including format-less
'Stop If True' rules, which block lower ones) are evaluated per cell in a second pass over that sheet
with the shared evaluator (checks/_cfeval.py: cellIs with literal or formula operands, text rules,
blanks / errors, and expression rules - comparisons, arithmetic, &, text / IS / logical / rounding
functions, absolute and relative references to the cell itself, to other cells of the sheet and to
other sheets, read through the value source).  Patrick 2026-10-05, "CONDITIONAL FORMATS never stop
a grading" / "assume the built-in Excel conditional formats always have visible text":
  * built-in conditional formats never count as hiding text: colour scales, data bars, icon sets and
    Excel's preset highlight styles (Light Red Fill with Dark Red Text FFC7CE/9C0006, Yellow Fill with
    Dark Yellow Text FFEB9C/9C6500, Green Fill with Dark Green Text C6EFCE/006100, Light Red Fill, Red
    Text, Red Border) - a cell whose font or fill comes from one is visible (its number format can still
    blank it);
  * a rule the evaluator cannot evaluate (COUNTIF, top10, aboveAverage, an untrusted referenced value
    ...), and any other way a conditional format leaves the verdict open (a CF font over rich-text
    runs, Excel's reading of a CF number format): the text is assumed VISIBLE, never a GradingError;
    every such assumption is recorded in stats.cf_assumptions (count + examples).
Formula values are read only for cells whose formatting could conceal them; an untrusted value
there makes the cell UNDECIDED (not a conditional-format matter).  Undecided cells raise
GradingError only when nothing else is certainly concealed (otherwise the verdict is fail either
way and they are listed in stats).

Not counted (see docs): fonts below a size, merged non-anchor values, data bars / icon sets
with showValue=0, a zero hidden by the sheet option showZeros=0 (General / single-section
formats; a format with an explicit zero section still prints it and is judged as usual -
Excel 2026-10-03; recorded in stats), table-style and pivot-style
formatting (not resolved: a colour-concealed cell inside a styled table, and dark directly
formatted text inside a TableStyleDark* table, are UNDECIDED).
"""
from __future__ import annotations

import itertools
import math
import re
from array import array
from collections import defaultdict
from functools import lru_cache
from typing import Optional

from dataclasses import dataclass

from ..core import numfmt as N
from ..core.refs import location, parse_range, range_to_str
from ..core.sheet import ExcelError
from ..core.styles import is_unknown
from ..errors import GradingError
from ._cfeval import (UNKNOWN, Unevaluable, cellis_fires, cf_env, expression_fires, is_unevaluable, literal,
                      prefetch, preset_style)
from ._cftext import text_rule_fires
from ._fills import DEFAULT_FONT, dxf_fill_paints, known_paint, paint_unknowns
from .base import Check

# ---------------------------------------------------------------------------- tunables
CONCEAL_LC = 12.0          # |APCA Lc| below this = text effectively the same as its background
HUE_EXEMPT = True          # ... unless the colours clearly differ in hue (prototype v5b 'chroma exemption'):
EQUILUM_RATIO = 1.1        #   never when the WCAG luminance ratio is below this (yellow on white: 1.07)
SAME_COLOUR_DE = 20.0      #   CIEDE2000 difference at least this
CHROMA_MIN_L = 40.0        #   neither colour dark (CIELAB L*): hue differences collapse at low lightness
CHROMA_MIN_C = 40.0        #   at least one colour clearly chromatic (CIELAB C*ab)
DARK_TABLE_PROXY = "404040"  # TableStyleDark* bodies are dark; directly formatted text invisible on this is undecided
MAX_UNDECIDED_IN_STATS = 10
ZERO_BLANK_COUNTS = True   # a zero rendered blank by its number format ('0;-0;;@') counts (handoff: any value)
SHOW_ZEROS_OFF_COUNTS = False  # sheet option showZeros=0 is not a number format: recorded, not counted
CF_COUNTS = True           # conditional formatting counts where it fires (evaluated per cell)
CF_BUILTIN_VISIBLE = True  # Patrick 2026-10-05: built-in CFs (colour scales, data bars, icon sets, preset styles)
                           # never count as hiding text
CF_ASSUME_VISIBLE = True   # Patrick 2026-10-05: where a conditional format leaves the verdict open (a rule that
                           # cannot be evaluated, ...), the text is assumed visible - recorded in stats
MAX_UNKNOWN_RULES = 8      # a cell under more unevaluable CF rules than this: assumed visible at once (recorded)
MAX_FLIP_COMBOS = 4000     # more reachable CF outcomes on a sheet than this: skip the flip pre-test, do the CF pass
MAX_ASSUMPTION_EXAMPLES = 10
WHITE = "FFFFFF"


@dataclass(frozen=True)
class Visible:
    """The text colour of a conditional-format outcome whose font or fill comes from a BUILT-IN
    conditional format (colour scale, preset highlight style): assumed visible (Patrick 2026-10-05).
    font = the colour the text would otherwise have (None: the cell's own), label = which built-in."""
    font: object
    label: str

# ---------------------------------------------------------------------------- APCA
_R, _G, _B = 0.2126729, 0.7151522, 0.0721750


def _lum(hex6: str) -> float:
    r, g, b = (int(hex6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return _R * r ** 2.4 + _G * g ** 2.4 + _B * b ** 2.4


def _clamp(y: float) -> float:
    return y if y > 0.022 else y + (0.022 - y) ** 1.414


@lru_cache(maxsize=65536)
def apca_lc(text: str, bg: str) -> float:
    """APCA-W3 0.0.98G-4g lightness contrast Lc (polarity-aware, ~ -108..106) of text on bg
    ('RRGGBB' or 'FFRRGGBB').  Same formula as clusters/hidden/final/apca.py."""
    yt, yb = _clamp(_lum(text[-6:])), _clamp(_lum(bg[-6:]))
    if abs(yb - yt) < 0.0005:
        return 0.0
    if yb > yt:                                    # dark text on a light background
        s = (yb ** 0.56 - yt ** 0.57) * 1.14
        out = 0.0 if s < 0.1 else s - 0.027
    else:                                          # light text on a dark background
        s = (yb ** 0.65 - yt ** 0.62) * 1.14
        out = 0.0 if s > -0.1 else s + 0.027
    return round(out * 100.0, 2)


def _lin(c: int) -> float:
    c = c / 255.0
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


@lru_cache(maxsize=65536)
def _lab(hex6: str) -> tuple:
    r, g, b = (_lin(int(hex6[i:i + 2], 16)) for i in (0, 2, 4))
    x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / 0.95047
    y = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b
    z = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / 1.08883
    f = lambda t: t ** (1 / 3) if t > 216 / 24389 else (24389 / 27 * t + 16) / 116   # noqa: E731
    fx, fy, fz = f(x), f(y), f(z)
    return 116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)


@lru_cache(maxsize=65536)
def _de2000(h1: str, h2: str) -> float:
    """CIEDE2000 difference (same formula as clusters/hidden/final/colordiff.py)."""
    L1, a1, b1 = _lab(h1)
    L2, a2, b2 = _lab(h2)
    C1, C2 = math.hypot(a1, b1), math.hypot(a2, b2)
    Cb = (C1 + C2) / 2
    G = 0.5 * (1 - math.sqrt(Cb ** 7 / (Cb ** 7 + 25 ** 7)))
    a1p, a2p = (1 + G) * a1, (1 + G) * a2
    C1p, C2p = math.hypot(a1p, b1), math.hypot(a2p, b2)
    h1p = math.degrees(math.atan2(b1, a1p)) % 360 if C1p else 0.0
    h2p = math.degrees(math.atan2(b2, a2p)) % 360 if C2p else 0.0
    dLp, dCp = L2 - L1, C2p - C1p
    if C1p * C2p == 0:
        dhp = 0.0
    else:
        dhp = h2p - h1p
        if dhp > 180:
            dhp -= 360
        elif dhp < -180:
            dhp += 360
    dHp = 2 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp / 2))
    Lbp, Cbp = (L1 + L2) / 2, (C1p + C2p) / 2
    if C1p * C2p == 0:
        hbp = h1p + h2p
    elif abs(h1p - h2p) <= 180:
        hbp = (h1p + h2p) / 2
    else:
        hbp = (h1p + h2p + 360) / 2 if h1p + h2p < 360 else (h1p + h2p - 360) / 2
    T = (1 - 0.17 * math.cos(math.radians(hbp - 30)) + 0.24 * math.cos(math.radians(2 * hbp))
         + 0.32 * math.cos(math.radians(3 * hbp + 6)) - 0.20 * math.cos(math.radians(4 * hbp - 63)))
    dth = 30 * math.exp(-((hbp - 275) / 25) ** 2)
    Rc = 2 * math.sqrt(Cbp ** 7 / (Cbp ** 7 + 25 ** 7))
    Sl = 1 + 0.015 * (Lbp - 50) ** 2 / math.sqrt(20 + (Lbp - 50) ** 2)
    Sc = 1 + 0.045 * Cbp
    Sh = 1 + 0.015 * Cbp * T
    Rt = -math.sin(math.radians(2 * dth)) * Rc
    return math.sqrt((dLp / Sl) ** 2 + (dCp / Sc) ** 2 + (dHp / Sh) ** 2 + Rt * (dCp / Sc) * (dHp / Sh))


def _wcag_ratio(a: str, b: str) -> float:
    la, lb = _lum_srgb(a[-6:]), _lum_srgb(b[-6:])
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _lum_srgb(hex6: str) -> float:
    r, g, b = (_lin(int(hex6[i:i + 2], 16)) for i in (0, 2, 4))
    return 0.2126729 * r + 0.7151522 * g + 0.0721750 * b


@lru_cache(maxsize=65536)
def hue_exempt(text: str, bg: str) -> bool:
    """Low lightness contrast is still visible when the two colours clearly differ in hue, are not dark and
    at least one is clearly chromatic (green text on a red fill, [Cyan] or [Green] text on white).
    Equiluminant pairs (yellow on white) and dark pairs (black on navy, blue on navy) are not exempt.
    Prototype v5b (clusters/hidden/final/hidden_checks.chroma_exempt)."""
    if not HUE_EXEMPT:
        return False
    t, b = text[-6:], bg[-6:]
    if _wcag_ratio(t, b) < EQUILUM_RATIO:
        return False
    lt, lb = _lab(t), _lab(b)
    ct, cb = math.hypot(lt[1], lt[2]), math.hypot(lb[1], lb[2])
    return _de2000(t, b) >= SAME_COLOUR_DE and min(lt[0], lb[0]) >= CHROMA_MIN_L and max(ct, cb) >= CHROMA_MIN_C


def conceals(text, bgs) -> Optional[bool]:
    """Text colour is invisible against every background candidate: True / False, or None when an
    unresolvable colour (core.styles.UnknownColour, finding 47-R2-08) decides it - an unknown text colour,
    or unknown backgrounds when every known one conceals.  Callers treat None as 'may conceal'.
    A Visible text colour (a built-in conditional format, Patrick 2026-10-05) never conceals."""
    if isinstance(text, Visible):
        return False
    if is_unknown(text):
        return None
    unknown_bg = False
    for b in bgs:
        if is_unknown(b):
            unknown_bg = True
        elif not (abs(apca_lc(text, b)) < CONCEAL_LC and not hue_exempt(text, b)):
            return False
    return None if unknown_bg else True


def _hex6(c):
    """'FFRRGGBB' / 'RRGGBB' -> 'RRGGBB'; an UnknownColour passes through unchanged."""
    return c if is_unknown(c) else c[-6:]


def _cname(c) -> str:
    """A colour for messages: 'RRGGBB', or the unresolvable reference and why."""
    return str(c) if is_unknown(c) else c[-6:]


# ---------------------------------------------------------------------------- CF evaluation
_DARK_TABLE_RX = re.compile(r"^TableStyleDark\d+$", re.I)
parse_operand = literal            # a CF operand as a literal value, or UNKNOWN (kept for callers / tests)


def eval_rule(rule, value, where=None, env=None):
    """True / False when the rule fires / does not fire for this (populated) value; UNKNOWN when it
    depends on what the caller does not know (a position-dependent rule without `where`); an
    Unevaluable when it cannot be evaluated (checks/_cfeval.py).  where = (row, col, anchor_row,
    anchor_col); env = the sheet's _cfeval.CfEnv (other cells' values; None: none can be read)."""
    t = rule.type
    is_err = isinstance(value, ExcelError)
    if t == "cellIs":
        return cellis_fires(rule, value, env, where)
    if t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        needle = rule.text
        if needle is None:
            return Unevaluable(f"{t} rule without its text")
        if is_err:
            return t == "notContainsText"
        return text_rule_fires(t, needle, _as_text(value))     # containsText: SEARCH wildcards (Excel 2026-10-03)
    if t == "containsBlanks":
        return False                             # only populated cells are judged
    if t == "notContainsBlanks":
        return True
    if t == "containsErrors":
        return is_err
    if t == "notContainsErrors":
        return not is_err
    if t == "expression":
        if len(rule.formulas) != 1:
            return Unevaluable(f"expression rule with {len(rule.formulas)} formulas")
        return expression_fires(rule.formulas[0], env, where, value)
    return Unevaluable(f"rule type {t!r} is not evaluated")


def _as_text(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return N.general_text(float(v))
    return str(v)


def _rule_text(rule) -> str:
    what = rule.type or "?"
    if rule.formulas:
        what += " " + ", ".join(repr(f) for f in rule.formulas[:2])
    elif rule.text is not None:
        what += f" {rule.text!r}"
    return what


# ---------------------------------------------------------------------------- formatting model
class _Style:
    """Per style index: font colour, background candidates, number format, risk flags; defaulted = why an
    unresolvable colour of the style was read as Excel's default (None: none was)."""
    __slots__ = ("font", "bgs", "bg_desc", "fmt", "fmt_err", "colour_conceal", "fmt_risky", "risky", "no_own_fill",
                 "direct_font", "defaulted")


def _bg_of_paint(p) -> tuple[list, str]:
    if p is None or p.kind == "none" or p.effective is None:
        return [WHITE], "no fill, i.e. the white sheet"
    if p.kind == "gradient":
        return ([_hex6(c) for c in p.stops] + ([] if is_unknown(p.effective) else [p.effective[-6:]]),
                "gradient " + "→".join(_cname(c) for c in p.stops))
    if p.pattern and p.pattern != "solid":
        return [_hex6(p.effective)], f"{p.pattern} pattern ≈{_cname(p.effective)}"
    return [_hex6(p.effective)], f"fill {_cname(p.effective)}"


_NO_EFFECT = (None, None, None, None)     # (font, bgs, bg_desc, fmt): no conditional format applies
_UNDECIDED = object()                     # a cell this check cannot decide (recorded by _undecided)
_PACK = 16384                             # (row, col) -> row * _PACK + col in compact arrays


def _sig_unknown(sig) -> bool:
    """A verdict signature (font concealed, format profile) that an unresolvable colour decides."""
    fcon, prof = sig
    return fcon is None or (len(prof) == 3 and any(p is None for p in prof[2]))


def _populated(value) -> bool:
    """The displayed value is not empty: numbers, booleans, errors, text with a non-space char."""
    if value is None:
        return False
    if isinstance(value, str) and not isinstance(value, ExcelError):
        return bool(value.strip())
    return True


class C94(Check):
    number = 94
    key = "Potential Dangers/No white-on-white hiding"
    needs_cells = True
    needs_values = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        self.st = wb.styles
        self.palette = list(self.st.indexed_palette) if self.st.custom_indexed else None
        self._tags_defaulted: set = set()       # [ColorN] tags read as black (recorded once each)
        self._styles: dict = {}
        self._fmt_cache: dict = {}
        self.rich_idx = self._rich_coloured_sst()
        self.n_judged = 0
        self.n_values_read = 0
        self.n_cf_cells = 0
        self.n_concealed = 0
        self.show_zeros_off: list = []
        self.zeros_off = False
        self.n_zero_hidden_by_sheet = 0         # zeros hidden by showZeros=0 (format without a zero section)
        self.hidden_value_bars: list = []
        self.cf_second_pass: list = []
        self.examples: dict = {}
        self.pending: dict = {}
        self.cf_assumed = 0                     # cells assumed visible: a conditional format left the verdict open
        self.cf_assumed_examples: list = []
        self.cf_unevaluable: dict = {}          # rule (sheet, sqref, text) -> why it cannot be evaluated
        self.cf_builtin_visible = 0             # cells a built-in conditional format would conceal: visible
        self.cf_builtin_examples: list = []
        self._builtin_seen: set = set()
        self.und_total = 0                      # cells this check cannot decide (committed sheets)
        self.und_examples: list = []            # their first reasons
        self.und_cells = array("q")             # current sheet: packed (r, c) of undecided cells
        self.und_reasons: dict = {}             # current sheet: packed -> reason (first few)

    # -------------------------------------------------------------- formatting model
    def _rich_coloured_sst(self) -> set:
        """Shared-string indexes (as their <v> text) with a run that sets its own font colour."""
        out = set()
        for i, runs in self.wb.shared_string_runs.items():
            if any(r.font is not None and r.font.color is not None for r in runs):
                out.add(str(i))
        return out

    def _fmt_risky(self, code: str, bgs: tuple, colours: bool = True) -> bool:
        """Can this format print nothing, or print in a colour invisible on bgs, for SOME value?
        (Then the cell's value must be read to decide.)  colours=False: a conditional-format font
        colour applies, which beats the format's colour tags (Excel 2026-10-03), so only blanks count."""
        key = (code, bgs, colours)
        v = self._fmt_cache.get(key)
        if v is None:
            f = N.parse_format(code)
            v = not f.verified
            for sec in f.sections:
                if sec.is_empty:
                    v = True
                elif sec.color is not None and colours:
                    rgb = self._tag_rgb(sec.color)
                    if rgb and conceals(rgb, bgs) is not False:
                        v = True
            if not v:   # sections that render blank for some values ('#,###' on 0, '?' digits, space literals)
                probes = [0.0, 1e-9, 0.4, -1e-9, -0.4, 1.0, -1.0]
                for sec in f.sections:          # conditional sections: probe around each threshold ('[>2]"";0')
                    if sec.condition is not None:
                        t = float(sec.condition[1])
                        probes += [t, t + 1, t - 1, t + 0.5, t - 0.5, t * 2, t * 10, t / 2, -t, -t * 10]
                v = any(N.render(p, code, indexed_palette=self.palette).is_blank for p in probes) or \
                    N.render("x", code, indexed_palette=self.palette).is_blank
            self._fmt_cache[key] = v
        return v

    def _tag_rgb(self, tag):
        """RRGGBB of a number-format colour tag ([Red], [Color10] ...), None when it names no colour.  A [ColorN]
        whose custom <indexedColors> entry is not a colour is read as Excel's default for a font colour,
        automatic = black (Patrick 2026-10-05: every attempt graded; recorded once per tag)."""
        if tag and self.st.unknown_indexed:
            m = re.match(r"^colou?r\s*(\d{1,2})$", tag.strip(), re.I)
            if m and 1 <= int(m.group(1)) <= 56 and (int(m.group(1)) + 7) in self.st.unknown_indexed:
                c = self.st.indexed_colour(int(m.group(1)) + 7)
                if is_unknown(c):
                    if tag not in self._tags_defaulted:
                        self._tags_defaulted.add(tag)
                        self.note_default("unresolved_colour", f"number-format colour tag [{tag}]: {c} read as "
                                                               f"automatic (black)")
                    return DEFAULT_FONT[-6:]
                return c
        return N.color_tag_rgb(tag, self.palette)

    def _known(self, c, why: list, what: str):
        """An unresolvable font colour read as automatic (black), the reason appended to `why`."""
        if is_unknown(c):
            why.append(f"{what} {c} read as automatic (black)")
            return DEFAULT_FONT[-6:]
        return _hex6(c)

    def _style(self, s: int) -> _Style:
        x = self._styles.get(s)
        if x is None:
            x = _Style()
            why = []
            # Patrick 2026-10-05 (every attempt graded): an unresolvable colour is read as Excel's default for its
            # slot - the font colour as automatic (black), the fill as none (recorded per cell in cell())
            x.font = self._known(self.st.font_color(s), why, "font colour")
            raw_fill = self.st.cell_fill(s)
            fill = known_paint(raw_fill)
            if fill is not raw_fill:
                why.append(f"fill colour {', '.join(str(k) for k in paint_unknowns(raw_fill))} read as no fill")
            x.defaulted = "; ".join(why) or None
            x.bgs, x.bg_desc = _bg_of_paint(fill)
            x.bgs = tuple(x.bgs)
            x.fmt_err = None
            try:
                x.fmt = N.resolve_format(self.st.num_fmt_id(s), self.st.num_fmts)
                N.parse_format(x.fmt)            # an unreadable code raises here, not at a cell
            except GradingError as e:
                # Patrick 2026-10-06 (attempt 2777, measured in Excel): an unreadable format is General
                x.fmt, x.fmt_err = "General", str(e)
            x.colour_conceal = conceals(x.font, x.bgs)        # None: an unresolvable colour decides
            x.fmt_risky = self._fmt_risky(x.fmt, x.bgs)
            x.risky = x.colour_conceal is not False or x.fmt_risky
            x.no_own_fill = fill.effective is None
            x.direct_font = self.st.xf(s).font_id != self.st.xf(0).font_id
            self._styles[s] = x
        return x

    def _state(self, st: _Style, eff) -> tuple:
        """(colour concealed, format risky) of a style under one CF effect, value-independent.
        (False, False) = visible for every value; used by the second-pass screen only."""
        font, bgs, _d, fmt = eff
        bgs = st.bgs if bgs is None else bgs
        fmt = st.fmt if fmt is None else fmt
        return (conceals(st.font if font is None else font, bgs), self._fmt_risky(fmt, bgs, colours=font is None))

    def _fmt_profile(self, fmt: str, bgs: tuple, tags: bool, fcon: bool) -> tuple:
        """What a number format adds to the verdict, for EVERY value (review 2026-10-04), given whether
        the font colour conceals (fcon) and whether the format's colour tags apply (no CF font):
          ('plain', z)    it never prints nothing and no tag changes the outcome: the font decides
                          (z = whether showZeros=0 hides a zero under it, on such sheets only);
          ('blank', fmt)  the font is visible and no tag conceals: only blanking matters;
          (fmt, tags, which tags conceal on bgs)  otherwise.
        Two outcomes with equal font concealment and equal profiles give the same verdict for every
        value; the coarse 'risky' flag does not (a CF ';;;' over '0;-0;;@')."""
        key = ("profile", fmt, bgs, tags, fcon, self.zeros_off)
        v = self._fmt_cache.get(key)
        if v is None:
            f = N.parse_format(fmt)
            pattern = ()
            if tags:
                # True / False per tag, None where an unresolvable colour decides (never 'visible')
                pattern = tuple((conceals(rgb, bgs) if rgb else False) for rgb in
                                (self._tag_rgb(sec.color) for sec in f.sections if sec.color is not None))
            blanks = self._fmt_risky(fmt, bgs, colours=False)       # can print nothing for some value
            z = N.zero_hidden_by_show_zeros_off(fmt) if self.zeros_off else None
            if fcon is False and not any(p is not False for p in pattern):      # colour never conceals
                v = ("blank", fmt) if blanks else ("plain", None)
            elif fcon is True and not pattern and not blanks:       # the (concealing) font always shows
                v = ("plain", z)
            else:
                v = (fmt, tags, pattern)
            self._fmt_cache[key] = v
        return v

    def _sig(self, st: _Style, eff) -> tuple:
        """The verdict of a plain (non-rich) cell of this style under one CF outcome, as a function of
        its value: (font colour concealed, format profile).  Equal signatures = equal verdicts."""
        font, bgs, _d, fmt = eff
        bgs = st.bgs if bgs is None else bgs
        fmt = st.fmt if fmt is None else fmt
        fcon = conceals(st.font if font is None else font, bgs)
        return (fcon, self._fmt_profile(fmt, bgs, font is None, fcon))

    def _runs_colours(self, cell, base_font) -> Optional[list]:
        """Colours of the visible rich-text runs (None for plain cells).  A run without its own
        colour shows the cell's font colour; an unresolvable run colour is read as automatic (black)
        (Patrick 2026-10-05: every attempt graded; recorded)."""
        runs = cell.rich_runs
        if not runs:
            return None
        out = []
        why = []
        for run in runs:
            if not (run.text or "").strip():
                continue
            if run.font is not None and run.font.color is not None:
                out.append(self._known(self.st.resolve(run.font.color, "font"), why, "rich-text run colour"))
            else:
                out.append(base_font)
        if why:
            self.note_default("unresolved_colour", f"{location(cell.sheet, cell.ref)}: {'; '.join(why)}")
        return out or None

    def _is_rich(self, cell) -> bool:
        if cell.t == "s":
            return cell.raw in self.rich_idx
        if cell.t == "inlineStr":
            runs = cell.rich_runs
            return bool(runs) and any(r.font is not None and r.font.color is not None for r in runs)
        return False

    # -------------------------------------------------------------- the decision
    def _blank_alternatives(self, value, fmt: str, vt) -> set:
        """Whether the format prints nothing for this value.  Every reading is measured in Excel
        since 2026-10-03 (logicals go through the text section, '' = General, spaces-only codes
        print nothing), so this is a single outcome - except under a code the engine marks
        unverified (certain=False, e.g. '0 days': unquoted date letters among digit placeholders),
        where both outcomes are possible."""
        r = N.render(value, fmt, custom_formats=self.st.num_fmts, value_type=vt,
                     date1904=self.wb.date1904, indexed_palette=self.palette)
        return {r.is_blank}          # an unverified code too: the engine's best rendering (Patrick 2026-10-05)

    def _decide(self, cell, value, st: _Style, runs, effects, cf_note: Optional[str] = None) -> Optional[str]:
        """Why the populated cell is concealed, or None when it is visible.  effects: alternative
        conditional-format outcomes [(font, bgs, bg_desc, fmt)].  Every alternative interpretation must
        agree, otherwise the cell is undecided (no guessing) - EXCEPT where the conditional formatting
        alone leaves it open (the cell's own formatting decides it, the CF outcomes disagree: a rule that
        cannot be evaluated, a CF font over rich-text runs, a CF number format or colour Excel's reading
        of which is not verified): then the text is assumed visible (Patrick 2026-10-05) and the
        assumption recorded.  cf_note: why the CF is open (the unevaluable rules), for the record."""
        if not _populated(value):
            return None
        outcomes = self._outcomes(cell, value, st, runs, effects, count=True)
        if len(outcomes) <= 1:
            return next(iter(outcomes.values()), None)
        if CF_ASSUME_VISIBLE and effects != [_NO_EFFECT]:
            base = self._outcomes(cell, value, st, runs, [_NO_EFFECT], count=False)
            if len(base) == 1:
                # the cell's own formatting decides; only the conditional formatting leaves it open
                self._assume_visible(cell, value, cf_note or (
                    "the conditional-format outcomes disagree and Excel's rendering of the one that conceals is "
                    f"not verified ({outcomes[True]})"))
                return None
        self._undecided(cell, f"cannot decide whether {cell.sheet}!{cell.ref} (value {value!r}, format {st.fmt!r}) "
                              f"is concealed: the outcome depends on an unresolvable colour or Excel rendering this "
                              f"check cannot determine ({outcomes[True]})")
        return _UNDECIDED

    def _assume_visible(self, cell, value, why: str):
        """Record a cell assumed visible because a conditional format left its verdict open."""
        self.cf_assumed += 1
        if len(self.cf_assumed_examples) < MAX_ASSUMPTION_EXAMPLES:
            self.cf_assumed_examples.append(f"{location(cell.sheet, cell.ref)} = {value!r}: {why}"[:300])

    def _outcomes(self, cell, value, st: _Style, runs, effects, count: bool = True) -> dict:
        """{concealed?: reason} over every interpretation of the cell under the given CF outcomes
        (_decide).  count=False: no statistics are updated (a second look at the same cell)."""
        vt = "e" if isinstance(value, ExcelError) else ("d" if cell.t == "d" and not cell.is_formula_result else None)
        outcomes: dict = {}
        # showZeros=0 (Excel 2026-10-03): an exact numeric 0 is hidden by the sheet option unless its
        # format has an explicit zero section, which is then printed (and judged) as usual
        zero_hidden_by_sheet = None
        if self.zeros_off and isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0 \
                and not SHOW_ZEROS_OFF_COUNTS:
            zero_hidden_by_sheet = True
        for cf_font, cf_bgs, cf_bg_desc, cf_fmt in effects:
            fmt = cf_fmt if cf_fmt is not None else st.fmt
            bgs, bg_desc = (cf_bgs, cf_bg_desc) if cf_bgs is not None else (st.bgs, st.bg_desc)
            r = N.render(value, fmt, custom_formats=self.st.num_fmts, value_type=vt,
                         date1904=self.wb.date1904, indexed_palette=self.palette)
            blanks = self._blank_alternatives(value, fmt, vt)
            primary = ("rich-text runs", runs) if runs else ("font", [st.font])
            alts = []
            tag = self._tag_rgb(r.color_tag) if r.color_tag else None    # the colour tag of the section used
            #                                      (text / logicals: the text section's only; errors: none) -
            #                                      Excel 2026-10-03; an UnknownColour for an invalid palette entry
            if isinstance(cf_font, Visible):
                # a built-in conditional format paints the font or the fill: never hides text (Patrick
                # 2026-10-05); recorded when the text would otherwise be concealed
                alts.append((cf_font.label, [cf_font]))
                if count:
                    real = [cf_font.font] if cf_font.font is not None else (runs or [st.font])
                    if any(conceals(c, (b,)) is not False for c in real for b in bgs):
                        self._note_builtin(cell, value, cf_font.label, bg_desc)
            elif cf_font is not None:
                # a conditional-format font colour beats a number-format colour tag (Excel 2026-10-03)
                alts.append(("conditional-format font", [cf_font]))
                if runs:                         # CF font over rich-text runs: unverified
                    alts.append(primary)
            elif tag is not None:
                # (an unverified code included: the engine's best rendering, the tag of the section it used)
                alts.append(("number-format colour", [tag]))
                if runs and count:
                    # number-format colour vs rich-text runs is not verified: the engine's best rendering, the
                    # tag applies as on a plain cell (Patrick 2026-10-05: every attempt graded)
                    self.note_default("unverified_number_format", f"{location(cell.sheet, cell.ref)}: number-format "
                                                                  f"colour over rich-text runs, the tag applies")
            else:
                alts.append(primary)
            if not r.certain and count:
                # Excel's reading of this format is unverified (numfmt certain=False): the engine's best rendering
                # (Patrick 2026-10-05: every attempt graded)
                self.note_default("unverified_number_format", f"{location(cell.sheet, cell.ref)} = {value!r} under "
                                                              f"{fmt!r}: graded as rendered ({r.text!r})")
            if zero_hidden_by_sheet is not None:
                # showZeros=0 and the format has no zero section: the sheet option hides the zero
                # (not counted, SHOW_ZEROS_OFF_COUNTS); None = unmeasured (conditional format)
                hz = N.zero_hidden_by_show_zeros_off(fmt) if zero_hidden_by_sheet else False
                if hz is None:
                    # unmeasured (a format with conditional sections): the engine's best rendering, the format's
                    # own display, is judged as usual (Patrick 2026-10-05: every attempt graded)
                    if count:
                        self.note_default("unverified_number_format", f"{location(cell.sheet, cell.ref)}: zero under "
                                                                      f"{fmt!r} on a showZeros=0 sheet judged as shown")
                elif hz:
                    if count:
                        self.n_zero_hidden_by_sheet += 1
                    outcomes.setdefault(False, None)
                    continue
            for blank in blanks:
                for label, cols in alts:
                    states = [conceals(c, bgs) for c in cols]
                    hidden = [c for c, x in zip(cols, states) if x is True]
                    maybe = [c for c, x in zip(cols, states) if x is None]
                    reasons = []
                    if blank:
                        if isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0 \
                                and not ZERO_BLANK_COUNTS:
                            pass
                        elif len(blanks) > 1:            # unverified format: blank is one possible reading
                            why = N.parse_format(fmt).unverified_why or "Excel's reading of it is not verified"
                            reasons.append(f"its number format {fmt!r} may print nothing for its value ({why})")
                        else:
                            reasons.append(f"its number format {fmt!r} prints nothing for its value")
                    if hidden:
                        c0 = hidden[0]
                        lc = min(abs(apca_lc(c0, b)) for b in bgs)
                        reasons.append(f"its text colour {c0} ({label}) on its own background ({bg_desc}) has "
                                       f"APCA contrast Lc {lc:.1f} < {CONCEAL_LC:g}")
                    if reasons or not maybe:
                        outcomes.setdefault(bool(reasons), "; ".join(reasons) or None)
                    else:
                        # an unresolvable text or background colour decides: concealed or not (47-R2-08)
                        t0 = maybe[0]
                        if is_unknown(t0):
                            why = (f"its text colour ({label}) is an {t0}, so whether it shows on its own background "
                                   f"({bg_desc}) is unknown")
                        else:
                            why = (f"its text colour {t0} ({label}) may be invisible on its own background ({bg_desc}), "
                                   f"a colour that cannot be resolved")
                        outcomes.setdefault(True, why)
                        outcomes.setdefault(False, None)
        return outcomes

    def _note_builtin(self, cell, value, label: str, bg_desc: str):
        k = (cell.sheet, cell.row, cell.col)
        if k in self._builtin_seen:
            return
        self._builtin_seen.add(k)
        self.cf_builtin_visible += 1
        if len(self.cf_builtin_examples) < MAX_ASSUMPTION_EXAMPLES:
            self.cf_builtin_examples.append(f"{location(cell.sheet, cell.ref)} = {value!r}: {label} ({bg_desc}) would "
                                            f"conceal its text; assumed visible"[:300])

    def _undecided(self, cell, reason: str):
        k = cell.row * _PACK + cell.col
        self.und_cells.append(k)
        if len(self.und_reasons) < MAX_UNDECIDED_IN_STATS:
            self.und_reasons.setdefault(k, reason)

    def _commit_undecided(self):
        self.und_total += len(self.und_cells)
        for reason in self.und_reasons.values():
            if len(self.und_examples) < MAX_UNDECIDED_IN_STATS:
                self.und_examples.append(reason)
        self.und_cells, self.und_reasons = array("q"), {}

    def _value(self, cell):
        """The cell's value, or _UNDECIDED when it is an untrusted formula result: the cell is skipped
        (Patrick 2026-10-05: every attempt graded; stats.defaults.untrusted_value)."""
        if cell.is_formula_result:
            self.n_values_read += 1
            if not cell.value_trusted:
                self.note_default("untrusted_value", f"{location(cell.sheet, cell.ref)} (source={cell.value_source})")
                return _UNDECIDED
            return cell.value
        return cell.value

    def _judge(self, cell, st: _Style, effects, rich: bool):
        value = self._value(cell)
        if value is not _UNDECIDED and _populated(value):
            self._judge_known(cell, st, effects, rich, value)

    # -------------------------------------------------------------- sheet level (first pass)
    def sheet_start(self, head):
        self.hits = defaultdict(list)            # description -> [(r, c)]
        self.und_cells, self.und_reasons = array("q"), {}
        self.styles_seen: set = set()
        self.has_rich = False
        self.table_boxes = self._styled_tables(head.info)
        self.dark_boxes = [b for b, style, _n in self.table_boxes if _DARK_TABLE_RX.match(style)]
        v = head.view
        self.zeros_off = v is not None and v.show_zeros is False
        if self.zeros_off:
            self.show_zeros_off.append(head.name)

    def _styled_tables(self, info) -> list:
        """[(box, style name, table name)] of tables carrying a table style (tableStyleInfo name)."""
        out = []
        for t in self.wb.tables(info):
            root = self.wb.xml(t.part)
            style = None
            if root is not None:
                for el in root:
                    if el.tag.rsplit("}", 1)[-1] == "tableStyleInfo":
                        style = el.get("name")
            b = parse_range(t.ref or "")
            if style and b:
                out.append((b, style, t.display_name or t.name or "?"))
        return out

    def cell(self, cell):
        if cell.is_blank:
            return
        s = cell.s
        st = self._styles.get(s) or self._style(s)
        self.styles_seen.add(s)
        if st.defaulted:
            self.note_default("unresolved_colour", f"{location(cell.sheet, cell.ref)}: {st.defaulted}")
        if st.fmt_err:
            self.note_default("unreadable_number_format", f"{location(cell.sheet, cell.ref)}: {st.fmt_err}")
        rich = (cell.t == "s" or cell.t == "inlineStr") and self._is_rich(cell)
        if rich:
            self.has_rich = True
        elif not st.risky and not (self.dark_boxes and self._dark_table_risk(cell, st)):
            return
        self._judge(cell, st, [_NO_EFFECT], rich)

    def _dark_table_risk(self, cell, st) -> bool:
        """Dark (or unresolvable), directly formatted text with no fill of its own inside a
        TableStyleDark* table."""
        return st.no_own_fill and st.direct_font and conceals(st.font, (DARK_TABLE_PROXY,)) is not False and \
            any(r1 <= cell.row <= r2 and c1 <= cell.col <= c2 for r1, c1, r2, c2 in self.dark_boxes)

    # -------------------------------------------------------------- conditional formatting
    def _cf_rules(self, tail, sheet) -> list:
        """[(priority, order, rule, ranges, effects, scale, anchor, builtin)] of rules that can change the
        font colour, fill or number format, or that stop lower rules (stopIfTrue, even with no format or a
        border-only one: effects [_NO_EFFECT]).  effects = alternative outcomes [(font, bgs, bg_desc, fmt)]
        when the rule fires (the dxf fill is what Excel paints: bgColor only, see _fills.py); scale =
        colour-scale stop colours or None; anchor = top-left cell of the first range (CF formulas are
        relative to it); builtin = the name of the BUILT-IN conditional format the rule is (a colour scale,
        one of Excel's preset highlight styles) or None: its font / fill never hides text (Patrick
        2026-10-05).  Data bars and icon sets change no font or fill (with showValue=0 they are listed)."""
        out = []
        order = 0
        for cf in tail.conditional_formats:
            if not cf.ranges:
                continue
            anchor = (cf.ranges[0][0], cf.ranges[0][1])
            for rule in cf.rules:
                order += 1
                prio = rule.priority if rule.priority is not None else 10 ** 9
                if rule.type == "colorScale":
                    # a built-in format: the fill it paints never hides text (CF_BUILTIN_VISIBLE); an unresolvable
                    # stop colour is read as no fill, the white sheet (Patrick 2026-10-05: every attempt graded)
                    cols = tuple(WHITE if is_unknown(x) else x[-6:]
                                 for x in (self.st.resolve(c, "fill") or "FFFFFFFF" for c in rule.colors))
                    if cols:
                        out.append((prio, order, rule, cf.ranges, [_NO_EFFECT], cols, anchor, "a colour scale"))
                    continue
                if rule.type in ("dataBar", "iconSet"):
                    if rule.show_value is False:
                        self.hidden_value_bars.append(location(sheet, ",".join(cf.sqref.split())))
                    continue
                d = rule.dxf if rule.dxf is not None else self.st.dxf(rule.dxf_id)
                if d is None:
                    if rule.stop_if_true:
                        # 'Stop If True' with no format (review 2026-10-04): when it fires, lower-priority
                        # rules are not applied, so it is kept for its stop effect (as in check 66)
                        out.append((prio, order, rule, cf.ranges, [_NO_EFFECT], None, anchor, None))
                    continue
                font = None
                why = []
                if d.font is not None and d.font.color is not None:
                    # an unresolvable dxf font colour: automatic (black) (Patrick 2026-10-05: every attempt graded)
                    font = self._known(self.st.resolve(d.font.color, "font"), why, "conditional-format font colour")
                cols, together = dxf_fill_paints(self.st, d.fill)
                if any(is_unknown(x) for x in cols):
                    # an unresolvable dxf fill colour: no fill (a gradient drops that stop) - the cell's own fill shows
                    why.append(f"conditional-format fill colour {', '.join(str(x) for x in cols if is_unknown(x))} "
                               f"read as no fill")
                    cols = [x for x in cols if not is_unknown(x)] if together else []
                if why:
                    self.note_default("unresolved_colour", f"conditional format on "
                                                           f"{location(sheet, ','.join(cf.sqref.split()))}: "
                                                           f"{'; '.join(why)}")
                if not cols:
                    fills = [(None, None)]
                elif together:
                    fills = [(tuple(_hex6(x) for x in cols), "conditional gradient " + "→".join(_cname(x) for x in cols))]
                else:
                    fills = [((_hex6(x),), f"conditional fill {_cname(x)}") for x in cols]
                fmt = None
                if d.num_fmt_code is not None:
                    fmt = d.num_fmt_code
                elif d.num_fmt_id is not None:
                    fmt = N.resolve_format(d.num_fmt_id, self.st.num_fmts)
                if font is None and fills == [(None, None)] and fmt is None and not rule.stop_if_true:
                    continue                      # no visible effect (bold, border ...) and no stop
                builtin = None
                fill_rgb = cols[0] if cols else None
                if CF_BUILTIN_VISIBLE and not together and len(cols) <= 1 and not is_unknown(font) \
                        and not is_unknown(fill_rgb):
                    name = preset_style(font, fill_rgb, fmt is not None)
                    if name:
                        builtin = f"Excel's preset style '{name}'"
                effects = [(font, b, bd, fmt) for b, bd in fills]
                out.append((prio, order, rule, cf.ranges, effects, None, anchor, builtin))
        out.sort(key=lambda x: (x[0], x[1]))
        return out

    @staticmethod
    def _compose(cover, fired, pick):
        """One CF outcome (font, bgs, bg_desc, fmt) of the covering rules in priority order: the first firing
        rule that sets a property wins it; a firing stopIfTrue rule ends the composition.  When a built-in
        conditional format (colour scale, preset style) sets the font or the fill, the outcome's text is
        Visible (Patrick 2026-10-05)."""
        font = bgs = bg_desc = fmt = None
        builtin = None
        for x, f, eff in zip(cover, fired, pick):
            if not f:
                continue
            set_any = False
            if x[5] is not None:                 # colour scale: sets the fill
                if bgs is None:
                    bgs, bg_desc = x[5], "colour-scale fill " + "→".join(_cname(c) for c in x[5])
                    set_any = True
            else:
                if font is None and eff[0] is not None:
                    font = eff[0]
                    set_any = True
                if bgs is None and eff[1] is not None:
                    bgs, bg_desc = eff[1], eff[2]
                    set_any = True
                if fmt is None and eff[3] is not None:
                    fmt = eff[3]
            if set_any and x[7] is not None and builtin is None:
                builtin = x[7]
            if x[2].stop_if_true:
                break
        if builtin is not None and not isinstance(font, Visible):
            font = Visible(font, builtin)
        return (font, bgs, bg_desc, fmt)

    def _combos(self, cover, fired_known=None) -> Optional[list]:
        """All CF outcomes [(font, bgs, bg_desc, fmt)] for the covering rules (priority order, _compose).
        fired_known: True / False / UNKNOWN / an Unevaluable per rule (None = all unknown).  Unknown and
        unevaluable rules branch on firing; rules with alternative effects branch on the alternative.  None
        when more than MAX_UNKNOWN_RULES rules branch."""
        n = len(cover)
        known = fired_known if fired_known is not None else [UNKNOWN] * n
        unknown = [i for i, f in enumerate(known) if f is UNKNOWN or is_unevaluable(f)]
        if len(unknown) > MAX_UNKNOWN_RULES:
            return None
        out = []
        for combo in itertools.product((True, False), repeat=len(unknown)):
            fired = list(known)
            for k, i in enumerate(unknown):
                fired[i] = combo[k]
            choices = [x[4] if f and x[5] is None else [None] for x, f in zip(cover, fired)]
            for pick in itertools.product(*choices):
                out.append(self._compose(cover, fired, pick))
        return list(dict.fromkeys(out))

    def _may_flip(self, rules) -> bool:
        """Cheap value-independent test: could any conditional-format outcome change the verdict of
        any style used by a populated cell of this sheet, for any value?  When it cannot, no second
        pass and no extra value reads are needed.
        Every outcome some set of firing rules can produce (_reachable: rules composed in priority
        order, the first firing rule sets each property, stopIfTrue ends) is compared with the style's
        own formatting by its verdict signature (_sig: font concealment + the format's full profile,
        not a coarse 'risky' flag - review 2026-10-04: a CF ';;;' or '[White]0' over a base '0;-0;;@'
        was missed, and so were a colour-scale fill under a dxf font and orders lost when two rules
        share a dxf).  Rule conditions and ranges are ignored, so this is a superset of what any one
        cell can see; the second pass then decides per cell."""
        if self.has_rich and any(x[5] is not None or any(e[0] is not None or e[1] is not None for e in x[4])
                                 for x in rules):
            return True                          # run colours are not in the style screen below
        reach = self._reachable(rules)
        if reach is None:
            return True
        outcomes = [(f, b, None, m) for f, b, m in reach]
        seen = set()
        for s_ in self.styles_seen:
            st = self._style(s_)
            k = (st.font, st.bgs, st.fmt)
            if k in seen:
                continue
            seen.add(k)
            base = self._sig(st, _NO_EFFECT)
            if _sig_unknown(base):
                return True                      # an unresolvable colour: no value-independent screen
            for e in outcomes:
                sig = self._sig(st, e)
                if sig != base or _sig_unknown(sig):
                    return True
                if isinstance(e[0], Visible):
                    # a built-in format that would conceal this style is evaluated per cell too, so that every
                    # cell it keeps visible is recorded (stats.cf_builtin_visible)
                    real = self._sig(st, (e[0].font, e[1], None, e[3]))
                    if real != base or _sig_unknown(real):
                        return True
        return False

    def _reachable(self, rules) -> Optional[set]:
        """Every (font, bgs, fmt) outcome that some set of firing rules produces, whatever their
        conditions and ranges: rules in priority order, a firing rule fills each property not yet set
        (its alternative effects branch), a built-in format that sets the font or the fill makes the text
        Visible, a firing stopIfTrue rule ends the composition.  None when more than MAX_FLIP_COMBOS
        outcomes."""
        live = {(None, None, None, None)}         # partial outcomes (font, bgs, fmt, built-in) open to lower rules
        done = set()                              # outcomes ended by a firing stopIfTrue rule
        for x in rules:
            if x[5] is not None:
                effs = {(None, x[5], None)}
            else:
                effs = {(e[0], e[1], e[3]) for e in x[4]}
            fired = set()
            for p in live:
                for e in effs:
                    builtin = p[3]
                    if builtin is None and x[7] is not None and \
                            ((p[0] is None and e[0] is not None) or (p[1] is None and e[1] is not None)):
                        builtin = x[7]
                    fired.add((p[0] if p[0] is not None else e[0], p[1] if p[1] is not None else e[1],
                               p[2] if p[2] is not None else e[2], builtin))
            if x[2].stop_if_true:
                done |= fired
            else:
                live |= fired
            if len(live) + len(done) > MAX_FLIP_COMBOS:
                return None
        # as _compose: an outcome a built-in format took part in keeps its text Visible
        return {(Visible(f, b) if b is not None else f, bgs, fmt) for f, bgs, fmt, b in live | done}

    def _cover(self, r: int, c: int) -> list:
        """Rules (priority order) whose ranges contain (r, c); boxes indexed per column."""
        p = self._p
        col = p["by_col"].get(c)
        if col is None:
            col = [(b[0], b[2], i) for i, x in enumerate(p["rules"]) for b in x[3] if b[1] <= c <= b[3]]
            p["by_col"][c] = col
        ids = sorted({i for r1, r2, i in col if r1 <= r <= r2})
        return ids

    def sheet_end(self, head, tail):
        rules = self._cf_rules(tail, head.name) if CF_COUNTS else []
        if rules and self._may_flip(rules):
            boxes = [b for x in rules for b in x[3]]
            keep = defaultdict(list)
            for desc, cells in self.hits.items():
                for r, c in cells:
                    if not any(b[0] <= r <= b[2] and b[1] <= c <= b[3] for b in boxes):
                        keep[desc].append((r, c))
            # undecided cells inside the CF ranges are re-judged with their conditional formatting
            und = array("q", (k for k in self.und_cells if not any(b[0] <= k // _PACK <= b[2] and
                                                                    b[1] <= k % _PACK <= b[3] for b in boxes)))
            kept = set(und) if self.und_reasons else set()
            und_reasons = {k: v for k, v in self.und_reasons.items() if k in kept}
            self.und_cells, self.und_reasons = array("q"), {}
            # the cells the rules' formulas read (same sheet or others): one stream per referenced sheet
            env = cf_env(self.wb, head.name)
            prefetch(env, tail.conditional_formats)
            self.pending[head.name] = {
                "hits": keep, "rules": rules, "tables": self.table_boxes, "by_col": {}, "screen": {},
                "und_cells": und, "und_reasons": und_reasons, "env": env,
                "bbox": (min(b[0] for b in boxes), min(b[1] for b in boxes),
                         max(b[2] for b in boxes), max(b[3] for b in boxes))}
            self.cf_second_pass.append(head.name)
            self.request_second_pass(head.name, None)
        else:
            self._commit(head.name, self.hits)
            self._commit_undecided()
        self.hits = defaultdict(list)

    def _commit(self, sheet, hits):
        for desc, cells in hits.items():
            self.n_concealed += len(cells)
            self.add_cell_mistakes(sheet, cells, "{n} populated cell(s) {range} cannot be seen: " + desc + ".")

    # -------------------------------------------------------------- second pass (CF sheets only)
    def second_pass_start(self, head):
        p = self.pending[head.name]
        self.hits = p["hits"]
        self.table_boxes = p["tables"]
        self.dark_boxes = [b for b, style, _n in self.table_boxes if _DARK_TABLE_RX.match(style)]
        self.und_cells, self.und_reasons = p.pop("und_cells", array("q")), p.pop("und_reasons", {})
        v = head.view
        self.zeros_off = v is not None and v.show_zeros is False
        self._p = p

    def second_pass_cell(self, cell):
        if cell.is_blank:
            return
        p = self._p
        r, c = cell.row, cell.col
        bb = p["bbox"]
        if not (bb[0] <= r <= bb[2] and bb[1] <= c <= bb[3]):
            return
        ids = self._cover(r, c)
        if not ids:
            return
        cover = [p["rules"][i] for i in ids]
        st = self._style(cell.s)
        rich = (cell.t == "s" or cell.t == "inlineStr") and self._is_rich(cell)
        if not rich:
            # value-independent screen: visible under every CF outcome -> no value needed
            key = (cell.s, tuple(ids))
            visible = p["screen"].get(key)
            if visible is None:
                combos = self._combos(cover)
                # (a built-in format's outcome is screened with the colours it masks, so that a cell it keeps
                # visible is evaluated and recorded)
                visible = combos is not None and all(
                    self._state(st, e) == (False, False) and
                    (not isinstance(e[0], Visible) or self._state(st, (e[0].font,) + tuple(e[1:])) == (False, False))
                    for e in combos + [_NO_EFFECT])
                p["screen"][key] = visible
            if visible and not (self.dark_boxes and self._dark_table_risk(cell, st)):
                return
        value = self._value(cell)
        if value is _UNDECIDED or not _populated(value):
            return
        self.n_cf_cells += 1
        fired = []
        open_rules = []
        for x in cover:
            if x[5] is not None:      # colour scale: applies to numbers only
                fired.append(isinstance(value, (int, float)) and not isinstance(value, bool))
            else:
                f = eval_rule(x[2], value, (r, c) + x[6], p["env"])
                if f is UNKNOWN or is_unevaluable(f):
                    why = f.why if is_unevaluable(f) else "it depends on what this check does not know"
                    open_rules.append(f"{_rule_text(x[2])} ({why})")
                    self._note_unevaluable(cell.sheet, x, why)
                fired.append(f)
        cf_note = (f"conditional-format rule(s) that cannot be evaluated: {'; '.join(open_rules)}"
                   if open_rules else None)
        effects = self._combos(cover, fired)
        if effects is None:
            # Patrick 2026-10-05: a conditional format never stops a grading - assumed visible, recorded
            self.n_judged += 1
            self._assume_visible(cell, value, f"covered by more than {MAX_UNKNOWN_RULES} conditional-format rules "
                                              f"that cannot be evaluated ({cf_note})")
            return
        self._judge_known(cell, st, effects, rich, value, cf_note)

    def _note_unevaluable(self, sheet: str, x, why: str):
        """Record a rule met in a judged cell that cannot be evaluated (stats.cf_assumptions)."""
        k = (sheet, x[1])
        if k not in self.cf_unevaluable and len(self.cf_unevaluable) < 1000:
            where = ",".join(range_to_str(*b) for b in x[3][:3]) + (",..." if len(x[3]) > 3 else "")
            self.cf_unevaluable[k] = f"{location(sheet, where)}: {_rule_text(x[2])} ({why})"[:300]

    def _judge_known(self, cell, st, effects, rich, value, cf_note: Optional[str] = None):
        self.n_judged += 1
        runs = self._runs_colours(cell, st.font) if rich else None
        desc = self._decide(cell, value, st, runs, effects, cf_note)
        if desc is _UNDECIDED:
            return
        if desc is None:
            # visible under its own formatting; a dark table style could still paint a dark fill under dark,
            # directly formatted text (Excel's direct formatting wins over the table style's font)
            if self.table_boxes and st.no_own_fill and st.direct_font and all(e[1] is None for e in effects) and \
                    any(conceals(col, (DARK_TABLE_PROXY,)) is not False for col in (runs or [st.font])):
                for (r1, c1, r2, c2), style, tname in self.table_boxes:
                    if r1 <= cell.row <= r2 and c1 <= cell.col <= c2 and _DARK_TABLE_RX.match(style):
                        # table-style fills are not resolved: assumed visible (Patrick 2026-10-05: every attempt
                        # graded)
                        self.note_default("table_style_visible", f"{location(cell.sheet, cell.ref)}: dark, directly "
                                                                 f"formatted text in table '{tname}' styled {style!r}")
                        return
            return
        if "contrast" in desc:
            for (r1, c1, r2, c2), style, tname in self.table_boxes:
                if r1 <= cell.row <= r2 and c1 <= cell.col <= c2:
                    # table-style colours are not resolved: assumed visible (Patrick 2026-10-05: every attempt graded)
                    self.note_default("table_style_visible", f"{location(cell.sheet, cell.ref)} looks concealed "
                                                             f"({desc}) but lies in table '{tname}' styled {style!r}")
                    return
        self.hits[desc].append((cell.row, cell.col))
        self.examples.setdefault(desc, f"{location(cell.sheet, cell.ref)} = {value!r}"[:120])

    def second_pass_end(self, head, tail):
        self._commit(head.name, self.hits)
        self._commit_undecided()
        self.hits = defaultdict(list)
        self.pending.pop(head.name, None)

    # -------------------------------------------------------------- verdict
    def finish(self) -> dict:
        if self.und_total and not self.n_concealed:
            # nothing is certainly concealed, so the undecided cells decide the verdict: no guessing
            more = f" (and {self.und_total - 1} more undecided cell(s))" if self.und_total > 1 else ""
            raise GradingError(f"{self.key}: {self.und_examples[0]}{more}")
        stats = {"concealed_cells": self.n_concealed,
                 "undecided_cells": self.und_total,
                 "undecided_examples": [reason[:200] for reason in self.und_examples],
                 "hue_exempt": HUE_EXEMPT,
                 "cells_judged": self.n_judged,
                 "formula_values_read": self.n_values_read,
                 "cf_second_pass_sheets": self.cf_second_pass,
                 "cf_cells_evaluated": self.n_cf_cells,
                 "conceal_lc": CONCEAL_LC,
                 "zero_blank_counts": ZERO_BLANK_COUNTS,
                 "show_zeros_off_sheets": self.show_zeros_off,
                 "zeros_hidden_by_show_zeros_off": self.n_zero_hidden_by_sheet,
                 "databar_iconset_value_hidden": self.hidden_value_bars[:10],
                 # Patrick 2026-10-05: conditional formats never stop a grading - every assumption is recorded
                 "cf_assumptions": {
                     "policy": "where a conditional format leaves a cell's verdict open (a rule that cannot be "
                               "evaluated, a CF font over rich-text runs, a CF format or colour Excel's reading of "
                               "which is not verified) the text is assumed visible (Patrick 2026-10-05)",
                     "cells": self.cf_assumed,
                     "examples": self.cf_assumed_examples,
                     "unevaluable_rules": len(self.cf_unevaluable),
                     "unevaluable_rule_examples": list(self.cf_unevaluable.values())[:MAX_ASSUMPTION_EXAMPLES]},
                 "cf_builtin_visible": {
                     "policy": "built-in conditional formats (colour scales, data bars, icon sets, Excel's preset "
                               "highlight styles) never count as hiding text (Patrick 2026-10-05)",
                     "cells": self.cf_builtin_visible,
                     "examples": self.cf_builtin_examples},
                 "examples": dict(list(self.examples.items())[:8])}
        return self.verdict("No populated cell is concealed by its formatting.",
                            "{n} block(s) of populated cells concealed by formatting "
                            f"({self.n_concealed} cell(s)).", stats)
