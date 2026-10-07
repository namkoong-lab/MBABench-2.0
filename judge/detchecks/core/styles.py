"""styles.xml + theme reader and Excel colour resolution (stdlib only).

Colours are resolved to ARGB strings 'FFRRGGBB' (alpha always FF: Excel ignores the
stored alpha byte for cell fonts and fills; openpyxl writes 00RRGGBB for opaque colours).

Resolution rules (ECMA-376 18.8.19 CT_Color, verified against Excel):
  rgb       -> that colour (alpha ignored); 6 or 8 hex digits
  indexed   -> the workbook's own <indexedColors> palette if present, else the legacy
               64-colour default palette (note indexed 13 = FFFF00, bright yellow);
               64 = system foreground (window text: black), 65 = system background
               (window: white)
  theme     -> theme part clrScheme slot, cell index order 0=lt1, 1=dk1, 2=lt2, 3=dk2,
               4..9=accent1..6, 10=hlink, 11=folHlink (same as judge/utils/theme_palette.py;
               so theme 0 is WHITE and theme 1 BLACK in the default Office theme)
  auto      -> font/border: black (window text); fill: white (window background)
  tint      -> applied to the base colour with Excel's own integer HLS arithmetic
               (HLSMAX = 240, see excel_tint()); reproduces Excel's rendered colours
               exactly on all 46 Office-theme swatches tested, e.g. theme 1 + tint
               0.49998 -> 808080 (theme_palette.py's float version gives 7F7F7F).
Anything else is UNRESOLVABLE and resolves to an explicit UnknownColour (never None, never black
or 'no fill'): an rgb that is not 6 or 8 hex digits ('FFGGFF00', ''), a theme index outside the 12
slots, an indexed value outside the palette other than 64 / 65 (e.g. 81, the system tooltip-text
colour Excel writes for comment fonts), an invalid entry of a custom <indexedColors> palette, a tint
that is not a number in [-1, 1].  Building the style table never raises for one; a check that meets
one decides whether its verdict depends on the colour and raises GradingError only then (finding
47-R2-08, 2026-10-04; docs/reader.md section 5).
"""
from __future__ import annotations

import math
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Optional

from ..errors import GradingError

# ECMA-376 default indexed palette (0..63), RRGGBB
DEFAULT_INDEXED = (
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
)
# Office 2013+ default theme, cell theme-index order (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink)
DEFAULT_THEME = ("FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
                 "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72")
THEME_SLOTS = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6", "hlink", "folHlink")

# Built-in number formats (ECMA-376 18.8.30, en-US).  Locale-dependent ids (5-8, 27-36,
# 50-81) are not listed; detchecks.core.numfmt (if present) is the authority for rendering.
BUILTIN_NUM_FMTS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}

# share of the foreground colour a hatch pattern paints (approximation for effective colour)
PATTERN_FG_SHARE = {
    "solid": 1.0, "darkGray": 0.75, "mediumGray": 0.5, "lightGray": 0.25, "gray125": 0.125,
    "gray0625": 0.0625, "darkHorizontal": 0.5, "darkVertical": 0.5, "darkDown": 0.5,
    "darkUp": 0.5, "darkGrid": 0.75, "darkTrellis": 0.75, "lightHorizontal": 0.25,
    "lightVertical": 0.25, "lightDown": 0.25, "lightUp": 0.25, "lightGrid": 0.4,
    "lightTrellis": 0.4,
}

_HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")
_HEX8 = re.compile(r"^[0-9A-Fa-f]{8}$")
BLACK = "FF000000"
WHITE = "FFFFFFFF"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _b(el, name: str, default: bool = False) -> bool:
    v = el.get(name)
    if v is None:
        return default
    return v.strip().lower() in ("1", "true", "on")


def _bool_el(el) -> bool:
    """<b/>, <b val="1"/> -> True; <b val="0"/> -> False."""
    v = el.get("val")
    return v is None or v.strip().lower() in ("1", "true", "on")


def _int(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _float(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


# ----------------------------------------------------------------------------- tint (Excel HLS)
_HLSMAX = 240


def _rgb_to_hls(r: int, g: int, b: int) -> tuple[int, int, int]:
    mx, mn = max(r, g, b), min(r, g, b)
    lum = ((mx + mn) * _HLSMAX + 255) // 510
    if mx == mn:
        return _HLSMAX * 2 // 3, lum, 0
    d = mx - mn
    if lum <= _HLSMAX // 2:
        sat = (d * _HLSMAX + (mx + mn) // 2) // (mx + mn)
    else:
        sat = (d * _HLSMAX + (510 - mx - mn) // 2) // (510 - mx - mn)
    rd = ((mx - r) * (_HLSMAX // 6) + d // 2) // d
    gd = ((mx - g) * (_HLSMAX // 6) + d // 2) // d
    bd = ((mx - b) * (_HLSMAX // 6) + d // 2) // d
    if r == mx:
        hue = bd - gd
    elif g == mx:
        hue = _HLSMAX // 3 + rd - bd
    else:
        hue = 2 * _HLSMAX // 3 + gd - rd
    if hue < 0:
        hue += _HLSMAX
    if hue > _HLSMAX:
        hue -= _HLSMAX
    return hue, lum, sat


def _hue_to_rgb(n1: int, n2: int, hue: int) -> int:
    if hue < 0:
        hue += _HLSMAX
    if hue > _HLSMAX:
        hue -= _HLSMAX
    if hue < _HLSMAX // 6:
        return n1 + ((n2 - n1) * hue + _HLSMAX // 12) // (_HLSMAX // 6)
    if hue < _HLSMAX // 2:
        return n2
    if hue < _HLSMAX * 2 // 3:
        return n1 + ((n2 - n1) * (_HLSMAX * 2 // 3 - hue) + _HLSMAX // 12) // (_HLSMAX // 6)
    return n1


def _hls_to_rgb(hue: int, lum: int, sat: int) -> tuple[int, int, int]:
    if sat == 0:
        v = (lum * 255 + _HLSMAX // 2) // _HLSMAX
        return v, v, v
    if lum <= _HLSMAX // 2:
        m2 = (lum * (_HLSMAX + sat) + _HLSMAX // 2) // _HLSMAX
    else:
        m2 = lum + sat - (lum * sat + _HLSMAX // 2) // _HLSMAX
    m1 = 2 * lum - m2
    conv = lambda h: max(0, min(255, (_hue_to_rgb(m1, m2, h) * 255 + _HLSMAX // 2) // _HLSMAX))
    return conv(hue + _HLSMAX // 3), conv(hue), conv(hue - _HLSMAX // 3)


def excel_tint(hex6: str, tint: float) -> str:
    """Apply an OOXML tint the way Excel renders it.  hex6 'RRGGBB' -> 'RRGGBB'.
    Integer HLS on a 0..240 scale; darken L' = floor(L*(1+t)); lighten
    L' = floor(L*(1-t) + 240 - floor(240*(1-t)))."""
    if not tint:
        return hex6.upper()
    t = max(-1.0, min(1.0, float(tint)))
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    hue, lum, sat = _rgb_to_hls(r, g, b)
    if t < 0:
        lum2 = math.floor(lum * (1.0 + t))
    else:
        lum2 = math.floor(lum * (1.0 - t) + _HLSMAX - math.floor(_HLSMAX * (1.0 - t)))
    lum2 = max(0, min(_HLSMAX, int(lum2)))
    return "%02X%02X%02X" % _hls_to_rgb(hue, lum2, sat)


# ----------------------------------------------------------------------------- dataclasses
@dataclass(frozen=True)
class Color:
    """A colour reference exactly as stored (one of rgb / theme / indexed / auto) + tint."""
    rgb: Optional[str] = None       # as written (6 or 8 hex digits when valid), upper-case
    theme: Optional[int] = None
    indexed: Optional[int] = None
    auto: bool = False
    tint: float = 0.0
    bad: Optional[str] = None       # an attribute as stored that is not a valid value ('theme="x"', 'tint="2"')

    @staticmethod
    def from_el(el) -> Optional["Color"]:
        if el is None:
            return None
        a = el.attrib
        bad = None
        tint = 0.0
        if a.get("tint") is not None:
            t = _float(a.get("tint"))
            if t is None or not math.isfinite(t) or not -1.0 <= t <= 1.0:
                bad = f'tint="{a.get("tint")}"'
            else:
                tint = t
        if a.get("rgb") is not None:
            return Color(rgb=a["rgb"].strip().upper(), tint=tint, bad=bad)
        if a.get("theme") is not None:
            t = _int(a["theme"])
            return Color(theme=t if t is not None else -1, tint=tint,
                         bad=bad if t is not None else f'theme="{a["theme"]}"')
        if a.get("indexed") is not None:
            i = _int(a["indexed"])
            return Color(indexed=i if i is not None else -1, tint=tint,
                         bad=bad if i is not None else f'indexed="{a["indexed"]}"')
        if (a.get("auto") or "").lower() in ("1", "true"):
            return Color(auto=True)
        return None

    def describe(self) -> str:
        t = f"{self.tint:+.3f}" if self.tint else ""
        b = f" ({self.bad})" if self.bad else ""
        if self.rgb is not None:
            return f"rgb:{self.rgb}{t}{b}"
        if self.theme is not None:
            return f"theme:{self.theme}{t}{b}"
        if self.indexed is not None:
            return f"indexed:{self.indexed}{t}{b}"
        return "auto" if self.auto else "none"


@dataclass(frozen=True)
class UnknownColour:
    """What Styles.resolve returns for a colour reference whose rendering is not known (see the module
    doc): never a colour string, so it cannot be mistaken for black, white or 'no fill' - code that
    slices it like 'FFRRGGBB' fails loudly.  ref = the reference as stored (Color.describe()), why =
    why it cannot be resolved.  Hashable; equal references compare equal (Excel renders one stored
    value one way, whatever that is)."""
    ref: str
    why: str

    def __str__(self) -> str:
        return f"unresolvable colour {self.ref} ({self.why})"


def is_unknown(colour) -> bool:
    """True for an UnknownColour (an unresolvable colour reference)."""
    return isinstance(colour, UnknownColour)


@dataclass(frozen=True)
class Font:
    name: Optional[str] = None
    sz: Optional[float] = None
    b: bool = False
    i: bool = False
    u: Optional[str] = None        # None | single | double | singleAccounting | doubleAccounting
    strike: bool = False
    color: Optional[Color] = None  # None = no <color> element (= automatic)
    vert_align: Optional[str] = None
    family: Optional[int] = None
    scheme: Optional[str] = None
    # for dxf fonts: which properties were explicitly given (a dxf font is a partial override)
    specified: frozenset = frozenset()


@dataclass(frozen=True)
class Fill:
    kind: str = "none"             # none | pattern | gradient
    pattern: Optional[str] = None  # patternType ('solid', 'gray125', ...); None for none/gradient
    fg: Optional[Color] = None
    bg: Optional[Color] = None
    gradient_type: Optional[str] = None
    stops: tuple = ()              # ((position, Color), ...)


@dataclass(frozen=True)
class BorderEdge:
    style: Optional[str] = None
    color: Optional[Color] = None


@dataclass(frozen=True)
class Border:
    left: BorderEdge = BorderEdge()
    right: BorderEdge = BorderEdge()
    top: BorderEdge = BorderEdge()
    bottom: BorderEdge = BorderEdge()
    diagonal: BorderEdge = BorderEdge()


@dataclass(frozen=True)
class Alignment:
    horizontal: Optional[str] = None   # general/left/center/right/fill/justify/centerContinuous/distributed
    vertical: Optional[str] = None
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int = 0
    text_rotation: int = 0


@dataclass(frozen=True)
class Xf:
    num_fmt_id: int = 0
    font_id: int = 0
    fill_id: int = 0
    border_id: int = 0
    xf_id: Optional[int] = None        # parent cellStyleXfs index (cellXfs only)
    alignment: Alignment = Alignment()
    locked: bool = True
    hidden: bool = False               # protection hidden (formula hidden when sheet protected)
    quote_prefix: bool = False
    apply: dict = field(default_factory=dict, hash=False, compare=False)  # applyFont etc. as stored


@dataclass(frozen=True)
class Dxf:
    font: Optional[Font] = None
    fill: Optional[Fill] = None
    num_fmt_id: Optional[int] = None
    num_fmt_code: Optional[str] = None
    border: Optional[Border] = None


@dataclass(frozen=True)
class CellStyle:
    name: str
    xf_id: int
    builtin_id: Optional[int]


@dataclass(frozen=True)
class FillPaint:
    """What a fill paints.  effective = the single colour a viewer perceives:
    solid -> fg; other patterns -> fg blended over bg by pattern density; gradient ->
    average of the stops.  None means 'no fill' (the cell shows the sheet background, white);
    an UnknownColour means the paint depends on a colour that cannot be resolved."""
    kind: str
    pattern: Optional[str]
    fg: Optional[str]
    bg: Optional[str]
    effective: Optional[str]
    stops: tuple = ()


# ----------------------------------------------------------------------------- parsing helpers
def parse_font(el, partial: bool = False) -> Font:
    kw = {}
    spec = set()
    for c in el:
        ln = _local(c.tag)
        if ln == "name":
            kw["name"] = c.get("val")
        elif ln == "sz":
            kw["sz"] = _float(c.get("val"))
        elif ln == "b":
            kw["b"] = _bool_el(c)
        elif ln == "i":
            kw["i"] = _bool_el(c)
        elif ln == "strike":
            kw["strike"] = _bool_el(c)
        elif ln == "u":
            v = c.get("val") or "single"
            kw["u"] = None if v == "none" else v
        elif ln == "color":
            kw["color"] = Color.from_el(c)
        elif ln == "vertAlign":
            kw["vert_align"] = c.get("val")
        elif ln == "family":
            kw["family"] = _int(c.get("val"))
        elif ln == "scheme":
            kw["scheme"] = c.get("val")
        else:
            continue
        spec.add(ln)
    return Font(specified=frozenset(spec), **kw)


def parse_fill(el, dxf: bool = False) -> Fill:
    for c in el:
        ln = _local(c.tag)
        if ln == "patternFill":
            pt = c.get("patternType")
            if pt is None:
                # a cell fill without patternType paints nothing; in a dxf it means solid
                pt = "solid" if dxf else "none"
            fg = bg = None
            for g in c:
                gl = _local(g.tag)
                if gl == "fgColor":
                    fg = Color.from_el(g)
                elif gl == "bgColor":
                    bg = Color.from_el(g)
            if pt == "none":
                return Fill("none", None, fg, bg)
            return Fill("pattern", pt, fg, bg)
        if ln == "gradientFill":
            stops = []
            for s in c:
                if _local(s.tag) == "stop":
                    col = None
                    for g in s:
                        if _local(g.tag) == "color":
                            col = Color.from_el(g)
                    stops.append((_float(s.get("position"), 0.0), col))
            return Fill("gradient", None, None, None, c.get("type") or "linear", tuple(stops))
    return Fill("none")


def parse_border(el) -> Border:
    kw = {}
    for c in el:
        ln = _local(c.tag)
        if ln in ("left", "right", "top", "bottom", "diagonal", "start", "end"):
            col = None
            for g in c:
                if _local(g.tag) == "color":
                    col = Color.from_el(g)
            key = {"start": "left", "end": "right"}.get(ln, ln)
            kw[key] = BorderEdge(c.get("style"), col)
    return Border(**kw)


def parse_xf(el, cell_xf: bool) -> Xf:
    al = Alignment()
    locked, hidden = True, False
    for c in el:
        ln = _local(c.tag)
        if ln == "alignment":
            al = Alignment(c.get("horizontal"), c.get("vertical"), _b(c, "wrapText"), _b(c, "shrinkToFit"),
                           _int(c.get("indent"), 0) or 0, _int(c.get("textRotation"), 0) or 0)
        elif ln == "protection":
            locked = _b(c, "locked", True)
            hidden = _b(c, "hidden", False)
    apply = {k: v for k, v in el.attrib.items() if k.startswith("apply")}
    return Xf(_int(el.get("numFmtId"), 0) or 0, _int(el.get("fontId"), 0) or 0, _int(el.get("fillId"), 0) or 0,
              _int(el.get("borderId"), 0) or 0, _int(el.get("xfId")) if cell_xf else None, al, locked, hidden,
              _b(el, "quotePrefix"), apply)


def parse_theme(xml_bytes: Optional[bytes]) -> tuple:
    """12 RRGGBB colours in cell theme-index order; slot-by-slot fallback to the Office default."""
    if not xml_bytes:
        return DEFAULT_THEME
    slots: dict[str, str] = {}
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return DEFAULT_THEME
    for scheme in root.iter():
        if _local(scheme.tag) != "clrScheme":
            continue
        for slot in scheme:
            name = _local(slot.tag)
            for c in slot:
                cl = _local(c.tag)
                val = c.get("val") if cl == "srgbClr" else (c.get("lastClr") or c.get("val")) if cl == "sysClr" else None
                if val and _HEX6.match(val):
                    slots[name] = val.upper()
                    break
        break
    return tuple(slots.get(s, DEFAULT_THEME[i]) for i, s in enumerate(THEME_SLOTS))


def _blend(fg: str, bg: str, w: float) -> str:
    """ARGB blend, w = share of fg."""
    out = []
    for i in (2, 4, 6):
        f, b = int(fg[i:i + 2], 16), int(bg[i:i + 2], 16)
        out.append(int(round(f * w + b * (1 - w))))
    return "FF" + "".join(f"{v:02X}" for v in out)


# ----------------------------------------------------------------------------- Styles
class Styles:
    """Parsed styles.xml (+ theme colours).  All lists are indexed exactly as in the file;
    out-of-range ids resolve to the defaults (index 0 / Font()) the way Excel tolerates them."""

    def __init__(self):
        self.num_fmts: dict[int, str] = {}          # custom formats from <numFmts>
        self.fonts: list[Font] = []
        self.fills: list[Fill] = []
        self.borders: list[Border] = []
        self.cell_xfs: list[Xf] = []
        self.cell_style_xfs: list[Xf] = []
        self.cell_styles: list[CellStyle] = []
        self.dxfs: list[Dxf] = []
        self.indexed_palette: tuple = DEFAULT_INDEXED   # RRGGBB, custom <indexedColors> applied
        self.custom_indexed: bool = False
        self.unknown_indexed: dict = {}                 # palette index -> invalid <rgbColor rgb> as stored
        #                                                 (its palette slot holds a placeholder; resolve() and
        #                                                 indexed_colour() return an UnknownColour for it)
        self.theme_colors: tuple = DEFAULT_THEME        # RRGGBB in cell theme-index order
        self.theme_present: bool = False
        self._fill_paint_cache: dict[int, FillPaint] = {}
        self._font_argb_cache: dict[int, str] = {}

    @classmethod
    def parse(cls, styles_xml: Optional[bytes], theme_xml: Optional[bytes] = None, where: str = "") -> "Styles":
        st = cls()
        st.theme_colors = parse_theme(theme_xml)
        st.theme_present = bool(theme_xml)
        if styles_xml:
            try:
                root = ET.fromstring(styles_xml)
            except ET.ParseError as e:
                raise GradingError(f"XML parse error in {where}:styles.xml: {e}") from e
            for sect in root:
                ln = _local(sect.tag)
                if ln == "numFmts":
                    for e in sect:
                        i = _int(e.get("numFmtId"))
                        if i is not None:
                            st.num_fmts[i] = e.get("formatCode") or ""
                elif ln == "fonts":
                    st.fonts = [parse_font(f) for f in sect if _local(f.tag) == "font"]
                elif ln == "fills":
                    st.fills = [parse_fill(f) for f in sect if _local(f.tag) == "fill"]
                elif ln == "borders":
                    st.borders = [parse_border(b) for b in sect if _local(b.tag) == "border"]
                elif ln == "cellXfs":
                    st.cell_xfs = [parse_xf(x, True) for x in sect if _local(x.tag) == "xf"]
                elif ln == "cellStyleXfs":
                    st.cell_style_xfs = [parse_xf(x, False) for x in sect if _local(x.tag) == "xf"]
                elif ln == "cellStyles":
                    for c in sect:
                        if _local(c.tag) == "cellStyle":
                            st.cell_styles.append(CellStyle(c.get("name") or "", _int(c.get("xfId"), 0) or 0,
                                                            _int(c.get("builtinId"))))
                elif ln == "dxfs":
                    for d in sect:
                        if _local(d.tag) != "dxf":
                            continue
                        font = fill = border = None
                        nf_id = nf_code = None
                        for c in d:
                            cl = _local(c.tag)
                            if cl == "font":
                                font = parse_font(c, partial=True)
                            elif cl == "fill":
                                fill = parse_fill(c, dxf=True)
                            elif cl == "numFmt":
                                nf_id, nf_code = _int(c.get("numFmtId")), c.get("formatCode")
                            elif cl == "border":
                                border = parse_border(c)
                        st.dxfs.append(Dxf(font, fill, nf_id, nf_code, border))
                elif ln == "colors":
                    for c in sect:
                        if _local(c.tag) == "indexedColors":
                            vals = []
                            for rc in c:
                                v = (rc.get("rgb") or "").strip()
                                if _HEX8.match(v) or _HEX6.match(v):
                                    vals.append(v[-6:].upper())
                                else:
                                    st.unknown_indexed[len(vals)] = v      # not a colour: unknown, not black
                                    vals.append("000000")                  # placeholder (never read as a colour)
                            if vals:
                                st.indexed_palette = tuple(vals) + DEFAULT_INDEXED[len(vals):]
                                st.custom_indexed = True
        if not st.fonts:
            st.fonts = [Font(name="Calibri", sz=11.0)]
        if not st.fills:
            st.fills = [Fill("none"), Fill("pattern", "gray125")]
        if not st.cell_xfs:
            st.cell_xfs = [Xf()]
        return st

    # ------------------------------------------------------------------ lookups (never raise)
    def xf(self, s: Optional[int]) -> Xf:
        if s is not None and 0 <= s < len(self.cell_xfs):
            return self.cell_xfs[s]
        return self.cell_xfs[0]

    def font(self, s: Optional[int]) -> Font:
        fid = self.xf(s).font_id
        return self.fonts[fid] if 0 <= fid < len(self.fonts) else self.fonts[0]

    def fill(self, s: Optional[int]) -> Fill:
        fid = self.xf(s).fill_id
        return self.fills[fid] if 0 <= fid < len(self.fills) else Fill("none")

    def border(self, s: Optional[int]) -> Border:
        bid = self.xf(s).border_id
        return self.borders[bid] if 0 <= bid < len(self.borders) else Border()

    def num_fmt_id(self, s: Optional[int]) -> int:
        return self.xf(s).num_fmt_id

    def num_fmt_code(self, s: Optional[int]) -> str:
        """Format code of a cell style: custom <numFmt> first, then the built-in table;
        unknown built-in ids -> 'General'."""
        i = self.xf(s).num_fmt_id
        return self.num_fmts.get(i) or BUILTIN_NUM_FMTS.get(i, "General")

    def style_name(self, s: Optional[int]) -> Optional[str]:
        """Name of the cell style (cellStyles entry) the xf derives from ('Normal', 'Note', ...)."""
        xf_id = self.xf(s).xf_id
        if xf_id is None:
            return None
        for cs in self.cell_styles:
            if cs.xf_id == xf_id:
                return cs.name
        return None

    def dxf(self, dxf_id: Optional[int]) -> Optional[Dxf]:
        if dxf_id is not None and 0 <= dxf_id < len(self.dxfs):
            return self.dxfs[dxf_id]
        return None

    # ------------------------------------------------------------------ colours
    def resolve(self, color: Optional[Color], role: str = "font"):
        """Color -> 'FFRRGGBB', or an UnknownColour when the reference cannot be resolved (never
        raises; never guessed as black / white / no fill).  role 'font'|'border' (auto/None ->
        black) or 'fill' (auto -> white; None -> None = no colour)."""
        if color is None:
            return BLACK if role != "fill" else None
        if color.auto:
            return BLACK if role != "fill" else WHITE
        if color.bad is not None:
            return UnknownColour(color.describe(), f"{color.bad} is not a valid value")
        base = None
        if color.rgb is not None:
            h = color.rgb
            if not (_HEX8.match(h) or _HEX6.match(h)):
                return UnknownColour(color.describe(), "rgb is not 6 or 8 hex digits")
            base = h[-6:].upper()
        elif color.theme is not None:
            if 0 <= color.theme < len(self.theme_colors):
                base = self.theme_colors[color.theme]
            else:
                return UnknownColour(color.describe(), "theme index outside the theme's 12 colour slots")
        elif color.indexed is not None:
            i = color.indexed
            if i == 64:
                base = "000000"          # system foreground (window text)
            elif i == 65:
                base = "FFFFFF"
            elif i in self.unknown_indexed:
                return UnknownColour(color.describe(), f"custom <indexedColors> entry {i} "
                                                       f"(rgb={self.unknown_indexed[i]!r}) is not a colour")
            elif 0 <= i < len(self.indexed_palette):
                base = self.indexed_palette[i]
            else:
                return UnknownColour(color.describe(), "indexed value outside the palette (only 64 and 65, the "
                                                       "window text / background, are known system colours)")
        else:
            return BLACK if role != "fill" else None
        if color.tint:
            base = excel_tint(base, color.tint)
        return "FF" + base

    def indexed_colour(self, i: int):
        """'FFRRGGBB' of palette entry i as a number-format [ColorN] tag uses it (N + 7), or an
        UnknownColour for an invalid custom <indexedColors> entry; None outside the palette."""
        if i in self.unknown_indexed:
            return UnknownColour(f"indexed:{i}", f"custom <indexedColors> entry {i} "
                                                 f"(rgb={self.unknown_indexed[i]!r}) is not a colour")
        if 0 <= i < len(self.indexed_palette):
            return "FF" + self.indexed_palette[i]
        return None

    def font_color(self, s: Optional[int]):
        """Rendered font colour of a cell style (no conditional formatting / number-format colour):
        'FFRRGGBB', or an UnknownColour when the font's colour cannot be resolved."""
        v = self._font_argb_cache.get(s)
        if v is None:
            v = self.resolve(self.font(s).color, "font")
            self._font_argb_cache[s] = v
        return v

    def fill_paint(self, fill: Fill, dxf: bool = False) -> FillPaint:
        """Resolve a Fill.  For a solid cell fill the paint is fgColor.  For a dxf
        (conditional-format) solid fill pass dxf=True: Excel paints bgColor ONLY (measured by
        Patrick in Excel, 2026-10-03): fgColor is ignored; bgColor indexed 64 and rgb 00000000
        paint BLACK (alpha ignored); no bgColor -> no fill (FillPaint kind 'none').
        An unresolvable colour the paint uses makes `effective` that UnknownColour (a gradient stop:
        the stop is the UnknownColour and `effective` the first unknown stop); one the paint does not
        use (the bgColor of a solid cell fill, the fgColor of a dxf solid fill) changes nothing."""
        if fill.kind == "none":
            return FillPaint("none", None, None, None, None)
        if fill.kind == "gradient":
            cols = [self.resolve(c, "fill") for _p, c in fill.stops]
            cols = [c for c in cols if c]                   # a stop without a colour paints nothing
            unknown = next((c for c in cols if is_unknown(c)), None)
            eff = unknown
            if cols and unknown is None:
                n = len(cols)
                eff = "FF" + "".join(f"{int(round(sum(int(c[i:i + 2], 16) for c in cols) / n)):02X}" for i in (2, 4, 6))
            return FillPaint("gradient", None, None, None, eff, tuple(cols))
        # an 'auto' fgColor is Excel's "Automatic" pattern colour (black); an auto bgColor is the
        # window background (white)
        fg = (BLACK if fill.fg.auto else self.resolve(fill.fg, "fill")) if fill.fg is not None else None
        bg = self.resolve(fill.bg, "fill") if fill.bg is not None else None
        pt = fill.pattern or "solid"
        if pt == "solid":
            if dxf:
                if fill.bg is None:
                    return FillPaint("none", None, fg, None, None)
                # an auto bgColor in a dxf: Excel's "automatic" colour, as indexed 64 (black); not
                # measured on its own (indexed 64 was), see docs/reader.md
                eff = BLACK if fill.bg.auto else bg
            else:
                eff = fg if fill.fg is not None else BLACK   # solid without fgColor paints black
        else:
            fgc = fg or BLACK
            bgc = bg or WHITE
            if is_unknown(fgc) or is_unknown(bgc):
                eff = fgc if is_unknown(fgc) else bgc        # the blend depends on the unknown colour
            else:
                eff = _blend(fgc, bgc, PATTERN_FG_SHARE.get(pt, 0.5))
        return FillPaint("pattern", pt, fg, bg, eff)

    def cell_fill(self, s: Optional[int]) -> FillPaint:
        """Fill paint of a cell style (cached per style index; see fill_paint for unknown colours)."""
        p = self._fill_paint_cache.get(s)
        if p is None:
            p = self.fill_paint(self.fill(s))
            self._fill_paint_cache[s] = p
        return p

    def fill_color(self, s: Optional[int]):
        """Effective fill colour 'FFRRGGBB' of a cell style, None when there is no fill, an
        UnknownColour when the paint depends on an unresolvable colour."""
        return self.cell_fill(s).effective

    def dxf_fill_color(self, dxf_id: Optional[int]):
        """Colour a conditional-format dxf paints (bgColor only), None for no fill, an UnknownColour
        when that colour cannot be resolved."""
        d = self.dxf(dxf_id)
        if d is None or d.fill is None:
            return None
        return self.fill_paint(d.fill, dxf=True).effective

    def dxf_font_color(self, dxf_id: Optional[int]):
        """Font colour a dxf sets, None when the dxf does not set one, an UnknownColour when it
        cannot be resolved."""
        d = self.dxf(dxf_id)
        if d is None or d.font is None or d.font.color is None:
            return None
        return self.resolve(d.font.color, "font")

    @staticmethod
    def effective_style(cell_s: Optional[int], row_style: Optional[int], col_style: Optional[int]) -> int:
        """Style index that applies at a position: the cell's own s if a <c> exists; else the
        row style (row s with customFormat=1); else the column style (<col style>); else 0."""
        if cell_s is not None:
            return cell_s
        if row_style is not None:
            return row_style
        if col_style is not None:
            return col_style
        return 0


def rgb6(argb):
    """'FFRRGGBB' -> 'RRGGBB' (None and an UnknownColour pass through)."""
    if argb is None or is_unknown(argb):
        return argb
    return argb[-6:] if argb else None
