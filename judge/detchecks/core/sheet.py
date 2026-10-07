"""Streaming worksheet reader (stdlib xml.etree; never openpyxl.load_workbook).

A sheet part is read in three phases (SheetStream):
  1. head  - everything before <sheetData>: sheetPr, dimension, sheetViews,
             sheetFormatPr, cols.  Parsed from the raw bytes up to <sheetData.
  2. body  - <sheetData>: rows and cells, streamed with XMLPullParser and cleared as
             they go (memory stays flat on 140 MB sheet parts).  Or SKIPPED at byte level
             (no XML parsing at all) when nobody needs cells.
  3. tail  - everything after </sheetData>: mergeCells, conditionalFormatting,
             dataValidations, hyperlinks, ignoredErrors, sheetProtection, drawing /
             legacyDrawing, tableParts, extLst (x14 CF/DV, sparklines) ...

All elements are matched by LOCAL name, so prefixed (<x:c>) and Strict OOXML parts read
the same way.  Rows/cells without an r attribute get their positions inferred.
"""
from __future__ import annotations

import codecs
import io
import re
import xml.etree.ElementTree as ET
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from typing import Callable, Optional

from ..errors import GradingError
from . import formula as _fm
from .package import Package, SheetInfo, local as _plocal, rel_id, truthy
from .refs import MAX_COL, MAX_ROW, col_to_index, index_to_col, parse_range, parse_sqref, split_ref
from .styles import Color, Dxf, parse_border, parse_fill, parse_font

CHUNK = 1 << 20

_SD_START = re.compile(rb"<(?:[A-Za-z_][\w.\-]*:)?sheetData(?=[\s/>])")
_SD_END = re.compile(rb"</(?:[A-Za-z_][\w.\-]*:)?sheetData\s*>")
# start tag of an element: qualified name, attributes (quoted values may contain '>'), end
_START_TAG = re.compile(rb"<([A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?)"
                        rb"(?:\s+[^\s=/>]+\s*=\s*(?:\"[^\"]*\"|'[^']*'))*\s*/?>")
_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


def _find_root(buf) -> Optional[re.Match]:
    """Match of the ROOT element's start tag in a document prefix: skips the BOM, the XML
    declaration, processing instructions, comments and a DOCTYPE (a '<tool>' inside a
    comment is not the root).  None when the prefix holds no complete root start tag."""
    n = len(buf)
    i = 3 if buf[:3] == b"\xef\xbb\xbf" else 0
    while i < n:
        while i < n and buf[i:i + 1] in (b" ", b"\t", b"\r", b"\n"):
            i += 1
        if buf.startswith(b"<?", i):
            j = buf.find(b"?>", i + 2)
            if j < 0:
                return None
            i = j + 2
        elif buf.startswith(b"<!--", i):
            j = buf.find(b"-->", i + 4)
            if j < 0:
                return None
            i = j + 3
        elif buf.startswith(b"<!", i):            # <!DOCTYPE ...> (with an optional [internal subset])
            j = buf.find(b">", i)
            k = buf.find(b"[", i)
            if 0 <= k < j:
                k2 = buf.find(b"]", k)
                j = buf.find(b">", k2) if k2 >= 0 else -1
            if j < 0:
                return None
            i = j + 1
        else:
            return _START_TAG.match(buf, i)
    return None


def _grid_range(ref: str, what: str, where: str) -> tuple:
    """parse_range() that refuses unreadable or off-grid ranges (GradingError)."""
    b = parse_range(ref)
    if b is None or b[0] < 1 or b[1] < 1 or b[2] > MAX_ROW or b[3] > MAX_COL:
        raise GradingError(f"unreadable {what} {ref!r} in {where}")
    return b


_LOCAL: dict[str, str] = {}


def _loc(tag: str) -> str:
    v = _LOCAL.get(tag)
    if v is None:
        v = _LOCAL[tag] = tag.rsplit("}", 1)[-1]
    return v


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _i(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _t(v) -> bool:
    return v is not None and v.strip().lower() in ("1", "true", "on")


# ============================================================================ value types
class ExcelError(str):
    """A cached error value ('#N/A', '#DIV/0!', ...).  isinstance(v, ExcelError) separates it
    from a text constant that merely reads '#N/A'."""
    __slots__ = ()

    def __repr__(self):
        return f"ExcelError({str.__repr__(self)})"


@dataclass(frozen=True)
class Run:
    """One rich-text run.  font is a partial Font (only the properties the run sets;
    see Font.specified) or None when the run has no <rPr> (inherits the cell font)."""
    text: str
    font: object = None


def parse_runs(el) -> tuple[str, Optional[list]]:
    """<si> / <is> element -> (plain text, [Run] or None for plain strings).
    Phonetic runs (<rPh>) are excluded."""
    runs = []
    parts = []
    plain = None
    for c in el:
        ln = _loc(c.tag)
        if ln == "t":
            plain = (plain or "") + (c.text or "")
        elif ln == "r":
            txt = ""
            font = None
            for g in c:
                gl = _loc(g.tag)
                if gl == "t":
                    txt += g.text or ""
                elif gl == "rPr":
                    font = parse_font(g, partial=True)
            runs.append(Run(txt, font))
            parts.append(txt)
    if not runs:
        return (plain or ""), None
    return (plain or "") + "".join(parts), runs


# ============================================================================ head objects
@dataclass
class Pane:
    x_split: float = 0.0
    y_split: float = 0.0
    top_left_cell: Optional[str] = None
    active_pane: str = "topLeft"           # topLeft | topRight | bottomLeft | bottomRight
    state: str = "split"                   # split | frozen | frozenSplit


@dataclass
class Selection:
    pane: str = "topLeft"
    active_cell: Optional[str] = None
    active_cell_id: int = 0
    sqref: Optional[str] = None


@dataclass
class SheetView:
    workbook_view_id: int = 0
    tab_selected: bool = False
    zoom_scale: Optional[int] = None                    # as stored (None = absent = 100)
    zoom_scale_normal: Optional[int] = None
    zoom_scale_page_layout_view: Optional[int] = None
    zoom_scale_sheet_layout_view: Optional[int] = None
    view: str = "normal"                                # normal | pageBreakPreview | pageLayout
    top_left_cell: Optional[str] = None
    show_grid_lines: bool = True
    show_zeros: bool = True
    show_formulas: bool = False
    show_row_col_headers: bool = True
    show_outline_symbols: bool = True
    right_to_left: bool = False
    pane: Optional[Pane] = None
    selections: list = field(default_factory=list)      # [Selection] in file order
    attrs: dict = field(default_factory=dict)           # raw attributes

    @property
    def zoom(self) -> int:
        """Effective zoom of the view the sheet opens in (zoomScale, default 100)."""
        return self.zoom_scale if self.zoom_scale else 100

    @property
    def active_pane(self) -> str:
        return self.pane.active_pane if self.pane is not None else "topLeft"

    @property
    def active_selection(self) -> Optional[Selection]:
        """The selection of the ACTIVE pane (not simply the first <selection>).  Without a
        pane, the topLeft selection (a <selection> without pane= is topLeft)."""
        ap = self.active_pane
        for s in self.selections:
            if s.pane == ap:
                return s
        return None

    @property
    def active_cell(self) -> str:
        """Cursor cell the sheet opens with: the active pane's selection activeCell, or 'A1'
        when that pane has no selection (Excel's default)."""
        s = self.active_selection
        if s is not None and s.active_cell:
            return s.active_cell
        return "A1"


@dataclass
class SheetFormat:
    default_row_height: Optional[float] = None
    custom_height: bool = False
    zero_height: bool = False              # rows are hidden unless a <row> says otherwise
    default_col_width: Optional[float] = None
    base_col_width: Optional[float] = None  # None = absent (Excel default 8)
    outline_level_row: int = 0
    outline_level_col: int = 0
    thick_top: bool = False
    thick_bottom: bool = False


@dataclass
class ColInfo:
    min: int
    max: int
    width: Optional[float] = None
    hidden: bool = False
    custom_width: bool = False
    best_fit: bool = False
    outline_level: int = 0
    collapsed: bool = False
    style: Optional[int] = None            # <col style> (column default style), None if absent


def paint_cols(cols) -> list:
    """Disjoint, sorted (lo, hi, ColInfo) segments of a sheet's <col> entries: each entry, taken in the
    reader's order (by min, file order among equal mins - the order of SheetHead.cols), overrides the
    earlier ones on the columns it covers, as a whole (width, hidden, style ... all come from the winning
    entry; an entry without a width still wins and means the default width).  Columns are clamped to
    1..16384; an entry with min > max covers nothing.  Valid files never overlap; GPT-6 tooling writes
    e.g. F:G 9.0 followed by F:F 44.0, or C:XFD 18.0 followed by D:D 3.0 (187 corpus files).  The rule
    'the later entry wins' is the one Reasonable column widths (70) and Reasonable row heights (73)
    decided; no Excel measurement or toy settles what Excel itself does with overlapping entries."""
    segs: list = []
    ends: list = []
    for ci in cols:
        lo, hi = max(1, ci.min), min(MAX_COL, ci.max)
        if lo > hi:
            continue
        if not segs or lo > ends[-1]:
            segs.append((lo, hi, ci))
            ends.append(hi)
            continue
        i = bisect_left(ends, lo)                 # first segment ending at or after lo
        j = i
        while j < len(segs) and segs[j][0] <= hi:
            j += 1
        repl = []
        if i < j and segs[i][0] < lo:
            repl.append((segs[i][0], lo - 1, segs[i][2]))
        repl.append((lo, hi, ci))
        if i < j and segs[j - 1][1] > hi:
            repl.append((hi + 1, segs[j - 1][1], segs[j - 1][2]))
        segs[i:j] = repl
        ends[i:j] = [s[1] for s in repl]
    return segs


@dataclass
class SheetPr:
    tab_color: Optional[Color] = None
    code_name: Optional[str] = None
    filter_mode: bool = False
    summary_below: bool = True             # outlinePr
    summary_right: bool = True
    apply_styles: bool = False
    show_outline_symbols: bool = True
    fit_to_page: bool = False
    attrs: dict = field(default_factory=dict)


@dataclass
class SheetHead:
    """Everything before <sheetData>.  Valid from sheet_start() on."""
    info: SheetInfo
    sheet_pr: SheetPr = field(default_factory=SheetPr)
    dimension: Optional[str] = None
    views: list = field(default_factory=list)          # [SheetView]
    format: SheetFormat = field(default_factory=SheetFormat)
    cols: list = field(default_factory=list)           # [ColInfo] sorted by min (file order among equal mins)
    has_sheet_data: bool = True
    _col_segs: Optional[list] = field(default=None, repr=False)       # paint_cols(cols), built lazily
    _col_seg_starts: list = field(default_factory=list, repr=False)
    _col_segs_key: tuple = field(default=(), repr=False)              # (id, len) of the cols list it was built from

    @property
    def name(self) -> str:
        return self.info.name

    @property
    def state(self) -> str:
        return self.info.state

    @property
    def kind(self) -> str:
        return self.info.kind

    @property
    def index(self) -> int:
        return self.info.index

    @property
    def view(self) -> Optional[SheetView]:
        """The sheetView of workbook window 0 (else the first one); None if the sheet has none."""
        for v in self.views:
            if v.workbook_view_id == 0:
                return v
        return self.views[0] if self.views else None

    def col_segments(self) -> list:
        """Disjoint, sorted (lo, hi, ColInfo) runs of the columns the <col> entries cover, the later
        entry winning where entries overlap (paint_cols: reader's order = by min, file order among
        equal mins).  Columns outside every run are at the sheet defaults."""
        key = (id(self.cols), len(self.cols))
        if self._col_segs is None or self._col_segs_key != key:
            self.cols.sort(key=lambda c: c.min)
            self._col_segs = paint_cols(self.cols)
            self._col_seg_starts = [s[0] for s in self._col_segs]
            self._col_segs_key = key
        return self._col_segs

    def col_info(self, col: int) -> Optional[ColInfo]:
        """The <col> entry that applies to a column index, or None (column at defaults).  Where
        entries overlap (invalid, but written by GPT-6 tooling) the later entry wins (col_segments)."""
        if not self.cols:
            return None
        segs = self.col_segments()
        i = bisect_right(self._col_seg_starts, col) - 1
        if i >= 0 and col <= segs[i][1]:
            return segs[i][2]
        return None

    def col_run(self, col: int) -> tuple:
        """(lo, hi, ColInfo or None): the maximal run of columns around `col` (1..16384) that share
        the same applying entry (None = sheet defaults), for walking a range run by run."""
        if not self.cols:
            return 1, MAX_COL, None
        segs = self.col_segments()
        starts = self._col_seg_starts
        i = bisect_right(starts, col) - 1
        if i >= 0 and col <= segs[i][1]:
            return segs[i][0], segs[i][1], segs[i][2]
        lo = segs[i][1] + 1 if i >= 0 else 1
        hi = starts[i + 1] - 1 if i + 1 < len(segs) else MAX_COL
        return lo, hi, None

    def col_style(self, col: int) -> Optional[int]:
        ci = self.col_info(col)
        return ci.style if ci is not None else None


# ============================================================================ body objects
class Row:
    """A <row> element (delivered to row() before the row's cells)."""
    __slots__ = ("r", "ht", "custom_height", "hidden", "outline_level", "collapsed", "s",
                 "custom_format", "spans", "thick_top", "thick_bot", "attrs")

    def __init__(self, r, a):
        self.r = r
        self.ht = _f(a.get("ht"))
        self.custom_height = _t(a.get("customHeight"))
        self.hidden = _t(a.get("hidden"))
        self.outline_level = _i(a.get("outlineLevel"), 0) or 0
        self.collapsed = _t(a.get("collapsed"))
        s = a.get("s")
        self.s = _i(s) if s is not None else None
        self.custom_format = _t(a.get("customFormat"))
        self.spans = a.get("spans")
        self.thick_top = _t(a.get("thickTop"))
        self.thick_bot = _t(a.get("thickBot"))
        self.attrs = dict(a)

    @property
    def style(self) -> Optional[int]:
        """Row default style (applies to cells of this row that have no <c>): s when customFormat=1."""
        return self.s if self.custom_format else None

    def __repr__(self):
        return f"Row({self.r}, ht={self.ht}, hidden={self.hidden}, s={self.s})"


class SharedMaster:
    """The master of a shared-formula group (si).  Children are expanded with
    detchecks.core.formula.shift (the same code as formula.translate): it shifts the left
    end of 'B5:INDEX(...)' ranges and turns references pushed off the grid into #REF!, which
    openpyxl's Translator gets wrong.  A master text the formula lexer cannot read raises
    FormulaError (a GradingError)."""
    __slots__ = ("si", "text", "row", "col", "ref")

    def __init__(self, si, text, row, col, ref):
        self.si, self.text, self.row, self.col, self.ref = si, text, row, col, ref

    def translate(self, row: int, col: int) -> str:
        if row == self.row and col == self.col:
            return self.text
        return _fm.shift(self.text, row - self.row, col - self.col)


class Formula:
    """A cell's <f> element.

    text      formula text exactly as stored, without '=' (None for shared-formula children
              and for the empty <f/> markers Excel writes on some spill members)
    kind      'normal' | 'shared' | 'array' | 'dataTable'
    ref       ref attribute (array / dataTable range, or the shared group's range on the master)
    si        shared group index (str) for kind == 'shared'
    expanded  the formula text that applies to THIS cell: the stored text, or for a
              shared child the master's text shifted by the offset (formula.shift).
              Raises GradingError when a shared child has no master.
    """
    __slots__ = ("text", "kind", "ref", "si", "ca", "aca", "attrs", "_master", "_row", "_col", "_exp")

    def __init__(self, text, kind, ref, si, attrs, master, row, col):
        self.text = text
        self.kind = kind
        self.ref = ref
        self.si = si
        self.attrs = attrs
        self.ca = _t(attrs.get("ca"))
        self.aca = _t(attrs.get("aca"))
        self._master = master
        self._row = row
        self._col = col
        self._exp = None

    @property
    def is_shared_master(self) -> bool:
        return self.kind == "shared" and bool(self.text)

    @property
    def is_shared_child(self) -> bool:
        return self.kind == "shared" and not self.text

    @property
    def is_empty_marker(self) -> bool:
        """<f/> or <f ca="1"/> with no text and no shared group: not a formula of its own."""
        return self.kind == "normal" and not self.text

    @property
    def master_ref(self) -> Optional[str]:
        m = self._master
        return f"{index_to_col(m.col)}{m.row}" if m is not None else None

    @property
    def expanded(self) -> Optional[str]:
        if self._exp is not None:
            return self._exp
        if self.kind == "shared" and not self.text:
            if self._master is None:
                raise GradingError(f"shared formula si={self.si} at {index_to_col(self._col)}{self._row} "
                                   f"has no master formula before it")
            self._exp = self._master.translate(self._row, self._col)
        else:
            self._exp = self.text
        return self._exp

    def __repr__(self):
        return f"Formula({self.kind}, {self.text!r}, ref={self.ref}, si={self.si})"


class ArrayRange:
    """An array / dynamic-array (spill) / data-table range anchored at a formula cell."""
    __slots__ = ("r1", "c1", "r2", "c2", "anchor_ref", "formula")

    def __init__(self, b, anchor_ref, formula):
        self.r1, self.c1, self.r2, self.c2 = b
        self.anchor_ref = anchor_ref
        self.formula = formula

    @property
    def ref(self) -> str:
        return self.formula.ref


class Cell:
    """One <c> element.  Attributes:

    ref, row, col   'B7', 7, 2
    t               type attribute as stored ('n' when absent): n s str b e inlineStr d
    s               style index (cellXfs), 0 when absent
    raw             raw cached value text (<v> text; for inlineStr the concatenated <is> text);
                    None when the cell has no value element
    formula         Formula or None
    array           ArrayRange containing this cell when it is a NON-anchor member of an
                    array / spill / data-table range (its value is computed by array.formula)
    cm, vm          cell-metadata / value-metadata indexes (dynamic arrays, rich errors) or None
    value           decoded value (see values.py for the source): float | str | bool |
                    ExcelError | None.  NO FALLBACK: for a formula result (own formula, <f/>
                    marker or array member) it raises GradingError when the value is
                    untrusted, or when the reading check did not declare needs_values = True.
    value_trusted   whether value can be relied on (constants always; formula results per
                    provenance)
    unverified_value  the value WITHOUT those guards (diagnostics / messages only)
    """
    __slots__ = ("ref", "row", "col", "t", "s", "raw", "formula", "array", "cm", "vm", "_runs", "_ctx", "_rv")

    def __init__(self, ref, row, col, t, s, raw, formula, array, cm, vm, runs, ctx, rv):
        self.ref = ref
        self.row = row
        self.col = col
        self.t = t
        self.s = s
        self.raw = raw
        self.formula = formula
        self.array = array
        self.cm = cm
        self.vm = vm
        self._runs = runs
        self._ctx = ctx
        self._rv = rv

    # ---- formula helpers
    @property
    def has_formula(self) -> bool:
        """True when the cell carries a formula of its own (normal, shared master/child,
        array/dataTable anchor).  False for spill/array members and empty <f/> markers."""
        f = self.formula
        return f is not None and not f.is_empty_marker

    @property
    def is_formula_result(self) -> bool:
        """The cell's value comes from a formula: own formula (incl. an empty <f/> marker) or
        membership of an array / spill / data-table range."""
        return self.formula is not None or self.array is not None

    @property
    def formula_text(self) -> Optional[str]:
        """Formula that applies to this cell (expanded for shared children), else None."""
        return self.formula.expanded if self.has_formula else None

    # ---- values
    @property
    def has_value(self) -> bool:
        """A cached/stored value element is present and non-empty (an empty-string formula
        result t='str' counts as a value)."""
        if self.raw is None:
            return False
        return self.raw != "" or self.t == "str"

    @property
    def is_blank(self) -> bool:
        """No formula and no value (a styled empty cell)."""
        return self.formula is None and self.array is None and (self.raw is None or self.raw == "")

    @property
    def value(self):
        """Decoded value.  Raises GradingError for an untrusted formula result, and for any
        formula result read without needs_values (no fallback, see values.py)."""
        return self._ctx.value_of(self)

    @property
    def value_trusted(self) -> bool:
        return self._ctx.trusted(self)

    @property
    def unverified_value(self):
        """The value from its source (recalc copy or the file's cache) with NO trust or
        needs_values check.  Explicit opt-in for diagnostics, e.g. quoting a placeholder
        in an error message; never grade on it."""
        return self._ctx.unverified_value_of(self)

    @property
    def value_source(self) -> str:
        """'constant' | 'cached' | 'recalc'."""
        return self._ctx.source_of(self)

    @property
    def cached_value(self):
        """The delivered file's own stored value, decoded (ignores value_path)."""
        return self._ctx.decode(self.t, self.raw, self.ref)

    @property
    def rich_runs(self) -> Optional[list]:
        """[Run] for rich-text strings (inline or shared), None for plain values."""
        if self._runs is not None:
            return self._runs
        if self.t == "s" and self.raw:
            try:
                return self._ctx.sst_runs.get(int(self.raw))
            except ValueError:
                return None
        return None

    @property
    def sheet(self) -> str:
        return self._ctx.sheet_name

    def __repr__(self):
        f = f", f={self.formula.text!r}" if self.formula is not None else ""
        return f"Cell({self.ref}, t={self.t}, s={self.s}, raw={self.raw!r}{f})"


# ============================================================================ tail objects
@dataclass
class Merge:
    ref: str
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def is_single_cell(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2


@dataclass
class CfRule:
    type: Optional[str]
    priority: Optional[int] = None
    dxf_id: Optional[int] = None
    dxf: Optional[Dxf] = None              # inline <x14:dxf> of x14 rules (main rules: use dxf_id)
    operator: Optional[str] = None
    formulas: list = field(default_factory=list)
    stop_if_true: bool = False
    text: Optional[str] = None
    time_period: Optional[str] = None
    rank: Optional[int] = None
    percent: bool = False
    bottom: bool = False
    above_average: Optional[bool] = None
    equal_average: bool = False
    std_dev: Optional[int] = None
    colors: list = field(default_factory=list)   # [Color] of colorScale / dataBar
    cfvos: list = field(default_factory=list)    # [(type, val)]
    show_value: Optional[bool] = None            # dataBar / iconSet showValue (False = value hidden)
    icon_set: Optional[str] = None
    id: Optional[str] = None                     # x14 rule id / link id of a main rule's x14 twin
    attrs: dict = field(default_factory=dict)


@dataclass
class ConditionalFormat:
    sqref: str
    ranges: list                                 # [(r1, c1, r2, c2)]
    rules: list                                  # [CfRule]
    source: str = "main"                         # main | x14
    pivot: bool = False


@dataclass
class DataValidation:
    sqref: str
    ranges: list
    type: Optional[str] = None                   # list | whole | decimal | date | time | textLength | custom
    operator: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    allow_blank: bool = False
    show_error_message: bool = False
    show_input_message: bool = False
    error_style: Optional[str] = None
    source: str = "main"                         # main | x14
    attrs: dict = field(default_factory=dict)


@dataclass
class Hyperlink:
    ref: str
    rid: Optional[str]
    target: Optional[str]                        # relationship target (URL / file) when r:id given
    external: bool
    location: Optional[str]                      # in-workbook target ('Sheet2!A1')
    display: Optional[str] = None
    tooltip: Optional[str] = None


@dataclass
class IgnoredError:
    sqref: str
    ranges: list
    flags: dict                                  # {'numberStoredAsText': True, ...}


@dataclass
class Sparkline:
    sqref: str                                   # host cell(s) (<xm:sqref>), as stored
    formula: Optional[str]                       # data range (<xm:f>), as stored


@dataclass
class SparklineGroup:
    """An x14:sparklineGroup (extLst).  Its formulas reference cells like any formula."""
    type: str                                    # line | column | stacked
    date_range: Optional[str]                    # group <xm:f> (date axis range) or None
    sparklines: list                             # [Sparkline]
    attrs: dict = field(default_factory=dict)

    @property
    def formulas(self) -> list:
        """Every formula text of the group: date range first, then each sparkline's data range."""
        out = [self.date_range] if self.date_range else []
        return out + [s.formula for s in self.sparklines if s.formula]


@dataclass
class SheetTail:
    """Everything after </sheetData>.  Valid in sheet_end()."""
    merges: list = field(default_factory=list)                 # [Merge]
    conditional_formats: list = field(default_factory=list)    # [ConditionalFormat] (main + x14)
    data_validations: list = field(default_factory=list)       # [DataValidation] (main + x14)
    hyperlinks: list = field(default_factory=list)             # [Hyperlink]
    ignored_errors: list = field(default_factory=list)         # [IgnoredError]
    protection: Optional[dict] = None                          # sheetProtection attributes
    auto_filter: Optional[dict] = None                         # {'ref': ..., 'filter_columns': n}
    drawing_rid: Optional[str] = None
    legacy_drawing_rid: Optional[str] = None                   # VML (notes / form controls)
    legacy_drawing_hf_rid: Optional[str] = None
    picture_rid: Optional[str] = None                          # background picture
    table_part_rids: list = field(default_factory=list)
    has_ole_objects: bool = False
    has_controls: bool = False
    has_comments: bool = False                                 # legacy notes part related to the sheet
    has_threaded_comments: bool = False
    header_footer: dict = field(default_factory=dict)          # {'oddHeader': text, ...}
    page_setup: dict = field(default_factory=dict)
    sparkline_groups: list = field(default_factory=list)       # [SparklineGroup] (x14, extLst)
    other: list = field(default_factory=list)                  # local names of other elements seen;
                                                               # uninterpreted extLst content as 'extLst/<name>'

    @property
    def is_protected(self) -> bool:
        return bool(self.protection) and _t(self.protection.get("sheet"))


# ============================================================================ element appliers
def _sheet_view(el) -> SheetView:
    a = el.attrib
    v = SheetView(
        workbook_view_id=_i(a.get("workbookViewId"), 0) or 0,
        tab_selected=_t(a.get("tabSelected")),
        zoom_scale=_i(a.get("zoomScale")),
        zoom_scale_normal=_i(a.get("zoomScaleNormal")),
        zoom_scale_page_layout_view=_i(a.get("zoomScalePageLayoutView")),
        zoom_scale_sheet_layout_view=_i(a.get("zoomScaleSheetLayoutView")),
        view=a.get("view") or "normal",
        top_left_cell=a.get("topLeftCell"),
        show_grid_lines=not (a.get("showGridLines") or "1").strip().lower() in ("0", "false"),
        show_zeros=not (a.get("showZeros") or "1").strip().lower() in ("0", "false"),
        show_formulas=_t(a.get("showFormulas")),
        show_row_col_headers=not (a.get("showRowColHeaders") or "1").strip().lower() in ("0", "false"),
        show_outline_symbols=not (a.get("showOutlineSymbols") or "1").strip().lower() in ("0", "false"),
        right_to_left=_t(a.get("rightToLeft")),
        attrs=dict(a))
    for c in el:
        ln = _loc(c.tag)
        if ln == "pane":
            ca = c.attrib
            v.pane = Pane(_f(ca.get("xSplit")) or 0.0, _f(ca.get("ySplit")) or 0.0, ca.get("topLeftCell"),
                          ca.get("activePane") or "topLeft", ca.get("state") or "split")
        elif ln == "selection":
            ca = c.attrib
            v.selections.append(Selection(ca.get("pane") or "topLeft", ca.get("activeCell"),
                                          _i(ca.get("activeCellId"), 0) or 0, ca.get("sqref")))
    return v


def _cf_rule(el, source: str) -> CfRule:
    a = el.attrib
    r = CfRule(type=a.get("type"), priority=_i(a.get("priority")), dxf_id=_i(a.get("dxfId")),
               operator=a.get("operator"), stop_if_true=_t(a.get("stopIfTrue")), text=a.get("text"),
               time_period=a.get("timePeriod"), rank=_i(a.get("rank")), percent=_t(a.get("percent")),
               bottom=_t(a.get("bottom")),
               above_average=(None if a.get("aboveAverage") is None else _t(a.get("aboveAverage"))),
               equal_average=_t(a.get("equalAverage")), std_dev=_i(a.get("stdDev")), id=a.get("id"),
               attrs=dict(a))
    for c in el:
        ln = _loc(c.tag)
        if ln in ("formula", "f"):
            r.formulas.append(c.text or "")
        elif ln in ("colorScale", "dataBar", "iconSet"):
            if ln in ("dataBar", "iconSet"):
                r.show_value = (c.get("showValue") or "1").strip().lower() not in ("0", "false")
            if ln == "iconSet":
                r.icon_set = c.get("iconSet")
            for g in c:
                gl = _loc(g.tag)
                if gl == "cfvo":
                    fv = None
                    for h in g:
                        if _loc(h.tag) == "f":
                            fv = h.text
                    r.cfvos.append((g.get("type"), g.get("val") if g.get("val") is not None else fv))
                elif gl in ("color", "fillColor", "negativeFillColor", "borderColor", "negativeBorderColor", "axisColor"):
                    col = Color.from_el(g)
                    if col is not None:
                        r.colors.append(col)
        elif ln == "dxf":
            font = fill = border = None
            nf_id = nf_code = None
            for g in c:
                gl = _loc(g.tag)
                if gl == "font":
                    font = parse_font(g, partial=True)
                elif gl == "fill":
                    fill = parse_fill(g, dxf=True)
                elif gl == "numFmt":
                    nf_id, nf_code = _i(g.get("numFmtId")), g.get("formatCode")
                elif gl == "border":
                    border = parse_border(g)
            r.dxf = Dxf(font, fill, nf_id, nf_code, border)
        elif ln == "extLst":
            for g in c.iter():
                if _loc(g.tag) == "id" and g.text:
                    r.id = g.text.strip()
    return r


def _cf(el, source: str) -> ConditionalFormat:
    sq = el.get("sqref") or ""
    rules = []
    for c in el:
        ln = _loc(c.tag)
        if ln == "cfRule":
            rules.append(_cf_rule(c, source))
        elif ln == "sqref":
            sq = (c.text or "").strip()
    return ConditionalFormat(sq, parse_sqref(sq), rules, source, _t(el.get("pivot")))


def _dv(el, source: str) -> DataValidation:
    a = el.attrib
    sq = a.get("sqref") or ""
    f1 = f2 = None
    for c in el:
        ln = _loc(c.tag)
        if ln in ("formula1", "formula2"):
            txt = c.text
            for g in c:
                if _loc(g.tag) == "f":
                    txt = g.text
            if ln == "formula1":
                f1 = txt
            else:
                f2 = txt
        elif ln == "sqref":
            sq = (c.text or "").strip()
    return DataValidation(sq, parse_sqref(sq), a.get("type"), a.get("operator"), f1, f2,
                          _t(a.get("allowBlank")), _t(a.get("showErrorMessage")),
                          _t(a.get("showInputMessage")), a.get("errorStyle"), source, dict(a))


def _sparkline_group(g) -> SparklineGroup:
    date = None
    sps = []
    for c in g:
        cl = _loc(c.tag)
        if cl == "f":
            date = c.text
        elif cl == "sparklines":
            for sp in c:
                if _loc(sp.tag) != "sparkline":
                    continue
                f = sq = None
                for h in sp:
                    hl = _loc(h.tag)
                    if hl == "f":
                        f = h.text
                    elif hl == "sqref":
                        sq = (h.text or "").strip()
                sps.append(Sparkline(sq or "", f))
    return SparklineGroup(g.get("type") or "line", date, sps, dict(g.attrib))


def apply_top_level(el, head: SheetHead, tail: SheetTail, pkg: Package):
    """Apply one direct child of the <worksheet>/<chartsheet> root to head/tail."""
    ln = _loc(el.tag)
    a = el.attrib
    if ln == "sheetPr":
        sp = head.sheet_pr
        sp.attrs = dict(a)
        sp.code_name = a.get("codeName")
        sp.filter_mode = _t(a.get("filterMode"))
        for c in el:
            cl = _loc(c.tag)
            if cl == "tabColor":
                sp.tab_color = Color.from_el(c)
            elif cl == "outlinePr":
                sp.summary_below = (c.get("summaryBelow") or "1").strip().lower() not in ("0", "false")
                sp.summary_right = (c.get("summaryRight") or "1").strip().lower() not in ("0", "false")
                sp.apply_styles = _t(c.get("applyStyles"))
                sp.show_outline_symbols = (c.get("showOutlineSymbols") or "1").strip().lower() not in ("0", "false")
            elif cl == "pageSetUpPr":
                sp.fit_to_page = _t(c.get("fitToPage"))
    elif ln == "dimension":
        head.dimension = a.get("ref")
    elif ln == "sheetViews":
        head.views = [_sheet_view(v) for v in el if _loc(v.tag) == "sheetView"]
    elif ln == "sheetFormatPr":
        head.format = SheetFormat(
            default_row_height=_f(a.get("defaultRowHeight")), custom_height=_t(a.get("customHeight")),
            zero_height=_t(a.get("zeroHeight")), default_col_width=_f(a.get("defaultColWidth")),
            base_col_width=_f(a.get("baseColWidth")), outline_level_row=_i(a.get("outlineLevelRow"), 0) or 0,
            outline_level_col=_i(a.get("outlineLevelCol"), 0) or 0, thick_top=_t(a.get("thickTop")),
            thick_bottom=_t(a.get("thickBottom")))
    elif ln == "cols":
        for c in el:
            if _loc(c.tag) != "col":
                continue
            ca = c.attrib
            mn, mx = _i(ca.get("min")), _i(ca.get("max"))
            if mn is None or mx is None:
                continue
            st = ca.get("style")
            head.cols.append(ColInfo(mn, mx, _f(ca.get("width")), _t(ca.get("hidden")), _t(ca.get("customWidth")),
                                     _t(ca.get("bestFit")), _i(ca.get("outlineLevel"), 0) or 0,
                                     _t(ca.get("collapsed")), _i(st) if st is not None else None))
        head.cols.sort(key=lambda c: c.min)       # stable: file order among equal mins
        head._col_segs = None                     # col_segments() rebuilds lazily
    elif ln == "mergeCells":
        for c in el:
            if _loc(c.tag) != "mergeCell":
                continue
            ref = (c.get("ref") or "").strip()
            b = _grid_range(ref, "mergeCell ref", f"{pkg.file_name}:{head.info.part} (sheet {head.info.name!r})")
            tail.merges.append(Merge(ref.replace("$", ""), *b))
    elif ln == "conditionalFormatting":
        tail.conditional_formats.append(_cf(el, "main"))
    elif ln == "dataValidations":
        for c in el:
            if _loc(c.tag) == "dataValidation":
                tail.data_validations.append(_dv(c, "main"))
    elif ln == "hyperlinks":
        rels = {r.id: r for r in pkg.rels(head.info.part)} if head.info.part else {}
        for c in el:
            if _loc(c.tag) != "hyperlink":
                continue
            rid = rel_id(c)
            rel = rels.get(rid) if rid else None
            tail.hyperlinks.append(Hyperlink(c.get("ref") or "", rid, rel.target if rel else None,
                                             bool(rel and rel.external), c.get("location"), c.get("display"),
                                             c.get("tooltip")))
    elif ln == "ignoredErrors":
        for c in el:
            if _loc(c.tag) == "ignoredError":
                sq = c.get("sqref") or ""
                flags = {k: _t(v) for k, v in c.attrib.items() if k != "sqref"}
                tail.ignored_errors.append(IgnoredError(sq, parse_sqref(sq), flags))
    elif ln == "sheetProtection":
        tail.protection = {_loc(k): v for k, v in a.items()}
    elif ln == "autoFilter":
        tail.auto_filter = {"ref": a.get("ref"), "filter_columns": sum(1 for c in el if _loc(c.tag) == "filterColumn")}
    elif ln == "drawing":
        tail.drawing_rid = rel_id(el)
    elif ln == "legacyDrawing":
        tail.legacy_drawing_rid = rel_id(el)
    elif ln == "legacyDrawingHF":
        tail.legacy_drawing_hf_rid = rel_id(el)
    elif ln == "picture":
        tail.picture_rid = rel_id(el)
    elif ln == "tableParts":
        tail.table_part_rids = [rel_id(c) for c in el if _loc(c.tag) == "tablePart"]
    elif ln == "oleObjects":
        tail.has_ole_objects = True
    elif ln == "controls":
        tail.has_controls = True
    elif ln == "headerFooter":
        for c in el:
            if c.text:
                tail.header_footer[_loc(c.tag)] = c.text
    elif ln in ("pageSetup", "pageMargins", "printOptions"):
        tail.page_setup[ln] = dict(a)
    elif ln == "extLst":
        for ext in el:
            for x in ext:
                xl = _loc(x.tag)
                if xl in ("conditionalFormattings", "conditionalFormatting"):
                    for cf in (x if xl.endswith("s") else [x]):
                        if _loc(cf.tag) == "conditionalFormatting":
                            tail.conditional_formats.append(_cf(cf, "x14"))
                elif xl in ("dataValidations", "dataValidation"):
                    for dv in (x if xl.endswith("s") else [x]):
                        if _loc(dv.tag) == "dataValidation":
                            tail.data_validations.append(_dv(dv, "x14"))
                elif xl in ("sparklineGroups", "sparklineGroup"):
                    for g in (x if xl.endswith("s") else [x]):
                        if _loc(g.tag) == "sparklineGroup":
                            tail.sparkline_groups.append(_sparkline_group(g))
                else:
                    tail.other.append(f"extLst/{xl}")
    elif ln == "sheetData":
        pass
    else:
        tail.other.append(ln)


# ============================================================================ value context
class SheetContext:
    """Per-sheet value decoding / provenance (built by values.py, shared by all cells)."""

    def __init__(self, sheet_name: str, pkg: Package, policy=None, cursor=None, values_enabled: bool = False):
        self.sheet_name = sheet_name
        self.pkg = pkg
        self.policy = policy          # values.TrustPolicy (None = trust all cached values)
        self.cursor = cursor          # values.SheetValueCursor or None
        # formula results may be read through cell.value only when the stream was opened for a
        # check that declared needs_values (otherwise a value_path copy would be ignored silently)
        self.values_enabled = values_enabled

    @property
    def sst(self) -> list:
        return self.pkg.shared_strings          # loaded lazily, once per workbook

    @property
    def sst_runs(self) -> dict:
        return self.pkg.shared_string_runs

    def decode(self, t, raw, ref="?"):
        if raw is None:
            return None
        if t == "n":
            if raw == "":
                return None
            try:
                return float(raw)
            except ValueError:
                raise GradingError(f"{self.sheet_name}!{ref}: cannot decode number {raw!r}") from None
        if t == "s":
            if raw == "":
                return None
            try:
                return self.sst[int(raw)]
            except (ValueError, IndexError):
                raise GradingError(f"{self.sheet_name}!{ref}: shared string index {raw!r} out of range "
                                   f"({len(self.sst)} strings)") from None
        if t in ("str", "inlineStr"):
            return raw
        if t == "b":
            return raw.strip().lower() in ("1", "true")
        if t == "e":
            return ExcelError(raw)
        return raw   # 'd' (ISO 8601 text) and anything unknown: as stored

    def unverified_value_of(self, cell: "Cell"):
        if cell._rv is not None:
            return cell._rv[1]
        return self.decode(cell.t, cell.raw, cell.ref)

    def value_of(self, cell: "Cell"):
        """cell.value: constants as decoded; formula results only when values are enabled for
        this stream AND trusted.  Otherwise GradingError (no fallback)."""
        if not cell.is_formula_result:
            return self.decode(cell.t, cell.raw, cell.ref)
        if not self.values_enabled:
            raise GradingError(f"{self.sheet_name}!{cell.ref}: the value of a formula result was read by a "
                               f"check that does not declare needs_values = True")
        if not self.trusted(cell):
            prov = self.pkg.provenance
            why = (f"writer={prov.writer}, value_path={'given' if prov.value_path else 'none'}"
                   if prov is not None else "no provenance")
            raise GradingError(f"{self.sheet_name}!{cell.ref}: formula value is untrusted "
                               f"(source={self.source_of(cell)}, {why}); test cell.value_trusted first "
                               f"or use Check.require_value")
        return self.unverified_value_of(cell)

    def source_of(self, cell: "Cell") -> str:
        if not cell.is_formula_result:
            return "constant"
        return "recalc" if cell._rv is not None else "cached"

    def trusted(self, cell: "Cell") -> bool:
        if not cell.is_formula_result:
            return True
        if cell._rv is not None:
            return cell._rv[0] == "ok"
        if self.policy is None:
            return cell.has_value
        return self.policy.cached_trusted(cell)


# ============================================================================ streaming
class SheetStream:
    """Read one sheet part in phases.  Usage:

        ss = SheetStream(pkg, info)
        head = ss.read_head()
        tail = ss.read_body(on_row, on_cell, ctx)     # or ss.skip_body()
        ss.close()
    """

    def __init__(self, pkg: Package, info: SheetInfo):
        self.pkg = pkg
        self.info = info
        self.head: Optional[SheetHead] = None
        self.tail = SheetTail()
        self._fh = None
        self._head_bytes = b""
        self._rest = b""
        self._found = False
        self._qname = b""
        self._root_end = 0

    def close(self):
        if self._fh is not None:
            self._fh.close()
            self._fh = None

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def _where(self) -> str:
        return f"{self.pkg.file_name}:{self.info.part} (sheet {self.info.name!r})"

    def _finish_tail(self) -> SheetTail:
        rels = self.pkg.rels(self.info.part) if self.info.part else []
        self.tail.has_comments = any(r.type_short.lower() == "comments" for r in rels)
        self.tail.has_threaded_comments = any(r.type_short.lower() in ("threadedcomment", "threadedcomments")
                                              for r in rels)
        return self.tail

    # ------------------------------------------------------------ phase 1
    def read_head(self) -> SheetHead:
        head = SheetHead(self.info)
        self.head = head
        if not self.info.part or not self.pkg.has(self.info.part):
            raise GradingError(f"sheet {self.info.name!r}: part {self.info.part!r} is missing from {self.pkg.file_name}")
        if self.info.kind == "unknown":
            raise GradingError(f"sheet {self.info.name!r} in {self.pkg.file_name}: relationship type "
                               f"{self.info.rel_type!r} is not a worksheet, chartsheet, dialog or macro sheet; "
                               f"cannot tell how to read it")
        fh = self.pkg.open_part(self.info.part)
        first = fh.read(CHUNK)
        if first[:2] in (codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE) or first[:4] in (b"<\x00?\x00", b"\x00<\x00?"):
            # rare: UTF-16 sheet part -> transcode to UTF-8 in memory
            data = first + fh.read()
            fh.close()
            text = data.decode("utf-16")
            text = re.sub(r"^\ufeff?<\?xml[^>]*\?>", "", text)
            fh = io.BytesIO(text.encode("utf-8"))
            first = fh.read(CHUNK)
        self._fh = fh
        buf = bytearray(first)
        mroot = _find_root(buf)
        while mroot is None:
            chunk = fh.read(CHUNK) if len(buf) < 16 * CHUNK else b""
            if not chunk:
                raise GradingError(f"no root element start tag in {self._where()}")
            buf += chunk
            mroot = _find_root(buf)
        self._qname = bytes(mroot.group(1))
        self._root_end = mroot.end()
        searched = mroot.end()
        m = None
        while True:
            m = _SD_START.search(buf, max(mroot.end(), searched - 64))
            if m is not None:
                break
            searched = len(buf)
            chunk = fh.read(CHUNK)
            if not chunk:
                break
            buf += chunk
        if m is None:
            # no sheetData at all (chartsheet / empty sheet): the whole document is "head"
            self._found = False
            self._head_bytes = bytes(buf)
            head.has_sheet_data = False
            doc = self._head_bytes
        else:
            self._found = True
            self._head_bytes = bytes(buf[:m.start()])
            self._rest = bytes(buf[m.start():])
            doc = self._head_bytes + b"</" + self._qname + b">"
        try:
            root = ET.fromstring(doc)
        except ET.ParseError as e:
            raise GradingError(f"XML parse error in head of {self._where()}: {e}") from e
        for el in root:
            apply_top_level(el, head, self.tail, self.pkg)
        return head

    # ------------------------------------------------------------ phase 2+3, skipping cells
    def skip_body(self) -> SheetTail:
        """Skip <sheetData> at byte level and parse the tail."""
        if not self._found:
            return self._finish_tail()
        fh = self._fh
        buf = self._rest
        # self-closing <sheetData/>
        gt = buf.find(b">")
        while gt < 0:
            chunk = fh.read(CHUNK)
            if not chunk:
                raise GradingError(f"truncated sheetData start tag in {self._where()}")
            buf += chunk
            gt = buf.find(b">")
        if buf[gt - 1:gt] == b"/":
            tail_bytes = buf[gt + 1:] + fh.read()
        else:
            pos = 0
            while True:
                m = _SD_END.search(buf, pos)
                if m is not None:
                    tail_bytes = buf[m.end():] + fh.read()
                    break
                chunk = fh.read(CHUNK)
                if not chunk:
                    raise GradingError(f"</sheetData> not found in {self._where()}")
                keep = buf[-64:]
                buf = keep + chunk
                pos = 0
        # rebuild a minimal document: prolog + original root start tag + tail
        doc = self._head_bytes[:self._root_end] + tail_bytes
        try:
            root = ET.fromstring(doc)
        except ET.ParseError as e:
            raise GradingError(f"XML parse error in tail of {self._where()}: {e}") from e
        for el in root:
            apply_top_level(el, self.head, self.tail, self.pkg)
        return self._finish_tail()

    # ------------------------------------------------------------ phase 2+3, streaming cells
    def read_body(self, on_row: Optional[Callable], on_cell: Optional[Callable], ctx: SheetContext) -> SheetTail:
        """Stream rows/cells to on_row(Row) / on_cell(Cell), then parse the tail."""
        if not self._found:
            return self._finish_tail()
        head, tail, pkg = self.head, self.tail, self.pkg
        parser = ET.XMLPullParser(events=("start", "end"))
        try:
            parser.feed(self._head_bytes)
            for _ev, _el in parser.read_events():
                pass
        except ET.ParseError as e:
            raise GradingError(f"XML parse error in {self._where()}: {e}") from e
        # the parser now sits inside the root element (depth 1); head children already applied
        depth = 1
        in_sd = False
        sd_el = None
        cur_row = 0
        last_col = 0
        masters: dict = {}
        active: list = []          # open ArrayRange objects
        cursor = ctx.cursor
        loc = _loc
        LOCAL = _LOCAL
        col_cache = {}
        ref_match = _REF.match
        fh = self._fh
        data = self._rest
        while True:
            if data:
                try:
                    parser.feed(data)
                except ET.ParseError as e:
                    raise GradingError(f"XML parse error in {self._where()}: {e}") from e
                for ev, el in parser.read_events():
                    if ev == "start":
                        depth += 1
                        if in_sd:
                            if depth == 3:
                                a = el.attrib
                                r = a.get("r")
                                if r is not None:
                                    try:
                                        cur_row = int(r)
                                    except ValueError:
                                        raise GradingError(f"bad row number {r!r} in {self._where()}") from None
                                else:
                                    cur_row += 1
                                last_col = 0
                                if active:
                                    active = [x for x in active if x.r2 >= cur_row]
                                if on_row is not None:
                                    on_row(Row(cur_row, a))
                        elif depth == 2:
                            tag = el.tag
                            ln = LOCAL.get(tag) or loc(tag)
                            if ln == "sheetData":
                                in_sd = True
                                sd_el = el
                        continue
                    # ---------------- end events
                    if in_sd:
                        if depth == 4:
                            # <c> (anything else at this depth is ignored)
                            tag = el.tag
                            ln = LOCAL.get(tag) or loc(tag)
                            if ln == "c":
                                a = el.attrib
                                ref = a.get("r")
                                if ref is not None:
                                    m = ref_match(ref)
                                    if m is None:
                                        raise GradingError(f"bad cell reference {ref!r} in {self._where()}")
                                    letters = m.group(1)
                                    col = col_cache.get(letters)
                                    if col is None:
                                        col = col_cache[letters] = col_to_index(letters)
                                    row = int(m.group(2))
                                    if "$" in ref:
                                        ref = ref.replace("$", "")
                                else:
                                    row = cur_row
                                    col = last_col + 1
                                    ref = f"{index_to_col(col)}{row}"
                                last_col = col
                                t = a.get("t") or "n"
                                s = a.get("s")
                                s = int(s) if s else 0
                                raw = None
                                fel = None
                                runs = None
                                for ch in el:
                                    cl = LOCAL.get(ch.tag) or loc(ch.tag)
                                    if cl == "v":
                                        raw = ch.text
                                        if raw is None:
                                            raw = ""
                                    elif cl == "f":
                                        fel = ch
                                    elif cl == "is":
                                        raw, runs = parse_runs(ch)
                                formula = None
                                if fel is not None:
                                    fa = fel.attrib
                                    fk = fa.get("t") or "normal"
                                    ftext = fel.text
                                    si = fa.get("si")
                                    master = None
                                    if fk == "shared":
                                        if ftext:
                                            master = masters[si] = SharedMaster(si, ftext, row, col, fa.get("ref"))
                                        else:
                                            master = masters.get(si)
                                    formula = Formula(ftext, fk, fa.get("ref"), si, fa, master, row, col)
                                    if fk in ("array", "dataTable"):
                                        fr = fa.get("ref")
                                        if fr:
                                            b = _grid_range(fr, f"{fk} formula ref at {ref}", self._where())
                                            if (b[0], b[1], b[2], b[3]) != (row, col, row, col):
                                                active.append(ArrayRange(b, ref, formula))
                                member = None
                                if active and (formula is None or formula.is_empty_marker):
                                    for x in active:
                                        if x.r1 <= row <= x.r2 and x.c1 <= col <= x.c2:
                                            member = x
                                            break
                                rv = None
                                if cursor is not None and (formula is not None or member is not None):
                                    rv = cursor.get(row, col)
                                cm = a.get("cm")
                                vm = a.get("vm")
                                cell = Cell(ref, row, col, t, s, raw, formula, member,
                                            int(cm) if cm else None, int(vm) if vm else None, runs, ctx, rv)
                                if on_cell is not None:
                                    on_cell(cell)
                            el.clear()
                        elif depth == 3:
                            sd_el.clear()
                        elif depth == 2:
                            in_sd = False
                            sd_el.clear()
                    elif depth == 2:
                        apply_top_level(el, head, tail, pkg)
                        el.clear()
                    depth -= 1
            data = fh.read(CHUNK)
            if not data:
                break
        try:
            parser.close()
        except ET.ParseError as e:
            raise GradingError(f"XML parse error at end of {self._where()}: {e}") from e
        for ev, el in parser.read_events():   # normally nothing left but the root's end
            if ev == "start":
                depth += 1
                continue
            if depth == 2 and not in_sd:
                apply_top_level(el, head, tail, pkg)
            depth -= 1
        return self._finish_tail()


# ============================================================================ convenience
def load_sheet(pkg: Package, name: str, ctx: Optional[SheetContext] = None):
    """Materialise one sheet: (head, rows, cells, tail).  For tests / debugging of SMALL
    sheets only - checks must stream through the engine instead.  Formula values are
    readable (trusted ones only, as for a needs_values check)."""
    info = pkg.sheet(name)
    if info is None:
        raise GradingError(f"no sheet named {name!r} in {pkg.file_name}")
    if ctx is None:
        from .values import make_context
        ctx = make_context(pkg, info, None, needs_values=True)
    rows, cells = [], []
    with SheetStream(pkg, info) as ss:
        head = ss.read_head()
        tail = ss.read_body(rows.append, cells.append, ctx)
    return head, rows, cells, tail
