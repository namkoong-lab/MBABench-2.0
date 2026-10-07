"""73 Formatting/Reasonable row heights.

Rule (handoff bucket 2): a row fails only if it is over 60 pt AND more than twice as tall as its
content needs.  Every worksheet, dialog sheet and macro sheet of the delivered workbook is graded
(hidden sheets, case tabs and the Instructions sheet included: no exemption, no starting-file
diff).  Hidden rows (hidden="1", ht=0, or no height on a zeroHeight sheet) are ignored.

Row height = the row's stored ht when it has customHeight="1"; a stored ht WITHOUT customHeight
is ignored, and so is a sheetFormatPr defaultRowHeight without customHeight="1": Excel auto-fits
such rows to their content (measured by Patrick 2026-10-03: a non-custom 60 pt row and an 80 pt
non-custom sheet default open at 15 pt; CUSTOM_HEIGHT_ONLY), so they can never be too tall.  A
row without a (custom) height takes the sheet's custom defaultRowHeight, else one Normal-font
line.  Only rows taller than MAX_ROW_HEIGHT_PT are
examined ("candidates"); a candidate fails when height > EXCESS_FACTOR x need.  A sheet whose
default row height is itself excessive (over the cap and over 2x one Normal line) gets one
sheet-level mistake; its rows without a height of their own are covered by it and not listed
again row by row.

Content need of a row = the largest need of its displayed cells, never less than one line of the
Normal font (LINE_HEIGHT_PER_PT x Normal size: Excel sizes a row for the Normal font of its
unformatted cells too).  Need of one cell = lines x line height:
  * line height = LINE_HEIGHT_PER_PT (1.3) x font size in pt: the cell font's size, or for a
    rich-text string the largest size among its runs (a run without its own size has the cell's;
    the cell font itself counts only for text outside the runs);
  * lines = 1, unless wrap is on (wrapText, or horizontal/vertical justify/distributed, which
    Excel wraps); then a greedy word wrap of the text into the usable width: the column width
    (the merged width when the cell anchors a merge; hidden columns count 0) converted to pixels
    with Excel's own formula (ECMA-376 col/@width, the Normal font's maximum digit width at 96
    dpi) minus CELL_MARGIN_PX, minus the alignment indent (ECMA: one level = 3 spaces of the
    Normal font; both sides for distributed).  Text width = sum of per-character
    Helvetica/Arial advance widths (AFM table, stdlib) x font size x the face's width factor
    (FACE_WIDTH: measured once from the real font files, the wider of prose and capitals/digits;
    bold separately; unknown faces as Arial).  Typed line breaks start new lines (only with wrap
    on: LINE_BREAKS_NEED_WRAP); a word wider than a line breaks by characters;
  * numbers, booleans, errors and dates never wrap (1 line);
  * rotated text (textRotation 1..180) needs its width projected on the vertical
    (width x sin(angle) + lines x line height x cos(angle)); vertical stacked text (255) needs
    one line per character.
Merged cells: covered (non-anchor) cells of a merge are ignored; an anchor uses the merged
width; a merge spanning several rows is judged as a whole block (MERGE_ROWS_WHOLE_BLOCK, sanity
Rec 73-6, 2026-10-04): the block's height is the sum of its visible rows' heights, and each row
is credited with the merge's need in proportion to its own height (need x row / block), so a
row's ratio through the merge equals the block's ratio and a 48 pt title over 40 + 62 pt rows
(102 pt for 62.4 pt of text) passes.  With MERGE_ROWS_WHOLE_BLOCK = False the row gets the
merge's total need minus the other rows' heights instead (per-row share, the behaviour before
2026-10-04, which failed that row 6 at 2.8x).  An anchor in an earlier row that is not a
candidate is read in a second pass when the row cannot be decided otherwise.  Cells in hidden or
zero-width columns are not displayed and need nothing.

Values: a constant's text is read from the file.  A formula result's value is needed only for a
wrapped or rotated formula cell (an unwrapped formula cell needs one line of its font whatever
its value), and only for a candidate row that every value-independent cell - off-row merge
anchors included - leaves undecided.  Such rows are settled in finish(): first with trusted
values only (they never raise; every row they settle is listed correctly).  An untrusted value is
then required only while the workbook's verdict is still open (UNTRUSTED_ONLY_IF_VERDICT_NEEDS):
if any row already fails, the rows that would need an untrusted value are left undecided (listed
in stats) and nothing is raised; otherwise the first such value raises GradingError (no
fallback).

Calibration: Excel AutoFit figures from the toy Overview - one line of 10 pt -> Normal floor
14.3 pt (Excel 14.4); C54 of toy T3, 147 characters of Arial 10 in a 33.9-character column -> 4
lines x 13 = 52 pt (Excel 55.2, 4 lines); 18 pt title -> 23.4 (23.4); 48 pt title -> 62.4
(60.6).  Toy Fail rows come out at 3.5x-8.4x, the Instructions title row at 0.97x.  Against a
glyph-metric wrap (fontTools, real font files; detchecks/scratch/rowscols/xval73.py) on all 2,473
rows over 60 pt of 374 real attempts, this estimate never fails a row the glyph model passes.
"""
from __future__ import annotations

import math
import unicodedata
from array import array
from bisect import bisect_left, bisect_right

from ..core.refs import location, make_ref
from ..errors import GradingError
from .base import Check

# ---------------------------------------------------------------- rule constants
MAX_ROW_HEIGHT_PT = 60.0       # rubric rule of thumb: rows taller than ~60 pt are examined
EXCESS_FACTOR = 2.0            # ... and fail when more than 2x as tall as their content needs
LINE_HEIGHT_PER_PT = 1.3       # one text line = 1.3 x font size (pt)
DEFAULT_FONT_PT = 11.0         # font size when a font has no <sz>
WIDTH_SAFETY = 1.0             # multiplies every estimated text width (> 1 = more lines = more lenient)
CELL_MARGIN_PX = 5             # column pixels not available to wrapped text (margins + gridline)
PX_PER_PT = 96.0 / 72.0        # Excel's column geometry is defined at 96 dpi
DEFAULT_BASE_COL_WIDTH = 8.0   # baseColWidth when absent
LINE_BREAKS_NEED_WRAP = True   # typed line breaks start new lines only when wrap is on (Excel shows a
CUSTOM_HEIGHT_ONLY = True      # Excel 2026-10-03: heights without customHeight (row ht, sheet default) are
                               # auto-fitted by Excel, so only customHeight="1" heights are graded
                               # non-wrapped cell on one line)
WRAP_ALIGNMENTS = ("justify", "distributed")   # alignments Excel wraps like wrapText
INDENT_SPACES_PER_LEVEL = 3    # ECMA-376 alignment/@indent: one level = 3 spaces of the Normal font
INDENT_ALIGNMENTS = (None, "general", "left", "right", "distributed")   # alignments an indent applies to
MERGE_ROWS_WHOLE_BLOCK = True  # a merge over several rows is judged as one block (sanity Rec 73-6): each row
                               # is credited with the merge's need x (its height / the block's visible height);
                               # False = per-row share (merge need minus the other rows' heights, pre-2026-10-04)
UNTRUSTED_ONLY_IF_VERDICT_NEEDS = True   # an untrusted formula value is required (GradingError) only
                                         # when no row fails otherwise; False = also when the file
                                         # already fails (the behaviour before the review)
MAX_BUFFERED_CELLS = 500_000   # wrapped/rotated text cells buffered from candidate rows of one sheet;
                               # more raises (one-line cells are kept as compact column lists)
MAX_LISTED_OK_ROWS = 20

# ---------------------------------------------------------------- text metrics (stdlib only)
# Advance widths of Helvetica (Adobe AFM, 1/1000 em); Arial and Liberation Sans are metric-compatible.
_HELV_ASCII = (
    " 278 ! 278 \" 355 # 556 $ 556 % 889 & 667 ' 191 ( 333 ) 333 * 389 + 584 , 278 - 333 . 278 / 278 "
    "0 556 1 556 2 556 3 556 4 556 5 556 6 556 7 556 8 556 9 556 : 278 ; 278 < 584 = 584 > 584 ? 556 @ 1015 "
    "A 667 B 667 C 722 D 722 E 667 F 611 G 778 H 722 I 278 J 500 K 667 L 556 M 833 N 722 O 778 P 667 "
    "Q 778 R 722 S 667 T 611 U 722 V 667 W 944 X 667 Y 667 Z 611 [ 278 \\ 278 ] 278 ^ 469 _ 556 ` 333 "
    "a 556 b 556 c 500 d 556 e 556 f 278 g 556 h 556 i 222 j 222 k 500 l 222 m 833 n 556 o 556 p 556 "
    "q 556 r 333 s 500 t 278 u 556 v 500 w 722 x 500 y 500 z 500 { 334 | 260 } 334 ~ 584")
_HELV: dict = {}
_tok = _HELV_ASCII.split(" ")
_HELV[" "] = 0.278
for _i in range(2, len(_tok) - 1, 2):
    _HELV[_tok[_i]] = int(_tok[_i + 1]) / 1000.0
_HELV.update({"’": 0.222, "‘": 0.222, "“": 0.333, "”": 0.333, "–": 0.556,
              "—": 1.0, "•": 0.35, "…": 1.0, "×": 0.584, "÷": 0.584, "€": 0.556,
              "£": 0.556, "¥": 0.556, "°": 0.4, "±": 0.584, "·": 0.278, " ": 0.278,
              "→": 1.0, "←": 1.0, "↑": 1.0, "↓": 1.0, "≤": 0.584, "≥": 0.584,
              "−": 0.584, "≈": 0.584, "≠": 0.584, "\t": 0.278})
_HELV_DEFAULT = 0.556


def helv_em(ch: str, _bold: bool = False) -> float:
    """Advance width of one character in em (Helvetica/Arial metrics; accented letters as their
    base letter; East-Asian wide characters 1 em; anything else 0.556 em)."""
    w = _HELV.get(ch)
    if w is not None:
        return w
    base = unicodedata.normalize("NFD", ch)[:1]
    w = _HELV.get(base)
    if w is None:
        w = 1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else _HELV_DEFAULT
    _HELV[ch] = w
    return w


def text_em(text: str) -> float:
    return sum(helv_em(ch) for ch in text)


# Width of each face relative to Helvetica/Arial, (regular, bold): the larger of the prose and the
# capitals/digits ratios measured from the real font files (detchecks/scratch/rowscols/face_factors.py).
# Unknown faces are measured as Arial.
_W_ARIAL, _W_CALIBRI, _W_TIMES, _W_COURIER = (1.0, 1.076), (0.918, 0.939), (0.937, 0.976), (1.382, 1.382)
FACE_WIDTH = {
    "arial": _W_ARIAL, "helvetica": _W_ARIAL, "liberation sans": _W_ARIAL, "arimo": _W_ARIAL,
    "calibri": _W_CALIBRI, "calibri light": _W_CALIBRI, "carlito": _W_CALIBRI,
    "aptos narrow": (0.890, 0.924), "aptos": (0.969, 1.022), "aptos display": (0.969, 1.022),
    "times new roman": _W_TIMES, "tinos": _W_TIMES, "liberation serif": _W_TIMES,
    "courier new": _W_COURIER, "cousine": _W_COURIER, "liberation mono": _W_COURIER,
    "verdana": (1.149, 1.282), "tahoma": (1.003, 1.141), "georgia": (1.011, 1.156),
    "consolas": (1.266, 1.266), "cambria": (0.959, 1.031), "century gothic": (1.096, 1.094),
    "candara": (0.948, 0.963), "corbel": (0.923, 0.968), "arial narrow": (0.820, 0.883),
    "dejavu sans": (1.141, 1.291),
}
UNKNOWN_FACE_WIDTH = (1.0, 1.08)     # e.g. Roboto, Segoe UI: measured as Arial
# Maximum digit width of each face in em (Excel's column unit is the Normal font's digit width).
DIGIT_EM = {
    "arial": 0.556, "helvetica": 0.556, "liberation sans": 0.556, "arimo": 0.556,
    "calibri": 0.507, "calibri light": 0.507, "carlito": 0.507, "aptos narrow": 0.507,
    "aptos": 0.534, "aptos display": 0.534, "times new roman": 0.5, "tinos": 0.5, "liberation serif": 0.5,
    "courier new": 0.6, "cousine": 0.6, "liberation mono": 0.6, "verdana": 0.636, "tahoma": 0.546,
    "georgia": 0.614, "consolas": 0.55, "cambria": 0.554, "century gothic": 0.554, "candara": 0.552,
    "corbel": 0.524, "arial narrow": 0.456, "dejavu sans": 0.636, "roboto": 0.562,
}
DEFAULT_DIGIT_EM = 0.556
# Space advance of each face in em (alignment indent = 3 spaces of the Normal font; measured from the
# real font files: detchecks/scratch/rowscols/space_em.py).
SPACE_EM = {
    "arial": 0.278, "helvetica": 0.278, "liberation sans": 0.278, "arimo": 0.278,
    "calibri": 0.226, "calibri light": 0.226, "carlito": 0.226, "aptos narrow": 0.187,
    "aptos": 0.203, "aptos display": 0.203, "times new roman": 0.25, "tinos": 0.25, "liberation serif": 0.25,
    "courier new": 0.6, "cousine": 0.6, "liberation mono": 0.6, "verdana": 0.352, "tahoma": 0.312,
    "georgia": 0.241, "consolas": 0.55, "cambria": 0.22, "century gothic": 0.277, "candara": 0.217,
    "corbel": 0.2, "arial narrow": 0.228, "dejavu sans": 0.318,
}
DEFAULT_SPACE_EM = 0.278


def face_width(name, bold: bool) -> float:
    f = FACE_WIDTH.get((name or "").strip().lower(), UNKNOWN_FACE_WIDTH)
    return f[1] if bold else f[0]


def mdw_px(normal_face, normal_size: float) -> int:
    """Excel's maximum digit width (px at 96 dpi) of the Normal font."""
    em = DIGIT_EM.get((normal_face or "").strip().lower(), DEFAULT_DIGIT_EM)
    return max(1, int(round(em * (normal_size or DEFAULT_FONT_PT) * PX_PER_PT)))


def space_px(normal_face, normal_size: float) -> int:
    """Width (px at 96 dpi, whole pixels) of one space of the Normal font."""
    em = SPACE_EM.get((normal_face or "").strip().lower(), DEFAULT_SPACE_EM)
    return max(1, int(round(em * (normal_size or DEFAULT_FONT_PT) * PX_PER_PT)))


def col_px(width_ch: float, mdw: int) -> int:
    """ECMA-376 col/@width (characters incl. padding) -> pixels."""
    if width_ch <= 0:
        return 0
    return int(math.trunc(((256.0 * width_ch + math.trunc(128.0 / mdw)) / 256.0) * mdw))


def _paragraphs(text: str, wrap: bool) -> list:
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    if wrap or not LINE_BREAKS_NEED_WRAP:
        return t.split("\n")
    return [t.replace("\n", "")]


def wrap_lines(text: str, px_per_em: float, usable_px: float, em=None) -> int:
    """Greedy word wrap (Excel-like) of text into lines `usable_px` wide; a character is
    em(ch) x px_per_em pixels (em defaults to helv_em, this check's Helvetica/Arial table;
    Sufficient column widths (69) passes its own per-glyph widths in pixels with px_per_em 1).
    Each typed line break starts a paragraph (an empty paragraph is one line); a word wider
    than a line breaks by characters."""
    if em is None:
        em = helv_em
    usable = max(usable_px, 1.0)
    space = em(" ") * px_per_em
    total = 0
    for para in _paragraphs(text, True):
        lines = 1
        cur = 0.0                    # width used on the current line (0 = empty line)
        for word in para.split(" "):
            wp = sum(em(ch) for ch in word) * px_per_em
            if cur > 0:
                if cur + space + wp <= usable + 1e-6:
                    cur += space + wp
                    continue
                lines += 1
                cur = 0.0
            if wp <= usable + 1e-6:
                cur = wp
                continue
            for ch in word:          # a word wider than a line breaks by characters
                cw = em(ch) * px_per_em
                if cur > 0 and cur + cw > usable + 1e-6:
                    lines += 1
                    cur = 0.0
                cur += cw
        total += lines
    return total


def cell_need(text, size: float, wfac: float, wrap: bool, rot: int, usable_px: float):
    """(need in pt, lines) of one displayed cell.  text None = no text to lay out (a number,
    or a value-independent one-line cell); wfac = face width factor; usable_px = text width
    available (column or merged width in px minus CELL_MARGIN_PX)."""
    line = LINE_HEIGHT_PER_PT * size
    if text is None:
        text = ""
        if not rot:
            return line, 1
    if rot == 255:                                   # vertical stacked letters
        paras = _paragraphs(text, wrap)
        n = max(1, max(len(p) for p in paras))
        return n * line, n
    if rot:
        ang = rot if rot <= 90 else rot - 90
        th = math.radians(min(max(ang, 0), 90))
        paras = _paragraphs(text, wrap)
        w_pt = max(text_em(p) for p in paras) * size * wfac * WIDTH_SAFETY
        return w_pt * math.sin(th) + len(paras) * line * math.cos(th), len(paras)
    if not wrap:
        n = 1 if LINE_BREAKS_NEED_WRAP else len(_paragraphs(text, False))
        return n * line, n
    lines = wrap_lines(text, size * PX_PER_PT * wfac * WIDTH_SAFETY, usable_px)
    return lines * line, lines


class _Cand:
    """A candidate row (visible, taller than the cap) and its displayed cells."""
    __slots__ = ("r", "ht", "custom", "own", "cells", "small", "n_cells", "need", "why", "pending",
                 "deferred", "status")

    def __init__(self, r, ht, custom, own):
        self.r, self.ht, self.custom, self.own = r, ht, custom, own
        self.cells = []          # layout cells [(col, size, wfac, wrap, rot, text, cell_or_None, ref, indent_px)]
        self.small = {}          # one-line cells: font size -> array of columns
        self.n_cells = 0
        self.need = 0.0
        self.why = ""
        self.pending = []        # merges whose anchor (in a non-candidate row) is read in pass 2
        self.deferred = []       # [(d, merge, px, others)] formula cells whose value would decide
        self.status = None       # 'ok' | 'fail' | 'values' (open: needs formula values) | None (pass 2)

    def add(self, d) -> bool:
        """Buffer a described cell; True when it is a layout cell (counted against the cap)."""
        self.n_cells += 1
        col, size, _wfac, wrap, rot, text, cell = d[:7]
        one_line = (not rot and cell is None
                    and (text is None or (not wrap and (LINE_BREAKS_NEED_WRAP
                                                        or ("\n" not in text and "\r" not in text)))))
        if one_line:
            a = self.small.get(size)
            if a is None:
                a = self.small[size] = array("H")
            a.append(col)
            return False
        self.cells.append(d)
        return True

    def find(self, col):
        """The described cell in column col, or None."""
        for d in self.cells:
            if d[0] == col:
                return d
        for size, a in self.small.items():
            if col in a:
                return _one_line(self.r, col, size)
        return None

    def free(self):
        self.cells = []
        self.small = {}


def _one_line(r, col, size):
    return (col, size, 1.0, False, 0, None, None, make_ref(r, col), 0)


class _Geo:
    """Row/column geometry of one sheet, kept while its rows can still be decided."""
    __slots__ = ("head", "dflt", "cov", "hidden", "ht", "default_h", "zero_default")

    def __init__(self, head, dflt, cov, hidden, ht, default_h, zero_default):
        self.head, self.dflt, self.cov = head, dflt, cov
        self.hidden, self.ht, self.default_h, self.zero_default = hidden, ht, default_h, zero_default

    def merge_at(self, r, c):
        for m in self.cov.get(r, ()):
            if m.c1 <= c <= m.c2:
                return m
        return None

    def covering(self, r):
        return self.cov.get(r, ())

    def row_height(self, r) -> float:
        if r in self.hidden:
            return 0.0
        h = self.ht.get(r)
        if h is not None:
            return h
        return 0.0 if self.zero_default else self.default_h


class C73(Check):
    number = 73
    key = "Formatting/Reasonable row heights"
    needs_rows = True
    needs_cells = True
    needs_values = True          # only wrapped/rotated formula cells of rows still open are read
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    def start(self, wb):
        super().start(wb)
        st = wb.styles
        self.st = st
        nf = self._normal_font(st)
        self.normal_face = nf.name if nf is not None else None
        self.normal_size = float(nf.sz) if nf is not None and nf.sz and nf.sz > 0 else DEFAULT_FONT_PT
        self.mdw = mdw_px(self.normal_face, self.normal_size)
        self.space_px = space_px(self.normal_face, self.normal_size)
        self.floor = LINE_HEIGHT_PER_PT * self.normal_size
        self._sheets = []            # per-sheet stats (tab order)
        self._pass2 = {}             # sheet -> state for second pass
        self._values_read = set()
        self._style_cache = {}

    @staticmethod
    def _normal_font(st):
        """Font of the Normal cell style (builtinId 0), else fonts[0]; None without fonts."""
        fid = None
        for cs in st.cell_styles:
            if cs.builtin_id == 0 and cs.xf_id is not None and 0 <= cs.xf_id < len(st.cell_style_xfs):
                fid = st.cell_style_xfs[cs.xf_id].font_id
                break
        if fid is None or not (0 <= fid < len(st.fonts)):
            fid = 0
        return st.fonts[fid] if st.fonts else None

    def _style(self, s):
        v = self._style_cache.get(s)
        if v is None:
            font = self.st.font(s)
            al = self.st.xf(s).alignment
            size = float(font.sz) if font.sz and font.sz > 0 else DEFAULT_FONT_PT
            wrap = bool(al.wrap_text) or al.horizontal in WRAP_ALIGNMENTS or al.vertical in WRAP_ALIGNMENTS
            ind = 0
            if al.indent and al.indent > 0 and al.horizontal in INDENT_ALIGNMENTS:
                ind = (al.indent * INDENT_SPACES_PER_LEVEL * self.space_px
                       * (2 if al.horizontal == "distributed" else 1))
            v = self._style_cache[s] = (size, face_width(font.name, font.b), wrap, int(al.text_rotation or 0),
                                        font.name, font.b, ind)
        return v

    # ------------------------------------------------------------------ per sheet
    def sheet_start(self, head):
        f = head.format
        dh = f.default_row_height
        self._ignored_default = None
        if dh is not None and CUSTOM_HEIGHT_ONLY and not f.custom_height:
            self._ignored_default, dh = dh, None      # not custom: Excel auto-fits (2026-10-03)
        self._zero_default = bool(f.zero_height) or (dh is not None and dh <= 0)
        self._default_h = dh if dh is not None else self.floor
        # a default row height that is itself excessive: one sheet-level mistake covers every row
        # without a height of its own (they are not candidates as well)
        self._default_fail = (not self._zero_default and self._default_h > MAX_ROW_HEIGHT_PT
                              and self._default_h > EXCESS_FACTOR * self.floor)
        self._ht = {}            # r -> stored height (rows with ht)
        self._hidden = set()     # hidden rows
        self._cands = []
        self._cur = None
        self._n_buf = 0
        self._max_ht = 0.0
        self._n_default_rows = 0
        self._n_autofit = 0

    def row(self, row):
        r = row.r
        ht = row.ht
        if ht is not None and ht > 0 and CUSTOM_HEIGHT_ONLY and not row.custom_height:
            self._n_autofit += 1                 # stored height without customHeight: Excel auto-fits it
            ht = None
        if ht is not None:
            self._ht[r] = ht
        if row.hidden or (ht is not None and ht <= 0) or (ht is None and self._zero_default):
            self._hidden.add(r)
            self._cur = None
            return
        h = ht if ht is not None else self._default_h
        if h > self._max_ht:
            self._max_ht = h
        if ht is None and self._default_fail:
            self._n_default_rows += 1        # covered by the sheet-level mistake
            self._cur = None
        elif h > MAX_ROW_HEIGHT_PT:
            self._cur = _Cand(r, h, bool(row.custom_height) and ht is not None, ht is not None)
            self._cands.append(self._cur)
        else:
            self._cur = None

    def cell(self, cell):
        rec = self._cur
        if rec is None or cell.row != rec.r:
            return
        d = self._describe(cell)
        if d is not None and rec.add(d):
            self._n_buf += 1
            if self._n_buf > MAX_BUFFERED_CELLS:
                raise GradingError(f"{self.key}: more than {MAX_BUFFERED_CELLS:,} wrapped or rotated text cells in "
                                   f"rows taller than {MAX_ROW_HEIGHT_PT:g} pt on sheet {cell.sheet!r}; too many "
                                   f"to grade")

    def _describe(self, cell):
        """(col, size, wfac, wrap, rot, text, cell_or_None, ref, indent_px) for a displayed cell,
        else None.  text None + cell given = a formula result whose value is needed to lay it out."""
        size, wfac, wrap, rot, face, bold, ind = self._style(cell.s)
        if cell.is_formula_result:
            if wrap or rot:
                return (cell.col, size, wfac, wrap, rot, None, cell, cell.ref, ind)
            return (cell.col, size, wfac, wrap, rot, None, None, cell.ref, ind)   # one line, value-free
        if not cell.has_value:
            return None
        v = cell.value                                                       # a constant: always readable
        runs = cell.rich_runs
        if runs:
            # rich text shows each run in its own font: the largest run size and the widest run
            # face; a run without <rPr> (or without a size / face) has the cell's; the cell font
            # itself counts only for text outside the runs
            rsize, rwfac, in_runs = 0.0, 0.0, 0
            for rr in runs:
                if not rr.text:
                    continue
                in_runs += len(rr.text)
                f = rr.font
                if f is None:
                    s_, w_ = size, wfac
                else:
                    s_ = float(f.sz) if "sz" in f.specified and f.sz and f.sz > 0 else size
                    w_ = face_width(f.name if "name" in f.specified else face, f.b if "b" in f.specified else bold)
                rsize, rwfac = max(rsize, s_), max(rwfac, w_)
            if not isinstance(v, str) or len(v) > in_runs:
                rsize, rwfac = max(rsize, size), max(rwfac, wfac)
            if rsize > 0:
                size, wfac = rsize, rwfac
        text = self._text_of(v, rot) if cell.t != "d" else (v if rot else None)
        if text == "":
            return None
        return (cell.col, size, wfac, wrap, rot, text, None, cell.ref, ind)

    @staticmethod
    def _text_of(v, rot):
        """Text to lay out: '' = nothing displayed; None = a number (one line, never wraps);
        a rotated number is laid out as its digits."""
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, str):
            return v                      # text, or an error value (ExcelError is a str)
        if rot:
            return f"{v:.10g}" if isinstance(v, (int, float)) else str(v)
        return None

    # ------------------------------------------------------------------ geometry
    def _default_col_width(self, head) -> float:
        f = head.format
        if f.default_col_width is not None:
            return f.default_col_width
        base = f.base_col_width if f.base_col_width is not None else DEFAULT_BASE_COL_WIDTH
        px = math.ceil((base * self.mdw + 5) / 8.0) * 8                # Excel: whole 8-px steps
        return px / self.mdw

    def _col_width(self, head, c, dflt) -> float:
        ci = head.col_info(c)
        if ci is not None and ci.hidden:
            return 0.0
        w = ci.width if ci is not None and ci.width is not None else dflt
        return max(w, 0.0)

    def _geometry(self, d, geo, merge, r):
        """(px, span): the cell's (or merge's) width in px, and for a merge spanning several rows
        span = (own, others): the height of row r and the summed heights of the merge's other
        (visible) rows; else None.  Value-independent."""
        if merge is not None:
            px = sum(col_px(self._col_width(geo.head, c, geo.dflt), self.mdw) for c in range(merge.c1, merge.c2 + 1))
        else:
            px = col_px(self._col_width(geo.head, d[0], geo.dflt), self.mdw)
        span = None
        if merge is not None and merge.r2 > merge.r1:
            own = geo.row_height(r)
            others = sum(geo.row_height(x) for x in range(merge.r1, merge.r2 + 1) if x != r)
            span = (own, others)
        return px, span

    @staticmethod
    def _need(d, text, px, span, merge):
        """Need (pt) a cell with this text gives the row, and its explanation.  span = (own,
        others) for a merge over several rows: whole-block accounting (MERGE_ROWS_WHOLE_BLOCK)
        credits the row with need x own / (own + others); per-row share with need - others."""
        _col, size, wfac, wrap, rot, _t, _cell, ref, ind = d
        need, lines = cell_need(text, size, wfac, wrap, rot, px - CELL_MARGIN_PX - ind)
        what = (f"{lines} wrapped line{'s' if lines != 1 else ''}" if wrap and not rot and text is not None
                else ("rotated text" if rot else "one line"))
        why = f"{ref}: {what} of {size:g} pt text"
        if span is not None:
            own, others = span
            block = own + others
            if MERGE_ROWS_WHOLE_BLOCK:
                share = need * own / block if block > 0 else need
                why += (f" in merged {merge.ref} ({need:.1f} pt for the {block:.1f} pt block of rows "
                        f"{merge.r1}:{merge.r2}; this row's share {share:.1f} pt)")
                need = share
            else:
                why += f" in merged {merge.ref} ({need:.1f} pt, of which the other rows give {others:.1f} pt)"
                need = max(need - others, 0.0)
        elif merge is not None:
            why += f" in merged {merge.ref}"
        return need, why

    def _contribution(self, d, geo, merge, r):
        """Value-free: ((need, why), None), or (None, deferred) when the need depends on a
        formula value (deferred = (d, merge, px, span), settled in finish)."""
        px, span = self._geometry(d, geo, merge, r)
        if px <= 0:
            return (0.0, f"{d[7]} is in a hidden column"), None
        if d[6] is not None and d[5] is None:
            return None, (d, merge, px, span)
        return self._need(d, d[5], px, span, merge), None

    # ------------------------------------------------------------------ decision (value-free)
    def _decide(self, rec, geo, cand_by_row, *, final):
        """Settle rec from value-independent cells.  final=False: first pass (off-row anchors may
        be pending).  Status 'values' = only formula values can settle it (done in finish)."""
        r = rec.r
        thr = rec.ht / EXCESS_FACTOR
        best, why = self.floor, f"one line of the {self.normal_size:g} pt Normal font"
        deferred = []

        def take(d, m):
            nonlocal best, why
            res, dfr = self._contribution(d, geo, m, r)
            if dfr is not None:
                deferred.append(dfr)
            elif res[0] > best:
                best, why = res

        for d in rec.cells:
            m = geo.merge_at(r, d[0])
            if m is not None and (m.r1, m.c1) != (r, d[0]):
                continue                                        # covered cell: not displayed
            take(d, m)
        for size in sorted(rec.small, reverse=True):            # one-line cells, largest font first
            line = LINE_HEIGHT_PER_PT * size
            if line <= best:
                break
            for col in rec.small[size]:
                m = geo.merge_at(r, col)
                if m is not None and (m.r1, m.c1) != (r, col):
                    continue
                res, _ = self._contribution(_one_line(r, col, size), geo, m, r)
                if res[0] > best:
                    best, why = res
                if res[0] >= line:
                    break                                       # nothing at this size needs more
        # merges that start in an earlier row and cover this row
        off = []
        for m in geo.covering(r):
            if m.r1 == r:
                continue
            src = cand_by_row.get(m.r1)
            if src is None:
                off.append(m)
                continue
            d = src.find(m.c1)
            if d is not None:
                take(d, m)
        rec.need, rec.why, rec.deferred = best, why, deferred
        if best >= thr:
            rec.status = "ok"
            rec.deferred = []
        elif off and not final:
            rec.pending = off                                   # anchors first, values last
            rec.status = None
        elif deferred:
            rec.status = "values"
        else:
            rec.status = "fail"

    def sheet_end(self, head, tail):
        name = head.name
        cands = self._cands
        dflt = self._default_col_width(head)
        merges = [m for m in tail.merges if not m.is_single_cell]
        cand_rows = [c.r for c in cands]
        cand_by_row = {c.r: c for c in cands}
        # merge lookups restricted to candidate rows
        cov = {}
        for m in merges:
            i = bisect_left(cand_rows, m.r1)
            j = bisect_right(cand_rows, m.r2)
            for k in range(i, j):
                cov.setdefault(cand_rows[k], []).append(m)
        geo = _Geo(head, dflt, cov, self._hidden, self._ht, self._default_h, self._zero_default)
        for rec in cands:
            self._decide(rec, geo, cand_by_row, final=False)
        pending = [c for c in cands if c.status is None]
        if pending:
            cells = {(m.r1, m.c1) for c in pending for m in c.pending}
            self.request_second_pass(name, cells)
            self._pass2[name] = {"geo": geo, "cands": cands, "cand_by_row": cand_by_row, "pending": pending,
                                 "anchors": {}}
        else:
            for c in cands:          # decided (or waiting for values only): free the buffered cells
                c.free()
        self._sheets.append({"sheet": name, "state": head.state, "max_row_height": round(self._max_ht, 2),
                             "default_row_height": round(self._default_h, 2), "default_fail": self._default_fail,
                             "default_rows_written": self._n_default_rows, "cands": cands,
                             "autofit_heights_ignored": self._n_autofit,
                             "non_custom_default_ignored": self._ignored_default})

    # ------------------------------------------------------------------ second pass (rare)
    def second_pass_start(self, head):
        self._p2 = self._pass2.get(head.name)

    def second_pass_cell(self, cell):
        p = self._p2
        if p is None:
            return
        d = self._describe(cell)
        if d is not None:
            p["anchors"][(cell.row, cell.col)] = d

    def second_pass_end(self, head, tail):
        p = self._p2
        if p is None:
            return
        anchors = p["anchors"]
        for rec in p["pending"]:
            # fold the off-row anchors in as pseudo-candidate rows and decide for good
            extra = {}
            for m in rec.pending:
                d = anchors.get((m.r1, m.c1))
                if d is not None:
                    extra.setdefault(m.r1, _Cand(m.r1, 0.0, False, False)).add(d)
            by_row = dict(p["cand_by_row"])
            by_row.update(extra)
            rec.pending = []
            self._decide(rec, p["geo"], by_row, final=True)
        for c in p["cands"]:
            c.free()
        p["geo"] = p["anchors"] = p["cand_by_row"] = None
        self._p2 = None

    # ------------------------------------------------------------------ values (finish only)
    def _settle_with_values(self, rec, *, trusted_only):
        """Read the deferred formula values of an open row until it is settled.  trusted_only:
        skip untrusted values (the row stays 'values' when one of them would still be needed)."""
        thr = rec.ht / EXCESS_FACTOR
        best, why = rec.need, rec.why
        left = []
        for dfr in rec.deferred:
            d, m, px, span = dfr
            cell = d[6]
            if not cell.value_trusted:
                if trusted_only:
                    left.append(dfr)
                    continue
                # Patrick 2026-10-05 (every attempt graded): an untrusted value is skipped (it needs nothing)
                self.note_default("untrusted_value", f"{location(cell.sheet, d[7])} (source={cell.value_source})")
                continue
            v = self.require_value(cell)
            self._values_read.add((cell.sheet, d[7]))
            text = self._text_of(v, d[4])
            if text == "":
                continue                              # shows nothing
            need, w = self._need(d, text, px, span, m)
            if need > best:
                best, why = need, w
            if best >= thr:
                break
        rec.need, rec.why = best, why
        if best >= thr:
            rec.status, rec.deferred = "ok", []
        elif left:
            rec.status, rec.deferred = "values", left
        else:
            rec.status, rec.deferred = "fail", []

    def _settle_open_rows(self):
        """Settle rows only formula values can decide: trusted values first (always read, so the
        mistake list is complete); untrusted ones only while the verdict still needs them."""
        open_rows = [c for s in self._sheets for c in s["cands"] if c.status == "values"]
        if not open_rows:
            return
        for c in open_rows:
            self._settle_with_values(c, trusted_only=True)
        failed = any(s["default_fail"] or any(c.status == "fail" for c in s["cands"]) for s in self._sheets)
        if failed and UNTRUSTED_ONLY_IF_VERDICT_NEEDS:
            return                                   # the rest stay undecided (stats); nothing raised
        for c in open_rows:
            if c.status == "values":
                self._settle_with_values(c, trusted_only=False)     # untrusted values skipped (recorded)

    # ------------------------------------------------------------------ verdict
    def finish(self) -> dict:
        self._settle_open_rows()
        per_sheet = []
        undecided = []
        n_over = n_fail = 0
        for s in self._sheets:
            name = s["sheet"]
            hid = " (hidden sheet)" if s["state"] != "visible" else ""
            cands = s["cands"]
            if any(c.status is None for c in cands):
                raise GradingError(f"{self.key}: row decision left open on sheet {name!r}")
            fails = [c for c in cands if c.status == "fail"]
            und = [c for c in cands if c.status == "values"]
            n_over += len(cands)
            n_fail += len(fails)
            for c in und:
                undecided.append({"sheet": name, "row": c.r, "ht": c.ht, "need_without_values": round(c.need, 1),
                                  "cells_needing_values": sorted({dfr[0][7] for dfr in c.deferred})[:10]})
            # one mistake per run of consecutive failing rows
            runs = []
            for c in sorted(fails, key=lambda x: x.r):
                if runs and c.r == runs[-1][-1].r + 1:
                    runs[-1].append(c)
                else:
                    runs.append([c])
            for run in runs:
                a, b = run[0].r, run[-1].r
                hmin, hmax = min(c.ht for c in run), max(c.ht for c in run)
                top = max(run, key=lambda c: c.need)
                ratio = min(c.ht / c.need for c in run)
                hs = f"{hmax:g} pt" if hmin == hmax else f"{hmin:g}-{hmax:g} pt"
                if a == b:
                    lead = f"Row {a} on sheet '{name}'{hid} is {hs} tall but its content needs about {top.need:.1f} pt"
                else:
                    lead = (f"Rows {a}:{b} ({len(run)} rows) on sheet '{name}'{hid} are {hs} tall but their content "
                            f"needs at most {top.need:.1f} pt")
                if all(c.custom for c in run):
                    kind = "manually set height"
                elif all(c.own for c in run):
                    kind = "stored height"
                else:
                    kind = "stored or sheet-default height"
                self.add_mistake(location(name, f"{a}:{b}"),
                                 f"{lead} ({top.why}); {ratio:.2f}x the need ({kind}), over "
                                 f"{MAX_ROW_HEIGHT_PT:g} pt and more than {EXCESS_FACTOR:g}x what the content needs.")
            if s["default_fail"]:
                self.add_mistake(location(name),
                                 f"Sheet '{name}'{hid} has a default row height of {s['default_row_height']:g} pt "
                                 f"(sheetFormatPr): every row without a height of its own, empty rows included, "
                                 f"is that tall; one line needs about {self.floor:.1f} pt.")
            ok = [c for c in cands if c.status == "ok"]
            if cands or s["default_fail"]:
                per_sheet.append({
                    "sheet": name, "max_row_height": s["max_row_height"],
                    "rows_over_cap": len(cands), "rows_failing": len(fails), "rows_undecided": len(und),
                    "default_row_height": s["default_row_height"], "default_fail": s["default_fail"],
                    "default_height_rows_written": s["default_rows_written"],
                    "tall_rows_justified": [{"row": c.r, "ht": c.ht, "need": round(c.need, 1),
                                             "ratio": round(c.ht / c.need, 2), "why": c.why}
                                            for c in ok[:MAX_LISTED_OK_ROWS]],
                    "failing_rows": [{"row": c.r, "ht": c.ht, "need": round(c.need, 1),
                                      "ratio": round(c.ht / c.need, 2)} for c in fails[:MAX_LISTED_OK_ROWS]],
                })
        stats = {"n_sheets_checked": len(self._sheets), "rows_over_cap": n_over, "rows_failing": n_fail,
                 "non_custom_row_heights_ignored": sum(x["autofit_heights_ignored"] for x in self._sheets),
                 "non_custom_default_heights_ignored": {x["sheet"]: x["non_custom_default_ignored"] for x in self._sheets
                                                        if x["non_custom_default_ignored"] is not None},
                 "rows_undecided": len(undecided), "undecided_rows": undecided[:MAX_LISTED_OK_ROWS],
                 "normal_font_pt": self.normal_size, "one_line_floor_pt": round(self.floor, 2),
                 "formula_values_read": len(self._values_read),
                 "second_pass_sheets": sorted(self._pass2), "per_sheet": per_sheet,
                 "options": {"line_height_per_pt": LINE_HEIGHT_PER_PT,
                             "merge_rows_whole_block": MERGE_ROWS_WHOLE_BLOCK,
                             "untrusted_only_if_verdict_needs": UNTRUSTED_ONLY_IF_VERDICT_NEEDS,
                             "line_breaks_need_wrap": LINE_BREAKS_NEED_WRAP,
                             "custom_height_only": CUSTOM_HEIGHT_ONLY}}
        flagged = ", ".join(f"'{p['sheet']}'" for p in per_sheet if p["rows_failing"] or p["default_fail"])
        und_note = ""
        if undecided:
            und_note = (f" {len(undecided)} other row(s) over {MAX_ROW_HEIGHT_PT:g} pt were left undecided: only "
                        f"untrusted formula values could settle them, and the verdict does not depend on them "
                        f"(see stats).")
        return self.verdict(
            f"No excessive row heights: {n_over} row(s) over {MAX_ROW_HEIGHT_PT:g} pt on "
            f"{len(self._sheets)} sheet(s), each within {EXCESS_FACTOR:g}x its content's need.",
            f"{{n}} block(s) of excessively tall rows (over {MAX_ROW_HEIGHT_PT:g} pt and more than "
            f"{EXCESS_FACTOR:g}x their content's need) on: {flagged}.{und_note}", stats)
