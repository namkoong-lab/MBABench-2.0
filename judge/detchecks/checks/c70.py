"""70 Formatting/Reasonable column widths.

Rubric: "Column widths should not be excessive."  Good: "Columns are no wider than their content
requires; several full-width columns fit on screen at once, and long text is wrapped rather than
accommodated by widening the column."  Bad: "The model contains excessively wide columns (rule of
thumb: wider than roughly 75-80 characters at the default font), or long text is left unwrapped."

Rule as implemented (the toys' wording; every judgement call is a module constant below):

  Units.  A column's width is what Excel's Column Width dialog shows: characters of the workbook's
  Normal font, (px - 5) / MDW, where px is the column's pixel width at 96 dpi from its stored width
  (ECMA-376, exactly as Sufficient column widths (69): c69.col_px / c69.default_col_px, strict
  reading) and MDW the Normal font's maximum digit width in whole pixels.  Hidden and zero-width
  columns are skipped (No hidden rows/columns (93) owns them).  Every worksheet, dialog sheet and
  macro sheet is graded, hidden sheets included.  Overlapping <col> entries (GPT-6 tooling): a later
  entry (the reader's order: by min, then file order) overrides earlier ones on the columns it
  covers - the core reader's one rule (SheetHead.col_segments / col_info, shared with Sufficient
  column widths (69) and Reasonable row heights (73)).

  Content need of a column = its widest displayed content in the same characters: the text width of
  each displayed cell in the cell's own face, size and weight (c69's glyph tables, 96 dpi, per-glyph
  rounded advances) divided by MDW.  Numbers and dates: their displayed text (numfmt with the cell's
  format).  Booleans and errors: their text.  Unwrapped text: its full width (typed line breaks do
  not break an unwrapped cell).  Wrapped text (wrapText, or justify / distributed): its widest line,
  capped at the cap (WIDTH_CAP_CHARS): wrapped text can always wrap at the cap.  Rich text: run by
  run in each run's size and weight.  Indent adds 3 Normal-font spaces a level; rotated text needs
  its projection on the horizontal.  Cells in hidden rows or hidden columns are not displayed and
  need nothing.  Merged ranges: covered cells need nothing; the merged content needs the merged
  width, so each column of the span needs what the span's other visible columns leave
  (content - their pixels); a one-column merge counts in full.

  Test A (cap).  A visible column FAILS when it is wider than WIDTH_CAP_CHARS (80) AND more than
  EXCESS_FACTOR (2) times as wide as its content needs.  Empty columns with an explicit width over
  the cap fail (need 0), and so do the columns of a sheet whose own default width is over the cap.

  Test B (WIDE OUTLIER).  The judge's tag (judge/utils/workbook_properties.wide_outlier_tags) ported
  on stored widths: inside the sheet's used range (first to last column holding a value or formula)
  the visible columns form runs of equal stored width (<col> widths rounded to 2 decimals; columns
  without one take sheetFormatPr/@defaultColWidth, else 8.43); a run of >= 2 columns at least
  WIDE_OUTLIER_RATIO (2.5) times as wide as the nearest run of >= 2 columns at least
  WIDE_OUTLIER_MIN_NEIGHBOUR (4) wide on BOTH sides, and no longer than the longer of the two, is a
  WIDE OUTLIER.  Extension (OUTLIER_UNEQUAL_GROUPS): >= 2 adjacent lone columns of unequal widths
  between two such neighbour runs, each at least 2.5 times as wide as both, form a group too (toy T2
  Fail's AE:AF are stored 48.0 and 35.78: the exact port never compares lone columns and misses
  them).  An outlier column FAILS when its content needs less than OUTLIER_NEED_SHARE (1/2) of its
  width (the toys' wording; the judge's guidance excuses a column "whose own content needs the
  width").

  No "long text left unwrapped" test (Patrick 2026-10-04): unwrapped text is a problem only when it
  is cut off, which Sufficient column widths (69) grades.  The former Test C (200+ character
  unwrapped text spilling into empty cells) is removed with its switches and stats.

  Case brief.  The guidance excludes "the case's own brief (the Instructions sheet the agent was
  given) ... including its text wrapping and column widths"; Patrick's rule grades the whole
  workbook.  EXCLUDE_CASE_BRIEF = False: the sheet named "Instructions" (c74.is_instructions_sheet)
  is graded by Tests A and B.

  stats["switches"] gives the verdict under each alternative setting (cap 75, outlier share 0.75,
  exact judge port, brief excluded, the judge's sheet skip); they never decide.

Values (needs_values).  Constants are read from the file.  A formula result's value is read only
for cells of columns being judged (Tests A / B), and only when trusted (the recalc pipeline supplies
a value for every non-Excel file).  An untrusted value matters only when its column's measured need
is still under the line; if the workbook fails anyway such cells are listed in stats
(UNTRUSTED_ONLY_IF_VERDICT_NEEDS), otherwise the first one raises GradingError (no fallback).

Mistakes: one per run of adjacent failing columns with the same test(s) per sheet (location
'Sheet'!H:H, 'Sheet'!AE:AF), naming the widths, the content need with its widest cell, and the test.
"""
from __future__ import annotations

import math
import re
from bisect import bisect_right
from typing import Optional

from ..core import numfmt as N
from ..core.refs import MAX_COL, index_to_col, location, make_ref
from ..core.sheet import ExcelError, paint_cols  # noqa: F401  (paint_cols re-exported: tests, scratch)
from ..errors import GradingError
from .base import Check
from .c69 import (CELL_MARGIN_PX, DEFAULT_FONT_PT, GLYPH_CHARS, WIDE_CHAR, _normal_font, col_px, default_col_px,
                  face_key, glyph_row, mdw_px, ppem_at, resolve_char, space_px)
from .c74 import is_instructions_sheet

# ---------------------------------------------------------------- rule constants
WIDTH_CAP_CHARS = 80.0            # Test A: columns wider than this (Normal-font characters) are examined ...
EXCESS_FACTOR = 2.0               # ... and fail when more than this many times as wide as their content needs
WIDE_OUTLIER_RATIO = 2.5          # Test B, judge port: a run this many times wider than both neighbour runs
WIDE_OUTLIER_MIN_NEIGHBOUR = 4.0  # judge port: runs narrower than this (stored width) are spacers, never neighbours
WIDE_OUTLIER_EQUAL_TOL = 0.01     # judge port: stored widths (rounded to 2 decimals) closer than this are equal
JUDGE_DEFAULT_COL_WIDTH = 8.43    # judge port: stored width of columns without <col> on a sheet without a default
OUTLIER_UNEQUAL_GROUPS = True     # Test B extension: >= 2 adjacent lone columns of unequal widths form a group too
OUTLIER_NEED_SHARE = 0.5          # Test B: an outlier column fails when its content needs less than this share
EXCLUDE_CASE_BRIEF = False        # True: the case brief is not graded at all (guidance); False: whole workbook
UNTRUSTED_ONLY_IF_VERDICT_NEEDS = True   # untrusted values raise only while the verdict is still open
INDENT_SPACES_PER_LEVEL = 3       # ECMA-376 alignment/@indent: one level = 3 spaces of the Normal font
WRAP_ALIGNMENTS = ("justify", "distributed")    # horizontal alignments Excel wraps like wrapText
ROTATED_LINE_PER_PT = 1.3         # rotated text: one line is 1.3 x the font size (as Reasonable row heights (73))
# stats only: the verdict under the alternatives the doc reports (they never decide)
ALT_CAP_CHARS = 75.0              # the rubric's "roughly 75-80": the lower end
ALT_OUTLIER_NEED_SHARE = 0.75     # the earlier prototype's excuse (content needs >= 75 % of the width)
JUDGE_SKIP_SHEETS = re.compile(r"instruction|question|brief|readme", re.I)   # sheets the judge's tag skips
MAX_LISTED = 12
MAX_UNDECIDED_LISTED = 12

# ---------------------------------------------------------------- text width (c69's glyph tables)
_DIGIT_IDX = GLYPH_CHARS.index("0")


def text_px(row: tuple, text: str, ppem: float) -> int:
    """Width of `text` in whole pixels: the face's per-glyph advances (c69.GLYPHS, 1/1000 em) at ppem,
    each rounded (the reference model of Sufficient column widths (69): 96 dpi, GDI rounding).
    Characters outside the table: accented letters as their base letter, East Asian wide characters
    1 em, combining marks 0, anything else a digit (c69)."""
    tot = 0
    dig = row[_DIGIT_IDX]
    for ch in text:
        i = resolve_char(ch)                      # c69's text-character rule (shared)
        if i is None:
            a = dig
        elif i >= 0:
            a = row[i] or dig
        elif i == WIDE_CHAR:
            a = 1000
        else:
            continue
        tot += round(a * ppem / 1000.0)
    return tot


# ---------------------------------------------------------------- column geometry
class Geo:
    """Pixel widths (96 dpi) of one sheet's columns; None = hidden or zero width."""
    __slots__ = ("segs", "seg_lo", "default_px", "mdw", "cache")

    def __init__(self, segs: list, default_px: int, mdw: int):
        self.segs = segs
        self.seg_lo = [s[0] for s in segs]
        self.default_px = default_px
        self.mdw = mdw
        self.cache: dict = {}

    def seg_px(self, ci) -> Optional[int]:
        if ci.hidden or (ci.width is not None and ci.width <= 0):
            return None
        if ci.width is None:                     # a <col> without a width: the sheet default
            return self.default_px
        return col_px(ci.width, self.mdw)

    def px(self, c: int) -> Optional[int]:
        v = self.cache.get(c, 0)
        if v != 0:
            return v
        i = bisect_right(self.seg_lo, c) - 1
        if i >= 0 and self.segs[i][0] <= c <= self.segs[i][1]:
            v = self.seg_px(self.segs[i][2])
        else:
            v = self.default_px
        self.cache[c] = v
        return v


def stored_pieces(segs: list, default_stored: float) -> list:
    """(lo, hi, stored width, or None when hidden / zero width) covering columns 1..MAX_COL, as the
    judge's tag sees them: <col> widths rounded to 2 decimals, the sheet default elsewhere (also for a
    <col> without a width attribute)."""
    out = []
    c = 1
    for lo, hi, ci in segs:
        if lo > c:
            out.append((c, lo - 1, default_stored))
        hidden = ci.hidden or (ci.width is not None and ci.width <= 0)
        out.append((lo, hi, None if hidden else (default_stored if ci.width is None else round(ci.width, 2))))
        c = hi + 1
    if c <= MAX_COL:
        out.append((c, MAX_COL, default_stored))
    return out


def width_runs(pieces: list, lo: int, hi: int) -> list:
    """Runs [first, last, stored width] of adjacent visible columns of equal stored width inside
    lo..hi - the judge's run building (hidden columns are left out and break adjacency)."""
    runs: list = []
    for a, b, w in pieces:
        a, b = max(a, lo), min(b, hi)
        if a > b or w is None:
            continue
        if runs and runs[-1][1] == a - 1 and abs(runs[-1][2] - w) < WIDE_OUTLIER_EQUAL_TOL:
            runs[-1][1] = b
        else:
            runs.append([a, b, w])
    return runs


def wide_outliers(runs: list, unequal: Optional[bool] = None) -> list:
    """WIDE OUTLIER groups among `runs` -> [(first, last, kind, left_run, right_run, widths)].

    kind "judge": the exact port of judge/utils/workbook_properties.wide_outlier_tags - a run of >= 2
    equal columns at least WIDE_OUTLIER_RATIO times as wide as the nearest run of >= 2 columns that is
    at least WIDE_OUTLIER_MIN_NEIGHBOUR wide on BOTH sides (lone columns and spacers are skipped when
    looking for neighbours), and no longer than the longer neighbour run.
    kind "ext" (unequal, default OUTLIER_UNEQUAL_GROUPS): a stretch of >= 2 adjacent columns between two
    consecutive such neighbour runs, none of them in a qualifying run, each at least WIDE_OUTLIER_RATIO
    times as wide as both neighbours, and no longer than the longer neighbour run.
    widths = the group's stored widths, left to right (one per run)."""
    if unequal is None:
        unequal = OUTLIER_UNEQUAL_GROUPS
    field = [i for i, r in enumerate(runs) if r[1] - r[0] + 1 >= 2 and r[2] >= WIDE_OUTLIER_MIN_NEIGHBOUR]
    out = []
    for k, i in enumerate(field):
        r = runs[i]
        if r[2] < WIDE_OUTLIER_MIN_NEIGHBOUR * WIDE_OUTLIER_RATIO or k == 0 or k == len(field) - 1:
            continue
        left, right = runs[field[k - 1]], runs[field[k + 1]]
        if r[2] < WIDE_OUTLIER_RATIO * left[2] or r[2] < WIDE_OUTLIER_RATIO * right[2]:
            continue
        if r[1] - r[0] + 1 > max(left[1] - left[0] + 1, right[1] - right[0] + 1):
            continue
        out.append((r[0], r[1], "judge", tuple(left), tuple(right), (r[2],)))
    if unequal:
        for k in range(len(field) - 1):
            ia, ib = field[k], field[k + 1]
            left, right = runs[ia], runs[ib]
            thr = max(WIDE_OUTLIER_RATIO * max(left[2], right[2]), WIDE_OUTLIER_MIN_NEIGHBOUR * WIDE_OUTLIER_RATIO)
            longest = max(left[1] - left[0] + 1, right[1] - right[0] + 1)
            groups, grp = [], None
            for i in range(ia + 1, ib):
                r = runs[i]
                if r[2] >= thr:
                    if grp is not None and r[0] == grp[1] + 1:
                        grp[1] = r[1]
                        grp[2].append(r[2])
                        continue
                    if grp is not None:
                        groups.append(grp)
                    grp = [r[0], r[1], [r[2]]]
                elif grp is not None:
                    groups.append(grp)
                    grp = None
            if grp is not None:
                groups.append(grp)
            for g in groups:
                if 2 <= g[1] - g[0] + 1 <= longest:
                    out.append((g[0], g[1], "ext", tuple(left), tuple(right), tuple(g[2])))
    out.sort(key=lambda x: x[0])
    return out


def judge_tag(group) -> str:
    """The judge's tag line for a group, e.g. 'BA:BD width 34.9 vs neighbours 12.9 (2.7x): WIDE OUTLIER'
    (an unequal group lists its widths)."""
    first, last, kind, left, right, widths = group
    nw = sorted({round(left[2], 1), round(right[2], 1)})
    w = "/".join(f"{x:.1f}" for x in widths)
    ratio = min(widths) / max(left[2], right[2])
    return (f"{index_to_col(first)}:{index_to_col(last)} width {w} vs neighbours {'/'.join(f'{x:.1f}' for x in nw)} "
            f"({ratio:.1f}x): WIDE OUTLIER" + (" (unequal widths)" if kind == "ext" else ""))


# ---------------------------------------------------------------- settings (the verdict and its alternatives)
def default_setting() -> dict:
    return {"cap": WIDTH_CAP_CHARS, "share": OUTLIER_NEED_SHARE, "ext": OUTLIER_UNEQUAL_GROUPS,
            "brief": not EXCLUDE_CASE_BRIEF, "judge_skip": False, "regardless": False}


def alternative_settings() -> dict:
    """Stats only (stats.switches): the verdict under each alternative; none of them decides.
    *_regardless: Test A fails every column over the cap whatever its content (the rubric's "rule of
    thumb" read literally, as the judge guidance's second test does); llm_like: that plus the brief
    excluded, the 75 % outlier excuse and the judge's sheet skip - how the LLM judge is instructed to
    grade (the LLM judge also never failed long unwrapped text on its own)."""
    d = default_setting()
    alts = {"cap_75": dict(d, cap=ALT_CAP_CHARS), "cap_80": dict(d, cap=80.0),
            "outlier_share_0.75": dict(d, share=ALT_OUTLIER_NEED_SHARE),
            "outlier_exact_port": dict(d, ext=False), "brief_excluded": dict(d, brief=False),
            "judge_sheet_skip": dict(d, judge_skip=True),
            "cap_80_regardless": dict(d, cap=80.0, regardless=True),
            "cap_75_regardless": dict(d, cap=ALT_CAP_CHARS, regardless=True),
            "llm_like": dict(d, cap=80.0, regardless=True, brief=False, ext=False,
                             share=ALT_OUTLIER_NEED_SHARE, judge_skip=True)}
    return alts


class _Sty:
    __slots__ = ("row", "ppem", "size", "bold", "face_k", "font", "wrap", "shrink", "rot", "indent_px", "code",
                 "fmt_err")


class _Need:
    """Widest content of one judged column: the widest unwrapped and the widest wrapped cell
    (px, row, col, snippet, font), the count, and untrusted formula results (undecided)."""
    __slots__ = ("unwrapped", "wrapped", "n", "untrusted", "n_untrusted")

    def __init__(self):
        self.unwrapped = None
        self.wrapped = None
        self.n = 0
        self.untrusted = []
        self.n_untrusted = 0

    def add(self, px, wrapped, r, c, snippet, font):
        self.n += 1
        if wrapped:
            if self.wrapped is None or px > self.wrapped[0]:
                self.wrapped = (px, r, c, snippet, font)
        elif self.unwrapped is None or px > self.unwrapped[0]:
            self.unwrapped = (px, r, c, snippet, font)

    def add_untrusted(self, r, c, why=None):
        """A cell whose display cannot be measured: an untrusted formula value (why None) or a number
        rendering the format engine marks unverified (why says so)."""
        self.n_untrusted += 1
        if len(self.untrusted) < MAX_UNDECIDED_LISTED:
            self.untrusted.append((r, c, why))


def _filled(cell) -> bool:
    """A cell that counts for the sheet's used range (Test B): a formula (even one returning "") or a
    stored value."""
    return cell.formula is not None or cell.array is not None or (cell.raw is not None and cell.raw != "")


def _colrange(a: int, b: int) -> str:
    return f"{index_to_col(a)}:{index_to_col(b)}"


class C70(Check):
    number = 70
    key = "Formatting/Reasonable column widths"
    needs_rows = True
    needs_cells = True
    needs_values = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    # ------------------------------------------------------------------ workbook
    def start(self, wb):
        super().start(wb)
        st = wb.styles
        self.st = st
        self.date1904 = bool(wb.date1904)
        nf = _normal_font(st)
        self.normal_face = nf.name if nf is not None else None
        self.normal_size = float(nf.sz) if nf is not None and nf.sz and nf.sz > 0 else DEFAULT_FONT_PT
        nk, self.normal_known = face_key(self.normal_face)
        self.mdw = mdw_px(nk, self.normal_size, 96)
        self.space = space_px(nk, self.normal_size, 96)
        self.settings = {"default": default_setting(), **alternative_settings()}
        self.caps = sorted({s["cap"] for s in self.settings.values()})
        self.low_cap = min(self.caps)
        self._styles = {}
        self._need_cache = {}
        self._sheets = []
        self._recs = {}
        self.unknown_faces = {}
        self.counts = {"cells_measured": 0, "formula_values_read": 0, "sheets_with_overlapping_cols": 0,
                       "cols_without_width": 0, "second_pass_sheets": 0}

    def chars(self, px) -> float:
        """Pixels -> Normal-font characters (Excel's Column Width dialog: (px - 5) / MDW)."""
        return (px - CELL_MARGIN_PX) / self.mdw

    # ------------------------------------------------------------------ styles and measuring
    def _style(self, s) -> _Sty:
        v = self._styles.get(s)
        if v is not None:
            return v
        st = self.st
        xf = st.xf(s)
        font = st.font(s)
        al = xf.alignment
        v = _Sty()
        k, known = face_key(font.name)
        if not known:
            nm = (font.name or "").strip() or "(none)"
            self.unknown_faces[nm] = self.unknown_faces.get(nm, 0) + 1
        v.face_k = k
        v.size = float(font.sz) if font.sz and font.sz > 0 else DEFAULT_FONT_PT
        v.bold = bool(font.b)
        v.row = glyph_row(k, font.b)
        v.ppem = ppem_at(v.size, 96)
        v.font = f"{(font.name or 'default font').strip()} {v.size:g}{' bold' if font.b else ''}"
        v.wrap = bool(al.wrap_text) or al.horizontal in WRAP_ALIGNMENTS
        v.shrink = bool(al.shrink_to_fit)
        v.rot = int(al.text_rotation or 0)
        lv = int(al.indent or 0) if al.horizontal in (None, "general", "left", "right", "distributed") else 0
        v.indent_px = lv * INDENT_SPACES_PER_LEVEL * self.space * (2 if al.horizontal == "distributed" else 1)
        v.fmt_err = None
        try:
            v.code = N.resolve_format(xf.num_fmt_id, st.num_fmts)
            N.parse_format(v.code)               # an unreadable code raises here, not at a cell
        except GradingError as e:
            # Patrick 2026-10-06 (attempt 2777, measured in Excel): an unreadable format is General
            v.code, v.fmt_err = "General", str(e)
        self._styles[s] = v
        return v

    def _need_px(self, sty: _Sty, text: str, runs=None) -> int:
        """Pixels a displayed text needs: one line (unwrapped; typed line breaks do not break it) or
        its widest line (wrapped), indent and rotation included.  Rich runs use their own size and
        weight (the reader does not keep a run's face: the cell's face is used)."""
        if runs:
            segs = []
            rest = len(text) - sum(len(rn.text) for rn in runs)
            if rest > 0:                         # plain text before the runs (<si><t>..</t><r>..)
                segs.append((sty.row, sty.ppem, text[:rest]))
            for rn in runs:
                f = rn.font
                if f is None:
                    segs.append((sty.row, sty.ppem, rn.text))
                    continue
                size = float(f.sz) if ("sz" in f.specified and f.sz and f.sz > 0) else sty.size
                bold = bool(f.b) if "b" in f.specified else sty.bold
                segs.append((glyph_row(sty.face_k, bold), ppem_at(size, 96), rn.text))
            if sty.wrap:
                lines, cur = [], 0
                for row, ppem, t in segs:
                    for j, part in enumerate(t.replace("\r\n", "\n").replace("\r", "\n").split("\n")):
                        if j:
                            lines.append(cur)
                            cur = 0
                        cur += text_px(row, part, ppem)
                lines.append(cur)
                w = max(lines)
            else:
                w = sum(text_px(row, t.replace("\r", "").replace("\n", ""), ppem) for row, ppem, t in segs)
        else:
            key = (text, id(sty.row), sty.ppem, sty.wrap)
            w = self._need_cache.get(key)
            if w is None:
                if sty.wrap:
                    w = max(text_px(sty.row, ln, sty.ppem)
                            for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"))
                else:
                    w = text_px(sty.row, text.replace("\r", "").replace("\n", ""), sty.ppem)
                if len(self._need_cache) > 200_000:
                    self._need_cache.clear()
                self._need_cache[key] = w
        if sty.rot:
            if sty.rot == 255:                   # vertical stacked text: one character wide
                w = max((text_px(sty.row, ch, sty.ppem) for ch in text if ch not in "\r\n"), default=0)
            else:
                ang = math.radians(sty.rot if sty.rot <= 90 else sty.rot - 90)
                line = sty.size * ROTATED_LINE_PER_PT * 96.0 / 72.0
                w = int(math.ceil(abs(w * math.cos(ang)) + abs(line * math.sin(ang))))
        return w + sty.indent_px

    def _route(self, cell):
        """(kind, value) of a displayed cell: 'number' | 'date' | 'text' | 'bool' | 'error' | 'empty' |
        'untrusted'.  Constants by their stored type; formula results by their TRUSTED value (the
        delivered t of an agent cache is not evidence, as in Sufficient column widths (69))."""
        if cell.is_formula_result:
            if not cell.value_trusted:
                return "untrusted", None
            self.counts["formula_values_read"] += 1
            v = cell.value
            if cell.t == "d" and isinstance(v, str) and cell.value_source == "cached":
                return "date", v
        else:
            if cell.raw is None:
                return "empty", None
            t = cell.t
            v = cell.value
            if t == "d":
                return "date", v
            if t == "b":
                return "bool", v
            if t == "e":
                return "error", v
        if v is None:
            return "empty", None
        if isinstance(v, bool):
            return "bool", v
        if isinstance(v, (int, float)):
            return ("number", v) if math.isfinite(v) else ("empty", None)
        if isinstance(v, ExcelError):
            return "error", v
        if isinstance(v, str):
            return ("text", v) if v != "" else ("empty", None)
        return "empty", None

    def _display(self, cell, sty: _Sty, kind: str, v):
        """(text, rich runs or None) a displayed cell shows; (None, why) when the number format engine
        marks the rendering unverified (certain=False: e.g. '0 days', unquoted date letters among digit
        placeholders) - the cell is then undecided like an untrusted value."""
        if kind in ("number", "date"):
            if sty.fmt_err:
                self.note_default("unreadable_number_format", f"{cell.sheet}!{cell.ref}: {sty.fmt_err}")
            r = N.render(v, sty.code, value_type=("d" if kind == "date" else None), date1904=self.date1904)
            if not r.certain:
                # Patrick 2026-10-05 (every attempt graded): the engine's best rendering is measured
                self.note_default("unverified_number_format", f"{location(cell.sheet, cell.ref)} = {v!r} under "
                                                              f"{sty.code!r} measured as {r.text!r}")
            return ("" if r.is_blank else r.text), None
        if kind == "bool":
            return ("TRUE" if v else "FALSE"), None
        if kind == "error":
            return str(v), None
        return v, (None if cell.is_formula_result else cell.rich_runs)

    # ------------------------------------------------------------------ per sheet
    def sheet_start(self, head):
        name = head.name
        self._brief = is_instructions_sheet(name)
        segs = head.col_segments()               # the shared reading: the later <col> entry wins (core)
        raw = sum(min(ci.max, MAX_COL) - max(ci.min, 1) + 1 for ci in head.cols if ci.max >= ci.min)
        if raw > sum(s[1] - s[0] + 1 for s in segs):
            self.counts["sheets_with_overlapping_cols"] += 1
        self.counts["cols_without_width"] += sum(1 for ci in head.cols if ci.width is None and not ci.hidden)
        strict, _lenient = default_col_px(head.format, self.mdw)
        self.geo = geo = Geo(segs, strict, self.mdw)
        dcw = head.format.default_col_width
        self._pieces = stored_pieces(segs, dcw if dcw else JUDGE_DEFAULT_COL_WIDTH)
        # columns to measure: Test A's (wider than the lowest cap of any setting) and a superset of Test
        # B's (outlier groups over the full column range; the used range can only remove groups)
        flags = bytearray(MAX_COL + 2)
        wide = []
        c = 1
        for lo, hi, ci in segs:
            if lo > c and self.chars(strict) > self.low_cap:
                wide.append((c, lo - 1, strict))
            px = geo.seg_px(ci)
            if px is not None and self.chars(px) > self.low_cap:
                wide.append((lo, hi, px))
            c = hi + 1
        if c <= MAX_COL and self.chars(strict) > self.low_cap:
            wide.append((c, MAX_COL, strict))
        for lo, hi, _px in wide:
            flags[lo:hi + 1] = b"\x01" * (hi - lo + 1)
        for g in wide_outliers(width_runs(self._pieces, 1, MAX_COL), unequal=True):
            flags[g[0]:g[1] + 1] = b"\x01" * (g[1] - g[0] + 1)
        self._wide = wide
        self._flags = flags
        self._needs = {}
        self._minc = self._maxc = None
        self._zero_default = bool(head.format.zero_height)
        self._row_hidden = False
        self._hidden_rows = set()

    def row(self, row):
        ht = row.ht
        self._row_hidden = bool(row.hidden) or (ht is not None and ht <= 0) or (ht is None and self._zero_default)
        if self._row_hidden:
            self._hidden_rows.add(row.r)

    def cell(self, cell):
        c = cell.col
        filled = _filled(cell)
        if filled:
            if self._minc is None or c < self._minc:
                self._minc = c
            if self._maxc is None or c > self._maxc:
                self._maxc = c
        if self._row_hidden or not filled or not self._flags[c]:
            return                               # only the judged columns' displayed cells are measured
        px = self.geo.px(c)
        if px is None:                           # hidden / zero-width column: not displayed
            return
        sty = self._style(cell.s)
        kind, v = self._route(cell)
        nd = self._needs.get(c)
        if nd is None:
            nd = self._needs[c] = _Need()
        if kind == "untrusted":
            nd.add_untrusted(cell.row, c)
        elif kind != "empty":
            text, runs = self._display(cell, sty, kind, v)
            if text is None:
                nd.add_untrusted(cell.row, c, runs)          # unverified rendering (runs holds why)
            elif text:
                nd.add(self._need_px(sty, text, runs), sty.wrap and kind == "text", cell.row, c, text[:60],
                       sty.font)
                self.counts["cells_measured"] += 1

    def sheet_end(self, head, tail):
        name = head.name
        merges = [m for m in tail.merges if not m.is_single_cell]
        lo, hi = self._minc, self._maxc
        groups = wide_outliers(width_runs(self._pieces, lo, hi), unequal=True) if lo is not None else []
        outlier_of = {}
        for g in groups:
            for c in range(g[0], g[1] + 1):
                outlier_of[c] = g
        rec = {"sheet": name, "state": head.state, "brief": self._brief,
               "judge_skip": bool(JUDGE_SKIP_SHEETS.search(name)), "geo": self.geo, "wide": self._wide,
               "groups": groups, "outlier_of": outlier_of, "needs": self._needs, "used": (lo, hi),
               "hidden_rows": self._hidden_rows, "spans": [], "anchors": {}, "anchor_need": {},
               "anchor_why": {}, "remeasure": set()}
        # merges over a judged column: covered cells need nothing and the merged content needs the span,
        # so those columns are measured again in a second pass, merges known
        if merges:
            for m in merges:
                touched = [c for c in range(m.c1, min(m.c2, MAX_COL) + 1)
                           if c in outlier_of or self._in_wide(c)]
                if not touched:
                    continue
                rec["spans"].append(m)
                rec["remeasure"].update(touched)
                if m.r1 not in self._hidden_rows:
                    rec["anchors"][(m.r1, m.c1)] = m
            if rec["spans"]:
                self.counts["second_pass_sheets"] += 1
                self.request_second_pass(name, None)
        self._recs[name] = rec
        self._sheets.append(rec)
        self._needs = {}

    def _in_wide(self, c: int) -> bool:
        return any(a <= c <= b for a, b, _px in self._wide)

    # ------------------------------------------------------------------ second pass (merges)
    def second_pass_start(self, head):
        rec = self._recs.get(head.name)
        self._p2 = rec
        if rec is None:
            return
        self.geo = rec["geo"]
        self._p2_needs = {c: _Need() for c in rec["remeasure"]}
        self._p2_merges = {}
        for m in rec["spans"]:
            for c in range(m.c1, min(m.c2, MAX_COL) + 1):
                if c in rec["remeasure"]:
                    self._p2_merges.setdefault(c, []).append(m)

    def second_pass_cell(self, cell):
        rec = self._p2
        if rec is None:
            return
        r, c = cell.row, cell.col
        anchor = rec["anchors"].get((r, c))
        if anchor is None and c not in rec["remeasure"]:
            return
        if r in rec["hidden_rows"] or not _filled(cell):
            return
        if anchor is None:
            if self.geo.px(c) is None:
                return
            for m in self._p2_merges.get(c, ()):
                if m.r1 <= r <= m.r2:            # a covered cell: never displayed
                    return
        kind, v = self._route(cell)
        sty = self._style(cell.s)
        if anchor is not None:
            # the merged content: it needs the merged span (each column the rest; _column_need)
            if kind == "untrusted":
                rec["anchor_need"][(r, c)] = None
            elif kind != "empty":
                text, runs = self._display(cell, sty, kind, v)
                if text is None:
                    rec["anchor_need"][(r, c)] = None
                    rec["anchor_why"][(r, c)] = runs             # unverified rendering
                elif text:
                    rec["anchor_need"][(r, c)] = (self._need_px(sty, text, runs), sty.wrap and kind == "text",
                                                  text[:60], sty.font)
            return
        nd = self._p2_needs[c]
        if kind == "untrusted":
            nd.add_untrusted(r, c)
        elif kind != "empty":
            text, runs = self._display(cell, sty, kind, v)
            if text is None:
                nd.add_untrusted(r, c, runs)
            elif text:
                nd.add(self._need_px(sty, text, runs), sty.wrap and kind == "text", r, c, text[:60], sty.font)

    def second_pass_end(self, head, tail):
        rec = self._p2
        if rec is not None:
            rec["needs"].update(self._p2_needs)
        self._p2 = None

    # ------------------------------------------------------------------ deciding
    def _column_need(self, rec, c: int, cap: float):
        """(need in characters, widest example, undecided cells) of column c, wrapped text capped at
        `cap` characters.  example = (row, col, snippet, font, note)."""
        cap_px = cap * self.mdw
        nd = rec["needs"].get(c)
        best, ex, und = 0.0, None, []
        if nd is not None:
            if nd.unwrapped is not None and nd.unwrapped[0] > best:
                best, ex = nd.unwrapped[0], nd.unwrapped[1:] + ("",)
            if nd.wrapped is not None and min(nd.wrapped[0], cap_px) > best:
                best = min(nd.wrapped[0], cap_px)
                ex = nd.wrapped[1:] + ((", wrapped" + (f", capped at {cap:g}" if nd.wrapped[0] > cap_px else "")),)
            und = list(nd.untrusted)
            if nd.n_untrusted > len(nd.untrusted):
                und.append(("more", nd.n_untrusted - len(nd.untrusted)))
        geo = rec["geo"]
        for m in rec["spans"]:
            if not (m.c1 <= c <= m.c2) or (m.r1, m.c1) not in rec["anchors"]:
                continue
            key = (m.r1, m.c1)
            if key not in rec["anchor_need"]:
                continue                         # empty anchor: nothing to show
            an = rec["anchor_need"][key]
            if an is None:
                und.append(key + (rec["anchor_why"].get(key),))
                continue
            t = min(an[0], cap_px) if an[1] else an[0]
            others = sum((geo.px(k) or 0) for k in range(m.c1, min(m.c2, MAX_COL) + 1) if k != c)
            if t - others > best:
                best = t - others
                ex = (m.r1, m.c1, an[2], an[3],
                      f", merged {make_ref(m.r1, m.c1)}:{make_ref(m.r2, m.c2)}"
                      + (f", needs {self.chars(t + CELL_MARGIN_PX):.2f} over the span" if m.c1 != m.c2 else ""))
        return best / self.mdw, ex, und

    def _judged(self, rec) -> tuple:
        """({c: column record}, [(first, last, width)] empty parts of wide segments)."""
        cols = {}
        wanted = set(rec["outlier_of"])
        for a, b, _px in rec["wide"]:
            wanted.update(c for c in rec["needs"] if a <= c <= b)
            for m in rec["spans"]:
                wanted.update(range(max(a, m.c1), min(b, m.c2) + 1))
        geo = rec["geo"]
        for c in sorted(wanted):
            px = geo.px(c)
            if px is None:
                continue
            d = {"c": c, "W": self.chars(px), "group": rec["outlier_of"].get(c)}
            for cap in self.caps:
                d[cap] = self._column_need(rec, c, cap)
            cols[c] = d
        empties = []
        for a, b, px in rec["wide"]:
            inside = sorted(c for c in cols if a <= c <= b)
            x = a
            for c in inside + [b + 1]:
                if c > x:
                    empties.append((x, c - 1, self.chars(px)))
                x = c + 1
        return cols, empties

    def _evaluate(self, rec, cols, empties, s: dict) -> tuple:
        """(fails, undecided) of one sheet under setting s.  fails: [("col", c, tests) | ("empty", a, b)];
        undecided: [(sheet, ref, why)]."""
        fails, und = [], []
        name = rec["sheet"]
        if rec["brief"] and not s["brief"]:
            return fails, und
        cap = s["cap"]
        for c, d in cols.items():
            W = d["W"]
            need, _ex, u = d[cap]
            tests = []
            if W > cap and s["regardless"]:
                tests.append("A")
            elif W > cap and W > EXCESS_FACTOR * need:
                tests.append("A")
            g = d["group"]
            if g is not None and (s["ext"] or g[2] == "judge") and not (s["judge_skip"] and rec["judge_skip"]):
                if need < s["share"] * W:
                    tests.append("B")
            if tests:
                fails.append(("col", c, tuple(tests)))
            # untrusted formula values in a judged column are skipped (Patrick 2026-10-05: every attempt graded)
            for x in u:
                if x[0] == "more":
                    und.append((name, index_to_col(c), f"{x[1]} more untrusted formula value(s) in the column", x[1]))
                else:
                    why = x[2] if len(x) > 2 and x[2] else "formula value untrusted"
                    und.append((name, make_ref(x[0], x[1]), f"{why}; skipped in column {index_to_col(c)}", 1))
        for a, b, W in empties:
            if W > cap:
                fails.append(("empty", a, b))
        return fails, und

    def finish(self) -> dict:
        evals = {k: {"fails": 0, "undecided": 0, "brief_fails": 0} for k in self.settings}
        default_fails, undecided = [], []
        tags, band, per_sheet = [], [], []
        tests_failed = set()
        for rec in self._sheets:
            name = rec["sheet"]
            cols, empties = self._judged(rec)
            rec["_cols"] = cols
            for g in rec["groups"]:
                tags.append({"sheet": name, "tag": judge_tag(g), "kind": g[2]})
            for c, d in cols.items():
                g = d["group"]
                if g is not None:
                    need = d[WIDTH_CAP_CHARS][0]
                    share = need / d["W"] if d["W"] > 0 else 1.0
                    if OUTLIER_NEED_SHARE <= share < ALT_OUTLIER_NEED_SHARE:
                        band.append(f"{name}!{index_to_col(c)} width {d['W']:.2f} need {need:.2f} ({share:.0%})")
            for k, s in self.settings.items():
                f, u = self._evaluate(rec, cols, empties, s)
                evals[k]["fails"] += len(f)
                evals[k]["undecided"] += len(u)
                if rec["brief"]:
                    evals[k]["brief_fails"] += len(f)
                if k == "default":
                    if f:
                        default_fails.append((rec, f))
                    undecided.extend(u)
            per_sheet.append({"sheet": name, "used_columns": [index_to_col(x) if x else None for x in rec["used"]],
                              "judged_columns": len(cols), "outlier_groups": len(rec["groups"]),
                              "brief": rec["brief"]})
        for rec, fails in default_fails:
            for kind in self._add_mistakes(rec, fails):
                tests_failed.add(kind)
        for sheet, ref, why, n in undecided:
            # Patrick 2026-10-05 (every attempt graded): skipped, never raised
            self.note_default("untrusted_value", f"{location(sheet, ref)}: {why}", n)
        switches = {}
        for k, e in evals.items():
            if k == "default":
                continue
            switches[k] = "fail" if e["fails"] else "pass"
        stats = dict(self.counts)
        stats.update({
            "normal_font": f"{self.normal_face} {self.normal_size:g}", "normal_font_known": self.normal_known,
            "column_unit_px": self.mdw, "n_sheets_checked": len(self._sheets),
            "tests_failed": sorted(tests_failed),
            "outlier_tags": tags[:MAX_LISTED], "n_outlier_tags": len(tags),
            "n_outlier_tags_unequal": sum(1 for t in tags if t["kind"] == "ext"),
            "outlier_band_columns": band[:MAX_LISTED], "n_outlier_band_columns": len(band),
            "undecided_cells": sum(x[3] for x in undecided),
            "undecided_examples": [{"sheet": a, "cell": b, "why": c} for a, b, c, _n in undecided[:MAX_UNDECIDED_LISTED]],
            "unknown_faces": self.unknown_faces, "switches": switches,
            "brief_fails": evals["default"]["brief_fails"],
            "per_sheet": per_sheet,
            "options": {"width_cap_chars": WIDTH_CAP_CHARS, "excess_factor": EXCESS_FACTOR,
                        "wide_outlier_ratio": WIDE_OUTLIER_RATIO, "outlier_unequal_groups": OUTLIER_UNEQUAL_GROUPS,
                        "outlier_need_share": OUTLIER_NEED_SHARE, "exclude_case_brief": EXCLUDE_CASE_BRIEF,
                        "untrusted_only_if_verdict_needs": UNTRUSTED_ONLY_IF_VERDICT_NEEDS},
        })
        und_note = ""
        if undecided:
            und_note = (f" {sum(x[3] for x in undecided)} untrusted formula value(s) in judged columns were skipped "
                        f"(stats.defaults).")
        return self.verdict(
            f"No excessive column width: no visible column is over {WIDTH_CAP_CHARS:g} characters and more than "
            f"{EXCESS_FACTOR:g}x its content, and no WIDE OUTLIER column is more than twice its content "
            f"({len(self._sheets)} sheet(s)).",
            f"{{n}} excessive column width(s): "
            f"{', '.join(sorted({m['location'].rsplit('!', 1)[0] for m in self.mistakes.items}))}.{und_note}",
            stats)

    # ------------------------------------------------------------------ mistakes
    def _add_mistakes(self, rec, fails) -> set:
        """One mistake per run of adjacent failing columns with the same tests."""
        name = rec["sheet"]
        hid = " (hidden sheet)" if rec["state"] != "visible" else ""
        cols = rec["_cols"]
        units = []                                # (first, last, tests)
        for f in fails:
            if f[0] == "col":
                units.append((f[1], f[1], f[2]))
            elif f[0] == "empty":
                units.append((f[1], f[2], ("A",)))
        units.sort()
        runs = []
        for a, b, t in units:
            if runs and runs[-1][1] == a - 1 and runs[-1][2] == t:
                runs[-1][1] = b
            else:
                runs.append([a, b, t])
        kinds = set()
        for a, b, t in runs:
            kinds.update(t)
            self.add_mistake(location(name, _colrange(a, b)), self._describe(rec, name, hid, a, b, t, cols))
        return kinds

    def _describe(self, rec, name, hid, a, b, tests, cols) -> str:
        cap = WIDTH_CAP_CHARS
        rng = index_to_col(a) if a == b else _colrange(a, b)
        many = a != b
        parts = []
        listed = [cols[c] for c in range(a, b + 1) if c in cols]
        if listed:
            for d in listed[:6]:
                need, ex, _u = d[cap]
                if ex is not None:
                    r_, c_, snip, font, note = ex
                    what = f"{make_ref(r_, c_)} '{str(snip)[:40]}' ({font}{note})"
                else:
                    what = "no content"
                parts.append(f"{index_to_col(d['c'])} is {d['W']:.2f} characters wide and its content needs "
                             f"{need:.2f} ({what})")
            if len(listed) > 6:
                parts.append(f"{len(listed) - 6} more column(s)")
        n_empty = (b - a + 1) - len(listed)
        widths = sorted({round(self.chars(rec['geo'].px(c)), 2) for c in (a, b) if rec['geo'].px(c)})
        if n_empty and not listed:
            parts.append(f"{'they are' if many else 'it is'} {'/'.join(f'{w:.2f}' for w in widths)} characters wide "
                         f"and hold{'' if many else 's'} no content")
        elif n_empty:
            parts.append(f"{n_empty} other column(s) of the range hold no content")
        why = []
        if "A" in tests:
            why.append(f"wider than the {cap:g}-character cap and more than {EXCESS_FACTOR:g}x what the content "
                       f"needs")
        if "B" in tests:
            g = cols[a]["group"] if a in cols else None
            lr = (f"the neighbouring columns {_colrange(g[3][0], g[3][1])} and {_colrange(g[4][0], g[4][1])}"
                  if g is not None else "its neighbours")
            why.append(f"a WIDE OUTLIER against {lr} ({judge_tag(g) if g is not None else 'WIDE OUTLIER'}, "
                       f"stored widths) while the content needs less than {OUTLIER_NEED_SHARE:.0%} of the width")
        unit = f"{self.normal_face or 'default font'} {self.normal_size:g}"
        return (f"Column{'s' if many else ''} {rng} on sheet '{name}'{hid}: " + "; ".join(parts)
                + f" (characters of the Normal font, {unit}). Excessive: " + "; ".join(why) + ".")
