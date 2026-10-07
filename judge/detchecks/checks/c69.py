"""69 Formatting/Sufficient column widths.

Rubric: "All values display fully; no truncation or '###'" / bad: "Some columns truncate values or
show '###'".  Two rules, both "sure" (they fail a cell only when it fails under EVERY platform
model by more than SURE_TOL_PX = 2 px at 96 dpi; report.md approach C, kept by Patrick 2026-10-04):

NUMBERS ('####').  A visible numeric or date cell - constant or formula result, spill / array
members included - FAILS when the text Excel displays for it (its full number format, including a
conditional-format number format that applies) is wider than the column's text area under EVERY
platform model by more than SURE_TOL_PX.  Excel then shows '####' whatever the viewer's platform.

  * Platform models (MODELS): Windows 96 dpi and 120 dpi (Excel for Mac behaves like the 120 dpi
    grid: a 9 px digit for Calibri 11), each with per-glyph rounded advances (GDI) and fractional
    advances.  A cell that overflows under some models only is a "band" case: reported in stats,
    never a mistake.
  * Column text area = column pixels - CELL_MARGIN_PX (5), merged spans summed for the anchor cell
    (covered cells are not displayed).  Column pixels follow ECMA-376 col/@width:
    px = trunc(((256 * w + trunc(128 / MDW)) / 256) * MDW), where MDW is the Normal font's maximum
    digit width in whole pixels at the model's dpi (7 px for Calibri 11, Aptos Narrow 11 and Arial
    10 at 96 dpi; 9 px at 120 dpi).  A column without <col> is baseColWidth (8) digits + 5 px rounded
    up to a multiple of 8 px, or sheetFormatPr/@defaultColWidth through the formula; for a
    LibreOffice-style defaultColWidth the lenient reading trunc(w * MDW + 5) is also allowed in the
    non-reference models (Excel's reading of it is unverified).
  * Text width = sum of the glyph advances of the cell's own font face, size and weight (GLYPHS,
    1/1000 em per character, read once from the real font files with fontTools - see
    detchecks/scratch/colwidths/glyph_table.py).  Rounded model: each advance rounded to whole
    pixels at ppem = size * dpi / 72 (plain rounding; the fonts' hdmx tables agree with it for
    Arial, and Calibri / Aptos have none).  Padding tokens `_x` are measured as the glyph x;
    `?` blanks as a space (Excel: a digit width; lenient).  Unknown faces are measured as Calibri,
    the narrowest common face (a miss, never a false '###'), and listed in stats.
  * Alignment indent (left / right / distributed) takes INDENT_SPACES_PER_LEVEL (3) Normal-font
    spaces per level off the text area, as in Reasonable row heights (73).
  * General format never shows '####' while a shorter rendering exists: Excel rounds decimals away
    and falls back to d E+nn scientific notation.  The need of a General cell is therefore the
    narrowest of its full text, its 0-decimal rounding and its 1-digit scientific form.
  * Dates count (they are numbers; toy T3).  Negative / out-of-range dates show '#####' at any width
    and are not a width defect (NEGATIVE_DATE_FAILS = False; counted in stats).
  * Shrink-to-fit cells never show '####' (exempt).  Cells in hidden rows or columns (hidden flag,
    zero width / height) are not displayed and are skipped (No hidden rows/columns (93) owns them),
    except the anchor of a merge whose span still has a visible column and a visible row: Excel draws
    the merged value across the visible part, so it is measured against the span's visible columns in
    a second pass (review fix 69-F4; HIDDEN_ANCHOR_SPAN).
    Hidden sheets ARE graded (whole-workbook rule, as 73 and 66 do).  Fill, centre-across-selection
    and rotated numbers are skipped (how Excel clips them is unverified; counted in stats).
    Booleans and errors are not values this check measures.

CUT-OFF TEXT (Patrick 2026-10-04: long text anywhere in the workbook that cannot be fully seen fails;
text running into empty neighbours and staying visible is fine).  Option 3 (Patrick, 2026-10-04): only
a text showing at least TEXT_CLIP_MIN_CHARS (200) characters fails - shown_chars: its displayed text
(after the number format) without leading / trailing blanks and without typed line breaks; a shorter
cut-off text (a row label clipped by the next cell, ...) passes, as the toys and the judge guidance
treat it, and is counted in stats (text_cut_off_short, wrapped_cut_off_short).  Displayed text =
constants (numbers stored as text included) and trusted text formula results, through the cell's
number format (a text section can add to it or blank it).  Same models, glyph tables (characters
outside them: accented letters as their base letter, East Asian wide characters 1 em, combining
marks 0, anything else a digit), indent, tolerance and hidden-row / hidden-column / hidden-sheet
scope as the numbers.  Rich text is measured run by run (run size and weight; the reader keeps no
run face).
  (a) Unwrapped text (TEXT_CLIP) is cut off when its visible glyphs (blanks at the ends only push
      them) reach past the space it can use under EVERY model by more than SURE_TOL_PX (typed line
      breaks show as nothing).  Space = its column (or merged span) text area plus the
      contiguous EMPTY cells in the direction it overflows: general / left alignment to the right,
      right alignment to the left, center to both sides (each side must hold half the excess).
      Overflow stops at the first non-empty cell (a value, any formula result - "" included - a
      boolean or error, an empty-string constant), at a merged range (MERGES_BLOCK_OVERFLOW: merged
      cells take no overflow), and at the sheet's first / last column.  An empty cell in a hidden
      column lets the text through (0 px of room); a filled cell in a hidden column still stops it
      (HIDDEN_CELLS_BLOCK_OVERFLOW).  Text never spills out of a merge: an anchor's space is the
      merge's visible width.  Centre-across-selection text is centred over the cell and the
      contiguous empty cells to its right that carry the same alignment, then overflows both ways
      like centred text; fill text never spills (space = the cell and the contiguous empty
      fill-aligned cells to its right).  Shrink-to-fit text is never cut off; rotated text is
      skipped (counted).
  (b) Wrapped text (WRAPPED_TEXT_CLIP; wrapText, or justify / distributed alignment) is cut off when
      it needs more lines than its row shows - only in rows with a custom height (customHeight=1):
      Excel auto-fits every other row on open (docs/excel_measurements.md item 6).  Lines: the
      greedy word wrap of Reasonable row heights (73) (c73.wrap_lines) with this check's per-glyph
      widths in the cell's column (or merged span), under every model; the FEWEST lines count.
      Line height: c73's LINE_HEIGHT_PER_PT x font size.  The row shows height / line height lines;
      the text is cut off when more than WRAP_HIDDEN_LINES_TOL (half a line) is hidden.  A merge
      over several rows counts the custom heights of its visible rows; with any auto-fitted row in
      it, it is not graded (counted).

Values.  Constants are read from the file and routed by their stored type.  A formula result is
routed by its TRUSTED value, never by the delivered `t` attribute (review fix 69-F1: `t` is a cache
attribute - agent tools write t="str" placeholders with empty caches, LibreOffice a #NAME? placeholder -
so a number behind such a placeholder was never measured, even when a value copy held it).  A trusted
value (Excel cache, LibreOffice copy, unknown-writer cache) that is a number is measured; a trusted
text result goes to the text rules; a boolean / error result is not a value this check measures (the
recalc pipeline vets a LibreOffice copy's #NAME? / #VALUE! values, so with it those cells are trusted
errors and skipped, as ruled 2026-10-04).  An untrusted value (openpyxl / XlsxWriter cache, no cache, a
LibreOffice-written #NAME? / #VALUE! cache read without the pipeline, missing from the copy) leaves the
cell undecided whatever its placeholder type - unless its display cannot depend on the value: a covered
(non-anchor) merged cell is never displayed, and a format that prints nothing for every number AND
every text (';;;') shows nothing (review fix 69-F2; a conditional-format number format over the cell
keeps it undecided).  If the workbook fails on measured cells anyway, undecided cells are listed in
stats and nothing is raised (UNTRUSTED_ONLY_IF_VERDICT_NEEDS); otherwise the first one raises
GradingError (no fallback).  The same holds for a conditional-format number format whose rule cannot
be evaluated when the possible displays disagree, and for a rendering numfmt marks uncertain.

Conditional formats.  Sheets whose rules carry a number format (dxf numFmt) are streamed a second
time: every numeric cell inside such a range is rendered under the format its rules give it (c66.fires
with the shared evaluator, checks/_cfeval.py, decides which rules fire - including expressions that
read other cells and sheets), and fails when that display overflows.  Patrick 2026-10-05,
"CONDITIONAL FORMATS never stop a grading": a rule that cannot be evaluated is taken as OFF, a
conditional-format number format whose display is not verified gives way to the cell's own format,
and an ISO-date cell (t="d") is graded by its own format; every such assumption is recorded in
stats.cf_assumptions.  First-pass candidates inside those ranges are handed to the second pass.
(A conditional format does not change how text is displayed here.)
"""
from __future__ import annotations

import itertools
import math
import unicodedata
from bisect import bisect_left, bisect_right

from ..core import numfmt as N
from ..core.refs import MAX_COL, group_cells, index_to_col, location, make_ref, range_to_str
from ..core.sheet import ExcelError
from ..errors import GradingError
from .base import Check
from ._cfeval import cf_env, is_unevaluable, prefetch
from .c66 import UNKNOWN, fires, rule_text
from .c73 import LINE_HEIGHT_PER_PT, wrap_lines

# ---------------------------------------------------------------- rule constants
SURE_TOL_PX = 2.0                 # a cell fails when its overflow (96-dpi px) exceeds this under EVERY model
CELL_MARGIN_PX = 5                # column pixels not available to text (ECMA-376 padding)
MODELS = (("w96r", 96, True), ("w96f", 96, False), ("w120r", 120, True), ("w120f", 120, False))
REFERENCE_MODEL = "w96r"          # strict default-column reading; also the model quoted in messages
DEFAULT_FONT_PT = 11.0
DEFAULT_BASE_COL_WIDTH = 8.0
INDENT_SPACES_PER_LEVEL = 3       # ECMA-376 alignment/@indent: one level = 3 spaces of the Normal font
INDENT_ALIGNMENTS = ("left", "right", "distributed")
NEGATIVE_DATE_FAILS = False       # '#####' for a negative date/time is not a width defect
GENERAL_SHORTENS = True           # General format falls back to fewer decimals / scientific before '####'
UNTRUSTED_ONLY_IF_VERDICT_NEEDS = True   # undecided cells raise only while the verdict is still open
MAX_ASSUMPTION_EXAMPLES = 10
HIDDEN_ANCHOR_SPAN = True         # a merge anchor in a hidden row / column is measured against the visible span (69-F4)
# cut-off text (Patrick 2026-10-04; replaces the two narrow text rules TEXT_NUMBER_CLIP / TEXT_HIDDEN_COL_MAX_PX)
TEXT_CLIP = True                  # (a) unwrapped text that cannot be fully seen (blocked overflow) fails
WRAPPED_TEXT_CLIP = True          # (b) wrapped text needing more lines than its custom-height row shows fails
TEXT_CLIP_MIN_CHARS = 200         # Patrick 2026-10-04 (option 3): (a) and (b) fail only for a text showing at least
                                  # this many characters (shown_chars); shorter cut-off text is counted in stats only
WRAP_HIDDEN_LINES_TOL = 0.5       # (b) ... when more than this many lines are hidden (a line counts as shown
                                  #     when at least half of it is visible)
WRAP_ALIGNMENTS = ("justify", "distributed")    # alignments Excel wraps like wrapText (as 73)
MAX_FONT_PT = 409.0               # Excel's largest font size: a rich-text run size outside (0, 409] is out of range
                                  # (e.g. ChatGPT's writer stores 1000 for 10 pt); the cell's size is used, counted
MERGES_BLOCK_OVERFLOW = True      # (a) a merged range takes no overflowing text from a neighbour
HIDDEN_CELLS_BLOCK_OVERFLOW = True  # (a) a filled cell in a hidden column still stops overflowing text
MAX_UNDECIDED_LISTED = 12
MAX_BAND_LISTED = 8
MAX_TEXT_EXAMPLES = 8
TEXT_SNIPPET_CHARS = 30           # characters of a cut-off text quoted in a mistake
MAX_CANDIDATES = 2_000_000        # buffered overflow candidates per sheet; more raises
RENDER_CACHE_MAX = 400_000

# ---------------------------------------------------------------- glyph metrics (stdlib only)
# Advance widths in 1/1000 em of every character a rendered number or date can contain, per face and
# weight, read once from the real font files (detchecks/scratch/colwidths/glyph_table.py, fontTools).
# 0 = the face lacks the glyph (measured as its digit).  Order = GLYPH_CHARS.
GLYPH_CHARS = '0123456789.,\'()-+%/:$€£¥ ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!"#&*;<=>?@[\\]^_`{|}~\xa0−–—’°×…'
GLYPHS = {
    ("calibri", False): (507,507,507,507,507,507,507,507,507,507,252,250,221,303,303,306,498,715,386,268,507,507,507,507,226,579,544,533,615,488,459,631,623,252,319,520,420,855,646,662,517,673,543,459,487,642,567,890,519,487,468,479,525,423,525,498,305,471,525,229,239,455,229,799,525,527,525,525,349,391,335,525,452,715,433,453,395,326,401,498,682,498,268,498,498,498,463,894,307,386,307,498,498,291,314,460,314,498,226,498,498,905,250,339,498,690),
    ("calibri", True): (507,507,507,507,507,507,507,507,507,507,267,258,233,312,312,306,498,729,430,276,507,507,507,507,226,606,561,529,630,488,459,637,631,267,331,547,423,874,659,676,532,686,563,473,495,653,591,906,551,520,478,494,537,418,537,503,316,474,537,246,255,480,246,813,537,538,537,537,355,399,347,537,473,745,459,474,397,326,438,498,705,498,276,498,498,498,463,898,325,430,325,498,498,300,344,475,344,498,226,498,498,905,258,342,498,711),
    ("calibri light", False): (507,507,507,507,507,507,507,507,507,507,245,245,214,299,299,306,498,707,362,263,507,507,507,507,226,563,535,535,607,489,460,627,619,244,312,504,419,845,638,654,508,666,532,453,483,636,554,881,501,469,463,471,520,425,520,494,299,469,520,221,230,441,221,791,520,521,520,520,345,387,329,520,440,699,418,441,394,326,380,498,670,498,263,498,498,498,463,892,297,362,297,498,498,286,299,453,299,498,226,498,498,905,245,337,498,679),
    ("calibri light", True): (507,507,507,507,507,507,507,507,507,507,267,258,233,312,312,306,498,729,430,276,507,507,507,507,226,606,561,529,630,488,459,637,631,267,331,547,423,874,659,676,532,686,563,473,495,653,591,906,551,520,478,494,537,418,537,503,316,474,537,246,255,480,246,813,537,538,537,537,355,399,347,537,473,745,459,474,397,326,438,498,705,498,276,498,498,498,463,898,325,430,325,498,498,300,344,475,344,498,226,498,498,905,258,342,498,711),
    ("arial", False): (556,556,556,556,556,556,556,556,556,556,278,278,191,333,333,333,584,889,278,278,556,556,556,556,278,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,278,355,556,667,389,278,584,584,584,556,1015,278,278,278,469,556,333,334,260,334,584,278,584,556,1000,222,400,584,1000),
    ("arial", True): (556,556,556,556,556,556,556,556,556,556,278,278,238,333,333,333,584,889,278,333,556,556,556,556,278,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,333,474,556,722,389,333,584,584,584,611,975,333,278,333,584,556,333,389,280,389,584,278,584,556,1000,278,400,584,1000),
    ("arial narrow", False): (456,456,456,456,456,456,456,456,456,456,228,228,157,273,273,273,479,729,228,228,456,456,456,456,228,547,547,592,592,547,501,638,592,228,410,547,456,683,592,638,547,638,592,547,501,592,547,774,547,547,501,456,456,410,456,456,228,456,456,182,182,410,182,683,456,456,456,456,273,410,228,456,410,592,410,410,410,228,291,456,547,319,228,479,479,479,456,832,228,228,228,385,456,273,274,213,274,479,228,479,456,820,182,400,479,820),
    ("arial narrow", True): (456,456,456,456,456,456,456,456,456,456,228,228,195,273,273,273,479,729,228,273,456,456,456,456,228,592,592,592,592,547,501,638,592,228,456,592,501,683,592,638,547,638,592,547,501,592,547,774,547,547,501,456,501,456,501,456,273,501,501,228,228,456,228,729,501,501,501,501,319,456,273,501,456,638,456,456,410,273,389,456,592,319,273,479,479,479,501,800,273,228,273,479,456,273,319,230,319,479,456,479,456,820,228,400,479,820),
    ("aptos narrow", False): (507,507,507,507,507,507,507,507,507,507,260,260,192,304,304,306,507,761,313,260,507,507,507,507,187,534,553,633,627,509,480,646,646,243,305,522,457,722,645,667,529,667,555,519,436,621,531,813,505,493,470,487,514,479,514,485,278,450,507,224,224,448,243,784,507,505,514,514,309,443,296,510,414,656,406,414,399,326,339,498,588,498,260,507,507,507,455,806,271,313,271,507,416,499,271,462,271,507,187,507,416,832,247,365,507,792),
    ("aptos narrow", True): (507,507,507,507,507,507,507,507,507,507,267,267,207,307,307,306,507,752,354,267,507,507,507,507,181,553,558,634,633,515,481,641,653,272,333,562,459,734,646,658,544,658,569,538,450,618,549,840,537,526,489,499,530,496,530,506,304,461,525,251,252,482,271,806,525,517,530,530,336,459,319,525,445,677,444,445,417,326,377,498,610,498,267,507,507,507,459,766,304,354,304,507,405,485,304,466,304,507,181,507,405,809,267,377,507,815),
    ("aptos", False): (534,534,534,534,534,534,534,534,534,534,286,286,210,293,293,340,534,826,339,286,534,534,534,534,203,589,604,692,686,556,524,708,707,260,331,568,500,790,706,732,577,732,606,566,479,681,585,892,553,540,516,531,561,525,561,527,301,484,551,239,239,487,260,853,551,552,561,561,334,486,323,559,452,721,442,452,438,293,370,537,643,457,286,534,534,534,501,895,294,339,294,534,460,552,294,270,294,534,203,534,460,920,270,365,534,818),
    ("aptos", True): (534,534,534,534,534,534,534,534,534,534,300,300,232,293,293,340,534,841,387,300,534,534,534,534,203,620,619,707,707,573,537,718,729,294,367,623,512,821,723,737,605,737,634,598,505,696,617,940,599,586,547,552,586,549,586,556,335,511,579,268,268,534,296,889,579,573,586,586,369,516,355,582,494,761,495,494,466,293,420,537,677,457,300,534,534,534,518,900,337,387,337,534,460,552,337,270,337,534,203,534,460,920,298,378,534,850),
    ("times new roman", False): (500,500,500,500,500,500,500,500,500,500,250,250,180,333,333,333,564,833,278,278,500,500,500,500,250,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,333,408,500,778,500,278,564,564,564,444,921,333,278,333,469,500,333,480,200,480,541,250,564,500,1000,333,400,564,1000),
    ("times new roman", True): (500,500,500,500,500,500,500,500,500,500,250,250,278,333,333,333,570,1000,278,333,500,500,500,500,250,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,333,555,500,833,500,333,570,570,570,500,930,333,278,333,581,500,333,394,220,394,520,250,570,500,1000,333,400,570,1000),
    ("verdana", False): (636,636,636,636,636,636,636,636,636,636,364,364,269,454,454,454,818,1076,454,454,636,636,636,636,352,684,686,698,771,632,575,775,751,421,455,693,557,843,748,787,603,787,695,684,616,732,684,989,685,615,685,601,623,521,623,596,352,623,633,274,344,592,274,973,633,607,623,623,427,521,394,633,592,818,592,592,525,394,459,818,727,636,454,818,818,818,545,1000,454,454,454,818,636,636,635,454,635,818,352,818,636,1000,269,542,818,818),
    ("verdana", True): (711,711,711,711,711,711,711,711,711,711,361,361,332,543,543,480,867,1272,689,402,711,711,711,711,342,776,762,724,830,683,650,811,837,546,555,771,637,948,847,850,733,850,782,710,682,812,764,1128,764,737,692,668,699,588,699,664,422,699,712,342,403,671,342,1058,712,687,699,699,497,593,456,712,650,979,669,651,597,402,587,867,862,711,402,867,867,867,617,964,543,689,543,867,711,711,711,543,711,867,342,867,711,1000,332,587,867,1049),
    ("tahoma", False): (546,546,546,546,546,546,546,546,546,546,303,303,211,383,383,363,728,977,382,354,546,546,546,546,312,600,589,601,678,561,521,667,675,373,417,588,498,771,667,708,551,708,621,557,584,656,597,902,581,576,559,525,553,461,553,526,318,553,558,229,282,498,229,840,558,543,553,553,360,446,334,558,498,742,495,498,444,332,401,728,674,546,354,728,728,728,474,909,383,382,383,728,546,546,480,382,480,728,312,728,546,909,211,471,728,817),
    ("tahoma", True): (637,637,637,637,637,637,637,637,637,637,312,312,275,454,454,431,818,1199,577,363,637,637,637,637,293,685,686,667,757,615,581,745,764,483,500,696,572,893,771,770,657,770,726,633,612,739,675,1028,685,670,623,599,632,527,629,594,382,629,640,302,363,603,302,954,640,617,629,629,434,515,416,640,579,890,604,576,526,343,489,818,781,637,363,818,818,818,566,920,454,577,454,818,637,546,623,637,623,818,293,818,637,909,275,520,818,1000),
    ("georgia", False): (614,430,559,552,565,528,566,502,596,566,270,270,215,375,375,374,643,817,469,312,610,642,622,615,241,671,654,642,749,653,599,725,815,390,518,694,604,927,767,744,610,744,702,561,619,756,667,976,710,615,602,504,560,454,574,483,325,509,582,293,292,536,286,881,591,539,571,560,410,432,345,575,497,737,505,492,444,331,412,643,710,472,312,643,643,643,479,929,375,469,375,643,643,500,430,375,430,643,241,643,643,857,227,419,643,807),
    ("georgia", True): (701,490,626,625,649,599,648,554,676,648,328,328,269,447,447,379,703,879,472,367,641,715,690,732,254,758,757,715,834,721,671,807,913,446,595,817,686,1023,839,820,701,820,797,649,684,833,762,1126,809,732,689,596,646,531,663,572,393,577,680,354,346,632,344,1016,690,636,658,648,520,513,397,677,567,863,588,562,525,376,510,703,799,482,367,703,703,703,548,967,447,472,447,703,703,500,500,388,500,703,254,703,703,928,269,420,703,942),
    ("courier new", False): (600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600),
    ("courier new", True): (600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600,600),
    ("consolas", False): (550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550),
    ("consolas", True): (550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550,550),
    ("cambria", False): (554,554,554,554,554,554,554,554,554,554,205,205,237,382,382,332,554,890,490,264,506,628,528,609,220,623,611,562,662,575,537,611,687,324,307,629,537,815,681,653,568,653,621,496,593,648,604,921,571,570,538,488,547,441,555,488,303,494,552,278,266,524,271,832,558,531,556,547,414,430,338,552,504,774,483,504,455,286,393,619,688,427,264,554,554,554,422,885,350,490,350,554,371,285,387,316,387,712,220,554,500,1000,221,375,554,752),
    ("cambria", True): (592,592,592,592,592,592,592,592,592,592,232,232,251,408,408,337,592,976,505,280,543,640,556,641,220,652,651,573,705,578,551,646,722,350,341,682,551,846,679,695,614,695,662,513,639,676,634,961,619,604,566,535,591,469,597,531,326,520,597,314,302,592,308,890,604,569,597,591,461,459,365,597,531,798,525,531,479,335,422,618,740,453,280,592,592,592,452,921,368,505,368,592,371,285,393,320,393,592,220,592,500,1000,235,378,592,772),
    ("century gothic", False): (554,554,554,554,554,554,554,554,554,554,277,277,198,369,369,332,606,775,437,277,554,554,554,554,277,740,574,813,744,536,485,872,683,226,482,591,462,919,740,869,592,871,607,498,426,655,702,960,609,592,480,683,682,647,685,650,314,673,610,200,203,502,200,938,610,655,682,682,301,388,339,608,554,831,480,536,425,295,309,720,757,425,277,606,606,606,591,867,351,605,351,672,500,378,351,672,351,606,280,606,500,1000,351,400,606,1000),
    ("century gothic", True): (560,560,560,560,560,560,560,560,560,560,280,280,220,380,380,420,600,860,460,280,560,560,560,560,280,740,580,780,700,520,480,840,680,280,480,620,440,900,740,840,560,840,580,520,420,640,700,900,680,620,500,660,660,640,660,640,280,660,600,240,260,580,240,940,600,640,660,660,320,440,300,600,560,800,560,580,460,280,360,600,680,440,280,600,600,600,560,740,320,640,320,600,500,420,340,600,340,600,280,600,500,1000,280,400,600,1000),
    ("candara", False): (549,350,462,486,533,491,552,473,550,548,252,252,254,353,353,252,504,504,267,252,444,444,444,444,217,616,586,549,643,521,489,621,652,276,416,592,490,854,674,694,553,694,605,511,509,683,554,886,558,550,519,491,556,452,550,513,339,540,542,218,218,488,229,819,542,559,556,556,354,419,359,536,480,766,504,465,456,252,452,504,682,504,252,504,504,504,353,1009,353,267,353,504,504,274,353,252,353,504,217,504,500,1000,252,267,500,1000),
    ("candara", True): (534,334,479,471,537,473,529,465,536,528,252,252,252,353,353,252,504,504,267,252,444,444,444,444,217,630,598,552,645,527,491,622,653,279,399,589,485,864,665,697,556,697,596,523,526,658,574,917,587,552,515,491,555,445,550,512,362,551,550,249,249,510,258,832,550,559,560,555,385,419,363,541,489,788,515,489,435,252,457,504,682,504,252,504,504,504,353,1009,353,267,353,504,504,274,353,252,353,504,217,504,500,1000,252,267,500,1000),
    ("corbel", False): (514,448,511,453,517,480,524,428,515,524,264,264,196,300,300,333,513,805,275,264,513,513,513,513,200,635,593,589,671,551,504,668,667,246,375,605,521,822,696,730,570,739,592,552,555,671,607,882,584,594,585,490,536,438,536,498,318,531,532,232,241,487,232,829,527,533,531,523,335,404,347,518,462,712,455,479,444,264,362,655,669,538,267,513,513,513,426,1011,315,275,315,513,488,342,300,227,300,513,200,513,488,830,219,408,513,792),
    ("corbel", True): (515,495,499,488,528,489,550,488,556,551,292,317,210,306,306,333,527,827,294,292,527,527,527,527,206,657,623,584,689,568,521,672,698,276,402,632,538,826,728,734,601,760,628,577,570,700,638,910,636,634,601,515,557,441,558,513,328,550,560,249,263,516,252,852,555,554,550,550,350,424,379,545,491,749,521,514,468,292,395,665,692,528,317,527,527,527,449,1024,315,294,315,527,488,391,305,237,305,527,206,527,488,830,243,408,527,877),
    ("dejavu sans", False): (636,636,636,636,636,636,636,636,636,636,318,318,275,390,390,361,838,950,337,337,636,636,636,636,318,684,686,698,770,632,575,775,752,295,295,656,557,863,748,787,603,787,695,635,611,732,684,989,685,611,685,613,635,550,635,615,352,635,634,278,278,579,278,974,634,612,635,635,411,521,392,634,592,818,592,592,525,401,460,838,780,500,337,838,838,838,531,1000,390,337,390,838,500,500,636,337,636,838,318,838,500,1000,318,500,838,1000),
    ("dejavu sans", True): (696,696,696,696,696,696,696,696,696,696,380,380,306,457,457,415,838,1002,365,400,696,696,696,696,348,774,762,734,830,683,683,821,837,372,372,775,637,995,837,850,733,850,770,720,682,812,774,1103,771,724,725,675,716,593,716,678,435,716,712,343,343,665,343,1042,712,687,716,716,493,595,478,712,652,924,645,652,582,456,521,838,872,523,400,838,838,838,580,1000,457,365,457,838,500,500,712,365,712,838,348,838,500,1000,380,500,838,1000),
}
FACE_ALIASES = {
    "helvetica": "arial", "helvetica neue": "arial", "liberation sans": "arial", "arimo": "arial",
    "arial unicode ms": "arial", "roboto": "arial",       # Roboto digits 0.562 em vs Arial 0.556 (judge table)
    "carlito": "calibri", "aptos display": "aptos", "aptos serif": "cambria",
    "liberation serif": "times new roman", "tinos": "times new roman",
    "liberation mono": "courier new", "cousine": "courier new",
}
DEFAULT_FACE = "calibri"          # unknown faces: the narrowest common face (a miss, never a false ###)
_CHAR_INDEX = {ch: i for i, ch in enumerate(GLYPH_CHARS)}
_DIGIT_IDX = _CHAR_INDEX["0"]
_SPACE_IDX = _CHAR_INDEX[" "]


def face_key(name) -> tuple[str, bool]:
    """(table key, known?) for a font name."""
    k = (name or "").strip().lower()
    k = FACE_ALIASES.get(k, k)
    if (k, False) in GLYPHS:
        return k, True
    return DEFAULT_FACE, False


def glyph_row(face_k: str, bold: bool) -> tuple:
    row = GLYPHS.get((face_k, bool(bold)))
    if row is None:
        row = GLYPHS[(face_k, False)]
    return row


def adv_em1000(row: tuple, ch: str) -> int:
    """Advance of one character in 1/1000 em (a missing glyph counts as a digit)."""
    i = _CHAR_INDEX.get(ch)
    if i is None:
        return row[_DIGIT_IDX]
    return row[i] or row[_DIGIT_IDX]


def text_em1000(row: tuple, text: str) -> int:
    return sum(adv_em1000(row, ch) for ch in text)


def text_px(row: tuple, text: str, ppem: float, rounded: bool) -> float:
    """Width of text in pixels at ppem: per-glyph rounded advances (GDI) or fractional."""
    if rounded:
        return float(sum(round(adv_em1000(row, ch) * ppem / 1000.0) for ch in text))
    return text_em1000(row, text) * ppem / 1000.0


def ppem_at(size_pt: float, dpi: float) -> float:
    return (size_pt if size_pt and size_pt > 0 else DEFAULT_FONT_PT) * dpi / 72.0


def mdw_px(face_k: str, size_pt: float, dpi: float) -> int:
    """Excel's maximum digit width of the Normal font: whole pixels at the model's dpi."""
    row = glyph_row(face_k, False)
    p = ppem_at(size_pt, dpi)
    return max(1, int(max(round(row[_CHAR_INDEX[d]] * p / 1000.0) for d in "0123456789")))


def space_px(face_k: str, size_pt: float, dpi: float) -> int:
    row = glyph_row(face_k, False)
    return max(1, int(round(row[_SPACE_IDX] * ppem_at(size_pt, dpi) / 1000.0)))


def col_px(width_ch: float, mdw: int) -> int:
    """ECMA-376 col/@width (characters incl. padding) -> pixels."""
    if width_ch is None or width_ch <= 0:
        return 0
    return int(math.trunc(((256.0 * width_ch + math.trunc(128.0 / mdw)) / 256.0) * mdw))


def default_col_px(fmt, mdw: int) -> tuple[int, int]:
    """(strict, lenient) pixel width of a column without <col>: baseColWidth digits + 5 px rounded up
    to 8 px, or defaultColWidth through the ECMA formula (strict) / trunc(w * MDW + 5) (lenient:
    Excel's reading of a LibreOffice-written 8.5390625 is unverified)."""
    base = fmt.base_col_width if fmt.base_col_width is not None else DEFAULT_BASE_COL_WIDTH
    base_px = int(math.ceil((base * mdw + 5) / 8.0) * 8)
    dcw = fmt.default_col_width
    if dcw is not None and dcw > 0:
        strict = col_px(dcw, mdw)
        return strict, max(strict, int(math.trunc(dcw * mdw + 5)), base_px)
    return base_px, base_px


def general_alternatives(v: float, full: str) -> list:
    """Texts Excel may fall back to for a General-format number in a narrow column: the full text,
    the 0-decimal rounding, and 1-significant-digit scientific notation."""
    out = [full]
    if not GENERAL_SHORTENS or not isinstance(v, (int, float)) or isinstance(v, bool):
        return out
    if math.isfinite(v):
        if abs(v) < 1e15:
            r = int(math.floor(abs(v) + 0.5))
            out.append(("-" if v < 0 and r else "") + str(r))
        if v != 0:
            out.append(f"{v:.0E}")
    return out


def blank_for_any_number(code: str) -> bool:
    """True when the format prints nothing for EVERY number (';;;', ';;'): all its numeric sections are
    empty.  A numeric cell's display then cannot depend on its value (text results are never measured,
    booleans and errors neither), so an untrusted formula result under it is not undecided (69-F2).
    Conservative: a tag-only section ('[Red]', '[>100]') renders General and is not empty."""
    try:
        f = N.parse_format(code)
    except GradingError:
        return False
    if not f.numeric:
        return False
    return all(f.sections[i].is_empty for i in f.numeric)


def blank_for_any_value(code: str) -> bool:
    """True when the format prints nothing for every number AND every text (';;;': all numeric sections
    empty and a text section that shows nothing).  Under ';;' (no text section) text shows as typed, so
    since the cut-off text rule an untrusted formula result there can still show something and stays
    undecided (69-F2, narrowed 2026-10-04)."""
    if not blank_for_any_number(code):
        return False
    f = N.parse_format(code)
    if f.text_index is None:
        return False
    sec = f.sections[f.text_index]
    if sec.has_at or sec.kind == "general":
        return False
    try:
        return N.render("x", code, value_type="s").is_blank
    except GradingError:
        return False


# ---------------------------------------------------------------- text characters (cut-off text rule)
WIDE_CHAR, ZERO_CHAR = -1, -2
_RESOLVED: dict = {}


def resolve_char(ch: str):
    """Glyph-table index of a text character, or WIDE_CHAR (East Asian wide / full-width: 1 em), ZERO_CHAR
    (combining mark: 0), None (anything else: measured as a digit, the rule for a glyph the table lacks).
    Accented letters are measured as their base letter, a tab as a space.  Shared with Reasonable column
    widths (70)."""
    if ch in _RESOLVED:
        return _RESOLVED[ch]
    i = _CHAR_INDEX.get(ch)
    if i is None:
        if ch == "\t":
            i = _SPACE_IDX
        elif unicodedata.combining(ch):
            i = ZERO_CHAR
        elif unicodedata.east_asian_width(ch) in ("W", "F"):
            i = WIDE_CHAR
        else:
            base = unicodedata.normalize("NFD", ch)[:1]
            i = _CHAR_INDEX.get(base) if base != ch else None
    _RESOLVED[ch] = i
    return i


def char_em1000(row: tuple, ch: str) -> int:
    """Advance of one TEXT character in 1/1000 em (resolve_char; numbers keep adv_em1000)."""
    i = resolve_char(ch)
    if i is None:
        return row[_DIGIT_IDX]
    if i >= 0:
        return row[i] or row[_DIGIT_IDX]
    return 1000 if i == WIDE_CHAR else 0


def _deficits(e, kind: str, area: float, room_l: float, room_r: float) -> tuple:
    """(right, left) px by which a text's visible glyphs reach past the space they may use on each side
    (None: the text does not reach that side).  e = (reach right, reach left) from C69._unwrapped_ext,
    measured from the cell's left edge (left-aligned, fill), its right edge (right-aligned) or the centre
    of its area (centred, centre-across); area = the cell's (or span's) text area, rooms = the empty
    cells it may overflow into."""
    er, el = e
    if kind == "spill_both" or kind == "across":
        return er - area / 2.0 - room_r, el - area / 2.0 - room_l
    return (None if er is None else er - area - room_r, None if el is None else el - area - room_l)


def _worst(d) -> float:
    return max(x for x in d if x is not None)


def shown_chars(text: str) -> int:
    """Characters a cut-off text shows, for TEXT_CLIP_MIN_CHARS: its displayed text (the value after the
    cell's number format, e.g. with a text section's prefix / suffix), without leading / trailing blanks
    (spaces, tabs, non-breaking spaces, line breaks) and without the typed line breaks inside it (an
    unwrapped cell shows them as nothing, a wrapped one as a new line).  Every other character counts once,
    blanks inside the text included."""
    return len(text.strip().replace("\r", "").replace("\n", ""))


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


class _Style:
    __slots__ = ("shrink", "skip", "code", "fmt_err", "row", "face", "known", "size", "bold", "indent_levels",
                 "indent_sides", "horizontal", "wrap", "rot", "always_blank", "face_k", "text_kind", "text_mode",
                 "text_err", "line_pt")

    def __init__(self):
        self.skip = None


class _Model:
    __slots__ = ("name", "dpi", "rounded", "mdw", "space", "scale")

    def __init__(self, name, dpi, rounded, mdw, space):
        self.name, self.dpi, self.rounded, self.mdw, self.space = name, dpi, rounded, mdw, space
        self.scale = 96.0 / dpi          # model px -> 96-dpi px


class _SheetGeo:
    """Column text areas per model for one sheet (lazy, cached per column).  A column's width comes
    from the core reader's one rule (SheetHead.col_info / col_run): where <col> entries overlap, the
    later entry wins, as in Reasonable column widths (70) and Reasonable row heights (73)."""
    __slots__ = ("head", "models", "dflt", "cache")

    def __init__(self, head, models):
        self.head, self.models = head, models
        self.dflt = {}
        for m in models:
            strict, lenient = default_col_px(head.format, m.mdw)
            self.dflt[m.name] = strict if m.name == REFERENCE_MODEL else lenient
        self.cache = {}

    def col_pixels(self, c: int):
        """{model: px} for column c, or None when the column is hidden / zero width."""
        v = self.cache.get(c, 0)
        if v != 0:
            return v
        ci = self.head.col_info(c)
        if ci is not None and ci.hidden:
            v = None
        elif ci is not None and ci.width is not None:
            v = None if ci.width <= 0 else {m.name: col_px(ci.width, m.mdw) for m in self.models}
        else:
            v = dict(self.dflt)
        self.cache[c] = v
        return v

    def span_pixels(self, c1: int, c2: int):
        """{model: px} summed over columns c1..c2 (hidden columns count 0); None if all hidden."""
        tot = {m.name: 0 for m in self.models}
        any_vis = False
        for c in range(c1, min(c2, c1 + 16383) + 1):
            p = self.col_pixels(c)
            if p is None:
                continue
            any_vis = True
            for k, px in p.items():
                tot[k] += px
        return tot if any_vis else None

    def room(self, a: int, b: int, target=None) -> dict:
        """{model: px} summed over the columns a..b (hidden / zero-width columns count 0: an empty cell
        there lets text through without giving it room).  Walks runs of columns that share one applying
        <col> entry (SheetHead.col_run, the reading col_info gives); with target ({model: px}) it stops
        once every model has at least its target (the sum is then a lower bound, which is all a fit
        decision needs)."""
        tot = {m.name: 0 for m in self.models}
        a, b = max(a, 1), min(b, MAX_COL)
        if a > b:
            return tot
        x = a
        while x <= b:
            end = min(self.head.col_run(x)[1], b)
            px = self.col_pixels(x)
            if px is not None:
                n = end - x + 1
                for k in tot:
                    tot[k] += px[k] * n
                if target is not None and all(tot[k] >= target[k] for k in tot):
                    break
            x = end + 1
        return tot


class _CfRule:
    __slots__ = ("prio", "order", "rule", "boxes", "anchor", "fmt", "stop")


class C69(Check):
    number = 69
    key = "Formatting/Sufficient column widths"
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
        nk, _known = face_key(self.normal_face)
        self.normal_key = nk
        self.models = [_Model(n, dpi, rnd, mdw_px(nk, self.normal_size, dpi), space_px(nk, self.normal_size, dpi))
                       for n, dpi, rnd in MODELS]
        self.ref_model = next(m for m in self.models if m.name == REFERENCE_MODEL)
        self._styles = {}
        self._render_cache = {}
        self._width_cache = {}
        self._sheets = []            # per-sheet records
        self._pass2 = {}
        self.unknown_faces = {}
        self.n_values_read = 0
        self.n_undecided = 0
        self.undecided = []          # [{sheet, ref, why}]
        self._env = None
        self.cf_assumed = 0          # numeric cells measured under a conditional-format assumption (Patrick 2026-10-05)
        self.cf_assumed_decisive = 0
        self.cf_assumed_examples = []
        self.cf_unevaluable = {}     # (sheet, rule order) -> description
        self.band = []               # cells overflowing under some models only (stats)
        self.n_band = 0
        self.counts = {"numeric_cells": 0, "measured": 0, "sure_overflows": 0, "hidden_row_cells": 0,
                       "hidden_col_cells": 0, "hidden_anchor_cells": 0, "shrink_to_fit": 0, "skipped_alignment": 0,
                       "general_cells": 0, "negative_dates": 0, "blank_display": 0, "cf_cells": 0,
                       "formula_nonnumeric": 0, "undecided_not_displayed": 0}
        # cut-off text (a: unwrapped, b: wrapped); each count is described in docs/checks/69.md ("Mistakes")
        self.tcounts = {"text_cells": 0, "text_formula_results": 0, "text_hidden_col_cells": 0, "text_shrink": 0,
                        "text_rotated_skipped": 0, "text_blank_format": 0, "text_wider_than_cell": 0,
                        "text_covered_merged": 0, "text_fits_with_spill": 0, "text_cut_off": 0, "text_band": 0,
                        "text_center_across": 0, "text_fill": 0, "text_rows_under_half_line": 0,
                        "wrapped_text_cells": 0, "wrapped_auto_rows": 0, "wrapped_custom_rows": 0,
                        "wrapped_cut_off": 0, "wrapped_band": 0, "wrapped_merge_auto_rows": 0,
                        "wrapped_merged_in_auto_rows": 0, "text_hidden_anchor_cells": 0,
                        "text_cut_off_short": 0, "wrapped_cut_off_short": 0, "text_runs_size_out_of_range": 0}
        self.text_blockers = {}      # what stops graded cut-off text: value kinds, merge, sheet edge, own merge, fill
        self.text_blockers_short = {}   # ... and cut-off text under TEXT_CLIP_MIN_CHARS (counted only)
        self.text_short_examples = []
        self.wrapped_short_examples = []
        self.text_unknown_faces = {}
        self.text_band_examples = []
        self.text_short_row_examples = []
        self.wrapped_merged_auto_examples = []
        self._char_fns = {}          # (id(glyph row), ppem, rounded) -> character -> px (cached tables)

    # ------------------------------------------------------------------ styles
    def _style(self, s) -> _Style:
        v = self._styles.get(s)
        if v is not None:
            return v
        v = _Style()
        st = self.st
        xf = st.xf(s)
        font = st.font(s)
        al = xf.alignment
        v.shrink = bool(al.shrink_to_fit)
        v.horizontal = al.horizontal
        v.wrap = bool(al.wrap_text)
        v.rot = int(al.text_rotation or 0)
        v.skip = None
        if al.horizontal in ("fill", "centerContinuous"):
            v.skip = al.horizontal
        elif v.rot:
            v.skip = "rotated"
        v.fmt_err = None
        try:
            v.code = N.resolve_format(xf.num_fmt_id, st.num_fmts)
            N.parse_format(v.code)               # an unreadable code raises here, not at a cell
        except GradingError as e:
            # Patrick 2026-10-06 (attempt 2777, measured in Excel): Excel drops an unreadable format on
            # repair, so the cell shows General; recorded per cell (unreadable_number_format)
            v.code, v.fmt_err = "General", str(e)
        # an untrusted formula result is "not displayed whatever its value" only when the format prints
        # nothing for numbers AND text (';;;'); under ';;' a text result would show (cut-off text rule)
        v.always_blank = v.fmt_err is None and blank_for_any_value(v.code)
        k, known = face_key(font.name)
        v.face, v.known, v.face_k = (font.name or "").strip() or None, known, k
        v.row = glyph_row(k, font.b)
        v.size = float(font.sz) if font.sz and font.sz > 0 else DEFAULT_FONT_PT
        v.bold = bool(font.b)
        v.indent_levels = int(al.indent or 0) if al.indent and al.horizontal in INDENT_ALIGNMENTS else 0
        v.indent_sides = 2 if al.horizontal == "distributed" else 1
        # cut-off text: how a text in this style is laid out (direction of overflow, or why not graded)
        h = al.horizontal
        if v.rot:
            v.text_kind = "rotated"              # skipped (how Excel clips rotated text is unverified)
        elif v.shrink:
            v.text_kind = "shrink"               # shrinks to fit: never cut off
        elif v.wrap or h in WRAP_ALIGNMENTS or al.vertical in WRAP_ALIGNMENTS:
            v.text_kind = "wrap"                 # (b)
        elif h == "right":
            v.text_kind = "spill_left"           # right-aligned: overflows to the left
        elif h == "center":
            v.text_kind = "spill_both"           # half the excess on each side
        elif h == "centerContinuous":
            v.text_kind = "across"               # centred over its selection, then both ways
        elif h == "fill":
            v.text_kind = "fill"                 # never spills
        else:
            v.text_kind = "spill_right"          # general / left (/ unknown): overflows to the right
        v.text_mode, v.text_err = "plain", None
        if v.fmt_err is not None:
            v.text_err = v.fmt_err               # General: text shows plain; recorded when a text needs it
        else:
            try:
                f = N.parse_format(v.code)
                if f.text_index is not None and f.sections[f.text_index].kind != "general":
                    v.text_mode = "format"       # a text section decides what is shown (prefix, padding, blank)
            except GradingError as e:
                v.text_mode, v.text_err = "error", str(e)
        v.line_pt = LINE_HEIGHT_PER_PT * v.size  # Reasonable row heights (73)'s line height
        self._styles[s] = v
        return v

    # ------------------------------------------------------------------ measuring
    def _pads(self, code: str, section_index):
        """Padding characters (`_x` tokens) of the section used, measured as the glyph x instead of
        the space the renderer printed."""
        if section_index is None or "_" not in code:
            return ""
        try:
            secs = N.parse_format(code).sections
            toks = secs[section_index].tokens
        except (IndexError, AttributeError, GradingError):
            return ""
        return "".join(v for k, v in toks if k == "pad")

    def _need_px(self, sty: _Style, text: str, pads: str, m: _Model) -> float:
        """Text width under model m in model pixels, indent included."""
        key = (text, pads, id(sty.row), sty.size, m.name)
        w = self._width_cache.get(key)
        if w is None:
            p = ppem_at(sty.size, m.dpi)
            w = text_px(sty.row, text, p, m.rounded)
            if pads:
                w += text_px(sty.row, pads, p, m.rounded) - len(pads) * text_px(sty.row, " ", p, m.rounded)
            if len(self._width_cache) > RENDER_CACHE_MAX:
                self._width_cache.clear()
            self._width_cache[key] = w
        if sty.indent_levels:
            w += sty.indent_levels * INDENT_SPACES_PER_LEVEL * m.space * sty.indent_sides
        return w

    def _overflows(self, sty: _Style, texts: list, pads: str, pixels: dict) -> dict:
        """{model: overflow in 96-dpi px} using the narrowest of `texts` per model."""
        out = {}
        for m in self.models:
            avail = pixels[m.name] - CELL_MARGIN_PX
            need = min(self._need_px(sty, t, pads, m) for t in texts)
            out[m.name] = (need - avail) * m.scale
        return out

    def _render(self, v, t, code: str):
        key = (v, t, code)
        r = self._render_cache.get(key)
        if r is None:
            r = N.render(v, code, value_type=("d" if t == "d" else None), date1904=self.date1904)
            if len(self._render_cache) > RENDER_CACHE_MAX:
                self._render_cache.clear()
            self._render_cache[key] = r
        return r

    def _measure(self, sty: _Style, v, t, code: str, pixels: dict):
        """(status, overflows, text): status 'fit' | 'over' | 'band' | 'hash' | 'blank' | 'uncertain'."""
        r = self._render(v, t, code)
        if r.is_hash:
            return "hash", None, r.text
        if r.is_blank:
            return "blank", None, r.text
        texts = [r.text]
        if r.kind == "general":
            self.counts["general_cells"] += 1
            vn = v
            if t == "d":
                # an ISO-typed date under General shows its serial (the rendered text); the
                # shortenings start from that number, not from the ISO string (review fix 69-F5)
                try:
                    vn = float(r.text)
                except ValueError:
                    vn = None
            texts = general_alternatives(vn, r.text)
        pads = self._pads(code, r.section_index)
        # exact pre-screen: the fractional 96-dpi model is one of the models, so if it does not
        # overflow beyond the tolerance the cell is no sure overflow
        m96f = self.models[1]
        avail = pixels[m96f.name] - CELL_MARGIN_PX
        if min(self._need_px(sty, x, pads, m96f) for x in texts) - avail <= SURE_TOL_PX:
            return "fit", None, r.text
        ov = self._overflows(sty, texts, pads, pixels)
        if min(ov.values()) > SURE_TOL_PX:
            if not r.certain:
                return "uncertain", ov, r.text
            return "over", ov, r.text
        return "band", ov, r.text

    # ------------------------------------------------------------------ per sheet
    def sheet_start(self, head):
        f = head.format
        self.geo = _SheetGeo(head, self.models)
        self._head = head
        self._sheet_name = head.name
        self._zero_default = bool(f.zero_height)
        self._row_hidden = False
        self._hidden_rows = set()
        self._cur_row = None
        self._shown_rows = set()     # zeroHeight sheets only: rows with a visible <row> (absent rows are hidden)
        self._cands = []             # [(r, c, s, v, t, text, ov_min_96r, note)] sure overflows (base format)
        self._uncertain = []         # [(r, c, ref, why)]
        self._undec_sheet = []       # untrusted formula results [(r, c, ref, s)], settled at sheet_end
        self._hidden_valued = 0      # value / formula cells in hidden rows or columns (merge anchors among them: pass 2)
        self._sheet_counts = {"numeric_cells": 0, "sure_overflows": 0, "values_read": 0}
        # cut-off text: the current row (blockers are its filled cells) and the sheet's candidates
        self._row_ht = None          # custom height (pt) of the current row, None = auto-fitted by Excel
        self._row_style = None
        self._custom_ht = {}         # r -> custom height of every visible custom-height row (merged blocks)
        self._rf_cols = []           # filled columns of the current row, and what fills them
        self._rf_kinds = []
        self._r_sel = {}             # empty cells with fill / centre-across alignment: col -> horizontal
        self._r_cands = []           # text cells of the current row wider than their own cell
        self._tcands = []            # [(r, c, s, text, segments, ext, kind, formula, lb, lbk, rb, rbk, sel_end)]
        self._wcands = []            # wrapped text in custom-height rows that may need more lines than shown
        self._wauto = []             # wrapped multi-line text in auto-fitted rows (merge anchors counted only)
        self._sheet_short = [0, 0]   # cut-off texts under TEXT_CLIP_MIN_CHARS on this sheet: (a), (b)

    def row(self, row):
        self._flush_row()
        self._cur_row = row.r
        ht = row.ht
        self._row_hidden = bool(row.hidden) or (ht is not None and ht <= 0) or (ht is None and self._zero_default)
        if self._row_hidden:
            self._hidden_rows.add(row.r)
        elif self._zero_default:
            self._shown_rows.add(row.r)
        self._row_style = row.style
        self._row_ht = None
        if not self._row_hidden and row.custom_height and ht is not None and ht > 0:
            self._row_ht = ht
            self._custom_ht[row.r] = ht

    def _span_has_visible_row(self, r1: int, r2: int) -> bool:
        """Is any row r1..r2 displayed?  Rows without <row> are visible unless the sheet is zeroHeight."""
        n = r2 - r1 + 1
        if self._zero_default:
            shown = self._shown_rows
            if n <= len(shown):
                return any(r in shown for r in range(r1, r2 + 1))
            return any(r1 <= r <= r2 for r in shown)
        hidden = self._hidden_rows
        if n > len(hidden):
            return True
        return any(r not in hidden for r in range(r1, r2 + 1))

    def cell(self, cell):
        if self._row_hidden:
            if cell.raw is not None or cell.is_formula_result:
                self.counts["hidden_row_cells"] += 1
                self._hidden_valued += 1
            return
        if cell.row != self._cur_row:            # defensive: a cell outside its <row>
            self._flush_row()
            self._cur_row = cell.row
            self._row_ht = self._custom_ht.get(cell.row)
        c = cell.col
        if cell.is_formula_result:
            # a formula result stops overflowing text whatever it shows (even ""); its delivered type
            # is a cache attribute and says nothing reliable (69-F1): the trusted value decides
            self._rf_cols.append(c)
            self._rf_kinds.append("formula")
            routed = self._formula_value(cell)
            self._numeric(cell, cell.t, routed)
            kind, v, _t = routed
            if kind == "other" and isinstance(v, str) and not isinstance(v, ExcelError):
                self.tcounts["text_formula_results"] += 1
                self._text(cell, v, None, True)
            elif kind == "untrusted":
                self._text_untrusted(cell)
            return
        t = cell.t
        if t == "n" or t == "d":
            if cell.raw is not None and cell.raw != "":
                self._rf_cols.append(c)
                self._rf_kinds.append("date" if t == "d" else "number")
            else:
                self._note_empty(cell)
            if cell.raw is not None:
                self._numeric(cell, t)
            return
        if t in ("s", "inlineStr", "str"):
            if cell.raw is None:
                self._note_empty(cell)
                return
            self._rf_cols.append(c)                 # an empty-string constant stops overflow too
            self._rf_kinds.append("text")
            v = cell.value
            if isinstance(v, str):
                self._text(cell, v, cell.rich_runs, False)
            return
        # booleans, errors: not values this check measures, but they stop overflowing text
        if cell.raw is not None and cell.raw != "":
            self._rf_cols.append(c)
            self._rf_kinds.append("boolean" if t == "b" else "error value")
        else:
            self._note_empty(cell)

    # ------------------------------------------------------------------ numeric cells
    def _route(self, cell):
        """(kind, value, t) for any cell: constants by their stored type (a number / ISO date with a
        value element is 'number', anything else 'other'), formula results by _formula_value."""
        if cell.is_formula_result:
            return self._formula_value(cell)
        t = cell.t
        if (t == "n" or t == "d") and cell.raw is not None:
            return "number", cell.value, t
        return "other", None, t

    def _formula_value(self, cell):
        """What a formula result holds, decided by its TRUSTED value and never by the delivered `t`
        (a cache attribute: agent tools write t="str" placeholders with empty caches, LibreOffice a
        #NAME? placeholder).  -> (kind, value, t): kind 'number' (t 'n', or 'd' for a cached ISO
        date), 'other' (text, boolean, error, empty: not measured) or 'untrusted' (undecided)."""
        if not cell.value_trusted:
            return "untrusted", None, "n"
        v = cell.value
        if v is None or isinstance(v, bool):
            return "other", v, cell.t
        if isinstance(v, (int, float)):
            return "number", v, "n"
        if cell.t == "d" and isinstance(v, str) and cell.value_source == "cached":
            return "number", v, "d"
        return "other", v, cell.t                   # text (str), error (ExcelError)

    def _numeric(self, cell, t, routed=None):
        formula = cell.is_formula_result
        kind = "number"
        if formula:
            kind, v, t = routed if routed is not None else self._formula_value(cell)
            if kind == "other":
                self.counts["formula_nonnumeric"] += 1
                return
        else:
            v = cell.value
        self.counts["numeric_cells"] += 1
        self._sheet_counts["numeric_cells"] += 1
        sty = self._style(cell.s)
        if sty.shrink:
            self.counts["shrink_to_fit"] += 1
            return
        pixels = self.geo.col_pixels(cell.col)
        if pixels is None:
            self.counts["hidden_col_cells"] += 1
            self._hidden_valued += 1
            return
        if sty.skip:
            self.counts["skipped_alignment"] += 1
            return
        if not sty.known:
            self.unknown_faces[sty.face or "(none)"] = self.unknown_faces.get(sty.face or "(none)", 0) + 1
        if sty.fmt_err:
            self.note_default("unreadable_number_format", f"{location(cell.sheet, cell.ref)}: {sty.fmt_err}")
        if kind == "untrusted":
            # settled at sheet_end: dropped when never displayed (covered merged cell) or blank
            # whatever the value (';;;' outside conditional-format ranges), else undecided
            self._undec_sheet.append((cell.row, cell.col, cell.ref, cell.s))
            return
        if formula:
            self.n_values_read += 1
            self._sheet_counts["values_read"] += 1
        if v is None or isinstance(v, bool) or (t != "d" and not isinstance(v, (int, float))):
            return
        if t != "d" and not math.isfinite(v):
            return
        self.counts["measured"] += 1
        status, ov, text = self._measure(sty, v, t, sty.code, pixels)
        self._settle(status, ov, text, cell.row, cell.col, cell.s, v, t, cell.ref, sty, self._cands, self._uncertain)

    def _settle(self, status, ov, text, r, c, s, v, t, ref, sty, cands, uncertain):
        """File one measured cell: a sure overflow into `cands`, an unverified rendering into
        `uncertain` (as (r, c, ref, why)), everything else into the counts."""
        if status == "fit":
            return
        if status == "blank":
            self.counts["blank_display"] += 1
            return
        if status == "hash":
            self.counts["negative_dates"] += 1
            if NEGATIVE_DATE_FAILS:
                cands.append((r, c, s, v, t, text, 999.0, "negative date"))
            return
        if status == "uncertain":
            # Patrick 2026-10-05 (every attempt graded): the engine's best rendering is measured, and it overflows
            # under every model
            self.note_default("unverified_number_format", f"{location(self._sheet_name, ref)} = {v!r} under "
                                                          f"{sty.code!r} shows {text.strip()!r}")
            cands.append((r, c, s, v, t, text, ov[REFERENCE_MODEL], "an unverified number format, best rendering"))
            if len(cands) > MAX_CANDIDATES:
                raise GradingError(f"{self.key}: more than {MAX_CANDIDATES:,} overflowing cells on sheet "
                                   f"{self._sheet_name!r}; too many to grade")
            return
        if status == "band":
            self.n_band += 1
            if len(self.band) < MAX_BAND_LISTED:
                self.band.append({"cell": f"{self._sheet_name}!{ref}", "text": text,
                                  "overflow_px_by_model": {k: round(x, 1) for k, x in ov.items()}})
            return
        cands.append((r, c, s, v, t, text, ov[REFERENCE_MODEL], None))
        if len(cands) > MAX_CANDIDATES:
            raise GradingError(f"{self.key}: more than {MAX_CANDIDATES:,} overflowing cells on sheet "
                               f"{self._sheet_name!r}; too many to grade")

    # ------------------------------------------------------------------ cut-off text: measuring
    def _char_fn(self, row, ppem: float, rounded: bool):
        """Character -> px of a glyph row at ppem (each advance rounded to whole px, or fractional), cached."""
        key = (id(row), ppem, rounded)
        fn = self._char_fns.get(key)
        if fn is None:
            tab = {}

            def fn(ch, tab=tab, row=row, k=ppem / 1000.0, rounded=rounded):
                w = tab.get(ch)
                if w is None:
                    a = char_em1000(row, ch) * k
                    w = tab[ch] = float(round(a)) if rounded else a
                return w
            self._char_fns[key] = fn
        return fn

    def _segments(self, sty: _Style, text: str, runs) -> tuple:
        """((glyph row, size, text), ...) of a displayed text: rich runs in their own size and weight (the
        reader keeps no run face: the cell's face is used), plain text in the cell font."""
        if not runs:
            return ((sty.row, sty.size, text),)
        segs = []
        rest = len(text) - sum(len(rn.text) for rn in runs)
        if rest > 0:
            segs.append((sty.row, sty.size, text[:rest]))
        for rn in runs:
            if not rn.text:
                continue
            f = rn.font
            if f is None:
                segs.append((sty.row, sty.size, rn.text))
                continue
            size = sty.size
            if "sz" in f.specified and f.sz:
                if 0 < f.sz <= MAX_FONT_PT:
                    size = float(f.sz)
                else:
                    # out of Excel's range: how Excel shows it is unverified, so it never enlarges the text
                    self.tcounts["text_runs_size_out_of_range"] += 1
            bold = bool(f.b) if "b" in f.specified else sty.bold
            segs.append((glyph_row(sty.face_k, bold), size, rn.text))
        return tuple(segs)

    def _char_widths(self, segs, m: _Model) -> list:
        out = []
        for row, size, t in segs:
            out.extend(map(self._char_fn(row, ppem_at(size, m.dpi), m.rounded), t))
        return out

    def _text_parts(self, segs, n: int, i0: int, j0: int, m: _Model) -> tuple:
        """(leading blanks, visible part, trailing blanks) widths in model px of a displayed text of n
        characters whose visible part (first to last non-blank character) is [i0, j0)."""
        if len(segs) == 1:
            row, size, t = segs[0]
            fn = self._char_fn(row, ppem_at(size, m.dpi), m.rounded)
            if i0 == 0 and j0 == n:
                return 0.0, sum(map(fn, t)), 0.0
            return sum(map(fn, t[:i0])), sum(map(fn, t[i0:j0])), sum(map(fn, t[j0:]))
        w = self._char_widths(segs, m)
        return sum(w[:i0]), sum(w[i0:j0]), sum(w[j0:])

    def _indent_px(self, sty: _Style, m: _Model) -> float:
        if not sty.indent_levels:
            return 0.0
        return sty.indent_levels * INDENT_SPACES_PER_LEVEL * m.space * sty.indent_sides

    def _em_fn(self, row):
        """Character -> advance in 1/1000 em for a glyph row (cached)."""
        key = (id(row), 0, None)
        fn = self._char_fns.get(key)
        if fn is None:
            tab = {}

            def fn(ch, tab=tab, row=row):
                a = tab.get(ch)
                if a is None:
                    a = tab[ch] = char_em1000(row, ch)
                return a
            self._char_fns[key] = fn
        return fn

    def _unwrapped_ext(self, sty: _Style, disp: str, segs, kind: str, pixels: dict):
        """Per-model reach (right, left) of an unwrapped text's visible glyphs, in model px (see _deficits), or
        None when it fits its own cell (within SURE_TOL_PX) under EVERY model: it then needs no room at all.
        Blanks at the ends are invisible: a leading blank pushes left-aligned text right, a trailing one
        pushes right-aligned text left, and centred text is centred with both."""
        n = len(disp)
        centred = kind == "spill_both" or kind == "across"
        if len(segs) == 1:
            # cheap screen: the fractional width (one pass) plus half a pixel per glyph bounds every model
            row, size, t = segs[0]
            em = sum(map(self._em_fn(row), t))
            for m in self.models:
                w = em * ppem_at(size, m.dpi) / 1000.0 + (0.5 * n if m.rounded else 0.0)
                area = pixels[m.name] - CELL_MARGIN_PX
                reach = w / 2.0 - area / 2.0 if centred else self._indent_px(sty, m) + w - area
                if reach * m.scale > SURE_TOL_PX:
                    break
            else:
                return None
        i0 = n - len(disp.lstrip())
        j0 = len(disp.rstrip())
        ext = []
        over = False
        for m in self.models:
            lead, vis, trail = self._text_parts(segs, n, i0, j0, m)
            ind = self._indent_px(sty, m)
            if kind == "spill_right" or kind == "fill":
                e = (ind + lead + vis, None)
            elif kind == "spill_left":
                e = (None, ind + trail + vis)
            else:                                    # centred / centred across selection
                w = lead + vis + trail
                e = (w / 2.0 - trail, w / 2.0 - lead)
            if _worst(_deficits(e, kind, pixels[m.name] - CELL_MARGIN_PX, 0.0, 0.0)) * m.scale > SURE_TOL_PX:
                over = True
            ext.append(e)
        return tuple(ext) if over else None

    def _wrap_lines_models(self, row, size: float, text: str, pixels: dict, sty: _Style) -> tuple:
        """Lines a wrapped text needs under each model: Reasonable row heights (73)'s greedy word wrap
        (c73.wrap_lines) with this check's per-glyph widths, in the column (or span) text area minus the
        alignment indent."""
        out = []
        for m in self.models:
            usable = pixels[m.name] - CELL_MARGIN_PX - self._indent_px(sty, m)
            out.append(wrap_lines(text, 1.0, usable, em=self._char_fn(row, ppem_at(size, m.dpi), m.rounded)))
        return tuple(out)

    def _wrap_font(self, sty: _Style, disp: str, runs):
        """(glyph row, size) a wrapped text is laid out in: the cell font; for rich text the narrowest
        (smallest size, regular unless every run is bold) - lenient, so a cut-off is sure."""
        if not runs:
            return sty.row, sty.size
        segs = self._segments(sty, disp, runs)
        size = min(sz for _r, sz, t in segs if t)
        bold = all(r is glyph_row(sty.face_k, True) for r, _sz, t in segs if t) and \
            glyph_row(sty.face_k, True) is not glyph_row(sty.face_k, False)
        return glyph_row(sty.face_k, bold), size

    @staticmethod
    def _snippet(text: str) -> str:
        t = " ".join(text.split())
        return t if len(t) <= TEXT_SNIPPET_CHARS else t[:TEXT_SNIPPET_CHARS].rstrip() + "…"

    # ------------------------------------------------------------------ cut-off text: streaming
    def _note_empty(self, cell):
        """An empty cell lets overflowing text through; with fill / centre-across alignment it can also be
        part of a neighbour's fill / centre-across selection."""
        h = self._style(cell.s).horizontal
        if h == "fill" or h == "centerContinuous":
            self._r_sel[cell.col] = h

    def _text_untrusted(self, cell):
        """An untrusted formula result in a fill / centre-across cell: the number rule skips such cells
        (unverified), but a text there is graded, so its value is needed (undecided, settled at sheet_end)."""
        sty = self._style(cell.s)
        if sty.skip in ("fill", "centerContinuous") and not sty.shrink and self.geo.col_pixels(cell.col) is not None:
            self._undec_sheet.append((cell.row, cell.col, cell.ref, cell.s))

    def _display_text(self, cell, sty: _Style, s: str):
        """What a text shows under the cell's number format (a text section can add to it or blank it),
        or None when nothing visible is shown."""
        if sty.text_err:
            self.note_default("unreadable_number_format", f"{location(cell.sheet, cell.ref)}: {sty.text_err}")
        disp = s
        if sty.text_mode == "format":
            r = N.render(s, sty.code, value_type="s")
            if r.is_blank:
                return None
            disp = r.text
        return disp if disp.strip() else None

    def _text(self, cell, s: str, runs, formula: bool):
        """A displayed text (constant, or trusted text formula result): the cut-off text rules (a) / (b)."""
        tc = self.tcounts
        tc["text_cells"] += 1
        if not s:
            return                                   # "" shows nothing (it still stops overflow)
        sty = self._style(cell.s)
        kind = sty.text_kind
        if kind == "rotated":
            tc["text_rotated_skipped"] += 1
            return
        if kind == "shrink":
            tc["text_shrink"] += 1
            return
        pixels = self.geo.col_pixels(cell.col)
        if pixels is None:
            # not displayed; a merge anchor here is measured on the merge's visible part (69-F4, pass 2)
            tc["text_hidden_col_cells"] += 1
            self._hidden_valued += 1
            return
        disp = self._display_text(cell, sty, s)
        if disp is None:
            tc["text_blank_format"] += 1
            return
        if disp != s:
            runs = None                              # rearranged by a text section: measured in the cell font
        if not sty.known:
            nm = sty.face or "(none)"
            self.text_unknown_faces[nm] = self.text_unknown_faces.get(nm, 0) + 1
        if kind == "wrap":
            self._wrapped(cell, sty, disp, runs, pixels, formula)
            return
        if kind == "across":
            tc["text_center_across"] += 1
        elif kind == "fill":
            tc["text_fill"] += 1
        if self._row_ht is not None and self._row_ht < 0.5 * sty.line_pt:
            # stats only: one line of text in a custom-height row lower than half a line (not graded)
            tc["text_rows_under_half_line"] += 1
            if len(self.text_short_row_examples) < MAX_TEXT_EXAMPLES:
                self.text_short_row_examples.append(f"{self._sheet_name}!{cell.ref} row {cell.row} = "
                                                    f"{self._row_ht:g} pt: '{self._snippet(disp)}'")
        disp, segs = self._one_line(sty, disp, runs)
        if disp is None:
            return
        ext = self._unwrapped_ext(sty, disp, segs, kind, pixels)
        if ext is None:
            return
        tc["text_wider_than_cell"] += 1
        self._r_cands.append((cell.row, cell.col, cell.s, disp, segs, ext, kind, formula))

    def _one_line(self, sty: _Style, disp: str, runs):
        """(text, segments) of an unwrapped text as Excel shows it on one line: a typed line break shows as
        nothing.  (None, None) when nothing visible is left."""
        segs = self._segments(sty, disp, runs)
        if "\n" in disp or "\r" in disp:
            segs = tuple((r_, sz, t.replace("\r", "").replace("\n", "")) for r_, sz, t in segs)
            disp = "".join(t for _r, _sz, t in segs)
            if not disp.strip():
                return None, None
        return disp, segs

    def _wrapped(self, cell, sty: _Style, disp: str, runs, pixels: dict, formula: bool):
        """(b) wrapped text: candidate when its row has a custom height that may show fewer lines than
        the text needs (merges settled at sheet_end)."""
        tc = self.tcounts
        tc["wrapped_text_cells"] += 1
        text = disp.rstrip()                         # trailing blanks / line breaks show nothing
        row, size = self._wrap_font(sty, disp, runs)
        line_pt = LINE_HEIGHT_PER_PT * size
        ht = self._row_ht
        if ht is None:
            # auto-fitted row: Excel fits it to the text on open (docs/excel_measurements.md item 6).  Its
            # AutoFit ignores merged cells, though: a merge anchor needing several lines is counted at
            # sheet_end (stats only, not graded)
            tc["wrapped_auto_rows"] += 1
            m = self.models[3]
            usable = pixels[m.name] - CELL_MARGIN_PX - self._indent_px(sty, m)
            if wrap_lines(text, 1.0, usable, em=self._char_fn(row, ppem_at(size, m.dpi), m.rounded)) > 1:
                self._wauto.append((cell.row, cell.col, cell.s, text, row, size))
            return
        tc["wrapped_custom_rows"] += 1
        lines = self._wrap_lines_models(row, size, text, pixels, sty)
        if max(lines) - ht / line_pt <= WRAP_HIDDEN_LINES_TOL:
            return
        self._wcands.append((cell.row, cell.col, cell.s, text, row, size, line_pt, ht, lines, formula))

    def _flush_row(self):
        """End of a row: find what stops each overflowing text of the row (the nearest filled cell on
        each side; merges are applied at sheet_end)."""
        cands = self._r_cands
        if cands:
            cols, kinds = self._rf_cols, self._rf_kinds
            if not HIDDEN_CELLS_BLOCK_OVERFLOW:
                keep = [i for i, c in enumerate(cols) if self.geo.col_pixels(c) is not None]
                cols, kinds = [cols[i] for i in keep], [kinds[i] for i in keep]
            if any(cols[i] > cols[i + 1] for i in range(len(cols) - 1)):
                order = sorted(range(len(cols)), key=cols.__getitem__)
                cols, kinds = [cols[i] for i in order], [kinds[i] for i in order]
            n = len(cols)
            for r, c, s, disp, segs, ext, kind, formula in cands:
                i = bisect_left(cols, c)
                lb, lbk = (cols[i - 1], kinds[i - 1]) if i > 0 else (0, "edge")
                j = bisect_right(cols, c)
                rb, rbk = (cols[j], kinds[j]) if j < n else (MAX_COL + 1, "edge")
                sel_end = c
                if kind == "across" or kind == "fill":
                    sel_end = self._selection_end(c, rb, "centerContinuous" if kind == "across" else "fill")
                self._tcands.append((r, c, s, disp, segs, ext, kind, formula, lb, lbk, rb, rbk, sel_end))
            if len(self._tcands) > MAX_CANDIDATES:
                raise GradingError(f"{self.key}: more than {MAX_CANDIDATES:,} overflowing text cells on sheet "
                                   f"{self._sheet_name!r}; too many to grade")
            self._r_cands = []
        if self._rf_cols:
            self._rf_cols = []
            self._rf_kinds = []
        if self._r_sel:
            self._r_sel = {}

    def _selection_end(self, c: int, rb: int, h: str) -> int:
        """Last column of a fill / centre-across selection starting at c: the contiguous empty cells to its
        right with the same alignment (their own, or for a cell without <c> the row's or column's style)."""
        x = c + 1
        head = self._head
        while x < rb:
            hx = self._r_sel.get(x)
            if hx is None:
                s = self._row_style if self._row_style is not None else head.col_style(x)
                hx = self._style(s).horizontal if s is not None else None
            if hx != h:
                break
            x += 1
        return x - 1

    # ------------------------------------------------------------------ cut-off text: sheet end
    @staticmethod
    def _row_merges(merges, rows) -> dict:
        """{row: [(c1, c2, merge)]} of the multi-cell merges covering the given rows."""
        if not merges or not rows:
            return {}
        rs = sorted(rows)
        out = {}
        for m in merges:
            if m.is_single_cell:
                continue
            for k in range(bisect_left(rs, m.r1), bisect_right(rs, m.r2)):
                out.setdefault(rs[k], []).append((m.c1, m.c2, m))
        return out

    def _text_fit(self, ext, kind, own: dict, start: int, end: int, lb: int, rb: int, spill: bool):
        """Overflow of a text's visible glyphs past the space it can use, per model -> ({model: 96-dpi px},
        reference-model (right, left) deficits, area and rooms).  spill=False: no room beyond start..end."""
        d0 = [_deficits(ext[k], kind, own[m.name] - CELL_MARGIN_PX, 0.0, 0.0) for k, m in enumerate(self.models)]
        room_r = room_l = None
        if spill:
            need_r = {m.name: (d[0] if d[0] is not None and d[0] > 0 else 0.0) for m, d in zip(self.models, d0)}
            need_l = {m.name: (d[1] if d[1] is not None and d[1] > 0 else 0.0) for m, d in zip(self.models, d0)}
            if any(v > 0 for v in need_r.values()):
                room_r = self.geo.room(end + 1, rb - 1, need_r)
            if any(v > 0 for v in need_l.values()):
                room_l = self.geo.room(lb + 1, start - 1, need_l)
        over, ref = {}, None
        for k, m in enumerate(self.models):
            dr, dl = d0[k]
            rr = room_r[m.name] if room_r is not None else 0.0
            rl = room_l[m.name] if room_l is not None else 0.0
            if dr is not None:
                dr -= rr
            if dl is not None:
                dl -= rl
            over[m.name] = _worst((dr, dl)) * m.scale
            if m is self.ref_model:
                ref = {"dr": dr, "dl": dl, "area": own[m.name] - CELL_MARGIN_PX, "room_r": rr, "room_l": rl}
        return over, ref

    def _settle_text(self, merges) -> list:
        """(a): decide the sheet's overflowing unwrapped texts now that merges are known -> cut-off records."""
        cands = self._tcands
        out = []
        if not cands:
            return out
        tc = self.tcounts
        ml = self._merge_lookup(merges, [(x[0], x[1]) for x in cands])
        rowm = self._row_merges(merges, {x[0] for x in cands}) if MERGES_BLOCK_OVERFLOW else {}
        for r, c, s, disp, segs, ext, kind, formula, lb, lbk, rb, rbk, sel_end in cands:
            m = ml.get((r, c))
            if m is not None:
                if (m.r1, m.c1) != (r, c):
                    tc["text_covered_merged"] += 1   # a covered cell is not displayed
                    continue
                own = self.geo.span_pixels(m.c1, m.c2)
                if own is None:
                    continue
                start, end, spill = m.c1, m.c2, False
                why_r = why_l = ("own_merge", m)
            else:
                start, end = c, sel_end
                own = self.geo.col_pixels(c) if end == c else self.geo.span_pixels(c, end)
                if own is None:
                    continue
                spill = kind != "fill"
                why_r, why_l = (("fill", end) if kind == "fill" else ("cell", rb, rbk)), ("cell", lb, lbk)
                for c1, c2, mm in rowm.get(r, ()):
                    if end < c1 < rb:
                        rb, why_r = c1, ("merge", mm)
                    if lb < c2 < start:
                        lb, why_l = c2, ("merge", mm)
            over, ref = self._text_fit(ext, kind, own, start, end, lb, rb, spill)
            rec = self._text_decide(r, c, s, disp, segs, kind, formula, over, ref, why_r, why_l, m)
            if rec is not None:
                out.append(rec)
        return out

    def _text_decide(self, r, c, s, disp, segs, kind, formula, over, ref, why_r, why_l, merge):
        """A sure cut-off (every model past SURE_TOL_PX) -> its record; band / fit -> counted only."""
        tc = self.tcounts
        lo, hi = min(over.values()), max(over.values())
        if lo > SURE_TOL_PX:
            blockers = []
            if ref["dr"] is not None and ref["dr"] > 0:
                blockers.append(("right", why_r, self._blocker_phrase("right", why_r, r)))
            if ref["dl"] is not None and ref["dl"] > 0:
                blockers.append(("left", why_l, self._blocker_phrase("left", why_l, r)))
            n = shown_chars(disp)
            short = n < TEXT_CLIP_MIN_CHARS
            counts = self.text_blockers_short if short else self.text_blockers
            for _side, why, _phrase in blockers:
                b = self._blocker_kind(why)
                counts[b] = counts.get(b, 0) + 1
            if short:
                # cut off but shorter than TEXT_CLIP_MIN_CHARS: counted, not a mistake (Patrick, option 3)
                tc["text_cut_off_short"] += 1
                self._sheet_short[0] += 1
                if len(self.text_short_examples) < MAX_TEXT_EXAMPLES:
                    self.text_short_examples.append(
                        {"cell": f"{self._sheet_name}!{make_ref(r, c)}", "text": self._snippet(disp), "chars": n,
                         "blocked_by": [ph for _s, _w, ph in blockers],
                         "overflow_px_by_model": {k: round(x, 1) for k, x in over.items()}})
                return None
            tc["text_cut_off"] += 1
            return {"r": r, "c": c, "s": s, "disp": disp, "segs": segs, "kind": kind, "formula": formula,
                    "over": over[REFERENCE_MODEL], "over_by_model": over, "ref": ref, "blockers": blockers,
                    "merge": merge, "chars": n}
        if hi > SURE_TOL_PX:
            tc["text_band"] += 1
            if len(self.text_band_examples) < MAX_BAND_LISTED:
                self.text_band_examples.append({"cell": f"{self._sheet_name}!{make_ref(r, c)}",
                                                "text": self._snippet(disp),
                                                "overflow_px_by_model": {k: round(x, 1) for k, x in over.items()}})
        else:
            tc["text_fits_with_spill"] += 1
        return None

    def _hidden_anchor_text(self, cell, v, m, p):
        """69-F4 for text: a merge anchor in a hidden row / column shows its text across the merge's visible
        part.  Unwrapped text cannot overflow a merge; wrapped text must fit the visible rows of the merge
        (custom heights only)."""
        runs = None
        if cell.is_formula_result:
            if not isinstance(v, str) or isinstance(v, ExcelError):
                return
            s, formula = v, True
        else:
            if cell.t not in ("s", "inlineStr", "str") or cell.raw is None:
                return
            s, runs, formula = cell.value, cell.rich_runs, False
            if not isinstance(s, str):
                return
        if not s:
            return
        sty = self._style(cell.s)
        if sty.text_kind in ("rotated", "shrink"):
            return
        pixels = self.geo.span_pixels(m.c1, m.c2)
        if pixels is None:
            return
        disp = self._display_text(cell, sty, s)
        if disp is None:
            return
        if disp != s:
            runs = None
        self.tcounts["text_hidden_anchor_cells"] += 1
        if sty.text_kind == "wrap":
            text = disp.rstrip()
            row, size = self._wrap_font(sty, disp, runs)
            block = self._block_height(m, *p["row_info"])
            if block is None or block <= 0:
                return
            lines = self._wrap_lines_models(row, size, text, pixels, sty)
            rec = self._wrap_decide(cell.row, cell.col, cell.s, text, size, LINE_HEIGHT_PER_PT * size, block,
                                    lines, m, formula, pixels)
            if rec is not None:
                p["wrap_cut"].append(rec)
            return
        kind = sty.text_kind
        disp, segs = self._one_line(sty, disp, runs)
        if disp is None:
            return
        ext = self._unwrapped_ext(sty, disp, segs, kind, pixels)
        if ext is None:
            return
        over, ref = self._text_fit(ext, kind, pixels, m.c1, m.c2, 0, MAX_COL + 1, False)
        rec = self._text_decide(cell.row, cell.col, cell.s, disp, segs, kind, formula, over, ref,
                                ("own_merge", m), ("own_merge", m), m)
        if rec is not None:
            p["text_cut"].append(rec)

    def _blocker_phrase(self, side: str, why, r: int) -> str:
        """What stops a cut-off text on one side, for the mistake text."""
        k = why[0]
        if k == "own_merge":
            return f"the edge of its merged range {why[1].ref} (text never spills out of a merge)"
        if k == "merge":
            return f"the merged range {why[1].ref} (merged cells take no overflow)"
        if k == "fill":
            return "its fill alignment (fill text does not spill into other cells)"
        col, ck = why[1], why[2]
        if ck == "edge":
            return f"the {side} edge of the sheet"
        hidden = self.geo.col_pixels(col) is None
        art = {"number": "a number", "date": "a date", "text": "text", "formula": "a formula result",
               "boolean": "a boolean", "error value": "an error value"}.get(ck, ck)
        return f"{make_ref(r, col)} ({art}{', in hidden column ' + index_to_col(col) if hidden else ''})"

    def _blocker_kind(self, why) -> str:
        """Stats key of what stops a cut-off text."""
        if why[0] != "cell":
            return {"own_merge": "own merged range", "merge": "merged range", "fill": "fill alignment"}[why[0]]
        if why[2] == "edge":
            return "sheet edge"
        hidden = self.geo.col_pixels(why[1]) is None
        return f"{why[2]}{' (hidden column)' if hidden else ''}"

    def _block_height(self, m, hidden_rows, custom_ht, zero_default, shown_rows):
        """Height (pt) of the visible rows of a merge, or None when one of them is auto-fitted (Excel
        then decides its height: not graded)."""
        tot = 0.0
        for rr in range(m.r1, m.r2 + 1):
            if rr in hidden_rows or (zero_default and rr not in shown_rows):
                continue
            h = custom_ht.get(rr)
            if h is None:
                return None
            tot += h
        return tot

    def _settle_wrapped(self, merges) -> list:
        """(b): decide the sheet's wrapped-text candidates now that merges are known -> cut-off records."""
        tc = self.tcounts
        out = []
        cands = self._wcands
        ml = self._merge_lookup(merges, [(x[0], x[1]) for x in cands] + [(x[0], x[1]) for x in self._wauto])
        for r, c, s, text, row, size, line_pt, ht, lines, formula in cands:
            m = ml.get((r, c))
            block = ht
            px = self.geo.col_pixels(c)
            if m is not None:
                if (m.r1, m.c1) != (r, c):
                    continue                         # a covered cell is not displayed
                px = self.geo.span_pixels(m.c1, m.c2)
                if px is None:
                    continue
                if m.r2 > m.r1:
                    block = self._block_height(m, self._hidden_rows, self._custom_ht, self._zero_default,
                                               self._shown_rows)
                    if block is None:
                        tc["wrapped_merge_auto_rows"] += 1
                        continue
                lines = self._wrap_lines_models(row, size, text, px, self._style(s))
            rec = self._wrap_decide(r, c, s, text, size, line_pt, block, lines, m, formula, px)
            if rec is not None:
                out.append(rec)
        for r, c, s, text, row, size in self._wauto:
            m = ml.get((r, c))
            if m is None or (m.r1, m.c1) != (r, c):
                continue
            px = self.geo.span_pixels(m.c1, m.c2)
            if px is None:
                continue
            lines = min(self._wrap_lines_models(row, size, text, px, self._style(s)))
            rows = sum(1 for rr in range(m.r1, m.r2 + 1)
                       if rr not in self._hidden_rows and not (self._zero_default and rr not in self._shown_rows))
            if lines > max(rows, 1):
                tc["wrapped_merged_in_auto_rows"] += 1
                if len(self.wrapped_merged_auto_examples) < MAX_TEXT_EXAMPLES:
                    self.wrapped_merged_auto_examples.append(
                        f"{self._sheet_name}!{m.ref}: {lines} lines in {rows} auto-fitted row(s): '{self._snippet(text)}'")
        return out

    def _wrap_decide(self, r, c, s, text, size, line_pt, block, lines, merge, formula, px):
        """Sure cut-off of a wrapped text: even the fewest lines (any model) hide more than
        WRAP_HIDDEN_LINES_TOL lines of the block; band when only some models do."""
        tc = self.tcounts
        shown = block / line_pt
        lo, hi = min(lines), max(lines)
        if lo - shown > WRAP_HIDDEN_LINES_TOL:
            n = shown_chars(text)
            if n < TEXT_CLIP_MIN_CHARS:
                # cut off but shorter than TEXT_CLIP_MIN_CHARS: counted, not a mistake (Patrick, option 3)
                tc["wrapped_cut_off_short"] += 1
                self._sheet_short[1] += 1
                if len(self.wrapped_short_examples) < MAX_TEXT_EXAMPLES:
                    self.wrapped_short_examples.append(
                        {"cell": f"{self._sheet_name}!{make_ref(r, c)}", "text": self._snippet(text), "chars": n,
                         "lines": lo, "rows_show": round(shown, 2)})
                return None
            tc["wrapped_cut_off"] += 1
            return {"r": r, "c": c, "s": s, "text": text, "size": size, "line_pt": line_pt, "block": block,
                    "lines": lo, "lines_by_model": dict(zip((m.name for m in self.models), lines)),
                    "shown": shown, "merge": merge, "formula": formula, "width_px": px[REFERENCE_MODEL],
                    "chars": n}
        if hi - shown > WRAP_HIDDEN_LINES_TOL:
            tc["wrapped_band"] += 1
        return None

    # ------------------------------------------------------------------ conditional formats
    def _cf_rules(self, tail) -> list:
        """Rules that can change a number's display width (set a number format) or stop lower ones."""
        out, order = [], 0
        for cf in tail.conditional_formats:
            if not cf.ranges:
                continue
            anchor = (cf.ranges[0][0], cf.ranges[0][1])
            for rule in cf.rules:
                order += 1
                if rule.type in ("dataBar", "iconSet", "colorScale"):
                    continue
                fmt = None
                d = rule.dxf if rule.dxf is not None else self.st.dxf(rule.dxf_id)
                if d is not None:
                    if d.num_fmt_code is not None:
                        fmt = d.num_fmt_code
                    elif d.num_fmt_id is not None:
                        fmt = N.resolve_format(d.num_fmt_id, self.st.num_fmts)
                if fmt is None and not rule.stop_if_true:
                    continue
                x = _CfRule()
                x.prio = rule.priority if rule.priority is not None else 10 ** 9
                x.order, x.rule, x.boxes, x.anchor, x.fmt, x.stop = order, rule, list(cf.ranges), anchor, fmt, bool(rule.stop_if_true)
                out.append(x)
        out.sort(key=lambda x: (x.prio, x.order))
        if not any(x.fmt is not None for x in out):
            return []
        return out

    @staticmethod
    def _in_boxes(rules, r, c) -> list:
        return [i for i, x in enumerate(rules) if any(b[0] <= r <= b[2] and b[1] <= c <= b[3] for b in x.boxes)]

    def _cf_open(self, rules, ids, v, r, c) -> list:
        """[(rule, why)] of the covering rules that cannot be evaluated for this cell."""
        out = []
        for i in ids:
            f = fires(rules[i].rule, v, (r, c) + rules[i].anchor, self._env)
            if is_unevaluable(f):
                out.append((rules[i], f.why))
        return out

    def _cf_formats(self, rules, ids, v, r, c, branch: bool = False):
        """Set of number formats (None = the cell's own) the covering rules can produce.  A rule that
        cannot be evaluated is OFF (Patrick 2026-10-05); branch=True: it may fire too (is the assumption
        decisive?).  None when more than 8 rules branch."""
        fired = [fires(rules[i].rule, v, (r, c) + rules[i].anchor, self._env) for i in ids]
        fired = [((UNKNOWN if branch else False) if is_unevaluable(f) else f) for f in fired]
        unknown = [k for k, f in enumerate(fired) if f is UNKNOWN]
        if len(unknown) > 8:
            return None
        out = set()
        for combo in itertools.product((True, False), repeat=len(unknown)):
            fl = list(fired)
            for k, i in enumerate(unknown):
                fl[i] = combo[k]
            fmt = None
            for i, f in zip(ids, fl):
                if not f:
                    continue
                x = rules[i]
                if x.fmt is not None and fmt is None:
                    fmt = x.fmt
                if x.stop:
                    break
            out.add(fmt)
        return out

    # ------------------------------------------------------------------ sheet end
    @staticmethod
    def _merge_lookup(merges, cells):
        """{(r, c): Merge} for the given cells that lie in a multi-cell merge."""
        if not merges or not cells:
            return {}
        rows = sorted({rc[0] for rc in cells})
        by_row = {}
        for rc in cells:
            by_row.setdefault(rc[0], []).append(rc[1])
        out = {}
        for m in merges:
            if m.is_single_cell:
                continue
            i, j = bisect_left(rows, m.r1), bisect_right(rows, m.r2)
            for k in range(i, j):
                r = rows[k]
                for c in by_row[r]:
                    if m.c1 <= c <= m.c2:
                        out[(r, c)] = m
        return out

    def sheet_end(self, head, tail):
        self._flush_row()
        name = head.name
        rules = self._cf_rules(tail)
        self._env = cf_env(self.wb, name)            # the rules' formulas read other cells through it
        if rules:
            prefetch(self._env, tail.conditional_formats)
        cands = self._cands
        merges = tail.merges
        # merged spans: covered cells are not displayed; an anchor has the whole span
        ml = self._merge_lookup(merges, [(r, c) for r, c, *_ in cands])
        final = []
        for cand in cands:
            r, c, s, v, t, text, ov96, note = cand
            m = ml.get((r, c))
            if m is not None:
                if (m.r1, m.c1) != (r, c):
                    continue
                pixels = self.geo.span_pixels(m.c1, m.c2)
                if pixels is None:
                    continue
                status, ov, text = self._measure(self._style(s), v, t, self._style(s).code, pixels)
                if status not in ("over", "uncertain"):
                    if status == "band":
                        self.n_band += 1
                    continue
                cand = (r, c, s, v, t, text, ov[REFERENCE_MODEL], note)
            final.append(cand)
        text_cut = self._settle_text(merges)         # cut-off text (a), merges known
        wrap_cut = self._settle_wrapped(merges)      # cut-off text (b)
        if self._uncertain:
            # an unverified rendering in a covered (non-anchor) merged cell is never displayed (as 69-F2)
            mu = self._merge_lookup(merges, [(r, c) for r, c, *_ in self._uncertain])
            self._uncertain = [x for x in self._uncertain
                               if mu.get((x[0], x[1])) is None or (mu[(x[0], x[1])].r1, mu[(x[0], x[1])].c1) == (x[0], x[1])]
        undecided = self._settle_undecided(merges, rules)
        hidden_anchors = self._hidden_anchors(merges)
        rec = {"sheet": name, "state": head.state, "cands": final, "uncertain": self._uncertain,
               "undecided": undecided, "counts": self._sheet_counts, "cf_second_pass": False,
               "cf_undecided": [], "text_cut": text_cut, "wrap_cut": wrap_cut, "short": self._sheet_short}
        self._sheets.append(rec)
        if rules or hidden_anchors:
            rec["rules"] = rules
            rec["hidden_anchors"] = hidden_anchors
            rec["merges"] = [m for m in merges if not m.is_single_cell]
            rec["hidden_rows"] = self._hidden_rows
            rec["row_info"] = (self._hidden_rows, self._custom_ht, self._zero_default, self._shown_rows)
            rec["geo"] = self.geo
            rec["env"] = self._env
            self._pass2[name] = rec
        if rules:
            # cells inside a format-setting rule's ranges are judged in the second pass under every
            # possible outcome; first-pass candidates there are dropped here
            rec["cands"] = [x for x in final if not self._in_boxes(rules, x[0], x[1])]
            rec["uncertain"] = [x for x in self._uncertain if not self._in_boxes(rules, x[0], x[1])]
            rec["cf_second_pass"] = True
            self.request_second_pass(name, None)
        elif hidden_anchors:
            self.request_second_pass(name, list(hidden_anchors))
        self._cands = []
        self._tcands = []
        self._wcands = []
        self._wauto = []
        self._hidden_rows = set()
        self._shown_rows = set()
        self._custom_ht = {}
        self._undec_sheet = []

    def _settle_undecided(self, merges, rules) -> list:
        """Untrusted formula results of this sheet -> [(r, c, ref)] still undecided.  Dropped (69-F2): a
        covered (non-anchor) cell of a merge, which is never displayed, and a cell whose own format
        prints nothing for every number (';;;') unless a format-setting conditional rule covers it."""
        und = self._undec_sheet
        if not und:
            return []
        ml = self._merge_lookup(merges, [(r, c) for r, c, _ref, _s in und])
        fmt_rules = [x for x in rules if x.fmt is not None]
        out = []
        for r, c, ref, s in und:
            m = ml.get((r, c))
            if m is not None and (m.r1, m.c1) != (r, c):
                self.counts["undecided_not_displayed"] += 1
                continue
            if self._style(s).always_blank and not self._in_boxes(fmt_rules, r, c):
                self.counts["undecided_not_displayed"] += 1
                continue
            out.append((r, c, ref))
        return out

    def _hidden_anchors(self, merges) -> dict:
        """{(r1, c1): Merge} for merges whose anchor sits in a hidden row or column while the span still
        shows (a visible column and a visible row): Excel draws the merged value across the visible part
        (69-F4).  Pass 1 skipped those anchors as hidden; pass 2 measures them against the visible span."""
        if not HIDDEN_ANCHOR_SPAN or not self._hidden_valued or not merges:
            return {}
        out = {}
        for m in merges:
            if m.is_single_cell:
                continue
            row_hidden = not self._span_has_visible_row(m.r1, m.r1)
            col_hidden = self.geo.col_pixels(m.c1) is None
            if not (row_hidden or col_hidden):
                continue
            if self.geo.span_pixels(m.c1, m.c2) is None or not self._span_has_visible_row(m.r1, m.r2):
                continue
            out[(m.r1, m.c1)] = m
        return out

    # ------------------------------------------------------------------ second pass (CF formats, hidden anchors)
    def second_pass_start(self, head):
        self._p2 = self._pass2.get(head.name)
        self._sheet_name = head.name
        if self._p2 is not None:
            self.geo = self._p2["geo"]
            self._env = self._p2.get("env")
            self._sheet_short = self._p2["short"]      # hidden text anchors count on their own sheet
            self._p2_anchor = {(m.r1, m.c1): m for m in self._p2["merges"]}
            self._p2_merges = self._p2["merges"]

    def second_pass_cell(self, cell):
        p = self._p2
        if p is None:
            return
        key = (cell.row, cell.col)
        hidden_anchor = p["hidden_anchors"].get(key)
        rules = p["rules"]
        if hidden_anchor is None and (not rules or cell.row in p["hidden_rows"]):
            return
        ids = self._in_boxes(rules, cell.row, cell.col) if rules else []
        if hidden_anchor is None and not ids:
            return
        # routed exactly as in pass 1: constants by stored type, formula results by trusted value (69-F1)
        kind, v, t = self._route(cell)
        if kind == "other":
            if hidden_anchor is not None:
                self._hidden_anchor_text(cell, v, hidden_anchor, p)
            return
        sty = self._style(cell.s)
        if sty.shrink or sty.skip:
            return
        if hidden_anchor is not None:
            # 69-F4: pass 1 skipped this anchor as hidden; Excel shows it across the visible span
            pixels = self.geo.span_pixels(hidden_anchor.c1, hidden_anchor.c2)
            if pixels is None:
                return
            self.counts["hidden_anchor_cells"] += 1
            if sty.fmt_err:
                self.note_default("unreadable_number_format", f"{location(cell.sheet, cell.ref)}: {sty.fmt_err}")
            if kind == "untrusted":
                if sty.always_blank and not self._in_boxes([x for x in rules if x.fmt is not None], *key):
                    self.counts["undecided_not_displayed"] += 1
                else:
                    p["undecided"].append((cell.row, cell.col, cell.ref))
                return
            if cell.is_formula_result:
                self.n_values_read += 1
        else:
            pixels = self.geo.col_pixels(cell.col)
            if pixels is None:
                return
            m = self._p2_anchor.get(key)
            if m is None:
                # a covered (non-anchor) cell of a merge is not displayed
                for mm in self._p2_merges:
                    if mm.r1 <= cell.row <= mm.r2 and mm.c1 <= cell.col <= mm.c2:
                        return
            else:
                pixels = self.geo.span_pixels(m.c1, m.c2)
                if pixels is None:
                    return
            if kind == "untrusted":
                return                                   # settled in pass 1 (sheet_end)
        if v is None or isinstance(v, bool) or (t != "d" and not isinstance(v, (int, float))):
            return
        if t != "d" and not math.isfinite(v):
            return
        if not ids:
            # a hidden anchor outside every format-setting rule: its own format decides
            self.counts["measured"] += 1
            status, ov, text = self._measure(sty, v, t, sty.code, pixels)
            self._settle(status, ov, text, cell.row, cell.col, cell.s, v, t, cell.ref, sty, p["cands"], p["uncertain"])
            return
        self.counts["cf_cells"] += 1
        why = []
        if t == "d":
            # an ISO-date cell (t="d") is not evaluated against the rules: they are OFF (Patrick 2026-10-05)
            fmts, opened = {None}, []
            alt_fmts = {None} | {rules[i].fmt for i in ids if rules[i].fmt is not None}
            if len(alt_fmts) > 1:
                why.append("rules over an ISO-date cell (t=\"d\") are not evaluated: assumed off")
        else:
            fmts = self._cf_formats(rules, ids, v, cell.row, cell.col)
            opened = self._cf_open(rules, ids, v, cell.row, cell.col)
            alt_fmts = self._cf_formats(rules, ids, v, cell.row, cell.col, branch=True) if opened else fmts
            if opened:
                why.append("rule(s) assumed off: " + "; ".join(f"{rule_text(x.rule)} ({w})" for x, w in opened))
                for x, w in opened:
                    self._note_unevaluable(cell.sheet, x, w)
        measured = {}

        def measure(fmt):
            if fmt not in measured:
                code = fmt if fmt is not None else sty.code
                status, ov, text = self._measure(sty, v, t, code, pixels)
                measured[fmt] = (status, ov, text, code)
            return measured[fmt]

        results = {fmt: measure(fmt) for fmt in (fmts if fmts is not None else {None})}
        decisive = False
        if fmts is None or len({x[0] == "over" for x in results.values()}) > 1 or \
                any(fmt is not None and x[0] == "uncertain" for fmt, x in results.items()):
            # the conditional formatting leaves the display open (a CF format Excel's reading of which is not
            # verified ...): graded by the cell's own format (Patrick 2026-10-05)
            why.append("its conditional-format display is not verified (" +
                       ", ".join(sorted(repr(x[3]) for f, x in results.items() if f is not None)) +
                       "): graded by its own format")
            decisive = True
            results = {None: measure(None)}
        if not decisive and why:
            over_now = {x[0] == "over" for x in results.values()}
            try:
                alt = [measure(f) for f in (alt_fmts if alt_fmts is not None else {None})]
                decisive = alt_fmts is None or {x[0] == "over" for x in alt} != over_now or \
                    any(x[0] == "uncertain" for x in alt)
            except GradingError:
                decisive = True
        if why:
            self._assume(cell.sheet, cell.ref, v, "; ".join(why), decisive)
        if any(x[0] == "uncertain" for x in results.values()):
            # the cell's OWN format's display is not verified (not a conditional-format matter): the engine's best
            # rendering is measured (Patrick 2026-10-05: every attempt graded)
            self.note_default("unverified_number_format", f"{location(cell.sheet, cell.ref)} = {v!r} under "
                                                          f"{sty.code!r}: best rendering measured")
        over = [x for x in results.values() if x[0] in ("over", "uncertain")]
        if over and len(over) == len(results):
            best = max(over, key=lambda x: x[1][REFERENCE_MODEL])
            note = f'conditional-format number format "{best[3]}"' if best[3] != sty.code else None
            p["cands"].append((cell.row, cell.col, cell.s, v, t, best[2], best[1][REFERENCE_MODEL], note))

    def second_pass_end(self, head, tail):
        p = self._p2
        if p is not None:
            p["geo"] = None
        self._p2 = None

    def _assume(self, sheet: str, ref: str, v, why: str, decisive: bool):
        """Record a cell measured under a conditional-format assumption (stats.cf_assumptions)."""
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

    # ------------------------------------------------------------------ verdict
    def finish(self) -> dict:
        n_sure = 0
        per_sheet = []
        undecided = []
        blocks = {"numbers": 0, "unwrapped_text": 0, "wrapped_text": 0}
        for rec in self._sheets:
            name = rec["sheet"]
            hid = " (hidden sheet)" if rec["state"] != "visible" else ""
            cands = rec["cands"]
            n_sure += len(cands)
            rec["counts"]["sure_overflows"] = len(cands)
            for r, c, ref in rec["undecided"]:
                undecided.append({"sheet": name, "cell": ref, "why": "formula value untrusted"})
            for r, c, ref, why in rec["uncertain"] + rec["cf_undecided"]:
                undecided.append({"sheet": name, "cell": ref, "why": why})
            if cands:
                by_cell = {(x[0], x[1]): x for x in cands}
                for box in group_cells(by_cell):
                    cells = [by_cell[(rr, cc)] for rr in range(box[0], box[2] + 1) for cc in range(box[1], box[3] + 1)
                             if (rr, cc) in by_cell]
                    worst = max(cells, key=lambda x: x[6])
                    r, c, s, v, t, text, ov96, note = worst
                    sty = self._style(s)
                    rng = range_to_str(*box)
                    n = len(cells)
                    cols = index_to_col(box[1]) if box[1] == box[3] else f"{index_to_col(box[1])}:{index_to_col(box[3])}"
                    font = f"{sty.face or 'default font'} {sty.size:g}{' bold' if sty.bold else ''}"
                    desc = (f"{n} numeric cell{'s' if n != 1 else ''} {rng} on sheet '{name}'{hid} "
                            f"display{'' if n != 1 else 's'} '####' in Excel: column {cols} is too narrow for the "
                            f"value{'s' if n != 1 else ''}; e.g. {make_ref(r, c)} shows '{text.strip()}' ({font}, "
                            f"number format \"{sty.code}\"{', ' + note if note else ''}), about {ov96:.0f} px wider "
                            f"than the column's text area at 96 dpi and overflowing under every platform model "
                            f"(96/120 dpi, rounded or fractional glyph widths).")
                    self.add_mistake(location(name, rng), desc)
                    blocks["numbers"] += 1
            if TEXT_CLIP and rec["text_cut"]:
                blocks["unwrapped_text"] += self._text_mistakes(name, hid, rec["text_cut"])
            if WRAPPED_TEXT_CLIP and rec["wrap_cut"]:
                blocks["wrapped_text"] += self._wrap_mistakes(name, hid, rec["wrap_cut"])
            per_sheet.append({"sheet": name, **rec["counts"], "undecided": len(rec["undecided"]) + len(rec["uncertain"])
                              + len(rec["cf_undecided"]), "cf_second_pass": rec["cf_second_pass"],
                              "text_cut_off": len(rec["text_cut"]), "wrapped_text_cut_off": len(rec["wrap_cut"]),
                              "text_cut_off_short": rec["short"][0], "wrapped_text_cut_off_short": rec["short"][1]})
        self.counts["sure_overflows"] = n_sure
        for u in undecided:
            # Patrick 2026-10-05 (every attempt graded): an untrusted formula value is skipped, never raised
            self.note_default("untrusted_value", f"{location(u['sheet'], u['cell'])}: {u['why']}")
        stats = dict(self.counts)
        stats.update(self.tcounts)
        stats.update({
            "n_sheets_checked": len(self._sheets),
            "normal_font": f"{self.normal_face} {self.normal_size:g}",
            "normal_font_known": face_key(self.normal_face)[1],
            "column_unit_px": {m.name: m.mdw for m in self.models},
            "sure_tolerance_px": SURE_TOL_PX,
            "band_cells": self.n_band, "band_examples": self.band,
            "formula_values_read": self.n_values_read,
            "undecided_cells": len(undecided), "undecided_examples": undecided[:MAX_UNDECIDED_LISTED],
            "unknown_faces": self.unknown_faces,
            "text_unknown_faces": self.text_unknown_faces,
            "mistakes_by_rule": blocks,
            "text_blockers": dict(sorted(self.text_blockers.items(), key=lambda kv: -kv[1])),
            "text_blockers_short": dict(sorted(self.text_blockers_short.items(), key=lambda kv: -kv[1])),
            "text_cut_off_short_examples": self.text_short_examples,
            "wrapped_cut_off_short_examples": self.wrapped_short_examples,
            "text_band_examples": self.text_band_examples,
            "text_rows_under_half_line_examples": self.text_short_row_examples,
            "wrapped_merged_in_auto_rows_examples": self.wrapped_merged_auto_examples,
            "second_pass_sheets": sorted(self._pass2),
            # Patrick 2026-10-05: conditional formats never stop a grading - every assumption is recorded
            "cf_assumptions": {
                "policy": "a conditional-format rule that cannot be evaluated is OFF, and a conditional-format "
                          "number format whose display is not verified gives way to the cell's own format "
                          "(Patrick 2026-10-05)",
                "cells": self.cf_assumed, "decisive_cells": self.cf_assumed_decisive,
                "examples": self.cf_assumed_examples,
                "unevaluable_rules": len(self.cf_unevaluable),
                "unevaluable_rule_examples": list(self.cf_unevaluable.values())[:MAX_ASSUMPTION_EXAMPLES]},
            "per_sheet": per_sheet,
            "options": {"negative_date_fails": NEGATIVE_DATE_FAILS, "general_shortens": GENERAL_SHORTENS,
                        "untrusted_only_if_verdict_needs": UNTRUSTED_ONLY_IF_VERDICT_NEEDS,
                        "hidden_anchor_span": HIDDEN_ANCHOR_SPAN,
                        "text_clip": TEXT_CLIP, "wrapped_text_clip": WRAPPED_TEXT_CLIP,
                        "text_clip_min_chars": TEXT_CLIP_MIN_CHARS, "max_font_pt": MAX_FONT_PT,
                        "wrap_hidden_lines_tol": WRAP_HIDDEN_LINES_TOL, "line_height_per_pt": LINE_HEIGHT_PER_PT,
                        "merges_block_overflow": MERGES_BLOCK_OVERFLOW,
                        "hidden_cells_block_overflow": HIDDEN_CELLS_BLOCK_OVERFLOW,
                        "models": [m[0] for m in MODELS]},
        })
        und_note = ""
        if undecided:
            und_note = (f" {len(undecided)} cell(s) with an untrusted formula value were skipped (stats.defaults).")
        flagged = ", ".join(f"'{p['sheet']}'" for p in per_sheet
                            if p["sure_overflows"] or (TEXT_CLIP and p["text_cut_off"])
                            or (WRAPPED_TEXT_CLIP and p["wrapped_text_cut_off"]))
        n_short = self.tcounts["text_cut_off_short"] + self.tcounts["wrapped_cut_off_short"]
        short_note = (f" {n_short} cut-off text(s) shorter than {TEXT_CLIP_MIN_CHARS} characters are counted in stats "
                      f"only." if n_short else "")
        return self.verdict(
            f"No value is wider than its column and no long text is cut off: {self.counts['numeric_cells']} numeric "
            f"cell(s) and {self.tcounts['text_cells']} text cell(s) on {len(self._sheets)} sheet(s); nothing displays "
            f"'####' and no text of {TEXT_CLIP_MIN_CHARS} or more characters is cut off under every platform "
            f"model.{short_note}",
            f"{{n}} block(s) of values shown as '####' or text of {TEXT_CLIP_MIN_CHARS}+ characters cut off on: "
            f"{flagged}.{und_note}",
            stats)

    # ------------------------------------------------------------------ cut-off text: mistakes
    _ALIGN_WORDS = {"spill_right": "left-aligned", "spill_left": "right-aligned", "spill_both": "centred",
                    "across": "centred across selection", "fill": "fill-aligned"}

    def _font_text(self, sty: _Style) -> str:
        return f"{sty.face or 'default font'} {sty.size:g}{' bold' if sty.bold else ''}"

    def _hidden_chars(self, x) -> tuple:
        """(hidden characters, visible-part characters) of a cut-off text under the reference model: a
        character counts as hidden when any part of its glyph lies outside the space the text may use."""
        sty = self._style(x["s"])
        disp, ref, kind = x["disp"], x["ref"], x["kind"]
        m = self.ref_model
        w = self._char_widths(x["segs"], m)
        n = len(disp)
        i0, j0 = n - len(disp.lstrip()), len(disp.rstrip())
        total = sum(w)
        ind = self._indent_px(sty, m)
        area = ref["area"]
        if kind in ("spill_right", "fill"):
            x0 = ind
        elif kind == "spill_left":
            x0 = area - ind - total
        else:
            x0 = area / 2.0 - total / 2.0
        lo, hi = -ref["room_l"] - 1e-6, area + ref["room_r"] + 1e-6
        hidden, pos = 0, x0
        for k in range(n):
            if i0 <= k < j0 and (pos < lo or pos + w[k] > hi):
                hidden += 1
            pos += w[k]
        return hidden, j0 - i0

    def _text_mistakes(self, name: str, hid: str, cuts: list) -> int:
        """One mistake per rectangle of adjacent cut-off texts (unwrapped), naming its worst cell."""
        by_cell = {(x["r"], x["c"]): x for x in cuts}
        boxes = group_cells(by_cell)
        for box in boxes:
            cells = [by_cell[(rr, cc)] for rr in range(box[0], box[2] + 1) for cc in range(box[1], box[3] + 1)
                     if (rr, cc) in by_cell]
            x = max(cells, key=lambda y: y["over"])
            sty = self._style(x["s"])
            rng = range_to_str(*box)
            n = len(cells)
            phrases = []
            for _side, _why, phrase in x["blockers"]:
                if phrase not in phrases:
                    phrases.append(phrase)
            hc, nv = self._hidden_chars(x)
            hp = sum(d for d in (x["ref"]["dr"], x["ref"]["dl"]) if d is not None and d > 0)
            what = "a text formula result" if x["formula"] else "text"
            lead = (f"Text in {rng} on sheet '{name}'{hid} is cut off: " if n == 1 else
                    f"{n} text cells {rng} on sheet '{name}'{hid} are cut off; e.g. ")
            self.add_mistake(location(name, rng), lead + (
                f"{make_ref(x['r'], x['c'])} '{self._snippet(x['disp'])}' ({what}, {self._font_text(sty)}, "
                f"{self._ALIGN_WORDS[x['kind']]}) is cut off by {' and '.join(phrases)}: about {hc} of its {nv} "
                f"characters ({hp:.0f} px at 96 dpi) cannot be seen under any platform model (96/120 dpi, rounded or "
                f"fractional glyph widths)."))
        return len(boxes)

    def _wrap_mistakes(self, name: str, hid: str, cuts: list) -> int:
        """One mistake per rectangle of adjacent cut-off wrapped texts, naming its worst cell."""
        by_cell = {(x["r"], x["c"]): x for x in cuts}
        boxes = group_cells(by_cell)
        for box in boxes:
            cells = [by_cell[(rr, cc)] for rr in range(box[0], box[2] + 1) for cc in range(box[1], box[3] + 1)
                     if (rr, cc) in by_cell]
            x = max(cells, key=lambda y: y["lines"] - y["shown"])
            sty = self._style(x["s"])
            rng = range_to_str(*box)
            n = len(cells)
            m = x["merge"]
            chars = (x["width_px"] - CELL_MARGIN_PX) / self.ref_model.mdw
            where = (f"its merged range {m.ref} ({chars:.2f} characters wide)" if m is not None else
                     f"its {chars:.2f}-character column {index_to_col(x['c'])}")
            if m is not None and m.r2 > m.r1:
                rows = f"rows {m.r1}:{m.r2} of the merge have custom heights totalling {x['block']:g} pt"
            else:
                rows = f"row {x['r']} has a custom height of {x['block']:g} pt"
            what = "a wrapped text formula result" if x["formula"] else "wrapped text"
            lead = (f"Wrapped text in {rng} on sheet '{name}'{hid} is cut off at the bottom: " if n == 1 else
                    f"{n} wrapped text cells {rng} on sheet '{name}'{hid} are cut off at the bottom; e.g. ")
            self.add_mistake(location(name, rng), lead + (
                f"{make_ref(x['r'], x['c'])} '{self._snippet(x['text'])}' ({what}, {self._font_text(sty)}) needs "
                f"{x['lines']} lines in {where}, but {rows}: room for about {x['shown']:.1f} lines of "
                f"{x['line_pt']:.1f} pt, so about {x['lines'] - x['shown']:.1f} line(s) cannot be seen under any "
                f"platform model (96/120 dpi, rounded or fractional glyph widths)."))
        return len(boxes)
