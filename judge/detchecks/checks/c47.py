"""47 Formatting/No bright-yellow highlighting.

Rule (handoff bucket 2, Patrick 2026-10-02; full write-up in docs/checks/47.md):

Fail if any position of the delivered workbook is painted BRIGHT yellow, through
  - a cell's own fill (every <c> element: values, formulas, spill members, empty styled cells),
  - a row style (<row s=.. customFormat=1>: the row's positions without a <c>),
  - a column style (<col style=..>: the column's positions without a <c> or row style),
  - the default style 0 (every empty position of every sheet),
  - a conditional-format rule whose dxf fill is bright yellow (main and x14 rules), counted
    ALWAYS - whether or not it fires today (the rubric treats a CF on error-check rows as a
    yellow usage that needs documenting),
unless the workbook's own text explains yellow as an ongoing convention (a legend line, see
LEGEND below).  Even then, a bright-yellow cell whose own text or comment is a placeholder
marker (TBD, to do, ...) still fails (WIP_CONTENT_OVERRIDES_LEGEND).  A conditional-format fill
whose colour cannot be resolved is not counted (Patrick 2026-10-05: "CONDITIONAL FORMATS never stop
a grading" - the rule is taken as off for this check; recorded in stats.cf_assumptions); an
unresolvable fill colour of a cell, row, column or default style is read as no fill (Patrick 2026-10-05:
every attempt graded; stats.defaults.unresolved_colour).

Bright yellow = HSV hue in [47.5, 68] degrees, saturation >= 0.50, value >= 0.85 (FFFF00,
FFFF66, FFD700, FFCC00 ...).  Pale yellows (FFFFCC, FFF2CC, FFFF99, FFEB9C, FFFFE0) and
Excel's orange/gold FFC000 are NOT bright.  Colours are resolved by the core (theme + tint,
indexed palette incl. custom <indexedColors>); solid fills use fgColor, hatch patterns their
fg-over-bg blend, gradients count when ANY stop is bright yellow.  Fonts, sheet tabs, colour
scales, data bars and icon sets never count.  A fill colour the core cannot resolve (invalid rgb,
theme / indexed index beyond the palette ...: an UnknownColour, 47-R2-08) is read as Excel's default
for a fill, none (Patrick 2026-10-05: every attempt graded; until then every reading of it was graded
and the check raised when they disagreed).

LEGEND (the exception).  Text sources: constant text cells (shared / inline strings), cell
comments (legacy and threaded) and drawing text boxes, on every sheet.  A text "mentions"
yellow when it says yellow / FFFF00 without a pale qualifier ("light yellow", "pale yellow"
explain pale fills, not bright yellow).  Each clause that mentions yellow is classified, in order:
  NEGATED   "no yellow is used", "yellow: none", "yellow highlighting removed / replaced / cleaned
            up" (not: a scoped "none on output sheets", "except for the input cells", or a
            negation that qualifies the subject: "cells with no formula are shaded yellow")
  WIP       defines yellow as unfinished: TBD, TBC, to do, placeholder, WIP, draft, pending,
            to be confirmed/determined, not yet, still to, unconfirmed, dummy, preliminary, fill
            in later ...; also soft markers ("to be reviewed / needs review / to confirm / needs
            input / missing / follow-up") unless qualified as periodic ("reviewed quarterly" is
            the rubric's ongoing use) or conditional ("turns yellow when data is missing")
  DEFINES   any other definition: "Yellow fill = key outputs", "Yellow cells are inputs",
            "Inputs are (highlighted) in yellow", "the yellow input cells", "Error-check rows
            turn yellow when a check fails", "Enter assumptions in the yellow cells"
  MENTION   yellow appears but nothing is defined ("the logo is yellow", "Yellow (FFFF00)")
A short colour label ("Yellow", "Yellow fill:", "Yellow Fill (#FFFF00)") takes its definition
from the next text cells to its right on the same row ("Yellow | Key outputs"); a text left of
it is classified too, but only to restrict ("Review / Unfinished | Yellow | ..." is WIP); with
nothing on the right the left text is the meaning ("Inputs | Yellow"), a definition only when it
names cells ("Cost schedule | Yellow" in a RAG status column defines nothing).
One DEFINES line anywhere excuses the workbook's bright yellow (except WIP-content cells),
unless some text defines yellow as unfinished (WIP_LEGEND_BLOCKS_EXCUSE).
A legend SWATCH (a bright-yellow sample cell, or a sample merged across 2-3 empty cells) is
documentation, not highlighting, and is not an offender when no bright-yellow cell touches it
horizontally and it has a label:
  - a text 1-2 columns to its right (first non-empty cell) or within 3 columns to its left
    that STARTS with the colour word ("Yellow fill = key outputs", "Yellow | Unfinished - none
    in this file | [sample]", "Yellow fill | [sample] | Unfinished / review (none ...)" - the
    colour-word label wins over a plain text on the other side); or the sample's own text does
    (a bare colour label, or a DEFINES / NEGATED line - not a yellow note such as "Yellow flag:
    rate TBD"): the swatch is documentation and that text, classified as above, decides whether
    yellow is excused;
  - or a label WITHOUT the colour word, when the cell before the line holds no data (a legend
    header is allowed), the label is not WIP / negated and no number / formula sits in the 3
    columns after the label (after the sample, for a left label or a self-labelled sample):
      right label ("[sample] | Key output cells"): a legend / colour key / conventions header
        up to 15 rows above within 3 columns, or a colour key block (another colour's sample
        within 3 rows in the same column whose label names a cell category), or the label
        names a cell category, the cell right of it is empty and it does not head a table
        (no label / value row in the 2 rows below: "[yellow] | Inputs" over "Growth | 5%" is a
        section header with a marker);
      left label ("Key outputs | [sample]") or the sample's own text ([yellow 'Key output']):
        the label must name a cell category (inputs, assumptions, outputs, formulas, links,
        checks, hard-coded ...), plus a legend header or key block (or, for a left label,
        nothing after the sample and no table below).
    Such a label counts as a DEFINES line.
Every other bright-yellow cell is an ordinary offender.  A CF rule's fill is what Excel paints:
bgColor only (fgColor ignored; bgColor indexed 64 / 00000000 = black; no bgColor = no fill;
measured in Excel 2026-10-03, see _fills.py).
A text cell that cannot be decoded (shared-string index out of range) makes the check raise
only when the workbook has a bright-yellow source (the text could be its legend); otherwise the
verdict does not depend on it.
"""
from __future__ import annotations

import colorsys
import re
from collections import defaultdict
from typing import Optional

from ..core.refs import index_to_col, location, range_to_str, split_ref
from ..core.styles import FillPaint, is_unknown
from ..errors import GradingError
from ._fills import dxf_fill_paints, known_paint, paint_unknowns as _paint_unknowns
from .base import Check

# ---------------------------------------------------------------------------- tunables
HUE_MIN, HUE_MAX = 47.5, 68.0      # degrees (HSV); FFFF00 = 60, FFCC00 = 48.0, FFC000 = 45.2 (out)
SAT_MIN = 0.50                     # FFFF99 (0.40), FFEB9C (0.39), FFFFCC (0.20) are pale, not bright
VAL_MIN = 0.85                     # excludes olive / dark-yellow (CCCC00 = 0.80, 808000)
CF_COUNTS_ALWAYS = True            # a bright-yellow CF rule counts whether or not it fires today
WIP_CONTENT_OVERRIDES_LEGEND = True  # a yellow cell saying 'TBD' fails even under a positive legend
WIP_LEGEND_BLOCKS_EXCUSE = True    # a text that defines yellow as unfinished (WIP) anywhere: no line excuses yellow
TEMPLATE_LINES_EXCUSE = False      # a line about the case template's retained (pale) yellow does not excuse bright yellow
SWATCH_LABEL_MAX_DIST_RIGHT = 2    # label column distance for a swatch (label to the right)
SWATCH_LABEL_MAX_DIST_LEFT = 3     # colour-word label to the left ('Yellow | meaning | [sample]')
SWATCH_PLAIN_LEFT_DIST = 2         # key label WITHOUT the colour word to the left ('Key outputs | [sample]')
LEGEND_HEADER_ROWS_ABOVE = 15
LEGEND_HEADER_COL_DIST = 3
KEY_BLOCK_ROW_DIST = 3
KEY_LABEL_MAX_LEN = 60             # a colour-key label naming a category of cells is short
KEY_LINE_CLEAR_COLS = 3            # a legend line has no number / formula in the 3 columns right of its label
TABLE_ROWS_BELOW = 2               # a key label with label / value rows in the 2 rows below heads a table (47-R2-04)
MERGED_SWATCH_MAX_CELLS = 3        # a legend sample merged across up to 3 cells is still one sample (47-R2-10)
MAX_LINES_IN_STATS = 12

# ---------------------------------------------------------------------------- colour band


def hsv(argb: str) -> tuple[float, float, float]:
    h6 = argb[-6:]
    r, g, b = (int(h6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return h * 360.0, s, v


def is_bright_yellow(argb: Optional[str]) -> bool:
    """'FFRRGGBB' / 'RRGGBB' -> True when the colour is in the bright-yellow band."""
    if not argb or len(argb) < 6:
        return False
    try:
        h, s, v = hsv(argb)
    except ValueError:
        return False
    return HUE_MIN <= h <= HUE_MAX and s >= SAT_MIN and v >= VAL_MIN


# ---------------------------------------------------------------------------- legend text
_MENTION_RX = re.compile(
    r"(?<![\w#])(?:(?P<q>light|pale|soft|pastel|faint|muted|creamy?)[\s-]*)?(?:bright[\s-]*)?"
    r"yellow(?:s|ish)?(?![\w])|#?(?<![0-9A-Fa-f])(?:FF)?FFFF00(?![0-9A-Fa-f])", re.I)   # FFFF00 or ARGB FFFFFF00


def mentions_yellow(text: str) -> bool:
    """True when the text names (bright) yellow; 'light/pale yellow' does not count."""
    for m in _MENTION_RX.finditer(text or ""):
        if not m.group("q"):
            return True
    return False


_YW = r"(?:bright[\s-]*)?(?:yellow|#?(?:ff)?ffff00)"
_YNOUN = (r"(?:fill(?:ed|s)?|cells?|highlight(?:s|ed|ing)?|shad(?:ed|ing|e)|background|box(?:es)?|colou?r(?:ed)?|marking|"
          r"flag(?:s|ged)?|rows?|columns?)")
# a colour code written after the colour word: 'Yellow (FFFF00)', 'Yellow Fill (#FFFF00)', 'Yellow - #FFFF00'
_HEXCODE = (r"(?:\s*\(\s*(?:#|rgb\s*:?\s*|hex\s*:?\s*)?(?:[0-9a-f]{6}|[0-9a-f]{8}|\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3})\s*\)"
            r"|\s*[:=–—-]?\s*#?(?:ff)?ffff00\b)")
# words that name the paint itself, not a meaning ('Highlight colour: FFFF00', 'Fill (#FFFF00)')
_PAINTWORD = r"(?:fill|fills|filled|colou?rs?|highlight\w*|shading|background|code|hex|rgb|swatch|sample|yellow|bright)"
_EXC = r"(?:except|excluding|outside|other\s+than|besides|apart\s+from|but|beyond|only|unless|instead\s+of|rather\s+than)"
# a category of cells, as colour keys name them
_CAT = (r"(?:inputs?|assumptions?|outputs?|drivers?|(?:error[\s-]?)?checks?|hard[\s-]?cod\w*|formulas?|links?|results?|"
        r"answers?|constants?|scenarios?|switch(?:es)?|parameters?|estimates?|forecasts?|actuals?|historicals?|"
        r"overrides?|sources?|key\s+[\w-]+)")
# 'except for the input cells', 'only for inputs': yellow is used for that category
_EXCEPT_FOR_RX = re.compile(r"\b(?:except|other\s+than|apart\s+from|besides|save\s+for|but\s+for|only)\s+(?:for|on|in|as|to)?\s*"
                            r"(?:the\s+|our\s+|its\s+)?(?:[\w-]+\s+){0,2}?" + _CAT + r"\b", re.I)
# a scope that is only part of the workbook: 'none on output sheets', 'none in the Summary tab'
_PART_SCOPE_RX = re.compile(r"^\s+(?:on|in)\s+(?:the\s+)?(?!this\b|delivered\b|final\b|whole\b|entire\b|any\b|all\b)"
                            r"(?:[\w&-]+\s+){0,2}(?:sheets?|tabs?|worksheets?|pages?|sections?|schedules?)\b", re.I)

_NEG_RXS = [
    re.compile(r"\b(?P<neg>no|zero|without|nothing|none\s+of\s+the)\b(?P<mid>[^.;:=]{0,40}?)" + _YW + r"\b", re.I),
    re.compile(r"\bnot\s+(?:\w+\s+){0,2}?(?:in\s+)?" + _YW + r"\b", re.I),
    re.compile(_YW + r"\b[^.;]{0,60}?\b(?:(?:is|are|was|were|has\s+been|have\s+been)\s+)?(?:not\s+used|unused|not\s+present|"
               r"nowhere|removed|not\s+applied|absent|does\s+not\s+appear|doesn't\s+appear|do\s+not\s+appear|never\s+used|"
               r"no\s+longer|n/a)\b", re.I),
    re.compile(_YW + r"\b[^.;]{0,60}?\bnone\b", re.I),
    re.compile(r"\b(?:removed|cleared|deleted|eliminated|stripped)\b[^.;]{0,30}" + _YW + r"\b", re.I),
    # the yellow PAINT was replaced / cleaned up: 'Completed answers replace the template's yellow input fill',
    # 'Yellow source input fills are replaced with pale blue', 'All yellow flags were cleaned up'
    re.compile(r"\b(?:replac(?:e|es|ed|ing)|swap(?:s|ped)?|overr(?:ide|ides|idden|ode))\b[^.;]{0,40}?" + _YW +
               r"(?:[\s-]+[\w'-]+){0,3}?[\s-]+(?:fill|fills|highlight\w*|shading|colou?r(?:ing)?|background|formatting)\b", re.I),
    re.compile(_YW + r"\b[^.;]{0,40}?\b(?:fill|fills|highlight\w*|shading|colou?r(?:ing)?|background|formatting|flags?|"
               r"cells?|markers?)\b[^.;]{0,30}?\b(?:(?:is|are|was|were|has\s+been|have\s+been)\s+)?(?:replaced|cleaned(?:\s+up)?|"
               r"cleared|resolved|reverted|dropped|retired|swapped\s+out)\b", re.I),
    # a clause that opens with a negation over a list: 'No external links, macros, hidden items or unfinished yellow
    # cells' (any distance; same exception handling as the first pattern)
    re.compile(r"^\W*(?P<neg>no|there\s+(?:are|is|were|was)\s+no|nothing|without)\b(?P<mid>[^.;:=]*?)" + _YW + r"\b", re.I),
]


def _negated(cl: str) -> bool:
    for i, rx in enumerate(_NEG_RXS):
        for m in rx.finditer(cl):
            rest = cl[m.end():]
            if i in (0, 2, 7) and _EXCEPT_FOR_RX.search(rest):
                continue      # 'No yellow is used except for the input cells' scopes yellow, it does not negate it
            if i == 7 and re.search(r"\b" + _EXC + r"\b", m.group("mid"), re.I):
                continue
            if i == 0:
                if re.search(r"\b" + _EXC + r"\b", m.group("mid"), re.I):
                    continue  # 'no hard-coded numbers outside the yellow cells' negates hard-codes, not yellow
                modifier = m.group("neg").lower() == "without" or re.search(r"\bwith\s*$", cl[:m.start()], re.I)
                if modifier and re.search(r"\)|\b(?:are|is|were|was|shaded|highlighted|filled|colou?red|marked|shown)\b",
                                          m.group("mid"), re.I):
                    continue  # 'Cells with no formula are shaded yellow': 'no' qualifies the subject, not yellow
            if i == 3 and _PART_SCOPE_RX.match(cl[m.end():]):
                continue      # 'Yellow fill = inputs (none on output sheets)': none on PART of the workbook
            return True
    return False


# 'WIP' the marker, not 'WIP inventory' / 'WIP balance' (work-in-progress is a balance-sheet line, 47-R2-07)
_WIP_WORD = r"wip(?![\s-]+(?:inventor(?:y|ies)|stock|balances?|accounts?|goods|assets?|reserves?|provisions?))"
_WIP_HARD_RX = re.compile(
    r"\b(?:tbd|tbc|tba|to[\s-]?dos?|placeholders?|" + _WIP_WORD + r"|work[\s-]in[\s-]progress|unfinished|incomplete|drafts?|"
    r"not\s+(?:yet\s+)?final(?:i[sz]ed)?|to\s+be\s+(?:confirmed|determined|completed|finali[sz]ed|filled(?:\s+in)?|provided|"
    r"added|decided|agreed|defined|sourced|replaced)|pending|outstanding|open\s+(?:items?|points?|questions?)|"
    r"temporary|provisional|awaiting|fix\s?me|x{3,}|\?{2,}|"
    # 47-wip-vocab (review 2026-10-03): more ways of saying 'not done yet'
    r"to\s+(?:complete|finali[sz]e|be\s+finished)|still\s+to|yet\s+to|not\s+yet|"
    # ('unaudited' is a data-quality label like 'estimates' / 'actuals', not unfinished work: 47-R2-07)
    r"un(?:confirmed|verified|validated|finished|resolved|decided|checked|reviewed|sourced)|dummy|preliminary|"
    r"(?:fill(?:ed)?|complete[ds]?|updated?|confirm(?:ed)?|add(?:ed)?|finali[sz]ed?|enter(?:ed)?|check(?:ed)?|revisit(?:ed)?)"
    r"\s+(?:in\s+)?later)(?![\w])", re.I)
_WIP_SOFT_RX = re.compile(
    r"\b(?:to\s+be\s+(?:reviewed|updated|checked|verified|revisited|validated)|needs?\s+(?:a\s+)?(?:review|reviewing|check(?:ing)?|"
    r"updat(?:e|ing)|confirm(?:ation|ing)?|verif(?:y|ication))|review\s+me|check\s+me|for\s+review|requires?\s+(?:review|"
    r"confirmation|checking|verification)|under\s+review|to\s+check|to\s+verify|in\s+progress|"
    r"(?:update|replace|remove|delete|fill)\w*\s+(?:in\s+)?(?:before|prior\s+to)|"
    # 47-wip-vocab: soft (an ongoing convention when periodic or conditional)
    r"to\s+(?:confirm|update|revisit|review|validate|follow\s+up)|revisit|follow[\s-]?ups?|"
    r"needs?\s+(?:more\s+|further\s+|user\s+)?(?:input|inputs|data|values?|numbers?|figures?|filling|completing|completion)|"
    r"(?:inputs?|data|values?|figures?|numbers?)\s+(?:needed|required|missing|outstanding)|missing|"
    # 'replace with actuals' is an instruction (unfinished); 'were replaced with actuals' describes done work (47-R2-11)
    r"(?<!were\s)(?<!was\s)(?<!been\s)(?<!already\s)replac\w*\s+(?:with|by)\s+actuals?)\b", re.I)
_PERIODIC_RX = re.compile(
    r"\b(?:periodic(?:ally)?|annual(?:ly)?|yearly|quarterly|monthly|weekly|regular(?:ly)?|ongoing|recurring|"
    r"(?:each|every)\s+(?:year|quarter|month|period|cycle|close|forecast|reporting\s+period))\b", re.I)
# a soft marker inside a condition describes an ongoing check, not unfinished work: 'Error checks turn
# yellow when data is missing', 'Yellow = needs review if the balance check fails'
_CONDITIONAL_RX = re.compile(r"\b(?:when|whenever|if|once|in\s+case)\b", re.I)
# 'nothing here is TBD' denies WIP; it does not define yellow as WIP
_NEG_WIP_RX = re.compile(r"\b(?:no|none|nothing|not|never|zero)\b[^.;]{0,30}?\b(?:tbd|tbc|placeholders?|wip|unfinished|"
                         r"incomplete|drafts?|pending|to[\s-]?dos?)\b", re.I)
# a cross-reference does not define yellow: 'Yellow fill = inputs (see Checks tab for open items)',
# 'Yellow = inputs - see the Checks tab for open items' (47-R2-07)
_XREF_RX = re.compile(r"\(\s*(?:see|refer(?:\s+to)?|cf\.?|as\s+per|per|details?(?:\s+in)?)\b[^)]*\)|"
                      r"(?:[,–—]|\s-)\s*(?:see|refer\s+to|cf\.?)\b[^.;]*", re.I)


# Patrick 2026-10-06 (attempt 4026, Checks!A28 'REVIEW flags (yellow) mark sense-band breaches ... -
# "review me", nothing unfinished'; rubric: yellow for "flagging assumptions that require periodic
# review" is an allowed, documented purpose): a soft marker that is the QUOTED name of what the colour
# shows ('"review me" flag'), or that sits in a clause describing a flag / check convention, is an
# ongoing convention, not unfinished work.  Hard markers (TBD, unfinished, placeholder) are unchanged.
_QUOTED_RX = re.compile(r"\"[^\"]*\"|“[^”]*”|'[^']{1,40}'")
_FLAG_RX = re.compile(r"\b(?:flags?|flagged|flagging|breach\w*|alerts?|warnings?|exceptions?|sense[\s-]?bands?|"
                      r"tolerance|threshold\w*|error[\s-]?checks?)\b", re.I)


def _wip(cl: str) -> bool:
    c2 = _NEG_WIP_RX.sub(" ", _XREF_RX.sub(" ", cl))
    if _WIP_HARD_RX.search(c2):
        return True
    c3 = _QUOTED_RX.sub(" ", c2)
    return bool(_WIP_SOFT_RX.search(c3)) and not _PERIODIC_RX.search(c3) and not _CONDITIONAL_RX.search(c3) \
        and not _FLAG_RX.search(c3)


_SEP = r"\s*(?:=|:|–|—|->|→|\s-\s)\s*"
# after 'yellow cells are ...': words that describe the text itself or the fill's status, not a meaning
# ('its yellow fill is explained above' (GPT-6 2356), 'yellow cells are not inputs': 47-R2-03)
_META = (r"(?!(?:explained|described|documented|noted|covered|discussed|detailed|addressed|mentioned|listed|set\s+out|"
         r"outlined|summari[sz]ed|given|provided|intentional(?:ly)?|deliberate(?:ly)?|not)\b)")
# a noun that names cells, as a colour key's left-hand side does ('Input cells = yellow', 'Scenario selector: yellow')
_KEYNOUN = (r"(?:cells?|fields?|boxes|ranges?|rows?|columns?|values?|figures?|numbers?|items?|lines?|entries|tabs?|sheets?|"
            r"selectors?|toggles?|flags?|markers?|totals?|headers?|headings?|labels?)")
_DEF_RXS = [
    # 'Yellow fill = key outputs', 'Yellow: inputs', 'Yellow cells - assumptions'
    re.compile(r"\b" + _YW + r"(?:[\s-]+" + _YNOUN + r"|" + _HEXCODE + r")*" + _SEP + r"(?!#?(?:ff)?ffff00\b)[\w(\"']",
               re.I),
    # 'yellow-highlighted cells indicate ...', 'Bright yellow fill (#FFFF00) marks error-check rows'; only cell /
    # fill nouns (or a colour code) between the colour and the verb ('The Yellow River project indicates' is not one)
    re.compile(r"\b" + _YW + r"(?:[\s-]+" + _YNOUN + r"|" + _HEXCODE + r")*\s+(?:indicates?|denotes?|marks?|means?|"
               r"represents?|signals?|signif(?:y|ies)|identif(?:y|ies)|flags?|highlights?|designates?)\s+\w", re.I),
    # 'Yellow cells are inputs', 'Yellow fill is used for checks' (a cell/fill noun is required before are/is,
    # so 'yellow and green are the brand colours' is not a definition)
    re.compile(r"\b" + _YW + r"(?:[\s-]+" + _YNOUN + r")+(?:\s+[^\s.;:=]+){0,2}?\s+(?:are|is|contain|contains|hold|holds|"
               r"show|shows)\s+" + _META + r"\w", re.I),
    # 'Cells in yellow are inputs', 'cells shaded yellow contain hard-coded data'
    re.compile(r"\bcells?\s+(?:\w+\s+){0,2}?(?:in\s+)?" + _YW + r"(?:[\s-]+" + _YNOUN + r")?\s+(?:are|is|indicates?|"
               r"denotes?|marks?|means?|represents?|contain|contains|hold|holds|show|shows)\s+" + _META + r"\w", re.I),
    # 'Inputs are highlighted in yellow', 'key outputs shaded yellow'
    re.compile(r"\b(?:highlighted|shaded|filled|marked|colou?red|formatted|shown|flagged|identified|painted)\b(?:\s+\w+){0,2}?"
               r"\s+(?:in\s+|with\s+)?(?:a\s+)?" + _YW + r"\b", re.I),
    # 'the yellow input cells', 'yellow assumption cells', 'its yellow answer fill'
    re.compile(r"\b" + _YW + r"(?:[\s-]+(?:fill(?:ed)?|highlight(?:ed)?|shaded|background))?\s+(?:[\w-]+\s+)?"
               r"(?:inputs?|assumptions?|outputs?|drivers?|checks?|hard[\s-]?cod\w*|answers?)\b", re.I),
    # 'Input cells = yellow', 'Inputs: yellow fill', 'Scenario selector = yellow': the left-hand side must name a
    # category of cells or a cell noun (not 'Highlight colour: FFFF00', which only names the paint, and not a
    # status word beside a row label, 'Cost schedule | Yellow' in a RAG column: 47-R2-02)
    re.compile(r"(?<![\w-])(?:(?!" + _PAINTWORD + r"\b)[A-Za-z][\w-]*\s+){0,3}?(?!" + _PAINTWORD + r"\b)(?:" + _CAT + r"|"
               + _KEYNOUN + r")\)?" + _SEP + _YW + r"\b", re.I),
    # 'Yellow on Questions is the preserved source answer-cell fill' (clause starts with the colour,
    # then is/are, then a cell-category noun; 'Yellow and green are the brand colours' has none)
    re.compile(r"^\W*" + _YW + r"\b[^.;:=]{0,40}?\b(?:is|are)\s+" + _META + r"[^.;]{0,60}?\b(?:fill|cells?|inputs?|outputs?|"
               r"answers?|assumptions?|checks?|highlight\w*|drivers?)\b", re.I),
    # 'Assumptions (yellow) are updated each quarter', 'Key outputs (yellow fill)' (not 'Yellow Fill (#FFFF00)')
    re.compile(r"(?<![\w-])(?!" + _PAINTWORD + r"\b)[A-Za-z][\w-]*\s*\(\s*" + _YW + r"(?:[\s-]+" + _YNOUN + r")*\s*\)", re.I),
    # 47-positive-legends-missed (review 2026-10-03):
    # 'Inputs are in yellow', 'Input cells in yellow', 'Inputs use a yellow fill', 'Hard-codes in yellow'
    re.compile(r"\b" + _CAT + r"(?:[\s-]+(?:cells?|rows?|columns?|fields?|values?))?\s+(?:are\s+|is\s+)?(?:in|use|uses|using|"
               r"with|get|gets)\s+(?:a\s+|the\s+)?" + _YW + r"\b", re.I),
    # 'Assumptions are yellow', 'Assumptions to be updated annually are in yellow'
    re.compile(r"\b" + _CAT + r"(?:\s+[^\s.;:=]+){0,4}?\s+(?:are|is|appear|appears)\s+(?:shown\s+|kept\s+|always\s+)?(?:in\s+)?"
               + _YW + r"\b", re.I),
    # 'Error-check rows turn yellow when a check fails', 'Conditional formatting turns the cell yellow if ...'
    re.compile(r"\b(?:turns?|turned|turning|goes|go|becomes?|became|changes?\s+to|switch(?:es)?\s+to)\s+(?:the\s+)?"
               r"(?:(?:cell|row|column|line)s?\s+)?(?:to\s+)?" + _YW + r"\b", re.I),
    # 'Enter your assumptions in the yellow cells', 'Change only the yellow cells', 'Use the yellow cells to ...'
    re.compile(r"\b(?:enter|input|type|change|edit|use|modify|adjust|overwrite|amend)\w*\b[^.;]{0,40}?\b(?:the\s+|only\s+the\s+)?"
               + _YW + r"(?:[\s-]+(?:shaded|highlighted|filled))?\s+(?:cells?|fields?|boxes|inputs?|ranges?)\b", re.I),
    # 'Yellow is used only for inputs', 'No yellow highlighting is used except for the input cells'
    re.compile(r"\b" + _YW + r"(?:[\s-]+" + _YNOUN + r")*\s+(?:is|are)\s+(?:only\s+)?used\s+(?:only\s+)?(?:for|on|to|in|as)\s+\w",
               re.I),
    re.compile(r"\b" + _YW + r"\b[^.;]{0,60}?" + _EXCEPT_FOR_RX.pattern, re.I),
]


def _defines(cl: str) -> bool:
    return any(rx.search(cl) for rx in _DEF_RXS)


_CLAUSE_SPLIT_RX = re.compile(r"[\n\r;•·]+|(?<=[a-z0-9\)])\.\s+", re.I)
CLASS_ORDER = ("DEFINES", "WIP", "TEMPLATE", "NEGATED", "MENTION")
# a definition of the case template's own yellow, kept from the starting file ('Questions retains its original
# layout and yellow answer fill', 'Yellow on Questions is the preserved source answer-cell fill'): every case's
# yellow is pale (FFFFCC inputs, FFF2CC answers), so such a line explains no bright yellow (TEMPLATE_LINES_EXCUSE)
_TEMPLATE_RX = re.compile(r"\b(?:retain\w*|preserv\w*|original(?:ly)?|inherit\w*|template|as\s+received|legacy|"
                          r"carried\s+over|case[\s-](?:provided|supplied)|supplied|"
                          # the case's Questions answer cells (FFF2CC in every case): 'yellow answer cells', 'reserved answers'
                          r"answers?(?:[\s-](?:cells?|fills?|areas?|fields?|shading|locations?|cues?))?)\b", re.I)


def classify_clause(cl: str) -> str:
    if _negated(cl):
        return "NEGATED"
    if _wip(cl):
        return "WIP"
    if _defines(cl):
        return "TEMPLATE" if not TEMPLATE_LINES_EXCUSE and _TEMPLATE_RX.search(cl) else "DEFINES"
    return "MENTION"


# a clause that continues the previous one with a pronoun: 'Yellow is not used for inputs; it marks the
# output cells' - the pronoun stands for yellow (47-R2-11)
_PRONOUN_RX = re.compile(r"^\W*(it|this|these|they|that)\b", re.I)


def legend_clauses(text: str) -> list[str]:
    """The clauses of a text that mention bright yellow; a clause that follows one of them and
    opens with a pronoun ('it marks ...') is read with 'yellow' in the pronoun's place."""
    out, prev = [], False
    for cl in _CLAUSE_SPLIT_RX.split(text or ""):
        m = mentions_yellow(cl)
        if not m and prev:
            p = _PRONOUN_RX.match(cl)
            if p:
                cl = cl[:p.start(1)] + "yellow" + cl[p.end(1):]
                m = True
        if m:
            out.append(cl)
        prev = m
    return out


def classify_legend(text: str) -> Optional[str]:
    """Class of a text that mentions bright yellow: DEFINES | WIP | NEGATED | MENTION, or None
    when it does not mention bright yellow.  The most permissive clause wins: one clause that
    defines yellow as an ongoing convention makes the text a legend line."""
    classes = [classify_clause(cl) for cl in legend_clauses(text)]
    if not classes:
        return None
    for k in CLASS_ORDER:
        if k in classes:
            return k
    return "MENTION"


def classify_label(text: str) -> str:
    """Class of a swatch label WITHOUT the colour word ('Key outputs'): DEFINES | WIP | NEGATED | MENTION."""
    t = (text or "").strip()
    if sum(ch.isalpha() for ch in t) < 3:
        return "MENTION"
    if re.search(r"\b(?:no|none|not\s+used|n/a)\b", t, re.I) and len(t) < 40:
        return "NEGATED"
    if _wip(t):
        return "WIP"
    return "DEFINES"


# a colour-key label naming a CATEGORY of cells ('Key output cells', 'Hard-coded inputs', 'Error checks').
# 'totals' and 'results' are left out: 'Total costs' / 'Results' are ordinary row labels (review 2026-10-03).
_KEY_LABEL_RX = re.compile(r"\b(?:inputs?|assumptions?|outputs?|formulas?|calculations?|calcs?|links?|"
                           r"linked|checks?|hard[\s-]?cod\w*|drivers?|historicals?|actuals?|sources?|constants?|"
                           r"answers?|key\s+(?:figures?|metrics?|values?|cells?))\b", re.I)


def is_key_label(text: str) -> bool:
    """A short label that names a category of cells, as a colour key does ('Key output cells')."""
    t = (text or "").strip()
    return 0 < len(t) <= KEY_LABEL_MAX_LEN and bool(_KEY_LABEL_RX.search(t)) and classify_label(t) == "DEFINES"


# a bare colour label, optionally with its colour code: 'Yellow', 'Yellow fill:', 'Yellow Fill (#FFFF00)'
_BARE_LABEL_RX = re.compile(r"^\W*" + _YW + r"(?:[\s-]+" + _YNOUN + r")*" + _HEXCODE + r"?\W*$", re.I)
_LEADS_RX = re.compile(r"^\W*(?:" + _YW + r")\b", re.I)
_LEGEND_HDR_RX = re.compile(r"\b(?:legend|colou?r[\s-]*(?:key|code[sd]?|coding|conventions?|scheme|guide)|"
                            r"formatting\s+(?:key|conventions?|guide)|conventions|key\s+to\s+(?:colou?rs|formatting|shading)|"
                            r"glossary|cell\s+(?:key|styles|formatting))\b", re.I)
# the text OF a highlighted cell (or its comment) that marks it unfinished (rubric 'bad' case)
_WIP_CONTENT_RX = re.compile(r"^\W*(?:\?{2,}|x{3,})\W*$|\b(?:tbd|tbc|tba|to[\s-]?do|placeholder|" + _WIP_WORD + r"|fix\s?me|"
                             r"to\s+be\s+(?:confirmed|determined|completed|updated|provided|added)|update\s+later|"
                             r"confirm\s+with|check\s+this|to\s+confirm|not\s+yet\s+(?:confirmed|final\w*|known|available)|"
                             r"still\s+to|fill\s+in\s+later|dummy)\b", re.I)


def is_wip_content(text: Optional[str]) -> bool:
    return bool(text) and bool(_WIP_CONTENT_RX.search(text))


def _flags(txt: str) -> tuple:
    """(text, mentions_yellow, legend_header, key_label) of a cell text."""
    return (txt, mentions_yellow(txt), len(txt) <= 60 and bool(_LEGEND_HDR_RX.search(txt)),
            len(txt) <= KEY_LABEL_MAX_LEN and bool(_KEY_LABEL_RX.search(txt)))


def _short(t: str, n: int = 90) -> str:
    t = " ".join((t or "").split())
    return t if len(t) <= n else t[:n - 1] + "…"


# ---------------------------------------------------------------------------- unresolvable colours
# ---------------------------------------------------------------------------- check
# row-buffer cell kinds
K_TEXT, K_VALUE, K_FORMULA, K_EMPTY = "t", "v", "f", "e"


class C47(Check):
    """The rule above.  A fill colour the core cannot resolve (core.styles.UnknownColour: an invalid rgb, a
    theme / indexed index beyond the palette ...; finding 47-R2-08) is read as Excel's default for a fill:
    none (Patrick 2026-10-05: every attempt graded).  Every examined position (cell, row / column style,
    style 0) whose fill uses one is recorded in stats.defaults.unresolved_colour, and the colours in
    stats.unresolved_fill_colours.  Building the style table never raises, nor does a style no examined
    position uses."""
    number = 47
    key = "Formatting/No bright-yellow highlighting"
    needs_cells = True
    needs_rows = True
    sheet_kinds = ("worksheet", "dialogsheet", "macrosheet")

    # -------------------------------------------------------------- unresolvable fill colours: no fill
    def start(self, wb):
        super().start(wb)
        self.st = wb.styles
        self._unknown: dict = {}              # UnknownColour -> where it was first met
        self._keys_cache: dict = {}           # style index -> unknown colours its fill paint uses
        self._admit(self._style_keys(0), "the default cell style 0 (every empty cell)")
        self._start_body(wb)

    def _style_keys(self, s) -> tuple:
        k = self._keys_cache.get(s)
        if k is None:
            k = self._keys_cache[s] = _paint_unknowns(self.st.cell_fill(s))
        return k

    def _admit(self, keys, where: str):
        """A position whose fill uses unresolvable colour(s): read as no fill (recorded)."""
        if not keys:
            return
        for k in keys:
            self._unknown.setdefault(k, where)
        self.note_default("unresolved_colour", f"{where}: fill colour {', '.join(str(k) for k in keys)} read as no "
                                               f"fill")

    def _paint(self, s) -> FillPaint:
        """The fill paint of style s, an unresolvable colour painting nothing (cached)."""
        p = self._paint_cache.get(s)
        if p is None:
            p = self._paint_cache[s] = known_paint(self.st.cell_fill(s))
        return p

    def sheet_start(self, head):
        # column styles through the shared <col> reading (later entry wins where entries overlap)
        for lo, hi, ci in head.col_segments():
            if ci.style is not None:
                self._admit(self._style_keys(ci.style),
                            f"the column style of {location(head.name, f'{index_to_col(lo)}:{index_to_col(hi)}')}")
        self._sheet_start(head)

    def row(self, row):
        if row.style is not None:
            ks = self._style_keys(row.style)
            if ks:
                self._admit(ks, f"the row style of {location(self.sheet, f'{row.r}:{row.r}')}")
        self._row(row)

    def cell(self, cell):
        ks = self._style_keys(cell.s)
        if ks:
            self._admit(ks, f"{location(self.sheet, cell.ref)} (cell style {cell.s})")
        self._cell(cell)

    def sheet_end(self, head, tail):
        # conditional-format fills with an unresolvable colour are not counted either: such a rule is OFF here
        # (Patrick 2026-10-05: conditional formats never stop a grading), see _sheet_end
        self._sheet_end(head, tail)

    def finish(self) -> dict:
        if self._unknown:
            self.stats["unresolved_fill_colours"] = [{"colour": k.ref, "why": k.why, "first_used": w}
                                                     for k, w in self._unknown.items()]
        return self._finish_body()

    # -------------------------------------------------------------- workbook level (one reading)
    def _start_body(self, wb):
        self._paint_cache: dict = {}          # style index -> fill paint under this reading
        self._bright_cache: dict = {}         # style index -> bright colour ('FFRRGGBB') or None
        self._filled_cache: dict = {}         # style index -> has any non-none fill
        self._sst_flags: dict = {}            # shared-string index -> (mentions, header, wip)
        self.lines: list[dict] = []           # legend lines {where, text, class, how}
        self.offenders: dict = {}             # sheet -> {"cells": runs, ...}
        self.n_bright_cells = 0
        self.n_swatches = 0
        self.colours = set()
        self.cf_hits: list = []               # (sheet, sqref, colour, rule type)
        self.cf_unknown_fills: list = []      # CF rules whose fill colour cannot be resolved: not counted
        self.row_hits: list = []              # (sheet, r1, r2, colour)
        self.col_hits: list = []              # (sheet, c1, c2, colour)
        self.default_hits: list = []          # sheet names (style 0 bright)
        self.wip_cells: list = []             # (sheet, r, c, text, colour)
        self.unreadable: list = []            # locations of text cells whose text could not be decoded
        self.default_bright = self._bright(0)
        self._read_textboxes()

    def _bright(self, s) -> Optional[str]:
        """Bright-yellow colour painted by style s (own fill), else None."""
        v = self._bright_cache.get(s, 0)
        if v != 0:
            return v
        p = self._paint(s)
        out = None
        if p.kind == "gradient":
            out = next((c for c in p.stops if is_bright_yellow(c)), None)
        elif p.effective and is_bright_yellow(p.effective):
            out = p.effective
        self._bright_cache[s] = out
        return out

    def _filled(self, s) -> bool:
        v = self._filled_cache.get(s)
        if v is None:
            v = self._paint(s).effective is not None
            self._filled_cache[s] = v
        return v

    def _read_textboxes(self):
        """Drawing text boxes / shapes: each paragraph is one text unit (legend source)."""
        for si in self.wb.sheets:
            if not si.part:
                continue
            for r in self.wb.rels_of_type(si.part, "drawing"):
                if not r.part:
                    continue
                root = self.wb.xml(r.part)
                if root is None:
                    continue
                for body in root.iter():
                    if _ln(body.tag) != "txBody":
                        continue
                    paras = []
                    for p in body:
                        if _ln(p.tag) == "p":
                            txt = "".join((t.text or "") for t in p.iter() if _ln(t.tag) == "t")
                            if txt.strip():
                                paras.append(txt)
                    if paras:
                        self._classify_unit(f"{si.name} (text box)", "\n".join(paras), "textbox")

    def _classify_unit(self, where: str, text: str, how: str):
        cls = classify_legend(text)
        if cls is not None:
            self.lines.append({"where": where, "text": _short(text, 160), "class": cls, "how": how})

    # -------------------------------------------------------------- sheet level (one reading)
    def _sheet_start(self, head):
        self.sheet = head.name
        self.runs: list = []                  # offending own-fill runs (r, c1, c2, colour)
        self.swatch_cands: list = []          # (r, c, colour) + _swatch_label(...)
        self.headers: list = []               # (r, c) legend-header-like texts
        self.swatchlike: dict = {}            # (r, c) -> label: empty filled non-bright cells with a text label
        self.selfkeys: dict = {}              # (r, c) -> text: filled non-bright text cells naming a cell category
        self.row_r = None
        self.row_buf: list = []
        self.row_flag = False
        self.row_styles: list = []            # (r, colour)
        # column styles: the shared <col> reading (core.sheet.paint_cols, SheetHead.col_segments) - where
        # entries overlap the later one wins on the columns it covers, so an overridden style never paints
        for lo, hi, ci in head.col_segments():
            if ci.style is None:
                continue
            col = self._bright(ci.style)
            if col:
                self.col_hits.append((head.name, lo, hi, col))
                self.colours.add(col)
        if self.default_bright:
            self.default_hits.append(head.name)
            self.colours.add(self.default_bright)

    def _row(self, row):
        if self.row_buf:
            self._flush_row()
        self.row_r = row.r
        if row.style is not None:
            col = self._bright(row.style)
            if col:
                self.row_styles.append((row.r, col))
                self.colours.add(col)

    def _text_flags(self, cell) -> Optional[tuple]:
        """(text, mentions_yellow, legend_header, key_label) of a constant text cell; None when
        its text cannot be decoded (shared-string index out of range): recorded in self.unreadable,
        decisive only when the workbook has bright yellow (47-R2-09)."""
        if cell.t == "s":
            key = cell.raw
            f = self._sst_flags.get(key, 0)
            if f == 0:
                try:
                    txt = cell.value
                except GradingError:
                    f = None
                else:
                    f = _flags(txt if isinstance(txt, str) else "")
                self._sst_flags[key] = f
        else:
            try:
                txt = cell.value
            except GradingError:
                f = None
            else:
                f = _flags(txt if isinstance(txt, str) else "")
        if f is None:
            self.unreadable.append(location(self.sheet, cell.ref))
        return f

    def _cell(self, cell):
        r = cell.row
        if r != self.row_r:
            if self.row_buf:
                self._flush_row()
            self.row_r = r
        s = cell.s
        bright = self._bright(s)
        text = None
        if cell.formula is not None or cell.array is not None:
            kind = K_FORMULA
        elif cell.raw is None or cell.raw == "":
            kind = K_EMPTY
        elif cell.t in ("s", "inlineStr", "str"):
            flags = self._text_flags(cell)
            if flags is None:
                kind = K_VALUE                   # opaque: unreadable text is data, never a label or a sample
            elif not flags[0].strip():
                kind = K_EMPTY
            else:
                text, mention, header, keylabel = flags
                kind = K_TEXT
                if mention or header:
                    self.row_flag = True
                if header:
                    self.headers.append((r, cell.col))
                if keylabel and not bright and self._filled(s):
                    self.row_flag = True          # maybe another colour's self-labelled key sample
        else:
            kind = K_VALUE
        if bright:
            self.row_flag = True
        elif kind == K_EMPTY and self._filled(s):
            self.row_flag = True
        self.row_buf.append((cell.col, kind, text, bright, s))

    def _flush_row(self):
        buf, r = self.row_buf, self.row_r
        self.row_buf = []
        if self.swatch_cands:
            self._mark_table_rows(buf, r)
        if not self.row_flag:
            return
        self.row_flag = False
        buf.sort(key=lambda x: x[0])
        by_col = {x[0]: x for x in buf}
        texts = [(x[0], x[2]) for x in buf if x[1] == K_TEXT]
        # 1. legend lines from cells that mention yellow
        for i, (c, t) in enumerate(texts):
            if not mentions_yellow(t):
                continue
            line, extra = t, None
            if _BARE_LABEL_RX.match(t):
                right = [tt for cc, tt in texts[i + 1:] if cc - c <= 3][:2]
                left = [tt for cc, tt in texts[:i] if c - cc <= 3][-1:]
                if right:
                    line = f"{t.strip()} = " + " – ".join(x.strip() for x in right)
                    if left:
                        # 'Review / Unfinished | Yellow Fill (#FFFF00) | None remaining' (attempt 2226): the
                        # left cell is the meaning too; it can only restrict (WIP / NEGATED), never excuse
                        extra = f"{left[0].strip()} = {t.strip()}"
                elif left:
                    line = f"{left[0].strip()} = {t.strip()}"
            where = location(self.sheet, f"{index_to_col(c)}{r}")
            cls = classify_legend(line)
            if cls is not None:
                self.lines.append({"where": where, "text": _short(line, 160), "class": cls, "how": "cell"})
            if extra is not None and classify_legend(extra) in ("WIP", "NEGATED"):
                self.lines.append({"where": where, "text": _short(extra, 160), "class": classify_legend(extra),
                                   "how": "cell (label left of the colour)"})
        # 2. other colours' key samples: an empty filled non-bright cell with a text label beside it (right,
        #    or left), and a filled non-bright text cell naming a cell category ([blue 'Inputs']).  A header
        #    band (empty filled cell + title in the same fill) is not a sample: the label cell and the other
        #    neighbour must not share the sample's fill.
        for c, kind, text, bright, s in buf:
            if bright or not self._filled(s):
                continue
            fill = self._paint(s).effective
            nxt, prv = by_col.get(c + 1), by_col.get(c - 1)
            if kind == K_EMPTY:
                for lab, other in ((nxt, prv), (prv, nxt)):
                    if lab is not None and lab[1] == K_TEXT and self._paint(lab[4]).effective != fill and \
                            (other is None or self._paint(other[4]).effective != fill):
                        self.swatchlike[(r, c)] = lab[2]
                        break
            elif kind == K_TEXT and is_key_label(text) and \
                    all(x is None or self._paint(x[4]).effective != fill for x in (nxt, prv)):
                self.selfkeys[(r, c)] = text
        # 3. bright cells: swatch candidates or offenders
        run = None
        skip_to = 0
        for c, kind, text, bright, s in buf:
            if not bright or c <= skip_to:
                continue
            lft, rgt = by_col.get(c - 1), by_col.get(c + 1)
            # a sample merged across 2-3 empty cells ('[sample  ] | Key output cells', 47-R2-10): one candidate
            # for the span; sheet_end keeps it only when the span is a merged range of the sheet
            c2 = c
            if kind == K_EMPTY and not (lft and lft[3]):
                while c2 - c + 1 < MERGED_SWATCH_MAX_CELLS:
                    nx = by_col.get(c2 + 1)
                    if nx is not None and nx[1] == K_EMPTY and nx[3] == bright:
                        c2 += 1
                    else:
                        break
                after = by_col.get(c2 + 1)
                if after and after[3]:
                    c2 = c                        # the band goes on: not a sample
            if c2 > c:
                cand = self._swatch_label(c, c2, by_col, texts)
                if cand is not None:
                    cand["c2"] = c2
                    self.n_bright_cells += c2 - c + 1
                    self.colours.add(bright)
                    self.swatch_cands.append((r, c, bright, cand))
                    skip_to = c2
                    continue
            self.n_bright_cells += 1
            self.colours.add(bright)
            cand = None
            isolated = not (lft and lft[3]) and not (rgt and rgt[3])
            if kind == K_EMPTY and isolated:
                cand = self._swatch_label(c, c, by_col, texts)
            elif kind == K_TEXT and isolated and mentions_yellow(text) and \
                    (_BARE_LABEL_RX.match(text) or classify_legend(text) in ("DEFINES", "TEMPLATE", "NEGATED")):
                # self-labelled sample: the yellow cell's own text is its legend line ('Yellow fill =
                # ...', 'Unfinished / review me (yellow) - none in this file': the colour word may sit
                # anywhere, Patrick 2026-10-06); that line was classified in step 1 and decides whether
                # yellow is excused.  A yellow note that defines nothing or marks unfinished work
                # ('Yellow flag: rate TBD') is not one.
                self.n_swatches += 1
                continue
            elif kind == K_TEXT and isolated and len(text) <= KEY_LABEL_MAX_LEN and \
                    (lab := self._swatch_label(c, c, by_col, texts)) is not None and lab["mentions"]:
                # Patrick 2026-10-06 (attempt 1605 and 13 more): a legend sample that holds a short text
                # ('Unfinished / review me', 'sample', 'none', 'n/a') beside a label that mentions yellow
                # is documentation like an empty sample, whatever its text says; the label line decides
                # whether the rest of the yellow is excused
                cand = lab
            elif kind == K_TEXT and isolated and is_key_label(text) and self._right_clear(c, by_col):
                # self-labelled sample without the colour word ([yellow 'Key output'] under a legend header; a
                # description may follow it, a number / formula may not)
                cand = {"label": text, "mentions": False, "lc": c, "side": "self", "alone": True,
                        "left_is_data": self._left_is_data(c, by_col), "clear": True}
            if cand is not None:
                self.swatch_cands.append((r, c, bright, cand))
                continue
            if kind == K_TEXT and is_wip_content(text):
                self.wip_cells.append((self.sheet, r, c, text, bright))
            if run is not None and run[2] == c - 1 and run[3] == bright:
                run[2] = c
            else:
                if run is not None:
                    self.runs.append(tuple(run))
                run = [r, c, c, bright]
        if run is not None:
            self.runs.append(tuple(run))

    @staticmethod
    def _left_is_data(c, by_col) -> bool:
        """The cell left of column c holds a value / formula / text other than a legend header."""
        lft = by_col.get(c - 1)
        return lft is not None and lft[1] in (K_VALUE, K_FORMULA, K_TEXT) and not \
            (lft[1] == K_TEXT and _LEGEND_HDR_RX.search(lft[2]))

    @staticmethod
    def _right_clear(c, by_col, any_kind=False) -> bool:
        """No number / formula (any_kind: nothing at all) in the KEY_LINE_CLEAR_COLS columns right of c."""
        for d in range(1, KEY_LINE_CLEAR_COLS + 1):
            x = by_col.get(c + d)
            if x is None or (x[1] == K_EMPTY and not x[3]):
                continue
            if any_kind or x[1] in (K_VALUE, K_FORMULA) or x[3]:
                return False
        return True

    def _swatch_label(self, c1, c2, by_col, texts) -> Optional[dict]:
        """The label of an empty bright cell (columns c1..c2, one cell or a merged span), or None
        when it has none:
        {label, mentions (starts with the colour word), lc (label column), side ('right' | 'left'),
         alone (right label: the cell right of it is empty; left label: nothing right of the sample),
         left_is_data (the cell left of the sample / of a left label holds data),
         clear (no number / formula in the KEY_LINE_CLEAR_COLS columns right of the label / sample)}.
        A label that starts with the colour word wins on either side: 'Yellow fill | [sample] |
        Unfinished / review (none in delivered file)' is labelled by the left text (47-R2-01)."""
        left_is_data = self._left_is_data(c1, by_col)
        right = None
        for d in range(1, SWATCH_LABEL_MAX_DIST_RIGHT + 1):
            x = by_col.get(c2 + d)
            if x is None or x[1] == K_EMPTY and not x[3]:
                continue
            if x[1] != K_TEXT:
                break
            t = x[2]
            after = by_col.get(c2 + d + 1)
            alone = after is None or (after[1] == K_EMPTY and not after[3])
            # Patrick 2026-10-06: the colour word may sit anywhere in the label ('Unfinished / review me
            # (yellow fill) - none in the delivered file', attempt 1924), not only at its start
            right = {"label": t, "mentions": mentions_yellow(t), "lc": c2 + d,
                     "side": "right", "alone": alone, "left_is_data": left_is_data,
                     "clear": self._right_clear(c2 + d, by_col)}
            break
        if right is not None and right["mentions"]:
            return right
        for cc, t in reversed(texts):
            if c1 - SWATCH_LABEL_MAX_DIST_LEFT <= cc < c1 and mentions_yellow(t):
                return {"label": t, "mentions": True, "lc": cc, "side": "left", "alone": True, "left_is_data": False,
                        "clear": True}
        if right is not None:
            return right
        # a key label WITHOUT the colour word to the left ('Key outputs | [sample]'); nothing may follow the sample
        for d in range(1, SWATCH_PLAIN_LEFT_DIST + 1):
            x = by_col.get(c1 - d)
            if x is None or x[1] == K_EMPTY and not x[3]:
                continue
            if x[1] != K_TEXT or not self._right_clear(c2, by_col, any_kind=True):
                return None
            return {"label": x[2], "mentions": False, "lc": c1 - d, "side": "left", "alone": True,
                    "left_is_data": self._left_is_data(c1 - d, by_col), "clear": True}
        return None

    def _mark_table_rows(self, buf, r):
        """A label right above label / value rows heads a data table ('[yellow] | Inputs' over 'Growth | 5%'):
        a section header, not a colour-key line (47-R2-04).  Called for every row while swatch
        candidates exist; marks the candidates of the TABLE_ROWS_BELOW rows above this one."""
        near = [cand for cr, _c, _col, cand in self.swatch_cands
                if 0 < r - cr <= TABLE_ROWS_BELOW and not cand.get("heads_table")]
        if not near:
            return
        by_col = {x[0]: x for x in buf}
        for cand in near:
            lc = cand["lc"]
            for dc in (-1, 0, 1):
                x = by_col.get(lc + dc)
                if x is not None and x[1] == K_TEXT and not self._right_clear(lc + dc, by_col):
                    cand["heads_table"] = True
                    break

    def _sheet_end(self, head, tail):
        if self.row_buf:
            self._flush_row()
        name = head.name
        # swatches
        for r, c, colour, cand in self.swatch_cands:
            ok = False
            label = cand["label"]
            c2 = cand.get("c2", c)
            if c2 != c and not any(m.r1 == r == m.r2 and m.c1 == c and m.c2 == c2 for m in tail.merges):
                self.runs.append((r, c, c2, colour))     # adjacent yellow cells that are not one merged cell: a band
                continue
            if cand["mentions"]:
                # a sample beside a line that starts with the colour word; that line was already
                # classified as a legend line in _flush_row (it decides whether yellow is excused)
                ok = True
            elif not cand["left_is_data"] and cand["clear"] and classify_label(label) == "DEFINES":
                how = None
                key = is_key_label(label)
                if cand["side"] != "right" and not key:
                    pass              # a label left of the sample / in it must name a cell category
                elif self._legend_header(r, c):
                    how = "legend header"
                elif self._key_block(r, c):
                    how = "colour key block"
                elif key and cand["alone"] and cand["side"] != "self" and not cand.get("heads_table"):
                    how = "colour-key label"
                if how:
                    ok = True
                    self.lines.append({"where": location(name, f"{index_to_col(cand['lc'])}{r}"),
                                       "text": _short(label, 160), "class": "DEFINES",
                                       "how": f"label of swatch {range_to_str(r, c, r, c2)} ({cand['side']}, {how})"})
            if ok:
                self.n_swatches += 1
            else:
                self.runs.append((r, c, c2, colour))
                if cand["side"] == "self" and is_wip_content(label):
                    self.wip_cells.append((name, r, c, label, colour))
        # comments (legend lines + WIP comments on bright cells)
        if tail.has_comments or tail.has_threaded_comments:
            for cm in self.wb.comments(head.info):
                self._classify_unit(location(name, cm.ref) + " (comment)", cm.text, "comment")
                if is_wip_content(cm.text):
                    rc = _split(cm.ref)
                    if rc is not None:
                        col = self._bright_at(rc)
                        if col and not any(w[:3] == (name, rc[0], rc[1]) for w in self.wip_cells):
                            self.wip_cells.append((name, rc[0], rc[1], f"comment: {cm.text}", col))
        # row styles: group consecutive rows of the same colour
        rs = sorted(self.row_styles)
        i = 0
        while i < len(rs):
            j = i
            while j + 1 < len(rs) and rs[j + 1][0] == rs[j][0] + 1 and rs[j + 1][1] == rs[i][1]:
                j += 1
            self.row_hits.append((name, rs[i][0], rs[j][0], rs[i][1]))
            i = j + 1
        # conditional formats (dxf fills; colour scales, data bars and icon sets are not highlighting)
        for cf in tail.conditional_formats:
            for rule in cf.rules:
                if rule.type in ("colorScale", "dataBar", "iconSet"):
                    continue
                d = rule.dxf if rule.dxf is not None else self.st.dxf(rule.dxf_id)
                cols, together = dxf_fill_paints(self.st, d.fill if d is not None else None)
                unknown = [c for c in cols if is_unknown(c)]
                if unknown:
                    # whether it paints bright yellow is unknown: the rule is OFF for this check (Patrick
                    # 2026-10-05), recorded
                    self.cf_unknown_fills.append(f"{location(name, ','.join(cf.sqref.split()))}: {rule.type or '?'} "
                                                 f"rule fill {unknown[0]} - not counted"[:300])
                cols = [c for c in cols if c is not None and not is_unknown(c)]
                yellow = [x for x in cols if is_bright_yellow(x)]
                if not yellow or not CF_COUNTS_ALWAYS:
                    continue
                self.cf_hits.append((name, cf.sqref, yellow[0], rule.type or "?"))
                self.colours.add(yellow[0])
        if self.runs:
            self.offenders[name] = self.runs
        self.runs, self.swatch_cands, self.swatchlike, self.selfkeys, self.headers = [], [], {}, {}, []

    def _bright_at(self, rc) -> Optional[str]:
        r, c = rc
        for rr, c1, c2, col in self.runs:
            if rr == r and c1 <= c <= c2:
                return col
        return None

    def _legend_header(self, r, c) -> bool:
        """A legend / colour key / conventions header up to LEGEND_HEADER_ROWS_ABOVE rows above."""
        for hr, hc in self.headers:
            if 0 < r - hr <= LEGEND_HEADER_ROWS_ABOVE and abs(hc - c) <= LEGEND_HEADER_COL_DIST:
                return True
        return False

    def _key_block(self, r, c) -> bool:
        """Another colour's key sample within KEY_BLOCK_ROW_DIST rows in the same column whose label names a
        category of cells ('Formulas', 'Inputs'): a colour key, not a status column of section names."""
        for dr in range(-KEY_BLOCK_ROW_DIST, KEY_BLOCK_ROW_DIST + 1):
            if not dr:
                continue
            lab = self.swatchlike.get((r + dr, c)) or self.selfkeys.get((r + dr, c))
            if lab is not None and is_key_label(lab):
                return True
        return False

    # -------------------------------------------------------------- verdict (one reading)
    def _finish_body(self) -> dict:
        if self.unreadable and (self.n_bright_cells or self.row_hits or self.col_hits or self.default_hits
                                or self.cf_hits):
            # the unreadable text could be the legend that excuses this yellow, or the WIP line that blocks
            # it: the verdict depends on it (no yellow anywhere: it cannot, so no raise - 47-R2-09)
            raise GradingError(f"{len(self.unreadable)} text cell(s) could not be read ({self.unreadable[0]}"
                               f"{', ...' if len(self.unreadable) > 1 else ''}) and the workbook uses bright yellow: "
                               f"the unreadable text may be the legend that explains it")
        defines = [l for l in self.lines if l["class"] == "DEFINES"]
        wip_lines = [l for l in self.lines if l["class"] == "WIP"]
        blocked = bool(WIP_LEGEND_BLOCKS_EXCUSE and defines and wip_lines)
        excused = bool(defines) and not blocked
        others = sorted((l for l in self.lines if l["class"] != "DEFINES"),
                        key=lambda l: CLASS_ORDER.index(l["class"]))
        legend = f"{defines[0]['where']}: {defines[0]['text']!r}" if defines else ""
        if excused:
            why = f"excused by the legend {legend}"
        elif blocked:
            w = wip_lines[0]
            why = (f"the workbook's own text defines yellow as unfinished ({w['where']}: {w['text']!r}), so the "
                   f"other legend line {legend} does not make yellow an explained convention")
        elif others:
            o = others[0]
            what = {"WIP": "defines yellow as unfinished", "NEGATED": "says no yellow is used",
                    "TEMPLATE": "explains the case template's retained yellow, not this bright yellow",
                    "MENTION": "does not define what yellow means"}[o["class"]]
            why = f"the only text about yellow ({o['where']}: {o['text']!r}) {what}, so it does not explain the yellow"
        else:
            why = "no legend or other text in the workbook explains yellow"

        n_cells_flagged = 0
        if not excused:
            wip_by_sheet = defaultdict(list)
            for sheet, r, c, text, _col in self.wip_cells:
                wip_by_sheet[sheet].append((r, c, text))
            for sheet, runs in self.offenders.items():
                by_col = defaultdict(list)
                for r, c1, c2, col in runs:
                    by_col[col].append((r, c1, c2))
                for col, rr in by_col.items():
                    for (r1, c1, r2, c2) in _rects(rr):
                        n = (r2 - r1 + 1) * (c2 - c1 + 1)
                        n_cells_flagged += n
                        rng = range_to_str(r1, c1, r2, c2)
                        what = f"Cell {rng} is" if n == 1 else f"{n} cells {rng} are"
                        marks = [f"{index_to_col(c)}{r} says {_short(t, 40)!r}" for r, c, t in wip_by_sheet[sheet]
                                 if r1 <= r <= r2 and c1 <= c <= c2][:3]
                        extra = f" ({'; '.join(marks)}: a placeholder marker)" if marks else ""
                        self.add_mistake(location(sheet, rng),
                                         f"{what} filled bright yellow #{col[-6:]}{extra}; {why}.")
            for sheet, r1, r2, col in self.row_hits:
                rng = f"{r1}:{r2}"
                self.add_mistake(location(sheet, rng), f"Row style of row(s) {rng} paints the row bright yellow "
                                                       f"#{col[-6:]}; {why}.")
            for sheet, c1, c2, col in self.col_hits:
                rng = f"{index_to_col(c1)}:{index_to_col(c2)}"
                self.add_mistake(location(sheet, rng), f"Column style of column(s) {rng} paints the column bright "
                                                       f"yellow #{col[-6:]}; {why}.")
            for sheet in self.default_hits:
                self.add_mistake(sheet, f"The workbook's default cell style paints every empty cell of '{sheet}' "
                                        f"bright yellow #{self.default_bright[-6:]}; {why}.")
            for sheet, sq, col, typ in self.cf_hits:
                sqc = ",".join(sq.split())
                self.add_mistake(location(sheet, sqc), f"Conditional format ({typ}) on {sqc} paints bright yellow "
                                                       f"#{col[-6:]} when it fires; {why}.")
        elif WIP_CONTENT_OVERRIDES_LEGEND:
            for sheet, r, c, text, col in self.wip_cells:
                ref = f"{index_to_col(c)}{r}"
                n_cells_flagged += 1
                self.add_mistake(location(sheet, ref),
                                 f"Cell {ref} is filled bright yellow #{col[-6:]} and marks unfinished work "
                                 f"({_short(text, 60)!r}); a yellow legend does not excuse a placeholder left in "
                                 f"the deliverable.")
        stats = {
            "bright_cells": self.n_bright_cells,
            "swatches": self.n_swatches,
            "offending_cells_listed": n_cells_flagged,
            "row_style_hits": len(self.row_hits), "col_style_hits": len(self.col_hits),
            "default_style_sheets": len(self.default_hits), "cf_rules": len(self.cf_hits),
            "colours": sorted(c[-6:] for c in self.colours),
            "excused_by_legend": excused,
            "legend_lines": self.lines[:MAX_LINES_IN_STATS],
            "n_legend_lines": len(self.lines),
            "wip_content_cells": len(self.wip_cells),
            "band": {"hue": [HUE_MIN, HUE_MAX], "sat_min": SAT_MIN, "val_min": VAL_MIN},
            # Patrick 2026-10-05: conditional formats never stop a grading - every assumption is recorded
            "cf_assumptions": {
                "policy": "a conditional-format fill whose colour cannot be resolved is not counted as bright "
                          "yellow: the rule is taken as off (Patrick 2026-10-05)",
                "rules": len(self.cf_unknown_fills), "examples": self.cf_unknown_fills[:10]},
        }
        any_yellow = bool(self.offenders or self.row_hits or self.col_hits or self.default_hits or self.cf_hits)
        stats["legend_blocked_by_wip"] = blocked
        if excused and any_yellow:
            ps = f"Bright yellow is used but explained as an ongoing convention ({legend})."
        elif self.n_swatches and not any_yellow:
            ps = "No bright-yellow highlighting (only legend swatch sample cells)."
        else:
            ps = "No bright-yellow highlighting."
        return self.verdict(ps, "{n} bright-yellow highlighting finding(s) not explained as an ongoing convention.",
                            stats)


def _ln(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _split(ref: str):
    return split_ref((ref or "").replace("$", ""))


def _rects(runs):
    """Horizontal runs (r, c1, c2), any order -> rectangles (r1, c1, r2, c2): runs with the same
    column span on consecutive rows are stacked (same algorithm as refs.group_cells)."""
    out = []
    open_: dict = {}
    for r, c1, c2 in sorted(runs):
        key = (c1, c2)
        rect = open_.get(key)
        if rect is not None and rect[1] == r - 1:
            rect[1] = r
        else:
            if rect is not None:
                out.append((rect[0], c1, rect[1], c2))
            open_[key] = [r, r]
    for (c1, c2), (r1, r2) in open_.items():
        out.append((r1, c1, r2, c2))
    out.sort()
    return out
