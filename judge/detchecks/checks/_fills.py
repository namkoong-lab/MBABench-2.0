"""Helpers shared by checks 47 and 94 (group "fills"): what a conditional-format fill paints.

Measured by Patrick in Excel (2026-10-03, detchecks/out/excel_session_answers.md Q2): for a
conditional-format (dxf) SOLID fill Excel paints bgColor ONLY.  fgColor is ignored; bgColor
indexed 64 and bgColor rgb 00000000 paint BLACK (the alpha byte is ignored); no bgColor -> no
fill.  So ChatGPT's tool (fgColor=X + bgColor indexed 64) paints black, openpyxl's
PatternFill(start_color=X, fill_type="solid") (bgColor 00000000) paints black, and a dxf with
fgColor only paints nothing.  The core implements this (styles.fill_paint(..., dxf=True));
there is no ambiguity left, so nothing here raises.

Unresolvable colours (core.styles.UnknownColour) are read as Excel's default for their slot (Patrick
2026-10-05: every attempt graded): a fill as none (known_paint), a font colour as automatic, black.
"""
from __future__ import annotations

from typing import Optional

from ..core.styles import FillPaint, is_unknown

DEFAULT_FONT = "FF000000"           # automatic font colour (black)


def paint_unknowns(p) -> tuple:
    """The unresolvable colours (UnknownColour) a fill paint depends on, in order, without repeats."""
    out = []
    if p is None:
        return ()
    if p.kind == "gradient":
        out = [c for c in p.stops if is_unknown(c)]
    if is_unknown(p.effective):
        out.append(p.effective)
    return tuple(dict.fromkeys(out))


def _average(cols) -> Optional[str]:
    """Average of 'FFRRGGBB' colours (a gradient's effective colour, as core.styles computes it)."""
    if not cols:
        return None
    n = len(cols)
    return "FF" + "".join(f"{int(round(sum(int(c[i:i + 2], 16) for c in cols) / n)):02X}" for i in (2, 4, 6))


def known_paint(p) -> FillPaint:
    """A fill paint with every unresolvable colour read as Excel's default for a fill, none: a solid / pattern
    paint that uses one paints nothing; a gradient drops that stop (no stop left: nothing)."""
    if not paint_unknowns(p):
        return p
    if p.kind == "gradient":
        stops = tuple(c for c in p.stops if not is_unknown(c))
        if not stops:
            return FillPaint("none", None, None, None, None)
        return FillPaint("gradient", None, None, None, _average(stops), stops)
    return FillPaint("none", None, None, None, None)


def dxf_fill_paints(styles, fill) -> tuple[list[str], bool]:
    """(colours 'FFRRGGBB', together) for a conditional-format (dxf) fill.
    together=True: every colour is painted (gradient stops).  Otherwise one colour (the paint)
    or none.  ([], False) when the dxf paints no fill."""
    if fill is None or fill.kind == "none":
        return [], False
    p = styles.fill_paint(fill, dxf=True)
    if p.kind == "none":
        return [], False
    if fill.kind == "gradient":
        return list(p.stops), True
    return ([p.effective] if p.effective else []), False
