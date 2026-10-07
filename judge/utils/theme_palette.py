"""Theme-colour resolution for the judge's evidence (judge v7 tier 2, 2026-09-10).

Excel stores most colours as a theme slot plus a tint rather than an RGB
value. openpyxl exposes those as `Color(type="theme", theme=N, tint=t)` whose
`.rgb` is a descriptor error string, so until now the extractor emitted no
colour token at all for them and the judge read every theme-coloured cell as
default black text on no fill — blind on the blue-input / header-shading
checks for every golden and the Excel add-in cohorts.

`load_palette(workbook)` parses the workbook's own `theme1.xml`
(`workbook.loaded_theme`, raw bytes) into the 12 scheme slots and
`resolve(palette, theme_index, tint)` returns the 6-hex RRGGBB Excel would
paint, applying the tint in HLS the way the spec defines it. Workbooks with
no theme part fall back to the Office default palette.
"""

from __future__ import annotations

import colorsys
import re
from typing import Optional, Sequence
import xml.etree.ElementTree as ET

# clrScheme child order in theme1.xml
_SCHEME_ORDER = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink")
# Cell colour theme index -> scheme slot. Indexes 0-3 are swapped pairs:
# 0 = lt1 (window/background 1), 1 = dk1 (text 1), 2 = lt2, 3 = dk2.
_INDEX_TO_SLOT = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                  "accent4", "accent5", "accent6", "hlink", "folHlink")

# Office 2013+ default theme, used when a workbook carries no theme part.
DEFAULT_PALETTE: tuple[str, ...] = (
    "FFFFFF",  # 0 lt1
    "000000",  # 1 dk1
    "E7E6E6",  # 2 lt2
    "44546A",  # 3 dk2
    "4472C4",  # 4 accent1
    "ED7D31",  # 5 accent2
    "A5A5A5",  # 6 accent3
    "FFC000",  # 7 accent4
    "5B9BD5",  # 8 accent5
    "70AD47",  # 9 accent6
    "0563C1",  # 10 hlink
    "954F72",  # 11 folHlink
)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")


def _slot_hex(node) -> Optional[str]:
    """`<a:srgbClr val>` or `<a:sysClr lastClr>` under one scheme slot."""
    for child in node:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "srgbClr":
            val = child.get("val")
        elif tag == "sysClr":
            val = child.get("lastClr") or child.get("val")
        else:
            continue
        if val and _HEX6.match(val):
            return val.upper()
    return None


def parse_theme(theme_xml) -> tuple[str, ...]:
    """12-entry palette indexed by the CELL theme index (0 = lt1, 1 = dk1, ...).
    Falls back slot-by-slot to DEFAULT_PALETTE for anything unreadable."""
    slots: dict[str, str] = {}
    try:
        if isinstance(theme_xml, str):
            theme_xml = theme_xml.encode("utf-8")
        root = ET.fromstring(theme_xml)
        scheme = root.find(f".//{{{_A_NS}}}clrScheme")
        if scheme is not None:
            for child in scheme:
                name = child.tag.rsplit("}", 1)[-1]
                if name in _SCHEME_ORDER:
                    hx = _slot_hex(child)
                    if hx:
                        slots[name] = hx
    except Exception:  # noqa: BLE001 — a bad theme part must never break extraction
        slots = {}
    return tuple(
        slots.get(slot, DEFAULT_PALETTE[i]) for i, slot in enumerate(_INDEX_TO_SLOT)
    )


def load_palette(workbook) -> tuple[str, ...]:
    """The workbook's palette (see parse_theme); DEFAULT_PALETTE when the
    workbook has no theme part (openpyxl-authored files, some converters)."""
    theme = getattr(workbook, "loaded_theme", None)
    if not theme:
        return DEFAULT_PALETTE
    return parse_theme(theme)


def apply_tint(hex6: str, tint: float) -> str:
    """ECMA-376 tint: in HLS, L' = L*(1+tint) for tint<0, L + (1-L)*tint for tint>0."""
    if not tint:
        return hex6.upper()
    r, g, b = (int(hex6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    tint = max(-1.0, min(1.0, float(tint)))
    l = l * (1.0 + tint) if tint < 0 else l + (1.0 - l) * tint
    l = max(0.0, min(1.0, l))
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return "".join(f"{int(round(v * 255)):02X}" for v in (r2, g2, b2))


def resolve(palette: Optional[Sequence[str]], theme_index, tint=0) -> Optional[str]:
    """RRGGBB (6 hex) for a theme colour, or None when the index is unusable."""
    try:
        idx = int(theme_index)
    except (TypeError, ValueError):
        return None
    pal = palette or DEFAULT_PALETTE
    if not 0 <= idx < len(pal):
        return None
    try:
        return apply_tint(pal[idx], float(tint or 0))
    except (TypeError, ValueError):
        return pal[idx]


def resolve_color(palette, color) -> Optional[str]:
    """openpyxl Color of type 'theme' -> RRGGBB; None for any other type."""
    if color is None or getattr(color, "type", None) != "theme":
        return None
    return resolve(palette, getattr(color, "theme", None), getattr(color, "tint", 0) or 0)
