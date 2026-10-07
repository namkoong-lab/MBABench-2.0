"""Conditional-format text rules shared by checks 66 and 94.

containsText / notContainsText: Excel stores these rules with a SEARCH() formula
(NOT(ISERROR(SEARCH("x",A1)))) and evaluates that formula, so the rule's text uses SEARCH's
wildcards: '*' any run of characters, '?' one character, '~' escapes a following '*', '?' or
'~'; case-insensitive.  Measured by Patrick in Excel 2026-10-03
(detchecks/out/excel_session_answers.md Q3: "P*S" and "P?SS" both match "PASS").
beginsWith / endsWith are stored as LEFT()/RIGHT() comparisons: literal, no wildcards.
"""
from __future__ import annotations

import re
from functools import lru_cache


@lru_cache(maxsize=4096)
def _search_rx(needle: str):
    rx, i = [], 0
    while i < len(needle):
        ch = needle[i]
        if ch == "~" and i + 1 < len(needle) and needle[i + 1] in "*?~":
            rx.append(re.escape(needle[i + 1]))
            i += 2
            continue
        rx.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(rx), re.S | re.I)


def search_finds(needle: str, hay: str) -> bool:
    """Excel SEARCH(needle, hay) finds a match (wildcards, case-insensitive)."""
    return _search_rx(needle.casefold()).search(hay.casefold()) is not None


def text_rule_fires(rule_type: str, needle: str, hay: str) -> bool:
    """containsText / notContainsText (SEARCH wildcards), beginsWith / endsWith (literal),
    case-insensitive, for the cell's displayed-value text `hay`."""
    if rule_type in ("containsText", "notContainsText"):
        found = search_finds(needle, hay)
        return found if rule_type == "containsText" else not found
    h, n = hay.casefold(), needle.casefold()
    if rule_type == "beginsWith":
        return h.startswith(n)
    if rule_type == "endsWith":
        return h.endswith(n)
    raise ValueError(rule_type)
