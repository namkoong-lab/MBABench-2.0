"""Cell reference validation for formulas."""
import re
from typing import List

from openpyxl.workbook import Workbook

from .. import formula_validator

_ARITHMETIC = {"+", "-", "*", "/", "^"}
# With any of these in the formula the author is handling a non-number on
# purpose (=IFERROR(A1*2,0), =N(A1)+B1): never second-guess it.
_GUARD_FUNCTIONS = {"IFERROR", "IFNA", "ISERROR", "ISERR", "ISNUMBER", "ISTEXT", "ISNONTEXT",
                    "N", "T", "VALUE", "NUMBERVALUE", "TYPE", "ERROR.TYPE", "AGGREGATE"}
_DATE_LIKE = re.compile(r"^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}")


def _excel_reads_as_number(text: str) -> bool:
    """Excel coerces such text in arithmetic ("5", "1,200", "12%", "$3.50", a date)."""
    cleaned = text.strip().replace(",", "").replace("$", "").replace("€", "").replace("£", "").rstrip("%").strip()
    if not cleaned:
        return True  # "" + 1 is left to the recalculation to report
    try:
        float(cleaned)
        return True
    except ValueError:
        return bool(_DATE_LIKE.match(text))


def _validate_cell_references(formula: str, workbook: Workbook, current_worksheet: str = None) -> List[str]:
    """
    Refuse plain arithmetic on a cell that holds text (=B1*5 where B1 is a label
    is #VALUE!). Returns the list of error messages, [] when the formula is fine.

    Only a single-cell operand standing directly beside + - * / ^ (or under a
    leading minus or a trailing %) is looked at, on the sheet the reference
    names. Until 2026-09-19 the test was "a text cell within three characters of
    an operator OR A BRACKET", read on the CURRENT sheet whenever the sheet name
    was unquoted. That refused every lookup by label (XLOOKUP, MATCH, SUMIF,
    COUNTIF), every IF/SWITCH on a text selector, SUM over a range with a
    header, LEN/CONCAT/TEXTJOIN, and answer formulas such as =ROUND(Ratios!C20,2)
    whenever the answer sheet's own C20 held a unit label - 3,168 refusals in
    the v2 logs. When in doubt the formula is allowed: every write reports the
    recalculated value, so a real #VALUE! is still seen.
    """
    try:
        used = {formula_validator._bare_function_name(name) for name in
                formula_validator.extract_function_names(formula.upper())}
    except Exception:
        used = set()
    if used & _GUARD_FUNCTIONS:
        return []
    refs = list(formula_validator.iter_cell_refs(formula))

    errors: List[str] = []
    for ref in refs:
        if ref.end is not None:
            continue  # a range: SUM, lookups and counts skip or use text legitimately
        arithmetic = (
            (ref.prev[0] in ("OPERATOR-INFIX", "OPERATOR-PREFIX") and ref.prev[1] in _ARITHMETIC)
            or (ref.next[0] == "OPERATOR-INFIX" and ref.next[1] in _ARITHMETIC)
            or (ref.next[0] == "OPERATOR-POSTFIX" and ref.next[1] == "%")
        )
        if not arithmetic:
            continue
        sheet_name = ref.sheet or current_worksheet
        if not sheet_name or sheet_name not in workbook.sheetnames:
            continue
        try:
            value = workbook[sheet_name][ref.start].value
        except Exception:
            continue
        if not isinstance(value, str) or value.startswith("=") or _excel_reads_as_number(value):
            continue
        where = f"{ref.sheet}!{ref.start}" if ref.sheet else ref.start
        shown = value if len(value) <= 40 else value[:37] + "..."
        message = f"Cell {where} contains text '{shown}' but is used in arithmetic ({'+ - * / ^'})"
        if not ref.sheet and current_worksheet:
            message += f" in current worksheet '{current_worksheet}'"
        if message not in errors:
            errors.append(message)
    return errors
