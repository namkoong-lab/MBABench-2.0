"""
Formula validation utilities for Excel agent

Validates Excel formulas for:
- Invalid function names
- Potential undefined named ranges
- Basic syntax errors
"""

import re
from typing import Dict, Iterator, List, NamedTuple, Optional, Set, Tuple

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import FORMULAE as _OPENPYXL_FORMULAE

# Comprehensive list of valid Excel functions
VALID_EXCEL_FUNCTIONS: Set[str] = {
    # Math & Stats
    "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MAX", "MIN",
    "MEDIAN", "MODE", "STDEV", "STDEVP", "VAR", "VARP",
    "ROUND", "ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR", "TRUNC", "INT",
    "ABS", "SQRT", "POWER", "EXP", "LN", "LOG", "LOG10",
    "MOD", "QUOTIENT", "PRODUCT", "FACT", "GCD", "LCM",
    "SIGN", "RAND", "RANDBETWEEN",

    # Conditional & Logic
    "IF", "IFS", "AND", "OR", "NOT", "XOR", "TRUE", "FALSE",
    "IFERROR", "IFNA", "SWITCH",
    "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS",
    "COUNTIF", "COUNTIFS", "MAXIFS", "MINIFS",

    # Financial
    "PMT", "IPMT", "PPMT", "FV", "PV", "RATE", "NPER",
    "NPV", "IRR", "XIRR", "XNPV", "MIRR",
    "CUMIPMT", "CUMPRINC", "EFFECT", "NOMINAL",
    "SLN", "SYD", "DB", "DDB", "VDB",
    "PRICE", "YIELD", "DURATION", "MDURATION",

    # Lookup & Reference
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP",
    "INDEX", "MATCH", "XMATCH", "INDIRECT", "OFFSET", "CHOOSE",
    "ROW", "ROWS", "COLUMN", "COLUMNS",
    "ADDRESS", "AREAS", "TRANSPOSE",
    "GETPIVOTDATA", "HYPERLINK",

    # Text
    "CONCATENATE", "CONCAT", "TEXTJOIN",
    "LEFT", "RIGHT", "MID", "LEN", "TRIM",
    "UPPER", "LOWER", "PROPER",
    "SUBSTITUTE", "REPLACE", "REPT",
    "FIND", "FINDB", "SEARCH", "SEARCHB",
    "TEXT", "VALUE", "NUMBERVALUE",
    "CHAR", "CODE", "CLEAN", "EXACT",

    # Date & Time
    "DATE", "YEAR", "MONTH", "DAY", "TODAY", "NOW",
    "TIME", "HOUR", "MINUTE", "SECOND",
    "WEEKDAY", "WEEKNUM", "ISOWEEKNUM",
    "EOMONTH", "EDATE", "WORKDAY", "NETWORKDAYS",
    "DATEDIF", "DAYS", "DAYS360",
    "DATEVALUE", "TIMEVALUE",

    # Named intermediates. The house standards (House_Standards_v1.md,
    # prompt v14+) recommend LET and XMATCH; LibreOffice 24.8+ evaluates
    # both, so rejecting them here would only push the agent to nested
    # duplicates of the same expression.
    "LET",

    # Database & Array
    "SUMPRODUCT", "MMULT", "MDETERM", "MINVERSE",
    "TRANSPOSE", "UNIQUE", "SORT", "SORTBY", "FILTER",
    "SEQUENCE", "RANDARRAY", "ARRAY",

    # Information & Error Checking
    "ISBLANK", "ISERROR", "ISERR", "ISNA",
    "ISNUMBER", "ISTEXT", "ISLOGICAL", "ISREF",
    "ISFORMULA", "ISNONTEXT", "ISODD", "ISEVEN",
    "TYPE", "N", "NA", "ERROR.TYPE",
    "CELL", "INFO", "SHEET", "SHEETS",

    # Statistical (Extended)
    "AVERAGEA", "COUNTA", "MAXA", "MINA",
    "CORREL", "COVARIANCE.P", "COVARIANCE.S",
    "FORECAST", "FORECAST.LINEAR", "INTERCEPT", "SLOPE",
    "PEARSON", "RSQ", "STEYX", "TREND", "GROWTH",
    "PERCENTILE", "PERCENTILE.INC", "PERCENTILE.EXC",
    "QUARTILE", "QUARTILE.INC", "QUARTILE.EXC",
    "RANK", "RANK.AVG", "RANK.EQ",
    "STANDARDIZE", "Z.TEST",

    # Engineering (Common ones)
    "CONVERT", "BIN2DEC", "BIN2HEX", "BIN2OCT",
    "DEC2BIN", "DEC2HEX", "DEC2OCT",
    "HEX2BIN", "HEX2DEC", "HEX2OCT",
    "OCT2BIN", "OCT2DEC", "OCT2HEX",

    # Web & External Data
    "WEBSERVICE", "FILTERXML", "ENCODEURL",
}

# Functions added to Excel after 2007. The file format stores them with an
# "_xlfn." prefix (FILTER and SORT with "_xlfn._xlws."); openpyxl writes the
# text it is given, and an unprefixed name is "#NAME?" in both Excel and
# LibreOffice (LibreOffice re-saves it lowercased: "xlookup("). These are
# exactly the functions the House Standards prefer (XLOOKUP/XMATCH, IFS/SWITCH,
# LET), so add_future_function_prefixes() supplies the prefix on write.
FUTURE_FUNCTIONS: Set[str] = {
    # Excel 2010
    "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV", "CEILING.PRECISE",
    "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CONFIDENCE.NORM",
    "CONFIDENCE.T", "COVARIANCE.P", "COVARIANCE.S", "ERF.PRECISE", "ERFC.PRECISE", "EXPON.DIST",
    "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FLOOR.PRECISE", "GAMMA.DIST",
    "GAMMA.INV", "GAMMALN.PRECISE", "HYPGEOM.DIST", "LOGNORM.DIST", "LOGNORM.INV", "MODE.MULT",
    "MODE.SNGL", "NEGBINOM.DIST", "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST",
    "NORM.S.INV", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC",
    "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANK.AVG", "RANK.EQ", "STDEV.P", "STDEV.S",
    "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "VAR.P", "VAR.S",
    "WEIBULL.DIST", "WORKDAY.INTL", "Z.TEST",
    # Excel 2013
    "ACOT", "ACOTH", "ARABIC", "BASE", "BINOM.DIST.RANGE", "BITAND", "BITLSHIFT", "BITOR",
    "BITRSHIFT", "BITXOR", "CEILING.MATH", "COMBINA", "COT", "COTH", "CSC", "CSCH", "DAYS",
    "DECIMAL", "ENCODEURL", "FILTERXML", "FLOOR.MATH", "FORMULATEXT", "GAMMA", "GAUSS", "IFNA",
    "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA",
    "ISO.CEILING", "ISOWEEKNUM", "MUNIT", "NUMBERVALUE", "PDURATION", "PERMUTATIONA", "PHI",
    "RRI", "SEC", "SECH", "SHEET", "SHEETS", "SKEW.P", "UNICHAR", "UNICODE", "WEBSERVICE", "XOR",
    # Excel 2016 / 2019
    "CONCAT", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR", "IFS", "MAXIFS", "MINIFS", "SWITCH", "TEXTJOIN",
    # Excel 2021 / 365
    "ARRAYTOTEXT", "BYCOL", "BYROW", "CHOOSECOLS", "CHOOSEROWS", "DROP", "EXPAND", "HSTACK",
    "ISOMITTED", "LAMBDA", "LET", "MAKEARRAY", "MAP", "RANDARRAY", "REDUCE", "SCAN", "SEQUENCE",
    "SORTBY", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTSPLIT", "TOCOL", "TOROW", "UNIQUE",
    "VALUETOTEXT", "VSTACK", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH",
}
WORKSHEET_PREFIXED_FUNCTIONS: Set[str] = {"FILTER", "SORT"}   # stored as _xlfn._xlws.NAME

# The hand-kept list above refused real functions (YEARFRAC, MROUND, STDEV.S,
# NORM.S.DIST, AGGREGATE, ... - 534 refusals in the v2 logs). Every function
# openpyxl knows (Excel 2007's full set) and every later one is valid; a
# misspelt name is still refused.
VALID_EXCEL_FUNCTIONS |= set(_OPENPYXL_FORMULAE) | FUTURE_FUNCTIONS | WORKSHEET_PREFIXED_FUNCTIONS

# Common function name mistakes and their corrections
FUNCTION_SUGGESTIONS: Dict[str, str] = {
    "SUMPMT": "Use CUMIPMT for cumulative interest, or PMT for payment amount",
    "AVGIF": "Use AVERAGEIF",
    "AVGIFS": "Use AVERAGEIFS",
    "LOOKUP2": "Use VLOOKUP or XLOOKUP",
    "SUMAVERAGE": "Use SUM and AVERAGE separately",
    "COUNTUNIQUE": "Use SUMPRODUCT(1/COUNTIF(range,range))",
    "CONCATENATEIF": "Use TEXTJOIN with IF",
    "VLOOK": "Use VLOOKUP",
    "HLOOK": "Use HLOOKUP",
    "AVGERAGE": "Use AVERAGE (check spelling)",
    "SUMM": "Use SUM (check spelling)",
}


def extract_function_names(formula: str) -> List[str]:
    """
    Extract all function names from an Excel formula.

    Args:
        formula: Excel formula string (with or without leading '=')

    Returns:
        List of unique function names found in the formula
    """
    # Remove string literals to avoid false positives
    # Match both single and double quotes
    formula_cleaned = re.sub(r'"[^"]*"', '', formula)
    formula_cleaned = re.sub(r"'[^']*'", '', formula_cleaned)

    # Find function names (word characters followed by opening paren)
    # Pattern: word boundary, uppercase letter start, word chars, optional whitespace, opening paren
    functions = re.findall(r'\b([A-Z][A-Z0-9_.]*)\s*\(', formula_cleaned)

    return list(set(functions))  # Return unique names


def extract_potential_names(formula: str) -> List[str]:
    """
    Extract potential named ranges from formula (heuristic).

    Named ranges are typically:
    - All uppercase or mixed case
    - Not cell references (A1, B2, etc.)
    - Not worksheet references
    - Not function names

    Args:
        formula: Excel formula string

    Returns:
        List of potential named range references
    """
    # Remove string literals
    formula_cleaned = re.sub(r'"[^"]*"', '', formula)
    formula_cleaned = re.sub(r"'[^']*'", '', formula_cleaned)

    # Remove worksheet references (Sheet1!A1)
    formula_cleaned = re.sub(r"'[^']*'!", '', formula_cleaned)
    formula_cleaned = re.sub(r'\w+!', '', formula_cleaned)

    # Find words that could be named ranges
    # Pattern: uppercase letter or underscore, followed by word chars
    # Must be 3+ chars to avoid false positives on cell refs
    potential_names = re.findall(r'\b([A-Z_][A-Z0-9_]{2,})\b', formula_cleaned)

    # Filter out function names
    functions = set(extract_function_names(formula))

    # Filter out obvious cell references (A1, B2, etc.) and ranges (A1:B10)
    cell_pattern = re.compile(r'^[A-Z]{1,3}\d+$')

    names = [
        name for name in potential_names
        if name not in functions
        and name not in VALID_EXCEL_FUNCTIONS
        and not cell_pattern.match(name)
    ]

    return list(set(names))


def extract_cell_references(formula: str) -> List[str]:
    """
    Extract all cell references from an Excel formula.

    Examples:
        "=A1 + B2" → ["A1", "B2"]
        "=SUM(A1:A10)" → ["A1:A10"]
        "='Sheet1'!B1" → ["Sheet1!B1"]
        "=B1 * (B2 / 100)" → ["B1", "B2"]

    Args:
        formula: Excel formula string

    Returns:
        List of cell references found in the formula
    """
    # Remove ONLY double-quoted string literals to avoid false positives
    # DO NOT remove single quotes - they're part of worksheet references like 'Sheet Name'!A1
    formula_cleaned = re.sub(r'"[^"]*"', '', formula)

    # Pattern for cell references with optional worksheet qualifier
    # Matches: A1, B2, AA10, Sheet1!A1, 'Sheet Name'!A1, A1:B10

    # First, extract quoted worksheet references: 'Sheet Name'!A1
    quoted_refs = re.findall(r"'([^']+)'!([A-Z]+\d+(?::[A-Z]+\d+)?)", formula)
    refs_with_sheets = [f"{sheet}!{ref}" for sheet, ref in quoted_refs]

    # Extract unquoted worksheet references: Sheet1!A1
    unquoted_refs = re.findall(r'(\w+)!([A-Z]+\d+(?::[A-Z]+\d+)?)', formula_cleaned)
    refs_with_sheets.extend([f"{sheet}!{ref}" for sheet, ref in unquoted_refs])

    # Remove worksheet references from formula for plain cell extraction
    # Must match the same pattern as extraction to avoid orphaned cell refs
    formula_no_sheets = re.sub(r"'[^']+'![A-Z]+\d+(?::[A-Z]+\d+)?", '', formula_cleaned)
    formula_no_sheets = re.sub(r'\w+![A-Z]+\d+(?::[A-Z]+\d+)?', '', formula_no_sheets)

    # Extract plain cell references: A1, B2, A1:B10
    plain_refs = re.findall(r'\b([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\b', formula_no_sheets)

    # Combine all references
    all_refs = refs_with_sheets + plain_refs

    return list(set(all_refs))


def _cell_in_range(cell: str, start_cell: str, end_cell: str) -> bool:
    """
    Check if a cell falls within a given range.

    Args:
        cell: Cell address (e.g., "B2")
        start_cell: Range start (e.g., "B1")
        end_cell: Range end (e.g., "B5")

    Returns:
        bool: True if cell is within the range
    """
    import re
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    try:
        # Parse cell coordinates
        cell_col, cell_row = coordinate_from_string(cell.upper())
        start_col, start_row = coordinate_from_string(start_cell.upper())
        end_col, end_row = coordinate_from_string(end_cell.upper())

        # Convert column letters to numbers for comparison
        cell_col_num = column_index_from_string(cell_col)
        start_col_num = column_index_from_string(start_col)
        end_col_num = column_index_from_string(end_col)

        # Ensure start <= end for both row and column
        min_row, max_row = min(start_row, end_row), max(start_row, end_row)
        min_col, max_col = min(start_col_num, end_col_num), max(start_col_num, end_col_num)

        # Check if cell is within range
        return (min_row <= cell_row <= max_row and
                min_col <= cell_col_num <= max_col)

    except Exception:
        # If parsing fails, assume not in range
        return False


_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')
_CELL_REF = re.compile(
    r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?"          # optional sheet: 'My Sheet'! or Sheet1!
    r"\$?([A-Za-z]{1,3})\$?(\d+)"                     # first cell
    r"(?::\$?([A-Za-z]{1,3})\$?(\d+))?$"              # optional :second cell
)
# Functions that use only the address of their argument, never its value:
# =ROWS($A$4:A4) in A4 is a running counter, not a circular reference.
_REFERENCE_ONLY_FUNCTIONS = {"ROW", "ROWS", "COLUMN", "COLUMNS"}


def strip_string_literals(formula: str) -> str:
    """The formula with every "..." literal emptied, so text such as
    "ERROR!" or "Note (see" is not read as a sheet reference or a bracket."""
    return _STRING_LITERAL.sub('""', formula)


class CellRef(NamedTuple):
    """One cell or range operand of a formula, with its surroundings."""
    sheet: Optional[str]      # None = the formula's own sheet
    start: str                # "B1" ($ removed, upper case)
    end: Optional[str]        # "B3" for a range, None for a single cell
    function: Optional[str]   # innermost enclosing function, upper case, no prefix
    prev: Tuple[str, str]     # (type, value) of the neighbouring tokens,
    next: Tuple[str, str]     # white space skipped; ("", "") at either end


def _bare_function_name(token_value: str) -> str:
    name = token_value.rstrip("(").strip().upper()
    for prefix in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if name.startswith(prefix):
            name = name[len(prefix):]
    return name


def iter_cell_refs(formula: str) -> Iterator[CellRef]:
    """Cell and range operands of *formula*, read with openpyxl's tokenizer.

    Raises nothing: a formula the tokenizer cannot read yields no operands,
    and the callers then allow it - LibreOffice's recalculation, reported
    with every write, shows a real error anyway.
    """
    try:
        items = [tok for tok in Tokenizer(formula).items if tok.type != "WHITE-SPACE"]
    except Exception:
        return
    stack: List[Optional[str]] = []
    for idx, tok in enumerate(items):
        if tok.type == "FUNC" and tok.subtype == "OPEN":
            stack.append(_bare_function_name(tok.value))
        elif tok.type in ("PAREN", "ARRAY") and tok.subtype == "OPEN":
            stack.append(None)
        elif tok.type in ("FUNC", "PAREN", "ARRAY") and tok.subtype == "CLOSE":
            if stack:
                stack.pop()
        elif tok.type == "OPERAND" and tok.subtype == "RANGE":
            m = _CELL_REF.match(tok.value.strip())
            if not m:
                continue  # a defined name, a whole column/row, a table reference
            quoted, plain, c1, r1, c2, r2 = m.groups()
            sheet = quoted.replace("''", "'") if quoted else plain
            function = next((f for f in reversed(stack) if f), None)
            before = items[idx - 1] if idx else None
            after = items[idx + 1] if idx + 1 < len(items) else None
            yield CellRef(
                sheet=sheet,
                start=f"{c1}{r1}".upper(),
                end=f"{c2}{r2}".upper() if c2 else None,
                function=function,
                prev=(before.type, before.value) if before else ("", ""),
                next=(after.type, after.value) if after else ("", ""),
            )


def detect_circular_reference(cell: str, formula: str, worksheet: str = None) -> Dict[str, any]:
    """
    Detect if a formula refers to its own cell, directly or through a range.

    - Direct self-reference: B1 = B1 + 1, or Workings!B1 = Workings!B1 + 1
    - Range-based: B2 = SUM(B1:B3), absolute or not ($B$1:$B$3)
    A reference that names ANOTHER sheet is never circular (B2 = SUM(Sheet2!B1:B3)
    was refused until 2026-09-19 because the sheet name was ignored), and neither
    is the argument of ROW/ROWS/COLUMN/COLUMNS, which read an address, not a value.

    Args:
        cell: Target cell address (e.g., "B1", "A10")
        formula: Excel formula string
        worksheet: Optional worksheet name for context

    Returns:
        {"is_circular": bool, "error": str or None, "suggestion": str or None}
    """
    cell = cell.upper().replace("$", "")
    here = f"{worksheet}!" if worksheet else ""

    for ref in iter_cell_refs(formula):
        if ref.sheet is not None and worksheet is not None and ref.sheet.lower() != worksheet.lower():
            continue  # another sheet
        if ref.function in _REFERENCE_ONLY_FUNCTIONS:
            continue
        shown = f"{ref.sheet}!" if ref.sheet else ""
        if ref.end is None:
            if ref.start == cell:
                return {
                    "is_circular": True,
                    "error": f"CIRCULAR REFERENCE DETECTED: Cell {here}{cell} references itself directly in formula"
                             + (f" via {shown}{ref.start}" if shown else ""),
                    "suggestion": f"Remove {shown}{cell} from the formula or reference a different cell"
                                  f" (e.g., a cell on another worksheet: Assumptions!{cell})",
                }
        elif _cell_in_range(cell, ref.start, ref.end):
            return {
                "is_circular": True,
                "error": f"CIRCULAR REFERENCE DETECTED: Cell {cell} is within range {shown}{ref.start}:{ref.end} referenced in its own formula",
                "suggestion": f"Either exclude {cell} from the range {ref.start}:{ref.end} or place the formula in a different cell",
            }

    return {"is_circular": False, "error": None, "suggestion": None}


_LET_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")


def add_future_function_prefixes(formula: str) -> str:
    """*formula* as it has to be stored: post-2007 functions prefixed "_xlfn."
    (FILTER/SORT "_xlfn._xlws."), LET/LAMBDA names prefixed "_xlpm.".

    Written plainly, =XLOOKUP(...), =IFS(...), =LET(x,...) are "#NAME?" in Excel
    and in LibreOffice. Already-prefixed text is left alone; a formula the
    tokenizer cannot read is returned unchanged.
    """
    upper = formula.upper()
    if not any(name + "(" in upper for name in FUTURE_FUNCTIONS | WORKSHEET_PREFIXED_FUNCTIONS):
        return formula
    try:
        tok = Tokenizer(formula)
        items = tok.items
    except Exception:
        return formula

    # 1. names declared by LET (arguments 1, 3, 5, ... but the last) and
    #    LAMBDA (every argument but the last), to be prefixed inside that call.
    spans: List[Tuple[int, int, Set[str]]] = []
    for start, t in enumerate(items):
        if not (t.type == "FUNC" and t.subtype == "OPEN" and _bare_function_name(t.value) in ("LET", "LAMBDA")):
            continue
        depth, args, current = 1, [], []
        end = start
        for end in range(start + 1, len(items)):
            it = items[end]
            if it.subtype == "OPEN":
                depth += 1
            elif it.subtype == "CLOSE":
                depth -= 1
                if depth == 0:
                    break
            if depth == 1 and it.type == "SEP":
                args.append(current); current = []
            elif it.type != "WHITE-SPACE":
                current.append(it)
        args.append(current)
        is_let = _bare_function_name(t.value) == "LET"
        candidates = args[:-1][::2] if is_let else args[:-1]
        names = {a[0].value.lower() for a in candidates
                 if len(a) == 1 and a[0].type == "OPERAND" and a[0].subtype == "RANGE"
                 and _LET_NAME.match(a[0].value) and not _CELL_REF.match(a[0].value)}
        if names:
            spans.append((start, end, names))

    changed = False
    for idx, t in enumerate(items):
        if t.type == "FUNC" and t.subtype == "OPEN":
            raw = t.value
            if raw.lstrip().upper().startswith(("_XLFN.", "_XLWS.")):
                continue
            name = _bare_function_name(raw)
            if name in WORKSHEET_PREFIXED_FUNCTIONS:
                t.value, changed = f"_xlfn._xlws.{name}(", True
            elif name in FUTURE_FUNCTIONS:
                t.value, changed = f"_xlfn.{name}(", True
        elif t.type == "OPERAND" and t.subtype == "RANGE" and not t.value.lower().startswith("_xlpm."):
            if any(s <= idx <= e and t.value.lower() in names for s, e, names in spans):
                t.value, changed = "_xlpm." + t.value, True
    return "=" + "".join(it.value for it in items) if changed else formula


def validate_formula(formula: str, cell: str = None, worksheet: str = None) -> Dict[str, any]:
    """
    Validate an Excel formula for syntax and function name errors.

    Args:
        formula: Excel formula string (should start with '=')

    Returns:
        Dictionary with validation results:
        {
            "valid": bool,              # True if no errors
            "errors": List[str],        # List of error messages
            "warnings": List[str],      # List of warning messages
            "functions_used": List[str], # Functions found in formula
            "potential_names": List[str] # Potential named ranges
        }
    """
    errors: List[str] = []
    warnings: List[str] = []

    formula_stripped = formula.strip()

    # Check basic syntax
    if not formula_stripped.startswith('='):
        errors.append("Formula must start with '='")
        return {
            "valid": False,
            "errors": errors,
            "warnings": warnings,
            "functions_used": [],
            "potential_names": []
        }

    # Check for empty formula
    if formula_stripped == '=':
        errors.append("Formula is empty (only '=')")
        return {
            "valid": False,
            "errors": errors,
            "warnings": warnings,
            "functions_used": [],
            "potential_names": []
        }

    # Check for circular references if cell and/or worksheet provided
    if cell:
        circ_check = detect_circular_reference(cell, formula, worksheet)
        if circ_check["is_circular"]:
            errors.append(circ_check["error"])
            if circ_check["suggestion"]:
                warnings.append(f"Suggestion: {circ_check['suggestion']}")

    # Extract function names
    functions = extract_function_names(formula)

    # Check for invalid functions
    invalid_funcs = [f for f in functions if f not in VALID_EXCEL_FUNCTIONS]

    for func in invalid_funcs:
        suggestion = FUNCTION_SUGGESTIONS.get(
            func,
            "Verify this is a valid Excel function"
        )
        errors.append(f"Invalid function '{func}': {suggestion}")

    # Check for potential undefined named ranges
    potential_names = extract_potential_names(formula)

    for name in potential_names:
        warnings.append(
            f"Potential undefined name '{name}' - ensure it's defined or use cell reference instead"
        )

    # Check for common syntax issues (text literals may hold any bracket)
    unquoted = strip_string_literals(formula)
    paren_balance = unquoted.count('(') - unquoted.count(')')
    if paren_balance != 0:
        errors.append(
            f"Unbalanced parentheses: {abs(paren_balance)} {'extra opening' if paren_balance > 0 else 'extra closing'}"
        )

    bracket_balance = unquoted.count('[') - unquoted.count(']')
    if bracket_balance != 0:
        errors.append(
            f"Unbalanced brackets: {abs(bracket_balance)} {'extra opening' if bracket_balance > 0 else 'extra closing'}"
        )

    return {
        "valid": len(errors) == 0,
        "errors": errors,
        "warnings": warnings,
        "functions_used": functions,
        "potential_names": potential_names
    }


def validate_formulas_batch(formulas: List[str]) -> Dict[str, any]:
    """
    Validate multiple formulas and return aggregated results.

    Args:
        formulas: List of Excel formula strings

    Returns:
        {
            "total": int,
            "valid": int,
            "invalid": int,
            "results": List[Dict]  # Individual validation results
        }
    """
    results = [validate_formula(f) for f in formulas]

    return {
        "total": len(formulas),
        "valid": sum(1 for r in results if r["valid"]),
        "invalid": sum(1 for r in results if not r["valid"]),
        "results": results
    }
