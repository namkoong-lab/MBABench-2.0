"""Uncached-formula pre-flight (Class E of the 2026-09-01 evidence sweep).

The judge reads cached formula *results*, never live ones: extraction loads the
workbook twice (`data_only=False` for formulas, `data_only=True` for values) and
`_get_formatted_value` returns the empty string whenever the cached value is
``None``. A workbook saved without calculation — the classic openpyxl-produced
attempt — therefore reaches the judge as formulas with no values anywhere:

    [B4]1,234|FORMULA:=SUM(B1:B3)     <- calculated, value visible
    |FORMULA:=SUM(B1:B3)              <- never calculated, value gone

Nothing downstream notices. The Accuracy category (19% of the score, including
the single 10.36-point `Final calculation accuracy` check) then has no evidence
behind it, and the grade that comes out looks exactly like a real one. Under a
strictness rule ("undecided is a fail") it becomes a mechanical Accuracy zero —
scoring the attempt's *save path* rather than its quality.

This module refuses that grading instead.

Two censuses, workbook first
----------------------------
The CSV encoding cannot tell a formula that *legitimately returns the empty
string* (``=IF($A215="","",...)`` in the spare rows of a schedule) from one
that was never calculated — both print a blank display half. A workbook built
that way (2026-09-03: a 10,251-formula bond schedule with 7,221 blank-result
rows) was refused at "70% uncached" while being fully calculated.

The workbook XML does tell them apart, and every producer we see agrees:

    <c t="str"><f>..</f><v></v></c>   Excel / Excel Online / LibreOffice:
                                      calculated, result is the empty string
    <c><f>..</f><v /></c>             openpyxl: never calculated (no type,
                                      empty value element)
    <c t="n"><f>..</f><v>200</v></c>  calculated, numeric (or t="str"/"b"/"e"
                                      with content)

So when the source workbook is on disk, `check_case` decides on a *workbook*
census: a formula cell is uncached iff its ``<v>`` is absent or empty AND the
cell is not typed ``str``. The CSV census is still computed and recorded (it
measures exactly what the judge reads, and it is the only census available
when the CSVs came from cache and no workbook was staged), and it remains the
deciding basis when no workbook path is given. ``provenance[label]["basis"]``
says which one decided.

Enforcement is wired in `judge._prepare_case`, so every mode inherits it —
classic, agentic, and single-pass. Set ``JUDGE_SKIP_FORMULA_CACHE_CHECK=1`` to
grade anyway; the skip is recorded in ``scored_results.formula_cache`` so such
rows stay identifiable.

The fix, when it fires for real, is to re-extract with ``--run-calculation``
(a LibreOffice recalculation pass), or to repair the source workbook.
"""

import re
from pathlib import Path

try:
    from .logger import logger
    from .misc_utils import load_env_var
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger
    from misc_utils import load_env_var

SKIP_ENV = "JUDGE_SKIP_FORMULA_CACHE_CHECK"
DEFAULT_MAX_RATIO = 0.5

# A formula cell is uncached iff its display half is blank. Splitting on the
# first '|' is safe for that question specifically: display text may itself
# contain '|', but then the head is non-empty either way, which is the only
# thing being asked. Whitespace-only heads count as blank because
# `create_enhanced_cell_variants` tests `display_value.strip()`.
_FORMULA_MARKER = "FORMULA:"

# Worksheet XML: one <c> element per cell, self-closing or with content. The
# sheet parts are small enough (a few MB at most) to scan with a regex; no
# openpyxl load, which is the thing that loses the str-vs-untyped distinction.
# Tags may carry a namespace prefix: valid SpreadsheetML can bind the main
# namespace to a prefix (<s:c>, <s:f>, <s:v>), and 21 Codex coding workbooks
# do. Unprefixed-only patterns counted 0 formula cells in those files, so the
# workbook census passed them at ratio 0 while 7 were 80-91% uncached
# (2026-09-22).
_NS = r"(?:[A-Za-z_][\w.-]*:)?"
_CELL_RE = re.compile(rf"<{_NS}c\b([^>]*?)(?:/>|>(.*?)</{_NS}c>)", re.S)
_TYPE_RE = re.compile(r'\bt="([^"]*)"')
_VALUE_RE = re.compile(rf"<{_NS}v\b[^>]*?(?:/>|>(.*?)</{_NS}v>)", re.S)
_FORMULA_RE = re.compile(rf"<{_NS}f\b")


class FormulaCacheError(Exception):
    """A workbook's formulas carry no cached values; grading it is meaningless."""


def skip_requested() -> bool:
    import os

    return os.environ.get(SKIP_ENV) == "1"


def max_ratio() -> float:
    """Uncached fraction at or above which a workbook is refused."""
    raw = load_env_var("JUDGE_UNCACHED_FORMULA_MAX_RATIO", default=DEFAULT_MAX_RATIO)
    try:
        return float(raw)
    except (TypeError, ValueError):
        logger.warning(
            f"  Invalid JUDGE_UNCACHED_FORMULA_MAX_RATIO {raw!r}; "
            f"using {DEFAULT_MAX_RATIO}"
        )
        return DEFAULT_MAX_RATIO


# --------------------------------------------------------------- CSV census


_REF_PREFIX_RE = re.compile(r"^\[[A-Z]+\d+\]")
# Tags the extractor appends to the display half (never part of the value):
# spill anchors/children (tier 2, 2026-09-10) and what-if data tables (v7).
_DISPLAY_TAG_RE = re.compile(r"\s*\[(?:SPILL [^\]]*|SPILLED FROM [^\]]*|DATA TABLE [^\]]*|TEXT)\]")


def _census_field(field: str) -> str | None:
    """Classify one CSV field: 'cached', 'uncached', or None (not a formula)."""
    if _FORMULA_MARKER not in field:
        return None
    head = field.split("|", 1)[0]
    # Judge v7 (2026-09-09): a formula cell keeps its `[ref]` even when its
    # display is empty, so strip the address before deciding on the display.
    head = _REF_PREFIX_RE.sub("", head, count=1)
    # An uncached spill anchor / data-table cell still carries its tag; the
    # tag is not a value, so strip it too before deciding.
    head = _DISPLAY_TAG_RE.sub("", head)
    return "cached" if head.strip() else "uncached"


def census_csv_dir(csv_dir) -> dict:
    """Count formula cells and uncached formula cells across a workbook's CSVs.

    Reads the `*_full.csv` views only — they are the canonical encoding and are
    present even in pre-revision caches that lack the `*_data.csv` siblings.
    """
    import csv as csv_mod

    result = {"sheets": {}, "formula_cells": 0, "uncached_formula_cells": 0}
    if not csv_dir:
        return result
    directory = Path(csv_dir)
    if not directory.is_dir():
        return result

    for path in sorted(directory.glob("*_full.csv")):
        formulas = uncached = 0
        try:
            with open(path, "r", encoding="utf-8", newline="") as f:
                for row in csv_mod.reader(f):
                    for field in row:
                        kind = _census_field(field)
                        if kind is None:
                            continue
                        formulas += 1
                        if kind == "uncached":
                            uncached += 1
        except (OSError, UnicodeDecodeError, csv_mod.Error) as e:
            # A sheet we cannot read is not evidence of a cache problem; say so
            # rather than silently scoring it as clean.
            logger.warning(f"  [formula_cache] could not census {path.name}: {e}")
            result["sheets"][path.name] = {"unreadable": True}
            continue
        if formulas:
            result["sheets"][path.name] = {
                "formula_cells": formulas,
                "uncached_formula_cells": uncached,
            }
        result["formula_cells"] += formulas
        result["uncached_formula_cells"] += uncached

    result["ratio"] = (
        result["uncached_formula_cells"] / result["formula_cells"]
        if result["formula_cells"]
        else 0.0
    )
    return result


# ---------------------------------------------------------- workbook census


def _census_cell_xml(attrs: str, body: str | None) -> str | None:
    """Classify one <c> element: 'cached', 'empty_string', 'uncached', or
    None when the cell holds no formula."""
    if not body or not _FORMULA_RE.search(body):
        return None
    m = _VALUE_RE.search(body)
    has_value = m is not None and bool((m.group(1) or "").strip())
    if has_value:
        return "cached"
    t = _TYPE_RE.search(attrs)
    if t and t.group(1) == "str":
        # Excel / LibreOffice: a calculated formula whose result is "".
        return "empty_string"
    return "uncached"


def census_workbook(xlsx_path) -> dict:
    """Count formula cells, uncached formula cells and empty-string results by
    reading the worksheet XML parts directly.

    Returns ``{"available": False, ...}`` (with a reason) when the workbook is
    missing or unreadable, so callers fall back to the CSV census instead of
    treating an unreadable file as clean.
    """
    import zipfile

    result = {
        "available": False,
        "sheets": {},
        "formula_cells": 0,
        "uncached_formula_cells": 0,
        "empty_string_results": 0,
    }
    if not xlsx_path:
        result["reason"] = "no workbook path"
        return result
    path = Path(xlsx_path)
    if not path.is_file():
        result["reason"] = f"not a file: {path}"
        return result

    try:
        with zipfile.ZipFile(path) as z:
            parts = sorted(
                n
                for n in z.namelist()
                if n.startswith("xl/worksheets/") and n.endswith(".xml")
            )
            for part in parts:
                xml = z.read(part).decode("utf-8", errors="replace")
                formulas = uncached = empties = 0
                for m in _CELL_RE.finditer(xml):
                    kind = _census_cell_xml(m.group(1), m.group(2))
                    if kind is None:
                        continue
                    formulas += 1
                    if kind == "uncached":
                        uncached += 1
                    elif kind == "empty_string":
                        empties += 1
                if formulas:
                    result["sheets"][part.rsplit("/", 1)[-1]] = {
                        "formula_cells": formulas,
                        "uncached_formula_cells": uncached,
                        "empty_string_results": empties,
                    }
                result["formula_cells"] += formulas
                result["uncached_formula_cells"] += uncached
                result["empty_string_results"] += empties
    except (OSError, zipfile.BadZipFile, KeyError, UnicodeDecodeError) as e:
        logger.warning(f"  [formula_cache] could not census workbook {path.name}: {e}")
        result["reason"] = f"unreadable: {e}"
        return result

    result["available"] = True
    result["ratio"] = (
        result["uncached_formula_cells"] / result["formula_cells"]
        if result["formula_cells"]
        else 0.0
    )
    return result


# ------------------------------------------------------------------ decision


def _census_label(csv_dir, xlsx_path) -> dict:
    """Both censuses for one workbook, plus the deciding numbers.

    Top-level `formula_cells` / `uncached_formula_cells` / `ratio` are the
    DECIDING figures (workbook when available, else CSV); the raw censuses sit
    under `csv` and `workbook` so a row's provenance shows both.
    """
    csv_census = census_csv_dir(csv_dir)
    wb_census = census_workbook(xlsx_path)
    if wb_census["available"]:
        basis, deciding = "workbook", wb_census
    else:
        basis, deciding = "csv", csv_census
    return {
        "basis": basis,
        "formula_cells": deciding["formula_cells"],
        "uncached_formula_cells": deciding["uncached_formula_cells"],
        "ratio": deciding.get("ratio", 0.0),
        "csv": csv_census,
        "workbook": wb_census,
    }


def _describe(label: str, census: dict) -> str:
    text = (
        f"{label}: {census['uncached_formula_cells']} of "
        f"{census['formula_cells']} formula cells have no cached value "
        f"({census['ratio'] * 100:.0f}%, {census['basis']} census"
    )
    wb = census.get("workbook") or {}
    if wb.get("available") and wb.get("empty_string_results"):
        text += f"; {wb['empty_string_results']} calculated empty-string results"
    return text + ")"


def check_case(
    attempt_csv_dir,
    solution_csv_dir=None,
    attempt_xlsx=None,
    solution_xlsx=None,
) -> dict:
    """Census both workbooks and refuse the grading if either is uncalculated.

    Pass the staged ``.xlsx`` paths when they exist so the decision rests on the
    workbook census (which recognises calculated empty-string results); the CSV
    census decides only when no workbook is available.

    Returns provenance for `scored_results.formula_cache`. Raises
    FormulaCacheError when a census is at or above the configured ratio and the
    escape hatch is not set.
    """
    provenance = {
        "checked": True,
        "max_ratio": max_ratio(),
        "attempt": _census_label(attempt_csv_dir, attempt_xlsx),
        "solution": _census_label(solution_csv_dir, solution_xlsx),
    }

    offenders = [
        (label, provenance[label])
        for label in ("attempt", "solution")
        if provenance[label]["formula_cells"]
        and provenance[label]["ratio"] >= provenance["max_ratio"]
    ]
    provenance["offenders"] = [label for label, _ in offenders]

    for label in ("attempt", "solution"):
        c = provenance[label]
        if c["formula_cells"]:
            logger.info(f"  [formula_cache] {_describe(label, c)}")

    if not offenders:
        return provenance

    detail = "; ".join(_describe(label, c) for label, c in offenders)
    if skip_requested():
        provenance["skipped_via_env"] = True
        logger.warning(
            f"  [formula_cache] {SKIP_ENV}=1 — grading anyway despite: {detail}. "
            f"Accuracy verdicts from this grading are not evidence-based."
        )
        return provenance

    raise FormulaCacheError(
        f"{detail}. The judge reads cached results, so these formulas reach it "
        f"with no values and Accuracy cannot be graded from evidence. "
        f"Re-run with --run-calculation to recalculate through LibreOffice, or "
        f"repair the source workbook. Set {SKIP_ENV}=1 to grade anyway "
        f"(recorded in scored_results.formula_cache)."
    )
