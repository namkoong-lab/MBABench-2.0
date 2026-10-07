"""The answer-equivalence rulebook — ONE source, TWO consumers (judge v6).

`compare()` is what the harness answer checker (utils/answer_check.py) runs
on every Questions-sheet answer, and `render_rules_text()` is rendered
VERBATIM into the single-pass judge prompt (template_8), so the deterministic
engine and the LLM judge apply the same definition of "the same number" and
cannot drift apart.

The rules remove *noise*, not *standards*: a genuinely different value is a
mismatch no matter how it is dressed. Mined 2026-09-02 from 394 cited
"Final calculation accuracy" mistakes across the deprecated grading corpus
(tolerance / sign / percent-form / display / missing classes all attested).

Rules (numbering matches render_rules_text and the v6 spec §3):
  1. Tolerance is scale-aware: max(1e-6 relative noise band, rounding band).
     The rounding band is half a unit of the last decimal the golden carries
     — i.e. the attempt is THE SAME number when it ROUNDS TO the golden (the
     golden is a rounded figure; an unrounded attempt that rounds to it holds
     the same underlying value) — and a full unit only when BOTH sides are
     rounded figures (two correct roundings can sit one notch apart at a
     boundary). v6.4 (maintainer 2026-09-02): the band no longer requires the
     attempt to be rounded — 22 of 454 attempts failed the binary check on
     presentation alone; not rounding as instructed is instead judged
     under Rounding / Rounded outputs (by the judge — the v6.4 harness
     verdict for that check was retired in v6.6). The decimals
     the comparison runs at come from the golden's own ROUND(...,n) when
     present (the key's real precision — the 4-dp tasks store ROUND(x,4) on
     a %-formatted fraction, not 6 places), else from the header phrase.
  2. Sign convention, GUARDED by an outflow lexicon on the question text:
     |a| == |b| within tolerance is accepted ONLY for such rows. One-off
     flips on every other row stay failures — for a "difference" question
     the sign IS the answer. (maintainer 2026-09-02: lexicon only; the earlier
     "consistent across the block" clause was dropped as too loose.) v6.4:
     the core lexicon (expense/cost/spend/outflow/depreciation/amortization/
     capex/tax) is unconditional; an EXTENDED list of P&L line items the
     454-attempt sweep showed agents presenting negative by convention
     (raw materials, energy, SG&A/G&A, wages, R&D, marketing, interest
     (loan), repayments... — 219 flagged flips over 19 labels) is accepted
     only when no INFLOW_GUARD word (income/revenue/net/change/difference/
     growth/asset/value...) marks the row as a net or inflow quantity.
     v6.5 (maintainer 2026-09-10): loan-schedule "payment" / "principal" rows
     join the extended list — House Standards attempts sign them negative
     while the goldens are positive; "balance" rows stay guarded.
  3. Percent x100 / fraction form: when the Unit column says %, the question
     text says percent/rate/margin/..., the golden cell is %-formatted, or
     the attempt wrote a literal "42%": a == 100b or b == 100a is equal
     (within tolerance, applied on the golden's scale).
  4. Values, never rendered strings: parentheses = minus, $/commas/%
     stripped, "(480,051.30)" parses to -480051.30.
  5. Sentinel text synonyms ("Outside model horizon" ~ "N/A" ~ "beyond
     horizon") compare equal to each other — text-vs-text only.
  6. Numeric-as-text ("42", "$1,234.00") parses before comparing.
  7. Dates: ISO / US / Excel serial are equal when they name the same day.
  8. Zero forms: 0 / 0.0 / accounting "-" equal zero. A BLANK cell is
     "unanswered" — completeness, not equivalence.
  9. Units scale (x1000 / x1e6) is NEVER auto-accepted (maintainer): flagged
     as `possible_unit_scale_difference` for human review; verdict stays
     mismatch.

Hardcoded answers are detected by the checker (formula-vs-constant on the
answer cell), not here; the rulebook text describes the standard so the
judge applies it identically.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time as dt_time
from typing import Any, Optional

ABS_TOL = 1e-9
REL_TOL = 1e-6

# Rule 2 — question-text lexicon under which a sign flip is a convention,
# not an error. Word-boundary matched, case-insensitive. The CORE list is
# the maintainer's (2026-09-02): rows whose quantity is unambiguously an outflow,
# accepted unconditionally.
OUTFLOW_LEXICON_CORE = (
    "expense", "expenses", "cost", "costs", "spend", "spending", "spent",
    "outflow", "outflows", "depreciation", "amortization", "amortisation",
    "capex", "capital expenditure", "capital expenditures", "tax", "taxes",
)
# v6.4 EXTENDED list: the P&L line items the 454-attempt sweep showed agents
# presenting negative by convention (Fixings / FiveAutomotive / PlasticParts
# / PastaInc "amount for/of <line item>" rows). Accepted only when the
# INFLOW_GUARD below does not veto the row.
OUTFLOW_LEXICON_EXTENDED = (
    "opex", "cogs", "cost of goods sold", "raw material", "raw materials",
    "materials", "energy", "utilities", "rent", "wages", "salaries", "salary",
    "payroll", "sg&a", "g&a", "r&d", "research & development",
    "research and development", "marketing", "selling", "advertising",
    "logistics", "distribution", "legal", "compliance", "it & infrastructure",
    "insurance", "maintenance", "repairs", "fees", "purchases", "interest",
    "interest expense", "interest paid", "repayment", "repayments",
    "dividends paid", "dividend paid",
    # v6.5 (maintainer 2026-09-10): loan-schedule rows. Attempts built under the
    # House Standards ("costs and outflows negative, everywhere") sign the
    # scheduled payment and principal negative while the goldens are
    # positive; the magnitude is the answer. "Balance" stays guarded.
    "payment", "payments", "payment due", "principal", "principal amount",
)
OUTFLOW_LEXICON = OUTFLOW_LEXICON_CORE + OUTFLOW_LEXICON_EXTENDED
# Rule 2 guard (v6.4) — an EXTENDED-list hit is vetoed when the row is a net
# / inflow / balance / difference quantity, where the sign IS the answer
# ("Interest income", "Changes in Net Working Capital", "Intangible Assets
# growth"). Core words ("Income tax") are never vetoed.
INFLOW_GUARD = (
    "income", "revenue", "revenues", "inflow", "inflows", "receipt", "receipts",
    "proceeds", "received", "earned", "net", "change", "changes", "difference",
    "growth", "balance", "asset", "assets", "equity", "value", "margin",
    "profit", "ebit", "ebitda", "cash flow", "cashflow", "free cash",
)

# Rule 3 — question-text cues that the quantity is a percentage / ratio.
PERCENT_LEXICON = (
    "percent", "percentage", "%", "ratio", "rate", "margin", "irr", "yield",
    "growth", "return", "proportion", "cagr", "roi", "roe", "roa", "wacc",
    "share of",
)

# Rule 5 — attempt-side sentinel synonyms, all canonicalised to "n/a".
SENTINEL_SYNONYMS = {
    "n/a", "na", "n.a.", "n.a", "not applicable", "not available",
    "outside model horizon", "outside the model horizon", "beyond horizon",
    "beyond model horizon", "beyond the model horizon", "outside horizon",
    "out of horizon", "out of range", "not in model horizon",
    "outside of model horizon",
}

# Rule 8 — accounting zero renderings.
ZERO_FORMS = {"-", "--", "—", "–", "−"}

# Rule 9 — scale factors that are flagged, never accepted.
UNIT_SCALE_FACTORS = (1_000.0, 1_000_000.0, 1_000_000_000.0)

_WORD_NUMBERS = {
    "zero": 0, "no": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8,
}

_NUMERIC_TEXT_RE = re.compile(
    r"""^\s*
        (?P<open>\()?\s*
        (?P<sign>[-−–+])?\s*
        (?P<cur>[$€£¥])?\s*
        (?P<sign2>[-−–+])?\s*
        (?P<num>(\d{1,3}(,\d{3})+|\d+)(\.\d+)?|\.\d+)
        \s*(?P<pct>%)?\s*
        (?P<close>\))?\s*
        (?P<suffix>x|X|bps|pp)?\s*$""",
    re.VERBOSE,
)
_ISO_DATE_RE = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]00:00(?::00)?)?\s*$")
_US_DATE_RE = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")
_ROUND_RE = re.compile(r"ROUND\s*\(", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------


def norm_text(s: Any) -> str:
    """Casefold, collapse whitespace, unify quotes/dashes, strip edge punctuation."""
    if s is None:
        return ""
    t = str(s)
    t = (
        t.replace("’", "'").replace("‘", "'")
        .replace("“", '"').replace("”", '"')
        .replace("–", "-").replace("—", "-").replace("−", "-")
        .replace(" ", " ")
    )
    t = re.sub(r"\s+", " ", t).strip().casefold()
    return t.strip(" .:;")


def canonical_sentinel(s: str) -> Optional[str]:
    """'n/a' when *s* is a recognised sentinel synonym, else None."""
    t = norm_text(s).rstrip(".")
    return "n/a" if t in SENTINEL_SYNONYMS else None


# ---------------------------------------------------------------------------
# Scalar parsing (rules 4, 6, 7, 8)
# ---------------------------------------------------------------------------


@dataclass
class Scalar:
    kind: str                 # "empty" | "number" | "text" | "date" | "bool"
    value: Any = None
    pct_literal: bool = False  # the source text carried an explicit "%"
    raw: Any = None


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def parse_scalar(v: Any) -> Scalar:
    """Classify one cell value for comparison (rules 4/6/7/8)."""
    if v is None:
        return Scalar("empty", raw=v)
    if isinstance(v, bool):
        return Scalar("bool", v, raw=v)
    if _is_number(v):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return Scalar("text", str(v), raw=v)
        return Scalar("number", float(v), raw=v)
    if isinstance(v, datetime):
        return Scalar("date", v.date(), raw=v)
    if isinstance(v, date):
        return Scalar("date", v, raw=v)
    if isinstance(v, dt_time):
        return Scalar("text", v.isoformat(), raw=v)
    s = str(v)
    if s.strip() == "":
        return Scalar("empty", raw=v)
    st = s.strip()
    if st in ZERO_FORMS:
        return Scalar("number", 0.0, raw=v)
    m = _NUMERIC_TEXT_RE.match(st)
    if m:
        num = float(m.group("num").replace(",", ""))
        negative = bool(m.group("open") and m.group("close"))
        sign = m.group("sign") or m.group("sign2")
        if sign and sign != "+":
            negative = True  # "-" (inside or outside parentheses) is negative
        if negative:
            num = -num
        return Scalar("number", num, pct_literal=bool(m.group("pct")), raw=v)
    m = _ISO_DATE_RE.match(st)
    if m:
        try:
            return Scalar("date", date(int(m.group(1)), int(m.group(2)), int(m.group(3))), raw=v)
        except ValueError:
            pass
    m = _US_DATE_RE.match(st)
    if m:
        try:
            return Scalar("date", date(int(m.group(3)), int(m.group(1)), int(m.group(2))), raw=v)
        except ValueError:
            pass
    low = st.casefold()
    if low in ("yes", "true"):
        return Scalar("bool", True, raw=v)
    if low in ("no", "false"):
        return Scalar("bool", False, raw=v)
    return Scalar("text", st, raw=v)


def excel_serial_to_date(n: float) -> Optional[date]:
    """Excel 1900-system serial -> date (whole-day serials only)."""
    try:
        from openpyxl.utils.datetime import from_excel

        d = from_excel(n)
        if isinstance(d, datetime):
            return d.date()
        if isinstance(d, date):
            return d
    except Exception:  # noqa: BLE001 — best-effort conversion
        return None
    return None


# ---------------------------------------------------------------------------
# Precision directives (rule 1)
# ---------------------------------------------------------------------------


@dataclass
class Precision:
    dp: Optional[int]         # decimal places requested; None = no directive
    scale: str = "value"      # "value" | "percent" (dp applies to the % rendering)
    source: str = "none"      # "header" | "round" | "format" | "none"


def parse_precision_directive(header_text: Any) -> Precision:
    """Read the Questions-sheet header phrase into a Precision.

    Understands "round your answers to two decimal places", "four decimal
    places", "whole numbers", and a task-68 style "Format return cells as
    0.00%" (2 dp on the percent rendering).
    """
    t = norm_text(header_text)
    if not t:
        return Precision(None)
    m = re.search(r"(\d+|zero|one|two|three|four|five|six|seven|eight)\s+decimal\s+place", t)
    if m:
        tok = m.group(1)
        dp = int(tok) if tok.isdigit() else _WORD_NUMBERS[tok]
        return Precision(dp, "value", "header")
    if re.search(r"whole\s+number|nearest\s+(whole|integer|unit)", t):
        return Precision(0, "value", "header")
    m = re.search(r"0\.(0+)%", t)
    if m:
        return Precision(len(m.group(1)), "percent", "header")
    if re.search(r"\b0%", t):
        return Precision(0, "percent", "header")
    return Precision(None)


def formula_text(formula: Any) -> Optional[str]:
    """Formula text from a cell value: a '=...' string, or the `.text` of an
    openpyxl ArrayFormula / DataTableFormula object. Every golden answer
    formula is an ArrayFormula (dynamic-array XLOOKUP/LET wrappers), so
    reading only strings saw no ROUND() in any golden until v6.4."""
    if isinstance(formula, str):
        return formula
    t = getattr(formula, "text", None)
    return t if isinstance(t, str) else None


_ROUND_CALL_RE = re.compile(r"(?<![A-Z])ROUND\s*\(", re.IGNORECASE)


def round_dp_from_formula(formula: Any) -> Optional[int]:
    """The digit count of the OUTERMOST (leftmost) ROUND(...,n) in a formula,
    if any: walks to that call's matching parenthesis and reads its last
    argument, so LET()/IFNA()/INDEX() wrappers around the ROUND are fine.
    ROUNDUP/ROUNDDOWN/MROUND are not ROUND."""
    text = formula_text(formula)
    if not text:
        return None
    for m in _ROUND_CALL_RE.finditer(text):
        depth, j, start, args = 1, m.end(), m.end(), []
        while j < len(text) and depth:
            ch = text[j]
            if ch == '"':                       # skip string literals
                k = text.find('"', j + 1)
                j = len(text) if k == -1 else k
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    args.append(text[start:j])
            elif ch == "," and depth == 1:
                args.append(text[start:j])
                start = j + 1
            j += 1
        if depth == 0 and len(args) >= 2:
            mm = re.fullmatch(r"\s*(-?\d+)\s*", args[-1])
            if mm:
                return int(mm.group(1))
    return None


def percent_format(number_format: Any) -> bool:
    return isinstance(number_format, str) and "%" in number_format


def effective_decimals(precision: Precision, expected_pct_format: bool) -> Optional[int]:
    """Decimal places the directive means on the STORED value: a %-formatted
    fraction (0.4213 shown as 42.13%) carries two more than the header says."""
    if precision.dp is None:
        return None
    if precision.scale == "percent" or (precision.scale == "value" and expected_pct_format):
        return precision.dp + 2
    return precision.dp


def is_rounded_to(value: float, decimals: int) -> bool:
    """True when *value* carries no more than *decimals* decimal places, up to
    float representation noise — i.e. the agent actually rounded. The slack
    is relative 1e-12 of the scaled value (v6.4: the earlier 1e-6 slack let
    any value above ~10,000 count as rounded, which hid unrounded answers
    from the `n_unrounded` audit statistic)."""
    scaled = value * (10.0 ** decimals)
    return abs(scaled - round(scaled)) <= max(1e-9, abs(scaled) * 1e-12)


# Relative band beyond which a rounding allowance is "coarse" (flagged).
COARSE_ROUNDING_REL = 0.01


def tolerance_for(expected: float, got: float, precision: Precision,
                  expected_pct_format: bool, got_rounded: bool = True,
                  golden_rounded: bool = False, decimals: Optional[int] = None,
                  source: Optional[str] = None) -> tuple[float, str]:
    """Absolute tolerance for one comparison plus a label of what set it.

    Scale-aware by construction (maintainer 2026-09-02: "if the answer is 0.5,
    a .005 delta is a lot; if it is 5,000 it is not"):

      tolerance = max( relative noise band, rounding band )

      relative noise band = max(1e-9, 1e-6 * max(|a|,|b|)) — always applies;
          on a six-figure answer this alone forgives a penny.
      rounding band — half a unit of the last decimal the golden carries
          (`decimals`, from the golden's ROUND(...,n) else the header): an
          attempt within it ROUNDS TO the golden, so it holds the same
          underlying value whether or not the agent rounded (v6.4 — the
          old "unrounded attempt gets only the noise band" clause failed
          correct models on presentation; rounding compliance is judged
          under Rounding / Rounded outputs instead). A full unit only when
          BOTH the golden and the attempt are rounded figures (two correct
          roundings can sit one notch apart at a boundary).
    """
    noise = max(ABS_TOL, REL_TOL * max(abs(expected), abs(got)))
    dec = decimals if decimals is not None else effective_decimals(precision, expected_pct_format)
    if dec is None:
        return noise, "global_fallback"
    unit = 10.0 ** -dec
    both_rounded = bool(golden_rounded and got_rounded)
    band = unit if both_rounded else unit / 2.0
    # A hair of float slack keeps 0.005 vs 0.0050000001 from flipping.
    band *= 1 + 1e-6
    src = source or precision.source
    label = f"{src}_dp{dec}_{'full' if both_rounded else 'half'}_unit"
    if noise >= band:
        return noise, f"{label}+noise"
    return band, label


def numbers_equal(a: float, b: float, tol_abs: float) -> bool:
    return abs(a - b) <= tol_abs


# ---------------------------------------------------------------------------
# Context + comparison (rules 1-9)
# ---------------------------------------------------------------------------


@dataclass
class AnswerContext:
    label: str = ""                       # the question text
    unit: Any = None                      # Unit-column text, if any
    precision: Precision = field(default_factory=lambda: Precision(None))
    expected_formula: Any = None          # golden formula text (for ROUND)
    expected_number_format: Any = None    # golden cell number format
    percent_directive: bool = False       # sheet header declares %-as-decimal
    got_formula: Any = None               # attempt formula text (ROUND evidence)


def _word_hit(words, t: str) -> bool:
    return any(re.search(rf"(?<![a-z]){re.escape(w)}(?![a-z])", t) for w in words)


def is_outflow_label(label: Any) -> bool:
    """Core lexicon: unconditional. Extended lexicon (v6.4): only when no
    INFLOW_GUARD word marks the row as a net / inflow / balance quantity."""
    t = norm_text(label)
    if _word_hit(OUTFLOW_LEXICON_CORE, t):
        return True
    return _word_hit(OUTFLOW_LEXICON_EXTENDED, t) and not _word_hit(INFLOW_GUARD, t)


def is_percent_context(ctx: AnswerContext, got: Scalar) -> tuple[bool, str]:
    if got.pct_literal:
        return True, "literal_percent_sign"
    u = norm_text(ctx.unit)
    if u and ("%" in u or "percent" in u or u in ("pct", "pp", "bps")):
        return True, "unit_column"
    if percent_format(ctx.expected_number_format):
        return True, "golden_percent_format"
    if ctx.percent_directive:
        return True, "header_directive"
    t = norm_text(ctx.label)
    if any(re.search(rf"(?<![a-z]){re.escape(w)}(?![a-z])", t) if w != "%" else "%" in t
           for w in PERCENT_LEXICON):
        return True, "question_text"
    return False, ""


def effective_precision(ctx: AnswerContext) -> Precision:
    """Header directive first, then a per-answer ROUND() in the golden formula."""
    if ctx.precision.dp is not None:
        return ctx.precision
    n = round_dp_from_formula(ctx.expected_formula)
    if n is not None:
        return Precision(n, "value", "round")
    return Precision(None)


def compare(expected: Any, got: Any, ctx: Optional[AnswerContext] = None) -> dict:
    """Compare one golden answer against one attempt answer.

    Returns {verdict: match|mismatch|missing, rule: str|None, flags: [...],
             detail: str, expected_parsed, got_parsed, tolerance, abs_delta}.
    Never raises.
    """
    ctx = ctx or AnswerContext()
    e = parse_scalar(expected)
    g = parse_scalar(got)
    out = {
        "verdict": None,
        "rule": None,
        "flags": [],
        "detail": "",
        "expected_parsed": e.value if e.kind != "date" else e.value.isoformat(),
        "got_parsed": g.value if g.kind != "date" else g.value.isoformat(),
        "expected_kind": e.kind,
        "got_kind": g.kind,
        "tolerance": None,
        "abs_delta": None,
    }
    if g.kind == "empty":
        out["verdict"] = "missing"
        out["detail"] = "answer cell is blank (unanswered)"
        if e.kind == "empty":
            out["flags"].append("expected_also_blank")
        return out
    if e.kind == "empty":
        out["verdict"] = "missing"
        out["flags"].append("missing_expected")
        out["detail"] = "golden answer cell is blank"
        return out

    # --- numbers ---------------------------------------------------------
    if e.kind == "number" and g.kind == "number":
        return _compare_numbers(e, g, ctx, out)

    # --- dates (rule 7) ---------------------------------------------------
    if e.kind == "date" or g.kind == "date":
        ed = e.value if e.kind == "date" else (
            excel_serial_to_date(e.value) if e.kind == "number" else None)
        gd = g.value if g.kind == "date" else (
            excel_serial_to_date(g.value) if g.kind == "number" else None)
        if ed and gd and ed == gd:
            out.update(verdict="match", rule="date_form", detail="same calendar day")
        else:
            out.update(verdict="mismatch", rule=None,
                       detail=f"dates differ ({ed} vs {gd})" if ed and gd
                       else "date vs non-date")
        return out

    # --- booleans (Yes/No) -----------------------------------------------
    if e.kind == "bool" or g.kind == "bool":
        eb = e.value if e.kind == "bool" else _text_to_bool(e.value)
        gb = g.value if g.kind == "bool" else _text_to_bool(g.value)
        if eb is not None and gb is not None:
            if eb == gb:
                out.update(verdict="match", rule="exact", detail="same yes/no")
            else:
                out.update(verdict="mismatch", detail="opposite yes/no")
            return out
        out.update(verdict="mismatch", flags=["type_mismatch"],
                   detail=f"{e.kind} vs {g.kind}")
        return out

    # --- text vs text (rule 5) -------------------------------------------
    if e.kind == "text" and g.kind == "text":
        es, gs = canonical_sentinel(e.value), canonical_sentinel(g.value)
        if es and gs:
            out.update(verdict="match", rule="sentinel_synonym", detail="both n/a-class")
            return out
        if norm_text(e.value) == norm_text(g.value):
            out.update(verdict="match", rule="exact", detail="same text")
        else:
            out.update(verdict="mismatch", detail="text differs")
        return out

    # --- mixed kinds ------------------------------------------------------
    if e.kind == "number" and g.kind == "text":
        if canonical_sentinel(g.value):
            out.update(verdict="mismatch", flags=["sentinel_for_numeric"],
                       detail="attempt answered n/a-class text where a number was expected")
        else:
            out.update(verdict="mismatch", flags=["type_mismatch"],
                       detail="text where a number was expected")
        return out
    out.update(verdict="mismatch", flags=["type_mismatch"],
               detail=f"{e.kind} vs {g.kind}")
    return out


def _text_to_bool(v) -> Optional[bool]:
    t = norm_text(v)
    if t in ("yes", "y", "true"):
        return True
    if t in ("no", "n", "false"):
        return False
    return None


def stored_decimals(ctx: AnswerContext, precision: Precision,
                    e_pct_fmt: bool) -> tuple[Optional[int], str]:
    """Decimal places the comparison runs at, on the STORED value, and the
    source label. The golden's own ROUND(...,n) is authoritative when present
    — it IS the precision the key carries (the 4-dp tasks store ROUND(x,4)
    on a %-formatted fraction, so 6 places would hold attempts to a precision
    the key itself lacks); otherwise the header directive, with the
    percent-rendering adjustment."""
    n = round_dp_from_formula(ctx.expected_formula)
    if n is not None:
        return n, "round"
    return effective_decimals(precision, e_pct_fmt), precision.source


def _attempt_rounded(value: float, dec: Optional[int], ctx: AnswerContext) -> Optional[bool]:
    """Did the attempt round to the requested places? Value-based (no more
    decimals than requested, in whichever unit the agent answered) or a
    ROUND(...,n<=requested) in its formula. None when nothing was requested."""
    if dec is None:
        return None
    if is_rounded_to(value, dec):
        return True
    n = round_dp_from_formula(ctx.got_formula)
    return n is not None and n <= dec


def _tol(a: float, b: float, ctx: AnswerContext, precision: Precision,
         e_pct_fmt: bool) -> tuple[float, str]:
    dec, src = stored_decimals(ctx, precision, e_pct_fmt)
    return tolerance_for(
        a, b, precision, e_pct_fmt,
        got_rounded=bool(_attempt_rounded(b, dec, ctx)),
        golden_rounded=round_dp_from_formula(ctx.expected_formula) is not None,
        decimals=dec, source=src,
    )


def _coarse(a: float, tol: float, tol_src: str, out: dict) -> None:
    """Flag a match that leaned on a rounding band wider than 1% of the answer."""
    if "unit" in tol_src and a != 0 and tol / abs(a) > COARSE_ROUNDING_REL:
        out["flags"].append("coarse_rounding")


def _compare_numbers(e: Scalar, g: Scalar, ctx: AnswerContext, out: dict) -> dict:
    a, b = float(e.value), float(g.value)
    precision = effective_precision(ctx)
    e_pct_fmt = percent_format(ctx.expected_number_format)
    tol, tol_src = _tol(a, b, ctx, precision, e_pct_fmt)
    out["tolerance"] = tol
    out["tolerance_source"] = tol_src
    out["abs_delta"] = abs(a - b)
    # Rounding compliance (v6.4): recorded per answer as an audit statistic
    # (v6.6: no longer a harness verdict); never changes the equivalence verdict.
    dec, dec_src = stored_decimals(ctx, precision, e_pct_fmt)
    out["requested_decimals"] = dec
    out["decimals_source"] = dec_src
    out["attempt_rounded"] = _attempt_rounded(b, dec, ctx)

    if a == b:
        out.update(verdict="match", rule="exact", detail="identical values")
        return out
    if numbers_equal(a, b, tol):
        out.update(verdict="match", rule="tolerance",
                   detail=f"|delta| {abs(a - b):.6g} <= tolerance {tol:.6g} ({tol_src})")
        _coarse(a, tol, tol_src, out)
        return out

    pct_ctx, pct_why = is_percent_context(ctx, g)
    outflow = is_outflow_label(ctx.label)

    # Rule 3 — percent form, tolerance measured on the golden's scale.
    if pct_ctx:
        for factor, form in ((100.0, "attempt_x100"), (0.01, "attempt_/100")):
            b2 = b * factor
            tol2, src2 = _tol(a, b2, ctx, precision, e_pct_fmt)
            if numbers_equal(a, b2, tol2):
                out.update(verdict="match", rule="percent_form",
                           detail=f"{form} equals golden ({pct_why})")
                _coarse(a, tol2, src2, out)
                return out
            if outflow and numbers_equal(abs(a), abs(b2), tol2):
                out.update(verdict="match", rule="sign_outflow+percent_form",
                           detail=f"{form}, |value| equal on an outflow row ({pct_why})")
                return out

    # Rule 2 — sign convention, outflow rows only.
    if outflow and numbers_equal(abs(a), abs(b), tol):
        out.update(verdict="match", rule="sign_outflow",
                   detail="|value| equal; question names an outflow quantity")
        return out
    if numbers_equal(abs(a), abs(b), tol):
        out["flags"].append("sign_flip_not_outflow")

    # Rule 9 — flag scale differences, never accept them.
    if a != 0 and b != 0:
        ratio = abs(a / b)
        for f in UNIT_SCALE_FACTORS:
            if abs(ratio - f) <= f * 1e-6 or abs(ratio - 1 / f) <= (1 / f) * 1e-6:
                out["flags"].append("possible_unit_scale_difference")
                break

    out.update(verdict="mismatch",
               detail=f"|delta| {abs(a - b):.6g} > tolerance {tol:.6g} ({tol_src})")
    return out


# ---------------------------------------------------------------------------
# Prompt rendering (the second consumer)
# ---------------------------------------------------------------------------

# v6.4 (maintainer 2026-09-02, after the 454-attempt sweep): rounds-to-golden
# tolerance (attempt need not be rounded), golden ROUND() read from array
# formulas and authoritative for the decimals, outflow lexicon extended with
# an inflow guard, rounding compliance moved to Rounding / Rounded outputs.
# v6.3: prompt text trimmed (maintainer 2026-09-02); rules unchanged.
# v6.6 (2026-09-14, toy-reliability walkthrough): the harness verdict for
# Rounding / Rounded outputs is retired — the judge decides it from every
# final output a reader sees, and a number format counts as rounding. Rule
# 11 now only keeps rounding out of Accuracy; the per-answer rounding
# statistics stay recorded for audit.
# v6.7 (2026-10-06, judge v14): alternate accepted answers. The task author
# lists a second acceptable value for 85 questions over 7 tasks
# (judge/alternate_answers.yaml); utils/answer_check.py accepts an attempt
# that matches EITHER value under the same rules. Comparison rules unchanged;
# a constant typed in place of a formula is hardcoded whichever value it holds.
RULES_VERSION = "v6.7"


def render_rules_text() -> str:
    """The rulebook as prompt text. Rendered verbatim into template_8 so the
    judge and the harness share one definition of equality."""
    outflow = ", ".join(sorted({w for w in OUTFLOW_LEXICON if " " not in w}))
    guard = ", ".join(sorted({w for w in INFLOW_GUARD if " " not in w}))
    return f"""Answer-equivalence rulebook ({RULES_VERSION}) — how to decide whether an attempt's answer equals the golden answer. These rules remove noise, not standards: a genuinely different value is wrong however it is dressed.
  1. Tolerance is scale-aware. Two numbers are THE SAME when they differ by no more than the LARGER of: (a) one part in a million of the value, and (b) half a unit of the last decimal place the golden carries (two decimal places => 0.005; "whole numbers" => 0.5) — in other words, when the attempt's number ROUNDS TO the golden's. The golden answers are rounded figures, so an unrounded attempt that rounds to the golden holds the same underlying value: 100.63598 vs 100.64 — same; 480,051.30 vs 480,051.31 — same; 5,000.00 vs 5,000.004 — same; 0.505 vs 0.50 — same (on the boundary); 0.51 vs 0.50 — different (a 2% error); 100.646 vs 100.64 — different (it rounds to 100.65); 101 vs 100.9 under "whole numbers" — same, but 100 vs 100.9 is the wrong rounding. When BOTH the golden and the attempt are rounded figures, a full unit is allowed (two correct roundings can sit one notch apart). The decimal places come from the golden's own ROUND(...) when it has one, otherwise from the Questions header; for a percentage stored as a fraction (0.4213 shown as 42.13%) the header's places apply to the percentage rendering. Never fail an answer for a rounding difference inside that precision; never forgive a proportionally large error because the numbers are small. NOT rounding as instructed is a presentation matter judged under Rounding / Rounded outputs — it is never an accuracy mistake.
  2. Sign convention is accepted ONLY on outflow rows. If the question names an outflow quantity ({outflow}), an answer equal in magnitude but opposite in sign is THE SAME — except where the row is a net, inflow, balance or difference figure ({guard}): "Interest income", "Changes in Net Working Capital" and "Intangible Assets growth" keep their sign, and a flipped one is WRONG. On every other row the sign is part of the answer — a flipped "difference", "net", or "change" is WRONG.
  3. Percent form. When the unit is %, the question asks for a percent/rate/margin/return/ratio, the golden cell is percent-formatted, or the answer carries a literal "%", then 0.42 and 42 (and "42%") are THE SAME answer.
  4. Compare values, never rendered strings. (1,890,487.51) shown in parentheses IS -1890487.51; $, commas and % are display only; number formats, fonts and alignment belong to Formatting, not Accuracy.
  5. Sentinel text is interchangeable: "Outside model horizon", "N/A", "beyond horizon" and similar mean the same thing when compared with each other — but a sentinel where the golden has a NUMBER is a wrong answer.
  6. Numbers written as text ("42", "$1,234.00", "(480,051.30)") are read as numbers before comparing.
  7. Dates: 2027-01-01, 01/01/2027 and the Excel serial for that day are THE SAME date.
  8. Zero: 0, 0.0 and an accounting "-" are all zero. A BLANK answer cell is UNANSWERED (that is a completeness failure, not an equivalence question).
  9. Unit scale is never forgiven: 1,234 vs 1,234,000 is WRONG even if a thousands/millions convention explains it — record it as a mistake and say "possible unit-scale difference" so a reviewer can see the cause.
 10. Hardcoded answers: an answer typed as a constant instead of a live formula referencing the model is a mistake for Final calculation accuracy even when the number is right (the rubric requires the workbook to calculate it). Text answers (Yes/No, sentinels) may be literal.
 11. Rounding is never an accuracy matter: whether answers are rounded as instructed is judged under Rounding / Rounded outputs, never under Final calculation accuracy."""
