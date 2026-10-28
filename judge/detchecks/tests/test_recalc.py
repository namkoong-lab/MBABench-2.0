"""Unit tests for the recalculation pipeline (core/recalc.py) with FAKE LibreOffice / Excel
runners (no office application is launched), plus the per-check live switch.

    cd judge && python -m detchecks.tests.test_recalc

Plain asserts; also collectable by pytest.  Workbooks are built with test_checks_22's builder;
the recalc cache goes to a temporary workdir under detchecks/scratch/errors22/.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback

from detchecks.api import grade
from detchecks.checks import REGISTRY
from detchecks.checks.c22 import C22
from detchecks.core import recalc as R
from detchecks.core.values import ValueSource
from detchecks.errors import GradingError
from detchecks.tests import test_checks_22 as T
from detchecks.tests.test_checks_22 import APP_EXCEL, APP_LO, APP_OPX, book, c, sheet

_WORK = None


def workdir() -> str:
    global _WORK
    if _WORK is None:
        os.makedirs(T.SCRATCH, exist_ok=True)
        _WORK = tempfile.mkdtemp(prefix="recalc_cache_", dir=T.SCRATCH)
    return _WORK


def fake_runner(values: dict, calls: list, *, excel: bool = False):
    """A LibreOffice / Excel stand-in: writes a copy whose sheets hold `values`
    ({sheet name: [c(...) cells]}) and records each call."""
    def run(src, out_dir, policy):
        calls.append(src)
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".xlsx")
        book([(nm, sheet(cells)) for nm, cells in values.items()],
             app=APP_EXCEL if excel else APP_LO, excel=excel, path=out)
        return out
    return run


def never(src, out_dir, policy):
    raise AssertionError(f"runner must not be called for {src}")


def policy(**kw) -> R.RecalcPolicy:
    kw.setdefault("workdir", workdir())
    kw.setdefault("lo_runner", never)
    kw.setdefault("excel_runner", never)
    kw.setdefault("excel_available", True)
    return R.RecalcPolicy(**kw)


def raises(fn, *needles):
    try:
        fn()
    except GradingError as e:
        msg = str(e) + " " + " ".join(str(v) for v in getattr(e, "failures", {}).values())
        for n in needles:
            assert n in msg, (n, msg[:900])
        return e
    raise AssertionError(f"expected GradingError containing {needles}")


# ============================================================================ tests
def test_excel_saved_file_uses_its_cache():
    p = book([("S", sheet([c("A1", 2, f="1+1")]))])                      # Excel signature
    plan = R.ensure_values(p, workdir(), policy())
    assert plan.source == "cache" and plan.value_path is None and plan.errors_vetted and plan.writer == "excel"
    v = grade(p, checks=[22], recalc=policy())[C22.key]
    assert v["decision"] == "pass" and v["stats"]["values"]["source"] == "cache"


def test_libreoffice_copy_no_gaps_cached_once():
    p = book([("S", sheet([c("A1", 0, f="1/0"), c("A2", 0, f="SUM(1,2)")]))], app=APP_OPX, excel=False)
    calls: list = []
    lo = fake_runner({"S": [c("A1", "#DIV/0!", t="e", f="1/0"), c("A2", 3, f="SUM(1,2)")]}, calls)
    pol = policy(lo_runner=lo)
    plan = R.ensure_values(p, workdir(), pol)
    assert plan.source == "libreoffice" and plan.n_gaps == 0 and plan.errors_vetted and calls == [p]
    assert plan.n_lo_errors == 0 and os.path.exists(plan.value_path) and plan.writer == "openpyxl"
    plan2 = R.ensure_values(p, workdir(), pol)                            # second run: from the cache
    assert calls == [p] and plan2.timings.get("from_cache") is True and plan2.value_path == plan.value_path
    v = grade(p, checks=[22], recalc=pol)[C22.key]
    assert [m["location"] for m in v["mistakes"]] == ["S!A1"]
    assert v["stats"]["values"]["source"] == "libreoffice" and v["stats"]["values"]["n_gaps"] == 0
    # the copy's error values are vetted -> trusted; an unvetted manual copy keeps #NAME? untrusted
    vs = ValueSource(plan.value_path, errors_vetted=True)
    assert vs.cursor("S").get(1, 1)[0] == "ok"
    vs.close()


def test_genuine_errors_are_not_gaps():
    cells = [c("A1", 1), c("A2", 2),
             c("B1", 0, f="CONUTIFS(A1:A2,1)"),             # misspelled function
             c("B2", 0, f="NoSuchName*2"),                  # undefined name
             c("B3", 0, f='"abc"+1'),                       # scalar text arithmetic
             c("B4", 0, f="A1/0")]
    p = book([("S", sheet(cells))], app=APP_OPX, excel=False)
    lo = fake_runner({"S": [c("A1", 1), c("A2", 2), c("B1", "#NAME?", t="e", f="CONUTIFS(A1:A2,1)"),
                            c("B2", "#NAME?", t="e", f="NoSuchName*2"), c("B3", "#VALUE!", t="e", f='"abc"+1'),
                            c("B4", "#DIV/0!", t="e", f="A1/0")]}, [])
    pol = policy(lo_runner=lo)
    plan = R.ensure_values(p, workdir(), pol)
    assert plan.source == "libreoffice" and plan.n_gaps == 0 and plan.n_lo_errors == 3 and plan.n_genuine_lo_errors == 3
    v = grade(p, checks=[22], recalc=pol)[C22.key]
    assert v["decision"] == "fail" and v["stats"]["error_cells"] == 4, v["mistakes"]
    assert v["stats"]["error_cells_by_code"] == {"#DIV/0!": 1, "#NAME?": 2, "#VALUE!": 1}


SCAN = "SUM(_xlfn.SCAN(0,A1:A3,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)))"


def _gap_book():
    cells = [c("A1", 1), c("A2", 2), c("A3", 3), c("B1", 0, f=SCAN), c("B2", 0, f="B1*2")]
    return book([("S", sheet(cells))], app=APP_OPX, excel=False)


def _gap_lo(calls):
    return fake_runner({"S": [c("A1", 1), c("A2", 2), c("A3", 3), c("B1", "#NAME?", t="e", f=SCAN),
                              c("B2", "#NAME?", t="e", f="B1*2")]}, calls)


def test_gap_needs_excel_not_allowed_or_unavailable():
    p = _gap_book()
    calls: list = []
    # Excel off (the default, Patrick 2026-10-04): the LibreOffice copy is used as displayed, gaps
    # are recorded, nothing is raised
    plan = R.ensure_values(p, workdir(), policy(lo_runner=_gap_lo(calls), excel_allowed=False))
    assert plan.source == "libreoffice" and plan.errors_vetted and plan.n_gaps == 2 and calls == [p]
    assert "SCAN" in plan.gap_functions and any("could not compute 2" in n for n in plan.notes), plan.notes
    # the LibreOffice copy and the gap list were kept: the next call does not run LibreOffice again
    plan2 = R.ensure_values(p, workdir(), policy(lo_runner=never, excel_allowed=False))
    assert plan2.n_gaps == 2 and plan2.errors_vetted
    # Excel wanted but not available on the machine: loud error
    raises(lambda: R.ensure_values(p, workdir(), policy(lo_runner=never, excel_allowed=True, excel_available=False)),
           "Excel recalculation required", "not available")
    # through the engine with Excel off (judge v15, Patrick 2026-10-28): the only error cells are gap cells
    # (B1 uses SCAN / LAMBDA; B2 inherits its #NAME?), so No formula errors is UNDECIDED for this attempt
    # and tells the LLM which cells are gaps and why; the full gap list survived the plan cache
    v = grade(p, checks=[22, 92], recalc=policy(lo_runner=never, excel_allowed=False))
    u = v[C22.key]
    assert u["decision"] == "undecided" and u["mistakes"] == [] and u["stats"]["values"]["source"] == "libreoffice"
    assert u["stats"]["values"]["n_gaps"] == 2 and v[REGISTRY[92].key]["decision"] == "pass"
    assert u["stats"]["gap_error_cells"] == 2 and u["stats"]["undecided"]["decided_by"] == "llm", u["stats"]
    assert u["stats"]["gap_sheets"] == ["S"] and u["stats"]["error_cells"] == 0, u["stats"]
    note = u["llm_note"]
    assert "S!B1 shows #NAME? = =SUM(_xlfn.SCAN(" in note and "[LAMBDA, SCAN]" in note, note
    assert "S!B2 shows #NAME? = =B1*2" in note and "Every other cell of the workbook has been checked by code" in note
    assert "Judge whether each of these formulas is valid Excel" in note and "2 cell(s)" in note, note
    assert "undecided" in u["summary"] or "LLM judge decides" in u["summary"], u["summary"]


def test_gap_rerouted_to_excel():
    p = _gap_book()
    lo_calls: list = []
    xl_calls: list = []
    xl = fake_runner({"S": [c("A1", 1), c("A2", 2), c("A3", 3), c("B1", 6, f=SCAN), c("B2", 12, f="B1*2")]},
                     xl_calls, excel=True)
    # Excel is OFF by default (Patrick 2026-10-04, LibreOffice only); the reroute is tested with it switched on
    pol = policy(lo_runner=_gap_lo(lo_calls), excel_runner=xl, excel_allowed=True)
    plan = R.ensure_values(p, workdir(), pol)
    # the LibreOffice copy and gap list of the previous test (same file content, same hash) are
    # reused: LibreOffice is not run again, only Excel
    assert plan.source == "excel" and plan.errors_vetted and xl_calls == [p] and lo_calls == [], (xl_calls, lo_calls)
    assert plan.n_gaps == 2 and "SCAN" in plan.gap_functions and "LAMBDA" in plan.gap_functions
    assert "excel_s" in plan.timings and plan.value_path.endswith(os.path.join("excel", os.path.basename(p)))
    v = grade(p, checks=[22], recalc=pol)[C22.key]
    assert v["decision"] == "pass" and v["stats"]["values"]["source"] == "excel"
    assert v["stats"]["values"]["n_gaps"] == 2 and v["stats"]["values"]["gaps"][0]["ref"] == "B1"
    # cached: a third run touches neither office application
    plan3 = R.ensure_values(p, workdir(), policy())
    assert plan3.source == "excel" and plan3.timings.get("from_cache") is True


def test_classify_lo_error_table():
    uns = R.lo_unsupported((25, 8))
    defined = {"MYLAM": "LAMBDA(_xlpm.x,_xlpm.x*2)", "RATE_": "Inputs!$B$5"}
    tables = frozenset({"TBL"})

    def cls(text, value, is_array=False):
        return R.classify_lo_error(text, value, defined, tables, uns, is_array)
    assert cls(SCAN, "#NAME?")[0] is True and "SCAN" in cls(SCAN, "#NAME?")[1]
    assert cls("INDEX(B20#,1,2)", "#VALUE!")[0] is True                      # spill reference
    assert cls("MyLam(A1)", "#NAME?")[0] is True                              # LAMBDA name called
    assert cls("RATE_*2", "#VALUE!")[0] is True                               # a name: may be a range
    assert cls("CONUTIFS(A1:A3,1)", "#NAME?") == (False, ["CONUTIFS"], "genuine #NAME?: unknown function CONUTIFS")
    assert cls("NoSuchName*2", "#NAME?")[0] is False
    assert cls("Tbl[Col]*2", "#NAME?")[0] is True                             # table known: not undefined
    assert cls('"abc"+1', "#VALUE!")[0] is False
    assert cls("A1*B1", "#VALUE!")[0] is False
    assert cls("SUM(A1:A3)*2", "#VALUE!")[0] is True                          # range in a plain formula
    assert cls("MMULT(TRANSPOSE(A1:A3),B1:B3)", "#VALUE!")[0] is True
    assert cls("A1*B1", "#VALUE!", is_array=True)[0] is True
    assert cls("_xlfn.XLOOKUP(1,A1:A3,B1:B3)", "#NAME?")[0] is True            # valid function, LibreOffice still #NAME?
    # judge v15 (Patrick 2026-10-28): a future function stored WITHOUT _xlfn. is #NAME? in Excel too -> genuine
    for t in ("XLOOKUP(1,A1:A3,B1:B3)", "LET(_xlpm.x,1,_xlpm.x)", "LAMBDA(_xlpm.x,_xlpm.x)(1)", "IFS(A1>0,1,TRUE,0)",
              "FILTER(A1:A3,B1:B3)", "SUM(SEQUENCE(3))", "STDEV.S(A1:A3)"):
        gap, fns, why = cls(t, "#NAME?")
        assert gap is False and "without the _xlfn. prefix" in why, (t, gap, fns, why)
    assert cls("XLOOKUP(1,A1:A3,B1:B3)", "#VALUE!")[0] is False                 # whatever LibreOffice shows
    for t in ("_xlfn.LET(_xlpm.x,1,_xlpm.x)", "_xlfn._xlws.FILTER(A1:A3,B1:B3)", "_xlfn.STDEV.S(A1:A3)",
              "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)", "SUM(_xlfn.SEQUENCE(3))"):
        assert cls(t, "#NAME?")[0] is True, t                                   # prefixed: a gap (LibreOffice's problem)
    assert cls("CUBEVALUE(\"c\",\"m\")", "#NAME?") == (True, ["CUBEVALUE"], "LibreOffice cannot evaluate: CUBEVALUE")
    assert cls("REGEXTEST(A1,\"x\")", "#NAME?")[0] is False                     # 365 function: needs the prefix
    assert cls("_xlfn.REGEXTEST(A1,\"x\")", "#NAME?") == (True, ["REGEXTEST"], "LibreOffice cannot evaluate: REGEXTEST")
    for fn in ("TRANSLATE", "DETECTLANGUAGE", "BINOM.DIST.RANGE", "ECMA.CEILING", "CUBEMEMBER", "RTD", "CALL",
               "REGISTER.ID", "SQL.REQUEST", "DBCS", "PHONETIC", "THAIDIGIT", "YEN", "REGEXEXTRACT", "REGEXREPLACE"):
        assert fn in uns and fn in R.EXCEL_FUNCTIONS, fn
    assert R.LO_VERIFIED_VERSION == "25.8.7" and R.XLFN_REQUIRED <= R.EXCEL_FUNCTIONS
    assert not (R.XLFN_REQUIRED & {"SUM", "VLOOKUP", "CUBEVALUE", "RTD", "PHONETIC", "BAHTTEXT", "IFERROR", "SUMIFS"})
    assert cls("SUM(A1:A3", "#VALUE!")[0] is True                             # unparseable
    assert cls(None, "#VALUE!")[0] is True
    old = R.lo_unsupported((7, 4))
    assert "XLOOKUP" in old and "XLOOKUP" not in uns and "LAMBDA" in uns
    assert R.lo_version_of("LibreOffice/25.8.7.3$MacOSX_AARCH64 LibreOffice_project/x") == (25, 8)
    assert R.lo_version_of("Microsoft Excel") is None
    assert R.lo_version_of("LibreOffice/7.4.7.2$Linux_AARCH64 LibreOffice_project/40$Build-2") == (7, 4)
    assert "CONUTIFS" not in R.EXCEL_FUNCTIONS and {"XLOOKUP", "SCAN", "REGEXTEST", "NA"} <= R.EXCEL_FUNCTIONS


def test_gap_cells_split_the_verdict_v15():
    """Judge v15 (Patrick 2026-10-28), No formula errors (22) per attempt: (a) a clean workbook whose
    _xlfn.LET LibreOffice computes -> pass by code; (b) the same plus a real #DIV/0! -> fail by code;
    (c) a _xlfn.LAMBDA cell only -> undecided, the note names the cell; (d) an unprefixed =XLOOKUP(...)
    -> genuine #NAME?, fail by code; (e) a gap cell beside a genuine error -> fail by code, the gap cell
    in stats, never a mistake; a typed error constant and an implicit intersection fail whatever the
    gaps."""
    let = "_xlfn.LET(_xlpm.x,A1*2,_xlpm.x+1)"
    lam = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A1)"
    xl = "XLOOKUP(1,A1:A2,A1:A2)"

    def run_lo(cells, lo_cells, checks=(22,)):
        p = book([("S", sheet(cells))], app=APP_OPX, excel=False)
        pol = policy(lo_runner=fake_runner({"S": lo_cells}, []))
        return grade(p, checks=list(checks), recalc=pol)[C22.key]

    # (a)
    v = run_lo([c("A1", 1), c("B1", 0, f=let)], [c("A1", 1), c("B1", 3, f=let)])
    assert v["decision"] == "pass" and v["stats"]["decided_by"] == "code" and v["stats"]["gap_error_cells"] == 0, v
    assert v["stats"]["gap_information"] == "pipeline" and "llm_note" not in v
    # (b)
    v = run_lo([c("A1", 1), c("A2", 0), c("B1", 0, f=let), c("B2", 0, f="A1/A2")],
               [c("A1", 1), c("A2", 0), c("B1", 3, f=let), c("B2", "#DIV/0!", t="e", f="A1/A2")])
    assert v["decision"] == "fail" and [m["location"] for m in v["mistakes"]] == ["S!B2"], v["mistakes"]
    assert v["stats"]["decided_by"] == "code"
    # (c)
    v = run_lo([c("A1", 1), c("B1", 0, f=lam)], [c("A1", 1), c("B1", "#NAME?", t="e", f=lam)])
    assert v["decision"] == "undecided" and v["mistakes"] == [] and v["stats"]["gap_error_cells"] == 1, v
    assert "S!B1 shows #NAME? = =_xlfn.LAMBDA(" in v["llm_note"] and "[LAMBDA]" in v["llm_note"], v["llm_note"]
    assert v["stats"]["undecided"]["why"].startswith("Only LibreOffice gap cells show errors (1 cell(s)")
    # (d)
    v = run_lo([c("A1", 1), c("A2", 2), c("B1", 0, f=xl)], [c("A1", 1), c("A2", 2), c("B1", "#NAME?", t="e", f=xl)])
    assert v["decision"] == "fail" and [m["location"] for m in v["mistakes"]] == ["S!B1"], v
    assert v["stats"]["gap_error_cells"] == 0 and v["stats"]["values"]["n_gaps"] == 0, v["stats"]["values"]
    # (e): the gap cell is reported beside the genuine error, not as a mistake
    v = run_lo([c("A1", 1), c("B1", 0, f=lam), c("B2", 0, f=xl)],
               [c("A1", 1), c("B1", "#NAME?", t="e", f=lam), c("B2", "#NAME?", t="e", f=xl)])
    assert v["decision"] == "fail" and [m["location"] for m in v["mistakes"]] == ["S!B2"], v["mistakes"]
    assert v["stats"]["gap_error_cells"] == 1 and v["stats"]["error_cells"] == 1, v["stats"]
    assert "1 LibreOffice gap cell(s) show errors Excel would not (not counted)" in v["summary"], v["summary"]
    assert v["stats"]["gap_error_examples"] == ["S!B1 #NAME?: =_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A1) [LAMBDA]"], v["stats"]
    # a typed error constant beside a gap cell: fail by code
    v = run_lo([c("A1", "#N/A", t="e"), c("B1", 0, f=lam)], [c("A1", "#N/A", t="e"), c("B1", "#NAME?", t="e", f=lam)])
    assert v["decision"] == "fail" and v["stats"]["typed_error_constants"] == 1 and v["stats"]["gap_error_cells"] == 1
    # an implicit intersection beside a gap cell: fail by code (flagged from the formula text)
    ii = "A1:A3*2"
    v = run_lo([c("A1", 1), c("A2", 2), c("A3", 3), c("B1", 0, f=lam), c("C5", 0, f=ii)],
               [c("A1", 1), c("A2", 2), c("A3", 3), c("B1", "#NAME?", t="e", f=lam), c("C5", 2, f=ii)])
    assert v["decision"] == "fail" and v["stats"]["implicit_intersection_cells"] == 1, v
    assert [m["location"] for m in v["mistakes"]] == ["S!C5"] and v["stats"]["gap_error_cells"] == 1


def test_no_recalc_policy_keeps_old_behaviour():
    """Without a RecalcPolicy the engine grades from the file's caches under the trust policy
    (openpyxl caches untrusted -> GradingError), exactly as before."""
    p = book([("S", sheet([c("A1", 0, f="1/0")]))], app=APP_OPX, excel=False)
    v = grade(p, checks=[22])                     # skipped since Patrick 2026-10-05 (every attempt graded)
    assert next(iter(v.values()))["stats"]["defaults"]["untrusted_value"]["count"] >= 1, v
    assert isinstance(R.excel_available(), bool)


TESTS = [test_excel_saved_file_uses_its_cache, test_libreoffice_copy_no_gaps_cached_once,
         test_genuine_errors_are_not_gaps, test_gap_needs_excel_not_allowed_or_unavailable,
         test_gap_rerouted_to_excel, test_classify_lo_error_table, test_gap_cells_split_the_verdict_v15,
         test_no_recalc_policy_keeps_old_behaviour]


def main() -> int:
    failed = 0
    try:
        for t in TESTS:
            try:
                t()
                print(f"PASS {t.__name__}")
            except Exception:  # noqa: BLE001
                failed += 1
                print(f"FAIL {t.__name__}")
                traceback.print_exc()
    finally:
        for d in (_WORK, T._TMP):
            if d and os.path.isdir(d):
                shutil.rmtree(d, ignore_errors=True)
    print(f"\n{len(TESTS) - failed}/{len(TESTS)} tests passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
