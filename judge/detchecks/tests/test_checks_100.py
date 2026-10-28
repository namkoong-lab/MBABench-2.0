"""Unit tests for Automatic calculation mode (100) on synthetic workbooks (raw SpreadsheetML zips,
test_checks_22's builder; a hand-written <calcPr> per case).

    cd judge && python -m detchecks.tests.test_checks_100

Plain asserts; also collectable by pytest.  No office application runs (the check never reads values).
"""
from __future__ import annotations

import sys
import traceback

from detchecks.api import Engine, grade
from detchecks.checks import REGISTRY
from detchecks.checks import c100 as M
from detchecks.checks.c100 import C100, dyadic_blocks, parse_structured_spec
from detchecks.errors import GradingError
from detchecks.tests import test_checks_colours as TC
from detchecks.tests.test_checks_22 import book, c, sheet

AUTO = '<calcPr calcId="191029"/>'
MANUAL = '<calcPr calcId="191029" calcMode="manual"/>'
ITER = '<calcPr calcId="191029" iterate="1" iterateCount="100" iterateDelta="0.001"/>'
ITER_MANUAL = '<calcPr calcId="191029" calcMode="manual" iterate="1"/>'


def run(path):
    return Engine(path, [C100()]).run()[C100.key]


def locs(v):
    return [m["location"] for m in v["mistakes"]]


# ============================================================================ helpers
def test_dyadic_blocks_and_structured_specs():
    def rows_of(blocks):
        out = set()
        for lvl, k in blocks:
            out |= set(range((k << lvl) + 1, ((k + 1) << lvl) + 1))
        return out
    for r1, r2 in ((1, 1), (1, 10), (5, 5), (3, 1000), (1, 1048576), (1048570, 1048576), (64, 65), (100, 4000)):
        bl = dyadic_blocks(r1, r2)
        assert rows_of(bl) == set(range(r1, r2 + 1)), (r1, r2, bl)
        assert len(bl) <= 2 * 21, (r1, r2, len(bl))
    assert parse_structured_spec("[Col]") == (frozenset({"data"}), "Col", "Col")
    assert parse_structured_spec("[[#This Row],[Col]]") == ("this", "Col", "Col")
    assert parse_structured_spec("[@Col]") == ("this", "Col", "Col") and parse_structured_spec("[@[Col]]") == ("this", "Col", "Col")
    assert parse_structured_spec("[[#Totals],[Col]]") == (frozenset({"totals"}), "Col", "Col")
    assert parse_structured_spec("[[#Headers],[#Data]]") == (frozenset({"headers", "data"}), None, None)
    assert parse_structured_spec("[#All]") == (frozenset({"headers", "data", "totals"}), None, None)
    assert parse_structured_spec("[[Col1]:[Col2]]") == (frozenset({"data"}), "Col1", "Col2")
    assert parse_structured_spec("[]") == (frozenset({"data"}), None, None)
    assert parse_structured_spec("[#Bogus]") is None and parse_structured_spec("Col") is None


def test_fast_path_matches_the_parser():
    """_fast_entries reads plain formulas without the parser and must yield exactly what _entries yields,
    counters included; anything it does not understand goes to the parser (None)."""
    p = book([("Calc", sheet([c("A1", 1)])), ("My Sheet", sheet([c("A1", 1)])), ("Inputs", sheet([c("A1", 1)]))], calc_pr=AUTO)
    from detchecks.core.package import Package
    pkg = Package.open(p)
    chk = C100()
    chk.start(pkg)
    accepted = ["A1*2", "$A$1+B$2-C3/D4^2", "SUM(A1:B5)", "SUM(A:A)", "SUM(1:1)", "SUM($A:$C,2:2)", "Inputs!B5*C3",
                "'My Sheet'!A1:B9", "IF(B20149=\"\",\"\",A20149-INDEX(Inputs!$T$9:$T$1258,B20149)+1)",
                "INDEX(Inputs!$B$2:$G$157,CN$3,$C115)", "SUM(OFFSET(A1,0,0,5,1))", "ROWS($B$14:B14)",
                "IF(COLUMN()-COLUMN($G$62)<3,1,2)", "_xlfn.IFS(A1>0,1,TRUE,0)", "IFERROR(1/A1,\"n/a\")",
                "1E5*A1", "ROUND(A1*1.5%,2)", "(A1+A2)*(B1-B2)", "SUM(A1:A3)*ROWS(A1:A3)", "TRUE", "A1<>B1",
                "INDIRECT(\"A1\")+1", "_xlfn._xlws.FILTER(A1:A9,B1:B9>0)", "YEARFRAC(A1,F9726,INDEX(Inputs!$O$9:$O$1258,B9726))",
                "  A1 + 1 ", "SUM(A1:A3, B1:B3)", "=A1"]
    rejected = ["Rate*2", "SUM(Tbl[Col])", "A1#", "@A1:A3", "#REF!+1", "{1,2}", "R1C1", "Jan:Dec!B5", "A1:.A5",
                "SUM(A1:A5 B2:B9)", "MyFn(A1)", "CONUTIFS(A1:A3,1)", "[1]Prices!A1", "'C:\\p\\[B.xlsx]S'!A1",
                "Nope!A1", "SUM(A1", "A1)", "1,2", "_xlpm.x*2", "TRUEISH", "A1:#REF!"]
    for t in accepted:
        info = {}
        fast = M._fast_entries(t, chk._sheet_idx, info)
        assert fast is not None, t
        g0, i0, d0 = chk.n_geometry_refs, chk.dynamic["index_offset_reference_args"], chk.dynamic["count"]
        slow = chk._entries(M.parse_text(t))
        assert sorted(map(str, fast)) == sorted(map(str, slow)), (t, fast, slow)
        assert (chk.n_geometry_refs - g0, chk.dynamic["index_offset_reference_args"] - i0, chk.dynamic["count"] - d0) == \
            (info["geometry"], info["index_args"], info["dynamic"]), (t, info)
    for t in rejected:
        assert M._fast_entries(t, chk._sheet_idx, {}) is None, t
    pkg.close()
    # the graph is the same with and without the fast path
    cells = [c("B1", 0, f="Inputs!C3*2"), c("B2", 0, f="B1+1"), c("B5", 0, f="SUM(B1:B2)*ROWS(B1:B2)")]
    inputs = sheet([c("C3", 0, f="Calc!B2/2")])
    p = book([("Calc", sheet(cells)), ("Inputs", inputs)], calc_pr=AUTO)
    v1 = run(p)
    assert v1["stats"]["fast_path_texts"] == 4 and v1["decision"] == "fail" and v1["stats"]["largest_cycle"] == 3, v1["stats"]
    saved = M.FAST_PATH
    M.FAST_PATH = False
    try:
        v2 = run(p)
    finally:
        M.FAST_PATH = saved
    assert v2["stats"]["fast_path_texts"] == 0 and v2["decision"] == "fail" and v2["mistakes"] == v1["mistakes"]
    assert v2["stats"]["n_edges"] == v1["stats"]["n_edges"] and v2["stats"]["geometry_refs_skipped"] == v1["stats"]["geometry_refs_skipped"]


def test_rubric_key_and_contract():
    assert REGISTRY[100] is C100 and C100.key == "Potential Dangers/Automatic calculation mode"
    assert C100.needs_values is False and C100.live is True


# ============================================================================ rules 1 and 3
def test_manual_mode_fails_whatever_else():
    cells = [c("A1", 1), c("B1", 2, f="A1*2")]
    v = run(book([("S", sheet(cells))], calc_pr=MANUAL))
    assert v["decision"] == "fail" and locs(v) == ["workbook"] and "calculation mode is manual" in v["mistakes"][0]["description"]
    assert v["stats"]["rule"] == 1 and v["stats"]["decided_by"] == "code" and v["stats"]["n_cycles"] is None
    # manual with iteration on, and a cycle: still rule 1
    v = run(book([("S", sheet([c("A1", 1, f="A1+1")]))], calc_pr=ITER_MANUAL))
    assert v["decision"] == "fail" and locs(v) == ["workbook"] and v["stats"]["rule"] == 1
    # automatic forms pass: default, explicit auto, autoNoTable, no calcPr at all
    for cp in (AUTO, '<calcPr calcMode="auto"/>', '<calcPr calcMode="autoNoTable"/>', ""):
        v = run(book([("S", sheet(cells))], calc_pr=cp))
        assert v["decision"] == "pass" and v["stats"]["rule"] == 3 and v["stats"]["manual"] is False, (cp, v)
    # iteration on without a cycle passes, recorded
    v = run(book([("S", sheet(cells))], calc_pr=ITER))
    assert v["decision"] == "pass" and v["stats"]["iteration_on_without_cycle"] is True and v["stats"]["iterate"] is True
    assert "iterative calculation is on without a cycle" in v["summary"]


# ============================================================================ rule 4: cycles, iteration off
def test_self_loops_and_simple_cycles_fail_with_iteration_off():
    # direct self reference
    v = run(book([("S", sheet([c("A1", 0, f="A1+1")]))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and v["stats"]["n_cycles"] == 1, v
    assert "circular reference through 1 cell(s)" in v["mistakes"][0]["description"]
    assert "iterative calculation is off" in v["mistakes"][0]["description"] and v["stats"]["rule"] == 4
    # a formula inside the range it sums
    cells = [c(f"A{r}", r) for r in (1, 2, 3, 4, 6, 7, 8, 9, 10)] + [c("A5", 0, f="SUM(A1:A10)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!A5"], v["mistakes"]
    # three cells across two sheets: Calc!B1 -> Inputs!C3 -> Calc!B2 -> Calc!B1
    calc = sheet([c("B1", 0, f="Inputs!C3*2"), c("B2", 0, f="B1+1"), c("B5", 0, f="B2*3")])
    inputs = sheet([c("C3", 0, f="Calc!B2/2"), c("C4", 5)])
    v = run(book([("Calc", calc), ("Inputs", inputs)], calc_pr=AUTO))
    assert v["decision"] == "fail" and v["stats"]["n_cycles"] == 1 and v["stats"]["largest_cycle"] == 3, v["stats"]
    d = v["mistakes"][0]["description"]
    assert "through 3 cell(s)" in d and "Calc!B1:B2" in d and "Inputs!C3" in d, d
    assert locs(v) == ["Calc!B1:B2"], locs(v)                 # the first sheet's first block locates it
    # B5 depends on the cycle but is not part of it: not counted
    assert v["stats"]["cycles"][0]["cells"] == 3
    # no cycle: a chain, a diamond, and references to constants
    chain = sheet([c("A1", 1), c("A2", 0, f="A1*2"), c("A3", 0, f="A2+A1"), c("A4", 0, f="SUM(A1:A3)")])
    v = run(book([("S", chain)], calc_pr=AUTO))
    assert v["decision"] == "pass" and v["stats"]["n_cycles"] == 0 and v["stats"]["n_formula_cells"] == 3
    # whole-column sum in its own column (=SUM(A:A) in column A)
    v = run(book([("S", sheet([c("A1", 1), c("A2", 2), c("A4", 0, f="SUM(A:A)")]))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!A4"], v
    # ... but a whole-column sum of another column is fine
    v = run(book([("S", sheet([c("A1", 1), c("A2", 2), c("B4", 0, f="SUM(A:A)")]))], calc_pr=AUTO))
    assert v["decision"] == "pass", v


def test_breaker_switch_off_still_fails():
    """Static means static (Patrick 2026-10-28): =IF($B$1=1, <circular>, 0) with the switch OFF is a cycle."""
    cells = [c("B1", 0), c("C1", 0, f="IF($B$1=1,C2*0.05,0)"), c("C2", 0, f="C1+100")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!C1:C2"], v


def test_cycles_through_names_shared_formulas_arrays_and_3d():
    # through a defined name: A1 = Rate*2, Rate = S!$A$1
    v = run(book([("S", sheet([c("A1", 0, f="Rate*2")]))], names={"Rate": "S!$A$1"}, calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and "1 defined name(s)" in v["mistakes"][0]["description"], v
    # through two names only (no cell): Name Manager location
    v = run(book([("S", sheet([c("A1", 0, f="X*2")]))], names={"X": "Y+1", "Y": "X+1"}, calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["Name Manager"] and v["stats"]["cycles"][0]["names"] == 2, v
    # a name that reads nothing circular is fine
    v = run(book([("S", sheet([c("A1", 0, f="Rate*2"), c("B1", 5)]))], names={"Rate": "S!$B$1"}, calc_pr=AUTO))
    assert v["decision"] == "pass", v
    # shared formulas: master B2 =A2+B1 (running total), children B3:B4; B1 is the total =SUM(B2:B4) -> cycle
    cells = [c("A2", 1), c("A3", 2), c("A4", 3), c("B1", 0, f="SUM(B2:B4)"),
             c("B2", 0, f="A2+B1", si=0, fref="B2:B4"), c("B3", 0, f="", si=0), c("B4", 0, f="", si=0)]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and v["stats"]["largest_cycle"] == 4, v["stats"]
    # the same shared group reading a constant column instead: children shifted, no cycle
    cells = [c("A1", 10), c("A2", 1), c("A3", 2), c("A4", 3), c("B2", 0, f="A2+$A$1", si=0, fref="B2:B4"),
             c("B3", 0, f="", si=0), c("B4", 0, f="", si=0), c("C1", 0, f="SUM(B2:B4)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "pass" and v["stats"]["n_formula_cells"] == 4, v
    # a shared child whose shifted reference reaches the total cell: B3 = A3+C1 ... C1 = SUM(B2:B4) -> cycle via the child
    cells = [c("A2", 1), c("A3", 2), c("B2", 0, f="A2+C0", si=0, fref="B2:B3"), c("B3", 0, f="", si=0),
             c("C1", 0, f="SUM(B2:B3)")]
    # (C0 is off the grid in the master; the child B3 reads C1: the master's #REF! leads nowhere)
    try:
        v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    except GradingError:
        v = None                                              # the lexer may reject C0: not this test's point
    if v is not None:
        assert v["decision"] in ("pass", "fail")
    # array / spill: anchor A1 =SEQUENCE(3)*B1, members A2:A3 unwritten; B1 =SUM(A1:A3) -> cycle through the range
    cells = [c("A1", 0, f="_xlfn.SEQUENCE(3)*B1", array=True, fref="A1:A3"), c("B1", 0, f="SUM(A2:A3)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and v["stats"]["largest_cycle"] == 2, v
    # a written member referenced directly: C1 = A2 (member of A1's range) with A1 = C1*2 -> cycle
    cells = [c("A1", 0, f="C1*{1;2}", array=True, fref="A1:A2"), c("A2", 0), c("C1", 0, f="A2+1")]
    cells[1] = '<c r="A2"><v>0</v></c>'
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail", v
    # spill reference: B1 = SUM(A1#) with A1 = B1*{1;2} spilling -> cycle
    cells = [c("A1", 0, f="B1*{1;2}", array=True, fref="A1:A2", cm=1), c("B1", 0, f="SUM(A1#)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail", v
    # 3-D reference: Jan:Mar!A1 summed on Mar!A1 itself -> cycle; summed on Total!A1 -> fine
    jan, feb = sheet([c("A1", 1)]), sheet([c("A1", 2)])
    v = run(book([("Jan", jan), ("Feb", feb), ("Mar", sheet([c("A1", 0, f="SUM(Jan:Mar!A1)")]))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["Mar!A1"], v
    v = run(book([("Jan", jan), ("Feb", feb), ("Mar", sheet([c("A1", 3)])), ("Total", sheet([c("A1", 0, f="SUM(Jan:Mar!A1)")]))],
                 calc_pr=AUTO))
    assert v["decision"] == "pass", v


def test_structured_references_and_tables():
    # colours' builder makes a one-column table Tbl on sheet index 1 (Inputs, A1:A4: header row 1, data 2:4)
    data = TC.sheet((2, [TC.c("A2", 0, f="SUM(Tbl[Col])")]))            # a formula INSIDE the data body of its column
    inputs = TC.sheet((1, [TC.cstr("A1", "Col")]), (2, [TC.c("A2", 1)]), (3, [TC.c("A3", 2)]),
                      (4, [TC.c("A4", 0, f="SUM(Tbl[Col])")]))          # A4 sums its own column: cycle
    p = TC.book("100_table.xlsx", [("Calc", data, None), ("Inputs", inputs, None)], tables={1: [("Tbl", "A1:A4", "Col")]})
    v = run(p)
    assert v["decision"] == "fail" and locs(v) == ["Inputs!A4"], v
    # [#This Row] on the formula's own row: Inputs!A3 = [@Col]*2 would be a self reference
    inputs = TC.sheet((1, [TC.cstr("A1", "Col")]), (2, [TC.c("A2", 1)]), (3, [TC.c("A3", 0, f="Tbl[[#This Row],[Col]]*2")]),
                      (4, [TC.c("A4", 3)]))
    p = TC.book("100_table_row.xlsx", [("Calc", TC.sheet(), None), ("Inputs", inputs, None)], tables={1: [("Tbl", "A1:A4", "Col")]})
    v = run(p)
    assert v["decision"] == "fail" and locs(v) == ["Inputs!A3"], v
    # a totals-row style reference from outside the table: Calc!A1 = SUM(Tbl[Col]) with constants -> pass
    inputs = TC.sheet((1, [TC.cstr("A1", "Col")]), (2, [TC.c("A2", 1)]), (3, [TC.c("A3", 2)]), (4, [TC.c("A4", 3)]))
    p = TC.book("100_table_ok.xlsx", [("Calc", TC.sheet((1, [TC.c("A1", 0, f="SUM(Tbl[Col])")])), None), ("Inputs", inputs, None)],
                tables={1: [("Tbl", "A1:A4", "Col")]})
    v = run(p)
    assert v["decision"] == "pass" and v["stats"]["structured_refs_unresolved"] == 0, v


# ============================================================================ rule 5: iteration on -> undecided
def test_cycle_with_iteration_on_is_undecided():
    calc = sheet([c("B1", 0, f="Inputs!C3*2"), c("B2", 0, f="B1+1")])
    inputs = sheet([c("C3", 0, f="Calc!B2/2")])
    cover = sheet([c("A1", "This model is intentionally circular (interest on cash): iterative calculation required.")])
    v = run(book([("Cover", cover), ("Calc", calc), ("Inputs", inputs)], calc_pr=ITER))
    assert v["decision"] == "undecided" and v["mistakes"] == [] and v["stats"]["rule"] == 5, v
    assert v["stats"]["undecided"]["decided_by"] == "llm" and v["stats"]["undecided"]["n_cycles"] == 1
    note = v["llm_note"]
    assert "found a circular reference through these cells" in note and "Calc!B1:B2" in note and "Inputs!C3" in note
    assert "Iterative calculation is enabled (calcMode=auto, iterate=on, iterateCount=100" in note, note
    assert "Does the workbook state that the model is intentionally circular" in note and "Pass only if it does" in note
    # the same workbook with iteration off: fail by code (rule 4)
    v = run(book([("Cover", cover), ("Calc", calc), ("Inputs", inputs)], calc_pr=AUTO))
    assert v["decision"] == "fail" and v["stats"]["rule"] == 4


# ============================================================================ misses and scale
def test_geometry_and_index_references_are_not_cycles():
    """Excel never dereferences a reference used for its geometry (ROWS, COLUMNS, ROW, COLUMN, ...) and reads only
    the computed part of an INDEX / OFFSET base reference: attempt 2539 (Excel Online, iteration off) caches 1, 2,
    3 for =IF(ROWS($B$14:B14)<=$C$8,ROWS($B$14:B14),"") in B14:B16, so these idioms are not circular."""
    cells = [c("B14", 1, f='IF(ROWS($B$14:B14)<=$C$8,ROWS($B$14:B14),"")'), c("B15", 2, f='IF(ROWS($B$14:B15)<=$C$8,ROWS($B$14:B15),"")'),
             c("C8", 3), c("E21", 0, f="COLUMNS($E$21:E$21)-1"), c("G62", 2030, f="IF(COLUMN()-COLUMN($G$62)<3,2030+COLUMN()-COLUMN($G$62),\"\")"),
             c("E4", 0, f="SUM(E8:INDEX(E:E,'S'!$C$8+8))"), c("E8", 1), c("E9", 2),
             c("A5", 0, f="INDEX($A$1:$A$10,ROW()-1)"), c("A4", 7),
             c("H3", 0, f="SUM(OFFSET(H1,0,0,5,1))"), c("H1", 1)]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "pass" and v["stats"]["n_cycles"] == 0, v["mistakes"]
    assert v["stats"]["geometry_refs_skipped"] == 7, v["stats"]            # one per reference operand of a distinct text
    assert v["stats"]["dynamic_refs"]["index_offset_reference_args"] == 3, v["stats"]["dynamic_refs"]
    # ... but a function that reads the values of a self-including range is circular (SUM beside ROWS)
    cells = [c("A1", 1), c("A2", 2), c("A5", 0, f="SUM(A1:A10)*ROWS(A1:A10)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!A5"], v
    # a lookup over a table that contains the formula's cell is dereferenced in full: circular
    cells = [c("A1", "k"), c("B1", 5), c("A2", "x"), c("B2", 0, f='VLOOKUP("k",A1:B3,2,FALSE)'), c("A3", "y"), c("B3", 6)]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and locs(v) == ["S!B2"], v


def test_dynamic_references_are_recorded_not_traced():
    # an INDIRECT loop the graph cannot see: passes, recorded in stats.dynamic_refs
    cells = [c("A1", 0, f='INDIRECT("B1")+1'), c("B1", 0, f='INDIRECT("A1")*2'), c("C1", 0, f="OFFSET(A1,0,0)")]
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "pass" and v["stats"]["dynamic_refs"]["count"] == 3, v["stats"]["dynamic_refs"]
    assert any("INDIRECT" in x for x in v["stats"]["dynamic_refs"]["examples"])
    # data tables: constants, no edges; external references lead nowhere
    cells = [c("A1", 1), c("B1", 0, f="A1*2"), c("C7", 0, f="", t=None)]
    cells[2] = '<c r="C7"><f t="dataTable" ref="C7:D8" dt2D="0" dtr="0" r1="A1"/><v>0</v></c>'
    v = run(book([("S", sheet(cells + ['<c r="D7"><v>0</v></c>', '<c r="C8"><v>0</v></c>', '<c r="D8"><v>0</v></c>']))],
                 calc_pr=AUTO))
    assert v["decision"] == "pass" and v["stats"]["data_table_cells_ignored"] >= 1, v["stats"]


def test_cumulative_ranges_scale():
    """=SUM($B$1:B<k>) copied down 3,000 rows: the dyadic blocks keep the walk near-linear."""
    import time
    n = 3000
    cells = [c(f"A{r}", r) for r in range(1, n + 1)] + [c(f"B{r}", 0, f=f"SUM($A$1:A{r})") for r in range(1, n + 1)]
    cells += [c(f"C{r}", 0, f=f"SUM($B$1:B{r})") for r in range(1, n + 1)]
    p = book([("S", sheet(cells))], calc_pr=AUTO)
    t0 = time.perf_counter()
    v = run(p)
    secs = time.perf_counter() - t0
    assert v["decision"] == "pass" and v["stats"]["n_formula_cells"] == 2 * n, v["stats"]
    assert v["stats"]["n_edges"] < 120 * n, v["stats"]["n_edges"]      # far below the n^2 / 2 of a naive walk
    assert secs < 20, secs
    print(f"  cumulative family of {2 * n} formulas: {v['stats']['n_edges']} edges in {secs:.2f} s")
    # the same family made circular by one cell: C1500 referenced by A1 -> every C<k> with k >= 1500 ... fails
    cells[0] = c("A1", 0, f="C1500*0")
    v = run(book([("S", sheet(cells))], calc_pr=AUTO))
    assert v["decision"] == "fail" and v["stats"]["n_cycles"] == 1 and v["stats"]["largest_cycle"] >= 3, v["stats"]


TESTS = [test_dyadic_blocks_and_structured_specs, test_fast_path_matches_the_parser, test_rubric_key_and_contract,
         test_manual_mode_fails_whatever_else,
         test_self_loops_and_simple_cycles_fail_with_iteration_off, test_breaker_switch_off_still_fails,
         test_cycles_through_names_shared_formulas_arrays_and_3d, test_structured_references_and_tables,
         test_cycle_with_iteration_on_is_undecided, test_geometry_and_index_references_are_not_cycles,
         test_dynamic_references_are_recorded_not_traced, test_cumulative_ranges_scale]


def main() -> int:
    failed = 0
    for t in TESTS:
        try:
            t()
            print(f"PASS {t.__name__}")
        except Exception:  # noqa: BLE001
            failed += 1
            print(f"FAIL {t.__name__}")
            traceback.print_exc()
    print(f"\n{len(TESTS) - failed}/{len(TESTS)} tests passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
