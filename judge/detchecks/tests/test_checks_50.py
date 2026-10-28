"""Unit tests for Green font for cross-sheet links (50) added by its second review (2026-10-04).

The shared colour-rule tests (49 / 50 / 51) live in test_checks_colours.py; this module reuses
its micro-workbook helpers and only holds what is specific to 50.

    cd judge && python -m detchecks.tests.test_checks_50

Plain asserts; also collectable by pytest.
"""
from __future__ import annotations

import os
import shutil
import sys
import traceback

from detchecks.checks import colour_rules as CR
from . import test_checks_colours as T
from .test_checks_colours import BLACK, BLUE, GREEN, GREY50, WHITE, book, c, graded, locs, raises, run, sheet


def test_prefilter_rejects_shapes_that_cannot_be_pointers():
    """Finding 50-R2-prefilter-raise: a non-green formula that can never be a pure reference
    (a call other than the one-argument wrappers, two operands, unbalanced parentheses) is
    rejected by the pre-filter and never parsed, so an unknown table or malformed text in it
    cannot make 50 raise.  Every pointer form still passes the pre-filter."""
    pkg, clf = T._clf()
    try:
        pointers = ("Inputs!B5", "+Inputs!B5", "-Inputs!B5", "--Inputs!B5", "(Inputs!B5)", "( -( Inputs!B5 ) )",
                    "@Inputs!B5:B9", "Inputs!B5#", "+-@Inputs!B5#", "( Inputs!B5 )#", "Inputs!$A$5:$A$19",
                    "Inputs!B5:B9", "Inputs!XFD1048576", "'My Sheet'!A1", "Jan:Dec!B5", "Calc:Inputs!B5",
                    "_xlfn.ANCHORARRAY(Inputs!$H$641)", "_xlfn.SINGLE(Inputs!B5:B9)", "_xlfn.TOROW(Inputs!$C$5:$C$40)",
                    "_xlfn.TOCOL(Inputs!B5:C9)", "_xlfn.TRANSPOSE(Inputs!A1:B2)", "TRANSPOSE(Inputs!B5:B9)",
                    "_xlfn.TOROW((Inputs!A1:A5))", "_xludf.TOROW(Inputs!A1:A5)", "_xlfn._xlws.TOROW(Inputs!A1:A5)",
                    "_xlfn.TRANSPOSE(_xlfn.TOROW(Inputs!A1:A5))", "-_xlfn.TOROW(Inputs!A1:A5)",
                    "WACC", "-WACC", "WaccTwice", "Rng", "Rate?", "Tbl[Col]", "Tbl[[#This Row],[Col]]",
                    "Tbl[[#Totals],[Col]]", "Tbl[]", "Tbl", "=Inputs!B5", "= Inputs!B5")
        for t in pointers:
            assert clf.classify(t, "Calc").kind == CR.POINTER, (t, clf.classify(t, "Calc"))
            assert CR.pointer_shaped(t) and clf.may_be_pointer(t), t
        # never a pointer: rejected before any parse (several would raise in the classifier)
        never = ("SUM(Nope[Col])", "SUM(Inputs!A1:A5", "SUM(Inputs!A1:A5)", "Inputs!B5+Nope[Col]",
                 "Inputs!B5-Inputs!B6", "(Inputs!B5)-(Inputs!B6)", "Inputs!A1:A5 Inputs!B1:B5", "Inputs!A1 : A5",
                 "_xlfn.TOROW(Inputs!A1:A5,1)", "_xlfn.TOROW(SUM(Inputs!A1))", "_xlfn.TOROW(Nope[Col]",
                 "(Inputs!B5", "Inputs!B5)", "TOROW (Inputs!A1:A5)", "MYTOROW(Inputs!A1)", "TOROWX(Inputs!A1)",
                 "Sheet1!TOROW(Inputs!A1)", "IF(Inputs!A1,1,2)", "Inputs!B5*2", 'Inputs!B5&""', "INDEX(Inputs!B:B,5)",
                 "ROUND('Solution Model'!$D$10,2)", 'Instructions!B2&": "&Instructions!B4',
                 'HYPERLINK("#Inputs!A1","Go")', 'INDIRECT("Inputs!B5")', "(", "+", "")
        for t in never:
            assert not CR.pointer_shaped(t) and not clf.may_be_pointer(t), t
        # the parser tolerates a stray trailing operator and still sees one reference: such texts
        # stay accepted (the test is conservative with respect to the classifier)
        for t in ("Inputs!B5+", "Inputs!B5 +", "Nope[Col]+"):
            assert CR.pointer_shaped(t), t
    finally:
        pkg.close()
    # workbook level: black formulas that cannot be pointers never make 50 raise, whatever the
    # unknown table or the unbalanced text would do in the classifier; the real black link fails
    data = sheet((1, [c("A1", BLACK, "SUM(Nope[Col])"), c("B1", BLACK, "SUM(Inputs!A1:A5"),
                      c("C1", BLACK, "Inputs!B5+Nope[Col]"), c("D1", BLACK, "Inputs!A1:A5 Nope[Col]"),
                      c("E1", BLUE, "_xlfn.TOROW(Nope[Col]"), c("F1", WHITE, "Inputs!B5-Nope[Col]"),
                      c("G1", BLACK, "Inputs!B5")]))
    p = book("50_r2_prefilter.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 50)
    # v15: B1 (=SUM(Inputs!A1:A5, unbalanced) and G1 (=Inputs!B5) carry a direct reference to Inputs, so 50
    # charges them by the quick scan without a parse (the same scan 49 trusts to accept green cells); the
    # texts with an unknown table are parsed, fail to classify, and are skipped under the default
    assert set(locs(v)) == {"Calc!B1", "Calc!G1"}, locs(v)
    assert v["stats"]["cells_classified"] == 0 and v["stats"]["cells_charged_by_direct_reference"] == 2, v["stats"]
    assert v["stats"]["defaults"]["unparsable_formula"]["count"] == 5, v["stats"]["defaults"]
    # a text of pointer shape whose class cannot be worked out still raises (no guess)
    for t in ("Nope[Col]", "-Nope[Col]", "(Nope[Col])", "_xlfn.TOROW(Nope[Col])", "Nope[Col]#"):
        p = book("50_r2_pure_unknown.xlsx", [("Calc", sheet((1, [c("A1", BLACK, t)])), None), ("Inputs", sheet(), None)])
        graded("unparsable_formula", lambda: run(p, 50))          # skipped since Patrick 2026-10-05
    # green cells are never classified, whatever their text
    p = book("50_r2_green.xlsx", [("Calc", sheet((1, [c("A1", GREEN, "Nope[Col]"), c("B1", GREEN, "SUM(Nope[Col]")])), None),
                                  ("Inputs", sheet(), None)])
    assert run(p, 50)["decision"] == "pass"
    # 49 charges a colour that is wrong in every class: a grey unbalanced cross-sheet formula fails
    # as 'unclassified' (v15: whatever it references, grey is wrong; 50 skips what it cannot
    # classify, so the cell is charged once)
    p = book("50_r2_49_grey.xlsx", [("Calc", sheet((1, [c("A1", GREY50, "SUM(Inputs!A1:A5")])), None), ("Inputs", sheet(), None)])
    v = run(p, 49)
    assert locs(v) == ["Calc!A1"] and v["stats"]["unclassified_offending_cells"] == 1, (locs(v), v["stats"])
    p = book("50_r2_49_grey_tbl.xlsx", [("Calc", sheet((1, [c("A1", GREY50, "SUM(Nope[Col])")])), None)])
    v = run(p, 49)            # v15: grey is wrong in every class, so the unclassifiable cell fails 49 (charged once)
    assert locs(v) == ["Calc!A1"] and v["stats"]["unclassified_offending_cells"] == 1, (locs(v), v["stats"])
    assert run(p, 50)["decision"] == "pass"


TESTS = [test_prefilter_rejects_shapes_that_cannot_be_pointers]


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
        if T._TMP and os.path.isdir(T._TMP):
            shutil.rmtree(T._TMP, ignore_errors=True)
    print(f"\n{len(TESTS) - failed}/{len(TESTS)} tests passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
