"""Unit tests for check 22 (No formula errors) on synthetic micro-workbooks (raw SpreadsheetML
zips labelled as saved by Excel, so their caches are the display; openpyxl-labelled where cache
trust matters).

    cd judge && python -m detchecks.tests.test_checks_22

Plain asserts; also collectable by pytest.  Temporary files go to detchecks/scratch/errors22/.
The workbook builder here is shared with test_recalc.py.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
from xml.sax.saxutils import escape, quoteattr

from detchecks.api import Engine, grade
from detchecks.checks import REGISTRY
from detchecks.checks.c22 import C22, implicit_intersection_hits, may_intersect_outside, rich_error_map
from detchecks.core.package import Package
from detchecks.errors import GradingError

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.path.join(os.path.dirname(HERE), "scratch", "errors22")
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
XL_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'     # Excel's declaration
PY_DECL = "<?xml version='1.0' encoding='UTF-8'?>"                          # Python writers
RUBRIC = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompts", "rubrics", "rubric_9.json")
APP_EXCEL = "<Application>Microsoft Excel</Application><AppVersion>16.0300</AppVersion>"
APP_OPX = "<Application>Microsoft Excel Compatible / Openpyxl 3.1.5</Application>"
APP_LO = "<Application>LibreOffice/25.8.7.3$MacOSX_AARCH64 LibreOffice_project/x</Application>"

_TMP = None
_N = [0]


def tmp(name: str = "") -> str:
    global _TMP
    if _TMP is None:
        os.makedirs(SCRATCH, exist_ok=True)
        _TMP = tempfile.mkdtemp(prefix="test_errors22_", dir=SCRATCH)
    _N[0] += 1
    return os.path.join(_TMP, f"{_N[0]:03d}_{name or 'book'}.xlsx")


# ============================================================================ raw workbook builder
def c(ref, v=None, t=None, f=None, *, si=None, fref=None, array=False, s=0, vm=None, cm=None):
    """One <c>.  v: cached value (number / str / bool / error text with t='e'); f: formula text
    (f='' with si: shared child; fref: shared master range or array ref); array: t="array"."""
    attrs = f'r="{ref}"' + (f' s="{s}"' if s else "") + (f' cm="{cm}"' if cm else "") + (f' vm="{vm}"' if vm else "")
    tt = t
    if tt is None and v is not None:
        tt = "str" if (isinstance(v, str) and f is not None) else "inlineStr" if isinstance(v, str) else \
            "b" if isinstance(v, bool) else None
    tattr = f' t="{tt}"' if tt else ""
    if f is not None:
        if array:
            fx = f'<f t="array" ref="{fref or ref}">{escape(f)}</f>'
        elif si is not None:
            fx = f'<f t="shared" si="{si}"/>' if f == "" else f'<f t="shared" ref="{fref}" si="{si}">{escape(f)}</f>'
        else:
            fx = f"<f>{escape(f)}</f>"
        if v is None:
            return f"<c {attrs}{tattr}>{fx}</c>"
        vv = "1" if v is True else "0" if v is False else escape(str(v))
        return f"<c {attrs}{tattr}>{fx}<v>{vv}</v></c>"
    if v is None:
        return f"<c {attrs}/>"
    if tt == "inlineStr":
        return f'<c {attrs} t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
    vv = "1" if v is True else "0" if v is False else escape(str(v))
    return f"<c {attrs}{tattr}><v>{vv}</v></c>"


def sheet(cells, *, head="", tail="", rows_attr=None):
    rows: dict = {}
    for x in cells:
        ref = re.search(r'r="([A-Z]+)(\d+)"', x)
        rows.setdefault(int(ref.group(2)), []).append((ref.group(1), x))
    rows_attr = rows_attr or {}
    for r in rows_attr:
        rows.setdefault(r, [])

    def colkey(col):
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return n
    body = "".join(f'<row r="{r}"{rows_attr.get(r, "")}>'
                   + "".join(x for _k, x in sorted(rows[r], key=lambda y: colkey(y[0]))) + "</row>"
                   for r in sorted(rows))
    return f"{head}<sheetData>{body}</sheetData>{tail}"


def book(sheets, *, app=APP_EXCEL, excel=True, states=None, names=None, extra_parts=None, extra_wb_rels="",
         styles=None, calc_pr='<calcPr calcId="191029"/>', name="", path=None):
    """sheets: [(name, inner_xml)].  excel=True writes Excel's signature (double-quoted
    declarations everywhere, fileVersion appName='xl') so the writer is 'excel' and caches are
    the display; excel=False writes Python-style declarations (writer from `app`)."""
    path = path or tmp(name)
    decl = XL_DECL if excel else PY_DECL
    states = states or {}
    extra_parts = extra_parts or {}
    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    ov = [f'<Override PartName="/xl/workbook.xml" ContentType="{CT_XLSX}"/>']
    ov += [f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT_WS}"/>' for i in range(len(sheets))]
    for k in extra_parts:
        if k == "xl/metadata.xml":
            ov.append('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>')
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>' + "".join(ov) + "</Types>")
    root_rels = [f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>']
    if app is not None:
        root_rels.append(f'<Relationship Id="rId2" Type="{REL}/extended-properties" Target="docProps/app.xml"/>')
        z.writestr("docProps/app.xml", decl + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                   + app + "</Properties>")
    z.writestr("_rels/.rels", f'<Relationships xmlns="{PKG_REL}">' + "".join(root_rels) + "</Relationships>")
    wrels, sh = [], []
    for i, (nm, xml) in enumerate(sheets):
        st = f' state="{states[i]}"' if i in states else ""
        sh.append(f'<sheet name={quoteattr(nm)} sheetId="{i + 1}"{st} r:id="rId{i + 1}"/>')
        wrels.append(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>')
        z.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'{decl}<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{xml}</worksheet>')
    if styles is not None:
        wrels.append(f'<Relationship Id="rIdS" Type="{REL}/styles" Target="styles.xml"/>')
        z.writestr("xl/styles.xml", f'<styleSheet xmlns="{MAIN}">{styles}</styleSheet>')
    wrels.append(extra_wb_rels)
    fv = '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="28000"/>' if excel else ""
    dn = ""
    if names:
        dn = "<definedNames>" + "".join(f'<definedName name={quoteattr(k)}>{escape(v)}</definedName>'
                                        for k, v in names.items()) + "</definedNames>"
    z.writestr("xl/workbook.xml", f'{decl}<workbook xmlns="{MAIN}" xmlns:r="{REL}">{fv}<sheets>' + "".join(sh)
               + f"</sheets>{dn}{calc_pr}</workbook>")
    z.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{PKG_REL}">' + "".join(wrels) + "</Relationships>")
    for k, v in extra_parts.items():
        z.writestr(k, v)
    z.close()
    return path


def run(path, **kw):
    return Engine(path, [C22()], **kw).run()[C22.key]


def graded(kind, fn, *needles):
    """Patrick 2026-10-05 (every attempt graded): where the check used to raise, it decides by a default
    recorded in stats.defaults[kind]; needles that name a location ('S!A1') must appear in that record."""
    import json
    v = fn()
    d = (v["stats"].get("defaults") or {}).get(kind)
    assert d and d["count"] >= 1, (kind, v["stats"].get("defaults"))
    blob = json.dumps(d)
    for n in needles:
        if "!" in n:
            assert n.split("!")[-1] in blob, (n, d)
    return v


def raises(fn, *needles):
    try:
        fn()
    except GradingError as e:
        msg = str(e) + " " + " ".join(str(v) for v in getattr(e, "failures", {}).values())
        for n in needles:
            assert n in msg, (n, msg[:800])
        return e
    raise AssertionError(f"expected GradingError containing {needles}")


def locs(v):
    return [m["location"] for m in v["mistakes"]]


# ============================================================================ tests
def test_rubric_key():
    with open(RUBRIC) as fh:
        rub = json.load(fh)
    items = rub["checks"] if isinstance(rub, dict) and "checks" in rub else rub
    if isinstance(items, dict):
        items = list(items.values())
    keys = []
    for it in items:
        if isinstance(it, dict):
            cat = it.get("category") or it.get("Category")
            nm = it.get("name") or it.get("check") or it.get("Check")
            if cat and nm:
                keys.append(f"{cat}/{nm}")
    if keys:
        assert C22.key in keys, C22.key
    assert REGISTRY[22] is C22


def test_cached_errors_fail_and_trapped_or_text_pass():
    cells = [c("A1", 1), c("A2", 0),
             c("B1", "#DIV/0!", t="e", f="A1/A2"),            # displayed error
             c("B2", 0, f="IFERROR(A1/A2,0)"),                 # trapped: shows 0
             c("B3", "#N/A", f='IF(A1>0,"#N/A","ok")'),        # formula's TEXT result
             c("B4", "#N/A"),                                  # text constant
             c("B5", "#N/A", t="e", f="NA()")]                 # deliberate =NA() counts
    v = run(book([("S", sheet(cells))]))
    assert v["decision"] == "fail" and sorted(locs(v)) == ["S!B1", "S!B5"], locs(v)
    assert v["stats"]["error_cells_by_code"] == {"#DIV/0!": 1, "#N/A": 1}
    assert "=A1/A2" in v["mistakes"][0]["description"]
    v = run(book([("S", sheet([c("A1", 1), c("B2", 0, f="IFERROR(A1/0,0)"), c("B3", "#N/A", f='"#N/A"'), c("B4", "#REF!")]))]))
    assert v["decision"] == "pass", locs(v)
    assert v["stats"]["formula_cells"] == 2 and v["live"] is True    # live since judge v15 (Patrick 2026-10-28)
    assert v["stats"]["decided_by"] == "code" and v["stats"]["gap_information"] == "none"


def test_typed_error_constant_broken_ref_misspelling():
    cells = [c("A1", "#N/A", t="e"),                               # typed error constant
             c("B1", "#REF!", t="e", f="E79-#REF!"),               # broken reference
             c("C1", "#NAME?", t="e", f='CONUTIFS($G$1:$Z$1,FALSE)&" false"')]
    v = run(book([("S", sheet(cells))]))
    assert sorted(locs(v)) == ["S!A1", "S!B1", "S!C1"], locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "typed error value #N/A" in d["S!A1"] and "#REF!" in d["S!B1"] and "#NAME?" in d["S!C1"]
    assert v["stats"]["typed_error_constants"] == 1


def test_hidden_sheet_row_column_and_concealment_count():
    styles = ('<numFmts count="1"><numFmt numFmtId="164" formatCode=";;;"/></numFmts>'
              '<fonts count="2"><font><sz val="11"/></font><font><color rgb="FFFFFFFF"/><sz val="11"/></font></fonts>'
              '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
              '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
              '<cellXfs count="3"><xf numFmtId="0" fontId="0"/><xf numFmtId="164" fontId="0" applyNumberFormat="1"/>'
              '<xf numFmtId="0" fontId="1" applyFont="1"/></cellXfs>')
    hid = sheet([c("A1", "#DIV/0!", t="e", f="1/0")])
    vis = sheet([c("A1", "#DIV/0!", t="e", f="1/0", s=1),                          # ;;; format conceals it
                 c("B1", "#NUM!", t="e", f="SQRT(-1)", s=2),                       # white font
                 c("C3", "#VALUE!", t="e", f='"a"+1')],                            # hidden row 3, hidden col C
                head='<cols><col min="3" max="3" width="0" hidden="1" customWidth="1"/></cols>',
                rows_attr={3: ' hidden="1"'})
    v = run(book([("V", vis), ("H", hid)], states={1: "veryHidden"}, styles=styles))
    assert sorted(locs(v)) == ["H!A1", "V!A1", "V!B1", "V!C3"], locs(v)
    assert v["stats"]["hidden_sheets_with_errors"] == ["H"]
    assert any("[hidden sheet]" in m["description"] for m in v["mistakes"])


def test_array_members_listed_once_per_cell():
    cells = [c("B7", 0), c("B2", 1), c("B3", 2), c("B4", 3),
             c("D2", "#DIV/0!", t="e", f="$B$2:$B$4/$B$7", array=True, fref="D2:D4"),
             c("D3", "#DIV/0!", t="e"), c("D4", "#DIV/0!", t="e")]
    v = run(book([("S", sheet(cells))]))
    assert locs(v) == ["S!D2:D4"] and v["stats"]["error_cells"] == 3, (locs(v), v["stats"])
    assert "$B$2:$B$4/$B$7" in v["mistakes"][0]["description"]


RICH_META = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" '
             'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
             '<metadataTypes count="2"><metadataType name="XLDAPR" minSupportedVersion="120000"/>'
             '<metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>'
             '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
             '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
             '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">'
             '<xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>'
             '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
             '<valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata></metadata>')
RICH_RV = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">'
           '<rv s="0"><v>0</v><v>{et}</v><v>2</v><v>1</v></rv></rvData>')
RICH_ST = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">'
           '<s t="_error"><k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/></s>'
           '</rvStructures>')


def test_blocked_spill_is_reported_as_spill():
    cells = [c("B2", 1), c("B3", 2), c("B7", 4),
             c("D2", "#VALUE!", t="e", f="$B$2:$B$6/$B$7", array=True, fref="D2", cm=1, vm=1),
             c("D4", 0.1742)]
    parts = {"xl/metadata.xml": RICH_META, "xl/richData/rdrichvalue.xml": RICH_RV.format(et=8),
             "xl/richData/rdrichvaluestructure.xml": RICH_ST}
    p = book([("Summary", sheet(cells))], extra_parts=parts)
    with Package.open(p) as pkg:
        assert rich_error_map(pkg) == {1: "#SPILL!"}
    v = run(p)
    assert locs(v) == ["Summary!D2"] and "#SPILL!" in v["mistakes"][0]["description"], v["mistakes"]
    assert v["stats"]["error_cells_by_code"] == {"#SPILL!": 1}
    # #CALC! (type 13); without rich parts the error is still counted, as a rich #VALUE!
    parts["xl/richData/rdrichvalue.xml"] = RICH_RV.format(et=13)
    v = run(book([("Summary", sheet(cells))], extra_parts=parts))
    assert "#CALC!" in v["mistakes"][0]["description"]
    v = run(book([("Summary", sheet(cells))]))
    assert v["decision"] == "fail" and "#VALUE! (rich error)" in v["mistakes"][0]["description"]


def test_implicit_intersection_detector_and_prefilter():
    assert may_intersect_outside("AVERAGE($I$8:$R$8/$I$7:$R$7)", 41, 8)          # H41 outside I:R
    assert not may_intersect_outside("AVERAGE($I$8:$R$8/$I$7:$R$7)", 41, 10)     # J41 inside
    assert may_intersect_outside("SUM(A1:A10)", 12, 1)                           # conservative: the detector decides
    assert may_intersect_outside("A1:A10*2", 12, 4) and not may_intersect_outside("A1:A10*2", 5, 4)
    assert not may_intersect_outside('"A1:A10"&B2', 12, 4)                       # inside a string
    assert not may_intersect_outside("SUM(A:A)", 12, 4)                          # whole column: never flagged
    assert implicit_intersection_hits("AVERAGE($I$8:$R$8/$I$7:$R$7)", "H41")
    assert not implicit_intersection_hits("AVERAGE($I$8:$R$8/$I$7:$R$7)", "J41")
    assert not implicit_intersection_hits("SUM(A1:A10)", "A12")                  # aggregator: fine


def test_implicit_intersection_toy_t5_pattern():
    row7 = [c(f"{col}7", 100 + i) for i, col in enumerate("IJKLMNOPQR")]
    row8 = [c(f"{col}8", 80 + i) for i, col in enumerate("IJKLMNOPQR")]
    base = row7 + row8
    # FAIL: legacy formula at H41 (outside I:R) with a numeric cache; Excel shows #VALUE!
    v = run(book([("Solution Model", sheet(base + [c("H41", 0.8625, f="AVERAGE($I$8:$R$8/$I$7:$R$7)")]))],
                 calc_pr='<calcPr calcId="191029" fullCalcOnLoad="1"/>'))
    assert locs(v) == ["'Solution Model'!H41"], locs(v)
    assert "implicit intersection" in v["mistakes"][0]["description"] and v["stats"]["implicit_intersection_cells"] == 1
    # PASS: the same text as a dynamic-array formula
    v = run(book([("Solution Model", sheet(base + [c("H41", 0.8625, f="AVERAGE($I$8:$R$8/$I$7:$R$7)", array=True, cm=1)]))]))
    assert v["decision"] == "pass", locs(v)
    # PASS: the legacy formula inside the range's columns (J41 intersects)
    v = run(book([("Solution Model", sheet(base + [c("J41", 0.8625, f="AVERAGE($I$8:$R$8/$I$7:$R$7)")]))]))
    assert v["decision"] == "pass", locs(v)
    # shared formula: master D12 and child E12 both outside rows 1..10 -> both flagged, one block
    cells = [c(f"A{i}", i) for i in range(1, 11)]
    cells += [c("D12", 2, f="A1:A10*2", si=0, fref="D12:E12"), c("E12", 2, f="", si=0)]
    v = run(book([("S", sheet(cells))]))
    assert locs(v) == ["S!D12:E12"], locs(v)
    # the same formula in row 5 intersects: pass
    cells = [c(f"A{i}", i) for i in range(1, 11)] + [c("D5", 10, f="A1:A10*2")]
    assert run(book([("S", sheet(cells))]))["decision"] == "pass"


def test_untrusted_values_raise_without_the_pipeline():
    # openpyxl-labelled file: its cache is never trusted -> GradingError (no fallback)
    p = book([("S", sheet([c("A1", "#DIV/0!", t="e", f="1/0")]))], app=APP_OPX, excel=False)
    graded("untrusted_value", lambda: run(p), "No formula errors", "untrusted", "S!A1")
    # a formula without any cached value in an Excel-labelled file is untrusted too
    p = book([("S", sheet([c("A1", None, f="1/0")]))])
    graded("untrusted_value", lambda: run(p), "untrusted")


def test_value_path_copy_supplies_values():
    """A manual value_path (LibreOffice copy) decides; its #NAME?/#VALUE! stay untrusted unless
    the recalc pipeline vetted them (test_recalc)."""
    p = book([("S", sheet([c("A1", 0, f="1/0"), c("A2", 0, f="SUM(1,2)")]))], app=APP_OPX, excel=False)
    lo = book([("S", sheet([c("A1", "#DIV/0!", t="e", f="1/0"), c("A2", 3, f="SUM(1,2)")]))], app=APP_LO, excel=False)
    v = run(p, value_path=lo)
    assert locs(v) == ["S!A1"] and v["stats"]["values"]["source"] == "value_path"
    lo2 = book([("S", sheet([c("A1", "#NAME?", t="e", f="1/0"), c("A2", 3, f="SUM(1,2)")]))], app=APP_LO, excel=False)
    graded("untrusted_value", lambda: run(p, value_path=lo2), "untrusted")


def test_live_flag_on_verdicts():
    """Judge v15: No formula errors (22) and Negatives in parentheses (65) are live (Patrick 2026-10-26 /
    2026-10-28); the per-check switch and its note stay for the next recorded-only check."""
    p = book([("S", sheet([c("A1", 1)]))])
    v = grade(p, checks=[22, 65, 92])
    for n in (22, 65, 92):
        assert v[REGISTRY[n].key]["live"] is True and "live_note" not in v[REGISTRY[n].key], n
    from detchecks.checks import live_numbers
    assert {22, 65, 92} <= set(live_numbers())
    saved = C22.live
    C22.live = False
    try:
        v = grade(p, checks=[22])[C22.key]
        assert v["live"] is False and v["live_note"] == "recorded only; the LLM verdict stands at scoring"
        assert 22 not in live_numbers()
    finally:
        C22.live = saved


def test_grouping_and_summary():
    cells = [c(f"G{r}", "#DIV/0!", t="e", f=f"$D{r}/$D$21") for r in range(18, 22)]
    cells += [c("B3", "#N/A", t="e", f="XLOOKUP(1,A1:A2,B1:B2)")]
    v = run(book([("Budget vs. Actuals Variance Ana", sheet(cells))]))
    assert locs(v) == ["'Budget vs. Actuals Variance Ana'!B3", "'Budget vs. Actuals Variance Ana'!G18:G21"], locs(v)
    assert "5 cell(s)" in v["summary"] and "#DIV/0! x4" in v["summary"] and v["stats"]["n_mistakes"] == 2


TESTS = [test_rubric_key, test_cached_errors_fail_and_trapped_or_text_pass,
         test_typed_error_constant_broken_ref_misspelling, test_hidden_sheet_row_column_and_concealment_count,
         test_array_members_listed_once_per_cell, test_blocked_spill_is_reported_as_spill,
         test_implicit_intersection_detector_and_prefilter, test_implicit_intersection_toy_t5_pattern,
         test_untrusted_values_raise_without_the_pipeline, test_value_path_copy_supplies_values,
         test_live_flag_on_verdicts, test_grouping_and_summary]


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
        if _TMP and os.path.isdir(_TMP):
            shutil.rmtree(_TMP, ignore_errors=True)
    print(f"\n{len(TESTS) - failed}/{len(TESTS)} tests passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
