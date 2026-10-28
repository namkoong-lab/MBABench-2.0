"""Unit tests for the font-colour checks 49 / 50 / 51 (group "colours") on synthetic
micro-workbooks (raw OOXML zips, so every colour form and formula kind is under control).

    cd judge && python -m detchecks.tests.test_checks_colours

Plain asserts; also collectable by pytest.  Temporary files go to detchecks/scratch/colours/.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback
import zipfile

from detchecks.api import grade
from detchecks.checks import REGISTRY, colour_rules as CR
from detchecks.checks.c49 import C49
from detchecks.checks.c50 import C50
from detchecks.checks.c51 import C51
from detchecks.core.package import Package
from detchecks.errors import GradingError

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.path.join(os.path.dirname(HERE), "scratch", "colours")
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
K49, K50, K51 = C49.key, C50.key, C51.key

# font id == cellXfs index (style s) for every entry below
FONTS = [
    "",                                         # 0 no <color>: black
    '<color rgb="FF0000FF"/>',                  # 1 blue
    '<color rgb="FF00B050"/>',                  # 2 green
    '<color rgb="FFFF0000"/>',                  # 3 red
    '<color theme="1" tint="0.49998"/>',        # 4 Text 1 lighter 50% = 808080 grey (49/T4 trap)
    '<color theme="1"/>',                       # 5 dk1 = black
    '<color indexed="64"/>',                    # 6 system foreground = black
    '<color auto="1"/>',                        # 7 automatic = black
    '<color rgb="00000000"/>',                  # 8 openpyxl alpha 00 = black
    '<color theme="9"/>',                       # 9 accent6 70AD47 = green (50/T4 Pass)
    '<color theme="4"/>',                       # 10 accent1 4472C4 = blue (50/T4 Fail)
    '<color rgb="FF404040"/>',                  # 11 Text 1 lighter 25%: black band
    '<color rgb="FF595959"/>',                  # 12 Text 1 lighter 35%: not black
    '<color rgb="FFC00000"/>',                  # 13 dark red: red
    '<color rgb="FFFF6600"/>',                  # 14 orange: not red
    '<color theme="99"/>',                      # 15 unresolvable theme slot
    '<color indexed="12"/>',                    # 16 legacy palette 12 = 0000FF blue
    '<color theme="0"/>',                       # 17 lt1 = white
]
BLACK, BLUE, GREEN, RED, GREY50, DK1, IDX64, AUTO, A00, ACC6, ACC1, G404, G595, DRED, ORANGE, BADTHEME, IDX12, WHITE = range(18)
STYLES = ("<fonts count=\"%d\">%s</fonts>" % (len(FONTS), "".join(f"<font><sz val=\"11\"/>{c}<name val=\"Calibri\"/></font>"
                                                                  for c in FONTS))
          + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          + '<cellXfs count="%d">%s</cellXfs>' % (len(FONTS), "".join(
              f'<xf numFmtId="0" fontId="{i}" fillId="0" borderId="0" xfId="0" applyFont="1"/>' for i in range(len(FONTS))))
          + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
          # dxfs for the conditional-format excuse (Patrick 2026-10-06): 0 green "Good" font, 1 red "Bad"
          # font, 2 black font, 3 unresolvable font colour, 4 fill only (no font colour)
          + '<dxfs count="5"><dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
          + '<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
          + '<dxf><font><color rgb="FF000000"/></font></dxf><dxf><font><color theme="99"/></font></dxf>'
          + '<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>')
DXF_GREEN, DXF_RED, DXF_BLACK, DXF_BAD, DXF_FILL = range(5)


def cf(sqref, *dxf_ids, op="equal", values=('"OK"',)):
    """<conditionalFormatting> with one cellIs rule per dxf id (each comparing to the given values)."""
    rules = "".join(f'<cfRule type="cellIs" dxfId="{d}" priority="{i + 1}" operator="{op}">'
                    + "".join(f"<formula>{esc(v)}</formula>" for v in values) + "</cfRule>"
                    for i, d in enumerate(dxf_ids))
    return f'<conditionalFormatting sqref="{sqref}">{rules}</conditionalFormatting>'

_TMP = None


def tmp(name: str) -> str:
    global _TMP
    if _TMP is None:
        os.makedirs(SCRATCH, exist_ok=True)
        _TMP = tempfile.mkdtemp(prefix="test_colours_", dir=SCRATCH)
    return os.path.join(_TMP, name)


def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def c(ref, s=0, f=None, *, v="1", t=None, fattrs="", cm=None):
    """One <c>: f is the formula text (None = no <f>, '' = an empty <f/> marker)."""
    a = f' r="{ref}" s="{s}"' + (f' t="{t}"' if t else "") + (f' cm="{cm}"' if cm else "")
    fx = "" if f is None else (f"<f{fattrs}/>" if f == "" else f"<f{fattrs}>{esc(f)}</f>")
    vx = "" if v is None else f"<v>{v}</v>"
    return f"<c{a}>{fx}{vx}</c>"


def cstr(ref, text, s=0):
    """A constant inline-string cell."""
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{esc(text)}</t></is></c>'


def sheet(*rows_cells):
    """rows_cells: (row number, [cell xml, ...]) pairs -> <sheetData>."""
    return "<sheetData>" + "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in rows_cells) + "</sheetData>"


def book(name, sheets, *, names=(), ext_links=0, tables=None):
    """sheets: [(sheet name, sheetData xml, state)]; names: [(name, text, localSheetId|None)];
    ext_links: number of externalLink parts; tables: {sheet index: [(table name, ref, column)]}."""
    path = tmp(name)
    tables = tables or {}
    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    ov = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    ov += [f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
           for i in range(len(sheets))]
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>' + "".join(ov) + "</Types>")
    z.writestr("_rels/.rels", f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{RT}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wrels, sh = [], []
    tno = 0
    for i, (nm, data, state) in enumerate(sheets):
        st = f' state="{state}"' if state else ""
        sh.append(f'<sheet name="{esc(nm)}" sheetId="{i + 1}"{st} r:id="rId{i + 1}"/>')
        wrels.append(f'<Relationship Id="rId{i + 1}" Type="{RT}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>')
        tparts = ""
        srels = []
        for (tname, tref, col) in tables.get(i, []):
            tno += 1
            z.writestr(f"xl/tables/table{tno}.xml",
                       f'{DECL}<table xmlns="{MAIN}" id="{tno}" name="{tname}" displayName="{tname}" ref="{tref}">'
                       f'<tableColumns count="1"><tableColumn id="1" name="{col}"/></tableColumns></table>')
            srels.append(f'<Relationship Id="rIdT{tno}" Type="{RT}/table" Target="../tables/table{tno}.xml"/>')
            tparts += f'<tablePart r:id="rIdT{tno}"/>'
        if srels:
            z.writestr(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels", f'<Relationships xmlns="{PKG_REL}">{"".join(srels)}</Relationships>')
            tparts = f'<tableParts count="{len(srels)}">{tparts}</tableParts>'
        z.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'{DECL}<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{data}{tparts}</worksheet>')
    wrels.append(f'<Relationship Id="rIdS" Type="{RT}/styles" Target="styles.xml"/>')
    z.writestr("xl/styles.xml", f'{DECL}<styleSheet xmlns="{MAIN}">{STYLES}</styleSheet>')
    extra = ""
    if ext_links:
        refs = []
        for k in range(1, ext_links + 1):
            wrels.append(f'<Relationship Id="rIdE{k}" Type="{RT}/externalLink" Target="externalLinks/externalLink{k}.xml"/>')
            refs.append(f'<externalReference r:id="rIdE{k}"/>')
            z.writestr(f"xl/externalLinks/externalLink{k}.xml",
                       f'{DECL}<externalLink xmlns="{MAIN}" xmlns:r="{REL}"><externalBook r:id="rId1"><sheetNames>'
                       f'<sheetName val="Prices"/></sheetNames></externalBook></externalLink>')
            z.writestr(f"xl/externalLinks/_rels/externalLink{k}.xml.rels",
                       f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{RT}/externalLinkPath" '
                       f'Target="file:///C:\\Data\\Book{k}.xlsx" TargetMode="External"/></Relationships>')
        extra += "<externalReferences>" + "".join(refs) + "</externalReferences>"
    if names:
        extra += "<definedNames>" + "".join(
            f'<definedName name="{n}"{"" if sid is None else f" localSheetId=\"{sid}\""}>{esc(txt)}</definedName>'
            for n, txt, sid in names) + "</definedNames>"
    z.writestr("xl/workbook.xml", f'{DECL}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{"".join(sh)}</sheets>{extra}</workbook>')
    z.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{PKG_REL}">{"".join(wrels)}</Relationships>')
    z.close()
    return path


def run(path, n):
    return grade(path, checks=[n])[REGISTRY[n].key]


def locs(v):
    return [m["location"] for m in v["mistakes"]]


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
        msg = str(e) + " " + " ".join(str(x) for x in (e.failures or {}).values())
        for nd in needles:
            assert nd in msg, (nd, msg[:600])
        return e
    raise AssertionError(f"expected GradingError containing {needles}")


# ============================================================================ colour families
def test_colour_families():
    fam = CR.family
    for hx in ("000000", "1A1A1A", "0D0D26", "262626", "404040", "333333"):
        assert fam(hx) == "black", hx
    for hx in ("595959", "808080", "A6A6A6", "D0CECE", "FFFFFF", "0000FF", "0070C0", "4472C4", "002060"):
        assert fam(hx) == "other", (hx, fam(hx))
    for hx in ("00B050", "70AD47", "008000", "00FF00", "92D050", "548235", "006100", "375623", "339966"):
        assert fam(hx) == "green", hx
    for hx in ("FF0000", "C00000", "C0504D", "9C0006", "800000", "FF3B3B"):
        assert fam(hx) == "red", hx
    for hx in ("FF6600", "FF99CC", "FFC7CE", "ED7D31"):
        assert fam(hx) != "red", hx
    assert CR.colour_name("FF0000FF") == "blue" and CR.colour_name("FF808080") == "grey"
    assert CR.describe_colour("FF4472C4") == "blue (4472C4)"


# ============================================================================ classification
def _clf():
    p = book("clf.xlsx", [("Calc", sheet(), None), ("Inputs", sheet(), None), ("My Sheet", sheet(), None),
                          ("Jan", sheet(), None), ("Dec", sheet(), None)],
             names=[("WACC", "Inputs!$B$3", None), ("ConstName", "0.08", None), ("LocalCalc", "Calc!$A$1", 0),
                    ("ExtName", "[1]Prices!$A$1", None), ("WaccTwice", "WACC", None),
                    ("Rng", "Inputs!$B$3:$B$9", None), ("Rate?", "Inputs!$B$4", None), ("LocalIn", "0.5", 1)],
             ext_links=1, tables={1: [("Tbl", "A1:A4", "Col")]})
    pkg = Package.open(p)
    return pkg, CR.Classifier(pkg)


def test_classifier_kinds():
    pkg, clf = _clf()
    try:
        def k(text, own="Calc"):
            return clf.classify(text, own).kind
        # own sheet / no reference -> OWN
        for t in ("B2*C2", "Calc!B2*2", "'calc'!B2", "Calc!B2", "5", "TODAY()", '"text"', "SUM(A1:A9)",
                  'IF(A1,"Inputs!B5",0)', 'HYPERLINK("#Inputs!A1","Go")', "ConstName*2", "ConstName", "LocalCalc",
                  "INDIRECT(\"B\"&ROW())*2"):
            assert k(t) == CR.OWN, (t, k(t))
        # pure pointers to another sheet
        for t in ("Inputs!B5", "+Inputs!B5", "(Inputs!B5)", "-Inputs!B5", "Inputs!B5#", "@Inputs!B5:B9",
                  "_xlfn.ANCHORARRAY(Inputs!B5)", "_xlfn.SINGLE(Inputs!B5:B9)", "_xlfn.TOROW(Inputs!B5:B9)",
                  "_xlfn.TOCOL(Inputs!B5:C9)", "TRANSPOSE(Inputs!B5:B9)", "'My Sheet'!A1", "Inputs!$A$5:$A$19",
                  "WACC", "WaccTwice", "Rng", "Tbl[Col]", "Jan:Dec!B5", "( -( Inputs!B5 ) )"):
            assert k(t) == CR.POINTER, (t, k(t))
        # calculations that read another sheet
        for t in ("Inputs!B5*C3", "SUM(Inputs!A1:A9)", "Inputs!B5+0", "_xlfn.TOROW(Inputs!B5:B9,1)", "WACC*B2",
                  "SUM(Tbl[Col])", "SUM(Jan:Dec!B5)", "INDIRECT(\"Inputs!B5\")*2", "INDIRECT(\"Inputs!B5\")",
                  "Inputs!B5+Inputs!B6", "Inputs!#REF!", "INDEX(Inputs!B:B,5)"):
            assert k(t) == CR.CROSS, (t, k(t))
        # another workbook (also a pure pointer to it, and via a defined name)
        for t in ("[1]Prices!$C$5", "[1]Prices!$B$2*C25", "ExtName*2", "ExtName", "'C:\\Data\\[Book.xlsx]Prices'!$C$5",
                  "[1]!Rate", "Inputs!B5+[1]Prices!A1", "INDIRECT(\"'[Book.xlsx]Prices'!A1\")"):
            assert k(t) == CR.EXTERNAL, (t, k(t))
        # self-qualified on the other sheet: Inputs!B5 on Inputs is OWN
        assert k("Inputs!B5", "Inputs") == CR.OWN and k("Tbl[@Col]", "Inputs") == CR.OWN
        assert k("Tbl[Col]", "Inputs") == CR.OWN
        # computed INDIRECT address: unknown, flagged
        c_ = clf.classify("INDIRECT(A20)*2", "Calc")
        assert c_.kind == CR.OWN and c_.indirect_unresolved
        assert c_.indirect_cells == ((20, 1),)                   # the address is A20's content
        # ... but a literal qualifier before the computed part fixes the sheet (gpt6 1340/1570/1601)
        for t, kind in (("INDIRECT(\"'My Sheet'!\"&ADDRESS(G$4+1,$F8+1,4))", CR.CROSS),
                        ("INDIRECT(\"'My \"&\"Sheet'!I\"&(D10+7))", CR.CROSS),
                        ("INDIRECT(\"Calc!B\"&ROW())*2", CR.OWN),
                        ("INDIRECT(\"'[Book.xlsx]Prices'!B\"&ROW())", CR.EXTERNAL)):
            c_ = clf.classify(t, "Calc")
            assert c_.kind == kind and not c_.indirect_unresolved, (t, c_)
        assert not clf.classify("INDIRECT(\"Inputs!B5\")", "Calc").indirect_unresolved
        assert clf.classify("Inputs!B5*C3", "Calc").other_sheets == ("Inputs",)
        # unknown table, malformed formula: loud
        raises(lambda: clf.classify("SUM(Nope[Col])", "Calc"), "table 'Nope' is not defined")
        raises(lambda: clf.classify("SUM(Inputs!B5", "Calc"))
    finally:
        pkg.close()


def test_prefilters_are_conservative():
    pkg, clf = _clf()
    try:
        # may_be_pointer False -> never POINTER; may_read_other_book False -> never EXTERNAL
        texts = ["Inputs!B5", "Inputs!B5*2", 'IF(A1,"[x]",Inputs!B5)', "WACC", "WACC*2", "ExtName", "B2*C2",
                 "'My Sheet'!A1", "-Inputs!B5", "[1]Prices!A1", "'C:\\Data\\[Book.xlsx]Prices'!$C$5",
                 "INDIRECT(\"'[Book.xlsx]Prices'!A1\")", 'IF(A1,"a!b",2)', "Tbl[Col]", "Other.xlsx!Rate",
                 "INDIRECT(\"Inputs!B5\")", "SUM(Inputs!A1:A9)/2", "A1/'My Sheet'!B2", "Rate?", "-Rate?"]
        for t in texts:
            kind = clf.classify(t, "Calc").kind
            if kind == CR.POINTER:
                assert clf.may_be_pointer(t), t
            if kind == CR.EXTERNAL:
                assert clf.may_read_other_book(t), t
        assert clf.classify("Rate?", "Calc").kind == CR.POINTER and clf.may_be_pointer("Rate?")
        assert not clf.may_be_pointer("Inputs!B5*2") and not clf.may_be_pointer("B2")
        assert not clf.may_read_other_book("SUM(Inputs!A1:A9)/2")
        assert not clf.may_read_other_book('IF(A1,"[Act/360]",Inputs!B5)')       # '[' only inside a string
        assert clf.may_read_other_book('INDIRECT("[B.xlsx]S!A1")')               # ... unless INDIRECT
    finally:
        pkg.close()


# ============================================================================ check 49
def test_49_rules():
    data = sheet(
        (1, [c("A1", BLUE, "B2*C2"),             # OWN blue -> fail
             c("B1", GREEN, "Calc!B2*2"),        # self-qualified OWN in green -> fail (49/T3)
             c("C1", GREEN, "Inputs!B5*C3"),     # CROSS green -> pass
             c("D1", BLUE, "Inputs!B5*C3"),      # CROSS blue -> 50's business since v15 (was 49's)
             c("E1", BLACK, "Inputs!B5"),        # POINTER black -> 49 passes (50's business)
             c("F1", BLUE, "Inputs!B5"),         # POINTER blue -> not judged by 49 (50 fails it; ruling 2026-10-03)
             c("G1", RED, "[1]Prices!A1"),       # EXTERNAL red -> pass
             c("H1", GREEN, "[1]Prices!A1*2"),   # EXTERNAL green -> not judged here (51 fails it)
             c("I1", GREY50, "B2+1"),            # theme 1 + tint 0.5 = 808080 -> fail (49/T4)
             ]),
        (2, [c("A2", DK1, "B2+1"), c("B2", IDX64, "B2+1"), c("C2", AUTO, "B2+1"), c("D2", A00, "B2+1"),
             c("E2", G404, "B2+1"), c("F2", BLACK, "B2+1"),          # all black forms pass (49/T4 Pass)
             c("G2", G595, "B2+1"),              # 595959 -> fail
             c("H2", IDX12, "B2+1"),             # legacy indexed blue -> fail
             c("I2", BLUE, None, v="5"),         # a blue constant is not a formula -> ignored
             c("J2", BLUE, "5"),                 # constant-only formula is a formula: OWN -> fail
             ]),
    )
    p = book("49_rules.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)], ext_links=1)
    v = run(p, 49)
    assert v["decision"] == "fail"
    # rectangles are grouped per (reason, colour): A1 blue and B1 green stay separate
    assert sorted(locs(v)) == sorted(["Calc!A1", "Calc!B1", "Calc!I1", "Calc!G2", "Calc!H2", "Calc!J2"]), locs(v)
    st = v["stats"]
    assert st["offending_cells"] == 6, st
    assert st["offending_cells_by_reason"] == {"own": 6}, st
    assert st["pointer_cells_left_to_check_50"] == 1 and st["external_cells_left_to_check_51"] == 1, st
    assert st["cross_cells_left_to_check_50"] == 1, st                 # D1 (C1 is green: accepted unparsed)
    assert set(locs(run(p, 50))) == {"Calc!D1", "Calc!E1", "Calc!F1"}, locs(run(p, 50))   # D1 cross, E1 / F1 pointers
    txt = " ".join(m["description"] for m in v["mistakes"])
    assert "grey (808080)" in txt and "blue (0000FF)" in txt and "=B2*C2" in txt


def test_49_arrays_shared_datatable_hidden():
    data = sheet(
        # dynamic-array anchor black, two spilled members green: members are judged (OWN class)
        (1, [c("A1", BLACK, "SEQUENCE(3)", fattrs=' t="array" ref="A1:A3"', cm="1")]),
        (2, [c("A2", GREEN)]),
        (3, [c("A3", GREEN)]),
        # shared formula: master black, one child green (49/T2: child of =AAK$169+AAK$170)
        (5, [c("B5", BLACK, "B$3+B$4", fattrs=' t="shared" ref="B5:E5" si="0"'),
             c("C5", BLACK, "", fattrs=' t="shared" si="0"'), c("D5", GREEN, "", fattrs=' t="shared" si="0"'),
             c("E5", BLACK, "", fattrs=' t="shared" si="0"')]),
        # what-if data table: anchor + members, one member blue
        (7, [c("C7", BLACK, "", fattrs=' t="dataTable" ref="C7:D8" dt2D="0" dtr="0" r1="A1"'), c("D7", BLACK)]),
        (8, [c("C8", BLACK), c("D8", BLUE)]),
        # empty <f/> marker outside any array: not a formula cell
        (10, [c("A10", BLUE, "")]),
        # array anchor with no formula text and no value (ChatGPT writer, gpt6 2592): no formula
        (12, [c("A12", BLUE, "", v=None, fattrs=' t="array" ref="A12:A13"', cm="1")]),
        (13, [c("A13", BLUE, None, v=None)]),
    )
    hidden = sheet((1, [c("A1", BLUE, "B1*2")]))
    p = book("49_kinds.xlsx", [("Calc", data, None), ("Hid", hidden, "hidden")])
    v = run(p, 49)
    assert v["decision"] == "fail"
    assert set(locs(v)) == {"Calc!A2:A3", "Calc!D5", "Calc!D8", "Hid!A1"}, locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "spilled from A1" in d["Calc!A2:A3"] and "=D$3+D$4" in d["Calc!D5"], d     # child text expanded
    assert "data table" in d["Calc!D8"] and "(hidden sheet)" in d["Hid!A1"], d
    assert v["stats"]["empty_f_markers_ignored"] == 1 and v["stats"]["textless_array_cells_ignored"] == 2, v["stats"]
    for n in (50, 51):
        assert run(p, n)["stats"]["textless_array_cells_ignored"] == 2


def test_49_pass_and_contract():
    data = sheet((1, [c("A1", BLACK, "B1*2"), c("B1", GREEN, "Inputs!A1"), c("C1", GREEN, "Inputs!A1*B1"),
                      c("D1", BLACK, "Inputs!A1"), c("E1", BLUE, None, v="3")]))
    p = book("49_pass.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 49)
    assert v["decision"] == "pass" and v["mistakes"] == [] and v["stats"]["n_mistakes"] == 0, v
    assert v["engine"] == "harness" and v["stats"]["formula_cells"] == 4


def test_49_no_fallback():
    # an unresolvable font colour on a formula cell: read as automatic (black) (Patrick 2026-10-05: every
    # attempt graded; it raised until then), recorded in stats.defaults
    p = book("49_badtheme.xlsx", [("Calc", sheet((1, [c("A1", BADTHEME, "B1*2")])), None)])
    for n in (49, 50, 51):               # an own-sheet calculation in black passes all three
        v = run(p, n)
        assert v["decision"] == "pass" and v["stats"]["font_families"]["unresolved"] == 1, v
        d = v["stats"]["defaults"]["unresolved_colour"]
        assert d["count"] == 1 and "'Calc'!A1" in d["examples"][0] and "automatic (black)" in d["examples"][0], d
    # green formula whose INDIRECT address is computed from empty cells: nothing names a sheet -> skipped
    # (Patrick 2026-10-06, every attempt graded; it raised until then), recorded in stats.defaults
    p = book("49_indirect.xlsx", [("Calc", sheet((1, [c("A1", GREEN, "INDIRECT(A20&B20)*2")])), None),
                                  ("Inputs", sheet(), None)])
    graded("indirect_address_unknown", lambda: run(p, 49))
    # ... but a BLUE one is wrong whatever it reads, and a black one is fine
    p = book("49_indirect2.xlsx", [("Calc", sheet((1, [c("A1", BLUE, "INDIRECT(A20)*2"),
                                                       c("B1", BLACK, "INDIRECT(A20)*2")])), None)])
    assert locs(run(p, 49)) == ["Calc!A1"]
    # malformed formula text: raises only where the decision needs the class (green), not on black
    # cells, and not on blue ones (blue is wrong in every class: the cell fails)
    p = book("49_malformed.xlsx", [("Calc", sheet((1, [c("A1", GREEN, "SUM(C16:C22")])), None)])
    graded("unparsable_formula", lambda: run(p, 49))
    p = book("49_malformed_black.xlsx", [("Calc", sheet((1, [c("A1", BLACK, "SUM(C16:C22")])), None)])
    assert run(p, 49)["decision"] == "pass"


# ============================================================================ check 50
def test_50_rules():
    data = sheet(
        (1, [c("A1", BLACK, "Inputs!B5"),                 # pointer black -> fail
             c("B1", GREEN, "Inputs!B6"),                 # pointer green -> pass
             c("C1", ACC6, "Inputs!B7"),                  # theme accent6 70AD47 -> green -> pass (50/T4)
             c("D1", ACC1, "_xlfn.TOROW(Inputs!$C$5:$C$40)"),   # accent1 blue TOROW pull -> fail (50/T4 Fail)
             c("E1", BLACK, "'Calc'!G66+Calc!G85"),       # self-qualified own sheet: not a link (50/T3)
             c("F1", BLACK, "Inputs!B5*C3"),              # cross-sheet calculation in black -> fail (v15 rule)
             c("G1", BLACK, "[1]Prices!A1"),              # external pointer: 51's business
             c("H1", BLACK, "-Inputs!B8"),                # sign-flipped link -> fail
             c("I1", BLACK, "WACC"),                      # name on another sheet -> fail
             c("J1", BLACK, "Tbl[Col]"),                  # table on another sheet -> fail
             c("K1", BLACK, "Calc!K2"),                   # self-qualified pointer: own sheet, not a link
             c("L1", WHITE, "Inputs!B9"),                 # white title link -> fail (no light-on-dark exemption)
             ]),
        # spill: anchor black pure pull, members green -> only the anchor (50/T2)
        (3, [c("A3", BLACK, "_xlfn.ANCHORARRAY(Inputs!$H$641)", fattrs=' t="array" ref="A3:C3"', cm="1"),
             c("B3", GREEN), c("C3", GREEN)]),
        # spill: anchor + members all black (50/T1) -> one block
        (5, [c("A5", BLACK, "Inputs!$A$5:$A$7", fattrs=' t="array" ref="A5:A7"', cm="1")]),
        (6, [c("A6", BLACK)]),
        (7, [c("A7", BLACK)]),
    )
    p = book("50_rules.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)],
             names=[("WACC", "Inputs!$B$3", None)], ext_links=1, tables={1: [("Tbl", "A1:A4", "Col")]})
    v = run(p, 50)
    assert set(locs(v)) == {"Calc!A1", "Calc!D1", "Calc!F1", "Calc!H1:J1", "Calc!L1", "Calc!A3", "Calc!A5:A7"}, locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "blue (4472C4)" in d["Calc!D1"] and "white (FFFFFF)" in d["Calc!L1"]
    assert "calculation that references another sheet; font black (000000), must be green" in d["Calc!F1"], d["Calc!F1"]
    assert "link to another sheet" in d["Calc!A1"], d["Calc!A1"]
    assert v["stats"]["offending_cells_by_reason"] == {"pointer": 10, "cross": 1}, v["stats"]
    assert "1 calculation(s)" in v["summary"] and "10 pure link(s)" in v["summary"], v["summary"]


def test_50_shared_child_off_grid_and_pass():
    # shared master C2 =Inputs!A1048576; its child C3 is pushed off the grid (Inputs!#REF!): not a
    # pointer, so only C2 is flagged (guards the master-class shortcut; Excel itself never writes this)
    data = sheet((2, [c("C2", BLACK, "Inputs!A1048576", fattrs=' t="shared" ref="C2:C3" si="0"')]),
                 (3, [c("C3", BLACK, "", fattrs=' t="shared" si="0"')]),
                 (5, [c("C5", BLACK, "Inputs!A1", fattrs=' t="shared" ref="C5:C6" si="1"')]),
                 (6, [c("C6", BLACK, "", fattrs=' t="shared" si="1"')]))
    p = book("50_shared.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 50)
    # v15: the off-grid child =Inputs!#REF! still references Inputs, so it is a black cross-sheet formula
    assert locs(v) == ["Calc!C2", "Calc!C5:C6", "Calc!C3"], locs(v)
    assert "=Inputs!A2" in v["mistakes"][1]["description"] or "=Inputs!A1" in v["mistakes"][1]["description"]
    assert "calculation that references another sheet" in v["mistakes"][2]["description"], v["mistakes"][2]
    # no formula reads another sheet in a wrong colour -> pass, whatever the own-sheet formulas' colours
    p = book("50_pass.xlsx", [("Calc", sheet((1, [c("A1", BLUE, "B1*2"), c("B1", GREEN, "Inputs!A1*2"),
                                                  c("C1", GREEN, "Inputs!A1"), c("D1", BLACK, "SUM(B1:C1)"),
                                                  c("E1", BLACK, "Calc!B1*2")])), None),
                              ("Inputs", sheet(), None)])
    v = run(p, 50)
    assert v["decision"] == "pass", v
    # D1 never leaves its sheet (not parsed); the self-qualified E1 is parsed and found OWN (49's)
    assert v["stats"]["cells_classified"] == 1 and v["stats"]["own_sheet_cells_left_to_check_49"] == 1, v["stats"]
    # ... and a black cross-sheet calculation among own-sheet formulas fails on its own (v15)
    p = book("50_cross_black.xlsx", [("Calc", sheet((1, [c("A1", BLUE, "B1*2"), c("B1", BLACK, "Inputs!A1*2"),
                                                         c("C1", GREEN, "Inputs!A1")])), None), ("Inputs", sheet(), None)])
    v = run(p, 50)
    assert locs(v) == ["Calc!B1"] and v["stats"]["offending_cells_by_reason"] == {"cross": 1}, v


# ============================================================================ check 51
def test_51_rules():
    data = sheet((1, [c("A1", RED, "[1]Prices!$G$41"), c("B1", BLACK, "[1]Prices!$H$41"),   # one of three black (51/T2)
                      c("C1", RED, "[1]Prices!$I$41"),
                      c("D1", GREEN, "[1]Prices!$C$12"),              # green external (51/T3)
                      c("E1", BLACK, "[1]Prices!$B$2*C25"),           # inside a calculation (51/T4)
                      c("F1", DRED, "[1]Prices!A1"),                  # C00000 is red
                      c("G1", ORANGE, "[1]Prices!A2"),                # FF6600 is not
                      c("H1", BLACK, "ExtName*2"),                    # via a defined name
                      c("I1", BLACK, "'C:\\Data\\[Book1.xlsx]Prices'!A1"),   # path form
                      c("J1", BLACK, 'INDIRECT("\'[Book1.xlsx]Prices\'!A1")'),  # literal INDIRECT
                      c("K1", GREEN, "Inputs!A1"),                    # cross-sheet link: not external (51/T3 Pass)
                      c("L1", BLACK, 'IF(A1,"[1]Prices!A1",0)'),      # string literal: not a reference
                      ]))
    p = book("51_rules.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)],
             names=[("ExtName", "[1]Prices!$A$1", None)], ext_links=1)
    v = run(p, 51)
    assert sorted(locs(v)) == sorted(["Calc!B1", "Calc!D1", "Calc!E1", "Calc!G1", "Calc!H1:J1"]), locs(v)
    assert v["stats"]["offending_cells"] == 7 and v["stats"]["external_link_parts"][0]["index"] == 1


def test_51_default_pass():
    # no formula reads another workbook -> pass (satisfied by default), also with no formula at all
    p = book("51_none.xlsx", [("Calc", sheet((1, [c("A1", BLACK, "Inputs!A1"), c("B1", BLUE, "B2")])), None),
                              ("Inputs", sheet(), None)])
    v = run(p, 51)
    assert v["decision"] == "pass" and "satisfied" in v["summary"], v
    p = book("51_empty.xlsx", [("Calc", sheet((1, [c("A1", BLACK, None, v="1")])), None)])
    assert run(p, 51)["decision"] == "pass"


def test_all_three_together():
    data = sheet((1, [c("A1", BLUE, "[1]Prices!A1"), c("B1", BLACK, "Inputs!A1"), c("C1", GREEN, "B1*2")]))
    p = book("all.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)], ext_links=1)
    v = grade(p, checks=[49, 50, 51])
    assert locs(v[K49]) == ["Calc!C1"], locs(v[K49])          # the blue external A1 is 51's only
    assert locs(v[K50]) == ["Calc!B1"] and locs(v[K51]) == ["Calc!A1"], (locs(v[K50]), locs(v[K51]))


def test_rubric_keys():
    import json
    r = json.load(open(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompts", "rubrics", "rubric_9.json")))
    keys = [f"{cat}/{it['name']}" for cat, items in r.items() for it in items]
    for cls in (C49, C50, C51):
        assert keys[cls.number - 1] == cls.key and REGISTRY[cls.number] is cls


# ============================================================================ review fixes (2026-10-03)
def test_structured_refs_with_commas_are_pointers():
    """Finding 50-prefilter-structured-comma: Excel stores [@Col] as Tbl[[#This Row],[Col]]; the
    pre-filter must not drop such links because of the comma inside the brackets."""
    pkg, clf = _clf()
    try:
        forms = ("Tbl[[#This Row],[Col]]", "Tbl[[#Totals],[Col]]", "Tbl[[#Headers],[Col]]", "Tbl[[#Data],[Col]]",
                 "-Tbl[[#This Row],[Col]]", "( Tbl[[#This Row],[Col]] )", "Tbl[[#All],[Col]]")
        for t in forms:
            assert clf.classify(t, "Calc").kind == CR.POINTER, (t, clf.classify(t, "Calc"))
            assert clf.may_be_pointer(t), t
        for t in ("Tbl[[#This Row],[Col]]*2", 'Tbl[[#This Row],[Col]]&"x"', "SUM(Tbl[[#This Row],[Col]],1)"):
            assert clf.classify(t, "Calc").kind == CR.CROSS and not clf.may_be_pointer(t), t
    finally:
        pkg.close()
    data = sheet((2, [c("A2", BLACK, "Tbl[[#This Row],[Col]]"), c("B2", BLACK, "Tbl[[#Totals],[Col]]"),
                      c("C2", BLACK, "Tbl[[#Headers],[Col]]"), c("D2", BLACK, "Tbl[Col]"),
                      c("E2", BLUE, "Tbl[[#This Row],[Col]]"), c("F2", GREEN, "Tbl[[#This Row],[Col]]"),
                      c("G2", BLACK, "Tbl[[#This Row],[Col]]*2")]))
    p = book("50_struct_commas.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)],
             tables={1: [("Tbl", "A1:A4", "Col")]})
    assert set(locs(run(p, 50))) == {"Calc!A2:D2", "Calc!E2", "Calc!G2"}, locs(run(p, 50))   # G2: black CROSS (v15)
    assert locs(run(p, 49)) == []                             # the blue pointer E2 is 50's only


def test_leading_equals_pointer():
    """Finding 50-prefilter-leading-equals: a '=' stored inside <f> (non-Excel writer) is tolerated
    by the parser, so the pre-filter must tolerate it too."""
    data = sheet((1, [c("A1", BLACK, "=Inputs!B5"), c("B1", GREEN, "=Inputs!B6"), c("C1", BLACK, "=B2*2")]))
    p = book("50_leading_eq.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    assert locs(run(p, 50)) == ["Calc!A1"]
    assert run(p, 49)["decision"] == "pass"


def test_bare_table_name():
    """Finding bare-table-name-unresolved: =SUM(Tbl) / =Tbl read the table's sheet like Tbl[]."""
    pkg, clf = _clf()
    try:
        assert clf.classify("SUM(Tbl)", "Calc").kind == CR.CROSS
        assert clf.classify("Tbl", "Calc").kind == CR.POINTER and clf.may_be_pointer("Tbl")
        assert clf.classify("Tbl", "Inputs").kind == CR.OWN
        assert clf.classify('INDIRECT("Tbl")*2', "Calc").kind == CR.CROSS
        assert clf.classify("Nope*2", "Calc").kind == CR.OWN          # an unknown name reads nothing
    finally:
        pkg.close()
    data = sheet((1, [c("A1", GREEN, "SUM(Tbl)"), c("B1", BLACK, "Tbl"), c("C1", GREEN, "SUM(Tbl[])"),
                      c("D1", BLACK, "Tbl[]"), c("E1", GREEN, "Tbl"), c("F1", BLACK, "SUM(Tbl)*2")]))
    p = book("49_bare_table.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)],
             tables={1: [("Tbl", "A1:A4", "Col")]})
    assert run(p, 49)["decision"] == "pass", locs(run(p, 49))
    assert set(locs(run(p, 50))) == {"Calc!B1", "Calc!D1", "Calc!F1"}, locs(run(p, 50))   # F1: =SUM(Tbl)*2 black (v15)


def test_unresolved_colour_raises_only_where_needed():
    """An unresolvable font colour is read as automatic (black) (Patrick 2026-10-05: every attempt graded;
    until then 49 / 50 / 51 raised where the cell's class made the colour decide): a black calculation passes
    49, a black pointer fails 50, a black external link fails 51 - each recorded in stats.defaults."""
    data = sheet((1, [c("A1", BADTHEME, "SUM(B2:B9)"), c("B1", BLACK, "Inputs!B5"), c("C1", IDX12, "C2*2")]))
    p = book("unres_own.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 49)                                              # A1 read as black passes; C1 (indexed 12) fails
    assert locs(v) == ["Calc!C1"] and v["stats"]["defaults"]["unresolved_colour"]["count"] == 1, v["stats"]
    assert locs(run(p, 50)) == ["Calc!B1"]                      # B1: a real black pointer
    assert run(p, 51)["decision"] == "pass"
    p = book("unres_ptr.xlsx", [("Calc", sheet((1, [c("A1", BADTHEME, "Inputs!B5")])), None), ("Inputs", sheet(), None)])
    v = run(p, 50)
    assert locs(v) == ["Calc!A1"] and "black" in v["mistakes"][0]["description"], v
    assert run(p, 51)["decision"] == "pass"
    p = book("unres_ext.xlsx", [("Calc", sheet((1, [c("A1", BADTHEME, "[1]Prices!A1*2")])), None)], ext_links=1)
    assert locs(run(p, 51)) == ["Calc!A1"]
    assert run(p, 50)["decision"] == "pass"


def test_unknown_scope_names_raise_only_where_needed():
    """Finding raise-where-class-cannot-matter (b): a defined name whose localSheetId names no
    sheet raises only where its possible definitions could change the decision."""
    data = sheet((1, [c("A1", BLACK, "Rate*B1"), c("B1", BLACK, "Rate"), c("C1", GREEN, "Rate*2")]))
    p = book("badscope_const.xlsx", [("Calc", data, None)], names=[("Rate", "0.05", 7)])
    assert run(p, 50)["decision"] == "pass" and run(p, 51)["decision"] == "pass"
    assert locs(run(p, 49)) == ["Calc!C1"]                     # =Rate*2 reads nothing whatever the scope
    # a scope-less name pointing at another sheet
    names = [("Ptr", "Inputs!$B$1", 7), ("XB", "[1]Prices!$A$1", 7)]
    data = sheet((1, [c("A1", BLACK, "Ptr*2"), c("B1", GREEN, "Ptr*2+Inputs!A1"), c("C1", BLUE, "Ptr*2"),
                      c("D1", BLACK, "XB*2+[1]Prices!A2"), c("E1", RED, "XB*2+[1]Prices!A2")]))
    p = book("badscope_ref.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)], names=names, ext_links=1)
    v = run(p, 49)
    assert locs(v) == ["Calc!C1"]                              # blue fails whatever Ptr is (reason "unsure", v15)
    assert v["stats"]["offending_cells_by_reason"] == {"unsure": 1} and "wrong either way" in v["mistakes"][0]["description"]
    # v15: the black =Ptr*2 (A1) reads Inputs or nothing depending on the scope -> 50 cannot decide (no fallback);
    # the blue one (C1) is 49's charge, the direct externals are 51's
    raises(lambda: run(p, 50), "cannot be decided", "A1")
    assert locs(run(p, 51)) == ["Calc!D1"]                     # reads [1] directly anyway
    for f_, n, needle in (("Ptr", 50, "whether it is a link"), ("Ptr*2", 49, "cannot be decided"),
                          ("Ptr*2", 50, "cannot be decided"), ("XB*2", 51, "may read another workbook")):
        fam = GREEN if n == 49 else BLACK
        p = book(f"badscope_{n}.xlsx", [("Calc", sheet((1, [c("A1", fam, f_)])), None), ("Inputs", sheet(), None)],
                 names=names, ext_links=1)
        raises(lambda: run(p, n), needle)
    pkg = Package.open(book("badscope_clf.xlsx", [("Calc", sheet(), None), ("Inputs", sheet(), None)], names=names,
                            ext_links=1))
    try:
        clf = CR.Classifier(pkg)
        c_ = clf.classify("Ptr*2+Inputs!A1", "Calc")
        assert c_.kind == CR.CROSS and c_.unsure and c_.direct_other and not c_.unsure_ext, c_
        c_ = clf.classify("Ptr", "Calc")
        assert c_.unsure and c_.pure_name, c_
        c_ = clf.classify("XB*2", "Calc")
        assert c_.unsure_ext and not c_.direct_ext, c_
        c_ = clf.classify('Ptr*2+INDIRECT("Inputs!A1")', "Calc")
        assert c_.direct_other, c_                              # a literal INDIRECT is direct too
    finally:
        pkg.close()
    # a scope-less constant that may shadow a global pointer: =Rate may or may not be a link
    p = book("badscope_shadow.xlsx", [("Calc", sheet((1, [c("A1", BLACK, "Rate")])), None), ("Inputs", sheet(), None)],
             names=[("Rate", "0.05", 7), ("Rate", "Inputs!$B$2", None)])
    raises(lambda: run(p, 50), "whether it is a link")


def test_49_colour_wrong_in_every_class_needs_no_class():
    """Finding raise-where-class-cannot-matter (c): blue / grey / white fail in every class, so a
    formula that cannot be classified still fails (no raise); green needs the class."""
    data = sheet((1, [c("A1", BLUE, "SUM(C16:C22"), c("B1", GREY50, "SUM(C16:C22"), c("C1", BLUE, "B9*2")]))
    p = book("49_other_unclassified.xlsx", [("Calc", data, None)])
    v = run(p, 49)
    assert set(locs(v)) == {"Calc!A1", "Calc!B1", "Calc!C1"}, locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "could not be worked out" in d["Calc!A1"] and "only its own sheet" in d["Calc!C1"], d
    assert v["stats"]["unclassified_offending_cells"] == 2 and v["stats"]["offending_cells_by_reason"] == \
        {"unclassified": 2, "own": 1}, v["stats"]
    # v15: grey is wrong in every class (OWN here, CROSS / POINTER under 50, EXTERNAL under 51), so a grey formula
    # that cannot be classified fails here whatever its text ('[' included); 50 and 51 skip it, so it is charged
    # once.  Green needs the class (OWN fails, the rest pass) and a green one is skipped (default); red is right
    # for an external link, so a red one that may read another workbook needs its class too
    p = book("49_grey_unknown_table.xlsx", [("Calc", sheet((1, [c("A1", GREY50, "SUM(Nope[Col])")])), None)])
    v = run(p, 49)
    assert locs(v) == ["Calc!A1"] and v["stats"]["unclassified_offending_cells"] == 1, (locs(v), v["stats"])
    assert run(p, 50)["decision"] == "pass" and run(p, 51)["decision"] == "pass"
    p = book("49_green_unknown_table.xlsx", [("Calc", sheet((1, [c("A1", GREEN, "SUM(Nope[Col])")])), None)])
    graded("unparsable_formula", lambda: run(p, 49), "table 'Nope' is not defined")
    p = book("49_red_unknown_table.xlsx", [("Calc", sheet((1, [c("A1", RED, "SUM(Nope[Col])")])), None)])
    graded("unparsable_formula", lambda: run(p, 49), "table 'Nope' is not defined")


def test_indirect_addresses():
    """Finding 49-indirect-decidable-raises: syntactically known addresses are worked out; an
    address held in one own-sheet constant cell is read in a second pass."""
    pkg, clf = _clf()
    try:
        def k(t):
            return clf.classify(t, "Calc")
        for t, kind in (('INDIRECT("B"&ROW())*2', CR.OWN), ('INDIRECT("B"&(ROW()+1)&":C"&ROW())', CR.OWN),
                        ('INDIRECT(ADDRESS(ROW(),2,4,TRUE,"Inputs"))*1', CR.CROSS),
                        ('INDIRECT(ADDRESS(ROW(),2,4,TRUE,"Calc"))*1', CR.OWN),
                        ('INDIRECT(ADDRESS(1,1,,,"My Sheet"))', CR.CROSS), ("INDIRECT(ADDRESS(1,1))", CR.OWN),
                        ('INDIRECT(ADDRESS(1,1)&":"&ADDRESS(5,2))', CR.OWN), ("INDIRECT(ROW())", CR.OWN),
                        ('INDIRECT(ADDRESS(1,1,,,"[Book.xlsx]Prices"))', CR.EXTERNAL),
                        ('INDIRECT("R"&ROW()&"C2",FALSE)', CR.OWN)):
            assert k(t).kind == kind and not k(t).indirect_unresolved, (t, k(t))
        for t in ('INDIRECT("Rate"&ROW())', 'INDIRECT("B"&A1)', "INDIRECT(ADDRESS(1,1,1,TRUE,A1))",
                  'INDIRECT(TEXT(ROW(),"0"))'):
            assert k(t).kind == CR.OWN and k(t).indirect_unresolved and k(t).indirect_cells is None, (t, k(t))
        assert k("INDIRECT($A$20)+INDIRECT(Calc!B3)").indirect_cells == ((20, 1), (3, 2))
        assert k('INDIRECT(A20)+INDIRECT(B1&"x")').indirect_cells is None
        # Patrick 2026-10-06 (attempts 1458, 1725): a literal anywhere in the address that names another
        # sheet decides (CROSS); a literal '!' qualifier with a computed name is assumed CROSS; own-sheet
        # cells that are parts of an unknown address are read in the second pass
        r = k('IFERROR(INDIRECT(MID($I33,FIND("Inputs!",$I33),LEN($I33))),"")')
        assert r.kind == CR.CROSS and not r.indirect_assumed and r.other_sheets == ("Inputs",), r
        r = k('INDIRECT(MID($I33,FIND("Calc!",$I33),LEN($I33)))')
        assert r.kind == CR.OWN and not r.indirect_unresolved, r
        r = k('INDIRECT("\'"&B3&"\'!"&ADDRESS(22,D3))')
        assert r.kind == CR.CROSS and r.indirect_assumed and r.other_sheets == (), r
        r = k('INDIRECT(B3&"!"&ADDRESS(22,D3))')
        assert r.kind == CR.CROSS and r.indirect_assumed, r
        r = k('INDIRECT(MID(I33,FIND("!",I33)+1,99))')          # "!" inside FIND locates text: not a qualifier
        assert r.kind == CR.OWN and r.indirect_unresolved and r.indirect_cells is None \
            and r.indirect_part_cells == ((33, 9),), r
        assert k('INDIRECT("B"&A1)').indirect_part_cells == ((1, 1),)
    finally:
        pkg.close()
    data = sheet((1, [c("A1", GREEN, 'INDIRECT(ADDRESS(ROW(),2,4,TRUE,"Inputs"))*1'),
                      c("B1", GREEN, 'INDIRECT(ADDRESS(ROW(),2,4,TRUE,"Calc"))*1'),
                      c("C1", GREEN, 'INDIRECT("B"&ROW())*2')]))
    p = book("49_indirect_syntactic.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    assert locs(run(p, 49)) == ["Calc!B1:C1"], locs(run(p, 49))
    # address cells: A20 'Inputs!B2' (reads Inputs: green fine), B20 'C5' (own: fails), C20 empty
    # (#REF!, reads nothing: fails); shared master E1 =INDIRECT(A20) passes, its child F1
    # =INDIRECT(B20) fails; a quoted sheet in the address text counts too
    data = sheet((1, [c("A1", GREEN, "INDIRECT(A20)*2"), c("B1", GREEN, "INDIRECT(B20)*2"),
                      c("C1", GREEN, "INDIRECT(C20)*2"),
                      c("E1", GREEN, "INDIRECT(A$20)", fattrs=' t="shared" ref="E1:F1" si="0"'),
                      c("F1", GREEN, "", fattrs=' t="shared" si="0"'), c("G1", GREEN, "INDIRECT(G20)")]),
                 (20, [cstr("A20", "Inputs!B2"), cstr("B20", "C5"), cstr("G20", "'My Sheet'!A1")]))
    p = book("49_indirect_cells.xlsx", [("Calc", data, None), ("Inputs", sheet(), None), ("My Sheet", sheet(), None)])
    v = run(p, 49)
    assert set(locs(v)) == {"Calc!B1:C1", "Calc!F1"}, locs(v)
    assert v["stats"]["indirect_address_cells_read"] == 6, v["stats"]
    # the address comes from a formula cell: its value would be needed -> skipped, recorded (Patrick
    # 2026-10-06, every attempt graded; it raised until then)
    data = sheet((1, [c("A1", GREEN, "INDIRECT(A20)*2")]), (20, [c("A20", BLACK, '"Inputs!B"&2', t="str", v="Inputs!B2")]))
    p = book("49_indirect_formula_cell.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = graded("indirect_address_unknown", lambda: run(p, 49), "Calc!A20")
    assert v["decision"] == "pass", v
    # Patrick 2026-10-06 (attempts 1458 / 1725): A1 reads a sheet named in a literal inside MID/FIND
    # (passes, no default); B1's address is sheet-qualified with the name computed (passes, default
    # indirect_sheet_unverified); C1's address is cut from a constant cell whose text names Inputs
    # (passes); D1's from one that names no sheet (skipped, default); E1 is a plain own-sheet
    # address in green (fails, as before)
    data = sheet((1, [c("A1", GREEN, 'IFERROR(INDIRECT(MID($A$20,FIND("Inputs!",$A$20),LEN($A$20))),"")'),
                      c("B1", GREEN, "INDIRECT(\"'\"&B20&\"'!\"&ADDRESS(22,3))"),
                      c("C1", GREEN, "INDIRECT(MID(A20,5,99))"),
                      c("D1", GREEN, "INDIRECT(MID(C20,1,2))"),
                      c("E1", GREEN, 'INDIRECT("B"&ROW())*2')]),
                 (20, [cstr("A20", "see Inputs!B2"), c("B20", BLACK, '"Inp"&"uts"', t="str", v="Inputs"),
                       cstr("C20", "C5")]))
    p = book("49_indirect_text.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 49)
    assert locs(v) == ["Calc!E1"], locs(v)
    d = v["stats"]["defaults"]
    assert d["indirect_sheet_unverified"]["count"] == 1 and "B1" in d["indirect_sheet_unverified"]["examples"][0], d
    assert d["indirect_address_unknown"]["count"] == 1 and "no address part" in d["indirect_address_unknown"]["examples"][0], d


def test_49_green_direct_reference_needs_no_parse():
    """Finding C49-PERF-1: a green formula with a direct reference to another sheet is accepted
    without a full parse; the quick test never claims more than the classifier."""
    data = sheet((1, [c("A1", GREEN, "Inputs!B5*C3"), c("B1", GREEN, "SUM(Inputs!C16:C22"),   # malformed: accepted
                      c("C1", GREEN, "'My Sheet'!A1+1"), c("D1", GREEN, "Calc!B2*2"),            # own sheet: parsed, fails
                      c("E1", GREEN, 'IF(A1,"Inputs!B5",0)'), c("F1", GREEN, "Inputs!LocalIn*2")]))
    p = book("49_quick.xlsx", [("Calc", data, None), ("Inputs", sheet(), None), ("My Sheet", sheet(), None)],
             names=[("LocalIn", "0.5", 1)])
    v = run(p, 49)
    assert set(locs(v)) == {"Calc!D1:F1"}, locs(v)
    assert v["stats"]["green_cells_accepted_by_direct_reference"] == 3, v["stats"]
    pkg, clf = _clf()
    try:
        yes = ("Inputs!B5", "Inputs!B5*C3", "'My Sheet'!$A$1:B9", "Jan:Dec!B5", "[1]Prices!A1", "-Inputs!A1",
               "SUM(Inputs!A1:A9)/2", "A1/'My Sheet'!B2", "IF(Inputs!A1>0,1,2)", "@Inputs!B5:B9", "'[1]Prices'!A1*2",
               "x+'C:\\Data\\[Book.xlsx]Prices'!$C$5", "Inputs!XFD1048576",
               # INDIRECT whose address starts with a literal qualifier of another sheet / book
               "INDIRECT(\"'My Sheet'!B\"&(ROW()+1))*2", 'LET(_xlpm.x,INDIRECT( "Inputs!B" & A1 ),_xlpm.x)',
               "INDIRECT(\"'[B.xlsx]S'!B\"&ROW())")
        no = ("Calc!B2*2", "'calc'!B2", "Calc:Calc!B5", 'IF(A1,"Inputs!B5",0)', "Inputs!LocalIn*2", "Inputs!XFE1",
              "Inputs!A0", "Inputs!#REF!", "Tbl[Col]", "WACC*2", 'INDIRECT("Inputs!B5")', "Inputs!R1C1",
              'HYPERLINK("#Inputs!A1","Go")', "B2*C2",
              'INDIRECT("Calc!B"&ROW())', 'INDIRECT("Inputs!B"&"5")', 'IF(A1,"INDIRECT(""Inputs!B""&1)",0)',
              'INDIRECT("B"&ROW())', 'XINDIRECT("Inputs!B"&1)', 'INDIRECT("Inputs"&"!B5")')
        for t in yes:
            assert CR.quick_reads_other_sheet(t, "Calc"), t
            assert clf.classify(t, "Calc").kind != CR.OWN, t
        for t in no:
            assert not CR.quick_reads_other_sheet(t, "Calc"), t
    finally:
        pkg.close()


# ============================================================================ decisions 2026-10-03
def test_hyperlink_cells_any_colour():
    """Ruling 2026-10-03 (A1): a formula whose top-level call is HYPERLINK(...) may be any colour:
    ignored by 49, 50 and 51.  HYPERLINK nested in a calculating expression is an ordinary
    formula (a literal-only IFERROR / IFNA / IF guard is not one: see
    test_hyperlink_guarded_navigation_cells, 2026-10-04)."""
    assert CR.is_hyperlink_formula('HYPERLINK("#\'Cover\'!A1","Cover")')
    assert CR.is_hyperlink_formula('=HYPERLINK("#Inputs!A1", "Inputs")')
    assert CR.is_hyperlink_formula(' HYPERLINK(Inputs!B5)')
    assert CR.is_hyperlink_formula('HYPERLINK([1]Prices!A1,"x")')
    assert CR.is_hyperlink_formula('IFERROR(HYPERLINK("#A1"),"")')          # literal-only guard (2026-10-04)
    for t in ('HYPERLINK("#A1","x")&"y"', '(HYPERLINK("#A1","x"))', '+HYPERLINK("#A1","x")',
              'IF(A1,HYPERLINK("#A1","x"),"")', 'MYHYPERLINK("#A1")',
              'HYPERLINK("#A1"', "Inputs!B5", ""):
        assert not CR.is_hyperlink_formula(t), t
    data = sheet((1, [c("A1", BLUE, 'HYPERLINK("#\'Cover\'!A1","Cover")', t="str", v="Cover"),     # own -> any colour
                      c("B1", GREEN, 'HYPERLINK("#Inputs!A1","Inputs")', t="str", v="Inputs"),
                      c("C1", BLUE, 'HYPERLINK(Inputs!B5)', t="str", v="x"),        # pointer-like -> any colour
                      c("D1", BLACK, 'HYPERLINK([1]Prices!A1,"x")', t="str", v="x"),  # external -> any colour
                      c("E1", WHITE, 'HYPERLINK("#A1","x")', t="str", v="x")]),
                 (2, [c("A2", BLUE, 'IF(B9>0,HYPERLINK("#A1","x"),"")', t="str", v="x"),   # nested: OWN blue
                      c("B2", BLUE, 'HYPERLINK("#A1","x")&""', t="str", v="x")]))         # not top-level: OWN blue
    p = book("hyperlinks.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)], ext_links=1)
    v = grade(p, checks=[49, 50, 51])
    assert locs(v[K49]) == ["Calc!A2:B2"], locs(v[K49])
    assert v[K50]["decision"] == "pass" and v[K51]["decision"] == "pass", (locs(v[K50]), locs(v[K51]))
    assert v[K49]["stats"]["hyperlink_cells_ignored"] == 4, v[K49]["stats"]      # D1 is black: never looked at
    assert v[K51]["stats"]["hyperlink_cells_ignored"] == 5, v[K51]["stats"]      # D1 is not red: ignored too


def test_pointers_judged_only_by_50():
    """Ruling 2026-10-03 (A6): a pure pointer to another sheet is judged by 50 only (49 judges
    own-sheet formulas and cross-sheet calculations); externals by 51 only."""
    data = sheet((1, [c("A1", BLUE, "Inputs!B5"), c("B1", RED, "Inputs!B5"), c("C1", WHITE, "-Inputs!B6"),
                      c("D1", BLUE, "Inputs!B5*2"), c("E1", WHITE, "Inputs!B5+C1"),
                      c("F1", BLUE, "[1]Prices!A1"), c("G1", GREEN, "[1]Prices!A1")]))
    p = book("pointers_50_only.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)], ext_links=1)
    v = grade(p, checks=[49, 50, 51])
    assert locs(v[K49]) == [], locs(v[K49])                               # v15: cross calcs blue / white are 50's
    assert set(locs(v[K50])) == {"Calc!A1", "Calc!B1", "Calc!C1", "Calc!D1", "Calc!E1"}, locs(v[K50])
    assert v[K50]["stats"]["offending_cells_by_reason"] == {"pointer": 3, "cross": 2}, v[K50]["stats"]
    assert set(locs(v[K51])) == {"Calc!F1:G1"} or set(locs(v[K51])) == {"Calc!F1", "Calc!G1"}, locs(v[K51])
    # an unresolvable font colour on a pointer does not matter to 49 (it does to 50)
    p = book("pointer_bad_colour.xlsx", [("Calc", sheet((1, [c("A1", BADTHEME, "Inputs!B5")])), None), ("Inputs", sheet(), None)])
    assert run(p, 49)["decision"] == "pass"
    assert locs(run(p, 50)) == ["Calc!A1"]                                 # read as black (Patrick 2026-10-05)
    p = book("own_bad_colour.xlsx", [("Calc", sheet((1, [c("A1", BADTHEME, "B5*2")])), None), ("Inputs", sheet(), None)])
    assert run(p, 49)["decision"] == "pass"


def test_light_text_on_dark_fill_no_exemption():
    """Ruling 2026-10-03 (A2): white / light header formulas follow the normal colour rule
    (the font colour alone counts; the fill is never looked at)."""
    data = sheet((1, [c("A1", WHITE, '"FY"&B9', t="str", v="FY"),       # own-sheet header calc in white -> 49 fails
                      c("B1", WHITE, "Tax!B4"),                        # header link in white -> 50 fails
                      c("C1", GREEN, "Tax!C4")]))
    p = book("white_headers.xlsx", [("FCFF", data, None), ("Tax", sheet(), None)])
    v = grade(p, checks=[49, 50])
    assert locs(v[K49]) == ["FCFF!A1"] and locs(v[K50]) == ["FCFF!B1"], (locs(v[K49]), locs(v[K50]))



# ============================================================================ second review (2026-10-04)
def test_hyperlink_guarded_navigation_cells():
    """Finding 49-R2-1: the ruling 'HYPERLINK cells any colour' (not a calculation) also covers a
    navigation link guarded by IFERROR / IFNA / IF whose other arguments are literals, e.g. the
    ChatGPT writer's =IFERROR(HYPERLINK("#'Summary'!A1","OPEN"),"OPEN") (gpt6 2004, 2017, ...).
    A guard argument that reads a cell or calculates, or any other expression around the link,
    keeps the cell an ordinary formula."""
    yes = ('IFERROR(HYPERLINK("#\'Summary\'!A1","OPEN"),"OPEN")',
           '=IFERROR(HYPERLINK("#\'Summary\'!A1","Open "&B5),"Open")',      # the link's own arguments are free
           'IF(TRUE,HYPERLINK("#A1","go"),"")', 'IF(TRUE,HYPERLINK("#A1","go"),)',
           '_xlfn.IFNA(HYPERLINK("#A1","x"),"x")', 'IFERROR(IFERROR(HYPERLINK("#A1","x"),"y"),"z")',
           'IFERROR( HYPERLINK("#A1","x") , "x" )', 'IFERROR(HYPERLINK("#A1","x"),0)',
           'IFERROR(HYPERLINK("#A1","x"),TRUE)', 'IFERROR(HYPERLINK("#A1","x"),#N/A)',
           'IF(1,HYPERLINK("#A1","a"),HYPERLINK("#B1","b"))', 'iferror(hyperlink("#A1","go"),"go")')
    no = ('IF(B9>0,HYPERLINK("#A1","x"),"")', 'IF(A1,HYPERLINK("#A1","x"),"")',      # condition reads a cell
          'IFERROR(HYPERLINK("#A1","x"),C1*2)', 'IFERROR(HYPERLINK("#A1","x"),Rate)',  # fallback calculates / reads
          'IFERROR(HYPERLINK("#A1","x"),"a"&"b")', 'IFERROR(HYPERLINK("#A1","x"),-1)',
          'IFERROR(HYPERLINK("#A1","x"),(1))', 'IFERROR(HYPERLINK("#A1","x")&"","x")',
          'IFERROR(HYPERLINK("#A1","x"),"y")&""', 'IFERROR(HYPERLINK("#A1","x"),"y")#',
          'IFERROR(IF(J7=0,"No",IF(D7="x",HYPERLINK("#A1","y"),"z")),"w")',          # gpt6 2497: conditional link
          'CONCAT(HYPERLINK("#A1","x"))', 'SUM(IFERROR(HYPERLINK("#A1","x"),1))', 'IFS(TRUE,HYPERLINK("#A1","x"))',
          '_xludf.HYPERLINK("#A1")', '@HYPERLINK("#A1","x")', 'IFERROR(1,"HYPERLINK(")', 'IFERROR("x","y")',
          'IFERROR(HYPERLINK("#A1","x"', None)
    for t in yes:
        assert CR.is_hyperlink_formula(t), t
    for t in no:
        assert not CR.is_hyperlink_formula(t), t
    nav = 'IFERROR(HYPERLINK("#\'Summary\'!A1","OPEN"),"OPEN")'
    data = sheet((1, [c("A1", GREEN, nav, t="str", v="OPEN"),                            # own-sheet target, green
                      c("B1", BLUE, 'IFERROR(HYPERLINK("#Inputs!A1","Inputs"),"Inputs")', t="str", v="Inputs"),
                      c("C1", WHITE, 'IF(TRUE,HYPERLINK("#A1","go"),"")', t="str", v="go"),
                      c("D1", GREEN, '_xlfn.IFNA(HYPERLINK(Inputs!B5,"x"),"x")', t="str", v="x"),  # link built from another sheet
                      c("E1", BLUE, 'IFERROR(HYPERLINK([1]Prices!A1,"x"),"x")', t="str", v="x"),  # ... from another workbook
                      c("F1", GREEN, "B2*2")]),                                                   # a real own-sheet calc: fails
                 (2, [c("A2", BLUE, 'IF(B9>0,HYPERLINK("#A1","x"),"")', t="str", v="x"),          # condition reads B9: OWN blue
                      c("B2", BLUE, 'IFERROR(HYPERLINK("#A1","x"),C1*2)', t="str", v="x"),        # fallback calculates: OWN blue
                      c("C2", GREEN, 'IFERROR(HYPERLINK("#A1","x"),"y")&""', t="str", v="x")]),  # not top-level: OWN green
                 # a shared group of navigation cells (ChatGPT writes one per row): master + child ignored
                 (4, [c("A4", GREEN, 'IFERROR(HYPERLINK("#\'Summary\'!A1","OPEN"),"OPEN")', t="str", v="OPEN",
                        fattrs=' t="shared" ref="A4:A5" si="0"')]),
                 (5, [c("A5", GREEN, "", t="str", v="OPEN", fattrs=' t="shared" si="0"')]))
    p = book("hyperlinks_guarded.xlsx", [("Summary", data, None), ("Inputs", sheet(), None)], ext_links=1)
    v = grade(p, checks=[49, 50, 51])
    assert set(locs(v[K49])) == {"Summary!F1", "Summary!A2:B2", "Summary!C2"}, locs(v[K49])
    assert v[K50]["decision"] == "pass" and v[K51]["decision"] == "pass", (locs(v[K50]), locs(v[K51]))
    assert v[K49]["stats"]["hyperlink_cells_ignored"] == 7, v[K49]["stats"]
    assert v[K51]["stats"]["hyperlink_cells_ignored"] == 7, v[K51]["stats"]
    assert v[K49]["stats"]["offending_cells_by_reason"] == {"own": 4}, v[K49]["stats"]


def test_49_indirect_unknown_wording():
    """Finding 49-R2-3 (wording): a blue / red / grey cell whose only possible other-sheet reading
    is an INDIRECT address built from cell contents fails either way, so the address is not
    read and the mistake must not claim the formula uses only its own sheet.  v15: a CROSS cell
    (literal or direct reference to Inputs) is 50's; a BLACK =INDIRECT(A20) has its address read
    by 50 (A20 names Inputs -> a black cross-sheet formula -> fails 50)."""
    data = sheet((1, [c("A1", BLUE, "INDIRECT(A20)"),              # address 'Inputs!B5' never read: reason 'indirect'
                      c("B1", BLUE, 'INDIRECT("Inputs!B5")'),      # literal address: CROSS, blue -> 50 fails it
                      c("C1", GREEN, "INDIRECT(A20)"),             # green: the address cell is read -> reads Inputs, fine
                      c("D1", BLUE, "Inputs!B5"),                  # written pointer: 50's business
                      c("E1", RED, "INDIRECT(A20&B20)*2"),         # computed address, red: wrong either way
                      c("F1", BLACK, "INDIRECT(A20)"),             # black: 50 reads A20 -> reads Inputs -> must be green
                      c("G1", BLUE, "INDIRECT(A20)+Inputs!B1")]),  # reads Inputs directly: CROSS -> 50
                 (20, [cstr("A20", "Inputs!B5"), cstr("B20", "x")]))
    p = book("49_indirect_wording.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = run(p, 49)
    assert set(locs(v)) == {"Calc!A1", "Calc!E1"}, locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "INDIRECT address is built from cell contents" in d["Calc!A1"] and "wrong either way" in d["Calc!A1"], d
    assert "only its own sheet" not in d["Calc!A1"] and "only its own sheet" not in d["Calc!E1"], d
    st = v["stats"]
    assert st["offending_cells_by_reason"] == {"indirect": 2}, st
    assert st["indirect_address_cells_read"] == 1 and st["pointer_cells_left_to_check_50"] == 1, st
    assert st["cross_cells_left_to_check_50"] == 2, st                                     # B1, G1
    v = run(p, 50)
    assert set(locs(v)) == {"Calc!B1", "Calc!D1", "Calc!F1", "Calc!G1"}, locs(v)
    d = {m["location"]: m["description"] for m in v["mistakes"]}
    assert "calculation that references another sheet" in d["Calc!B1"] and "must be green" in d["Calc!F1"], d
    assert v["stats"]["indirect_address_cells_read"] == 1 and v["stats"]["own_sheet_cells_left_to_check_49"] == 2, \
        v["stats"]                                                                          # A1, E1 left to 49
    # the black INDIRECT whose address stays on its own sheet, or is held in a formula cell, is fine / skipped
    data = sheet((1, [c("A1", BLACK, "INDIRECT(A20)*2"), c("B1", BLACK, "INDIRECT(B20)*2"),
                      c("C1", BLACK, "INDIRECT(C20&D20)")]),
                 (20, [cstr("A20", "C5"), c("B20", BLACK, '"Inputs!B"&2', t="str", v="Inputs!B2"), cstr("C20", "In"),
                       cstr("D20", "puts!B5")]))
    p = book("50_indirect_black.xlsx", [("Calc", data, None), ("Inputs", sheet(), None)])
    v = graded("indirect_address_unknown", lambda: run(p, 50), "Calc!B20")
    assert v["decision"] == "pass" and v["stats"]["defaults"]["indirect_address_unknown"]["count"] == 2, v["stats"]


def test_near_empty_attempt_no_gate():
    """Ruling 2026-10-03 (A3): no completeness gate - a workbook without formulas passes 49, 50
    and 51 on the plain rule (the LLM's incomplete-attempt rule is not applied)."""
    p = book("near_empty.xlsx", [("Answers", sheet((1, [cstr("A1", "Answer:", BLUE)])), None)])
    v = grade(p, checks=[49, 50, 51])
    assert all(v[k]["decision"] == "pass" for k in (K49, K50, K51)), v


def test_unresolvable_colours_2026_10_04():
    """Finding 47-R2-08 in the core (2026-10-04): what the core used to guess now reaches 49 / 50 / 51 as an
    unresolvable colour (core styles.UnknownColour) - an invalid custom <indexedColors> entry (read as black
    before, so a formula in it passed 49 silently), an rgb of 7 digits (its last six were used), a tint that is
    not a number in [-1, 1] (ignored), indexed 81 - and raises only where the cell's class makes its colour
    decide, exactly as for a theme slot >= 12.  Valid custom palette entries resolve as before."""
    global STYLES
    fonts = ["", '<color indexed="14"/>', '<color rgb="0FF0000"/>', '<color theme="1" tint="2"/>',
             '<color indexed="81"/>', '<color indexed="13"/>']
    custom = ("<fonts count=\"%d\">%s</fonts>" % (len(fonts), "".join(f"<font><sz val=\"11\"/>{x}<name val=\"Calibri\"/></font>"
                                                                       for x in fonts))
              + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
              + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              + '<cellXfs count="%d">%s</cellXfs>' % (len(fonts), "".join(
                  f'<xf numFmtId="0" fontId="{i}" fillId="0" borderId="0" xfId="0" applyFont="1"/>' for i in range(len(fonts))))
              + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
              + '<colors><indexedColors>' + '<rgbColor rgb="FF000000"/>' * 13 + '<rgbColor rgb="FF0000FF"/>'
              + '<rgbColor rgb="not-a-colour"/></indexedColors></colors>')
    saved = STYLES
    STYLES = custom
    try:
        for s_, why in ((1, "custom <indexedColors> entry 14"), (2, "6 or 8 hex digits"), (3, 'tint="2"'),
                        (4, "64 and 65")):
            p = book(f"unres26_{s_}.xlsx", [("Calc", sheet((1, [c("A1", s_, "1+1")])), None)])
            v = run(p, 49)                                             # read as automatic (black): passes 49
            assert v["decision"] == "pass" and why in v["stats"]["defaults"]["unresolved_colour"]["examples"][0], v
            assert run(p, 50)["decision"] == "pass" and run(p, 51)["decision"] == "pass"
        p = book("unres26_ptr.xlsx", [("Calc", sheet((1, [c("A1", 1, "Inputs!B5")])), None), ("Inputs", sheet(), None)])
        v = run(p, 50)                                                 # a black pointer fails 50
        assert locs(v) == ["Calc!A1"] and "custom <indexedColors> entry 14" in \
            v["stats"]["defaults"]["unresolved_colour"]["examples"][0], v
        assert run(p, 49)["decision"] == "pass"                        # a pointer is not 49's business
        p = book("unres26_valid.xlsx", [("Calc", sheet((1, [c("A1", 5, "1+1")])), None)])
        v = run(p, 49)
        assert v["decision"] == "fail" and "0000FF" in v["mistakes"][0]["description"], v   # entry 13: blue
    finally:
        STYLES = saved


def test_conditional_format_font_colour_excuses():
    """Patrick 2026-10-06 (spot-check case 20, attempt 2348 Cover!C5 =Checks!F32, black, CF green when
    "OK" / red when "FAIL"): a conditional-format font colour in a family the cell's class accepts
    excuses a wrong base font, fires or not; it never makes a base-correct cell fail."""
    inputs = ("Inputs", sheet(), None)
    # 50: a black pointer painted green by one of two rules -> excused; only a red rule -> still fails;
    # a rule on another range excuses nothing; a fill-only dxf excuses nothing
    data = sheet((5, [c("C5", BLACK, "Inputs!F32"), c("D5", BLACK, "Inputs!F33"), c("E5", BLACK, "Inputs!F34"),
                      c("F5", BLACK, "Inputs!F35")]))
    p = book("cf_50.xlsx", [("Cover", data + cf("C5", DXF_GREEN, DXF_RED) + cf("D5", DXF_RED) + cf("A1:B9", DXF_GREEN)
                             + cf("F5", DXF_FILL), None), inputs])
    v = run(p, 50)
    assert set(locs(v)) == {"Cover!D5:F5"}, locs(v)
    assert v["stats"]["cf_excused"]["count"] == 1 and "Cover!C5" in v["stats"]["cf_excused"]["examples"][0], v["stats"]
    # 49: an own-sheet formula in blue painted black -> excused; painted green -> not (OWN needs black);
    # a CROSS calculation in blue painted green -> 50's since v15, excused there (green accepted)
    data = sheet((1, [c("A1", BLUE, "B1*2"), c("B1", BLUE, "C1*2"), c("C1", BLUE, "Inputs!B5*2")]))
    p = book("cf_49.xlsx", [("Calc", data + cf("A1", DXF_BLACK) + cf("B1", DXF_GREEN) + cf("C1", DXF_GREEN), None), inputs])
    v = run(p, 49)
    assert locs(v) == ["Calc!B1"], locs(v)
    assert v["stats"]["cf_excused"]["count"] == 1, v["stats"]
    v = run(p, 50)
    assert v["decision"] == "pass" and v["stats"]["cf_excused"]["count"] == 1, v["stats"]
    # one-directional: a base-correct green pointer painted red by a CF rule still passes 50
    p = book("cf_onedir.xlsx", [("Cover", sheet((1, [c("A1", GREEN, "Inputs!B5")])) + cf("A1", DXF_RED), None), inputs])
    v = run(p, 50)
    assert v["decision"] == "pass" and v["stats"]["cf_excused"]["count"] == 0, v
    # an unresolvable dxf font colour excuses nothing and is recorded
    p = book("cf_bad.xlsx", [("Cover", sheet((1, [c("A1", BLACK, "Inputs!B5")])) + cf("A1", DXF_BAD), None), inputs])
    v = run(p, 50)
    assert locs(v) == ["Cover!A1"] and v["stats"]["cf_unresolved_font_colours"]["count"] == 1, v["stats"]
    # 51: a black external link painted red -> excused
    p = book("cf_51.xlsx", [("Calc", sheet((1, [c("A1", BLACK, "[1]Prices!A1"), c("B1", BLACK, "[1]Prices!A2")]))
                             + cf("A1", DXF_RED), None)], ext_links=1)
    v = run(p, 51)
    assert locs(v) == ["Calc!B1"] and v["stats"]["cf_excused"]["count"] == 1, (locs(v), v["stats"])
    # the INDIRECT second pass goes through the same excuse: green =INDIRECT(A20) with A20 'C5' (own) painted black
    data = sheet((1, [c("A1", GREEN, "INDIRECT(A20)*2")]), (20, [cstr("A20", "C5")]))
    p = book("cf_pending.xlsx", [("Calc", data + cf("A1", DXF_BLACK), None), inputs])
    v = run(p, 49)
    assert v["decision"] == "pass" and v["stats"]["cf_excused"]["count"] == 1, v


TESTS = [test_colour_families, test_classifier_kinds, test_prefilters_are_conservative, test_49_rules,
         test_conditional_format_font_colour_excuses,
         test_49_arrays_shared_datatable_hidden, test_49_pass_and_contract, test_49_no_fallback, test_50_rules,
         test_50_shared_child_off_grid_and_pass, test_51_rules, test_51_default_pass, test_all_three_together,
         test_rubric_keys, test_structured_refs_with_commas_are_pointers, test_leading_equals_pointer,
         test_bare_table_name, test_unresolved_colour_raises_only_where_needed,
         test_unknown_scope_names_raise_only_where_needed, test_49_colour_wrong_in_every_class_needs_no_class,
         test_indirect_addresses, test_49_green_direct_reference_needs_no_parse,
         test_hyperlink_cells_any_colour, test_pointers_judged_only_by_50, test_light_text_on_dark_fill_no_exemption,
         test_hyperlink_guarded_navigation_cells, test_49_indirect_unknown_wording, test_near_empty_attempt_no_gate,
         test_unresolvable_colours_2026_10_04]


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
