"""Unit tests for Sufficient column widths (69) on synthetic micro-workbooks (raw SpreadsheetML zips
written here).

    cd judge && python -m detchecks.tests.test_checks_69

Plain asserts; also collectable by pytest.  Temporary files go to detchecks/scratch/colwidths/.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback
import zipfile
from xml.sax.saxutils import escape, quoteattr

from detchecks.api import Engine
from detchecks.checks import c69 as M
from detchecks.checks.c69 import C69, col_px, general_alternatives, glyph_row, mdw_px, text_px
from detchecks.errors import GradingError

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.path.join(os.path.dirname(HERE), "scratch", "colwidths")
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

_TMP = None
_N = [0]


def tmp(name: str = "") -> str:
    global _TMP
    if _TMP is None:
        os.makedirs(SCRATCH, exist_ok=True)
        _TMP = tempfile.mkdtemp(prefix="test_69_", dir=SCRATCH)
    _N[0] += 1
    return os.path.join(_TMP, f"{_N[0]:03d}_{name or 'book'}.xlsx")


# ============================================================================ builder
class Styles:
    """styles.xml builder; Normal = Calibri 11 unless given."""

    def __init__(self, normal=("Calibri", 11)):
        self.fonts = [f'<font><sz val="{normal[1]}"/><name val="{normal[0]}"/></font>']
        self.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
        self.numfmts = {}       # id -> code
        self.dxfs = []

    def fmt(self, code: str) -> int:
        for i, c_ in self.numfmts.items():
            if c_ == code:
                return i
        i = 164 + len(self.numfmts)
        self.numfmts[i] = code
        return i

    def xf(self, name="Calibri", size=11, bold=False, numfmt=0, shrink=False, wrap=False, horizontal=None,
           indent=0, rotation=0) -> int:
        self.fonts.append(f'<font>{"<b/>" if bold else ""}<sz val="{size}"/><name val="{name}"/></font>')
        fid = len(self.fonts) - 1
        nid = numfmt if isinstance(numfmt, int) else self.fmt(numfmt)
        al = []
        if shrink:
            al.append('shrinkToFit="1"')
        if wrap:
            al.append('wrapText="1"')
        if horizontal:
            al.append(f'horizontal="{horizontal}"')
        if indent:
            al.append(f'indent="{indent}"')
        if rotation:
            al.append(f'textRotation="{rotation}"')
        alx = f'<alignment {" ".join(al)}/>' if al else ""
        self.xfs.append(f'<xf numFmtId="{nid}" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" '
                        f'applyNumberFormat="1" applyAlignment="1">{alx}</xf>')
        return len(self.xfs) - 1

    def dxf(self, numfmt: str) -> int:
        self.dxfs.append(f'<dxf><numFmt numFmtId="{self.fmt(numfmt)}" formatCode={quoteattr(numfmt)}/></dxf>')
        return len(self.dxfs) - 1

    def xml(self) -> str:
        nf = "".join(f'<numFmt numFmtId="{i}" formatCode={quoteattr(c_)}/>' for i, c_ in self.numfmts.items())
        nfx = f'<numFmts count="{len(self.numfmts)}">{nf}</numFmts>' if nf else ""
        dx = f'<dxfs count="{len(self.dxfs)}">{"".join(self.dxfs)}</dxfs>' if self.dxfs else ""
        return (f'<styleSheet xmlns="{MAIN}">{nfx}<fonts count="{len(self.fonts)}">{"".join(self.fonts)}</fonts>'
                '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf fontId="0"/></cellStyleXfs>'
                f'<cellXfs count="{len(self.xfs)}">{"".join(self.xfs)}</cellXfs>'
                f'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>{dx}</styleSheet>')


def c(ref, v=None, s=0, f=None, t=None):
    """One <c>.  v: str (inline string), number, or formula cache (with f; None = no cache)."""
    sa = f' s="{s}"' if s else ""
    if f is not None:
        tt = t or ("str" if isinstance(v, str) else None)
        ta = f' t="{tt}"' if tt else ""
        vv = "" if v is None else f"<v>{escape(str(v))}</v>"
        return f'<c r="{ref}"{sa}{ta}><f>{escape(f)}</f>{vv}</c>'
    if v is None:
        return f'<c r="{ref}"{sa}/>'
    if t == "d":
        return f'<c r="{ref}"{sa} t="d"><v>{v}</v></c>'
    if isinstance(v, str):
        return f'<c r="{ref}"{sa} t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
    return f'<c r="{ref}"{sa}><v>{v!r}</v></c>'


def ws(rows=None, *, cols="", fmt='<sheetFormatPr defaultRowHeight="15"/>', tail=""):
    """rows: {r: (attrs, [cells])}."""
    body = []
    for r in sorted(rows or {}):
        attrs, cells = rows[r]
        body.append(f'<row r="{r}"{(" " + attrs) if attrs else ""}>{"".join(cells)}</row>')
    colx = f"<cols>{cols}</cols>" if cols else ""
    return f"{fmt}{colx}<sheetData>{''.join(body)}</sheetData>{tail}"


def col(lo, hi, width=None, hidden=False):
    w = f' width="{width}" customWidth="1"' if width is not None else ""
    h = ' hidden="1"' if hidden else ""
    return f'<col min="{lo}" max="{hi}"{w}{h}/>'


def book(sheets, styles=None, *, states=None, app=None, name=""):
    styles = styles or Styles()
    states = states or {}
    path = tmp(name)
    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    ov = [f'<Override PartName="/xl/workbook.xml" ContentType="{CT_XLSX}"/>']
    ov += [f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT_WS}"/>' for i in range(len(sheets))]
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>' + "".join(ov) + "</Types>")
    rr = f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
    if app:
        rr += f'<Relationship Id="rId2" Type="{REL}/extended-properties" Target="docProps/app.xml"/>'
        z.writestr("docProps/app.xml", '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                                       f'<Application>{escape(app)}</Application></Properties>')
    z.writestr("_rels/.rels", f'<Relationships xmlns="{PKG_REL}">{rr}</Relationships>')
    wrels, sh = [], []
    for i, (nm, xml) in enumerate(sheets):
        st = f' state="{states[i]}"' if i in states else ""
        sh.append(f'<sheet name={quoteattr(nm)} sheetId="{i + 1}"{st} r:id="rId{i + 1}"/>')
        wrels.append(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>')
        z.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'{DECL}<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{xml}</worksheet>')
    wrels.append(f'<Relationship Id="rIdS" Type="{REL}/styles" Target="styles.xml"/>')
    z.writestr("xl/styles.xml", styles.xml())
    z.writestr("xl/workbook.xml", f'{DECL}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>' + "".join(sh) + "</sheets></workbook>")
    z.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{PKG_REL}">' + "".join(wrels) + "</Relationships>")
    z.close()
    return path


def run(path):
    return Engine(path, [C69()]).run()[C69.key]


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
        msg = str(e) + " " + " ".join(str(x) for x in e.failures.values())
        for n in needles:
            assert n in msg, (n, msg[:600])
        return e
    raise AssertionError(f"expected GradingError containing {needles}")


class patched:
    def __init__(self, mod, **kw):
        self.mod, self.kw, self.old = mod, kw, {}

    def __enter__(self):
        for k, v in self.kw.items():
            self.old[k] = getattr(self.mod, k)
            setattr(self.mod, k, v)

    def __exit__(self, *a):
        for k, v in self.old.items():
            setattr(self.mod, k, v)


# Excel stores width = UI characters + 5/MDW (padding), e.g. UI 8.43 (default, 64 px) -> no <col>;
# UI 3.0 -> 3.7109375 (26 px); UI 12.0 -> 12.7109375 (89 px).
def stored(ui_chars: float, mdw: int = 7) -> float:
    return int((ui_chars * mdw + 5) / mdw * 256) / 256


# ============================================================================ pure functions
def test_69_metrics():
    assert mdw_px("calibri", 11, 96) == 7 and mdw_px("calibri", 11, 120) == 9
    assert mdw_px("arial", 10, 96) == 7 and mdw_px("arial", 10, 120) == 9
    assert mdw_px("aptos narrow", 11, 96) == 7
    assert col_px(8.7109375, 7) == 61          # ECMA-376 example
    assert col_px(stored(12.0), 7) == 89 and col_px(stored(3.0), 7) == 26 and col_px(stored(2.0), 7) == 19
    assert col_px(0, 7) == 0
    cal = glyph_row("calibri", False)
    assert text_px(cal, "2029", 96 * 11 / 72, True) == 28.0       # 4 digits x 7 px
    assert abs(text_px(cal, "2029", 96 * 11 / 72, False) - 4 * 0.507 * 96 * 11 / 72) < 1e-6
    assert text_px(glyph_row("arial", True), "0", 96 * 10 / 72, True) == 7.0
    assert text_px(cal, " ", 15, True) == text_px(cal, " ", 15, True)
    # General falls back to a 0-decimal rounding and 1-digit scientific notation
    assert general_alternatives(123456789.0, "123456789") == ["123456789", "123456789", "1E+08"]
    assert general_alternatives(0.75, "0.75")[1] == "1"
    assert general_alternatives(-2029.0, "-2029") == ["-2029", "-2029", "-2E+03"]
    assert general_alternatives(0.0, "0") == ["0", "0"]
    # unknown faces are measured as Calibri (narrowest common face)
    assert M.face_key("Segoe UI") == ("calibri", False)
    assert M.face_key("Helvetica Neue") == ("arial", True)
    assert M.face_key("Aptos Narrow") == ("aptos narrow", True)


# ============================================================================ toy traps rebuilt
def test_69_narrow_number_date_percent():
    """T1: a year and formatted numbers in a 3.0-character column; T3: dates in a 5.0 column; a percent."""
    st = Styles(("Aptos Narrow", 11))
    s_num = st.xf("Arial", 10, numfmt="#,##0.00;\\(#,##0.00\\)")
    s_date = st.xf("Arial", 10, numfmt="mm/dd/yyyy")
    s_pct = st.xf("Arial", 10, numfmt="0.0%")
    s_gen = st.xf("Arial", 10)
    rows = {2: ("", [c("J2", 2029, s_gen)]),
            7: ("", [c("J7", 22852.0, s_num), c("K7", 45658.0, s_date), c("L7", 0.125, s_pct)]),
            8: ("", [c("J8", 1234.5, s_num)])}
    narrow = col(10, 12, stored(3.0))
    v = run(book([("Solution Model", ws(rows, cols=narrow))], st))
    assert v["decision"] == "fail", v["summary"]
    # rectangles: horizontal runs first (J7:L7), then J8 (group_cells)
    assert locs(v) == ["'Solution Model'!J2", "'Solution Model'!J7:L7", "'Solution Model'!J8"], locs(v)
    assert "'####'" in v["mistakes"][0]["description"] and "2029" in v["mistakes"][0]["description"]
    assert "column J:L" in v["mistakes"][1]["description"]
    assert v["stats"]["sure_overflows"] == 5 and v["stats"]["column_unit_px"] == {"w96r": 7, "w96f": 7, "w120r": 9, "w120f": 9}
    # the same cells in a 12.0-character column fit (T1 Pass: 12.0 wide, widest needs 8.11)
    v = run(book([("Solution Model", ws(rows, cols=col(10, 12, stored(12.0))))], st))
    assert v["decision"] == "pass", v["mistakes"]
    assert v["stats"]["numeric_cells"] == 5 and v["stats"]["band_cells"] == 0


def test_69_one_digit_overflow_and_general():
    """Judge README calibration: 123456789 (format 0) in an untouched Calibri 11 column fails (one digit,
    >= 4 px under every model); 12345678 fits.  Under General the same number shortens to 1E+08."""
    st = Styles()
    s0 = st.xf(numfmt=1)                  # built-in '0'
    v = run(book([("S", ws({1: ("", [c("A1", 123456789, s0)]), 2: ("", [c("A2", 12345678, s0)])}))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    assert "123456789" in v["mistakes"][0]["description"]
    v = run(book([("S", ws({1: ("", [c("A1", 123456789)])}))], st))
    assert v["decision"] == "pass" and v["stats"]["general_cells"] == 1
    with patched(M, GENERAL_SHORTENS=False):
        v = run(book([("S", ws({1: ("", [c("A1", 123456789)])}))], st))
        assert v["decision"] == "fail"
    # General 0.75 in a 1.5-character column rounds to '1' instead of ####
    v = run(book([("S", ws({1: ("", [c("A1", 0.75)])}, cols=col(1, 1, stored(1.5))))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # thousands separator: '12,345,678' (#,##0) overflows the default column
    v = run(book([("S", ws({1: ("", [c("A1", 12345678, st.xf(numfmt=3))])}))], st))
    assert v["decision"] == "fail"


def test_69_shrink_merge_hidden():
    st = Styles()
    s_num = st.xf(numfmt="#,##0.00")
    s_shr = st.xf(numfmt="#,##0.00", shrink=True)
    narrow = col(1, 2, stored(3.0))
    # shrink-to-fit never shows ####
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.89, s_shr)])}, cols=narrow))], st))
    assert v["decision"] == "pass" and v["stats"]["shrink_to_fit"] == 1
    # merged A1:B1 (2 x 26 px = 52, text area 47): '1,234.50' (8 x 7 = 56) still overflows; '123.50' fits the
    # span but not column A alone
    merge = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
    v = run(book([("S", ws({1: ("", [c("A1", 123.5, s_num), c("B1", 999999.0, s_num)])}, cols=narrow, tail=merge))], st))
    assert v["decision"] == "pass", v["mistakes"]          # anchor fits the span; covered B1 is not displayed
    v = run(book([("S", ws({1: ("", [c("A1", 1234.5, s_num)])}, cols=narrow, tail=merge))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"]
    # hidden column, zero-width column, hidden row, zero-height row: not displayed
    for cols_, rows_ in ((col(1, 1, stored(3.0), hidden=True), {1: ("", [c("A1", 1234567.0, s_num)])}),
                         (col(1, 1, 0), {1: ("", [c("A1", 1234567.0, s_num)])}),
                         (narrow, {1: ('hidden="1"', [c("A1", 1234567.0, s_num)])}),
                         (narrow, {1: ('ht="0" customHeight="1"', [c("A1", 1234567.0, s_num)])})):
        v = run(book([("S", ws(rows_, cols=cols_))], st))
        assert v["decision"] == "pass", (cols_, rows_, v["mistakes"])
    assert v["stats"]["hidden_row_cells"] == 1
    # a hidden sheet is graded (whole workbook)
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.0, s_num)])}, cols=narrow))], st, states={0: "hidden"}))
    assert v["decision"] == "fail" and "(hidden sheet)" in v["mistakes"][0]["description"]


# ============================================================================ cut-off text (Patrick 2026-10-04)
# Calibri 11 (the default Normal font here): a digit '0' is 7 px at 96 dpi rounded (7.44 fractional) and 9 px at
# 120 dpi rounded (9.30 fractional).  Column text areas (px - 5): UI 5.0 -> 35 / 46, UI 10.0 -> 70 / 91,
# default (no <col>) -> 59 / 75.
TEN = "0" * 10              # 70 / 74.4 / 90 / 93.0 px


def rich(ref, runs, s=0):
    """Inline rich string: runs = [(text, size or None)]."""
    rr = "".join((f'<r><rPr><sz val="{sz}"/></rPr>' if sz else "<r>") + f'<t xml:space="preserve">{escape(t)}</t></r>'
                 for t, sz in runs)
    sa = f' s="{s}"' if s else ""
    return f'<c r="{ref}"{sa} t="inlineStr"><is>{rr}</is></c>'


def merges(*refs):
    return f'<mergeCells count="{len(refs)}">' + "".join(f'<mergeCell ref="{r}"/>' for r in refs) + "</mergeCells>"


def test_69_text_cut_off_by_neighbours():
    """Text that cannot be fully seen is cut off; text running into empty neighbours and staying visible is not
    (Patrick 2026-10-04).  Any length (Patrick 2026-10-28, judge v15): toy T4's B19 trap - a label cut off next
    to a filled C19 - and the two narrow rules of the first build (a number stored as text cut off; any text in a
    <= 12 px column blocked) are cases of the rule.  (Under v13/v14 only 200+ characters failed.)"""
    st = Styles(("Aptos Narrow", 11))
    s_txt = st.xf("Arial", 10)
    s_bold = st.xf("Arial", 10, bold=True)
    s_num = st.xf("Arial", 10, numfmt="#,##0")
    colb = col(2, 2, stored(17.0)) + col(3, 3, stored(34.0))
    # B51 overflows into the empty C51 and stays fully visible -> pass
    v = run(book([("S", ws({51: ("", [c("B51", "Real Estate Valuation (Income Approach)", s_txt)])}, cols=colb))], st))
    assert v["decision"] == "pass" and v["stats"]["text_fits_with_spill"] == 1, v["mistakes"]
    # T4's B19: 'Net Operating Income' (Arial 10 bold, 136 px at 96 dpi) in the 17.0 column B (119 px of text area)
    # next to a filled C19 -> cut off
    p19 = book([("Solution Model", ws({19: ("", [c("B19", "Net Operating Income", s_bold), c("C19", "Year 1", s_txt)])},
                                      cols=colb))], st)
    v = run(p19)
    assert v["decision"] == "fail" and locs(v) == ["'Solution Model'!B19"], v["mistakes"]
    d = v["mistakes"][0]["description"]
    assert ("Text in B19 on sheet 'Solution Model' is cut off: B19 'Net Operating Income' (text, Arial 10 bold, "
            "left-aligned) is cut off by C19 (text): about 2 of its 20 characters (17 px at 96 dpi)") in d, d
    assert v["stats"]["mistakes_by_rule"] == {"numbers": 0, "unwrapped_text": 1, "wrapped_text": 0}
    assert v["stats"]["text_blockers"] == {"text": 1} and v["stats"]["text_cut_off"] == 1
    with patched(M, TEXT_CLIP=False):
        v = run(p19)
        assert v["decision"] == "pass" and v["stats"]["text_cut_off"] == 1, v["mistakes"]
    # a label cut off by a number; the same label running across an empty styled cell and an absent one passes
    lab = "Revenue growth assumption"
    v = run(book([("S", ws({1: ("", [c("A1", lab, s_txt), c("B1", 0.05, s_num)])}))], st))
    assert v["decision"] == "fail" and "A1 'Revenue growth assumption'" in v["mistakes"][0]["description"]
    assert "cut off by B1 (a number)" in v["mistakes"][0]["description"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", lab, s_txt), c("B1", None, s_txt), c("D1", 0.05, s_num)])}))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # a formula returning "" is not empty for overflow: it cuts the label off
    v = run(book([("S", ws({1: ("", [c("A1", lab, s_txt), c("B1", "", s_txt, f='IF(1,"","")')])}))], st,
                 app="Microsoft Excel"))
    assert v["decision"] == "fail" and "cut off by B1 (a formula result)" in v["mistakes"][0]["description"], v
    # a number stored as text (the old TEXT_NUMBER_CLIP case): cut off by a filled B1; with room it passes
    narrow = col(1, 1, stored(5.0))
    v = run(book([("S", ws({1: ("", [c("A1", "1234567890123", s_txt), c("B1", "x", s_txt)])}, cols=narrow))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", "1234567890123", s_txt), c("B1", None, s_txt), c("D1", "x", s_txt)])},
                           cols=narrow))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # text in a 12 px column (the old TEXT_HIDDEN_COL_MAX_PX case), any length: blocked fails, room passes
    tiny = col(1, 1, stored(1.0))
    for text, b, want in (("Note", "B1", "fail"), ("Note", "C1", "pass"), ("ab", "B1", "fail")):
        v = run(book([("S", ws({1: ("", [c("A1", text, s_txt), c(b, 5.0, s_num)])}, cols=tiny))], st))
        assert v["decision"] == want, (text, b, v["mistakes"])
    # no length bar (2026-10-28): T4's B19 (20 characters) fails like any other cut-off text, and the stats carry
    # no "short" counters any more
    v = run(p19)
    assert v["decision"] == "fail" and v["stats"]["text_cut_off"] == 1 and v["stats"]["text_blockers"] == {"text": 1}
    assert v["stats"]["per_sheet"][0]["text_cut_off"] == 1 and v["stats"]["mistakes_by_rule"]["unwrapped_text"] == 1
    assert not any(k.endswith("_short") or "short_examples" in k for k in v["stats"]), sorted(v["stats"])
    assert v["stats"]["options"]["text_clip_min_chars"] == 0
    assert "text cut off on: 'Solution Model'" in v["summary"], v["summary"]


def test_69_text_alignment_directions():
    """General / left text overflows to the right, right-aligned to the left, centred to both sides (each side must
    hold half the excess); centre-across-selection is centred over its selection first; fill never spills;
    rotated text is skipped and shrink-to-fit text is never cut off; an indent takes room."""
    st = Styles()
    s_r = st.xf(horizontal="right")
    s_c = st.xf(horizontal="center")
    s_cc = st.xf(horizontal="centerContinuous")
    s_fill = st.xf(horizontal="fill")
    s_left = st.xf(horizontal="left")
    s_ind = st.xf(horizontal="left", indent=2)
    narrow = col(3, 3, stored(5.0))                         # C: 35 / 46 px of text area
    for cells, want, why in (([c("C1", TEN, s_r)], "pass", "right-aligned, B1 empty"),
                             ([c("B1", 1.0), c("C1", TEN, s_r)], "fail", "right-aligned, B1 filled"),
                             ([c("C1", TEN, s_r), c("D1", 1.0)], "pass", "right-aligned, only D1 filled"),
                             ([c("C1", TEN, s_c)], "pass", "centred, both sides empty"),
                             ([c("B1", 1.0), c("C1", TEN, s_c)], "fail", "centred, B1 filled"),
                             ([c("C1", TEN, s_c), c("D1", 1.0)], "fail", "centred, D1 filled"),
                             ([c("C1", TEN), c("B1", 1.0)], "pass", "general, only B1 filled")):
        cells = sorted(cells, key=lambda x: x.split('"')[1])
        v = run(book([("S", ws({1: ("", cells)}, cols=narrow))], st))
        assert v["decision"] == want, (why, v["mistakes"])
    v = run(book([("S", ws({1: ("", [c("B1", 1.0), c("C1", TEN, s_r)])}, cols=narrow))], st))
    assert "(text, Calibri 11, right-aligned) is cut off by B1 (a number)" in v["mistakes"][0]["description"], v
    # right-aligned text in column A / general text in column XFD: the sheet's edges stop them
    v = run(book([("S", ws({1: ("", [c("A1", TEN, s_r)])}, cols=col(1, 1, stored(5.0))))], st))
    assert v["decision"] == "fail" and "the left edge of the sheet" in v["mistakes"][0]["description"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("XFD1", TEN)])}, cols=col(16384, 16384, stored(5.0))))], st))
    assert v["decision"] == "fail" and "the right edge of the sheet" in v["mistakes"][0]["description"], v["mistakes"]
    # centre across selection: B1 (30 digits, 210 px) centred over B1:D1 (187 px) overflows 11.5 px on each side
    # into A1 and E1; a filled E1 cuts it off; without the selection (C1:D1 plain) it is centred on B1 alone
    thirty = "0" * 30
    sel = [c("B1", thirty, s_cc), c("C1", None, s_cc), c("D1", None, s_cc)]
    v = run(book([("S", ws({1: ("", sel)}))], st))
    assert v["decision"] == "pass" and v["stats"]["text_center_across"] == 1, v["mistakes"]
    v = run(book([("S", ws({1: ("", sel + [c("E1", 1.0)])}))], st))
    assert v["decision"] == "fail" and "cut off by E1 (a number)" in v["mistakes"][0]["description"], v["mistakes"]
    assert "centred across selection" in v["mistakes"][0]["description"]
    v = run(book([("S", ws({1: ("", [c("B1", thirty, s_cc), c("C1", None), c("D1", None)])}))], st))
    assert v["decision"] == "fail" and "cut off by the left edge of the sheet" in v["mistakes"][0]["description"], v
    # fill alignment never spills: 10 digits in a 5.0 column fail although B1 is empty; 4 fit; an empty
    # fill-aligned B1 extends the cell
    v = run(book([("S", ws({1: ("", [c("A1", TEN, s_fill)])}, cols=col(1, 1, stored(5.0))))], st))
    assert v["decision"] == "fail" and "its fill alignment" in v["mistakes"][0]["description"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", "0000", s_fill)])}, cols=col(1, 1, stored(5.0))))], st))
    assert v["decision"] == "pass", v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", TEN, s_fill), c("B1", None, s_fill)])}, cols=col(1, 1, stored(5.0))))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # rotated text is skipped (counted); shrink-to-fit text is never cut off
    for s_, stat in ((st.xf(rotation=90), "text_rotated_skipped"), (st.xf(shrink=True), "text_shrink")):
        v = run(book([("S", ws({1: ("", [c("A1", TEN, s_), c("B1", 1.0)])}, cols=col(1, 1, stored(5.0))))], st))
        assert v["decision"] == "pass" and v["stats"][stat] == 1, (stat, v["mistakes"])
    # an indent takes room: 9 digits (63 px) fit a 10.0 column (70 px) next to a filled B1; with indent 2
    # (2 x 3 Normal spaces = 18 px at 96 dpi, 24 at 120 dpi) they do not
    nine = "0" * 9
    v = run(book([("S", ws({1: ("", [c("A1", nine, s_left), c("B1", 1.0)])}, cols=col(1, 1, stored(10.0))))], st))
    assert v["decision"] == "pass", v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", nine, s_ind), c("B1", 1.0)])}, cols=col(1, 1, stored(10.0))))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]


def test_69_text_merges_hidden_columns_and_edges():
    st = Styles()
    narrow = col(1, 3, stored(5.0))                         # A:C 40 / 51 px each
    # a merged anchor has the merge's width and never spills: 10 digits fit A1:B1 (75 px of text area), 15 do
    # not although C1 is empty
    v = run(book([("S", ws({1: ("", [c("A1", TEN)])}, cols=narrow, tail=merges("A1:B1")))], st))
    assert v["decision"] == "pass", v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", "0" * 15)])}, cols=narrow, tail=merges("A1:B1")))], st))
    assert v["decision"] == "fail" and "the edge of its merged range A1:B1" in v["mistakes"][0]["description"], v
    # a covered cell is not displayed
    v = run(book([("S", ws({1: ("", [c("A1", "x"), c("B1", "0" * 15)])}, cols=narrow, tail=merges("A1:B1")))], st))
    assert v["decision"] == "pass" and v["stats"]["text_covered_merged"] == 1, v["mistakes"]
    # a merged range takes no overflow (MERGES_BLOCK_OVERFLOW): A1 is cut off by the empty merge B1:C1
    p = book([("S", ws({1: ("", [c("A1", TEN)])}, cols=narrow, tail=merges("B1:C1")))], st)
    v = run(p)
    assert v["decision"] == "fail" and "the merged range B1:C1" in v["mistakes"][0]["description"], v["mistakes"]
    with patched(M, MERGES_BLOCK_OVERFLOW=False):
        assert run(p)["decision"] == "pass"
    # hidden column B in the path: empty, it lets the text through but gives no room (C, D give it); filled, it
    # still stops it (HIDDEN_CELLS_BLOCK_OVERFLOW)
    hid = col(1, 1, stored(5.0)) + col(2, 2, stored(5.0), hidden=True) + col(3, 3, stored(5.0))
    v = run(book([("S", ws({1: ("", [c("A1", TEN)])}, cols=hid))], st))
    assert v["decision"] == "pass", v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", TEN), c("C1", 1.0)])}, cols=hid))], st))
    assert v["decision"] == "fail" and "cut off by C1 (a number)" in v["mistakes"][0]["description"], v["mistakes"]
    p = book([("S", ws({1: ("", [c("A1", TEN), c("B1", 1.0)])}, cols=hid))], st)
    v = run(p)
    assert v["decision"] == "fail" and "B1 (a number, in hidden column B)" in v["mistakes"][0]["description"], v
    assert v["stats"]["text_blockers"] == {"number (hidden column)": 1}
    with patched(M, HIDDEN_CELLS_BLOCK_OVERFLOW=False):
        assert run(p)["decision"] == "pass"
    # a merge spanning a hidden column has only its visible width: A1:C1 with B hidden is 75 px of text area,
    # too narrow for 12 digits (84 px); with B visible (115 px) they fit
    v = run(book([("S", ws({1: ("", [c("A1", "0" * 12)])}, cols=hid, tail=merges("A1:C1")))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", "0" * 12)])}, cols=narrow, tail=merges("A1:C1")))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # 69-F4 for text: a merge anchor in a hidden column shows its text across the visible part of the merge
    hcols = col(1, 1, hidden=True) + col(2, 2, stored(3.0))
    v = run(book([("S", ws({1: ("", [c("A1", "Long label text")])}, cols=hcols, tail=merges("A1:B1")))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and v["stats"]["text_hidden_anchor_cells"] == 1, v
    v = run(book([("S", ws({1: ("", [c("A1", "Long label text")])}, cols=col(1, 1, hidden=True) + col(2, 2, stored(40.0)),
                           tail=merges("A1:B1")))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # a text in a hidden column (not a merge anchor) or a hidden row is not displayed
    v = run(book([("S", ws({1: ("", [c("A1", TEN), c("B1", 1.0)]), 2: ('hidden="1"', [c("C2", TEN), c("D2", 1.0)])},
                           cols=col(1, 1, hidden=True) + col(3, 3, stored(5.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["text_hidden_col_cells"] == 1, v["mistakes"]
    # hidden sheets are graded
    v = run(book([("S", ws({1: ("", [c("A1", TEN), c("B1", 1.0)])}, cols=narrow))], st, states={0: "hidden"}))
    assert v["decision"] == "fail" and "(hidden sheet)" in v["mistakes"][0]["description"]


def test_69_wrapped_text_rows():
    """(b) wrapped text is cut off when it needs more lines than its custom-height row shows (more than half a
    line hidden); rows without customHeight are auto-fitted by Excel and never cut it (excel_measurements 6)."""
    st = Styles()
    s_wrap = st.xf(wrap=True)
    s_just = st.xf(horizontal="justify")
    w8 = "0" * 8                                            # 56 / 59.5 / 72 / 74.4 px: one per line in a 10.0 column
    two, four = f"{w8} {w8}", " ".join([w8] * 4)
    cols10 = col(1, 2, stored(10.0))
    custom = 'ht="15" customHeight="1"'
    p = book([("S", ws({1: (custom, [c("A1", four, s_wrap)])}, cols=cols10))], st)
    v = run(p)
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    d = v["mistakes"][0]["description"]
    assert ("Wrapped text in A1 on sheet 'S' is cut off at the bottom: A1 '00000000 00000000 00000000 000…' "
            "(wrapped text, Calibri 11) needs 4 lines in its 10.00-character column A, but row 1 has a custom height of "
            "15 pt: room for about 1.0 lines of 14.3 pt, so about 3.0 line(s) cannot be seen") in d, d
    assert v["stats"]["mistakes_by_rule"]["wrapped_text"] == 1 and v["stats"]["wrapped_custom_rows"] == 1
    with patched(M, WRAPPED_TEXT_CLIP=False):
        v = run(p)
        assert v["decision"] == "pass" and v["stats"]["wrapped_cut_off"] == 1, v["mistakes"]
    # the same height without customHeight, or no height at all: auto-fitted -> pass
    for attrs in ('ht="15"', ""):
        v = run(book([("S", ws({1: (attrs, [c("A1", four, s_wrap)])}, cols=cols10))], st))
        assert v["decision"] == "pass" and v["stats"]["wrapped_auto_rows"] == 1, (attrs, v["mistakes"])
    # two lines (28.6 pt): a 22 pt row shows 1.54 lines (0.46 hidden <= 0.5: pass); 20 pt shows 1.40 (fail)
    for ht, want in ((22, "pass"), (20, "fail")):
        v = run(book([("S", ws({1: (f'ht="{ht}" customHeight="1"', [c("A1", two, s_wrap)])}, cols=cols10))], st))
        assert v["decision"] == want, (ht, v["mistakes"])
    # horizontal justify wraps like wrapText; a trailing line break shows nothing
    v = run(book([("S", ws({1: (custom, [c("A1", four, s_just)])}, cols=cols10))], st))
    assert v["decision"] == "fail" and v["stats"]["wrapped_cut_off"] == 1, v["mistakes"]
    v = run(book([("S", ws({1: (custom, [c("A1", w8 + "\n", s_wrap)])}, cols=cols10))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # shrink-to-fit with wrap: never cut off
    v = run(book([("S", ws({1: (custom, [c("A1", four, st.xf(wrap=True, shrink=True))])}, cols=cols10))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # a merge over two custom rows (30 pt, 2.1 lines) still hides about 1.9 of 4 lines; with an auto-fitted row
    # in the block it is not graded (counted); a wider merge needs fewer lines
    rows2 = {1: (custom, [c("A1", four, s_wrap)]), 2: (custom, [c("A2", None, s_wrap)])}
    v = run(book([("S", ws(rows2, cols=cols10, tail=merges("A1:A2")))], st))
    assert v["decision"] == "fail" and "rows 1:2 of the merge have custom heights totalling 30 pt" in \
        v["mistakes"][0]["description"], v["mistakes"]
    rows2a = {1: (custom, [c("A1", four, s_wrap)]), 2: ("", [c("A2", None, s_wrap)])}
    v = run(book([("S", ws(rows2a, cols=cols10, tail=merges("A1:A2")))], st))
    assert v["decision"] == "pass" and v["stats"]["wrapped_merge_auto_rows"] == 1, v["mistakes"]
    v = run(book([("S", ws({1: (custom, [c("A1", two, s_wrap)])}, cols=cols10, tail=merges("A1:B1")))], st))
    assert v["decision"] == "pass", v["mistakes"]       # 150 px wide: both words on one line
    # Excel's AutoFit ignores merged cells: a wrapped merge anchor needing several lines in an auto-fitted row is
    # counted (not graded)
    v = run(book([("S", ws({1: ("", [c("A1", " ".join([w8] * 8), s_wrap)])}, cols=cols10, tail=merges("A1:B1")))], st))
    assert v["decision"] == "pass" and v["stats"]["wrapped_merged_in_auto_rows"] == 1, v["stats"]
    # a wrapped text formula result (Excel cache) is graded too
    v = run(book([("S", ws({1: (custom, [c("A1", four, s_wrap, f='REPT("0",8)&" "&REPT("0",8)')])}, cols=cols10))],
                 st, app="Microsoft Excel"))
    assert v["decision"] == "fail" and "a wrapped text formula result" in v["mistakes"][0]["description"], v
    # no length bar (2026-10-28): the 35-character text hidden in the 15 pt row fails on its own
    v = run(p)
    assert v["decision"] == "fail" and v["stats"]["wrapped_cut_off"] == 1 and "wrapped_cut_off_short" not in v["stats"]


def test_69_text_values_formats_and_fonts():
    st = Styles()
    cols8 = col(1, 1, stored(8.0))                          # A: 56 / 73 px of text area
    label = "far too long a label"
    # a trusted (Excel) text formula result is graded like a constant
    v = run(book([("S", ws({1: ("", [c("A1", label, f='"far too long a label"'), c("B1", 5.0)])}, cols=cols8))], st,
                 app="Microsoft Excel"))
    assert v["decision"] == "fail" and "(a text formula result, Calibri 11" in v["mistakes"][0]["description"], v
    assert v["stats"]["text_formula_results"] == 1
    # openpyxl: an untrusted formula result is undecided and raises when nothing else fails (no fallback) - unless its
    # format shows nothing for numbers AND text (';;;'); under ';;' a text result would still show (69-F2 narrowed)
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, f="C1&C1"), c("C1", "x")])}))], st, app=APP_OPX)),
           "cannot decide S!A1", "untrusted")
    v = run(book([("S", ws({1: ("", [c("A1", None, st.xf(numfmt=";;;"), f="C1&C1"), c("C1", "x")])}))], st, app=APP_OPX))
    assert v["decision"] == "pass" and v["stats"]["undecided_not_displayed"] == 1, v
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, st.xf(numfmt=";;"), f="C1&C1"), c("C1", "x")])}))], st,
                            app=APP_OPX)), "cannot decide S!A1")
    # ... an untrusted result in a centre-across cell is needed for the text rule (the number rule skips the cell)
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, st.xf(horizontal="centerContinuous"), f="C1&C1"),
                                                 c("C1", "x")])}))], st, app=APP_OPX)), "cannot decide S!A1")
    # with a value copy holding a long text behind the placeholder: cut off -> fail
    x = ws({1: ("", [c("A1", 0, f='IF(C1>0,"far too long a label","")'), c("B1", 5.0), c("C1", 1)])}, cols=cols8)
    copy = book([("S", ws({1: ("", [c("A1", label, f='IF(C1>0,"far too long a label","")'), c("B1", 5.0), c("C1", 1)])},
                          cols=cols8))], st)
    v = run_with(book([("S", x)], st, app=APP_OPX), value_path=copy)
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    # number formats: a text section can hide text (';;;') or lengthen it; 'Revenue' (52 px) fits A on its own
    s_hide = st.xf(numfmt=";;;")
    s_suffix = st.xf(numfmt='0;-0;0;@" (in millions)"')
    for s_, want in ((0, "pass"), (s_hide, "pass"), (s_suffix, "fail")):
        v = run(book([("S", ws({1: ("", [c("A1", "Revenue", s_), c("B1", 5.0)])}, cols=cols8))], st))
        assert v["decision"] == want, (s_, v["mistakes"])
    assert "'Revenue (in millions)'" in v["mistakes"][0]["description"]
    # blanks at the ends: trailing blanks are invisible, leading blanks push the text right
    v = run(book([("S", ws({1: ("", [c("A1", "Revenue" + " " * 30), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "pass", v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", " " * 6 + "Revenue"), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "fail", v["mistakes"]
    # an unwrapped cell shows typed line breaks as nothing: 'Rev\n\n\nenue' is 'Revenue' on one line (fits; three
    # breaks measured as glyphs would add about 21 px)
    v = run(book([("S", ws({1: ("", [c("A1", "Rev\n\n\nenue"), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # rich text: a 22 pt run doubles the width
    v = run(book([("S", ws({1: ("", [rich("A1", [("Revenue", 22)]), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "fail", v["mistakes"]
    v = run(book([("S", ws({1: ("", [rich("A1", [("Revenue", None)]), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # an unknown face is measured as Calibri and listed
    v = run(book([("S", ws({1: ("", [c("A1", "Revenue", st.xf("Segoe UI", 11)), c("B1", 5.0)])}, cols=cols8))], st))
    assert v["decision"] == "pass" and v["stats"]["text_unknown_faces"] == {"Segoe UI": 1}, v


def test_69_text_mistakes_band_and_stats():
    # adjacent cut-off labels form one mistake, the worst one quoted
    st = Styles(("Aptos Narrow", 11))
    s_txt = st.xf("Arial", 10)
    rows = {r: ("", [c(f"A{r}", ("Revenue growth assumption " + "x" * (r - 5)).strip(), s_txt), c(f"B{r}", 0.05)])
            for r in (5, 6, 7)}
    rows[9] = ("", [c("A9", "Cost of goods sold, net of rebates", s_txt), c("B9", 1.0)])
    v = run(book([("Model", ws(rows))], st))
    assert locs(v) == ["Model!A5:A7", "Model!A9"], locs(v)
    assert v["mistakes"][0]["description"].startswith("3 text cells A5:A7 on sheet 'Model' are cut off; e.g. A7 "), v
    assert v["stats"]["text_cut_off"] == 4 and v["stats"]["per_sheet"][0]["text_cut_off"] == 4
    # a band case: '709,557.13' as text (Arial 10) in a 9.0 column next to a filled B1 overflows under the
    # 96 dpi models only -> pass, counted
    v = run(book([("S", ws({1: ("", [c("A1", "709,557.13", s_txt), c("B1", 1.0)])}, cols=col(1, 1, stored(9.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["text_band"] == 1, v["stats"]["text_band_examples"]
    # a one-line text in a custom-height row lower than half a line is counted only
    v = run(book([("S", ws({1: ('ht="3" customHeight="1"', [c("A1", "Note", s_txt)])}))], st))
    assert v["decision"] == "pass" and v["stats"]["text_rows_under_half_line"] == 1, v["stats"]


def long_text(n: int) -> str:
    """A text of exactly n characters with blanks inside it but none at either end."""
    t = ("Assumption note " * (n // 16 + 2))[:n]
    return t[:-1] + "."


def test_69_text_clip_any_length():
    """Patrick 2026-10-28 (judge v15): cut-off text fails whatever its length - the 200-character bar of
    2026-10-04 ("option 3") is gone, for (a) unwrapped and (b) wrapped text alike.  shown_chars still words
    the mistake ("about 192 of its 200 characters"): the displayed text without leading / trailing blanks
    and typed line breaks."""
    assert not hasattr(M, "TEXT_CLIP_MIN_CHARS")
    assert M.shown_chars("  " + long_text(199) + " \n") == 199 and M.shown_chars(long_text(100) + "\n" + long_text(99)) == 199
    st = Styles()
    s_wrap = st.xf(wrap=True)
    s_suffix = st.xf(numfmt='0;-0;0;@" (bn)"')               # a text section adding 5 characters
    # (a) unwrapped, cut off by B1 (a number) in a default column: 12, 199 and 200 characters all fail
    for n in (12, 199, 200):
        v = run(book([("S", ws({1: ("", [c("A1", long_text(n)), c("B1", 5.0)])}))], st))
        assert v["decision"] == "fail" and v["stats"]["text_cut_off"] == 1, (n, v["stats"])
    d = v["mistakes"][0]["description"]
    assert "is cut off by B1 (a number): about 192 of its 200 characters" in d, d
    # blanks at the ends and typed line breaks do not count as shown; a number format's added text does
    for text, s_, shown in (("   " + long_text(199) + "   ", 0, 199),
                            (long_text(100) + "\n" + long_text(99), 0, 199),
                            (long_text(195), s_suffix, 200)):           # 195 + ' (bn)'
        v = run(book([("S", ws({1: ("", [c("A1", text, s_), c("B1", 5.0)])}))], st))
        assert v["decision"] == "fail" and f"of its {shown} characters" in v["mistakes"][0]["description"], \
            (repr(text[-12:]), s_, v["mistakes"])
    # (b) wrapped, in a 15 pt custom-height row of a 10.0 column: 35, 199 and 200 characters all fail
    custom = 'ht="15" customHeight="1"'
    for text in (long_text(35), long_text(199), long_text(200), long_text(150) + "\n" + long_text(49)):
        v = run(book([("S", ws({1: (custom, [c("A1", text, s_wrap)])}, cols=col(1, 1, stored(10.0))))], st))
        assert v["decision"] == "fail" and v["stats"]["wrapped_cut_off"] == 1, (len(text), v["stats"])
    assert v["stats"]["mistakes_by_rule"] == {"numbers": 0, "unwrapped_text": 0, "wrapped_text": 1}
    # a short text with room passes; a '####' number fails beside a cut-off label (two rules, two blocks)
    v = run(book([("S", ws({1: ("", [c("A1", long_text(12)), c("C1", 5.0)])}))], st))
    assert v["decision"] == "pass" and v["stats"]["text_fits_with_spill"] == 1, v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", long_text(50)), c("B1", 123456789, st.xf(numfmt=1))])},
                           cols=col(2, 2, stored(3.0))))], st))
    assert v["decision"] == "fail" and v["stats"]["mistakes_by_rule"] == {"numbers": 1, "unwrapped_text": 1, "wrapped_text": 0}


def test_69_text_run_size_out_of_range():
    """A rich-text run size outside Excel's 1-409 pt range (ChatGPT's writer stores 1000 for 10 pt, GPT-6 1331) never
    enlarges the text: the cell's size is used and the run is counted.  A 250-character paragraph whose runs say
    1000 / 1200 fits two lines of a 34 pt row in a 142-character column (measured at 1000 pt it needed 158 lines)."""
    st = Styles()
    s_wrap = st.xf("Arial", 10, wrap=True)
    para = [(long_text(120) + " ", 1000), (long_text(129), 1200)]
    v = run(book([("Instructions", ws({10: ('ht="34" customHeight="1"', [rich("B10", para, s_wrap)])},
                                      cols=col(2, 2, stored(141.29))))], st))
    assert v["decision"] == "pass" and v["stats"]["wrapped_cut_off"] == 0, v["mistakes"]
    assert v["stats"]["text_runs_size_out_of_range"] == 2, v["stats"]["text_runs_size_out_of_range"]
    # an in-range run size still counts: the same paragraph at 36 pt needs far more than the row shows
    para36 = [(long_text(120) + " ", 36), (long_text(129), 36)]
    v = run(book([("Instructions", ws({10: ('ht="34" customHeight="1"', [rich("B10", para36, s_wrap)])},
                                      cols=col(2, 2, stored(141.29))))], st))
    assert v["decision"] == "fail" and v["stats"]["wrapped_cut_off"] == 1, v["mistakes"]
    # unwrapped: a 1000 'pt' run is measured at the cell's 11 pt (fits a 10.0 column next to a filled B1)
    v = run(book([("S", ws({1: ("", [rich("A1", [("Revenue", 1000)]), c("B1", 5.0)])}, cols=col(1, 1, stored(10.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["text_cut_off"] == 0, v["stats"]


def test_69_formula_values():
    st = Styles()
    s_num = st.xf(numfmt="#,##0.00")
    narrow = col(1, 1, stored(3.0))
    # openpyxl-labelled file: the formula's cache is untrusted and nothing else decides -> GradingError
    x = ws({1: ("", [c("A1", 1234567.0, s_num, f="B1*2"), c("B1", 617283.5, s_num)])}, cols=col(1, 1, stored(3.0)) + col(2, 2, stored(20.0)))
    graded("untrusted_value", lambda: run(book([("S", x)], st, app="Microsoft Excel Compatible / Openpyxl 3.1.5")),
           "cannot decide S!A1", "untrusted", "writer=openpyxl")
    # the same with no cache at all
    x2 = ws({1: ("", [c("A1", None, s_num, f="B1*2"), c("B1", 617283.5, s_num)])}, cols=col(1, 1, stored(3.0)) + col(2, 2, stored(20.0)))
    graded("untrusted_value", lambda: run(book([("S", x2)], st)), "cannot decide S!A1")
    # Excel-labelled: the cache is read and the cell fails
    v = run(book([("S", x)], st, app="Microsoft Excel"))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and v["stats"]["formula_values_read"] == 1
    # a constant that fails elsewhere settles the verdict: the untrusted cell is left undecided, nothing raised
    x3 = ws({1: ("", [c("A1", 1234567.0, s_num, f="B1*2"), c("B1", 617283.5, s_num)]),
             2: ("", [c("A2", 9999999.0, s_num)])}, cols=col(1, 1, stored(3.0)) + col(2, 2, stored(20.0)))
    v = run(book([("S", x3)], st, app="Microsoft Excel Compatible / Openpyxl 3.1.5"))
    assert v["decision"] == "fail" and locs(v) == ["S!A2"]
    assert v["stats"]["undecided_cells"] == 1 and v["stats"]["undecided_examples"][0]["cell"] == "A1"
    assert "were skipped" in v["summary"] and v["stats"]["defaults"]["untrusted_value"]["count"] == 1, v
    with patched(M, UNTRUSTED_ONLY_IF_VERDICT_NEEDS=False):
        graded("untrusted_value", lambda: run(book([("S", x3)], st, app="Microsoft Excel Compatible / Openpyxl 3.1.5")), "cannot decide")
    # a shrink-to-fit formula cell never needs its value; a formula in a hidden column neither
    s_shr = st.xf(numfmt="#,##0.00", shrink=True)
    v = run(book([("S", ws({1: ("", [c("A1", None, s_shr, f="1/3")])}, cols=narrow))], st))
    assert v["decision"] == "pass" and v["stats"]["formula_values_read"] == 0
    v = run(book([("S", ws({1: ("", [c("A1", None, s_num, f="1/3")])}, cols=col(1, 1, stored(3.0), hidden=True)))], st))
    assert v["decision"] == "pass"
    # an ISO-date cell (t="d") in a 3.0 column fails; a text result never counts
    s_date = st.xf(numfmt=14)
    v = run(book([("S", ws({1: ("", [c("A1", "2026-01-01", s_date, t="d")])}, cols=narrow))], st))
    assert v["decision"] == "fail" and "1/1/2026" in v["mistakes"][0]["description"]


def test_69_conditional_formats():
    st = Styles()
    s_num = st.xf(numfmt="#,##0")
    d_wide = st.dxf("#,##0.000000")
    d_narrow = st.dxf('0.0,,"m"')
    cols8 = col(1, 2, stored(8.0))             # 61 px, text area 56
    cf_wide = (f'<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="{d_wide}" priority="1" '
               f'operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>')
    # 1234 fits as '1,234' but the firing CF format shows '1,234.000000' (12 chars) -> ####
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num)])}, cols=cols8, tail=cf_wide))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and "conditional-format" in v["mistakes"][0]["description"]
    assert v["stats"]["second_pass_sheets"] == ["S"] and v["stats"]["cf_cells"] == 1
    # the rule does not fire on a negative value: the base format decides (fits)
    v = run(book([("S", ws({1: ("", [c("A1", -1234.0, s_num)])}, cols=cols8, tail=cf_wide))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # a CF format can also narrow the display: 123,456,789 (####) shows '123.5m' under the rule -> pass
    cf_narrow = (f'<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="{d_narrow}" priority="1" '
                 f'operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>')
    v = run(book([("S", ws({1: ("", [c("A1", 123456789.0, s_num)])}, cols=cols8, tail=cf_narrow))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # an expression rule reading another cell is evaluated through the value source (Patrick 2026-10-05; it
    # raised 'cannot decide' until then): B1 = 1 > 0 -> the wide format -> ####; B1 = -1 -> the base fits
    cf_expr = (f'<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="{d_wide}" priority="1">'
               f'<formula>$B1&gt;0</formula></cfRule></conditionalFormatting>')
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num), c("B1", 1.0, s_num)])}, cols=cols8, tail=cf_expr))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and v["stats"]["cf_assumptions"]["cells"] == 0, v["stats"]
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num), c("B1", -1.0, s_num)])}, cols=cols8, tail=cf_expr))], st))
    assert v["decision"] == "pass", v["mistakes"]
    # a rule this check cannot evaluate is OFF (the base format decides) and recorded; decisive here
    cf_unev = (f'<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="{d_wide}" priority="1">'
               f'<formula>COUNTIF($B:$B,"&gt;0")&gt;0</formula></cfRule></conditionalFormatting>')
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num), c("B1", 1.0, s_num)])}, cols=cols8, tail=cf_unev))], st))
    a = v["stats"]["cf_assumptions"]
    assert v["decision"] == "pass" and a["cells"] == 1 and a["decisive_cells"] == 1 and "COUNTIF" in a["examples"][0], a
    # ... not decisive when both displays fit
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num), c("B1", 1.0, s_num)])}, cols=col(1, 2, stored(20.0)), tail=cf_unev))], st))
    assert v["decision"] == "pass" and v["stats"]["cf_assumptions"]["decisive_cells"] == 0, v["stats"]["cf_assumptions"]


def test_69_band_indent_unknown_face_negative_date():
    # T2 band example: '709,557.13' in Arial 10, UI 9.0 column, Normal Aptos Narrow 11: overflows under the
    # fractional 96-dpi model only -> pass, counted as a band cell
    st = Styles(("Aptos Narrow", 11))
    s_num = st.xf("Arial", 10, numfmt="#,##0.00")
    v = run(book([("S", ws({1: ("", [c("A1", 709557.13, s_num)])}, cols=col(1, 1, stored(9.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["band_cells"] == 1, v["stats"]
    assert v["stats"]["band_examples"][0]["overflow_px_by_model"]["w96f"] > 2
    # indent consumes width: '12,345,678' (8 x 7 + 2 x 4 = 64 px) in a 10.0 column (75 px, area 70) fits; with
    # indent 2 (2 x 3 Aptos Narrow spaces = 18 px) only 52 px are left under every model
    s_ind = st.xf("Arial", 10, numfmt="#,##0", horizontal="right", indent=2)
    s_plain = st.xf("Arial", 10, numfmt="#,##0", horizontal="right")
    v = run(book([("S", ws({1: ("", [c("A1", 12345678.0, s_plain), c("B1", 12345678.0, s_ind)])}, cols=col(1, 2, stored(10.0))))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!B1"], v["mistakes"]
    # an unknown face is measured as Calibri and listed
    s_seg = st.xf("Segoe UI", 10, numfmt="#,##0")
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.0, s_seg)])}, cols=col(1, 1, stored(12.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["unknown_faces"] == {"Segoe UI": 1}
    # a negative date shows ##### at any width: not a width defect
    s_date = st.xf("Arial", 10, numfmt=14)
    v = run(book([("S", ws({1: ("", [c("A1", -1.0, s_date)])}, cols=col(1, 1, stored(30.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["negative_dates"] == 1
    with patched(M, NEGATIVE_DATE_FAILS=True):
        v = run(book([("S", ws({1: ("", [c("A1", -1.0, s_date)])}, cols=col(1, 1, stored(30.0))))], st))
        assert v["decision"] == "fail"
    # rotated / fill / centre-across-selection numbers are skipped (unverified), counted
    s_rot = st.xf("Arial", 10, numfmt="#,##0", rotation=90)
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.0, s_rot)])}, cols=col(1, 1, stored(3.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["skipped_alignment"] == 1
    # accounting padding '_)' is measured as a parenthesis; a blank display (';;;') is nothing
    s_acc = st.xf("Arial", 10, numfmt=43)
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.0, s_acc)])}, cols=col(1, 1, stored(11.0))))], st))
    assert v["decision"] == "fail"                         # ' 1,234,567.00 ' + pad > 82 px
    s_blank = st.xf("Arial", 10, numfmt=";;;")
    v = run(book([("S", ws({1: ("", [c("A1", 1234567.0, s_blank)])}, cols=col(1, 1, stored(1.0))))], st))
    assert v["decision"] == "pass" and v["stats"]["blank_display"] == 1


def test_69_default_widths_and_grouping():
    st = Styles()
    s_num = st.xf(numfmt="#,##0.00")
    # no <col>: 64 px default (text area 59): '12,345.00' (7 x 7 + 2 x 4 = 57 px) fits; '123,456.00' (64 px)
    # overflows by 5 px at 96 dpi and by 5-7 px at 120 dpi (80 px default, area 75) -> sure
    v = run(book([("S", ws({1: ("", [c("A1", 12345.0, s_num)]), 2: ("", [c("A2", 123456.0, s_num)])}))], st))
    assert locs(v) == ["S!A2"] and v["stats"]["band_cells"] == 0, (v["mistakes"], v["stats"]["band_examples"])
    assert "about 5 px wider" in v["mistakes"][0]["description"]
    # a LibreOffice-style defaultColWidth of 8.5390625 (60 px strict, 64 lenient)
    fmt = '<sheetFormatPr defaultRowHeight="12.8" defaultColWidth="8.5390625"/>'
    v = run(book([("S", ws({1: ("", [c("A1", 123456.0, s_num)])}, fmt=fmt))], st))
    assert v["decision"] == "fail"
    # rectangles: J7:J9 and J12 as separate mistakes, counts in stats
    rows = {r: ("", [c(f"J{r}", 1234567.0, s_num)]) for r in (7, 8, 9, 12)}
    v = run(book([("Solution Model", ws(rows, cols=col(10, 10, stored(3.0))))], st))
    assert locs(v) == ["'Solution Model'!J7:J9", "'Solution Model'!J12"]
    assert v["stats"]["sure_overflows"] == 4 and v["stats"]["n_mistakes"] == 2
    assert v["stats"]["per_sheet"][0]["sure_overflows"] == 4
    # a number whose format is undefined (id 25 without <numFmt>) is General (Patrick 2026-10-06, attempt
    # 2777, measured in Excel: the format is dropped on repair), recorded; it raised until then
    st2 = Styles()
    st2.xfs.append('<xf numFmtId="25" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
    v = graded("unreadable_number_format", lambda: run(book([("S", ws({1: ("", [c("A1", 5.0, len(st2.xfs) - 1)])}))], st2)), "S!A1")
    assert v["decision"] == "pass", v
    # an invalid code (unterminated quoted literal, attempt 2777) likewise: the width is judged on General
    st3 = Styles()
    st3.numfmts[200] = '#,##0.00;\\-#,##0.00;"-""'
    st3.xfs.append('<xf numFmtId="200" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
    v = graded("unreadable_number_format", lambda: run(book([("S", ws({1: ("", [c("A1", 27643.237, len(st3.xfs) - 1)])},
                                                                  cols=col(1, 1, stored(3.0))))], st3)), "S!A1")
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v


# ============================================================================ second review (2026-10-04)
APP_OPX = "Microsoft Excel Compatible / Openpyxl 3.1.5"
APP_LO = "LibreOffice/7.6.4.1$MacOSX_X86_64"
SCAN = "SUM(_xlfn.SCAN(0,B1:B3,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)))"


def run_with(path, value_path=None, recalc=None):
    return Engine(path, [C69()], value_path=value_path, recalc=recalc).run()[C69.key]


def _pipeline(copy_path):
    """RecalcPolicy whose 'LibreOffice' returns copy_path (no soffice, no Excel; cache off)."""
    from detchecks.core import recalc as R

    def lo(src, out_dir, policy):
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, "copy.xlsx")
        shutil.copyfile(copy_path, out)
        return out

    def never(src, out_dir, policy):
        raise AssertionError("Excel must not be called")

    work = os.path.join(os.path.dirname(tmp("work")), "recalc_work")
    return R.RecalcPolicy(workdir=work, lo_runner=lo, excel_runner=never, excel_allowed=False,
                          excel_available=False, use_cache=False)


def test_69_review2_placeholder_types():
    """69-F1: a formula result is routed by its TRUSTED value, never by the delivered `t` (a cache attribute).
    An untrusted t="str" / "b" / "e" placeholder is undecided like a numeric one; with a value copy the copy's
    value type decides whether the cell is measured."""
    st = Styles()
    s0 = st.xf(numfmt=1)                                   # built-in '0'
    cols = col(1, 1, stored(3.0)) + col(2, 2, stored(40.0))
    big = 123456789
    # openpyxl t="str" placeholder with an empty cache: undecided -> GradingError (was a silent pass)
    x = ws({1: ("", [c("A1", "", s0, f="B1*1", t="str"), c("B1", big)])}, cols=cols)
    graded("untrusted_value", lambda: run(book([("S", x)], st, app=APP_OPX)), "cannot decide S!A1", "untrusted", "writer=openpyxl")
    # ... with a value copy holding the number: measured and failed (was a pass)
    copy = book([("S", ws({1: ("", [c("A1", big, s0, f="B1*1"), c("B1", big)])}, cols=cols))], st)
    v = run_with(book([("S", x)], st, app=APP_OPX), value_path=copy)
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    # a t="b" placeholder (openpyxl) whose copy value is a number: measured and failed
    xb = ws({1: ("", [c("A1", 0, s0, f="IF(B1>0,B1,FALSE)", t="b"), c("B1", big)])}, cols=cols)
    copy_b = book([("S", ws({1: ("", [c("A1", big, s0, f="IF(B1>0,B1,FALSE)"), c("B1", big)])}, cols=cols))], st)
    v = run_with(book([("S", xb)], st, app=APP_OPX), value_path=copy_b)
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    # the copy decides the other way too: a numeric placeholder whose copy value is text is never measured as a
    # number (it goes to the cut-off text rule instead: here it runs into the empty 40.0-wide B1 and is visible;
    # the referenced number sits in C1 since 2026-10-04 - in B1 it would cut the label off)
    xn = ws({1: ("", [c("A1", 0, s0, f='IF(C1>0,"far too long a label","")'), c("C1", big)])}, cols=cols)
    copy_t = book([("S", ws({1: ("", [c("A1", "far too long a label", s0, f='IF(C1>0,"far too long a label","")'),
                                       c("C1", big)])}, cols=cols))], st)
    v = run_with(book([("S", xn)], st, app=APP_OPX), value_path=copy_t)
    assert v["decision"] == "pass" and v["stats"]["formula_nonnumeric"] == 1, v["mistakes"]
    assert v["stats"]["text_formula_results"] == 1 and v["stats"]["numeric_cells"] == 1
    # a LibreOffice-written #NAME? cache read without the pipeline is untrusted (core policy) -> undecided
    xe = ws({1: ("", [c("A1", "#NAME?", s0, f="FOO(1)", t="e")])}, cols=cols)
    graded("untrusted_value", lambda: run(book([("S", xe)], st, app=APP_LO)), "cannot decide S!A1", "writer=libreoffice")
    # trusted (Excel) text and error results are not numbers this check measures (the text goes to the cut-off
    # text rule: it runs into the empty B1 and is visible; the error sits in C1 since 2026-10-04 - in B1 it would
    # cut the text off)
    xt = ws({1: ("", [c("A1", "far too long a label", s0, f='"far too long a label"', t="str"),
                      c("C1", "#DIV/0!", s0, f="1/0", t="e")])}, cols=cols)
    v = run(book([("S", xt)], st, app="Microsoft Excel"))
    assert v["decision"] == "pass" and v["stats"]["formula_nonnumeric"] == 2 and v["stats"]["numeric_cells"] == 0
    assert v["stats"]["text_formula_results"] == 1
    # the recalc pipeline (production path, Excel off): a LibreOffice gap (#NAME? on SCAN) is a vetted, trusted
    # error and is skipped without a loud error (ruling 2026-10-04); a number in the copy is measured
    xs = ws({1: ("", [c("A1", "", s0, f=SCAN, t="str"), c("B1", 1)]), 2: ("", [c("B2", 2)]), 3: ("", [c("B3", 3)])},
            cols=cols)
    copy_gap = book([("S", ws({1: ("", [c("A1", "#NAME?", s0, f=SCAN, t="e"), c("B1", 1)]),
                               2: ("", [c("B2", 2)]), 3: ("", [c("B3", 3)])}, cols=cols))], st, app=APP_LO)
    v = run_with(book([("S", xs)], st, app=APP_OPX), recalc=_pipeline(copy_gap))
    assert v["decision"] == "pass" and v["stats"]["formula_nonnumeric"] == 1 and v["stats"]["undecided_cells"] == 0, v
    copy_num = book([("S", ws({1: ("", [c("A1", big, s0, f=SCAN), c("B1", 1)]),
                               2: ("", [c("B2", 2)]), 3: ("", [c("B3", 3)])}, cols=cols))], st, app=APP_LO)
    v = run_with(book([("S", xs)], st, app=APP_OPX), recalc=_pipeline(copy_num))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]


def test_69_review2_iso_date_general():
    """69-F5: an ISO-typed date (t="d") under General shows its serial (46023): the General shortenings start
    from that number, so it overflows a 3.0 column (no '0' alternative any more) and fits a 12.0 one."""
    st = Styles()
    v = run(book([("S", ws({1: ("", [c("A1", "2026-01-01T00:00:00", t="d")])}, cols=col(1, 1, stored(3.0))))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and "46023" in v["mistakes"][0]["description"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", "2026-01-01T00:00:00", t="d")])}, cols=col(1, 1, stored(12.0))))], st))
    assert v["decision"] == "pass", v["mistakes"]


def test_69_review2_not_displayed_never_undecided():
    """69-F2: an untrusted formula result whose display cannot depend on its value is not undecided: a format
    that prints nothing for any number (';;;') and a covered (non-anchor) merged cell.  Controls: a
    format-setting conditional rule over the ';;;' cell, and the merge anchor, stay undecided."""
    st = Styles()
    s_blank = st.xf(numfmt=";;;")
    s0 = st.xf(numfmt=1)
    narrow = col(1, 2, stored(3.0))
    v = run(book([("S", ws({1: ("", [c("A1", None, s_blank, f="B1*1"), c("B1", 1)])}, cols=narrow))], st, app=APP_OPX))
    assert v["decision"] == "pass" and v["stats"]["undecided_cells"] == 0 and v["stats"]["undecided_not_displayed"] == 1
    d = st.dxf("0.000")
    cf = (f'<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="{d}" priority="1" operator="greaterThan">'
          f'<formula>0</formula></cfRule></conditionalFormatting>')
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, s_blank, f="B1*1"), c("B1", 1)])}, cols=narrow, tail=cf))],
                            st, app=APP_OPX)), "cannot decide S!A1")
    merge = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
    v = run(book([("S", ws({1: ("", [c("A1", 5, s0), c("B1", None, s0, f="A1*2")])}, cols=narrow, tail=merge))], st, app=APP_OPX))
    assert v["decision"] == "pass" and v["stats"]["undecided_not_displayed"] == 1, v
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, s0, f="C1*2"), c("C1", 5)])}, cols=narrow, tail=merge))],
                            st, app=APP_OPX)), "cannot decide S!A1")
    # an unverified rendering (numfmt certain=False) in a covered merged cell is not undecided either
    s_unc = st.xf(numfmt="#,##0.000000;[>100]0")
    v = run(book([("S", ws({1: ("", [c("A1", 5, s0), c("B1", -1234567.0, s_unc)])}, cols=narrow, tail=merge))], st))
    assert v["decision"] == "pass" and v["stats"]["undecided_cells"] == 0, v
    graded("unverified_number_format", lambda: run(book([("S", ws({1: ("", [c("A1", -1234567.0, s_unc)])}, cols=narrow))], st)), "not verified")


def test_69_review2_hidden_merge_anchor():
    """69-F4: a merge anchor in a hidden column (or row) is shown across the span's visible part; it is measured
    there in a second pass.  Nothing visible in the span (or the anchor's row hidden in a one-row merge): skipped."""
    st = Styles()
    s0 = st.xf(numfmt=1)
    big = 123456789                                         # needs 63 px; 3.0 column = 26 px (area 21)
    merge = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
    row1 = {1: ("", [c("A1", big, s0)])}
    v = run(book([("S", ws(row1, cols=col(1, 1, hidden=True) + col(2, 2, stored(3.0)), tail=merge))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    assert v["stats"]["hidden_anchor_cells"] == 1 and v["stats"]["second_pass_sheets"] == ["S"]
    v = run(book([("S", ws(row1, cols=col(1, 1, hidden=True) + col(2, 2, stored(40.0)), tail=merge))], st))
    assert v["decision"] == "pass" and v["stats"]["hidden_anchor_cells"] == 1
    v = run(book([("S", ws(row1, cols=col(1, 2, hidden=True), tail=merge))], st))
    assert v["decision"] == "pass" and v["stats"]["hidden_anchor_cells"] == 0
    with patched(M, HIDDEN_ANCHOR_SPAN=False):
        v = run(book([("S", ws(row1, cols=col(1, 1, hidden=True) + col(2, 2, stored(3.0)), tail=merge))], st))
        assert v["decision"] == "pass" and v["stats"]["second_pass_sheets"] == []
    # hidden row: a vertical merge shows the anchor in the next row (column A's width); a one-row merge shows nothing
    vmerge = '<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>'
    rows_h = {1: ('hidden="1"', [c("A1", big, s0)]), 2: ("", [c("A2", None, s0)])}
    v = run(book([("S", ws(rows_h, cols=col(1, 1, stored(3.0)), tail=vmerge))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    v = run(book([("S", ws({1: ('hidden="1"', [c("A1", big, s0)])}, cols=col(1, 2, stored(3.0)), tail=merge))], st))
    assert v["decision"] == "pass" and v["stats"]["hidden_anchor_cells"] == 0
    # zeroHeight sheet: row 2 shown by its own height -> the anchor shows there; without a shown row -> nothing
    zfmt = '<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/>'
    rows_z = {1: ("", [c("A1", big, s0)]), 2: ('ht="15" customHeight="1"', [c("A2", None, s0)])}
    v = run(book([("S", ws(rows_z, cols=col(1, 1, stored(3.0)), fmt=zfmt, tail=vmerge))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"], v["mistakes"]
    v = run(book([("S", ws({1: ("", [c("A1", big, s0)])}, cols=col(1, 1, stored(3.0)), fmt=zfmt, tail=vmerge))], st))
    assert v["decision"] == "pass" and v["stats"]["hidden_anchor_cells"] == 0
    # an untrusted formula anchor there is undecided (raises) unless its format prints nothing
    s_blank = st.xf(numfmt=";;;")
    hid = col(1, 1, hidden=True) + col(2, 2, stored(3.0)) + col(3, 3, stored(40.0))
    graded("untrusted_value", lambda: run(book([("S", ws({1: ("", [c("A1", None, s0, f="C1*1"), c("C1", big)])}, cols=hid, tail=merge))],
                            st, app=APP_OPX)), "cannot decide S!A1")
    v = run(book([("S", ws({1: ("", [c("A1", None, s_blank, f="C1*1"), c("C1", big)])}, cols=hid, tail=merge))], st, app=APP_OPX))
    assert v["decision"] == "pass" and v["stats"]["undecided_not_displayed"] == 1, v
    # Excel-cached anchor value: measured
    v = run(book([("S", ws({1: ("", [c("A1", big, s0, f="C1*1"), c("C1", big)])}, cols=hid, tail=merge))], st,
                 app="Microsoft Excel"))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and v["stats"]["formula_values_read"] == 1, v["mistakes"]
    # a conditional-format number format over the hidden anchor: '1,234' fits B (8.0 = 61 px) but the firing
    # rule's '1,234.000000' does not
    s_num = st.xf(numfmt="#,##0")
    d_wide = st.dxf("#,##0.000000")
    cf = (f'<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="{d_wide}" priority="1" operator="greaterThan">'
          f'<formula>0</formula></cfRule></conditionalFormatting>')
    v = run(book([("S", ws({1: ("", [c("A1", 1234.0, s_num)])}, cols=col(1, 1, hidden=True) + col(2, 2, stored(8.0)),
                           tail=merge + cf))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!A1"] and "conditional-format" in v["mistakes"][0]["description"]
    v = run(book([("S", ws({1: ("", [c("A1", -1234.0, s_num)])}, cols=col(1, 1, hidden=True) + col(2, 2, stored(8.0)),
                           tail=merge + cf))], st))
    assert v["decision"] == "pass", v["mistakes"]


def test_69_overlapping_cols_later_entry_wins():
    """Overlapping <col> entries (GPT-6 tooling: C:XFD 18 followed by D:D 3 and E:E 3): the core reader's one rule,
    the later entry wins (SheetHead.col_info / col_run, shared with 70 and 73).  Before the fix every column after
    E read the sheet default (8.43): F1's 123456789 showed '####' and C3's long note was cut off at G."""
    st = Styles()
    s0 = st.xf(numfmt=1)                      # built-in '0'
    cols = col(3, 16384, stored(18.0)) + col(4, 4, stored(3.0)) + col(5, 5, stored(3.0))
    rows = {1: ("", [c("D1", 123456789, s0), c("F1", 123456789, s0)])}
    v = run(book([("S", ws(rows, cols=cols))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!D1"], v["mistakes"]       # D is 3 wide; F is 18 wide
    # a 210-character note in C3 cut off by H3: room C..G = 80 + 3 + 3 + 80 + 80 characters (fits); with F and G
    # at the default it would be about 745 px against the note's ~1,200 px
    cols = col(3, 16384, stored(80.0)) + col(4, 4, stored(3.0)) + col(5, 5, stored(3.0))
    rows = {3: ("", [c("C3", long_text(210)), c("H3", 5.0)])}
    v = run(book([("S", ws(rows, cols=cols))], st))
    assert v["decision"] == "pass" and not v["stats"]["text_cut_off"], (v["mistakes"], v["stats"])
    # a later E:G entry narrows F and G too: room 80 + 3 + 3 + 3 + 3 characters, the note is cut off
    cols_narrow = col(3, 16384, stored(80.0)) + col(4, 4, stored(3.0)) + col(5, 7, stored(3.0))
    v = run(book([("S", ws(rows, cols=cols_narrow))], st))
    assert v["decision"] == "fail" and locs(v) == ["S!C3"], v["mistakes"]       # room 80+3+3+3+3: cut off


TESTS = [test_69_metrics, test_69_narrow_number_date_percent, test_69_one_digit_overflow_and_general,
         test_69_shrink_merge_hidden, test_69_text_cut_off_by_neighbours, test_69_text_alignment_directions,
         test_69_text_merges_hidden_columns_and_edges, test_69_wrapped_text_rows,
         test_69_text_values_formats_and_fonts, test_69_text_mistakes_band_and_stats, test_69_text_clip_any_length,
         test_69_text_run_size_out_of_range,
         test_69_formula_values, test_69_conditional_formats,
         test_69_band_indent_unknown_face_negative_date, test_69_default_widths_and_grouping,
         test_69_review2_placeholder_types, test_69_review2_iso_date_general,
         test_69_review2_not_displayed_never_undecided, test_69_review2_hidden_merge_anchor,
         test_69_overlapping_cols_later_entry_wins]


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
