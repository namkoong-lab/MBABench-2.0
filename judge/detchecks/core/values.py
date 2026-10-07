"""Value provenance: which writer produced the workbook, whether formula caches can be
trusted, and (optionally) formula values taken from a LibreOffice recalculation copy.

Writer classes (detect_writer):
  excel        Application 'Microsoft Excel' / 'Microsoft Excel Online' / 'Microsoft Macintosh
               Excel' AND evidence that Excel itself wrote the parts: AppVersion present,
               <fileVersion appName="xl">, and Excel's double-quoted XML declaration on the
               workbook part, app.xml and every sheet part.  Agents' Python tools copy app.xml
               and fileVersion from the starting file but re-serialise the parts (single-quoted
               or no declaration), so such a file is 'unknown', not 'excel'.  Not XlsxWriter.
  xlsxwriter   Application 'Microsoft Excel' but XlsxWriter's signature: fileVersion
               rupBuild="4505" (lastEdited 4), or AppVersion 12.0000 + calcId 124519 +
               fullCalcOnLoad.  Its formula caches are all-zero placeholders.
  openpyxl     Application contains 'openpyxl' (openpyxl >= 3: 'Microsoft Excel Compatible /
               Openpyxl x.y'), or dc:creator 'openpyxl' with calcId 124519.  openpyxl
               never computes formulas (no cached values).
  libreoffice  Application 'LibreOffice...' or fileVersion appName="Calc".
  unknown      anything else (ChatGPT's spreadsheet writer, Google Sheets, .NET libraries...)

Trust policy for a FORMULA cell's cached <v> (TrustPolicy.cached_trusted):
  - no cached value (no <v>, or empty <v> on a non-string formula)  -> untrusted
  - writer xlsxwriter or openpyxl                                     -> untrusted
  - writer libreoffice and the cache is #NAME? or #VALUE!             -> untrusted ("unmeasured")
  - otherwise (excel, libreoffice, unknown with a cache)              -> trusted
Constants (cells without a formula and outside array ranges) are always trusted.

With value_path (a LibreOffice recalculation copy of the same workbook, same sheet names)
formula-cell values come from the copy instead (cell.value_source == 'recalc'): the copy's
sheet is streamed in lockstep with the delivered sheet (merge-join on (row, col)); only
values are taken from it - never styles or structure.  #NAME? / #VALUE! from LibreOffice,
a missing cell or a missing sheet in the copy -> untrusted.  The recalc pipeline
(core/recalc.py) opens the copy with errors_vetted=True once its gap scan has shown that those
error values are genuine (or when the copy was made by Excel): they are then trusted values.

NO FALLBACK for values: cell.value of a formula result raises GradingError when the value
is untrusted, and when the stream was not opened for a check with needs_values = True (so
a forgotten flag cannot silently read caches while a value_path copy was supplied).  A check
that needs values sets needs_values = True and tests cell.value_trusted first, or calls
Check.require_value(cell).  cell.unverified_value bypasses both guards (diagnostics only).
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
import zlib
from dataclasses import dataclass, field
from typing import Optional

from ..errors import GradingError
from .package import Package, SheetInfo, truthy
from .refs import col_to_index
from .sheet import ExcelError, SheetContext, _REF, _loc, parse_runs

UNMEASURED_LO_ERRORS = ("#NAME?", "#VALUE!")
WRITERS = ("excel", "xlsxwriter", "openpyxl", "libreoffice", "unknown")
UNTRUSTED_WRITERS = ("xlsxwriter", "openpyxl")


@dataclass
class Provenance:
    writer: str                        # excel | xlsxwriter | openpyxl | libreoffice | unknown
    evidence: str                      # why (human readable)
    formula_caches_trusted: bool       # policy for cached values of formula cells (see module doc)
    full_calc_on_load: bool            # calcPr fullCalcOnLoad="1": Excel recalculates on open
    application: Optional[str] = None
    app_version: Optional[str] = None
    value_path: Optional[str] = None   # LibreOffice recalculation copy, when supplied
    value_writer: Optional[str] = None
    value_sheets: list = field(default_factory=list)   # sheet names present in the copy
    value_kind: Optional[str] = None   # recalc pipeline: 'cache' | 'libreoffice' | 'excel' (None: manual value_path / none)
    recalc: Optional[dict] = None      # recalc pipeline summary (core/recalc.py ValuePlan.summary())

    def as_dict(self) -> dict:
        return {"writer": self.writer, "evidence": self.evidence,
                "formula_caches_trusted": self.formula_caches_trusted,
                "full_calc_on_load": self.full_calc_on_load, "application": self.application,
                "value_path": self.value_path, "value_writer": self.value_writer,
                "value_kind": self.value_kind, "recalc": self.recalc}


_DECL = re.compile(rb"^\s*<\?xml\s([^?]*)\?>")
EXCEL_APPS = ("microsoft excel", "microsoft macintosh excel")


def _decl_style(pkg: Package, part: Optional[str]) -> Optional[str]:
    """XML declaration style of a part: 'double' (Excel / LibreOffice / XlsxWriter),
    'single' (Python ElementTree / lxml), 'none', 'unreadable' (corrupt member), or None
    when the part is absent."""
    if not part or not pkg.has(part):
        return None
    try:
        with pkg.open_part(part) as fh:
            head = fh.read(256)
    except (zipfile.BadZipFile, zlib.error, OSError, EOFError, RuntimeError):
        return "unreadable"
    if head[:3] == b"\xef\xbb\xbf":
        head = head[3:]
    m = _DECL.match(head)
    if m is None:
        return "none"
    return "single" if b"'" in m.group(1) else "double"


def _not_excel_evidence(pkg: Package) -> list[str]:
    """Reasons why an Excel-labelled file was NOT saved by Excel (empty list = Excel's own)."""
    why = []
    if not (pkg.app.app_version or "").strip():
        why.append("no AppVersion in app.xml")
    if (pkg.file_version.get("appName") or "").lower() != "xl":
        why.append(f"workbook.xml fileVersion appName={pkg.file_version.get('appName')!r}")
    app_part = None
    for r in pkg.rels(""):
        if r.type_short.lower() in ("extended-properties", "extendedproperties") and r.part:
            app_part = r.part
    parts = [("workbook part", pkg.main_part), ("app.xml", app_part or pkg.member("docProps/app.xml"))]
    parts += [(f"sheet {s.name!r}", s.part) for s in pkg.sheets if s.part]
    for what, part in parts:
        st = _decl_style(pkg, part)
        if st is not None and st != "double":
            why.append({"single": f"{what} has a single-quoted XML declaration",
                        "none": f"{what} has no XML declaration"}.get(st, f"{what} is unreadable"))
            break
    return why


def detect_writer(pkg: Package) -> tuple[str, str]:
    app = (pkg.app.application or "").strip()
    al = app.lower()
    fv = pkg.file_version or {}
    cp = pkg.calc_pr or {}
    creator = (pkg.core.creator or "").strip().lower()
    if "openpyxl" in al:
        return "openpyxl", f"Application={app!r}"
    if al.startswith("libreoffice") or (fv.get("appName") or "").lower() == "calc":
        return "libreoffice", f"Application={app!r}, fileVersion appName={fv.get('appName')!r}"
    if al.startswith(EXCEL_APPS):
        if fv.get("rupBuild") == "4505" or (pkg.app.app_version == "12.0000" and cp.get("calcId") == "124519"
                                            and truthy(cp.get("fullCalcOnLoad"))):
            return "xlsxwriter", (f"Application={app!r} with XlsxWriter signature (fileVersion rupBuild="
                                  f"{fv.get('rupBuild')!r}, AppVersion={pkg.app.app_version!r}, calcPr={cp})")
        if creator == "openpyxl" and cp.get("calcId") == "124519":
            return "openpyxl", "dc:creator=openpyxl, calcId 124519 (older openpyxl labels itself Microsoft Excel)"
        why = _not_excel_evidence(pkg)
        if why:
            return "unknown", (f"Application={app!r} but not saved by Excel: {'; '.join(why)} "
                               f"(app.xml / fileVersion copied from another file); creator={pkg.core.creator!r}")
        return "excel", (f"Application={app!r} AppVersion={pkg.app.app_version!r}, fileVersion appName='xl' "
                         f"rupBuild={fv.get('rupBuild')!r}, Excel-style XML declarations")
    if creator == "openpyxl" and cp.get("calcId") == "124519":
        return "openpyxl", "dc:creator=openpyxl, calcId 124519"
    return "unknown", f"Application={app or None!r}, fileVersion={fv or None}, creator={pkg.core.creator!r}"


class TrustPolicy:
    def __init__(self, writer: str):
        self.writer = writer
        self.caches_trusted = writer not in UNTRUSTED_WRITERS

    def cached_trusted(self, cell) -> bool:
        if not cell.has_value:
            return False
        if not self.caches_trusted:
            return False
        if self.writer == "libreoffice" and cell.t == "e" and (cell.raw or "").strip() in UNMEASURED_LO_ERRORS:
            return False
        return True


def detect_provenance(pkg: Package, value_path: Optional[str] = None) -> Provenance:
    writer, ev = detect_writer(pkg)
    prov = Provenance(writer, ev, writer not in UNTRUSTED_WRITERS, truthy(pkg.calc_pr.get("fullCalcOnLoad")),
                      pkg.app.application, pkg.app.app_version, value_path)
    pkg.provenance = prov
    return prov


# ---------------------------------------------------------------------------- recalc copy
class ValueSource:
    """A recalculation copy (LibreOffice or Excel) opened for values only.

    errors_vetted=False (default, a manually supplied value_path): a copy value of #NAME? /
    #VALUE! is 'unmeasured' -> untrusted.  errors_vetted=True (set by the recalc pipeline,
    core/recalc.py, after its gap scan found no LibreOffice gap, or for an Excel-made copy):
    those error values are the display Excel would show too and are trusted."""

    def __init__(self, value_path: str, errors_vetted: bool = False):
        self.path = value_path
        self.errors_vetted = errors_vetted
        self.pkg = Package.open(value_path)
        if not self.pkg.is_spreadsheetml:
            raise GradingError(f"value_path {value_path} is not a readable .xlsx: {self.pkg.unsupported_reason}",
                               path=value_path)
        self.writer = detect_writer(self.pkg)[0]

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.pkg.sheets]

    def cursor(self, sheet_name: str) -> "SheetValueCursor":
        info = self.pkg.sheet(sheet_name)
        return SheetValueCursor(self.pkg, info, self.errors_vetted)

    def close(self):
        self.pkg.close()


class SheetValueCursor:
    """Monotonic (row, col) lookup into one sheet of the recalc copy.  get() must be called
    with increasing (row, col) - the engine does so while streaming the delivered sheet.
    Returns ('ok', value) | ('unmeasured', value) | ('missing', None)."""

    MISSING = ("missing", None)

    def __init__(self, pkg: Package, info: Optional[SheetInfo], errors_vetted: bool = False):
        self.pkg = pkg
        self.info = info
        self.errors_vetted = errors_vetted
        self._gen = self._cells() if info is not None and info.part else iter(())
        self._cur = next(self._gen, None)

    def _cells(self):
        sst = None
        cur_row = 0
        last_col = 0
        sheet_data = None
        with self.pkg.open_part(self.info.part) as fh:
            for ev, el in ET.iterparse(fh, events=("start", "end")):
                ln = _loc(el.tag)
                if ev == "start":
                    if ln == "row":
                        r = el.get("r")
                        cur_row = int(r) if r else cur_row + 1
                        last_col = 0
                    elif ln == "sheetData":
                        sheet_data = el
                    continue
                if ln == "c":
                    ref = el.get("r")
                    m = _REF.match(ref) if ref else None
                    if m is not None:
                        row, col = int(m.group(2)), col_to_index(m.group(1))
                    else:
                        row, col = cur_row, last_col + 1
                    last_col = col
                    t = el.get("t") or "n"
                    raw = None
                    for ch in el:
                        cl = _loc(ch.tag)
                        if cl == "v":
                            raw = ch.text if ch.text is not None else ""
                        elif cl == "is":
                            raw = parse_runs(ch)[0]
                    el.clear()
                    if raw is None:
                        continue
                    if t == "s":
                        if sst is None:
                            sst = self.pkg.shared_strings
                        try:
                            val = sst[int(raw)]
                        except (ValueError, IndexError):
                            continue
                    elif t == "n":
                        try:
                            val = float(raw) if raw != "" else None
                        except ValueError:
                            continue
                    elif t == "b":
                        val = raw.strip().lower() in ("1", "true")
                    elif t == "e":
                        val = ExcelError(raw)
                    else:
                        val = raw
                    status = ("unmeasured" if (t == "e" and raw.strip() in UNMEASURED_LO_ERRORS
                                               and not self.errors_vetted) else "ok")
                    if val is None and t == "n":
                        continue
                    yield row, col, status, val
                elif ln == "row":
                    # drop the finished row from <sheetData> too (an emptied element left in the
                    # tree costs ~90 bytes per row: 88 MB per million rows)
                    el.clear()
                    if sheet_data is not None:
                        sheet_data.clear()

    def get(self, row: int, col: int):
        cur = self._cur
        while cur is not None and (cur[0] < row or (cur[0] == row and cur[1] < col)):
            cur = next(self._gen, None)
        self._cur = cur
        if cur is not None and cur[0] == row and cur[1] == col:
            return (cur[2], cur[3])
        return self.MISSING


def make_context(pkg: Package, info: SheetInfo, value_source: Optional[ValueSource],
                 needs_values: bool = False) -> SheetContext:
    """Per-sheet SheetContext: decoding + trust policy.  needs_values=True enables reading
    formula results through cell.value (trusted ones only) and opens the recalc cursor when a
    value_source is given."""
    prov = pkg.provenance or detect_provenance(pkg)
    policy = TrustPolicy(prov.writer)
    cursor = None
    if needs_values and value_source is not None:
        cursor = value_source.cursor(info.name)
    return SheetContext(info.name, pkg, policy, cursor, values_enabled=needs_values)
