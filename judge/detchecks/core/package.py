"""OOXML package reader: format classification, content types, relationships,
workbook.xml, external links, document properties, drawings/charts, pivot caches,
tables and comments.  Stdlib only (zipfile + xml.etree); never openpyxl.

Everything is matched by LOCAL element/attribute name, so Strict OOXML
(http://purl.oclc.org/ooxml/...) and namespace-prefixed parts (<x:workbook>) read
exactly like Transitional ones.

Entry point:  pkg = Package.open(path)
  - never raises for a bad/unsupported format: it sets pkg.format and
    pkg.unsupported_reason, and pkg.is_spreadsheetml is False.  (Check 77 must be
    able to grade such files; every other check raises GradingError on them.)
  - raises GradingError only when the path does not exist / cannot be read at all.
"""
from __future__ import annotations

import io
import os
import posixpath
import re
import struct
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any, Optional
from urllib.parse import unquote

from ..errors import GradingError

ZIP_MAGIC = b"PK\x03\x04"
ZIP_EMPTY_MAGIC = b"PK\x05\x06"
OLE2_MAGIC = bytes.fromhex("d0cf11e0a1b11ae1")

MAIN_CT = {   # keys lower-case (MIME types compare case-insensitively)
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "xlsx",
    "application/vnd.ms-excel.sheet.macroenabled.main+xml": "xlsm",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "xltx",
    "application/vnd.ms-excel.template.macroenabled.main+xml": "xltm",
    "application/vnd.ms-excel.addin.macroenabled.main+xml": "xlam",
    "application/vnd.ms-excel.sheet.binary.macroenabled.main": "xlsb",
}
# content types that say nothing about the part (then the extension decides and the root
# element of the main part is checked)
GENERIC_CT = ("", "application/xml", "text/xml", "application/octet-stream")
SPREADSHEETML_FORMATS = ("xlsx", "xlsm", "xltx", "xltm", "xlam")
# namespaces of the <workbook> root element (Transitional, Strict)
SML_MAIN_NS = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main",
               "http://purl.oclc.org/ooxml/spreadsheetml/main")

# relationship type (last path component, lower-case) -> sheet kind.  Any other type on a
# <sheet> gives kind 'unknown', which the engine refuses to grade (GradingError).
SHEET_KINDS = {
    "worksheet": "worksheet",
    "chartsheet": "chartsheet",
    "dialogsheet": "dialogsheet",
    "xlmacrosheet": "macrosheet",
    "xlintlmacrosheet": "macrosheet",
}


# ----------------------------------------------------------------------------- xml helpers
def local(tag: str) -> str:
    """Local part of an ElementTree tag or attribute name ('{ns}row' -> 'row')."""
    return tag.rsplit("}", 1)[-1] if tag[:1] == "{" else tag.rsplit(":", 1)[-1]


def attr(el, name: str, default=None):
    """Attribute by local name (also finds r:id / strict-namespaced attributes)."""
    v = el.get(name)
    if v is not None:
        return v
    for k, val in el.attrib.items():
        if k[:1] == "{" and k.rsplit("}", 1)[-1] == name:
            return val
    return default


def rel_id(el) -> Optional[str]:
    """r:id of an element, in the Transitional or the Strict relationships namespace."""
    for k, v in el.attrib.items():
        if k[:1] == "{" and k.endswith("}id") and "relationships" in k:
            return v
    return None


def truthy(v, default: bool = False) -> bool:
    if v is None:
        return default
    return str(v).strip().lower() in ("1", "true", "on")


def children(el, name: str):
    return [c for c in el if local(c.tag) == name]


def child(el, name: str):
    for c in el:
        if local(c.tag) == name:
            return c
    return None


def text_of(el) -> str:
    """Concatenated text of all <t> descendants (rich runs included, phonetic <rPh> runs excluded)."""
    parts: list[str] = []
    _collect_t(el, parts)
    return "".join(parts)


def _collect_t(el, parts):
    for c in el:
        ln = local(c.tag)
        if ln == "t":
            parts.append(c.text or "")
        elif ln in ("r",):
            _collect_t(c, parts)
        # rPh (phonetic) and phoneticPr are skipped on purpose


def parse_xml(data: bytes, what: str):
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise GradingError(f"XML parse error in {what}: {e}") from e


# ----------------------------------------------------------------------------- dataclasses
@dataclass(frozen=True)
class Rel:
    id: str
    type: str                 # full relationship type URI
    type_short: str           # last path component ('worksheet', 'styles', 'externalLinkPath', ...)
    target: str               # Target attribute as written
    external: bool            # TargetMode="External"
    part: Optional[str]       # resolved zip member name (None for external targets / missing parts)


@dataclass
class SheetInfo:
    index: int                # 0-based position in workbook order (tab order)
    name: str
    sheet_id: Optional[str]
    state: str                # visible | hidden | veryHidden
    rid: Optional[str]
    part: Optional[str]       # zip member name of the sheet part (None if the relationship is broken)
    kind: str                 # worksheet | chartsheet | dialogsheet | macrosheet | unknown
    rel_type: Optional[str] = None   # full relationship type URI (None if the relationship is missing)

    @property
    def hidden(self) -> bool:
        return self.state in ("hidden", "veryHidden")


@dataclass
class DefinedName:
    """One <definedName>.  `scope` raises GradingError when localSheetId is present but
    does not name a sheet (out of range or not an integer): such a name is neither
    workbook-scoped nor attached to a known sheet, so using it would be a guess."""
    name: str
    text: str                 # raw definition (formula text without '=')
    local_sheet_id: Optional[int]
    _scope: Optional[str]     # resolved sheet name (read it through .scope)
    hidden: bool
    builtin: bool             # name starts with '_xlnm.' (Print_Area, _FilterDatabase, ...)
    function: bool = False
    vb_procedure: bool = False
    xlm: bool = False
    comment: Optional[str] = None
    attrs: dict = field(default_factory=dict)
    scope_error: Optional[str] = None   # why the scope is unknown (bad localSheetId), else None

    @property
    def scope(self) -> Optional[str]:
        """Sheet name for sheet-scoped names, None for workbook scope."""
        if self.scope_error is not None:
            raise GradingError(f"defined name {self.name!r}: {self.scope_error}")
        return self._scope


@dataclass
class ExternalLink:
    index: int                # n of the '[n]' prefix used in formulas (1-based, workbook order)
    rid: Optional[str]
    part: Optional[str]
    kind: Optional[str]       # externalBook | ddeLink | oleLink | None (part missing)
    target: Optional[str]     # path/URL of the linked file (from the part's relationships)
    target_external: bool
    sheet_names: list = field(default_factory=list)
    defined_names: list = field(default_factory=list)   # [(name, refersTo, sheetId)]
    has_cached_data: bool = False


@dataclass
class AppProps:
    application: Optional[str] = None
    app_version: Optional[str] = None
    company: Optional[str] = None
    raw: dict = field(default_factory=dict)   # local name -> text for simple children


@dataclass
class CoreProps:
    creator: Optional[str] = None
    last_modified_by: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None
    title: Optional[str] = None
    raw: dict = field(default_factory=dict)


@dataclass
class Chart:
    sheet: Optional[str]      # sheet whose drawing holds the chart (None if unknown)
    part: str
    formulas: list            # every <*:f> text (series, categories, titles, data labels)


@dataclass
class Drawing:
    sheet: str
    part: str
    chart_parts: list
    textlinks: list           # shape textlink="..." formulas


@dataclass
class PivotCache:
    cache_id: Optional[str]
    part: Optional[str]
    source_type: Optional[str]
    source_ref: Optional[str]
    source_sheet: Optional[str]
    source_name: Optional[str]
    source_target: Optional[str]   # external workbook target when the source is external


@dataclass
class Table:
    sheet: str
    part: str
    name: Optional[str]
    display_name: Optional[str]
    ref: Optional[str]
    header_row_count: int
    totals_row_count: int
    columns: list             # [{'name', 'calculated', 'totals_formula', 'totals_function'}]


@dataclass
class Comment:
    ref: str
    author: Optional[str]
    text: str
    threaded: bool = False


# ----------------------------------------------------------------------------- package
class Package:
    """An opened workbook file.  See module docstring."""

    def __init__(self, path: str):
        self.path = path
        self.file_name = os.path.basename(path)
        self.extension = os.path.splitext(path)[1].lower()
        self.size = 0
        self.container = "missing"      # zip | ole2 | other | empty | missing
        self.format = "unknown"         # xlsx xlsm xltx xltm xlam xlsb xls ods encrypted corrupt unknown
        self.unsupported_reason: Optional[str] = None
        self.zip: Optional[zipfile.ZipFile] = None
        self.names: list[str] = []
        self._lower: dict[str, str] = {}
        self.ole_streams: Optional[list[str]] = None
        self.content_types_default: dict[str, str] = {}
        self.content_types_override: dict[str, str] = {}
        self.main_part: Optional[str] = None
        self.main_content_type: Optional[str] = None
        self.vba_parts: list[str] = []          # macro carriers related from the workbook part
        self.vba_file_present = False           # any member named *vbaProject.bin
        self.conformance: Optional[str] = None  # transitional | strict
        # workbook level
        self.sheets: list[SheetInfo] = []
        self.defined_names: list[DefinedName] = []
        self.calc_pr: dict = {}
        self.book_views: list[dict] = []
        self.workbook_pr: dict = {}
        self.file_version: dict = {}
        self.workbook_protection: Optional[dict] = None
        self.external_links: list[ExternalLink] = []
        self.pivot_cache_refs: list[tuple] = []  # (cacheId, rid)
        self.app = AppProps()
        self.core = CoreProps()
        self.task_meta: dict = {}
        self.provenance = None                  # set by values.detect_provenance()
        self.value_source = None                # set by the engine: the recalc copy formula values come from
                                                # (core.lookup reads referenced cells through it)
        self._rels_cache: dict[str, list[Rel]] = {}
        self._styles = None
        self._sst = None
        self._sst_rich = None
        self._charts = None
        self._drawings = None
        self._pivots = None

    # ------------------------------------------------------------------ opening
    @classmethod
    def open(cls, path: str) -> "Package":
        pkg = cls(path)
        if not os.path.isfile(path):
            raise GradingError(f"workbook not found: {path}", path=path)
        try:
            pkg.size = os.path.getsize(path)
            with open(path, "rb") as fh:
                magic = fh.read(8)
        except OSError as e:
            raise GradingError(f"cannot read {path}: {e}", path=path) from e
        if not magic:
            pkg.container, pkg.format = "empty", "unknown"
            pkg.unsupported_reason = "empty file"
            return pkg
        if magic == OLE2_MAGIC:
            pkg.container = "ole2"
            pkg.ole_streams = ole2_stream_names(path)
            streams = {s.lower() for s in (pkg.ole_streams or [])}
            if "encryptedpackage" in streams:
                pkg.format = "encrypted"
                pkg.unsupported_reason = "password-protected (encrypted) OOXML workbook (OLE2 EncryptedPackage)"
            else:
                pkg.format = "xls"
                pkg.unsupported_reason = "legacy binary .xls workbook (OLE2 compound file), not SpreadsheetML XML"
            return pkg
        if magic[:4] not in (ZIP_MAGIC, ZIP_EMPTY_MAGIC):
            pkg.container = "other"
            pkg.format = {".csv": "csv", ".txt": "csv", ".ods": "ods"}.get(pkg.extension, "unknown")
            pkg.unsupported_reason = f"not a zip or OLE2 file (magic {magic[:4]!r})"
            return pkg
        pkg.container = "zip"
        try:
            pkg.zip = zipfile.ZipFile(path)
            pkg.names = pkg.zip.namelist()
        except (zipfile.BadZipFile, OSError, ValueError) as e:
            pkg.format = "corrupt"
            pkg.unsupported_reason = f"corrupt zip: {e}"
            return pkg
        pkg._lower = {n.lower().lstrip("/"): n for n in pkg.names}
        try:
            pkg._classify()
        except Exception as e:  # noqa: BLE001 - malformed [Content_Types].xml / root rels
            pkg.format = "corrupt"
            pkg.unsupported_reason = f"package structure unreadable: {e}"
            return pkg
        if pkg.format in SPREADSHEETML_FORMATS:
            # a broken workbook part keeps the classified format (check 77 still needs it) but
            # makes the package unparseable: every other check raises GradingError on it
            try:
                pkg._read_workbook()
            except GradingError as e:
                pkg.unsupported_reason = str(e)
            except Exception as e:  # noqa: BLE001 - recorded, surfaced by the engine as GradingError
                pkg.unsupported_reason = f"workbook part unreadable: {type(e).__name__}: {e}"
        return pkg

    @property
    def is_spreadsheetml(self) -> bool:
        """True when the workbook and sheet parts are SpreadsheetML XML this reader can parse."""
        return self.format in SPREADSHEETML_FORMATS and self.unsupported_reason is None

    @property
    def has_vba(self) -> bool:
        """A VBA project (or XLM macro sheet) is related from the workbook part and present."""
        return bool(self.vba_parts)

    def close(self):
        if self.zip is not None:
            self.zip.close()
            self.zip = None

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ------------------------------------------------------------------ parts
    def member(self, part: Optional[str]) -> Optional[str]:
        """Actual zip member name for a part name (case-insensitive, leading '/' ignored)."""
        if not part:
            return None
        return self._lower.get(part.lstrip("/").lower())

    def has(self, part: Optional[str]) -> bool:
        return self.member(part) is not None

    def read(self, part: Optional[str]) -> Optional[bytes]:
        m = self.member(part)
        if m is None or self.zip is None:
            return None
        return self.zip.read(m)

    def open_part(self, part: str):
        m = self.member(part)
        if m is None or self.zip is None:
            raise GradingError(f"part {part!r} missing from {self.file_name}", path=self.path)
        return self.zip.open(m)

    def part_size(self, part: str) -> int:
        m = self.member(part)
        return self.zip.getinfo(m).file_size if m and self.zip else 0

    def xml(self, part: Optional[str]):
        data = self.read(part)
        if data is None:
            return None
        return parse_xml(data, f"{self.file_name}:{part}")

    # ------------------------------------------------------------------ content types
    def _read_content_types(self):
        data = self.read("[Content_Types].xml")
        if data is None:
            return
        root = parse_xml(data, "[Content_Types].xml")
        for el in root:
            ln = local(el.tag)
            if ln == "Default":
                self.content_types_default[(el.get("Extension") or "").lower()] = el.get("ContentType") or ""
            elif ln == "Override":
                self.content_types_override[(el.get("PartName") or "").lstrip("/").lower()] = el.get("ContentType") or ""

    def content_type_of(self, part: str) -> Optional[str]:
        p = part.lstrip("/").lower()
        if p in self.content_types_override:
            return self.content_types_override[p]
        return self.content_types_default.get(posixpath.splitext(p)[1].lstrip("."))

    # ------------------------------------------------------------------ relationships
    @staticmethod
    def rels_part_for(part: str) -> str:
        if part in ("", "/"):
            return "_rels/.rels"
        d, f = posixpath.split(part.lstrip("/"))
        return posixpath.join(d, "_rels", f + ".rels")

    def rels(self, part: str) -> list[Rel]:
        """Relationships of a part ('' = package root), targets resolved to zip member names."""
        key = part or ""
        if key in self._rels_cache:
            return self._rels_cache[key]
        out: list[Rel] = []
        data = self.read(self.rels_part_for(key))
        if data is not None:
            root = parse_xml(data, self.rels_part_for(key))
            base = posixpath.dirname(key.lstrip("/"))
            for el in root:
                if local(el.tag) != "Relationship":
                    continue
                tgt = el.get("Target") or ""
                typ = el.get("Type") or ""
                ext = (el.get("TargetMode") or "").lower() == "external"
                resolved = None
                if not ext:
                    t = tgt.replace("\\", "/")
                    full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
                    resolved = self.member(full)
                    if resolved is None and "%" in full:
                        resolved = self.member(unquote(full))   # URL-encoded target ('My%20Sheet.xml')
                out.append(Rel(el.get("Id") or "", typ, typ.rstrip("/").rsplit("/", 1)[-1], tgt, ext, resolved))
        self._rels_cache[key] = out
        return out

    def rel_by_id(self, part: str, rid: Optional[str]) -> Optional[Rel]:
        if rid is None:
            return None
        for r in self.rels(part):
            if r.id == rid:
                return r
        return None

    def rels_of_type(self, part: str, *types: str) -> list[Rel]:
        tl = {t.lower() for t in types}
        return [r for r in self.rels(part) if r.type_short.lower() in tl]

    def reachable_parts(self) -> dict[str, set]:
        """BFS over the relationship graph from the package root: part -> set of relationship
        types it is reached by.  Parts nothing relates are not included."""
        seen: dict[str, set] = {}
        queue = [""]
        while queue:
            p = queue.pop(0)
            for r in self.rels(p):
                if r.external or r.part is None:
                    continue
                if r.part not in seen:
                    seen[r.part] = set()
                    queue.append(r.part)
                seen[r.part].add(r.type_short)
        return seen

    # ------------------------------------------------------------------ classification
    def _classify(self):
        self._read_content_types()
        main = None
        for r in self.rels(""):
            if r.type_short.lower() == "officedocument" and r.part:
                main = r.part
                break
        if main is None:
            for cand in ("xl/workbook.xml", "xl/workbook.bin"):
                if self.has(cand):
                    main = self.member(cand)
                    break
        self.main_part = main
        self.main_content_type = self.content_type_of(main) if main else None
        self.vba_file_present = any(n.lower().endswith("vbaproject.bin") for n in self.names)
        if main is None:
            if self.has("content.xml") and self.has("mimetype"):
                self.format = "ods"
                self.unsupported_reason = "OpenDocument spreadsheet, not SpreadsheetML"
            else:
                self.format = "unknown"
                self.unsupported_reason = "zip without a workbook part"
            return
        ct = (self.main_content_type or "").strip().lower()
        if main.lower().endswith(".bin") or MAIN_CT.get(ct) == "xlsb":
            self.format = "xlsb"
            self.unsupported_reason = "binary .xlsb workbook (xl/workbook.bin), not SpreadsheetML XML"
        elif ct not in MAIN_CT and ct not in GENERIC_CT:
            # e.g. a Word or PowerPoint package renamed to .xlsx: not a workbook at all
            self.format = "unknown"
            self.unsupported_reason = (f"zip package whose main part {main!r} has content type "
                                       f"{self.main_content_type!r}, not a SpreadsheetML workbook")
            return
        else:
            fmt = MAIN_CT.get(ct)
            if fmt is None:
                # no (or a generic) content type: the extension decides; _read_workbook then
                # checks that the main part really is a SpreadsheetML <workbook>
                fmt = {".xlsx": "xlsx", ".xlsm": "xlsm", ".xltx": "xltx", ".xltm": "xltm",
                       ".xlam": "xlam"}.get(self.extension, "xlsx")
            self.format = fmt
        # macro carriers Excel loads: workbook-part relationships of type vbaProject / XLM macro sheets
        for r in self.rels(main):
            t = r.type_short.lower()
            if t in ("vbaproject", "xlmacrosheet", "xlintlmacrosheet") and r.part:
                self.vba_parts.append(r.part)

    # ------------------------------------------------------------------ workbook.xml
    def _read_workbook(self):
        root = self.xml(self.main_part)
        if root is None:
            raise GradingError(f"workbook part {self.main_part!r} missing")
        ns = root.tag[1:].split("}", 1)[0] if root.tag.startswith("{") else ""
        if local(root.tag) != "workbook" or ns not in SML_MAIN_NS:
            raise GradingError(f"main part {self.main_part!r} is <{local(root.tag)}> in namespace {ns or None!r}, "
                               f"not a SpreadsheetML <workbook>")
        self.conformance = "strict" if "purl.oclc.org" in ns else "transitional"
        wb_rels = {r.id: r for r in self.rels(self.main_part)}
        for el in root:
            ln = local(el.tag)
            if ln == "sheets":
                for s in el:
                    if local(s.tag) != "sheet":
                        continue
                    rid = rel_id(s)
                    rel = wb_rels.get(rid)
                    kind = SHEET_KINDS.get(rel.type_short.lower(), "unknown") if rel else "unknown"
                    self.sheets.append(SheetInfo(
                        index=len(self.sheets), name=s.get("name") or "", sheet_id=s.get("sheetId"),
                        state=(s.get("state") or "visible"), rid=rid, part=rel.part if rel else None,
                        kind=kind, rel_type=rel.type if rel else None))
            elif ln == "definedNames":
                for d in el:
                    if local(d.tag) != "definedName":
                        continue
                    lsid = d.get("localSheetId")
                    try:
                        idx = int(lsid) if lsid not in (None, "") else None
                        bad = None
                    except ValueError:
                        idx = None
                        bad = f"localSheetId {lsid!r} is not an integer"
                    name = d.get("name") or ""
                    self.defined_names.append(DefinedName(
                        name=name, text=(d.text or ""), local_sheet_id=idx, _scope=None,
                        hidden=truthy(d.get("hidden")), builtin=name.lower().startswith("_xlnm."),
                        function=truthy(d.get("function")), vb_procedure=truthy(d.get("vbProcedure")),
                        xlm=truthy(d.get("xlm")), comment=d.get("comment"), attrs=dict(d.attrib),
                        scope_error=bad))
            elif ln == "calcPr":
                self.calc_pr = dict(el.attrib)
            elif ln == "bookViews":
                self.book_views = [dict(v.attrib) for v in el if local(v.tag) == "workbookView"]
            elif ln == "workbookPr":
                self.workbook_pr = dict(el.attrib)
            elif ln == "fileVersion":
                self.file_version = dict(el.attrib)
            elif ln == "workbookProtection":
                self.workbook_protection = {local(k): v for k, v in el.attrib.items()}
            elif ln == "externalReferences":
                n = 0
                for x in el:
                    if local(x.tag) != "externalReference":
                        continue
                    n += 1
                    self.external_links.append(self._read_external_link(n, rel_id(x), wb_rels))
            elif ln == "pivotCaches":
                for x in el:
                    if local(x.tag) == "pivotCache":
                        self.pivot_cache_refs.append((x.get("cacheId"), rel_id(x)))
        if not self.sheets:
            raise GradingError(f"workbook part {self.main_part!r} lists no sheets (Excel requires at least one)")
        for d in self.defined_names:
            if d.local_sheet_id is None:
                continue
            if 0 <= d.local_sheet_id < len(self.sheets):
                d._scope = self.sheets[d.local_sheet_id].name
            else:
                d.scope_error = (f"localSheetId {d.local_sheet_id} does not name a sheet "
                                 f"(the workbook has {len(self.sheets)})")
        self._read_doc_props()

    def _read_external_link(self, n: int, rid: Optional[str], wb_rels: dict) -> ExternalLink:
        rel = wb_rels.get(rid)
        part = rel.part if rel else None
        link = ExternalLink(index=n, rid=rid, part=part, kind=None, target=None, target_external=False)
        if part is None:
            return link
        root = self.xml(part)
        if root is None:
            return link
        part_rels = {r.id: r for r in self.rels(part)}
        for el in root:
            ln = local(el.tag)
            if ln in ("externalBook", "ddeLink", "oleLink"):
                link.kind = ln
                trid = rel_id(el)
                tr = part_rels.get(trid)
                if tr is None and part_rels:
                    tr = next(iter(part_rels.values()))
                if tr is not None:
                    link.target = tr.target
                    link.target_external = tr.external
                if ln == "ddeLink":
                    link.target = link.target or f"{el.get('ddeService')}|{el.get('ddeTopic')}"
                for sub in el.iter():
                    sl = local(sub.tag)
                    if sl == "sheetName":
                        link.sheet_names.append(sub.get("val") or "")
                    elif sl == "definedName" and sub is not el:
                        link.defined_names.append((sub.get("name"), sub.get("refersTo"), sub.get("sheetId")))
                    elif sl == "sheetData":
                        link.has_cached_data = True
        return link

    def _read_doc_props(self):
        app_part = core_part = None
        for r in self.rels(""):
            t = r.type_short.lower()
            if t in ("extended-properties", "extendedproperties") and r.part:
                app_part = r.part
            elif t in ("core-properties", "coreproperties") and r.part:
                core_part = r.part
        app_part = app_part or self.member("docProps/app.xml")
        core_part = core_part or self.member("docProps/core.xml")
        if app_part:
            try:
                root = self.xml(app_part)
            except GradingError:
                root = None
            if root is not None:
                for el in root:
                    if len(el) == 0:
                        self.app.raw[local(el.tag)] = (el.text or "").strip()
                self.app.application = self.app.raw.get("Application")
                self.app.app_version = self.app.raw.get("AppVersion")
                self.app.company = self.app.raw.get("Company")
        if core_part:
            try:
                root = self.xml(core_part)
            except GradingError:
                root = None
            if root is not None:
                for el in root:
                    self.core.raw[local(el.tag)] = (el.text or "").strip()
                self.core.creator = self.core.raw.get("creator")
                self.core.last_modified_by = self.core.raw.get("lastModifiedBy")
                self.core.created = self.core.raw.get("created")
                self.core.modified = self.core.raw.get("modified")
                self.core.title = self.core.raw.get("title")

    # ------------------------------------------------------------------ lookups
    def sheet(self, name: str) -> Optional[SheetInfo]:
        """SheetInfo by name (exact first, then case-insensitive like Excel)."""
        for s in self.sheets:
            if s.name == name:
                return s
        low = name.lower()
        for s in self.sheets:
            if s.name.lower() == low:
                return s
        return None

    @property
    def active_tab(self) -> int:
        try:
            return int((self.book_views[0] if self.book_views else {}).get("activeTab", 0))
        except ValueError:
            return 0

    @property
    def date1904(self) -> bool:
        return truthy(self.workbook_pr.get("date1904"))

    # ------------------------------------------------------------------ styles / strings
    @property
    def styles(self):
        if self._styles is None:
            from .styles import Styles
            styles_part = theme_part = None
            for r in self.rels(self.main_part or ""):
                t = r.type_short.lower()
                if t == "styles" and r.part:
                    styles_part = r.part
                elif t == "theme" and r.part:
                    theme_part = r.part
            self._styles = Styles.parse(self.read(styles_part) if styles_part else None,
                                        self.read(theme_part) if theme_part else None,
                                        where=self.file_name)
        return self._styles

    def _load_sst(self):
        strings: list[str] = []
        rich: dict[int, list] = {}
        part = None
        for r in self.rels(self.main_part or ""):
            if r.type_short.lower() == "sharedstrings" and r.part:
                part = r.part
        if part:
            from .sheet import parse_runs
            with self.open_part(part) as fh:
                try:
                    for _ev, el in ET.iterparse(fh, events=("end",)):
                        if local(el.tag) == "si":
                            txt, runs = parse_runs(el)
                            if runs is not None:
                                rich[len(strings)] = runs
                            strings.append(txt)
                            el.clear()
                except ET.ParseError as e:
                    raise GradingError(f"XML parse error in {self.file_name}:{part}: {e}") from e
        self._sst, self._sst_rich = strings, rich

    @property
    def shared_strings(self) -> list[str]:
        """Plain text of every shared string (rich runs concatenated, phonetic runs dropped)."""
        if self._sst is None:
            self._load_sst()
        return self._sst

    @property
    def shared_string_runs(self) -> dict[int, list]:
        """{sst index: [Run, ...]} for rich-text shared strings only."""
        if self._sst_rich is None:
            self._load_sst()
        return self._sst_rich

    # ------------------------------------------------------------------ per-sheet side parts
    def sheet_rels(self, si: SheetInfo) -> list[Rel]:
        return self.rels(si.part) if si.part else []

    def tables(self, si: SheetInfo) -> list[Table]:
        out = []
        for r in self.rels_of_type(si.part or "", "table"):
            if not r.part:
                continue
            root = self.xml(r.part)
            if root is None:
                continue
            cols = []
            tc = child(root, "tableColumns")
            for c in (tc if tc is not None else []):
                if local(c.tag) != "tableColumn":
                    continue
                calc = child(c, "calculatedColumnFormula")
                tot = child(c, "totalsRowFormula")
                cols.append({"name": c.get("name"), "calculated": calc.text if calc is not None else None,
                             "totals_formula": tot.text if tot is not None else None,
                             "totals_function": c.get("totalsRowFunction")})
            out.append(Table(si.name, r.part, root.get("name"), root.get("displayName"), root.get("ref"),
                             int(root.get("headerRowCount", 1) or 1), int(root.get("totalsRowCount", 0) or 0),
                             cols))
        return out

    def comments(self, si: SheetInfo) -> list[Comment]:
        """Legacy notes (comments part) and threaded comments of a sheet."""
        out: list[Comment] = []
        for r in self.rels_of_type(si.part or "", "comments"):
            root = self.xml(r.part) if r.part else None
            if root is None:
                continue
            authors = []
            a = child(root, "authors")
            for x in (a if a is not None else []):
                authors.append(x.text or "")
            cl = child(root, "commentList")
            for c in (cl if cl is not None else []):
                if local(c.tag) != "comment":
                    continue
                t = child(c, "text")
                try:
                    au = authors[int(c.get("authorId", 0))]
                except (ValueError, IndexError):
                    au = None
                out.append(Comment(c.get("ref") or "", au, text_of(t) if t is not None else ""))
        for r in self.rels_of_type(si.part or "", "threadedComment", "threadedComments"):
            root = self.xml(r.part) if r.part else None
            if root is None:
                continue
            for c in root:
                if local(c.tag) == "threadedComment":
                    t = child(c, "text")
                    out.append(Comment(c.get("ref") or "", c.get("personId"),
                                       (t.text or "") if t is not None else "", threaded=True))
        return out

    # ------------------------------------------------------------------ drawings / charts
    def _load_drawings(self):
        drawings: list[Drawing] = []
        charts: list[Chart] = []
        seen_charts = set()
        for si in self.sheets:
            if not si.part:
                continue
            for r in self.rels_of_type(si.part, "drawing"):
                if not r.part:
                    continue
                textlinks = []
                try:
                    root = self.xml(r.part)
                except GradingError:
                    root = None
                if root is not None:
                    for el in root.iter():
                        tl = el.get("textlink")
                        if tl:
                            textlinks.append(tl)
                cps = []
                for cr in self.rels_of_type(r.part, "chart", "chartEx"):
                    if not cr.part:
                        continue
                    cps.append(cr.part)
                    if cr.part in seen_charts:
                        continue
                    seen_charts.add(cr.part)
                    croot = self.xml(cr.part)
                    fs = [el.text for el in croot.iter() if local(el.tag) == "f" and el.text] if croot is not None else []
                    charts.append(Chart(si.name, cr.part, fs))
                drawings.append(Drawing(si.name, r.part, cps, textlinks))
        self._drawings, self._charts = drawings, charts

    @property
    def charts(self) -> list[Chart]:
        if self._charts is None:
            self._load_drawings()
        return self._charts

    @property
    def drawings(self) -> list[Drawing]:
        if self._drawings is None:
            self._load_drawings()
        return self._drawings

    @property
    def pivot_caches(self) -> list[PivotCache]:
        if self._pivots is None:
            out = []
            wb_rels = {r.id: r for r in self.rels(self.main_part or "")}
            for cache_id, rid in self.pivot_cache_refs:
                rel = wb_rels.get(rid)
                pc = PivotCache(cache_id, rel.part if rel else None, None, None, None, None, None)
                root = self.xml(rel.part) if rel and rel.part else None
                if root is not None:
                    cs = child(root, "cacheSource")
                    if cs is not None:
                        pc.source_type = cs.get("type")
                        ws = child(cs, "worksheetSource")
                        if ws is not None:
                            pc.source_ref, pc.source_sheet, pc.source_name = ws.get("ref"), ws.get("sheet"), ws.get("name")
                            srid = rel_id(ws)
                            if srid:
                                sr = self.rel_by_id(rel.part, srid)
                                pc.source_target = sr.target if sr else None
                out.append(pc)
            self._pivots = out
        return self._pivots


# ----------------------------------------------------------------------------- OLE2 directory
def ole2_stream_names(path: str) -> Optional[list[str]]:
    """Entry names of a Compound File Binary (OLE2) container (MS-CFB), stdlib only.
    None when the structure cannot be parsed."""
    try:
        with open(path, "rb") as fh:
            data = fh.read()
        if data[:8] != OLE2_MAGIC:
            return None
        ssz = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        n_fat = struct.unpack_from("<I", data, 0x2C)[0]
        dir_start = struct.unpack_from("<I", data, 0x30)[0]
        difat_start, n_difat = struct.unpack_from("<II", data, 0x44)
        difat = list(struct.unpack_from("<109I", data, 0x4C))

        def sec(i):
            off = ssz * (i + 1)
            return data[off:off + ssz]

        nxt, seen = difat_start, 0
        while nxt < 0xFFFFFFFA and seen < n_difat:
            vals = struct.unpack_from(f"<{ssz // 4}I", sec(nxt))
            difat += list(vals[:-1])
            nxt = vals[-1]
            seen += 1
        fat: list[int] = []
        for fs in difat[:n_fat]:
            if fs >= 0xFFFFFFFA:
                continue
            b = sec(fs)
            fat += list(struct.unpack_from(f"<{len(b) // 4}I", b))
        names, s, hops = [], dir_start, 0
        while s < 0xFFFFFFFA and hops < 100000:
            b = sec(s)
            for off in range(0, len(b) - 127, 128):
                nlen = struct.unpack_from("<H", b, off + 0x40)[0]
                typ = b[off + 0x42]
                if typ in (1, 2, 5) and 2 <= nlen <= 64:
                    names.append(b[off:off + nlen - 2].decode("utf-16-le", "ignore"))
            s = fat[s] if s < len(fat) else 0xFFFFFFFE
            hops += 1
        return names
    except Exception:  # noqa: BLE001 - a malformed container just yields no names
        return None
