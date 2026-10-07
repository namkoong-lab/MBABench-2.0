"""77 Formatting/File extension (.xlsx).

Rule as implemented (handoff bucket 1 row 77; see docs/checks/77.md):

  PASS  the delivered file name ends .xlsx (any letter case) and the file really is a
        macro-free .xlsx workbook package (main content type "...spreadsheetml.sheet.main+xml",
        no macro part loaded); or
        the name ends .xlsm and the file really is a macro-enabled workbook package
        ("...ms-excel.sheet.macroEnabled.main+xml") that loads at least one macro part.
  FAIL  everything else: .xls, .xlsb, .csv, .ods, templates/add-ins (.xltx, .xltm, .xlam),
        no extension; an .xlsm without macros; macros in an .xlsx; a name whose extension
        does not match the package (legacy .xls / .xlsb / CSV / ODS bytes or a macro-enabled
        package named .xlsx, a plain .xlsx package named .xlsm) - Excel refuses to open a
        file whose extension and format disagree.

Delivered name: production stages every delivery as ai_attempt.xlsx, so the ORIGINAL name
comes from task_meta["delivered_filename"] when given (else the file's own name);
stats["filename_source"] says which.  A file that carries the staging name itself
(STAGED_FILE_NAMES) without a delivered name raises GradingError: its own name is not the
delivered one (judge guidance: "never from the staged name"), so grading it would be a guess.

"Macro part" = a part Excel loads as macros: a VBA project (vbaProject relationship from the
workbook part, the part present and a real VBA project: an OLE2 compound file whose root
holds a "VBA" storage with a "dir" stream and a "PROJECT" stream, MS-OVBA 2.2) or an Excel
4.0 macro sheet (xlMacrosheet / xlIntlMacrosheet relationship).  A *vbaProject.bin nothing
relates is not loaded (OPC: parts are reached through relationships) and does not count; it
is listed in stats.  A related part that is not a compound file, has a broken compound-file
header, or lacks the VBA storage/streams is not a VBA project (invalid_vba_parts).  An empty
VBA project (structure present, no modules) counts as present (no source-code decoding).

This is the only check that grades files the reader cannot parse (accepts_unparsed).
No guess (GradingError) when the decision needs what cannot be known: the staging name
without a delivered name; an empty, corrupt or password-encrypted file delivered as
.xlsx/.xlsm; a package whose main content type is generic (application/xml ...) under an
.xlsx/.xlsm name; a related vbaProject part whose compound-file directory cannot be walked
being the only possible macro carrier; an Excel 5.0 dialog sheet being the only possible
macro carrier (whether it needs .xlsm is unverified).  A name that is not .xlsx/.xlsm fails
on the name alone, whatever the bytes.
"""
from __future__ import annotations

import struct

from ..core.package import MAIN_CT
from ..errors import GradingError
from .base import Check

ALLOWED = (".xlsx", ".xlsm")
OLE2_MAGIC = bytes.fromhex("d0cf11e0a1b11ae1")
TASK_META_KEY = "delivered_filename"
PROBLEM_META_KEY = "delivered_filename_problem"   # why the caller has no delivered name (e.g. "missing", "malformed: ...")
SIDECAR_KEY = "original_filename"     # the key production's _attempt_origin.json uses (named in the error only)
STAGED_FILE_NAMES = ("ai_attempt.xlsx",)   # grade.py stages every delivery under this name
DIALOG_SHEET_POLICY = "error"         # Excel 5.0 dialog sheet as the only macro-like part: "error" (GradingError,
                                      # unverified) | "macro" (counts as a macro part) | "not_macro" (ignored)

KIND_LABEL = {
    "xlsx": "a macro-free Excel workbook package (.xlsx content)",
    "xlsm": "a macro-enabled Excel workbook package (.xlsm content)",
    "xltx": "an Excel template package (.xltx content)",
    "xltm": "a macro-enabled Excel template package (.xltm content)",
    "xlam": "an Excel add-in package (.xlam content)",
    "xlsb": "a binary Excel workbook (.xlsb: xl/workbook.bin, not XML)",
}


def delivered_name(wb) -> tuple:
    """(original delivered file name, where it came from) - or (None, why) when the delivered name is unknown: a
    malformed task_meta name, or the staging name without a delivered name (Patrick 2026-10-05: every attempt
    graded - the format is then judged from the file's content, C77.finish)."""
    meta = getattr(wb, "task_meta", None) or {}
    if TASK_META_KEY in meta and meta[TASK_META_KEY] is not None:
        v = meta[TASK_META_KEY]
        if not isinstance(v, str) or not v.strip():
            return None, f"malformed: task_meta[{TASK_META_KEY!r}] is {v!r}, not a file name"
        name = v.strip().replace("\\", "/").rsplit("/", 1)[-1]
        if not name:
            return None, f"malformed: task_meta[{TASK_META_KEY!r}]={v!r} names a folder, not a file"
        return name, f"task_meta.{TASK_META_KEY}"
    if wb.file_name.lower() in STAGED_FILE_NAMES:
        why = str(meta.get(PROBLEM_META_KEY) or "missing")
        if SIDECAR_KEY in meta:
            why += f" (task_meta has {SIDECAR_KEY!r} but not {TASK_META_KEY!r})"
        return None, why
    return wb.file_name, "file_name"


def extension_of(name: str) -> str:
    """Lower-case extension including the dot ('' when none); '.xlsx' for a file named '.xlsx'."""
    k = name.rfind(".")
    return name[k:].lower() if k >= 0 and k < len(name) - 1 else ""


# ----------------------------------------------------------------------------- VBA project
class _Unreadable(Exception):
    pass


def _cfb_directory(data: bytes) -> list[tuple[str, int, int, int, int]]:
    """Directory entries (name, type, left, right, child) by entry index of a Compound File
    Binary (MS-CFB) with a valid header.  Raises _Unreadable when the directory cannot be
    walked (sectors out of range, a chain loop)."""
    ssz = 1 << struct.unpack_from("<H", data, 0x1E)[0]
    n_fat = struct.unpack_from("<I", data, 0x2C)[0]
    dir_start = struct.unpack_from("<I", data, 0x30)[0]
    difat_start, n_difat = struct.unpack_from("<II", data, 0x44)
    difat = list(struct.unpack_from("<109I", data, 0x4C))
    n_sectors = -(-(len(data) - ssz) // ssz)       # a short last sector is padded (FREESECT bytes)

    def sec(i: int) -> bytes:
        if i >= n_sectors:
            raise _Unreadable(f"sector {i} is beyond the end of the file")
        return data[ssz * (i + 1):ssz * (i + 2)].ljust(ssz, b"\xff")

    nxt, seen = difat_start, set()
    while nxt < 0xFFFFFFFA and len(seen) < n_difat:
        if nxt in seen:
            raise _Unreadable("DIFAT chain loops")
        seen.add(nxt)
        vals = struct.unpack_from(f"<{ssz // 4}I", sec(nxt))
        difat += list(vals[:-1])
        nxt = vals[-1]
    fat: list[int] = []
    for fs in difat[:n_fat]:
        if fs < 0xFFFFFFFA:
            fat += list(struct.unpack_from(f"<{ssz // 4}I", sec(fs)))
    entries, s, visited = [], dir_start, set()
    while s < 0xFFFFFFFA:
        if s in visited:
            raise _Unreadable("directory chain loops")
        visited.add(s)
        b = sec(s)
        for off in range(0, ssz, 128):
            nlen = struct.unpack_from("<H", b, off + 0x40)[0]
            typ = b[off + 0x42]
            name = b[off:off + max(0, min(nlen, 64) - 2)].decode("utf-16-le", "replace")
            left, right, child = struct.unpack_from("<III", b, off + 0x44)
            entries.append((name, typ, left, right, child))
        if s >= len(fat):
            raise _Unreadable(f"directory sector {s} has no FAT entry")
        s = fat[s]
    if not entries or entries[0][1] != 5:
        raise _Unreadable("the first directory entry is not the root storage")
    return entries


def _children(entries, idx: int) -> dict[str, tuple[int, int]]:
    """{upper-case name: (entry index, type)} of the direct children of storage `idx`
    (the red-black tree under its child pointer)."""
    out, stack, seen = {}, [entries[idx][4]], set()
    while stack:
        i = stack.pop()
        if i >= 0xFFFFFFFA:
            continue
        if i >= len(entries) or i in seen:
            raise _Unreadable(f"directory entry {i} is out of range or reached twice")
        seen.add(i)
        name, typ, left, right, _child = entries[i]
        if typ in (1, 2):
            out[name.upper()] = (i, typ)
        stack += [left, right]
    return out


def vba_project_status(data: bytes | None) -> tuple[str, str]:
    """('vba' | 'invalid' | 'unreadable', why) for the bytes of a related vbaProject part.
    'vba': an OLE2 compound file whose root holds a VBA storage with a dir stream and a
    PROJECT stream (MS-OVBA 2.2.1/2.3.1: all three are required).  'invalid': definitely not
    a VBA project Excel can load (not OLE2, a header that breaks MS-CFB's MUST rules, or the
    VBA structure missing).  'unreadable': a valid header whose directory cannot be walked."""
    if not data or data[:8] != OLE2_MAGIC:
        return "invalid", "not an OLE2 compound file"
    if len(data) < 512:
        return "invalid", "a compound file shorter than its 512-byte header"
    bom, shift = struct.unpack_from("<HH", data, 0x1C)
    if bom != 0xFFFE or shift not in (9, 12):
        return "invalid", "a damaged compound-file header (byte order / sector size)"
    try:
        entries = _cfb_directory(data)
        root = _children(entries, 0)
        vba = root.get("VBA")
        if vba is None or vba[1] != 1:
            return "invalid", "a compound file without a VBA storage (not a VBA project)"
        if root.get("PROJECT", (0, 0))[1] != 2:
            return "invalid", "a compound file without a PROJECT stream (not a VBA project)"
        if _children(entries, vba[0]).get("DIR", (0, 0))[1] != 2:
            return "invalid", "a compound file whose VBA storage has no dir stream (not a VBA project)"
    except _Unreadable as e:
        return "unreadable", f"a compound file whose directory cannot be read ({e})"
    except struct.error as e:
        return "unreadable", f"a compound file whose directory cannot be read (truncated: {e})"
    return "vba", "a VBA project"


def ole2_kind(streams) -> str:
    s = {x.lower() for x in (streams or [])}
    if "workbook" in s or "book" in s:
        return "a legacy Excel 97-2003 binary workbook (.xls, BIFF in an OLE2 compound file)"
    if streams is None:
        return "an OLE2 compound file whose directory cannot be read (not an Excel workbook package)"
    return "an OLE2 compound file without an Excel workbook stream (not an Excel workbook package)"


class C77(Check):
    number = 77
    key = "Formatting/File extension (.xlsx)"
    accepts_unparsed = True
    sheet_kinds = ()

    def wants_sheet(self, info) -> bool:
        return False                       # package level only

    # ------------------------------------------------------------------ package facts
    def _macro_facts(self) -> dict:
        wb = self.wb
        real, invalid, unreadable = [], {}, {}
        if wb.container != "zip" or wb.format == "corrupt" or not wb.main_part:
            return {"macro_parts": [], "invalid_vba_parts": {}, "unreadable_vba_parts": {},
                    "missing_vba_parts": [], "unrelated_vba_files": [], "dialog_sheets": []}
        for part in wb.vba_parts:
            ct_rel = [r.type_short.lower() for r in wb.rels(wb.main_part) if r.part == part]
            if "vbaproject" in ct_rel:
                status, why = vba_project_status(wb.read(part))
                if status == "vba":
                    real.append(part)
                else:
                    (invalid if status == "invalid" else unreadable)[part] = why
            else:
                real.append(part)          # an Excel 4.0 (XLM) macro sheet related from the workbook
        related = {p.lower() for p in wb.vba_parts}
        orphans = [n for n in wb.names if n.lower().endswith("vbaproject.bin") and n.lower() not in related]
        dialogs = [r.part or r.target for r in wb.rels(wb.main_part) if r.type_short.lower() == "dialogsheet"]
        missing = [r.target for r in wb.rels(wb.main_part)
                   if r.type_short.lower() in ("vbaproject", "xlmacrosheet", "xlintlmacrosheet") and r.part is None]
        if DIALOG_SHEET_POLICY == "macro":
            real += [d for d in dialogs if d not in real]
        elif DIALOG_SHEET_POLICY not in ("error", "not_macro"):
            raise GradingError(f"{self.key}: unknown DIALOG_SHEET_POLICY {DIALOG_SHEET_POLICY!r}")
        return {"macro_parts": real, "invalid_vba_parts": invalid, "unreadable_vba_parts": unreadable,
                "missing_vba_parts": missing, "unrelated_vba_files": orphans, "dialog_sheets": dialogs}

    def _content(self) -> tuple[str | None, str]:
        """(package kind or None, description).  Kind is one of KIND_LABEL's keys, or a
        non-workbook marker ('xls', 'ole2', 'ods', 'csv', 'zip', 'other'); None = unknown."""
        wb = self.wb
        if wb.container == "ole2":
            if wb.format == "encrypted":
                return None, "a password-encrypted package (OLE2 EncryptedPackage) whose content cannot be read"
            k = ole2_kind(wb.ole_streams)
            if wb.ole_streams is None:      # could be a damaged encrypted .xlsx as well as an .xls
                return None, k
            return ("xls" if "BIFF" in k else "ole2"), k
        if wb.container == "empty":
            return None, "an empty file"
        if wb.container == "other":
            what = {"csv": "a text/CSV file", "ods": "an OpenDocument file"}.get(wb.format, "not a zip or OLE2 file")
            return ("csv" if wb.format == "csv" else "other"), f"{what} ({wb.unsupported_reason})"
        if wb.format == "corrupt":
            return None, f"a damaged zip package ({wb.unsupported_reason})"
        if wb.format == "ods":
            return "ods", "an OpenDocument spreadsheet (.ods), not an Excel workbook"
        if wb.format == "xlsb":
            return "xlsb", KIND_LABEL["xlsb"]
        if wb.format == "unknown":
            return "zip", f"a zip package that is not an Excel workbook ({wb.unsupported_reason})"
        ct = (wb.main_content_type or "").strip().lower()
        kind = MAIN_CT.get(ct)
        if kind is None:
            return None, (f"a package whose workbook part has the generic content type {wb.main_content_type!r}, "
                          f"so its type (.xlsx or .xlsm) is not declared")
        return kind, KIND_LABEL[kind]

    def _content_extension(self, kind, macros: bool) -> str:
        """The extension the file's content implies when the delivered name is unknown (Patrick 2026-10-05): '' when
        the content is no Excel workbook (an empty or damaged file, a zip that is not a workbook)."""
        wb = self.wb
        if wb.container == "ole2":
            return ".xls"                              # OLE2 = .xls (an encrypted package is OLE2 too)
        if kind in ("xlsb", "xltx", "xltm", "xlam", "ods", "csv"):
            return "." + kind
        if wb.container == "zip" and wb.format != "corrupt" and kind in ("xlsx", "xlsm", None) and wb.main_part:
            return ".xlsm" if macros else ".xlsx"
        return ""

    # ------------------------------------------------------------------ verdict
    def finish(self) -> dict:
        wb = self.wb
        if wb is None:
            raise GradingError(f"{self.key}: no workbook")
        name, source = delivered_name(wb)
        kind, what = self._content()
        facts = self._macro_facts()
        macros = bool(facts["macro_parts"])
        unknown = None
        if name is None:
            # Patrick 2026-10-05 (every attempt graded): the delivered name is unknown (no or a malformed
            # _attempt_origin.json) - the format is judged from the file's content: a zip workbook without a VBA
            # project = .xlsx, with a real one = .xlsm, OLE2 = .xls, an .xlsb zip = .xlsb
            unknown = source
            ext = self._content_extension(kind, macros)
            if kind is None and ext in ALLOWED:
                kind = ext[1:]                         # a workbook package that does not declare its type
            name = f"{wb.file_name} (delivered name unknown)"
            source = f"content (delivered name unknown: {unknown})"
            self.note_default("delivered_name_unknown", f"{wb.file_name}: {unknown}; judged from its content as "
                                                        f"{ext or 'no Excel workbook'}")
        else:
            ext = extension_of(name)
        stats = {"delivered_filename": None if unknown else name, "filename_source": source,
                 "staged_file_name": wb.file_name,
                 "extension": ext, "container": wb.container, "reader_format": wb.format,
                 "package_kind": kind, "package": what, "main_part": wb.main_part,
                 "main_content_type": wb.main_content_type, "has_macros": macros, **facts,
                 "ole2_streams": (wb.ole_streams or [])[:12] if wb.container == "ole2" else None,
                 "application": getattr(getattr(wb, "app", None), "application", None),
                 "unsupported_reason": wb.unsupported_reason, "size": wb.size}

        def fail(desc: str):
            self.add_mistake(name, desc, severity="major")
            return self.verdict("unreachable", desc, stats)

        if unknown is not None:
            stats["delivered_name_unknown"] = unknown
            stats["judged_from_content_as"] = ext or None
            if ext not in ALLOWED:
                return fail(f"Delivered file name unknown ({unknown}); judged from its content, the file is {what}: "
                            f"not a standard .xlsx workbook (or .xlsm with macros).")
        if ext not in ALLOWED:
            shown = ext or "no extension"
            return fail(f"Delivered as '{name}' ({shown}): not a standard .xlsx workbook (or .xlsm with macros). "
                        f"The file is {what}.")
        if kind is None:
            raise GradingError(f"{self.key}: '{name}' is named {ext} but is {what}; its format cannot be verified")
        target = ext[1:]                                   # "xlsx" | "xlsm"
        if kind != target:
            if kind in ("xlsx", "xlsm"):
                why = ("the name says .xlsm but the package is a plain macro-free .xlsx workbook"
                       if target == "xlsm" else
                       "the name says .xlsx but the package is declared macro-enabled (.xlsm content)")
            else:
                why = f"the name says {ext} but the file is {what}"
            return fail(f"Delivered as '{name}': {why}. Extension and file format disagree (Excel refuses "
                        f"to open such a file); save it as a real {ext} workbook.")
        if facts["unreadable_vba_parts"] and not macros:
            part, why = next(iter(facts["unreadable_vba_parts"].items()))
            raise GradingError(f"{self.key}: '{name}' ({ext}) relates the VBA part {part}, {why}; whether it "
                               f"holds a VBA project (macros) cannot be known")
        undecided_dialogs = facts["dialog_sheets"] and not macros and DIALOG_SHEET_POLICY == "error"
        if undecided_dialogs:
            raise GradingError(f"{self.key}: '{name}' ({ext}) has no VBA project or macro sheet but has Excel 5.0 "
                               f"dialog sheet(s) {facts['dialog_sheets'][:3]}; whether dialog sheets need .xlsm "
                               f"is unverified (DIALOG_SHEET_POLICY)")
        if target == "xlsx":
            if macros:
                return fail(f"Delivered as '{name}' but the workbook carries macros "
                            f"({', '.join(facts['macro_parts'][:3])}); a workbook with macros must be .xlsm.")
            return self.verdict(f"Delivered as '{name}': a standard macro-free .xlsx workbook"
                                + (f" ({source})." if source != "file_name" else "."), "unreachable", stats)
        # .xlsm
        if macros:
            return self.verdict(f"Delivered as '{name}': a macro-enabled .xlsm workbook with macros "
                                f"({', '.join(facts['macro_parts'][:3])}).", "unreachable", stats)
        extra = []
        if facts["missing_vba_parts"]:
            extra.append(f"its macro relationship points to a missing part {facts['missing_vba_parts'][0]}")
        if facts["invalid_vba_parts"]:
            part, why = next(iter(facts["invalid_vba_parts"].items()))
            extra.append(f"its VBA part {part} is not a VBA project ({why})")
        if facts["unrelated_vba_files"]:
            extra.append(f"{facts['unrelated_vba_files'][0]} is in the zip but not attached to the workbook, so "
                         f"Excel does not load it")
        return fail(f"Delivered as '{name}' (macro-enabled .xlsm) but the workbook contains no VBA project or "
                    f"macro sheet" + (f" ({'; '.join(extra)})" if extra else "")
                    + "; a workbook without macros should be saved as .xlsx.")
