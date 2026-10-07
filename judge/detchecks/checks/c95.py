"""95 Potential Dangers/No external links.

Rule as implemented (see docs/checks/95.md):
fail when the workbook references another workbook (or another external data source)
anywhere:
  * every external link the workbook declares (<externalReferences> -> externalLink part:
    externalBook, DDE or OLE link), whether or not a formula uses it;
  * every formula site whose text references another workbook - [1]Sheet!A1,
    '[1]My Sheet'!A1, [1]!Name, [Book.xlsx]Sheet!A1, 'C:\\dir\\[Book.xlsx]Sheet'!A1,
    'C:\\dir\\Book.xlsx'!Name, Book.xlsx!Name (when no sheet has that name), [1]!MyFn(..):
    cell formulas (shared children of a flagged master included), defined names (used or
    not, hidden included), conditional formats and data validations (main and x14),
    chart series, sparklines, shape text links and form-control links (ctrlProps); CF
    formulas include colour-scale / data-bar / icon-set thresholds (cfvo);
  * pivot caches fed from another file, and live data connections (xl/connections.xml)
    that are not the workbook's own Data Model / worksheet ranges and not deleted.
Not links: hyperlinks (web or file), text that names a file ("Q3_comps.xlsx"), string
literals ("[Book.xlsx]Sheet1!A1"), [0]!Name (this workbook), unreferenced externalLink
parts in the zip (stats only).
Exception (rubric "unless required by the case context"): task_meta
{'requires_external_links': True} turns every finding into stats and passes; absent or
False = not required.  Power Query connections cannot be decoded: if one exists and nothing
else fails the file, the check raises GradingError (no guess).
"""
from __future__ import annotations

import re
from collections import Counter

from ..core import formula as F
from ..core.package import local, truthy
from ..errors import GradingError
from ._fscan import (FormulaScanCheck, ctrl_formulas, name_label, name_location, parse_text,
                     plural, short, sqref_location)

META_KEY = "requires_external_links"
COUNT_DATA_CONNECTIONS = True        # live ODBC/OLEDB/text/web connections make the file not self-contained
COUNT_FILE_HYPERLINKS = False        # handoff: hyperlinks are not links (stats only)
_RE_BOOKISH = re.compile(r"\.\w{1,5}'?!")
_RE_SHEETFILE = re.compile(r"\.(?:xl[a-z]{1,2}|csv|ods)(?:[#?].*)?$", re.I)


def may_be_external(text: str) -> bool:
    """Conservative pre-filter: False guarantees no reference to another workbook (string
    literals masked).  External forms all carry '[', a path separator, '://', or a file
    name right before '!' (Book.xlsx!Name / 'C:\\p\\Book.xlsx'!Name)."""
    if not text or ("[" not in text and "!" not in text):
        return False
    m = F.mask_strings(text)
    return "[" in m or "\\" in m or "://" in m or bool(_RE_BOOKISH.search(m))


class C95(FormulaScanCheck):
    number = 95
    key = "Potential Dangers/No external links"

    def start(self, wb):
        super().start(wb)
        meta = (wb.task_meta or {}).get(META_KEY)
        if meta is not None and not isinstance(meta, bool):
            raise GradingError(f"{self.key}: task_meta[{META_KEY!r}] must be true/false, got {meta!r}")
        self.required = bool(meta)
        self._sheets_low = {s.lower() for s in self.sheet_names}
        self._links = {}
        for link in wb.external_links:
            if link.part is None or link.kind is None:
                raise GradingError(f"{self.key}: external reference {link.index} (r:id {link.rid}) has no readable "
                                   f"externalLink part; cannot tell what it links to")
            self._links[str(link.index)] = link
        self._site_uses = Counter()       # target label -> formula sites using it
        self.n_unknown_file_sheets = 0
        self.unknown_file_sheet_examples: list = []
        self._name_items = []
        flagged = []
        for i, dn in enumerate(wb.defined_names):
            self.site = name_location(dn)
            r = self.safe_classify(dn.text, None)
            if r is not None:
                flagged.append(i)
                self._name_items.append((i, dn, r))
        self.set_flagged_names(flagged)
        self.n_hyperlinks_to_files = 0
        self.hyperlink_examples: list = []
        self.power_query: list = []
        self.deleted_connections = 0
        self.internal_connections = 0

    # ------------------------------------------------------------------ rule
    def _label(self, ext: str) -> str:
        if ext.isdigit():
            link = self._links.get(ext)
            return f"[{ext}] {link.target}" if link is not None else f"[{ext}] (no external link with that number)"
        return ext

    def classify(self, text, sheet):
        if not may_be_external(text):
            return None
        f = parse_text(text)
        targets, raws = [], []
        for o in f.operands:
            # Other.xlsx!A1 with no sheet 'Other.xlsx': the parser reads it as a (missing) sheet of this
            # workbook, not as another workbook - stats only, question for Patrick
            if o.is_reference and o.external is None and o.sheet is not None and o.sheet_end is None \
                    and _RE_SHEETFILE.search(o.sheet) and o.sheet.lower() not in self._sheets_low:
                self.n_unknown_file_sheets += 1
                if len(self.unknown_file_sheet_examples) < 5:
                    self.unknown_file_sheet_examples.append(f"{self.site}: {o.raw}")
        for o in f.external_references(sheet_names=self.sheet_names):
            targets.append(self._label(o.external if o.external is not None else o.sheet))
            raws.append(o.raw)
        for c in f.functions:
            if c.qualifier:
                q = F.parse_qualifier(c.qualifier)
                ext = q.external if q.external is not None else (
                    q.sheet if q.ambiguous_book and q.sheet.lower() not in self._sheets_low else None)
                if ext is not None:
                    targets.append(self._label(ext))
                    raws.append(f"{c.qualifier}!{c.raw[len(c.prefix):]}(")
        if not targets:
            return None
        key = frozenset(targets)
        for t in key:
            self._site_uses[t] += 1
        return key, tuple(dict.fromkeys(raws))

    def record(self, cell, result):
        if cell.formula is not None and not cell.formula.text:      # shared child: count its use too
            for t in result[0]:
                self._site_uses[t] += 1
        super().record(cell, result)

    def describe_cells(self, sheet, rng, n, targets, raws):
        cells = "Formula" if n == 1 else f"{n} formulas"
        return (f"{cells} in {rng} reference{'s' if n == 1 else ''} another workbook: "
                f"{'; '.join(sorted(targets))} (e.g. {', '.join(raws[:2])}).")

    def describe_rule(self, what, sheet, sqref, keys, texts):
        targets = sorted({t for k, _ in keys for t in k})
        return f"{what} on {sqref} references another workbook: {'; '.join(targets)} (={short(texts[0], 100)})."

    def sheet_end(self, head, tail):
        super().sheet_end(head, tail)
        for g in tail.sparkline_groups:
            hosts = [sp.sqref for sp in g.sparklines if sp.sqref]
            for t, host in ([(g.date_range, " ".join(hosts))] if g.date_range else []) + \
                    [(sp.formula, sp.sqref) for sp in g.sparklines if sp.formula]:
                self.site = f"sparkline on '{head.name}'"
                r = self.classify(t, head.name)
                if r is not None:
                    self._sheet_items.append((sqref_location(head.name, host) if host else head.name,
                                              f"Sparkline(s) on {','.join(host.split()) or '?'} read another "
                                              f"workbook: {'; '.join(sorted(r[0]))} (={short(t, 100)})."))
        for h in tail.hyperlinks:
            tgt = (h.target or "").split("#", 1)[0]
            if h.external and _RE_SHEETFILE.search(tgt):
                self.n_hyperlinks_to_files += 1
                if len(self.hyperlink_examples) < 5:
                    self.hyperlink_examples.append(f"'{head.name}'!{h.ref} -> {h.target}")
                if COUNT_FILE_HYPERLINKS:
                    self._sheet_items.append((sqref_location(head.name, h.ref),
                                              f"Hyperlink in {h.ref} opens another workbook file: {h.target}."))

    # ------------------------------------------------------------------ workbook-level places
    def _workbook_places(self):
        wb = self.wb
        items = []
        for i, dn, (targets, raws) in self._name_items:
            items.append((name_location(dn),
                          f"The {name_label(dn)} refers to another workbook: {'; '.join(sorted(targets))} "
                          f"(={short(dn.text, 100)}); {self.usage_text(i)}."))
        for ch in wb.charts:
            for t in ch.formulas:
                self.site = f"chart {ch.part}"
                r = self.classify(t, ch.sheet)
                if r is not None:
                    items.append((ch.sheet or f"Chart: {ch.part}",
                                  f"A chart on '{ch.sheet}' ({ch.part}) reads another workbook: "
                                  f"{'; '.join(sorted(r[0]))} (={short(t, 100)})."))
            for rel in wb.rels(ch.part):
                if rel.external and rel.type_short.lower() != "hyperlink":
                    items.append((ch.sheet or f"Chart: {ch.part}",
                                  f"A chart on '{ch.sheet}' ({ch.part}) takes its data from an "
                                  f"external file {rel.target} ({rel.type_short})."))
        for dr in wb.drawings:
            for t in dr.textlinks:
                self.site = f"shape text link on '{dr.sheet}'"
                r = self.classify(t, dr.sheet)
                if r is not None:
                    items.append((dr.sheet, f"A shape on '{dr.sheet}' ({dr.part}) shows a value from another "
                                            f"workbook: {'; '.join(sorted(r[0]))} (={short(t, 100)})."))
        for sheet, part, a, t in ctrl_formulas(wb):
            self.site = f"form control on '{sheet}'"
            r = self.classify(t, sheet)
            if r is not None:
                items.append((sheet, f"A form control on '{sheet}' ({part}, {a}) links to another workbook: "
                                     f"{'; '.join(sorted(r[0]))} (={short(t, 100)})."))
        for pc in wb.pivot_caches:
            ext = [r for r in (wb.rels(pc.part) if pc.part else []) if r.external]
            if pc.source_target or ext:
                tgt = pc.source_target or ext[0].target
                items.append((f"Pivot cache: {pc.part}", f"A pivot cache ({pc.part}) is fed from another file: {tgt}."))
        if COUNT_DATA_CONNECTIONS:
            items.extend(self._connections())
        # declared links last, so their use counts include every site classified above
        links = []
        for k, link in self._links.items():
            n = self._site_uses.get(self._label(k), 0)
            sheets = f"; cached sheets {', '.join(repr(s) for s in link.sheet_names[:5])}" if link.sheet_names else ""
            kind = {"externalBook": "workbook", "ddeLink": "DDE", "oleLink": "OLE"}.get(link.kind, link.kind)
            used = f"used by {plural(n, 'formula site')}" if n else "no formula uses it (Excel still lists it under Edit Links)"
            links.append((f"External link [{k}]",
                          f"The workbook declares an external {kind} link [{k}] to {link.target}{sheets} "
                          f"({link.part}); {used}."))
        return links + items

    def _connections(self):
        wb = self.wb
        out = []
        parts = [r.part for r in wb.rels_of_type(wb.main_part or "", "connections") if r.part]
        for part in parts:
            root = wb.xml(part)
            if root is None:
                continue
            for el in root:
                if local(el.tag) != "connection":
                    continue
                name = el.get("name") or ""
                subs = list(el.iter())
                src, cmd = None, ""
                for s in subs:
                    ln = local(s.tag)
                    if ln == "dbPr":
                        src = s.get("connection") or src
                        cmd += " " + (s.get("command") or "")
                    elif ln == "textPr":
                        src = s.get("sourceFile") or src
                    elif ln == "webPr":
                        src = s.get("url") or src
                blob = " ".join(x for x in (name, src or "", cmd, el.get("description") or "") if x)
                if truthy(el.get("deleted")):
                    self.deleted_connections += 1
                elif ("ThisWorkbookDataModel" in blob or "$Embedded$" in blob or name.startswith("WorksheetConnection_")
                      or any(local(s.tag) == "rangePr" for s in subs)):
                    self.internal_connections += 1
                elif "Microsoft.Mashup.OleDb" in blob or "$Workbook$" in blob:
                    self.power_query.append(name or "?")
                else:
                    out.append((f"Data connection: {name or '?'}",
                                f"The workbook keeps a live data connection '{name}' (type {el.get('type')}) to "
                                f"{short(src or 'an external source', 120)}; the file is not self-contained."))
        for n in wb.names:
            nl = n.lower()
            if nl.startswith("customxml/item") and nl.endswith(".xml") and "props" not in nl:
                data = wb.read(n) or b""
                if b"DataMashup" in data[:4000] and not self.power_query:
                    self.power_query.append(f"DataMashup in {n}")
        return out

    # ------------------------------------------------------------------ verdict
    def finish(self):
        wb_items = self._workbook_places()
        self._wb_items[:0] = wb_items
        found = len(self._wb_items) + len(self._sheet_items)
        linked_parts = {l.part.lower() for l in self._links.values()}
        orphans = [n for n in self.wb.names if n.lower().startswith("xl/externallinks/externallink")
                   and n.lower().endswith(".xml") and n.lower() not in linked_parts]
        if self.power_query and not found and not self.required:
            raise GradingError(f"{self.key}: Power Query present ({', '.join(self.power_query[:3])}); its sources "
                               f"live in M code that this check does not decode, so it cannot tell whether the "
                               f"workbook reads another file")
        stats = {"n_formula_cells": self.n_formula_cells, "n_formula_texts": self.n_formula_texts,
                 "n_flagged_cells": self.n_flagged_cells, "n_external_links": len(self._links),
                 "n_names": len(self.wb.defined_names), "n_names_flagged": len(self._name_items),
                 "n_cf_rules": self.n_cf_rules, "n_data_validations": self.n_dv,
                 "requires_external_links": self.required, "orphan_external_link_parts": orphans,
                 "hyperlinks_to_files_not_counted": self.n_hyperlinks_to_files,
                 "hyperlink_examples": self.hyperlink_examples,
                 "power_query": self.power_query, "deleted_connections": self.deleted_connections,
                 "internal_connections": self.internal_connections,
                 "refs_to_missing_file_named_sheets_not_counted": self.n_unknown_file_sheets,
                 "missing_file_named_sheet_examples": self.unknown_file_sheet_examples}
        if self.required:
            stats["links_found_but_required"] = [loc for loc, _ in (self._wb_items + self._sheet_items)][:25]
            stats["n_links_found_but_required"] = found
            return self.verdict(
                f"The task requires external links ({META_KEY}); {found} external-link place(s) found and accepted.",
                "{n}", stats)
        self.emit()
        return self.verdict(
            f"No external links: no link parts, and none of {plural(self.n_formula_cells, 'formula cell')}, "
            f"{plural(len(self.wb.defined_names), 'defined name')}, CF/DV rules, charts or connections "
            f"reference another workbook.",
            "{n} place(s) link to another workbook or external source.", stats)
