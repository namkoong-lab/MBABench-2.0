"""Grading engine.

    from detchecks.api import grade
    verdicts = grade("attempt.xlsx", checks=[92, 74])
    # {"Potential Dangers/No hidden sheets": {...verdict...}, "Formatting/No merged cells": {...}}

grade() parses workbook-level data once, then makes ONE streaming pass per sheet shared by
all selected checks (visitor hooks, see detchecks/checks/base.py), then an optional targeted
second pass, then collects the verdicts.

NO FALLBACK: any problem - unreadable or unsupported file, a check raising, a reader error,
a verdict breaking the contract - ends in GradingError naming the check(s) and the file.
grade() never returns a partial result; the verdicts of checks that did finish are attached
to the exception (GradingError.verdicts) for inspection only.
"""
from __future__ import annotations

import time
import traceback
from typing import Iterable, Optional

from .checks import REGISTRY, get as get_check
from .checks.base import Check, validate_verdict
from .core.package import Package
from .core.recalc import RecalcPolicy, ensure_values
from .core.sheet import SheetStream
from .core.values import ValueSource, detect_provenance, make_context
from .errors import GradingError

LIVE_NOTE = "recorded only; the LLM verdict stands at scoring"


def grade(path: str, *, checks: Optional[Iterable[int]] = None, value_path: Optional[str] = None,
          task_meta: Optional[dict] = None, recalc: Optional[RecalcPolicy] = None) -> dict:
    """Grade one delivered workbook.

    path        the delivered file (never the LibreOffice copy)
    checks      rubric numbers to run (default: every registered check)
    value_path  optional recalculation copy (manual override); formula values are taken from it
    task_meta   optional dict passed to checks as wb.task_meta
    recalc      a RecalcPolicy: when a selected check needs values and no value_path is given,
                the recalc pipeline (core/recalc.py) decides where values come from (Excel's
                own caches, a LibreOffice copy, or an Excel copy).  None = no pipeline (the
                file's caches under the trust policy, as before).
    Returns {rubric key: verdict}.  Every verdict carries "live" (False for a check whose
    Python verdict is recorded only, with "live_note").  Raises GradingError (no fallback)."""
    numbers = sorted(REGISTRY) if checks is None else [int(n) for n in checks]
    if not numbers:
        raise GradingError(f"grade({path!r}): no checks selected", path=path)
    dups = sorted({n for n in numbers if numbers.count(n) > 1})
    if dups:
        raise GradingError(f"grade({path!r}): check number(s) {dups} selected more than once", path=path)
    instances = [get_check(n)() for n in numbers]
    return Engine(path, instances, value_path=value_path, task_meta=task_meta, recalc=recalc).run()


class Engine:
    def __init__(self, path: str, checks: list[Check], *, value_path: Optional[str] = None,
                 task_meta: Optional[dict] = None, recalc: Optional[RecalcPolicy] = None):
        self.path = path
        self.checks = checks
        self.value_path = value_path
        self.recalc = recalc
        self.plan = None                      # core.recalc.ValuePlan when the pipeline ran
        self.task_meta = task_meta or {}
        self.failures: dict[str, str] = {}
        self.dead: set = set()
        self.timings: dict = {}
        self.retry_later = False              # a failure came from the machine (LibreOfficeUnavailable), not the file

    # ------------------------------------------------------------------ failure handling
    def _fail(self, chk: Check, exc: BaseException, where: str = ""):
        if chk in self.dead:
            return
        self.dead.add(chk)
        if isinstance(exc, GradingError):
            msg = str(exc)
        else:
            tb = traceback.extract_tb(exc.__traceback__)
            last = tb[-1] if tb else None
            at = f" at {last.filename.rsplit('/', 1)[-1]}:{last.lineno}" if last else ""
            msg = f"internal error {type(exc).__name__}: {exc}{at}"
        self.failures[chk.key] = f"check {chk.number} ({chk.key}) on {self.path}{(' [' + where + ']') if where else ''}: {msg}"

    def _call(self, chk: Check, fn, *args, where: str = ""):
        try:
            return fn(*args)
        except Exception as e:  # noqa: BLE001 - every exception becomes a GradingError for that check
            self._fail(chk, e, where)
            return None

    def _alive(self, checks):
        return [c for c in checks if c not in self.dead]

    # ------------------------------------------------------------------ run
    def run(self) -> dict:
        keys = [c.key for c in self.checks]
        if not keys:
            raise GradingError(f"no checks to run on {self.path}", path=self.path)
        if len(set(keys)) != len(keys):
            raise GradingError(f"a check is selected more than once for {self.path}: {keys}", path=self.path)
        try:
            return self._run()
        except GradingError as e:
            if not e.failures:
                # raised before any check ran (missing / unreadable file): every check failed
                e.failures = {k: str(e) for k in keys}
                e.check = e.check or keys[0]
                e.path = e.path or self.path
            raise
        except Exception as e:  # noqa: BLE001 - an engine-level bug still surfaces loudly, naming every check
            msg = f"engine error while grading {self.path}: {type(e).__name__}: {e}"
            raise GradingError(msg, check=keys[0], path=self.path, failures={k: msg for k in keys}) from e

    def _run(self) -> dict:
        t0 = time.perf_counter()
        pkg = Package.open(self.path)          # GradingError if the file is missing/unreadable
        pkg.task_meta = self.task_meta
        value_source = None
        try:
            if not pkg.is_spreadsheetml:
                self._run_unparsed(pkg)
            else:
                detect_provenance(pkg, self.value_path)
                needs = [c for c in self.checks if c.needs_values]
                if self.value_path and needs:
                    try:
                        value_source = ValueSource(self.value_path)
                        pkg.provenance.value_writer = value_source.writer
                        pkg.provenance.value_sheets = value_source.sheet_names()
                    except Exception as e:  # noqa: BLE001 - only the checks that need values fail
                        for c in needs:
                            self._fail(c, GradingError(f"value_path unusable: {e}"))
                elif self.recalc is not None and needs:
                    # the recalc pipeline decides where formula values come from; when it cannot
                    # (Excel needed and unavailable, LibreOffice failure) only value checks fail
                    try:
                        self.plan = ensure_values(self.path, self.recalc.workdir, self.recalc, pkg=pkg)
                        pkg.provenance.value_kind = self.plan.source
                        pkg.provenance.recalc = self.plan.summary()
                        value_source = self.plan.value_source()
                        if value_source is not None:
                            pkg.provenance.value_path = value_source.path
                            pkg.provenance.value_writer = value_source.writer
                            pkg.provenance.value_sheets = value_source.sheet_names()
                    except GradingError as e:
                        # LibreOfficeUnavailable (memory wait timed out, every try failed) marks the whole
                        # grading retry_later: the attempt is re-run later, when memory is free
                        self.retry_later = self.retry_later or bool(getattr(e, "retry_later", False))
                        for c in needs:
                            self._fail(c, GradingError(f"values unavailable: {e}"))
                    except Exception as e:  # noqa: BLE001
                        for c in needs:
                            self._fail(c, GradingError(f"recalc pipeline error {type(e).__name__}: {e}"))
                pkg.value_source = value_source      # core.lookup (cells a conditional format refers to)
                self._run_parsed(pkg, value_source)
            verdicts = self._finish(pkg)
        finally:
            if value_source is not None:
                value_source.close()
            pkg.close()
        self.timings["total_s"] = round(time.perf_counter() - t0, 3)
        if self.failures:
            keys = ", ".join(self.failures)
            msg = f"{len(self.failures)} check(s) could not grade {self.path}: {keys}\n" + \
                  "\n".join(f"  - {m}" for m in self.failures.values())
            raise GradingError(msg, check=next(iter(self.failures)), path=self.path,
                               failures=dict(self.failures), verdicts=verdicts, retry_later=self.retry_later)
        return verdicts

    def _run_unparsed(self, pkg: Package):
        for chk in self.checks:
            if chk.accepts_unparsed:
                self._call(chk, chk.start, pkg)
            else:
                self._fail(chk, GradingError(
                    f"cannot grade a {pkg.format!r} file ({pkg.unsupported_reason}); only SpreadsheetML "
                    f".xlsx/.xlsm workbooks are supported by this check"))

    def _run_parsed(self, pkg: Package, value_source):
        for chk in self.checks:
            self._call(chk, chk.start, pkg, where="start")
        for info in pkg.sheets:
            active = self._alive(self.checks)
            if not active:
                break
            sheet_checks = [c for c in active if self._call(c, c.wants_sheet, info, where=f"wants_sheet {info.name}")]
            if sheet_checks:
                self._stream_sheet(pkg, info, sheet_checks, value_source, first_pass=True)
        # optional targeted second pass.  Requested names are resolved like Excel does (exact,
        # then case-insensitive) to the workbook's own sheet; an unknown name fails the check.
        requests: dict[str, dict] = {}       # actual sheet name -> {check: cells (None = all)}
        for chk in self._alive(self.checks):
            for sheet_name, cells in getattr(chk, "_second_pass", {}).items():
                info = pkg.sheet(sheet_name)
                if info is None:
                    self._fail(chk, GradingError(f"second pass requested for unknown sheet {sheet_name!r}"))
                    continue
                per = requests.setdefault(info.name, {})
                if chk in per and (per[chk] is None or cells is None):
                    per[chk] = None
                elif chk in per:
                    per[chk] = set(per[chk]) | set(cells)
                else:
                    per[chk] = cells
        for info in pkg.sheets:
            reqs = [(c, cells) for c, cells in requests.get(info.name, {}).items() if c not in self.dead]
            if reqs:
                self._stream_sheet(pkg, info, reqs, value_source, first_pass=False)

    def _stream_sheet(self, pkg: Package, info, checks, value_source, first_pass: bool):
        where = f"sheet {info.name!r}" + ("" if first_pass else " (second pass)")
        t0 = time.perf_counter()
        if first_pass:
            chks = checks
        else:
            chks = [c for c, _ in checks]
            wanted = {c: cells for c, cells in checks}
        ss = SheetStream(pkg, info)
        try:
            try:
                head = ss.read_head()
            except Exception as e:  # noqa: BLE001 - reader error fails every check on this sheet
                for c in chks:
                    self._fail(c, e, where)
                return
            body_checks = []
            for c in chks:
                if first_pass:
                    r = self._call(c, c.sheet_start, head, where=where)
                    if c in self.dead:
                        continue
                    if (c.needs_cells or c.needs_rows) and r is not False:
                        body_checks.append(c)
                else:
                    self._call(c, c.second_pass_start, head, where=where)
                    if c not in self.dead:
                        body_checks.append(c)
            try:
                if body_checks:
                    ctx = make_context(pkg, info, value_source,
                                       needs_values=any(c.needs_values for c in body_checks))
                    on_row, on_cell = self._dispatchers(body_checks, info, first_pass,
                                                        None if first_pass else wanted)
                    tail = ss.read_body(on_row, on_cell, ctx)
                else:
                    tail = ss.skip_body()
            except Exception as e:  # noqa: BLE001 - reader error (or a hook error escaping dispatch)
                for c in self._alive(chks):
                    self._fail(c, e, where)
                return
            for c in self._alive(chks):
                if first_pass:
                    self._call(c, c.sheet_end, head, tail, where=where)
                else:
                    self._call(c, c.second_pass_end, head, tail, where=where)
        finally:
            ss.close()
            self.timings.setdefault("sheets", []).append(
                {"sheet": info.name, "pass": 1 if first_pass else 2, "s": round(time.perf_counter() - t0, 3),
                 "streamed_cells": bool(chks) and any((c.needs_cells or c.needs_rows) for c in chks)})

    def _dispatchers(self, checks, info, first_pass: bool, wanted):
        engine = self
        if first_pass:
            row_hooks = [(c, c.row) for c in checks if c.needs_rows]
            cell_hooks = [(c, c.cell) for c in checks if c.needs_cells]
        else:
            row_hooks = []
            cell_hooks = [(c, c.second_pass_cell) for c in checks]

        def on_row(row):
            bad = False
            for c, h in row_hooks:
                try:
                    h(row)
                except Exception as e:  # noqa: BLE001
                    engine._fail(c, e, f"sheet {info.name!r} row {row.r}")
                    bad = True
            if bad:
                row_hooks[:] = [(c, h) for c, h in row_hooks if c not in engine.dead]

        if first_pass:
            def on_cell(cell):
                bad = False
                for c, h in cell_hooks:
                    try:
                        h(cell)
                    except Exception as e:  # noqa: BLE001
                        engine._fail(c, e, f"{info.name}!{cell.ref}")
                        bad = True
                if bad:
                    cell_hooks[:] = [(c, h) for c, h in cell_hooks if c not in engine.dead]
        else:
            def on_cell(cell):
                key = (cell.row, cell.col)
                for c, h in cell_hooks:
                    cells = wanted[c]
                    if cells is not None and key not in cells:
                        continue
                    if c in engine.dead:
                        continue
                    try:
                        h(cell)
                    except Exception as e:  # noqa: BLE001
                        engine._fail(c, e, f"{info.name}!{cell.ref} (second pass)")
        return (on_row if row_hooks else None), (on_cell if cell_hooks else None)

    def _finish(self, pkg=None) -> dict:
        verdicts = {}
        values_stats = None
        prov = getattr(pkg, "provenance", None) if pkg is not None else None
        if prov is not None and (prov.value_kind or prov.value_path):
            values_stats = {"writer": prov.writer, "source": prov.value_kind or "value_path",
                            "value_path": prov.value_path, "value_writer": prov.value_writer}
            if prov.recalc:
                values_stats.update({k: prov.recalc.get(k) for k in
                                     ("n_gaps", "gap_functions", "gaps", "n_lo_errors", "timings", "lo_version", "notes")})
        for chk in self._alive(self.checks):
            v = self._call(chk, chk.finish, where="finish")
            if chk in self.dead:
                continue
            try:
                validate_verdict(v)
            except GradingError as e:
                self._fail(chk, GradingError(f"verdict contract violated: {e}"))
                continue
            # per-check live switch: every check runs and records its verdict; a non-live one
            # is marked so scoring keeps the LLM's verdict (handoff ruling 2026-10-04)
            v["live"] = bool(getattr(chk, "live", True))
            if not v["live"]:
                v["live_note"] = LIVE_NOTE
            if values_stats is not None and chk.needs_values:
                v["stats"].setdefault("values", values_stats)
            verdicts[chk.key] = v
        return verdicts
