"""Deterministic rubric checks in the judge (judge v13).

`detchecks` (judge/detchecks/, API in detchecks/docs/reader.md) grades rubric_9 checks in
Python from the DELIVERED workbook. This module is the only place the judge calls it:

    run = run_det_checks(task_folder, rubric_path=..., weights_path=..., mode=None)
    harness_verdicts = merge_harness_verdicts(answer_check_verdicts, run.harness_verdicts)
    single_pass_judge_case(..., harness_verdicts=harness_verdicts, det_checks=run.for_judge())

What runs
  project_configs.yaml det_checks.live           the Python verdict replaces the LLM's at scoring
  project_configs.yaml det_checks.recorded_only  the Python verdict is recorded beside the LLM's,
                                                 and the LLM's counts
  Both are rubric numbers. Each number is pinned to its (category, name) in DET_CHECK_NAMES and
  checked against the loaded rubric and the detchecks registry; any mismatch refuses to grade
  (numbering drift must never move a verdict onto the wrong check). A configured live check
  must also be live in detchecks (Check.live). A check is graded only when the task's
  rubric-suitability gate leaves it applicable (retired checks never are) and the effective
  weights score it: the same gate judge._apply_harness_verdicts applies.

Representation (the judge's harness_verdicts format, key "<Category>/<check name>")
  live           engine "harness", decision, summary, mistakes, fallback_reason None,
                 family "det_checks", live True, check_no, n_mistakes, stats
  recorded only  the same with engine "llm", live False and fallback_reason = detchecks'
                 LIVE_NOTE ("recorded only; the LLM verdict stands at scoring")
  `engine` keeps its one meaning for every reader: "harness" = this verdict may replace the
  LLM's; "llm" = the LLM's verdict stands (here with the Python decision recorded beside it).
  judge._apply_harness_verdicts records decision, llm_decision and `agreed` for both kinds and
  overlays engine "harness" entries only.

The switch (--det-checks on every driver; default from det_checks.enabled: true -> harness)
  harness  live Python verdicts count in the recorded total (judge v13 default)
  llm      everything runs and is recorded; the LLM's verdicts count (shadow run;
           scored_results.accuracy_engine.total_score_harness still shows the v13 total)
  off      nothing runs; scored_results.det_checks says so
  --accuracy-check keeps deciding the answer check (Final calculation accuracy) on its own.

No fallback (maintainer, 2026-10-02)
  A check that cannot grade the file raises detchecks' GradingError; run_det_checks writes the
  failure to det_checks.json and raises DetChecksError naming every failing check by title and
  the file. Every driver calls it FIRST - before the answer check (which recalculates through
  LibreOffice too, so a file refused here never reaches it) and the LLM - so the attempt fails
  the way the formula-cache refusal does (grade.py logs FAILED, no grading row, the batch
  continues) with no API spend.
  Nothing here ever turns a failure into an LLM verdict. This holds in `llm` mode too: when the
  checks run, they run loudly.

Values (detchecks/docs/recalc.md)
  Structure and styles come from task_folder/ai_attempt.xlsx, the delivered file (never
  temp_recalculated/; a legacy .xls delivery is the one exception, below). Formula values: an
  Excel-saved file's own caches (any size); any other
  writer's file is recalculated by LibreOffice (paths.libreoffice_path) into task_folder/
  det_checks_recalc/ (the copy and the private profiles), which run_det_checks deletes itself
  as soon as the checks return or raise, in every driver (_remove_recalc_dir; det_checks.json
  stays). Excel recalculation is OFF (det_checks.excel_recalc); cells LibreOffice
  cannot compute are used as displayed.
  Every attempt is graded, whatever its size (maintainer 2026-10-05): det_checks.libreoffice_max_mb
  is 0 = no limit in production (a positive value, for test runs only, refuses larger files that
  need LibreOffice: "not graded: too large"). Memory safety instead (detchecks/core/lo_guard.py),
  for EVERY LibreOffice run of a grading - this pipeline, the answer check's recalculation and
  --run-calculation (run_libreoffice below): one LibreOffice at a time on the machine
  (det_checks.libreoffice_lock_path, default /tmp/mbabench_libreoffice.lock, waited for at most
  det_checks.libreoffice_max_lock_wait_minutes), started only once
  det_checks.libreoffice_min_free_pct of memory is free (waiting at most
  det_checks.libreoffice_max_wait_minutes), a failed run (crash, no copy, timeout) retried
  det_checks.libreoffice_retries times with the timeout doubled. When that is exhausted the grading
  fails loudly with retry_later set (DetChecksError / LibreOfficeUnavailable): FAILED, no grading row,
  the batch continues, and the drivers list the attempt at the end of the run to be re-run when
  the machine has memory to spare (retry_later_report).
  LibreOffice never outlives the grading (detchecks/docs/recalc.md, "The LibreOffice process"):
  its paths are encoded file URLs (a task folder with a space or '%'); it runs in the grader's
  process group under a watchdog that kills it when the grader dies (SIGKILL included); a
  SIGTERM / SIGHUP handler, installed by startup_check and run_det_checks from the main thread,
  kills it before the grader dies of the signal; and _reap_libreoffice sweeps this grading's
  private profiles (raw or URL form) when the checks return or raise.

Legacy .xls deliveries (maintainer 2026-10-05: "just keep doing whatever v12 did or does")
  grade.py stages the first .xlsx / .xlsm / .xls delivery as ai_attempt.xlsx without
  converting it. legacy_xls() tells from the bytes (as detchecks.core.package classifies them,
  whatever the name) that the staged file is a legacy binary Excel workbook: an OLE2 compound
  file holding a Workbook / Book stream. Then, as in judge v12:
  - with --run-calculation (run_det_checks(..., run_calculation=True); v12 grades LibreOffice's
    re-saved copy): LibreOffice converts it to .xlsx through the pipeline's own LibreOffice step
    (the memory guard, private profile, watchdog and reaper above; a test-run size limit applies)
    into det_checks_recalc/xls_converted/ai_attempt.xlsx, and every check except File extension
    (.xlsx) (77) grades that copy: structure, styles and values. LibreOffice recalculates every
    formula of an .xls it loads, so the copy's values are LibreOffice's own; the recalculation
    pipeline takes them from the copy (no second LibreOffice run; gaps as usual). File extension
    (.xlsx) (77) grades the delivered file and its delivered name, and fails. Recorded in
    det_checks.json "xls_conversion", scored_results.det_checks.xls_conversion and each verdict's
    stats.graded_on ("libreoffice_xlsx_conversion" / "delivered_file"). The copy goes with
    det_checks_recalc/. The judge's own --run-calculation re-save (temp_recalculated/, after the
    checks and the answer check, v12's code) converts the file again; the checks never read it.
  - without --run-calculation: unchanged, the checks other than File extension (.xlsx) (77) cannot
    read .xls bytes, so the grading fails loudly before the LLM call (the error says to re-run
    with --run-calculation; det_checks.json "xls_conversion" says why); v12 fails too (openpyxl
    cannot open the file).
  .xlsx / .xlsm deliveries, an encrypted package, an OLE2 file without a workbook stream (or with
  an unreadable directory) and .xlsb (never staged) are graded exactly as before.

Task metadata
  delivered_filename       from the _attempt_origin.json sidecar (original_filename). Without
                           it, or with a malformed sidecar (read_origin: not JSON, not an object,
                           no non-empty original_filename), File extension (.xlsx) (77) judges
                           the format from the file's content (maintainer 2026-10-05: every
                           attempt graded; task_meta delivered_filename_problem says why, and
                           the verdict's stats.defaults records it).
  requires_external_links  always False (maintainer: no task requires external links).

Artefacts
  task_folder/det_checks.json  everything (config, gate, task_meta, the full verdicts and the
                               full recalculation block, or the failure and its stage);
                               written in every case, also when a config error, a rubric
                               drift or a suitability refusal stops the run;
                               judge._finalize_case copies it into the output dir
  scored_results.det_checks    compact block (status, mode, per-check engine / decision / live /
                               n_mistakes / summary capped at DB_SUMMARY_CAP characters), the
                               recalculation block `values` ONCE (no local value_path) and, for a
                               legacy .xls delivery only, `xls_conversion`
  stats in harness_verdicts    -> scored_results.accuracy_engine.checks[key].stats: a value check's
                               stats.values is the reference {"ref": "det_checks.values"}; in the
                               row payload every list longer than DB_LIST_CAP keeps its first
                               DB_LIST_CAP items, the original lengths under "db_capped"
"""
from __future__ import annotations

import dataclasses
import hashlib
import json
import math
import os
import shutil
import time
from pathlib import Path

from detchecks.errors import GradingError

try:
    from .logger import logger
    from .misc_utils import load_env_var, libreoffice_path as _libreoffice_path
    from . import rubric_suitability, workbook_properties
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger
    from misc_utils import load_env_var, libreoffice_path as _libreoffice_path
    import rubric_suitability
    import workbook_properties

FAMILY = "det_checks"
MODES = ("harness", "llm", "off")
ATTEMPT_FILENAME = "ai_attempt.xlsx"
ARTEFACT_FILENAME = "det_checks.json"
RECALC_DIRNAME = "det_checks_recalc"
# A legacy .xls delivery graded with --run-calculation (module doc): LibreOffice's .xlsx conversion goes to
# det_checks_recalc/xls_converted/, and every verdict's stats.graded_on says which file it graded.
XLS_COPY_DIRNAME = "xls_converted"
GRADED_ON_COPY = "libreoffice_xlsx_conversion"
GRADED_ON_DELIVERED = "delivered_file"

# The name each rubric number must carry (rubric_9 flat numbering, as in
# rubric_suitability.RETIRED_CHECK_NAMES). The config lists decide what runs and what counts;
# a pin only permits it. A number missing here, or whose rubric position or detchecks key
# carries another name, refuses to grade.
DET_CHECK_NAMES = {
    22: ("Error Checks", "No formula errors"),
    29: ("Error Checks", "Clean Name Manager"),
    47: ("Formatting", "No bright-yellow highlighting"),
    49: ("Formatting", "Black font for calculations"),
    50: ("Formatting", "Green font for cross-sheet links"),
    51: ("Formatting", "Red font for external links"),
    61: ("Formatting", "Consistent zoom level"),
    62: ("Formatting", "Active cell reset to A1 on all sheets"),
    65: ("Formatting", "Negatives in parentheses"),
    66: ("Formatting", "Zeros as dashes"),
    69: ("Formatting", "Sufficient column widths"),
    70: ("Formatting", "Reasonable column widths"),
    73: ("Formatting", "Reasonable row heights"),
    74: ("Formatting", "No merged cells"),
    77: ("Formatting", "File extension (.xlsx)"),
    80: ("Formulas", "Avoid volatile functions"),
    87: ("Formulas", "Avoid whole-column references"),
    92: ("Potential Dangers", "No hidden sheets"),
    93: ("Potential Dangers", "No hidden rows/columns"),
    94: ("Potential Dangers", "No white-on-white hiding"),
    95: ("Potential Dangers", "No external links"),
}
_NUMBER_BY_KEY = {f"{cat}/{name}": no for no, (cat, name) in DET_CHECK_NAMES.items()}
_FILE_EXTENSION_CHECK = 77


class DetChecksError(GradingError):
    """The deterministic checks could not grade the delivered workbook. The grading must
    fail here, before the LLM (no fallback). Carries GradingError's check/path/failures/
    verdicts; the message names every failing check by title and the file."""


class DetChecksConfigError(Exception):
    """project_configs.yaml det_checks does not match the rubric or the detchecks registry."""


@dataclasses.dataclass(frozen=True)
class DetChecksSettings:
    enabled: bool
    live: tuple
    recorded_only: tuple
    excel_recalc: bool
    libreoffice_path: str
    libreoffice_timeout_s: float
    libreoffice_max_mb: float = 0.0          # 0 = no size limit (production, maintainer 2026-10-05)
    libreoffice_retries: int = 3
    libreoffice_min_free_pct: float = 25.0
    libreoffice_max_wait_s: float = 3600.0
    libreoffice_lock_path: str = ""          # "" = detchecks.core.lo_guard.default_lock_path()
    libreoffice_max_lock_wait_s: float = 3 * 3600.0   # longest wait for the lock (maintainer 2026-10-05)

    def size_limited(self) -> bool:
        return self.libreoffice_max_mb > 0

    def size_text(self) -> str:
        return f"files up to {self.libreoffice_max_mb:g} MB" if self.size_limited() else "no size limit"

    def guard_settings(self):
        """The machine-wide LibreOffice guard (detchecks/core/lo_guard.py) these settings describe."""
        from detchecks.core import lo_guard

        return lo_guard.GuardSettings(lock_path=self.libreoffice_lock_path or None,
                                      min_free_pct=self.libreoffice_min_free_pct,
                                      max_wait_s=self.libreoffice_max_wait_s,
                                      max_lock_wait_s=self.libreoffice_max_lock_wait_s,
                                      retries=self.libreoffice_retries)

    def guard_text(self) -> str:
        from detchecks.core import lo_guard

        return (f"one LibreOffice at a time (lock {self.libreoffice_lock_path or lo_guard.default_lock_path()}, "
                f"waited for at most {self.libreoffice_max_lock_wait_s / 60:g} min), "
                f"started at >= {self.libreoffice_min_free_pct:g}% free memory (waiting at most "
                f"{self.libreoffice_max_wait_s / 60:g} min), {self.libreoffice_retries} retries with the timeout "
                f"doubled")


@dataclasses.dataclass
class DetChecksRun:
    """What run_det_checks returns: the entries to merge into harness_verdicts, the compact
    block for scored_results.det_checks, and the artefact written into the task folder."""
    mode: str
    harness_verdicts: dict
    summary: dict
    artefact_path: Path | None = None

    def for_judge(self) -> dict:
        """The `det_checks` argument of judge.single_pass_judge_case / _finalize_case."""
        return {
            "mode": self.mode,
            "summary": self.summary,
            "artefact": str(self.artefact_path) if self.artefact_path else None,
        }


# --------------------------------------------------------------------------- config
def label(number: int) -> str:
    """The maintainer's naming convention: 'No hidden sheets (92)'."""
    pin = DET_CHECK_NAMES.get(int(number))
    return f"{pin[1]} ({number})" if pin else f"check {number}"


def label_for_key(key: str) -> str:
    no = _NUMBER_BY_KEY.get(key)
    return label(no) if no is not None else key


def _as_bool(value) -> bool:
    return str(value).strip().lower() in ("1", "true", "yes", "on")


def _numbers(env_key: str) -> tuple:
    raw = str(load_env_var(env_key, default="") or "").strip()
    if not raw or raw.lower() in ("none", "null", "[]"):
        return ()
    out = []
    for tok in raw.replace(";", ",").split(","):
        tok = tok.strip()
        if not tok:
            continue
        if not tok.isdigit():
            raise DetChecksConfigError(f"{env_key}: {tok!r} is not a check number")
        out.append(int(tok))
    if len(set(out)) != len(out):
        raise DetChecksConfigError(f"{env_key}: a check number is listed twice ({raw!r})")
    return tuple(out)


def load_settings() -> DetChecksSettings:
    """det_checks.* from project_configs.yaml (env MBABENCHJUDGE_DET_CHECKS_*); the
    LibreOffice binary is the judge's own paths.libreoffice_path."""
    try:
        settings = _read_settings()
    except ValueError as e:
        raise DetChecksConfigError(f"det_checks: {e}") from None
    bad = []
    if not settings.libreoffice_timeout_s > 0:
        bad.append(f"det_checks.libreoffice_timeout_seconds ({settings.libreoffice_timeout_s}) must be positive")
    if not settings.libreoffice_max_mb >= 0:
        bad.append(f"det_checks.libreoffice_max_mb ({settings.libreoffice_max_mb}) must be 0 / null (no limit) "
                   f"or positive")
    bad += [f"det_checks.libreoffice_* settings: {p}" for p in settings.guard_settings().problems()]
    if bad:
        raise DetChecksConfigError("; ".join(bad))
    return settings


def _config_number(env_key: str, default, cast=float):
    """A numeric det_checks value; '', 'none' and 'null' mean the default (a YAML null is not exported)."""
    raw = load_env_var(env_key, default=default)
    if raw is None or str(raw).strip().lower() in ("", "none", "null"):
        raw = default
    try:
        value = cast(float(raw)) if cast is int and float(raw) == int(float(raw)) else cast(raw)
    except (TypeError, ValueError, OverflowError):
        raise ValueError(f"{env_key.lower()} is {raw!r}, not a number") from None
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError(f"{env_key.lower()} is {raw!r}, not a finite number")
    return value


def _read_settings() -> DetChecksSettings:
    return DetChecksSettings(
        enabled=_as_bool(load_env_var("DET_CHECKS_ENABLED", default="false")),
        live=_numbers("DET_CHECKS_LIVE"),
        recorded_only=_numbers("DET_CHECKS_RECORDED_ONLY"),
        excel_recalc=_as_bool(load_env_var("DET_CHECKS_EXCEL_RECALC", default="false")),
        libreoffice_path=_libreoffice_path(required=False),
        libreoffice_timeout_s=_config_number("DET_CHECKS_LIBREOFFICE_TIMEOUT_SECONDS", 600),
        libreoffice_max_mb=_config_number("DET_CHECKS_LIBREOFFICE_MAX_MB", 0),
        libreoffice_retries=_config_number("DET_CHECKS_LIBREOFFICE_RETRIES", 3, int),
        libreoffice_min_free_pct=_config_number("DET_CHECKS_LIBREOFFICE_MIN_FREE_PCT", 25),
        libreoffice_max_wait_s=60.0 * _config_number("DET_CHECKS_LIBREOFFICE_MAX_WAIT_MINUTES", 60),
        libreoffice_lock_path=str(load_env_var("DET_CHECKS_LIBREOFFICE_LOCK_PATH", default="") or "").strip(),
        libreoffice_max_lock_wait_s=60.0 * _config_number("DET_CHECKS_LIBREOFFICE_MAX_LOCK_WAIT_MINUTES", 180),
    )


def resolve_mode(cli_value: str | None = None, settings: DetChecksSettings | None = None) -> str:
    """--det-checks value, else the config default (enabled -> harness, disabled -> off)."""
    if cli_value is not None:
        if cli_value not in MODES:
            raise ValueError(f"--det-checks must be one of {MODES}, got {cli_value!r}")
        return cli_value
    settings = settings or load_settings()
    return "harness" if settings.enabled else "off"


def add_det_checks_arg(parser) -> None:
    """--det-checks harness|llm|off (judge v13). Shared by every single-pass driver."""
    parser.add_argument(
        "--det-checks",
        dest="det_checks",
        choices=list(MODES),
        default=None,
        help=(
            "(Single-pass) Deterministic rubric checks (utils/det_checks.py): 'harness' = the "
            "Python verdicts of the live checks replace the LLM's in the recorded total (judge "
            "v13); 'llm' = they run and are recorded beside the LLM's, which count (shadow); "
            "'off' = they do not run. Default: project_configs.yaml det_checks.enabled "
            "(true -> harness). A check that cannot grade the file fails the attempt before "
            "the LLM call in both 'harness' and 'llm' (no fallback)."
        ),
    )


def configured_checks(settings: DetChecksSettings, rubric: dict, rubric_src=None) -> dict:
    """{number: live?} for the configured checks, validated against the pins, the loaded
    rubric (read from `rubric_src`, named in the errors) and the detchecks registry. Positions
    beyond the rubric are dropped: those checks do not exist there.
    Raises DetChecksConfigError."""
    from detchecks.checks import REGISTRY

    overlap = sorted(set(settings.live) & set(settings.recorded_only))
    if overlap:
        raise DetChecksConfigError(
            f"det_checks: {[label(n) for n in overlap]} listed as both live and recorded_only")
    flat = [(cat, c["name"]) for cat, checks in rubric.items() for c in checks]
    where = f"the loaded rubric ({rubric_src})" if rubric_src else "the loaded rubric"
    plan = {}
    for no in list(settings.live) + list(settings.recorded_only):
        live = no in settings.live
        pin = DET_CHECK_NAMES.get(no)
        if pin is None:
            named = f"{flat[no - 1][1]} ({no})" if 0 < no <= len(flat) else f"check {no} (beyond {where})"
            raise DetChecksConfigError(
                f"det_checks names {named}, which has no entry in DET_CHECK_NAMES "
                f"(utils/det_checks.py) - add the (category, name) pin first")
        cls = REGISTRY.get(no)
        if cls is None:
            raise DetChecksConfigError(f"det_checks names {label(no)}, which detchecks does not implement")
        if cls.key != f"{pin[0]}/{pin[1]}":
            raise DetChecksConfigError(
                f"{label(no)} is pinned to {pin!r} but detchecks registers {cls.key!r} under that number "
                f"- refusing to grade")
        if no > len(flat):
            continue
        if flat[no - 1] != pin:
            raise DetChecksConfigError(
                f"{label(no)} is pinned to {pin!r} but {where} has {flat[no - 1]!r} at that position "
                f"- numbering drift; refusing to grade")
        if live and not getattr(cls, "live", True):
            raise DetChecksConfigError(
                f"det_checks.live lists {label(no)}, but detchecks marks it not live (Check.live = False); "
                f"move it to det_checks.recorded_only or change the check")
        plan[no] = live
    return plan


def startup_check(rubric_path, mode: str | None = None) -> str:
    """Validate det_checks against the run's rubric once, before a driver grades anything,
    and log what will run. Returns the resolved mode; raises DetChecksConfigError, so a bad
    config stops the run before any download or API call (rather than failing every
    attempt one by one)."""
    settings = load_settings()
    mode = resolve_mode(mode, settings)
    if mode == "off":
        logger.info("det_checks: off - the LLM decides every check (judge v13 rows without Python verdicts)")
        return mode
    _install_termination_reaper()
    rubric = json.loads(Path(rubric_path).read_text(encoding="utf-8"))
    plan = configured_checks(settings, rubric, rubric_path)
    logger.info(
        f"det_checks: mode={mode}; live: {[label(n) for n, live in sorted(plan.items()) if live]}; "
        f"recorded only: {[label(n) for n, live in sorted(plan.items()) if not live]}; "
        f"LibreOffice {settings.libreoffice_path} ({settings.size_text()}, first timeout "
        f"{settings.libreoffice_timeout_s:.0f} s; {settings.guard_text()}); Excel recalculation "
        f"{'ON' if settings.excel_recalc else 'off'}")
    if settings.size_limited():
        logger.warning(f"det_checks: libreoffice_max_mb = {settings.libreoffice_max_mb:g} refuses larger files not "
                       f"saved by Excel - a TEST-RUN setting; production grades every attempt (0 = no limit)")
    return mode


# --------------------------------------------------------------------------- gate
def _judge_rubric(task_folder: Path, rubric_path) -> tuple[dict, Path]:
    """The rubric the judge itself will load: a task folder's own rubric.json wins over the
    configured one (excel_utils.copy_support_files -> output_dir/rubric.json)."""
    local = Path(task_folder) / "rubric.json"
    src = local if local.exists() else Path(rubric_path)
    return json.loads(src.read_text(encoding="utf-8")), src


def scored_checks(task_folder, rubric: dict, weights: dict) -> tuple[set, dict]:
    """{(category, name)} the judge scores for this task, and the suitability provenance:
    the applicable checks after rubric_suitability.load_for_case (annotation + retired
    checks; the same call single_pass_judge_case makes) that the effective weights still
    score. Raises SuitabilityError exactly where the judge would."""
    suit = rubric_suitability.load_for_case(Path(task_folder), rubric)
    if suit is None:
        applicable = {(cat, c["name"]) for cat, checks in rubric.items() for c in checks}
        effective = weights
        provenance = {"gated": False}
        if rubric_suitability.skip_requested():
            provenance["skipped_via_env"] = True
    else:
        applicable = {(cat, n) for cat, names in suit["applicable"].items() for n in names}
        effective = rubric_suitability.build_effective_weights(weights, suit["excluded"])
        p = suit.get("provenance") or {}
        provenance = {k: p.get(k) for k in ("gated", "source", "annotator", "excluded_count",
                                             "retired_checks", "skipped_via_env") if k in p}
    weighted = {
        (cat, e["name"])
        for cat, entries in (effective or {}).items()
        if cat != "CategoryWeights" and isinstance(entries, list)
        for e in entries
        if isinstance(e, dict) and "name" in e
    }
    return applicable & weighted, provenance


# --------------------------------------------------------------------------- helpers
def merge_harness_verdicts(base: dict | None, det: dict | None) -> dict:
    """The answer check's entries plus the deterministic family's. Keys are disjoint by
    construction (Accuracy/* and Rounding/* vs the detchecks keys); an overlap is a bug."""
    merged = dict(base or {})
    overlap = sorted(set(merged) & set(det or {}))
    if overlap:
        raise ValueError(f"harness_verdicts: answer-check and deterministic-check keys overlap: {overlap}")
    merged.update(det or {})
    return merged


def json_safe(obj):
    """A copy of `obj` made of JSON-native values only: scores.json is dumped without
    default= and scored_results goes into a JSONB column (no NaN, no tuples or sets as keys,
    no str subclasses such as detchecks' ExcelError)."""
    if obj is None or isinstance(obj, bool):
        return obj
    if isinstance(obj, int):
        return int(obj)
    if isinstance(obj, float):
        return float(obj) if math.isfinite(obj) else str(obj)
    if isinstance(obj, str):
        return str(obj)
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [json_safe(v) for v in obj]
    if isinstance(obj, (set, frozenset)):
        items = [json_safe(v) for v in obj]
        try:
            return sorted(items)
        except TypeError:
            return sorted(items, key=repr)
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return json_safe(dataclasses.asdict(obj))
    if isinstance(obj, os.PathLike):
        return os.fspath(obj)
    if isinstance(obj, (bytes, bytearray)):
        return bytes(obj).decode("utf-8", "replace")
    return str(obj)


_CODE_SHA = None
# The grading code: what code_sha fingerprints. Nothing else under detchecks/ (scratch/ is
# git-ignored and full of probes, tests/, tools/, docs/, out/) may move the fingerprint, so the
# same commit always records the same code_sha.
CODE_SHA_FILES = ("__init__.py", "api.py", "errors.py")
CODE_SHA_DIRS = ("core", "checks")


def code_sha_files(root) -> list:
    """The .py files code_sha hashes under a detchecks package folder `root`, in hashing order."""
    root = Path(root)
    files = [root / f for f in CODE_SHA_FILES if (root / f).is_file()]
    for d in CODE_SHA_DIRS:
        files += [p for p in (root / d).rglob("*.py") if "__pycache__" not in p.parts]
    return sorted(files, key=lambda p: p.relative_to(root).as_posix())


def code_sha_of(root) -> str:
    """Fingerprint of the grading code in the detchecks package folder `root`: relative path and
    bytes of detchecks/__init__.py, api.py, errors.py and every .py under core/ and checks/."""
    root = Path(root)
    h = hashlib.sha256()
    for p in code_sha_files(root):
        h.update(p.relative_to(root).as_posix().encode())
        h.update(b"\0")
        h.update(p.read_bytes())
        h.update(b"\0")
    return h.hexdigest()[:16]


def code_sha() -> str:
    """Fingerprint of the grading code (code_sha_of the imported detchecks), recorded per grading."""
    global _CODE_SHA
    if _CODE_SHA is None:
        import detchecks

        _CODE_SHA = code_sha_of(Path(detchecks.__file__).resolve().parent)
    return _CODE_SHA


# --------------------------------------------------------------------------- row payload
# What lands in gradings.scored_results (and scores.json): the recalculation block `values` once,
# at scored_results.det_checks.values, each value check's stats.values a reference to it (it was
# copied into every value check: 7 identical copies, ~9 KB each with a LibreOffice gap list), and
# every list longer than DB_LIST_CAP cut to its first DB_LIST_CAP items, the original length
# recorded under "db_capped". det_checks.json (the bundle) keeps the full verdicts and the full
# values block.
DB_LIST_CAP = 10
DB_SUMMARY_CAP = 300
VALUES_REF = "det_checks.values"


def _cap_lists(obj, path: str, capped: dict):
    """A copy of `obj` (JSON-safe) whose lists keep their first DB_LIST_CAP items; `capped` gets
    {json path: original length} for every list cut."""
    if isinstance(obj, dict):
        return {k: _cap_lists(v, f"{path}.{k}" if path else k, capped) for k, v in obj.items()}
    if isinstance(obj, list):
        if len(obj) > DB_LIST_CAP:
            capped[path] = len(obj)
            obj = obj[:DB_LIST_CAP]
        return [_cap_lists(v, f"{path}[{i}]", capped) for i, v in enumerate(obj)]
    return obj


def _db_capped(obj: dict) -> dict:
    capped: dict = {}
    out = _cap_lists(obj, "", capped)
    if capped:
        out["db_capped"] = capped
    return out


def _db_values(values: dict | None) -> dict | None:
    """The shared recalculation block as stored once in scored_results.det_checks.values: no local
    value_path (deleted with det_checks_recalc/), lists capped (gaps: first DB_LIST_CAP of n_gaps)."""
    if values is None:
        return None
    return _db_capped({k: v for k, v in values.items() if k != "value_path"})


def _db_stats(stats: dict, shared_values: dict | None) -> dict:
    """A check's stats as stored in scored_results.accuracy_engine.checks: stats.values -> the
    reference VALUES_REF when it is the shared block, long lists capped."""
    out = dict(stats)
    if shared_values is not None and out.get("values") == shared_values:
        out["values"] = {"ref": VALUES_REF}
    return _db_capped(out)


def _short(text, n: int = DB_SUMMARY_CAP) -> str:
    text = str(text or "")
    return text if len(text) <= n else text[: n - 3] + "..."


def _file_sha256(path: Path) -> str | None:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return None


def _write_artefact(path: Path, record: dict) -> None:
    try:
        path.write_text(json.dumps(json_safe(record), indent=1, allow_nan=False), encoding="utf-8")
    except OSError as e:
        logger.warning(f"  [det_checks] could not write {path}: {e}")


def _size_guarded_libreoffice(limit_mb: float):
    """The recalculation pipeline's LibreOffice step (RecalcPolicy.lo_runner). Production has no size
    limit (limit_mb 0; maintainer 2026-10-05: every attempt is graded, whatever its size - memory is
    protected by the guard in detchecks/core/lo_guard.py instead). A positive limit_mb
    (det_checks.libreoffice_max_mb, test runs only) refuses a larger file with GradingError, so the value
    checks fail and the grading stops loudly before the LLM call ("not graded: too large"). It only
    fires when LibreOffice is actually needed: an Excel-saved file of any size is read from its own
    caches and never gets here."""

    def _run(src, out_dir, policy):
        size = os.path.getsize(src)
        if limit_mb and limit_mb > 0 and size > limit_mb * 1_000_000:
            raise GradingError(
                f"{src} is {size / 1_000_000:.1f} MB, over det_checks.libreoffice_max_mb "
                f"({limit_mb:g} MB, a test-run setting; production uses 0 = no limit): a file not saved by "
                f"Excel needs a LibreOffice recalculation for its formula values, which this configuration "
                f"does not run on files this large - not graded: too large")
        from detchecks.core import recalc as det_recalc   # looked up per call (tests swap it)

        return det_recalc.libreoffice_recalc(src, out_dir, policy)

    return _run


def _lo_log(msg: str) -> None:
    logger.info(f"  [LibreOffice] {msg}")


def run_libreoffice(run_once, *, timeout_s: float, what: str, settings: DetChecksSettings | None = None):
    """run_once(timeout) - one LibreOffice run with that timeout - under the machine-wide memory guard
    configured in det_checks.libreoffice_* (detchecks/core/lo_guard.py): the lock, the memory wait,
    the retries with the timeout doubled. Used by every LibreOffice run of a grading outside the
    detchecks pipeline (the answer check, --run-calculation); the pipeline itself gets the same
    settings through its RecalcPolicy. Raises detchecks.errors.LibreOfficeUnavailable (retry_later)
    when the memory wait times out or every try failed."""
    from detchecks.core import lo_guard

    settings = settings or load_settings()
    return lo_guard.run_guarded(run_once, timeout_s=timeout_s, settings=settings.guard_settings(), what=what,
                                log=_lo_log)


def retry_later_ids(results) -> list:
    """Attempt ids of failed gradings whose cause is the machine (retry_later: LibreOffice could not run
    now), in result order."""
    return [r.get("attempt_id") for r in results or [] if not r.get("success") and r.get("retry_later")]


def retry_later_report(results) -> list:
    """The end-of-run lines (empty when none) that list the attempts LibreOffice could not grade now, so
    the operator re-runs them when the machine has memory to spare."""
    ids = retry_later_ids(results)
    if not ids:
        return []
    return [f"LIBREOFFICE: {len(ids)} attempt(s) NOT graded because LibreOffice could not run (the machine did "
            f"not have the memory within det_checks.libreoffice_max_wait_minutes, the LibreOffice lock stayed busy longer "
            f"than det_checks.libreoffice_max_lock_wait_minutes, or every retry crashed / timed "
            f"out) - no grading row was written.",
            f"  Re-run them when the machine has memory to spare: --attempt-ids "
            f"{' '.join(str(i) for i in ids)}"]


def _reap_libreoffice(workdir: Path) -> list[int]:
    """Kill any LibreOffice process still running on one of this grading's private profiles.

    detchecks' libreoffice_recalc already kills its run on a timeout, an exception or an
    interrupt, runs soffice under a watchdog that kills it when the grader dies, and installs a
    SIGTERM handler (detchecks/core/recalc.py); this sweep is the last net when the checks return
    or raise. It matches the profile folder workdir/_lo_profiles/ in a command line as a raw
    path or as the percent-encoded file URL soffice is given (a task folder name with a space or
    '%'), and kills each match with everything below it. The profiles live under this attempt's
    own workdir, so the match can only hit this grading's processes - never another job's."""
    from detchecks.core import recalc as det_recalc

    marker = Path(workdir) / "_lo_profiles"
    killed = det_recalc.reap_profiles(str(marker))
    if killed:
        logger.warning(f"  [det_checks] killed leftover LibreOffice process(es) {killed} on {marker}{os.sep}")
    return killed


def _remove_recalc_dir(workdir: Path) -> None:
    """Delete this grading's det_checks_recalc/ (LibreOffice copies, profiles) once the checks
    have returned or raised, whichever driver ran them. det_checks.json lives in the task folder
    and stays."""
    workdir = Path(workdir)
    if workdir.name != RECALC_DIRNAME or not workdir.exists():
        return
    shutil.rmtree(workdir, ignore_errors=True)
    if workdir.exists():
        logger.warning(f"  [det_checks] could not remove {workdir}")


def _install_termination_reaper() -> None:
    """detchecks' SIGTERM / SIGHUP handler (kill this process's LibreOffice runs, then die of the
    signal as before). Main thread only (Python's rule): startup_check runs there in every
    driver before any worker starts, and judge.py grades in the main thread."""
    from detchecks.core import recalc as det_recalc

    det_recalc.install_termination_reaper()


# --------------------------------------------------------------------------- legacy .xls deliveries
XLS_NOT_CONVERTED_NOTE = ("a legacy .xls delivery is graded only with --run-calculation, on LibreOffice's .xlsx "
                          "conversion, as judge v12 grades it; without --run-calculation judge v12 fails too")
XLS_VALUES_NOTE = ("values from LibreOffice's .xlsx conversion of the delivered legacy .xls (LibreOffice recalculates "
                   "every formula of an .xls it loads); libreoffice_s is the conversion, no second LibreOffice run")


def legacy_xls(path) -> bool:
    """True when `path` holds a legacy binary Excel workbook (.xls): an OLE2 compound file whose directory
    lists a Workbook (BIFF8) or Book (BIFF5) stream - told from the bytes as detchecks.core.package
    classifies a file, never from the name (every delivery is staged as ai_attempt.xlsx). False for a zip
    package (.xlsx, .xlsm, ...), an encrypted package (OLE2 EncryptedPackage), an OLE2 file without a
    workbook stream or whose directory cannot be read, and a missing file: those are graded (or refused)
    exactly as before."""
    from detchecks.core.package import OLE2_MAGIC, Package

    try:
        with open(path, "rb") as fh:
            if fh.read(len(OLE2_MAGIC)) != OLE2_MAGIC:
                return False
        pkg = Package.open(str(path))
    except (OSError, GradingError):
        return False
    try:
        return pkg.format == "xls" and any(s.lower() in ("workbook", "book") for s in pkg.ole_streams or ())
    finally:
        pkg.close()


def _converted_copy_runner(copy: Path):
    """RecalcPolicy.lo_runner for grading LibreOffice's .xlsx conversion of a delivered .xls: LibreOffice
    recalculated every formula while converting it (it recalculates any .xls it loads), so the
    recalculation pipeline takes the values from the copy itself instead of starting LibreOffice again.
    Asked for any other file, it raises (an internal error, loud)."""

    def _run(src, out_dir, policy):
        if Path(src).resolve() != copy.resolve():
            raise GradingError(f"internal: the recalculation pipeline asked to recalculate {src}, not the converted "
                               f"copy {copy}")
        return str(copy)

    return _run


def _grade_xls_conversion(attempt: Path, selected: list, task_meta: dict, policy, block: dict) -> dict:
    """Grade a delivered legacy .xls as judge v12 does with --run-calculation (module doc): File extension
    (.xlsx) (77) on the delivered file; every other selected check on LibreOffice's .xlsx conversion, made
    by the pipeline's own LibreOffice step (policy.lo_runner: the memory guard, retries, watchdog and the
    test-run size limit) into <workdir>/xls_converted/, values from that copy. Fills `block`
    (det_checks.json "xls_conversion"). Returns the verdicts; both parts always run, then one GradingError
    names every check that could not grade (retry_later when LibreOffice could not run now)."""
    from detchecks import api as det_api

    def key(no):
        return "/".join(DET_CHECK_NAMES[no])

    on_delivered = [n for n in selected if n == _FILE_EXTENSION_CHECK]
    on_copy = [n for n in selected if n != _FILE_EXTENSION_CHECK]
    verdicts, failures = {}, {}
    retry_later = False
    if on_delivered:
        try:
            verdicts.update(det_api.grade(str(attempt), checks=on_delivered, task_meta=task_meta))
        except GradingError as e:
            failures.update(e.failures or {key(n): str(e) for n in on_delivered})
            verdicts.update(e.verdicts or {})
            retry_later = retry_later or e.retry_later
    copy = None
    if on_copy:
        t0 = time.perf_counter()
        try:
            copy = Path(policy.lo_runner(str(attempt), str(Path(policy.workdir) / XLS_COPY_DIRNAME), policy))
        except GradingError as e:             # LibreOffice failed or could not run now, or the test-run size limit
            msg = f"no .xlsx conversion of the delivered legacy .xls {attempt}: {e}"
            failures.update({key(n): msg for n in on_copy})
            retry_later = retry_later or e.retry_later
        else:
            block.update(converted=True, copy=str(copy), copy_bytes=copy.stat().st_size,
                         copy_sha256=_file_sha256(copy), conversion_s=round(time.perf_counter() - t0, 2))
            logger.info(f"  [det_checks] legacy .xls converted to .xlsx by LibreOffice in {block['conversion_s']} s: "
                        f"{len(on_copy)} check(s) grade that copy ({copy})")
            reuse = dataclasses.replace(policy, lo_runner=_converted_copy_runner(copy))
            try:
                verdicts.update(det_api.grade(str(copy), checks=on_copy, task_meta=task_meta, recalc=reuse))
            except GradingError as e:
                failures.update(e.failures or {key(n): str(e) for n in on_copy})
                verdicts.update(e.verdicts or {})
                retry_later = retry_later or e.retry_later
    for no in selected:                       # which file each verdict graded (a failed run's finished ones too)
        stats = (verdicts.get(key(no)) or {}).get("stats")
        if isinstance(stats, dict):
            stats["graded_on"] = GRADED_ON_DELIVERED if no == _FILE_EXTENSION_CHECK else GRADED_ON_COPY
            values = stats.get("values")
            if copy is not None and isinstance(values, dict) and XLS_VALUES_NOTE not in (values.get("notes") or []):
                values["notes"] = list(values.get("notes") or []) + [XLS_VALUES_NOTE]
                values["timings"] = {**(values.get("timings") or {}), "libreoffice_s": block["conversion_s"]}
    if failures:
        where = f"LibreOffice's .xlsx conversion {copy}" if copy is not None else "no .xlsx conversion"
        raise GradingError("\n".join([f"{len(failures)} check(s) could not grade the delivered legacy .xls {attempt} "
                                      f"({where}): {', '.join(failures)}"] + [f"  - {m}" for m in failures.values()]),
                           check=next(iter(failures)), path=str(attempt), failures=failures, verdicts=verdicts,
                           retry_later=retry_later)
    return verdicts


def _xls_note(block: dict | None) -> str | None:
    """The line a failure message adds for a legacy .xls delivery (None for any other file)."""
    if not block:
        return None
    if not block.get("run_calculation"):
        return (f"(the delivered workbook is a legacy binary .xls, staged as {ATTEMPT_FILENAME}: "
                f"{XLS_NOT_CONVERTED_NOTE} - re-run with --run-calculation)")
    return (f"(the delivered workbook is a legacy binary .xls: graded with --run-calculation on LibreOffice's .xlsx "
            f"conversion, {label(_FILE_EXTENSION_CHECK)} on the delivered file)")


def _delivered_text(delivered) -> str:
    return f"delivered as {delivered!r}" if delivered else "delivered file name unknown"


def read_origin(task_folder) -> tuple:
    """(sidecar, problem) for task_folder/_attempt_origin.json: problem None (it names the
    delivered file), "missing" (no sidecar) or "malformed: <why>" (unreadable or not JSON, a JSON
    value that is not an object, or no non-empty string 'original_filename'). The sidecar dict
    is returned whenever it is an object."""
    p = Path(task_folder) / workbook_properties.ORIGIN_FILENAME
    if not p.exists():
        return None, "missing"
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, ValueError) as e:
        return None, f"malformed: not readable as JSON ({type(e).__name__}: {str(e)[:120]})"
    if not isinstance(data, dict):
        return None, f"malformed: a JSON {type(data).__name__}, not an object"
    name = data.get("original_filename")
    if "original_filename" not in data:
        return data, "malformed: no 'original_filename'"
    if not isinstance(name, str) or not name.strip():
        return data, f"malformed: 'original_filename' is {name!r}, not a file name"
    return data, None


def _origin_note(task_folder: Path, problem: str | None) -> str:
    sidecar = Path(task_folder) / workbook_properties.ORIGIN_FILENAME
    if problem == "missing":
        why = f"no {workbook_properties.ORIGIN_FILENAME} in {task_folder}"
    else:
        why = f"{sidecar} is {problem}"
    return (f"({why}: the delivered file name is unknown, so {label(_FILE_EXTENSION_CHECK)} judges the format "
            f"from the file's content - maintainer 2026-10-05: every attempt graded)")


def _failure_message(e: GradingError, attempt: Path, delivered, task_folder: Path,
                     selected: list, settings: DetChecksSettings, origin_problem: str | None = None,
                     xls_block: dict | None = None) -> str:
    failures = dict(getattr(e, "failures", None) or {})
    head = (f"deterministic checks could not grade {attempt} ({_delivered_text(delivered)}): "
            f"{len(failures) or 'one or more'} of {len(selected)} check(s) failed; no fallback "
            f"(judge v13), the grading stops before the LLM call")
    lines = [head]
    for key, msg in failures.items():
        lines.append(f"  - {label_for_key(key)}: {msg}")
    if not failures:
        lines.append(f"  - {e}")
    key77 = "/".join(DET_CHECK_NAMES[_FILE_EXTENSION_CHECK])
    if delivered is None and key77 in failures:
        lines.append("  " + _origin_note(task_folder, origin_problem or "missing"))
    if _xls_note(xls_block):
        lines.append("  " + _xls_note(xls_block))
    if any("LibreOffice" in str(m) for m in (list(failures.values()) or [e])):
        lines.append(
            f"  (LibreOffice binary: {settings.libreoffice_path}, from project_configs.yaml "
            f"paths.libreoffice_path; first timeout {settings.libreoffice_timeout_s:.0f} s, "
            f"det_checks.libreoffice_timeout_seconds; {settings.size_text()}, det_checks.libreoffice_max_mb; "
            f"{settings.guard_text()}, det_checks.libreoffice_*)")
    if getattr(e, "retry_later", False):
        lines.append("  RETRY LATER: LibreOffice could not run now (memory, the lock, or every retry crashed / timed out); "
                     "this attempt is not graded - re-run it when the machine has memory to spare")
    return "\n".join(lines)


# --------------------------------------------------------------------------- run


def run_det_checks(task_folder, *, rubric_path, weights_path, mode: str | None = None,
                   run_calculation: bool = False) -> DetChecksRun:
    """Grade the applicable configured checks on task_folder/ai_attempt.xlsx (module doc).

    run_calculation: the driver's --run-calculation. It matters for a legacy .xls delivery only, which
    is then graded on LibreOffice's .xlsx conversion as judge v12 grades it (module doc, "Legacy .xls
    deliveries"); without it such a delivery fails loudly, as in v12. Nothing else depends on it.

    Returns a DetChecksRun; raises DetChecksError (no fallback) when a check cannot grade the
    file, DetChecksConfigError on a config/rubric/registry mismatch, and SuitabilityError where
    the judge itself would. Writes task_folder/det_checks.json in every case: the verdicts, the
    "off" / "no_applicable_checks" record, or status "error" with the stage that failed
    (config, gate, grade) and the error - whatever the exception, which is then re-raised as is."""
    t0 = time.perf_counter()
    task_folder = Path(task_folder)
    artefact = task_folder / ARTEFACT_FILENAME
    ctx = {"stage": "config", "written": False,
           "record": {"status": None, "mode": mode, "file": {"path": str(task_folder / ATTEMPT_FILENAME)}}}
    try:
        return _run_det_checks(task_folder, artefact, rubric_path, weights_path, mode, t0, ctx,
                               bool(run_calculation))
    except BaseException as e:
        if not ctx["written"]:              # the grade stage writes its own, fuller record
            rec = dict(ctx["record"])
            rec.update(status="error", stage=ctx["stage"], error=f"{type(e).__name__}: {e}",
                       seconds=round(time.perf_counter() - t0, 3))
            _write_artefact(artefact, rec)
        raise


def _run_det_checks(task_folder: Path, artefact: Path, rubric_path, weights_path, mode, t0,
                    ctx: dict, run_calculation: bool = False) -> DetChecksRun:
    settings = load_settings()
    mode = resolve_mode(mode, settings)
    ctx["record"]["mode"] = mode
    attempt = task_folder / ATTEMPT_FILENAME
    if mode == "off":
        summary = {"status": "off", "mode": "off"}
        _write_artefact(artefact, {**summary, "note": "det_checks off: the LLM decides every check"})
        ctx["written"] = True
        logger.info("  [det_checks] off: nothing graded, the LLM decides every check")
        return DetChecksRun("off", {}, summary, artefact)

    from detchecks import api as det_api
    from detchecks.core.recalc import RecalcPolicy

    rubric, rubric_src = _judge_rubric(task_folder, rubric_path)
    ctx["record"]["rubric"] = str(rubric_src)
    plan = configured_checks(settings, rubric, rubric_src)
    ctx["stage"] = "gate"
    weights = json.loads(Path(weights_path).read_text(encoding="utf-8"))
    scored, suitability = scored_checks(task_folder, rubric, weights)
    selected = [no for no in sorted(plan) if DET_CHECK_NAMES[no] in scored]
    not_applicable = [no for no in sorted(plan) if no not in selected]

    origin, origin_problem = read_origin(task_folder)
    delivered = None if origin_problem else origin["original_filename"]
    task_meta = {"requires_external_links": False}
    if delivered:
        task_meta["delivered_filename"] = delivered
    elif origin_problem:                    # 77 then judges the format from the content (maintainer 2026-10-05)
        task_meta["delivered_filename_problem"] = origin_problem
    if origin_problem and _FILE_EXTENSION_CHECK in selected:
        logger.warning(f"  [det_checks] {_origin_note(task_folder, origin_problem)}")
    elif origin_problem and origin_problem != "missing":
        logger.warning(f"  [det_checks] {task_folder / workbook_properties.ORIGIN_FILENAME} is {origin_problem}; "
                       f"{label(_FILE_EXTENSION_CHECK)} is not graded for this task, so nothing depends on it")
    workdir = task_folder / RECALC_DIRNAME
    policy = RecalcPolicy(
        workdir=str(workdir),
        libreoffice_path=settings.libreoffice_path,
        lo_timeout_s=settings.libreoffice_timeout_s,
        excel_allowed=settings.excel_recalc,
        lo_runner=_size_guarded_libreoffice(settings.libreoffice_max_mb),
        lo_retries=settings.libreoffice_retries,
        lo_min_free_pct=settings.libreoffice_min_free_pct,
        lo_max_wait_s=settings.libreoffice_max_wait_s,
        lo_max_lock_wait_s=settings.libreoffice_max_lock_wait_s,
        lo_lock_path=settings.libreoffice_lock_path or None,
        lo_log=_lo_log,
    )
    record = {
        "status": None,
        "mode": mode,
        "code_sha": code_sha(),
        "file": {
            "path": str(attempt),
            "bytes": attempt.stat().st_size if attempt.exists() else None,
            "sha256": _file_sha256(attempt),
            "delivered_filename": delivered,
            "origin_sidecar": origin_problem is None,
            "origin_problem": origin_problem,
        },
        "config": {
            "live": list(settings.live),
            "recorded_only": list(settings.recorded_only),
            "excel_recalc": settings.excel_recalc,
            "libreoffice_path": settings.libreoffice_path,
            "libreoffice_timeout_s": settings.libreoffice_timeout_s,
            "libreoffice_max_mb": settings.libreoffice_max_mb,
            "libreoffice_retries": settings.libreoffice_retries,
            "libreoffice_min_free_pct": settings.libreoffice_min_free_pct,
            "libreoffice_max_wait_s": settings.libreoffice_max_wait_s,
            "libreoffice_lock_path": policy.guard_settings().lock(),
            "libreoffice_max_lock_wait_s": settings.libreoffice_max_lock_wait_s,
            "workdir": str(workdir),
        },
        "rubric": str(rubric_src),
        "suitability": suitability,
        "task_meta": task_meta,
        "graded": selected,
        "not_applicable": not_applicable,
    }
    ctx["record"] = record
    ctx["stage"] = "grade"
    base_summary = {"mode": mode, "code_sha": record["code_sha"], "delivered_filename": delivered,
                    "graded": selected, "not_applicable": not_applicable, "artefact": ARTEFACT_FILENAME}

    if not selected:
        record["status"] = "no_applicable_checks"
        record["seconds"] = round(time.perf_counter() - t0, 3)
        _write_artefact(artefact, record)
        ctx["written"] = True
        logger.info(f"  [det_checks] mode={mode}: no configured check is applicable for this task "
                    f"({len(not_applicable)} not applicable)")
        summary = {"status": "no_applicable_checks", **base_summary, "checks": {},
                   "seconds": record["seconds"]}
        return DetChecksRun(mode, {}, json_safe(summary), artefact)

    logger.info(f"  [det_checks] mode={mode}: grading {len(selected)} check(s) on {attempt.name} "
                f"({_delivered_text(delivered)}); not applicable: {[label(n) for n in not_applicable] or 'none'}")
    _install_termination_reaper()
    xls_block = None                          # det_checks.json "xls_conversion": a legacy .xls delivery only
    try:
        if legacy_xls(attempt):
            xls_block = {"delivered_format": "xls", "run_calculation": run_calculation, "converted": False}
            record["xls_conversion"] = xls_block
            if run_calculation:
                on_copy = [n for n in selected if n != _FILE_EXTENSION_CHECK]
                xls_block.update(converted_by="libreoffice", graded_on_copy=on_copy,
                                 graded_on_delivered=[n for n in selected if n == _FILE_EXTENSION_CHECK])
                if not on_copy:
                    xls_block["note"] = (f"only {label(_FILE_EXTENSION_CHECK)} is graded for this task, on the "
                                         f"delivered file: nothing to convert")
                logger.info(f"  [det_checks] the delivered workbook is a legacy binary .xls: with --run-calculation "
                            f"it is graded, as judge v12 does, on LibreOffice's .xlsx conversion "
                            f"({label(_FILE_EXTENSION_CHECK)} on the delivered file)")
            else:
                xls_block["note"] = XLS_NOT_CONVERTED_NOTE
                logger.warning(f"  [det_checks] the delivered workbook is a legacy binary .xls and this run has no "
                               f"--run-calculation: {XLS_NOT_CONVERTED_NOTE}")
        if xls_block and run_calculation:
            verdicts = _grade_xls_conversion(attempt, selected, task_meta, policy, xls_block)
        else:
            verdicts = det_api.grade(str(attempt), checks=selected, task_meta=task_meta, recalc=policy)
        missing = sorted(set("/".join(DET_CHECK_NAMES[n]) for n in selected) - set(verdicts))
        if missing:
            raise GradingError(f"detchecks returned no verdict for {missing} on {attempt}",
                               path=str(attempt), failures={k: "no verdict returned" for k in missing})
    except GradingError as e:
        record["status"] = "error"
        record["stage"] = "grade"
        record["error"] = str(e)
        record["failures"] = dict(e.failures or {})
        record["finished_verdicts"] = dict(e.verdicts or {})
        record["retry_later"] = bool(getattr(e, "retry_later", False))
        record["seconds"] = round(time.perf_counter() - t0, 3)
        _write_artefact(artefact, record)
        ctx["written"] = True
        msg = _failure_message(e, attempt, delivered, task_folder, selected, settings, origin_problem, xls_block)
        logger.error(f"  [det_checks] {msg}")
        raise DetChecksError(msg, check=e.check, path=str(attempt), failures=e.failures,
                             verdicts=e.verdicts, retry_later=bool(getattr(e, "retry_later", False))) from e
    except Exception as e:  # noqa: BLE001 - anything else is just as loud, with the artefact written
        record["status"] = "error"
        record["stage"] = "grade"
        record["error"] = f"{type(e).__name__}: {e}"
        record["seconds"] = round(time.perf_counter() - t0, 3)
        _write_artefact(artefact, record)
        ctx["written"] = True
        msg = (f"deterministic checks could not grade {attempt} ({_delivered_text(delivered)}): "
               f"{type(e).__name__}: {e}; no fallback (judge v13), the grading stops before the LLM call")
        logger.error(f"  [det_checks] {msg}")
        raise DetChecksError(msg, path=str(attempt)) from e
    finally:
        _reap_libreoffice(workdir)
        _remove_recalc_dir(workdir)

    all_stats = {no: json_safe(verdicts["/".join(DET_CHECK_NAMES[no])].get("stats") or {}) for no in selected}
    # the recalculation block: one per grading (detchecks attaches the same one to every value check)
    values = next((s["values"] for s in all_stats.values() if isinstance(s.get("values"), dict)), None)
    harness_verdicts, checks_block = {}, {}
    for no in selected:
        key = "/".join(DET_CHECK_NAMES[no])
        v = verdicts[key]
        live = plan[no] and bool(v.get("live", True))
        stats = all_stats[no]
        mistakes = json_safe(list(v.get("mistakes") or []))
        entry = {
            "engine": "harness" if live else "llm",
            "decision": v["decision"],
            "summary": str(v.get("summary") or ""),
            "mistakes": mistakes,
            "fallback_reason": None if live else (v.get("live_note") or det_api.LIVE_NOTE),
            "family": FAMILY,
            "live": live,
            "check_no": no,
            "n_mistakes": stats.get("n_mistakes", len(mistakes)),
            "stats": _db_stats(stats, values),
        }
        harness_verdicts[key] = entry
        checks_block[key] = {"check_no": no, "engine": entry["engine"], "decision": entry["decision"],
                             "live": live, "n_mistakes": entry["n_mistakes"], "summary": _short(entry["summary"])}

    seconds = round(time.perf_counter() - t0, 3)
    record.update(status="ok", seconds=seconds, values=values, verdicts=verdicts)
    _write_artefact(artefact, record)
    ctx["written"] = True
    summary = {"status": "ok", **base_summary, "values": _db_values(values), "checks": checks_block,
               "db_list_cap": DB_LIST_CAP, "seconds": seconds}
    if xls_block:                             # compact: no local copy path (deleted with det_checks_recalc/)
        summary["xls_conversion"] = {k: xls_block[k] for k in ("delivered_format", "run_calculation", "converted",
                                                               "converted_by", "graded_on_delivered", "conversion_s")
                                     if k in xls_block}
    summary = json_safe(summary)
    fails = [label(e["check_no"]) + ("" if e["live"] else " [recorded only]")
             for e in harness_verdicts.values() if e["decision"] == "fail"]
    graded_on = " on LibreOffice's .xlsx conversion of the delivered .xls" if (xls_block or {}).get("converted") else ""
    logger.info(f"  [det_checks] {len(selected)} graded{graded_on} in {seconds:.1f} s "
                f"(values: {(values or {}).get('source') or 'not needed'}); "
                f"fail: {fails or 'none'}")
    return DetChecksRun(mode, harness_verdicts, summary, artefact)
