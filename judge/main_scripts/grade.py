"""Grade recorded attempts with the MBABench judge.

Attempts come from every `<output_root>/attempts/**/task_attempts.jsonl` a
pipeline run wrote; tasks, goldens, starting files and the rubric suitability
annotation from `<data_root>/tasks/task_id=<N>/`. Each attempt's workbook is
staged into a task folder under judge/scratch/, graded (deterministic checks,
answer check, then the single-pass judge), and recorded as one row in
`<output_root>/gradings/gradings.jsonl` with its files under
`<output_root>/gradings/<grading_id>/`. No database, no object store.

Roots: $MBABENCH_DATA_ROOT / $MBABENCH_OUTPUT_ROOT, else config/config.yaml
local.data_root / local.output_root, else data/ and outputs/ under the repo.

Usage:
    python judge/main_scripts/grade.py --all --dry-run
    python judge/main_scripts/grade.py --all --workers 2
    python judge/main_scripts/grade.py --agent-model-name claude-code/claude-fable-5-1 --task-ids 1 2
    python judge/main_scripts/grade.py --attempt-ids 1759830000123 --regrade
"""

import argparse
import importlib.util
import json
import re
import shutil
import sys
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

# Ensure judge/ directory is in Python path for local imports
_judge_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_judge_root))

from utils import local_store, rubric_suitability, workbook_properties
from utils.answer_check import run_answer_check, summary_block
from utils import det_checks as det_checks_mod
from utils.det_checks import (
    add_det_checks_arg,
    merge_harness_verdicts,
    retry_later_ids,
    retry_later_report,
    run_det_checks,
)
from detchecks.errors import LibreOfficeUnavailable
from utils.excel_utils import find_golden_solution_file
from utils.judge_identity import resolve_judge_identity
from utils.llm_utils import get_client
from utils.local_store import LocalStoreError
from utils.logger import add_log_file, logger, remove_log_file
from utils.misc_utils import (
    load_env_var,
    load_project_configs,
    relative_path_from_project_root,
)

# Import the judge from the sibling module (no __init__.py, so use importlib)
_spec = importlib.util.spec_from_file_location(
    "_judge_module", str(Path(__file__).parent / "judge.py")
)
_judge_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_judge_mod)
single_pass_judge_case = _judge_mod.single_pass_judge_case

SINGLE_PASS_MAX_ROUNDS = int(load_env_var("SINGLE_PASS_MAX_ROUNDS", default=500))

_EXCEL_EXTS = frozenset(local_store.EXCEL_EXTS)


def add_accuracy_check_arg(parser):
    """--accuracy-check harness|llm and --max-forced-rounds.

    Which engine's verdict on the harness-decidable Accuracy checks (Final
    calculation accuracy; the zero-answers case of Deliverable completeness)
    counts in the recorded total. Both engines' verdicts and totals are
    always recorded in scored_results.accuracy_engine, so this only picks
    the number that lands in the row — no re-run is needed to compare.
    """
    parser.add_argument(
        "--accuracy-check",
        dest="accuracy_check",
        choices=["harness", "llm"],
        default="harness",
        help=(
            "Which engine decides the harness-measurable Accuracy checks in "
            "the recorded total: 'harness' (default) = the deterministic "
            "Questions-sheet answer checker; 'llm' = the judge's own verdict. "
            "Both are always recorded."
        ),
    )
    parser.add_argument(
        "--max-forced-rounds",
        dest="max_forced_rounds",
        type=int,
        default=None,
        help=(
            "Cap on forced-finalization rounds after the main round budget is "
            "exhausted (default: config single_pass.max_forced_rounds). Smoke "
            "tests set this low along with --max-tool-rounds so a capped run "
            "cannot keep spending."
        ),
    )


# ---------------------------------------------------------------------------
# Task folder staging
# ---------------------------------------------------------------------------


def extract_file_refs(refs):
    """(name, absolute path) pairs from a resolved file column
    (local_store.build_attempt gives [{name, path}])."""
    out = []
    for item in refs or []:
        if isinstance(item, dict):
            path = str(item.get("path") or "")
            if path:
                out.append((item.get("name") or Path(path).name, path))
        elif isinstance(item, str) and item:
            out.append((Path(item).name, item))
    return out


def copy_file(source, dest_path):
    """Copy *source* (an absolute local path) to *dest_path*."""
    src = Path(source)
    if not src.exists():
        raise FileNotFoundError(f"Source file not found: {src}")
    shutil.copy(str(src), str(dest_path))


def setup_task_folder(attempt, scratch_run_dir):
    """Create a task folder under scratch with the layout the judge expects.

        attempt_files        -> first xlsx       -> ai_attempt.xlsx
        task solution_files  -> xlsx             -> solution/<name>.xlsx
                             -> pdf / txt        -> <name> (context)
        task starting_files  -> first xlsx       -> starting/starting_workbook.xlsx
                             -> pdf              -> <name> (context)
        rubric_suitability.json of the task      -> rubric_suitability.json

    Returns the task folder, or None on failure.
    """
    attempt_id = attempt["attempt_id"]
    task_id = attempt["task_id"]
    task_name = attempt["task_name"] or f"task_{task_id}"
    agent_model_name = attempt.get("agent_model_name") or "unknown"
    safe_agent_model = re.sub(r"[^A-Za-z0-9._-]", "_", agent_model_name)

    task_folder = (
        scratch_run_dir
        / f"{task_name}__task_{task_id}__attempt_{attempt_id}__agent_model={safe_agent_model}"
    )
    task_folder.mkdir(parents=True, exist_ok=True)

    # --- ai_attempt.xlsx from the attempt row's attempt_files ---
    xlsx_refs = [
        (n, s)
        for n, s in extract_file_refs(attempt.get("attempt_files"))
        if Path(n).suffix.lower() in _EXCEL_EXTS
    ]
    if not xlsx_refs:
        logger.error(f"  No xlsx found in attempt_files for attempt {attempt_id}")
        logger.error(f"  attempt_files: {attempt.get('attempt_files')}")
        return None

    name, source = xlsx_refs[0]
    try:
        copy_file(source, task_folder / "ai_attempt.xlsx")
        logger.info(f"  ai_attempt.xlsx <- {local_store.display_path(source)}")
    except Exception as e:
        logger.error(f"  Failed to copy attempt file '{name}': {e}")
        return None
    # Provenance sidecar: the staged name is always ai_attempt.xlsx, so the
    # delivered filename / extension (rubric check 77) is only recoverable
    # from here. Best-effort; never blocks grading.
    try:
        (task_folder / workbook_properties.ORIGIN_FILENAME).write_text(
            json.dumps(
                {"original_filename": name, "source": str(source), "attempt_id": attempt_id},
                indent=2,
            ),
            encoding="utf-8",
        )
    except Exception as e:  # noqa: BLE001
        logger.warning(f"  Could not write attempt origin sidecar: {e}")

    # --- Solution xlsx + context files from the task's solution_files ---
    solution_refs = extract_file_refs(attempt.get("task_solution_files"))
    if not solution_refs:
        logger.error(f"  No solution_files listed for task '{task_name}'")
        return None

    solution_dir = task_folder / "solution"
    solution_dir.mkdir(exist_ok=True)
    has_solution_xlsx = False
    for name, source in solution_refs:
        ext = Path(name).suffix.lower()
        if ext in _EXCEL_EXTS:
            try:
                copy_file(source, solution_dir / name)
                logger.info(f"  solution/{name} <- {local_store.display_path(source)}")
                has_solution_xlsx = True
            except Exception as e:
                logger.error(f"  Failed to copy solution file '{name}': {e}")
        elif ext in (".pdf", ".txt"):
            # Context files -> task folder root (the judge finds them there)
            try:
                copy_file(source, task_folder / name)
                logger.info(f"  {name} (context) <- {local_store.display_path(source)}")
            except Exception as e:
                logger.error(f"  Failed to copy context file '{name}': {e}")
    if not has_solution_xlsx:
        logger.error(f"  No solution xlsx staged for task '{task_name}'")
        return None

    # --- Starting workbook from the task's starting_files ---
    # Staged as starting/starting_workbook.xlsx: the subdirectory keeps it
    # invisible to find_golden_solution_file's task-folder scan, and the
    # fixed stem gives a deterministic extraction directory. The judge
    # serves it as read_file source='starting' so guidance rules that turn
    # on inherited-vs-agent-authored content can be checked, not guessed.
    # Best-effort: a task without one grades without it.
    starting_refs = extract_file_refs(attempt.get("task_starting_files"))
    starting_xlsx_refs = [
        (n, s) for n, s in starting_refs if Path(n).suffix.lower() in _EXCEL_EXTS
    ]
    if starting_xlsx_refs:
        name, source = starting_xlsx_refs[0]
        try:
            (task_folder / "starting").mkdir(parents=True, exist_ok=True)
            copy_file(source, task_folder / "starting" / "starting_workbook.xlsx")
            logger.info(f"  starting/starting_workbook.xlsx <- {local_store.display_path(source)}")
        except Exception as e:
            logger.warning(f"  Failed to copy starting workbook '{name}': {e} — grading without it")

    # --- Context PDFs from the task's starting_files ---
    # Starting files often include a "Questions.pdf" while the solution side
    # has a "Questions with Answers.pdf". Both are staged here; the merger
    # dedupes the question-only variant when an answer PDF is present.
    for name, source in starting_refs:
        if Path(name).suffix.lower() != ".pdf":
            continue
        dest = task_folder / name
        if dest.exists():
            logger.info(f"  Skipping starting context '{name}' (already staged)")
            continue
        try:
            copy_file(source, dest)
            logger.info(f"  {name} (context, from starting) <- {local_store.display_path(source)}")
        except Exception as e:
            logger.error(f"  Failed to copy starting context file '{name}': {e}")

    # --- Rubric suitability annotation of the task ---
    # A missing annotation is not an error here: the judge enforces the
    # refusal rule itself (JUDGE_SKIP_SUITABILITY=1 grades ungated).
    try:
        annotation, source = rubric_suitability.load_bundled(task_id)
        annotation["_staging"] = {"source": local_store.display_path(source)}
        (task_folder / rubric_suitability.STAGED_FILENAME).write_text(
            json.dumps(annotation, indent=2), encoding="utf-8"
        )
        logger.info(f"  {rubric_suitability.STAGED_FILENAME} <- {local_store.display_path(source)}")
    except rubric_suitability.SuitabilityError as e:
        logger.warning(f"  {e}")
    except Exception as e:  # noqa: BLE001
        logger.warning(f"  Could not stage the suitability annotation: {e}")

    return task_folder


# ---------------------------------------------------------------------------
# Grading one attempt
# ---------------------------------------------------------------------------


def grade_single_attempt(
    attempt,
    client,
    rubric_path,
    rubric_weight_path,
    template_path,
    model,
    scratch_run_dir,
    nocall=False,
    run_calculation=False,
    cached_solution_csv_dir=None,
    cached_attempt_csv_dir=None,
    cached_starting_csv_dir=None,
    attempt_sheet_name_filter=False,
    ignore_sheets=None,
    max_tool_rounds=SINGLE_PASS_MAX_ROUNDS,
    reasoning_effort=None,
    accuracy_check="harness",
    max_forced_rounds=None,
    det_checks=None,
):
    """Grade a single attempt. Returns a result dict.

    `accuracy_check` ("harness" | "llm"): which engine's verdict on the
    harness-decidable Accuracy checks lands in the recorded total. Both are
    always recorded (scored_results.accuracy_engine).

    `det_checks` ("harness" | "llm" | "off"; None = project_configs.yaml
    det_checks.enabled): the deterministic rubric checks (utils/det_checks.py).
    They grade the delivered ai_attempt.xlsx first - BEFORE the answer check
    and the judge - and are not score-neutral: a check that cannot grade it
    fails this attempt (success False, logged FAILED, no grading row) before
    any LLM spend, like the formula-cache refusal.
    """
    attempt_id = attempt["attempt_id"]
    task_name = attempt["task_name"] or f"task_{attempt['task_id']}"

    logger.info(f"\n{'=' * 60}")
    logger.info(
        f"Grading attempt {attempt_id} for task '{task_name}' "
        f"(task_id={attempt['task_id']})"
    )
    logger.info(f"Agent: {attempt['agent_model_name']} ({attempt['agent_model_type']})")
    logger.info("=" * 60)

    start_time = time.time()

    logger.info("[Setup] Creating task folder...")
    task_folder = setup_task_folder(attempt, scratch_run_dir)
    if task_folder is None:
        return {
            "attempt_id": attempt_id,
            "task_id": attempt["task_id"],
            "success": False,
            "error": "Failed to set up task folder — missing files",
        }
    logger.info(f"  Task folder: {task_folder}")

    # Per-attempt log (its own thread's records under --workers > 1)
    log_path = str(task_folder / "grade.log")
    add_log_file(log_path)

    ac_result = None
    ac_artifact = task_folder / "answer_check.json"
    try:
        # Deterministic rubric checks FIRST, before the answer check and the
        # judge, and NOT score-neutral: no fallback. A DetChecksError lands in
        # the except below exactly like the formula-cache refusal (FAILED,
        # success False, no grading row, the batch continues), before any API
        # spend and before the answer check opens the workbook. LibreOffice
        # runs under the machine-wide memory guard, and when it cannot run now
        # (memory, every retry failed) the error carries retry_later: the
        # attempt is listed at the end of the run to be re-run later. They
        # read the delivered ai_attempt.xlsx, so they must run before
        # prune_workbook_copies. A delivery in the legacy .xls format is graded
        # with --run-calculation on LibreOffice's .xlsx conversion; without it
        # the grading fails.
        det_run = run_det_checks(
            task_folder,
            rubric_path=rubric_path,
            weights_path=rubric_weight_path,
            mode=det_checks,
            run_calculation=run_calculation,
        )

        # Harness answer check, also BEFORE the judge: it only needs the two
        # workbooks, and its verdicts are handed to the scoring layer. The
        # artifact is written into the task folder now and copied into the
        # grading's output_dir after the judge returns. Failures never block
        # grading — the judge's own verdicts then stand.
        try:
            solution_xlsx = find_golden_solution_file(Path(task_folder))
            hardcoded_counts = str(
                load_env_var("SINGLE_PASS_HARDCODED_COUNTS", default="true")
            ).strip().lower() in ("1", "true", "yes")
            ac_result = run_answer_check(
                Path(task_folder) / "ai_attempt.xlsx",
                solution_xlsx,
                output_json_path=ac_artifact,
                hardcoded_counts=hardcoded_counts,
                task_id=attempt["task_id"],   # alternate accepted answers
            )
            logger.info(f"  [answer_check] {summary_block(ac_result)}")
        except LibreOfficeUnavailable:
            # LibreOffice could not run now (machine-wide guard exhausted):
            # fail loudly, re-run later - never an answer check that reads
            # uncomputed answers as unanswered
            raise
        except Exception as e:  # noqa: BLE001 — score-neutral by design
            logger.warning(f"  [answer_check] skipped on error: {e}")
            ac_result = {"status": "error", "error": str(e), "harness_verdicts": {}}
        harness_verdicts = merge_harness_verdicts(
            ac_result.get("harness_verdicts") or {}, det_run.harness_verdicts
        )

        logger.info("[Judge] Running single_pass_judge_case...")
        result = single_pass_judge_case(
            harness_verdicts=harness_verdicts,
            accuracy_engine=accuracy_check,
            det_checks=det_run.for_judge(),
            max_forced_rounds=max_forced_rounds,
            task_folder=str(task_folder),
            client=client,
            rubric_path=rubric_path,
            template_path=template_path,
            rubric_weight_path=rubric_weight_path,
            model=model,
            nocall=nocall,
            run_calculation=run_calculation,
            attempt_model=attempt["agent_model_name"],
            cached_solution_csv_dir=cached_solution_csv_dir,
            cached_attempt_csv_dir=cached_attempt_csv_dir,
            cached_starting_csv_dir=cached_starting_csv_dir,
            attempt_sheet_name_filter=attempt_sheet_name_filter,
            ignore_sheets=ignore_sheets,
            max_tool_rounds=max_tool_rounds,
            reasoning_effort=reasoning_effort,
        )

        if result is None:
            # nocall — nothing to record
            return {
                "attempt_id": attempt_id,
                "task_id": attempt["task_id"],
                "success": True,
                "skipped": True,
                "task_folder": str(task_folder),
            }

        elapsed = time.time() - start_time
        output_dir = Path(result["output_dir"])

        ai_judgement = {}
        ai_judgement_path = output_dir / "ai_judgement.json"
        if ai_judgement_path.exists():
            with open(ai_judgement_path) as f:
                ai_judgement = json.load(f)

        token_tracking = {}
        token_tracking_path = output_dir / "token_tracking.json"
        if token_tracking_path.exists():
            with open(token_tracking_path) as f:
                token_tracking = json.load(f)

        conversation = []
        conversation_path = output_dir / "conversation_messages.json"
        if conversation_path.exists():
            with open(conversation_path) as f:
                conversation = json.load(f)

        # Answer-check artifact rides with the grading's files; its summary
        # lands in scored_results.answer_check (the harness verdicts and the
        # engine that decided each check are in scored_results.accuracy_engine).
        answer_check_summary = summary_block(ac_result) if ac_result else None
        try:
            if ac_artifact.exists():
                shutil.copy(str(ac_artifact), str(output_dir / "answer_check.json"))
        except OSError as e:
            logger.warning(f"  [answer_check] could not copy artifact: {e}")

        scores = {
            "accuracy_grade": result.get("accuracy_score") or 0,
            "formula_grade": result.get("formula_score") or 0,
            "format_grade": result.get("formatting_score") or 0,
            "final_score": result.get("final_score") or 0,
        }
        scored_results = result.get("score_results", {})
        if answer_check_summary is not None and isinstance(scored_results, dict):
            scored_results["answer_check"] = answer_check_summary

        # Warn loudly if expected scores are missing from the judge result.
        # Completeness is defined by the configured weights file (via
        # result["missing_categories"], computed in _finalize_case against
        # the weights' CategoryWeights).
        if result.get("score_results"):
            missing_scores = list(result.get("missing_categories") or [])
            if result.get("final_score") is None:
                missing_scores.append("final_score")
        else:
            missing_scores = ["all categories (no score_results)"]
        if missing_scores:
            logger.warning(
                f"  WARNING: judge returned incomplete scores for attempt "
                f"{attempt_id}! Missing: {missing_scores}. "
                f"This likely means rubric_weight_path was not provided, a "
                f"category was skipped on parse failure, or weights "
                f"calculation failed. Scores will default to 0."
            )

        # Categories whose judge output could not be parsed after all retries.
        parse_failures = result.get("parse_failures") or {}
        hard_parse_failures = [
            cat for cat, info in parse_failures.items() if not info.get("success", True)
        ]
        if hard_parse_failures:
            logger.warning(
                f"  WARNING: judge failed to parse output for categories: "
                f"{hard_parse_failures}. Affected categories contribute 0 to the "
                f"score and this grading will be recorded as failed."
            )

        # Silent-scoring hazards detected inside calculate_scores: unscored
        # checks, empty categories, duplicates, and total_mistakes/len mismatches.
        scoring_warnings = result.get("scoring_warnings") or {}
        has_scoring_warnings = any(scoring_warnings.get(k) for k in scoring_warnings)
        if has_scoring_warnings:
            logger.warning(
                f"  WARNING: scoring warnings for attempt {attempt_id}: "
                f"{ {k: v for k, v in scoring_warnings.items() if v} }. "
                f"This grading will be recorded as failed."
            )

        return {
            "attempt_id": attempt_id,
            "task_id": attempt["task_id"],
            "success": True,
            "task_folder": str(task_folder),
            "output_dir": str(output_dir),
            "scores": scores,
            "scored_results": scored_results,
            "ai_judgement": ai_judgement,
            "conversation": conversation,
            "token_tracking": token_tracking,
            "elapsed_seconds": round(elapsed, 2),
            "cost": token_tracking.get("total_cost", 0),
            "solution_context_reduced": result.get("solution_context_reduced", False),
            "attempt_context_reduced": result.get("attempt_context_reduced", False),
            "context_reduced_details": result.get("context_reduced_details"),
            "parse_failures": result.get("parse_failures"),
            "hard_parse_failures": hard_parse_failures,
            "missing_scores": missing_scores,
            "scoring_warnings": scoring_warnings,
            "has_scoring_warnings": has_scoring_warnings,
            "solution_csv_dir": result.get("solution_csv_dir"),
            "attempt_csv_dir": result.get("attempt_csv_dir"),
            "starting_csv_dir": result.get("starting_csv_dir"),
            # The versions this grading actually ran under.
            "versions": result.get("versions"),
            # Effective reasoning effort (identity pin, or a --reasoning-effort
            # override), recorded as grader_reasoning.
            "judge_reasoning": result.get("judge_reasoning"),
            "grader_identity": result.get("grader_identity"),
        }

    except Exception as e:
        elapsed = time.time() - start_time
        # retry_later: LibreOffice could not run now (memory wait timed out or
        # every retry failed) - listed at the end of the run, to be re-run
        # when the machine has memory to spare
        retry_later = bool(getattr(e, "retry_later", False))
        logger.error(f"  FAILED{' (LibreOffice could not run - re-run later)' if retry_later else ''}: {e}")
        traceback.print_exc()
        return {
            "attempt_id": attempt_id,
            "task_id": attempt["task_id"],
            "success": False,
            "error": str(e),
            "retry_later": retry_later,
            "traceback": traceback.format_exc(),
            "task_folder": str(task_folder),
            "elapsed_seconds": round(elapsed, 2),
        }
    finally:
        remove_log_file(log_path)


# ---------------------------------------------------------------------------
# Scratch hygiene
# ---------------------------------------------------------------------------

_KEEP_IN_OUTPUT_DIR = frozenset({"judge_conversation_logs"})


def _is_workbook_export_dir(path):
    """True for a CSV-export folder of one workbook (no subfolders, has CSVs).

    Used to find the export folders of a FAILED grading, whose result dict
    carries no *_csv_dir keys. Deliberately narrow: judge_conversation_logs
    (json/yaml) and any future results folder are left alone.
    """
    try:
        entries = list(path.iterdir())
    except OSError:
        return False
    return any(p.is_file() and p.suffix == ".csv" for p in entries) and not any(
        p.is_dir() for p in entries
    )


def prune_workbook_copies(result):
    """Delete an attempt's local workbook copies once the attempt is finished.

    Removes the staged attempt/solution/starting workbooks and their CSV
    exports under judge_results/ (hundreds of MB per attempt for the biggest
    workbooks). They are re-copyable inputs and re-derivable exports, and the
    graded bundle itself is under <output_root>/gradings/. Scores, logs, and
    judge conversations stay. Only paths inside the attempt's own task folder
    are touched, so the shared CSV caches are never affected. Returns bytes
    freed. Also det_checks_recalc/, which utils.det_checks.run_det_checks
    already deletes itself when the checks finish (a no-op then; kept as a
    net); det_checks.json stays with the logs.
    """
    task_folder = Path(result["task_folder"]).resolve()
    output_dir = Path(result.get("output_dir") or task_folder / "judge_results")
    targets = [
        result.get("solution_csv_dir"),
        result.get("attempt_csv_dir"),
        result.get("starting_csv_dir"),
        task_folder / "ai_attempt.xlsx",
        task_folder / "solution",
        task_folder / "starting",
        task_folder / det_checks_mod.RECALC_DIRNAME,
    ]
    if output_dir.is_dir():
        targets += [
            d
            for d in output_dir.iterdir()
            if d.is_dir()
            and d.name not in _KEEP_IN_OUTPUT_DIR
            and _is_workbook_export_dir(d)
        ]
    freed = 0
    for target in targets:
        if not target:
            continue
        path = Path(target).resolve()
        if task_folder not in path.parents or not path.exists():
            continue
        try:
            if path.is_dir():
                freed += sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
                shutil.rmtree(path)
            else:
                freed += path.stat().st_size
                path.unlink()
        except OSError as e:
            logger.warning(f"  Could not remove {path}: {e}")
    logger.info(f"  Removed local workbook copies ({freed / 1e6:.0f} MB)")
    return freed


def enforce_cache_cap(cache_base, cap_bytes, min_age_seconds=3600):
    """Evict the oldest entries of a CSV cache directory to keep it under a cap.

    The per-attempt cache is unbounded by nature — one entry per attempt, about
    40 MB each. Entries are whole `attempt_id=…` folders; only ones untouched
    for min_age_seconds are candidates, so a concurrent grader copying a fresh
    entry is never pulled out from under it. Best effort: errors are logged,
    never raised, and a cap of 0 disables eviction. Returns bytes freed.
    """
    cache_base = Path(cache_base)
    if cap_bytes <= 0 or not cache_base.is_dir():
        return 0
    entries, total = [], 0
    for entry in cache_base.iterdir():
        if not entry.is_dir():
            continue
        try:
            size = sum(f.stat().st_size for f in entry.rglob("*") if f.is_file())
            entries.append((entry.stat().st_mtime, size, entry))
        except OSError:
            continue  # evicted by a concurrent grader
        total += size
    if total <= cap_bytes:
        return 0
    freed, cutoff = 0, time.time() - min_age_seconds
    for mtime, size, entry in sorted(entries):
        if total - freed <= cap_bytes:
            break
        if mtime > cutoff:
            continue
        try:
            shutil.rmtree(entry)
            freed += size
        except OSError as e:
            logger.warning(f"  Could not evict cache entry {entry.name}: {e}")
    logger.info(
        f"  Cache {cache_base.name}: {total / 1e9:.1f} GB over the "
        f"{cap_bytes / 1e9:.0f} GB cap — evicted {freed / 1e9:.1f} GB"
    )
    return freed


# ---------------------------------------------------------------------------
# Recording a grading
# ---------------------------------------------------------------------------


def build_grading_row(attempt, result, model):
    """The grading row for a result, with the JSON columns as live objects.

    `id`, `created_at`, `updated_at`, `raw_files_path` and `raw_files` are
    filled by write_grading.
    """
    versions = result.get("versions") or {}
    if versions.get("PROMPT_VERSION") and versions.get("JUDGE_VERSION"):
        PROMPT_VERSION = versions["PROMPT_VERSION"]
        JUDGE_VERSION = versions["JUDGE_VERSION"]
    else:
        PROMPT_VERSION = load_env_var("SINGLE_PASS_PROMPT_VERSION", required=True)
        JUDGE_VERSION = load_env_var("SINGLE_PASS_VERSION", required=True)
    RUBRIC_VERSION = load_env_var("JUDGE_RUBRIC_VERSION", required=True)
    RUBRIC_WEIGHT_VERSION = load_env_var("JUDGE_RUBRIC_WEIGHT_VERSION", required=True)

    if not result.get("scores"):
        raise ValueError("Result missing 'scores' field with grading details")
    for key in ("accuracy_grade", "formula_grade", "format_grade"):
        if key not in result["scores"]:
            raise ValueError(f"Result 'scores' missing expected key: {key}")

    # A grading is "failed" whenever we can't trust the numeric scores:
    #   - the judge raised (result["success"] is False)
    #   - any category's output couldn't be parsed after all retries
    #   - calculate_scores skipped a category and left it as None
    #   - calculate_scores flagged a silent-scoring hazard (unscored/empty/
    #     duplicate/mistake-count-mismatch)
    hard_parse_failures = result.get("hard_parse_failures") or []
    missing_scores = result.get("missing_scores") or []
    scoring_warnings = result.get("scoring_warnings") or {}
    has_scoring_warnings = bool(result.get("has_scoring_warnings"))
    is_failed = (
        not result["success"]
        or bool(hard_parse_failures)
        or bool(missing_scores)
        or has_scoring_warnings
    )
    if not result["success"]:
        failed_reason = result.get("error")
    elif hard_parse_failures:
        failed_reason = f"Parse failed for categories: {', '.join(hard_parse_failures)}"
    elif missing_scores:
        failed_reason = f"Missing scores: {', '.join(missing_scores)}"
    elif has_scoring_warnings:
        warn_keys = [k for k, v in scoring_warnings.items() if v]
        failed_reason = f"Scoring warnings: {', '.join(warn_keys)}"
    else:
        failed_reason = None

    errors_blob = {}
    if result.get("parse_failures"):
        errors_blob["parse_failures"] = result["parse_failures"]
    if has_scoring_warnings:
        errors_blob["scoring_warnings"] = {k: v for k, v in scoring_warnings.items() if v}

    return {
        "id": None,
        "task_id": attempt["task_id"],
        "attempt_id": attempt["attempt_id"],
        "grader_model": model,
        "grader_prompts": result.get("conversation", []),
        "grader_response": result.get("ai_judgement", {}),
        "accuracy_grade": result.get("scores", {}).get("accuracy_grade", 0),
        "formula_grade": result.get("scores", {}).get("formula_grade", 0),
        "format_grade": result.get("scores", {}).get("format_grade", 0),
        "rubric_version": RUBRIC_VERSION,
        "rubric_weight_version": RUBRIC_WEIGHT_VERSION,
        "prompt_version": PROMPT_VERSION,
        "scored_results": result.get("scored_results", {}),
        "time_elapsed_min": round(result.get("elapsed_seconds", 0) / 60, 4),
        "cost": round(result.get("cost", 0), 6),
        "raw_files_path": result.get("raw_files_path", ""),
        "raw_files": result.get("raw_files", []),
        "errors_encountered": errors_blob or None,
        "failed": is_failed,
        "failed_reason": failed_reason,
        "deprecated": False,
        "deprecated_reason": None,
        "solution_context_reduced": result.get("solution_context_reduced", False),
        "attempt_context_reduced": result.get("attempt_context_reduced", False),
        "context_reduced_details": result.get("context_reduced_details") or None,
        "agentic_mode": True,
        "judge_version": JUDGE_VERSION,
        "created_at": None,
        "updated_at": None,
        # Effective reasoning effort (identity-pinned unless overridden).
        "grader_reasoning": result.get("judge_reasoning"),
    }


_write_lock = threading.Lock()


def write_grading(attempt, result, model):
    """Record a grading under `<output_root>/gradings/`; returns its id.

    The judge's files are copied to `gradings/<grading_id>/` (listed in
    raw_files) and one row is appended to gradings/gradings.jsonl.
    """
    with _write_lock:
        grading_id = local_store.new_local_id()
        row = build_grading_row(attempt, result, model)
        output_dir = result.get("output_dir")
        if output_dir and Path(output_dir).is_dir():
            raw_files_path, raw_files = local_store.stage_grading_files(output_dir, grading_id)
            row["raw_files_path"] = raw_files_path
            row["raw_files"] = raw_files
            # Keep the result dict in step so the run summary names the kept
            # copy, not the scratch directory that is about to be pruned.
            result["raw_files_path"] = raw_files_path
            result["raw_files"] = raw_files
            logger.info(f"  Staged {len(raw_files)} files -> {raw_files_path}/")
        else:
            logger.warning("  No output directory to stage; the row has no files")
        row.update(id=grading_id, created_at=local_store.now_iso(), updated_at=None)
        local_store.append_grading(row)
    return grading_id


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def _print_dry_run(attempts, model, skip_ids):
    logger.info(f"\n{'=' * 60}")
    logger.info("DRY RUN — would grade the following attempts:")
    logger.info("=" * 60)
    for a in attempts:
        graded = " [already graded, skipped without --regrade]" if a["attempt_id"] in skip_ids else ""
        logger.info(
            f"  Attempt {a['attempt_id']}: task='{a['task_name']}' "
            f"(id={a['task_id']}), model={a['agent_model_name']}, "
            f"failed={a['agent_failed']}{graded}"
        )
        # The whole point of the local store is which files it picks; a dry
        # run that did not show them could not be checked.
        for label, column, want_first in (
            ("ai_attempt.xlsx", "attempt_files", True),
            ("solution/", "task_solution_files", False),
            ("starting/", "task_starting_files", False),
        ):
            refs = extract_file_refs(a.get(column))
            if want_first:
                refs = [(n, s) for n, s in refs if Path(n).suffix.lower() in _EXCEL_EXTS][:1]
            for name, src in refs:
                exists = "" if Path(src).exists() else "  [MISSING]"
                logger.info(f"      {label:<16} {local_store.display_path(src)}{exists}")
        suit = rubric_suitability.bundled_path(a["task_id"])
        exists = "" if suit.exists() else "  [MISSING]"
        logger.info(f"      {'suitability':<16} {local_store.display_path(suit)}{exists}")
    logger.info(f"\nTotal: {len(attempts)} attempts; grader: {model}")
    logger.info(
        f"Gradings would be written to {local_store.display_path(local_store.gradings_dir())}/"
    )


def main(args):
    """Select attempts, grade each (in up to --workers threads), record results."""
    load_project_configs()

    rubric_path = str(relative_path_from_project_root(load_env_var("JUDGE_RUBRIC", required=True)))
    rubric_weight_path = str(
        relative_path_from_project_root(load_env_var("JUDGE_RUBRIC_WEIGHT", required=True))
    )
    template_path = str(
        relative_path_from_project_root(load_env_var("SINGLE_PASS_PROMPT_TEMPLATE", required=True))
    )

    model = args.model
    # Fail fast on an unregistered grader label (also covers --dry-run).
    identity = resolve_judge_identity(model)
    # ...or if the deterministic checks' config does not match the rubric.
    det_checks_mod.startup_check(rubric_path, args.det_checks)

    logger.info(
        f"Data root:   {local_store.data_root()}\n"
        f"Output root: {local_store.output_root()}"
    )

    scratch_base = str(
        relative_path_from_project_root(load_env_var("PATHS_SCRATCH_PATH", default="./scratch"))
    )
    # The uuid suffix makes every process's scratch tree its own: two
    # same-second processes grading the same attempt would otherwise share
    # one task folder and destroy each other's artifacts.
    run_id = f"{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    scratch_run_dir = Path(scratch_base) / "grade_runs" / run_id
    scratch_run_dir.mkdir(parents=True, exist_ok=True)

    run_log_path = str(scratch_run_dir / "run.log")
    add_log_file(run_log_path)
    logger.info(f"Run directory: {scratch_run_dir}")
    logger.info(f"Run log: {run_log_path}")

    try:
        # Select attempts
        try:
            attempts, note = local_store.select_attempts(
                attempt_ids=args.attempt_ids or None,
                task_ids=args.task_ids,
                agent_model_names=args.agent_model_name,
            )
        except LocalStoreError as e:
            logger.error(str(e))
            return
        if not note["searched"]:
            logger.error(
                f"No task_attempts.jsonl under {local_store.display_path(local_store.attempts_dir())}/ "
                f"— run a pipeline first, or point ${local_store.repo_config.OUTPUT_ROOT_ENV} at its output root."
            )
            return
        logger.info(f"Attempt rows read from: {', '.join(note['searched'])}")
        if "total_rows" in note:
            logger.info(
                f"{note['matched']} attempt row(s) matched the selection"
                + (f" (of {note['total_rows']})" if note["matched"] != note["total_rows"] else "")
                + (f"; {note['deprecated']} deprecated skipped" if note["deprecated"] else "")
                + (f"; {note['no_workbook']} without a workbook on disk skipped" if note["no_workbook"] else "")
            )
        if not attempts:
            logger.info("No matching attempts found")
            return

        # Already graded under this grader? Skip unless --regrade.
        skip_ids = set()
        if not args.regrade:
            already = local_store.graded_attempt_ids(model)
            skip_ids = {a["attempt_id"] for a in attempts if a["attempt_id"] in already}
            if skip_ids:
                logger.info(
                    f"{len(skip_ids)} attempt(s) already have a grading under {model} in "
                    f"{local_store.display_path(local_store.gradings_dir() / local_store.GRADINGS_FILENAME)}; "
                    f"skipped (--regrade grades them again)"
                )

        if args.dry_run:
            _print_dry_run(attempts, model, skip_ids)
            return

        attempts = [a for a in attempts if a["attempt_id"] not in skip_ids]
        if not attempts:
            logger.info("Nothing left to grade")
            return

        client = get_client(identity)

        # Persistent caches for extracted CSVs — avoids re-extracting across
        # runs. The suffix is the cache generation; older generations must
        # never be reused, their dirs are left untouched.
        cache_root = Path(scratch_base) / "grade_cache"
        solution_cache_base = cache_root / "solution_csv_cache_v9"
        attempt_cache_base = cache_root / "attempt_csv_cache_v9"
        starting_cache_base = cache_root / "starting_csv_cache_v9"
        for d in (solution_cache_base, attempt_cache_base, starting_cache_base):
            d.mkdir(parents=True, exist_ok=True)
        attempt_filter = args.attempt_sheet_name_filter
        # Namespace attempt cache by filter state — filtered extractions are a
        # strict subset/rename of unfiltered ones, so they must not share a path.
        attempt_cache_suffix = "__sheet_filtered" if attempt_filter else ""

        cache_lock = threading.Lock()
        solution_csv_cache = {}  # task_id -> cached dir path
        attempt_csv_cache = {}  # attempt_id -> cached dir path
        starting_csv_cache = {}  # task_id -> cached dir path

        def _find_cached(index, key, cache_dir, what):
            with cache_lock:
                cached = index.get(key)
            if not cached and cache_dir.exists() and list(cache_dir.glob("*.csv")):
                cached = str(cache_dir)
                with cache_lock:
                    index[key] = cached
                logger.info(f"  Found persistent {what} CSV cache: {cache_dir}")
            return cached

        def _persist(index, key, src_dir, cache_dir, what):
            with cache_lock:
                if key in index:
                    return
            if not cache_dir.exists():
                try:
                    shutil.copytree(src_dir, str(cache_dir))
                    logger.info(f"  Persisted {what} CSVs: {cache_dir}")
                except (FileExistsError, shutil.Error) as e:
                    # Concurrent graders on different attempts of the same
                    # task race here; the cache is shared and identical, so
                    # losing the race is harmless.
                    logger.info(
                        f"  {what} CSV cache written concurrently "
                        f"({e.__class__.__name__}); using it"
                    )
            with cache_lock:
                index[key] = str(cache_dir)

        total = len(attempts)

        def _grade(i, attempt):
            task_id = attempt["task_id"]
            attempt_id = attempt["attempt_id"]
            cached_dir = _find_cached(
                solution_csv_cache, task_id, solution_cache_base / f"task_id={task_id}", "solution"
            )
            cached_attempt_dir = _find_cached(
                attempt_csv_cache, attempt_id,
                attempt_cache_base / f"attempt_id={attempt_id}{attempt_cache_suffix}", "attempt",
            )
            cached_starting_dir = _find_cached(
                starting_csv_cache, task_id, starting_cache_base / f"task_id={task_id}", "starting"
            )
            logger.info(f"\n[{i + 1}/{total}] Processing attempt {attempt_id}...")

            result = grade_single_attempt(
                attempt=attempt,
                client=client,
                rubric_path=rubric_path,
                rubric_weight_path=rubric_weight_path,
                template_path=template_path,
                model=model,
                scratch_run_dir=scratch_run_dir,
                nocall=args.nocall,
                run_calculation=args.run_calculation,
                cached_solution_csv_dir=cached_dir,
                cached_attempt_csv_dir=cached_attempt_dir,
                cached_starting_csv_dir=cached_starting_dir,
                attempt_sheet_name_filter=attempt_filter,
                ignore_sheets=args.ignore_sheets,
                max_tool_rounds=args.max_tool_rounds,
                reasoning_effort=args.reasoning_effort,
                accuracy_check=args.accuracy_check,
                max_forced_rounds=args.max_forced_rounds,
                det_checks=args.det_checks,
            )

            finished = result["success"] and not result.get("skipped")
            if finished and result.get("solution_csv_dir"):
                _persist(solution_csv_cache, task_id, result["solution_csv_dir"],
                         solution_cache_base / f"task_id={task_id}", "solution")
            if finished and result.get("starting_csv_dir"):
                _persist(starting_csv_cache, task_id, result["starting_csv_dir"],
                         starting_cache_base / f"task_id={task_id}", "starting")
            if finished and result.get("attempt_csv_dir"):
                _persist(attempt_csv_cache, attempt_id, result["attempt_csv_dir"],
                         attempt_cache_base / f"attempt_id={attempt_id}{attempt_cache_suffix}", "attempt")
                enforce_cache_cap(attempt_cache_base, args.cache_cap_gb * 1_000_000_000)

            # Record the grading whenever we have a usable result with scores,
            # even if it is marked failed (e.g. parse failures) — that way the
            # failure is visible in gradings.jsonl instead of being dropped.
            if not result.get("skipped") and result.get("scores"):
                try:
                    grading_id = write_grading(attempt, result, model)
                    result["grading_id"] = grading_id
                    logger.info(f"  Wrote grading: id={grading_id}")
                except Exception as e:
                    logger.error(f"  Failed to write grading: {e}")
                    result["write_error"] = str(e)

            # Drop the local workbook copies once the attempt is finished:
            # after its grade is stored, or right away if the grading failed —
            # a failure's copies are re-copyable inputs, and its logs, which
            # are what gets read, are kept either way.
            if (
                not args.keep_workbook_copies
                and result.get("task_folder")
                and not result.get("skipped")
            ):
                if result.get("grading_id") or not result.get("success"):
                    prune_workbook_copies(result)
            return result

        if args.workers > 1:
            with ThreadPoolExecutor(max_workers=args.workers, thread_name_prefix="grader") as pool:
                results = list(pool.map(lambda ia: _grade(*ia), enumerate(attempts)))
        else:
            results = [_grade(i, a) for i, a in enumerate(attempts)]

        # Save run summary
        summary = {
            "run_id": run_id,
            "model": model,
            "grader_identity": identity.settings(),
            "rubric_path": rubric_path,
            "template_path": template_path,
            "total_attempts": len(attempts),
            "successful": sum(1 for r in results if r["success"]),
            "failed": sum(1 for r in results if not r["success"]),
            # LibreOffice could not run for these (memory / every retry failed):
            # not graded, no row - re-run them when memory is free
            "retry_later_attempt_ids": retry_later_ids(results),
            "results": results,
        }
        summary_path = scratch_run_dir / "run_summary.json"
        with open(summary_path, "w") as f:
            json.dump(summary, f, indent=2, default=str)

        logger.info(f"\n{'=' * 60}")
        logger.info("GRADING RUN COMPLETE")
        logger.info("=" * 60)
        logger.info(f"  Total: {len(results)}")
        logger.info(f"  Successful: {summary['successful']}")
        logger.info(f"  Failed: {summary['failed']}")
        for line in retry_later_report(results):
            logger.warning(f"  {line}")
        total_cost = sum(r.get("cost", 0) for r in results if r["success"])
        logger.info(f"  Total cost: ${total_cost:.6f}")
        logger.info(f"  Gradings: {local_store.display_path(local_store.gradings_dir())}/")
        logger.info(f"  Run directory: {scratch_run_dir}")
        logger.info(f"  Summary: {summary_path}")
        logger.info("=" * 60)

        for r in results:
            if not r["success"]:
                status = "FAILED - LibreOffice, re-run later" if r.get("retry_later") else "FAILED"
            elif r.get("hard_parse_failures") or r.get("missing_scores"):
                status = "PARSE_FAILED"
            elif r.get("has_scoring_warnings"):
                status = "SCORING_WARN"
            else:
                status = "OK"
            parts = [f"  attempt {r['attempt_id']}: [{status}]"]
            if r.get("scores"):
                s = r["scores"]
                parts.append(
                    f"total={s['final_score']:.2f} "
                    f"A={s['accuracy_grade']:.2f} "
                    f"F={s['formula_grade']:.2f} "
                    f"Fmt={s['format_grade']:.2f}"
                )
            if r.get("hard_parse_failures"):
                parts.append(f"parse_failed={r['hard_parse_failures']}")
            if r.get("missing_scores"):
                parts.append(f"missing_scores={r['missing_scores']}")
            if r.get("has_scoring_warnings"):
                sw = r.get("scoring_warnings") or {}
                counts = {
                    "unscored": sum(len(v) for v in (sw.get("unscored_checks") or {}).values()),
                    "empty_cats": len(sw.get("empty_category_judgements") or []),
                    "dupes": sum(len(v) for v in (sw.get("duplicate_judgements") or {}).values()),
                    "mismatches": len(sw.get("mistake_count_mismatches") or []),
                }
                parts.append("scoring_warn=" + " ".join(f"{k}={v}" for k, v in counts.items() if v))
            if r.get("grading_id"):
                parts.append(f"grading_id={r['grading_id']}")
            if r.get("error"):
                parts.append(r["error"])
            logger.info(" | ".join(parts))
        # repeated last, so the operator cannot miss it
        for line in retry_later_report(results):
            logger.warning(line)

    finally:
        remove_log_file(run_log_path)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    load_project_configs()

    JUDGE_MODEL = load_env_var("JUDGE_DEFAULT_GRADER", required=True)

    parser = argparse.ArgumentParser(
        description="Grade recorded attempts with the MBABench judge.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview everything that would be graded (and which files it would use)
  python judge/main_scripts/grade.py --all --dry-run

  # Grade every attempt with a workbook on disk, two at a time
  python judge/main_scripts/grade.py --all --workers 2

  # One pipeline's attempts on two tasks
  python judge/main_scripts/grade.py --agent-model-name claude-code/claude-fable-5-1 --task-ids 1 2

  # Grade an attempt again under the default grader
  python judge/main_scripts/grade.py --attempt-ids 1759830000123 --regrade

  # Stage the files without calling the model
  python judge/main_scripts/grade.py --all --nocall
""",
    )

    # Selection: --all, or any combination of the filters.
    sel = parser.add_argument_group(
        "selection",
        "Which attempts to grade (at least one flag). Attempts come from every "
        "attempts/**/task_attempts.jsonl under the output root; rows that are "
        "deprecated or whose workbook is not on disk are skipped.",
    )
    sel.add_argument("--all", action="store_true", help="Every attempt whose workbook exists.")
    sel.add_argument("--attempt-ids", type=int, nargs="+", help="Exactly these attempt ids.")
    sel.add_argument(
        "--agent-model-name", type=str, nargs="+", metavar="LABEL",
        help="Only attempts of these agent_model_name labels.",
    )
    sel.add_argument("--task-ids", type=int, nargs="+", help="Only attempts of these tasks.")
    sel.add_argument(
        "--regrade", action="store_true",
        help="Grade attempts that already have a grading row under this grader "
             "(default: skip them).",
    )

    parser.add_argument(
        "--model",
        default=JUDGE_MODEL,
        help=f"Grader label from judge_identities.yaml (default: {JUDGE_MODEL})",
    )
    parser.add_argument(
        "--workers", type=int, default=1,
        help="Attempts graded concurrently (default 1). LibreOffice runs stay "
             "serialised machine-wide by the det_checks memory guard.",
    )
    parser.add_argument(
        "--run-calculation",
        action="store_true",
        help="Run Excel formula calculations via LibreOffice before extracting CSVs",
    )
    parser.add_argument(
        "--attempt-sheet-name-filter",
        dest="attempt_sheet_name_filter",
        action="store_true",
        default=False,
        help="Only keep attempt sheets starting with 'answers_' or 'model_' (off by default).",
    )
    parser.add_argument(
        "--ignore-sheets",
        nargs="+",
        default=[],
        help=(
            "Sheet names to drop from attempt, solution and starting workbook "
            "before grading (case-insensitive). Default: none — every sheet, "
            "including the cover, is served to the judge."
        ),
    )
    parser.add_argument(
        "--max-tool-rounds",
        type=int,
        default=SINGLE_PASS_MAX_ROUNDS,
        help=f"Max tool-calling rounds per grading (default: {SINGLE_PASS_MAX_ROUNDS})",
    )
    parser.add_argument(
        "--reasoning-effort",
        type=str,
        default=None,
        choices=["none", "minimal", "low", "medium", "high", "xhigh", "max"],
        help=(
            "Override the reasoning effort pinned by the grader's identity "
            "(default: the identity's effort). Models without thinking "
            "support may reject the kwarg."
        ),
    )
    add_accuracy_check_arg(parser)
    add_det_checks_arg(parser)

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List the attempts (and files) that would be graded, then stop",
    )
    parser.add_argument(
        "--nocall",
        action="store_true",
        help="Stage files and run the deterministic checks, but skip the model call",
    )
    parser.add_argument(
        "--cache-cap-gb",
        type=float,
        default=5.0,
        help=(
            "Keep the per-attempt CSV cache under this many GB by evicting its "
            "oldest entries (default 5; 0 disables eviction). The per-task "
            "solution and starting caches are bounded by the task count and "
            "are never evicted."
        ),
    )
    parser.add_argument(
        "--keep-workbook-copies",
        action="store_true",
        help=(
            "Keep each attempt's local workbook copies and CSV exports after "
            "its grade is saved. By default they are deleted once the grading "
            "is recorded."
        ),
    )

    args = parser.parse_args()
    if not (args.all or args.attempt_ids or args.agent_model_name or args.task_ids):
        parser.error("select attempts with --all, --attempt-ids, --agent-model-name and/or --task-ids")
    if args.workers < 1:
        parser.error("--workers must be >= 1")

    logger.info(f"Running grade.py with parameters: {json.dumps(vars(args), indent=2, default=str)}")

    main(args)
