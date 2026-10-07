"""Grade one assembled task folder with the MBABench judge.

The folder holds ai_attempt.xlsx (the delivered workbook), solution/<golden>.xlsx,
optionally starting/starting_workbook.xlsx, any context .pdf/.txt files and the
task's rubric_suitability.json (grade.py stages all of this from the downloaded
benchmark; a hand-assembled folder works the same way).
"""

import hashlib
import json
import shutil
import sys
import time
import uuid
from pathlib import Path

# Ensure judge/ directory is in Python path for local imports (mirrors
# grade.py; needed since the pyproject stopped installing `utils`).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openai import OpenAI
from utils.excel_utils import (
    copy_support_files,
    find_golden_solution_file,
    prepare_directory_files,
    process_case_files,
)
from utils import (
    answer_rules,
    anthropic_native,
    formula_cache,
    openai_responses,
    rubric_guidance,
    rubric_suitability,
    workbook_properties,
)
from utils.judge_identity import resolve_judge_identity
from utils.llm_utils import (
    calculate_cost,
    get_client,
    get_native_anthropic_client,
    strip_unsupported_anthropic_images,
    usage_breakdown,
)
from utils.logger import add_log_file, logger, remove_log_file
from utils.misc_utils import (
    CHECK_ORDER as DEFAULT_CHECK_ORDER,
    dump_messages_yaml,
    get_absolute_path,
    load_env_var,
    load_project_configs,
    relative_path_from_project_root,
    str2bool,
)
from utils.prompt_utils import (
    compile_prompt,
    encode_file_to_base64,
    numbered_rubric_checks,
    render_rubric_checks_flat,
)
from utils.trajectory import TrajectoryRecorder
from utils.answer_check import run_answer_check, summary_block
from utils.det_checks import add_det_checks_arg, merge_harness_verdicts, run_det_checks
from detchecks.errors import LibreOfficeUnavailable

### Obtain constants
load_project_configs()
JUDGE_MODEL = load_env_var("JUDGE_DEFAULT_GRADER", required=True)
RUBRIC_MAX_MISTAKES = int(
    load_env_var("JUDGE_RUBRIC_MAX_MISTAKES", default=1),
)
# Single-pass judge: one tool-calling conversation over every applicable
# check. Its limits come from project_configs.yaml single_pass.*.
SINGLE_PASS_CONTEXT_TOKEN_LIMIT = int(
    load_env_var("SINGLE_PASS_CONTEXT_TOKEN_LIMIT", default=850_000),
)
READ_FILE_MAX_CELLS = int(
    load_env_var("SINGLE_PASS_READ_FILE_MAX_CELLS", default=5000),
)
# 500 rounds is deliberately high enough to be effectively unbound; set a
# production value from measured round usage, not by feel (deliberation is
# the one lever measured to improve reproducibility).
SINGLE_PASS_MAX_ROUNDS = int(
    load_env_var("SINGLE_PASS_MAX_ROUNDS", default=500),
)
# Forced-finalization ceiling: a single pass can enter forced finalization
# with 100+ pending checks and models often record ~one per round.
SINGLE_PASS_MAX_FORCED_ROUNDS = int(
    load_env_var("SINGLE_PASS_MAX_FORCED_ROUNDS", default=50),
)

### Custom Errors


class RubricWeightConsistencyError(Exception):
    """Raised when rubric and weight files are inconsistent."""

    pass


### Scoring functions
def validate_rubric_weights_consistency(rubric_path: str, weights_path: str) -> None:
    """Validate that the rubric and weights files are consistent.

    Checks:
    1. All categories in weights exist in rubric and vice versa.
    2. Check names within each category match between rubric and weights.
    3. CategoryWeights has all expected categories and sums to ~1.

    Raises:
        RubricWeightConsistencyError: If any inconsistency is found.
    """
    with open(rubric_path, "r", encoding="utf-8") as f:
        rubric = json.load(f)
    with open(weights_path, "r", encoding="utf-8") as f:
        weights = json.load(f)

    # Categories are defined by the rubric file itself (the 12-category /
    # 132-check rubric_9). The consistency contract is rubric <-> weights
    # agreement, not a fixed category list.
    expected_categories = [k for k in rubric.keys() if k != "CategoryWeights"]
    errors = []

    # Check CategoryWeights
    if "CategoryWeights" not in weights:
        errors.append("Weights file missing 'CategoryWeights' key.")
    else:
        cat_weights = weights["CategoryWeights"][0]
        for cat in expected_categories:
            if cat not in cat_weights:
                errors.append(f"CategoryWeights missing category: {cat}")
        weight_sum = sum(cat_weights.get(cat, 0) for cat in expected_categories)
        if not (0.99 <= weight_sum <= 1.01):
            errors.append(f"CategoryWeights must sum to 1, got {weight_sum:.4f}")

    # Check each category's checks match by name
    for cat in expected_categories:
        rubric_names = []
        if cat in rubric:
            if isinstance(rubric[cat], list):
                rubric_names = [item["name"] for item in rubric[cat] if "name" in item]
        else:
            errors.append(f"Category '{cat}' missing from rubric file.")

        weight_names = []
        if cat in weights:
            if isinstance(weights[cat], list):
                weight_names = [item["name"] for item in weights[cat] if "name" in item]
        else:
            errors.append(f"Category '{cat}' missing from weights file.")

        # Compare names
        rubric_set = set(rubric_names)
        weight_set = set(weight_names)
        in_rubric_only = rubric_set - weight_set
        in_weights_only = weight_set - rubric_set
        if in_rubric_only:
            errors.append(
                f"{cat}: checks in rubric but not in weights: {sorted(in_rubric_only)}"
            )
        if in_weights_only:
            errors.append(
                f"{cat}: checks in weights but not in rubric: {sorted(in_weights_only)}"
            )

    if errors:
        raise RubricWeightConsistencyError(
            "Rubric/weights inconsistency:\n  " + "\n  ".join(errors)
        )


def calculate_check_score(mistakes: int, max_mistakes: int = 5) -> float:
    """Calculate score for a single check, normalized to 0-1 range."""
    raw_score = max(0, max_mistakes - mistakes)
    return raw_score / max_mistakes


def calculate_scores(all_responses: dict, weights: dict, max_mistakes: int = 5) -> dict:
    """Calculate weighted scores from judgement results and weights.

    Matches checks between judgement and weights by the 'name' field.

    Conservative on silent-failure edges:
      - Checks present in weights but missing from judgement score 0 (and are
        recorded in scoring_warnings["unscored_checks"]), rather than defaulting
        to 0 mistakes / full credit.
      - An empty judgement list for a category is recorded in
        scoring_warnings["empty_category_judgements"]; combined with the rule
        above, every weighted check ends up unscored.
      - Duplicate check names keep the worst (max mistakes) rather than
        last-write-wins, and the incident is recorded in
        scoring_warnings["duplicate_judgements"].
      - Mistake counts are taken from len(mistakes); if the model also
        reported `total_mistakes` and it disagrees, the mismatch is recorded
        in scoring_warnings["mistake_count_mismatches"] but the structured
        list wins.

    Args:
        all_responses: Judgement results dict {category: [check_items...]}.
        weights: Weights dict with CategoryWeights and per-category check weights.
        max_mistakes: Maximum mistakes before score is 0 (default 5).

    Returns:
        Dictionary with check_scores, criteria_scores, total_score (0-100),
        and scoring_warnings.
    """
    results = {
        "check_scores": {},
        "criteria_scores": {},
        "total_score": 0.0,
        "scoring_warnings": {
            "unscored_checks": {},
            "empty_category_judgements": [],
            "duplicate_judgements": {},
            "mistake_count_mismatches": [],
            "fail_without_mistakes": [],
        },
    }
    scoring_warnings = results["scoring_warnings"]
    category_weights = weights["CategoryWeights"][0]

    total_score = 0.0

    # Score every category the weights file defines (12 for rubric_9) — the
    # weights file is the scoring contract.
    for category in category_weights.keys():
        category_data = all_responses.get(category, [])
        if not isinstance(category_data, list):
            logger.warning(
                f"  Skipping score for {category}: response is not a list (parse failure?). See category_data: {category_data}"
            )
            continue

        if not category_data:
            scoring_warnings["empty_category_judgements"].append(category)

        # Build name -> mistake count from judgement
        judgement_by_name = {}
        for item in category_data:
            name = item.get("name")
            if not name:
                logger.warning(
                    f"  Skipping item in {category} with missing name: {item}"
                )
                continue

            mistakes_list = item.get("mistakes", [])
            actual_mistakes = (
                len(mistakes_list) if isinstance(mistakes_list, list) else 0
            )
            if str(item.get("decision", "")).strip().lower() == "fail" and actual_mistakes == 0:
                actual_mistakes = max_mistakes
                scoring_warnings["fail_without_mistakes"].append(
                    {"category": category, "name": name}
                )
            if "total_mistakes" in item and item["total_mistakes"] != actual_mistakes:
                scoring_warnings["mistake_count_mismatches"].append(
                    {
                        "category": category,
                        "name": name,
                        "claimed_total_mistakes": item["total_mistakes"],
                        "actual_mistakes_len": actual_mistakes,
                    }
                )
            m = actual_mistakes

            if name in judgement_by_name:
                dupes = scoring_warnings["duplicate_judgements"].setdefault(
                    category, {}
                )
                dupes.setdefault(name, [judgement_by_name[name]]).append(m)
                judgement_by_name[name] = max(judgement_by_name[name], m)
            else:
                judgement_by_name[name] = m

        # Calculate scores for each check in weights
        category_check_scores = {}
        weighted_sum = 0.0
        total_weight = 0.0

        for check_weight in weights.get(category, []):
            check_name = check_weight["name"]
            weight = check_weight["weight"]

            if check_name in judgement_by_name:
                mistakes = judgement_by_name[check_name]
                unscored = False
            else:
                # Not evaluated by the judge — conservative penalty instead of
                # silently crediting as "0 mistakes".
                mistakes = max_mistakes
                unscored = True
                scoring_warnings["unscored_checks"].setdefault(category, []).append(
                    check_name
                )

            check_score = calculate_check_score(mistakes, max_mistakes)
            weighted_score = check_score * weight
            weighted_sum += weighted_score
            total_weight += weight

            category_check_scores[check_name] = {
                "mistakes": mistakes,
                "score": check_score,
                "weight": weight,
                "weighted_score": weighted_score,
                "unscored": unscored,
            }

        category_normalized_score = (
            weighted_sum / total_weight * 100 if total_weight > 0 else 0.0
        )

        cat_weight = category_weights.get(category, 0)
        results["check_scores"][category] = category_check_scores
        results["criteria_scores"][category] = {
            "weighted_sum": weighted_sum,
            "total_weight": total_weight,
            "normalized_score": category_normalized_score,
            "category_weight": cat_weight,
        }

        total_score += category_normalized_score * cat_weight

    results["total_score"] = total_score
    return results


### Main Judge Function


def _delete_ignored_sheet_files(directories, ignore_sheets):
    """Delete CSV/format files for sheets in *ignore_sheets* from each directory.

    Files are named ``<safe_sheet_name>_full.csv`` and
    ``<safe_sheet_name>_additional_format.txt``; matching is case-insensitive
    against the file stem (i.e. the safe sheet name). Applies uniformly to
    fresh extractions and to dirs populated from the persistent CSV cache.
    """
    if not ignore_sheets:
        return
    targets = {s.lower() for s in ignore_sheets}
    for d in directories:
        if not d:
            continue
        d_path = Path(d)
        if not d_path.exists():
            continue
        for csv_path in d_path.glob("*_full.csv"):
            stem = csv_path.name.removesuffix("_full.csv")
            if stem.lower() in targets:
                csv_path.unlink()
                fmt = d_path / f"{stem}_additional_format.txt"
                fmt.unlink(missing_ok=True)
                logger.info(f"  Ignored sheet '{stem}': removed from {d_path}")


def _prepare_case(
    task_folder: str,
    rubric_path: str,
    rubric_weight_path: str = None,
    use_existing: bool = True,
    run_calculation: bool = False,
    cached_solution_csv_dir: str = None,
    cached_attempt_csv_dir: str = None,
    cached_starting_csv_dir: str = None,
    attempt_sheet_name_filter: bool = False,
    ignore_sheets: list[str] | None = None,
) -> dict:
    """Shared setup for the judge workflow: logging, validation, file processing.

    Handles:
    - Cache directory and log file setup
    - Version info loading (single_pass.* / rubric env vars)
    - Rubric/weights validation (Step 0)
    - Case file processing: xlsx to CSV extraction (Step 1)
    - Support file copying

    Returns a dict with all state needed by downstream judge functions.
    """
    cache_dir = (
        Path(load_env_var("PATHS_SCRATCH_PATH", default="scratch"))
        / "judge_cache"
        / f"run_{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
    )
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_log_path = str(cache_dir / "judge.log")
    add_log_file(cache_log_path)
    logger.info(f"Writing logs to temporary cache path: {cache_log_path}")

    task_folder_name = Path(task_folder).name
    task_path = Path(task_folder)

    JUDGE_VERSION = load_env_var("SINGLE_PASS_VERSION", required=True)
    PROMPT_VERSION = load_env_var("SINGLE_PASS_PROMPT_VERSION", required=True)
    RUBRIC_VERSION = load_env_var("JUDGE_RUBRIC_VERSION", required=True)
    RUBRIC_WEIGHT_VERSION = load_env_var("JUDGE_RUBRIC_WEIGHT_VERSION", default=None)
    CHECK_ORDER = load_env_var("JUDGE_CHECK_ORDER", default=DEFAULT_CHECK_ORDER).split(",")

    start_time = time.time()

    # Step 0: Validate rubric/weights
    weights_data = None
    if rubric_weight_path:
        logger.info("\n[Step 0] Validating rubric/weights consistency...")
        try:
            validate_rubric_weights_consistency(rubric_path, rubric_weight_path)
            logger.info("  Rubric and weights files are consistent.")
            with open(rubric_weight_path, "r", encoding="utf-8") as f:
                weights_data = json.load(f)
        except RubricWeightConsistencyError as e:
            logger.info(f"  ERROR: {e}")
            raise

    # Step 1: Process case files
    logger.info("\n[Step 1] Processing case files...")
    try:
        golden_solution_path = find_golden_solution_file(task_path)
        logger.info(f"  Golden solution file: {golden_solution_path.name}")
    except Exception as e:
        logger.info(f"  Error finding golden solution file: {e}")
        raise

    ai_attempt_path = task_path / "ai_attempt.xlsx"
    golden_solution_stem = golden_solution_path.stem

    files_to_process = []
    if not cached_attempt_csv_dir:
        files_to_process.append(str(ai_attempt_path))
    if not cached_solution_csv_dir:
        files_to_process.append(str(golden_solution_path))

    # The starting workbook (what the agent was handed before doing any work)
    # is staged by grade.setup_task_folder as starting/starting_workbook.xlsx
    # — the subdirectory keeps it invisible to find_golden_solution_file's
    # task-folder scan, and the fixed stem gives it a deterministic
    # extraction directory. It is optional: a standalone judge run without
    # it grades exactly as before.
    starting_workbook_path = task_path / "starting" / "starting_workbook.xlsx"
    have_starting = bool(cached_starting_csv_dir) or starting_workbook_path.exists()
    if have_starting and not cached_starting_csv_dir:
        files_to_process.append(str(starting_workbook_path))

    if cached_solution_csv_dir:
        logger.info(f"  Using cached solution CSVs from: {cached_solution_csv_dir}")
    if cached_attempt_csv_dir:
        logger.info(f"  Using cached attempt CSVs from: {cached_attempt_csv_dir}")
    if cached_starting_csv_dir:
        logger.info(f"  Using cached starting CSVs from: {cached_starting_csv_dir}")

    # When a cache is used, the extraction for that file is skipped but
    # `use_existing` must still be True for the remaining file so prior
    # extractions are honored.
    effective_use_existing = (
        True
        if (
            cached_solution_csv_dir
            or cached_attempt_csv_dir
            or cached_starting_csv_dir
        )
        else use_existing
    )
    result = process_case_files(
        files_to_process,
        task_folder,
        use_existing=effective_use_existing,
        run_calculation=run_calculation,
        attempt_sheet_name_filter=attempt_sheet_name_filter,
    )
    output_dir = result["output_dir"]
    workbook_dirs = result.get("workbook_dirs", {})

    if cached_solution_csv_dir:
        dest_solution_dir = output_dir / golden_solution_stem
        if not dest_solution_dir.exists():
            shutil.copytree(cached_solution_csv_dir, str(dest_solution_dir))
        workbook_dirs[golden_solution_stem] = str(dest_solution_dir)

        cached_files = sorted(
            f.name for f in Path(dest_solution_dir).iterdir() if f.is_file()
        )
        logger.info(
            f"  Copied {len(cached_files)} cached solution CSV files to: "
            f"{dest_solution_dir}"
        )
        for fname in cached_files:
            logger.info(f"    {fname}")

    if cached_attempt_csv_dir:
        dest_attempt_dir = output_dir / "ai_attempt"
        if not dest_attempt_dir.exists():
            shutil.copytree(cached_attempt_csv_dir, str(dest_attempt_dir))
        workbook_dirs["ai_attempt"] = str(dest_attempt_dir)

        cached_files = sorted(
            f.name for f in Path(dest_attempt_dir).iterdir() if f.is_file()
        )
        logger.info(
            f"  Copied {len(cached_files)} cached attempt CSV files to: "
            f"{dest_attempt_dir}"
        )
        for fname in cached_files:
            logger.info(f"    {fname}")

    if cached_starting_csv_dir:
        dest_starting_dir = output_dir / "starting_workbook"
        if not dest_starting_dir.exists():
            shutil.copytree(cached_starting_csv_dir, str(dest_starting_dir))
        workbook_dirs["starting_workbook"] = str(dest_starting_dir)
        logger.info(
            f"  Copied cached starting-workbook CSVs to: {dest_starting_dir}"
        )

    logger.info(f"  Files processed and saved to: {output_dir}")

    copied_files = copy_support_files(
        task_path,
        output_dir,
        default_rubric_path=rubric_path,
    )
    logger.info(f"  Copied {len(copied_files)} support files to output directory")

    ai_attempt_dir = workbook_dirs.get("ai_attempt")
    golden_solution_dir = workbook_dirs.get(golden_solution_stem)
    starting_workbook_dir = workbook_dirs.get("starting_workbook")

    # Drop sheets the caller explicitly asked to ignore (--ignore-sheets;
    # nothing by default — the cover sheet is graded content). Applied
    # after both fresh extraction and cache copy so the artifact on disk
    # reflects exactly what the judge will see.
    _delete_ignored_sheet_files(
        [ai_attempt_dir, golden_solution_dir, starting_workbook_dir], ignore_sheets
    )

    # Detect context file
    context_file_path = None
    context_pdf = output_dir / "context.pdf"
    context_txt = output_dir / "context.txt"
    if context_pdf.exists():
        context_file_path = context_pdf
    elif context_txt.exists():
        context_file_path = context_txt

    rubric_json_path = output_dir / "rubric.json"

    # Uncached-formula pre-flight. The judge reads cached formula results; a
    # workbook saved without calculation reaches it as formulas with no values,
    # which silently guts Accuracy. Refuse rather than produce a grade that
    # looks real. JUDGE_SKIP_FORMULA_CACHE_CHECK=1 grades anyway (recorded).
    # The staged workbooks are passed so the decision rests on their XML,
    # which (unlike the CSVs) tells a calculated empty-string result from a
    # never-calculated cell; the CSV census decides only when no workbook is
    # on disk (cached CSVs, standalone folders).
    logger.info("\n[Step 1b] Formula-cache pre-flight...")
    # --run-calculation re-saves the attempt through LibreOffice into
    # temp_recalculated/ (process_all_worksheets) and extracts the CSVs from
    # THAT copy, so the census must read the same copy: censusing the original
    # here refused every recalculated grading (2026-09-05, the three coding
    # attempts the full run could not grade), which made the flag the error
    # message recommends a no-op.
    census_attempt_xlsx = ai_attempt_path if ai_attempt_path.exists() else None
    recalculated_attempt = task_path / "temp_recalculated" / "ai_attempt.xlsx"
    if run_calculation and recalculated_attempt.exists():
        census_attempt_xlsx = recalculated_attempt
        logger.info(
            f"  [formula_cache] censusing the recalculated attempt workbook: "
            f"{recalculated_attempt}"
        )
    formula_cache_provenance = formula_cache.check_case(
        ai_attempt_dir,
        golden_solution_dir,
        attempt_xlsx=census_attempt_xlsx,
        solution_xlsx=golden_solution_path if golden_solution_path.exists() else None,
    )
    if run_calculation:
        formula_cache_provenance["attempt"]["recalculated_workbook"] = (
            recalculated_attempt.exists()
        )

    return {
        "cache_dir": cache_dir,
        "cache_log_path": cache_log_path,
        "starting_workbook_dir": starting_workbook_dir,
        "task_folder_name": task_folder_name,
        "task_path": task_path,
        "output_dir": output_dir,
        "workbook_dirs": workbook_dirs,
        "golden_solution_path": golden_solution_path,
        "golden_solution_stem": golden_solution_stem,
        "ai_attempt_dir": ai_attempt_dir,
        "golden_solution_dir": golden_solution_dir,
        "weights_data": weights_data,
        "context_file_path": context_file_path,
        "rubric_json_path": rubric_json_path,
        "formula_cache_provenance": formula_cache_provenance,
        "start_time": start_time,
        "versions": {
            "JUDGE_VERSION": JUDGE_VERSION,
            "PROMPT_VERSION": PROMPT_VERSION,
            "RUBRIC_VERSION": RUBRIC_VERSION,
            "RUBRIC_WEIGHT_VERSION": RUBRIC_WEIGHT_VERSION,
        },
        "CHECK_ORDER": CHECK_ORDER,
    }


def _file_md5(path) -> str | None:
    """md5 of a file's bytes, or None if unreadable (trajectory header only)."""
    try:
        with open(path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except OSError:
        return None


def _make_recorder(output_dir, **header_fields) -> TrajectoryRecorder:
    """Build the per-grading trajectory recorder and write its header.

    Enabled by default; set judge.record_trajectory: false (env
    JUDGE_RECORD_TRAJECTORY) to disable. The file lives in judge_results/
    so it is kept with the rest of the grading artifacts.
    """
    enabled = str(
        load_env_var("JUDGE_RECORD_TRAJECTORY", default="true")
    ).strip().lower() in ("1", "true", "yes")
    recorder = TrajectoryRecorder(Path(output_dir) / "trajectory.jsonl", enabled=enabled)
    if enabled:
        recorder.record_header(**header_fields)
    return recorder


def _resolve_category_score(criteria_scores: dict, *names) -> float | None:
    """Return the first named category's normalized score, else None.

    Aliases cover rubric naming drift ("Formula" vs rubric_9 "Formulas").
    """
    for name in names:
        score = criteria_scores.get(name, {}).get("normalized_score")
        if score is not None:
            return score
    return None


def tag_check_graders(score_results: dict) -> dict:
    """Write `grader` ("deterministic" | "llm") on every rubric item of
    score_results["check_scores"] (which grader decided each item must be
    readable from the grading row alone).

    "deterministic": the Python verdict is the one scored - the item is in
    scored_results.accuracy_engine.checks with engine "harness" and counted
    true (the deterministic rubric checks and the answer-check items the
    harness could measure). "llm": every other item, including the
    recorded-only checks and a Python check that fell back to the LLM.
    Idempotent. Returns {"deterministic": n, "llm": n}."""
    applied = ((score_results.get("accuracy_engine") or {}).get("checks")) or {}
    decided = {
        key for key, v in applied.items()
        if isinstance(v, dict) and v.get("engine") == "harness" and v.get("counted")
    }
    counts = {"deterministic": 0, "llm": 0}
    for category, checks in (score_results.get("check_scores") or {}).items():
        for name, item in (checks or {}).items():
            if not isinstance(item, dict):
                continue
            grader = "deterministic" if f"{category}/{name}" in decided else "llm"
            item["grader"] = grader
            counts[grader] += 1
    return counts


def _apply_harness_verdicts(all_responses, harness_verdicts, weights_data):
    """Overlay the harness's measured decisions onto a COPY of the judgement.

    For every "<Category>/<name>" the harness measured (engine == "harness")
    and that the weights file scores (suitability-gated checks are left
    alone), the LLM's item is replaced by the harness item; the LLM's
    decision/summary/mistakes are preserved on the item as `llm_*` so the
    audit trail survives. A check the LLM never recorded is inserted.

    Returns (overlaid_responses, provenance) where provenance maps each
    harness-addressable check to {engine, decision, fallback_reason,
    llm_decision, agreed, ...stats}.

    Two producers feed it: the answer check (utils/answer_check.py) and,
    from judge v13, the deterministic rubric checks (utils/det_checks.py,
    entries carry family "det_checks", live, check_no, n_mistakes, stats).
    A recorded-only deterministic verdict arrives as engine "llm" WITH a
    Python decision: it is recorded (decision, llm_decision, agreed) and
    never overlaid.
    """
    import copy as _copy

    overlaid = _copy.deepcopy(all_responses)
    provenance: dict = {}
    weighted = {
        (cat, cw["name"])
        for cat, entries in (weights_data or {}).items()
        if cat != "CategoryWeights" and isinstance(entries, list)
        for cw in entries
        if isinstance(cw, dict) and "name" in cw
    }
    for key, hv in (harness_verdicts or {}).items():
        cat, _, name = key.partition("/")
        items = overlaid.setdefault(cat, []) if isinstance(overlaid.get(cat, []), list) else None
        llm_item = None
        if items is not None:
            for it in items:
                if isinstance(it, dict) and it.get("name") == name:
                    llm_item = it
                    break
        prov = {
            "engine": hv.get("engine", "llm"),
            "decision": hv.get("decision"),
            "fallback_reason": hv.get("fallback_reason"),
            "llm_decision": (llm_item or {}).get("decision"),
            "llm_mistake_count": len((llm_item or {}).get("mistakes") or []),
        }
        for stat in ("n_questions", "n_match", "n_mismatch", "n_missing",
                     "n_unlocated", "n_answered", "n_hardcoded", "n_unrounded",
                     "rounding_directive", "fraction_correct",
                     "rules_fired", "flags", "hardcoded_counts", "rules_version",
                     # judge v13 deterministic checks (utils/det_checks.py)
                     "family", "live", "check_no", "n_mistakes", "stats"):
            if stat in hv:
                prov[stat] = hv[stat]
        # Agreement wherever the harness has a decision, overlaid or not
        # (judge v13: recorded-only deterministic verdicts are compared too).
        if hv.get("decision") is not None:
            prov["agreed"] = (
                (llm_item or {}).get("decision") == hv.get("decision")
                if llm_item else None
            )
        if hv.get("engine") != "harness":
            provenance[key] = prov
            continue
        if (cat, name) not in weighted:
            prov["engine"] = "llm"
            prov["fallback_reason"] = "check not scored for this task (suitability-gated)"
            provenance[key] = prov
            continue
        if items is None:
            prov["engine"] = "llm"
            prov["fallback_reason"] = "category judgement unparseable"
            provenance[key] = prov
            continue
        new_item = {
            # The LLM's own check id; for a check it never recorded, the
            # rubric number the producer supplies (deterministic checks).
            "check": (llm_item or {}).get("check")
            or (str(hv["check_no"]) if hv.get("check_no") is not None else None),
            "name": name,
            "decision": hv["decision"],
            "summary": hv.get("summary", ""),
            "mistakes": list(hv.get("mistakes") or []),
            "decided_by": "harness",
            "llm_decision": (llm_item or {}).get("decision"),
            "llm_summary": (llm_item or {}).get("summary"),
            "llm_mistakes": list((llm_item or {}).get("mistakes") or []),
        }
        if llm_item is not None:
            idx = items.index(llm_item)
            items[idx] = new_item
        else:
            items.append(new_item)
        provenance[key] = prov
    return overlaid, provenance


def _finalize_case(
    all_responses,
    output_dir,
    weights_data,
    token_tracking,
    model,
    attempt_model,
    task_folder_name,
    golden_solution_files,
    ai_attempt_files,
    context_file_path,
    start_time,
    cache_log_path,
    versions,
    golden_solution_dir=None,
    ai_attempt_dir=None,
    starting_workbook_dir=None,
    parse_failures=None,
    solution_context_reduced=False,
    attempt_context_reduced=False,
    context_reduced_details=None,
    agentic=False,
    auto_routed=False,
    reasoning_effort=None,
    grader_identity=None,
    recorder=None,
    suitability_provenance=None,
    formula_cache_provenance=None,
    harness_verdicts=None,
    accuracy_engine="llm",
    det_checks=None,
):
    """Shared finalization: save judgement, calculate scores, write metadata.

    `harness_verdicts` (judge v6, utils.answer_check.harness_verdicts) carries
    the deterministic answer checker's decisions for the checks it could
    measure. BOTH engines are always scored — `total_score_llm` and
    `total_score_harness` ride in scored_results.accuracy_engine — and
    `accuracy_engine` ("harness" | "llm") picks which one is THE score
    (total_score, criteria_scores) that lands in the grading row. So the
    harness-vs-LLM comparison never needs a second grading run.

    Judge v13: `harness_verdicts` also carries the deterministic rubric
    checks (entries with family "det_checks", utils/det_checks.py) and
    `det_checks` is that run's {"mode", "summary", "artefact"}. Each family
    has its own switch: `accuracy_engine` decides the answer check's
    verdicts, det_checks["mode"] ("harness" | "llm" | "off") the
    deterministic family's. total_score_harness still overlays every live
    measured verdict; the recorded total overlays only the families whose
    switch is "harness" — accuracy_engine.effective is "harness" (all
    counted), "llm" (none) or "mixed" (some), and each check's provenance
    says `counted`. scored_results.det_checks carries the run's summary and
    det_checks.json is copied into output_dir.
    """
    output_dir = Path(output_dir)

    # Seal the trajectory first so trajectory.jsonl.gz is on disk before the
    # caller copies output_dir into the gradings folder.
    if recorder is not None:
        traj_path = recorder.close()
        if traj_path:
            logger.info(f"  Trajectory recorded to: {traj_path}")

    # Warn if expected token_tracking keys are missing
    _expected_tt_keys = {
        "total_message_size",
        "total_message_size_with_images",
        "total_tokens",
        "total_prompt_tokens",
        "total_completion_tokens",
        "total_cost",
        "evaluations",
    }
    _missing_tt = _expected_tt_keys - set(token_tracking.keys())
    if _missing_tt:
        logger.warning(
            f"  token_tracking missing expected keys: {sorted(_missing_tt)}. "
            f"Defaulting to 0 for missing values."
        )

    if golden_solution_files is None:
        logger.warning("  golden_solution_files is None; expected a dict.")
        golden_solution_files = {}
    if ai_attempt_files is None:
        logger.warning("  ai_attempt_files is None; expected a dict.")
        ai_attempt_files = {}

    # Save ai_judgement
    logger.info("\n[Save] Saving AI judgement...")
    ai_judgement_path = output_dir / "ai_judgement.json"
    with open(ai_judgement_path, "w", encoding="utf-8") as f:
        json.dump(all_responses, f, indent=2)
    logger.info(f"  AI judgement saved to: {ai_judgement_path}")

    # Judge v13: the deterministic checks' full record rides with the bundle.
    det_mode = (det_checks or {}).get("mode") or "harness"
    det_artefact = (det_checks or {}).get("artefact")
    if det_artefact and Path(det_artefact).exists():
        try:
            if Path(det_artefact).resolve() != (output_dir / "det_checks.json").resolve():
                shutil.copy(str(det_artefact), str(output_dir / "det_checks.json"))
        except OSError as e:
            logger.warning(f"  [det_checks] could not copy artefact: {e}")

    # Calculate scores
    score_results = None
    if weights_data:
        logger.info("\n[Score] Calculating scores...")
        score_results_llm = calculate_scores(
            all_responses, weights_data, max_mistakes=RUBRIC_MAX_MISTAKES
        )
        score_results = score_results_llm
        engine_block = {
            "mode": accuracy_engine,
            "effective": "llm",
            "checks": {},
            "total_score_llm": score_results_llm["total_score"],
            "total_score_harness": None,
        }
        if det_checks is not None:
            engine_block["det_checks_mode"] = det_mode
        # Which switch governs which producer (family): the answer check
        # follows --accuracy-check, the deterministic checks --det-checks.
        family_modes = {"answer_check": accuracy_engine, "det_checks": det_mode}
        if harness_verdicts:
            harness_responses, applied = _apply_harness_verdicts(
                all_responses, harness_verdicts, weights_data
            )
            engine_block["checks"] = applied
            overlaid = [k for k, v in applied.items() if v["engine"] == "harness"]
            counted = [
                k for k in overlaid
                if family_modes.get(applied[k].get("family") or "answer_check") == "harness"
            ]
            for k, v in applied.items():
                v["counted"] = k in counted
            if overlaid:
                score_results_harness = calculate_scores(
                    harness_responses, weights_data, max_mistakes=RUBRIC_MAX_MISTAKES
                )
                engine_block["total_score_harness"] = score_results_harness["total_score"]
                with open(output_dir / "ai_judgement_harness.json", "w", encoding="utf-8") as f:
                    json.dump(harness_responses, f, indent=2)
                if counted and len(counted) == len(overlaid):
                    score_results = score_results_harness
                    engine_block["effective"] = "harness"
                elif counted:
                    # Only some families count: score the LLM judgement with
                    # just their verdicts overlaid.
                    effective_responses, _ = _apply_harness_verdicts(
                        all_responses,
                        {k: harness_verdicts[k] for k in counted},
                        weights_data,
                    )
                    score_results = calculate_scores(
                        effective_responses, weights_data, max_mistakes=RUBRIC_MAX_MISTAKES
                    )
                    engine_block["effective"] = "mixed"
                logger.info(
                    f"  Accuracy engine: mode={accuracy_engine}"
                    + (f", det_checks={det_mode}" if det_checks is not None else "")
                    + f" -> effective={engine_block['effective']} "
                    f"({len(counted)} of {len(overlaid)} harness verdict(s) counted); "
                    f"total llm={score_results_llm['total_score']:.2f} harness="
                    f"{score_results_harness['total_score']:.2f} recorded="
                    f"{score_results['total_score']:.2f}"
                )
            elif accuracy_engine == "harness":
                logger.info(
                    "  Accuracy engine: harness requested but nothing measurable "
                    "— LLM verdicts stand (reasons in scored_results.accuracy_engine)"
                )
        score_results["accuracy_engine"] = engine_block
        tag_check_graders(score_results)
        if det_checks is not None:
            score_results["det_checks"] = det_checks.get("summary")
        if suitability_provenance is not None:
            # Phase A provenance rides inside scored_results (and scores.json)
            # so every grading records exactly which checks were gated out.
            score_results["rubric_suitability"] = suitability_provenance
        if formula_cache_provenance is not None:
            # Same idea for the uncached-formula census: a row graded under
            # JUDGE_SKIP_FORMULA_CACHE_CHECK=1 must stay identifiable, and the
            # per-workbook counts explain any Accuracy anomaly after the fact.
            score_results["formula_cache"] = formula_cache_provenance
        scores_path = output_dir / "scores.json"
        with open(scores_path, "w", encoding="utf-8") as f:
            json.dump(score_results, f, indent=2)
        logger.info(f"  Scores saved to: {scores_path}")

        logger.info("\n  Score Summary:")
        for cat in score_results["criteria_scores"].keys():
            if cat in score_results["criteria_scores"]:
                cs = score_results["criteria_scores"][cat]
                logger.info(
                    f"    {cat}: {cs['normalized_score']:.2f}/100 "
                    f"(weight: {cs['category_weight']:.2f}, "
                    f"contribution: {cs['normalized_score'] * cs['category_weight']:.2f})"
                )
        logger.info(f"    TOTAL: {score_results['total_score']:.2f}/100")

    # Save token tracking
    token_tracking["model"] = model
    token_tracking_path = output_dir / "token_tracking.json"
    with open(token_tracking_path, "w", encoding="utf-8") as f:
        json.dump(token_tracking, f, indent=2)
    logger.info(f"  Token tracking saved to: {token_tracking_path}")

    # Create metadata
    elapsed_time = time.time() - start_time
    metadata_path = output_dir / "_metadata.json"
    metadata_dict = {
        "task_folder": task_folder_name,
        "grader_model": model,
        "judge_reasoning": reasoning_effort,
        "grader_identity": grader_identity,
        "judge_mode": "agentic" if agentic else "non-agentic",
        "auto_routed": auto_routed,
        "attempt_model": attempt_model,
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "judge_version": versions["JUDGE_VERSION"],
        "prompt_version": versions["PROMPT_VERSION"],
        "rubric_version": versions["RUBRIC_VERSION"],
        "rubric_weight_version": versions["RUBRIC_WEIGHT_VERSION"],
        "rubric_max_mistakes": RUBRIC_MAX_MISTAKES,
        "total_prompt_tokens": token_tracking.get("total_prompt_tokens", 0),
        "total_completion_tokens": token_tracking.get("total_completion_tokens", 0),
        "total_tokens": token_tracking.get("total_tokens", 0),
        "total_cost": round(token_tracking.get("total_cost", 0), 6),
        "elapsed_time_seconds": round(elapsed_time, 2),
        "files_considered": {
            "golden_solution": (
                sorted(golden_solution_files.keys()) if golden_solution_files else []
            ),
            "ai_attempt": (sorted(ai_attempt_files.keys()) if ai_attempt_files else []),
            "context": context_file_path.name if context_file_path else None,
        },
    }
    if score_results:
        metadata_dict["total_score"] = score_results["total_score"]
        metadata_dict["criteria_scores"] = {
            cat: data["normalized_score"]
            for cat, data in score_results["criteria_scores"].items()
        }
    if suitability_provenance is not None:
        metadata_dict["rubric_suitability"] = suitability_provenance
    if formula_cache_provenance is not None:
        metadata_dict["formula_cache"] = formula_cache_provenance
    if det_checks is not None:
        _det_summary = det_checks.get("summary") or {}
        metadata_dict["det_checks"] = {
            k: _det_summary.get(k)
            for k in ("status", "mode", "code_sha", "graded", "not_applicable", "values")
            if k in _det_summary
        }

    if metadata_path.exists():
        try:
            with open(metadata_path, "r", encoding="utf-8") as f:
                existing_metadata = json.load(f)
            if isinstance(existing_metadata, dict):
                existing_metadata.update(metadata_dict)
                metadata_dict = existing_metadata
        except (json.JSONDecodeError, IOError):
            pass

    with open(metadata_path, "w", encoding="utf-8") as f:
        json.dump(metadata_dict, f, indent=2)
    logger.info(f"  Metadata saved to: {metadata_path}")

    # Log summary
    logger.info("\n" + "=" * 80)
    logger.info("EVALUATION COMPLETE")
    logger.info("=" * 80)
    total_msg_size = token_tracking.get("total_message_size", 0)
    total_msg_size_img = token_tracking.get("total_message_size_with_images", 0)
    total_tokens = token_tracking.get("total_tokens", 0)
    logger.info(f"\nToken Usage & Cost Summary:")
    logger.info(
        f"  Total message size: {total_msg_size:,} characters "
        f"(with images: {total_msg_size_img:,})"
    )
    logger.info(f"  Total tokens used: {total_tokens:,}")
    logger.info(
        f"    - Prompt tokens: {token_tracking.get('total_prompt_tokens', 0):,}"
    )
    logger.info(
        f"    - Completion tokens: "
        f"{token_tracking.get('total_completion_tokens', 0):,}"
    )
    if total_tokens > 0:
        logger.info(f"  Average ratio: {total_msg_size / total_tokens:.2f} chars/token")
    logger.info(f"  Total cost: ${token_tracking.get('total_cost', 0):.6f}")
    if token_tracking.get("evaluations"):
        logger.info("\n  Evaluations:")
        for cat, data in token_tracking["evaluations"].items():
            logger.info(
                f"    {cat}: {data['message_size']:,} chars "
                f"(with images: {data.get('message_size_with_images', 0):,}) -> "
                f"{data['total_tokens']:,} tokens "
                f"({data.get('chars_per_token', 0):.2f} chars/token) | "
                f"${data.get('cost', 0):.6f}"
            )
    logger.info("=" * 80)

    # Build result
    result = {
        "ai_judgement": str(ai_judgement_path),
        "output_dir": str(output_dir),
        "solution_csv_dir": golden_solution_dir,
        "attempt_csv_dir": ai_attempt_dir,
        "starting_csv_dir": starting_workbook_dir,
        "solution_context_reduced": solution_context_reduced,
        "attempt_context_reduced": attempt_context_reduced,
        "context_reduced_details": context_reduced_details,
        "auto_routed": auto_routed,
        "judge_reasoning": reasoning_effort,
        "grader_identity": grader_identity,
        # The versions this grading actually ran under; grade.py records
        # them in the grading row rather than re-reading the env.
        "versions": versions,
    }
    if score_results:
        result["score_results"] = score_results
        criteria_scores = score_results["criteria_scores"]
        # Legacy per-category keys feed the three per-category grade columns
        # of the grading row. Category names differ across rubric editions
        # ("Formula" vs rubric_9 "Formulas"), so resolve through aliases; a
        # rubric that lacks the category yields None for that column rather
        # than failing the run.
        result["accuracy_score"] = _resolve_category_score(
            criteria_scores, "Accuracy"
        )
        result["formula_score"] = _resolve_category_score(
            criteria_scores, "Formula", "Formulas"
        )
        result["formatting_score"] = _resolve_category_score(
            criteria_scores, "Formatting"
        )
        result["final_score"] = score_results["total_score"]
        # Score completeness is judged against the configured weights file
        # (the scoring contract), not a fixed category list: any weighted
        # category without a normalized score (e.g. skipped on parse
        # failure) means the totals can't be trusted.
        result["missing_categories"] = [
            cat
            for cat in weights_data["CategoryWeights"][0]
            if criteria_scores.get(cat, {}).get("normalized_score") is None
        ]

        scoring_warnings = score_results.get("scoring_warnings") or {}
        if any(scoring_warnings.get(k) for k in scoring_warnings):
            result["scoring_warnings"] = scoring_warnings
            logger.info("\nScoring Warnings Summary:")
            if scoring_warnings.get("unscored_checks"):
                logger.info("  Unscored checks (weighted but not evaluated):")
                for cat, names in scoring_warnings["unscored_checks"].items():
                    logger.info(f"    {cat}: {names}")
            if scoring_warnings.get("empty_category_judgements"):
                logger.info(
                    f"  Empty category judgements: "
                    f"{scoring_warnings['empty_category_judgements']}"
                )
            if scoring_warnings.get("duplicate_judgements"):
                logger.info("  Duplicate judgement names (kept max mistakes):")
                for cat, dupes in scoring_warnings["duplicate_judgements"].items():
                    for name, ms in dupes.items():
                        logger.info(f"    {cat}/{name}: reported mistakes={ms}")
            if scoring_warnings.get("mistake_count_mismatches"):
                logger.info(
                    "  total_mistakes vs len(mistakes) mismatches "
                    "(using len(mistakes)):"
                )
                for m in scoring_warnings["mistake_count_mismatches"]:
                    logger.info(
                        f"    {m['category']}/{m['name']}: "
                        f"claimed={m['claimed_total_mistakes']} "
                        f"actual={m['actual_mistakes_len']}"
                    )
            if scoring_warnings.get("fail_without_mistakes"):
                logger.info(
                    "  Decision='fail' with no mistakes recorded "
                    "(scored as max_mistakes):"
                )
                for f in scoring_warnings["fail_without_mistakes"]:
                    logger.info(f"    {f['category']}/{f['name']}")

    if parse_failures:
        result["parse_failures"] = parse_failures
        logger.info("\nJSON Parse Failures Summary:")
        for category, info in parse_failures.items():
            logger.info(
                f"  {category}: {info['count']} failed parse attempts recorded."
            )
            if info["count"] >= 1:
                logger.info(
                    f"    Sample failed response: {info['responses'][0][:500]}..."
                )

    remove_log_file(cache_log_path)
    shutil.copy(cache_log_path, str(output_dir / "judge.log"))
    return result


### Agentic Judge

AGENTIC_JUDGE_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": (
                "Read a rectangular range from a CSV file in the AI attempt, "
                "golden solution, or starting workbook directories. You must "
                "specify the row and column "
                "range to extract. Use the sheet dimensions and any formatting "
                "metadata provided to decide which ranges to "
                "inspect. HARD LIMIT: a single call may not cover more than "
                "5000 cells ((end_row - start_row + 1) * (end_col - start_col + "
                "1)). Requests exceeding this are rejected; split large regions "
                "into multiple smaller calls."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source": {
                        "type": "string",
                        "enum": ["attempt", "solution", "starting"],
                        "description": (
                            "Which directory to read from: 'attempt' for the AI "
                            "attempt workbook, 'solution' for the golden "
                            "solution, 'starting' for the workbook the agent "
                            "was given before doing any work (when available)."
                        ),
                    },
                    "filename": {
                        "type": "string",
                        "description": (
                            "The CSV filename to read, e.g. 'Sheet1_full.csv'."
                        ),
                    },
                    "start_row": {
                        "type": "integer",
                        "description": (
                            "First row to include (1-based, inclusive). Row 1 is "
                            "the first row of the spreadsheet."
                        ),
                    },
                    "end_row": {
                        "type": "integer",
                        "description": ("Last row to include (1-based, inclusive)."),
                    },
                    "start_col": {
                        "type": "string",
                        "description": (
                            "First column letter to include (inclusive), e.g. 'A'."
                        ),
                    },
                    "end_col": {
                        "type": "string",
                        "description": (
                            "Last column letter to include (inclusive), e.g. 'Z'."
                        ),
                    },
                },
                "required": [
                    "source",
                    "filename",
                    "start_row",
                    "end_row",
                    "start_col",
                    "end_col",
                ],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "record_check",
            "description": (
                "Record (upsert) a pass/fail decision and summary for a rubric "
                "check in the current category. Must be called for every check "
                "letter before you stop. Calling again for the same letter "
                "overwrites the decision/summary but preserves any mistakes "
                "already appended via append_mistake."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "check": {
                        "type": "string",
                        "description": ("The rubric check letter (e.g. 'A')."),
                    },
                    "decision": {
                        "type": "string",
                        "enum": ["pass", "fail"],
                        "description": "Your pass/fail decision for this check.",
                    },
                    "summary": {
                        "type": "string",
                        "description": (
                            "Brief explanation of your assessment for this check."
                        ),
                    },
                },
                "required": ["check", "decision", "summary"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "append_mistake",
            "description": (
                "Append a single mistake to a check that has already been "
                "recorded via record_check. Call once per mistake. The "
                "corresponding record_check must have been called first."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "check": {
                        "type": "string",
                        "description": "The rubric check letter (e.g. 'A').",
                    },
                    "location": {
                        "type": "string",
                        "description": (
                            "Cell / range reference where the mistake occurs "
                            "(e.g. 'Sheet1!B7')."
                        ),
                    },
                    "description": {
                        "type": "string",
                        "description": "What is wrong.",
                    },
                    "severity": {
                        "type": "string",
                        "enum": ["minor", "major"],
                        "description": "Severity of the mistake.",
                    },
                },
                "required": ["check", "location", "description", "severity"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_working_judgement",
            "description": (
                "Return the current working judgement state for this category "
                "as JSON. Useful to review what you've recorded so far before "
                "finalizing."
            ),
            "parameters": {"type": "object", "properties": {}},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "evict_tool_results",
            "description": (
                "Drop all prior rounds' read_file / scratchpad tool-call and "
                "tool-result messages from the wire context (rounds strictly "
                "before before_round). The findings from those reads should "
                "already have been absorbed into your working judgement via "
                "record_check / append_mistake. Use this when the context "
                "pressure signal tells you context is filling up."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "before_round": {
                        "type": "integer",
                        "description": (
                            "Evict every tool-call/tool-result pair whose round "
                            "index is strictly less than this number. Round "
                            "numbering starts at 1 for the first API call."
                        ),
                    },
                },
                "required": ["before_round"],
            },
        },
    },
]


def _build_single_pass_tools() -> list[dict]:
    """The single-pass toolset, derived from AGENTIC_JUDGE_TOOLS.

    Same five tools; three deltas:
      - read_file gains an optional `view` parameter. Twelve-category mode
        keys the served view off the category being graded; a single pass
        has no current category, so the model states what it is looking for
        and the same serving rule applies (data view by default, styled view
        for formatting work, structure metadata for structure work).
      - record_check / append_mistake address checks by their global check
        NUMBER (the rubric list is globally numbered, not lettered).
    Derived programmatically so the range parameters and limits can never
    drift from the shared definition.
    """
    import copy

    tools = copy.deepcopy(AGENTIC_JUDGE_TOOLS)
    by_name = {t["function"]["name"]: t["function"] for t in tools}

    by_name["read_file"]["parameters"]["properties"]["view"] = {
        "type": "string",
        "enum": ["data", "formatting", "structure"],
        "description": (
            "Optional; default 'data'. 'data': values, formulas and number "
            "formats with cell-style segments stripped — use for everything "
            "except style/layout work. 'formatting': full cell-style "
            "segments (font, size, colors, fill, alignment, borders), plus "
            "the sheet's formatting metadata (merged ranges, frozen panes) "
            "appended to the first read of each sheet — use when grading "
            "Formatting checks. 'structure': the data view plus the sheet "
            "formatting metadata — use when grading Structure checks."
        ),
    }

    by_name["record_check"]["description"] = (
        "Record (upsert) a pass/fail decision and summary for a rubric "
        "check. Must be called for every check number before you stop. "
        "Calling again for the same check number overwrites the "
        "decision/summary but preserves any mistakes already appended via "
        "append_mistake."
    )
    by_name["record_check"]["parameters"]["properties"]["check"][
        "description"
    ] = "The check's number as shown in the rubric list (e.g. '17')."
    by_name["append_mistake"]["parameters"]["properties"]["check"][
        "description"
    ] = "The check's number as shown in the rubric list (e.g. '17')."
    by_name["get_working_judgement"]["description"] = (
        "Return the current working judgement state as JSON. Useful to "
        "review what you've recorded so far and which check numbers are "
        "still pending before finalizing."
    )
    return tools


SINGLE_PASS_JUDGE_TOOLS = _build_single_pass_tools()


def _build_file_metadata(directory: str) -> dict:
    """Build metadata dict for CSV files in a directory.

    Returns:
        dict mapping filenames to {"format_info": str|None}.
        Only includes *_full.csv files. format_info is the content of the
        corresponding *_additional_format.txt file (which already contains
        sheet dimensions, merged cells, and frozen panes info).
    """
    metadata = {}
    if not directory or not Path(directory).exists():
        return metadata

    dir_path = Path(directory)
    for csv_file in sorted(dir_path.glob("*_full.csv")):
        base_name = csv_file.stem.replace("_full", "")
        format_txt = dir_path / f"{base_name}_additional_format.txt"
        format_info = None
        if format_txt.exists():
            try:
                format_info = format_txt.read_text(encoding="utf-8")
            except Exception:
                pass
        metadata[csv_file.name] = {
            "format_info": format_info,
        }
    return metadata


def _col_letter_to_index(col_str: str) -> int:
    """Convert an Excel column letter (e.g. 'A', 'Z', 'AA') to a 1-based index."""
    col_str = col_str.upper().strip()
    result = 0
    for ch in col_str:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


_DIMS_RE = None


def _render_files_text_dims_only(file_list, file_metadata=None) -> str:
    """File listing with one dimensions line per sheet (template 5+).

    The full formatting metadata (merged cells / frozen panes) is no longer
    inlined here — it accompanies read_file results in the Formatting and
    Structure categories instead — so the listing is byte-identical across
    categories and sits in the prompt-cache-stable prefix.
    """
    global _DIMS_RE
    if _DIMS_RE is None:
        import re as _re

        _DIMS_RE = _re.compile(r"Sheet Dimensions: (\d+) rows x (\d+) columns")
    lines = []
    for fname in file_list:
        lines.append(f"    {fname}")
        meta = (file_metadata or {}).get(fname)
        info = (meta or {}).get("format_info")
        if info:
            m = _DIMS_RE.search(info)
            if m:
                lines.append(
                    f"      Dimensions: {m.group(1)} rows x {m.group(2)} columns"
                )
    return "\n".join(lines)


def _prompt_version_at_least(version_str, minimum: int) -> bool:
    """True when the agentic prompt version's major number >= minimum."""
    try:
        return int(str(version_str).split(".")[0]) >= minimum
    except (TypeError, ValueError):
        return False


def _build_agentic_context_messages(context_file_path) -> list[dict]:
    """Build context messages for the agentic judge, supporting .txt and PDFs.

    PDFs (and other binary types) are base64-encoded and attached via
    image_url, mirroring the standard judge's context handling. Returns an
    empty list if no context file is provided.
    """
    if not context_file_path:
        return []

    context_file_path = Path(context_file_path)
    ext = context_file_path.suffix.lower()

    if ext == ".txt":
        try:
            with open(context_file_path, "r", encoding="utf-8") as f:
                content = f.read()
            return [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"Case context:\n{content}"},
                    ],
                }
            ]
        except UnicodeDecodeError:
            logger.warning(f"  Could not read context file: {context_file_path}")
            # Fall through to base64 encoding below
            pass

    base64_content, mime_type = encode_file_to_base64(context_file_path)
    return [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": f"Case context ({context_file_path.name}):"},
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{mime_type};base64,{base64_content}"},
                },
            ],
        }
    ]


def _measure_message_chars(msg) -> int:
    """Return the character count of a single conversation message.

    Handles plain dicts (system/user/tool messages) and SDK response objects
    (assistant messages that may carry tool_calls with arguments).
    """
    total = 0

    # Dict messages (system, user, tool, or manually constructed assistant)
    if isinstance(msg, dict):
        content = msg.get("content")
        if isinstance(content, str):
            total += len(content)
        elif isinstance(content, list):
            # Multi-part content (e.g. text + image blocks)
            for part in content:
                if isinstance(part, dict):
                    total += len(part.get("text", ""))
        # Manually-constructed tool_calls (shouldn't normally happen, but be safe)
        for tc in msg.get("tool_calls", []):
            if isinstance(tc, dict):
                func = tc.get("function", {})
                total += len(func.get("name", ""))
                total += len(func.get("arguments", ""))
        return total

    # SDK ChatCompletionMessage objects (from choice.message)
    content = getattr(msg, "content", None)
    if isinstance(content, str):
        total += len(content)

    tool_calls = getattr(msg, "tool_calls", None)
    if tool_calls:
        for tc in tool_calls:
            func = getattr(tc, "function", None)
            if func:
                total += len(getattr(func, "name", "") or "")
                total += len(getattr(func, "arguments", "") or "")

    return total


def _msg_to_dict(m):
    """Convert an SDK message or plain dict to a plain dict for logging."""
    if isinstance(m, dict):
        return m

    # Handles ChatCompletionMessage (pydantic v2) — preserves role, content,
    # tool_calls, and any reasoning fields rather than stringifying.
    if hasattr(m, "model_dump"):
        return m.model_dump(exclude_none=True)

    logger.warning(
        f"  Warning: encountered non-dict message of type {type(m)}; \n content: {m}"
    )

    return {"role": "unknown", "content": str(m)}


class WorkingJudgement:
    """Per-category scratchpad state mutated by record_check / append_mistake."""

    def __init__(self, category: str, check_letters: list[str]):
        self.category = category
        self.check_letters = list(check_letters)
        self.working: dict[str, dict] = {}
        self.pending: set[str] = set(self.check_letters)

    def record_check(self, letter: str, decision: str, summary: str) -> str:
        if letter not in self.check_letters:
            return (
                f"Error: unknown check '{letter}' for {self.category}. "
                f"Valid letters: {', '.join(self.check_letters)}."
            )
        if decision not in ("pass", "fail"):
            return f"Error: decision must be 'pass' or 'fail' (got '{decision}')."
        if letter in self.working:
            self.working[letter]["decision"] = decision
            self.working[letter]["summary"] = summary
        else:
            self.working[letter] = {
                "check": letter,
                "decision": decision,
                "summary": summary,
                "mistakes": [],
            }
        self.pending.discard(letter)
        return (
            f"Recorded check {letter} as {decision}. Coverage: {self.coverage_str()}."
        )

    def append_mistake(
        self, letter: str, location: str, description: str, severity: str
    ) -> str:
        if letter not in self.check_letters:
            return (
                f"Error: unknown check '{letter}' for {self.category}. "
                f"Valid letters: {', '.join(self.check_letters)}."
            )
        if letter not in self.working:
            return (
                f"Error: cannot append_mistake to '{letter}' before "
                f"record_check. Call record_check('{letter}', ...) first."
            )
        if severity not in ("minor", "major"):
            return f"Error: severity must be 'minor' or 'major' (got '{severity}')."
        self.working[letter]["mistakes"].append(
            {
                "location": location,
                "description": description,
                "severity": severity,
            }
        )
        count = len(self.working[letter]["mistakes"])
        return (
            f"Appended mistake to {letter} (now {count} mistake"
            f"{'s' if count != 1 else ''}). "
            f"Coverage: {self.coverage_str()}."
        )

    def coverage_str(self) -> str:
        marks = " ".join(f"{l}\u2713" for l in self.check_letters if l in self.working)
        pending_letters = [l for l in self.check_letters if l in self.pending]
        if pending_letters:
            return f"{marks or '(none yet)'} | pending: {', '.join(pending_letters)}"
        return marks or "(none)"

    def fails_missing_mistakes(self) -> list[str]:
        """Letters recorded 'fail' with zero appended mistakes.

        calculate_scores treats such checks as a scoring hazard
        (fail_without_mistakes) and the whole grading is marked failed \u2014
        the loop nudges the model to repair them before accepting completion.
        """
        return [
            l
            for l in self.check_letters
            if l in self.working
            and self.working[l].get("decision") == "fail"
            and not self.working[l].get("mistakes")
        ]

    def finalize(self) -> list[dict]:
        return [self.working[l] for l in self.check_letters if l in self.working]


def _execute_scratchpad_tool(tool_call, working: WorkingJudgement) -> str:
    """Dispatch record_check / append_mistake / get_working_judgement calls."""
    name = tool_call.function.name
    try:
        args = json.loads(tool_call.function.arguments or "{}")
    except json.JSONDecodeError as e:
        return f"Error: could not parse tool arguments JSON: {e}"

    if name == "record_check":
        letter = str(args.get("check", "")).strip().upper()
        decision = str(args.get("decision", "")).strip().lower()
        summary = str(args.get("summary", ""))
        return working.record_check(letter, decision, summary)

    if name == "append_mistake":
        letter = str(args.get("check", "")).strip().upper()
        location = str(args.get("location", ""))
        description = str(args.get("description", ""))
        severity = str(args.get("severity", "")).strip().lower()
        return working.append_mistake(letter, location, description, severity)

    if name == "get_working_judgement":
        return json.dumps(
            {
                "category": working.category,
                "working": {
                    l: working.working[l]
                    for l in working.check_letters
                    if l in working.working
                },
                "pending": [l for l in working.check_letters if l in working.pending],
            },
            indent=2,
        )

    return f"Error: unknown scratchpad tool '{name}'."


class AgenticCategoryLoop:
    """Holds wire messages, append-only transcript, and per-round bookkeeping.

    Wire (`messages`) is what gets sent to the API and can be mutated by
    eviction. `transcript` is append-only and never reordered; evicted
    entries are flagged with `_evicted_at_round` rather than deleted.
    Every mutation to `messages` goes through `append_wire` so the two stay
    in sync.
    """

    def __init__(self, category: str, check_letters: list[str], seed_messages: list):
        self.category = category
        self.working = WorkingJudgement(category, check_letters)
        self.messages: list = list(seed_messages)
        # parallel meta to `messages`: {"round": int, "tag": str|None}
        self.message_meta: list[dict] = [
            {"round": 0, "tag": "setup"} for _ in seed_messages
        ]
        self.transcript: list[dict] = [
            {
                "type": "msg",
                "msg": _msg_to_dict(m),
                "tag": "setup",
                "round": 0,
                "_evicted_at_round": None,
            }
            for m in seed_messages
        ]
        self.round: int = 0
        self.last_prompt_tokens: int = 0

    def append_wire(self, msg, tag: str | None = None) -> None:
        self.messages.append(msg)
        self.message_meta.append({"round": self.round, "tag": tag})
        self.transcript.append(
            {
                "type": "msg",
                "msg": _msg_to_dict(msg),
                "tag": tag,
                "round": self.round,
                "_evicted_at_round": None,
            }
        )

    def log_event(self, event_type: str, **data) -> None:
        self.transcript.append(
            {
                "type": "event",
                "event": event_type,
                "round": self.round,
                **data,
            }
        )

    def append_synthetic_final(self, judgement: list[dict]) -> None:
        """Record the reconstructed final judgement in the transcript."""
        payload = json.dumps({self.category: judgement}, indent=2)
        self.transcript.append(
            {
                "type": "msg",
                "msg": {"role": "assistant", "content": payload},
                "tag": "synthetic_final",
                "round": self.round,
                "_evicted_at_round": None,
            }
        )

    def evict(self, before_round: int) -> str:
        """Drop tool-call / tool-result pairs with round < before_round."""
        to_evict_indices = [
            i
            for i, meta in enumerate(self.message_meta)
            if 0 < meta["round"] < before_round
            and meta.get("tag") in ("model_tool_call", "tool_result")
        ]
        if not to_evict_indices:
            return (
                "No eligible rounds to evict. evict_tool_results only drops "
                "tool-call/tool-result pairs from rounds before the given "
                "round that still remain in wire context."
            )

        dropped_chars = sum(
            _measure_message_chars(self.messages[i]) for i in to_evict_indices
        )
        rounds_evicted = sorted(
            {self.message_meta[i]["round"] for i in to_evict_indices}
        )
        first_r, last_r = rounds_evicted[0], rounds_evicted[-1]
        stub = {
            "role": "user",
            "content": (
                f"[evicted rounds {first_r}-{last_r} - read_file results "
                f"absorbed into working judgement]"
            ),
        }
        evict_set = set(to_evict_indices)
        new_msgs, new_meta = [], []
        inserted = False
        for i, (m, meta) in enumerate(zip(self.messages, self.message_meta)):
            if i in evict_set:
                if not inserted:
                    new_msgs.append(stub)
                    new_meta.append({"round": self.round, "tag": "evict_stub"})
                    inserted = True
                continue
            new_msgs.append(m)
            new_meta.append(meta)
        self.messages = new_msgs
        self.message_meta = new_meta

        for entry in self.transcript:
            if entry.get("type") != "msg":
                continue
            if entry.get("_evicted_at_round") is not None:
                continue
            if entry.get("tag") not in ("model_tool_call", "tool_result"):
                continue
            if first_r <= entry.get("round", -1) <= last_r:
                entry["_evicted_at_round"] = self.round

        self.log_event(
            "evict",
            dropped_rounds=[first_r, last_r],
            dropped_messages=len(to_evict_indices),
            dropped_chars=dropped_chars,
        )
        return (
            f"Evicted {len(rounds_evicted)} round(s) ({first_r}-{last_r}). "
            f"Wire context reduced by ~{dropped_chars // 1000}K chars."
        )


def _wire_char_total(messages) -> int:
    """Sum of character counts across all wire messages."""
    return sum(_measure_message_chars(m) for m in messages)


def _estimate_wire_tokens(wire_chars: int, chars_per_token: float) -> int:
    """Convert wire chars to an estimated token count using a calibrated ratio.

    `chars_per_token` is computed from the most recent API call's
    ``usage.prompt_tokens`` vs. the wire chars at call time. Falls back to
    a conservative 2.5 when no calibration is available (CSV-dense content
    is ~1-2 chars/token; English prose is ~4).
    """
    cpt = chars_per_token if chars_per_token and chars_per_token > 0 else 2.5
    return int(wire_chars / cpt)


def _build_pressure_signal(
    prompt_tokens: int, limit: int, rounds_elapsed: int
) -> tuple[str, str]:
    """Return (status_line, tier) for the current context pressure level.

    Tier is one of 'low' (<65%), 'advisory' (65-80%), 'strong' (80-90%),
    'forced' (>=90%). Thresholds moved 10/20/80 -> 65/80/90 (2026-09):
    the old ladder told every observed run to wrap up from 20% of the
    limit — sol's mean peak was 43%, so it spent effectively its whole
    grading being told to finalize, which is a live candidate for why it
    deliberated least and swung most between repeats.
    """
    pct = (prompt_tokens / limit * 100) if limit > 0 else 0.0

    def _k(n: int) -> str:
        return f"~{n // 1000}K" if n >= 1000 else f"~{n}"

    status = (
        f"[context: {_k(prompt_tokens)} / {_k(limit)} tokens "
        f"({pct:.0f}%). {rounds_elapsed} rounds elapsed.]"
    )
    if pct < 65:
        tier = "low"
    elif pct < 80:
        tier = "advisory"
    elif pct < 90:
        tier = "strong"
    else:
        tier = "forced"

    if tier == "advisory":
        status += (
            "\nConsider evict_tool_results to drop rounds whose findings "
            "you've already recorded."
        )
    elif tier == "strong":
        status += "\nYou must evict stale reads or finalize on your next turn."
    elif tier == "forced":
        status += (
            "\nYou must finalize now. Call record_check for every pending "
            "check or evict stale reads before calling more read_file."
        )
    return status, tier


# Token density assumed for a tool result the gate is ABOUT to add. The
# global calibrated ratio reflects the context mix so far — mostly prose on
# early rounds (~3.5-4 chars/token) — but big read results are dense CSV
# with formulas. MEASURED, not guessed: the canary's twice-crashed
# conversation (attempt 353, 2.07M wire chars vs 1,022,753 real tokens)
# tokenized at 2.03 chars/token; 2.4 was still ~18% optimistic and the
# burst squeaked through again. 1.8 gives ~11% margin below the measured
# density. Over-refusing a rare prose-dense result errs safe and costs one
# recoverable retry.
_READ_GATE_RESULT_CPT = 1.8
# Tokens held back from the budget for the round's own overhead (the next
# pressure note, nudges, the model's reply) and residual estimator error.
_READ_GATE_RESERVE_TOKENS = 30_000


def _read_refusal_check(
    result_chars: int, current_tokens: int, limit: int
) -> tuple[bool, int, int]:
    """Would serving a read of `result_chars` blow the context budget?

    Returns (refuse, projected_tokens, current_tokens). Measured, not
    estimated: the caller executes the read locally first (CSV slicing is
    free) and gates on the ACTUAL result size, so the clamped-range trap —
    refusing a request whose rectangle is huge but whose real content is
    small — cannot occur.

    `current_tokens` is the caller's running estimate for THIS round: the
    start-of-round wire converted at the live usage calibration, plus every
    in-round addition converted at CSV density as it lands — so a burst of
    parallel reads gates each read against the true running total, in the
    density of what was actually added. The budget is `limit` minus a
    fixed reserve.

    Why this exists (canary step 1, 2026-09-01): single-pass sol gradings
    issued bursts of parallel 5000-cell reads that took wire context from
    ~85K to 700K+ tokens in ONE round, and the next request 400'd at the
    provider's hard input cap (922K) — the pressure ladder warns between
    rounds and can never stop an intra-round burst. This gate is the hard
    guarantee; the refusal it produces is recoverable (evict, then re-read).
    """
    projected_tokens = current_tokens + int(result_chars / _READ_GATE_RESULT_CPT)
    budget = limit - _READ_GATE_RESERVE_TOKENS
    return projected_tokens >= budget, projected_tokens, current_tokens


def _read_refusal_message(
    result_chars: int, projected_tokens: int, current_tokens: int,
    limit: int, current_round: int,
) -> str:
    return (
        f"REFUSED (context budget): this read returned ~{result_chars:,} "
        f"characters, which would take the conversation to "
        f"~{projected_tokens // 1000}K of the {limit // 1000}K-token limit "
        f"(currently ~{current_tokens // 1000}K). The result was NOT added "
        f"to context. To proceed: (1) record any findings you already have "
        f"via record_check / append_mistake, (2) call "
        f"evict_tool_results(before_round={current_round}) to drop old read "
        f"results you no longer need, then (3) re-issue this read — it is "
        f"always retryable after a successful eviction. Or request a "
        f"smaller range instead."
    )


def _execute_read_file(tool_call, attempt_dir, solution_dir, category=None,
                       format_notes=None, starting_dir=None):
    """Execute a read_file tool call, extracting the specified row/column range from a CSV.

    When `category` is given (prompt template 5+), requests for *_full.csv
    resolve through the category-keyed serving rule (2026-08 cost lever):
      - Formatting          -> the full view (style FORMAT segments included)
      - every other category -> the *_data.csv sibling (style segments
        stripped; number-format rendering kept), falling back to the full
        view if the sibling is missing (pre-revision cache).
    Formatting and Structure reads also append the sheet's formatting
    metadata (merged cells / frozen panes) once per sheet per category,
    tracked in `format_notes`. Presented filenames stay *_full.csv — the
    resolution is an internal indirection.

    `starting_dir` (2026-09) serves the workbook the agent was given before
    doing any work, so inherited content can be distinguished from the
    agent's own — several guidance rules depend on that distinction. Absent
    (an older staging), source='starting' returns a clear error.
    """
    import csv as csv_mod

    args = json.loads(tool_call.function.arguments)
    source = args.get("source", "")
    filename = args.get("filename", "")
    start_row = args.get("start_row")
    end_row = args.get("end_row")
    start_col = args.get("start_col")
    end_col = args.get("end_col")

    if source == "attempt":
        base_dir = attempt_dir
    elif source == "solution":
        base_dir = solution_dir
    elif source == "starting":
        base_dir = starting_dir
    else:
        return (
            f"Error: Invalid source '{source}'. Use 'attempt', 'solution' "
            f"or 'starting'."
        )

    if not base_dir or not Path(base_dir).exists():
        return f"Error: {source} directory not available."

    file_path = Path(base_dir) / filename
    if filename.endswith("_additional_format.txt"):
        return (
            f"Error: '{filename}' is not directly readable — its contents are "
            f"already inlined under the corresponding *_full.csv entry in the "
            f"file list."
        )
    if filename.endswith("_data.csv"):
        return (
            f"Error: '{filename}' is not directly readable — request the "
            f"corresponding *_full.csv instead."
        )
    if not file_path.exists():
        available = sorted(
            f.name
            for f in Path(base_dir).iterdir()
            if f.is_file()
            and not f.name.endswith("_additional_format.txt")
            and not f.name.endswith("_data.csv")
        )
        return (
            f"Error: File '{filename}' not found in {source} directory. "
            f"Available files: {', '.join(available)}"
        )

    # Category-keyed serving indirection (template 5+ only).
    if category is not None and category != "Formatting" and filename.endswith(
        "_full.csv"
    ):
        data_sibling = Path(base_dir) / (
            filename[: -len("_full.csv")] + "_data.csv"
        )
        if data_sibling.exists():
            file_path = data_sibling

    # Prevent path traversal
    try:
        file_path.resolve().relative_to(Path(base_dir).resolve())
    except ValueError:
        return "Error: Invalid file path."

    # Validate range parameters
    if start_row is None or end_row is None or start_col is None or end_col is None:
        return (
            "Error: All range parameters are required: "
            "start_row, end_row, start_col, end_col."
        )

    try:
        start_row = int(start_row)
        end_row = int(end_row)
    except (TypeError, ValueError):
        return "Error: start_row and end_row must be integers."

    if start_row < 1 or end_row < start_row:
        return (
            f"Error: Invalid row range {start_row}-{end_row}. "
            f"Rows are 1-based and end_row must be >= start_row."
        )

    try:
        start_col_idx = _col_letter_to_index(start_col)
        end_col_idx = _col_letter_to_index(end_col)
    except (TypeError, AttributeError):
        return "Error: start_col and end_col must be column letters (e.g. 'A', 'Z')."

    if start_col_idx < 1 or end_col_idx < start_col_idx:
        return (
            f"Error: Invalid column range {start_col}-{end_col}. "
            f"end_col must be >= start_col."
        )

    num_rows = end_row - start_row + 1
    num_cols = end_col_idx - start_col_idx + 1
    cells = num_rows * num_cols
    if cells > READ_FILE_MAX_CELLS:
        return (
            f"Error: requested range is {cells} cells "
            f"({num_rows} rows x {num_cols} cols), which exceeds the "
            f"{READ_FILE_MAX_CELLS}-cell per-call limit. Split the request "
            f"into multiple smaller read_file calls."
        )

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            reader = csv_mod.reader(f)
            all_rows = list(reader)
    except UnicodeDecodeError:
        return f"Error: File '{filename}' could not be read as text."

    total_rows = len(all_rows)
    total_cols = max((len(r) for r in all_rows), default=0)

    if start_row > total_rows:
        return f"Error: start_row {start_row} exceeds file row count ({total_rows})."

    # Clamp end_row to actual file size
    end_row = min(end_row, total_rows)

    # Extract the requested range (1-based to 0-based)
    extracted = []
    for row in all_rows[start_row - 1 : end_row]:
        # Slice columns (1-based col index to 0-based)
        row_slice = row[start_col_idx - 1 : end_col_idx]
        extracted.append(row_slice)

    # Format as CSV text
    import io

    output = io.StringIO()
    writer = csv_mod.writer(output)
    writer.writerows(extracted)
    result_text = output.getvalue()

    # Add a header with range info
    header = (
        f"[Range: rows {start_row}-{end_row}, columns {start_col}-{end_col} "
        f"| File: {total_rows} rows x {total_cols} cols]\n"
    )

    # Formatting/Structure categories carry the sheet's formatting metadata
    # (merged cells / frozen panes) with the first read of each sheet —
    # under template 5 the file listings no longer inline it.
    appendix = ""
    if (
        category in ("Formatting", "Structure")
        and format_notes is not None
        and filename.endswith("_full.csv")
    ):
        note_key = (source, filename)
        if note_key not in format_notes:
            fmt_txt = Path(base_dir) / (
                filename[: -len("_full.csv")] + "_additional_format.txt"
            )
            if fmt_txt.exists():
                format_notes.add(note_key)
                try:
                    appendix = (
                        "\n[Sheet formatting metadata — sent once per sheet "
                        "for this category]\n" + fmt_txt.read_text(encoding="utf-8")
                    )
                except OSError:
                    appendix = ""

    return header + result_text + appendix


def single_pass_judge_case(
    task_folder: str,
    client: OpenAI,
    rubric_path: str,
    template_path: str,
    rubric_weight_path: str = None,
    model: str = JUDGE_MODEL,
    nocall: bool = False,
    use_existing: bool = True,
    attempt_model: str = None,
    run_calculation: bool = False,
    cached_solution_csv_dir: str = None,
    cached_attempt_csv_dir: str = None,
    cached_starting_csv_dir: str = None,
    attempt_sheet_name_filter: bool = False,
    ignore_sheets: list[str] | None = None,
    max_tool_rounds: int = SINGLE_PASS_MAX_ROUNDS,
    reasoning_effort: str | None = None,
    harness_verdicts: dict | None = None,
    accuracy_engine: str = "llm",
    max_forced_rounds: int | None = None,
    det_checks: dict | None = None,
):
    """The MBABench judge: ONE tool-calling conversation over every applicable check.

    `max_forced_rounds` caps the forced-finalization phase (default: the
    config's single_pass.max_forced_rounds). Smoke tests set it low together
    with max_tool_rounds — the config loader overwrites env vars, so the env
    override the README implies does not work; this parameter does.

    `det_checks` (judge v13): utils.det_checks.DetChecksRun.for_judge() of
    the deterministic checks the caller ran BEFORE this call; their verdicts
    are already in `harness_verdicts`. Passed through to _finalize_case
    (mode switch, scored_results.det_checks, det_checks.json). The LLM never
    sees either.

    Checks are globally numbered (the rubric's flattened order, the same
    1..132 numbering the suitability annotations validate against; gating
    leaves gaps rather than renumbering). Rows record the judge/prompt
    versions from config `single_pass.*`.
    """
    identity = resolve_judge_identity(model)
    if reasoning_effort is None:
        reasoning_effort = identity.effort
    elif reasoning_effort != identity.effort:
        logger.warning(
            f"reasoning_effort {reasoning_effort!r} overrides the effort "
            f"pinned by {model!r} ({identity.effort!r}); the effective value "
            f"is what gets recorded"
        )

    prep = _prepare_case(
        task_folder=task_folder,
        rubric_path=rubric_path,
        rubric_weight_path=rubric_weight_path,
        use_existing=use_existing,
        run_calculation=run_calculation,
        cached_solution_csv_dir=cached_solution_csv_dir,
        cached_attempt_csv_dir=cached_attempt_csv_dir,
        cached_starting_csv_dir=cached_starting_csv_dir,
        attempt_sheet_name_filter=attempt_sheet_name_filter,
        ignore_sheets=ignore_sheets,
    )

    output_dir = prep["output_dir"]
    cache_log_path = prep["cache_log_path"]
    golden_solution_dir = prep["golden_solution_dir"]
    ai_attempt_dir = prep["ai_attempt_dir"]
    starting_workbook_dir = prep["starting_workbook_dir"]
    weights_data = prep["weights_data"]
    context_file_path = prep["context_file_path"]
    rubric_json_path = prep["rubric_json_path"]
    start_time = prep["start_time"]
    task_folder_name = prep["task_folder_name"]

    # The versions this grading runs under (config single_pass.* and the
    # rubric constants); _finalize_case threads them into scores.json /
    # _metadata.json and grade.py records them in the grading row.
    versions = dict(prep["versions"])

    guidance = rubric_guidance.load_guidance(rubric_path)

    logger.info("=" * 80)
    logger.info("Single-Pass Judge Evaluation Workflow")
    logger.info("=" * 80)
    logger.info(
        f"Grading task: {task_folder_name}, model: {model}, "
        f"rubric: {versions['RUBRIC_VERSION']}, "
        f"judge version: {versions['JUDGE_VERSION']}"
    )
    logger.info("=" * 80)

    golden_solution_files = (
        prepare_directory_files(golden_solution_dir) if golden_solution_dir else {}
    )
    ai_attempt_files = prepare_directory_files(ai_attempt_dir) if ai_attempt_dir else {}
    starting_workbook_files = (
        prepare_directory_files(starting_workbook_dir)
        if starting_workbook_dir
        else {}
    )

    # Judge v6: list sheets in the workbook's TRUE tab order (the rubric
    # grades tab order; an alphabetical listing made that check unjudgeable)
    # and serve the properties block. Older caches without the properties
    # file degrade to the alphabetical listing.
    attempt_props = workbook_properties.load_properties(ai_attempt_dir)
    solution_props = workbook_properties.load_properties(golden_solution_dir)
    starting_props = workbook_properties.load_properties(starting_workbook_dir)
    attempt_file_list = workbook_properties.order_file_list(
        [f for f in ai_attempt_files if not f.endswith("_additional_format.txt")
         and not f.endswith(".json")],
        attempt_props,
    )
    solution_file_list = workbook_properties.order_file_list(
        [f for f in golden_solution_files if not f.endswith("_additional_format.txt")
         and not f.endswith(".json")],
        solution_props,
    )
    starting_file_list = workbook_properties.order_file_list(
        [f for f in starting_workbook_files if not f.endswith("_additional_format.txt")
         and not f.endswith(".json")],
        starting_props,
    )

    logger.info(f"\n  Attempt files: {attempt_file_list}")
    logger.info(f"  Solution files: {solution_file_list}")
    if starting_file_list:
        logger.info(f"  Starting-workbook files: {starting_file_list}")

    attempt_file_metadata = _build_file_metadata(ai_attempt_dir)
    solution_file_metadata = _build_file_metadata(golden_solution_dir)
    starting_file_metadata = (
        _build_file_metadata(starting_workbook_dir) if starting_workbook_dir else {}
    )

    context_messages = _build_agentic_context_messages(context_file_path)
    if context_file_path:
        logger.info(f"  Context file: {context_file_path.name}")

    if nocall:
        logger.info("\n--nocall flag set. Skipping API calls.")
        remove_log_file(cache_log_path)
        shutil.copy(cache_log_path, str(output_dir / "judge.log"))
        return

    with open(str(rubric_json_path), "r", encoding="utf-8") as _rf:
        _rubric_data = json.load(_rf)

    # Per-task suitability gating; the flat list keeps each check's ORIGINAL
    # global number, so gating produces gaps by design (numbers must be
    # stable across tasks — do not renumber).
    suitability = rubric_suitability.load_for_case(prep["task_path"], _rubric_data)
    numbered_all = numbered_rubric_checks(_rubric_data)
    if suitability is not None:
        _applicable_names = {
            cat: set(names) for cat, names in suitability["applicable"].items()
        }
        numbered_gated = [
            (no, cat, check)
            for no, cat, check in numbered_all
            if check["name"] in _applicable_names.get(cat, set())
        ]
        weights_data = rubric_suitability.build_effective_weights(
            weights_data, suitability["excluded"]
        )
        suitability_provenance = {"gated": True, **suitability["provenance"]}
        logger.info(
            f"  Suitability gating: {suitability_provenance['excluded_count']} "
            f"not_applicable check(s) excluded; "
            f"{suitability_provenance['applicable_count']} applicable"
        )
        _staged_annotation = prep["task_path"] / rubric_suitability.STAGED_FILENAME
        if _staged_annotation.exists():
            shutil.copy(
                str(_staged_annotation),
                str(output_dir / rubric_suitability.STAGED_FILENAME),
            )
    else:
        numbered_gated = numbered_all
        suitability_provenance = {"gated": False}
        if rubric_suitability.skip_requested():
            suitability_provenance["skipped_via_env"] = True

    check_ids = [str(no) for no, _, _ in numbered_gated]
    id_to_cat_name = {
        str(no): (cat, check["name"]) for no, cat, check in numbered_gated
    }

    recorder = _make_recorder(
        output_dir,
        mode="single_pass",
        model=model,
        reasoning_effort=reasoning_effort,
        attempt_model=attempt_model,
        versions=versions,
        check_order=prep["CHECK_ORDER"],
        rubric={"path": str(rubric_json_path), "md5": _file_md5(rubric_json_path)},
        weights={
            "path": str(rubric_weight_path) if rubric_weight_path else None,
            "md5": _file_md5(rubric_weight_path) if rubric_weight_path else None,
        },
        template_path=str(template_path),
        limits={
            "max_tool_rounds": max_tool_rounds,
            "context_token_limit": SINGLE_PASS_CONTEXT_TOKEN_LIMIT,
        },
        files={
            "golden_solution": sorted(golden_solution_files),
            "ai_attempt": sorted(ai_attempt_files),
            "starting_workbook": sorted(starting_workbook_files),
            "context": context_file_path.name if context_file_path else None,
        },
    )
    recorder.record_event(
        "rubric_suitability", **{k: v for k, v in suitability_provenance.items()
                                 if k in ("gated", "source", "excluded_count",
                                          "skipped_via_env")}
    )

    token_tracking = {
        "evaluations": {},
        "total_message_size": 0,
        "total_message_size_with_images": 0,
        "total_tokens": 0,
        "total_prompt_tokens": 0,
        "total_completion_tokens": 0,
        "total_cost": 0.0,
    }

    logger.info(
        f"\n[Single-pass] Starting evaluation over {len(check_ids)} checks..."
    )

    # Template 8+ (judge v6): guidance rendered as part of each check's
    # standard, the answer-equivalence rulebook under the Accuracy category,
    # and the workbook-properties block in the seed. Template 7 keeps the
    # advisory footnote form byte-for-byte.
    hardened = _prompt_version_at_least(versions["PROMPT_VERSION"], 8)
    rubric_checks_text = render_rubric_checks_flat(
        numbered_gated,
        guidance,
        guidance_style="standard" if hardened else "footnote",
        category_extras=(
            {"Accuracy": answer_rules.render_rules_text()} if hardened else None
        ),
    )
    # Evidence served for a retired check alone stays out of all three blocks
    # (workbook_properties.STYLED_EMPTY_CHECK); un-retiring renders it again.
    _retired_checks = suitability_provenance.get("retired_checks") or []
    compile_kwargs = dict(
        rubric_checks_text=rubric_checks_text,
        check_ids_text=", ".join(check_ids),
        num_checks=str(len(check_ids)),
        attempt_properties_text=workbook_properties.render_properties_text(
            attempt_props,
            set(attempt_file_list),
            origin=workbook_properties.load_origin(prep["task_path"]),
            retired_checks=_retired_checks,
        ),
        solution_properties_text=workbook_properties.render_properties_text(
            solution_props, set(solution_file_list), retired_checks=_retired_checks
        ),
        starting_properties_text=(
            workbook_properties.render_properties_text(
                starting_props, set(starting_file_list), retired_checks=_retired_checks
            )
            if starting_file_list
            else "  (starting workbook not available for this attempt)"
        ),
        attempt_files_text=_render_files_text_dims_only(
            attempt_file_list, attempt_file_metadata
        ),
        solution_files_text=_render_files_text_dims_only(
            solution_file_list, solution_file_metadata
        ),
        starting_files_text=(
            _render_files_text_dims_only(starting_file_list, starting_file_metadata)
            if starting_file_list
            else "  (starting workbook not available for this attempt)"
        ),
        general_guidance=(guidance or {}).get("general"),
    )
    if context_messages:
        compile_kwargs["context_messages"] = context_messages

    stages = compile_prompt(template_path, **compile_kwargs)
    seed_messages = list(stages[0])

    ALL = "all_checks"
    state = AgenticCategoryLoop(ALL, check_ids, seed_messages)
    format_notes_served: set = set()
    fail_nudge_count = 0

    # OpenAI + a real reasoning tier must route via /v1/responses (probed
    # 2026-09-01: chat/completions rejects function tools with any
    # reasoning_effort except 'none'). The adapter keeps the loop speaking
    # chat shapes; this side-table carries reasoning items across rounds.
    use_responses_api = openai_responses.wants_responses_api(
        identity.provider, reasoning_effort
    )
    responses_state = openai_responses.ReasoningState()
    if use_responses_api:
        logger.info(
            f"  [api] /v1/responses (OpenAI reasoning_effort="
            f"{reasoning_effort!r} with tools)"
        )
    # Anthropic graders speak the native Messages API in single-pass mode
    # (judge v6): the OpenAI-compat endpoint has no prompt caching, ignores
    # reasoning_effort and reports no cached tokens. The adapter keeps the
    # loop on chat shapes; this state replays thinking blocks across rounds.
    use_native_anthropic = anthropic_native.wants_native_anthropic(identity.provider)
    native_state = anthropic_native.NativeState()
    native_client = None
    if use_native_anthropic:
        native_client = get_native_anthropic_client(identity)
        logger.info(
            f"  [api] anthropic native messages (effort="
            f"{anthropic_native.map_effort(reasoning_effort)!r}, prompt caching on)"
        )

    cumulative_metrics = {
        "message_size": 0,
        "message_size_with_images": 0,
        "prompt_tokens": 0,
        "completion_tokens": 0,
        "total_tokens": 0,
        # Prompt-cache detail (judge v6): cached input is billed at a
        # fraction of the input rate; the cost meter needs these to be honest.
        "cached_tokens": 0,
        "cache_write_tokens": 0,
    }
    tool_call_stats: dict[str, int] = {
        t["function"]["name"]: 0 for t in SINGLE_PASS_JUDGE_TOOLS
    }
    tool_call_args_log: list[dict] = []
    for m in state.messages:
        size = _measure_message_chars(m)
        cumulative_metrics["message_size"] += size
        cumulative_metrics["message_size_with_images"] += size

    def _serving_category(parsed_args: dict) -> str:
        """Map the model's declared view onto the category-keyed serving rule."""
        view = str(parsed_args.get("view") or "").strip().lower()
        if view == "formatting":
            return "Formatting"
        if view == "structure":
            return "Structure"
        return "_data_view"  # any non-Formatting label serves the data view

    # Per-round flow bookkeeping for the read-refusal gate: `round_tokens`
    # is the running token estimate of everything in flight THIS round —
    # the request context at the calibrated ratio, plus each already-served
    # addition at CSV density — so a burst of parallel reads is gated
    # against its own running total, in the density of what was actually
    # added (a prose-calibrated ratio understated a CSV burst by ~30% and
    # let 1.02M real tokens through an 850K gate). Reset by _run_round.
    flow = {"round_tokens": 0, "refused": 0, "evicted": 0}

    def _dispatch_tool(tc, _args_parsed, phase: str) -> str:
        name = tc.function.name
        if name == "read_file":
            if phase == "forced_finalization":
                return (
                    "Error: tool 'read_file' is disabled in forced "
                    "finalization. Only record_check / append_mistake / "
                    "get_working_judgement are allowed."
                )
            result = _execute_read_file(
                tc,
                ai_attempt_dir,
                golden_solution_dir,
                category=_serving_category(_args_parsed),
                format_notes=format_notes_served,
                starting_dir=starting_workbook_dir,
            )
            # Context-budget gate on the MEASURED result (errors are tiny
            # and always served). Recoverable: the model evicts and retries.
            if not result.startswith("Error:"):
                refuse, projected, current = _read_refusal_check(
                    len(result),
                    flow["round_tokens"],
                    SINGLE_PASS_CONTEXT_TOKEN_LIMIT,
                )
                if refuse:
                    flow["refused"] += 1
                    logger.info(
                        f"      read_file REFUSED: +{len(result):,} chars "
                        f"would reach ~{projected // 1000}K of "
                        f"{SINGLE_PASS_CONTEXT_TOKEN_LIMIT // 1000}K tokens"
                    )
                    return _read_refusal_message(
                        len(result), projected, current,
                        SINGLE_PASS_CONTEXT_TOKEN_LIMIT, state.round,
                    )
            return result
        if name == "evict_tool_results":
            if phase == "forced_finalization":
                return (
                    "Error: tool 'evict_tool_results' is disabled in forced "
                    "finalization. Only record_check / append_mistake / "
                    "get_working_judgement are allowed."
                )
            try:
                before_round = int(_args_parsed.get("before_round", 0))
            except (ValueError, TypeError) as e:
                return f"Error: invalid arguments for evict_tool_results: {e}"
            before_round = min(before_round, state.round)
            result = state.evict(before_round)
            if result.startswith("Evicted "):
                flow["evicted"] += 1
                # The wire shrank; rebase the running round total on the
                # post-eviction context so subsequent reads gate correctly.
                flow["round_tokens"] = _estimate_wire_tokens(
                    _wire_char_total(state.messages), chars_per_token
                )
            return result
        if name in ("record_check", "append_mistake", "get_working_judgement"):
            return _execute_scratchpad_tool(tc, state.working)
        return f"Error: unknown tool '{name}'."

    def _run_round(phase: str, tools: list) -> str:
        """One API call + tool execution. Returns 'ok', 'stop', 'retry' or 'break'."""
        nonlocal chars_per_token, api_retries
        wire_chars_at_call = _wire_char_total(state.messages)
        flow.update(
            round_tokens=_estimate_wire_tokens(wire_chars_at_call, chars_per_token),
            refused=0,
            evicted=0,
        )
        _msgs = state.messages
        if identity.provider in ("anthropic", "openai", "tensorblock"):
            _msgs = strip_unsupported_anthropic_images(_msgs)
        _create_kwargs = {
            "model": identity.model,
            "messages": _msgs,
            "tools": tools,
        }
        if (
            reasoning_effort is not None
            and not use_responses_api
            and not use_native_anthropic
        ):
            _create_kwargs["reasoning_effort"] = (
                "none" if identity.provider in ("openai", "tensorblock") else reasoning_effort
            )
        if use_native_anthropic:
            _api = "anthropic_messages"
            _effective_effort = anthropic_native.map_effort(reasoning_effort)
        elif use_responses_api:
            _api = "responses"
            _effective_effort = reasoning_effort
        else:
            _api = "chat_completions"
            _effective_effort = _create_kwargs.get("reasoning_effort")
        _request_params = {"reasoning_effort": _effective_effort, "api": _api}
        _call_t0 = time.time()
        try:
            if use_native_anthropic:
                response = anthropic_native.create(
                    native_client,
                    model=identity.model,
                    messages=_msgs,
                    chat_tools=tools,
                    reasoning_effort=reasoning_effort,
                    state=native_state,
                )
            elif use_responses_api:
                response = openai_responses.create(
                    client,
                    model=identity.model,
                    messages=_msgs,
                    chat_tools=tools,
                    reasoning_effort=reasoning_effort,
                    state=responses_state,
                )
            else:
                response = client.chat.completions.create(**_create_kwargs)
        except Exception as e:
            recorder.record_call(
                mode="single_pass",
                category=ALL,
                round=state.round,
                purpose=phase,
                model=model,
                request_messages=_create_kwargs["messages"],
                request_params=_request_params,
                tools=[t["function"]["name"] for t in _create_kwargs["tools"]],
                error=str(e),
                t0=_call_t0,
            )
            err_str = str(e)
            if "maximum context length" in err_str or (
                "400" in err_str and "context" in err_str.lower()
            ):
                logger.error(
                    f"    Context-length overflow (round {state.round}): {e}. "
                    f"Stopping; partial judgement will be saved."
                )
                state.log_event("context_overflow", error=err_str[:500])
                recorder.record_event(
                    "context_overflow", category=ALL, round=state.round,
                    error=err_str[:500],
                )
                return "break"
            api_retries += 1
            if api_retries > max_api_retries:
                logger.error(f"    Giving up after {max_api_retries} API retries: {e}")
                return "break"
            wait = min(2**api_retries + 1, 30)
            logger.warning(
                f"    API error (round {state.round}, retry "
                f"{api_retries}/{max_api_retries}): {e}. Retrying in {wait}s..."
            )
            recorder.record_event(
                "api_retry", category=ALL, round=state.round,
                retry=api_retries, error=str(e)[:500],
            )
            time.sleep(wait)
            return "retry"

        recorder.record_call(
            mode="single_pass",
            category=ALL,
            round=state.round,
            purpose=phase,
            model=model,
            request_messages=_create_kwargs["messages"],
            request_params=_request_params,
            tools=[t["function"]["name"] for t in _create_kwargs["tools"]],
            response=response,
            t0=_call_t0,
        )

        if not response.choices:
            err_detail = (
                (response.model_extra or {}).get("error")
                or getattr(response, "error", None)
                or "no error field"
            )
            api_retries += 1
            if api_retries > max_api_retries:
                logger.error(
                    f"    Giving up after {max_api_retries} retries: "
                    f"empty choices. err={err_detail}"
                )
                state.log_event(
                    "empty_choices_giveup", error=str(err_detail)[:500]
                )
                return "break"
            wait = min(2**api_retries + 1, 30)
            logger.warning(
                f"    Empty choices (round {state.round}, retry "
                f"{api_retries}/{max_api_retries}): err={err_detail}. "
                f"Retrying in {wait}s..."
            )
            state.log_event("empty_choices", error=str(err_detail)[:500])
            recorder.record_event(
                "empty_choices", category=ALL, round=state.round,
                retry=api_retries, error=str(err_detail)[:500],
            )
            time.sleep(wait)
            return "retry"
        api_retries = 0

        usage = response.usage
        if usage:
            ub = usage_breakdown(usage)
            cumulative_metrics["prompt_tokens"] += ub["prompt_tokens"]
            cumulative_metrics["completion_tokens"] += ub["completion_tokens"]
            cumulative_metrics["total_tokens"] += ub["total_tokens"]
            cumulative_metrics["cached_tokens"] += ub["cached_tokens"]
            cumulative_metrics["cache_write_tokens"] += ub["cache_write_tokens"]
            if ub["prompt_tokens"]:
                # prompt_tokens is the TOTAL input (cached included) on every
                # path, so the chars/token calibration and the read gate are
                # unaffected by caching.
                state.last_prompt_tokens = ub["prompt_tokens"]
                if wire_chars_at_call > 0:
                    chars_per_token = wire_chars_at_call / ub["prompt_tokens"]

        msg = response.choices[0].message
        msg_size = _measure_message_chars(msg)
        cumulative_metrics["message_size"] += msg_size
        cumulative_metrics["message_size_with_images"] += msg_size
        flow["round_tokens"] += int(msg_size / _READ_GATE_RESULT_CPT)

        if msg.tool_calls:
            state.append_wire(msg, tag="model_tool_call")
            for tc in msg.tool_calls:
                name = tc.function.name
                args_preview = str(tc.function.arguments or "")[:200]
                logger.info(f"      Tool call: {name}({args_preview})")
                tool_call_stats[name] = tool_call_stats.get(name, 0) + 1
                try:
                    _args_parsed = json.loads(tc.function.arguments or "{}")
                except (json.JSONDecodeError, TypeError):
                    _args_parsed = {"_raw": tc.function.arguments}
                tool_call_args_log.append(
                    {
                        "round": state.round,
                        "phase": phase,
                        "tool": name,
                        "tool_call_id": tc.id,
                        "arguments": _args_parsed,
                    }
                )
                tool_result = _dispatch_tool(tc, _args_parsed, phase)
                tool_msg = {
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "content": tool_result,
                }
                state.append_wire(tool_msg, tag="tool_result")
                recorder.record_event(
                    "tool_exec",
                    category=ALL,
                    round=state.round,
                    phase=phase,
                    tool=name,
                    args=_args_parsed,
                    result=tool_result,
                )
                tsize = _measure_message_chars(tool_msg)
                cumulative_metrics["message_size"] += tsize
                cumulative_metrics["message_size_with_images"] += tsize
                flow["round_tokens"] += int(tsize / _READ_GATE_RESULT_CPT)
            return "ok"

        if msg.content:
            state.append_wire(msg, tag="assistant_text")
        return "stop"

    def _nudge(content: str, phase: str, **event) -> None:
        nudge_msg = {"role": "user", "content": content}
        state.append_wire(nudge_msg, tag="nudge")
        recorder.record_event(
            "nudge", category=ALL, round=state.round, phase=phase, **event
        )
        nsize = _measure_message_chars(nudge_msg)
        cumulative_metrics["message_size"] += nsize
        cumulative_metrics["message_size_with_images"] += nsize

    def _pending_ids() -> list[str]:
        return [i for i in state.working.check_letters if i in state.working.pending]

    api_retries = 0
    max_api_retries = 5
    chars_per_token = 0.0
    # Rounds in a row where reads were refused for context budget and the
    # model performed no successful eviction. A model that cannot fit its
    # own context is in a broken state — after 3 such rounds the grading
    # stops loudly (partial saved, row marked failed) instead of letting
    # forced finalization dress truncated evidence up as a real grade.
    consecutive_refusal_rounds = 0

    round_idx = 0
    while round_idx < max_tool_rounds:
        state.round = round_idx + 1
        if state.round % 25 == 0 or state.round == 1:
            logger.info(
                f"    Round {state.round}... "
                f"(recorded {len(state.working.working)}/{len(check_ids)})"
            )
        else:
            logger.info(f"    Round {state.round}...")

        wire_chars_pre = _wire_char_total(state.messages)
        wire_tokens_est = _estimate_wire_tokens(wire_chars_pre, chars_per_token)
        status, tier = _build_pressure_signal(
            wire_tokens_est,
            SINGLE_PASS_CONTEXT_TOKEN_LIMIT,
            state.round - 1,
        )
        logger.info(f"      Pressure ({tier}): {status.splitlines()[0]}")
        recorder.record_event(
            "pressure",
            category=ALL,
            round=state.round,
            tier=tier,
            estimated_tokens=wire_tokens_est,
            limit=SINGLE_PASS_CONTEXT_TOKEN_LIMIT,
        )
        pressure_msg = {"role": "user", "content": status}
        state.append_wire(pressure_msg, tag="pressure_note")
        psize = _measure_message_chars(pressure_msg)
        cumulative_metrics["message_size"] += psize
        cumulative_metrics["message_size_with_images"] += psize

        outcome = _run_round("main", SINGLE_PASS_JUDGE_TOOLS)
        if outcome == "retry":
            continue
        if outcome == "break":
            break
        if outcome == "ok":
            if flow["refused"] and not flow["evicted"]:
                consecutive_refusal_rounds += 1
                if consecutive_refusal_rounds >= 3:
                    logger.error(
                        f"    Context-budget deadlock: reads refused for "
                        f"{consecutive_refusal_rounds} consecutive rounds "
                        f"with no eviction. Stopping; partial judgement "
                        f"will be saved."
                    )
                    state.log_event(
                        "read_refusal_deadlock",
                        rounds=consecutive_refusal_rounds,
                    )
                    recorder.record_event(
                        "read_refusal_deadlock",
                        category=ALL,
                        round=state.round,
                        rounds=consecutive_refusal_rounds,
                    )
                    break
            else:
                consecutive_refusal_rounds = 0
            round_idx += 1
            continue

        # outcome == "stop": no tool calls — finalization rules
        if state.working.pending:
            missing = _pending_ids()
            preview = ", ".join(missing[:40]) + (
                f" (+{len(missing) - 40} more)" if len(missing) > 40 else ""
            )
            _nudge(
                f"You haven't recorded decisions for {len(missing)} check(s): "
                f"{preview}. Call record_check for each before concluding.",
                "main",
                pending_count=len(missing),
            )
            logger.info(f"      Nudged: {len(missing)} pending")
            round_idx += 1
            continue

        missing_mistakes = state.working.fails_missing_mistakes()
        if missing_mistakes and fail_nudge_count < 4:
            fail_nudge_count += 1
            _nudge(
                f"You recorded {', '.join(missing_mistakes)} as 'fail' but "
                f"appended no mistake. For each, either call append_mistake "
                f"with the concrete cell/range location and description of "
                f"the issue you found, or call record_check again with "
                f"decision 'pass' if there is in fact no concrete issue.",
                "fail_without_mistakes",
                checks=missing_mistakes,
            )
            logger.info(f"      Nudged: fail without mistakes {missing_mistakes}")
            round_idx += 1
            continue

        logger.info(
            f"      Model stopped with all checks recorded: "
            f"{len(state.working.working)}/{len(check_ids)}"
        )
        break

    # Forced finalization — the escape hatch when the round budget runs out, with a
    # ceiling sized for the whole rubric rather than one category.
    if round_idx >= max_tool_rounds and state.working.pending:
        missing_initial = _pending_ids()
        logger.warning(
            f"    Max rounds ({max_tool_rounds}) exhausted with "
            f"{len(missing_initial)} pending checks. Entering forced "
            f"finalization."
        )
        state.log_event(
            "forced_finalization_start",
            pending=missing_initial,
            max_tool_rounds=max_tool_rounds,
        )
        recorder.record_event(
            "forced_finalization_start",
            category=ALL,
            round=state.round,
            pending=missing_initial,
        )
        _nudge(
            f"You have exhausted the maximum number of tool-calling rounds "
            f"({max_tool_rounds}) but still have {len(missing_initial)} "
            f"pending checks: {', '.join(missing_initial)}. You must now "
            f"record your pass/fail decisions for ALL remaining pending "
            f"checks using record_check, based on the evidence you have "
            f"already gathered. No further file reads or evictions are "
            f"permitted; only record_check, append_mistake, and "
            f"get_working_judgement tools are available. Work in batches: "
            f"record at most 20 checks per turn, then stop and wait for the "
            f"next turn — do not deliberate over the whole list before "
            f"emitting any tool call. Output your best judgement now.",
            "forced_finalization",
            pending_count=len(missing_initial),
        )

        forced_round = 0
        api_retries = 0
        forced_cap = (
            max_forced_rounds if max_forced_rounds is not None
            else SINGLE_PASS_MAX_FORCED_ROUNDS
        )
        # Tool declarations stay IDENTICAL to the main loop's (Gemini binds
        # thought signatures to the request config — a reduced tool list
        # mid-conversation 400s). Restriction is enforced at execution time
        # in _dispatch_tool.
        while forced_round < forced_cap and state.working.pending:
            forced_round += 1
            state.round = max_tool_rounds + forced_round
            logger.info(
                f"    Forced finalization round {forced_round}/{forced_cap}..."
            )
            outcome = _run_round("forced_finalization", SINGLE_PASS_JUDGE_TOOLS)
            if outcome == "retry":
                forced_round -= 1
                continue
            if outcome == "break":
                break
            if outcome == "stop" and state.working.pending:
                missing_now = _pending_ids()
                _nudge(
                    f"You still have not recorded decisions for "
                    f"{len(missing_now)} check(s): "
                    f"{', '.join(missing_now[:40])}. Call record_check for "
                    f"the next batch of at most 20 pending checks now, then "
                    f"stop for the next turn.",
                    "forced_finalization",
                    pending_count=len(missing_now),
                )

        state.log_event(
            "forced_finalization_end",
            pending_after=_pending_ids(),
            forced_rounds_used=forced_round,
        )
        recorder.record_event(
            "forced_finalization_end",
            category=ALL,
            round=state.round,
            pending_after=_pending_ids(),
            forced_rounds_used=forced_round,
        )

    # Regroup the flat verdicts by category — scoring is unchanged and keys
    # on (category, check name), so after this point everything downstream
    # behaves exactly as before.
    all_responses: dict[str, list] = {}
    parse_failures: dict = {}

    if state.working.pending:
        missing = _pending_ids()
        logger.warning(
            f"    WARNING: finishing with {len(missing)} pending checks "
            f"(round {state.round})"
        )
        by_cat: dict[str, list] = {}
        for check_id in missing:
            cat, name = id_to_cat_name[check_id]
            by_cat.setdefault(cat, []).append(f"{check_id} ({name})")
        for cat, entries in by_cat.items():
            parse_failures[cat] = {
                "success": False,
                "count": len(entries),
                "responses": [f"Missing decisions for checks: {entries}"],
                "pending_checks": entries,
            }

    final_judgement = state.working.finalize()
    for item in final_judgement:
        check_id = item.get("check")
        cat, name = id_to_cat_name[check_id]
        item["name"] = name
        all_responses.setdefault(cat, []).append(item)

    state.append_synthetic_final(final_judgement)

    recorder.record_outcome(
        category=ALL,
        judgement=final_judgement,
        coverage=f"{len(state.working.working)}/{len(check_ids)}",
        pending=_pending_ids(),
        rounds_used=state.round,
        tool_call_stats=tool_call_stats,
    )

    token_tracking["evaluations"][ALL] = cumulative_metrics
    cost_info = calculate_cost(
        model,
        cumulative_metrics["prompt_tokens"],
        cumulative_metrics["completion_tokens"],
        cached_tokens=cumulative_metrics.get("cached_tokens", 0),
        cache_write_tokens=cumulative_metrics.get("cache_write_tokens", 0),
        provider=identity.provider,
    )
    token_tracking["evaluations"][ALL]["cost"] = cost_info["total_cost"]
    token_tracking["evaluations"][ALL]["cache_savings"] = cost_info["cache_savings"]
    token_tracking["total_cached_tokens"] = cumulative_metrics.get("cached_tokens", 0)
    token_tracking["total_cache_write_tokens"] = cumulative_metrics.get(
        "cache_write_tokens", 0
    )
    token_tracking["cache_savings"] = cost_info["cache_savings"]
    token_tracking["api"] = (
        "anthropic_messages" if use_native_anthropic
        else ("responses" if use_responses_api else "chat_completions")
    )
    token_tracking["evaluations"][ALL]["chars_per_token"] = (
        round(
            cumulative_metrics["message_size"] / cumulative_metrics["prompt_tokens"],
            2,
        )
        if cumulative_metrics["prompt_tokens"] > 0
        else 0
    )
    token_tracking["total_message_size"] = cumulative_metrics["message_size"]
    token_tracking["total_message_size_with_images"] = cumulative_metrics[
        "message_size_with_images"
    ]
    token_tracking["total_tokens"] = cumulative_metrics["total_tokens"]
    token_tracking["total_prompt_tokens"] = cumulative_metrics["prompt_tokens"]
    token_tracking["total_completion_tokens"] = cumulative_metrics[
        "completion_tokens"
    ]
    token_tracking["total_cost"] = cost_info["total_cost"]

    logger.info(
        f"    Tokens: {cumulative_metrics['prompt_tokens']:,} prompt + "
        f"{cumulative_metrics['completion_tokens']:,} completion | "
        f"Cost: ${cost_info['total_cost']:.6f}"
    )
    logger.info(
        f"    Final coverage: {len(state.working.working)}/{len(check_ids)} "
        f"checks recorded"
    )

    conversation_logs_dir = output_dir / "judge_conversation_logs"
    conversation_logs_dir.mkdir(parents=True, exist_ok=True)
    conversation_path = conversation_logs_dir / "conversation_messages_all_checks.json"
    with open(conversation_path, "w", encoding="utf-8") as f:
        json.dump(state.transcript, f, indent=2)
    dump_messages_yaml(state.transcript, conversation_path.with_suffix(".yaml"))

    total_tool_calls = sum(tool_call_stats.values())
    stats_str = (
        ", ".join(f"{k}={v}" for k, v in sorted(tool_call_stats.items()))
        if total_tool_calls
        else "(none)"
    )
    logger.info(f"    Tool calls: {total_tool_calls} total | {stats_str}")

    tool_calls_path = conversation_logs_dir / "tool_calls_all_checks.json"
    with open(tool_calls_path, "w", encoding="utf-8") as f:
        json.dump(
            {"stats": tool_call_stats, "calls": tool_call_args_log}, f, indent=2
        )

    return _finalize_case(
        all_responses=all_responses,
        output_dir=output_dir,
        weights_data=weights_data,
        token_tracking=token_tracking,
        model=model,
        attempt_model=attempt_model,
        task_folder_name=task_folder_name,
        golden_solution_files=golden_solution_files,
        ai_attempt_files=ai_attempt_files,
        context_file_path=context_file_path,
        start_time=start_time,
        cache_log_path=cache_log_path,
        versions=versions,
        golden_solution_dir=golden_solution_dir,
        ai_attempt_dir=ai_attempt_dir,
        starting_workbook_dir=starting_workbook_dir,
        parse_failures=parse_failures,
        agentic=True,
        reasoning_effort=reasoning_effort,
        grader_identity=identity.settings(),
        recorder=recorder,
        suitability_provenance=suitability_provenance,
        formula_cache_provenance=prep.get("formula_cache_provenance"),
        harness_verdicts=harness_verdicts,
        accuracy_engine=accuracy_engine,
        det_checks=det_checks,
    )


# ============================================================================
# CLI Entry Point
# ============================================================================


def main(args):
    """Grade args.folder_to_grade: deterministic checks, answer check, judge."""
    load_project_configs(verbose=True)

    # Resolve paths from config
    rubric_path = str(
        relative_path_from_project_root(load_env_var("JUDGE_RUBRIC", required=True))
    )
    rubric_weight_path = str(
        relative_path_from_project_root(load_env_var("JUDGE_RUBRIC_WEIGHT", required=True))
    )
    template_path = str(
        relative_path_from_project_root(
            load_env_var("SINGLE_PASS_PROMPT_TEMPLATE", required=True)
        )
    )

    # Fail fast on an unregistered label, before any file work.
    identity = resolve_judge_identity(args.model)
    client = get_client(identity)

    # Deterministic rubric checks first, before the answer check and the
    # judge, and NOT score-neutral: a check that cannot grade the delivered
    # workbook raises DetChecksError here (no fallback, no API call, and no
    # answer-check LibreOffice run). Every LibreOffice run goes through the
    # machine-wide memory guard; when it cannot run now the error carries
    # retry_later (re-run later). A hand-assembled folder needs the
    # _attempt_origin.json sidecar for File extension (.xlsx) (77). A legacy
    # .xls delivery is graded only with --run-calculation, on LibreOffice's
    # .xlsx conversion. Same wiring as grade.grade_single_attempt.
    task_folder = Path(args.folder_to_grade)
    det_run = run_det_checks(
        task_folder,
        rubric_path=rubric_path,
        weights_path=rubric_weight_path,
        mode=args.det_checks,
        run_calculation=args.run_calculation,
    )

    # Harness answer check: deterministic Questions-sheet comparison whose
    # verdicts the judge can adopt for the Accuracy checks (--accuracy-check
    # harness). Score-neutral on failure: the LLM's own verdicts then stand.
    harness_verdicts = {}
    try:
        solution_xlsx = find_golden_solution_file(task_folder)
        hardcoded_counts = str(
            load_env_var("SINGLE_PASS_HARDCODED_COUNTS", default="true")
        ).strip().lower() in ("1", "true", "yes")
        ac_result = run_answer_check(
            task_folder / "ai_attempt.xlsx",
            solution_xlsx,
            output_json_path=task_folder / "answer_check.json",
            hardcoded_counts=hardcoded_counts,
        )
        logger.info(f"  [answer_check] {summary_block(ac_result)}")
        harness_verdicts = ac_result.get("harness_verdicts") or {}
    except LibreOfficeUnavailable:
        # LibreOffice could not run now (machine-wide guard exhausted): fail
        # loudly, re-run later
        raise
    except Exception as e:  # noqa: BLE001 — score-neutral by design
        logger.warning(f"  [answer_check] skipped on error: {e}")
    harness_verdicts = merge_harness_verdicts(harness_verdicts, det_run.harness_verdicts)

    single_pass_judge_case(
        task_folder=args.folder_to_grade,
        client=client,
        rubric_path=rubric_path,
        template_path=template_path,
        rubric_weight_path=rubric_weight_path,
        model=args.model,
        nocall=args.nocall,
        use_existing=not args.no_use_existing,
        run_calculation=args.run_calculation,
        attempt_sheet_name_filter=args.attempt_sheet_name_filter,
        ignore_sheets=args.ignore_sheets,
        max_tool_rounds=args.max_tool_rounds,
        max_forced_rounds=args.max_forced_rounds,
        reasoning_effort=args.reasoning_effort,
        harness_verdicts=harness_verdicts,
        accuracy_engine=args.accuracy_check,
        det_checks=det_run.for_judge(),
    )


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description=(
            "Grade one assembled task folder (ai_attempt.xlsx, solution/, "
            "starting/, context files, rubric_suitability.json) with the "
            "MBABench judge. Results go to <folder>/judge_results/."
        )
    )
    parser.add_argument(
        "-f",
        "--folder-to-grade",
        required=True,
        help="Path to folder containing the attempt, solution files, and possibly context files. "
        "Relative paths are interpreted from the project root directory.",
    )
    parser.add_argument(
        "--model",
        default=JUDGE_MODEL,
        help=f"Grader label from judge_identities.yaml (default: {JUDGE_MODEL})",
    )
    parser.add_argument(
        "--nocall",
        action="store_true",
        help="Skip API calls (for testing file preparation only)",
    )
    parser.add_argument(
        "--no-use-existing",
        default=True,
        type=str2bool,
        help="Force re-extraction of CSV files even if they already exist",
    )
    parser.add_argument(
        "--run-calculation",
        action="store_true",
        help="Run Excel formula calculations via LibreOffice before extracting CSVs",
    )
    parser.add_argument(
        "--attempt-sheet-name-filter",
        action="store_true",
        help="Filter attempt sheets to only include those starting with 'answers_' or 'model_', stripping the prefix",
    )
    parser.add_argument(
        "--ignore-sheets",
        nargs="+",
        default=None,
        help=(
            "Sheet names to drop from attempt, solution and starting workbook "
            "before grading (case-insensitive). Default: none — every sheet, "
            "including the cover, is served to the judge."
        ),
    )
    parser.add_argument(
        "--accuracy-check",
        choices=["harness", "llm"],
        default="harness",
        help=(
            "Which engine decides the answer-accuracy checks: 'harness' "
            "(default) adopts the deterministic Questions-sheet comparison "
            "where it is confident, 'llm' keeps the judge's own verdicts. Both "
            "are always recorded in scores.json."
        ),
    )
    add_det_checks_arg(parser)
    parser.add_argument(
        "--max-tool-rounds",
        type=int,
        default=SINGLE_PASS_MAX_ROUNDS,
        help=f"Maximum number of tool-calling rounds (default: {SINGLE_PASS_MAX_ROUNDS})",
    )
    parser.add_argument(
        "--max-forced-rounds",
        type=int,
        default=None,
        help=(
            "Cap on forced-finalization rounds after the main round budget is "
            "exhausted (default: config single_pass.max_forced_rounds)."
        ),
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

    args = parser.parse_args()
    args.folder_to_grade = get_absolute_path(args.folder_to_grade)

    logger.info(
        f"Running local judge with parameters: {json.dumps(vars(args), indent=2)}"
    )

    main(args)
