"""Per-task rubric suitability gating.

Each task's annotation (`<data_root>/tasks/task_id=<N>/rubric_suitability.json`)
marks each of the 132 rubric_9 checks applicable or not_applicable for that
task. The judge skips not_applicable checks entirely — never prompted, never
scored — and renormalizes within the category (emergent: the scorer divides by
the summed weight of the checks it is given). CategoryWeights are untouched
unless a category loses every check (defensive; no task does today).

`conditional` flags are recorded for audit, ignored for scoring. The
annotation's rubric_version hash is stored as provenance only — validation is
by (no, category, name) match against the loaded rubric, which is exact 132/132.

Enforcement (wired in judge.single_pass_judge_case): a grading without a
staged annotation refuses to run; JUDGE_SKIP_SUITABILITY=1 grades ungated and
records that in scored_results.rubric_suitability.
"""

import copy
import json
import os
from pathlib import Path

try:
    from .logger import logger
except ImportError:  # imported as a bare module (utils/ on sys.path)
    from logger import logger

STAGED_FILENAME = "rubric_suitability.json"
SKIP_ENV = "JUDGE_SKIP_SUITABILITY"

# Retired checks. The rubric keeps its 132-position
# numbering — toys, annotations and every grading to date are keyed by
# position — so a check is retired by RULE, not by deletion: it is forced
# not_applicable on every task, never prompted, never scored, and the
# category rescales around it exactly as suitability gating does. Which
# numbers are retired comes from project_configs.yaml (judge.retired_checks,
# "28,37,101"); this table pins the NAME each number must carry so a
# regenerated/renumbered rubric can never retire the wrong check. A pin only
# permits retirement — the config list decides — so un-retiring a check is a
# config edit and its pin stays here (28 is meant to return once redefined).
RETIRED_ENV_KEY = "JUDGE_RETIRED_CHECKS"
RETIRED_CHECK_NAMES = {
    28: ("Error Checks", "No unused formatting"),   # judge v11, 2026-09-19
    37: ("Flexibility", "M&A / divestiture flexibility"),
    101: ("Purpose & Scope", "Architecture suited to audience"),
}


def _flat(rubric: dict) -> list:
    return [(cat, c["name"]) for cat, checks in rubric.items() for c in checks]


def retired_check_numbers(rubric: dict | None = None) -> list[int]:
    """Configured retired check numbers, validated against `rubric` when given.

    Raises SuitabilityError if a configured number is not in
    RETIRED_CHECK_NAMES or the rubric's check at that position carries a
    different (category, name) — refusing to grade beats retiring the wrong
    check. Positions beyond the rubric are ignored.
    """
    try:
        from .misc_utils import load_env_var
    except ImportError:  # bare-module import path
        from misc_utils import load_env_var
    raw = str(load_env_var(RETIRED_ENV_KEY, default="") or "").strip()
    if not raw or raw.lower() in ("none", "null", "[]"):
        return []
    numbers = []
    for tok in raw.replace(";", ",").split(","):
        tok = tok.strip()
        if not tok:
            continue
        if not tok.isdigit():
            raise SuitabilityError(f"{RETIRED_ENV_KEY}: {tok!r} is not a check number")
        numbers.append(int(tok))
    out = []
    for no in sorted(set(numbers)):
        if no not in RETIRED_CHECK_NAMES:
            raise SuitabilityError(
                f"{RETIRED_ENV_KEY} names check {no}, which has no entry in "
                f"RETIRED_CHECK_NAMES — add the (category, name) pin before retiring it"
            )
        if rubric is not None:
            flat = _flat(rubric)
            if no > len(flat):
                continue
            if flat[no - 1] != RETIRED_CHECK_NAMES[no]:
                raise SuitabilityError(
                    f"retired check {no} is pinned to {RETIRED_CHECK_NAMES[no]!r} but the loaded "
                    f"rubric has {flat[no - 1]!r} at that position — numbering drift; refusing to grade"
                )
        out.append(no)
    return out


def _apply_retirement(annotation: dict, retired: list[int]) -> list[int]:
    """Force retired numbers to not_applicable in a (validated-shape)
    annotation. Returns the numbers actually changed or already excluded."""
    applied = []
    for e in annotation.get("rubrics", []):
        if e.get("no") in retired:
            e["verdict"] = "not_applicable"
            e["retired"] = True
            applied.append(e["no"])
    return applied


def _all_applicable_annotation(rubric: dict) -> dict:
    rubrics = [
        {"no": i, "category": cat, "name": name, "verdict": "applicable", "conditional": False}
        for i, (cat, name) in enumerate(_flat(rubric), 1)
    ]
    return {"annotator": "retired-checks-only", "created_at": None, "rubric_version": None,
            "complete": True, "rubrics": rubrics}


class SuitabilityError(Exception):
    """A v2 grading cannot proceed without a valid suitability annotation."""


# --------------------------------------------------------------------------
# Per-task annotation under <data_root> (grade.py stages it into the task folder)
# --------------------------------------------------------------------------


def bundled_path(task_id: int) -> Path:
    """Where the task's annotation lives in the downloaded benchmark."""
    try:
        from . import local_store
    except ImportError:  # bare-module import path
        import local_store
    return local_store.suitability_path(task_id)


def load_bundled(task_id: int) -> tuple[dict, str]:
    """The task's annotation as (annotation, source path); raises when absent
    or not complete."""
    path = bundled_path(task_id)
    if not path.is_file():
        raise SuitabilityError(
            f"no rubric suitability annotation for task {task_id}: {path} does "
            f"not exist (it ships with the task folder in the downloaded "
            f"benchmark). Set {SKIP_ENV}=1 to grade ungated."
        )
    annotation = json.loads(path.read_text(encoding="utf-8"))
    if annotation.get("complete") is not True:
        raise SuitabilityError(f"annotation {path} is not complete")
    return annotation, str(path)


# --------------------------------------------------------------------------
# Validation + filter construction (pure; used by judge.py and tests)
# --------------------------------------------------------------------------


def validate_annotation(annotation: dict, rubric: dict) -> None:
    """Require an exact (no, category, name) match against the loaded rubric.

    The rubric's flattened order (category key order, then within-category
    order) must equal the annotation's `no` 1..N sequence. Refuses on any
    drift — do not grade against a rubric edition the annotation didn't see.
    """
    entries = annotation.get("rubrics")
    if not isinstance(entries, list) or not entries:
        raise SuitabilityError("annotation has no 'rubrics' entries")

    flat = [(cat, c["name"]) for cat, checks in rubric.items() for c in checks]
    if len(entries) != len(flat):
        raise SuitabilityError(
            f"annotation covers {len(entries)} checks, rubric has {len(flat)}"
        )
    mismatches = []
    for e in entries:
        no = e.get("no")
        if not isinstance(no, int) or not (1 <= no <= len(flat)):
            mismatches.append(f"bad 'no': {no!r}")
            continue
        expected = flat[no - 1]
        if (e.get("category"), e.get("name")) != expected:
            mismatches.append(
                f"no={no}: annotation ({e.get('category')!r}, {e.get('name')!r}) "
                f"!= rubric {expected!r}"
            )
    if mismatches:
        raise SuitabilityError(
            "annotation does not match the loaded rubric:\n  "
            + "\n  ".join(mismatches[:10])
            + (f"\n  ... {len(mismatches) - 10} more" if len(mismatches) > 10 else "")
        )
    verdicts = {e.get("verdict") for e in entries}
    unknown = verdicts - {"applicable", "not_applicable"}
    if unknown:
        raise SuitabilityError(f"unknown verdict value(s): {sorted(unknown)}")


def build_suitability(annotation: dict, rubric: dict, source: str = None) -> dict:
    """Validated annotation -> filter + provenance.

    Returns {
      "applicable": {category: [names]},   # in rubric order
      "excluded":   {category: [names]},   # in rubric order
      "provenance": {... for scored_results / _metadata.json ...},
    }
    """
    validate_annotation(annotation, rubric)
    by_no = {e["no"]: e for e in annotation["rubrics"]}
    applicable: dict[str, list] = {}
    excluded: dict[str, list] = {}
    no = 0
    for cat, checks in rubric.items():
        applicable[cat] = []
        excluded[cat] = []
        for c in checks:
            no += 1
            entry = by_no[no]
            if entry["verdict"] == "applicable":
                applicable[cat].append(c["name"])
            else:
                excluded[cat].append(c["name"])
    n_excluded = sum(len(v) for v in excluded.values())
    provenance = {
        "source": source,
        "annotator": annotation.get("annotator"),
        "created_at": annotation.get("created_at"),
        "rubric_version": annotation.get("rubric_version"),
        "applicable_count": sum(len(v) for v in applicable.values()),
        "excluded_count": n_excluded,
        "excluded": {cat: names for cat, names in excluded.items() if names},
        "conditional_flags": sorted(
            e["no"] for e in annotation["rubrics"] if e.get("conditional")
        ),
        "retired_checks": sorted(
            e["no"] for e in annotation["rubrics"] if e.get("retired")
        ),
    }
    return {"applicable": applicable, "excluded": excluded, "provenance": provenance}


def build_effective_weights(weights_data: dict, excluded: dict) -> dict:
    """Deep-copy the weights and drop excluded checks from each category.

    Renormalization within a category is emergent — calculate_scores divides
    by the summed weight of the checks present. If a category loses every
    check (cannot happen with today's annotations), it is dropped from both
    the per-category lists and CategoryWeights, and CategoryWeights is
    renormalized over the remainder — logged loudly.
    """
    eff = copy.deepcopy(weights_data)
    emptied = []
    for cat, names in excluded.items():
        if not names or cat not in eff:
            continue
        drop = set(names)
        eff[cat] = [c for c in eff[cat] if c["name"] not in drop]
        if not eff[cat]:
            emptied.append(cat)
    if emptied:
        logger.warning(
            f"  SUITABILITY: categories with ZERO applicable checks: {emptied} — "
            f"dropping them and renormalizing CategoryWeights (defensive path; "
            f"no known annotation does this)"
        )
        cw = eff["CategoryWeights"][0]
        for cat in emptied:
            cw.pop(cat, None)
            eff.pop(cat, None)
        total = sum(cw.values())
        if total > 0:
            eff["CategoryWeights"] = [{c: w / total for c, w in cw.items()}]
    return eff


# --------------------------------------------------------------------------
# Case-level loader (judge.py entry point)
# --------------------------------------------------------------------------


def skip_requested() -> bool:
    return os.environ.get(SKIP_ENV) == "1"


def load_for_case(task_path: Path, rubric: dict) -> dict | None:
    """Load the staged annotation for a task folder.

    Returns the build_suitability() dict, or None to grade ungated (only with
    the JUDGE_SKIP_SUITABILITY=1 escape hatch and no retired checks — the
    caller records the skip in scored_results).

    Raises SuitabilityError when the annotation is missing or invalid and no
    escape hatch is set.
    """
    staged = Path(task_path) / STAGED_FILENAME
    retired = retired_check_numbers(rubric)

    def _retired_only(extra: dict) -> dict | None:
        """No annotation in play: retire by rule alone, or None."""
        if not retired:
            return None
        annotation = _all_applicable_annotation(rubric)
        _apply_retirement(annotation, retired)
        out = build_suitability(annotation, rubric)
        out["provenance"].update({"gated": False, **extra})
        logger.info(f"  SUITABILITY: retired check(s) {retired} excluded by rule (no task annotation)")
        return out

    if skip_requested():
        logger.warning(
            f"  SUITABILITY: {SKIP_ENV}=1 — grading UNGATED (all checks except retired "
            f"{retired or 'none'}); recorded in scored_results"
        )
        return _retired_only({"skipped_via_env": True})
    if not staged.exists():
        raise SuitabilityError(
            f"grading requires a rubric suitability annotation, but {staged} is "
            f"missing (grade.py stages it from the task folder under the data "
            f"root; for standalone judge.py runs copy "
            f"tasks/task_id=<N>/{STAGED_FILENAME} there). Set {SKIP_ENV}=1 to "
            f"grade ungated."
        )
    annotation = json.loads(staged.read_text())
    if annotation.get("complete") is not True:
        raise SuitabilityError(f"staged annotation {staged} is not complete")
    meta = annotation.get("_staging", {})
    if retired:
        validate_annotation(annotation, rubric)   # shape first, so the numbers mean what we think
        _apply_retirement(annotation, retired)
        logger.info(f"  SUITABILITY: retired check(s) {retired} forced not_applicable")
    out = build_suitability(annotation, rubric, source=meta.get("source"))
    out["provenance"]["gated"] = True
    return out
