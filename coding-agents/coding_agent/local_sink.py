"""Local sink: an attempt recorded under <output_root>/attempts/.

    <output_root>/attempts/<agent_model_name>/
      task_attempts.jsonl                one row per attempt (columns below)
      task_id=<N>/<YYYYmmdd_HHMMSS>/     the attempt's files: solution.xlsx
                                         first, then PROMPT.md, transcript,
                                         telemetry, verdict, trajectory, the
                                         run config and the prompt snapshot

The row carries the shared column set every pipeline writes and the judge
(judge/main_scripts/grade.py) reads; `id` is the millisecond epoch of the
write (unique across the lanes on one machine). File columns hold POSIX
paths relative to <output_root>. The first Excel file in attempt_files is
the workbook the judge grades.

A cohort label containing a slash (claudecode_anthropic/claude-fable-5-1-max)
nests one directory — intended.
"""
import fcntl
import json
import shutil
import time
from datetime import datetime
from pathlib import Path

from .repo_config import output_relative

ATTEMPTS_SUBDIR = "attempts"
ATTEMPTS_JSONL = "task_attempts.jsonl"

# The attempt-row columns, in order. Identical across the four pipelines.
ATTEMPT_COLUMNS = (
    "id", "task_id", "agent_model_name", "agent_model_type", "attempt_files", "prompt_files",
    "start_time", "end_time", "time_taken_min", "cost", "prompt_version", "agent_failed",
    "agent_failed_reason", "deprecated", "created_at", "context_reduced", "deprecated_reason",
    "updated_at", "extra_configs",
)

# The runner-side artifacts copied from the attempt dir, in this order.
ATTEMPT_ARTIFACTS = ("transcript.jsonl", "telemetry.json", "verdict.json",
                     "trajectory.jsonl.gz", "run_config.yaml")


def _iso(dt: datetime) -> str:
    """ISO-8601 with an offset."""
    return (dt if dt.tzinfo else dt.astimezone()).isoformat()


def cohort_dir(output_root: Path, agent_model_name: str) -> Path:
    return Path(output_root).joinpath(ATTEMPTS_SUBDIR, *agent_model_name.split("/"))


def attempted_task_ids(output_root: Path, agent_model_name: str,
                       prompt_version: int | None = None) -> set:
    """Task ids this cohort already has a non-failed, non-deprecated row for
    (under this prompt version) — the skip-if-attempted rule a relaunch
    applies."""
    path = cohort_dir(output_root, agent_model_name) / ATTEMPTS_JSONL
    if not path.exists():
        return set()
    done = set()
    for line in path.read_text().splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            row = json.loads(line)
        except ValueError:  # a half-written line from a killed lane
            continue
        if row.get("deprecated") or row.get("agent_failed"):
            continue
        if prompt_version is not None and row.get("prompt_version") != prompt_version:
            continue
        if row.get("task_id") is not None:
            done.add(int(row["task_id"]))
    return done


def _append_row(path: Path, row: dict) -> None:
    """Append one JSON line under an exclusive lock: lanes on one machine
    append to the same cohort file concurrently."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            f.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")
            f.flush()
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


def write_attempt(output_root: Path, agent_model_name: str, task_id: int,
                  started_at: datetime, ended_at: datetime,
                  solution_path: Path | None, workspace: Path, attempt_dir: Path,
                  prompt_paths: list, time_taken_min: float, cost,
                  prompt_version: int, agent_failed: bool,
                  agent_failed_reason: str | None, extra_configs: dict) -> tuple:
    """Copy the artifacts and append the row. Returns (row, attempt folder)."""
    ts = ended_at.strftime("%Y%m%d_%H%M%S")
    dest = cohort_dir(output_root, agent_model_name) / f"task_id={task_id}" / ts
    dest.mkdir(parents=True, exist_ok=True)

    # Attempt artifacts — solution first (the judge takes the first xlsx).
    attempt_files = []
    if solution_path and Path(solution_path).exists():
        shutil.copy2(solution_path, dest / "solution.xlsx")
        attempt_files.append(output_relative(dest / "solution.xlsx"))
    prompt_md = Path(workspace) / "PROMPT.md"
    if prompt_md.exists():
        shutil.copy2(prompt_md, dest / "PROMPT.md")
        attempt_files.append(output_relative(dest / "PROMPT.md"))
    for name in ATTEMPT_ARTIFACTS:
        src = Path(attempt_dir) / name
        if src.exists():
            shutil.copy2(src, dest / name)
            attempt_files.append(output_relative(dest / name))

    # Prompt snapshot (system + template + attachments + workspace extras).
    prompt_files = []
    for src in prompt_paths:
        src = Path(src)
        shutil.copy2(src, dest / src.name)
        prompt_files.append(output_relative(dest / src.name))

    now = datetime.now().astimezone()
    row = {
        "id": int(time.time() * 1000),
        "task_id": task_id,
        "agent_model_name": agent_model_name,
        "agent_model_type": "coding_cli",
        "attempt_files": attempt_files,
        "prompt_files": prompt_files,
        "start_time": _iso(started_at),
        "end_time": _iso(ended_at),
        "time_taken_min": time_taken_min,
        "cost": cost,
        "prompt_version": prompt_version,
        "agent_failed": agent_failed,
        "agent_failed_reason": agent_failed_reason,
        "deprecated": False,
        "created_at": _iso(now),
        "context_reduced": None,
        "deprecated_reason": None,
        "updated_at": _iso(now),
        "extra_configs": extra_configs,
    }
    assert tuple(row) == ATTEMPT_COLUMNS
    _append_row(cohort_dir(output_root, agent_model_name) / ATTEMPTS_JSONL, row)
    return row, dest
