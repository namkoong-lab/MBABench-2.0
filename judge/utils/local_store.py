"""The local store: the downloaded benchmark under `<data_root>/`, run output
under `<output_root>/`.

This is the whole filesystem side of a grading. The judge itself is untouched;
what lives here is only WHERE an attempt, its task and the golden solution are
read from, and WHERE the grading row and its staged files are written.

Layout (the contract shared by the four agent pipelines and the judge):

  <data_root>/tasks.jsonl                          one row per task (same fields as task.json)
  <data_root>/tasks/task_id=<N>/task.json          task_id, task_name, starting_files, solution_files, ...
                                /rubric_suitability.json  per-task rubric applicability annotation
                                /starting_files/<name>    what the agent received
                                /solution_files/<name>    the golden solution (+ context files)
  <output_root>/attempts/<agent_model_name>/task_attempts.jsonl   one row per attempt
  <output_root>/attempts/<agent_model_name>/task_id=<N>/<stamp>/  that attempt's files
  <output_root>/gradings/gradings.jsonl            one gradings-shaped row per grading
  <output_root>/gradings/<grading_id>/             that grading's staged files

Paths stored in task.json are POSIX and relative to <data_root>; paths stored
in an attempt row (`attempt_files`, `prompt_files`) are POSIX and relative to
<output_root>. An absolute path is accepted as is.

Roots, first hit wins (utils.repo_config): $MBABENCH_DATA_ROOT /
$MBABENCH_OUTPUT_ROOT, config/config.yaml `local.data_root` /
`local.output_root`, then `data` / `outputs` under the repository root.
"""

from __future__ import annotations

import fcntl
import json
import shutil
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    from . import repo_config
    from .logger import logger
except ImportError:  # imported as a bare module (utils/ on sys.path)
    import repo_config
    from logger import logger

REPO_ROOT = repo_config.REPO_ROOT

TASKS_DIRNAME = "tasks"
TASK_FILENAME = "task.json"
SUITABILITY_FILENAME = "rubric_suitability.json"
ATTEMPTS_DIRNAME = "attempts"
ATTEMPTS_FILENAME = "task_attempts.jsonl"
GRADINGS_DIRNAME = "gradings"
GRADINGS_FILENAME = "gradings.jsonl"

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")

# The attempt-row columns every pipeline writes, in order (the judge reads
# id, task_id, agent_model_name, agent_model_type, attempt_files,
# prompt_files, prompt_version, agent_failed and deprecated).
ATTEMPT_COLUMNS = (
    "id", "task_id", "agent_model_name", "agent_model_type", "attempt_files",
    "prompt_files", "start_time", "end_time", "time_taken_min", "cost",
    "prompt_version", "agent_failed", "agent_failed_reason", "deprecated",
    "created_at", "context_reduced", "deprecated_reason", "updated_at",
    "extra_configs",
)

# Every column of a grading row, in order, as grade.py writes them.
GRADING_COLUMNS = (
    "id",
    "task_id",
    "attempt_id",
    "grader_model",
    "grader_prompts",
    "grader_response",
    "accuracy_grade",
    "formula_grade",
    "format_grade",
    "rubric_version",
    "rubric_weight_version",
    "prompt_version",
    "scored_results",
    "time_elapsed_min",
    "cost",
    "raw_files_path",
    "raw_files",
    "errors_encountered",
    "failed",
    "failed_reason",
    "deprecated",
    "deprecated_reason",
    "solution_context_reduced",
    "attempt_context_reduced",
    "context_reduced_details",
    "agentic_mode",
    "judge_version",
    "created_at",
    "updated_at",
    "grader_reasoning",
)


class LocalStoreError(Exception):
    """The local store cannot serve what was asked of it."""


# ---------------------------------------------------------------------------
# Roots and paths
# ---------------------------------------------------------------------------


def data_root() -> Path:
    return repo_config.data_root()


def output_root() -> Path:
    return repo_config.output_root()


def resolve_data_path(stored) -> Path:
    """Absolute path for a path stored in task.json (relative to <data_root>)."""
    p = Path(str(stored)).expanduser()
    return p if p.is_absolute() else data_root() / p


def resolve_output_path(stored) -> Path:
    """Absolute path for a path stored in an attempt row (relative to <output_root>)."""
    p = Path(str(stored)).expanduser()
    return p if p.is_absolute() else output_root() / p


def display_path(path) -> str:
    """`path` as a repo-relative POSIX string when it is inside the repo."""
    p = Path(path).resolve()
    try:
        return p.relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return p.as_posix()


def output_relative(path) -> str:
    """`path` as a POSIX string relative to <output_root> (how rows store it)."""
    p = Path(path).resolve()
    try:
        return p.relative_to(output_root().resolve()).as_posix()
    except ValueError:
        return p.as_posix()


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------


def tasks_dir() -> Path:
    return data_root() / TASKS_DIRNAME


def task_dir(task_id: int) -> Path:
    return tasks_dir() / f"task_id={task_id}"


def suitability_path(task_id: int) -> Path:
    """Where the task's rubric suitability annotation is."""
    return task_dir(task_id) / SUITABILITY_FILENAME


def load_task(task_id: int) -> dict:
    """The task.json row for `task_id`."""
    path = task_dir(task_id) / TASK_FILENAME
    if not path.is_file():
        raise LocalStoreError(
            f"no task {task_id}: {display_path(path)} does not exist. Download "
            f"the benchmark first (scripts/download_dataset.py) or point "
            f"${repo_config.DATA_ROOT_ENV} at it."
        )
    return json.loads(path.read_text(encoding="utf-8"))


def task_ids_present() -> list[int]:
    """Every task id with a task.json under <data_root>/tasks/, ascending."""
    root = tasks_dir()
    if not root.is_dir():
        return []
    ids = []
    for d in root.iterdir():
        if d.is_dir() and d.name.startswith("task_id=") and (d / TASK_FILENAME).is_file():
            try:
                ids.append(int(d.name.split("=", 1)[1]))
            except ValueError:
                continue
    return sorted(ids)


# ---------------------------------------------------------------------------
# Attempts
# ---------------------------------------------------------------------------


def attempts_dir() -> Path:
    return output_root() / ATTEMPTS_DIRNAME


def gradings_dir() -> Path:
    return output_root() / GRADINGS_DIRNAME


def attempt_row_files() -> list[Path]:
    """Every `attempts/**/task_attempts.jsonl` under the output root."""
    root = attempts_dir()
    if not root.is_dir():
        return []
    return sorted(root.rglob(ATTEMPTS_FILENAME))


def iter_rows(path: Path):
    with Path(path).open(encoding="utf-8") as f:
        for lineno, line in enumerate(f, 1):
            line = line.strip()
            if not line:
                continue
            try:
                yield json.loads(line)
            except json.JSONDecodeError as e:
                raise LocalStoreError(
                    f"{display_path(path)}:{lineno} is not valid JSON: {e}"
                ) from e


def attempt_index() -> dict:
    """{attempt_id: (row, source path)} across every attempts JSONL.

    An id present twice is an error, not a silent winner: the two rows would
    grade to two different workbooks under one attempt_id.
    """
    index: dict = {}
    for path in attempt_row_files():
        for row in iter_rows(path):
            attempt_id = row.get("id")
            if attempt_id is None:
                raise LocalStoreError(f"{display_path(path)}: an attempt row has no 'id'")
            if attempt_id in index:
                first = display_path(index[attempt_id][1])
                raise LocalStoreError(
                    f"attempt id {attempt_id} appears in both {first} and "
                    f"{display_path(path)}. Ids are millisecond epochs and must "
                    f"be unique; remove or renumber one of the rows."
                )
            index[attempt_id] = (row, path)
    return index


def _file_refs(column, resolve) -> list[dict] | None:
    """A file column as [{name, path}] with absolute paths.

    Accepts a list of relative/absolute path strings (the pipelines' shape),
    a list of {"name", "path"} dicts, a single string, or a {name: path} dict.
    """
    if column is None:
        return None
    if isinstance(column, str):
        column = [column]
    if isinstance(column, dict):
        column = [{"name": k, "path": v} for k, v in column.items()]
    refs = []
    for item in column:
        stored = item.get("path", "") if isinstance(item, dict) else str(item)
        if not stored:
            continue
        name = (
            item.get("name")
            if isinstance(item, dict) and item.get("name")
            else Path(stored).name
        )
        refs.append({"name": name, "path": str(resolve(stored))})
    return refs


def workbook_path(row: dict) -> Path | None:
    """The attempt's judged workbook: the first Excel file in attempt_files."""
    for ref in _file_refs(row.get("attempt_files"), resolve_output_path) or []:
        if ref["name"].lower().endswith(EXCEL_EXTS):
            return Path(ref["path"])
    return None


def build_attempt(row: dict, task: dict) -> dict:
    """The attempt dict grade.py grades, from an attempt row and its task.json."""
    return {
        "attempt_id": row["id"],
        "task_id": row["task_id"],
        "prompt_files": _file_refs(row.get("prompt_files"), resolve_output_path),
        "attempt_files": _file_refs(row.get("attempt_files"), resolve_output_path),
        "agent_model_name": row.get("agent_model_name"),
        "agent_model_type": row.get("agent_model_type"),
        "prompt_version": row.get("prompt_version"),
        "agent_failed": row.get("agent_failed"),
        "task_name": task.get("task_name"),
        "task_starting_files": _file_refs(task.get("starting_files"), resolve_data_path),
        "task_solution_files": _file_refs(task.get("solution_files"), resolve_data_path),
    }


def select_attempts(attempt_ids=None, task_ids=None, agent_model_names=None) -> tuple[list[dict], dict]:
    """Attempts to grade, and a note on what was passed over.

    With `attempt_ids`, exactly those rows (unknown ids raise; a deprecated
    row or a missing workbook is reported and left for the grading to refuse).
    Otherwise every row that is not deprecated and whose workbook is on disk,
    narrowed by `task_ids` and `agent_model_names` when given, ordered by
    (task_id, id). Tasks are loaded once each.
    """
    index = attempt_index()
    tasks: dict[int, dict] = {}

    def _task(task_id):
        if task_id not in tasks:
            tasks[task_id] = load_task(task_id)
        return tasks[task_id]

    note = {"searched": [display_path(p) for p in attempt_row_files()]}
    if attempt_ids:
        missing = [i for i in attempt_ids if i not in index]
        if missing:
            raise LocalStoreError(
                f"no attempt row for id(s) {missing}. Searched "
                f"{', '.join(note['searched']) or '(no task_attempts.jsonl under ' + display_path(attempts_dir()) + ')'}."
            )
        rows = [index[i][0] for i in dict.fromkeys(attempt_ids)]
        for r in rows:
            if r.get("deprecated"):
                logger.warning(f"  attempt {r['id']} is marked deprecated in its row")
        return [build_attempt(r, _task(r["task_id"])) for r in rows], note

    rows = [r for r, _ in index.values()]
    note["total_rows"] = len(rows)
    if task_ids is not None:
        wanted = set(task_ids)
        rows = [r for r in rows if r.get("task_id") in wanted]
    if agent_model_names is not None:
        wanted_labels = set(agent_model_names)
        rows = [r for r in rows if r.get("agent_model_name") in wanted_labels]
    note["matched"] = len(rows)
    live = [r for r in rows if not r.get("deprecated")]
    note["deprecated"] = len(rows) - len(live)
    present = []
    for r in live:
        wb = workbook_path(r)
        if wb is not None and wb.exists():
            present.append(r)
    note["no_workbook"] = len(live) - len(present)
    present.sort(key=lambda r: (r.get("task_id"), r["id"]))
    return [build_attempt(r, _task(r["task_id"])) for r in present], note


# ---------------------------------------------------------------------------
# Gradings
# ---------------------------------------------------------------------------


def grading_rows() -> list[dict]:
    path = gradings_dir() / GRADINGS_FILENAME
    return list(iter_rows(path)) if path.is_file() else []


def graded_attempt_ids(grader_model: str | None = None) -> set:
    """Attempt ids that already have a non-deprecated grading row (under
    `grader_model` when given)."""
    out = set()
    for row in grading_rows():
        if row.get("deprecated"):
            continue
        if grader_model is not None and row.get("grader_model") != grader_model:
            continue
        out.add(row.get("attempt_id"))
    return out


_last_local_id = 0


def new_local_id() -> int:
    """A millisecond-epoch id, strictly increasing within this process.

    Two gradings finishing in the same millisecond would otherwise collide
    and overwrite each other's staged folder.
    """
    global _last_local_id
    local_id = max(int(time.time() * 1000), _last_local_id + 1)
    _last_local_id = local_id
    return local_id


def stage_grading_files(output_dir, grading_id: int) -> tuple[str, list[str]]:
    """Copy a grading's files into `<output_root>/gradings/<id>/`.

    Returns (path relative to <output_root>, sorted relative POSIX paths of
    the files) for the row's raw_files_path / raw_files.
    """
    dest = gradings_dir() / str(grading_id)
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        raise LocalStoreError(f"{display_path(dest)} already exists")
    shutil.copytree(str(output_dir), str(dest))
    raw_files = sorted(
        p.relative_to(dest).as_posix() for p in dest.rglob("*") if p.is_file()
    )
    return output_relative(dest), raw_files


def append_grading(row: dict) -> Path:
    """Append one gradings-shaped row to `<output_root>/gradings/gradings.jsonl`
    under a file lock, so parallel workers can share the file."""
    unknown = sorted(set(row) - set(GRADING_COLUMNS))
    missing = sorted(set(GRADING_COLUMNS) - set(row))
    if unknown or missing:
        raise LocalStoreError(
            f"grading row does not match the gradings columns "
            f"(unknown: {unknown}, missing: {missing})"
        )
    path = gradings_dir() / GRADINGS_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    line = json.dumps({c: row[c] for c in GRADING_COLUMNS}, default=str) + "\n"
    with path.open("a", encoding="utf-8") as f:
        fcntl.flock(f.fileno(), fcntl.LOCK_EX)
        try:
            f.write(line)
            f.flush()
        finally:
            fcntl.flock(f.fileno(), fcntl.LOCK_UN)
    return path


def now_iso() -> str:
    """Timestamp in the rows' format: ISO-8601 with an offset."""
    return datetime.now(timezone.utc).isoformat()
