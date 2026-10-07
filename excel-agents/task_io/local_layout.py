"""Path conventions of the downloaded benchmark (<data_root>) and of runs
(<output_root>). Pure path arithmetic — nothing here reads config.

    <data_root>/tasks/task_id=<N>/task.json
    <data_root>/tasks/task_id=<N>/starting_files/<name>
    <output_root>/attempts/<agent_model_name>/task_attempts.jsonl
    <output_root>/attempts/<agent_model_name>/task_id=<N>/<YYYYmmdd_HHMMSS>/<file>

Paths recorded inside the data are POSIX and relative to <data_root>; paths
recorded in an attempt row are POSIX and relative to <output_root>. A label
containing `/` (e.g. `openpyxl_anthropic/claude-x`) nests one directory level.
"""

from __future__ import annotations

from pathlib import Path, PurePosixPath

TASKS_DIR = "tasks"
TASK_DIR_PREFIX = "task_id="
TASK_ROW_FILE = "task.json"
ATTEMPTS_DIR = "attempts"
ATTEMPTS_FILE = "task_attempts.jsonl"


def data_path(data_root: Path, rel: str) -> Path:
    """On-disk location of a path recorded relative to <data_root>."""
    p = PurePosixPath(str(rel).replace("\\", "/"))
    return Path(rel) if p.is_absolute() else Path(data_root).joinpath(*p.parts)


def cohort_dir(output_root: Path, agent_model_name: str) -> Path:
    """<output_root>/attempts/<agent_model_name>/"""
    return Path(output_root).joinpath(ATTEMPTS_DIR, *agent_model_name.split("/"))


def attempts_file(output_root: Path, agent_model_name: str) -> Path:
    """<output_root>/attempts/<agent_model_name>/task_attempts.jsonl"""
    return cohort_dir(output_root, agent_model_name) / ATTEMPTS_FILE


def output_relative(output_root: Path, path: Path) -> str:
    """POSIX path of `path` relative to <output_root>; absolute when outside it."""
    p = Path(path).resolve()
    try:
        return p.relative_to(Path(output_root).resolve()).as_posix()
    except ValueError:
        return p.as_posix()
