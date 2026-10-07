"""Task source: the downloaded benchmark under <data_root>/tasks/.

    <data_root>/tasks/task_id=<N>/task.json         the task row
    <data_root>/tasks/task_id=<N>/starting_files/   what the agent receives

task.json (paths are POSIX, relative to <data_root>):

    {"task_id": 1, "task_name": "ApfelInc", "difficulty": "Medium-Hard", ...,
     "starting_files": ["tasks/task_id=1/starting_files/ApfelInc.xlsx"],
     "solution_files": ["tasks/task_id=1/solution_files/ApfelInc - Solution.xlsx"]}

LocalSource stages the starting files, in the order the row lists them, into
a staging dir the workspace is then seeded from. No database, object store
or credentials involved.
"""
import json
from dataclasses import dataclass
from pathlib import Path

from .repo_config import data_path


@dataclass
class TaskSpec:
    task_id: int
    task_name: str
    starting_files: list  # list[Path] once staged locally


def task_dir(data_root: Path, task_id: int) -> Path:
    return Path(data_root) / "tasks" / f"task_id={task_id}"


def load_task_row(data_root: Path, task_id: int) -> dict:
    """The task.json row for one task, verbatim."""
    path = task_dir(data_root, task_id) / "task.json"
    if not path.exists():
        raise LookupError(
            f"No task with id {task_id} ({path} does not exist). The benchmark "
            f"lives under local.data_root (scripts/download_dataset.py)."
        )
    return json.loads(path.read_text())


def local_task_ids(data_root: Path, first: int | None = None, last: int | None = None) -> list[int]:
    """Every task id present under <data_root>/tasks/, ascending, optionally
    bounded to [first, last]."""
    root = Path(data_root) / "tasks"
    if not root.is_dir():
        raise LookupError(
            f"No benchmark at {root}. Download it with scripts/download_dataset.py, "
            f"or point local.data_root in config/config.yaml (or MBABENCH_DATA_ROOT) at it."
        )
    ids = []
    for p in root.glob("task_id=*/task.json"):
        tid = int(json.loads(p.read_text())["task_id"])
        if (first is None or tid >= first) and (last is None or tid <= last):
            ids.append(tid)
    return sorted(ids)


class LocalSource:
    """Reads the task row and stages its starting files."""

    def __init__(self, task_id: int, data_root: Path):
        self.task_id = task_id
        self.data_root = Path(data_root)

    def fetch(self, staging_dir: Path) -> TaskSpec:
        row = load_task_row(self.data_root, self.task_id)
        starting = row.get("starting_files") or []
        if not starting:
            raise ValueError(f"Task {self.task_id} has no starting_files")

        staging_dir.mkdir(parents=True, exist_ok=True)
        local_files = []
        for rel in starting:  # the order the row records
            src = data_path(str(rel))
            if not src.is_file():
                raise FileNotFoundError(
                    f"Task {self.task_id} lists {rel}, which is missing at {src}; "
                    f"re-run scripts/download_dataset.py"
                )
            dest = staging_dir / Path(str(rel)).name
            dest.write_bytes(src.read_bytes())
            if dest.stat().st_size == 0:
                raise IOError(f"Starting file is empty: {src}")
            local_files.append(dest)
        return TaskSpec(
            task_id=self.task_id,
            task_name=str(row.get("task_name") or f"task_{self.task_id}"),
            starting_files=local_files,
        )
