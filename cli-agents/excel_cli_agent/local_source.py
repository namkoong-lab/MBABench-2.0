"""The task source: the downloaded benchmark under <data_root>/tasks/.

Each task is one folder, `tasks/task_id=<N>/`, holding `task.json` and the
files it names (paths inside it are POSIX, relative to <data_root>):

    {
      "task_id": 1,
      "task_name": "ApfelInc",
      "difficulty": "Medium-Hard",
      "description": "...",
      "starting_files": ["tasks/task_id=1/starting_files/ApfelInc.xlsx"],
      "solution_files": ["tasks/task_id=1/solution_files/ApfelInc - Solution.xlsx"]
    }

Only `starting_files` reach the agent; `solution_files` are the judge's.
"""

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

from .repo_config import resolve_data_path

# The per-task folder name; also the glob that enumerates them.
TASK_DIR_GLOB = "task_id=*"
TASK_JSON = "task.json"


@dataclass
class LocalTask:
    """One task.json row. The two file lists hold absolute local paths;
    everything else in the row stays in `row` so nothing is lost."""

    task_id: int
    task_name: str
    starting_files: List[str] = field(default_factory=list)
    solution_files: List[str] = field(default_factory=list)
    row: Dict[str, Any] = field(default_factory=dict)


class BundleNotAvailable(RuntimeError):
    """The benchmark is not where the config says it is."""


class LocalTaskSource:
    """Read-only view of <data_root>/tasks/."""

    def __init__(self, data_root: Path):
        self.data_root = Path(data_root)
        self._tasks: Optional[List[LocalTask]] = None

    @property
    def tasks_dir(self) -> Path:
        return self.data_root / "tasks"

    def load_all(self) -> List[LocalTask]:
        """Every task present, ordered by id.

        Ordered explicitly rather than by directory-listing order: `sorted()`
        on a filesystem glob is lexicographic (task_id=10 before task_id=2).
        """
        if self._tasks is not None:
            return self._tasks
        if not self.tasks_dir.is_dir():
            raise BundleNotAvailable(
                f"No benchmark at {self.tasks_dir}. Download it with "
                "scripts/download_dataset.py, or point local.data_root in "
                "config/config.yaml (or MBABENCH_DATA_ROOT) at it."
            )
        tasks: List[LocalTask] = []
        for task_dir in self.tasks_dir.glob(TASK_DIR_GLOB):
            manifest = task_dir / TASK_JSON
            if manifest.is_file():
                tasks.append(self._load_task(manifest))
        if not tasks:
            raise BundleNotAvailable(
                f"{self.tasks_dir} holds no {TASK_DIR_GLOB}/{TASK_JSON}; the "
                "benchmark is empty or only partly downloaded."
            )
        self._tasks = sorted(tasks, key=lambda t: t.task_id)
        return self._tasks

    def _load_task(self, manifest: Path) -> LocalTask:
        row = json.loads(manifest.read_text(encoding="utf-8"))
        return LocalTask(
            task_id=int(row["task_id"]),
            task_name=str(row.get("task_name") or f"task_{row['task_id']}"),
            starting_files=self._resolve_files(row, "starting_files", manifest, required=True),
            solution_files=self._resolve_files(row, "solution_files", manifest, required=False),
            row=row,
        )

    @staticmethod
    def _resolve_files(row: Dict[str, Any], column: str, manifest: Path,
                       required: bool) -> List[str]:
        """Absolute local paths for one file list of a task row.

        A starting file that is not on disk is fatal here rather than at
        workspace setup: an attempt run without its starting workbook is the
        empty-context defect, and the download is the only place it can
        come from.
        """
        out: List[str] = []
        for rel in (row.get(column) or []):
            path = resolve_data_path(rel)
            if not path.is_file() or path.stat().st_size == 0:
                if not required:
                    continue
                raise BundleNotAvailable(
                    f"{manifest.parent.name}: {column} entry {rel} is missing "
                    f"or empty at {path}. Re-run scripts/download_dataset.py."
                )
            out.append(str(path))
        return out

    def get(self, task_id: int) -> Optional[LocalTask]:
        """The task with this id, or None."""
        for task in self.load_all():
            if task.task_id == int(task_id):
                return task
        return None

    def find_by_name(self, name: str) -> Optional[LocalTask]:
        """Name lookup, exact first, then with spaces normalised to underscores."""
        normalized = name.replace(" ", "_")
        for task in self.load_all():
            if task.task_name == name:
                return task
        for task in self.load_all():
            if task.task_name.replace(" ", "_") == normalized:
                return task
        return None
