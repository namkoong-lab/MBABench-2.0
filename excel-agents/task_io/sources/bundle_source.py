"""BundleTaskSource — the benchmark tasks under <data_root>/tasks/.

Reads <data_root>/tasks/task_id=<N>/task.json (placed there by
scripts/download_dataset.py) and yields one TaskSpec per task, ascending id:

    {
      "task_id": 1,
      "task_name": "ApfelInc",
      "difficulty": "Medium-Hard", "difficulty_score": 3.5,
      "model_type": "Scenario / Sensitivity Analysis",
      "description": "...",
      "estimated_solve_time_hours": 2.63, "solve_time_range": "2-3h",
      "starting_files": ["tasks/task_id=1/starting_files/ApfelInc.xlsx"],
      "solution_files": ["tasks/task_id=1/solution_files/ApfelInc - Solution.xlsx"]
    }

upload_files are the starting files, read in place (nothing is copied).
Filters, applied in this order:

    task_ids               -> only these ids (default: every task present)
    skip_already_attempted -> drop tasks that already have a row in
                              <output_root>/attempts/<agent_model_name>/task_attempts.jsonl
                              for this agent_model_name and prompt_version with
                              agent_failed false and deprecated false — the
                              file the local sink appends to, so a re-run of
                              the same config resumes where it stopped.

No database, no object store, no network.
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Iterator

from ..base import TaskSpec
from ..local_layout import (
    TASK_DIR_PREFIX,
    TASK_ROW_FILE,
    TASKS_DIR,
    attempts_file,
    data_path,
)

logger = logging.getLogger(__name__)

_TASK_DIR_RE = re.compile(rf"^{re.escape(TASK_DIR_PREFIX)}(\d+)$")

# task.json field names (fixed by the dataset).
ID_FIELD = "task_id"
NAME_FIELD = "task_name"
FILES_FIELD = "starting_files"


def _same_version(a: Any, b: Any) -> bool:
    """prompt_version equality across the int the config carries and
    whatever a JSON row recorded (int, or a numeric string)."""
    if a is None or b is None:
        return a is b
    try:
        return int(a) == int(b)
    except (TypeError, ValueError):
        return str(a) == str(b)


def load_attempted_task_ids(
    attempts_log: Path, agent_model_name: str, prompt_version: Any
) -> set[int]:
    """Task ids with a non-failed, non-deprecated row for this
    (agent_model_name, prompt_version) in the local attempts log."""
    if not attempts_log.is_file():
        return set()
    done: set[int] = set()
    with open(attempts_log) as f:
        for lineno, line in enumerate(f, 1):
            line = line.strip()
            if not line:
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                logger.warning(f"{attempts_log}:{lineno}: not JSON; ignored")
                continue
            if row.get("agent_model_name") != agent_model_name:
                continue
            if not _same_version(row.get("prompt_version"), prompt_version):
                continue
            if row.get("agent_failed") or row.get("deprecated"):
                continue
            try:
                done.add(int(row["task_id"]))
            except (KeyError, TypeError, ValueError):
                continue
    return done


class BundleTaskSource:
    def __init__(
        self,
        *,
        data_root: str | Path,
        output_root: str | Path,
        agent_model_name: str,
        prompt_version: int | str | None,
        task_ids: list[int] | None = None,
        skip_already_attempted: bool = True,
    ):
        self.data_root = Path(data_root)
        self.output_root = Path(output_root)
        self.tasks_dir = self.data_root / TASKS_DIR
        if not self.tasks_dir.is_dir():
            raise ValueError(
                f"bundle task source: {self.tasks_dir} is not a directory. "
                f"Download the benchmark first (scripts/download_dataset.py) or "
                f"point MBABENCH_DATA_ROOT / local.data_root in "
                f"<repo>/config/config.yaml at where it lives."
            )
        self.agent_model_name = agent_model_name
        self.prompt_version = prompt_version
        self.task_ids = [int(t) for t in (task_ids or [])]
        self.skip_already_attempted = skip_already_attempted
        self.attempts_log = attempts_file(self.output_root, agent_model_name)

    # --- rows --------------------------------------------------------------

    def _task_dirs(self) -> list[tuple[int, Path]]:
        found: list[tuple[int, Path]] = []
        for entry in self.tasks_dir.iterdir():
            m = _TASK_DIR_RE.match(entry.name)
            if m and (entry / TASK_ROW_FILE).is_file():
                found.append((int(m.group(1)), entry / TASK_ROW_FILE))
        return sorted(found)  # ascending task_id

    @staticmethod
    def _load_row(path: Path) -> dict[str, Any]:
        with open(path) as f:
            row = json.load(f)
        if not isinstance(row, dict) or ID_FIELD not in row or NAME_FIELD not in row:
            raise ValueError(
                f"bundle task source: {path} is not a task.json "
                f"(needs {ID_FIELD!r} and {NAME_FIELD!r})"
            )
        return row

    def _select_rows(self) -> list[dict[str, Any]]:
        wanted = set(self.task_ids)
        done = (
            load_attempted_task_ids(
                self.attempts_log, self.agent_model_name, self.prompt_version
            )
            if self.skip_already_attempted
            else set()
        )
        rows: list[dict[str, Any]] = []
        for tid, path in self._task_dirs():
            if wanted and tid not in wanted:
                continue
            if tid in done:
                continue
            row = self._load_row(path)
            if int(row[ID_FIELD]) != tid:
                raise ValueError(
                    f"bundle task source: {path} says task_id={row[ID_FIELD]!r} "
                    f"but sits in a task_id={tid} folder"
                )
            rows.append(row)
        if done:
            logger.info(
                f"skip_already_attempted: {len(done)} task(s) already have a "
                f"successful row in {self.attempts_log}"
            )
        return rows

    # --- spec assembly -----------------------------------------------------

    def _starting_files(self, row: dict[str, Any]) -> list[Path]:
        resolved: list[Path] = []
        for rel in row.get(FILES_FIELD) or []:
            p = data_path(self.data_root, str(rel))
            if not p.is_file():
                raise FileNotFoundError(
                    f"bundle task source: task_id={row[ID_FIELD]} names starting "
                    f"file {rel!r} but {p} is missing — the download under "
                    f"{self.data_root} is incomplete."
                )
            resolved.append(p)
        return resolved

    @staticmethod
    def _metadata_for(row: dict[str, Any]) -> dict[str, Any]:
        return {
            "source_kind": "bundle",
            "db_task_id": int(row[ID_FIELD]),
            "overrides": {},
            "difficulty": row.get("difficulty"),
            "model_type": row.get("model_type"),
        }

    # --- public API --------------------------------------------------------

    def iter_tasks(self) -> Iterator[TaskSpec]:
        rows = self._select_rows()
        logger.info(
            f"{type(self).__name__} matched {len(rows)} task(s) "
            f"(ids={self.task_ids or 'any'}) under {self.tasks_dir}"
        )
        for row in rows:
            tid = int(row[ID_FIELD])
            if not (row.get(FILES_FIELD) or []):
                logger.warning(
                    f"Task task_id={tid} name={row[NAME_FIELD]!r} has no "
                    f"{FILES_FIELD}; skipping."
                )
                continue
            yield TaskSpec(
                task_id=str(tid),
                task_name=str(row[NAME_FIELD]),
                upload_files=self._starting_files(row),
                solution_name=None,
                metadata=self._metadata_for(row),
            )

    def close(self) -> None:
        return None
