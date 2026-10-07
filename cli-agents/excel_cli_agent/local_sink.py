"""The attempt sink: <output_root>/attempts/ on local disk.

The batch runner copies one attempt's file set under

    <output_root>/attempts/<agent_model_name>/task_id=<N>/<YYYYmmdd_HHMMSS>/

and appends one JSON object per attempt to

    <output_root>/attempts/<agent_model_name>/task_attempts.jsonl

carrying exactly the columns below, in this order, with the two file columns
holding POSIX paths relative to <output_root>. The judge
(judge/main_scripts/grade.py) reads these rows; the first Excel file in
`attempt_files` is the workbook it grades.

IDS. A row's `id` is the millisecond epoch at which it was written: unique
across lanes running concurrently on one machine, and monotonic.

RESUME. The same JSONL is what skip-if-attempted reads, so a relaunched lane
picks up where it stopped.
"""

import fcntl
import json
import os
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .repo_config import to_output_relative

# The attempt-row columns, in order. A row has all of them and nothing else:
# the four pipelines write the same shape and the judge reads it.
ATTEMPT_COLUMNS: Tuple[str, ...] = (
    "id",
    "task_id",
    "agent_model_name",
    "agent_model_type",
    "attempt_files",
    "prompt_files",
    "start_time",
    "end_time",
    "time_taken_min",
    "cost",
    "prompt_version",
    "agent_failed",
    "agent_failed_reason",
    "deprecated",
    "created_at",
    "context_reduced",
    "deprecated_reason",
    "updated_at",
    "extra_configs",
)

ATTEMPTS_SUBDIR = "attempts"
ATTEMPTS_JSONL = "task_attempts.jsonl"


def iso(dt: Optional[datetime]) -> Optional[str]:
    """ISO-8601 with an offset. The runner builds its times with
    datetime.fromtimestamp(), which is naive local time; astimezone() on a
    naive value attaches this machine's offset."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.astimezone()
    return dt.isoformat()


class LocalAttemptSink:
    """Writes one cohort's attempts under <output_root>/attempts/<agent_model_name>/."""

    def __init__(self, output_root: Path, agent_model_name: str):
        self.output_root = Path(output_root)
        self.agent_model_name = agent_model_name

    # -- locations -------------------------------------------------------

    @property
    def cohort_dir(self) -> Path:
        """attempts/<label> — a label with a slash nests one level, by design."""
        return self.output_root.joinpath(ATTEMPTS_SUBDIR, *self.agent_model_name.split("/"))

    @property
    def jsonl_path(self) -> Path:
        return self.cohort_dir / ATTEMPTS_JSONL

    def attempt_dir(self, task_id: int, timestamp: str) -> Path:
        return self.cohort_dir / f"task_id={task_id}" / timestamp

    # -- writing ---------------------------------------------------------

    @staticmethod
    def new_attempt_id() -> int:
        """Millisecond epoch — see the module docstring."""
        return int(time.time() * 1000)

    def store_files(self, files: Iterable[Tuple[Path, str]], task_id: int,
                    timestamp: str) -> List[str]:
        """Copy one attempt's artifacts in; return their output-relative paths.

        `files` is a (source, relative destination) list in the order the
        row should record — solution.xlsx first, which is what makes
        attempt_files[0] the judged workbook.
        """
        dest_root = self.attempt_dir(task_id, timestamp)
        stored: List[str] = []
        for src, rel in files:
            dest = dest_root / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            stored.append(to_output_relative(dest))
        return stored

    def append_row(self, row: Dict[str, Any]) -> Dict[str, Any]:
        """Append one attempt row, filling the columns it did not set.

        The write is a single line under an exclusive lock on the file, so
        concurrent lanes appending to one cohort interleave whole rows
        rather than half-lines.
        """
        missing = set(row) - set(ATTEMPT_COLUMNS)
        if missing:
            raise ValueError(
                f"{sorted(missing)} are not attempt-row columns; a row must "
                "carry the shared column set exactly"
            )
        complete = {col: row.get(col) for col in ATTEMPT_COLUMNS}
        self.jsonl_path.parent.mkdir(parents=True, exist_ok=True)
        line = json.dumps(complete, ensure_ascii=False, default=str) + "\n"
        with open(self.jsonl_path, "a", encoding="utf-8") as fh:
            fcntl.flock(fh, fcntl.LOCK_EX)
            try:
                fh.write(line)
                fh.flush()
                os.fsync(fh.fileno())
            finally:
                fcntl.flock(fh, fcntl.LOCK_UN)
        return complete

    # -- reading (skip-if-attempted / resume) -----------------------------

    def read_rows(self) -> List[Dict[str, Any]]:
        """Every row this cohort has recorded, oldest first.

        A truncated final line is skipped rather than fatal: a lane killed
        mid-write must not stop the relaunch that is meant to recover it.
        """
        path = self.jsonl_path
        if not path.is_file():
            return []
        rows: List[Dict[str, Any]] = []
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue
        return rows

    def has_attempt(self, task_id: int, prompt_version: Optional[int] = None) -> bool:
        """True if a non-failed, non-deprecated row exists for this task (and
        prompt version) — the skip-if-attempted rule a relaunch applies."""
        for row in self.read_rows():
            if int(row.get("task_id") or -1) != int(task_id):
                continue
            if row.get("deprecated") or row.get("agent_failed"):
                continue
            if prompt_version is not None and row.get("prompt_version") != prompt_version:
                continue
            return True
        return False
