"""LocalAttemptSink — the record of an attempt under <output_root>.

    <output_root>/attempts/<agent_model_name>/task_id=<N>/<YYYYmmdd_HHMMSS>/<file>
        the solution workbook FIRST (the file the judge grades), then every
        log the attempt produced, then the prompts JSON
    <output_root>/attempts/<agent_model_name>/task_attempts.jsonl
        one task_attempts-shaped row per attempt (task_io/sinks/attempt_row.py)

`id` is the millisecond epoch time of the write; file paths are POSIX paths
relative to <output_root>; timestamps are ISO-8601 with an offset. Rows are
appended under an exclusive fcntl lock so parallel lanes can share one file.
Every file is copied, so the runner may delete its staging directory
afterwards (retains_files = True). skip_already_attempted in the bundle
source reads the same file this appends to.
"""

from __future__ import annotations

import json
import os
import logging
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from ..base import AttemptResult
from ..local_layout import (
    TASK_DIR_PREFIX,
    attempts_file,
    cohort_dir,
    output_relative,
)
from .attempt_row import TASK_ATTEMPTS_COLUMNS, attempt_row, db_task_id

logger = logging.getLogger(__name__)

_TIMESTAMP_FMT = "%Y%m%d_%H%M%S"


def _with_offset(dt: datetime) -> datetime:
    """A naive datetime (the runner records local wall-clock time) gets the
    local offset attached; an aware one stands."""
    return dt.astimezone() if dt.tzinfo is None else dt


def _json_default(o: Any):
    if isinstance(o, datetime):
        return _with_offset(o).isoformat()
    if isinstance(o, Path):
        return o.as_posix()
    raise TypeError(f"Object of type {o.__class__.__name__} is not JSON serializable")


def append_locked(path: Path, line: str) -> None:
    """Append one line under an exclusive advisory lock (released on close)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a") as f:
        try:
            import fcntl

            fcntl.flock(f.fileno(), fcntl.LOCK_EX)
        except (ImportError, OSError):
            pass  # non-POSIX: no lock semantics available
        f.write(line + "\n")
        f.flush()


class LocalAttemptSink:
    # Every file handed to publish() is copied under output_root before the
    # row is written, so the caller's staging copies are disposable.
    retains_files = True

    def __init__(
        self,
        *,
        output_root: str | Path,
        agent_model_name: str,
        agent_model_type: str = "excel",
        prompt_version: int | str | None,
        extra_configs: dict[str, Any] | None = None,
    ):
        if not agent_model_name:
            raise ValueError(
                "local sink: agent_model_name is required. It is derived from "
                "resolve_agent_identity(cfg) — an empty value means the "
                "resolver returned an invalid AgentIdentity."
            )
        self.output_root = Path(output_root)
        self.agent_model_name = agent_model_name
        self.agent_model_type = agent_model_type
        self.prompt_version = prompt_version
        self.extra_configs = dict(extra_configs or {})
        self.cohort_dir = cohort_dir(self.output_root, agent_model_name)
        self.log_path = attempts_file(self.output_root, agent_model_name)
        self._last_id = 0
        # Nothing is created here: a dry run must leave outputs/ untouched.
        # The cohort folder and the attempts log appear on the first write.
        probe = self.output_root
        while not probe.exists() and probe.parent != probe:
            probe = probe.parent
        try:
            if not os.access(probe, os.W_OK):
                raise PermissionError(f"{probe} is not writable")
        except OSError as e:
            raise ValueError(
                f"local sink: cannot write under {self.cohort_dir} "
                f"({type(e).__name__}: {e}). Set MBABENCH_OUTPUT_ROOT or "
                f"local.output_root in <repo>/config/config.yaml."
            ) from e
        logger.info(
            f"Sink: local -> {self.cohort_dir}/ (rows appended to {self.log_path.name})"
        )

    # --- layout ------------------------------------------------------------

    def attempt_dir(self, task_id: int | str, timestamp: str) -> Path:
        return self.cohort_dir / f"{TASK_DIR_PREFIX}{task_id}" / timestamp

    def describe_destination(self, task_id: int | str) -> str:
        """One line for --dry-run: where this task's attempt would land."""
        folder = self.attempt_dir(task_id, "<YYYYmmdd_HHMMSS>")
        return (
            f"{folder}/ (solution workbook, logs, prompts JSON) + one row "
            f"appended to {self.log_path}"
        )

    # --- internals ---------------------------------------------------------

    @staticmethod
    def _existing(paths, what: str) -> list[Path]:
        out: list[Path] = []
        for p in paths:
            if p is None:
                continue
            p = Path(p)
            if not p.exists():
                logger.warning(f"Sink: skipping missing {what} {p}")
                continue
            out.append(p)
        return out

    def _attempt_files(self, result: AttemptResult) -> list[Path]:
        # Solution workbook FIRST (the judged file), then every log.
        return self._existing((result.solution_file, *result.log_files), "file")

    def _prompt_files(self, result: AttemptResult) -> list[Path]:
        return self._existing(result.prompt_files or [], "prompt file")

    def _copy_all(self, files: list[Path], dest_dir: Path) -> list[str]:
        recorded: list[str] = []
        for local in files:
            dest = dest_dir / local.name
            shutil.copy2(local, dest)
            recorded.append(output_relative(self.output_root, dest))
        return recorded

    def _make_attempt_dir(self, task_id: int, timestamp: str) -> Path:
        """A fresh directory for this attempt; a second publish of the same
        task in the same second (another lane) gets a numbered sibling."""
        suffix = 1
        while True:
            name = timestamp if suffix == 1 else f"{timestamp}_{suffix}"
            dest_dir = self.attempt_dir(task_id, name)
            try:
                dest_dir.mkdir(parents=True, exist_ok=False)
                return dest_dir
            except FileExistsError:
                suffix += 1

    def _next_id(self) -> int:
        row_id = int(time.time() * 1000)
        if row_id <= self._last_id:  # two publishes inside one millisecond
            row_id = self._last_id + 1
        self._last_id = row_id
        return row_id

    # --- public API --------------------------------------------------------

    def publish(self, result: AttemptResult) -> None:
        task_id = db_task_id(result)
        now = datetime.now()
        dest_dir = self._make_attempt_dir(task_id, now.strftime(_TIMESTAMP_FMT))

        attempt_paths = self._copy_all(self._attempt_files(result), dest_dir)
        prompt_paths = self._copy_all(self._prompt_files(result), dest_dir)

        row = attempt_row(
            result, self, attempt_files=attempt_paths, prompt_files=prompt_paths
        )
        row["id"] = self._next_id()
        row["created_at"] = _with_offset(now)
        assert tuple(row) == TASK_ATTEMPTS_COLUMNS

        append_locked(self.log_path, json.dumps(row, default=_json_default))
        logger.info(
            f"Sink: recorded attempt id={row['id']} for task_id={task_id} "
            f"status={result.status} attempt_files={len(attempt_paths)} "
            f"prompt_files={len(prompt_paths)} -> {dest_dir}/"
        )

    def close(self) -> None:
        return None
