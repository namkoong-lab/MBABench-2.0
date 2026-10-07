"""One `task_attempts` row from one AttemptResult.

The row shape is shared by every MBABench pipeline (the judge reads it):

    id, task_id, agent_model_name, agent_model_type, attempt_files,
    prompt_files, start_time, end_time, time_taken_min, cost, prompt_version,
    agent_failed, agent_failed_reason, deprecated, created_at,
    context_reduced, deprecated_reason, updated_at, extra_configs

Nothing here touches the filesystem; the sink fills `id` / `created_at` and
serialises the datetimes.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Protocol

from ..base import AttemptResult

TASK_ATTEMPTS_COLUMNS: tuple[str, ...] = (
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


class AttemptRowConfig(Protocol):
    """What a sink must know to label a row; the sink passes `self`."""

    agent_model_name: str
    agent_model_type: str
    prompt_version: int | str | None
    extra_configs: dict[str, Any]


def task_metadata(result: AttemptResult) -> dict:
    extra = result.extra or {}
    meta = extra.get("task_metadata")
    return meta if isinstance(meta, dict) else {}


def db_task_id(result: AttemptResult) -> int:
    """task_attempts.task_id (an int) for this result: the source's
    metadata['db_task_id'], else a numeric result.task_id."""
    value = task_metadata(result).get("db_task_id")
    if value is not None:
        return int(value)
    try:
        return int(result.task_id)
    except (TypeError, ValueError) as e:
        raise ValueError(
            f"task_attempts sink: task_id must resolve to an int, got "
            f"{result.task_id!r}. Ensure the source populates "
            f"spec.metadata['db_task_id'] or yields numeric task_ids."
        ) from e


def extra_configs_payload(
    result: AttemptResult, base: dict[str, Any] | None
) -> dict[str, Any] | None:
    """task_attempts.extra_configs: the sink's identity settings merged with
    the runner's per-attempt stamps (result.extra['extra_configs']); None
    when there is nothing to record."""
    payload = dict(base or {})
    per_attempt = (result.extra or {}).get("extra_configs")
    if isinstance(per_attempt, dict):
        payload |= per_attempt
    return payload or None


def attempt_row(
    result: AttemptResult,
    cfg: AttemptRowConfig,
    *,
    attempt_files: list[str],
    prompt_files: list[str],
) -> dict[str, Any]:
    """The task_attempts row for `result`, keyed by column in table order.

    `attempt_files` / `prompt_files` are the stored locations (POSIX paths
    relative to <output_root>) in storage order: solution workbook first,
    then the logs; the prompts JSON separately. `cost` is always None (GUI
    runs are subscription-based) and `deprecated` always False. `id` and
    `created_at` are None here; the sink fills them in.
    """
    start_dt = datetime.fromisoformat(result.started_at)
    end_dt = datetime.fromisoformat(result.finished_at)
    time_taken_min = (result.duration_seconds or 0.0) / 60.0

    agent_failed = result.status != "success"
    agent_failed_reason: str | None = None
    if agent_failed:
        extra = result.extra or {}
        agent_failed_reason = (
            extra.get("error") or extra.get("failure_reason") or result.status
        )

    row = {
        "id": None,
        "task_id": db_task_id(result),
        "agent_model_name": cfg.agent_model_name,
        "agent_model_type": cfg.agent_model_type,
        "attempt_files": list(attempt_files),
        "prompt_files": list(prompt_files),
        "start_time": start_dt,
        "end_time": end_dt,
        "time_taken_min": time_taken_min,
        "cost": None,
        "prompt_version": cfg.prompt_version,
        "agent_failed": agent_failed,
        "agent_failed_reason": agent_failed_reason,
        "deprecated": False,
        "created_at": None,
        "context_reduced": None,
        "deprecated_reason": None,
        "updated_at": None,
        "extra_configs": extra_configs_payload(result, cfg.extra_configs),
    }
    assert tuple(row) == TASK_ATTEMPTS_COLUMNS
    return row
