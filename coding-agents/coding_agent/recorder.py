"""Recorder: persist a validated attempt under <output_root>/attempts/.

  <output_root>/attempts/<agent_model_name>/task_id=<N>/<ts>/   the artifacts,
       solution.xlsx FIRST in attempt_files (the judge grades the first xlsx),
       then PROMPT.md, transcript, telemetry, verdict, trajectory, run config,
       and the prompt snapshot (system prompt, template, attachments, extras)
  <output_root>/attempts/<agent_model_name>/task_attempts.jsonl   one row per
       attempt; extra_configs records the settings the attempt ran under
       (RunConfig.extra_configs(): identity, sandbox image, harness defaults,
       house_standards {version, file, sha256}, prompt_extras).

Verdicts infra_failure / needs_review write NO row (held locally in the
attempt dir, so a relaunch picks the task up again); success / timeout /
agent_failure write a row. See local_sink.py for the row shape.
"""
import json
from datetime import datetime

from .config import RunConfig, template_attachments
from .local_sink import write_attempt
from .prompt_builder import prompt_extra_paths, prompt_extras_provenance, prompt_file_paths
from .sandbox import SandboxResult
from .task_source import TaskSpec
from .validate import Verdict
from .workspace import Attempt

RECORDABLE = {"success", "timeout", "agent_failure"}


def _extras_stamp(cfg: RunConfig) -> dict:
    """{"prompt_extras": {...}} when the template stages extra files, else {}."""
    extras = prompt_extras_provenance(cfg)
    return {"prompt_extras": extras} if extras else {}


def record(cfg: RunConfig, spec: TaskSpec, attempt: Attempt, sandbox: SandboxResult,
           verdict: Verdict, telemetry: dict, prompt_version: int) -> dict:
    summary = {
        "agent_model_name": cfg.agent_model_name,
        "task_id": spec.task_id,
        "task_name": spec.task_name,
        "status": verdict.status,
        "reason": verdict.reason,
        "duration_min": round(sandbox.duration_seconds / 60.0, 2),
        "prompt_version": prompt_version,
        "cost_usd": telemetry.get("cost_usd"),
        "tokens": telemetry.get("totals"),
        "extra_configs": {**cfg.extra_configs(), **_extras_stamp(cfg)},
        "recorded": False,
    }

    if verdict.status not in RECORDABLE:
        summary["note"] = "not recorded (infra_failure/needs_review are held in the attempt dir)"
        (attempt.attempt_dir / "summary.json").write_text(json.dumps(summary, indent=2))
        return summary

    # The prompt snapshot: system prompt, the template actually used, its
    # attachments, its workspace extras.
    prompt_paths = [*prompt_file_paths(cfg), *template_attachments(cfg),
                    *[src for src, _ in prompt_extra_paths(cfg)]]
    agent_failed = verdict.status != "success"

    row, dest = write_attempt(
        output_root=cfg.output_root,
        agent_model_name=cfg.agent_model_name,
        task_id=spec.task_id,
        started_at=attempt.started_at,
        ended_at=datetime.now(),
        solution_path=verdict.solution_path,
        workspace=attempt.workspace,
        attempt_dir=attempt.attempt_dir,
        prompt_paths=prompt_paths,
        time_taken_min=round(sandbox.duration_seconds / 60.0, 2),
        cost=telemetry.get("cost_usd"),
        prompt_version=prompt_version,
        agent_failed=agent_failed,
        agent_failed_reason=None if not agent_failed else verdict.reason,
        extra_configs={**cfg.extra_configs(), **_extras_stamp(cfg)},
    )
    summary["attempt_id"] = row["id"]
    summary["recorded"] = True
    summary["output_dir"] = str(dest)
    summary["attempt_files"] = row["attempt_files"]
    (attempt.attempt_dir / "summary.json").write_text(json.dumps(summary, indent=2))
    return summary
