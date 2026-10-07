"""Single-task runner — the core of the coding-agent pipeline.

One invocation = one attempt of one benchmark task:

    python -m coding_agent.run_task --config run_configs/claude_code_fable_5_1_max.yaml --task-id 11

The task comes from <data_root>/tasks/task_id=<N>/ and the attempt is
recorded under <output_root>/attempts/<agent_model_name>/ (repo_config). A
range of tasks is run_sweep.py, which calls this once per task.

Exit codes: 0 success | 2 agent_failure | 3 timeout | 4 infra_failure |
            5 needs_review (orchestrators branch on these).
"""
import argparse
import gzip
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from .agents import agent_env, build_command
from .config import load_config, load_dotenv_if_present, resolve_secrets, template_attachments
from .local_sink import cohort_dir
from .prompt_builder import build_prompt, prompt_extra_paths, prompt_file_paths
from .recorder import record
from .repo_config import describe_local_target
from .sandbox import run_in_sandbox
from .task_source import LocalSource
from .telemetry import parse_transcript, write_telemetry
from .validate import validate, write_verdict
from .workspace import create_attempt, seed_template_attachments

EXIT_CODES = {"success": 0, "agent_failure": 2, "timeout": 3,
              "infra_failure": 4, "needs_review": 5}


def _snapshot_run_inputs(cfg, attempt) -> None:
    """Copy the run config and the exact prompt files — plus any template
    attachment, which is prompt material the agent read — into the attempt
    dir, so the local record says what the attempt ran with even if the
    repo files change later."""
    if cfg.config_path and cfg.config_path.exists():
        shutil.copy2(cfg.config_path, attempt.attempt_dir / "run_config.yaml")
    prompts_dir = attempt.attempt_dir / "prompts"
    prompts_dir.mkdir(exist_ok=True)
    for path in [*prompt_file_paths(cfg), *template_attachments(cfg)]:
        shutil.copy2(path, prompts_dir / path.name)
    for src, _ws_name in prompt_extra_paths(cfg):
        shutil.copy2(src, prompts_dir / src.name)


def build_source(cfg, task_id: int) -> LocalSource:
    """The task source for one task id."""
    return LocalSource(task_id, cfg.data_root)


def stage_attempt(cfg, source):
    """Fetch the task, seed the workspace, write PROMPT.md — everything an
    attempt does before the agent starts. Shared with run_sweep --dry-run so
    a dry run resolves exactly what a real one would.

    Returns (spec, attempt, seeded attachments, prompt_version)."""
    staging = cfg.workspaces_dir / f"_staging_{datetime.now():%Y%m%d_%H%M%S}_{os.getpid()}"
    try:
        spec = source.fetch(staging)
        # The template's attachments (v12: House_Standards_v1.md) ride into
        # starting_files/ with the task inputs; a missing one is caught here,
        # before any row can be written.
        seeded = seed_template_attachments(cfg, spec)
        attempt = create_attempt(cfg.workspaces_dir, spec)
        _snapshot_run_inputs(cfg, attempt)
    finally:
        shutil.rmtree(staging, ignore_errors=True)  # inputs are copied into the workspace
    _, prompt_version = build_prompt(cfg, spec, attempt.workspace, attempt=attempt)
    return spec, attempt, seeded, prompt_version


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run one coding-agent task attempt")
    parser.add_argument("--config", required=True, help="run config YAML (run_configs/*.yaml)")
    parser.add_argument("--task-id", type=int, required=True, help="benchmark task id (tasks/task_id=<N>)")
    args = parser.parse_args(argv)

    load_dotenv_if_present()
    cfg = load_config(args.config)
    api_key = resolve_secrets(cfg)

    print(f"▶ {cfg.agent_model_name} | sandbox={cfg.sandbox.mode} ({cfg.sandbox.image})")
    print(f"  pinned by identity: {json.dumps(cfg.identity.settings())}")
    print(f"  local roots: {describe_local_target()}")
    print(f"  attempts: {cohort_dir(cfg.output_root, cfg.agent_model_name)}")

    # 1. Fetch task + seed workspace + prompt (any failure here is infra, no
    #    row written). Same helper run_sweep --dry-run resolves with.
    try:
        spec, attempt, seeded, prompt_version = stage_attempt(cfg, build_source(cfg, args.task_id))
    except Exception as e:  # noqa: BLE001 — classified, reported, non-zero exit
        print(f"❌ infra_failure during task staging: {e}")
        return EXIT_CODES["infra_failure"]

    print(f"▶ task {spec.task_id} ({spec.task_name}) | {cfg.agent.cli}/{cfg.agent.model}")
    print(f"  attempt dir: {attempt.attempt_dir}")
    if seeded:
        print(f"  seeded template attachments: {', '.join(p.name for p in seeded)}")
    prompt_size = (attempt.workspace / "PROMPT.md").stat().st_size
    print(f"  prompt_version={prompt_version} ({prompt_size:,} bytes)")

    # 2. Run the agent in the sandbox.
    cmd = build_command(cfg.agent, relay=cfg.record_trajectory and cfg.sandbox.mode == "docker")
    env = agent_env(cfg.agent, cfg.api_key_env, api_key)
    sandbox = run_in_sandbox(cfg, cmd, env, attempt.workspace, attempt.attempt_dir)
    print(f"  agent finished: exit={sandbox.exit_code} "
          f"timed_out={sandbox.timed_out} duration={sandbox.duration_seconds:.0f}s")

    # 2b. Compress the trajectory capture (if any) into an attempt artifact.
    traj = attempt.attempt_dir / "trajectory" / "trajectory.jsonl"
    if traj.exists() and traj.stat().st_size:
        with open(traj, "rb") as fin, gzip.open(attempt.attempt_dir / "trajectory.jsonl.gz", "wb") as fout:
            shutil.copyfileobj(fin, fout)
        steps = sum(1 for _ in open(traj, errors="replace"))
        print(f"  trajectory: {steps} API calls captured "
              f"({(attempt.attempt_dir / 'trajectory.jsonl.gz').stat().st_size:,} bytes gz)")

    # 3. Telemetry (best-effort, never fatal).
    telemetry = parse_transcript(sandbox.transcript_path, cfg.agent.cli)
    write_telemetry(attempt.attempt_dir, telemetry)
    if telemetry.get("totals"):
        print(f"  tokens: {telemetry['totals']} | cost: {telemetry.get('cost_usd')}")

    # 4. Verdict.
    verdict = validate(attempt, sandbox, cfg.limits.junk_seconds)
    write_verdict(attempt, verdict, sandbox)
    print(f"  verdict: {verdict.status} — {verdict.reason}")

    # 5. Record.
    try:
        summary = record(cfg, spec, attempt, sandbox, verdict, telemetry, prompt_version)
    except Exception as e:  # noqa: BLE001 — recording failure must be loud but classified
        print(f"❌ infra_failure during recording (attempt artifacts kept at "
              f"{attempt.attempt_dir}): {e}")
        return EXIT_CODES["infra_failure"]

    print(json.dumps(summary, indent=2))
    return EXIT_CODES[verdict.status]


if __name__ == "__main__":
    sys.exit(main())
