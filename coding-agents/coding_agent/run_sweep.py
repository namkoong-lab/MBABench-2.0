"""Batch entry point: every task in a run config's range, one attempt each.

    python -m coding_agent.run_sweep --config run_configs/claude_code_fable_5_1_max.yaml [--dry-run]

run_task.py is still the unit of work — a sweep decides which task ids to call
it with and in what order, and calls it as its own process, so an attempt
behaves exactly as it does on its own. Tasks are the ids present under
<data_root>/tasks/, ascending, inside the range (`tasks:` in the config,
narrowed by --task-ids or --start/--end); tasks this cohort already has a
non-failed row for in <output_root>/attempts/<agent_model_name>/
task_attempts.jsonl (under this prompt version) are skipped, so re-running
the same command is the resume path. An attempt that ends as infra_failure
or needs_review writes no row and is picked up again the same way.

--dry-run resolves each task the way a real attempt does (task fetched,
workspace seeded, PROMPT.md written) and prints what the agent would get and
where the result would land, without starting a container. It needs no API
key and no Docker.
"""
import argparse
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

from .config import TaskRange, load_config, load_dotenv_if_present, resolve_secrets
from .local_sink import attempted_task_ids, cohort_dir
from .prompt_builder import parse_prompt_version, template_name
from .repo_config import describe_local_target
from .run_task import EXIT_CODES, build_source, stage_attempt
from .task_source import local_task_ids

OK_EXITS = {EXIT_CODES["success"], EXIT_CODES["agent_failure"], EXIT_CODES["timeout"]}


def parse_task_ids(spec: str) -> list:
    """"1-101", "3", "1,5,9-12" -> a sorted list of ids."""
    ids = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part.lstrip("-"):
            first, _, last = part.partition("-")
            ids.update(range(int(first), int(last) + 1))
        else:
            ids.add(int(part))
    return sorted(ids)


def dry_run_one(cfg, task_id: int) -> int:
    """Resolve one task as far as the container and print what it resolved to."""
    print(f"\n=== task {task_id} (dry run) ===")
    try:
        spec, attempt, seeded, prompt_version = stage_attempt(cfg, build_source(cfg, task_id))
    except Exception as e:  # noqa: BLE001 — same classification a real run gives it
        print(f"❌ infra_failure during task staging: {e}")
        return EXIT_CODES["infra_failure"]
    try:
        prompt = attempt.workspace / "PROMPT.md"
        print(f"  task_name: {spec.task_name}")
        print(f"  PROMPT.md: {prompt} ({prompt.stat().st_size:,} bytes, prompt_version={prompt_version})")
        for rel in sorted(attempt.manifest):
            print(f"  staged: {rel}")
        if seeded:
            print(f"  template attachments: {', '.join(p.name for p in seeded)}")
        dest = cohort_dir(cfg.output_root, cfg.agent_model_name)
        print(f"  would write: {dest}/task_id={task_id}/<YYYYmmdd_HHMMSS>/ (+ a row in {dest}/task_attempts.jsonl)")
        print(f"  would run: {cfg.agent.cli}/{cfg.agent.model} in {cfg.sandbox.image} (not started)")
        return 0
    finally:
        shutil.rmtree(attempt.attempt_dir, ignore_errors=True)
        # a dry run leaves no trace: drop the workspaces root too when it is now empty
        try:
            attempt.attempt_dir.parent.rmdir()
        except OSError:
            pass


def attempt_one(config_path: str, task_id: int, capture: bool) -> tuple:
    """One attempt, as its own run_task process."""
    cmd = [sys.executable, "-m", "coding_agent.run_task",
           "--config", config_path, "--task-id", str(task_id)]
    proc = subprocess.run(cmd, capture_output=capture, text=True)
    return task_id, proc.returncode, (proc.stdout or "") + (proc.stderr or "") if capture else ""


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Run every task in a run config's range, one attempt each")
    parser.add_argument("--config", required=True, help="run config YAML (run_configs/*.yaml)")
    parser.add_argument("--task-ids", help='ids to attempt, e.g. "1-101" or "1,5,9-12" '
                                           "(default: the config's tasks range)")
    parser.add_argument("--start", type=int, help="first task id (alternative to --task-ids)")
    parser.add_argument("--end", type=int, help="last task id")
    parser.add_argument("--workers", type=int, default=1,
                        help="attempts to run at once (default 1, sequential)")
    parser.add_argument("--dry-run", action="store_true",
                        help="resolve every task and print what it would run; start nothing "
                             "(no API key, no Docker)")
    parser.add_argument("--redo", action="store_true",
                        help="attempt every task in range, even ones that already have a row")
    parser.add_argument("--list", action="store_true",
                        help="print the task ids this sweep would attempt and exit")
    args = parser.parse_args(argv)

    load_dotenv_if_present()
    cfg = load_config(args.config)
    if args.workers < 1:
        parser.error("--workers must be at least 1")
    # Same fail-fast the per-attempt runner does, once instead of per task —
    # except on a dry run or --list, which start no agent and need no key.
    if not (args.dry_run or args.list):
        resolve_secrets(cfg)

    rng = cfg.tasks or TaskRange()
    if args.start is not None or args.end is not None:
        rng = TaskRange(first=args.start or rng.first, last=args.end or rng.last)
    wanted = set(parse_task_ids(args.task_ids)) if args.task_ids else None

    prompt_version = parse_prompt_version(cfg.system_prompt, template_name(cfg.template_version))

    ids = local_task_ids(cfg.data_root, rng.first, rng.last)
    if wanted is not None:
        ids = [i for i in ids if i in wanted]

    done = set() if args.redo else attempted_task_ids(cfg.output_root, cfg.agent_model_name, prompt_version)
    todo = [i for i in ids if i not in done]
    print(f"▶ {cfg.agent_model_name} pv={prompt_version} tasks {rng}"
          + (" (dry run)" if args.dry_run else f" workers={args.workers}"))
    print(f"  local roots: {describe_local_target()}")
    print(f"  {len(ids)} in range, {len(ids) - len(todo)} already recorded, {len(todo)} to run")
    if args.list:
        print(" ".join(str(i) for i in todo))
        return 0

    if args.dry_run:
        worst = max((dry_run_one(cfg, i) for i in todo), default=0)
        print(f"\n▶ dry run finished: {len(todo)} tasks resolved, nothing started")
        return 0 if worst == 0 else worst

    results = {}
    if args.workers == 1:
        for n, task_id in enumerate(todo, 1):
            print(f"\n=== [{n}/{len(todo)}] task {task_id} ===", flush=True)
            results[task_id] = attempt_one(args.config, task_id, capture=False)[1]
    else:
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = [pool.submit(attempt_one, args.config, i, True) for i in todo]
            for n, future in enumerate(futures, 1):
                task_id, code, output = future.result()
                print(f"\n=== [{n}/{len(todo)}] task {task_id} -> exit {code} ===", flush=True)
                print(output, end="", flush=True)
                results[task_id] = code

    bad = {i: c for i, c in results.items() if c not in OK_EXITS}
    print(f"\n▶ sweep finished: {len(todo)} attempted, {len(bad)} ended on infra/review "
          f"({', '.join(f'{i}:{c}' for i, c in sorted(bad.items())) or 'none'})")
    return 0 if not bad else max(bad.values())


if __name__ == "__main__":
    sys.exit(main())
