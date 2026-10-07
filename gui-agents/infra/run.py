"""task-io-driven runner for gui-agents.

Reads configs from infra/configs/, builds the task source (the downloaded
benchmark under <data_root>/tasks/) and the attempt sink (local files under
<output_root>/attempts/<agent_model_name>/), and drives the
claude_web_engine.py subprocess once per task.

Config layering (later wins), all of it project-wide — there is no
per-task override layer:

    1. infra/configs/configs.default.yaml   (schema + defaults)
    2. infra/configs/configs.yaml           (optional machine overrides:
                                             Chrome port, profile)
    3. --run-config <file>                  (the per-experiment overlay)

Where tasks are read and attempts written is not part of that stack: the
shared config <repo>/config/config.yaml names both roots (local.data_root /
local.output_root), and MBABENCH_DATA_ROOT / MBABENCH_OUTPUT_ROOT override
them. See infra/configs/repo_config.py.

Per task, the engine config is built by selecting the active provider block
from the merged cfg and assembling the dict the engine expects.

Usage (from gui-agents/):
    python -m infra.run --dry-run --run-config <file>   # resolve tasks, prompts,
                                                        # attachments and output
                                                        # paths; no browser
    python -m infra.run -y --run-config <file>          # real run
    python -m infra.run --task-id 7 --run-config <file> # one task
    python -m infra.run --start 0 --end 1 ...           # slice the task list

Exit-code contract:
    0   ran >=1 task and every attempt succeeded (also: --dry-run, or the
        user declined the interactive confirmation)
    1   ran >=1 task and >=1 attempt failed
    2   config / preflight / CLI error — nothing was attempted
    3   the source yielded no tasks (filters excluded everything, empty
        slice, or every task already has a successful attempt) —
        nothing was attempted.
    4   environment gate blocked the run before any task started (CDP
        lock held by another run), or the account hit a usage cap /
        signed-out session mid-run
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import logging
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import yaml

_THIS_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _THIS_DIR.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from claude_web_agent.claude_web_engine import (  # noqa: E402
    _sanitize_name,
    resolve_prompts,
)
from infra.configs import (  # noqa: E402
    ConfigError,
    describe_local_roots,
    describe_prompt_version,
    load_configs,
    resolve_agent_identity,
    resolve_prompt_attachments,
    resolve_prompt_files,
)
from task_io import (  # noqa: E402
    AttemptResult,
    TaskSpec,
    build_sink,
    build_source,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger("infra.run")

PROVIDER_AGENT_TYPE = {"claude": "claude_web", "chatgpt": "chatgpt_web"}

# Exit codes — see the module docstring for the full contract.
EXIT_OK = 0
EXIT_TASK_FAILED = 1
EXIT_CONFIG_ERROR = 2
EXIT_NO_TASKS = 3
EXIT_ENV_BLOCKED = 4

# run_engine's sentinel for "engine subprocess exceeded the deadman and was
# killed" — matches coreutils timeout(1) so log readers recognize it.
ENGINE_RC_TIMEOUT = 124

def _ns_to_dict(obj):
    """Recursively convert SimpleNamespace (from load_configs) to plain dicts."""
    if isinstance(obj, SimpleNamespace):
        return {k: _ns_to_dict(v) for k, v in vars(obj).items()}
    if isinstance(obj, list):
        return [_ns_to_dict(v) for v in obj]
    return obj


def build_engine_config(
    cfg: SimpleNamespace, spec: TaskSpec, agent_folder: str | None = None
) -> dict:
    """Assemble the full engine-input dict for one task.

    All overrides (defaults + configs.yaml + --run-config) are already
    baked into `cfg` by the loader. This function only projects the
    active provider block + task fields into the shape the engine expects:
        {agent_type, prompts, prompts_file, prompt_version, task_name,
         task_id, upload_files, solution_name, prompt_attachments,
         <provider>_web: {...}}

    Prompt selection also lands here rather than only in main(): this is the
    function that decides what the engine receives, so a caller that skips
    main() (tests, any future entrypoint) must not silently produce a config
    with no prompts. main() resolves first so it can log the choice once per
    run, in which case cfg.prompts_file is already set and this is a no-op.

    *agent_folder* is the resolved `AgentIdentity.agent_folder`; it becomes
    the engine's `session.agent_name`, which the engine stamps into solution
    filenames and completion JSONs. main() passes it so the file says which
    model produced it (`claude_haiku_4_5`) rather than which provider lane
    (`claude_web`). Identity is resolved by the caller, not here, because an
    unknown axis combination must surface as the resolver's error or a
    preflight message — not as a build failure for a config that preflight
    is about to reject anyway. Omitted, the engine falls back to its
    per-provider default.
    """
    if not getattr(cfg, "prompts_file", None):
        cfg.prompts_file = resolve_prompt_files(cfg)
    # Same rule for the version's attachments: main() resolves them once
    # (and refuses a missing file) before any task is built; a caller that
    # skipped main() gets them here so no path sends the prompt without the
    # files it points the agent at. `is None` rather than falsiness — an
    # empty list is a resolved answer, not a missing one.
    if getattr(cfg, "prompt_attachments", None) is None:
        cfg.prompt_attachments = [str(p) for p in resolve_prompt_attachments(cfg)]
    base = _ns_to_dict(cfg)

    provider = cfg.provider.kind
    agent_type = PROVIDER_AGENT_TYPE.get(provider, "claude_web")
    provider_block_key = f"{provider}_web"

    engine_config: dict = {
        "agent_type": agent_type,
        "prompts": list(base.get("prompts") or []),
        "prompt_version": base.get("prompt_version"),
    }

    # prompts_file carries the registry's selection to the engine, which
    # expands it into `prompts`. main() has already filled it from
    # prompt_version.
    if base.get("prompts_file"):
        engine_config["prompts_file"] = base["prompts_file"]

    # The version's attachments go AFTER the task's starting files: the
    # workbook stays first (that is the order the agents upload and verify
    # tiles in), and every task of the run gets the same files. They are
    # appended here, not to spec.upload_files, so the source's record of the
    # task's own files is untouched. `prompt_attachments` is also carried
    # separately so preflight, the prompts JSON and --dry-run can tell them
    # apart from the case files.
    attachments = list(base.get("prompt_attachments") or [])
    engine_config |= {
        "task_name": spec.task_name,
        "task_id": spec.task_id,
        "upload_files": [str(p) for p in spec.upload_files] + attachments,
        provider_block_key: copy.deepcopy(base.get(provider_block_key, {}) or {}),
    }
    if attachments:
        engine_config["prompt_attachments"] = attachments

    if agent_folder:
        block = engine_config[provider_block_key]
        block["session"] = (block.get("session") or {}) | {"agent_name": agent_folder}

    if spec.solution_name:
        engine_config["solution_name"] = spec.solution_name

    return engine_config


def _preflight_provider(provider: str, section: dict, errors: list[str]) -> None:
    """Provider checks: the identity resolver already refused unknown axis
    combinations; this catches what the agent itself cannot tolerate."""
    if provider == "claude":
        if section.get("model") is None:
            errors.append(
                "claude_web.model is null — the agent calls .lower() on it and crashes. "
                "Set claude_web.model to one of the values listed in "
                "configs.default.yaml."
            )


def preflight_check(engine_config: dict, provider: str) -> list[str]:
    """Collect all problems before we touch the browser. Empty list = OK."""
    errors: list[str] = []
    section_key = f"{provider}_web"
    section = engine_config.get(section_key, {}) or {}

    _preflight_provider(provider, section, errors)

    # chatgpt_web.project_id is optional: null/empty falls back to the
    # chatgpt.com homepage (no project scope), mirroring the claude_web
    # project_id fallback. The agent logs a warning in that case.
    # project_slug is likewise optional — some ChatGPT project URLs have no
    # slug (e.g. https://chatgpt.com/g/g-p-{id}/project with no -{slug}).

    # prompt_version labels every completion JSON and the task_attempts row.
    # The engine refuses a config without one; catch it here so --dry-run
    # shows it rather than the browser opening first. Only reachable by
    # nulling the key explicitly — configs.default.yaml supplies 204.
    if engine_config.get("prompt_version") is None:
        errors.append(
            "prompt_version is null. It names the prompt the attempt was "
            "run with, so a run cannot record one without it. Set it to a "
            "version in tasks_configs/prompts/registry.yaml (it stays a "
            "label even when prompts_file bypasses the registry)."
        )

    # prompts_file must exist and be non-empty — a benchmark run with the
    # wrong (or empty) prompt is worse than one that never starts.
    pf = engine_config.get("prompts_file")
    if pf:
        repo_root = Path(__file__).resolve().parent.parent
        for raw in [pf] if isinstance(pf, str) else pf:
            p = Path(raw)
            if not p.is_absolute():
                cand = repo_root / p
                p = cand if cand.exists() else Path.cwd() / p
            if not p.exists():
                errors.append(f"prompts_file not found: {raw}")
            elif p.stat().st_size == 0:
                errors.append(f"prompts_file is empty: {raw}")
    elif not engine_config.get("prompts"):
        errors.append(
            "No prompts configured — set prompts_file (preferred) or prompts."
        )

    # Upload files must exist on disk. The version's attachments ride in
    # upload_files and are checked by the same loop; they are named as such
    # in the message because the fix is different (restore the file /
    # re-register the version, not the task).
    upload_files = engine_config.get("upload_files") or []
    attachments = [str(a) for a in engine_config.get("prompt_attachments") or []]
    version = engine_config.get("prompt_version")
    for a in attachments:
        # Guards the append in build_engine_config: a declared attachment
        # that never reached upload_files would send the prompt without the
        # file it tells the agent to read.
        if a not in [str(u) for u in upload_files]:
            errors.append(
                f"prompt_version={version} attachment {a} is not in "
                f"upload_files — the runner must upload it with the task files"
            )
    for raw in upload_files:
        resolved = Path(str(raw))
        if not resolved.exists():
            if str(raw) in attachments:
                errors.append(
                    f"prompt_version={version} attachment not found: {resolved}"
                )
                continue
            errors.append(f"upload file not found: {resolved}")

    return errors


def collect_log_files(run_dir: Path) -> list[Path]:
    """Every log the attempt produced, in upload order.

    json_logs/ holds one completion_*.json per agent attempt (the engine
    retries internally); logs/ holds the chat transcript and the runtime
    log. The staging directory belongs to this attempt alone, so everything
    under it is in scope.
    """
    patterns = ("json_logs/*.json", "logs/conversations/*.json", "logs/*.log")
    return [
        p for pattern in patterns for p in sorted(run_dir.glob(pattern)) if p.is_file()
    ]


def _write_prompts_file(
    run_dir: Path, task_name: str, engine_config: dict, started: datetime
) -> Path | None:
    """Materialize the per-task prompt payload so the sink can store it.

    Records the prompt TEXT, not the file paths: this JSON is the only
    artifact that survives as evidence of what the agent was actually
    asked, and a path is not evidence once the file changes.

    `prompts` is empty in engine_config whenever the run uses `prompts_file`
    (the normal path — the engine expands it in the subprocess, after this
    record is written), so expand it here through the engine's own resolver
    rather than reading the key. That resolver is also the one the engine
    will use, so the recorded text is what the agent receives, not a second
    guess at it.

    Returns None only if the run genuinely has no prompts to log — the sink
    treats a missing path as "no prompt_files to record" rather than a
    failure."""
    prompts = engine_config.get("prompts") or []
    if not prompts and engine_config.get("prompts_file"):
        try:
            prompts = resolve_prompts(dict(engine_config)).get("prompts") or []
        except (FileNotFoundError, OSError) as e:
            # Preflight already checked existence, so this is close to
            # unreachable — but losing the record must not lose the run.
            logger.warning(f"Could not record prompt text: {e}")
            prompts = []
    if not prompts:
        return None
    run_dir.mkdir(parents=True, exist_ok=True)
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in task_name)
    ts = started.strftime("%Y%m%d_%H%M%S")
    path = run_dir / f"prompts_{safe_name}_{ts}.json"
    path.write_text(json.dumps(_prompts_payload(engine_config, prompts), indent=2))
    return path


def _prompts_payload(engine_config: dict, prompts: list[str]) -> dict:
    """The prompts JSON body: the text sent, the version, the files it came
    from and every attachment by content."""
    pf = engine_config.get("prompts_file")
    return {
        "prompts": prompts,
        "prompt_version": engine_config.get("prompt_version"),
        "prompts_file": [pf] if isinstance(pf, str) else list(pf or []),
        # The version's attachments, by content: the agent was handed these
        # files with the prompt, so the record must carry what they said,
        # not where they were.
        "attachments": [
            _attachment_record(a)
            for a in engine_config.get("prompt_attachments") or []
        ],
    }


def _attachment_record(path: str) -> dict:
    """{name, path, sha256, text} for one attachment, read now.

    Preflight already proved the file exists; if it still cannot be read the
    record keeps the name and the error rather than aborting the attempt —
    losing the evidence must not lose the run.
    """
    p = Path(path)
    try:
        data = p.read_bytes()
    except OSError as e:
        logger.warning(f"Could not record attachment {p.name}: {e}")
        return {"name": p.name, "path": str(p), "error": str(e)}
    return {
        "name": p.name,
        "path": str(p),
        "sha256": hashlib.sha256(data).hexdigest(),
        "text": data.decode("utf-8", errors="replace"),
    }


def _dry_run_report(
    spec: TaskSpec, engine_config: dict, sink, prompt_text_shown: set[str]
) -> None:
    """What one task would send and where its attempt would land — resolved
    the way a real run resolves it, without a browser or provider."""
    try:
        prompts = resolve_prompts(dict(engine_config)).get("prompts") or []
    except (FileNotFoundError, OSError) as e:
        logger.error(f"[DRY RUN] prompt text unavailable: {e}")
        prompts = []
    pf = engine_config.get("prompts_file")
    files = [pf] if isinstance(pf, str) else list(pf or [])
    version = engine_config.get("prompt_version")
    attachments = [str(a) for a in engine_config.get("prompt_attachments") or []]
    uploads = [str(u) for u in engine_config.get("upload_files") or []]

    logger.info(f"[DRY RUN] task_id={spec.task_id} {spec.task_name}")
    turns = "1 turn" if len(prompts) == 1 else f"{len(prompts)} turns"
    logger.info(f"[DRY RUN]   prompt_version={version} -> {turns}")
    for i, text in enumerate(prompts):
        name = files[i] if i < len(files) else "(inline)"
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]
        logger.info(
            f"[DRY RUN]     turn {i + 1}: {name} ({len(text):,} chars, "
            f"sha256 {digest})"
        )
    task_files = uploads[: len(uploads) - len(attachments)] if attachments else uploads
    logger.info(
        "[DRY RUN]   upload_files: "
        + ", ".join(task_files)
        + (
            " | prompt_version attachments after them: " + ", ".join(attachments)
            if attachments
            else ""
        )
    )
    describe = getattr(sink, "describe_destination", None)
    if callable(describe):
        try:
            task_id = int(spec.metadata.get("db_task_id") or spec.task_id)
        except (TypeError, ValueError, AttributeError):
            task_id = spec.task_id
        logger.info(f"[DRY RUN]   sink: {describe(task_id)}")
    # The prompt text itself, once per distinct prompt set (it is the same
    # for every task of a run).
    for i, text in enumerate(prompts):
        name = files[i] if i < len(files) else "(inline)"
        if name in prompt_text_shown:
            continue
        prompt_text_shown.add(name)
        logger.info(f"[DRY RUN] prompt text as sent, turn {i + 1} ({name}):")
        print(text)
        print("<<< end of prompt text >>>", flush=True)


def _workbook_rank(path: Path) -> tuple[bool, int, int]:
    """(clears_quality_floor, size_bytes, sheet_count) for ranking candidates.

    Never raises: an unreadable candidate simply fails the floor, so a corrupt
    file can never outrank a good one. Uses the same thresholds as
    check_output_quality so selection and the verdict agree.
    """
    try:
        size = path.stat().st_size
    except OSError:
        return False, 0, 0
    try:
        import openpyxl

        wb = openpyxl.load_workbook(path, read_only=True)
        n_sheets = len(wb.sheetnames)
        wb.close()
    except ImportError:
        # Without openpyxl, rank on size alone rather than rejecting everything.
        return size >= QUALITY_MIN_BYTES, size, 0
    except Exception:
        return False, size, 0
    return (size >= QUALITY_MIN_BYTES and n_sheets >= QUALITY_MIN_SHEETS), size, n_sheets


def _input_digests(inputs) -> set[str]:
    """sha1 of each starting file, for spotting a copy of one in solutions/."""
    out = set()
    for raw in inputs or []:
        try:
            out.add(hashlib.sha1(Path(raw).read_bytes()).hexdigest())
        except Exception:
            continue
    return out


def find_solution_file(
    run_dir: Path, task_name: str, solution_name: str | None, after: datetime,
    inputs=None,
) -> Path | None:
    """The attempt's deliverable workbook.

    An agent can leave SEVERAL .xlsx in solutions/: a long ChatGPT run
    routinely writes scratch workbooks alongside the real one, and every
    candidate gets renamed to include the task name, so name matching alone
    cannot separate them. Ranking by mtime picks whichever landed LAST, which
    is the scratch file as often as the deliverable (2026-08-27: a 6 KB
    bellman_test.xlsx beat the real 2.5 MB workbook on task 24).

    So rank by quality floor first, then size, then recency. If nothing clears
    the floor the newest still wins, so this never returns None where the old
    ordering returned a file — a degraded output stays the caller's problem to
    report, not this function's to hide.
    """
    solutions = run_dir / "solutions"
    if not solutions.exists():
        return None
    # Must match rename_solution_file's sanitizer; otherwise task names with
    # stripped chars (e.g. '&') fail to match the on-disk filename.
    needle = _sanitize_name(solution_name or task_name).lower()
    matches: list[Path] = []
    for p in solutions.glob("*.xlsx"):
        if p.stat().st_mtime < after.timestamp():
            continue
        if needle in p.name.lower():
            matches.append(p)
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]

    # A COPY OF THE INPUT is never the deliverable. It is routinely among the
    # downloads, it clears the quality floor, and ranking is size-first — so a
    # starting workbook larger than the model's output wins and the attempt is
    # failed by the quality gate with the real deliverable sitting beside it
    # (live 2026-09-23, task 99 TAM: 4.4 MB input copy chosen over the 2.6 MB
    # model with 87,198 formulas). Drop those candidates before ranking, but
    # never drop the last one — a caller with nothing left is worse off.
    digests = _input_digests(inputs)
    if digests and len(matches) > 1:
        kept = []
        for p in matches:
            try:
                if hashlib.sha1(p.read_bytes()).hexdigest() in digests:
                    logger.info(
                        f"Ignoring {p.name}: byte-identical to a starting file"
                    )
                    continue
            except Exception:
                pass
            kept.append(p)
        if kept:
            matches = kept
        if len(matches) == 1:
            return matches[0]

    scored = []
    for p in matches:
        ok, size, n_sheets = _workbook_rank(p)
        scored.append((ok, size, p.stat().st_mtime, p, n_sheets))
    # Path is deliberately outside the sort key — it is never compared.
    scored.sort(key=lambda t: (t[0], t[1], t[2]), reverse=True)
    best = scored[0]
    logger.info(
        f"{len(scored)} candidate workbooks in solutions/; chose "
        f"{best[3].name} ({best[1]:,} bytes, {best[4]} sheets, "
        f"quality_ok={best[0]}). Not chosen: "
        + ", ".join(f"{s[3].name} ({s[1]:,} bytes)" for s in scored[1:])
    )
    return best[3]


def collect_extra_workbooks(run_dir: Path, solution_file: Path | None) -> list[Path]:
    """Every OTHER workbook the agent left in solutions/.

    Stored after the solution so a mis-pick costs a re-pointer instead of
    the whole run: the bytes are kept either way. Before this, only the
    chosen file was kept, so picking wrong meant the real workbook was lost
    (2026-08-27, task 24 — 131 minutes nearly lost).

    Order matters downstream: the sink stores [solution, *log_files] and the
    judge grades the FIRST xlsx, so these must be appended after the logs,
    never prepended.
    """
    solutions = run_dir / "solutions"
    if not solutions.exists():
        return []
    try:
        chosen = solution_file.resolve() if solution_file else None
    except OSError:
        chosen = None
    extras = []
    for p in sorted(solutions.glob("*.xlsx")):
        if not p.is_file():
            continue
        try:
            if chosen is not None and p.resolve() == chosen:
                continue
        except OSError:
            continue
        extras.append(p)
    if extras:
        logger.info(
            f"Preserving {len(extras)} non-selected workbook(s) alongside the "
            f"solution: " + ", ".join(p.name for p in extras)
        )
    return extras


# Post-run output quality gate. Heuristic floor to catch obviously-degraded
# workbooks (e.g. a tiny stub produced when a ChatGPT "content failed to load"
# disruption truncated the model). Tunable.
QUALITY_MIN_BYTES = 12000
QUALITY_MIN_SHEETS = 3


def check_output_quality(
    solution_file: Path | None, input_files: list | None = None
) -> tuple[bool, str]:
    """Return (ok, reason). ok=False flags a suspected-degraded output.

    2026-09-11 (full run, attempts 1434/1461/1523): ChatGPT Work mode lists the
    user's own uploaded workbook among the sandbox artifacts, so when the sandbox
    dies mid-task and the model says "I couldn't produce a .xlsx", the backend
    download still returns the STARTING FILE (or a formula-free checkpoint) and
    the row was recorded success. Two extra checks close that: a workbook that is
    byte-identical to any input file, or that carries no formulas at all, is not
    a deliverable.

    Runs in infra.run AFTER the engine subprocess returns — i.e. OUTSIDE the
    engine's own retry loop — so it only records a verdict and never triggers
    a re-run. Fails OPEN on inspection errors (can't break the pipeline over a
    false alarm) except a missing/unopenable workbook, which IS a failure.
    """
    if solution_file is None:
        return False, "no solution workbook was produced"
    try:
        size = Path(solution_file).stat().st_size
    except OSError as e:
        return False, f"solution file not accessible: {e}"
    try:
        import openpyxl

        wb = openpyxl.load_workbook(solution_file, read_only=True)
        n_sheets = len(wb.sheetnames)
        wb.close()
    except ImportError:
        logger.warning("openpyxl unavailable; skipping output quality gate")
        return True, ""
    except Exception as e:
        return False, (
            f"solution workbook could not be opened ({type(e).__name__}); "
            f"likely corrupt/partial"
        )
    if size < QUALITY_MIN_BYTES or n_sheets < QUALITY_MIN_SHEETS:
        return False, (
            f"output below quality floor: {size} bytes / {n_sheets} sheet(s) "
            f"(min {QUALITY_MIN_BYTES} bytes, {QUALITY_MIN_SHEETS} sheets) — "
            f"suspected content-load disruption"
        )
    identical = _identical_input(solution_file, input_files or [])
    if identical is not None:
        return False, (
            f"solution workbook is byte-identical to the input file "
            f"{identical.name} — the agent produced no deliverable"
        )
    n_formulas = _count_formulas(solution_file)
    if n_formulas == 0:
        return False, (
            "solution workbook contains no formulas — not a model "
            "(starting file or checkpoint picked up as the artifact)"
        )
    return True, ""


def _identical_input(solution_file: Path, input_files: list) -> Path | None:
    """The input file whose bytes equal the solution's, else None. Fails open."""
    import hashlib

    def digest(p: Path) -> str:
        h = hashlib.sha256()
        with open(p, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()

    try:
        sol = digest(Path(solution_file))
        for raw in input_files:
            p = Path(raw)
            if p.suffix.lower() in (".xlsx", ".xlsm") and p.is_file() and digest(p) == sol:
                return p
    except Exception as e:  # noqa: BLE001
        logger.warning(f"identical-input check skipped: {e}")
    return None


def _count_formulas(solution_file: Path) -> int | None:
    """Formula cells across all sheets, read from the sheet XML so namespaced
    writers (<x:f>) count too. None on any error (fails open)."""
    import re
    import zipfile

    try:
        with zipfile.ZipFile(solution_file) as z:
            parts = [n for n in z.namelist() if "worksheets/" in n and n.endswith(".xml")]
            return sum(
                len(re.findall(rb"<(?:\w+:)?f[\s>]", z.read(n))) for n in parts
            )
    except Exception as e:  # noqa: BLE001
        logger.warning(f"formula count skipped: {e}")
        return None


# Signatures of a lane-level environment stop that is NOT a usage cap: the
# browser session no longer authenticates. Signing in is the operator's job,
# so the lane must stop rather than spend its remaining tasks on failed rows.
AUTH_LOST_LOG_SIGNATURES = (
    "Authentication required",
)


USAGE_CAP_LOG_SIGNATURES = (
    "Rate limit persisted",            # claude: limit banner during the wait
    "Usage/plan limit persisted",      # chatgpt: limit banner during the wait
    "Rate limited before prompts",     # either: get_state() saw the banner
    "Usage limit reached",             # claude 5-hour session banner text
    "Add credits to continue",         # chatgpt usage-credit exhaustion
)


def usage_cap_hit(run_dir: Path) -> str | None:
    """Why the task's failure looks like an ACCOUNT CAP, else None.

    2026-09-11 (full run): four separate cap events (two ChatGPT, two Claude)
    each left the lane walking its list and recording a failed row per task
    every few minutes until a human killed it. A cap is a property of the
    account, not the task, so the runner stops the lane instead (exit 4) —
    the row for the task that hit it is still published, and a relaunch after
    the reset resumes at that task because skip_already_attempted ignores
    failed rows. Reads the engine's completion JSONs (task_status
    'rate_limited' = banner seen before prompts) and its log (banner seen
    mid-wait). Fails open on any read error.
    """
    try:
        for cj in sorted(Path(run_dir).glob("json_logs/*.json")):
            try:
                data = json.loads(cj.read_text())
            except Exception:
                continue
            for t in data.get("tasks", []) or []:
                if t.get("task_status") == "rate_limited":
                    return f"engine reported task_status=rate_limited ({cj.name})"
        for lg in sorted(Path(run_dir).glob("logs/*.log")):
            try:
                text = lg.read_text(errors="ignore")
            except Exception:
                continue
            for sig in USAGE_CAP_LOG_SIGNATURES + AUTH_LOST_LOG_SIGNATURES:
                if sig in text:
                    return f"engine log contains {sig!r} ({lg.name})"
    except Exception as e:  # noqa: BLE001
        logger.warning(f"usage-cap check skipped: {e}")
    return None


def _kill_engine_tree(proc: subprocess.Popen) -> None:
    """SIGTERM the engine's whole process group, escalate to SIGKILL.

    The engine is launched with start_new_session=True, so killing its group
    reaps anything it spawned without touching run.py's own group (and,
    critically, without signalling the shared Chrome — that lives in a
    separate service/session and is never a child of the engine)."""
    if proc.poll() is not None:
        return
    try:
        pgid = os.getpgid(proc.pid)
        os.killpg(pgid, signal.SIGTERM)
    except (ProcessLookupError, PermissionError, OSError):
        proc.terminate()
    try:
        proc.wait(timeout=10)
        return
    except subprocess.TimeoutExpired:
        pass
    try:
        pgid = os.getpgid(proc.pid)
        os.killpg(pgid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError, OSError):
        proc.kill()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        logger.error(f"engine pid={proc.pid} survived SIGKILL (?)")


def run_engine(
    engine_config: dict,
    engine_script: Path,
    timeout: int | None,
    cdp_port: int | None = None,
) -> int:
    """Run the engine subprocess, streaming its output. Returns its exit
    code, or ENGINE_RC_TIMEOUT if the deadman killed it.

    The deadman is enforced by proc.wait(timeout=...) on the main thread
    while a daemon thread pumps stdout. The previous implementation read
    stdout to EOF on the main thread BEFORE calling wait(timeout=...), so a
    wedged engine that stopped writing but never exited (e.g. a hung
    Playwright await after the renderer died) blocked forever and the
    timeout never even started — the deadman existed but could not fire.

    The finally-block guarantees the engine's process group is reaped on
    ANY exit from this function (deadman, KeyboardInterrupt, SIGTERM via
    the SystemExit handler in main, unexpected exception) — run.py never
    leaves an orphaned engine driving the shared Chrome.
    """
    # The CDP port is embedded in the temp-config filename so external
    # supervisors (e.g. infra/overnight) can pgrep-scope the engine child to
    # ONE browser instance — parallel runs on distinct ports must never sweep
    # or kill each other's engine.
    port_tag = f"cdp{cdp_port}_" if cdp_port is not None else ""
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".yaml",
        prefix=f"gui_agents_{port_tag}{engine_config.get('task_name', 'task')}_",
        delete=False,
    ) as f:
        yaml.safe_dump(engine_config, f, default_flow_style=False)
        tmp_path = Path(f.name)
    proc: subprocess.Popen | None = None
    try:
        cmd = [
            sys.executable,
            str(engine_script),
            "--config",
            str(tmp_path),
            "--no-hold",
        ]
        logger.info(f"Engine: {' '.join(cmd)}")
        if timeout:
            logger.info(f"Engine deadman: {timeout}s")
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            start_new_session=True,
        )
        assert proc.stdout is not None

        def _pump(stream) -> None:
            for line in iter(stream.readline, ""):
                print(line, end="", flush=True)

        pump = threading.Thread(target=_pump, args=(proc.stdout,), daemon=True)
        pump.start()
        try:
            rc = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            logger.error(f"Engine exceeded {timeout}s deadman — killing process group")
            _kill_engine_tree(proc)
            rc = ENGINE_RC_TIMEOUT
        pump.join(timeout=5)  # flush whatever output remains
        return rc
    finally:
        if proc is not None:
            _kill_engine_tree(proc)
        try:
            tmp_path.unlink()
        except Exception:
            pass


def _resolve_cdp_port(cfg: SimpleNamespace, provider: str) -> int | None:
    """Active provider block's browser.cdp_port, or None if not configured."""
    try:
        return int(getattr(cfg, f"{provider}_web").browser.cdp_port)
    except (AttributeError, TypeError, ValueError):
        return None


def _acquire_cdp_lock(port: int):
    """Advisory exclusive lock on the shared Chrome's CDP port.

    Two engines driving one Chrome corrupt BOTH runs (interleaved clicks,
    stolen focus), so a second run.py targeting the same port must fail
    fast instead of starting. flock releases automatically when this
    process exits — including a SIGKILL — so a dead run can never wedge
    the port shut.

    Returns (lock_file_handle, None) on success — caller must keep the
    handle alive for the duration of the run — or (None, holder_info) if
    another process holds the lock. Fails OPEN (None, None handled by
    caller as acquired) is deliberately NOT used: no fcntl means no lock
    semantics anywhere, so we just skip locking on such platforms.
    """
    try:
        import fcntl
    except ImportError:
        return None, None  # non-POSIX: skip locking entirely
    lock_path = Path(tempfile.gettempdir()) / f"gui_agents_cdp_{port}.lock"
    try:
        fh = open(lock_path, "a+")
    except OSError as e:
        # E.g. the file exists owned by another user. Locking is advisory
        # protection — skip it rather than block a legitimate run.
        logger.warning(f"CDP lock unavailable ({e}); continuing without it")
        return None, None
    try:
        fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        fh.seek(0)
        holder = fh.read().strip() or "unknown pid"
        fh.close()
        return None, holder
    fh.seek(0)
    fh.truncate()
    fh.write(f"pid={os.getpid()} started={datetime.now().isoformat()}")
    fh.flush()
    return fh, None


def _default_deadman(engine_config: dict, provider: str) -> int | None:
    """Conservative ceiling for the engine subprocess when --timeout is not
    given: every legitimate run — all retry attempts at their own per-task
    budget — fits under it with a wide grace margin, so it only fires on a
    truly wedged engine, whose internal guard cannot preempt a hung await.
    Returns None if the config keys aren't available, leaving the subprocess
    unlimited."""
    section = engine_config.get(f"{provider}_web", {}) or {}
    try:
        per_task = int(section.get("max_sec_per_task") or 0)
        retry = section.get("retry", {}) or {}
        attempts = int(retry.get("max_total_attempts") or 0)
    except (TypeError, ValueError):
        return None
    if per_task <= 0 or attempts <= 0:
        return None
    return per_task * attempts + 1800


def _staging_dir(cfg: SimpleNamespace, task_name: str, started: datetime) -> Path:
    """The working directory for one attempt.

    Everything the run produces lands here — the engine's solutions/,
    json_logs/ and logs/, plus the prompts JSON this module writes. The sink
    copies the lot under <output_root> and _clear_staging removes the
    directory afterwards, so the only lasting copy is the sink's.

    The name carries this runner's pid: two lanes on one machine running the
    same task list (e.g. a fable and a sol cohort over the same ids) started
    task 69 in the same second on 2026-09-06 and shared one directory — both
    engines' logs and completion JSONs interleaved, either lane's
    find_solution_file could have picked the OTHER cohort's workbook, and
    the first lane to publish would have rmtree'd the other's live staging
    dir. Timestamp + task name alone is not unique across concurrent runs.
    """
    stem = _sanitize_name(task_name)
    return (
        Path(cfg.paths.scratch_dir)
        / "attempts"
        / f"{started.strftime('%Y%m%d_%H%M%S')}_{stem}_p{os.getpid()}"
    )


def _clear_staging(run_dir: Path, sink) -> None:
    """Delete the attempt's staging directory once the sink has the files.

    A sink that only records paths would keep it — removing it would destroy
    the only copy of the workbook.
    """
    if not getattr(sink, "retains_files", False):
        logger.info(f"Run files kept at {run_dir} (sink records paths only)")
        return
    shutil.rmtree(run_dir, ignore_errors=True)


def _confirm_tasks(specs: list[TaskSpec], attachments: list[str] = ()) -> bool:
    """Print the loaded task list and ask the user to confirm."""
    print(f"\nAbout to run {len(specs)} task(s):")
    for i, spec in enumerate(specs):
        files = ", ".join(p.name for p in spec.upload_files) or "(no files)"
        print(f"  [{i}] {spec.task_name}  —  {files}")
    if attachments:
        names = ", ".join(Path(a).name for a in attachments)
        print(f"  + uploaded with every task (prompt_version attachments): {names}")
    try:
        answer = input("\nProceed? [y/N]: ").strip().lower()
    except EOFError:
        # Non-interactive stdin — treat as "no" unless --yes was passed.
        return False
    return answer in {"y", "yes"}


def main() -> int:
    parser = argparse.ArgumentParser(description="gui-agents runner (task-io driven)")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--end", type=int, default=None)
    parser.add_argument("--timeout", type=int, default=None)
    parser.add_argument(
        "--run-config",
        default=None,
        help=(
            "Overlay a run-specific YAML (provider, model, prompt_version, "
            "source.filters.task_ids, ...) on top of configs.yaml as a 3rd "
            "config layer. See infra/configs/run_configs/ for the examples. "
            "Relative paths resolve from the gui-agents/ root."
        ),
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="Skip the interactive 'proceed?' confirmation.",
    )
    parser.add_argument(
        "--task-id",
        type=int,
        default=None,
        help=(
            "Run exactly one task (by task_id). Pins source.filters.task_ids "
            "to this value and disables skip_already_attempted so the run "
            "proceeds even if an earlier attempt exists."
        ),
    )
    parser.add_argument(
        "--skip-if-attempted",
        action="store_true",
        help=(
            "Force source.filters.skip_already_attempted=True — with "
            "--task-id this overrides its default re-run behavior, so a "
            "task that already has a successful attempt becomes a no-op "
            "(exit 3) instead of a duplicate attempt."
        ),
    )
    args = parser.parse_args()

    # SIGTERM (systemctl stop, orchestrator kill) must run our finally
    # blocks — Python's default handler exits immediately, which would
    # orphan the engine subprocess (it lives in its own process group, so
    # the terminal's signals do not reach it). SystemExit unwinds through
    # run_engine's finally, reaping the engine tree.
    signal.signal(signal.SIGTERM, lambda signum, frame: sys.exit(128 + signum))

    run_config_path: Path | None = None
    if args.run_config is not None:
        run_config_path = Path(args.run_config)
        if not run_config_path.is_absolute():
            run_config_path = _REPO_ROOT / run_config_path
        if not run_config_path.exists():
            logger.error(f"--run-config file not found: {run_config_path}")
            return EXIT_CONFIG_ERROR
        with open(run_config_path) as f:
            run_config_data = yaml.safe_load(f) or {}
        if not isinstance(run_config_data, dict):
            logger.error(
                f"--run-config must be a YAML mapping at top level: "
                f"{run_config_path}"
            )
            return EXIT_CONFIG_ERROR

    try:
        cfg = load_configs(run_config_path=run_config_path)
    except ConfigError as e:
        logger.error(f"Config load failed:\n{e}")
        return EXIT_CONFIG_ERROR

    if args.task_id is not None:
        filters = getattr(cfg.source, "filters", None)
        if filters is None:
            filters = SimpleNamespace()
            cfg.source.filters = filters
        filters.task_ids = [args.task_id]
        filters.skip_already_attempted = bool(args.skip_if_attempted)
    elif args.skip_if_attempted:
        filters = getattr(cfg.source, "filters", None)
        if filters is None:
            filters = SimpleNamespace()
            cfg.source.filters = filters
        filters.skip_already_attempted = True

    provider = cfg.provider.kind

    # Prompt selection. `prompt_version` picks the files through
    # tasks_configs/prompts/registry.yaml, so the DB label and the text the
    # agent receives are the same decision.
    #
    # `prompts_file` is a deprecated config key, kept out of the schema and
    # gated in infra/configs/loader.py, which warns when a config still sets
    # it. It reaches cfg only in that case, and it still wins here so a
    # config mid-migration sends what it says it sends.
    if getattr(cfg, "prompts_file", None):
        logger.info(
            f"prompts_file={cfg.prompts_file} — prompt registry skipped; "
            f"prompt_version={getattr(cfg, 'prompt_version', None)!r} is a "
            f"label only for this run."
        )
    else:
        try:
            cfg.prompts_file = resolve_prompt_files(cfg)
        except ConfigError as e:
            logger.error(f"Prompt selection failed:\n{e}")
            return EXIT_CONFIG_ERROR
        logger.info(describe_prompt_version(cfg, cfg.prompts_file))

    # The version's attachments (e.g. the house standards), resolved once
    # here so a missing file stops the run before any task is built, and so
    # every task's upload_files and prompts JSON see the same absolute paths.
    # Resolved on the deprecated prompts_file path too: the recorded
    # prompt_version still promises them.
    try:
        cfg.prompt_attachments = [str(p) for p in resolve_prompt_attachments(cfg)]
    except ConfigError as e:
        logger.error(f"Prompt attachment resolution failed:\n{e}")
        return EXIT_CONFIG_ERROR
    if cfg.prompt_attachments:
        logger.info(
            "attachments (uploaded after each task's starting files): "
            + ", ".join(cfg.prompt_attachments)
        )

    # The sink writes cfg.agent.prompt_version to task_attempts. Copy the
    # top-level value across so the recorded version is by construction the
    # one that selected the prompts, rather than a second key that can drift
    # from it. resolve_prompt_files already refused any disagreement.
    if getattr(cfg, "prompt_version", None) is not None:
        cfg.agent.prompt_version = cfg.prompt_version

    # Where tasks are read from and attempts written to (shared config /
    # env). Logged BEFORE the build so a missing download still shows what
    # was attempted.
    logger.info(f"Local layout: {describe_local_roots()}")

    engine_script = _REPO_ROOT / "claude_web_agent" / "claude_web_engine.py"
    if not engine_script.exists():
        logger.error(f"Engine not found: {engine_script}")
        return EXIT_CONFIG_ERROR

    try:
        source = build_source(cfg)
        sink = build_sink(cfg)
    except ValueError as e:
        # Build failures here are user-facing: a missing download, an
        # unwritable output root, an unknown source/sink kind, an unknown
        # identity. Swallow the traceback and log the message raised.
        logger.error(f"Source/sink build failed:\n{e}")
        return EXIT_CONFIG_ERROR

    identity = resolve_agent_identity(cfg)
    logger.info(
        f"agent identity: model_name={identity.model_name!r} "
        f"agent_folder={identity.agent_folder!r} "
        f"agent_model_type={identity.agent_model_type!r}"
    )

    succeeded = failed = 0
    cdp_lock = None
    try:
        specs = list(source.iter_tasks())
        specs = specs[args.start : args.end]
        logger.info(f"Loaded {len(specs)} task(s) from source kind={cfg.source.kind}")

        if not specs:
            logger.warning(
                "No tasks to run (source filters excluded everything, or "
                "the slice is empty)."
            )
            return EXIT_NO_TASKS

        # Build + preflight every task BEFORE the user-confirmation prompt.
        # If any task has null configs or missing files, abort here so the
        # user never sees "About to run" for a run that's guaranteed to fail.
        prepared: list[tuple[TaskSpec, dict]] = []
        had_errors = False
        for spec in specs:
            engine_config = build_engine_config(cfg, spec, identity.agent_folder)
            errors = preflight_check(engine_config, provider)
            if errors:
                had_errors = True
                logger.error(
                    f"Preflight failed for task {spec.task_name!r} "
                    f"(provider={provider!r}):"
                )
                for e in errors:
                    logger.error(f"  - {e}")
            else:
                prepared.append((spec, engine_config))
        if had_errors:
            logger.error(
                "Fix the --run-config and re-run. configs.default.yaml lists "
                "every available key."
            )
            return EXIT_CONFIG_ERROR

        if not args.dry_run and not args.yes:
            if not _confirm_tasks(specs, cfg.prompt_attachments):
                logger.info("Aborted by user.")
                return 0

        # Environment gate — only for real runs (dry-run never touches the
        # browser). Fails BEFORE any task starts, with a distinct exit code,
        # so callers can tell "environment blocked" from "task failed".
        cdp_port = _resolve_cdp_port(cfg, provider)
        if not args.dry_run and cdp_port is not None:
            cdp_lock, holder = _acquire_cdp_lock(cdp_port)
            if holder is not None:
                logger.error(
                    f"Another run already drives Chrome on CDP port "
                    f"{cdp_port} ({holder}) — two engines on one browser "
                    f"corrupt both runs. Wait for it or use a different "
                    f"browser/port."
                )
                return EXIT_ENV_BLOCKED

        cap_stop: str | None = None
        prompt_text_shown: set[str] = set()
        for i, (spec, engine_config) in enumerate(prepared):
            idx = args.start + i
            logger.info(f"\n{'=' * 60}\nTASK {idx}: {spec.task_name}\n{'=' * 60}")

            if args.dry_run:
                _dry_run_report(spec, engine_config, sink, prompt_text_shown)
                logger.info("[DRY RUN] engine_config:")
                print(yaml.safe_dump(engine_config, default_flow_style=False))
                continue

            started = datetime.now()
            run_dir = _staging_dir(cfg, spec.task_name, started)
            engine_config["run_dir"] = str(run_dir)
            prompts_file = _write_prompts_file(
                run_dir, spec.task_name, engine_config, started
            )

            # Deadman: --timeout wins; otherwise derive a conservative
            # ceiling from the provider's own retry budget (None = the old
            # unlimited behavior if those keys aren't configured).
            deadman = (
                args.timeout
                if args.timeout is not None
                else _default_deadman(engine_config, provider)
            )
            rc = run_engine(engine_config, engine_script, deadman, cdp_port)
            finished = datetime.now()
            solution_file = find_solution_file(
                run_dir, spec.task_name, spec.solution_name, started,
                inputs=list(spec.upload_files or []),
            )
            if rc == 0:
                status = "success"
            elif rc == ENGINE_RC_TIMEOUT:
                status = "timeout"
            else:
                status = "failed"

            # Post-engine output quality gate. Runs outside the engine's retry
            # loop, so it only RECORDS the verdict — it never re-runs. A degraded
            # output is marked failed-with-reason (no retry, no deprecation);
            # the stub is still stored for inspection.
            quality_reason: str | None = None
            if status == "success":
                ok, reason = check_output_quality(
                    solution_file, list(spec.upload_files or [])
                )
                if not ok:
                    status = "failed"
                    quality_reason = reason
                    logger.warning(f"Output quality gate FAILED: {reason}")

            if status == "success":
                succeeded += 1
            else:
                failed += 1

            extra: dict = {
                "return_code": rc,
                "task_metadata": dict(spec.metadata or {}),
            }
            if quality_reason:
                extra["failure_reason"] = quality_reason

            result = AttemptResult(
                task_id=spec.task_id,
                task_name=spec.task_name,
                agent_model_name=identity.model_name,
                prompt_version=cfg.agent.prompt_version,
                status=status,
                solution_file=solution_file,
                # Extras last: attempt_files is [solution, *log_files] and the
                # judge grades the first xlsx, so the solution stays first.
                log_files=collect_log_files(run_dir)
                + collect_extra_workbooks(run_dir, solution_file),
                started_at=started.isoformat(),
                finished_at=finished.isoformat(),
                duration_seconds=round((finished - started).total_seconds(), 2),
                prompt_files=[prompts_file] if prompts_file else [],
                extra=extra,
            )
            # Decide BEFORE publish/clear: _clear_staging may delete run_dir.
            cap_reason = usage_cap_hit(run_dir) if status != "success" else None
            sink.publish(result)
            _clear_staging(run_dir, sink)
            if cap_reason:
                cap_stop = cap_reason
                logger.error(
                    f"ACCOUNT BLOCKED on task {spec.task_name} (usage cap, or a "
                    f"signed-out browser session): {cap_reason}. "
                    f"Stopping this lane (exit {EXIT_ENV_BLOCKED}) — the remaining "
                    f"{len(prepared) - i - 1} task(s) are untouched; relaunch the "
                    f"same run config once the account resets or is signed in "
                    f"again."
                )
                break

        logger.info(f"\nDone. succeeded={succeeded} failed={failed}")
    finally:
        if cdp_lock is not None:
            try:
                cdp_lock.close()  # closing the fd releases the flock
            except Exception:
                pass
        source.close()
        sink.close()

    if cap_stop:
        return EXIT_ENV_BLOCKED
    return EXIT_OK if failed == 0 else EXIT_TASK_FAILED


if __name__ == "__main__":
    sys.exit(main())
