"""Workspace: the agent's disposable per-attempt desk.

Layout handed to the agent:
    workspace/
      PROMPT.md
      starting_files/<inputs>
      solution.xlsx        (agent must create)

The seeded-file manifest (sha256 of every input) lives OUTSIDE the workspace,
in the attempt dir, so validation can prove the agent produced new work.
Template attachments (v12: House_Standards_v1.md) are seeded through the same
path and so appear in the manifest and the PROMPT.md listing — acceptable: an
agent that hands back the standards file as solution.xlsx fails the hash
check like any other copied input.
"""
import hashlib
import json
import os
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .config import RunConfig, template_attachments
from .task_source import TaskSpec


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def seed_template_attachments(cfg: RunConfig, spec: TaskSpec) -> list[Path]:
    """Append the template's declared attachments to spec.starting_files so
    create_attempt copies them beside the task inputs. Returns what was
    added. Raises (-> infra_failure, no row) if one is missing or would
    shadow a task input — a silent overwrite either way would change the
    task the agent sees."""
    attachments = template_attachments(cfg)
    task_names = {p.name for p in spec.starting_files}
    clash = [p.name for p in attachments if p.name in task_names]
    if clash:
        raise FileExistsError(
            f"template attachment(s) {clash} collide with a task starting file"
        )
    spec.starting_files = list(spec.starting_files) + attachments
    return attachments


@dataclass
class Attempt:
    attempt_dir: Path      # holds workspace/ + all runner-side artifacts
    workspace: Path
    manifest: dict         # {relative filename: sha256}
    started_at: datetime


def create_attempt(workspaces_root: Path, spec: TaskSpec) -> Attempt:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    label = f"task{spec.task_id}" if spec.task_id is not None else spec.task_name
    safe_label = "".join(c if c.isalnum() or c in "-_" else "_" for c in str(label))[:60]
    # ts alone collides when two tracks start the same second — include pid
    attempt_dir = workspaces_root / f"{safe_label}_{ts}_{os.getpid()}"
    workspace = attempt_dir / "workspace"
    files_dir = workspace / "starting_files"
    files_dir.mkdir(parents=True, exist_ok=False)

    manifest = {}
    for src in spec.starting_files:
        dest = files_dir / src.name
        shutil.copy2(src, dest)
        if dest.stat().st_size == 0:
            raise IOError(f"Seeded file is empty: {dest}")
        manifest[f"starting_files/{src.name}"] = sha256_file(dest)
    if not manifest:
        raise IOError("Workspace seeded with zero files")

    (attempt_dir / "manifest.json").write_text(json.dumps(manifest, indent=2))
    return Attempt(
        attempt_dir=attempt_dir,
        workspace=workspace,
        manifest=manifest,
        started_at=datetime.now(),
    )
