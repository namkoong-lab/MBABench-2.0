"""Resolve a run's `prompt_version` to the prompt file(s) it selects.

The version is the single key that picks the text, so a task_attempts row
tells you what the agent was asked to do. The map lives in
tasks_configs/prompts/registry.yaml (data, not code, so adding a prompt set
is a one-file change).

This registry, keyed by `prompt_version`, is the one way to choose prompts.
The pre-registry keys that bypassed it are deprecated and gated in
infra/configs/loader.py.

Typical usage:

    from infra.configs import resolve_prompt_files
    files = resolve_prompt_files(cfg)     # ["tasks_configs/prompts/v2_1.txt"]
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import yaml

from .loader import ConfigError

_REPO_ROOT = Path(__file__).resolve().parents[2]
REGISTRY_PATH = _REPO_ROOT / "tasks_configs" / "prompts" / "registry.yaml"


class PromptVersionError(ConfigError):
    """The run's prompt_version cannot be resolved to prompt files."""


@dataclass(frozen=True)
class PromptVersion:
    """One registered version: one number, one prompt set."""

    version: int
    files: tuple[str, ...]  # repo-root-relative, in send order
    label: str = ""
    description: str = ""
    # Files uploaded WITH the task's starting files whenever this version
    # runs (e.g. the house standards). Repo-root-relative; `..` segments are
    # allowed so a version can point at a monorepo-level file. Declared here
    # rather than in a run config so the recorded prompt_version and the
    # attachments the agent saw cannot disagree.
    attachments: tuple[str, ...] = ()


def load_registry(path: Path = REGISTRY_PATH) -> dict[int, PromptVersion]:
    """Parse registry.yaml into {version: PromptVersion}.

    Raises PromptVersionError on a missing/malformed file rather than
    returning an empty map — a silently empty registry would surface as
    "unknown prompt_version" and send the operator hunting the wrong bug.
    """
    if not path.exists():
        raise PromptVersionError(
            f"Prompt registry not found: {path}. It maps prompt_version to "
            f"prompt files; without it no run can resolve a prompt_version."
        )
    try:
        with open(path) as f:
            data = yaml.safe_load(f) or {}
    except yaml.YAMLError as e:
        raise PromptVersionError(f"Prompt registry is not valid YAML ({path}):\n{e}")

    versions = data.get("versions")
    if not isinstance(versions, dict) or not versions:
        raise PromptVersionError(
            f"Prompt registry has no `versions:` mapping: {path}"
        )

    out: dict[int, PromptVersion] = {}
    for raw_version, entry in versions.items():
        try:
            version = int(raw_version)
        except (TypeError, ValueError):
            raise PromptVersionError(
                f"Prompt registry key {raw_version!r} is not an integer "
                f"version ({path})."
            )
        if not isinstance(entry, dict):
            raise PromptVersionError(
                f"Prompt registry entry {version} must be a mapping ({path})."
            )
        files = entry.get("files")
        if isinstance(files, str):
            files = [files]
        if not isinstance(files, list) or not files:
            raise PromptVersionError(
                f"Prompt registry entry {version} has no `files:` list ({path})."
            )
        attachments = entry.get("attachments") or []
        if isinstance(attachments, str):
            attachments = [attachments]
        if not isinstance(attachments, list):
            raise PromptVersionError(
                f"Prompt registry entry {version} has a non-list "
                f"`attachments:` ({path})."
            )
        out[version] = PromptVersion(
            version=version,
            files=tuple(str(f) for f in files),
            label=entry.get("label") or "",
            description=(entry.get("description") or "").strip(),
            attachments=tuple(str(a) for a in attachments),
        )
    return out


def _config_prompt_version(cfg: SimpleNamespace) -> Any:
    """The run's prompt_version, with the two config keys reconciled.

    `prompt_version` (top level, engine-facing) and `agent.prompt_version`
    (the value written to task_attempts.prompt_version) are separate knobs
    that have always defaulted to the same number. A config that moved only
    one of them would send one prompt set and label the row with another, so
    a disagreement is refused rather than silently resolved.
    """
    top = getattr(cfg, "prompt_version", None)
    agent = getattr(getattr(cfg, "agent", None), "prompt_version", None)
    if top is not None and agent is not None and top != agent:
        raise PromptVersionError(
            f"prompt_version={top!r} but agent.prompt_version={agent!r}. "
            f"The first selects the prompt files and is sent to the engine; "
            f"the second is written to task_attempts.prompt_version. A "
            f"mismatch records the row under a version the agent was never "
            f"sent. Set both to the same value (or only one)."
        )
    return top if top is not None else agent


def resolve_prompt_files(
    cfg: SimpleNamespace, registry_path: Path = REGISTRY_PATH
) -> list[str]:
    """Repo-relative prompt file paths for this run, in send order.

    Does not read the prompt text — existence is checked by run.py's
    preflight and the text is loaded by the engine's resolve_prompts(), both
    of which resolve relative paths against the repo root.
    """
    version = _config_prompt_version(cfg)
    if version is None:
        raise PromptVersionError(
            "No prompt_version set — nothing identifies which prompt to "
            "send. Set prompt_version to an entry in "
            "tasks_configs/prompts/registry.yaml."
        )

    registry = load_registry(registry_path)
    try:
        version = int(version)
    except (TypeError, ValueError):
        raise PromptVersionError(
            f"prompt_version={version!r} is not an integer. Known versions: "
            f"{sorted(registry)}."
        )

    entry = registry.get(version)
    if entry is None:
        known = ", ".join(
            f"{v} ({registry[v].label})" for v in sorted(registry)
        )
        raise PromptVersionError(
            f"prompt_version={version} is not in the prompt registry "
            f"({registry_path.name}). Known versions: {known}. Add an entry "
            f"there for the prompt set this run should send."
        )

    return list(entry.files)


def resolve_prompt_attachments(
    cfg: SimpleNamespace, registry_path: Path = REGISTRY_PATH
) -> list[Path]:
    """Absolute paths of the files this run's prompt_version attaches.

    Resolved against the repo root (so `../house_standards/...` reaches the
    monorepo). Every declared attachment must exist: a version that promises
    the agent a file it cannot deliver would record rows under a prompt the
    agent never fully received, so a missing file is refused here, before
    any browser opens.
    """
    version = _config_prompt_version(cfg)
    if version is None:
        return []
    registry = load_registry(registry_path)
    try:
        entry = registry.get(int(version))
    except (TypeError, ValueError):
        entry = None
    if entry is None:
        return []
    out: list[Path] = []
    for rel in entry.attachments:
        path = (registry_path.parents[2] / rel).resolve()
        if not path.is_file():
            raise PromptVersionError(
                f"prompt_version={entry.version} declares attachment {rel!r} "
                f"but {path} does not exist. Restore the file or register a "
                f"new version without it; never run a version whose "
                f"attachments cannot be delivered."
            )
        out.append(path)
    return out


def describe_prompt_version(cfg: SimpleNamespace, files: list[str]) -> str:
    """One log line naming the resolved prompt set and where it came from."""
    version = _config_prompt_version(cfg)
    try:
        entry = load_registry().get(int(version)) if version is not None else None
    except (PromptVersionError, TypeError, ValueError):
        entry = None
    label = f" {entry.label!r}" if entry and entry.label else ""
    n = len(files)
    turns = "1 turn" if n == 1 else f"{n} turns"
    line = f"prompt_version={version}{label} -> {turns}: {', '.join(files)}"
    if entry and entry.attachments:
        line += f"; attachments: {', '.join(entry.attachments)}"
    return line
